"""
Processa arquivos de seguradoras (Carlos Venancio) — v6

Uso:
  python processa_seguradoras.py -i entrada -o saida -c config.json --data 30/09/2025 --log-dir logs

Dependências: pandas, openpyxl, xlrd
"""

import argparse
import os
import re
import sys
import json
import uuid
import unicodedata
from pathlib import Path
from datetime import datetime

import pandas as pd

# ----------------------- Utilidades -----------------------

def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.lower()
    s = re.sub(r"[\W_]+", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_header_map(df_columns):
    return {normalize_text(col): col for col in df_columns}


def find_best_match_column(header_map, variants):
    # Match exato (normalizado)
    for v in variants:
        nv = normalize_text(v)
        if nv in header_map:
            return header_map[nv]
    # Heurística por substring
    for v in variants:
        nv = normalize_text(v)
        for norm_col, orig_col in header_map.items():
            if nv and (nv in norm_col or norm_col in nv):
                return orig_col
    return None


def detect_columns(df, col_synonyms: dict) -> dict:
    header_map = build_header_map(df.columns)
    colmap = {}
    for canon, variants in col_synonyms.items():
        match = find_best_match_column(header_map, variants)
        if match:
            colmap[canon] = match
    return colmap

def drop_columns_by_contains(df, tokens):
    """    Remove colunas cujo cabeçalho normalizado contenha qualquer 
    token (case/acentos-insensível).
    """
    tokens = [normalize_text(t) for t in (tokens or []) if str(t).strip()]
    if not tokens:
        return df, []
    header_map = {col: normalize_text(col) for col in df.columns}
    to_drop = []
    for orig, norm in header_map.items():
        for t in tokens:
            if t and t in norm:
                to_drop.append(orig)
                break
    if to_drop:
        df = df.drop(columns=to_drop, errors='ignore')
    return df, to_drop


def parse_date_series(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def br_number_to_float_series(s: pd.Series) -> pd.Series:
    import numpy as np
    def conv(x):
        if pd.isna(x):
            return np.nan
        txt = str(x)
        txt = re.sub(r"[^0-9,.-]", "", txt)
        txt = txt.strip()
        if "," in txt and "." in txt:
            txt = txt.replace(".", "")
        txt = txt.replace(",", ".")
        try:
            return float(txt)
        except Exception:
            return np.nan
    return s.apply(conv)


def recompute_status_by_dates(df, col_fim, out_col="status_automatico", today_override=None):
    """Recalcula status (VIGENTE/VENCIDA) e grava em out_col (default: status_automatico)."""
    today = pd.Timestamp(today_override.date()) if today_override is not None else pd.Timestamp(datetime.now().date())
    status = []
    for _, row in df.iterrows():
        fim = row[col_fim] if col_fim in df.columns else pd.NaT
        if pd.notna(fim):
            status.append("VIGENTE" if today <= fim else "VENCIDA")
        else:
            status.append(None)
    df[out_col] = status
    return df


def sort_for_latest(df, col_data_emissao, col_num_endosso):
    sort_cols, ascending = [], []
    if col_data_emissao in df.columns:
        sort_cols.append(col_data_emissao); ascending.append(False)
    if col_num_endosso in df.columns:
        tmp = pd.to_numeric(df[col_num_endosso], errors="coerce")
        if tmp.notna().mean() >= 0.5:
            df = df.copy(); df["_endosso_num"] = tmp
            sort_cols.append("_endosso_num"); ascending.append(False)
        else:
            sort_cols.append(col_num_endosso); ascending.append(False)
    if sort_cols:
        return df.sort_values(by=sort_cols, ascending=ascending)
    return df


def read_any_file(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext == ".xlsx":
        return pd.read_excel(path, engine="openpyxl")
    elif ext == ".xls":
        return pd.read_excel(path, engine="xlrd")
    elif ext == ".csv":
        try:
            return pd.read_csv(path, encoding="utf-8-sig", sep=None, engine="python")
        except Exception:
            return pd.read_csv(path, encoding="latin-1", sep=None, engine="python")
    else:
        raise ValueError(f"Extensão não suportada: {ext}")

# -------------- Data de referência (status) --------------

def parse_reference_date(s: str):
    """Converte uma string em pd.Timestamp (date-only). Aceita '30/09/2025', '2025-09-30', '30-09-2025'."""
    if not s:
        return None
    ts = pd.to_datetime(s, dayfirst=True, errors='coerce')
    if pd.isna(ts):
        return None
    return pd.Timestamp(ts.date())

# ----------------------- Núcleo -----------------------

OUTPUT_COLUMNS_ORDER = [
    "num_apolice",
    "apolice_susep",
    "num_endosso",
    "tipo_endosso",
    "data_emissao",
    "data_inicio_vigencia",
    "data_fim_vigencia",
    "status_apolice",
    "status_automatico",
    "is",
]

def decide_mode(filename: str, cfg: dict) -> str:
    # 1) Sufixos no nome (se habilitado)
    if cfg.get("rules", {}).get("suffix_overrides", True):
        for kw in cfg["rules"].get("suffix_keywords", {}).get("agrupar", []):
            if re.search(kw, filename, flags=re.IGNORECASE):
                return "agrupar"
        for kw in cfg["rules"].get("suffix_keywords", {}).get("ultimo", []):
            if re.search(kw, filename, flags=re.IGNORECASE):
                return "ultimo"
    # 2) Padrões por seguradora/nome
    for rule in cfg["rules"].get("insurer_patterns", []):
        if re.search(rule["pattern"], filename, flags=re.IGNORECASE):
            return rule["mode"]
    # 3) Default
    return cfg["rules"].get("default_mode", "ultimo")


def process_file(path: Path, out_dir: Path, cfg: dict, ref_date: "pd.Timestamp|None" = None) -> dict:
    filename = path.name
    mode = decide_mode(filename, cfg)

    try:
        raw = read_any_file(path)
    except Exception as e:
        return {"file": filename, "status": "erro_leitura", "detalhe": str(e)}
    if raw.empty:
        return {"file": filename, "status": "vazio"}
    
    # Remoção de colunas por tokens (primeira etapa) 
    drop_tokens = (cfg.get('drop_columns_contains') or cfg.get('rules', {}).get('drop_columns_contains'))
    colunas_dropadas = []
    if drop_tokens:
        raw, colunas_dropadas = drop_columns_by_contains(raw, drop_tokens)
        
    raw.drop_duplicates(inplace=True, ignore_index=True)

    colmap = detect_columns(raw, cfg["column_synonyms"]) 
    present = [c for c in OUTPUT_COLUMNS_ORDER if c in colmap]   
    if not present:
        return {"file": filename, "status": "colunas_nao_encontradas", "detalhe": str(list(raw.columns))}
    
    df = raw[[colmap[c] for c in present]].copy()
    df.columns = present

    # Normalizações
    if "data_emissao" in df.columns:
        df["data_emissao"] = parse_date_series(df["data_emissao"])
    if "data_inicio_vigencia" in df.columns:
        df["data_inicio_vigencia"] = parse_date_series(df["data_inicio_vigencia"])
    if "data_fim_vigencia" in df.columns:
        df["data_fim_vigencia"] = parse_date_series(df["data_fim_vigencia"])
    if "is" in df.columns:
        df["is"] = br_number_to_float_series(df["is"])

    # Regras por modo
    if mode == "agrupar":
        group_keys = [c for c in ["num_apolice"] if c in df.columns] or [present[0]]
        agg_map = {}
        if "is" in df.columns: agg_map["is"] = "sum"
        if "data_emissao" in df.columns: agg_map["data_emissao"] = "max"
        if "data_inicio_vigencia" in df.columns: agg_map["data_inicio_vigencia"] = "min"
        if "data_fim_vigencia" in df.columns: agg_map["data_fim_vigencia"] = "max"
        
        for t in ["tipo_endosso", "num_endosso", "status_apolice", "apolice_susep"]:
            if t in df.columns: agg_map[t] = "last"
        result = df.groupby(group_keys, dropna=False).agg(agg_map).reset_index()
       
        if "data_fim_vigencia" in result.columns:
            result = recompute_status_by_dates(result, "data_fim_vigencia", today_override=ref_date)
    else:  
        # ultimo
        if "num_apolice" not in df.columns:
            return {"file": filename, "status": "sem_num_apolice_para_ultimo"}
        ordered = sort_for_latest(
            df,
            "data_emissao" if "data_emissao" in df.columns else None,
            "num_endosso" if "num_endosso" in df.columns else None,
        )
        result = ordered.drop_duplicates(subset=["num_apolice"], keep="first").copy()
        if "data_fim_vigencia" in result.columns:
            result = recompute_status_by_dates(result, "data_fim_vigencia", today_override=ref_date)

    # Reordenar colunas
    final_cols = [c for c in OUTPUT_COLUMNS_ORDER if c in result.columns]
    result = result[final_cols]

    out_dir.mkdir(parents=True, exist_ok=True)
    out_name = f"{path.stem}__{mode}_automatico.xlsx"
    out_path = out_dir / out_name

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Aba com o resultado padronizado
        result.to_excel(writer, index=False, sheet_name="UNIQUE")        
        # Aba com os dados originais lidos do arquivo de entrada
        try:
            raw.to_excel(writer, index=False, sheet_name="original")
        except Exception:
            # Em caso de dados muito grandes ou tipos não suportados, converte para string
            raw_copy = raw.astype(str)
            raw_copy.to_excel(writer, index=False, sheet_name="original")

    return {
        "file": filename,
        "status": "ok",
        "mode": mode,
        "linhas_entrada": len(raw),
        "linhas_saida": len(result),
        "colunas_detectadas": list(result.columns),
        "colunas_dropadas": colunas_dropadas,
        "saida": str(out_path)
    }


def main():
    parser = argparse.ArgumentParser(description="Processa arquivos de seguradoras (agrupar ou último valor por apólice).")
    parser.add_argument("-i", "--input", required=True, help="Pasta de entrada com .xlsx/.xls/.csv")
    parser.add_argument("-o", "--output", required=True, help="Pasta de saída para salvar os UNIQUE")
    parser.add_argument("-c", "--config", default="config.json", help="Caminho do arquivo de configuração (JSON)")
    parser.add_argument("--data", dest="ref_date", default=None, help="Data de referência para cálculo do status (ex.: 30/09/2025)")
    parser.add_argument("--log-dir", dest="log_dir", default=None, help="Diretório para logs históricos (default: <saida>/_historico)")
    args = parser.parse_args()

    in_dir = Path(args.input)
    out_dir = Path(args.output)
    cfg_path = Path(args.config)

    if not in_dir.exists() or not in_dir.is_dir():
        print(f"ERRO: pasta de entrada inválida: {in_dir}")
        sys.exit(1)

    if not cfg_path.exists():
        print(f"ERRO: arquivo de configuração não encontrado: {cfg_path}")
        sys.exit(1)

    with open(cfg_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    ref_dt = parse_reference_date(args.ref_date) if args.ref_date else None

    run_id = str(uuid.uuid4())
    log_dir = Path(args.log_dir) if args.log_dir else (out_dir / "_historico")
    log_dir.mkdir(parents=True, exist_ok=True)

    files = [p for p in in_dir.iterdir() if p.suffix.lower() in [".xlsx", ".xls", ".csv"]]
    if not files:
        print("Nenhum arquivo .xlsx/.xls/.csv encontrado na pasta de entrada.")
        sys.exit(0)

    resumos = []
    per_file_logs = []
    for p in files:
        resumo = process_file(p, out_dir, cfg, ref_date=ref_dt)
        resumos.append(resumo)
        if resumo.get("status") == "ok":
            print(f"[OK] {p.name} -> {resumo.get('mode')} | linhas: {resumo.get('linhas_entrada')} -> {resumo.get('linhas_saida')}")
        else:
            print(f"[ERRO] {p.name} -> {resumo.get('status')} | {resumo.get('detalhe', '')}")
        per_file_logs.append({
            "run_id": run_id,
            "arquivo": p.name,
            "modo": resumo.get("mode"),
            "status": resumo.get("status"),
            "linhas_entrada": resumo.get("linhas_entrada"),
            "linhas_saida": resumo.get("linhas_saida"),
            "saida": resumo.get("saida"),
            "colunas_detectadas": ";".join(resumo.get("colunas_detectadas", [])) if resumo.get("colunas_detectadas") else None,
            "erro_detalhe": resumo.get("detalhe")
        })

    resumo_df = pd.DataFrame(resumos)
    effective_ref = (ref_dt if ref_dt is not None else pd.Timestamp(datetime.now().date()))
    resumo_df['data_referencia_base'] = pd.Timestamp(effective_ref.date())
    resumo_df['run_id'] = run_id

    exec_log = pd.DataFrame([{
        'run_id': run_id,
        'data_execucao': pd.Timestamp(datetime.now()),
        'data_referencia_status': pd.Timestamp(effective_ref.date()),
        'input_dir': str(in_dir.resolve()),
        'output_dir': str(out_dir.resolve()),
        'total_arquivos_encontrados': len(files),
        'total_processados_ok': int(sum(1 for r in resumos if r.get('status') == 'ok')),
        'total_erros': int(sum(1 for r in resumos if r.get('status') != 'ok'))
    }])

    resumo_path = out_dir / "_resumo_processamento.xlsx"
    with pd.ExcelWriter(resumo_path, engine="openpyxl") as writer:
        resumo_df.to_excel(writer, index=False, sheet_name="resumo")
        exec_log.to_excel(writer, index=False, sheet_name="log_execucao")

    # CSV históricos (append)
    exec_csv = log_dir / 'execucoes.csv'
    files_csv = log_dir / 'arquivos.csv'

    if exec_csv.exists():
        exec_log.to_csv(exec_csv, mode='a', index=False, header=False)
    else:
        exec_log.to_csv(exec_csv, mode='w', index=False, header=True)

    per_file_df = pd.DataFrame(per_file_logs)
    if files_csv.exists():
        per_file_df.to_csv(files_csv, mode='a', index=False, header=False)
    else:
        per_file_df.to_csv(files_csv, mode='w', index=False, header=True)

    print(f"\nResumo salvo em: {resumo_path}\nLog histórico (execuções): {exec_csv}\nLog histórico (arquivos): {files_csv}")


if __name__ == "__main__":
    main()
