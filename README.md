
# Pipeline de Processamento de Circularização Apólices (Venâncio, Carlos)

Este pacote padroniza colunas, aplica a regra de **agrupar** (somar IS por apólice) ou **último** (manter último valor emitido por apólice) **por arquivo**, e salva resultados individuais + um **_resumo**.

## 1) Requisitos
- Python 3.9+
- Pacotes: `pandas`, `openpyxl`, `xlrd`

Instale (uma vez):
```bash
pip install pandas openpyxl xlrd
```

## 2) Estrutura
1. **Local (No seu notebook)**
```
./
  gui_processa_seguradoras.py # Interface Gráfica do Usuário (GUI)
  processa_seguradoras.py
  config.json
  entrada/   # coloque aqui os .xlsx/.xls/.csv da seguradora
  saida/     # resultados serão salvos aqui
```
2. **Rede**
```
./
  gui_processa_seguradoras.py # Interface Gráfica do Usuário (GUI)
  processa_seguradoras.py
  config.json 
```

## 3) Como usar
```bash
python processa_seguradoras.py -i entrada -o saida -c config.json
```

## 4) Decisão da regra por arquivo
A ordem de decisão é:
1. **Sufixos no nome** (se `rules.suffix_overrides=true`):
   - "agrupar": match em qualquer keyword de `rules.suffix_keywords.agrupar` (ex.: `_agrupar`, `consolidado`, `soma`)
   - "ultimo":  match em qualquer keyword de `rules.suffix_keywords.ultimo` (ex.: `_ultimo`, `mais_recente`, `final`)
2. **Padrões por seguradora** (`rules.insurer_patterns`), ex.: `JUNTO` → agrupar; `CESCE` → último.
3. **Default** (`rules.default_mode`), hoje = `ultimo`.

> Se quiser fixar sempre por seguradora, basta **remover** `_agrupar` / `_ultimo` dos nomes e manter os padrões em `insurer_patterns`.

## 5) Colunas detectadas
São mapeadas por sinônimos (ver `column_synonyms` no `config.json`). Exemplos:
- `cd_apolice`, `Documento`, `Apólice mãe` → `num_apolice`
- `Vl Is Tomada`, `vl_is`, `IS` → `is`
- `Dt Emissao` → `data_emissao`; `Dt Inicio Vigencia` → `data_inicio_vigencia`; `Dt Final Vigencia` → `data_fim_vigencia`

## 6) Regras por modo
### AGRUPAR
- **Chaves**: `num_apolice` (+ `apolice_susep` se existir)
- **IS**: soma
- **Datas**: emissão = máx; início = mín; fim = máx
- **Status**: recalculado por `data_fim_vigencia` (VIGENTE se hoje ≤ fim)

### ÚLTIMO
- Ordena por `data_emissao` (desc) e `num_endosso` (desc), e mantém **1ª ocorrência** de cada `num_apolice`.
- **Status**: idem (se não houver coluna de status).

## 7) Saídas
- Um arquivo por entrada: `NOME_ARQUIVO__{agrupar|ultimo}_automatico.xlsx` (aba `UNIQUE` e aba `original`).
  ##### UNIQUE: as informações que foram tratas;
  ##### original: as informações que foram tratas
- Um `_resumo_processamento.xlsx` com status de cada arquivo.

## 8) Dúvidas comuns
**Não reconheceu nenhuma coluna**: adicione o nome que apareceu no cabeçalho da seguradora ao array correspondente em `column_synonyms`.

**CSV com separador estranho ou acentuação**: o script tenta `utf-8-sig` e depois `latin-1` e autodetecta separador.

**Data americana (MM/DD/YYYY)**: o conversor aceita formatos diversos e usa `dayfirst=True` por padrão; se necessário, padronize a coluna antes de rodar.

## Interface gráfica (Tkinter)
Se preferir, rode via GUI para escolher as pastas de entrada/saída e parâmetros:
```bash
python gui_processa_seguradoras.py
```
Na janela, informe:
- **Pasta de entrada** (obrigatório)
- **Pasta de saída** (obrigatório)
- **config.json** (opcional, se não estiver no mesmo diretório)
- **Data de referência** (opcional)
- **Diretório de logs** (opcional)

O botão **Executar** iniciará o pipeline e mostrará os logs na própria tela.

###### **Observação**: GUI front-end" refere-se à Interface Gráfica do Usuário (GUI), que é a parte visual de um aplicativo ou site com a qual o usuário interage diretamente.