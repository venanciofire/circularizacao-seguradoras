
"""
GUI (Tkinter) para o pipeline de apólices — Venâncio, Carlos

Como executar:
    python gui_processa_seguradoras.py

Pré-requisitos:
    - processa_seguradoras.py no mesmo diretório
    - Python 3.9+ com tkinter instalado (vem por padrão na maioria das distros/instalações)
"""

import os
import sys
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

SCRIPT_NAME = 'processa_seguradoras.py'

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Circularização de Apólices - GUI')
        self.geometry('720x520')
        self.minsize(680, 480)

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.config_path = tk.StringVar(value=str(Path('config.json').resolve()) if Path('config.json').exists() else '')
        self.ref_date = tk.StringVar()
        self.log_dir = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        pad = {'padx': 8, 'pady': 6}
        frm = ttk.Frame(self)
        frm.pack(fill='both', expand=True)

        # Entrada
        row = 0
        ttk.Label(frm, text='Pasta de entrada (obrigatório):').grid(row=row, column=0, sticky='w', padx= 8, pady= 6)
        ent_in = ttk.Entry(frm, textvariable=self.input_dir, width=60)
        ent_in.grid(row=row, column=1, sticky='we', padx= 8, pady= 6)
        ttk.Button(frm, text='Procurar…', command=self._ask_input).grid(row=row, column=2, padx= 8, pady= 6)

        # Saída
        row += 1
        ttk.Label(frm, text='Pasta de saída (obrigatório):').grid(row=row, column=0, sticky='w', padx= 8, pady= 6)
        ent_out = ttk.Entry(frm, textvariable=self.output_dir, width=60)
        ent_out.grid(row=row, column=1, sticky='we', padx= 8, pady= 6)
        ttk.Button(frm, text='Procurar…', command=self._ask_output).grid(row=row, column=2, padx= 8, pady= 6)

        # Config
        row += 1
        ttk.Label(frm, text='Arquivo de configuração (config.json):').grid(row=row, column=0, sticky='w', padx= 8, pady= 6)
        ent_cfg = ttk.Entry(frm, textvariable=self.config_path, width=60)
        ent_cfg.grid(row=row, column=1, sticky='we', padx= 8, pady= 6)
        ttk.Button(frm, text='Localizar…', command=self._ask_config).grid(row=row, column=2, padx= 8, pady= 6)

        # Data de referência
        row += 1
        ttk.Label(frm, text='Data de referência (opcional):').grid(row=row, column=0, sticky='w', padx= 8, pady= 6)
        ttk.Entry(frm, textvariable=self.ref_date, width=20).grid(row=row, column=1, sticky='w', padx= 8, pady= 6)
        ttk.Label(frm, text='Ex.: 30/09/2025').grid(row=row, column=2, sticky='w', padx= 8, pady= 6)

        # Log dir
        row += 1
        ttk.Label(frm, text='Diretório de logs históricos (opcional):').grid(row=row, column=0, sticky='w', padx= 8, pady= 6)
        ent_log = ttk.Entry(frm, textvariable=self.log_dir, width=60)
        ent_log.grid(row=row, column=1, sticky='we', padx= 8, pady= 6)
        ttk.Button(frm, text='Procurar…', command=self._ask_logdir).grid(row=row, column=2, padx= 8, pady= 6)

        # Run controls
        row += 1
         # Botões: Fechar, Executar, Abrir pasta de saída
        self.btn_run = ttk.Button(frm, text='Executar', command=self._run_pipeline)
        self.btn_run.grid(row=row, column=0, sticky='w', padx= 8, pady= 6)

        self.btn_close = ttk.Button(frm, text='Fechar', command=self.destroy)
        self.btn_close.grid(row=row, column=1, sticky='w', padx= 8, pady= 6)

        # Barra de progresso
        self.prog = ttk.Progressbar(frm, mode='indeterminate')
        self.prog.grid(row=row, column=2, sticky='we', padx= 8, pady= 6)

        # Output log
        row += 1
        ttk.Label(frm, text='Saída do processo:').grid(row=row, column=0, sticky='w', padx= 8, pady= 6)
        row += 1
        self.txt = tk.Text(frm, height=18, wrap='word')
        self.txt.grid(row=row, column=0, columnspan=3, sticky='nsew', padx=8, pady=(0,8))
        yscroll = ttk.Scrollbar(frm, command=self.txt.yview)
        self.txt.configure(yscrollcommand=yscroll.set)
        # grid weights
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(row, weight=1)

    def _ask_input(self):
        d = filedialog.askdirectory(title='Selecione a pasta de entrada')
        if d:
            self.input_dir.set(d)

    def _ask_output(self):
        d = filedialog.askdirectory(title='Selecione a pasta de saída')
        if d:
            self.output_dir.set(d)

    def _ask_config(self):
        f = filedialog.askopenfilename(title='Selecione o config.json', filetypes=[('JSON', '*.json'), ('Todos', '*.*')])
        if f:
            self.config_path.set(f)

    def _ask_logdir(self):
        d = filedialog.askdirectory(title='Selecione a pasta de logs históricos')
        if d:
            self.log_dir.set(d)

    def _append_log(self, text):
        self.txt.insert('end', text)
        self.txt.see('end')
    def _validate(self):
        in_dir = self.input_dir.get().strip()
        out_dir = self.output_dir.get().strip()
        cfg = self.config_path.get().strip()
        if not in_dir:
            messagebox.showerror('Validação', 'Informe a pasta de entrada.')
            return False
        if not Path(in_dir).exists():
            messagebox.showerror('Validação', 'A pasta de entrada não existe.')
            return False
        if not out_dir:
            messagebox.showerror('Validação', 'Informe a pasta de saída.')
            return False
        if cfg and not Path(cfg).exists():
            messagebox.showerror('Validação', 'Arquivo de configuração não encontrado.')
            return False
        if not Path(SCRIPT_NAME).exists():
            messagebox.showerror('Validação', f'Não encontrei "{SCRIPT_NAME}" no diretório atual.')
            return False
        return True
        
    def _run_pipeline(self):
        if not self._validate():
            return
        self.btn_run.configure(state='disabled')
        self.prog.start(10)
        self.txt.delete('1.0', 'end')

        # Build command
        cmd = [sys.executable, str(Path(SCRIPT_NAME).resolve()),
               '-i', self.input_dir.get().strip(),
               '-o', self.output_dir.get().strip()]
        if self.config_path.get().strip():
            cmd += ['-c', self.config_path.get().strip()]
        if self.ref_date.get().strip():
            cmd += ['--data', self.ref_date.get().strip()]
        if self.log_dir.get().strip():
            cmd += ['--log-dir', self.log_dir.get().strip()]

        def worker():
            try:                
                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
                stdout = proc.stdout
                if stdout is None:
                    # Fallback: captura tudo ao final se stdout não estiver disponível
                    combined, _ = proc.communicate()
                    if combined:
                        self._append_log(combined)
                else:
                    for line in stdout:
                        self._append_log(line)
                ret = proc.wait()
                if ret == 0:
                    self._append_log('Concluído com sucesso.')
                else:
                    self._append_log(f"Finalizado com código {ret}. Verifique acima.")
            except Exception as e:
                messagebox.showerror('Erro na execução', str(e))

            finally:
                self.prog.stop()
                self.btn_run.configure(state='normal')               

        threading.Thread(target=worker, daemon=True).start()

if __name__ == '__main__':
    app = App()
    app.mainloop()