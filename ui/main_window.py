import calendar as _cal
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Callable

from config import AppConfig, save_config

POLOS = [
    "Todos os polos",
    "UTN",
    "UTS",
    "Campos",
    "Lagos",
    "Macaé",
    "Noroeste",
    "Magé",
    "Niterói",
    "São Gonçalo",
    "Serrana",
    "Sul",
]

_MESES_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]


# ── Calendar popup ────────────────────────────────────────────────────────────

class _CalPopup(tk.Toplevel):
    def __init__(self, anchor: tk.Widget, data_atual: date, callback):
        super().__init__(anchor.winfo_toplevel())
        self.title("")
        self.resizable(False, False)
        self._cb = callback
        self._ano = data_atual.year
        self._mes = data_atual.month
        self._sel = data_atual

        self._build()
        self._render()
        self._posicionar(anchor)
        self.lift()
        self.grab_set()
        self.bind("<Escape>", lambda e: self.destroy())

    def _posicionar(self, anchor: tk.Widget):
        self.update_idletasks()
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height() + 2
        sw = anchor.winfo_screenwidth()
        w = self.winfo_reqwidth()
        if x + w > sw - 10:
            x = sw - w - 10
        self.geometry(f"+{x}+{y}")

    def _build(self):
        frm = tk.Frame(self, bg="#f0f0f0", padx=8, pady=6)
        frm.pack(fill="both", expand=True)

        nav = tk.Frame(frm, bg="#f0f0f0")
        nav.pack(fill="x", pady=(0, 4))
        tk.Button(nav, text="◀", bd=0, bg="#f0f0f0", relief="flat",
                  activebackground="#dde", font=("Segoe UI", 9), cursor="hand2",
                  command=self._prev).pack(side="left")
        self._lbl = tk.Label(nav, bg="#f0f0f0",
                             font=("Segoe UI", 9, "bold"), width=18, anchor="center")
        self._lbl.pack(side="left", expand=True)
        tk.Button(nav, text="▶", bd=0, bg="#f0f0f0", relief="flat",
                  activebackground="#dde", font=("Segoe UI", 9), cursor="hand2",
                  command=self._next).pack(side="right")

        hdr = tk.Frame(frm, bg="#f0f0f0")
        hdr.pack()
        for d in ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"]:
            tk.Label(hdr, text=d, font=("Segoe UI", 8, "bold"),
                     bg="#3465a4", fg="white", width=3, anchor="center",
                     padx=3, pady=3).pack(side="left", padx=1, pady=(0, 2))

        self._grid = tk.Frame(frm, bg="#f0f0f0")
        self._grid.pack()

    def _render(self):
        for w in self._grid.winfo_children():
            w.destroy()
        self._lbl.config(text=f"{_MESES_PT[self._mes - 1]} {self._ano}")

        for semana in _cal.monthcalendar(self._ano, self._mes):
            row = tk.Frame(self._grid, bg="#f0f0f0")
            row.pack()
            for dia in semana:
                if dia == 0:
                    tk.Label(row, width=3, bg="#f0f0f0").pack(side="left", padx=1, pady=1)
                else:
                    is_sel = (dia == self._sel.day and
                              self._mes == self._sel.month and
                              self._ano == self._sel.year)
                    bg = "#3465a4" if is_sel else "#f0f0f0"
                    fg = "white" if is_sel else "#333333"
                    tk.Button(
                        row, text=str(dia), width=3, bd=0, relief="flat",
                        bg=bg, fg=fg, activebackground="#b8ccee",
                        font=("Segoe UI", 9), cursor="hand2",
                        command=lambda d=dia: self._pick(d),
                    ).pack(side="left", padx=1, pady=1)

    def _pick(self, dia: int):
        self._cb(date(self._ano, self._mes, dia))
        self.destroy()

    def _prev(self):
        if self._mes == 1:
            self._mes, self._ano = 12, self._ano - 1
        else:
            self._mes -= 1
        self._render()

    def _next(self):
        if self._mes == 12:
            self._mes, self._ano = 1, self._ano + 1
        else:
            self._mes += 1
        self._render()


# ── Date entry widget ─────────────────────────────────────────────────────────

class _DateEntry(ttk.Frame):
    def __init__(self, parent, valor: datetime = None, **kwargs):
        super().__init__(parent, **kwargs)
        d = (valor or datetime.now()).date()
        self._date = d
        self._var = tk.StringVar(value=d.strftime("%d/%m/%Y"))

        self._entry = ttk.Entry(self, textvariable=self._var, width=11,
                                font=("Segoe UI", 10))
        self._entry.pack(side="left")
        self._entry.bind("<FocusOut>", self._parse_input)
        self._entry.bind("<Return>", self._parse_input)
        self._var.trace_add("write", self._on_type)
        ttk.Button(self, text="📅", width=3, command=self._open).pack(side="left", padx=(2, 0))
        self._popup = None

    def _on_type(self, *args):
        pass

    def _parse_input(self, event=None):
        try:
            val = self._var.get().strip()
            parts = val.split("/")
            if len(parts) == 3:
                day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
                new_date = date(year, month, day)
                self._date = new_date
                self._var.set(new_date.strftime("%d/%m/%Y"))
        except (ValueError, IndexError):
            self._var.set(self._date.strftime("%d/%m/%Y"))

    def _open(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
            self._popup = None
            return
        self._popup = _CalPopup(self._entry, self._date, self._set)

    def _set(self, d: date):
        self._date = d
        self._var.set(d.strftime("%d/%m/%Y"))
        self._popup = None

    def get_date(self) -> datetime:
        return datetime(self._date.year, self._date.month, self._date.day)


# ── Main window ───────────────────────────────────────────────────────────────

class MainWindow(tk.Tk):
    def __init__(self, cfg: AppConfig, on_executar: Callable,
                 on_parar: Callable, on_tratar: Callable):
        super().__init__()
        self.title("DataHub v2")
        self.resizable(False, False)
        self._cfg = cfg
        self._on_executar = on_executar
        self._on_parar = on_parar
        self._on_tratar = on_tratar
        self._build()
        self._centralizar()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build(self):
        self._nb = ttk.Notebook(self)
        self._nb.pack(fill="both", expand=True, padx=8, pady=8)
        self._build_execucao()
        self._build_tratamento()
        self._build_configuracoes()

    # ── Aba Execução ──────────────────────────────────────────────────────────

    def _build_execucao(self):
        frm = ttk.Frame(self._nb, padding=10)
        self._nb.add(frm, text="  Execução  ")

        # Período
        frm_per = ttk.LabelFrame(frm, text="Período", padding=10)
        frm_per.pack(fill="x", pady=(0, 8))

        hoje = datetime.now()
        ttk.Label(frm_per, text="De:").grid(row=0, column=0, sticky="w")
        self._data_ini = _DateEntry(frm_per, hoje - timedelta(days=6))
        self._data_ini.grid(row=0, column=1, padx=(6, 20))
        ttk.Label(frm_per, text="Até:").grid(row=0, column=2, sticky="w")
        self._data_fim = _DateEntry(frm_per, hoje)
        self._data_fim.grid(row=0, column=3, padx=6)

        # Polo
        frm_polo = ttk.LabelFrame(frm, text="Polo", padding=10)
        frm_polo.pack(fill="x", pady=(0, 8))
        self._polo_exec = tk.StringVar(value="Todos os polos")
        ttk.Combobox(frm_polo, textvariable=self._polo_exec, values=POLOS,
                     state="readonly", width=30).pack(anchor="w")

        # Modo de execução
        frm_exec = ttk.LabelFrame(frm, text="Execução", padding=10)
        frm_exec.pack(fill="x", pady=(0, 8))
        self._modo = tk.StringVar(value="uma_vez")
        for txt, val in [("Uma vez", "uma_vez"), ("A cada 30 min", "30min"), ("A cada 60 min", "60min")]:
            ttk.Radiobutton(frm_exec, text=txt, variable=self._modo,
                            value=val).pack(side="left", padx=8)

        # Botões
        frm_btn = ttk.Frame(frm)
        frm_btn.pack(fill="x", pady=(0, 8))
        self._btn_executar = ttk.Button(frm_btn, text="▶  Executar", command=self._executar)
        self._btn_executar.pack(side="left", padx=(0, 8))
        self._btn_parar = ttk.Button(frm_btn, text="⏹  Parar",
                                     command=self._parar, state="disabled")
        self._btn_parar.pack(side="left")

        # Log
        frm_log = ttk.LabelFrame(frm, text="Log", padding=6)
        frm_log.pack(fill="both", expand=True)
        self._log_exec_text = tk.Text(
            frm_log, height=12, width=68, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            wrap="word", relief="flat",
        )
        sc = ttk.Scrollbar(frm_log, command=self._log_exec_text.yview)
        self._log_exec_text.config(yscrollcommand=sc.set)
        self._log_exec_text.pack(side="left", fill="both", expand=True)
        sc.pack(side="right", fill="y")

    # ── Aba Configurações ─────────────────────────────────────────────────────

    def _build_configuracoes(self):
        frm = ttk.Frame(self._nb, padding=15)
        self._nb.add(frm, text="  Configurações  ")
        self._vars_cfg: dict[str, tk.Variable] = {}

        # GeoOnline
        geo = ttk.LabelFrame(frm, text="GeoOnline", padding=10)
        geo.pack(fill="x", pady=(0, 10))
        for i, (label, key, secret) in enumerate([
            ("URL",   "geo_url",   False),
            ("Conta", "geo_conta", False),
            ("Login", "geo_login", False),
            ("Senha", "geo_senha", True),
        ]):
            ttk.Label(geo, text=label).grid(row=i, column=0, sticky="w", pady=4)
            var = tk.StringVar()
            self._vars_cfg[key] = var
            ttk.Entry(geo, textvariable=var, width=45,
                      show="*" if secret else "").grid(row=i, column=1, padx=(10, 0), pady=4)

        # Geral
        gen = ttk.LabelFrame(frm, text="Geral", padding=10)
        gen.pack(fill="x", pady=(0, 10))

        ttk.Label(gen, text="Pasta local").grid(row=0, column=0, sticky="w", pady=4)
        v_pasta = tk.StringVar()
        self._vars_cfg["pasta_local"] = v_pasta
        ttk.Entry(gen, textvariable=v_pasta, width=38).grid(row=0, column=1, padx=(10, 0))
        ttk.Button(gen, text="...", width=3,
                   command=lambda: v_pasta.set(
                       filedialog.askdirectory() or v_pasta.get())
                   ).grid(row=0, column=2, padx=4)

        v_sp = tk.BooleanVar()
        self._vars_cfg["sp_enabled"] = v_sp
        ttk.Checkbutton(gen, text="Enviar para SharePoint", variable=v_sp,
                        command=self._toggle_sp).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(10, 4))

        ttk.Label(gen, text="Pasta SharePoint").grid(row=2, column=0, sticky="w", pady=4)
        v_sp_pasta = tk.StringVar()
        self._vars_cfg["sp_pasta"] = v_sp_pasta
        self._entry_sp = ttk.Entry(gen, textvariable=v_sp_pasta, width=38)
        self._entry_sp.grid(row=2, column=1, padx=(10, 0))
        self._btn_sp = ttk.Button(gen, text="...", width=3,
                                  command=lambda: v_sp_pasta.set(
                                      filedialog.askdirectory() or v_sp_pasta.get()))
        self._btn_sp.grid(row=2, column=2, padx=4)

        v_n8n = tk.BooleanVar()
        self._vars_cfg["n8n_enabled"] = v_n8n
        ttk.Checkbutton(gen, text="Enviar análises para n8n", variable=v_n8n,
                        command=self._toggle_n8n).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 4))

        ttk.Label(gen, text="Webhook n8n").grid(row=4, column=0, sticky="w", pady=4)
        v_n8n_url = tk.StringVar()
        self._vars_cfg["n8n_webhook"] = v_n8n_url
        self._entry_n8n = ttk.Entry(gen, textvariable=v_n8n_url, width=45)
        self._entry_n8n.grid(row=4, column=1, padx=(10, 0), pady=(4, 4))

        ttk.Label(gen, text="Atualizar base mensal").grid(row=5, column=0, sticky="w", pady=(10, 4))
        v_base = tk.BooleanVar(value=True)
        self._vars_cfg["base_mensal_enabled"] = v_base
        ttk.Checkbutton(gen, text="(mensal)", variable=v_base).grid(
            row=5, column=1, sticky="w", pady=(10, 4))

        ttk.Button(frm, text="💾  Salvar configurações",
                   command=self._salvar_config).pack(anchor="e", pady=8)

        self._carregar_config()
        self._toggle_sp()

    # ── Aba Tratamento ────────────────────────────────────────────────────────

    def _build_tratamento(self):
        frm = ttk.Frame(self._nb, padding=10)
        self._nb.add(frm, text="  Tratamento  ")

        # Frame de ferramentas
        frm_tools = ttk.LabelFrame(frm, text="Ferramentas", padding=10)
        frm_tools.pack(fill="x", pady=(0, 8))

        self._tool_var = tk.StringVar(value="tratar")
        for texto, valor in [
            ("Tratar Incidências", "tratar"),
            ("XLSX → CSV", "xlsx_csv"),
            ("CSV → XLSX", "csv_xlsx"),
            ("Merge Arquivos", "merge"),
        ]:
            ttk.Radiobutton(frm_tools, text=texto, variable=self._tool_var,
                            value=valor, command=self._toggle_tool_ui).pack(anchor="w", padx=10)

        # Painel: Tratar
        self._frm_tratar = ttk.Frame(frm)
        ttk.Label(self._frm_tratar, text="Arquivo de entrada").pack(anchor="w")
        frm_arq1 = ttk.Frame(self._frm_tratar)
        frm_arq1.pack(fill="x", pady=(0, 4))
        self._arquivo_var = tk.StringVar()
        ttk.Entry(frm_arq1, textvariable=self._arquivo_var, width=54).pack(side="left")
        ttk.Button(frm_arq1, text="...", width=3,
                   command=self._browse_arquivo).pack(side="left", padx=(6, 0))

        # Painel: Conversor
        self._frm_conversor = ttk.Frame(frm)
        ttk.Label(self._frm_conversor, text="Arquivo de entrada").pack(anchor="w")
        frm_arq2 = ttk.Frame(self._frm_conversor)
        frm_arq2.pack(fill="x", pady=(0, 4))
        self._conv_arquivo_var = tk.StringVar()
        ttk.Entry(frm_arq2, textvariable=self._conv_arquivo_var, width=54).pack(side="left")
        ttk.Button(frm_arq2, text="...", width=3,
                   command=self._browse_conv_arquivo).pack(side="left", padx=(6, 0))
        ttk.Label(self._frm_conversor, text="Separador (CSV)").pack(anchor="w")
        self._conv_sep_var = tk.StringVar(value=";")
        ttk.Entry(self._frm_conversor, textvariable=self._conv_sep_var, width=4).pack(anchor="w", pady=(0, 4))

        # Painel: Merge
        self._frm_merge = ttk.Frame(frm)
        ttk.Label(self._frm_merge, text="Arquivos de entrada (selecione múltiplos)").pack(anchor="w")
        frm_m1 = ttk.Frame(self._frm_merge)
        frm_m1.pack(fill="x", pady=(0, 4))
        self._merge_arquivos_var = tk.StringVar()
        ttk.Entry(frm_m1, textvariable=self._merge_arquivos_var, width=54).pack(side="left")
        ttk.Button(frm_m1, text="...", width=3,
                   command=self._browse_merge).pack(side="left", padx=(6, 0))
        ttk.Label(self._frm_merge, text="Arquivo de saída").pack(anchor="w")
        frm_m2 = ttk.Frame(self._frm_merge)
        frm_m2.pack(fill="x", pady=(0, 4))
        self._merge_saida_var = tk.StringVar()
        ttk.Entry(frm_m2, textvariable=self._merge_saida_var, width=54).pack(side="left")
        ttk.Button(frm_m2, text="...", width=3,
                   command=self._browse_merge_saida).pack(side="left", padx=(6, 0))
        ttk.Label(self._frm_merge, text="Coluna chave (opcional, remove duplicatas)").pack(anchor="w")
        self._merge_chave_var = tk.StringVar()
        ttk.Entry(self._frm_merge, textvariable=self._merge_chave_var, width=30).pack(anchor="w", pady=(0, 4))

        # Botão
        frm_btn = ttk.Frame(frm)
        frm_btn.pack(fill="x", pady=(8, 8))
        self._btn_tratar = ttk.Button(frm_btn, text="▶  Executar", command=self._executar_tool)
        self._btn_tratar.pack(side="left")

        # Log
        frm_log = ttk.LabelFrame(frm, text="Log", padding=6)
        frm_log.pack(fill="both", expand=True)
        self._log_trat_text = tk.Text(
            frm_log, height=12, width=68, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            wrap="word", relief="flat",
        )
        sc2 = ttk.Scrollbar(frm_log, command=self._log_trat_text.yview)
        self._log_trat_text.config(yscrollcommand=sc2.set)
        self._log_trat_text.pack(side="left", fill="both", expand=True)
        sc2.pack(side="right", fill="y")

        self._toggle_tool_ui()

    # ── Callbacks Execução ────────────────────────────────────────────────────

    def _executar(self):
        data_ini = self._data_ini.get_date()
        data_fim = self._data_fim.get_date()
        if data_ini > data_fim:
            messagebox.showerror("Erro de data",
                                 "Data inicial não pode ser maior que a final.")
            return
        self._btn_executar.config(state="disabled")
        self._btn_parar.config(state="normal")
        self._on_executar(
            selecionados={"geonline": True},
            data_ini=data_ini,
            data_fim=data_fim,
            modo=self._modo.get(),
            polo=self._polo_exec.get(),
        )

    def _parar(self):
        self._on_parar()
        self._btn_executar.config(state="normal")
        self._btn_parar.config(state="disabled")

    # ── Callbacks Configurações ───────────────────────────────────────────────

    def _toggle_sp(self):
        estado = "normal" if self._vars_cfg["sp_enabled"].get() else "disabled"
        self._entry_sp.config(state=estado)
        self._btn_sp.config(state=estado)

    def _toggle_n8n(self):
        estado = "normal" if self._vars_cfg["n8n_enabled"].get() else "disabled"
        self._entry_n8n.config(state=estado)

    def _carregar_config(self):
        cfg = self._cfg
        self._vars_cfg["geo_url"].set(cfg.geonline.url)
        self._vars_cfg["geo_conta"].set(cfg.geonline.conta)
        self._vars_cfg["geo_login"].set(cfg.geonline.login)
        self._vars_cfg["geo_senha"].set(cfg.geonline.senha)
        self._vars_cfg["pasta_local"].set(cfg.pasta_local)
        self._vars_cfg["sp_enabled"].set(cfg.sharepoint.enabled)
        self._vars_cfg["sp_pasta"].set(cfg.sharepoint.pasta)
        self._vars_cfg["n8n_enabled"].set(cfg.n8n.enabled)
        self._vars_cfg["n8n_webhook"].set(cfg.n8n.webhook)
        self._vars_cfg["base_mensal_enabled"].set(cfg.base_mensal_enabled)
        self._toggle_sp()
        self._toggle_n8n()

    def _salvar_config(self):
        cfg = self._cfg
        cfg.geonline.url   = self._vars_cfg["geo_url"].get().strip()
        cfg.geonline.conta = self._vars_cfg["geo_conta"].get().strip()
        cfg.geonline.login = self._vars_cfg["geo_login"].get().strip()
        cfg.geonline.senha = self._vars_cfg["geo_senha"].get().strip()
        cfg.pasta_local    = self._vars_cfg["pasta_local"].get().strip()
        cfg.sharepoint.enabled = self._vars_cfg["sp_enabled"].get()
        cfg.sharepoint.pasta   = self._vars_cfg["sp_pasta"].get().strip()
        cfg.n8n.enabled = self._vars_cfg["n8n_enabled"].get()
        cfg.n8n.webhook = self._vars_cfg["n8n_webhook"].get().strip()
        cfg.webhook_n8n = cfg.n8n.webhook
        cfg.base_mensal_enabled = self._vars_cfg["base_mensal_enabled"].get()

        if not cfg.pasta_local:
            messagebox.showwarning("Atenção", "Informe a pasta local.")
            return

        save_config(cfg)
        messagebox.showinfo("Salvo", "Configurações salvas com sucesso!")

    # ── Callbacks Tratamento ──────────────────────────────────────────────────

    def _toggle_tool_ui(self):
        tool = self._tool_var.get()
        for f in (self._frm_tratar, self._frm_conversor, self._frm_merge):
            f.pack_forget()
        labels = {"tratar": "▶  Tratar", "xlsx_csv": "▶  Converter", "csv_xlsx": "▶  Converter", "merge": "▶  Merge"}
        self._btn_tratar.config(text=labels.get(tool, "▶  Executar"))
        if tool == "tratar":
            self._frm_tratar.pack(fill="x", pady=(0, 4))
        elif tool in ("xlsx_csv", "csv_xlsx"):
            self._frm_conversor.pack(fill="x", pady=(0, 4))
        elif tool == "merge":
            self._frm_merge.pack(fill="x", pady=(0, 4))

    def _browse_arquivo(self):
        path = filedialog.askopenfilename(
            title="Selecionar arquivo",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")],
        )
        if path:
            self._arquivo_var.set(path)

    def _browse_conv_arquivo(self):
        tool = self._tool_var.get()
        # Corrige lógica: se é xlsx_csv, buscamos xlsx; se é csv_xlsx, buscamos csv
        if tool == "xlsx_csv":
            ext = "*.xlsx"
        else:
            ext = "*.csv"
        path = filedialog.askopenfilename(
            title="Selecionar arquivo",
            filetypes=[("Arquivos", ext), ("Todos", "*.*")],
        )
        if path:
            self._conv_arquivo_var.set(path)

    def _browse_merge(self):
        paths = filedialog.askopenfilenames(
            title="Selecionar arquivos",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")],
        )
        if paths:
            self._merge_arquivos_var.set("; ".join(paths))

    def _browse_merge_saida(self):
        path = filedialog.asksaveasfilename(
            title="Salvar como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
        )
        if path:
            self._merge_saida_var.set(path)

    def _executar_tool(self):
        tool = self._tool_var.get()
        self._btn_tratar.config(state="disabled")

        if tool == "tratar":
            arq = self._arquivo_var.get().strip()
            if not arq:
                messagebox.showwarning("Atenção", "Selecione um arquivo de entrada.")
                self._btn_tratar.config(state="normal")
                return
            p = Path(arq)
            if not p.exists():
                messagebox.showerror("Erro", f"Arquivo não encontrado:\n{arq}")
                self._btn_tratar.config(state="normal")
                return
            self._on_tratar(
                arquivo=p,
                polo="Todos os polos",
                log_fn=self.log_tratamento,
            )
        elif tool in ("xlsx_csv", "csv_xlsx"):
            import threading
            arq = self._conv_arquivo_var.get().strip()
            sep = self._conv_sep_var.get()
            if not arq:
                messagebox.showwarning("Atenção", "Selecione um arquivo de entrada.")
                self._btn_tratar.config(state="normal")
                return
            def _run():
                try:
                    from core.processors.utilidades import xlsx_para_csv, csv_para_xlsx
                    if tool == "xlsx_csv":
                        saida = xlsx_para_csv(arq, sep=sep, log_fn=self.log_tratamento)
                    else:
                        saida = csv_para_xlsx(arq, sep=sep, log_fn=self.log_tratamento)
                    self.log_tratamento(f"✅ Convertido: {saida}")
                except Exception as e:
                    self.log_tratamento(f"❌ Erro: {e}")
                finally:
                    self._btn_tratar.config(state="normal")
            threading.Thread(target=_run, daemon=True).start()
        elif tool == "merge":
            import threading
            arquivos_str = self._merge_arquivos_var.get().strip()
            saida = self._merge_saida_var.get().strip()
            chave = self._merge_chave_var.get().strip() or None
            if not arquivos_str or not saida:
                messagebox.showwarning("Atenção", "Informe arquivos de entrada e saída.")
                self._btn_tratar.config(state="normal")
                return
            arquivos = [a.strip() for a in arquivos_str.split(";") if a.strip()]
            if len(arquivos) < 2:
                messagebox.showwarning("Atenção", "Informe pelo menos 2 arquivos.")
                self._btn_tratar.config(state="normal")
                return
            def _run():
                try:
                    from core.processors.utilidades import merge_arquivos
                    saida_final = merge_arquivos(arquivos, saida, col_chave=chave, log_fn=self.log_tratamento)
                    self.log_tratamento(f"✅ Merge concluído: {saida_final}")
                except Exception as e:
                    self.log_tratamento(f"❌ Erro: {e}")
                finally:
                    self._btn_tratar.config(state="normal")
            threading.Thread(target=_run, daemon=True).start()

    # ── Log público ───────────────────────────────────────────────────────────

    def log(self, mensagem: str):
        """Thread-safe — log da aba Execução."""
        self.after(0, self._append_log, self._log_exec_text, mensagem)

    def log_tratamento(self, mensagem: str):
        """Thread-safe — log da aba Tratamento."""
        print(f"[DEBUG TRAT] {mensagem}")
        try:
            self.after(0, self._append_log, self._log_trat_text, mensagem)
        except Exception as e:
            print(f"Erro ao atualizar log: {e}")

    def _append_log(self, widget: tk.Text, mensagem: str):
        hora = datetime.now().strftime("%H:%M:%S")
        widget.config(state="normal")
        widget.insert("end", f"{hora}  {mensagem}\n")
        widget.see("end")
        widget.config(state="disabled")

    def set_rodando(self, rodando: bool):
        self.after(0, self._set_rodando, rodando)

    def _set_rodando(self, rodando: bool):
        self._btn_executar.config(state="disabled" if rodando else "normal")
        self._btn_parar.config(state="normal" if rodando else "disabled")

    def set_tratando(self, tratando: bool):
        self.after(0, lambda: self._btn_tratar.config(
            state="disabled" if tratando else "normal"))

    def mostrar_concluido(self, mensagem: str = "Fluxo concluído com sucesso!"):
        self.after(0, lambda: messagebox.showinfo(
            "DataHub v2", mensagem, parent=self))

    # ── Centralizar ───────────────────────────────────────────────────────────

    def _centralizar(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = self.winfo_width()
        h  = self.winfo_height()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")
