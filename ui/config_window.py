import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from config import AppConfig, save_config

try:
    from core.updater import (
        local_version, check_for_update, REPO,
        prepare_update, launch_updater, is_installed,
    )
except Exception:
    # fallbacks caso o modulo nao esteja disponivel
    local_version = lambda: "?"
    check_for_update = lambda repo=None: None
    REPO = ""
    prepare_update = lambda u, p=None: ""
    launch_updater = lambda b: None
    is_installed = lambda: False


class ConfigWindow(tk.Toplevel):
    def __init__(self, parent, cfg: AppConfig):
        super().__init__(parent)
        self.title("Configurações — DataHub v2")
        self.resizable(False, False)
        self.grab_set()
        self._cfg = cfg
        self._vars: dict[str, tk.StringVar | tk.BooleanVar] = {}
        self._build()
        self._carregar(cfg)
        self._centralizar(parent)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        self._aba_geonline(nb)
        self._aba_geral(nb)
        self._aba_operview(nb)
        self._aba_atualizacao(nb)

        frame_btn = ttk.Frame(self)
        frame_btn.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(frame_btn, text="Salvar", command=self._salvar).pack(side="right", padx=5)
        ttk.Button(frame_btn, text="Cancelar", command=self.destroy).pack(side="right")

    def _aba_geonline(self, nb):
        frm = ttk.Frame(nb, padding=15)
        nb.add(frm, text="GeoOnline")
        campos = [
            ("URL", "geo_url"),
            ("Conta", "geo_conta"),
            ("Login", "geo_login"),
            ("Senha", "geo_senha"),
        ]
        for i, (label, key) in enumerate(campos):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=4)
            var = tk.StringVar()
            show = "*" if key == "geo_senha" else ""
            ttk.Entry(frm, textvariable=var, width=45, show=show).grid(row=i, column=1, padx=(10, 0), pady=4)
            self._vars[key] = var

    def _aba_geral(self, nb):
        frm = ttk.Frame(nb, padding=15)
        nb.add(frm, text="Geral")

        # Pasta local
        ttk.Label(frm, text="Pasta local").grid(row=0, column=0, sticky="w", pady=4)
        v_pasta = tk.StringVar()
        self._vars["pasta_local"] = v_pasta
        ttk.Entry(frm, textvariable=v_pasta, width=38).grid(row=0, column=1, padx=(10, 0))
        ttk.Button(frm, text="...", width=3,
                   command=lambda: v_pasta.set(filedialog.askdirectory() or v_pasta.get())
                   ).grid(row=0, column=2, padx=4)

        # SharePoint toggle
        v_sp = tk.BooleanVar()
        self._vars["sp_enabled"] = v_sp
        ttk.Checkbutton(frm, text="Enviar para SharePoint", variable=v_sp,
                        command=self._toggle_sp).grid(row=1, column=0, columnspan=2, sticky="w", pady=(12, 4))

        ttk.Label(frm, text="Pasta SharePoint").grid(row=2, column=0, sticky="w", pady=4)
        v_sp_pasta = tk.StringVar()
        self._vars["sp_pasta"] = v_sp_pasta
        self._entry_sp = ttk.Entry(frm, textvariable=v_sp_pasta, width=38)
        self._entry_sp.grid(row=2, column=1, padx=(10, 0))
        self._btn_sp = ttk.Button(frm, text="...", width=3,
                                  command=lambda: v_sp_pasta.set(filedialog.askdirectory() or v_sp_pasta.get()))
        self._btn_sp.grid(row=2, column=2, padx=4)

        # n8n webhook
        ttk.Label(frm, text="Webhook n8n").grid(row=3, column=0, sticky="w", pady=(12, 4))
        v_n8n = tk.StringVar()
        self._vars["webhook_n8n"] = v_n8n
        ttk.Entry(frm, textvariable=v_n8n, width=45).grid(row=3, column=1, padx=(10, 0), pady=(12, 4))

        self._toggle_sp()

    # ── Operview (origem + gravação do login) ──────────────────────────────────

    def _aba_operview(self, nb):
        frm = ttk.Frame(nb, padding=15)
        nb.add(frm, text="Operview")

        campos = [
            ("URL", "op_url"),
            ("Login", "op_login"),
            ("Senha", "op_senha"),
        ]
        for i, (label, key) in enumerate(campos):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=4)
            var = tk.StringVar()
            show = "*" if key == "op_senha" else ""
            ttk.Entry(frm, textvariable=var, width=45, show=show).grid(
                row=i, column=1, padx=(10, 0), pady=4)
            self._vars[key] = var

        ttk.Button(
            frm, text="🎬 Gravar login (clique na conta e na senha)",
            command=self._gravar_login,
        ).grid(row=3, column=0, columnspan=2, pady=12, sticky="w")
        ttk.Label(
            frm,
            text=("Abre o Operview, clique na conta salva e no campo de senha,\n"
                  "depois Confirme na janela que abrir. Salva login_posicoes.txt no Desktop."),
            foreground="#5b6b7b",
        ).grid(row=4, column=0, columnspan=2, sticky="w")

    def _gravar_login(self):
        from core.downloader.operview import gravar_login_operview
        # reflete o que está na tela no cfg antes de gravar
        self._cfg.operview.url = self._vars["op_url"].get().strip()
        self._cfg.operview.login = self._vars["op_login"].get().strip()
        self._cfg.operview.senha = self._vars["op_senha"].get().strip()
        try:
            gravar_login_operview(parent=self, cfg=self._cfg.operview)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Erro na gravação", str(e), parent=self)

    # ── Atualização (GitHub) ─────────────────────────────────────────────────

    def _aba_atualizacao(self, nb):
        frm = ttk.Frame(nb, padding=15)
        nb.add(frm, text="Atualização")

        ttk.Label(frm, text=f"Versão atual: {local_version()}").grid(
            row=0, column=0, sticky="w", pady=4
        )
        self._update_status = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self._update_status, foreground="#1a7f37").grid(
            row=1, column=0, sticky="w", pady=4
        )
        ttk.Button(
            frm, text="Verificar atualização", command=self._verificar_update
        ).grid(row=2, column=0, pady=10)
        if REPO:
            ttk.Label(
                frm, text=f"Repositório: {REPO}", foreground="#5b6b7b"
            ).grid(row=3, column=0, sticky="w", pady=2)

    def _verificar_update(self):
        if not is_installed():
            self._update_status.set(
                "Executando a partir do código-fonte. Atualize via git/pull do repositório."
            )
            return
        self._update_status.set("Verificando...")
        self.update_idletasks()
        try:
            res = check_for_update(REPO)
        except Exception as e:  # noqa: BLE001
            self._update_status.set(f"Erro ao verificar: {e}")
            return
        if not res:
            self._update_status.set(
                "DataHub está atualizado (ou não foi possível verificar a conexão)."
            )
            return
        tag, url = res
        if not messagebox.askyesno(
            "Atualização disponível",
            f"Nova versão {tag} disponível.\n"
            "Deseja baixar e atualizar agora?\n"
            "O aplicativo será reiniciado.",
        ):
            self._update_status.set("Atualização adiada.")
            return
        self._update_status.set(f"Baixando {tag}...")
        self.update_idletasks()
        try:
            bat = prepare_update(url, progress=None)
        except Exception as e:  # noqa: BLE001
            self._update_status.set(f"Falha no download: {e}")
            return
        launch_updater(bat)  # encerra este processo e dispara o update

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _toggle_sp(self):
        estado = "normal" if self._vars["sp_enabled"].get() else "disabled"
        self._entry_sp.config(state=estado)
        self._btn_sp.config(state=estado)

    def _carregar(self, cfg: AppConfig):
        self._vars["geo_url"].set(cfg.geonline.url)
        self._vars["geo_conta"].set(cfg.geonline.conta)
        self._vars["geo_login"].set(cfg.geonline.login)
        self._vars["geo_senha"].set(cfg.geonline.senha)
        self._vars["op_url"].set(cfg.operview.url)
        self._vars["op_login"].set(cfg.operview.login)
        self._vars["op_senha"].set(cfg.operview.senha)
        self._vars["pasta_local"].set(cfg.pasta_local)
        self._vars["sp_enabled"].set(cfg.sharepoint.enabled)
        self._vars["sp_pasta"].set(cfg.sharepoint.pasta)
        self._vars["webhook_n8n"].set(cfg.webhook_n8n)
        self._toggle_sp()

    def _salvar(self):
        cfg = self._cfg
        cfg.geonline.url   = self._vars["geo_url"].get().strip()
        cfg.geonline.conta = self._vars["geo_conta"].get().strip()
        cfg.geonline.login = self._vars["geo_login"].get().strip()
        cfg.geonline.senha = self._vars["geo_senha"].get().strip()
        cfg.operview.url   = self._vars["op_url"].get().strip()
        cfg.operview.login = self._vars["op_login"].get().strip()
        cfg.operview.senha = self._vars["op_senha"].get().strip()
        cfg.pasta_local    = self._vars["pasta_local"].get().strip()
        cfg.sharepoint.enabled = self._vars["sp_enabled"].get()
        cfg.sharepoint.pasta   = self._vars["sp_pasta"].get().strip()
        cfg.webhook_n8n    = self._vars["webhook_n8n"].get().strip()

        if not cfg.pasta_local:
            messagebox.showwarning("Atenção", "Informe a pasta local.", parent=self)
            return

        save_config(cfg)
        messagebox.showinfo("Salvo", "Configurações salvas com sucesso!", parent=self)
        self.destroy()

    def _centralizar(self, parent):
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_x(), parent.winfo_y()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")
