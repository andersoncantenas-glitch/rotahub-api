# -*- coding: utf-8 -*-
import json
import logging
import os
import re
import subprocess
import tempfile
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from app.db.connection import get_db
from app.services.api_client import _build_api_url, _call_api
from app.context import AppContext
from app.ui.components.tree_helpers import enable_treeview_sorting, tree_insert_aligned
from app.ui.components.page_base import PageBase
from app.utils.async_ui import run_async_ui
from app.utils.excel_helpers import upper
from app.utils.formatters import format_date_time, safe_float, safe_int


class HomePage(PageBase):
    def __init__(self, parent, app, context=None):
        self.context = context if isinstance(context, AppContext) else AppContext(
            config=None,
            hooks={},
        )
        cfg = self.context.config
        hooks = self.context.hooks or {}

        self.API_BASE_URL = str(getattr(cfg, "api_base_url", "") or "") if cfg is not None else ""
        self.API_SYNC_TIMEOUT = float(getattr(cfg, "api_sync_timeout", 8.0) or 8.0) if cfg is not None else 8.0
        self.APP_DIR = str(getattr(cfg, "app_dir", "") or "") if cfg is not None else ""
        self.APP_ENV = str(getattr(cfg, "app_env", "") or "") if cfg is not None else ""
        self.APP_VERSION = str(getattr(cfg, "app_version", "") or "") if cfg is not None else ""
        self.CHANGELOG_URL = str(getattr(cfg, "changelog_url", "") or "") if cfg is not None else ""
        self.COMPANY_ID = str(getattr(cfg, "company_id", "") or "") if cfg is not None else ""
        config_file = str(getattr(cfg, "config_file", "") or "") if cfg is not None else ""
        config_source = str(getattr(cfg, "config_source", "") or "") if cfg is not None else ""
        self.CONFIG_SOURCE = config_file if (config_file and os.path.exists(config_file)) else config_source
        self.DB_PATH = str(getattr(cfg, "db_path", "") or "") if cfg is not None else ""
        self.IS_FROZEN = bool(getattr(cfg, "is_frozen", False)) if cfg is not None else False
        self.RESOURCE_DIR = str(getattr(cfg, "resource_dir", "") or "") if cfg is not None else ""
        self.SETUP_DOWNLOAD_URL = str(getattr(cfg, "setup_download_url", "") or "") if cfg is not None else ""
        self.SUPPORT_EMAIL = str(getattr(cfg, "support_email", "") or "") if cfg is not None else ""
        self.SUPPORT_WHATSAPP = str(getattr(cfg, "support_whatsapp", "") or "") if cfg is not None else ""
        self.SYNC_MODE = str(getattr(cfg, "sync_mode", "") or "") if cfg is not None else ""
        self.TENANT_ID = str(getattr(cfg, "tenant_id", "") or "") if cfg is not None else ""
        self.UPDATE_MANIFEST_URL = str(getattr(cfg, "update_manifest_url", "") or "") if cfg is not None else ""
        self.USER_DATA_DIR = str(getattr(cfg, "data_root", "") or "") if cfg is not None else ""

        self.apply_window_icon = hooks.get("apply_window_icon", lambda _win: None)
        self.can_read_from_api = hooks.get("can_read_from_api", lambda: False)
        self.is_desktop_api_sync_enabled = hooks.get("is_desktop_api_sync_enabled", lambda: False)
        self.resolve_equipe_nomes = hooks.get("resolve_equipe_nomes", lambda equipe_raw: str(equipe_raw or ""))
        self.db_has_column = hooks.get("db_has_column", lambda *_args, **_kwargs: False)
        self.fetch_programacao_itens = hooks.get("fetch_programacao_itens", lambda *_args, **_kwargs: [])

        super().__init__(parent, app, "Home")
        self._api_job = None
        self._load_seq = 0
        self._update_notice_shown = False
        self._remote_version = "-"
        self._remote_setup_url = self.SETUP_DOWNLOAD_URL
        self._remote_changelog_url = self.CHANGELOG_URL
        self._alerts_text = ""
        self._runtime_diag_rows = {}
        self.body.grid_rowconfigure(0, weight=1)

        home_grid = ttk.Frame(self.body, style="Content.TFrame")
        home_grid.grid(row=0, column=0, sticky="nsew")
        home_grid.grid_columnconfigure(0, weight=3)
        home_grid.grid_columnconfigure(1, weight=1)
        home_grid.grid_rowconfigure(0, weight=0)
        home_grid.grid_rowconfigure(1, weight=1)
        home_grid.grid_rowconfigure(2, weight=0)

        card = ttk.Frame(home_grid, style="Card.TFrame", padding=14)
        card.grid(row=0, column=0, sticky="ew", padx=(0, 12), pady=(0, 12))
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=0)

        ttk.Label(card, text="Painel Operacional", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.lbl_clock = tk.Label(
            card,
            text="--/--/---- --:--:--",
            bg="white",
            fg="#111827",
            font=("Segoe UI", 10, "bold")
        )
        self.lbl_clock.grid(row=0, column=1, sticky="e")

        ttk.Label(
            card,
            text=(
                "Visao inicial da operacao do dia.\n"
                "Entre direto no fluxo principal sem passar por telas de apoio."
            ),
            style="CardLabel.TLabel",
            justify="left"
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        quick = ttk.Frame(card, style="CardInset.TFrame", padding=(12, 10))
        quick.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        for idx in range(3):
            quick.grid_columnconfigure(idx, weight=1, uniform="home_quick")
        ttk.Label(quick, text="Fluxo principal", style="InsetTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )
        ttk.Button(
            quick,
            text=self.app.get_routine_nav_label("Programacao"),
            style="Primary.TButton",
            command=lambda: self.app.show_page("Programacao"),
        ).grid(
            row=1, column=0, sticky="ew", padx=(0, 4), pady=(0, 6)
        )
        ttk.Button(
            quick,
            text=self.app.get_routine_nav_label("Recebimentos"),
            style="Ghost.TButton",
            command=lambda: self.app.show_page("Recebimentos"),
        ).grid(
            row=1, column=1, sticky="ew", padx=4, pady=(0, 6)
        )
        ttk.Button(
            quick,
            text=self.app.get_routine_nav_label("Despesas"),
            style="Ghost.TButton",
            command=lambda: self.app.show_page("Despesas"),
        ).grid(
            row=1, column=2, sticky="ew", padx=(4, 0), pady=(0, 6)
        )
        ttk.Button(
            quick,
            text=self.app.get_routine_nav_label("Relatorios"),
            style="Ghost.TButton",
            command=lambda: self.app.show_page("Relatorios"),
        ).grid(
            row=2, column=0, sticky="ew", padx=(0, 4)
        )
        ttk.Button(
            quick,
            text=self.app.get_routine_nav_label("Rotas"),
            style="Ghost.TButton",
            command=lambda: self.app.show_page("Rotas"),
        ).grid(
            row=2, column=1, sticky="ew", padx=4
        )
        ttk.Button(
            quick,
            text=self.app.get_routine_nav_label("Cadastros"),
            style="Ghost.TButton",
            command=lambda: self.app.show_page("Cadastros"),
        ).grid(
            row=2, column=2, sticky="ew", padx=(4, 0)
        )

        resumo = ttk.Frame(card, style="CardInset.TFrame", padding=(12, 10))
        resumo.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        for idx in range(3):
            resumo.grid_columnconfigure(idx, weight=1, uniform="home_summary")
        ttk.Label(resumo, text="Indicadores rapidos", style="InsetTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )
        self.lbl_total_prog = self._build_summary_tile(resumo, 0, "Programações Ativas")
        self.lbl_total_vendas = self._build_summary_tile(resumo, 1, "Vendas Importadas")
        self.lbl_total_clientes_ativos = self._build_summary_tile(resumo, 2, "Clientes Ativos")

        self._build_system_info_panel(home_grid)

        card_rotas = ttk.Frame(home_grid, style="Card.TFrame", padding=(12, 12, 12, 10))
        card_rotas.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        card_rotas.grid_rowconfigure(2, weight=1)
        card_rotas.grid_columnconfigure(0, weight=1)

        top_rotas = ttk.Frame(card_rotas, style="CardInset.TFrame")
        top_rotas.grid(row=0, column=0, columnspan=2, sticky="ew")
        top_rotas.grid_columnconfigure(0, weight=1)
        ttk.Label(top_rotas, text="Rotas Ativas", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(top_rotas, text="Atualizar", style="Ghost.TButton", command=self.on_show).grid(row=0, column=1, sticky="e")
        ttk.Label(
            card_rotas,
            text="Duplo clique em uma rota para abrir a pre-visualizacao.",
            style="CardLabel.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        cols = ["COD", "MOTORISTA", "VEICULO", "DATA"]
        self.tree_rotas = ttk.Treeview(card_rotas, columns=cols, show="headings", height=12)
        self.tree_rotas.grid(row=2, column=0, sticky="nsew", pady=(6, 0))

        vsb = ttk.Scrollbar(card_rotas, orient="vertical", command=self.tree_rotas.yview)
        self.tree_rotas.configure(yscrollcommand=vsb.set)
        vsb.grid(row=2, column=1, sticky="ns", pady=(6, 0))

        hsb = ttk.Scrollbar(card_rotas, orient="horizontal", command=self.tree_rotas.xview)
        self.tree_rotas.configure(xscrollcommand=hsb.set)
        hsb.grid(row=3, column=0, sticky="ew", pady=(8, 0))

        for c in cols:
            self.tree_rotas.heading(c, text=c)
            self.tree_rotas.column(c, width=150, minwidth=110, anchor="w", stretch=True)

        self.tree_rotas.column("COD", width=165, minwidth=150, stretch=False)
        self.tree_rotas.column("MOTORISTA", width=230, minwidth=180, stretch=True)
        self.tree_rotas.column("VEICULO", width=180, minwidth=150, stretch=True)
        self.tree_rotas.column("DATA", width=120, minwidth=110, anchor="center", stretch=False)

        self.tree_rotas.bind("<Double-1>", self._open_rota_preview)

        enable_treeview_sorting(
            self.tree_rotas,
            numeric_cols=set(),
            money_cols=set(),
            date_cols={"DATA"}
        )

        pend = ttk.Frame(home_grid, style="Card.TFrame", padding=10)
        pend.grid(row=2, column=0, sticky="ew", padx=(0, 12))
        pend.grid_columnconfigure((0, 1, 2), weight=1)
        ttk.Label(pend, text="Pendências do Dia", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.lbl_pend_rotas = self._build_pending_stat(pend, 0, "Rotas em aberto")
        self.lbl_pend_prest = self._build_pending_stat(pend, 1, "Prestações pendentes")
        self.lbl_pend_desp = self._build_pending_stat(pend, 2, "Sem despesa lançada")

        self.lbl_footer_sync = ttk.Label(
            self.footer_left,
            text="Ultima atualizacao: --/--/---- --:--:--",
            style="CardLabel.TLabel",
        )
        self.lbl_footer_sync.pack(side="left")
        self.lbl_footer_source = ttk.Label(
            self.footer_left,
            text="  |  Fonte: -",
            style="CardLabel.TLabel",
        )
        self.lbl_footer_source.pack(side="left", padx=(8, 0))
        ttk.Button(
            self.footer_right,
            text=f"Atualizar {self.app.get_routine_nav_label('Home')}",
            style="Ghost.TButton",
            command=self.on_show,
        ).pack(side="right")
        ttk.Button(
            self.footer_right,
            text=f"Abrir {self.app.get_routine_nav_label('Programacao')}",
            style="Primary.TButton",
            command=lambda: self.app.show_page("Programacao"),
        ).pack(side="right", padx=(0, 8))

        self._clock_job = None
        self._update_clock()
        self._build_api_status_badge()
        self._update_api_status()
        self._check_updates(silent=True)
        self.after(300, self._maybe_show_post_update_notifications)

    def _build_system_info_panel(self, parent):
        panel = ttk.Frame(parent, style="Card.TFrame", padding=(12, 12, 12, 10))
        panel.grid(row=0, column=1, rowspan=3, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)

        ttk.Label(panel, text="Base e Sistema", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        version_box = ttk.Frame(panel, style="CardInset.TFrame")
        version_box.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        version_box.grid_columnconfigure(0, weight=1)

        self.lbl_version_local = ttk.Label(
            version_box,
            text=f"Versão local: {self.APP_VERSION}",
            background="white",
            foreground="#111827",
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_version_local.grid(row=0, column=0, sticky="w")

        self.lbl_version_remote = ttk.Label(
            version_box,
            text="Versao disponivel: -",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9),
        )
        self.lbl_version_remote.grid(row=1, column=0, sticky="w", pady=(2, 0))

        diag = ttk.Frame(panel, style="CardInset.TFrame", padding=(10, 8))
        diag.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        diag.grid_columnconfigure(0, weight=1)
        ttk.Label(diag, text="Diagnóstico do Ambiente", style="InsetTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.lbl_runtime_env = self._build_info_row(diag, 1, "Ambiente")
        self.lbl_runtime_db = self._build_info_row(diag, 2, "Banco")
        self.lbl_runtime_persist = self._build_info_row(diag, 3, "Persistência")
        self.lbl_runtime_api = self._build_info_row(diag, 4, "API")
        self.lbl_runtime_cfg = self._build_info_row(diag, 5, "Config")
        self._refresh_runtime_diagnostics()

        actions = ttk.Frame(panel, style="CardInset.TFrame")
        actions.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        actions.grid_columnconfigure((0, 1), weight=1)

        ttk.Button(actions, text="Histórico", style="Ghost.TButton", command=self._open_changelog).grid(row=0, column=0, sticky="ew")
        ttk.Button(actions, text="Baixar Setup", style="Ghost.TButton", command=self._open_setup).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Button(actions, text="Atualizar versão", style="Primary.TButton", command=self._check_updates).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(actions, text="Publicar versao", style="Ghost.TButton", command=self._publish_version).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        support = ttk.Frame(panel, style="CardInset.TFrame")
        support.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(support, text="Suporte Técnico", style="InsetTitle.TLabel").grid(row=0, column=0, sticky="w")

        wpp_txt = f"WhatsApp: {self.SUPPORT_WHATSAPP}" if self.SUPPORT_WHATSAPP else "WhatsApp: não definido"
        mail_txt = self.SUPPORT_EMAIL if self.SUPPORT_EMAIL else "E-mail: não definido"
        ttk.Label(support, text=wpp_txt, background="white", foreground="#2563EB", font=("Segoe UI", 9, "underline")).grid(row=1, column=0, sticky="w")
        ttk.Label(support, text=mail_txt, background="white", foreground="#111827", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w")

        alerts = ttk.Frame(panel, style="CardInset.TFrame")
        alerts.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(alerts, text="Alertas", style="InsetTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.lbl_alerts = ttk.Label(
            alerts,
            text="Sem alertas.",
            background="white",
            foreground="#6B7280",
            justify="left",
            wraplength=255,
        )
        self.lbl_alerts.grid(row=1, column=0, sticky="w", pady=(2, 0))

    def _refresh_runtime_diagnostics(self):
        db_dir = os.path.dirname(self.DB_PATH) or "."
        self._set_runtime_diag_value("Ambiente", f"{self.APP_ENV} | tenant={self.TENANT_ID} | empresa={self.COMPANY_ID} | sync={self.SYNC_MODE}")
        self._set_runtime_diag_value("Banco", self.DB_PATH)
        self._set_runtime_diag_value("Persistência", db_dir)
        self._set_runtime_diag_value("API", self.API_BASE_URL or "-")
        self._set_runtime_diag_value("Config", self.CONFIG_SOURCE)

    def _open_url_safe(self, url: str, label: str):
        if not url:
            messagebox.showwarning("Aviso", f"URL de {label} não configurada.")
            return
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao abrir {label}:\n{e}")

    def _open_setup(self):
        url = self._remote_setup_url or self.SETUP_DOWNLOAD_URL
        self._open_url_safe(url, "download")

    def _can_auto_update_from_url(self, url: str) -> bool:
        try:
            u = str(url or "").strip().lower()
            if not u:
                return False
            if "releases/latest" in u or u.endswith("/releases"):
                return False
            parsed = urllib.parse.urlparse(u)
            path = (parsed.path or "").lower()
            return path.endswith(".exe")
        except Exception:
            return False

    def _download_setup_to_temp(self, url: str, version: str) -> str:
        u = str(url or "").strip()
        if not u:
            raise RuntimeError("URL de setup não configurada.")
        base_tmp = os.path.join(tempfile.gettempdir(), "RotaHubDesktop")
        os.makedirs(base_tmp, exist_ok=True)
        fname = f"RotaHubDesktop_Setup_{(version or 'latest').replace('/', '_')}.exe"
        dst = os.path.join(base_tmp, fname)
        req = urllib.request.Request(
            u,
            method="GET",
            headers={"User-Agent": "RotaHubDesktop-Updater/1.0"},
        )
        with urllib.request.urlopen(req, timeout=90) as resp:
            data = resp.read()
        if not data:
            raise RuntimeError("Download vazio do instalador.")
        with open(dst, "wb") as f:
            f.write(data)
        return dst

    def _run_setup_installer(self, setup_path: str):
        if not os.path.exists(setup_path):
            raise RuntimeError("Arquivo do instalador não encontrado.")
        if os.name == "nt":
            os.startfile(setup_path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen([setup_path], shell=False)

    def _mark_update_pending(self, target_version: str, setup_path: str):
        st = self._load_update_state()
        st["pending_update_version"] = str(target_version or "").strip()
        st["pending_setup_path"] = str(setup_path or "").strip()
        st["pending_update_started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._save_update_state(st)

    def _clear_update_pending_if_installed(self):
        st = self._load_update_state()
        pending_version = str(st.get("pending_update_version") or "").strip()
        if not pending_version:
            return
        if self._version_tuple(self.APP_VERSION) < self._version_tuple(pending_version):
            return
        st.pop("pending_update_version", None)
        st.pop("pending_setup_path", None)
        st.pop("pending_update_started_at", None)
        self._save_update_state(st)

    def _update_now(self):
        local_v = self._version_tuple(self.APP_VERSION)
        remote_v = self._version_tuple(self._remote_version)
        if remote_v <= local_v:
            messagebox.showinfo(
                "Atualização",
                f"Versão local ({self.APP_VERSION}) já é igual ou superior à remota ({self._remote_version or '-'}).\n"
                "Nenhum download será executado.",
            )
            return

        url = self._remote_setup_url or self.SETUP_DOWNLOAD_URL
        if not url:
            messagebox.showwarning(
                "Atualização",
                "Setup não configurado no manifesto.\nUse 'Baixar Setup' com um link válido para o .exe.",
            )
            return

        if not self._can_auto_update_from_url(url):
            messagebox.showinfo(
                "Atualização",
                "Nao foi possivel atualizar automaticamente porque o link atual nao aponta para um arquivo .exe.\n\n"
                "Use o botão 'Baixar Setup' para baixar manualmente.",
            )
            self._open_setup()
            return

        try:
            target_v = self._remote_version or "latest"
            self.set_status(f"STATUS: Baixando atualização {target_v}...")
            setup_path = self._download_setup_to_temp(url, target_v)
            self._mark_update_pending(target_v, setup_path)
            self._run_setup_installer(setup_path)
            messagebox.showinfo(
                "Atualização",
                "Instalador aberto com sucesso.\n\n"
                f"Versão atual em execução: {self.APP_VERSION}\n"
                f"Versão alvo da instalação: {target_v}\n\n"
                "Conclua a instalação, feche o Desktop e abra novamente.\n"
                "A versão local só muda na tela depois da reabertura.",
            )
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao atualizar automaticamente:\n{e}")
        finally:
            self.set_status("STATUS: Atualização finalizada (manual/automática).")

    def _open_changelog(self):
        url = self._remote_changelog_url or self.CHANGELOG_URL
        self._open_url_safe(url, "histórico")

    @staticmethod
    def _version_tuple(v: str):
        try:
            parts = [int(x) for x in re.findall(r"\d+", str(v or ""))]
            return tuple(parts) if parts else (0,)
        except Exception:
            return (0,)

    @staticmethod
    def _is_same_release_line(local_v: tuple, remote_v: tuple) -> bool:
        # Considera mesma linha de release quando o major é igual.
        # Ex.: 1.0.0 e 1.0.1 (mesma linha), 1.0.0 e 9.0.0 (linha diferente).
        if not local_v or not remote_v:
            return False
        return int(local_v[0]) == int(remote_v[0])

    @staticmethod
    def _suggest_next_version(v: str) -> str:
        parts = [int(x) for x in re.findall(r"\d+", str(v or ""))]
        while len(parts) < 3:
            parts.append(0)
        parts = parts[:3]
        parts[2] += 1
        return ".".join(str(x) for x in parts)

    def _publish_version(self):
        if self.IS_FROZEN or self.APP_ENV != "development":
            messagebox.showinfo(
                "Publicar versao",
                "Use esta funcao somente no ambiente development (python main.py),\n"
                "onde o repositorio Git esta disponivel.",
            )
            return

        updates_dir = os.path.join(self.APP_DIR, "updates")
        manifest_path = os.path.join(updates_dir, "version.json")
        changelog_path = os.path.join(updates_dir, "changelog.txt")
        os.makedirs(updates_dir, exist_ok=True)

        manifest = {}
        if os.path.exists(manifest_path):
            try:
                with open(manifest_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    manifest = loaded
            except Exception:
                manifest = {}

        current_v = str(manifest.get("version") or self._remote_version or self.APP_VERSION).strip()
        suggested = self._suggest_next_version(current_v)
        new_version = simpledialog.askstring(
            "Publicar versao",
            "Informe a nova versao (ex.: 9.0.1):",
            initialvalue=suggested,
            parent=self,
        )
        if new_version is None:
            return
        new_version = str(new_version).strip()
        if not re.match(r"^\d+(?:\.\d+){2,3}$", new_version):
            messagebox.showerror("ERRO", "Versao invalida. Use formato como 9.0.1")
            return

        notes = simpledialog.askstring(
            "Publicar versao",
            "Resumo da versao (1 linha):",
            parent=self,
        )
        if notes is None:
            return
        notes = str(notes).strip()
        if not notes:
            messagebox.showerror("ERRO", "Resumo da versao e obrigatorio.")
            return

        alert_msg = simpledialog.askstring(
            "Publicar versao",
            "Alerta opcional para usuarios (pode deixar vazio):",
            parent=self,
        )
        if alert_msg is None:
            alert_msg = ""
        alert_msg = str(alert_msg).strip()

        setup_url = str(
            manifest.get("setup_url") or self.SETUP_DOWNLOAD_URL or "https://github.com/andersoncantenas-glitch/rotahub-api/releases/latest"
        ).strip()
        changelog_url = str(
            manifest.get("changelog_url")
            or self.CHANGELOG_URL
            or "https://raw.githubusercontent.com/andersoncantenas-glitch/rotahub-api/main/updates/changelog.txt"
        ).strip()

        new_manifest = {
            "version": new_version,
            "setup_url": setup_url,
            "changelog_url": changelog_url,
            "alert": alert_msg,
            "alerts": [notes] + ([alert_msg] if alert_msg else []),
        }

        try:
            with open(manifest_path, "w", encoding="utf-8") as f:
                json.dump(new_manifest, f, ensure_ascii=False, indent=2)
                f.write("\n")

            old_changelog = ""
            if os.path.exists(changelog_path):
                with open(changelog_path, "r", encoding="utf-8", errors="replace") as f:
                    old_changelog = f.read()

            stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            entry = f"{stamp} - v{new_version}\n- {notes}\n"
            if alert_msg:
                entry += f"- Alerta: {alert_msg}\n"
            entry += "\n"

            with open(changelog_path, "w", encoding="utf-8") as f:
                f.write(entry + old_changelog)

            self._apply_manifest(new_manifest)
            messagebox.showinfo(
                "OK",
                "Arquivos de atualizacao atualizados com sucesso.\n\n"
                "Proximo passo:\n"
                "git add updates/version.json updates/changelog.txt\n"
                f'git commit -m "release {new_version}"\n'
                "git push origin main",
            )
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao publicar versao:\n{e}")

    def _apply_manifest(self, manifest: dict):
        self._clear_update_pending_if_installed()
        remote_version = str(manifest.get("version") or "-").strip() or "-"
        self._remote_version = remote_version
        self._remote_setup_url = str(manifest.get("setup_url") or self._remote_setup_url or "").strip()
        self._remote_changelog_url = str(manifest.get("changelog_url") or self._remote_changelog_url or "").strip()

        alerts = manifest.get("alerts")
        if isinstance(alerts, list):
            alert_txt = "\n".join(str(x) for x in alerts if str(x).strip())
        else:
            alert_txt = str(manifest.get("alert") or "").strip()
        self._alerts_text = alert_txt

        local_v = self._version_tuple(self.APP_VERSION)
        remote_v = self._version_tuple(remote_version)
        is_candidate = (remote_v > local_v) and self._is_same_release_line(local_v, remote_v)
        has_new_version = remote_v > local_v
        if remote_v <= local_v:
            # Evita instalar downgrade quando manifesto remoto estiver atrasado.
            self._remote_setup_url = ""
        remote_display = remote_version if (remote_version != "-" and has_new_version) else "-"

        st = self._load_update_state()
        pending_version = str(st.get("pending_update_version") or "").strip()
        if pending_version and self._version_tuple(self.APP_VERSION) < self._version_tuple(pending_version):
            if remote_display != "-":
                remote_display = f"{remote_display} | instalacao pendente: {pending_version}"
            else:
                remote_display = f"instalacao pendente: {pending_version}"

        self.lbl_version_remote.config(text=f"Versao disponivel: {remote_display}")
        self.lbl_alerts.config(text=alert_txt or "Sem alertas.")

        if is_candidate:
            self.lbl_version_remote.config(foreground="#B45309")
        else:
            self.lbl_version_remote.config(foreground="#166534")

    def _check_updates(self, silent=False):
        try:
            def _fetch_api_version() -> str:
                req = urllib.request.Request(_build_api_url("openapi.json"), method="GET")
                with urllib.request.urlopen(req, timeout=8) as resp:
                    data = resp.read().decode("utf-8", errors="replace")
                openapi = json.loads(data)
                info = openapi.get("info") if isinstance(openapi, dict) else {}
                return str((info or {}).get("version") or "").strip()

            manifest = None
            if self.UPDATE_MANIFEST_URL:
                req = urllib.request.Request(self.UPDATE_MANIFEST_URL, method="GET")
                with urllib.request.urlopen(req, timeout=8) as resp:
                    data = resp.read().decode("utf-8", errors="replace")
                loaded = json.loads(data)
                if isinstance(loaded, dict):
                    manifest = loaded

            if manifest is None:
                api_version = _fetch_api_version()
                if not api_version:
                    raise ValueError("Nao foi possivel obter versao remota (manifesto e API).")
                manifest = {
                    "version": api_version,
                    "setup_url": self.SETUP_DOWNLOAD_URL,
                    "changelog_url": self.CHANGELOG_URL,
                    "alert": "",
                    "alerts": [
                        "Versao remota obtida via API (fallback).",
                        "Publique updates/version.json para habilitar update automatizado completo.",
                    ],
                }
            else:
                # Manifesto pode ficar atrasado no GitHub/Render.
                # Compara com a versao da API e usa a maior para evitar downgrade.
                try:
                    api_version = _fetch_api_version()
                except Exception:
                    api_version = ""
                if api_version and self._version_tuple(api_version) > self._version_tuple(str(manifest.get("version") or "")):
                    alerts = manifest.get("alerts")
                    if not isinstance(alerts, list):
                        alerts = []
                    alerts.append(
                        "Manifesto de update esta atrasado. Usando versao detectada na API."
                    )
                    manifest["version"] = api_version
                    manifest["setup_url"] = ""
                    manifest["alerts"] = alerts

            self._apply_manifest(manifest)

            if not silent:
                local_v = self._version_tuple(self.APP_VERSION)
                remote_v = self._version_tuple(self._remote_version)
                is_candidate = (remote_v > local_v) and self._is_same_release_line(local_v, remote_v)
                if is_candidate:
                    ask = messagebox.askyesno(
                        "Atualizacao disponivel",
                        f"Versao local: {self.APP_VERSION}\nVersao disponivel: {self._remote_version}\n\n"
                        "Deseja atualizar agora?",
                    )
                    if ask:
                        self._update_now()
                else:
                    messagebox.showinfo("Atualizacoes", f"Voce ja esta na versao mais recente ({self.APP_VERSION}).")
        except Exception as e:
            if not silent:
                messagebox.showerror("ERRO", f"Falha ao verificar atualizacoes:\n{e}")

    def _update_state_path(self):
        base_dir = self.USER_DATA_DIR if self.IS_FROZEN else self.APP_DIR
        return os.path.join(base_dir, "update_state.json")

    def _load_update_state(self):
        try:
            p = self._update_state_path()
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            logging.debug("Falha ignorada")
        return {}

    def _save_update_state(self, data: dict):
        try:
            p = self._update_state_path()
            os.makedirs(os.path.dirname(p) or ".", exist_ok=True)
            with open(p, "w", encoding="utf-8") as f:
                json.dump(data or {}, f, ensure_ascii=False, indent=2)
                f.write("\n")
        except Exception:
            logging.debug("Falha ignorada")

    def _changelog_preview(self, max_lines=18):
        # 1) tenta arquivo local do repositorio
        local_candidates = [
            os.path.join(self.APP_DIR, "updates", "changelog.txt"),
            os.path.join(self.RESOURCE_DIR, "updates", "changelog.txt"),
        ]
        for p in local_candidates:
            try:
                if os.path.exists(p):
                    with open(p, "r", encoding="utf-8", errors="replace") as f:
                        lines = [ln.rstrip() for ln in f.read().splitlines() if ln.strip()]
                    if lines:
                        return "\n".join(lines[:max_lines])
            except Exception:
                logging.debug("Falha ignorada")

        # 2) tenta URL remota (se disponivel)
        url = self._remote_changelog_url or self.CHANGELOG_URL
        if url:
            try:
                req = urllib.request.Request(url, method="GET")
                with urllib.request.urlopen(req, timeout=6) as resp:
                    data = resp.read().decode("utf-8", errors="replace")
                lines = [ln.rstrip() for ln in data.splitlines() if ln.strip()]
                if lines:
                    return "\n".join(lines[:max_lines])
            except Exception:
                logging.debug("Falha ignorada")

        return "Sem detalhes de changelog disponíveis."

    def _maybe_show_post_update_notifications(self):
        if self._update_notice_shown:
            return
        self._update_notice_shown = True

        st = self._load_update_state()
        last_seen = str(st.get("last_seen_version") or "").strip()

        if self._version_tuple(self.APP_VERSION) <= self._version_tuple(last_seen):
            return

        alert_block = (self._alerts_text or "").strip()
        changelog = self._changelog_preview()
        notes = self._extract_latest_notes(changelog, self.APP_VERSION)
        lines = [
            f"Versão atual: {self.APP_VERSION}",
            f"Versão anterior vista: {last_seen or '-'}",
            "",
        ]
        if alert_block:
            lines.append("Alertas:")
            lines.extend([f"- {x.strip()}" for x in str(alert_block).splitlines() if x.strip()])
            lines.append("")
        lines.append("O que foi atualizado:")
        lines.extend([f"- {n}" for n in notes if str(n).strip()])
        msg = "\n".join(lines)
        self._show_update_notification_window(msg)

    def _extract_latest_notes(self, changelog_text: str, version: str):
        try:
            lines = [ln.strip() for ln in str(changelog_text or "").splitlines() if ln.strip()]
            version_tag = f"v{version}".lower()
            notes = []
            inside_target = False

            for ln in lines:
                ln_low = ln.lower()
                if re.search(r"\bv\d+\.\d+\.\d+(?:\.\d+)?\b", ln_low):
                    if version_tag in ln_low:
                        inside_target = True
                        continue
                    if inside_target:
                        break
                    inside_target = False
                    continue

                if inside_target and ln.startswith("-"):
                    item = ln.lstrip("-").strip()
                    if item:
                        notes.append(item)

            if notes:
                return notes[:20]

            # Fallback: pega as primeiras linhas tipo item do changelog
            fallback = []
            for ln in lines:
                if ln.startswith("-"):
                    item = ln.lstrip("-").strip()
                    if item:
                        fallback.append(item)
                if len(fallback) >= 12:
                    break
            if fallback:
                return fallback
        except Exception:
            logging.debug("Falha ignorada")

        return ["Sem detalhes de atualização informados."]

    def _show_update_notification_window(self, text_msg: str):
        win = tk.Toplevel(self)
        win.title("Novidades da Atualização")
        self.apply_window_icon(win)
        win.geometry("760x520")
        win.minsize(620, 420)
        win.transient(self.winfo_toplevel())
        try:
            win.grab_set()
        except Exception:
            logging.debug("Falha ignorada")

        root = ttk.Frame(win, style="Content.TFrame", padding=14)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        ttk.Label(
            root,
            text="Atualizações instaladas",
            style="CardTitle.TLabel",
        ).grid(row=0, column=0, sticky="w")

        txt = tk.Text(
            root,
            wrap="word",
            font=("Segoe UI", 10),
            bg="white",
            fg="#111827",
            relief="solid",
            bd=1,
        )
        txt.grid(row=1, column=0, sticky="nsew", pady=(8, 8))
        txt.insert("1.0", text_msg)
        txt.config(state="disabled")

        btns = ttk.Frame(root, style="Content.TFrame")
        btns.grid(row=2, column=0, sticky="e")

        def _close_and_mark():
            st = self._load_update_state()
            st["last_seen_version"] = self.APP_VERSION
            st["last_seen_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._save_update_state(st)
            try:
                win.destroy()
            except Exception:
                logging.debug("Falha ignorada")

        ttk.Button(btns, text="Marcar como lido", style="Primary.TButton", command=_close_and_mark).pack(side="right")
        ttk.Button(btns, text="Fechar", style="Ghost.TButton", command=_close_and_mark).pack(side="right", padx=(0, 8))
        win.protocol("WM_DELETE_WINDOW", _close_and_mark)

    def _build_api_status_badge(self):
        self.lbl_api_status = tk.Label(
            self.header_right,
            text="API: verificando...",
            bg="#E5E7EB",
            fg="#111827",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
        )
        self.lbl_api_status.pack(anchor="e")

    def _update_api_status(self):
        if not self.can_read_from_api():
            self.lbl_api_status.config(text=f"API: LOCAL ONLY ({self.APP_ENV})", bg="#DBEAFE", fg="#1D4ED8")
            if self._api_job:
                try:
                    self.after_cancel(self._api_job)
                except Exception:
                    logging.debug("Falha ignorada")
            self._api_job = self.after(15000, self._update_api_status)
            return

        online = False
        integracao_ok = None
        api_host = "-"
        try:
            url = _build_api_url("openapi.json")
            try:
                api_host = urllib.parse.urlparse(url).netloc or url
            except Exception:
                api_host = url
            req = urllib.request.Request(url, method="GET")
            with urllib.request.urlopen(req, timeout=min(max(self.API_SYNC_TIMEOUT, 3.0), 10.0)) as resp:
                online = int(getattr(resp, "status", 0) or 0) == 200
        except Exception:
            online = False
            if api_host == "-":
                try:
                    api_host = urllib.parse.urlparse(self.API_BASE_URL).netloc or self.API_BASE_URL
                except Exception:
                    api_host = self.API_BASE_URL

        if online and self.can_read_from_api():
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            if desktop_secret:
                try:
                    _call_api(
                        "GET",
                        "admin/motoristas/acesso",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    integracao_ok = True
                except Exception:
                    integracao_ok = False

        if online and integracao_ok is False:
            self.lbl_api_status.config(
                text=f"API: PARCIAL ({api_host})",
                bg="#FEF3C7",
                fg="#92400E",
            )
        elif online:
            self.lbl_api_status.config(text=f"API: ONLINE ({api_host})", bg="#DCFCE7", fg="#166534")
        else:
            self.lbl_api_status.config(text=f"API: OFFLINE ({api_host})", bg="#FEE2E2", fg="#991B1B")

        if self._api_job:
            try:
                self.after_cancel(self._api_job)
            except Exception:
                logging.debug("Falha ignorada")
        self._api_job = self.after(15000, self._update_api_status)

    def _build_stat(self, parent, col, title, value):
        c = ttk.Frame(parent, style="Card.TFrame", padding=16)
        c.grid(row=0, column=col, sticky="ew", padx=6)
        ttk.Label(c, text=title, style="CardLabel.TLabel").pack(anchor="w")
        lbl = ttk.Label(c, text=value, font=("Segoe UI", 24, "bold"), background="white")
        lbl.pack(anchor="w", pady=(6, 0))
        return lbl

    def _build_summary_tile(self, parent, col, title, value="0", value_font=("Segoe UI", 21, "bold"), top_pad=(0, 0)):
        card = ttk.Frame(parent, style="CardInset.TFrame", padding=(10, 8))
        card.grid(row=1, column=col, sticky="ew", padx=4, pady=top_pad)
        ttk.Label(card, text=title, style="InsetTitle.TLabel").pack(anchor="w")
        if value_font == ("Segoe UI", 21, "bold"):
            lbl = ttk.Label(card, text=value, style="InsetValue.TLabel")
        else:
            lbl = ttk.Label(card, text=value, style="InsetStrong.TLabel", font=value_font)
        lbl.pack(anchor="w", pady=(6, 0))
        return lbl

    def _build_info_row(self, parent, row, title):
        item = ttk.Frame(parent, style="CardInset.TFrame")
        item.grid(row=row, column=0, sticky="ew", pady=(8 if row == 1 else 6, 0))
        item.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(item, style="CardInset.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        title_lbl = ttk.Label(header, text=f"{title}:", style="InsetTitle.TLabel")
        title_lbl.grid(row=0, column=0, sticky="w")

        btn = ttk.Button(
            header,
            text="Expandir",
            style="Toggle.TButton",
            command=lambda item_key=title: self._toggle_runtime_diag_item(item_key),
        )
        btn.grid(row=0, column=1, sticky="e", padx=(8, 0))

        content = ttk.Frame(item, style="CardInset.TFrame")
        content.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        content.grid_columnconfigure(0, weight=1)

        value = tk.Text(
            content,
            height=1,
            wrap="none",
            font=("Segoe UI", 9),
            background="white",
            foreground="#111827",
            relief="flat",
            bd=0,
            padx=6,
            pady=6,
            highlightthickness=1,
            highlightbackground="#E5E7EB",
            highlightcolor="#93C5FD",
            insertwidth=0,
        )
        value.grid(row=0, column=0, sticky="ew")
        value.configure(state="disabled")

        xscroll = ttk.Scrollbar(content, orient="horizontal", command=value.xview)
        xscroll.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        value.configure(xscrollcommand=xscroll.set)
        xscroll.grid_remove()
        content.grid_remove()

        title_lbl.bind("<Button-1>", lambda _event, item_key=title: self._toggle_runtime_diag_item(item_key), add="+")
        title_lbl.configure(cursor="hand2")
        value.bind("<Configure>", lambda _event, item_key=title: self._refresh_runtime_diag_scrollbar(item_key), add="+")

        self._runtime_diag_rows[title] = {
            "content": content,
            "button": btn,
            "text": value,
            "scrollbar": xscroll,
        }
        return value

    def _set_runtime_diag_value(self, item_key, value):
        row = self._runtime_diag_rows.get(item_key)
        if not row:
            return
        text_widget = row["text"]
        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")
        text_widget.insert("1.0", str(value or "-"))
        text_widget.configure(state="disabled")
        self.after_idle(lambda key=item_key: self._refresh_runtime_diag_scrollbar(key))

    def _refresh_runtime_diag_scrollbar(self, item_key):
        row = self._runtime_diag_rows.get(item_key)
        if not row:
            return
        text_widget = row["text"]
        scrollbar = row["scrollbar"]
        try:
            start, end = text_widget.xview()
        except Exception:
            return
        if (end - start) >= 0.999:
            scrollbar.grid_remove()
        else:
            scrollbar.grid()

    def _toggle_runtime_diag_item(self, item_key):
        row = self._runtime_diag_rows.get(item_key)
        if not row:
            return
        content = row["content"]
        button = row["button"]
        if content.winfo_ismapped():
            content.grid_remove()
            button.configure(text="Expandir")
            return
        content.grid()
        button.configure(text="Ocultar")
        self.after_idle(lambda key=item_key: self._refresh_runtime_diag_scrollbar(key))

    def _build_pending_stat(self, parent, col, title):
        c = ttk.Frame(parent, style="CardInset.TFrame", padding=(10, 8))
        c.grid(row=1, column=col, sticky="ew", padx=4, pady=(8, 0))
        ttk.Label(c, text=title, style="InsetTitle.TLabel").pack(anchor="w")
        lbl = ttk.Label(c, text="0", style="InsetValueSmall.TLabel")
        lbl.pack(anchor="w", pady=(4, 0))
        return lbl

    def _update_clock(self):
        now = datetime.now().strftime("%d/%m/%Y  %H:%M:%S")
        self.lbl_clock.config(text=now)

        if self._clock_job:
            try:
                self.after_cancel(self._clock_job)
            except Exception:
                logging.debug("Falha ignorada")

        self._clock_job = self.after(1000, self._update_clock)

    def _home_api_rows(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret or not self.can_read_from_api():
            return []
        try:
            resp = _call_api(
                "GET",
                "desktop/monitoramento/rotas",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            api_rows = []
            if isinstance(resp, dict):
                api_rows = resp.get("rotas") if isinstance(resp.get("rotas"), list) else []
            elif isinstance(resp, list):
                api_rows = resp
            out = []
            finais = {"FINALIZADA", "FINALIZADO", "CANCELADA", "CANCELADO"}
            for r in (api_rows or []):
                if not isinstance(r, dict):
                    continue
                st = upper(str(r.get("status_operacional") or r.get("status") or "").strip())
                if st in finais:
                    continue
                codigo = upper(str(r.get("codigo_programacao") or "").strip())
                if not codigo:
                    continue
                motorista = upper(str(r.get("motorista") or "").strip())
                veiculo = upper(str(r.get("veiculo") or "").strip())
                data_ref = str(r.get("recorded_at") or r.get("data_criacao") or "").strip()
                out.append((codigo, motorista, veiculo, data_ref))
            return out
        except Exception:
            logging.debug("Falha ao buscar rotas da Home na API", exc_info=True)
            return []

    def _home_api_overview(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret or not self.can_read_from_api():
            return None
        try:
            resp = _call_api(
                "GET",
                "desktop/overview",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            if not isinstance(resp, dict):
                return None
            return {
                "total_prog": safe_int(resp.get("total_programacoes_ativas"), 0),
                "total_vendas": safe_int(resp.get("total_vendas_importadas"), 0),
                "total_clientes_ativos": safe_int(resp.get("total_clientes_ativos"), 0),
                "pending": resp.get("pendencias") if isinstance(resp.get("pendencias"), dict) else None,
                "rotas": resp.get("rotas") if isinstance(resp.get("rotas"), list) else [],
            }
        except Exception:
            logging.debug("Falha ao buscar overview da Home na API", exc_info=True)
            return None

    @staticmethod
    def _home_local_not_finalized_where(cols_prog: set) -> str:
        clauses = [
            "UPPER(TRIM(COALESCE(status, ''))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')",
        ]
        if "status_operacional" in cols_prog:
            clauses.append(
                "UPPER(TRIM(COALESCE(status_operacional, ''))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')"
            )
        if "finalizada_no_app" in cols_prog:
            clauses.append("COALESCE(finalizada_no_app, 0) = 0")
        if "data_chegada" in cols_prog:
            clauses.append("TRIM(COALESCE(data_chegada, '')) = ''")
        if "hora_chegada" in cols_prog:
            clauses.append("TRIM(COALESCE(hora_chegada, '')) = ''")
        if "km_final" in cols_prog:
            clauses.append("COALESCE(km_final, 0) = 0")
        return " AND ".join(clauses)

    def _load_pending_counts(self, total_prog_hint: int = 0) -> dict:
        out = {
            "rotas_abertas": max(int(total_prog_hint or 0), 0),
            "prestacao_pendente": 0,
            "sem_despesa": 0,
        }
        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            except Exception:
                cols_prog = set()

            status_col = "status_operacional" if "status_operacional" in cols_prog else "status"
            status_expr = f"UPPER(TRIM(COALESCE({status_col}, '')))"
            where_final = f"{status_expr} IN ('FINALIZADA', 'FINALIZADO')"
            if "prestacao_status" in cols_prog:
                where_prest = "UPPER(TRIM(COALESCE(prestacao_status, 'PENDENTE'))) <> 'FECHADA'"
            else:
                where_prest = "1=1"
            where_base = f"{where_final} AND {where_prest}"
            try:
                cur.execute(f"SELECT COUNT(*) FROM programacoes WHERE {where_base}")
                out["prestacao_pendente"] = safe_int((cur.fetchone() or [0])[0], 0)
            except Exception:
                out["prestacao_pendente"] = 0
            try:
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
                if cur.fetchone():
                    cur.execute(
                        f"""
                        SELECT COUNT(*)
                        FROM programacoes p
                        WHERE {where_base}
                          AND NOT EXISTS (
                              SELECT 1
                              FROM despesas d
                              WHERE UPPER(TRIM(COALESCE(d.codigo_programacao, '')))
                                  = UPPER(TRIM(COALESCE(p.codigo_programacao, '')))
                          )
                        """
                    )
                    out["sem_despesa"] = safe_int((cur.fetchone() or [0])[0], 0)
                else:
                    out["sem_despesa"] = 0
            except Exception:
                out["sem_despesa"] = 0
        return out

    def _refresh_home_footer(self, source: str):
        stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.lbl_footer_sync.config(text=f"Ultima atualizacao: {stamp}")
        self.lbl_footer_source.config(text=f"  |  Fonte: {source}")

    @staticmethod
    def _format_home_data(value) -> str:
        txt = str(value or "").strip()
        if not txt:
            return ""
        if len(txt) >= 19 and txt[4:5] == "-" and txt[7:8] == "-":
            try:
                dt = datetime.strptime(txt[:19], "%Y-%m-%d %H:%M:%S")
                return dt.strftime("%d/%m/%Y %H:%M")
            except Exception:
                pass
        if len(txt) >= 10 and txt[4:5] == "-" and txt[7:8] == "-":
            try:
                dt = datetime.strptime(txt[:10], "%Y-%m-%d")
                return dt.strftime("%d/%m/%Y")
            except Exception:
                pass
        return txt

    def _load_home_payload(self):
        total_prog = 0
        total_vendas = 0
        total_clientes_ativos = 0
        rows = []
        source = "LOCAL"

        api_overview = self._home_api_overview()
        if api_overview:
            total_prog = safe_int(api_overview.get("total_prog"), 0)
            total_vendas = safe_int(api_overview.get("total_vendas"), 0)
            total_clientes_ativos = safe_int(api_overview.get("total_clientes_ativos"), 0)
            rows = api_overview.get("rotas") or []
            source = "API CENTRAL"
            pending = api_overview.get("pending") or self._load_pending_counts(total_prog_hint=total_prog)
            return {
                "total_prog": total_prog,
                "total_vendas": total_vendas,
                "total_clientes_ativos": total_clientes_ativos,
                "rows": rows,
                "source": source,
                "pending": pending,
            }

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            except Exception:
                cols_prog = set()
            where_ativas = self._home_local_not_finalized_where(cols_prog)

            try:
                cur.execute("SELECT COUNT(*) FROM programacoes WHERE " + where_ativas)
                r = cur.fetchone()
                total_prog = (r[0] if r else 0) or 0
            except Exception:
                try:
                    cur.execute("SELECT COUNT(*) FROM programacoes")
                    r = cur.fetchone()
                    total_prog = (r[0] if r else 0) or 0
                except Exception:
                    total_prog = 0

            try:
                cur.execute("SELECT COUNT(*) FROM vendas_importadas")
                r = cur.fetchone()
                total_vendas = (r[0] if r else 0) or 0
            except Exception:
                total_vendas = 0

            try:
                cur.execute("""
                    SELECT COUNT(DISTINCT cod_cliente)
                    FROM programacao_itens
                    WHERE codigo_programacao IN (
                        SELECT codigo_programacao
                        FROM programacoes
                        WHERE """ + where_ativas + """
                    )
                """)
                r = cur.fetchone()
                total_clientes_ativos = (r[0] if r else 0) or 0
            except Exception:
                total_clientes_ativos = 0

            api_rows = self._home_api_rows()
            if api_rows:
                rows = api_rows
                total_prog = len(api_rows)
                source = "API CENTRAL"
            else:
                data_row_expr = "COALESCE(data_criacao,'')" if "data_criacao" in cols_prog else ("COALESCE(data,'')" if "data" in cols_prog else "''")
                try:
                    cur.execute(f"""
                        SELECT codigo_programacao, COALESCE(motorista,''), COALESCE(veiculo,''), {data_row_expr} AS data_criacao
                        FROM programacoes
                        WHERE """ + where_ativas + """
                        ORDER BY id DESC
                        LIMIT 80
                    """)
                    rows = cur.fetchall() or []
                except Exception:
                    try:
                        cur.execute(f"""
                            SELECT codigo_programacao, COALESCE(motorista,''), COALESCE(veiculo,''), {data_row_expr} AS data_criacao
                            FROM programacoes
                            ORDER BY id DESC
                            LIMIT 80
                        """)
                        rows = cur.fetchall() or []
                    except Exception:
                        rows = []

        pending = self._load_pending_counts(total_prog_hint=total_prog)
        return {
            "total_prog": total_prog,
            "total_vendas": total_vendas,
            "total_clientes_ativos": total_clientes_ativos,
            "rows": rows,
            "source": source,
            "pending": pending,
        }

    def _apply_home_payload(self, seq, payload):
        if seq != self._load_seq or not self.winfo_exists():
            return
        payload = payload or {}
        rows = payload.get("rows") or []
        source = payload.get("source") or "LOCAL"
        pending = payload.get("pending") or {}

        self.tree_rotas.delete(*self.tree_rotas.get_children())
        for r in rows:
            if isinstance(r, dict):
                codigo = upper(r.get("codigo_programacao") or "")
                motorista = upper(r.get("motorista") or "")
                veiculo = upper(r.get("veiculo") or "")
                data_criacao = str(r.get("data_criacao") or r.get("recorded_at") or "")
            else:
                codigo = upper(r[0] if len(r) > 0 else "")
                motorista = upper(r[1] if len(r) > 1 else "")
                veiculo = upper(r[2] if len(r) > 2 else "")
                data_criacao = str(r[3] if len(r) > 3 else "")
            data_fmt = self._format_home_data(data_criacao)
            tree_insert_aligned(self.tree_rotas, "", "end", (codigo, motorista, veiculo, data_fmt))

        self.lbl_total_prog.config(text=str(payload.get("total_prog", 0)))
        self.lbl_total_vendas.config(text=str(payload.get("total_vendas", 0)))
        self.lbl_total_clientes_ativos.config(text=str(payload.get("total_clientes_ativos", 0)))
        self.lbl_pend_rotas.config(text=str(pending.get("rotas_abertas", 0)))
        self.lbl_pend_prest.config(text=str(pending.get("prestacao_pendente", 0)))
        self.lbl_pend_desp.config(text=str(pending.get("sem_despesa", 0)))
        self._refresh_home_footer(source)

        nome = self.app.user.get("nome", "")
        is_admin = bool(self.app.user.get("is_admin", False))
        self.set_status(f"STATUS: Logado como {nome} (ADMIN: {is_admin}). Fonte Home: {source}.")

    def _handle_home_load_error(self, seq, exc):
        if seq != self._load_seq or not self.winfo_exists():
            return
        logging.error("Falha ao carregar Home", exc_info=(type(exc), exc, exc.__traceback__))
        self.set_status("STATUS: Falha ao carregar a Home.")

    def on_show(self):
        self._load_seq += 1
        seq = self._load_seq
        self.set_status("STATUS: Carregando Home...")
        run_async_ui(
            self,
            self._load_home_payload,
            lambda payload, seq=seq: self._apply_home_payload(seq, payload),
            lambda exc, seq=seq: self._handle_home_load_error(seq, exc),
        )

    # =========================================================
    # PREVIEW ROTA âââââ€š¬Å¡¬ââââ‚¬Å¡¬ PUXA DADOS REAIS DA PROGRAMAÃÆââ‚¬â„¢ââââ‚¬Å¡¬¡ÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢O ATIVA
    # =========================================================
    def _open_rota_preview(self, event=None):
        sel = self.tree_rotas.selection()
        if not sel:
            return

        vals = self.tree_rotas.item(sel[0], "values")
        if not vals:
            return

        codigo = upper(vals[0])

        win = tk.Toplevel(self)
        win.title(f"Pre-visualizacao - {codigo}")
        win.geometry("1050x600")
        win.grab_set()
        win.resizable(True, True)

        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)
        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(2, weight=1)

        lbl_title = ttk.Label(frm, text=f"ROTA ATIVA: {codigo}", font=("Segoe UI", 18, "bold"))
        lbl_title.grid(row=0, column=0, sticky="w", pady=(0, 6))

        header = tk.Frame(frm, bg="#0B2A6F", bd=0, highlightthickness=0)
        header.grid(row=1, column=0, sticky="ew", pady=(8, 10))
        header.grid_columnconfigure((0, 1, 2, 3), weight=1)
        lbl_status = tk.Label(header, text="Status: ", font=("Segoe UI", 10, "bold"), bg="#0B2A6F", fg="#FFFFFF")
        lbl_motorista = tk.Label(header, text="Motorista: ", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_equipe = tk.Label(header, text="Equipe: ", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_nf = tk.Label(header, text="NF: ", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")

        lbl_adiant = tk.Label(header, text="Adiantamento: R$ 0,00", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_local_rota = tk.Label(header, text="Local da Rota: ", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_saida = tk.Label(header, text="Saida: -", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_local_carreg = tk.Label(header, text="Carregou em: ", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")

        lbl_kg_carregado = tk.Label(header, text="KG Carregado: 0,00", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_caixas_carreg = tk.Label(header, text="Caixas Carregadas: 0", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_caixas_entreg = tk.Label(header, text="Caixas Entregues: 0", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_saldo = tk.Label(header, text="Saldo (NF - Granja): 0,00", font=("Segoe UI", 10, "bold"), bg="#0B2A6F", fg="#FFFFFF")
        lbl_km = tk.Label(header, text="KM: ", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_media_carreg = tk.Label(header, text="Media Carregada: 0,000", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_caixa_final = tk.Label(header, text="Caixa Final: -", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_subst = tk.Label(header, text="Substituicao: -", font=("Segoe UI", 10), bg="#0B2A6F", fg="#FFFFFF")
        lbl_fonte = tk.Label(
            header,
            text="Fonte: LOCAL",
            font=("Segoe UI", 10, "bold"),
            bg="#0B2A6F",
            fg="#FBBF24",
        )

        lbl_status.grid(row=0, column=0, sticky="w", padx=6, pady=2)
        lbl_motorista.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        lbl_equipe.grid(row=0, column=2, sticky="w", padx=6, pady=2)
        lbl_nf.grid(row=0, column=3, sticky="w", padx=6, pady=2)

        lbl_adiant.grid(row=1, column=0, sticky="w", padx=6, pady=2)
        lbl_local_rota.grid(row=1, column=1, sticky="w", padx=6, pady=2)
        lbl_saida.grid(row=1, column=2, sticky="w", padx=6, pady=2)
        lbl_local_carreg.grid(row=1, column=3, sticky="w", padx=6, pady=2)

        lbl_kg_carregado.grid(row=2, column=0, sticky="w", padx=6, pady=2)
        lbl_caixas_carreg.grid(row=2, column=1, sticky="w", padx=6, pady=2)
        lbl_caixas_entreg.grid(row=2, column=2, sticky="w", padx=6, pady=2)
        lbl_saldo.grid(row=2, column=3, sticky="w", padx=6, pady=2)
        lbl_km.grid(row=3, column=0, sticky="w", padx=6, pady=2)
        lbl_media_carreg.grid(row=3, column=1, sticky="w", padx=6, pady=2)
        lbl_caixa_final.grid(row=3, column=2, sticky="w", padx=6, pady=2)
        lbl_subst.grid(row=3, column=3, sticky="w", padx=6, pady=2)
        lbl_fonte.grid(row=4, column=0, columnspan=4, sticky="w", padx=6, pady=(2, 4))

        table_wrap = ttk.Frame(frm)
        table_wrap.grid(row=2, column=0, sticky="nsew")
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        cols = [
            "COD CLIENTE", "NOME", "CX", "KG", "PRECO", "PEDIDO",
            "STATUS_PEDIDO", "CAIXAS_ATUAL", "PRECO_ATUAL", "ALTERADO_EM",
            "RECEBIDO_VALOR", "MORTALIDADE"
        ]
        tree = ttk.Treeview(table_wrap, columns=cols, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")
        detalhes_por_iid = {}
        cabecalho_ctx = {
            "codigo": codigo,
            "nf_numero": "",
            "motorista": "",
            "veiculo": "",
            "equipe": "",
            "local_carreg": "",
            "local_rota": "",
        }
        tree.tag_configure("st_entregue", foreground="#1B5E20")
        tree.tag_configure("st_pendente", foreground="#E65100")
        tree.tag_configure("st_cancelado", foreground="#B71C1C")
        tree.tag_configure("st_alterado", foreground="#0D47A1")
        tree.tag_configure("st_default", foreground="#111827")

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=130, anchor="w")

        tree.column("COD CLIENTE", width=100, anchor="center")
        tree.column("NOME", width=260, anchor="w")
        tree.column("CX", width=70, anchor="center")
        tree.column("KG", width=90, anchor="center")
        tree.column("PRECO", width=90, anchor="center")
        tree.column("PEDIDO", width=140, anchor="center")
        tree.column("STATUS_PEDIDO", width=140, anchor="center")
        tree.column("CAIXAS_ATUAL", width=110, anchor="center")
        tree.column("PRECO_ATUAL", width=110, anchor="center")
        tree.column("ALTERADO_EM", width=150, anchor="center")
        tree.column("RECEBIDO_VALOR", width=120, anchor="center")
        tree.column("MORTALIDADE", width=110, anchor="center")

        try:
            enable_treeview_sorting(
                tree,
                numeric_cols={"CX", "CAIXAS_ATUAL", "MORTALIDADE"},
                money_cols={"PRECO", "PRECO_ATUAL", "RECEBIDO_VALOR"},
                date_cols={"ALTERADO_EM"}
            )
        except Exception:
            logging.debug("Falha ignorada")

        def _fmt_money(v):
            try:
                v = float(v or 0.0)
                s = f"{v:,.2f}"
                return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return "R$ 0,00"

        def _normalize_unit_price(v):
            """
            Compatibilidade: alguns bancos guardam preço unitário em centavos
            (ex.: 630 para R$ 6,30). Ajusta apenas para exibição no preview.
            """
            val = safe_float(v, 0.0)
            try:
                is_integer_like = abs(val - round(val)) < 1e-9
            except Exception:
                is_integer_like = False
            if is_integer_like and 100.0 <= abs(val) <= 100000.0:
                return val / 100.0
            return val

        def _fmt_pedido(v):
            s = str(v or "").strip()
            if not s:
                return ""
            try:
                n = float(s.replace(",", "."))
                if abs(n - round(n)) < 1e-9:
                    return str(int(round(n)))
            except Exception:
                pass
            return s

        def _table_cols(cur, table_name: str):
            try:
                cur.execute(f"PRAGMA table_info({table_name})")
                return {str(r[1]).lower() for r in (cur.fetchall() or [])}
            except Exception:
                return set()

        def _collect_transferencias(cod_cli: str, pedido: str):
            out = []
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='transferencias'")
                    if not cur.fetchone():
                        return out

                    cols = _table_cols(cur, "transferencias")
                    cond_pedido = "COALESCE(TRIM(pedido),'')=COALESCE(TRIM(?),'')" if "pedido" in cols else "1=1"
                    params = [codigo, cod_cli]
                    if "pedido" in cols:
                        params.append(pedido or "")

                    cur.execute(
                        f"""
                        SELECT *
                        FROM transferencias
                        WHERE (
                            (COALESCE(TRIM(codigo_origem),'')=COALESCE(TRIM(?),'') AND COALESCE(TRIM(cod_cliente),'')=COALESCE(TRIM(?),''))
                            OR
                            (COALESCE(TRIM(codigo_destino),'')=COALESCE(TRIM(?),'') AND COALESCE(TRIM(cod_cliente),'')=COALESCE(TRIM(?),''))
                        )
                          AND {cond_pedido}
                        ORDER BY COALESCE(atualizado_em, criado_em, '') DESC
                        LIMIT 30
                        """,
                        [codigo, cod_cli, codigo, cod_cli, *([] if "pedido" not in cols else [pedido or ""])],
                    )
                    rows = cur.fetchall() or []
                    for r in rows:
                        d = dict(r) if hasattr(r, "keys") else {}
                        if not d:
                            continue
                        out.append({
                            "status": str(d.get("status") or ""),
                            "origem_prog": str(d.get("codigo_origem") or ""),
                            "destino_prog": str(d.get("codigo_destino") or ""),
                            "qtd": safe_int(d.get("qtd_caixas"), 0),
                            "obs": str(d.get("obs") or ""),
                            "mot_origem": str(d.get("motorista_origem") or ""),
                            "mot_destino": str(d.get("motorista_destino") or ""),
                            "criado_em": str(d.get("criado_em") or ""),
                            "atualizado_em": str(d.get("atualizado_em") or ""),
                        })
            except Exception:
                logging.debug("Falha ao coletar transferências do cliente", exc_info=True)
            return out

        def _collect_item_logs(cod_cli: str, pedido: str):
            out = []
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens_log'")
                    if not cur.fetchone():
                        return out
                    cols = _table_cols(cur, "programacao_itens_log")
                    if "codigo_programacao" not in cols or "cod_cliente" not in cols:
                        return out
                    has_pedido = "pedido" in cols
                    order_col = "id" if "id" in cols else ("criado_em" if "criado_em" in cols else None)
                    where = "WHERE COALESCE(TRIM(codigo_programacao),'')=COALESCE(TRIM(?),'') AND COALESCE(TRIM(cod_cliente),'')=COALESCE(TRIM(?),'')"
                    params = [codigo, cod_cli]
                    if has_pedido:
                        where += " AND COALESCE(TRIM(pedido),'')=COALESCE(TRIM(?),'')"
                        params.append(pedido or "")
                    order_sql = f"ORDER BY {order_col} DESC" if order_col else ""
                    cur.execute(f"SELECT * FROM programacao_itens_log {where} {order_sql} LIMIT 40", tuple(params))
                    rows = cur.fetchall() or []
                    for r in rows:
                        d = dict(r) if hasattr(r, "keys") else {}
                        if not d:
                            continue
                        evento = str(d.get("evento") or d.get("tipo") or d.get("acao") or "")
                        payload_raw = d.get("payload_json") if "payload_json" in d else (d.get("payload") or d.get("detalhes") or "")
                        payload_obj = {}
                        if isinstance(payload_raw, dict):
                            payload_obj = payload_raw
                            payload_txt = json.dumps(payload_raw, ensure_ascii=False)
                        else:
                            payload_txt = str(payload_raw or "")
                            if payload_txt:
                                try:
                                    tmp = json.loads(payload_txt)
                                    if isinstance(tmp, dict):
                                        payload_obj = tmp
                                except Exception:
                                    payload_obj = {}
                        dt = str(
                            d.get("created_at")
                            or d.get("criado_em")
                            or d.get("alterado_em")
                            or d.get("data_registro")
                            or ""
                        )
                        out.append({"evento": evento, "payload": payload_txt, "payload_obj": payload_obj, "quando": dt})
            except Exception:
                logging.debug("Falha ao coletar logs do item", exc_info=True)
            return out

        def _extract_event_context(logs_item):
            """
            Extrai horário/localização a partir dos logs do item.
            Mantém fallback para múltiplos formatos de payload.
            """
            ctx = {"quando": "", "latitude": "", "longitude": "", "endereco": "", "cidade": "", "bairro": ""}
            for lg in logs_item or []:
                if not ctx["quando"]:
                    ctx["quando"] = str(lg.get("quando") or "").strip()
                payload_obj = lg.get("payload_obj") if isinstance(lg, dict) else {}
                if not isinstance(payload_obj, dict):
                    payload_obj = {}
                for k in ("lat_evento", "latitude", "lat", "cliente_lat", "cliente_latitude"):
                    if not ctx["latitude"] and payload_obj.get(k) not in (None, ""):
                        ctx["latitude"] = str(payload_obj.get(k)).strip()
                        break
                for k in ("lon_evento", "longitude", "lon", "lng", "cliente_lon", "cliente_longitude"):
                    if not ctx["longitude"] and payload_obj.get(k) not in (None, ""):
                        ctx["longitude"] = str(payload_obj.get(k)).strip()
                        break
                for k in ("endereco", "endereco_cliente"):
                    if not ctx["endereco"] and payload_obj.get(k):
                        ctx["endereco"] = str(payload_obj.get(k)).strip()
                        break
                for k in ("cidade", "cidade_cliente"):
                    if not ctx["cidade"] and payload_obj.get(k):
                        ctx["cidade"] = str(payload_obj.get(k)).strip()
                        break
                for k in ("bairro", "bairro_cliente"):
                    if not ctx["bairro"] and payload_obj.get(k):
                        ctx["bairro"] = str(payload_obj.get(k)).strip()
                        break
            return ctx

        def _collect_cliente_location(cod_cli: str):
            loc = {
                "latitude": "",
                "longitude": "",
                "endereco": "",
                "cidade": "",
                "bairro": "",
                "amostras": 0,
                "ultima_coleta_em": "",
                "ultima_status": "",
                "ultima_origem": "",
                "lat_recorrente": "",
                "lon_recorrente": "",
                "endereco_recorrente": "",
                "cidade_recorrente": "",
                "bairro_recorrente": "",
                "recorrencia_qtd": 0,
            }
            try:
                with get_db() as conn:
                    cur = conn.cursor()

                    def _has_table(table_name: str) -> bool:
                        try:
                            cur.execute(
                                "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
                                (table_name,),
                            )
                            return cur.fetchone() is not None
                        except Exception:
                            return False

                    cols = _table_cols(cur, "clientes")
                    if cols:
                        lat_col = "latitude" if "latitude" in cols else ("lat" if "lat" in cols else "")
                        lon_col = "longitude" if "longitude" in cols else ("lon" if "lon" in cols else "")
                        end_col = "endereco" if "endereco" in cols else ""
                        cid_col = "cidade" if "cidade" in cols else ""
                        bai_col = "bairro" if "bairro" in cols else ""
                        sel = [
                            (lat_col if lat_col else "''") + " as latitude",
                            (lon_col if lon_col else "''") + " as longitude",
                            (end_col if end_col else "''") + " as endereco",
                            (cid_col if cid_col else "''") + " as cidade",
                            (bai_col if bai_col else "''") + " as bairro",
                        ]
                        cur.execute(
                            f"SELECT {', '.join(sel)} FROM clientes WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?) LIMIT 1",
                            (cod_cli,),
                        )
                        r = cur.fetchone()
                        if r:
                            rr = dict(r) if hasattr(r, "keys") else {}
                            for k in ("latitude", "longitude", "endereco", "cidade", "bairro"):
                                loc[k] = str(rr.get(k) or "")

                    if _has_table("cliente_localizacao_amostras"):
                        cur.execute(
                            """
                            SELECT COUNT(*) AS qtd, MAX(COALESCE(registrado_em, '')) AS ultima
                            FROM cliente_localizacao_amostras
                            WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)
                            """,
                            (cod_cli,),
                        )
                        r_count = cur.fetchone()
                        if r_count:
                            loc["amostras"] = safe_int(r_count["qtd"], 0)
                            loc["ultima_coleta_em"] = str(r_count["ultima"] or "")

                        cur.execute(
                            """
                            SELECT latitude, longitude, endereco, cidade, bairro, status_pedido, origem, registrado_em
                            FROM cliente_localizacao_amostras
                            WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)
                            ORDER BY COALESCE(registrado_em, '') DESC, id DESC
                            LIMIT 1
                            """,
                            (cod_cli,),
                        )
                        r_last = cur.fetchone()
                        if r_last:
                            rr_last = dict(r_last) if hasattr(r_last, "keys") else {}
                            loc["ultima_status"] = str(rr_last.get("status_pedido") or "")
                            loc["ultima_origem"] = str(rr_last.get("origem") or "")
                            if rr_last.get("registrado_em"):
                                loc["ultima_coleta_em"] = str(rr_last.get("registrado_em") or "")
                            for k in ("latitude", "longitude", "endereco", "cidade", "bairro"):
                                if not str(loc.get(k) or "").strip() and rr_last.get(k) not in (None, ""):
                                    loc[k] = str(rr_last.get(k) or "")

                        cur.execute(
                            """
                            SELECT latitude, longitude, endereco, cidade, bairro, COUNT(*) AS freq
                            FROM cliente_localizacao_amostras
                            WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)
                              AND (
                                    latitude IS NOT NULL
                                 OR longitude IS NOT NULL
                                 OR TRIM(COALESCE(endereco, '')) <> ''
                                 OR TRIM(COALESCE(cidade, '')) <> ''
                                 OR TRIM(COALESCE(bairro, '')) <> ''
                              )
                            GROUP BY latitude, longitude, endereco, cidade, bairro
                            ORDER BY freq DESC, MAX(COALESCE(registrado_em, '')) DESC
                            LIMIT 1
                            """,
                            (cod_cli,),
                        )
                        r_best = cur.fetchone()
                        if r_best:
                            rr_best = dict(r_best) if hasattr(r_best, "keys") else {}
                            loc["lat_recorrente"] = str(rr_best.get("latitude") or "")
                            loc["lon_recorrente"] = str(rr_best.get("longitude") or "")
                            loc["endereco_recorrente"] = str(rr_best.get("endereco") or "")
                            loc["cidade_recorrente"] = str(rr_best.get("cidade") or "")
                            loc["bairro_recorrente"] = str(rr_best.get("bairro") or "")
                            loc["recorrencia_qtd"] = safe_int(rr_best.get("freq"), 0)
            except Exception:
                logging.debug("Falha ao coletar localização do cliente", exc_info=True)
            return loc

        def _fmt_distancia_rota(v):
            raw = str(v or "").strip()
            if not raw:
                return "-"
            try:
                val = float(raw.replace(",", "."))
            except Exception:
                return raw
            casas = 0 if abs(val) >= 100 else 1
            return f"{val:.{casas}f}".replace(".", ",")

        def _fmt_confianca_localizacao(v):
            raw = str(v or "").strip()
            if not raw:
                return "-"
            try:
                val = float(raw.replace(",", "."))
            except Exception:
                return raw
            if val <= 1:
                val *= 100.0
            casas = 0 if val >= 100 else 1
            return f"{val:.{casas}f}%".replace(".", ",")

        def _build_cliente_detalhe_text(item_data: dict):
            cod_cli = str(item_data.get("cod_cliente") or "")
            nome_cli = str(item_data.get("nome") or "")
            pedido = _fmt_pedido(item_data.get("pedido") or "")
            cx_orig = safe_int(item_data.get("cx"), 0)
            cx_atual = safe_int(item_data.get("caixas_atual"), 0) or cx_orig
            kg = safe_float(item_data.get("kg"), 0.0)
            preco_orig = safe_float(item_data.get("preco"), 0.0)
            preco_atual = safe_float(item_data.get("preco_atual"), 0.0) or preco_orig
            status_item = upper(item_data.get("status_pedido") or "")
            alterado_em = str(item_data.get("alterado_em") or "")
            alterado_por = str(item_data.get("alterado_por") or "")
            mortalidade = safe_int(item_data.get("mortalidade"), 0)
            recebido_valor = safe_float(item_data.get("recebido_valor"), 0.0)
            peso_previsto = safe_float(item_data.get("peso_previsto"), 0.0)
            forma_receb = str(item_data.get("forma_recebimento") or item_data.get("recebido_forma") or "")
            obs_receb = str(item_data.get("obs_recebimento") or item_data.get("recebido_obs") or "")
            alteracao_tipo = str(item_data.get("alteracao_tipo") or "")
            alteracao_detalhe = str(item_data.get("alteracao_detalhe") or "")
            ordem_sugerida = safe_int(item_data.get("ordem_sugerida"), 0)
            eta = str(item_data.get("eta") or "").strip()
            distancia = item_data.get("distancia")
            confianca_localizacao = item_data.get("confianca_localizacao")
            tem_roteirizacao = bool(
                ordem_sugerida > 0
                or eta
                or (distancia is not None and str(distancia).strip() != "")
                or (confianca_localizacao is not None and str(confianca_localizacao).strip() != "")
            )

            media_cli = (kg / cx_orig) if cx_orig > 0 else 0.0
            valor_orig = float(cx_orig) * float(preco_orig)
            valor_atual = float(cx_atual) * float(preco_atual)
            delta_cx = cx_atual - cx_orig
            delta_preco = preco_atual - preco_orig

            local = _collect_cliente_location(cod_cli)
            transferencias = _collect_transferencias(cod_cli, pedido)
            logs_item = _collect_item_logs(cod_cli, pedido)
            evt_ctx = _extract_event_context(logs_item)
            if not alterado_em:
                alterado_em = str(item_data.get("evento_datahora") or evt_ctx.get("quando", ""))
            for lk in ("latitude", "longitude", "endereco", "cidade", "bairro"):
                if not str(local.get(lk) or "").strip():
                    local[lk] = str(evt_ctx.get(lk) or "").strip()
            if not str(local.get("latitude") or "").strip():
                local["latitude"] = str(item_data.get("lat_evento") or "").strip()
            if not str(local.get("longitude") or "").strip():
                local["longitude"] = str(item_data.get("lon_evento") or "").strip()
            if not str(local.get("endereco") or "").strip():
                local["endereco"] = str(item_data.get("endereco_evento") or "").strip()
            if not str(local.get("cidade") or "").strip():
                local["cidade"] = str(item_data.get("cidade_evento") or "").strip()
            if not str(local.get("bairro") or "").strip():
                local["bairro"] = str(item_data.get("bairro_evento") or "").strip()
            checklist = [
                ("Status recebido do APK", bool(status_item)),
                ("Hora da alteracao", bool(alterado_em)),
                (
                    "Alteracao do pedido",
                    bool(alteracao_tipo or alteracao_detalhe or delta_cx != 0 or abs(delta_preco) > 0.0001),
                ),
                ("Sugestao de roteirizacao", tem_roteirizacao),
                (
                    "Coordenadas do evento",
                    bool(str(local.get("latitude") or "").strip() and str(local.get("longitude") or "").strip()),
                ),
                (
                    "Endereco do evento",
                    bool(
                        str(local.get("endereco") or "").strip()
                        or str(local.get("cidade") or "").strip()
                        or str(local.get("bairro") or "").strip()
                    ),
                ),
                ("Amostras armazenadas", safe_int(local.get("amostras"), 0) > 0),
            ]

            linhas = []
            linhas.append(f"PROGRAMAÇÃO: {cabecalho_ctx.get('codigo') or codigo}")
            linhas.append(f"NOTA FISCAL: {cabecalho_ctx.get('nf_numero') or '-'}")
            linhas.append(
                f"MOTORISTA: {cabecalho_ctx.get('motorista') or '-'}    VEICULO: {cabecalho_ctx.get('veiculo') or '-'}"
            )
            eq_raw = cabecalho_ctx.get("equipe") or ""
            linhas.append(f"EQUIPE: {self.resolve_equipe_nomes(eq_raw) or eq_raw or '-'}")
            linhas.append(
                f"LOCAL CARREGADO: {cabecalho_ctx.get('local_carreg') or '-'}    "
                f"LOCAL ROTA: {cabecalho_ctx.get('local_rota') or '-'}"
            )
            linhas.append("-" * 110)
            linhas.append(f"CLIENTE: {cod_cli} - {nome_cli}")
            linhas.append(f"PEDIDO: {pedido or '-'}")
            linhas.append(f"STATUS PEDIDO: {status_item or '-'}")
            linhas.append(f"HORA ALTERAÇÃO/ENTREGA: {alterado_em or '-'}")
            linhas.append(f"ALTERADO POR: {alterado_por or '-'}")
            linhas.append(f"TIPO ALTERAÇÃO: {alteracao_tipo or '-'}")
            linhas.append(f"DETALHE ALTERAÇÃO: {alteracao_detalhe or '-'}")
            linhas.append("-" * 110)
            linhas.append("[CARREGAMENTO / ENTREGA]")
            linhas.append(f"CAIXAS ORIGEM: {cx_orig}")
            linhas.append(f"CAIXAS ATUAL: {cx_atual}    DELTA: {delta_cx:+d}")
            linhas.append(f"PESO (KG): {kg:.2f}    PESO PREVISTO: {peso_previsto:.2f}")
            linhas.append(f"MÉDIA CLIENTE (KG/CX): {media_cli:.3f}")
            linhas.append(f"PREÇO ORIGEM: {_fmt_money(preco_orig)}")
            linhas.append(f"PREÇO ATUAL: {_fmt_money(preco_atual)}    DELTA: {_fmt_money(delta_preco)}")
            linhas.append(f"VALOR ORIGEM: {_fmt_money(valor_orig)}")
            linhas.append(f"VALOR ATUAL: {_fmt_money(valor_atual)}")
            linhas.append(f"RECEBIDO: {_fmt_money(recebido_valor)}    FORMA: {forma_receb or '-'}")
            linhas.append(f"OBS RECEBIMENTO: {obs_receb or '-'}")
            linhas.append(f"MORTALIDADE (AVES): {mortalidade}")
            linhas.append("-" * 110)
            linhas.append("[ROTEIRIZACAO / IA]")
            linhas.append(f"ORDEM SUGERIDA: {ordem_sugerida if ordem_sugerida > 0 else '-'}")
            linhas.append(f"ETA: {eta or '-'}")
            linhas.append(f"DISTANCIA: {_fmt_distancia_rota(distancia) or '-'}")
            linhas.append(
                f"CONFIANCA LOCALIZACAO: {_fmt_confianca_localizacao(confianca_localizacao) or '-'}"
            )
            linhas.append("-" * 110)
            linhas.append("[LOCALIZACAO CLIENTE]")
            linhas.append(f"LATITUDE: {local.get('latitude') or '-'}")
            linhas.append(f"LONGITUDE: {local.get('longitude') or '-'}")
            linhas.append(f"ENDEREÇO: {local.get('endereco') or '-'}")
            linhas.append(f"CIDADE: {local.get('cidade') or '-'}    BAIRRO: {local.get('bairro') or '-'}")
            linhas.append(
                f"AMOSTRAS COLETADAS: {safe_int(local.get('amostras'), 0)}    ULTIMA LEITURA: {local.get('ultima_coleta_em') or '-'}"
            )
            linhas.append(
                "LOCAL MAIS RECORRENTE: "
                f"{local.get('endereco_recorrente') or '-'} / {local.get('cidade_recorrente') or '-'} / {local.get('bairro_recorrente') or '-'} "
                f"(freq. {safe_int(local.get('recorrencia_qtd'), 0)})"
            )
            linhas.append("[CHECKLIST APK / SINCRONIA]")
            for rotulo, ok in checklist:
                linhas.append(f"{rotulo}: {'OK' if ok else 'PENDENTE'}")
            linhas.append("-" * 110)
            linhas.append("[TRANSFERÊNCIAS / DIRECIONAMENTO]")
            if not transferencias:
                linhas.append("Sem transferências registradas para este cliente/pedido.")
            else:
                for i, t in enumerate(transferencias, start=1):
                    linhas.append(
                        f"{i:02d}. STATUS={t['status'] or '-'} | ORIGEM={t['origem_prog'] or '-'} -> DESTINO={t['destino_prog'] or '-'} "
                        f"| QTD={t['qtd']} | MOT={t['mot_origem'] or '-'} -> {t['mot_destino'] or '-'} "
                        f"| EM={t['atualizado_em'] or t['criado_em'] or '-'}"
                    )
                    if t["obs"]:
                        linhas.append(f"    OBS: {t['obs']}")
            linhas.append("-" * 110)
            linhas.append("[LOGS / OCORRÊNCIAS]")
            if not logs_item:
                linhas.append("Sem logs detalhados para este cliente/pedido.")
            else:
                for i, lg in enumerate(logs_item, start=1):
                    linhas.append(f"{i:02d}. {lg['quando'] or '-'} | {lg['evento'] or '-'}")
                    if lg["payload"]:
                        linhas.append(f"    {str(lg['payload'])[:500]}")
            return "\n".join(linhas)

        def _open_cliente_detalhe(event=None):
            iid = None
            try:
                if event is not None:
                    iid = tree.identify_row(event.y)
            except Exception:
                iid = None
            if not iid:
                sel_i = tree.selection()
                if sel_i:
                    iid = sel_i[0]
            if not iid:
                return

            item_data = detalhes_por_iid.get(iid, {})
            if not item_data:
                vals = tree.item(iid, "values") or ()
                if vals:
                    item_data = {
                        "cod_cliente": vals[0] if len(vals) > 0 else "",
                        "nome": vals[1] if len(vals) > 1 else "",
                        "cx": vals[2] if len(vals) > 2 else 0,
                        "kg": vals[3] if len(vals) > 3 else 0,
                        "preco": vals[4] if len(vals) > 4 else 0,
                        "pedido": vals[5] if len(vals) > 5 else "",
                        "status_pedido": vals[6] if len(vals) > 6 else "",
                        "caixas_atual": vals[7] if len(vals) > 7 else 0,
                        "preco_atual": vals[8] if len(vals) > 8 else 0,
                        "alterado_em": vals[9] if len(vals) > 9 else "",
                        "recebido_valor": vals[10] if len(vals) > 10 else 0,
                        "mortalidade": vals[11] if len(vals) > 11 else 0,
                        "ordem_sugerida": "",
                        "eta": "",
                        "distancia": "",
                        "confianca_localizacao": "",
                    }

            texto = _build_cliente_detalhe_text(item_data)
            w = tk.Toplevel(win)
            w.title(f"Detalhes do Cliente - {item_data.get('cod_cliente', '')}")
            w.geometry("1060x720")
            w.transient(win)
            w.grab_set()

            area = ttk.Frame(w, padding=10)
            area.pack(fill="both", expand=True)
            area.grid_rowconfigure(1, weight=1)
            area.grid_columnconfigure(0, weight=1)

            def _coord(v):
                s = str(v or "").strip().replace(",", ".")
                try:
                    return float(s)
                except Exception:
                    return None

            cod_cli = str(item_data.get("cod_cliente") or "")
            nome_cli = str(item_data.get("nome") or "")
            pedido = _fmt_pedido(item_data.get("pedido") or "")
            status_pedido = upper(item_data.get("status_pedido") or "")
            cx_orig = safe_int(item_data.get("cx"), 0)
            cx_atual = safe_int(item_data.get("caixas_atual"), 0) or cx_orig
            kg = safe_float(item_data.get("kg"), 0.0)
            preco_orig = safe_float(item_data.get("preco"), 0.0)
            preco_atual = safe_float(item_data.get("preco_atual"), 0.0) or preco_orig
            mortalidade = safe_int(item_data.get("mortalidade"), 0)
            recebido_valor = safe_float(item_data.get("recebido_valor"), 0.0)
            peso_previsto = safe_float(item_data.get("peso_previsto"), 0.0)
            alterado_em = str(item_data.get("alterado_em") or "")
            alterado_por = str(item_data.get("alterado_por") or "")
            forma_receb = str(item_data.get("forma_recebimento") or item_data.get("recebido_forma") or "")
            obs_receb = str(item_data.get("obs_recebimento") or item_data.get("recebido_obs") or "")
            alteracao_tipo = str(item_data.get("alteracao_tipo") or "")
            alteracao_detalhe = str(item_data.get("alteracao_detalhe") or "")
            ordem_sugerida = safe_int(item_data.get("ordem_sugerida"), 0)
            eta = str(item_data.get("eta") or "").strip()
            distancia = item_data.get("distancia")
            confianca_localizacao = item_data.get("confianca_localizacao")
            tem_roteirizacao = bool(
                ordem_sugerida > 0
                or eta
                or (distancia is not None and str(distancia).strip() != "")
                or (confianca_localizacao is not None and str(confianca_localizacao).strip() != "")
            )
            media_cli = (kg / cx_orig) if cx_orig > 0 else 0.0
            valor_orig = float(cx_orig) * float(preco_orig)
            valor_atual = float(cx_atual) * float(preco_atual)
            delta_cx = cx_atual - cx_orig
            delta_preco = preco_atual - preco_orig

            local = _collect_cliente_location(cod_cli)
            transferencias = _collect_transferencias(cod_cli, pedido)
            logs_item = _collect_item_logs(cod_cli, pedido)
            evt_ctx = _extract_event_context(logs_item)
            if not alterado_em:
                alterado_em = str(item_data.get("evento_datahora") or evt_ctx.get("quando", ""))
            for lk in ("latitude", "longitude", "endereco", "cidade", "bairro"):
                if not str(local.get(lk) or "").strip():
                    local[lk] = str(evt_ctx.get(lk) or "").strip()
            if not str(local.get("latitude") or "").strip():
                local["latitude"] = str(item_data.get("lat_evento") or "").strip()
            if not str(local.get("longitude") or "").strip():
                local["longitude"] = str(item_data.get("lon_evento") or "").strip()
            if not str(local.get("endereco") or "").strip():
                local["endereco"] = str(item_data.get("endereco_evento") or "").strip()
            if not str(local.get("cidade") or "").strip():
                local["cidade"] = str(item_data.get("cidade_evento") or "").strip()
            if not str(local.get("bairro") or "").strip():
                local["bairro"] = str(item_data.get("bairro_evento") or "").strip()
            lat = _coord(local.get("latitude"))
            lon = _coord(local.get("longitude"))

            st_norm = upper(status_pedido)
            st_entregue = {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}
            st_cancelado = {"CANCELADO", "CANCELADA"}
            st_em_rota = {"EM_ROTA", "EM ROTA", "INICIADA"}

            def _status_pin_colors(st: str):
                s = upper(st or "")
                if s in st_entregue:
                    return ("#16A34A", "#14532D", "#14532D")
                if s in st_cancelado:
                    return ("#DC2626", "#7F1D1D", "#991B1B")
                if s in st_em_rota:
                    return ("#2563EB", "#1E3A8A", "#1E3A8A")
                if s in {"ALTERADO"}:
                    return ("#F59E0B", "#92400E", "#92400E")
                return ("#6B7280", "#374151", "#374151")

            # Header executivo
            head = ttk.Frame(area, style="Card.TFrame", padding=8)
            head.grid(row=0, column=0, sticky="ew", pady=(0, 8))
            for i in range(4):
                head.grid_columnconfigure(i, weight=1)

            ttk.Label(head, text=f"Cliente: {cod_cli} - {nome_cli}", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
            ttk.Label(head, text=f"Pedido: {pedido or '-'}", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w")
            ttk.Label(head, text=f"Status: {status_pedido or '-'}", font=("Segoe UI", 10, "bold")).grid(row=0, column=3, sticky="w")

            ttk.Label(head, text=f"Programação: {codigo}").grid(row=1, column=0, sticky="w")
            ttk.Label(head, text=f"NF: {cabecalho_ctx.get('nf_numero') or '-'}").grid(row=1, column=1, sticky="w")
            ttk.Label(head, text=f"Motorista: {cabecalho_ctx.get('motorista') or '-'}").grid(row=1, column=2, sticky="w")
            ttk.Label(head, text=f"Veiculo: {cabecalho_ctx.get('veiculo') or '-'}").grid(row=1, column=3, sticky="w")

            nb = ttk.Notebook(area)
            nb.grid(row=1, column=0, sticky="nsew")

            tab_resumo = ttk.Frame(nb, style="Card.TFrame", padding=10)
            tab_hist = ttk.Frame(nb, style="Card.TFrame", padding=10)
            tab_mapa = ttk.Frame(nb, style="Card.TFrame", padding=10)
            nb.add(tab_resumo, text="Resumo")
            nb.add(tab_hist, text="Histórico")
            nb.add(tab_mapa, text="Mapa")

            # TAB RESUMO
            tab_resumo.grid_columnconfigure(0, weight=1)
            tab_resumo.grid_columnconfigure(1, weight=1)
            grp_ent = ttk.LabelFrame(tab_resumo, text="Entrega / Comercial", padding=10)
            grp_ent.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 8))
            grp_rst = ttk.LabelFrame(tab_resumo, text="Rastreabilidade", padding=10)
            grp_rst.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 8))
            grp_geo = ttk.LabelFrame(tab_resumo, text="Localização do Cliente", padding=10)
            grp_geo.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
            grp_check = ttk.LabelFrame(tab_resumo, text="Checklist APK / Sincronia", padding=10)
            grp_check.grid(row=1, column=1, sticky="nsew", padx=(6, 0))

            l_ent = [
                f"Caixas origem: {cx_orig}",
                f"Caixas atual: {cx_atual} (Delta {delta_cx:+d})",
                f"Peso (KG): {kg:.2f}",
                f"Peso previsto (KG): {peso_previsto:.2f}",
                f"Média cliente (KG/CX): {media_cli:.3f}",
                f"Preço origem: {_fmt_money(preco_orig)}",
                f"Preço atual: {_fmt_money(preco_atual)} (Delta {_fmt_money(delta_preco)})",
                f"Valor origem: {_fmt_money(valor_orig)}",
                f"Valor atual: {_fmt_money(valor_atual)}",
                f"Recebido: {_fmt_money(recebido_valor)} | Forma: {forma_receb or '-'}",
                f"Mortalidade (aves): {mortalidade}",
            ]
            for i, t in enumerate(l_ent):
                ttk.Label(grp_ent, text=t).grid(row=i, column=0, sticky="w", pady=1)

            l_rst = [
                f"Local carregado: {cabecalho_ctx.get('local_carreg') or '-'}",
                f"Local rota: {cabecalho_ctx.get('local_rota') or '-'}",
                f"Equipe: {self.resolve_equipe_nomes(cabecalho_ctx.get('equipe') or '') or (cabecalho_ctx.get('equipe') or '-')}",
                f"Hora da entrega/alteração: {alterado_em or '-'}",
                f"Alterado por: {alterado_por or '-'}",
                f"Tipo de alteração: {alteracao_tipo or '-'}",
                f"Detalhe da alteração: {alteracao_detalhe or '-'}",
                f"Observação recebimento: {obs_receb or '-'}",
                f"Transferências encontradas: {len(transferencias)}",
                f"Ocorrências/logs: {len(logs_item)}",
                f"Ordem sugerida: {ordem_sugerida if ordem_sugerida > 0 else '-'}",
                f"ETA previsto: {eta or '-'}",
                f"Distancia estimada: {_fmt_distancia_rota(distancia) or '-'}",
                f"Confianca localizacao: {_fmt_confianca_localizacao(confianca_localizacao) or '-'}",
            ]
            for i, t in enumerate(l_rst):
                ttk.Label(grp_rst, text=t).grid(row=i, column=0, sticky="w", pady=1)

            ttk.Label(grp_geo, text=f"Latitude: {local.get('latitude') or '-'}").grid(row=0, column=0, sticky="w", padx=(0, 20))
            ttk.Label(grp_geo, text=f"Longitude: {local.get('longitude') or '-'}").grid(row=0, column=1, sticky="w")
            ttk.Label(grp_geo, text=f"Endereço: {local.get('endereco') or '-'}").grid(row=1, column=0, columnspan=2, sticky="w")
            ttk.Label(grp_geo, text=f"Cidade: {local.get('cidade') or '-'}").grid(row=2, column=0, sticky="w")
            ttk.Label(grp_geo, text=f"Bairro: {local.get('bairro') or '-'}").grid(row=2, column=1, sticky="w")
            ttk.Label(grp_geo, text=f"Amostras coletadas: {safe_int(local.get('amostras'), 0)}").grid(row=3, column=0, sticky="w", pady=(6, 0))
            ttk.Label(grp_geo, text=f"Ultima leitura: {local.get('ultima_coleta_em') or '-'}").grid(row=3, column=1, sticky="w", pady=(6, 0))
            ttk.Label(grp_geo, text=f"Origem mais recente: {local.get('ultima_origem') or '-'}").grid(row=4, column=0, sticky="w")
            ttk.Label(grp_geo, text=f"Status mais recente: {local.get('ultima_status') or '-'}").grid(row=4, column=1, sticky="w")
            ttk.Label(
                grp_geo,
                text=(
                    f"Ponto recorrente: {local.get('endereco_recorrente') or '-'}"
                    f" | {local.get('cidade_recorrente') or '-'}"
                    f" | {local.get('bairro_recorrente') or '-'}"
                ),
                wraplength=430,
                justify="left",
            ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))
            ttk.Label(
                grp_geo,
                text=(
                    f"Coordenadas recorrentes: {local.get('lat_recorrente') or '-'} / {local.get('lon_recorrente') or '-'}"
                    f"  | Frequencia: {safe_int(local.get('recorrencia_qtd'), 0)}"
                ),
            ).grid(row=6, column=0, columnspan=2, sticky="w")

            checklist = [
                ("Status recebido do APK", bool(status_pedido)),
                ("Hora da alteracao", bool(alterado_em)),
                (
                    "Alteracao do pedido capturada",
                    bool(alteracao_tipo or alteracao_detalhe or delta_cx != 0 or abs(delta_preco) > 0.0001),
                ),
                ("Sugestao de roteirizacao", tem_roteirizacao),
                (
                    "Latitude / Longitude salvas",
                    bool(str(local.get("latitude") or "").strip() and str(local.get("longitude") or "").strip()),
                ),
                (
                    "Endereco / cidade / bairro salvos",
                    bool(
                        str(local.get("endereco") or "").strip()
                        or str(local.get("cidade") or "").strip()
                        or str(local.get("bairro") or "").strip()
                    ),
                ),
                ("Amostras historicas registradas", safe_int(local.get("amostras"), 0) > 0),
            ]
            for i, (label, ok) in enumerate(checklist):
                ttk.Label(grp_check, text=label + ":").grid(row=i, column=0, sticky="w", pady=2)
                ttk.Label(
                    grp_check,
                    text="OK" if ok else "PENDENTE",
                    foreground="#16A34A" if ok else "#B45309",
                    font=("Segoe UI", 9, "bold"),
                ).grid(row=i, column=1, sticky="w", padx=(8, 0), pady=2)

            # TAB HISTORICO
            tab_hist.grid_rowconfigure(1, weight=1)
            tab_hist.grid_columnconfigure(0, weight=1)

            timeline_wrap = ttk.LabelFrame(tab_hist, text="Linha do Tempo", padding=8)
            timeline_wrap.grid(row=0, column=0, sticky="ew", pady=(0, 8))
            tl_canvas = tk.Canvas(
                timeline_wrap,
                height=96,
                bg="#F9FAFB",
                highlightthickness=1,
                highlightbackground="#E5E7EB",
            )
            tl_canvas.pack(fill="x", expand=True)

            step_data = [
                ("Programado", True, f"Prog {codigo}"),
                ("Alterado", bool(alterado_em or delta_cx != 0 or abs(delta_preco) > 0.0001), (alterado_em[:16] if alterado_em else "Sem alteração")),
                ("Em rota", st_norm in st_em_rota, ("Status em rota" if st_norm in st_em_rota else "Não iniciado")),
                ("Entregue", st_norm in st_entregue, ("Pedido entregue" if st_norm in st_entregue else "Pendente")),
                ("Transferido", bool(transferencias), (f"{len(transferencias)} transferência(s)" if transferencias else "Sem transferência")),
            ]

            w_tl = 900
            h_tl = 96
            tl_canvas.config(width=w_tl, height=h_tl)
            x0, x1 = 34, w_tl - 34
            y_line = 42
            if len(step_data) > 1:
                step = (x1 - x0) / (len(step_data) - 1)
            else:
                step = 0
            for idx in range(len(step_data) - 1):
                xi = x0 + idx * step
                xj = x0 + (idx + 1) * step
                seg_color = "#22C55E" if (step_data[idx][1] and step_data[idx + 1][1]) else "#D1D5DB"
                tl_canvas.create_line(xi, y_line, xj, y_line, fill=seg_color, width=3)
            for idx, (nome_etapa, ativo, subt) in enumerate(step_data):
                xi = x0 + idx * step
                fill = "#16A34A" if ativo else "#FFFFFF"
                out = "#14532D" if ativo else "#9CA3AF"
                text_col = "#14532D" if ativo else "#6B7280"
                tl_canvas.create_oval(xi - 8, y_line - 8, xi + 8, y_line + 8, fill=fill, outline=out, width=2)
                tl_canvas.create_text(xi, y_line + 20, text=nome_etapa, fill=text_col, font=("Segoe UI", 9, "bold"))
                tl_canvas.create_text(xi, y_line + 38, text=str(subt)[:26], fill="#6B7280", font=("Segoe UI", 8))

            txt = tk.Text(tab_hist, wrap="word", font=("Consolas", 10))
            txt.grid(row=1, column=0, sticky="nsew")
            y = ttk.Scrollbar(tab_hist, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=y.set)
            y.grid(row=1, column=1, sticky="ns")
            txt.insert("1.0", texto)
            txt.configure(state="disabled")

            # TAB MAPA
            tab_mapa.grid_columnconfigure(0, weight=1)
            tab_mapa.grid_rowconfigure(1, weight=1)
            ttk.Label(tab_mapa, text="Prévia de localização da entrega", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
            map_card = tk.Canvas(tab_mapa, width=520, height=320, bg="#F3F4F6", highlightthickness=1, highlightbackground="#D1D5DB")
            map_card.grid(row=1, column=0, sticky="nsew")
            map_card.create_text(16, 14, anchor="nw", text=f"Cliente: {nome_cli}", fill="#111827", font=("Segoe UI", 10, "bold"))
            map_card.create_text(16, 34, anchor="nw", text=f"Cidade/Bairro: {(local.get('cidade') or '-')} / {(local.get('bairro') or '-')}", fill="#374151", font=("Segoe UI", 9))
            map_card.create_text(16, 52, anchor="nw", text=f"Endereço: {local.get('endereco') or '-'}", fill="#374151", font=("Segoe UI", 9))
            map_card.create_rectangle(20, 82, 500, 300, outline="#CBD5E1", width=1)

            if lat is not None and lon is not None:
                # Enquadramento aproximado BR para mini-preview.
                lat_min, lat_max = -34.0, 5.5
                lon_min, lon_max = -74.0, -32.0
                x_norm = (lon - lon_min) / (lon_max - lon_min)
                y_norm = 1.0 - ((lat - lat_min) / (lat_max - lat_min))
                x_norm = max(0.0, min(1.0, x_norm))
                y_norm = max(0.0, min(1.0, y_norm))
                px = 20 + int(x_norm * (500 - 20))
                py = 82 + int(y_norm * (300 - 82))
                pin_fill, pin_outline, pin_text = _status_pin_colors(status_pedido)
                map_card.create_oval(px - 6, py - 6, px + 6, py + 6, fill=pin_fill, outline=pin_outline)
                map_card.create_text(px + 10, py - 10, anchor="sw", text=f"Entrega ({status_pedido or '-'})", fill=pin_text, font=("Segoe UI", 9, "bold"))
                map_card.create_text(26, 286, anchor="sw", text=f"Lat/Lon: {lat:.6f}, {lon:.6f}", fill="#1F2937", font=("Segoe UI", 9))
            else:
                map_card.create_text(260, 192, text="Sem latitude/longitude no cadastro do cliente", fill="#6B7280", font=("Segoe UI", 10, "bold"))

            links = ttk.Frame(tab_mapa, style="Card.TFrame")
            links.grid(row=2, column=0, sticky="e", pady=(8, 0))

            def _open_osm():
                if lat is None or lon is None:
                    messagebox.showwarning("Mapa", "Latitude/Longitude não disponíveis para este cliente.")
                    return
                webbrowser.open(f"https://www.openstreetmap.org/?mlat={lat}&mlon={lon}#map=16/{lat}/{lon}")

            def _open_google():
                if lat is not None and lon is not None:
                    webbrowser.open(f"https://www.google.com/maps?q={lat},{lon}")
                    return
                q = str(local.get("endereco") or nome_cli or "").strip()
                if not q:
                    messagebox.showwarning("Mapa", "Não há endereço para pesquisar.")
                    return
                webbrowser.open(f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(q)}")

            ttk.Button(links, text="Abrir no OpenStreetMap", style="Ghost.TButton", command=_open_osm).pack(side="left", padx=(0, 6))
            ttk.Button(links, text="Abrir no Google Maps", style="Ghost.TButton", command=_open_google).pack(side="left")

            rodape = ttk.Frame(area)
            rodape.grid(row=2, column=0, sticky="e", pady=(8, 0))
            ttk.Button(rodape, text="Fechar", style="Ghost.TButton", command=w.destroy).pack(side="right")

        tree.bind("<Double-1>", _open_cliente_detalhe, add="+")

        def _load_preview():
            try:
                tree.delete(*tree.get_children())
                detalhes_por_iid.clear()
            except Exception:
                logging.debug("Falha ignorada")

            motorista = ""
            veiculo = ""
            equipe = ""
            status = ""
            nf_numero = ""
            local_rota = ""
            local_carreg = ""
            saida_data = ""
            saida_hora = ""
            data_saida = ""
            hora_saida = ""
            km_inicial = 0.0
            km_final = 0.0

            adiant = 0.0
            nf_kg = 0.0
            nf_kg_carregado = 0.0
            nf_caixas = 0
            nf_saldo = 0.0
            media_carregada = 0.0
            caixa_final = 0
            aves_por_cx = 6
            subst_txt = "-"
            api_rota = None
            api_clientes = None
            fonte_dados = "LOCAL"
            api_erro = ""

            # âÅâ€œââ‚¬¦ helper local pra não quebrar se não existir ainda no arquivo
            def _db_has_column_best_effort(cur, table, col):
                try:
                    if "db_has_column" in globals():
                        return self.db_has_column(cur, table, col)
                except Exception:
                    logging.debug("Falha ignorada")
                try:
                    cur.execute(f"PRAGMA table_info({table})")
                    cols = [str(r[1] or "").lower() for r in cur.fetchall()]
                    return str(col or "").lower() in cols
                except Exception:
                    return False

            # 1) HEADER
            try:
                with get_db() as conn:
                    cur = conn.cursor()

                    cols_chk = [
                        "status",
                        "nf_numero",
                        "local_rota",
                        "tipo_rota",
                        "local_carregamento",
                        "local_carreg",
                        "granja_carregada",
                        "local_carregado",
                        "saida_data",
                        "saida_hora",
                        "data_saida",
                        "hora_saida",
                        "adiantamento",
                        "adiantamento_rota",
                        "kg_carregado",
                        "nf_kg",
                        "nf_kg_carregado",
                        "nf_caixas",
                        "caixas_carregadas",
                        "qnt_cx_carregada",
                        "nf_saldo",
                        "km_inicial",
                        "km_final",
                        "media",
                        "qnt_aves_por_cx",
                        "aves_caixa_final",
                        "qnt_aves_caixa_final",
                    ]
                    existing = {c: _db_has_column_best_effort(cur, "programacoes", c) for c in cols_chk}

                    select_parts = ["motorista", "veiculo", "equipe"]
                    select_parts.append("status" if existing["status"] else "'' as status")
                    select_parts.append("nf_numero" if existing["nf_numero"] else "'' as nf_numero")
                    if existing["local_rota"] and existing["tipo_rota"]:
                        select_parts.append("COALESCE(NULLIF(local_rota, ''), NULLIF(tipo_rota, ''), '') as local_rota")
                    elif existing["local_rota"]:
                        select_parts.append("local_rota")
                    elif existing["tipo_rota"]:
                        select_parts.append("tipo_rota as local_rota")
                    else:
                        select_parts.append("'' as local_rota")
                    if existing["local_carregamento"] and existing["granja_carregada"] and existing["local_carregado"] and existing["local_carreg"]:
                        select_parts.append(
                            "COALESCE(NULLIF(local_carregamento, ''), NULLIF(granja_carregada, ''), NULLIF(local_carregado, ''), NULLIF(local_carreg, ''), '') as local_carregamento"
                        )
                    elif existing["local_carregamento"] and existing["granja_carregada"] and existing["local_carregado"]:
                        select_parts.append(
                            "COALESCE(NULLIF(local_carregamento, ''), NULLIF(granja_carregada, ''), NULLIF(local_carregado, ''), '') as local_carregamento"
                        )
                    elif existing["local_carregamento"] and existing["granja_carregada"]:
                        select_parts.append(
                            "COALESCE(NULLIF(local_carregamento, ''), NULLIF(granja_carregada, ''), '') as local_carregamento"
                        )
                    elif existing["local_carregamento"] and existing["local_carregado"]:
                        select_parts.append(
                            "COALESCE(NULLIF(local_carregamento, ''), NULLIF(local_carregado, ''), '') as local_carregamento"
                        )
                    elif existing["local_carregamento"] and existing["local_carreg"]:
                        select_parts.append(
                            "COALESCE(NULLIF(local_carregamento, ''), NULLIF(local_carreg, ''), '') as local_carregamento"
                        )
                    elif existing["local_carregamento"]:
                        select_parts.append("local_carregamento")
                    elif existing["granja_carregada"] and existing["local_carregado"]:
                        select_parts.append(
                            "COALESCE(NULLIF(granja_carregada, ''), NULLIF(local_carregado, ''), '') as local_carregamento"
                        )
                    elif existing["granja_carregada"] and existing["local_carreg"]:
                        select_parts.append(
                            "COALESCE(NULLIF(granja_carregada, ''), NULLIF(local_carreg, ''), '') as local_carregamento"
                        )
                    elif existing["granja_carregada"]:
                        select_parts.append("granja_carregada as local_carregamento")
                    elif existing["local_carregado"] and existing["local_carreg"]:
                        select_parts.append(
                            "COALESCE(NULLIF(local_carregado, ''), NULLIF(local_carreg, ''), '') as local_carregamento"
                        )
                    elif existing["local_carregado"]:
                        select_parts.append("local_carregado as local_carregamento")
                    elif existing["local_carreg"]:
                        select_parts.append("local_carreg as local_carregamento")
                    else:
                        select_parts.append("'' as local_carregamento")
                    select_parts.append("saida_data" if existing["saida_data"] else "'' as saida_data")
                    select_parts.append("saida_hora" if existing["saida_hora"] else "'' as saida_hora")
                    select_parts.append("data_saida" if existing["data_saida"] else "'' as data_saida")
                    select_parts.append("hora_saida" if existing["hora_saida"] else "'' as hora_saida")
                    select_parts.append("km_inicial" if existing["km_inicial"] else "0 as km_inicial")
                    select_parts.append("km_final" if existing["km_final"] else "0 as km_final")

                    if existing["adiantamento"] and existing["adiantamento_rota"]:
                        select_parts.append(
                            "CASE WHEN COALESCE(adiantamento, 0) <> 0 THEN adiantamento "
                            "ELSE COALESCE(adiantamento_rota, 0) END as adiantamento"
                        )
                    elif existing["adiantamento"]:
                        select_parts.append("adiantamento")
                    elif existing["adiantamento_rota"]:
                        select_parts.append("adiantamento_rota as adiantamento")
                    else:
                        select_parts.append("0 as adiantamento")

                    select_parts.append("COALESCE(nf_kg, 0) as nf_kg" if existing["nf_kg"] else "0 as nf_kg")

                    if existing["nf_kg_carregado"] and existing["kg_carregado"]:
                        select_parts.append(
                            "CASE "
                            "WHEN COALESCE(nf_kg_carregado, 0) > 0 THEN nf_kg_carregado "
                            "WHEN COALESCE(kg_carregado, 0) > 0 THEN kg_carregado "
                            "ELSE COALESCE(nf_kg_carregado, kg_carregado, 0) "
                            "END as nf_kg_carregado"
                        )
                    elif existing["nf_kg_carregado"]:
                        select_parts.append("COALESCE(nf_kg_carregado, 0) as nf_kg_carregado")
                    elif existing["kg_carregado"]:
                        select_parts.append("COALESCE(kg_carregado, 0) as nf_kg_carregado")
                    else:
                        select_parts.append("0 as nf_kg_carregado")
                    if existing["nf_caixas"] and existing["caixas_carregadas"] and existing["qnt_cx_carregada"]:
                        select_parts.append(
                            "CASE "
                            "WHEN COALESCE(nf_caixas, 0) > 0 THEN nf_caixas "
                            "WHEN COALESCE(caixas_carregadas, 0) > 0 THEN caixas_carregadas "
                            "WHEN COALESCE(qnt_cx_carregada, 0) > 0 THEN qnt_cx_carregada "
                            "ELSE 0 END as nf_caixas"
                        )
                    elif existing["nf_caixas"] and existing["caixas_carregadas"]:
                        select_parts.append(
                            "CASE "
                            "WHEN COALESCE(nf_caixas, 0) > 0 THEN nf_caixas "
                            "WHEN COALESCE(caixas_carregadas, 0) > 0 THEN caixas_carregadas "
                            "ELSE 0 END as nf_caixas"
                        )
                    elif existing["nf_caixas"] and existing["qnt_cx_carregada"]:
                        select_parts.append(
                            "CASE "
                            "WHEN COALESCE(nf_caixas, 0) > 0 THEN nf_caixas "
                            "WHEN COALESCE(qnt_cx_carregada, 0) > 0 THEN qnt_cx_carregada "
                            "ELSE 0 END as nf_caixas"
                        )
                    elif existing["caixas_carregadas"] and existing["qnt_cx_carregada"]:
                        select_parts.append(
                            "CASE "
                            "WHEN COALESCE(caixas_carregadas, 0) > 0 THEN caixas_carregadas "
                            "WHEN COALESCE(qnt_cx_carregada, 0) > 0 THEN qnt_cx_carregada "
                            "ELSE 0 END as nf_caixas"
                        )
                    elif existing["nf_caixas"]:
                        select_parts.append("COALESCE(nf_caixas, 0) as nf_caixas")
                    elif existing["caixas_carregadas"]:
                        select_parts.append("COALESCE(caixas_carregadas, 0) as nf_caixas")
                    elif existing["qnt_cx_carregada"]:
                        select_parts.append("COALESCE(qnt_cx_carregada, 0) as nf_caixas")
                    else:
                        select_parts.append("0 as nf_caixas")
                    select_parts.append("nf_saldo" if existing["nf_saldo"] else "0 as nf_saldo")
                    select_parts.append("media" if existing["media"] else "0 as media")
                    select_parts.append("qnt_aves_por_cx" if existing["qnt_aves_por_cx"] else "6 as qnt_aves_por_cx")
                    if existing["aves_caixa_final"] and existing["qnt_aves_caixa_final"]:
                        select_parts.append("COALESCE(aves_caixa_final, qnt_aves_caixa_final, 0) as caixa_final")
                    elif existing["aves_caixa_final"]:
                        select_parts.append("COALESCE(aves_caixa_final, 0) as caixa_final")
                    elif existing["qnt_aves_caixa_final"]:
                        select_parts.append("COALESCE(qnt_aves_caixa_final, 0) as caixa_final")
                    else:
                        select_parts.append("0 as caixa_final")

                    cur.execute(f"""
                        SELECT {", ".join(select_parts)}
                        FROM programacoes
                        WHERE codigo_programacao=?
                        LIMIT 1
                    """, (codigo,))
                    row = cur.fetchone()

                    if row:
                        motorista = row[0] or ""
                        veiculo = row[1] or ""
                        equipe = row[2] or ""
                        status = row[3] or ""
                        nf_numero = row[4] or ""
                        local_rota = row[5] or ""
                        local_carreg = row[6] or ""
                        saida_data = row[7] or ""
                        saida_hora = row[8] or ""
                        data_saida = row[9] or ""
                        hora_saida = row[10] or ""
                        km_inicial = safe_float(row[11], 0.0)
                        km_final = safe_float(row[12], 0.0)
                        adiant = safe_float(row[13], 0.0)
                        nf_kg = safe_float(row[14], 0.0)
                        nf_kg_carregado = safe_float(row[15], 0.0)
                        nf_caixas = safe_int(row[16], 0)
                        nf_saldo = safe_float(row[17], 0.0)
                        media_carregada = safe_float(row[18], 0.0)
                        aves_por_cx = safe_int(row[19], 6) or 6
                        caixa_final = safe_int(row[20], 0)

            except Exception:
                logging.debug("Falha ignorada")

            # 1.2) opcional: fonte central (API) para refletir status/carregamento do app.
            # Mantem a mesma regra da Home: so consulta API quando a integracao Desktop/API esta ativa.
            api_mode = self.is_desktop_api_sync_enabled()
            if api_mode:
                try:
                    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                    if desktop_secret:
                        resposta = _call_api(
                            "GET",
                            f"desktop/rotas/{codigo}",
                            extra_headers={"X-Desktop-Secret": desktop_secret},
                        )
                        if isinstance(resposta, dict):
                            api_rota = resposta.get("rota") if isinstance(resposta.get("rota"), dict) else None
                            api_clientes = resposta.get("clientes") if isinstance(resposta.get("clientes"), list) else None
                            if api_rota is not None:
                                fonte_dados = "API CENTRAL"
                except Exception:
                    api_erro = "API indisponivel"
                    logging.debug("Falha ao puxar preview da API central", exc_info=True)

            if api_rota:
                try:
                    motorista = str(api_rota.get("motorista") or motorista or "")
                    veiculo = str(api_rota.get("veiculo") or veiculo or "")
                    equipe = str(api_rota.get("equipe") or equipe or "")
                    status = str(api_rota.get("status_operacional") or api_rota.get("status") or status or "")
                    nf_numero = str(api_rota.get("nf_numero") or nf_numero or "")
                    local_rota = str(api_rota.get("local_rota") or api_rota.get("tipo_rota") or local_rota or "")
                    local_carreg = str(
                        api_rota.get("local_carregado")
                        or api_rota.get("local_carregamento")
                        or api_rota.get("granja_carregada")
                        or local_carreg
                        or ""
                    )
                    saida_data = str(api_rota.get("saida_data") or saida_data or "")
                    saida_hora = str(api_rota.get("saida_hora") or saida_hora or "")
                    data_saida = str(api_rota.get("data_saida") or data_saida or "")
                    hora_saida = str(api_rota.get("hora_saida") or hora_saida or "")
                    km_inicial = safe_float(api_rota.get("km_inicial"), km_inicial)
                    km_final = safe_float(api_rota.get("km_final"), km_final)
                    adiant = safe_float(api_rota.get("adiantamento"), adiant)
                    nf_kg = safe_float(api_rota.get("nf_kg"), nf_kg)
                    nf_kg_carregado = safe_float(
                        api_rota.get("nf_kg_carregado") if api_rota.get("nf_kg_carregado") is not None else api_rota.get("kg_carregado"),
                        nf_kg_carregado,
                    )
                    nf_caixas = safe_int(
                        api_rota.get("nf_caixas") if api_rota.get("nf_caixas") is not None else api_rota.get("caixas_carregadas"),
                        nf_caixas,
                    )
                    nf_saldo = safe_float(api_rota.get("nf_saldo"), nf_saldo)
                    media_carregada = safe_float(api_rota.get("media"), media_carregada)
                    aves_por_cx = safe_int(api_rota.get("qnt_aves_por_cx"), aves_por_cx) or aves_por_cx
                    caixa_final = safe_int(
                        api_rota.get("caixa_final") if api_rota.get("caixa_final") is not None else api_rota.get("aves_caixa_final"),
                        caixa_final,
                    )
                except Exception:
                    logging.debug("Falha ignorada")

            # 1.1) HISTORICO DE SUBSTITUICAO (se tabela existir)
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='rota_substituicoes'")
                    if cur.fetchone():
                        cur.execute(
                            """
                            SELECT status, origem_motorista_nome, destino_motorista_nome, solicitado_em, aceito_em
                            FROM rota_substituicoes
                            WHERE codigo_programacao=?
                            ORDER BY solicitado_em DESC
                            LIMIT 1
                            """,
                            (codigo,),
                        )
                        rs = cur.fetchone()
                        if rs:
                            st = upper(rs[0] or "")
                            origem = upper(rs[1] or "")
                            destino = upper(rs[2] or "")
                            when = str(rs[4] or rs[3] or "").strip()
                            if st == "PENDENTE_ACEITE":
                                subst_txt = f"Substituicao: EM_TRANSFERENCIA ({origem} -> {destino})"
                            elif st == "ACEITA":
                                subst_txt = f"Substituicao: ACEITA ({origem} -> {destino}) {when}"
                            elif st == "RECUSADA":
                                subst_txt = f"Substituicao: RECUSADA ({origem} -> {destino})"
            except Exception:
                logging.debug("Falha ignorada")

            # 2) ITENS: base local + merge com API (sem perder campos locais)
            itens_local = []
            try:
                if "fetch_programacao_itens" in globals():
                    itens_local = self.fetch_programacao_itens(codigo) or []
            except Exception:
                itens_local = []

            def _item_key(raw):
                if not isinstance(raw, dict):
                    return ("", "", "", 0)
                cod = upper((raw.get("cod_cliente") or "").strip())
                pedido_k = _fmt_pedido(raw.get("pedido", ""))
                nome = upper((raw.get("nome_cliente") or "").strip())
                cx = safe_int(raw.get("qnt_caixas", 0), 0)
                return (cod, pedido_k, nome, cx)

            itens = itens_local
            if isinstance(api_clientes, list):
                try:
                    local_map = {}
                    for it_loc in itens_local:
                        if isinstance(it_loc, dict):
                            local_map[_item_key(it_loc)] = dict(it_loc)
                    merged = []
                    for it_api in api_clientes:
                        if not isinstance(it_api, dict):
                            continue
                        k = _item_key(it_api)
                        base = local_map.pop(k, {})
                        it_m = dict(base)
                        it_m.update(it_api)
                        merged.append(it_m)
                    for it_rest in local_map.values():
                        merged.append(it_rest)
                    if merged:
                        itens = merged
                except Exception:
                    logging.debug("Falha ao mesclar itens locais/API no preview", exc_info=True)

            try:
                itens_indexados = list(enumerate(itens or []))

                def _ordem_preview_sort(pair):
                    idx, raw = pair
                    ordem = safe_int((raw or {}).get("ordem_sugerida"), 0) if isinstance(raw, dict) else 0
                    tem_ordem = ordem > 0
                    return (0 if tem_ordem else 1, ordem if tem_ordem else 10**9, idx)

                itens = [raw for _, raw in sorted(itens_indexados, key=_ordem_preview_sort)]
            except Exception:
                logging.debug("Falha ao ordenar itens por ordem_sugerida no preview", exc_info=True)

            total_cx_programado = 0
            total_cx_entregue = 0
            total_kg = 0.0

            for it in itens:
                cod_cliente = it.get("cod_cliente", "")
                nome = it.get("nome_cliente", "")
                cx = safe_int(it.get("qnt_caixas", 0), 0)
                kg_base = safe_float(it.get("kg", 0.0), 0.0)
                kg_previsto = safe_float(it.get("peso_previsto", 0.0), 0.0)
                media_calc = safe_float(media_carregada, 0.0)
                # Compatibilidade: algumas bases salvam media em gramas (ex.: 3455 -> 3.455)
                if media_calc > 20:
                    media_calc = media_calc / 1000.0
                kg_calc = 0.0
                if media_calc > 0 and cx > 0:
                    kg_calc = float(cx) * float(max(aves_por_cx, 1)) * media_calc

                if kg_previsto > 0:
                    kg = kg_previsto
                elif kg_base > 0:
                    kg = kg_base
                else:
                    kg = kg_calc
                preco = safe_float(it.get("preco", 0.0), 0.0)
                pedido = _fmt_pedido(it.get("pedido", ""))
                status_pedido = upper((it.get("status_pedido", "") or "").strip()) or "PENDENTE"
                caixas_atual = safe_int(it.get("caixas_atual", 0), 0)
                preco_atual = safe_float(it.get("preco_atual", 0.0), 0.0)
                alterado_em = it.get("alterado_em", "") or ""
                recebido_valor = safe_float(it.get("valor_recebido", 0.0), 0.0)
                mortalidade = safe_int(it.get("mortalidade_aves", 0), 0)

                total_cx_programado += cx
                total_kg += kg

                st_norm = upper(status_pedido)
                caixas_ref = caixas_atual if caixas_atual > 0 else cx
                if st_norm in {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}:
                    total_cx_entregue += max(caixas_ref, 0)

                if st_norm in {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO"}:
                    status_tag = "st_entregue"
                elif st_norm in {"PENDENTE", "EM_ROTA", "EM ROTA"}:
                    status_tag = "st_pendente"
                elif st_norm in {"CANCELADO", "CANCELADA"}:
                    status_tag = "st_cancelado"
                elif st_norm in {"ALTERADO"}:
                    status_tag = "st_alterado"
                else:
                    status_tag = "st_default"

                iid = tree_insert_aligned(tree, "", "end", (
                    cod_cliente,
                    nome,
                    cx,
                    f"{kg:.2f}".replace(".", ","),
                    _fmt_money(_normalize_unit_price(preco)),
                    pedido,
                    status_pedido,
                    caixas_atual,
                    _fmt_money(_normalize_unit_price(preco_atual)),
                    alterado_em,
                    _fmt_money(recebido_valor),
                    mortalidade
                ), tags=(status_tag,))
                detalhes_por_iid[iid] = {
                    "cod_cliente": cod_cliente,
                    "nome": nome,
                    "cx": cx,
                    "kg": kg,
                    "preco": _normalize_unit_price(preco),
                    "pedido": pedido,
                    "status_pedido": status_pedido,
                    "caixas_atual": caixas_atual,
                    "preco_atual": _normalize_unit_price(preco_atual),
                    "alterado_em": alterado_em,
                    "alterado_por": it.get("alterado_por", "") if isinstance(it, dict) else "",
                    "recebido_valor": recebido_valor,
                    "mortalidade": mortalidade,
                    "peso_previsto": kg_previsto,
                    "forma_recebimento": (
                        it.get("forma_recebimento", it.get("recebido_forma", "")) if isinstance(it, dict) else ""
                    ),
                    "obs_recebimento": (
                        it.get("obs_recebimento", it.get("recebido_obs", "")) if isinstance(it, dict) else ""
                    ),
                    "alteracao_tipo": it.get("alteracao_tipo", "") if isinstance(it, dict) else "",
                    "alteracao_detalhe": it.get("alteracao_detalhe", "") if isinstance(it, dict) else "",
                    "evento_datahora": it.get("evento_datahora", "") if isinstance(it, dict) else "",
                    "lat_evento": it.get("lat_evento", "") if isinstance(it, dict) else "",
                    "lon_evento": it.get("lon_evento", "") if isinstance(it, dict) else "",
                    "endereco_evento": it.get("endereco_evento", "") if isinstance(it, dict) else "",
                    "cidade_evento": it.get("cidade_evento", "") if isinstance(it, dict) else "",
                    "bairro_evento": it.get("bairro_evento", "") if isinstance(it, dict) else "",
                    "ordem_sugerida": it.get("ordem_sugerida", "") if isinstance(it, dict) else "",
                    "eta": it.get("eta", "") if isinstance(it, dict) else "",
                    "distancia": it.get("distancia", "") if isinstance(it, dict) else "",
                    "confianca_localizacao": it.get("confianca_localizacao", "") if isinstance(it, dict) else "",
                }

            # 3) labels
            lbl_status.config(text=f"Status: {status or ''}")
            lbl_motorista.config(text=f"Motorista: {motorista or ''}")
            equipe_txt = self.resolve_equipe_nomes(equipe)
            lbl_equipe.config(text=f"Equipe: {equipe_txt or ''}")
            lbl_nf.config(text=f"NF: {nf_numero or ''}")
            cabecalho_ctx["nf_numero"] = nf_numero or ""
            cabecalho_ctx["motorista"] = motorista or ""
            cabecalho_ctx["veiculo"] = veiculo or ""
            cabecalho_ctx["equipe"] = equipe or ""
            cabecalho_ctx["local_carreg"] = local_carreg or ""
            cabecalho_ctx["local_rota"] = local_rota or ""

            lbl_adiant.config(text=f"Adiantamento: {_fmt_money(adiant)}")
            lbl_local_rota.config(text=f"Local da Rota: {local_rota or ''}")
            lbl_local_carreg.config(text=f"Carregou em: {local_carreg or ''}")

            saida_txt = "-"
            if saida_data or saida_hora:
                saida_txt = format_date_time(saida_data, saida_hora) or "-"
            elif data_saida or hora_saida:
                saida_txt = format_date_time(data_saida, hora_saida) or "-"
            lbl_saida.config(text=f"Saida: {saida_txt}")

            if nf_kg_carregado <= 0:
                nf_kg_carregado = total_kg
            # Se o campo de carregamento vier vazio, usa o total programado como fallback
            # para não exibir "Caixas Carregadas: 0" no preview.
            if nf_caixas <= 0:
                nf_caixas = total_cx_programado
            if abs(nf_saldo) < 0.0001:
                nf_saldo = round(max(nf_kg - nf_kg_carregado, 0.0), 2)

            lbl_kg_carregado.config(text=f"KG Carregado: {nf_kg_carregado:.2f}".replace(".", ","))
            lbl_caixas_carreg.config(text=f"Caixas Carregadas: {nf_caixas}")
            lbl_caixas_entreg.config(text=f"Caixas Entregues: {total_cx_entregue}")
            lbl_saldo.config(text=f"Saldo (NF - Granja): {nf_saldo:.2f}".replace(".", ","))
            lbl_media_carreg.config(text=f"Media Carregada: {media_carregada:.3f}".replace(".", ","))
            lbl_caixa_final.config(text=f"Caixa Final: {caixa_final if caixa_final > 0 else '-'}")
            lbl_subst.config(text=subst_txt)
            if fonte_dados == "API CENTRAL":
                lbl_fonte.config(text="Fonte: API CENTRAL", fg="#86EFAC")
            else:
                txt_fonte = "Fonte: LOCAL"
                if api_erro:
                    txt_fonte = f"Fonte: LOCAL ({api_erro})"
                lbl_fonte.config(text=txt_fonte, fg="#FBBF24")
            if km_inicial or km_final:
                lbl_km.config(text=f"KM: {km_inicial:.2f} -> {km_final:.2f}".replace(".", ","))
            else:
                lbl_km.config(text="KM: ")

            try:
                if motorista or veiculo:
                    lbl_title.config(text=f"ROTA ATIVA: {codigo}    {upper(motorista)}    {upper(veiculo)}")
            except Exception:
                logging.debug("Falha ignorada")

        _load_preview()

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, sticky="e", pady=(6, 0))
        ttk.Button(btns, text="\U0001F504 ATUALIZAR", style="Ghost.TButton", command=_load_preview).pack(side="right")




__all__ = ["HomePage"]
