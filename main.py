# -*- coding: utf-8 -*-
# ==========================
# ===== INCIO DA PARTE 1 =====
# ==========================
import json
import logging
import os
import re
import sqlite3
import sys
import tempfile
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import shutil
import subprocess
from tkinter import filedialog, messagebox, simpledialog, ttk
try:
    import pandas as pd
except Exception:
    pd = None
from datetime import datetime, timedelta
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
except Exception:
    canvas = None
    A4 = None
import random
import string
import secrets
import ctypes
import time
from app.db.connection import (
    SyncedConnection,
    SyncedCursor,
    _configure_sqlite,
    _is_mutating_sql,
    _is_sql_mirror_enabled,
    _normalize_sql_params,
    _normalize_sql_scalar,
    _push_sql_mutations_to_api,
    configure_connection,
    db_connect,
    get_db,
)
from app.db.migrations import safe_add_column, table_has_column
from db_bootstrap import ensure_permission_system
from app.security.passwords import hash_password_pbkdf2, verify_password_pbkdf2
from app.services.api_client import (
    API_CACHE_TTLS,
    API_GET_CACHE,
    SyncError,
    _api_cache_key,
    _api_cache_ttl,
    _build_api_url,
    _call_api,
    _friendly_sync_error,
    _invalidate_api_cache,
    configure_api_client,
)
from app.services.api_binding import ensure_system_api_binding
from app.services.auth_service import (
    _notify_admin_seed,
    autenticar_usuario,
    ensure_admin_user,
)
from app.services.programacao_service import fetch_programacao_itens as fetch_programacao_itens_service
from app.services.runtime_flags import can_read_from_api, is_desktop_api_sync_enabled
from app.ui.components.cadastro_crud import CadastroCRUD
from app.context import AppContext
from app.ui.components.page_base import PageBase
from app.ui.components.tree_helpers import (
    _align_values_to_tree_columns,
    enable_treeview_sorting,
    tree_insert_aligned,
)
from app.ui.components.entry_helpers import bind_entry_smart
from app.ui.pages.backup_exportar_page import (
    BackupExportarPage,
    configure_backup_exportar_page_dependencies,
)
from app.ui.pages.cadastros_page import CadastrosPage
from app.ui.pages.centro_custos_page import (
    CentroCustosPage,
    configure_centro_custos_page_dependencies,
)
from app.ui.pages.clientes_import_page import (
    ClientesImportPage,
    configure_clientes_import_page_dependencies,
)
from app.ui.pages.home_page import (
    HomePage,
)
from app.ui.pages.permissions_page import (
    PermissionsPage,
)
from app.ui.pages.relatorios_page import (
    RelatoriosPage,
)
from app.ui.pages.system_tools_page import (
    SystemToolsPage,
)
from app.ui.styles import apply_style
from app.utils.async_ui import run_async_ui
from app.utils.excel_helpers import (
    excel_engine_for,
    guess_col,
    require_excel_support,
    require_pandas,
    upper,
)
from app.utils.formatters import (
    fmt_money,
    format_date_br_short,
    format_date_time,
    normalize_date,
    normalize_date_time_components,
    normalize_time,
    now_str,
    safe_float,
    safe_int,
    safe_money,
)
from app.utils.text_fix import _mojibake_score, fix_mojibake_text
from app.utils.validators import (
    normalize_phone,
)
from app.repositories.programacao_repository import db_has_column as repository_db_has_column
from database_runtime import enqueue_sql_statements, log_startup_diagnostics, process_sync_queue, validate_database_identity
from runtime_config import apply_process_environment, ensure_runtime_files, load_app_config

# =========================================================
# CONSTANTES E CONFIGURAÇÕES
# =========================================================
APP_W, APP_H = 1360, 780
DB_NAME = "rota_granja"
APP_CONFIG = load_app_config("desktop")
apply_process_environment(APP_CONFIG)
ensure_runtime_files(APP_CONFIG)

APP_TITLE_DESKTOP = APP_CONFIG.app_title
IS_FROZEN = APP_CONFIG.is_frozen
APP_DIR = APP_CONFIG.app_dir
RESOURCE_DIR = APP_CONFIG.resource_dir
USER_DATA_DIR = APP_CONFIG.data_root
DEFAULT_DB_PATH = APP_CONFIG.db_path
DB_PATH = APP_CONFIG.db_path
API_BASE_URL = APP_CONFIG.api_base_url
API_SYNC_TIMEOUT = APP_CONFIG.api_sync_timeout
APP_ENV = APP_CONFIG.app_env
APP_VERSION = APP_CONFIG.app_version
TENANT_ID = APP_CONFIG.tenant_id
COMPANY_ID = APP_CONFIG.company_id
SYNC_MODE = APP_CONFIG.sync_mode
CONFIG_SOURCE = APP_CONFIG.config_file if os.path.exists(APP_CONFIG.config_file) else APP_CONFIG.config_source
CONFIG_FILE = APP_CONFIG.config_file
UPDATE_MANIFEST_URL = APP_CONFIG.update_manifest_url
SETUP_DOWNLOAD_URL = APP_CONFIG.setup_download_url
CHANGELOG_URL = APP_CONFIG.changelog_url
SUPPORT_WHATSAPP = APP_CONFIG.support_whatsapp
SUPPORT_EMAIL = APP_CONFIG.support_email
LOG_LEVEL = APP_CONFIG.log_level
ENABLE_API_SYNC = APP_CONFIG.sync_enabled
ENABLE_SQL_MIRROR = APP_CONFIG.sql_mirror_api
UPDATE_CHANNEL = APP_CONFIG.update_channel
TENANT_MODE = APP_CONFIG.tenant_mode
ALLOW_SEED_DB = APP_CONFIG.allow_seed_db
ALLOW_REMOTE_WRITE = APP_CONFIG.allow_remote_write
ALLOW_VERSION_UPDATE = APP_CONFIG.allow_version_update
ALLOW_REMOTE_READ = APP_CONFIG.allow_remote_read
SOURCE_OF_TRUTH = APP_CONFIG.source_of_truth
SCHEMA_VERSION = APP_CONFIG.schema_version

configure_api_client(
    api_base_url=API_BASE_URL,
    api_sync_timeout=API_SYNC_TIMEOUT,
    app_env=APP_ENV,
    tenant_id=TENANT_ID,
    company_id=COMPANY_ID,
)

os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)


def enable_high_dpi_awareness():
    """Melhora consistencia visual no Windows com escalas 100/125/150%."""
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return
    except Exception:
        logging.debug("Falha ignorada")
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        logging.debug("Falha ignorada")


def compute_initial_window_size(win) -> tuple[int, int]:
    """Define tamanho inicial proporcional ao monitor atual."""
    try:
        sw = int(win.winfo_screenwidth() or APP_W)
        sh = int(win.winfo_screenheight() or APP_H)
    except Exception:
        return APP_W, APP_H

    width = max(1180, min(int(sw * 0.92), 1760))
    height = max(700, min(int(sh * 0.9), 1040))
    return width, height


def compute_sidebar_width(win) -> int:
    try:
        sw = int(win.winfo_screenwidth() or APP_W)
    except Exception:
        return 280
    if sw <= 1366:
        return 240
    if sw <= 1600:
        return 255
    if sw <= 1920:
        return 270
    return 290


ROUTINE_DEFINITIONS = (
    {
        "page_name": "Home",
        "sidebar_key": "home",
        "code": 100,
        "icon": "\U0001F3E0",
        "nav_label": "Home",
        "title_label": "Home",
    },
    {
        "page_name": "Cadastros",
        "sidebar_key": "cadastros",
        "code": 300,
        "icon": "\U0001F4CB",
        "nav_label": "Cadastros",
        "title_label": "Cadastros",
    },
    {
        "page_name": "Rotas",
        "sidebar_key": "rotas",
        "code": 200,
        "icon": "\U0001F5FA",
        "nav_label": "Rotas",
        "title_label": "Rotas",
    },
    {
        "page_name": "ImportarVendas",
        "sidebar_key": "vendas",
        "code": 400,
        "icon": "\U0001F4E5",
        "nav_label": "Importar Vendas",
        "title_label": "Importar Vendas",
        "aliases": ("Importar Vendas (Excel)",),
    },
    {
        "page_name": "Programacao",
        "sidebar_key": "programacao",
        "code": 500,
        "icon": "\U0001F4C5",
        "nav_label": "Programacao",
        "title_label": "Programacao",
    },
    {
        "page_name": "Recebimentos",
        "sidebar_key": "recebimentos",
        "code": 600,
        "icon": "\U0001F4B5",
        "nav_label": "Recebimentos",
        "title_label": "Recebimentos",
    },
    {
        "page_name": "Despesas",
        "sidebar_key": "despesas",
        "code": 700,
        "icon": "\U0001F4B8",
        "nav_label": "Despesas",
        "title_label": "Despesas",
    },
    {
        "page_name": "Escala",
        "sidebar_key": "escala",
        "code": 800,
        "icon": "\U0001F4CA",
        "nav_label": "Escala",
        "title_label": "Escala",
    },
    {
        "page_name": "CentroCustos",
        "sidebar_key": "centro_custos",
        "code": 900,
        "icon": "\U0001F4C8",
        "nav_label": "Centro de Custos",
        "title_label": "Centro de Custos",
        "aliases": ("Centro de Custos",),
    },
    {
        "page_name": "Relatorios",
        "sidebar_key": "relatorios",
        "code": 1000,
        "icon": "\U0001F4D1",
        "nav_label": "Relatorios",
        "title_label": "Relatorios",
    },
    {
        "page_name": "BackupExportar",
        "sidebar_key": "backup",
        "code": 1100,
        "icon": "\U0001F4E6",
        "nav_label": "Backup / Exportar",
        "title_label": "Backup / Exportar",
        "aliases": ("Backup / Exportar",),
    },
    {
        "page_name": "Permissoes",
        "sidebar_key": "permissoes",
        "code": 1200,
        "icon": "\U0001F512",
        "nav_label": "Gerenciar Permissões",
        "title_label": "Gerenciar Permissões",
        "aliases": ("Gerenciar Permissões",),
    },
    {
        "page_name": "Ferramentas",
        "sidebar_key": "ferramentas",
        "code": 1300,
        "icon": "\U0001F527",
        "nav_label": "Ferramentas do Sistema",
        "title_label": "Ferramentas do Sistema",
        "aliases": ("Ferramentas do Sistema", "Ferramentas Sistema"),
    },
)

# Mapa central das rotinas: altere apenas aqui para renumerar o sistema
# sem mexer nas chaves internas das paginas ou no fluxo atual.
ROUTINE_BY_PAGE = {item["page_name"]: item for item in ROUTINE_DEFINITIONS}
ROUTINE_BY_ALIAS = {}
for _routine_item in ROUTINE_DEFINITIONS:
    _aliases = {
        _routine_item.get("page_name"),
        _routine_item.get("nav_label"),
        _routine_item.get("title_label"),
    }
    _aliases.update(_routine_item.get("aliases") or ())
    for _alias in _aliases:
        _alias_key = str(_alias or "").strip().lower()
        if _alias_key:
            ROUTINE_BY_ALIAS[_alias_key] = _routine_item["page_name"]


def get_routine_meta(name_or_alias):
    raw = str(name_or_alias or "").strip()
    if not raw:
        return None
    meta = ROUTINE_BY_PAGE.get(raw)
    if meta:
        return meta
    resolved = ROUTINE_BY_ALIAS.get(raw.lower())
    return ROUTINE_BY_PAGE.get(resolved) if resolved else None


def get_routine_code(name_or_alias) -> str:
    meta = get_routine_meta(name_or_alias)
    return str(meta.get("code") or "").strip() if meta else ""


def format_routine_nav_label(name_or_alias, include_code: bool = False) -> str:
    meta = get_routine_meta(name_or_alias)
    if not meta:
        return str(name_or_alias or "").strip()
    label = str(meta.get("nav_label") or meta.get("title_label") or meta.get("page_name") or "").strip()
    code = get_routine_code(name_or_alias)
    if include_code and code:
        return f"{code} - {label}"
    return label


def format_routine_title(name_or_alias, include_code: bool = False) -> str:
    meta = get_routine_meta(name_or_alias)
    if not meta:
        return str(name_or_alias or "").strip()
    label = str(meta.get("title_label") or meta.get("nav_label") or meta.get("page_name") or "").strip()
    routine_text = f"Rotina {label}".strip()
    code = get_routine_code(name_or_alias)
    if include_code and code:
        return f"{code} - {routine_text}"
    return routine_text


def get_routine_sidebar_key(name_or_alias):
    meta = get_routine_meta(name_or_alias)
    return str(meta.get("sidebar_key") or "").strip() if meta else ""


def format_subroutine_code(name_or_alias, submenu_index) -> str:
    # Reserva a dezena/centena base para a rotina principal e usa +1..+99
    # para subrotinas/tabs/submenus da mesma area.
    code = get_routine_code(name_or_alias)
    if not code:
        return ""
    try:
        base_code = int(code)
        offset = int(str(submenu_index).strip())
    except Exception:
        return ""
    if offset < 1 or offset > 99:
        return ""
    return str(base_code + offset)


enable_high_dpi_awareness()


if is_desktop_api_sync_enabled() and not os.environ.get("ROTA_SECRET", "").strip():
    logging.warning(
        "Sincronizacao Desktop<->API esta ativa, mas ROTA_SECRET nao foi definido. "
        "Operacoes protegidas da API podem falhar ate a configuracao da chave."
    )

# Log global em DEBUG (pedido do usuario)
logging.basicConfig(
    level=getattr(logging, str(LOG_LEVEL or "DEBUG").upper(), logging.DEBUG),
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logging.info(
    "Runtime carregado | env=%s | tenant=%s | company=%s | sync=%s | db=%s | api=%s | config=%s | source=%s | channel=%s",
    APP_ENV,
    TENANT_ID,
    COMPANY_ID,
    SYNC_MODE,
    DB_PATH,
    API_BASE_URL,
    CONFIG_SOURCE,
    SOURCE_OF_TRUTH,
    UPDATE_CHANNEL,
)


def apply_window_icon(win):
    try:
        if os.name == "nt":
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("rotahub.desktop.app")
    except Exception:
        logging.debug("Falha ignorada")

    base = RESOURCE_DIR
    candidates = [
        os.path.join(base, "assets", "app_icon.ico"),
        os.path.join(base, "app_icon.ico"),
    ]
    for path in candidates:
        if not os.path.exists(path):
            continue
        try:
            win.iconbitmap(path)
            try:
                png_fallback = os.path.join(base, "assets", "app_icon.png")
                if os.path.exists(png_fallback):
                    icon_png = tk.PhotoImage(file=png_fallback)
                    win._app_icon_png = icon_png
                    win.iconphoto(True, icon_png)
            except Exception:
                logging.debug("Falha ignorada")
            return
        except Exception:
            logging.debug("Falha ignorada")


def _wrap_widget_init(widget_cls):
    original = widget_cls.__init__

    def wrapped(self, *args, **kwargs):
        if "text" in kwargs:
            kwargs["text"] = fix_mojibake_text(kwargs.get("text"))
        return original(self, *args, **kwargs)

    widget_cls.__init__ = wrapped


def _install_text_repair_hooks():
    # Labels e botoes (ttk/tk)
    for cls in (ttk.Label, ttk.Button, ttk.Checkbutton, ttk.LabelFrame, ttk.Combobox, tk.Label, tk.Button):
        try:
            _wrap_widget_init(cls)
        except Exception:
            logging.debug("Falha ao instalar hook de texto em widget")

    # valores de combobox
    try:
        _combo_init = ttk.Combobox.__init__

        def wrapped_combo_init(self, *args, **kwargs):
            vals = kwargs.get("values")
            if isinstance(vals, (list, tuple)):
                kwargs["values"] = [fix_mojibake_text(v) for v in vals]
            return _combo_init(self, *args, **kwargs)

        ttk.Combobox.__init__ = wrapped_combo_init
    except Exception:
        logging.debug("Falha ao instalar hook de combobox")

    # Notebook tabs (texto das abas)
    try:
        _nb_add = ttk.Notebook.add
        _nb_insert = ttk.Notebook.insert
        _nb_tab = ttk.Notebook.tab

        def wrapped_nb_add(self, child, **kw):
            if "text" in kw:
                kw["text"] = fix_mojibake_text(kw.get("text"))
            return _nb_add(self, child, **kw)

        def wrapped_nb_insert(self, pos, child, **kw):
            if "text" in kw:
                kw["text"] = fix_mojibake_text(kw.get("text"))
            return _nb_insert(self, pos, child, **kw)

        def wrapped_nb_tab(self, tab_id, option=None, **kw):
            if "text" in kw:
                kw["text"] = fix_mojibake_text(kw.get("text"))
            return _nb_tab(self, tab_id, option, **kw)

        ttk.Notebook.add = wrapped_nb_add
        ttk.Notebook.insert = wrapped_nb_insert
        ttk.Notebook.tab = wrapped_nb_tab
    except Exception:
        logging.debug("Falha ao instalar hook em Notebook")

    # titulo da janela
    try:
        _wm_title = tk.Wm.title

        def wrapped_wm_title(self, string=None):
            if string is not None:
                string = fix_mojibake_text(str(string))
            return _wm_title(self, string)

        tk.Wm.title = wrapped_wm_title
    except Exception:
        logging.debug("Falha ao instalar hook de title")

    # Treeview headings
    try:
        _tree_heading = ttk.Treeview.heading

        def wrapped_heading(self, column, option=None, **kw):
            if "text" in kw:
                kw["text"] = fix_mojibake_text(kw.get("text"))
            return _tree_heading(self, column, option, **kw)

        ttk.Treeview.heading = wrapped_heading
    except Exception:
        logging.debug("Falha ao instalar hook de heading")

    # Atualizacoes dinamicas de texto apos criacao de widget (config/configure)
    try:
        _widget_configure = tk.Widget.configure
        _widget_config = tk.Widget.config

        def wrapped_widget_configure(self, cnf=None, **kw):
            if "text" in kw:
                kw["text"] = fix_mojibake_text(kw.get("text"))
            if "values" in kw and isinstance(kw.get("values"), (list, tuple)):
                kw["values"] = [fix_mojibake_text(v) for v in kw["values"]]
            return _widget_configure(self, cnf, **kw)

        def wrapped_widget_config(self, cnf=None, **kw):
            if "text" in kw:
                kw["text"] = fix_mojibake_text(kw.get("text"))
            if "values" in kw and isinstance(kw.get("values"), (list, tuple)):
                kw["values"] = [fix_mojibake_text(v) for v in kw["values"]]
            return _widget_config(self, cnf, **kw)

        tk.Widget.configure = wrapped_widget_configure
        tk.Widget.config = wrapped_widget_config
    except Exception:
        logging.debug("Falha ao instalar hook global de configure/config")

    # Treeview insert

    # messagebox
    def _wrap_messagebox(func):
        def wrapped(title=None, message=None, *args, **kwargs):
            if title is not None:
                title = fix_mojibake_text(str(title))
            if message is not None:
                message = fix_mojibake_text(str(message))
            return func(title, message, *args, **kwargs)

        return wrapped

    for fn_name in (
        "showinfo",
        "showwarning",
        "showerror",
        "askquestion",
        "askokcancel",
        "askretrycancel",
        "askyesno",
        "askyesnocancel",
    ):
        fn = getattr(messagebox, fn_name, None)
        if callable(fn):
            try:
                setattr(messagebox, fn_name, _wrap_messagebox(fn))
            except Exception:
                logging.debug("Falha ao instalar hook em messagebox")


_install_text_repair_hooks()

# =========================================================
# FUNÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢âââââ€š¬Å¡¬¡ÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢âââââ€š¬Å¡¬¢ES UTILITÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢Ãâââ€š¬Å¡ÂÂÂÂRIAS
# =========================================================
def format_equipe_nomes(ajudante1: str, ajudante2: str, fallback: str = "") -> str:
    a1 = upper(ajudante1)
    a2 = upper(ajudante2)
    nomes = " / ".join([n for n in [a1, a2] if n])
    if nomes:
        return nomes
    return upper(fallback)


def format_ajudante_nome(nome: str, sobrenome: str, fallback: str = "") -> str:
    n = upper(nome)
    s = upper(sobrenome)
    full = " ".join([p for p in [n, s] if p]).strip()
    if full:
        return full
    return upper(fallback)


def resolve_equipe_nomes(equipe_raw: str) -> str:
    raw = str(equipe_raw or "").strip()
    if any(sep in raw for sep in ["|", ",", ";", "/"]):
        parts = [p.strip() for p in re.split(r"[|,;/]+", raw) if str(p or "").strip()]
        nomes = [resolve_equipe_nomes(p) for p in parts]
        nomes = [n for n in nomes if n]
        if nomes:
            return " / ".join(nomes)
    codigo = upper(raw)
    if not codigo:
        return ""
    try:
        with get_db() as conn:
            cur = conn.cursor()

            # ajudantes (novo cadastro): programacoes.equipe guarda o id do ajudante
            try:
                cur.execute(
                    "SELECT nome, sobrenome FROM ajudantes WHERE UPPER(CAST(id AS TEXT))=UPPER(?) LIMIT 1",
                    (codigo,),
                )
                r = cur.fetchone()
                if r:
                    nome = r["nome"] if hasattr(r, "keys") else r[0]
                    sobrenome = r["sobrenome"] if hasattr(r, "keys") else r[1]
                    full = format_ajudante_nome(nome, sobrenome, codigo)
                    if full:
                        return full
            except Exception:
                logging.debug("Falha ignorada")

            # equipes (ajudante1/ajudante2)
            try:
                cur.execute(
                    "SELECT ajudante1, ajudante2 FROM equipes WHERE UPPER(codigo)=UPPER(?) LIMIT 1",
                    (codigo,),
                )
                r = cur.fetchone()
                if r:
                    nomes = format_equipe_nomes(r[0], r[1], codigo)
                    if nomes:
                        return nomes
            except Exception:
                logging.debug("Falha ignorada")

            # equipe_integrantes (quando existir)
            try:
                cur.execute(
                    "SELECT nome FROM equipe_integrantes WHERE equipe_codigo=? ORDER BY nome ASC",
                    (codigo,),
                )
                rows = cur.fetchall() or []
                nomes = [upper(r[0]) for r in rows if r and r[0]]
                if nomes:
                    return " / ".join(nomes)
            except Exception:
                logging.debug("Falha ignorada")

            # Fallbacks antigos
            for sql, params in [
                ("SELECT integrantes FROM equipes WHERE codigo=? LIMIT 1", (codigo,)),
                ("SELECT integrantes FROM equipes WHERE cod_equipe=? LIMIT 1", (codigo,)),
                ("SELECT nomes FROM equipes WHERE codigo=? LIMIT 1", (codigo,)),
                ("SELECT nomes FROM equipes WHERE cod_equipe=? LIMIT 1", (codigo,)),
                ("SELECT nome FROM equipes WHERE codigo=? LIMIT 1", (codigo,)),
                ("SELECT nome FROM equipes WHERE cod_equipe=? LIMIT 1", (codigo,)),
            ]:
                try:
                    cur.execute(sql, params)
                    r = cur.fetchone()
                    if r and r[0]:
                        return upper(r[0])
                except Exception:
                    logging.debug("Falha ignorada")
    except Exception:
        logging.debug("Falha ignorada")

    return codigo

configure_connection(
    DB_PATH,
    call_api=_call_api,
    is_sync_enabled=is_desktop_api_sync_enabled,
)


def normalize_datetime_column(series):
    """Normaliza colunas datetime/DATE em DataFrame para YYYY-MM-DD HH:MM."""
    try:
        import pandas as _pd
        ser = series.astype(str)
        dt = _pd.to_datetime(ser, errors="coerce", dayfirst=True)
        return dt.dt.strftime("%Y-%m-%d %H:%M").fillna("")
    except Exception:
        return series

def normalize_date_column(series):
    """Normaliza colunas de data em DataFrame para YYYY-MM-DD."""
    try:
        import pandas as _pd
        ser = series.astype(str)
        dt = _pd.to_datetime(ser, errors="coerce", dayfirst=True)
        return dt.dt.strftime("%Y-%m-%d").fillna("")
    except Exception:
        return series

def require_openpyxl() -> bool:
    try:
        import openpyxl  # noqa: F401
        return True
    except Exception:
        try:
            messagebox.showerror(
                "ERRO",
                "Exportação Excel requer o pacote 'openpyxl'.\n\n"
                "Instale com: pip install openpyxl"
            )
        except Exception:
            logging.debug("Falha ignorada")
        return False

def require_xlrd() -> bool:
    try:
        import xlrd  # noqa: F401  # type: ignore[import-not-found]
        return True
    except Exception:
        try:
            messagebox.showerror(
                "ERRO",
                "Importação de .xls requer o pacote 'xlrd'.\n\n"
                "Instale com: pip install xlrd"
            )
        except Exception:
            logging.debug("Falha ignorada")
        return False


def require_reportlab() -> bool:
    if canvas is None or A4 is None:
        try:
            messagebox.showerror(
                "ERRO",
                "Geração de PDF requer o pacote 'reportlab'.\n\n"
                "Instale com: pip install reportlab"
            )
        except Exception:
            logging.debug("Falha ignorada")
        return False
    return True

def generate_program_code():
    """Gera um código curto para programação (PG + YYYY + sequência há 2 dÃÂÂgitos)."""
    now = datetime.now()
    prefix = f"PG{now.strftime('%Y')}"
    suffix = 1

    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT codigo_programacao FROM programacoes WHERE codigo_programacao LIKE ? ORDER BY codigo_programacao DESC LIMIT 1",
                (f"{prefix}%",),
            )
            row = cur.fetchone()
            if row:
                tail = (row[0] or "").strip()
                tail = tail[len(prefix) :]
                digits = "".join(ch for ch in tail if ch.isdigit())
                if digits:
                    suffix = int(digits) + 1
    except Exception:
        logging.debug("Falha ao calcular próximo código de programação")

    return f"{prefix}{suffix:02d}"

def generate_motorista_code(motorista_id: int = None) -> str:
    """Gera um código simples e único pro motorista (ex.: MOT000123 / MOTAB12CD)."""
    base = "MOT"
    if motorista_id is not None:
        try:
            return f"{base}{int(motorista_id):06d}"
        except Exception:
            logging.debug("Falha ignorada")
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"{base}{suffix}"

def generate_usuario_code(usuario_id: int = None) -> str:
    """Gera um código simples e único pro usuário (ex.: USR000123 / USRAB12CD)."""
    base = "USR"
    if usuario_id is not None:
        try:
            return f"{base}{int(usuario_id):06d}"
        except Exception:
            logging.debug("Falha ignorada")
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"{base}{suffix}"

def generate_motorista_password(length: int = 6) -> str:
    """Senha numérica simples (pra bater com a API atual)."""
    try:
        length = int(length)
    except Exception:
        length = 6
    if length < 4:
        length = 4
    return "".join(random.choices("0123456789", k=length))

def generate_usuario_password(length: int = 6) -> str:
    """Senha numérica simples pro usuário."""
    return generate_motorista_password(length)

# =========================================================
# MIGRAÇÃO DE BANCO DE DADOS
# =========================================================
def db_init():
    """Inicializa/atualiza banco de dados com migrações"""
    with get_db() as conn:
        cur = conn.cursor()

        # MOTORISTAS
        cur.execute("""
            CREATE TABLE IF NOT EXISTS motoristas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT,
                cpf TEXT,
                telefone TEXT
            )
        """)
        safe_add_column(cur, "motoristas", "codigo", "TEXT")
        safe_add_column(cur, "motoristas", "senha", "TEXT")
        safe_add_column(cur, "motoristas", "status", "TEXT DEFAULT 'ATIVO'")
        # Novos motoristas devem entrar no fluxo mobile sem bloqueio manual.
        safe_add_column(cur, "motoristas", "acesso_liberado", "INTEGER DEFAULT 1")
        safe_add_column(cur, "motoristas", "acesso_liberado_por", "TEXT")
        safe_add_column(cur, "motoristas", "acesso_liberado_em", "TEXT")
        safe_add_column(cur, "motoristas", "acesso_obs", "TEXT")
        try:
            cur.execute("""
                UPDATE motoristas
                SET status='ATIVO'
                WHERE status IS NULL OR TRIM(status)=''
            """)
        except Exception:
            logging.debug("Falha ignorada")

        # VENDEDORES (login do app vendedor)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS vendedores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE,
                nome TEXT,
                telefone TEXT,
                cidade_base TEXT,
                status TEXT DEFAULT 'ATIVO',
                senha TEXT,
                ultimo_login_em TEXT,
                ultimo_login_ip TEXT
            )
        """)
        safe_add_column(cur, "vendedores", "codigo", "TEXT")
        safe_add_column(cur, "vendedores", "nome", "TEXT")
        safe_add_column(cur, "vendedores", "telefone", "TEXT")
        safe_add_column(cur, "vendedores", "cidade_base", "TEXT")
        safe_add_column(cur, "vendedores", "status", "TEXT DEFAULT 'ATIVO'")
        safe_add_column(cur, "vendedores", "senha", "TEXT")
        safe_add_column(cur, "vendedores", "ultimo_login_em", "TEXT")
        safe_add_column(cur, "vendedores", "ultimo_login_ip", "TEXT")
        try:
            cur.execute("""
                UPDATE vendedores
                SET status='ATIVO'
                WHERE status IS NULL OR TRIM(status)=''
            """)
        except Exception:
            logging.debug("Falha ignorada")
        try:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_vendedores_codigo ON vendedores(codigo)")
        except Exception as e:
            logging.exception("Falha ao criar indice de vendedores (codigo): %s", e)

        # USURIOS
        cur.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT,
                permissoes TEXT,
                cpf TEXT,
                
                telefone TEXT
            )
        """)
        # âÅâ€œââ‚¬¦ ADICIONADO (pra login + gerar senha/codigo)
        safe_add_column(cur, "usuarios", "codigo", "TEXT")
        safe_add_column(cur, "usuarios", "senha", "TEXT")

        # PERMISSÕES DO SISTEMA - Catálogo
        cur.execute("""
            CREATE TABLE IF NOT EXISTS permissoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                modulo TEXT NOT NULL,
                nome_permissao TEXT NOT NULL,
                descricao TEXT,
                ativo INTEGER DEFAULT 1
            )
        """)
        
        # USUÁRIO-PERMISSÕES - Associação
        cur.execute("""
            CREATE TABLE IF NOT EXISTS usuario_permissoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id INTEGER NOT NULL,
                permissao_id INTEGER NOT NULL,
                concedida_em TEXT DEFAULT (datetime('now')),
                concedida_por TEXT,
                UNIQUE(usuario_id, permissao_id),
                FOREIGN KEY(usuario_id) REFERENCES usuarios(id),
                FOREIGN KEY(permissao_id) REFERENCES permissoes(id)
            )
        """)
        
        # LOGS DO SISTEMA - Ferramentas e operações críticas
        cur.execute("""
            CREATE TABLE IF NOT EXISTS sistema_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo_acao TEXT,
                descricao TEXT,
                usuario TEXT,
                status TEXT,
                resultado_texto TEXT,
                executado_em TEXT DEFAULT (datetime('now'))
            )
        """)

        # VECULOS
        cur.execute("""
            CREATE TABLE IF NOT EXISTS veiculos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                placa TEXT,
                modelo TEXT,
                
                capacidade_cx INTEGER
            )
        """)
        safe_add_column(cur, "veiculos", "capacidade_cx", "INTEGER")
        # Migração: se banco antigo usava capacidade_c, copiar para capacidade_cx
        try:
            if table_has_column(cur, "veiculos", "capacidade_c") and table_has_column(cur, "veiculos", "capacidade_cx"):
                cur.execute("""
                    UPDATE veiculos
                    SET capacidade_cx = CAST(capacidade_c AS INTEGER)
                    WHERE (capacidade_cx IS NULL OR capacidade_cx = '')
                      AND capacidade_c IS NOT NULL
                      AND capacidade_c <> ''
                """)
        except Exception as e:
            logging.exception("Falha ao migrar capacidade_c -> capacidade_cx: %s", e)

        # EQUIPES
        cur.execute("""
            CREATE TABLE IF NOT EXISTS equipes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT,
                ajudante1 TEXT,
                ajudante2 TEXT
            )
        """)

        # AJUDANTES (novo cadastro)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ajudantes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT,
                sobrenome TEXT,
                telefone TEXT,
                status TEXT DEFAULT 'ATIVO'
            )
        """)
        safe_add_column(cur, "ajudantes", "nome", "TEXT")
        safe_add_column(cur, "ajudantes", "sobrenome", "TEXT")
        safe_add_column(cur, "ajudantes", "telefone", "TEXT")
        safe_add_column(cur, "ajudantes", "status", "TEXT DEFAULT 'ATIVO'")
        try:
            cur.execute("""
                UPDATE ajudantes
                SET status='ATIVO'
                WHERE status IS NULL OR TRIM(status)=''
            """)
        except Exception:
            logging.debug("Falha ignorada")

        # Migra base legada de equipes -> ajudantes (melhor esforço)
        try:
            cur.execute("SELECT ajudante1, ajudante2 FROM equipes")
            for row in cur.fetchall() or []:
                for raw_nome in [row[0] if row else "", row[1] if row else ""]:
                    nome_full = upper(raw_nome)
                    if not nome_full:
                        continue
                    parts = nome_full.split()
                    nome = parts[0] if parts else nome_full
                    sobrenome = " ".join(parts[1:]) if len(parts) > 1 else ""
                    cur.execute(
                        """
                        SELECT 1 FROM ajudantes
                        WHERE UPPER(nome)=UPPER(?) AND UPPER(COALESCE(sobrenome,''))=UPPER(?)
                        LIMIT 1
                        """,
                        (nome, sobrenome),
                    )
                    if cur.fetchone():
                        continue
                    cur.execute(
                        "INSERT INTO ajudantes (nome, sobrenome, telefone) VALUES (?, ?, ?)",
                        (nome, sobrenome, ""),
                    )
        except Exception:
            logging.debug("Falha ignorada")

        
        # CLIENTES
        cur.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cod_cliente TEXT UNIQUE,
                nome_cliente TEXT,
                endereco TEXT,
                bairro TEXT,
                cidade TEXT,
                uf TEXT,
                telefone TEXT,
                rota TEXT
            )
        """)
        safe_add_column(cur, "clientes", "cod_cliente", "TEXT")
        safe_add_column(cur, "clientes", "nome_cliente", "TEXT")
        safe_add_column(cur, "clientes", "endereco", "TEXT")
        safe_add_column(cur, "clientes", "bairro", "TEXT")
        safe_add_column(cur, "clientes", "cidade", "TEXT")
        safe_add_column(cur, "clientes", "uf", "TEXT")
        safe_add_column(cur, "clientes", "telefone", "TEXT")
        safe_add_column(cur, "clientes", "rota", "TEXT")
        safe_add_column(cur, "clientes", "vendedor", "TEXT")
        safe_add_column(cur, "clientes", "latitude", "REAL")
        safe_add_column(cur, "clientes", "longitude", "REAL")

        try:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_clientes_cod ON clientes(cod_cliente)")
        except Exception as e:
            logging.exception("Falha ao criar indice de clientes (cod_cliente): %s", e)

        # RASCUNHO COMPARTILHADO DO APP VENDEDOR
        cur.execute("""
            CREATE TABLE IF NOT EXISTS vendedor_rascunho_itens (
                id TEXT PRIMARY KEY,
                cod_cliente TEXT NOT NULL,
                nome_cliente TEXT NOT NULL,
                cidade TEXT,
                bairro TEXT,
                endereco TEXT,
                vendedor_cadastro TEXT,
                vendedor_origem TEXT NOT NULL,
                preco REAL DEFAULT 0,
                caixas INTEGER DEFAULT 0,
                status TEXT DEFAULT 'PENDENTE',
                observacao TEXT DEFAULT '',
                alerta_codigo_programacao TEXT,
                alerta_status_rota TEXT,
                criado_em TEXT,
                atualizado_em TEXT,
                criado_por_codigo TEXT,
                atualizado_por_codigo TEXT
            )
        """)
        safe_add_column(cur, "vendedor_rascunho_itens", "cod_cliente", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "nome_cliente", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "cidade", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "bairro", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "endereco", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "vendedor_cadastro", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "vendedor_origem", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "preco", "REAL DEFAULT 0")
        safe_add_column(cur, "vendedor_rascunho_itens", "caixas", "INTEGER DEFAULT 0")
        safe_add_column(cur, "vendedor_rascunho_itens", "status", "TEXT DEFAULT 'PENDENTE'")
        safe_add_column(cur, "vendedor_rascunho_itens", "observacao", "TEXT DEFAULT ''")
        safe_add_column(cur, "vendedor_rascunho_itens", "alerta_codigo_programacao", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "alerta_status_rota", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "criado_em", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "atualizado_em", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "criado_por_codigo", "TEXT")
        safe_add_column(cur, "vendedor_rascunho_itens", "atualizado_por_codigo", "TEXT")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS vendedor_pre_programacoes (
                id TEXT PRIMARY KEY,
                titulo TEXT NOT NULL,
                observacao TEXT DEFAULT '',
                status TEXT DEFAULT 'ABERTA',
                criado_em TEXT,
                atualizado_em TEXT,
                criado_por_codigo TEXT,
                atualizado_por_codigo TEXT
            )
        """)
        safe_add_column(cur, "vendedor_pre_programacoes", "titulo", "TEXT")
        safe_add_column(cur, "vendedor_pre_programacoes", "observacao", "TEXT DEFAULT ''")
        safe_add_column(cur, "vendedor_pre_programacoes", "status", "TEXT DEFAULT 'ABERTA'")
        safe_add_column(cur, "vendedor_pre_programacoes", "criado_em", "TEXT")
        safe_add_column(cur, "vendedor_pre_programacoes", "atualizado_em", "TEXT")
        safe_add_column(cur, "vendedor_pre_programacoes", "criado_por_codigo", "TEXT")
        safe_add_column(cur, "vendedor_pre_programacoes", "atualizado_por_codigo", "TEXT")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS vendedor_pre_programacao_itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pre_programacao_id TEXT NOT NULL,
                rascunho_item_id TEXT NOT NULL,
                ordem INTEGER DEFAULT 0,
                criado_em TEXT,
                atualizado_em TEXT
            )
        """)
        safe_add_column(cur, "vendedor_pre_programacao_itens", "pre_programacao_id", "TEXT")
        safe_add_column(cur, "vendedor_pre_programacao_itens", "rascunho_item_id", "TEXT")
        safe_add_column(cur, "vendedor_pre_programacao_itens", "ordem", "INTEGER DEFAULT 0")
        safe_add_column(cur, "vendedor_pre_programacao_itens", "criado_em", "TEXT")
        safe_add_column(cur, "vendedor_pre_programacao_itens", "atualizado_em", "TEXT")
        try:
            cur.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_pre_programacao_item_unique "
                "ON vendedor_pre_programacao_itens(pre_programacao_id, rascunho_item_id)"
            )
        except Exception:
            pass
        try:
            cur.execute(
                "CREATE INDEX IF NOT EXISTS idx_pre_programacao_item_ordem "
                "ON vendedor_pre_programacao_itens(pre_programacao_id, ordem, id)"
            )
        except Exception:
            pass

        # VENDAS IMPORTADAS
        cur.execute("""
            CREATE TABLE IF NOT EXISTS vendas_importadas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido TEXT,
                data_venda TEXT,
                cliente TEXT,
                nome_cliente TEXT,
                vendedor TEXT,
                produto TEXT,
                vr_total REAL,
                qnt REAL,
                cidade TEXT,
                valor_unitario REAL,
                observacao TEXT,
                selecionada INTEGER DEFAULT 0
            )
        """)
        safe_add_column(cur, "vendas_importadas", "selecionada", "INTEGER DEFAULT 0")
        safe_add_column(cur, "vendas_importadas", "data_venda", "TEXT")
        safe_add_column(cur, "vendas_importadas", "vr_total", "REAL")
        safe_add_column(cur, "vendas_importadas", "qnt", "REAL")
        safe_add_column(cur, "vendas_importadas", "cidade", "TEXT")

        # PROGRAMAÇÕES
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT,
                data_criacao TEXT,
                motorista TEXT,
                veiculo TEXT,
                equipe TEXT,
                kg_estimado REAL
            )
        """)
        safe_add_column(cur, "programacoes", "codigo_programacao", "TEXT")
        safe_add_column(cur, "programacoes", "data_criacao", "TEXT")
        safe_add_column(cur, "programacoes", "motorista", "TEXT")
        safe_add_column(cur, "programacoes", "veiculo", "TEXT")
        safe_add_column(cur, "programacoes", "equipe", "TEXT")
        safe_add_column(cur, "programacoes", "kg_estimado", "REAL")
        # Regra FOB/CIF: estimativa pode ser por KG (CIF) ou por CX (FOB)
        safe_add_column(cur, "programacoes", "tipo_estimativa", "TEXT DEFAULT 'KG'")
        safe_add_column(cur, "programacoes", "caixas_estimado", "INTEGER DEFAULT 0")
        # Auditoria de criação/edição da programação
        safe_add_column(cur, "programacoes", "usuario_criacao", "TEXT")
        safe_add_column(cur, "programacoes", "usuario_ultima_edicao", "TEXT")

        try:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_programacoes_codigo ON programacoes(codigo_programacao)")
        except Exception as e:
            logging.exception("Falha ao criar indice de programacoes (codigo_programacao): %s", e)

        # NOVAS COLUNAS PARA STATUS E ROTAS ATIVAS
        safe_add_column(cur, "programacoes", "status", "TEXT DEFAULT 'ATIVA'")
        safe_add_column(cur, "programacoes", "status_operacional", "TEXT")
        safe_add_column(cur, "programacoes", "finalizada_no_app", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "prestacao_status", "TEXT DEFAULT 'PENDENTE'")
        safe_add_column(cur, "programacoes", "tipo_rota", "TEXT")
        safe_add_column(cur, "programacoes", "granja_carregada", "TEXT")
        safe_add_column(cur, "programacoes", "data_saida", "TEXT")
        safe_add_column(cur, "programacoes", "hora_saida", "TEXT")
        safe_add_column(cur, "programacoes", "data_chegada", "TEXT")
        safe_add_column(cur, "programacoes", "hora_chegada", "TEXT")
        safe_add_column(cur, "programacoes", "diaria_motorista_valor", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "adiantamento", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "num_nf", "TEXT")
        safe_add_column(cur, "programacoes", "kg_carregado", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "media", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "qnt_aves_caixa_final", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "aves_caixa_final", "INTEGER DEFAULT 0")

        # âÅâ€œââ‚¬¦ SUA TELA DESPESAS USA "adiantamento_rota" (ADICIONADO)
        safe_add_column(cur, "programacoes", "adiantamento_rota", "REAL DEFAULT 0")

        # NOVAS MIGRAÇÕES (NF / KM / CÉDULAS)
        safe_add_column(cur, "programacoes", "nf_numero", "TEXT")
        safe_add_column(cur, "programacoes", "nf_kg", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_preco", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_caixas", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_kg_carregado", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_kg_vendido", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_saldo", "REAL DEFAULT 0")
        # Ocorrências do transbordo (mortalidade) - impactam o KG útil da NF.
        safe_add_column(cur, "programacoes", "mortalidade_transbordo_aves", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "mortalidade_transbordo_kg", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "obs_transbordo", "TEXT")

        safe_add_column(cur, "programacoes", "km_inicial", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "km_final", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "litros", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "km_rodado", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "media_km_l", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "custo_km", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "rota_observacao", "TEXT")

        safe_add_column(cur, "programacoes", "ced_200_qtd", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "ced_100_qtd", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "ced_50_qtd", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "ced_20_qtd", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "ced_10_qtd", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "ced_5_qtd", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "ced_2_qtd", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "valor_dinheiro", "REAL DEFAULT 0")

        # garante prestação pendente em bases antigas (não sobrescreve FECHADA)
        try:
            if table_has_column(cur, "programacoes", "prestacao_status"):
                cur.execute("""
                    UPDATE programacoes
                    SET prestacao_status='PENDENTE'
                    WHERE (prestacao_status IS NULL OR prestacao_status='')
                """)
        except Exception as e:
            logging.exception("Falha ao garantir prestacao_status pendente: %s", e)

        # ITENS DA PROGRAMAÇÃO
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacao_itens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT,
                cod_cliente TEXT,
                nome_cliente TEXT,
                qnt_caixas INTEGER,
                kg REAL,
                preco REAL,
                endereco TEXT,
                vendedor TEXT,
                pedido TEXT,
                produto TEXT
            )
        """)
        safe_add_column(cur, "programacao_itens", "endereco", "TEXT")
        safe_add_column(cur, "programacao_itens", "pedido", "TEXT")
        safe_add_column(cur, "programacao_itens", "vendedor", "TEXT")
        safe_add_column(cur, "programacao_itens", "produto", "TEXT")
        safe_add_column(cur, "programacao_itens", "observacao", "TEXT")
        safe_add_column(cur, "programacao_itens", "status_pedido", "TEXT")
        safe_add_column(cur, "programacao_itens", "caixas_atual", "INTEGER")
        safe_add_column(cur, "programacao_itens", "preco_atual", "REAL")
        safe_add_column(cur, "programacao_itens", "alterado_em", "TEXT")
        safe_add_column(cur, "programacao_itens", "alterado_por", "TEXT")

        # CONTROLE/LOG (sincronização app)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacao_itens_controle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT,
                cod_cliente TEXT,
                status_pedido TEXT,
                caixas_atual INTEGER,
                preco_atual REAL,
                alterado_em TEXT,
                alterado_por TEXT,
                mortalidade_aves INTEGER DEFAULT 0,
                peso_previsto REAL DEFAULT 0,
                valor_recebido REAL DEFAULT 0,
                forma_recebimento TEXT,
                obs_recebimento TEXT,
                lat_evento REAL,
                lon_evento REAL,
                endereco_evento TEXT,
                cidade_evento TEXT,
                bairro_evento TEXT
            )
        """)
        safe_add_column(cur, "programacao_itens_controle", "peso_previsto", "REAL DEFAULT 0")
        safe_add_column(cur, "programacao_itens_controle", "lat_evento", "REAL")
        safe_add_column(cur, "programacao_itens_controle", "lon_evento", "REAL")
        safe_add_column(cur, "programacao_itens_controle", "endereco_evento", "TEXT")
        safe_add_column(cur, "programacao_itens_controle", "cidade_evento", "TEXT")
        safe_add_column(cur, "programacao_itens_controle", "bairro_evento", "TEXT")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacao_itens_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT,
                cod_cliente TEXT,
                payload_json TEXT,
                registrado_em TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS cliente_localizacao_amostras (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cod_cliente TEXT NOT NULL,
                codigo_programacao TEXT,
                pedido TEXT,
                latitude REAL,
                longitude REAL,
                endereco TEXT,
                cidade TEXT,
                bairro TEXT,
                status_pedido TEXT,
                motorista TEXT,
                origem TEXT DEFAULT 'APP',
                registrado_em TEXT DEFAULT (datetime('now'))
            )
        """)
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_cli_loc_amostras_cliente ON cliente_localizacao_amostras(cod_cliente, registrado_em DESC)")
        except Exception as e:
            logging.exception("Falha ao criar indice cliente_localizacao_amostras(cod_cliente): %s", e)

        # RECEBIMENTOS
        cur.execute("""
            CREATE TABLE IF NOT EXISTS recebimentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT,
                cod_cliente TEXT,
                nome_cliente TEXT,
                valor REAL,
                forma_pagamento TEXT,
                observacao TEXT,
                num_nf TEXT,
                data_registro TEXT
            )
        """)

        # DESPESAS
        cur.execute("""
            CREATE TABLE IF NOT EXISTS despesas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT,
                descricao TEXT,
                valor REAL,
                data_registro TEXT
            )
        """)

        # MIGRAÇÕES PARA RELATÓRIOS
        safe_add_column(cur, "despesas", "tipo_despesa", "TEXT DEFAULT 'ROTA'")
        safe_add_column(cur, "despesas", "categoria", "TEXT")
        safe_add_column(cur, "despesas", "motorista", "TEXT")
        safe_add_column(cur, "despesas", "veiculo", "TEXT")
        safe_add_column(cur, "despesas", "observacao", "TEXT")  # âÅâ€œââ‚¬¦ sua tela usa isso

        # NDICES (performance)
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_despesas_programacao ON despesas(codigo_programacao)")
        except Exception as e:
            logging.exception("Falha ao criar indice despesas(codigo_programacao): %s", e)
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_receb_programacao ON recebimentos(codigo_programacao)")
        except Exception as e:
            logging.exception("Falha ao criar indice recebimentos(codigo_programacao): %s", e)
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_receb_cliente ON recebimentos(cod_cliente)")
        except Exception as e:
            logging.exception("Falha ao criar indice recebimentos(cod_cliente): %s", e)
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_prog_itens_programacao ON programacao_itens(codigo_programacao)")
        except Exception as e:
            logging.exception("Falha ao criar indice programacao_itens(codigo_programacao): %s", e)
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_prog_itens_ctrl_prog_cliente ON programacao_itens_controle(codigo_programacao, cod_cliente)")
        except Exception as e:
            logging.exception("Falha ao criar indice programacao_itens_controle(codigo_programacao, cod_cliente): %s", e)
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_prog_itens_log_prog ON programacao_itens_log(codigo_programacao)")
        except Exception as e:
            logging.exception("Falha ao criar indice programacao_itens_log(codigo_programacao): %s", e)

    # âÅâ€œââ‚¬¦ garante ADMIN (senha segura e temporaria)
    ensure_admin_user()
    
    # âÅâ€œââ‚¬¦ inicializa sistema de permissões
    try:
        with get_db() as conn:
            ensure_permission_system(conn)
    except Exception as e:
        logging.warning("Erro ao inicializar sistema de permissões: %s", e)

# =========================================================
# COMPONENTE CRUD GENÉRICO (mesma lógica, layout mais robusto)
# =========================================================
# ==========================
# ===== CADASTRO CRUD (ATUALIZADO) =====
# ==========================

class Sidebar(ttk.Frame):
    def __init__(self, parent, app):
        sidebar_width = compute_sidebar_width(parent)
        super().__init__(parent, style="Sidebar.TFrame", width=sidebar_width)
        self.app = app
        self.pack_propagate(False)
        self.configure(width=sidebar_width)

        self.buttons = {}

        # Topo
        top = ttk.Frame(self, style="Sidebar.TFrame", padding=(18, 20))
        top.pack(fill="x")

        ttk.Label(top, text="ROTAHUB DESKTOP", style="SidebarLogo.TLabel").pack(anchor="w")
        ttk.Label(top, text="Centralizando sua operacao do inicio ao fim.", style="SidebarSmall.TLabel").pack(anchor="w", pady=(2, 0))

        ttk.Separator(self).pack(fill="x", padx=12, pady=(8, 12))

        # âÅâ€œââ‚¬¦ Menu com scroll (evita quebra/sumir item em telas pequenas)
        wrap = ttk.Frame(self, style="Sidebar.TFrame")
        wrap.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(wrap, highlightthickness=0, bd=0, bg="#2B2F8F")
        self.canvas.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        vsb.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=vsb.set)

        self.menu = ttk.Frame(self.canvas, style="Sidebar.TFrame", padding=(12, 8))
        self.menu_id = self.canvas.create_window((0, 0), window=self.menu, anchor="nw")

        def _on_frame_configure(event=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas_configure(event):
            self.canvas.itemconfig(self.menu_id, width=event.width)

        self.menu.bind("<Configure>", _on_frame_configure)
        self.canvas.bind("<Configure>", _on_canvas_configure)

        # Itens do menu
        for routine in app.iter_routines():
            page_name = routine["page_name"]
            btn_text = f"{routine['icon']} {app.get_routine_nav_label(page_name)}"
            self._add_btn(
                routine["sidebar_key"],
                btn_text,
                lambda page_name=page_name: app.show_page(page_name),
            )

        # Rodapé
        bottom = ttk.Frame(self, style="Sidebar.TFrame", padding=(12, 14))
        bottom.pack(fill="x")

        ttk.Button(bottom, text="\u23FB SAIR", style="Danger.TButton", command=self._safe_quit).pack(fill="x")

    def _safe_quit(self):
        """Evita travas se houver janelas abertas/toplevels"""
        if hasattr(self.app, "request_close"):
            try:
                self.app.request_close()
                return
            except Exception:
                logging.debug("Falha ignorada")
        try:
            self.app.quit()
        except Exception:
            try:
                self.app.destroy()
            except Exception:
                logging.debug("Falha ignorada")

    def _add_btn(self, key, text, cmd):
        """Adiciona botão ao menu lateral"""
        b = ttk.Button(self.menu, text=text, style="Side.TButton", command=cmd)
        b.pack(fill="x", pady=2)
        self.buttons[key] = b

    def set_active(self, page_key):
        """Destaca botão ativo"""
        for k, b in self.buttons.items():
            b.config(style="SideActive.TButton" if k == page_key else "Side.TButton")


class App(tk.Tk):
    def __init__(self, user=None):
        super().__init__()
        # Evita "flash" preto durante carga inicial pesada das páginas.
        self.withdraw()

        # âÅâ€œââ‚¬¦ Segurança básica: sempre garante chaves esperadas
        base_user = {"nome": "ADMIN", "is_admin": True}
        if isinstance(user, dict):
            base_user.update(user)
        self.user = base_user

        self.title(APP_TITLE_DESKTOP)
        apply_window_icon(self)
        init_w, init_h = compute_initial_window_size(self)
        self.geometry(f"{init_w}x{init_h}")
        self.minsize(1080, 680)

        try:
            self.state("zoomed")
        except Exception:
            logging.debug("Falha ignorada")

        apply_style(self)
        try:
            self._base_tk_scaling = float(self.tk.call("tk", "scaling"))
        except Exception:
            self._base_tk_scaling = 1.0
        self._ui_zoom = 1.0
        db_init()

        if not self.ensure_integracao_sistema("Inicializacao do sistema", force_probe=True):
            self.after(80, self.destroy)
            return

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(0, minsize=compute_sidebar_width(self))

        self.sidebar = Sidebar(self, self)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        try:
            self.sidebar.canvas._wheel_scroll_target = True
        except Exception:
            logging.debug("Falha ignorada")

        self.container = ttk.Frame(self, style="Content.TFrame")
        self.container.grid(row=0, column=1, sticky="nsew")
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.pages = {}
        self._page_factories = {}
        self.current_page_name = None
        self._register_page_factories()
        self._build_initial_pages()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self.request_close)

        self.show_page("Home")
        self.update_idletasks()
        self.deiconify()

    def ensure_integracao_sistema(self, contexto: str = "Operacao", force_probe: bool = False) -> bool:
        return ensure_system_api_binding(context=contexto, parent=self, force_probe=force_probe)

    def request_close(self):
        """Confirma fechamento da aplicação."""
        try:
            ok = messagebox.askyesno("Confirmar saida", "Deseja realmente fechar o sistema?")
        except Exception:
            ok = True
        if not ok:
            return
        try:
            self.quit()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self.destroy()
        except Exception:
            logging.debug("Falha ignorada")

    def _register_page_factories(self):
        """Registra factories para criação lazy das páginas."""
        self._page_factories = {
            "Home": HomePage,
            "Cadastros": CadastrosPage,
            "Rotas": RotasPage,
            "ImportarVendas": ImportarVendasPage,
            "Programacao": ProgramacaoPage,
            "Recebimentos": RecebimentosPage,
            "Despesas": DespesasPage,
            "Escala": EscalaPage,
            "CentroCustos": CentroCustosPage,
            "Relatorios": RelatoriosPage,
            "BackupExportar": BackupExportarPage,
            "Permissoes": PermissionsPage,
            "Ferramentas": SystemToolsPage,
        }

    def _create_page_if_needed(self, name):
        p = self.pages.get(name)
        if p:
            return p
        factory = self._page_factories.get(name)
        if not factory:
            return None
        context = self._build_page_context(name)
        if context is not None:
            p = factory(self.container, self, context=context)
        else:
            p = factory(self.container, self)
        p.grid(row=0, column=0, sticky="nsew")
        self.pages[name] = p
        return p

    def _build_page_context(self, name: str):
        if name == "Home":
            return AppContext(
                config=APP_CONFIG,
                hooks={
                    "apply_window_icon": apply_window_icon,
                    "can_read_from_api": can_read_from_api,
                    "is_desktop_api_sync_enabled": is_desktop_api_sync_enabled,
                    "resolve_equipe_nomes": resolve_equipe_nomes,
                },
            )
        if name == "Relatorios":
            return AppContext(
                config=APP_CONFIG,
                hooks={
                    "build_folha_retorno_operacional": build_folha_retorno_operacional,
                    "can_read_from_api": can_read_from_api,
                    "db_has_column": db_has_column,
                    "ensure_system_api_binding": ensure_system_api_binding,
                    "fetch_programacao_itens": lambda codigo, limit=8000: (
                        lambda r: (r.get("data") or []) if isinstance(r, dict) and bool(r.get("ok", False)) else []
                    )(fetch_programacao_itens_contract(codigo, limit=limit)),
                    "is_desktop_api_sync_enabled": is_desktop_api_sync_enabled,
                    "normalize_date_column": normalize_date_column,
                    "normalize_datetime_column": normalize_datetime_column,
                    "resolve_equipe_nomes": resolve_equipe_nomes,
                    "require_reportlab": require_reportlab,
                },
            )
        return None

    def _build_initial_pages(self):
        """Cria apenas páginas essenciais na inicialização para evitar travamento."""
        self._create_page_if_needed("Home")

    def iter_routines(self):
        return ROUTINE_DEFINITIONS

    def get_routine_meta(self, name_or_alias):
        return get_routine_meta(name_or_alias)

    def get_routine_code(self, name_or_alias) -> str:
        return get_routine_code(name_or_alias)

    def get_routine_nav_label(self, name_or_alias, include_code: bool = False) -> str:
        return format_routine_nav_label(name_or_alias, include_code=include_code)

    def get_routine_title(self, name_or_alias, include_code: bool = False) -> str:
        return format_routine_title(name_or_alias, include_code=include_code)

    def get_routine_sidebar_key(self, name_or_alias):
        return get_routine_sidebar_key(name_or_alias)

    def get_subroutine_code(self, name_or_alias, submenu_index) -> str:
        return format_subroutine_code(name_or_alias, submenu_index)

    def show_page(self, name):
        """Exibe página e atualiza menu lateral"""
        sidebar_key = self.get_routine_sidebar_key(name)
        if sidebar_key:
            self.sidebar.set_active(sidebar_key)

        page = self._create_page_if_needed(name)
        if not page:
            messagebox.showwarning("ATENÇÃO", f"Rotina '{self.get_routine_nav_label(name)}' não encontrada.")
            return

        page.tkraise()
        self.current_page_name = name
        try:
            self.title(f"{APP_TITLE_DESKTOP} | {self.get_routine_title(name)}")
        except Exception:
            logging.debug("Falha ignorada")
        started = time.perf_counter()

        def _run_on_show():
            try:
                if page.winfo_exists() and self.current_page_name == name:
                    if hasattr(page, "refresh_header_title"):
                        page.refresh_header_title()
                    page.on_show()
                    elapsed_ms = (time.perf_counter() - started) * 1000.0
                    logging.info("Rotina %s exibida | %.0f ms", self.get_routine_title(name), elapsed_ms)
            except Exception as e:
                messagebox.showerror("ERRO", f"Erro ao abrir rotina '{self.get_routine_title(name)}':\n\n{e}")

        self.after_idle(_run_on_show)

    def refresh_programacao_comboboxes(self):
        """Atualiza comboboxes em páginas relevantes"""
        for nm in ["Programacao", "Despesas", "Relatorios", "Recebimentos", "Home", "CentroCustos"]:
            p = self.pages.get(nm)
            if p and hasattr(p, "refresh_comboboxes"):
                try:
                    p.refresh_comboboxes()
                except Exception:
                    logging.debug("Falha ignorada")

            if nm == "Home" and p and hasattr(p, "on_show"):
                try:
                    p.on_show()
                except Exception:
                    logging.debug("Falha ignorada")

    # =========================================================
    # Atalhos globais de teclado
    # =========================================================
    def _bind_shortcuts(self):
        self.bind("<F2>", self._shortcut_insert)
        self.bind("<F3>", self._shortcut_edit)
        self.bind("<F5>", self._shortcut_refresh)
        self.bind("<Delete>", self._shortcut_delete)
        self.bind("<Control-s>", self._shortcut_save)
        self.bind("<Control-S>", self._shortcut_save)
        self.bind("<Control-l>", self._shortcut_clear)
        self.bind("<Control-L>", self._shortcut_clear)
        self.bind("<Control-p>", self._shortcut_print)
        self.bind("<Control-P>", self._shortcut_print)
        self.bind("<Escape>", self._shortcut_escape)
        self.bind("<Return>", self._shortcut_enter)
        self.bind("<Control-0>", self._shortcut_reset_zoom)
        self.bind("<Control-Key-0>", self._shortcut_reset_zoom)
        self.bind_all("<MouseWheel>", self._on_global_mousewheel, add="+")
        self.bind_all("<Shift-MouseWheel>", self._on_global_shift_mousewheel, add="+")
        self.bind_all("<Control-MouseWheel>", self._on_global_zoom_mousewheel, add="+")
        self.bind_all("<Button-4>", self._on_global_mousewheel_linux, add="+")
        self.bind_all("<Button-5>", self._on_global_mousewheel_linux, add="+")
        self.bind_all("<Shift-Button-4>", self._on_global_shift_mousewheel_linux, add="+")
        self.bind_all("<Shift-Button-5>", self._on_global_shift_mousewheel_linux, add="+")
        self.bind_all("<Control-Button-4>", self._on_global_zoom_mousewheel_linux, add="+")
        self.bind_all("<Control-Button-5>", self._on_global_zoom_mousewheel_linux, add="+")

    def _is_descendant_of(self, widget, ancestor):
        try:
            w = widget
            while w is not None:
                if w == ancestor:
                    return True
                w = getattr(w, "master", None)
        except Exception:
            return False
        return False

    def _find_canvas_scroll_target(self, widget):
        try:
            w = widget
            while w is not None:
                if isinstance(w, (ttk.Treeview, tk.Text, tk.Listbox)):
                    return None
                if getattr(w, "_wheel_scroll_target", False):
                    return w
                w = getattr(w, "master", None)
        except Exception:
            return None
        page = self._active_page()
        if page and hasattr(page, "body_canvas") and self._is_descendant_of(widget, page):
            return page.body_canvas
        if hasattr(self, "sidebar") and self._is_descendant_of(widget, self.sidebar):
            return getattr(self.sidebar, "canvas", None)
        return None

    def _mousewheel_units(self, event):
        try:
            if getattr(event, "num", None) == 4:
                return -3
            if getattr(event, "num", None) == 5:
                return 3
            delta = int(getattr(event, "delta", 0) or 0)
            if delta == 0:
                return 0
            base = max(1, int(abs(delta) / 120))
            return -base if delta > 0 else base
        except Exception:
            return 0

    def _scroll_canvas(self, event, axis="y"):
        target = self._find_canvas_scroll_target(getattr(event, "widget", None))
        if not target:
            return None
        units = self._mousewheel_units(event)
        if units == 0:
            return None
        try:
            if axis == "x":
                target.xview_scroll(units, "units")
            else:
                target.yview_scroll(units, "units")
            return "break"
        except Exception:
            return None

    def _refresh_zoom_layout(self):
        try:
            sidebar_width = compute_sidebar_width(self)
            self.grid_columnconfigure(0, minsize=sidebar_width)
            self.sidebar.configure(width=sidebar_width)
            self.sidebar.update_idletasks()
            for page in self.pages.values():
                if hasattr(page, "body_canvas"):
                    page.body_canvas.update_idletasks()
                    try:
                        page.body_canvas.configure(scrollregion=page.body_canvas.bbox("all"))
                    except Exception:
                        logging.debug("Falha ignorada")
            self.update_idletasks()
        except Exception:
            logging.debug("Falha ignorada")

    def _set_ui_zoom(self, zoom_value):
        zoom_clamped = min(1.35, max(0.85, float(zoom_value)))
        if abs(zoom_clamped - self._ui_zoom) < 0.001:
            return
        self._ui_zoom = zoom_clamped
        try:
            self.tk.call("tk", "scaling", self._base_tk_scaling * self._ui_zoom)
        except Exception:
            logging.debug("Falha ignorada")
            return
        self.after_idle(self._refresh_zoom_layout)

    def _shortcut_reset_zoom(self, _e=None):
        self._set_ui_zoom(1.0)
        return "break"

    def _on_global_mousewheel(self, event):
        state = int(getattr(event, "state", 0) or 0)
        if state & 0x4 or state & 0x1:
            return None
        return self._scroll_canvas(event, axis="y")

    def _on_global_shift_mousewheel(self, event):
        state = int(getattr(event, "state", 0) or 0)
        if state & 0x4:
            return None
        return self._scroll_canvas(event, axis="x")

    def _on_global_zoom_mousewheel(self, event):
        units = self._mousewheel_units(event)
        if units == 0:
            return None
        self._set_ui_zoom(self._ui_zoom + (-0.05 * units))
        return "break"

    def _on_global_mousewheel_linux(self, event):
        state = int(getattr(event, "state", 0) or 0)
        if state & 0x4 or state & 0x1:
            return None
        return self._scroll_canvas(event, axis="y")

    def _on_global_shift_mousewheel_linux(self, event):
        state = int(getattr(event, "state", 0) or 0)
        if state & 0x4:
            return None
        return self._scroll_canvas(event, axis="x")

    def _on_global_zoom_mousewheel_linux(self, event):
        units = self._mousewheel_units(event)
        if units == 0:
            return None
        self._set_ui_zoom(self._ui_zoom + (-0.05 * units))
        return "break"

    def _active_page(self):
        return self.pages.get(self.current_page_name) if self.current_page_name else None

    def _try_call(self, page, names):
        if not page:
            return False
        for nm in names:
            fn = getattr(page, nm, None)
            if callable(fn):
                try:
                    fn()
                    return True
                except Exception:
                    logging.debug("Falha ignorada")
        return False

    def _is_text_input_focus(self):
        try:
            w = self.focus_get()
            if w is None:
                return False
            cls = str(w.winfo_class()).upper()
            return cls in {"ENTRY", "TEXT", "TENTRY", "TCOMBOBOX", "SPINBOX"}
        except Exception:
            return False

    def _close_top_modal_if_any(self):
        try:
            tops = [w for w in self.winfo_children() if isinstance(w, tk.Toplevel) and w.winfo_exists()]
            if not tops:
                return False
            tops[-1].destroy()
            return True
        except Exception:
            return False

    def _shortcut_escape(self, _e=None):
        # ESC: fecha modal atual; se não houver, volta para Home.
        if self._close_top_modal_if_any():
            return "break"
        if self.current_page_name and self.current_page_name != "Home":
            try:
                self.show_page("Home")
                return "break"
            except Exception:
                logging.debug("Falha ignorada")
        return None

    def _shortcut_insert(self, _e=None):
        page = self._active_page()
        if self._try_call(page, ["inserir_linha", "inserir_cliente_manual", "_open_registrar_rapido"]):
            return "break"
        return None

    def _shortcut_edit(self, _e=None):
        page = self._active_page()
        if self._try_call(page, ["editar", "_editar_linha_selecionada"]):
            return "break"
        return None

    def _shortcut_delete(self, _e=None):
        # DEL: não intercepta quando estiver digitando em campo.
        if self._is_text_input_focus():
            return None
        page = self._active_page()
        if self._try_call(page, ["excluir", "remover_linha", "_excluir_linha_selecionada", "zerar_recebimento"]):
            return "break"
        return None

    def _shortcut_refresh(self, _e=None):
        page = self._active_page()
        if self._try_call(page, ["refresh_data", "_refresh_all", "carregar", "carregar_programacao", "refresh_comboboxes"]):
            return "break"
        return None

    def _shortcut_save(self, _e=None):
        page = self._active_page()
        if self._try_call(page, ["salvar_programacao", "salvar_recebimento", "salvar_tudo", "salvar_alteracoes", "salvar"]):
            return "break"
        return None

    def _shortcut_print(self, _e=None):
        page = self._active_page()
        if self._try_call(page, ["imprimir_romaneios_programacao", "imprimir_pdf", "imprimir_resumo", "gerar_pdf"]):
            return "break"
        return None

    def _shortcut_clear(self, _e=None):
        if self._is_text_input_focus():
            return None
        page = self._active_page()
        if self._try_call(page, ["limpar", "limpar_tudo", "limpar_itens", "_clear_form_recebimento", "_limpar_filtros_relatorio", "_limpar_busca_despesas"]):
            return "break"
        return None

    def _shortcut_enter(self, _e=None):
        # ENTER global: executa ação principal da tela quando não está digitando em campo.
        if self._is_text_input_focus():
            return None
        page = self._active_page()
        if self._try_call(page, ["salvar_programacao", "salvar_recebimento", "salvar_tudo", "gerar_resumo", "refresh_data", "carregar"]):
            return "break"
        return None


class RotasPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Rotas")
        self.body.grid_rowconfigure(0, weight=1)

        card = ttk.Frame(self.body, style="Card.TFrame", padding=16)
        card.grid(row=0, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        ttk.Label(card, text="Monitoramento de Rotas", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            card,
            text="Atualizacao via GPS do app mobile. Selecione uma rota para abrir no mapa.",
            style="CardLabel.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 10))

        cols = ("CODIGO", "MOTORISTA", "VEICULO", "STATUS", "ULT_GPS", "LAT", "LON", "VEL", "PRECISAO")
        self.tree = ttk.Treeview(card, columns=cols, show="headings", height=16)
        self.tree.grid(row=2, column=0, sticky="nsew")
        self.tree.tag_configure("st_em_rota", background="#E8F4FF")
        self.tree.tag_configure("st_carregada", background="#FFF8E1")
        self.tree.tag_configure("st_iniciada", background="#E8F4FF")
        self.tree.tag_configure("spd_baixa", foreground="#1B5E20")
        self.tree.tag_configure("spd_media", foreground="#E65100")
        self.tree.tag_configure("spd_alta", foreground="#B71C1C")

        ysb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        ysb.grid(row=2, column=1, sticky="ns")

        xsb = ttk.Scrollbar(card, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=xsb.set)
        xsb.grid(row=3, column=0, sticky="ew")

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor="w")
        self.tree.column("CODIGO", width=130, anchor="center")
        self.tree.column("STATUS", width=110, anchor="center")
        self.tree.column("ULT_GPS", width=150, anchor="center")
        self.tree.column("LAT", width=110, anchor="center")
        self.tree.column("LON", width=110, anchor="center")
        self.tree.column("VEL", width=90, anchor="center")
        self.tree.column("PRECISAO", width=100, anchor="center")

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"LAT", "LON", "VEL", "PRECISAO"},
            money_cols=set(),
            date_cols=set(),
        )
        self.tree.bind("<Double-1>", lambda _e: self._abrir_mapa_selecionado())

        ttk.Button(self.footer_right, text="\U0001F504 ATUALIZAR", style="Ghost.TButton", command=self.carregar).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(self.footer_right, text="\U0001F5FA MAPA SELECIONADO", style="Primary.TButton", command=self._abrir_mapa_selecionado).grid(
            row=0, column=1, padx=4
        )
        ttk.Button(self.footer_right, text="\U0001F5FA MAPA DE TODAS", style="Primary.TButton", command=self._abrir_mapa_todas).grid(
            row=0, column=2, padx=4
        )

        self._rows_cache = []
        self._load_seq = 0
        self._last_source = "LOCAL"
        self._last_error = ""
        self._refresh_job = None
        self._refresh_ms = 10000
        self._refresh_var = tk.StringVar(value="10s")
        ttk.Label(self.footer_left, text="Atualização automática:", style="CardLabel.TLabel").grid(
            row=0, column=0, padx=(0, 6), sticky="w"
        )
        self.cb_refresh = ttk.Combobox(
            self.footer_left,
            state="readonly",
            width=8,
            textvariable=self._refresh_var,
            values=["5s", "10s", "30s"],
        )
        self.cb_refresh.grid(row=0, column=1, padx=(0, 6), sticky="w")
        self.cb_refresh.bind("<<ComboboxSelected>>", self._on_change_refresh_interval)
        self.bind("<Destroy>", self._on_destroy, add="+")

    def _fmt_num(self, v, casas=6):
        try:
            if v is None or str(v).strip() == "":
                return ""
            return f"{float(v):.{casas}f}".replace(".", ",")
        except Exception:
            return ""

    def _fmt_2(self, v):
        try:
            if v is None or str(v).strip() == "":
                return ""
            return f"{float(v):.2f}".replace(".", ",")
        except Exception:
            return ""

    def _fetch_rows(self):
        rows = []
        self._last_source = "LOCAL"
        self._last_error = ""

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and can_read_from_api():
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
                if isinstance(api_rows, list):
                    for r in api_rows:
                        if not isinstance(r, dict):
                            continue
                        rows.append(
                            {
                                "codigo_programacao": str(r.get("codigo_programacao") or "").strip(),
                                "motorista": str(r.get("motorista") or "").strip(),
                                "veiculo": str(r.get("veiculo") or "").strip(),
                                "status": str(
                                    r.get("status_operacional")
                                    or r.get("status")
                                    or r.get("status_base")
                                    or ""
                                ).strip(),
                                "lat": r.get("lat"),
                                "lon": r.get("lon"),
                                "speed": r.get("speed"),
                                "accuracy": r.get("accuracy"),
                                "recorded_at": str(r.get("recorded_at") or "").strip(),
                            }
                        )
                    self._last_source = "API CENTRAL"
                    return rows
            except Exception as e:
                self._last_error = str(e or "")

        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()

            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='rota_gps_pings'")
            has_gps = cur.fetchone() is not None
            cur.execute("PRAGMA table_info(programacoes)")
            cols_prog = {str(c[1]).lower() for c in (cur.fetchall() or [])}
            if "status_operacional" in cols_prog:
                status_expr = "COALESCE(NULLIF(TRIM(p.status_operacional), ''), COALESCE(p.status, ''))"
            else:
                status_expr = "COALESCE(p.status, '')"

            # Filtro defensivo: não listar rotas já encerradas mesmo em bases legadas.
            conds = [
                "UPPER(TRIM(COALESCE(p.status, ''))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')"
            ]
            if "status_operacional" in cols_prog:
                conds.append(
                    "UPPER(TRIM(COALESCE(p.status_operacional, ''))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')"
                )
            if "finalizada_no_app" in cols_prog:
                conds.append("COALESCE(p.finalizada_no_app, 0)=0")
            if "data_chegada" in cols_prog:
                conds.append("TRIM(COALESCE(p.data_chegada, ''))=''")
            if "hora_chegada" in cols_prog:
                conds.append("TRIM(COALESCE(p.hora_chegada, ''))=''")
            if "km_final" in cols_prog:
                conds.append("COALESCE(p.km_final, 0)=0")
            status_where = " AND ".join(conds)

            if has_gps:
                sql = """
                    SELECT
                        p.codigo_programacao,
                        COALESCE(p.motorista, '') AS motorista,
                        COALESCE(p.veiculo, '') AS veiculo,
                        """ + status_expr + """ AS status,
                        g.lat,
                        g.lon,
                        g.speed,
                        g.accuracy,
                        g.recorded_at
                    FROM programacoes p
                    LEFT JOIN (
                        SELECT r1.codigo_programacao, r1.lat, r1.lon, r1.speed, r1.accuracy, r1.recorded_at
                        FROM rota_gps_pings r1
                        INNER JOIN (
                            SELECT codigo_programacao, MAX(id) AS max_id
                            FROM rota_gps_pings
                            GROUP BY codigo_programacao
                        ) r2 ON r2.max_id = r1.id
                    ) g ON g.codigo_programacao = p.codigo_programacao
                    WHERE """ + status_where + """
                    ORDER BY p.id DESC
                    LIMIT 300
                """
            else:
                sql = """
                    SELECT
                        p.codigo_programacao,
                        COALESCE(p.motorista, '') AS motorista,
                        COALESCE(p.veiculo, '') AS veiculo,
                        """ + status_expr + """ AS status,
                        NULL AS lat,
                        NULL AS lon,
                        NULL AS speed,
                        NULL AS accuracy,
                        NULL AS recorded_at
                    FROM programacoes p
                    WHERE """ + status_where + """
                    ORDER BY p.id DESC
                    LIMIT 300
                """
            cur.execute(sql)
            for r in cur.fetchall() or []:
                rows.append(dict(r))
        return rows

    def _apply_rows(self, seq, rows):
        if seq != self._load_seq or not self.winfo_exists():
            return
        self._rows_cache = rows or []
        self.tree.delete(*self.tree.get_children())

        com_gps = 0
        for r in self._rows_cache:
            lat = r.get("lat")
            lon = r.get("lon")
            if lat is not None and lon is not None:
                com_gps += 1
            status_txt = upper(r.get("status", ""))
            tags = []
            if status_txt in {"EM_ROTA", "EM ROTA"}:
                tags.append("st_em_rota")
            elif status_txt == "CARREGADA":
                tags.append("st_carregada")
            elif status_txt == "INICIADA":
                tags.append("st_iniciada")
            try:
                spd = float(r.get("speed")) if r.get("speed") is not None else 0.0
            except Exception:
                spd = 0.0
            if spd >= 60:
                tags.append("spd_alta")
            elif spd >= 30:
                tags.append("spd_media")
            elif spd > 0:
                tags.append("spd_baixa")
            tree_insert_aligned(
                self.tree,
                "",
                "end",
                (
                    r.get("codigo_programacao", ""),
                    upper(r.get("motorista", "")),
                    upper(r.get("veiculo", "")),
                    status_txt,
                    r.get("recorded_at", "") or "",
                    self._fmt_num(lat, 6),
                    self._fmt_num(lon, 6),
                    self._fmt_2(r.get("speed")),
                    self._fmt_2(r.get("accuracy")),
                ),
                tags=tuple(tags),
            )
        fonte = self._last_source
        if self._last_error:
            fonte = f"LOCAL (fallback API: {self._last_error})"
        self.set_status(f"STATUS: {len(self._rows_cache)} rota(s) ativa(s). GPS em {com_gps} rota(s). Fonte: {fonte}.")

    def _handle_rows_error(self, seq, exc):
        if seq != self._load_seq or not self.winfo_exists():
            return
        logging.error("Falha ao carregar rotas", exc_info=(type(exc), exc, exc.__traceback__))
        self.set_status("STATUS: Falha ao carregar monitoramento de rotas.")

    def carregar(self, async_load=True):
        self._load_seq += 1
        seq = self._load_seq
        self.set_status("STATUS: Carregando monitoramento de rotas...")
        if not async_load:
            try:
                rows = self._fetch_rows()
            except Exception as exc:
                self._handle_rows_error(seq, exc)
                return
            self._apply_rows(seq, rows)
            return

        run_async_ui(
            self,
            self._fetch_rows,
            lambda rows, seq=seq: self._apply_rows(seq, rows),
            lambda exc, seq=seq: self._handle_rows_error(seq, exc),
        )

    def _start_auto_refresh(self):
        self._stop_auto_refresh()
        self._refresh_job = self.after(self._refresh_ms, self._auto_refresh_tick)

    def _stop_auto_refresh(self):
        if self._refresh_job:
            try:
                self.after_cancel(self._refresh_job)
            except Exception:
                logging.debug("Falha ignorada")
            self._refresh_job = None

    def _auto_refresh_tick(self):
        self._refresh_job = None
        try:
            if self.winfo_exists() and self.winfo_ismapped():
                self.carregar()
                self._refresh_job = self.after(self._refresh_ms, self._auto_refresh_tick)
        except Exception:
            logging.debug("Falha ignorada")

    def _on_destroy(self, _event=None):
        self._stop_auto_refresh()

    def _on_change_refresh_interval(self, _event=None):
        sel = (self._refresh_var.get() or "").strip().lower()
        if sel == "5s":
            self._refresh_ms = 5000
        elif sel == "30s":
            self._refresh_ms = 30000
        else:
            self._refresh_ms = 10000
        self._start_auto_refresh()
        self.set_status(
            f"STATUS: atualização automática ajustada para {sel or '10s'}."
        )

    def _build_map_points(self, rows):
        points = []
        for r in rows:
            try:
                lat = float(r.get("lat"))
                lon = float(r.get("lon"))
            except Exception:
                continue
            points.append(
                {
                    "codigo": str(r.get("codigo_programacao", "")),
                    "motorista": str(r.get("motorista", "")),
                    "veiculo": str(r.get("veiculo", "")),
                    "status": str(r.get("status", "")),
                    "lat": lat,
                    "lon": lon,
                    "updated_at": str(r.get("recorded_at", "") or ""),
                    "speed": r.get("speed"),
                    "accuracy": r.get("accuracy"),
                }
            )
        return points

    def _open_map_html(self, points, title):
        if not points:
            messagebox.showwarning("Rotas", "Nenhuma coordenada GPS disponivel para exibir no mapa.")
            return

        center_lat = points[0]["lat"]
        center_lon = points[0]["lon"]
        points_json = json.dumps(points, ensure_ascii=False)
        title_js = json.dumps(title, ensure_ascii=False)

        html = f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>{title}</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    html, body, #map {{ height: 100%; margin: 0; }}
    .top {{ position: fixed; top: 8px; left: 8px; right: 8px; z-index: 9999; background:#fff; border:1px solid #ddd; border-radius:8px; padding:8px 10px; font-family:Segoe UI, Arial; font-size:14px; }}
  </style>
</head>
<body>
  <div class="top"></div>
  <div id="map"></div>
  <script>
    const title = {title_js};
    const points = {points_json};
    document.querySelector('.top').textContent = title + ' | veiculos no mapa: ' + points.length;
    const map = L.map('map').setView([{center_lat}, {center_lon}], 11);
    L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap'
    }}).addTo(map);

    const bounds = [];
    points.forEach(p => {{
      const marker = L.marker([p.lat, p.lon]).addTo(map);
      const speed = (p.speed === null || p.speed === undefined || p.speed === '') ? '-' : Number(p.speed).toFixed(1);
      const acc = (p.accuracy === null || p.accuracy === undefined || p.accuracy === '') ? '-' : Number(p.accuracy).toFixed(1);
      const gmaps = `https://www.google.com/maps?q=${{p.lat}},${{p.lon}}`;
      marker.bindPopup(
        `<b>${{p.codigo}}</b><br/>` +
        `Motorista: ${{p.motorista || '-'}}<br/>` +
        `Veiculo: ${{p.veiculo || '-'}}<br/>` +
        `Status: ${{p.status || '-'}}<br/>` +
        `Atualizacao: ${{p.updated_at || '-'}}<br/>` +
        `Vel: ${{speed}} km/h<br/>` +
        `Precisao: ${{acc}} m<br/>` +
        `<a href="${{gmaps}}" target="_blank">Abrir no Google Maps</a>`
      );
      bounds.push([p.lat, p.lon]);
    }});
    if (bounds.length > 1) {{
      map.fitBounds(bounds, {{padding: [30, 30]}});
    }}
  </script>
</body>
</html>
"""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".html", delete=False, encoding="utf-8") as fp:
            fp.write(html)
            html_path = fp.name
        webbrowser.open(f"file:///{html_path.replace(os.sep, '/')}")

    def _abrir_mapa_todas(self):
        points = self._build_map_points(self._rows_cache)
        self._open_map_html(points, "Rastreamento em tempo real - Todas as rotas")

    def _abrir_mapa_selecionado(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Rotas", "Selecione uma rota na tabela.")
            return
        vals = self.tree.item(sel[0], "values") or []
        if not vals:
            return
        codigo = str(vals[0]).strip()
        row = next((r for r in self._rows_cache if str(r.get("codigo_programacao", "")).strip() == codigo), None)
        if not row:
            messagebox.showwarning("Rotas", "Rota selecionada nao encontrada no cache.")
            return
        points = self._build_map_points([row])
        self._open_map_html(points, f"Rastreamento - {codigo}")

    def on_show(self):
        self.carregar()
        self._start_auto_refresh()


# ==========================
# ===== FIM DA PARTE 2 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INICIO DA PARTE X (FINAL / SEM DUPLICIDADE) =====
# ==========================

# =========================================================
# UTILITRIOS DE BANCO / BUSCA (PROGRAMACOES ATIVAS / ITENS)
# =========================================================

_SAFE_SQL_IDENT = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")

def _safe_ident(name: str) -> str:
    """
    Valida identificador SQL (tabela/coluna) para evitar SQL injection.
    SQLite não aceita parametrizar PRAGMA table_info(?)  então validamos.
    """
    name = str(name or "").strip()
    if not _SAFE_SQL_IDENT.match(name):
        raise ValueError(f"Identificador SQL inválido: {name!r}")
    return name


def db_has_column(cur, table_name: str, column_name: str) -> bool:
    """
    Verifica se uma coluna existe numa tabela (SQLite).
    Segurança: não altera nada, só consulta PRAGMA.
    """
    return repository_db_has_column(cur, table_name, column_name)


def fetch_programacoes_ativas(limit: int = 400):
    """
    Retorna lista de programações ATIVAS (compatÃÂÂvel com bases antigas e novas).
    SaÃÂÂda: lista de dicts: {codigo, motorista, veiculo, equipe, data_criacao, status}
    """
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if desktop_secret and can_read_from_api():
        try:
            resp = _call_api(
                "GET",
                "desktop/overview",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            api_rows = resp.get("rotas") if isinstance(resp, dict) else []
            out_api = []
            for r in (api_rows or [])[: max(int(limit or 0), 1)]:
                if not isinstance(r, dict):
                    continue
                out_api.append(
                    {
                        "codigo": upper(r.get("codigo_programacao") or ""),
                        "motorista": upper(r.get("motorista") or ""),
                        "veiculo": upper(r.get("veiculo") or ""),
                        "equipe": upper(r.get("equipe") or ""),
                        "data_criacao": str(r.get("data_criacao") or ""),
                        "status": upper(r.get("status_operacional") or r.get("status") or "ATIVA"),
                    }
                )
            if out_api:
                return out_api
        except Exception:
            logging.debug("Falha ao buscar programacoes ativas na API", exc_info=True)

    try:
        equipes_map = {}
        with get_db() as conn:
            cur = conn.cursor()

            cols_prog = {str(r[1]).lower() for r in (cur.execute("PRAGMA table_info(programacoes)").fetchall() or [])}
            data_expr = "COALESCE(data_criacao,'')" if "data_criacao" in cols_prog else ("COALESCE(data,'')" if "data" in cols_prog else "''")
            status_expr = (
                "COALESCE(NULLIF(TRIM(status_operacional), ''), COALESCE(status, ''))"
                if "status_operacional" in cols_prog
                else ("COALESCE(status,'')" if "status" in cols_prog else "'ATIVA'")
            )
            where_parts = []
            if "status_operacional" in cols_prog or "status" in cols_prog:
                where_parts.append(
                    "UPPER(TRIM(COALESCE("
                    + ("NULLIF(status_operacional, '')" if "status_operacional" in cols_prog else "status")
                    + ", COALESCE(status, 'ATIVA')))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')"
                    if "status_operacional" in cols_prog
                    else "UPPER(TRIM(COALESCE(status,'ATIVA'))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')"
                )
            if "finalizada_no_app" in cols_prog:
                where_parts.append("COALESCE(finalizada_no_app,0)=0")
            where_sql = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
            cur.execute(
                f"""
                SELECT
                    COALESCE(codigo_programacao,''),
                    COALESCE(motorista,''),
                    COALESCE(veiculo,''),
                    COALESCE(equipe,''),
                    {data_expr} as data_criacao,
                    {status_expr} as status
                FROM programacoes
                {where_sql}
                ORDER BY id DESC
                LIMIT ?
                """,
                (limit,),
            )

            rows = cur.fetchall()

            try:
                cur.execute("SELECT id, nome, sobrenome FROM ajudantes")
                for er in cur.fetchall() or []:
                    ajudante_id = str(er[0] if er else "").strip()
                    nomes = format_ajudante_nome(er[1] if er else "", er[2] if er else "", "")
                    if ajudante_id and nomes:
                        equipes_map[upper(ajudante_id)] = nomes
            except Exception:
                logging.debug("Falha ignorada")

            # fallback legado
            if not equipes_map:
                try:
                    cur.execute("SELECT codigo, ajudante1, ajudante2 FROM equipes")
                    for er in cur.fetchall() or []:
                        codigo = upper(er[0] if er else "")
                        nomes = format_equipe_nomes(er[1] if er else "", er[2] if er else "", "")
                        if codigo and nomes:
                            equipes_map[codigo] = nomes
                except Exception:
                    logging.debug("Falha ignorada")

        
        out = []
        for r in rows:
            equipe_raw = r[3] or ""
            equipe_key = upper(equipe_raw)
            if any(sep in str(equipe_raw or "") for sep in ["|", ",", ";", "/"]):
                equipe_nome = resolve_equipe_nomes(equipe_raw)
            else:
                equipe_nome = equipes_map.get(equipe_key, equipe_raw)
            out.append({
                "codigo": r[0] or "",
                "motorista": r[1] or "",
                "veiculo": r[2] or "",
                "equipe": equipe_nome or "",
                "data_criacao": r[4] or "",
                "status": r[5] or "",
            })
        return out


    except Exception:
        return []


def fetch_programacao_itens_contract(codigo_programacao: str, limit: int = 8000) -> dict:
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    result = fetch_programacao_itens_service(
        codigo_programacao=codigo_programacao,
        limit=limit,
        upper=upper,
        safe_int=safe_int,
        safe_float=safe_float,
        get_db=get_db,
        call_api=_call_api,
        quote=urllib.parse.quote,
        is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
        desktop_secret=desktop_secret,
    )
    if isinstance(result, dict):
        return result
    if isinstance(result, list):
        return {"ok": True, "data": result, "error": None, "source": "local"}
    return {"ok": False, "data": [], "error": "Retorno inesperado do service de programação.", "source": "local"}


def fetch_programacao_meta_relatorio(codigo_programacao: str) -> dict:
    codigo_programacao = upper(str(codigo_programacao or "").strip())
    if not codigo_programacao:
        return {}

    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if desktop_secret and is_desktop_api_sync_enabled():
        try:
            resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(codigo_programacao)}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            rota = resp.get("rota") if isinstance(resp, dict) else None
            if isinstance(rota, dict):
                return dict(rota)
        except Exception:
            logging.debug("Falha ao buscar meta da programacao via API", exc_info=True)

    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(programacoes)")
            cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}

            def _expr(*names, default="''"):
                for name in names:
                    if name.lower() in cols:
                        return f"COALESCE({name}, '')"
                return default

            def _expr_num(*names, default="0"):
                for name in names:
                    if name.lower() in cols:
                        return f"COALESCE({name}, 0)"
                return default

            cur.execute(
                f"""
                SELECT
                    {_expr('codigo_programacao')} as codigo_programacao,
                    {_expr('data_criacao')} as data_criacao,
                    {_expr('motorista')} as motorista,
                    {_expr('veiculo')} as veiculo,
                    {_expr('equipe')} as equipe,
                    {_expr('status')} as status,
                    {_expr('status_operacional')} as status_operacional,
                    {_expr('prestacao_status')} as prestacao_status,
                    {_expr('num_nf', 'nf_numero')} as num_nf,
                    {_expr('local_rota', 'tipo_rota')} as local_rota,
                    {_expr('local_carregamento', 'granja_carregada', 'local_carregado')} as local_carregamento,
                    {_expr('data_saida')} as data_saida,
                    {_expr('hora_saida')} as hora_saida,
                    {_expr('data_chegada')} as data_chegada,
                    {_expr('hora_chegada')} as hora_chegada,
                    {_expr_num('nf_kg')} as nf_kg,
                    {_expr_num('nf_caixas', 'caixas_carregadas')} as nf_caixas,
                    {_expr_num('nf_kg_carregado', 'kg_carregado')} as nf_kg_carregado,
                    {_expr_num('nf_kg_vendido')} as nf_kg_vendido,
                    {_expr_num('nf_saldo')} as nf_saldo,
                    {_expr_num('nf_preco')} as nf_preco,
                    {_expr_num('media')} as media,
                    {_expr_num('aves_caixa_final', 'qnt_aves_caixa_final')} as aves_caixa_final,
                    {_expr_num('mortalidade_transbordo_aves')} as mortalidade_transbordo_aves,
                    {_expr_num('mortalidade_transbordo_kg')} as mortalidade_transbordo_kg,
                    {_expr('obs_transbordo')} as obs_transbordo
                FROM programacoes
                WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)
                ORDER BY id DESC
                LIMIT 1
                """,
                (codigo_programacao,),
            )
            row = cur.fetchone()
            if not row:
                return {}

            keys = [
                "codigo_programacao",
                "data_criacao",
                "motorista",
                "veiculo",
                "equipe",
                "status",
                "status_operacional",
                "prestacao_status",
                "num_nf",
                "local_rota",
                "local_carregamento",
                "data_saida",
                "hora_saida",
                "data_chegada",
                "hora_chegada",
                "nf_kg",
                "nf_caixas",
                "nf_kg_carregado",
                "nf_kg_vendido",
                "nf_saldo",
                "nf_preco",
                "media",
                "aves_caixa_final",
                "mortalidade_transbordo_aves",
                "mortalidade_transbordo_kg",
                "obs_transbordo",
            ]
            return {k: row[i] for i, k in enumerate(keys)}
    except Exception:
        logging.debug("Falha ao buscar meta da programacao no banco", exc_info=True)
    return {}


def _parse_alteracao_campos(item: dict) -> tuple[str, str, str]:
    detalhe = str((item or {}).get("alteracao_detalhe") or "").strip()
    alterado_por = str((item or {}).get("alterado_por") or "").strip().upper()
    para_quem = ""
    motivo = ""
    autorizado = ""

    if detalhe:
        m = re.search(r"PARA\s+QUEM\s*[:=-]\s*([^|;]+)", detalhe, flags=re.IGNORECASE)
        if m:
            para_quem = m.group(1).strip()
        m = re.search(r"(POR\s+QUE|MOTIVO)\s*[:=-]\s*([^|;]+)", detalhe, flags=re.IGNORECASE)
        if m:
            motivo = m.group(2).strip()
        m = re.search(r"(AUTORIZADO\s+POR|QUEM\s+AUTORIZOU)\s*[:=-]\s*([^|;]+)", detalhe, flags=re.IGNORECASE)
        if m:
            autorizado = m.group(2).strip()

    if not motivo and detalhe:
        motivo = detalhe
    if not autorizado:
        autorizado = alterado_por

    return para_quem or "-", motivo or "-", autorizado or "-"


def build_folha_retorno_operacional(prog: str) -> str:
    prog = upper(str(prog or "").strip())
    if not prog:
        return "FOLHA DE RETORNO OPERACIONAL\n\nProgramacao nao informada."

    meta = fetch_programacao_meta_relatorio(prog)
    itens_res = fetch_programacao_itens_contract(prog)
    itens = itens_res.get("data") if isinstance(itens_res, dict) and bool(itens_res.get("ok", False)) else []
    if not meta and not itens:
        return (
            f"FOLHA DE RETORNO OPERACIONAL - {prog}\n"
            + "=" * 118
            + "\n\nProgramacao nao encontrada."
        )

    motorista = str(meta.get("motorista") or "")
    veiculo = str(meta.get("veiculo") or "")
    equipe_txt = resolve_equipe_nomes(str(meta.get("equipe") or ""))
    status = upper(str(meta.get("status_operacional") or meta.get("status") or ""))
    prest = upper(str(meta.get("prestacao_status") or ""))
    nf = str(meta.get("num_nf") or "")
    local_rota = str(meta.get("local_rota") or "")
    local_carreg = str(meta.get("local_carregamento") or "")
    data_saida, hora_saida = normalize_date_time_components(meta.get("data_saida"), meta.get("hora_saida"))
    data_chegada, hora_chegada = normalize_date_time_components(meta.get("data_chegada"), meta.get("hora_chegada"))
    data_criacao = format_date_br_short(str(meta.get("data_criacao") or ""))

    media_carregada = safe_float(meta.get("media"), 0.0)
    if media_carregada > 20:
        media_carregada = media_carregada / 1000.0
    aves_por_caixa = safe_int(meta.get("aves_caixa_final"), 0)
    if aves_por_caixa <= 0:
        aves_por_caixa = 6

    status_entregue = {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}
    status_cancelado = {"CANCELADO", "CANCELADA"}

    tot_clientes = len(itens)
    tot_entregues = 0
    tot_cancelados = 0
    tot_alterados = 0
    tot_cx_prog = 0
    tot_cx_ent = 0
    tot_kg_prog = 0.0
    tot_kg_ent = 0.0
    tot_mort_aves = safe_int(meta.get("mortalidade_transbordo_aves"), 0)
    tot_mort_kg = safe_float(meta.get("mortalidade_transbordo_kg"), 0.0)
    tot_valor = 0.0
    divergencias = 0
    linhas = []
    ocorrencias = []

    for item in itens:
        status_item = upper(str(item.get("status_pedido") or "PENDENTE").strip()) or "PENDENTE"
        cx_prog = max(safe_int(item.get("qnt_caixas"), 0), 0)
        cx_alt = max(safe_int(item.get("caixas_atual"), 0), 0)
        preco_orig = safe_float(item.get("preco"), 0.0)
        preco_final = safe_float(item.get("preco_atual"), 0.0) or preco_orig
        kg_orig = safe_float(item.get("kg"), 0.0)
        mort_aves = max(safe_int(item.get("mortalidade_aves"), 0), 0)
        alterado = bool(
            str(item.get("alterado_em") or "").strip()
            or str(item.get("alterado_por") or "").strip()
            or str(item.get("alteracao_tipo") or "").strip()
            or str(item.get("alteracao_detalhe") or "").strip()
            or (cx_alt > 0 and cx_alt != cx_prog)
            or abs(preco_final - preco_orig) > 0.0001
        )

        if status_item in status_cancelado:
            cx_ent = 0
        elif status_item in status_entregue:
            cx_ent = cx_alt if cx_alt > 0 else cx_prog
        else:
            cx_ent = cx_alt if cx_alt > 0 else 0

        kg_base = safe_float(item.get("peso_previsto"), 0.0)
        if kg_base <= 0:
            kg_base = kg_orig
        if kg_base <= 0 and cx_ent > 0 and media_carregada > 0:
            kg_base = float(cx_ent) * float(max(aves_por_caixa, 1)) * media_carregada
        elif kg_base <= 0 and cx_prog > 0 and media_carregada > 0:
            kg_base = float(cx_prog) * float(max(aves_por_caixa, 1)) * media_carregada

        total_aves_cliente = max((cx_ent if cx_ent > 0 else cx_prog) * max(aves_por_caixa, 1), 0)
        media_cliente = (kg_base / total_aves_cliente) if total_aves_cliente > 0 and kg_base > 0 else media_carregada
        mort_kg = mort_aves * media_cliente if media_cliente > 0 else 0.0
        kg_ent = kg_base if status_item not in status_cancelado else 0.0
        valor_total = max((kg_ent * preco_final), 0.0) if status_item not in status_cancelado else 0.0
        delta_kg = kg_ent - kg_orig if kg_orig > 0 else 0.0

        tot_cx_prog += cx_prog
        tot_cx_ent += cx_ent
        tot_kg_prog += max(kg_orig, 0.0)
        tot_kg_ent += kg_ent
        tot_mort_aves += mort_aves
        tot_mort_kg += mort_kg
        tot_valor += valor_total
        if status_item in status_entregue:
            tot_entregues += 1
        if status_item in status_cancelado:
            tot_cancelados += 1
        if alterado:
            tot_alterados += 1
        if abs(delta_kg) >= 0.01:
            divergencias += 1

        linhas.append(
            {
                "cod": str(item.get("cod_cliente") or "")[:6],
                "cliente": upper(str(item.get("nome_cliente") or "-"))[:28],
                "st": status_item[:9],
                "cx": f"{cx_prog}/{cx_ent}",
                "med": f"{safe_float(media_cliente, 0.0):.3f}",
                "kg": f"{safe_float(kg_ent, 0.0):.2f}",
                "preco": f"{safe_float(preco_final, 0.0):.2f}",
                "valor": f"{safe_float(valor_total, 0.0):.2f}",
                "mort": f"{safe_int(mort_aves, 0)}/{safe_float(mort_kg, 0.0):.2f}",
                "div": (f"{delta_kg:+.2f}" if abs(delta_kg) >= 0.01 else "-"),
            }
        )

        if status_item in status_cancelado or alterado:
            para_quem, motivo, autorizado = _parse_alteracao_campos(item)
            ocorrencias.append(
                {
                    "cliente": upper(str(item.get("nome_cliente") or "-"))[:24],
                    "st": status_item[:10],
                    "cancel": "SIM" if status_item in status_cancelado else "NAO",
                    "alter": "SIM" if alterado else "NAO",
                    "para": para_quem[:16],
                    "motivo": motivo[:38],
                    "aut": autorizado[:16],
                }
            )

    kg_carregado = safe_float(meta.get("nf_kg_carregado"), 0.0)
    if kg_carregado <= 0:
        kg_carregado = tot_kg_prog
    caixas_carregadas = safe_int(meta.get("nf_caixas"), 0)
    if caixas_carregadas <= 0:
        caixas_carregadas = tot_cx_prog
    media_resumo = media_carregada
    if media_resumo <= 0 and caixas_carregadas > 0 and kg_carregado > 0:
        media_resumo = kg_carregado / float(caixas_carregadas * max(aves_por_caixa, 1))

    lines = []
    lines.append("=" * 118)
    lines.append(f"{'FOLHA DE RETORNO OPERACIONAL':^118}")
    lines.append("=" * 118)
    lines.append(
        f"CODIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {status or '-'}   PRESTACAO: {prest or '-'}"
    )
    lines.append(
        f"MOTORISTA: {upper(motorista or '-')}   VEICULO: {upper(veiculo or '-')}   EQUIPE: {equipe_txt or '-'}"
    )
    lines.append(
        f"NF: {nf or '-'}   LOCAL ROTA: {upper(local_rota or '-')}   CARREGOU EM: {upper(local_carreg or '-')}"
    )
    lines.append(
        f"SAIDA: {f'{data_saida} {hora_saida}'.strip() or '-'}   CHEGADA: {f'{data_chegada} {hora_chegada}'.strip() or '-'}"
    )
    lines.append("-" * 118)
    lines.append(
        f"CLIENTES: {tot_clientes}   ENTREGUES: {tot_entregues}   CANCELADOS: {tot_cancelados}   ALTERADOS: {tot_alterados}"
    )
    lines.append(
        f"KG CARREGADOS: {kg_carregado:.2f}   CX CARREGADAS: {caixas_carregadas}   MEDIA CARREGADA: {media_resumo:.3f}"
    )
    lines.append(
        f"KG ENTREGUE: {tot_kg_ent:.2f}   CX ENTREGUES: {tot_cx_ent}   MORTALIDADE TOTAL: {tot_mort_aves} AVES / {tot_mort_kg:.2f} KG"
    )
    lines.append(
        f"VALOR TOTAL ENTREGUE: {fmt_money(tot_valor)}   DIVERGENCIAS DE PESO: {divergencias} CLIENTE(S)"
    )
    lines.append("-" * 118)
    lines.append(
        f"{'COD':<6} {'CLIENTE':<28} {'STATUS':<9} {'CX P/E':>7} {'MEDIA':>7} {'KG':>10} {'PRECO':>10} {'VALOR':>12} {'MORT A/KG':>14} {'DIV KG':>9}"
    )
    lines.append("-" * 118)
    if not linhas:
        lines.append("Sem itens de entrega para esta programacao.")
    else:
        for row in linhas:
            lines.append(
                f"{row['cod']:<6} {row['cliente']:<28} {row['st']:<9} {row['cx']:>7} {row['med']:>7} "
                f"{row['kg']:>10} {row['preco']:>10} {row['valor']:>12} {row['mort']:>14} {row['div']:>9}"
            )
    lines.append("-" * 118)
    lines.append("[OCORRENCIAS / AJUSTES]")
    lines.append(
        f"{'CLIENTE':<24} {'STATUS':<10} {'CANC':<4} {'ALT':<4} {'PARA QUEM':<16} {'POR QUE':<38} {'AUTORIZOU':<16}"
    )
    lines.append("-" * 118)
    if not ocorrencias:
        lines.append("Sem cancelamentos ou alteracoes registradas.")
    else:
        for row in ocorrencias[:40]:
            lines.append(
                f"{row['cliente']:<24} {row['st']:<10} {row['cancel']:<4} {row['alter']:<4} "
                f"{row['para']:<16} {row['motivo']:<38} {row['aut']:<16}"
            )
        if len(ocorrencias) > 40:
            lines.append(f"... e mais {len(ocorrencias) - 40} ocorrencia(s).")
    return "\n".join(lines)


def format_prog_display(p: dict) -> str:
    """
    Formata texto amigável p/ lista/combobox sem quebrar.
    Ex: "ABC123 (2026-01-18) - JOAO - ABC1D23"
    """
    try:
        codigo = upper(p.get("codigo", ""))
        data = str(p.get("data_criacao", "") or "")[:10]
        motorista = upper(p.get("motorista", ""))[:18]
        veiculo = upper(p.get("veiculo", ""))[:12]
        if data:
            return f"{codigo} ({data}) - {motorista} - {veiculo}"
        return f"{codigo} - {motorista} - {veiculo}"
    except Exception:
        return upper(str(p.get("codigo", "")))

# ==========================
# ===== FIM DA PARTE X (FINAL) =====
# ==========================


# ==========================
# ===== INCIO DA PARTE 3 (FINAL / SEM DUPLICIDADE) =====
# ==========================

# ===== FIM DA PARTE 3 (FINAL) =====
# ==========================

configure_clientes_import_page_dependencies(
    ensure_system_api_binding=ensure_system_api_binding,
    is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
)

configure_centro_custos_page_dependencies(
    is_desktop_api_sync_enabled=is_desktop_api_sync_enabled,
)

configure_backup_exportar_page_dependencies(
    DB_PATH=DB_PATH,
    APP_CONFIG=APP_CONFIG,
    normalize_date_column=normalize_date_column,
)

# ==========================
# ===== INCIO DA PARTE 4 (ATUALIZADA) =====
# ==========================

# =========================================================
# 4.0 IMPORTAÇÃO DE VENDAS
# =========================================================
class ImportarVendasPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "ImportarVendas")
        self._load_seq = 0
        self._import_seq = 0
        self._vinc_seq = 0

        top = ttk.Frame(self.body, style="Content.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(5, weight=1)
        self.btn_importar = ttk.Button(top, text="IMPORTAR EXCEL", style="Primary.TButton", command=self.importar_excel)
        self.btn_importar.grid(row=0, column=0, padx=6)
        self.btn_limpar = ttk.Button(top, text="LIMPAR TUDO", style="Danger.TButton", command=self.limpar_tudo)
        self.btn_limpar.grid(row=0, column=1, sticky="w", padx=6)
        self.btn_atualizar = ttk.Button(top, text="ATUALIZAR", style="Ghost.TButton", command=self.carregar)
        self.btn_atualizar.grid(row=0, column=2, padx=6)
        self.lbl_info = ttk.Label(
            top,
            text="Selecione as vendas que irao para Programacao (duplo clique marca/desmarca).",
            background="#F4F6FB",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_info.grid(row=0, column=3, sticky="w", padx=10)
        self.pb_loading = ttk.Progressbar(top, mode="indeterminate", length=150)
        self.pb_loading.grid(row=0, column=4, padx=(12, 8), sticky="w")
        self.lbl_loading = ttk.Label(
            top,
            text="",
            background="#F4F6FB",
            foreground="#2563EB",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_loading.grid(row=0, column=5, sticky="w")

        card = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        card.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        # Filtros
        filt = ttk.Frame(card, style="Card.TFrame")
        filt.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        filt.grid_columnconfigure(1, weight=1)
        filt.grid_columnconfigure(7, weight=1)

        ttk.Label(filt, text="Buscar:", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_busca = ttk.Entry(filt, style="Field.TEntry")
        self.ent_busca.grid(row=0, column=1, sticky="ew", padx=6)
        self.ent_busca.bind("<Return>", lambda e: self.carregar())
        ttk.Button(filt, text="FILTRAR", style="Ghost.TButton", command=self.carregar).grid(row=0, column=2, padx=6)
        ttk.Button(filt, text="MARCAR", style="Primary.TButton", command=self.marcar_selecionadas).grid(row=0, column=3, padx=6)
        ttk.Button(filt, text="MARCAR TODOS", style="Warn.TButton", command=lambda: self.set_all_selected(1)).grid(row=0, column=4, padx=6)
        ttk.Button(filt, text="DESMARCAR TODOS", style="Ghost.TButton", command=lambda: self.set_all_selected(0)).grid(row=0, column=5, padx=6)
        ttk.Label(filt, text="Programação ativa:", style="CardLabel.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.cb_prog_vinculo = ttk.Combobox(filt, state="readonly")
        self.cb_prog_vinculo.grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=(10, 0))
        self.btn_vincular_prog = ttk.Button(
            filt,
            text="VINCULAR A PROGRAMAÇÃO",
            style="Primary.TButton",
            command=self.vincular_selecionadas_programacao,
        )
        self.btn_vincular_prog.grid(row=1, column=3, columnspan=3, padx=6, pady=(10, 0), sticky="ew")
        # Tabela (com scroll horizontal)
        cols = ["ID", "SEL", "PEDIDO", "DATA", "CLIENTE", "NOME COMPLETO", "PRODUTO", "VR TOTAL", "QNT", "CIDADE", "VENDEDOR"]
        table_wrap = ttk.Frame(card, style="Card.TFrame")
        table_wrap.grid(row=2, column=0, sticky="nsew")
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, minwidth=90, anchor="w")

        self.tree.column("ID", width=70, minwidth=60, anchor="center")
        self.tree.column("SEL", width=55, minwidth=55, anchor="center")
        self.tree.column("DATA", width=110, minwidth=90, anchor="center")
        self.tree.column("VR TOTAL", width=110, minwidth=90, anchor="e")
        self.tree.column("QNT", width=90, minwidth=70, anchor="center")
        self.tree.column("CIDADE", width=160, minwidth=120, anchor="w")
        self.tree.column("VENDEDOR", width=160, minwidth=120, anchor="w")

        self.tree.bind("<Double-1>", self.toggle_selected)
        self.tree.bind("<Button-1>", self._on_tree_click_toggle_sel, add="+")
        self.tree.bind("<Control-a>", self._shortcut_select_all_tree)
        self.tree.bind("<Control-A>", self._shortcut_select_all_tree)
        self.tree.bind("<space>", self._shortcut_toggle_selected_rows)
        self.tree.bind("<Delete>", self._shortcut_delete_selected_rows)

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"ID", "QNT"},
            money_cols={"VR TOTAL"},
            date_cols={"DATA"}
        )

    def _set_busy(self, busy, text=""):
        state = "disabled" if busy else "normal"
        for btn in (self.btn_importar, self.btn_limpar, self.btn_atualizar, getattr(self, "btn_vincular_prog", None)):
            if not btn:
                continue
            try:
                btn.configure(state=state)
            except Exception:
                logging.debug("Falha ao ajustar estado de botoes da tela de importar vendas")
        if busy:
            self.lbl_loading.configure(text=text or "Processando...")
            try:
                self.pb_loading.start(10)
            except Exception:
                logging.debug("Falha ao iniciar progressbar de importar vendas")
        else:
            try:
                self.pb_loading.stop()
            except Exception:
                logging.debug("Falha ao parar progressbar de importar vendas")
            self.lbl_loading.configure(text=text or "")

    def on_show(self):
        self.set_status("STATUS: Importação e seleção de vendas para programação.")
        self.refresh_programacoes_vinculo()
        self.carregar()

    def _is_programacao_vinculavel(self, status_raw, status_op_raw, prest_raw):
        st = upper(str(status_raw or "").strip())
        st_op = upper(str(status_op_raw or "").strip())
        prest = upper(str(prest_raw or "").strip())
        if st in {"EM ENTREGAS", "EM_ENTREGAS"} or st_op in {"EM ENTREGAS", "EM_ENTREGAS"}:
            return False
        if st_op in {"CANCELADA", "CANCELADO", "FINALIZADA", "FINALIZADO"}:
            return False
        if st in {"CANCELADA", "CANCELADO"}:
            return False
        if prest == "FECHADA":
            return False
        return (st in {"ATIVA", "ABERTA", "PROGRAMADA", "EM_ROTA", "EM ROTA", "INICIADA", "CARREGADA"} or
                st_op in {"ATIVA", "ABERTA", "PROGRAMADA", "EM_ROTA", "EM ROTA", "INICIADA", "CARREGADA"} or
                not (st or st_op))

    def _fetch_programacoes_vinculo(self):
        encontrados = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/programacoes?modo=todas&limit=300",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                for r in ((resp or {}).get("programacoes") or []):
                    if not isinstance(r, dict):
                        continue
                    codigo = upper((r or {}).get("codigo_programacao") or "")
                    if not codigo:
                        continue
                    if self._is_programacao_vinculavel(r.get("status"), r.get("status_operacional"), r.get("prestacao_status")):
                        encontrados.append(codigo)
                if encontrados:
                    return encontrados
            except Exception:
                logging.debug("Falha ao carregar programacoes vinculaveis via API.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(programacoes)")
            cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            status_expr = "COALESCE(status,'')" if "status" in cols else "''"
            status_op_expr = "COALESCE(status_operacional,'')" if "status_operacional" in cols else "''"
            prest_expr = "COALESCE(prestacao_status,'')" if "prestacao_status" in cols else "''"
            cur.execute(
                f"""
                SELECT COALESCE(codigo_programacao,''), {status_expr}, {status_op_expr}, {prest_expr}
                FROM programacoes
                WHERE COALESCE(codigo_programacao,'') <> ''
                ORDER BY id DESC
                LIMIT 300
                """
            )
            for r in (cur.fetchall() or []):
                codigo = upper((r[0] if r else "") or "")
                if not codigo:
                    continue
                if self._is_programacao_vinculavel(r[1] if len(r) > 1 else "", r[2] if len(r) > 2 else "", r[3] if len(r) > 3 else ""):
                    encontrados.append(codigo)
        return list(dict.fromkeys(encontrados))

    def _apply_programacoes_vinculo(self, seq, values):
        if seq != self._vinc_seq or not self.winfo_exists():
            return
        atual = upper((self.cb_prog_vinculo.get() or "").strip())
        values = values or []
        self.cb_prog_vinculo["values"] = values
        if atual in values:
            self.cb_prog_vinculo.set(atual)
        elif values:
            self.cb_prog_vinculo.set(values[0])
        else:
            self.cb_prog_vinculo.set("")

    def _handle_programacoes_vinculo_error(self, seq, exc):
        if seq != self._vinc_seq or not self.winfo_exists():
            return
        logging.error("Falha ao carregar programacoes para vinculo", exc_info=(type(exc), exc, exc.__traceback__))

    def refresh_programacoes_vinculo(self):
        self._vinc_seq += 1
        seq = self._vinc_seq
        run_async_ui(
            self,
            self._fetch_programacoes_vinculo,
            lambda values, seq=seq: self._apply_programacoes_vinculo(seq, values),
            lambda exc, seq=seq: self._handle_programacoes_vinculo_error(seq, exc),
        )

    def _collect_vinculo_target_ids(self):
        itens = tuple(self.tree.selection() or ())
        if not itens:
            marcados = []
            for iid in (self.tree.get_children() or ()):
                vals = self.tree.item(iid, "values") or ()
                if len(vals) > 1 and str(vals[1]).strip() == "[x]":
                    marcados.append(iid)
            itens = tuple(marcados)
        if not itens:
            return []

        ids = []
        seen = set()
        for iid in itens:
            vals = self.tree.item(iid, "values") or ()
            rid = safe_int(vals[0] if vals else 0, 0)
            if rid > 0 and rid not in seen:
                seen.add(rid)
                ids.append(rid)
        return ids

    def _fetch_vinculo_rows_from_tree(self, ids):
        wanted = {safe_int(x, 0) for x in (ids or []) if safe_int(x, 0) > 0}
        if not wanted:
            return []

        rows = []
        for iid in (self.tree.get_children() or ()):
            vals = self.tree.item(iid, "values") or ()
            rid = safe_int(vals[0] if vals else 0, 0)
            if rid not in wanted:
                continue
            rows.append(
                {
                    "id": rid,
                    "pedido": str(vals[2] if len(vals) > 2 else ""),
                    "data_venda": str(vals[3] if len(vals) > 3 else ""),
                    "cliente": str(vals[4] if len(vals) > 4 else ""),
                    "nome_cliente": str(vals[5] if len(vals) > 5 else ""),
                    "produto": str(vals[6] if len(vals) > 6 else ""),
                    "vr_total": safe_float(vals[7] if len(vals) > 7 else 0, 0.0),
                    "qnt": safe_float(vals[8] if len(vals) > 8 else 0, 0.0),
                    "cidade": str(vals[9] if len(vals) > 9 else ""),
                    "vendedor": str(vals[10] if len(vals) > 10 else ""),
                }
            )
        return rows

    def _fetch_programacao_vinculo_snapshot(self, codigo_programacao):
        codigo = upper(str(codigo_programacao or "").strip())
        if not codigo:
            return None

        usuario = upper(str((getattr(self.app, "user", {}) or {}).get("nome", "")).strip()) or "ADMIN"
        meta = {
            "codigo_programacao": codigo,
            "data_criacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "motorista": "",
            "motorista_id": 0,
            "motorista_codigo": "",
            "codigo_motorista": "",
            "veiculo": "",
            "equipe": "",
            "kg_estimado": 0.0,
            "tipo_estimativa": "KG",
            "caixas_estimado": 0,
            "status": "ATIVA",
            "status_operacional": "",
            "prestacao_status": "PENDENTE",
            "local_rota": "",
            "local_carregamento": "",
            "adiantamento": 0.0,
            "usuario_criacao": usuario,
            "usuario_ultima_edicao": usuario,
        }
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(codigo)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
                clientes = resp.get("clientes") if isinstance(resp, dict) else []
                if isinstance(rota, dict):
                    meta.update(
                        {
                            "data_criacao": str(rota.get("data_criacao") or rota.get("data") or meta["data_criacao"]),
                            "motorista": upper(str(rota.get("motorista") or "")),
                            "motorista_id": safe_int(rota.get("motorista_id"), 0),
                            "motorista_codigo": upper(str(rota.get("motorista_codigo") or rota.get("codigo_motorista") or "")),
                            "codigo_motorista": upper(str(rota.get("codigo_motorista") or rota.get("motorista_codigo") or "")),
                            "veiculo": upper(str(rota.get("veiculo") or "")),
                            "equipe": upper(str(rota.get("equipe") or "")),
                            "kg_estimado": safe_float(rota.get("kg_estimado"), 0.0),
                            "tipo_estimativa": upper(str(rota.get("tipo_estimativa") or "KG").strip()) or "KG",
                            "caixas_estimado": safe_int(rota.get("caixas_estimado"), 0),
                            "status": upper(str(rota.get("status") or "ATIVA").strip()) or "ATIVA",
                            "status_operacional": upper(str(rota.get("status_operacional") or "").strip()),
                            "prestacao_status": upper(str(rota.get("prestacao_status") or "PENDENTE").strip()) or "PENDENTE",
                            "local_rota": upper(str(rota.get("local_rota") or rota.get("tipo_rota") or "")),
                            "local_carregamento": upper(
                                str(
                                    rota.get("local_carregamento")
                                    or rota.get("granja_carregada")
                                    or rota.get("local_carregado")
                                    or rota.get("local_carreg")
                                    or ""
                                )
                            ),
                            "adiantamento": safe_float(rota.get("adiantamento"), safe_float(rota.get("adiantamento_rota"), 0.0)),
                            "usuario_criacao": upper(
                                str(
                                    rota.get("usuario_criacao")
                                    or rota.get("usuario")
                                    or rota.get("criado_por")
                                    or usuario
                                ).strip()
                            )
                            or usuario,
                            "usuario_ultima_edicao": upper(
                                str(rota.get("usuario_ultima_edicao") or usuario).strip()
                            )
                            or usuario,
                        }
                    )
                itens = []
                if isinstance(clientes, list):
                    itens = [
                        {
                            "cod_cliente": str(it.get("cod_cliente") or ""),
                            "nome_cliente": str(it.get("nome_cliente") or ""),
                            "produto": str(it.get("produto") or ""),
                            "endereco": str(it.get("endereco") or ""),
                            "qnt_caixas": safe_int(it.get("qnt_caixas"), 0),
                            "kg": safe_float(it.get("kg"), 0.0),
                            "preco": safe_float(it.get("preco"), 0.0),
                            "vendedor": str(it.get("vendedor") or ""),
                            "pedido": str(it.get("pedido") or ""),
                            "obs": str(it.get("obs") or it.get("observacao") or ""),
                        }
                        for it in clientes
                        if isinstance(it, dict)
                    ]
                return {"meta": meta, "itens": itens}
            except Exception:
                logging.debug("Falha ao carregar snapshot da programacao via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(programacoes)")
            cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            if "codigo_programacao" not in cols_prog:
                return None
            data_expr = "COALESCE(data_criacao,'')" if "data_criacao" in cols_prog else ("COALESCE(data,'')" if "data" in cols_prog else "''")
            select_cols = [
                data_expr,
                "COALESCE(motorista,'')",
                "COALESCE(veiculo,'')",
                "COALESCE(equipe,'')",
                "COALESCE(kg_estimado,0)",
                "COALESCE(tipo_estimativa,'KG')" if "tipo_estimativa" in cols_prog else "'KG'",
                "COALESCE(caixas_estimado,0)" if "caixas_estimado" in cols_prog else "0",
                "COALESCE(status,'ATIVA')" if "status" in cols_prog else "'ATIVA'",
                "COALESCE(status_operacional,'')" if "status_operacional" in cols_prog else "''",
                "COALESCE(prestacao_status,'PENDENTE')" if "prestacao_status" in cols_prog else "'PENDENTE'",
                "COALESCE(local_rota,'')" if "local_rota" in cols_prog else ("COALESCE(tipo_rota,'')" if "tipo_rota" in cols_prog else "''"),
                (
                    "COALESCE(local_carregamento,'')"
                    if "local_carregamento" in cols_prog
                    else (
                        "COALESCE(granja_carregada,'')"
                        if "granja_carregada" in cols_prog
                        else ("COALESCE(local_carregado,'')" if "local_carregado" in cols_prog else "''")
                    )
                ),
                "COALESCE(adiantamento,0)" if "adiantamento" in cols_prog else ("COALESCE(adiantamento_rota,0)" if "adiantamento_rota" in cols_prog else "0"),
                "COALESCE(motorista_id,0)" if "motorista_id" in cols_prog else "0",
                "COALESCE(motorista_codigo,'')" if "motorista_codigo" in cols_prog else ("COALESCE(codigo_motorista,'')" if "codigo_motorista" in cols_prog else "''"),
                "COALESCE(codigo_motorista,'')" if "codigo_motorista" in cols_prog else ("COALESCE(motorista_codigo,'')" if "motorista_codigo" in cols_prog else "''"),
                "COALESCE(usuario_criacao,'')" if "usuario_criacao" in cols_prog else "''",
                "COALESCE(usuario_ultima_edicao,'')" if "usuario_ultima_edicao" in cols_prog else "''",
            ]
            cur.execute(
                f"SELECT {', '.join(select_cols)} FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                (codigo,),
            )
            row = cur.fetchone()
            if not row:
                return None
            meta.update(
                {
                    "data_criacao": str(row[0] or meta["data_criacao"]),
                    "motorista": upper(str(row[1] or "")),
                    "veiculo": upper(str(row[2] or "")),
                    "equipe": upper(str(row[3] or "")),
                    "kg_estimado": safe_float(row[4], 0.0),
                    "tipo_estimativa": upper(str(row[5] or "KG").strip()) or "KG",
                    "caixas_estimado": safe_int(row[6], 0),
                    "status": upper(str(row[7] or "ATIVA").strip()) or "ATIVA",
                    "status_operacional": upper(str(row[8] or "").strip()),
                    "prestacao_status": upper(str(row[9] or "PENDENTE").strip()) or "PENDENTE",
                    "local_rota": upper(str(row[10] or "")),
                    "local_carregamento": upper(str(row[11] or "")),
                    "adiantamento": safe_float(row[12], 0.0),
                    "motorista_id": safe_int(row[13], 0),
                    "motorista_codigo": upper(str(row[14] or "")),
                    "codigo_motorista": upper(str(row[15] or "")),
                    "usuario_criacao": upper(str(row[16] or usuario).strip()) or usuario,
                    "usuario_ultima_edicao": upper(str(row[17] or usuario).strip()) or usuario,
                }
            )
        itens_res = fetch_programacao_itens_contract(codigo)
        itens = itens_res.get("data") if isinstance(itens_res, dict) and bool(itens_res.get("ok", False)) else []
        return {"meta": meta, "itens": itens or []}

    def _resolve_clientes_endereco_map(self, codigos):
        codes = sorted({upper(str(c or "").strip()) for c in (codigos or []) if str(c or "").strip()})
        if not codes:
            return {}
        placeholders = ",".join(["?"] * len(codes))
        endereco_map = {}
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    f"SELECT UPPER(COALESCE(cod_cliente,'')), UPPER(COALESCE(endereco,'')) FROM clientes WHERE UPPER(COALESCE(cod_cliente,'')) IN ({placeholders})",
                    tuple(codes),
                )
                for cod, endereco in (cur.fetchall() or []):
                    endereco_map[upper(cod or "")] = upper(endereco or "")
        except Exception:
            logging.debug("Falha ao resolver enderecos dos clientes para vinculo.", exc_info=True)
        return endereco_map

    def _row_venda_to_programacao_item(self, row, endereco_map):
        cliente = upper(str((row or {}).get("cliente") or "").strip())
        pedido = upper(str((row or {}).get("pedido") or "").strip())
        nome_cliente = upper(str((row or {}).get("nome_cliente") or "").strip())
        produto = upper(str((row or {}).get("produto") or "").strip())
        cidade = upper(str((row or {}).get("cidade") or "").strip())
        vendedor = upper(str((row or {}).get("vendedor") or "").strip())
        qnt = safe_float((row or {}).get("qnt"), 0.0)
        vr_total = safe_float((row or {}).get("vr_total"), 0.0)
        caixas = safe_int((row or {}).get("qnt_caixas_vinculo"), -1)
        if caixas <= 0:
            caixas = max(safe_int(round(qnt), 0), 1 if qnt > 0 else 1)
        preco = (vr_total / qnt) if qnt > 0 else 0.0
        endereco = upper(str(endereco_map.get(cliente) or cidade or "").strip())
        return {
            "cod_cliente": cliente,
            "nome_cliente": nome_cliente,
            "produto": produto,
            "endereco": endereco,
            "qnt_caixas": caixas,
            "kg": 0.0,
            "preco": safe_float(preco, 0.0),
            "vendedor": vendedor,
            "pedido": pedido,
            "obs": "",
        }

    def _prompt_caixas_vinculo_row(self, row, index_atual, total_itens):
        result = {"ok": False, "value": None}
        default_qtd = safe_int((row or {}).get("qnt"), 0)
        if default_qtd <= 0:
            default_qtd = 1

        win = tk.Toplevel(self)
        win.title(f"Vincular Cliente {index_atual}/{total_itens}")
        win.transient(self.winfo_toplevel())
        win.resizable(False, False)
        win.grab_set()

        box = ttk.Frame(win, style="Card.TFrame", padding=16)
        box.grid(row=0, column=0, sticky="nsew")
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(0, weight=1)
        box.grid_columnconfigure(1, weight=1)

        nome = upper(str((row or {}).get("nome_cliente") or ""))
        vendedor = upper(str((row or {}).get("vendedor") or ""))
        produto = upper(str((row or {}).get("produto") or ""))
        preco_total = safe_float((row or {}).get("vr_total"), 0.0)
        qtd_excel = safe_float((row or {}).get("qnt"), 0.0)
        preco_unit = (preco_total / qtd_excel) if qtd_excel > 0 else 0.0

        ttk.Label(box, text=f"Cliente {index_atual} de {total_itens}", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(box, text="Nome:", style="CardLabel.TLabel").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Label(box, text=nome or "-", style="CardValue.TLabel").grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(box, text="Vendedor:", style="CardLabel.TLabel").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Label(box, text=vendedor or "-", style="CardValue.TLabel").grid(row=2, column=1, sticky="w", pady=3)
        ttk.Label(box, text="Produto:", style="CardLabel.TLabel").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Label(box, text=produto or "-", style="CardValue.TLabel").grid(row=3, column=1, sticky="w", pady=3)
        ttk.Label(box, text="Preço:", style="CardLabel.TLabel").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Label(
            box,
            text=f"R$ {preco_unit:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            style="CardValue.TLabel",
        ).grid(row=4, column=1, sticky="w", pady=3)
        ttk.Label(box, text="Quantidade de caixas:", style="CardLabel.TLabel").grid(row=5, column=0, sticky="w", pady=(10, 3))
        ent_qtd = ttk.Entry(box, style="Field.TEntry", width=12)
        ent_qtd.grid(row=5, column=1, sticky="w", pady=(10, 3))
        ent_qtd.insert(0, str(default_qtd))
        ent_qtd.focus_set()
        ent_qtd.selection_range(0, "end")

        btns = ttk.Frame(box, style="Card.TFrame")
        btns.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)

        def _confirm():
            qtd = safe_int(ent_qtd.get(), 0)
            if qtd <= 0:
                messagebox.showwarning("ATENÇÃO", "Informe uma quantidade de caixas maior que zero.", parent=win)
                return
            result["ok"] = True
            result["value"] = qtd
            win.destroy()

        def _cancel():
            result["ok"] = False
            result["value"] = None
            win.destroy()

        ttk.Button(btns, text="CANCELAR", style="Ghost.TButton", command=_cancel).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btns, text="CONFIRMAR", style="Primary.TButton", command=_confirm).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        ent_qtd.bind("<Return>", lambda e: _confirm())
        ent_qtd.bind("<Escape>", lambda e: _cancel())
        win.protocol("WM_DELETE_WINDOW", _cancel)
        try:
            win.update_idletasks()
            px = self.winfo_rootx() + max((self.winfo_width() - win.winfo_reqwidth()) // 2, 40)
            py = self.winfo_rooty() + max((self.winfo_height() - win.winfo_reqheight()) // 2, 40)
            win.geometry(f"+{px}+{py}")
        except Exception:
            logging.debug("Falha ignorada ao centralizar dialogo de caixas.")
        win.wait_window()
        return result["value"] if result["ok"] else None

    def _collect_vinculo_caixas(self, rows):
        prepared = []
        total = len(rows or [])
        for idx, row in enumerate((rows or []), start=1):
            qtd = self._prompt_caixas_vinculo_row(row, idx, total)
            if qtd is None:
                return None
            novo = dict(row or {})
            novo["qnt_caixas_vinculo"] = qtd
            prepared.append(novo)
        return prepared

    def _merge_programacao_items(self, itens_existentes, novos_itens):
        merged = []
        seen = set()
        for src in list(itens_existentes or []) + list(novos_itens or []):
            item = {
                "cod_cliente": upper(str((src or {}).get("cod_cliente") or "").strip()),
                "nome_cliente": upper(str((src or {}).get("nome_cliente") or "").strip()),
                "produto": upper(str((src or {}).get("produto") or "").strip()),
                "endereco": upper(str((src or {}).get("endereco") or "").strip()),
                "qnt_caixas": safe_int((src or {}).get("qnt_caixas"), 0),
                "kg": safe_float((src or {}).get("kg"), 0.0),
                "preco": safe_float((src or {}).get("preco"), 0.0),
                "vendedor": upper(str((src or {}).get("vendedor") or "").strip()),
                "pedido": upper(str((src or {}).get("pedido") or "").strip()),
                "obs": upper(str((src or {}).get("obs") or (src or {}).get("observacao") or "").strip()),
            }
            key = (item["cod_cliente"], item["pedido"], item["produto"])
            if not item["cod_cliente"] or not item["nome_cliente"] or key in seen:
                continue
            seen.add(key)
            merged.append(item)
        return merged

    def _persist_vinculo_programacao_local(self, meta, itens, ids, usada_em):
        codigo = upper(str((meta or {}).get("codigo_programacao") or "").strip())
        with get_db() as conn:
            cur = conn.cursor()
            existentes = set()
            try:
                cur.execute(
                    """
                    SELECT UPPER(COALESCE(cod_cliente,'')), UPPER(COALESCE(pedido,'')), UPPER(COALESCE(produto,''))
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                    """,
                    (codigo,),
                )
                existentes = {
                    (upper(r[0] or ""), upper(r[1] or ""), upper(r[2] or ""))
                    for r in (cur.fetchall() or [])
                }
            except Exception:
                logging.debug("Falha ao consultar itens existentes da programacao local.", exc_info=True)

            for item in (itens or []):
                key = (
                    upper(str(item.get("cod_cliente") or "").strip()),
                    upper(str(item.get("pedido") or "").strip()),
                    upper(str(item.get("produto") or "").strip()),
                )
                if key in existentes:
                    continue
                cur.execute(
                    """
                    INSERT INTO programacao_itens
                        (codigo_programacao, cod_cliente, nome_cliente, qnt_caixas, kg, preco, endereco, vendedor, pedido, produto, observacao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        codigo,
                        key[0],
                        upper(str(item.get("nome_cliente") or "").strip()),
                        safe_int(item.get("qnt_caixas"), 0),
                        safe_float(item.get("kg"), 0.0),
                        safe_float(item.get("preco"), 0.0),
                        upper(str(item.get("endereco") or "").strip()),
                        upper(str(item.get("vendedor") or "").strip()),
                        key[1],
                        key[2],
                        upper(str(item.get("obs") or item.get("observacao") or "").strip()),
                    ),
                )
                existentes.add(key)

            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                sets = []
                vals = []
                total_caixas = sum(safe_int(it.get("qnt_caixas"), 0) for it in (itens or []))
                total_quilos = round(sum(safe_float(it.get("kg"), 0.0) for it in (itens or [])), 2)
                if "total_caixas" in cols_prog:
                    sets.append("total_caixas=?")
                    vals.append(total_caixas)
                if "quilos" in cols_prog:
                    sets.append("quilos=?")
                    vals.append(total_quilos)
                if "status" in cols_prog:
                    sets.append("status='ATIVA'")
                if "usuario_ultima_edicao" in cols_prog:
                    sets.append("usuario_ultima_edicao=?")
                    vals.append(upper(str((meta or {}).get("usuario_ultima_edicao") or "").strip()))
                if sets:
                    vals.append(codigo)
                    cur.execute(f"UPDATE programacoes SET {', '.join(sets)} WHERE codigo_programacao=?", tuple(vals))
            except Exception:
                logging.debug("Falha ao atualizar totais da programacao local apos vinculo.", exc_info=True)

            if ids:
                cur.executemany(
                    """
                    UPDATE vendas_importadas
                       SET usada=1,
                           usada_em=?,
                           codigo_programacao=?,
                           selecionada=0
                     WHERE id=? AND IFNULL(usada,0)=0
                    """,
                    [(usada_em, codigo, rid) for rid in ids],
                )

    def _legacy_vincular_selecionadas_programacao_api_vinculo(self):
        # Mantido apenas por compatibilidade com binds legados.
        return self.vincular_selecionadas_programacao()

    # -------------------------
    # Helpers segurança/normalização
    # -------------------------
    def _norm(self, v):
        return upper(str(v or "").strip())

    def _legacy_vincular_selecionadas_programacao_merge_sem_caixas(self):
        # Mantido apenas por compatibilidade com binds legados.
        return self.vincular_selecionadas_programacao()

    def vincular_selecionadas_programacao(self):
        prog = upper((self.cb_prog_vinculo.get() or "").strip())
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione a programacao ativa para vincular as vendas.")
            return

        ids = self._collect_vinculo_target_ids()
        if not ids:
            messagebox.showwarning("ATENCAO", "Selecione uma ou mais vendas para vincular.")
            return

        rows = self._fetch_vinculo_rows_from_tree(ids)
        if not rows:
            messagebox.showwarning("ATENCAO", "Nao foi possivel carregar os dados das vendas selecionadas.")
            return

        snapshot = self._fetch_programacao_vinculo_snapshot(prog)
        if not snapshot:
            messagebox.showwarning("ATENCAO", f"Programacao nao encontrada para vinculo: {prog}.")
            return

        meta = dict((snapshot or {}).get("meta") or {})
        if not self._is_programacao_vinculavel(
            meta.get("status"),
            meta.get("status_operacional"),
            meta.get("prestacao_status"),
        ):
            messagebox.showwarning(
                "ATENCAO",
                f"A programacao {prog} nao pode receber vendas porque esta em rota de entregas.",
            )
            return

        rows = self._collect_vinculo_caixas(rows)
        if rows is None:
            self.set_status("STATUS: Vinculo cancelado pelo usuario antes de definir as caixas.")
            return

        endereco_map = self._resolve_clientes_endereco_map([(r or {}).get("cliente") for r in rows])
        novos_itens = [self._row_venda_to_programacao_item(row, endereco_map) for row in (rows or [])]
        itens_merged = self._merge_programacao_items((snapshot or {}).get("itens") or [], novos_itens)
        if not itens_merged:
            messagebox.showwarning("ATENCAO", "Nenhum item valido foi preparado para a programacao.")
            return

        meta["codigo_programacao"] = prog
        total_caixas = sum(safe_int(it.get("qnt_caixas"), 0) for it in itens_merged)
        total_quilos = round(sum(safe_float(it.get("kg"), 0.0) for it in itens_merged), 2)
        usada_em = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()

        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                _call_api(
                    "POST",
                    "desktop/rotas/upsert",
                    payload={
                        "codigo_programacao": prog,
                        "data_criacao": str(meta.get("data_criacao") or usada_em),
                        "motorista": upper(str(meta.get("motorista") or "")),
                        "motorista_id": safe_int(meta.get("motorista_id"), 0),
                        "motorista_codigo": upper(str(meta.get("motorista_codigo") or meta.get("codigo_motorista") or "")),
                        "codigo_motorista": upper(str(meta.get("codigo_motorista") or meta.get("motorista_codigo") or "")),
                        "veiculo": upper(str(meta.get("veiculo") or "")),
                        "equipe": upper(str(meta.get("equipe") or "")),
                        "kg_estimado": safe_float(meta.get("kg_estimado"), 0.0),
                        "tipo_estimativa": upper(str(meta.get("tipo_estimativa") or "KG").strip()) or "KG",
                        "caixas_estimado": safe_int(meta.get("caixas_estimado"), 0),
                        "status": upper(str(meta.get("status") or "ATIVA").strip()) or "ATIVA",
                        "local_rota": upper(str(meta.get("local_rota") or "")),
                        "local_carregamento": upper(str(meta.get("local_carregamento") or "")),
                        "adiantamento": safe_float(meta.get("adiantamento"), 0.0),
                        "total_caixas": total_caixas,
                        "quilos": total_quilos,
                        "usuario_criacao": upper(str(meta.get("usuario_criacao") or "").strip()),
                        "usuario_ultima_edicao": upper(str(meta.get("usuario_ultima_edicao") or meta.get("usuario_criacao") or "").strip()),
                        "linked_venda_ids": ids,
                        "vendas_usada_em": usada_em,
                        "itens": itens_merged,
                    },
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
            except Exception as exc:
                logging.debug("Falha ao anexar vendas diretamente a programacao via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel anexar as vendas a programacao na API central.\n\n"
                    "A gravacao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                    f"Detalhe: {exc}",
                )
                return
        else:
            self._persist_vinculo_programacao_local(meta, itens_merged, ids, usada_em)

        self.carregar()
        self.refresh_programacoes_vinculo()
        self.set_status(
            f"STATUS: {len(ids)} venda(s) anexada(s) a programacao {prog}. "
            "As vendas vinculadas sairam da lista e a rota ja pode refletir os clientes."
        )

    def _excel_text(self, v):
        """Normaliza valor textual vindo do Excel, removendo NaN/NaT e espaços."""
        try:
            s = str(v).strip()
        except Exception:
            return ""
        su = upper(s)
        if su in {"", "NAN", "NAT", "NONE", "NULL", "<NA>"}:
            return ""
        return s

    def _is_invalid_token(self, v):
        return upper(str(v or "").strip()) in {"", "NAN", "NAT", "NONE", "NULL", "<NA>"}

    def _clean_pedido(self, v):
        """
        Pedido costuma vir como float (ex.: 46522.0). Converte para inteiro textual quando aplicável.
        """
        s = self._excel_text(v)
        if not s:
            return ""
        try:
            f = float(str(s).replace(",", "."))
            if abs(f - int(f)) < 1e-9:
                return str(int(f))
            return str(f).rstrip("0").rstrip(".")
        except Exception:
            return s

    def _normalize_data_venda(self, v):
        """
        Normaliza data da venda para chave estavel (YYYY-MM-DD quando possivel).
        Evita duplicidade falsa por diferença de formato no Excel.
        """
        raw = self._excel_text(v)
        if not raw:
            return ""
        nd = normalize_date(raw)
        if nd:
            return nd
        return raw

    def _ensure_vendas_usada_cols(self):
        """Garante colunas para evitar reutilização (não quebra bases antigas)."""
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS vendas_importadas (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        pedido TEXT,
                        data_venda TEXT,
                        cliente TEXT,
                        nome_cliente TEXT,
                        vendedor TEXT,
                        produto TEXT,
                        vr_total REAL,
                        qnt REAL,
                        cidade TEXT,
                        valor_unitario REAL,
                        observacao TEXT,
                        selecionada INTEGER DEFAULT 0,
                        usada INTEGER DEFAULT 0,
                        usada_em TEXT,
                        codigo_programacao TEXT
                    )
                """)
                cur.execute("PRAGMA table_info(vendas_importadas)")
                cols = [str(r[1]).lower() for r in cur.fetchall()]

                if "usada" not in cols:
                    cur.execute("ALTER TABLE vendas_importadas ADD COLUMN usada INTEGER DEFAULT 0")
                if "usada_em" not in cols:
                    cur.execute("ALTER TABLE vendas_importadas ADD COLUMN usada_em TEXT")
                if "codigo_programacao" not in cols:
                    cur.execute("ALTER TABLE vendas_importadas ADD COLUMN codigo_programacao TEXT")
        except Exception as e:
            logging.exception("Falha ao garantir colunas de vendas_importadas: %s", e)

    def _importar_vendas_worker(self, path):
        df = pd.read_excel(path, engine=excel_engine_for(path))

        col_pedido = guess_col(df.columns, ["numero pedido", "num pedido", "n pedido", "pedido"])
        col_data = guess_col(df.columns, ["data venda", "data", "dt"])
        col_cliente = guess_col(df.columns, ["cod cliente", "codigo cliente", "cliente", "cod"])
        col_nome = guess_col(df.columns, ["nome completo", "nome cliente", "razao", "nome"])
        col_prod = guess_col(df.columns, ["descricao do produto", "produto", "descr", "item"])
        col_vr_total = guess_col(df.columns, ["vr. total", "vr total", "valor total", "total"])
        col_qnt = guess_col(df.columns, ["qnt", "qtd", "quantidade"])
        col_cidade = guess_col(df.columns, ["cidade", "municipio"])
        col_vend = guess_col(df.columns, ["nome do vendedor", "vendedor", "vend"])
        col_obs = guess_col(df.columns, ["obs", "observ", "observacao"])

        missing = []
        if not col_pedido:
            missing.append("Numero Pedido")
        if not col_cliente:
            missing.append("Cliente")
        if not col_nome:
            missing.append("Nome Completo")
        if not col_prod:
            missing.append("Descricao do Produto")
        if not col_vr_total:
            missing.append("Vr. Total")
        if not col_qnt:
            missing.append("Qnt.")
        if missing:
            raise ValueError("Nao identifiquei as colunas: " + ", ".join(missing))

        opcionais_ausentes = []
        if not col_data:
            opcionais_ausentes.append("Data")
        if not col_cidade:
            opcionais_ausentes.append("Cidade")
        if not col_vend:
            opcionais_ausentes.append("Nome do Vendedor")
        if not col_obs:
            opcionais_ausentes.append("Observacao")

        total = 0
        ignoradas = 0
        ignoradas_invalidas = 0
        payload_rows = []

        self._ensure_vendas_usada_cols()

        for _, r in df.iterrows():
            pedido = self._clean_pedido(r.get(col_pedido, ""))
            if self._is_invalid_token(pedido):
                ignoradas_invalidas += 1
                continue

            data_venda = self._normalize_data_venda(r.get(col_data, "")) if col_data else ""
            cliente = self._excel_text(r.get(col_cliente, ""))
            nome_cliente = self._excel_text(r.get(col_nome, "")) if col_nome else ""
            vendedor = self._excel_text(r.get(col_vend, "")) if col_vend else ""
            produto = self._excel_text(r.get(col_prod, "")) if col_prod else ""
            vr_total = safe_float(r.get(col_vr_total, 0)) if col_vr_total else 0.0
            qnt = safe_float(r.get(col_qnt, 0)) if col_qnt else 0.0
            cidade = self._excel_text(r.get(col_cidade, "")) if col_cidade else ""
            valor_unit = (vr_total / qnt) if qnt else 0.0
            obs = self._excel_text(r.get(col_obs, "")) if col_obs else ""

            pedido_u = self._norm(pedido)
            cliente_u = self._norm(cliente)
            produto_u = self._norm(produto)
            nome_u = self._norm(nome_cliente)
            if (not pedido_u) or (not cliente_u) or (not nome_u) or (not produto_u):
                ignoradas_invalidas += 1
                continue

            payload_rows.append(
                {
                    "pedido": pedido_u,
                    "data_venda": data_venda,
                    "cliente": cliente_u,
                    "nome_cliente": nome_u,
                    "vendedor": self._norm(vendedor),
                    "produto": produto_u,
                    "vr_total": float(vr_total or 0),
                    "qnt": float(qnt or 0),
                    "cidade": self._norm(cidade),
                    "valor_unitario": float(valor_unit or 0),
                    "observacao": self._norm(obs),
                }
            )

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            resp = _call_api(
                "POST",
                "desktop/vendas-importadas/importar",
                payload={"rows": payload_rows},
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            total = safe_int((resp or {}).get("importadas"), 0)
            ignoradas += safe_int((resp or {}).get("ignoradas"), 0)
        else:
            with get_db() as conn:
                cur = conn.cursor()
                existing_keys = set()
                try:
                    cur.execute("""
                        SELECT UPPER(TRIM(COALESCE(pedido,''))),
                               UPPER(TRIM(COALESCE(cliente,''))),
                               UPPER(TRIM(COALESCE(produto,''))),
                               COALESCE(TRIM(data_venda),'')
                        FROM vendas_importadas
                    """)
                    existing_keys = {
                        (
                            str(r[0] or ""),
                            str(r[1] or ""),
                            str(r[2] or ""),
                            str(r[3] or ""),
                        )
                        for r in (cur.fetchall() or [])
                    }
                except Exception:
                    existing_keys = set()
                for row in payload_rows:
                    data_venda = str(row.get("data_venda") or "")
                    pedido_u = str(row.get("pedido") or "")
                    cliente_u = str(row.get("cliente") or "")
                    produto_u = str(row.get("produto") or "")
                    key = (
                        upper(pedido_u.strip()),
                        upper(cliente_u.strip()),
                        upper(produto_u.strip()),
                        data_venda.strip(),
                    )
                    if key in existing_keys:
                        ignoradas += 1
                        continue
                    cur.execute("""
                        INSERT INTO vendas_importadas
                        (pedido, data_venda, cliente, nome_cliente, vendedor, produto, vr_total, qnt, cidade, valor_unitario, observacao, selecionada, usada, usada_em, codigo_programacao)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, '', '')
                    """, (
                        pedido_u,
                        data_venda,
                        cliente_u,
                        str(row.get("nome_cliente") or ""),
                        str(row.get("vendedor") or ""),
                        produto_u,
                        float(row.get("vr_total") or 0),
                        float(row.get("qnt") or 0),
                        str(row.get("cidade") or ""),
                        float(row.get("valor_unitario") or 0),
                        str(row.get("observacao") or ""),
                    ))
                    existing_keys.add(key)
                    total += 1

        msg = f"Vendas importadas: {total}"
        if ignoradas:
            msg += f"\nIgnoradas (duplicadas/invalidas): {ignoradas}"
        if ignoradas_invalidas:
            msg += f"\nIgnoradas (linhas invalidas/NaN): {ignoradas_invalidas}"
        if opcionais_ausentes:
            msg += "\nCampos opcionais nao encontrados (preenchidos em branco): " + ", ".join(opcionais_ausentes)
        return msg

    def _on_importar_vendas_done(self, seq, msg):
        if seq != self._import_seq or not self.winfo_exists():
            return
        self._set_busy(False, "Importação concluída")
        if hasattr(self, "ent_busca"):
            try:
                self.ent_busca.delete(0, "end")
            except Exception:
                logging.debug("Falha ao limpar busca apos importacao")
        messagebox.showinfo("OK", msg)
        self.carregar()

    def _on_importar_vendas_error(self, seq, exc):
        if seq != self._import_seq or not self.winfo_exists():
            return
        self._set_busy(False, "Falha na importação")
        logging.error("Falha ao importar vendas via Excel", exc_info=(type(exc), exc, exc.__traceback__))
        self.set_status("STATUS: Falha ao importar vendas.")
        messagebox.showerror("ERRO", str(exc))

    def importar_excel(self):
        path = filedialog.askopenfilename(
            title="IMPORTAR VENDAS (EXCEL)",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if not path:
            return

        if not (require_pandas() and require_excel_support(path)):
            return

        self._import_seq += 1
        seq = self._import_seq
        self._set_busy(True, "Importando vendas do Excel...")
        self.set_status("STATUS: Importando vendas do Excel...")
        run_async_ui(
            self,
            lambda path=path: self._importar_vendas_worker(path),
            lambda msg, seq=seq: self._on_importar_vendas_done(seq, msg),
            lambda exc, seq=seq: self._on_importar_vendas_error(seq, exc),
        )

    def _fetch_vendas_rows(self, busca):
        self._ensure_vendas_usada_cols()

        rows = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                path = f"desktop/vendas-importadas?limit=5000"
                if busca:
                    path += f"&busca={urllib.parse.quote(busca)}"
                resp = _call_api(
                    "GET",
                    path,
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rows_api = resp.get("rows") if isinstance(resp, dict) else []
                if isinstance(rows_api, list):
                    rows = [
                        (
                            safe_int(r.get("id"), 0),
                            safe_int(r.get("selecionada"), 0),
                            r.get("pedido"),
                            r.get("data_venda"),
                            r.get("cliente"),
                            r.get("nome_cliente"),
                            r.get("produto"),
                            safe_float(r.get("vr_total"), 0.0),
                            safe_float(r.get("qnt"), 0.0),
                            r.get("cidade"),
                            r.get("vendedor"),
                            r.get("codigo_programacao"),
                        )
                        for r in rows_api
                        if isinstance(r, dict)
                    ]
            except Exception:
                logging.debug("Falha ao carregar vendas importadas via API; usando fallback local.", exc_info=True)

        if not rows:
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    if busca:
                        cur.execute("""
                            SELECT id, selecionada, pedido, data_venda, cliente, nome_cliente, produto, vr_total, qnt, cidade, vendedor, codigo_programacao
                            FROM vendas_importadas
                            WHERE (IFNULL(usada,0)=0)
                              AND TRIM(COALESCE(codigo_programacao,''))=''
                              AND (
                                pedido LIKE ? OR cliente LIKE ? OR nome_cliente LIKE ? OR vendedor LIKE ? OR produto LIKE ?
                              )
                            ORDER BY id DESC
                        """, (f"%{busca}%", f"%{busca}%", f"%{busca}%", f"%{busca}%", f"%{busca}%"))
                    else:
                        cur.execute("""
                            SELECT id, selecionada, pedido, data_venda, cliente, nome_cliente, produto, vr_total, qnt, cidade, vendedor, codigo_programacao
                            FROM vendas_importadas
                            WHERE (IFNULL(usada,0)=0)
                              AND TRIM(COALESCE(codigo_programacao,''))=''
                            ORDER BY id DESC
                        """)
                    rows = cur.fetchall() or []
            except Exception:
                logging.debug("Falha ao carregar vendas importadas no fallback local.", exc_info=True)
                rows = []
        return rows

    def _apply_vendas_rows(self, seq, rows):
        if seq != self._load_seq or not self.winfo_exists():
            return
        self._set_busy(False, f"Lista carregada: {len(rows)} vendas")

        self.tree.delete(*self.tree.get_children())

        selected_count = 0
        for row in rows:
            rid = row[0]
            selecionada = 1 if row[1] == 1 else 0
            if selecionada:
                selected_count += 1

            sel = "[x]" if selecionada else "[ ]"
            valor = row[7] if row[7] is not None else 0
            try:
                valor_txt = f"{float(valor):.3f}"
            except Exception:
                valor_txt = "0.000"

            qnt_val = row[8] if row[8] is not None else 0
            try:
                qnt_txt = f"{float(qnt_val):.3f}"
            except Exception:
                qnt_txt = "0.000"

            tree_insert_aligned(
                self.tree,
                "",
                "end",
                (rid, sel, row[2], row[3], row[4], row[5], row[6], valor_txt, qnt_txt, row[9], row[10]),
            )

        self.set_status(f"STATUS: {len(rows)} registros carregados (NAO usadas)  Selecionadas: {selected_count}.")

    def _handle_vendas_load_error(self, seq, exc):
        if seq != self._load_seq or not self.winfo_exists():
            return
        self._set_busy(False, "Falha ao carregar vendas")
        logging.error("Falha ao carregar vendas importadas", exc_info=(type(exc), exc, exc.__traceback__))
        self.set_status("STATUS: Falha ao carregar vendas importadas.")

    def carregar(self, async_load=True):
        busca = self._norm(self.ent_busca.get()) if hasattr(self, "ent_busca") else ""
        self._load_seq += 1
        seq = self._load_seq
        self._set_busy(True, "Carregando lista de vendas...")
        self.set_status("STATUS: Carregando vendas importadas...")
        if not async_load:
            try:
                rows = self._fetch_vendas_rows(busca)
            except Exception as exc:
                self._handle_vendas_load_error(seq, exc)
                return
            self._apply_vendas_rows(seq, rows)
            return

        run_async_ui(
            self,
            lambda busca=busca: self._fetch_vendas_rows(busca),
            lambda rows, seq=seq: self._apply_vendas_rows(seq, rows),
            lambda exc, seq=seq: self._handle_vendas_load_error(seq, exc),
        )

    def toggle_selected(self, event=None):
        # âÅâ€œââ‚¬¦ só alterna se clicar em uma célula (evita bug ao clicar em cabeçalho)
        if event is not None:
            region = self.tree.identify("region", event.x, event.y)
            if region != "cell":
                return

        item = self.tree.selection()
        if not item:
            return

        self._toggle_selected_iid(item[0])

        self.carregar()

    def _shortcut_select_all_tree(self, event=None):
        try:
            items = self.tree.get_children()
            if items:
                self.tree.selection_set(items)
                self.tree.focus(items[0])
        except Exception:
            logging.debug("Falha ao selecionar todas as vendas importadas.", exc_info=True)
        return "break"

    def _shortcut_toggle_selected_rows(self, event=None):
        itens = tuple(self.tree.selection() or ())
        if not itens:
            focus = self.tree.focus()
            if focus:
                itens = (focus,)
        if not itens:
            return "break"

        for iid in itens:
            self._toggle_selected_iid(iid)
        self.carregar()
        self.set_status(f"STATUS: Marcacao atualizada para {len(itens)} venda(s).")
        return "break"

    def _shortcut_delete_selected_rows(self, event=None):
        itens = tuple(self.tree.selection() or ())
        if not itens:
            focus = self.tree.focus()
            if focus:
                itens = (focus,)
        if not itens:
            return "break"

        ids = []
        seen = set()
        for iid in itens:
            vals = self.tree.item(iid, "values") or ()
            if not vals:
                continue
            rid = safe_int(vals[0], 0)
            if rid > 0 and rid not in seen:
                seen.add(rid)
                ids.append(rid)

        if not ids:
            return "break"

        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja excluir {len(ids)} venda(s) importada(s) selecionada(s)?",
        ):
            return "break"

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                ids_csv = ",".join(str(x) for x in ids)
                _call_api(
                    "DELETE",
                    f"desktop/vendas-importadas/ids?ids={urllib.parse.quote(ids_csv)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                self.carregar()
                self.set_status(f"STATUS: {len(ids)} venda(s) importada(s) excluida(s).")
                return "break"
            except Exception as exc:
                logging.debug("Falha ao excluir vendas importadas via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel excluir as vendas na API central.\n\n"
                    "A exclusao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                    f"Detalhe: {exc}",
                )
                return "break"

        with get_db() as conn:
            cur = conn.cursor()
            cur.executemany(
                "DELETE FROM vendas_importadas WHERE id=? AND IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''",
                [(rid,) for rid in ids],
            )

        self.carregar()
        self.set_status(f"STATUS: {len(ids)} venda(s) importada(s) excluida(s).")
        return "break"

    def _toggle_selected_iid(self, iid):
        vals = self.tree.item(iid, "values") or ()
        if not vals:
            return
        rid = safe_int(vals[0], 0)
        if rid <= 0:
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                _call_api(
                    "POST",
                    f"desktop/vendas-importadas/{rid}/toggle-selecao",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                return
            except Exception as exc:
                logging.debug("Falha ao alternar selecao via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel atualizar a selecao na API central.\n\n"
                    "A alteracao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                    f"Detalhe: {exc}",
                )
                return

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE vendas_importadas
                SET selecionada = CASE WHEN selecionada=1 THEN 0 ELSE 1 END
                WHERE id=? AND IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''
            """, (rid,))

    def _on_tree_click_toggle_sel(self, event=None):
        """
        Clique simples na coluna SEL alterna marcação da linha.
        """
        if event is None:
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        try:
            col_idx = int(str(col_id).replace("#", "")) - 1
            cols = list(self.tree["columns"] or [])
            if col_idx < 0 or col_idx >= len(cols):
                return
            if cols[col_idx] != "SEL":
                return
        except Exception:
            return

        self._toggle_selected_iid(row_id)
        self.carregar()
        return "break"

    def set_all_selected(self, val):
        self._ensure_vendas_usada_cols()

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                _call_api(
                    "POST",
                    f"desktop/vendas-importadas/marcar-todas?selected={int(val)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                self.carregar()
                return
            except Exception as exc:
                logging.debug("Falha ao marcar/desmarcar todas via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel atualizar a selecao na API central.\n\n"
                    "A alteracao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                    f"Detalhe: {exc}",
                )
                return

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute(
                "UPDATE vendas_importadas SET selecionada=? WHERE IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''",
                (int(val),),
            )
        self.carregar()

    def marcar_selecionadas(self):
        """
        Marca apenas as vendas selecionadas na tabela (suporta sele??o m?ltipla com Shift/Ctrl).
        """
        self._ensure_vendas_usada_cols()
        itens = self.tree.selection() or ()
        if not itens:
            messagebox.showwarning("ATENCAO", "Selecione uma ou mais vendas para marcar.")
            return

        ids = []
        seen = set()
        for iid in itens:
            vals = self.tree.item(iid, "values") or ()
            if not vals:
                continue
            rid = safe_int(vals[0], 0)
            if rid > 0 and rid not in seen:
                seen.add(rid)
                ids.append(rid)

        if not ids:
            messagebox.showwarning("ATENCAO", "Nao foi possivel identificar as vendas selecionadas.")
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                ids_csv = ",".join(str(x) for x in ids)
                _call_api(
                    "POST",
                    f"desktop/vendas-importadas/marcar-ids?ids={urllib.parse.quote(ids_csv)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                self.carregar()
                self.set_status(f"STATUS: {len(ids)} venda(s) marcada(s) a partir da selecao.")
                return
            except Exception as exc:
                logging.debug("Falha ao marcar ids via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel atualizar a selecao na API central.\n\n"
                    "A alteracao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                    f"Detalhe: {exc}",
                )
                return

        with get_db() as conn:
            cur = conn.cursor()
            cur.executemany(
                "UPDATE vendas_importadas SET selecionada=1 WHERE id=? AND IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''",
                [(rid,) for rid in ids],
            )

        self.carregar()
        self.set_status(f"STATUS: {len(ids)} venda(s) marcada(s) a partir da selecao.")

    def limpar_tudo(self):
        if not messagebox.askyesno("CONFIRMAR", "Deseja apagar TODAS as vendas importadas?"):
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                _call_api(
                    "DELETE",
                    "desktop/vendas-importadas",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                self.carregar()
                return
            except Exception as exc:
                logging.debug("Falha ao limpar vendas importadas via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel limpar as vendas na API central.\n\n"
                    "A limpeza local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                    f"Detalhe: {exc}",
                )
                return

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM vendas_importadas WHERE IFNULL(usada,0)=0 AND TRIM(COALESCE(codigo_programacao,''))=''")

        self.carregar()

class ProgramacaoPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Programacao")

        self._editing = None
        self._prog_cols_checked = False  # evita PRAGMA/ALTER toda hora
        self._vendas_cols_checked = False  # evita PRAGMA/ALTER toda hora
        self._equipe_display_map = {}
        self._veiculos_lookup = {}
        self._editing_programacao_codigo = ""
        self._loaded_venda_ids = []

        # -------------------------
        # Cabeçalho (dados da programação)
        # -------------------------
        card = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=5, minsize=420)
        card.grid_columnconfigure(1, weight=6, minsize=500)

        team = ttk.Frame(card, style="CardInset.TFrame", padding=(12, 10))
        team.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        team.grid_columnconfigure(0, weight=3, minsize=120)
        team.grid_columnconfigure(1, weight=3, minsize=120)
        team.grid_columnconfigure(2, weight=4, minsize=210)

        route = ttk.Frame(card, style="CardInset.TFrame", padding=(12, 10))
        route.grid(row=0, column=1, sticky="nsew", padx=(0, 10))
        route.grid_columnconfigure(0, weight=3, minsize=110)
        route.grid_columnconfigure(1, weight=3, minsize=150)
        route.grid_columnconfigure(2, weight=3, minsize=130)
        route.grid_columnconfigure(3, weight=2, minsize=110)

        control = ttk.Frame(card, style="CardInset.TFrame", padding=(12, 10))
        control.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        control.grid_columnconfigure(0, weight=0, minsize=95)
        control.grid_columnconfigure(1, weight=0, minsize=95)
        control.grid_columnconfigure(2, weight=2, minsize=170)
        control.grid_columnconfigure(3, weight=2, minsize=180)
        control.grid_columnconfigure(4, weight=2, minsize=180)
        control.grid_columnconfigure(5, weight=2, minsize=180)

        ttk.Label(team, text="Motorista", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.cb_motorista = ttk.Combobox(team, state="readonly")
        self.cb_motorista.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(team, text="Veiculo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.cb_veiculo = ttk.Combobox(team, state="readonly")
        self.cb_veiculo.grid(row=1, column=1, sticky="ew", padx=(0, 8))

        ttk.Label(team, text="Ajudantes", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.btn_ajudantes = ttk.Button(
            team,
            text="\U0001F465 Selecionar ajudantes",
            style="Ghost.TButton",
            command=self._open_ajudantes_selector,
        )
        self.btn_ajudantes.grid(row=1, column=2, sticky="ew")
        self.lbl_ajudantes_sel = ttk.Label(
            team,
            text="Nenhum selecionado",
            style="CardLabel.TLabel",
            wraplength=190,
        )
        self.lbl_ajudantes_sel.grid(row=2, column=2, sticky="w", pady=(6, 0))

        ttk.Label(route, text="Local da Rota", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.cb_local_rota = ttk.Combobox(route, state="readonly", values=["SERRA", "SERTAO"])
        self.cb_local_rota.grid(row=1, column=0, sticky="ew", padx=(0, 10))

        ttk.Label(route, text="Estimativa (KG/CX)", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        frm_estim = ttk.Frame(route, style="CardInset.TFrame")
        frm_estim.grid(row=1, column=1, sticky="ew", padx=(0, 10))
        frm_estim.grid_columnconfigure(1, weight=1)
        self.cb_estimativa_tipo = ttk.Combobox(frm_estim, state="readonly", values=["KG", "CX"], width=5)
        self.cb_estimativa_tipo.grid(row=0, column=0, sticky="w")
        self.cb_estimativa_tipo.set("KG")
        self.cb_estimativa_tipo.bind("<<ComboboxSelected>>", lambda _e: self._on_estimativa_tipo_change())
        self.ent_kg = ttk.Entry(frm_estim, style="Field.TEntry")
        self.ent_kg.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.ent_kg.bind("<KeyRelease>", lambda _e: self._refresh_total_caixas_field())
        self.lbl_estimado_hint = ttk.Label(route, text="KG estimado", style="CardLabel.TLabel")
        self.lbl_estimado_hint.grid(row=2, column=1, sticky="w", pady=(6, 0))

        ttk.Label(route, text="Carregamento", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.ent_carregamento_prog = ttk.Entry(route, style="Field.TEntry")
        self.ent_carregamento_prog.grid(row=1, column=2, sticky="ew", padx=(0, 10))
        bind_entry_smart(self.ent_carregamento_prog, "text")

        ttk.Label(route, text="Total Caixas", style="CardLabel.TLabel").grid(row=0, column=3, sticky="w")
        self.ent_total_caixas_prog = ttk.Entry(route, style="Field.TEntry", state="readonly")
        self.ent_total_caixas_prog.grid(row=1, column=3, sticky="ew")

        ttk.Label(control, text="Adiantamento (R$)", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_adiantamento_prog = ttk.Entry(control, style="Field.TEntry")
        self.ent_adiantamento_prog.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self.ent_adiantamento_prog.insert(0, "0,00")
        self._bind_money_entry(self.ent_adiantamento_prog)

        ttk.Label(control, text="Codigo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.ent_codigo = ttk.Entry(control, style="Field.TEntry", state="readonly")
        self.ent_codigo.grid(row=1, column=1, sticky="ew", padx=(0, 10))
        ttk.Label(control, text="Status", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.lbl_prog_status_badge = ttk.Label(
            control,
            text="Status atual: -",
            background="#F5F8FE",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_prog_status_badge.grid(row=1, column=2, sticky="ew")
        self.btn_salvar_prog = ttk.Button(
            control,
            text="SALVAR PROGRAMAÇÃO",
            style="Primary.TButton",
            command=self.salvar_programacao
        )
        self.btn_salvar_prog.grid(row=1, column=3, sticky="ew", padx=(10, 8))
        self.btn_editar_prog = ttk.Button(
            control,
            text="EDITAR PROGRAMAÇÃO",
            style="Ghost.TButton",
            command=self.carregar_programacao_para_edicao
        )
        self.btn_editar_prog.grid(row=1, column=4, sticky="ew", padx=(0, 8))
        self.btn_excluir_prog = ttk.Button(
            control,
            text="EXCLUIR PROGRAMAÇÃO",
            style="Danger.TButton",
            command=self.excluir_programacao,
        )
        self.btn_excluir_prog.grid(row=1, column=5, sticky="ew")

        rank_wrap = ttk.Frame(card, style="Card.TFrame")
        rank_wrap.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        rank_wrap.grid_columnconfigure(0, weight=1)
        rank_wrap.grid_columnconfigure(1, weight=0)

        ttk.Label(rank_wrap, text="Recomendacoes da equipe", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.btn_toggle_rankings = ttk.Button(
            rank_wrap,
            text="Ocultar",
            style="Toggle.TButton",
            command=self._toggle_programacao_rankings,
        )
        self.btn_toggle_rankings.grid(row=0, column=1, sticky="e")

        self.rankings_content = ttk.Frame(rank_wrap, style="CardInset.TFrame", padding=(12, 8))
        self.rankings_content.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.rankings_content.grid_columnconfigure(0, weight=1)

        self.lbl_motorista_rank = ttk.Label(
            self.rankings_content,
            text="Ranking de motoristas: carregando...",
            background="#F5F8FE",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold"),
            justify="left",
            anchor="w",
        )
        self.lbl_motorista_rank.grid(row=0, column=0, sticky="ew")

        self.lbl_ajudantes_rank = ttk.Label(
            self.rankings_content,
            text="Ranking de ajudantes: carregando...",
            background="#F5F8FE",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold"),
            justify="left",
            anchor="w",
        )
        self.lbl_ajudantes_rank.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self._rankings_collapsed = False
        self._ajudantes_rows = []
        self._ajudantes_selected_keys = []
        self._ajudantes_mode = "ajudantes"
        self._ajudantes_popup = None
        self._ajudantes_tree = None
        self._ajudantes_tree_key_by_iid = {}
        self._ajudantes_tree_iid_by_key = {}
        self._ajudantes_filter_var = tk.StringVar(value="")
        self._ranking_periodo_dias = 30
        self._motoristas_ranking = []
        self._ajudantes_ranking = []

        # -------------------------
        # Itens (vendas / edição)
        # -------------------------
        card2 = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        card2.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.body.grid_rowconfigure(1, weight=1)

        card2.grid_columnconfigure(0, weight=1)
        card2.grid_rowconfigure(1, weight=1)

        top2 = ttk.Frame(card2, style="Card.TFrame")
        top2.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        top2.grid_columnconfigure(0, weight=1)

        acoes_linha_1 = ttk.Frame(top2, style="Card.TFrame")
        acoes_linha_1.grid(row=0, column=0, sticky="ew")
        for i in range(4):
            acoes_linha_1.grid_columnconfigure(i, weight=1, uniform="prog_actions_1")

        ttk.Button(
            acoes_linha_1,
            text="CARREGAR VENDAS",
            style="Warn.TButton",
            command=self.carregar_vendas_selecionadas
        ).grid(row=0, column=0, padx=4, sticky="ew")

        ttk.Button(acoes_linha_1, text="\u2795 INSERIR LINHA", style="Ghost.TButton",
                   command=self.inserir_linha).grid(row=0, column=1, padx=4, sticky="ew")

        ttk.Button(acoes_linha_1, text="\u2796 REMOVER LINHA", style="Danger.TButton",
                   command=self.remover_linha).grid(row=0, column=2, padx=4, sticky="ew")

        ttk.Button(acoes_linha_1, text="\U0001F9F9 LIMPAR ITENS", style="Danger.TButton",
                   command=self.limpar_itens).grid(row=0, column=3, padx=4, sticky="ew")

        acoes_linha_2 = ttk.Frame(top2, style="Card.TFrame")
        acoes_linha_2.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        acoes_linha_2.grid_columnconfigure(0, weight=1)
        ttk.Button(
            acoes_linha_2,
            text="IMPRIMIR ROMANEIOS",
            style="Ghost.TButton",
            command=self.imprimir_romaneios_programacao
        ).grid(row=0, column=0, padx=4, sticky="ew")

        ttk.Label(
            top2,
            text="Dica: duplo clique para editar Endereco/Caixas/Preco/Vendedor/Pedido/Obs. ENTER confirma, ESC cancela.",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        ).grid(row=2, column=0, padx=6, pady=(8, 0), sticky="w")

        cols = ["COD CLIENTE", "NOME CLIENTE", "PRODUTO", "ENDERECO", "CAIXAS", "KG", "PRECO", "VENDEDOR", "PEDIDO", "OBS"]

        table_wrap = ttk.Frame(card2, style="Card.TFrame")
        table_wrap.grid(row=1, column=0, sticky="nsew")
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings")
        self.tree.configure(
            displaycolumns=(
                "COD CLIENTE",
                "NOME CLIENTE",
                "PRODUTO",
                "ENDERECO",
                "CAIXAS",
                "PRECO",
                "VENDEDOR",
                "PEDIDO",
                "OBS",
            )
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=135, minwidth=90, anchor="w", stretch=True)

        self.tree.column("COD CLIENTE", width=110, minwidth=90, stretch=False)
        self.tree.column("ENDERECO", width=220, minwidth=150)
        self.tree.column("NOME CLIENTE", width=210, minwidth=150)
        self.tree.column("PRODUTO", width=140, minwidth=110)
        self.tree.column("VENDEDOR", width=130, minwidth=100)
        self.tree.column("PEDIDO", width=110, minwidth=90, stretch=False)
        self.tree.column("OBS", width=180, minwidth=120)
        self.tree.column("CAIXAS", width=75, minwidth=70, anchor="center", stretch=False)
        self.tree.column("KG", width=85, minwidth=75, anchor="center", stretch=False)
        self.tree.column("PRECO", width=90, minwidth=80, anchor="e", stretch=False)

        self.tree.bind("<Double-1>", self._start_edit_cell)
        self.tree.bind("<MouseWheel>", self._on_tree_scroll, add=True)
        self.tree.bind("<Button-4>", self._on_tree_scroll, add=True)  # linux
        self.tree.bind("<Button-5>", self._on_tree_scroll, add=True)  # linux

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"CAIXAS", "KG"},
            money_cols={"PRECO"},
            date_cols=set()
        )

        self._combo_load_seq = 0
        self._refresh_total_caixas_field()

    # -------------------------
    # Utilitários
    # -------------------------
    def _motorista_display(self, nome: str, codigo: str) -> str:
        nome = upper(nome)
        codigo = upper(codigo)
        if codigo:
            return f"{nome} ({codigo})"
        return nome

    def _parse_motorista_display(self, s: str):
        s = upper(s).strip()
        if not s:
            return "", ""
        m = re.match(r"^(.*)\s+\(([A-Z0-9_-]+)\)\s*$", s)
        if m:
            return upper(m.group(1)), upper(m.group(2))
        return upper(s), ""

    def _normalize_local_rota(self, v: str) -> str:
        """
        Normaliza Local da Rota para valores canônicos:
        - SERRA
        - SERTAO
        Tolerante a acentos/mojibake (ex.: SERTÃO, SERTÃO).
        """
        txt = str(v or "").strip()
        txt = fix_mojibake_text(txt) if txt else ""
        txt = upper(txt)
        # Remove qualquer caractere não alfanumérico para comparação robusta.
        key = re.sub(r"[^A-Z0-9]", "", txt)
        if key.startswith("SERRA"):
            return "SERRA"
        if key.startswith("SERT"):
            return "SERTAO"
        return ""

    def normalizar_indicador(self, lista_valores):
        valores = [safe_float(v, 0.0) for v in (lista_valores or [])]
        if not valores:
            return []
        minimo = min(valores)
        maximo = max(valores)
        if maximo == minimo:
            return [0.0 for _ in valores]
        return [(v - minimo) / (maximo - minimo) for v in valores]

    def _ranking_status_normalizado(self, status_raw) -> str:
        return upper(str(status_raw or "").strip())

    def _ranking_is_ativa(self, status_raw: str) -> bool:
        return self._ranking_status_normalizado(status_raw) in {"ATIVA", "EM_ROTA", "EM ROTA", "INICIADA", "CARREGADA"}

    def _ranking_is_cancelada(self, status_raw: str) -> bool:
        return self._ranking_status_normalizado(status_raw) in {"CANCELADA", "CANCELADO"}

    def _ranking_parse_data_programacao(self, raw: str):
        txt = str(raw or "").strip()
        if not txt:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%d/%m/%y %H:%M:%S", "%d/%m/%y"):
            try:
                return datetime.strptime(txt[:19], fmt) if "H" in fmt else datetime.strptime(txt[:10], fmt)
            except Exception:
                pass
        try:
            return datetime.fromisoformat(txt.replace(" ", "T"))
        except Exception:
            return None

    def _ranking_parse_data_hora(self, data_raw: str, hora_raw: str):
        dt_data = self._ranking_parse_data_programacao(data_raw)
        if dt_data is None:
            return None
        nt = normalize_time(str(hora_raw or "").strip()) if hora_raw else ""
        hh = mm = ss = 0
        if nt:
            try:
                p = (nt + ":00:00").split(":")
                hh = safe_int(p[0], 0)
                mm = safe_int(p[1], 0)
                ss = safe_int(p[2], 0)
            except Exception:
                hh = mm = ss = 0
        return dt_data.replace(hour=hh, minute=mm, second=ss, microsecond=0)

    def _ranking_calc_horas_trabalhadas(self, data_saida: str, hora_saida: str, data_chegada: str, hora_chegada: str) -> float:
        dt_saida = self._ranking_parse_data_hora(data_saida, hora_saida)
        dt_chegada = self._ranking_parse_data_hora(data_chegada, hora_chegada)
        if not dt_saida or not dt_chegada:
            return 0.0
        diff = (dt_chegada - dt_saida).total_seconds() / 3600.0
        if diff <= 0:
            return 0.0
        return min(round(diff, 2), 72.0)

    def _listar_programacoes_para_score(self, periodo_dias: int = 30):
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(programacoes)")
            cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            if "data_saida" in cols and "data_criacao" in cols and "data" in cols:
                data_ref_expr = "COALESCE(data_saida,data_criacao,data,'')"
            elif "data_saida" in cols and "data_criacao" in cols:
                data_ref_expr = "COALESCE(data_saida,data_criacao,'')"
            elif "data_saida" in cols and "data" in cols:
                data_ref_expr = "COALESCE(data_saida,data,'')"
            elif "data_criacao" in cols and "data" in cols:
                data_ref_expr = "COALESCE(data_criacao,data,'')"
            elif "data_saida" in cols:
                data_ref_expr = "COALESCE(data_saida,'')"
            elif "data_criacao" in cols:
                data_ref_expr = "COALESCE(data_criacao,'')"
            elif "data" in cols:
                data_ref_expr = "COALESCE(data,'')"
            else:
                data_ref_expr = "''"
            status_op_expr = "COALESCE(status_operacional,'')" if "status_operacional" in cols else "''"
            finalizada_expr = "COALESCE(finalizada_no_app,0)" if "finalizada_no_app" in cols else "0"
            motorista_codigo_expr = "COALESCE(motorista_codigo,'')" if "motorista_codigo" in cols else ("COALESCE(codigo_motorista,'')" if "codigo_motorista" in cols else "''")
            km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
            data_saida_expr = "COALESCE(data_saida,'')" if "data_saida" in cols else "''"
            hora_saida_expr = "COALESCE(hora_saida,'')" if "hora_saida" in cols else "''"
            data_chegada_expr = "COALESCE(data_chegada,'')" if "data_chegada" in cols else "''"
            hora_chegada_expr = "COALESCE(hora_chegada,'')" if "hora_chegada" in cols else "''"
            cur.execute(
                f"""
                SELECT
                    COALESCE(codigo_programacao,'') AS codigo_programacao,
                    {data_ref_expr} AS data_ref,
                    COALESCE(motorista,'') AS motorista,
                    {motorista_codigo_expr} AS motorista_codigo,
                    COALESCE(equipe,'') AS equipe,
                    COALESCE(status,'') AS status,
                    {status_op_expr} AS status_operacional,
                    {finalizada_expr} AS finalizada_no_app,
                    {km_expr} AS km_rodado,
                    {data_saida_expr} AS data_saida,
                    {hora_saida_expr} AS hora_saida,
                    {data_chegada_expr} AS data_chegada,
                    {hora_chegada_expr} AS hora_chegada
                FROM programacoes
                ORDER BY id DESC
                """
            )
            rows = cur.fetchall() or []

        try:
            cutoff = datetime.now() - timedelta(days=max(int(periodo_dias or 30), 1))
        except Exception:
            cutoff = datetime.now() - timedelta(days=30)

        out = []
        for row in rows:
            data_ref = row["data_ref"] if hasattr(row, "keys") else row[1]
            dt_ref = self._ranking_parse_data_programacao(data_ref)
            if dt_ref is not None and dt_ref < cutoff:
                continue
            status = self._ranking_status_normalizado(
                (row["status_operacional"] if hasattr(row, "keys") else row[6])
                or (row["status"] if hasattr(row, "keys") else row[5])
            )
            if not status and safe_int(row["finalizada_no_app"] if hasattr(row, "keys") else row[7], 0) == 1:
                status = "FINALIZADA"
            if self._ranking_is_cancelada(status):
                continue
            out.append(
                {
                    "codigo_programacao": upper(row["codigo_programacao"] if hasattr(row, "keys") else row[0]),
                    "data_ref": dt_ref,
                    "motorista": upper(row["motorista"] if hasattr(row, "keys") else row[2]),
                    "motorista_codigo": upper(row["motorista_codigo"] if hasattr(row, "keys") else row[3]),
                    "equipe": str(row["equipe"] if hasattr(row, "keys") else row[4] or ""),
                    "status": status,
                    "km_rodado": safe_float(row["km_rodado"] if hasattr(row, "keys") else row[8], 0.0),
                    "data_saida": str(row["data_saida"] if hasattr(row, "keys") else row[9] or ""),
                    "hora_saida": str(row["hora_saida"] if hasattr(row, "keys") else row[10] or ""),
                    "data_chegada": str(row["data_chegada"] if hasattr(row, "keys") else row[11] or ""),
                    "hora_chegada": str(row["hora_chegada"] if hasattr(row, "keys") else row[12] or ""),
                }
            )
        return out

    def _listar_candidatos_motoristas(self):
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(motoristas)")
            cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            cur.execute(
                "SELECT COALESCE(id,0), COALESCE(nome,''), COALESCE(codigo,''), COALESCE(status,'ATIVO'), "
                + ("COALESCE(acesso_liberado,1)" if "acesso_liberado" in cols else "1")
                + " FROM motoristas ORDER BY nome, codigo"
            )
            rows = cur.fetchall() or []
        candidatos = []
        for row in rows:
            nome = upper(row[1])
            codigo = upper(row[2])
            status = self._ranking_status_normalizado(row[3]) or "ATIVO"
            acesso_liberado = safe_int(row[4], 1)
            bloqueado = acesso_liberado == 0 or status in {"INATIVO", "BLOQUEADO", "DESATIVADO"}
            candidatos.append(
                {
                    "id": safe_int(row[0], 0),
                    "nome": nome,
                    "codigo": codigo,
                    "display": self._motorista_display(nome, codigo),
                    "status": status,
                    "elegivel": bool(nome) and status == "ATIVO" and not bloqueado,
                }
            )
        return candidatos

    def _listar_candidatos_ajudantes(self):
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, COALESCE(nome,''), COALESCE(sobrenome,''), COALESCE(status,'ATIVO'), COALESCE(telefone,'') FROM ajudantes ORDER BY nome, sobrenome")
            rows = cur.fetchall() or []
        candidatos = []
        for row in rows:
            ajudante_id = str(safe_int(row[0], 0))
            nome = upper(row[1])
            sobrenome = upper(row[2])
            status = self._ranking_status_normalizado(row[3]) or "ATIVO"
            candidatos.append(
                {
                    "id": ajudante_id,
                    "nome": nome,
                    "display": format_ajudante_nome(nome, sobrenome, ajudante_id),
                    "telefone": normalize_phone(row[4]),
                    "status": status,
                    "elegivel": bool(nome) and status == "ATIVO",
                }
            )
        return candidatos

    def _resolver_motorista_programacao_key(self, programacao: dict, por_codigo: dict, por_nome: dict):
        codigo = upper(programacao.get("motorista_codigo") or "")
        nome = upper(programacao.get("motorista") or "")
        if codigo and codigo in por_codigo:
            return por_codigo[codigo]
        if nome and nome in por_nome:
            return por_nome[nome]
        return ""

    def _resolver_ajudantes_programacao_keys(self, equipe_raw: str, por_id: dict, por_nome: dict):
        partes = [p.strip() for p in re.split(r"[|,;/]+", str(equipe_raw or "")) if str(p or "").strip()]
        out = []
        vistos = set()
        for parte in partes:
            if parte in por_id and parte not in vistos:
                vistos.add(parte)
                out.append(parte)
                continue
            nome_resolvido = upper(resolve_equipe_nomes(parte) or parte)
            ajudante_id = por_nome.get(nome_resolvido, "")
            if ajudante_id and ajudante_id not in vistos:
                vistos.add(ajudante_id)
                out.append(ajudante_id)
        return out

    def calcular_metricas_motoristas(self, periodo_dias: int = 30):
        candidatos = self._listar_candidatos_motoristas()
        programacoes = self._listar_programacoes_para_score(periodo_dias=periodo_dias)
        now = datetime.now()
        por_codigo = {c["codigo"]: str(c["id"]) for c in candidatos if c.get("codigo")}
        por_nome = {c["nome"]: str(c["id"]) for c in candidatos if c.get("nome")}
        conflitos = set()
        for prog in programacoes:
            if self._ranking_is_ativa(prog.get("status")):
                key = self._resolver_motorista_programacao_key(prog, por_codigo, por_nome)
                if key:
                    conflitos.add(key)
        metricas = {}
        for cand in candidatos:
            if not cand.get("elegivel"):
                continue
            cand_id = str(cand["id"])
            metricas[cand_id] = {
                "id": cand["id"],
                "nome": cand["nome"],
                "codigo": cand.get("codigo", ""),
                "display": cand.get("display") or cand["nome"],
                "status": cand.get("status", "ATIVO"),
                "total_horas_trabalhadas": 0.0,
                "total_km_rodado": 0.0,
                "total_viagens": 0,
                "dias_desde_ultima_programacao": max(safe_int(periodo_dias, 30), 1),
                "score_final": 0.0,
                "motivo_resumido": "",
                "disponivel": cand_id not in conflitos,
                "sem_base_historica": True,
                "_last_dt": None,
            }
        for prog in programacoes:
            cand_id = self._resolver_motorista_programacao_key(prog, por_codigo, por_nome)
            if not cand_id or cand_id not in metricas:
                continue
            horas = self._ranking_calc_horas_trabalhadas(prog.get("data_saida", ""), prog.get("hora_saida", ""), prog.get("data_chegada", ""), prog.get("hora_chegada", ""))
            item = metricas[cand_id]
            item["total_horas_trabalhadas"] += horas
            item["total_km_rodado"] += safe_float(prog.get("km_rodado"), 0.0)
            item["total_viagens"] += 1
            dt_ref = prog.get("data_ref")
            if isinstance(dt_ref, datetime) and (item["_last_dt"] is None or dt_ref > item["_last_dt"]):
                item["_last_dt"] = dt_ref
                item["dias_desde_ultima_programacao"] = max((now - dt_ref).days, 0)
        for item in metricas.values():
            item["total_horas_trabalhadas"] = round(float(item["total_horas_trabalhadas"]), 2)
            item["total_km_rodado"] = round(float(item["total_km_rodado"]), 2)
            item["sem_base_historica"] = item["total_horas_trabalhadas"] <= 0 and item["total_km_rodado"] <= 0 and item["total_viagens"] <= 0 and item["_last_dt"] is None
        return metricas

    def calcular_metricas_ajudantes(self, periodo_dias: int = 30):
        candidatos = self._listar_candidatos_ajudantes()
        programacoes = self._listar_programacoes_para_score(periodo_dias=periodo_dias)
        now = datetime.now()
        por_id = {str(c["id"]): str(c["id"]) for c in candidatos if c.get("id")}
        por_nome = {upper(c.get("display") or ""): str(c["id"]) for c in candidatos if c.get("display")}
        conflitos = set()
        for prog in programacoes:
            if self._ranking_is_ativa(prog.get("status")):
                conflitos.update(self._resolver_ajudantes_programacao_keys(prog.get("equipe", ""), por_id, por_nome))
        metricas = {}
        for cand in candidatos:
            if not cand.get("elegivel"):
                continue
            cand_id = str(cand["id"])
            metricas[cand_id] = {
                "id": cand_id,
                "nome": cand.get("display") or cand.get("nome") or "",
                "status": cand.get("status", "ATIVO"),
                "total_horas_trabalhadas": 0.0,
                "total_viagens": 0,
                "dias_desde_ultima_programacao": max(safe_int(periodo_dias, 30), 1),
                "score_final": 0.0,
                "motivo_resumido": "",
                "disponivel": cand_id not in conflitos,
                "sem_base_historica": True,
                "_last_dt": None,
            }
        for prog in programacoes:
            ajudante_ids = self._resolver_ajudantes_programacao_keys(prog.get("equipe", ""), por_id, por_nome)
            if not ajudante_ids:
                continue
            horas = self._ranking_calc_horas_trabalhadas(prog.get("data_saida", ""), prog.get("hora_saida", ""), prog.get("data_chegada", ""), prog.get("hora_chegada", ""))
            dt_ref = prog.get("data_ref")
            for cand_id in ajudante_ids:
                if cand_id not in metricas:
                    continue
                item = metricas[cand_id]
                item["total_horas_trabalhadas"] += horas
                item["total_viagens"] += 1
                if isinstance(dt_ref, datetime) and (item["_last_dt"] is None or dt_ref > item["_last_dt"]):
                    item["_last_dt"] = dt_ref
                    item["dias_desde_ultima_programacao"] = max((now - dt_ref).days, 0)
        for item in metricas.values():
            item["total_horas_trabalhadas"] = round(float(item["total_horas_trabalhadas"]), 2)
            item["sem_base_historica"] = item["total_horas_trabalhadas"] <= 0 and item["total_viagens"] <= 0 and item["_last_dt"] is None
        return metricas

    def calcular_score_motorista(self, metricas: dict, contexto: dict):
        horas_score = 1.0 - safe_float(contexto["horas_norm"].get(metricas["id"], 0.0), 0.0)
        km_score = 1.0 - safe_float(contexto["km_norm"].get(metricas["id"], 0.0), 0.0)
        viagens_score = 1.0 - safe_float(contexto["viagens_norm"].get(metricas["id"], 0.0), 0.0)
        descanso_score = safe_float(contexto["dias_norm"].get(metricas["id"], 0.0), 0.0)
        metricas["horas_score"] = horas_score
        metricas["km_score"] = km_score
        metricas["viagens_score"] = viagens_score
        metricas["descanso_score"] = descanso_score
        return round((horas_score * 0.35) + (km_score * 0.30) + (viagens_score * 0.20) + (descanso_score * 0.15), 6)

    def calcular_score_ajudante(self, metricas: dict, contexto: dict):
        horas_score = 1.0 - safe_float(contexto["horas_norm"].get(metricas["id"], 0.0), 0.0)
        viagens_score = 1.0 - safe_float(contexto["viagens_norm"].get(metricas["id"], 0.0), 0.0)
        descanso_score = safe_float(contexto["dias_norm"].get(metricas["id"], 0.0), 0.0)
        metricas["horas_score"] = horas_score
        metricas["viagens_score"] = viagens_score
        metricas["descanso_score"] = descanso_score
        return round((horas_score * 0.50) + (viagens_score * 0.30) + (descanso_score * 0.20), 6)

    def _motivo_resumido_motorista(self, item: dict, periodo_dias: int):
        if safe_float(item.get("descanso_score"), 0.0) >= max(safe_float(item.get("horas_score"), 0.0), safe_float(item.get("km_score"), 0.0), safe_float(item.get("viagens_score"), 0.0)):
            return "Maior intervalo desde a última programação"
        if safe_float(item.get("horas_score"), 0.0) >= 0.60 and safe_float(item.get("viagens_score"), 0.0) >= 0.60:
            return "Menos horas e menos viagens no período"
        return f"Menor carga acumulada nos últimos {safe_int(periodo_dias, 30)} dias"

    def _motivo_resumido_ajudante(self, item: dict, periodo_dias: int):
        if safe_float(item.get("descanso_score"), 0.0) >= max(safe_float(item.get("horas_score"), 0.0), safe_float(item.get("viagens_score"), 0.0)):
            return "Maior intervalo desde a última programação"
        if safe_float(item.get("horas_score"), 0.0) >= 0.60 and safe_float(item.get("viagens_score"), 0.0) >= 0.60:
            return "Menos horas e menos viagens no período"
        return f"Menor carga acumulada nos últimos {safe_int(periodo_dias, 30)} dias"

    def ranquear_motoristas(self, periodo_dias: int = 30):
        metricas = self.calcular_metricas_motoristas(periodo_dias=periodo_dias)
        elegiveis = [dict(v) for v in metricas.values() if v.get("disponivel")]
        if not elegiveis:
            return []
        todos_sem_historico = all(v.get("sem_base_historica") for v in elegiveis)
        if not todos_sem_historico:
            ids = [v["id"] for v in elegiveis]
            contexto = {
                "horas_norm": dict(zip(ids, self.normalizar_indicador([v["total_horas_trabalhadas"] for v in elegiveis]))),
                "km_norm": dict(zip(ids, self.normalizar_indicador([v["total_km_rodado"] for v in elegiveis]))),
                "viagens_norm": dict(zip(ids, self.normalizar_indicador([v["total_viagens"] for v in elegiveis]))),
                "dias_norm": dict(zip(ids, self.normalizar_indicador([v["dias_desde_ultima_programacao"] for v in elegiveis]))),
            }
            for item in elegiveis:
                item["score_final"] = self.calcular_score_motorista(item, contexto)
                item["motivo_resumido"] = self._motivo_resumido_motorista(item, periodo_dias)
        else:
            for item in elegiveis:
                item["score_final"] = 0.0
                item["motivo_resumido"] = "Sem base histórica suficiente"
        elegiveis.sort(key=lambda item: (-round(safe_float(item.get("score_final"), 0.0), 8), safe_float(item.get("total_horas_trabalhadas"), 0.0), safe_int(item.get("total_viagens"), 0), safe_float(item.get("total_km_rodado"), 0.0), -safe_int(item.get("dias_desde_ultima_programacao"), 0), upper(item.get("nome") or "")))
        for pos, item in enumerate(elegiveis, start=1):
            item["posicao_ranking"] = pos
            item["score_exibicao"] = round(safe_float(item.get("score_final"), 0.0) * 100.0, 2)
            item.pop("_last_dt", None)
        return elegiveis

    def ranquear_ajudantes(self, periodo_dias: int = 30):
        metricas = self.calcular_metricas_ajudantes(periodo_dias=periodo_dias)
        elegiveis = [dict(v) for v in metricas.values() if v.get("disponivel")]
        if not elegiveis:
            return []
        todos_sem_historico = all(v.get("sem_base_historica") for v in elegiveis)
        if not todos_sem_historico:
            ids = [v["id"] for v in elegiveis]
            contexto = {
                "horas_norm": dict(zip(ids, self.normalizar_indicador([v["total_horas_trabalhadas"] for v in elegiveis]))),
                "viagens_norm": dict(zip(ids, self.normalizar_indicador([v["total_viagens"] for v in elegiveis]))),
                "dias_norm": dict(zip(ids, self.normalizar_indicador([v["dias_desde_ultima_programacao"] for v in elegiveis]))),
            }
            for item in elegiveis:
                item["score_final"] = self.calcular_score_ajudante(item, contexto)
                item["motivo_resumido"] = self._motivo_resumido_ajudante(item, periodo_dias)
        else:
            for item in elegiveis:
                item["score_final"] = 0.0
                item["motivo_resumido"] = "Sem base histórica suficiente"
        elegiveis.sort(key=lambda item: (-round(safe_float(item.get("score_final"), 0.0), 8), safe_float(item.get("total_horas_trabalhadas"), 0.0), safe_int(item.get("total_viagens"), 0), -safe_int(item.get("dias_desde_ultima_programacao"), 0), upper(item.get("nome") or "")))
        for pos, item in enumerate(elegiveis, start=1):
            item["posicao_ranking"] = pos
            item["score_exibicao"] = round(safe_float(item.get("score_final"), 0.0) * 100.0, 2)
            item.pop("_last_dt", None)
        return elegiveis

    def _formatar_top_ranking(self, titulo: str, ranking: list, limite: int = 3):
        if not ranking:
            return f"{titulo}: sem candidatos elegíveis."
        top = ranking[:max(safe_int(limite, 3), 1)]
        partes = [f"{item['posicao_ranking']}. {item['nome']} ({item['score_exibicao']:.2f})" for item in top]
        motivo = str(top[0].get("motivo_resumido") or "").strip()
        if len(motivo) > 44:
            motivo = motivo[:41].rstrip() + "..."
        return (f"{titulo}: " + " • ".join(partes) + (f"  |  Motivo lider: {motivo}" if motivo else ""))

    def _toggle_programacao_rankings(self):
        self._rankings_collapsed = not bool(getattr(self, "_rankings_collapsed", False))
        if self._rankings_collapsed:
            self.rankings_content.grid_remove()
            self.btn_toggle_rankings.configure(text="Mostrar")
        else:
            self.rankings_content.grid()
            self.btn_toggle_rankings.configure(text="Ocultar")

    def _apply_programacao_rankings(self):
        periodo = max(safe_int(getattr(self, "_ranking_periodo_dias", 30), 30), 1)
        try:
            self._motoristas_ranking = self.ranquear_motoristas(periodo_dias=periodo)
            por_nome = {upper(item.get("display") or item.get("nome") or ""): item for item in self._motoristas_ranking}
            por_codigo = {upper(item.get("codigo") or ""): item for item in self._motoristas_ranking if item.get("codigo")}
            atuais = [str(v or "").strip() for v in list(self.cb_motorista["values"] or []) if str(v or "").strip()]
            ordenados = []
            sobras = []
            for valor in atuais:
                nome, codigo = self._parse_motorista_display(valor)
                item = por_codigo.get(upper(codigo)) if codigo else None
                if not item:
                    item = por_nome.get(upper(valor)) or por_nome.get(upper(nome))
                if item:
                    ordenados.append((safe_int(item.get("posicao_ranking"), 9999), valor))
                else:
                    sobras.append(valor)
            ordenados.sort(key=lambda kv: (kv[0], upper(kv[1])))
            self.cb_motorista["values"] = [v for _pos, v in ordenados] + sorted(sobras, key=upper)
            self.lbl_motorista_rank.config(text=self._formatar_top_ranking("Top motoristas", self._motoristas_ranking))
        except Exception:
            logging.exception("Falha ao aplicar ranking de motoristas na programação")
            self._motoristas_ranking = []
            self.lbl_motorista_rank.config(text="Top motoristas: ranking indisponível.")

        try:
            self._ajudantes_ranking = self.ranquear_ajudantes(periodo_dias=periodo)
            if self._ajudantes_mode != "ajudantes":
                self.lbl_ajudantes_rank.config(text="Top ajudantes: ranking individual indisponível no modo equipes.")
                return
            por_id = {str(item.get("id") or ""): item for item in self._ajudantes_ranking}
            ordenados = []
            sobras = []
            for row in list(self._ajudantes_rows or []):
                key = str(row.get("key", "")).strip()
                item = por_id.get(key)
                if item:
                    ordenados.append((safe_int(item.get("posicao_ranking"), 9999), row))
                else:
                    sobras.append(row)
            ordenados.sort(key=lambda kv: (kv[0], upper(str(kv[1].get("label", "")))))
            self._ajudantes_rows = [row for _pos, row in ordenados] + sorted(sobras, key=lambda r: upper(str(r.get("label", ""))))
            self._rebuild_ajudantes_popup_list()
            self._refresh_ajudantes_selected_label()
            self.lbl_ajudantes_rank.config(text=self._formatar_top_ranking("Top ajudantes", self._ajudantes_ranking))
        except Exception:
            logging.exception("Falha ao aplicar ranking de ajudantes na programação")
            self._ajudantes_ranking = []
            self.lbl_ajudantes_rank.config(text="Top ajudantes: ranking indisponível.")

    def _format_money_from_digits(self, digits: str) -> str:
        digits = re.sub(r"\D", "", str(digits or ""))
        if not digits:
            return "0,00"
        if len(digits) == 1:
            int_part = "0"
            cents = "0" + digits
        elif len(digits) == 2:
            int_part = "0"
            cents = digits
        else:
            int_part = digits[:-2]
            cents = digits[-2:]
        int_part = int_part.lstrip("0") or "0"

        parts = []
        while len(int_part) > 3:
            parts.insert(0, int_part[-3:])
            int_part = int_part[:-3]
        if int_part:
            parts.insert(0, int_part)
        int_part = ".".join(parts) if parts else "0"
        return f"{int_part},{cents}"

    def _bind_money_entry(self, ent: tk.Entry):
        def _on_focus_in(_e=None):
            try:
                v = str(ent.get() or "").strip()
                if v in {"0", "0,00", "0.00", "R$ 0,00", "R$0,00"}:
                    ent.delete(0, "end")
                else:
                    ent.selection_range(0, "end")
            except Exception:
                logging.debug("Falha ignorada")

        def _on_key_release(_e=None):
            try:
                v = str(ent.get() or "")
                digits = re.sub(r"\D", "", v)
                if digits:
                    ent.delete(0, "end")
                    ent.insert(0, self._format_money_from_digits(digits))
                    ent.icursor("end")
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out(_e=None):
            try:
                v = str(ent.get() or "").strip()
                if not v:
                    ent.insert(0, "0,00")
                else:
                    ent.delete(0, "end")
                    ent.insert(0, self._format_money_from_digits(v))
            except Exception:
                logging.debug("Falha ignorada")

        ent.bind("<FocusIn>", _on_focus_in)
        ent.bind("<KeyRelease>", _on_key_release)
        ent.bind("<FocusOut>", _on_focus_out)

    def _equipe_display(self, ajudante_id: str, nome: str, sobrenome: str, telefone: str) -> str:
        return format_ajudante_nome(nome, sobrenome, ajudante_id)

    def _parse_equipe_display(self, s: str) -> str:
        s = upper(s).strip()
        if not s:
            return ""
        try:
            return self._equipe_display_map.get(s, s)
        except Exception:
            return s

    # Nova versao: seletor em tabela (nome/sobrenome/telefone)
    def _set_ajudantes_options(self, rows, mode: str = "ajudantes"):
        prev_selected = set(self._ajudantes_selected_keys or [])
        self._ajudantes_mode = mode
        self._ajudantes_rows = list(rows or [])
        valid_keys = [str(r.get("key", "")).strip() for r in self._ajudantes_rows if str(r.get("key", "")).strip()]
        self._ajudantes_selected_keys = [k for k in self._ajudantes_selected_keys if k in valid_keys]
        for k in valid_keys:
            if k in prev_selected and k not in self._ajudantes_selected_keys:
                self._ajudantes_selected_keys.append(k)
        self._rebuild_ajudantes_popup_list()
        self._sync_tree_selection_from_state()
        self._refresh_ajudantes_selected_label()

    def _enforce_ajudantes_limit(self):
        max_sel = 2 if self._ajudantes_mode == "ajudantes" else 1
        if len(self._ajudantes_selected_keys) > max_sel:
            self._ajudantes_selected_keys = self._ajudantes_selected_keys[:max_sel]
            if self._ajudantes_mode == "ajudantes":
                messagebox.showwarning("ATENCAO", "Selecione no maximo 2 ajudantes.")
            else:
                messagebox.showwarning("ATENCAO", "Selecione apenas 1 equipe.")

    def _get_row_by_key(self, key: str):
        for r in (self._ajudantes_rows or []):
            if str(r.get("key", "")).strip() == str(key).strip():
                return r
        return None

    def _get_selected_labels(self):
        out = []
        for key in self._ajudantes_selected_keys:
            r = self._get_row_by_key(key)
            if not r:
                continue
            label = str(r.get("label", "")).strip()
            if label:
                out.append(label)
        return out

    def _get_selected_ajudantes(self):
        out = []
        for key in self._ajudantes_selected_keys:
            r = self._get_row_by_key(key)
            if not r:
                continue
            value = str(r.get("value", "")).strip()
            if value:
                out.append(value)
        return out

    def _refresh_ajudantes_selected_label(self):
        labels = self._get_selected_labels()
        qtd = len(labels)
        if qtd <= 0:
            self.btn_ajudantes.configure(text="\U0001F465 Selecionar ajudantes")
            self.lbl_ajudantes_sel.configure(text="Nenhum selecionado")
            return
        self.btn_ajudantes.configure(text="\U0001F465 Selecionar ajudantes")
        self.lbl_ajudantes_sel.configure(text=" | ".join(labels[:2]) if qtd <= 2 else f"{qtd} selecionados")

    def _get_filtered_ajudantes_rows(self):
        q = upper(self._ajudantes_filter_var.get().strip()) if hasattr(self, "_ajudantes_filter_var") else ""
        if not q:
            return list(self._ajudantes_rows or [])
        out = []
        for row in (self._ajudantes_rows or []):
            bucket = " ".join([
                str(row.get("nome", "")),
                str(row.get("sobrenome", "")),
                str(row.get("telefone", "")),
                str(row.get("label", "")),
            ])
            if q in upper(bucket):
                out.append(row)
        return out

    def _sync_tree_selection_from_state(self):
        if not (self._ajudantes_tree and self._ajudantes_tree.winfo_exists()):
            return
        try:
            self._ajudantes_tree.selection_remove(*self._ajudantes_tree.selection())
        except Exception:
            logging.debug("Falha ignorada")
        for key in self._ajudantes_selected_keys:
            iid = self._ajudantes_tree_iid_by_key.get(key)
            if iid:
                self._ajudantes_tree.selection_add(iid)

    def _on_ajudantes_tree_select(self, _e=None):
        if not (self._ajudantes_tree and self._ajudantes_tree.winfo_exists()):
            return
        selected_iids = list(self._ajudantes_tree.selection())
        keys = [self._ajudantes_tree_key_by_iid.get(iid, "") for iid in selected_iids]
        self._ajudantes_selected_keys = [k for k in keys if k]
        before = list(self._ajudantes_selected_keys)
        self._enforce_ajudantes_limit()
        if before != self._ajudantes_selected_keys:
            self._sync_tree_selection_from_state()
        self._refresh_ajudantes_selected_label()

    def _destroy_ajudantes_popup(self):
        try:
            if self._ajudantes_popup and self._ajudantes_popup.winfo_exists():
                self._ajudantes_popup.destroy()
        except Exception:
            logging.debug("Falha ignorada")
        self._ajudantes_popup = None
        self._ajudantes_tree = None
        self._ajudantes_tree_key_by_iid = {}
        self._ajudantes_tree_iid_by_key = {}
        try:
            self._ajudantes_filter_var.set("")
        except Exception:
            logging.debug("Falha ignorada")

    def _confirm_ajudantes_popup(self):
        self._refresh_ajudantes_selected_label()
        self._destroy_ajudantes_popup()

    def _clear_ajudantes_selection(self):
        self._ajudantes_selected_keys = []
        self._sync_tree_selection_from_state()
        self._refresh_ajudantes_selected_label()

    def _rebuild_ajudantes_popup_list(self):
        if not (self._ajudantes_tree and self._ajudantes_tree.winfo_exists()):
            return
        try:
            self._ajudantes_tree.delete(*self._ajudantes_tree.get_children())
        except Exception:
            logging.debug("Falha ignorada")
        self._ajudantes_tree_key_by_iid = {}
        self._ajudantes_tree_iid_by_key = {}
        for row in self._get_filtered_ajudantes_rows():
            key = str(row.get("key", "")).strip()
            nome = str(row.get("nome", "")).strip()
            sobrenome = str(row.get("sobrenome", "")).strip()
            telefone = str(row.get("telefone", "")).strip()
            iid = tree_insert_aligned(self._ajudantes_tree, "", "end", (nome, sobrenome, telefone))
            if key:
                self._ajudantes_tree_key_by_iid[iid] = key
                self._ajudantes_tree_iid_by_key[key] = iid
        self._sync_tree_selection_from_state()

    def _open_ajudantes_selector(self):
        if self._ajudantes_popup and self._ajudantes_popup.winfo_exists():
            self._ajudantes_popup.lift()
            self._ajudantes_popup.focus_force()
            return

        top = tk.Toplevel(self)
        top.title("Selecionar ajudantes")
        top.transient(self.winfo_toplevel())
        top.grab_set()
        top.geometry("980x620")
        top.minsize(820, 520)
        self._ajudantes_popup = top
        top.protocol("WM_DELETE_WINDOW", self._destroy_ajudantes_popup)

        frame = ttk.Frame(top, style="Card.TFrame", padding=10)
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        hint = "Selecione ate 2 ajudantes." if self._ajudantes_mode == "ajudantes" else "Selecione 1 equipe."
        ttk.Label(frame, text=hint, style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))

        search_row = ttk.Frame(frame, style="Card.TFrame")
        search_row.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        search_row.grid_columnconfigure(1, weight=1)
        ttk.Label(search_row, text="BUSCAR", style="CardLabel.TLabel").pack(side="left", padx=(0, 8))
        ent_search = ttk.Entry(search_row, textvariable=self._ajudantes_filter_var, style="Field.TEntry")
        ent_search.pack(side="left", fill="x", expand=True)
        ent_search.bind("<KeyRelease>", lambda _e: self._rebuild_ajudantes_popup_list())
        ent_search.focus_set()

        list_wrap = ttk.Frame(frame, style="Card.TFrame")
        list_wrap.grid(row=2, column=0, sticky="nsew")
        list_wrap.grid_columnconfigure(0, weight=1)
        list_wrap.grid_rowconfigure(0, weight=1)

        cols = ("nome", "sobrenome", "telefone")
        self._ajudantes_tree = ttk.Treeview(list_wrap, columns=cols, show="headings", selectmode="extended")
        self._ajudantes_tree.heading("nome", text="Nome")
        self._ajudantes_tree.heading("sobrenome", text="Sobrenome")
        self._ajudantes_tree.heading("telefone", text="Telefone")
        self._ajudantes_tree.column("nome", width=230, anchor="w")
        self._ajudantes_tree.column("sobrenome", width=230, anchor="w")
        self._ajudantes_tree.column("telefone", width=180, anchor="w")
        self._ajudantes_tree.grid(row=0, column=0, sticky="nsew")
        self._ajudantes_tree.bind("<<TreeviewSelect>>", self._on_ajudantes_tree_select)

        vsb = ttk.Scrollbar(list_wrap, orient="vertical", command=self._ajudantes_tree.yview)
        self._ajudantes_tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        self._rebuild_ajudantes_popup_list()

        footer = ttk.Frame(frame, style="Card.TFrame")
        footer.grid(row=3, column=0, sticky="ew", pady=(12, 2))
        footer.grid_columnconfigure(0, weight=1)
        ttk.Button(footer, text="LIMPAR", style="Ghost.TButton", command=self._clear_ajudantes_selection).pack(side="left")
        ttk.Button(footer, text="CONFIRMAR", style="Primary.TButton", command=self._confirm_ajudantes_popup).pack(side="right")

        # Centraliza para evitar popup abrindo "cortado" em telas menores.
        try:
            top.update_idletasks()
            w = top.winfo_width()
            h = top.winfo_height()
            sw = top.winfo_screenwidth()
            sh = top.winfo_screenheight()
            x = max(int((sw - w) / 2), 0)
            y = max(int((sh - h) / 2), 0)
            top.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            logging.debug("Falha ignorada")

    def _resolve_equipe_ajudantes(self, equipe_codigo: str) -> str:
        return resolve_equipe_nomes(equipe_codigo)

    def _get_row_values(self, iid):
        vals = self.tree.item(iid, "values") or ("", "", "", "", "", "", "", "", "", "")
        vals = list(vals) + [""] * (10 - len(vals))
        return [str(v or "").strip() for v in vals[:10]]

    def _collect_tree_item_rows(self):
        rows = []
        try:
            for iid in self.tree.get_children():
                rows.append(self._get_row_values(iid))
        except Exception:
            logging.debug("Falha ao coletar itens da programacao", exc_info=True)
        return rows

    def _load_existing_programacao_item_rows(self, codigo: str):
        rows = []
        itens_res = fetch_programacao_itens_contract(codigo)
        itens = itens_res.get("data") if isinstance(itens_res, dict) and bool(itens_res.get("ok", False)) else []
        for it in (itens or []):
            if not isinstance(it, dict):
                continue
            rows.append([
                upper(it.get("cod_cliente", "")),
                upper(it.get("nome_cliente", "")),
                upper(it.get("produto", "")),
                upper(it.get("endereco", "")),
                str(safe_int(it.get("qnt_caixas"), 0)),
                f"{safe_float(it.get('kg'), 0.0):.2f}",
                f"{safe_float(it.get('preco'), 0.0):.2f}",
                upper(it.get("vendedor", "")),
                upper(it.get("pedido", "")),
                upper(it.get("obs", "")),
            ])
        return rows

    def _next_programacao_codigo(self) -> str:
        prefix = f"PG{datetime.now().strftime('%Y')}"
        max_suffix = 0

        def _consume(codigo_ref):
            nonlocal max_suffix
            codigo_up = upper(str(codigo_ref or "").strip())
            if not codigo_up.startswith(prefix):
                return
            tail = codigo_up[len(prefix):]
            digits = "".join(ch for ch in tail if ch.isdigit())
            seq = safe_int(digits, 0) if digits else 0
            if seq > max_suffix:
                max_suffix = seq

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/programacoes?modo=todas&limit=1200",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                for row in ((resp or {}).get("programacoes") or []):
                    if isinstance(row, dict):
                        _consume(row.get("codigo_programacao"))
            except Exception:
                logging.debug("Falha ao calcular proximo codigo de programacao via API.", exc_info=True)

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    "SELECT codigo_programacao FROM programacoes WHERE codigo_programacao LIKE ?",
                    (f"{prefix}%",),
                )
                for row in (cur.fetchall() or []):
                    _consume(row[0] if row else "")
        except Exception:
            logging.debug("Falha ao calcular proximo codigo de programacao local.", exc_info=True)

        return f"{prefix}{max_suffix + 1:02d}"

    def _increment_programacao_codigo(self, codigo: str) -> str:
        codigo_up = upper(str(codigo or "").strip())
        prefix = f"PG{datetime.now().strftime('%Y')}"
        if not codigo_up.startswith(prefix):
            return self._next_programacao_codigo()
        tail = codigo_up[len(prefix):]
        digits = "".join(ch for ch in tail if ch.isdigit())
        seq = safe_int(digits, 0)
        return f"{prefix}{seq + 1:02d}"

    def _refresh_next_programacao_codigo_preview(self, force: bool = False):
        if not hasattr(self, "ent_codigo"):
            return
        if upper(str(self._editing_programacao_codigo or "").strip()):
            return
        codigo_atual = upper(str(self.ent_codigo.get() or "").strip())
        if codigo_atual and not force:
            return
        try:
            proximo_codigo = self._next_programacao_codigo()
            self.ent_codigo.config(state="normal")
            self.ent_codigo.delete(0, "end")
            self.ent_codigo.insert(0, proximo_codigo)
            self.ent_codigo.config(state="readonly")
        except Exception:
            logging.debug("Falha ao atualizar preview do codigo da programacao.", exc_info=True)

    def _on_tree_scroll(self, event=None):
        if self._editing:
            self._commit_edit()

    def _fetch_programacao_combobox_data(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and can_read_from_api():
            try:
                mot_resp = _call_api("GET", "desktop/cadastros/motoristas", extra_headers={"X-Desktop-Secret": desktop_secret})
                vei_resp = _call_api("GET", "desktop/cadastros/veiculos", extra_headers={"X-Desktop-Secret": desktop_secret})
                aju_resp = _call_api("GET", "desktop/cadastros/ajudantes", extra_headers={"X-Desktop-Secret": desktop_secret})

                valores_motoristas = []
                for r in (mot_resp or []):
                    if not isinstance(r, dict):
                        continue
                    status_m = upper(r.get("status") or "ATIVO")
                    if status_m != "ATIVO":
                        continue
                    valores_motoristas.append(self._motorista_display(r.get("nome") or "", r.get("codigo") or ""))

                equipe_display_map = {}
                rows_ajudantes = []
                for r in (aju_resp or []):
                    if not isinstance(r, dict):
                        continue
                    ajudante_id = str(r.get("id") or "").strip()
                    nome_full = upper(r.get("nome") or "")
                    parts = nome_full.split(" ", 1)
                    nome = parts[0] if parts else nome_full
                    sobrenome = parts[1] if len(parts) > 1 else ""
                    display = self._equipe_display(ajudante_id, nome, sobrenome, "")
                    if not display:
                        continue
                    rows_ajudantes.append(
                        {
                            "key": ajudante_id,
                            "value": ajudante_id,
                            "label": display,
                            "nome": upper(nome),
                            "sobrenome": upper(sobrenome),
                            "telefone": "",
                        }
                    )
                    if ajudante_id:
                        equipe_display_map[upper(display)] = ajudante_id
                veiculos_lookup = {}
                veiculos_values = []
                for r in (vei_resp or []):
                    if not isinstance(r, dict):
                        continue
                    placa = upper((r or {}).get("placa") or "")
                    if not placa:
                        continue
                    veiculos_values.append(placa)
                    cap_raw = r.get("capacidade_cx")
                    if cap_raw in (None, ""):
                        cap_raw = r.get("capacidade_c")
                    veiculos_lookup[placa] = {
                        "placa": placa,
                        "capacidade_cx": safe_int(cap_raw, -1),
                    }
                return {
                    "motoristas": valores_motoristas,
                    "veiculos": veiculos_values,
                    "veiculos_lookup": veiculos_lookup,
                    "ajudantes": rows_ajudantes,
                    "ajudantes_mode": "ajudantes",
                    "equipe_display_map": equipe_display_map,
                }
            except Exception:
                logging.debug("Falha ao carregar comboboxes da Programacao via API", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()

            valores_motoristas = []
            try:
                cur.execute("PRAGMA table_info(motoristas)")
                cols_m = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                if "status" in cols_m:
                    cur.execute(
                        "SELECT nome, codigo FROM motoristas "
                        "WHERE UPPER(COALESCE(status,'ATIVO'))='ATIVO' ORDER BY nome"
                    )
                else:
                    cur.execute("SELECT nome, codigo FROM motoristas ORDER BY nome")
                for r in cur.fetchall():
                    valores_motoristas.append(self._motorista_display(r[0], r[1]))
            except Exception:
                cur.execute("SELECT nome FROM motoristas ORDER BY nome")
                valores_motoristas = [r[0] for r in cur.fetchall()]

            cur.execute("PRAGMA table_info(veiculos)")
            cols_vei = [str(r[1]).lower() for r in cur.fetchall()]
            cap_col = "capacidade_cx"
            if "capacidade_cx" not in cols_vei and "capacidade_c" in cols_vei:
                cap_col = "capacidade_c"
            try:
                cur.execute(f"SELECT placa, {cap_col} FROM veiculos ORDER BY placa")
                veiculos_rows = cur.fetchall()
            except Exception:
                cur.execute("SELECT placa FROM veiculos ORDER BY placa")
                veiculos_rows = [(r[0], None) for r in cur.fetchall()]
            valores_veiculos = []
            veiculos_lookup = {}
            for r in veiculos_rows:
                placa = upper((r[0] if r else "") or "")
                if not placa:
                    continue
                valores_veiculos.append(placa)
                cap_raw = r[1] if len(r) > 1 else None
                veiculos_lookup[placa] = {
                    "placa": placa,
                    "capacidade_cx": safe_int(cap_raw, -1),
                }

            equipe_display_map = {}
            try:
                cur.execute("PRAGMA table_info(ajudantes)")
                cols_aj = [str(r[1]).lower() for r in cur.fetchall()]
                if "status" in cols_aj:
                    cur.execute("""
                        SELECT id, nome, sobrenome, telefone
                        FROM ajudantes
                        WHERE UPPER(COALESCE(status, 'ATIVO'))='ATIVO'
                        ORDER BY nome, sobrenome
                    """)
                else:
                    cur.execute("SELECT id, nome, sobrenome, telefone FROM ajudantes ORDER BY nome, sobrenome")
                rows_ajudantes = []
                for r in cur.fetchall():
                    ajudante_id = str(r[0] if r else "").strip()
                    nome = r[1] if r else ""
                    sobrenome = r[2] if r else ""
                    telefone = normalize_phone(r[3] if r else "")
                    display = self._equipe_display(ajudante_id, nome, sobrenome, "")
                    if display:
                        rows_ajudantes.append({
                            "key": ajudante_id,
                            "value": ajudante_id,
                            "label": display,
                            "nome": upper(nome),
                            "sobrenome": upper(sobrenome),
                            "telefone": telefone,
                        })
                        if ajudante_id:
                            equipe_display_map[upper(display)] = ajudante_id
                return {
                    "motoristas": valores_motoristas,
                    "veiculos": valores_veiculos,
                    "veiculos_lookup": veiculos_lookup,
                    "ajudantes": rows_ajudantes,
                    "ajudantes_mode": "ajudantes",
                    "equipe_display_map": equipe_display_map,
                }
            except Exception:
                try:
                    cur.execute("SELECT codigo, ajudante1, ajudante2 FROM equipes ORDER BY codigo")
                    rows_equipes = []
                    for r in cur.fetchall():
                        codigo = r[0] if r else ""
                        ajudante1 = r[1] if r else ""
                        ajudante2 = r[2] if r else ""
                        display = format_equipe_nomes(ajudante1, ajudante2, codigo)
                        if display:
                            rows_equipes.append({
                                "key": upper(codigo),
                                "value": upper(codigo),
                                "label": display,
                                "nome": display,
                                "sobrenome": "",
                                "telefone": "",
                            })
                            if codigo:
                                equipe_display_map[upper(display)] = upper(codigo)
                    return {
                        "motoristas": valores_motoristas,
                        "veiculos": valores_veiculos,
                        "veiculos_lookup": veiculos_lookup,
                        "ajudantes": rows_equipes,
                        "ajudantes_mode": "equipes",
                        "equipe_display_map": equipe_display_map,
                    }
                except Exception:
                    return {
                        "motoristas": valores_motoristas,
                        "veiculos": valores_veiculos,
                        "veiculos_lookup": veiculos_lookup,
                        "ajudantes": [],
                        "ajudantes_mode": "ajudantes",
                        "equipe_display_map": {},
                    }

    def _apply_programacao_combobox_data(self, seq, data):
        if seq != self._combo_load_seq or not self.winfo_exists():
            return
        data = data or {}
        self.cb_motorista["values"] = data.get("motoristas") or []
        self.cb_veiculo["values"] = data.get("veiculos") or []
        self._veiculos_lookup = data.get("veiculos_lookup") or {}
        self._equipe_display_map = data.get("equipe_display_map") or {}
        self._set_ajudantes_options(data.get("ajudantes") or [], mode=data.get("ajudantes_mode") or "ajudantes")
        self._apply_programacao_rankings()
        self.set_status("STATUS: Carregue vendas e ajuste dados antes de salvar a programaÃ§Ã£o.")

    def _handle_programacao_combobox_error(self, seq, exc):
        if seq != self._combo_load_seq or not self.winfo_exists():
            return
        logging.error("Falha ao carregar comboboxes da Programacao", exc_info=(type(exc), exc, exc.__traceback__))
        self.set_status("STATUS: Falha ao carregar cadastros da programaÃ§Ã£o.")

    def refresh_comboboxes(self, async_load=True):
        self._combo_load_seq += 1
        seq = self._combo_load_seq
        if not async_load:
            try:
                data = self._fetch_programacao_combobox_data()
            except Exception as exc:
                self._handle_programacao_combobox_error(seq, exc)
                return
            self._apply_programacao_combobox_data(seq, data)
            return

        run_async_ui(
            self,
            self._fetch_programacao_combobox_data,
            lambda data, seq=seq: self._apply_programacao_combobox_data(seq, data),
            lambda exc, seq=seq: self._handle_programacao_combobox_error(seq, exc),
        )

    def on_show(self):
        self.set_status("STATUS: Carregue vendas e ajuste dados antes de salvar a programação.")
        self.refresh_comboboxes()
        self._on_estimativa_tipo_change()
        self._refresh_next_programacao_codigo_preview()
        self._refresh_programacao_status_badge()

    def _reset_form_after_save(self):
        """Renova a tela para nova programação após salvar."""
        try:
            self.refresh_comboboxes()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self.cb_motorista.set("")
            self.cb_veiculo.set("")
            self.cb_local_rota.set("")
            self.cb_estimativa_tipo.set("KG")
            self.ent_kg.delete(0, "end")
            self.ent_carregamento_prog.delete(0, "end")
            self.ent_adiantamento_prog.delete(0, "end")
            self.ent_adiantamento_prog.insert(0, "0,00")
            self.ent_codigo.config(state="normal")
            self.ent_codigo.delete(0, "end")
            self.ent_codigo.config(state="readonly")
            self._editing_programacao_codigo = ""
            self._loaded_venda_ids = []
            self._ajudantes_selected_keys = []
            self._refresh_ajudantes_selected_label()
            self.limpar_itens()
            self._on_estimativa_tipo_change()
            self._refresh_next_programacao_codigo_preview(force=True)
            self._refresh_programacao_status_badge()
        except Exception:
            logging.debug("Falha ignorada")

    def _on_estimativa_tipo_change(self):
        tipo = upper((self.cb_estimativa_tipo.get() if hasattr(self, "cb_estimativa_tipo") else "KG") or "KG")
        if tipo == "CX":
            self.lbl_estimado_hint.config(text="Modo FOB (Caixas estimadas)")
        else:
            self.lbl_estimado_hint.config(text="Modo CIF (KG estimado)")
        self._refresh_total_caixas_field()

    def _get_caixas_estimadas_header(self) -> int:
        try:
            tipo = upper((self.cb_estimativa_tipo.get() or "KG").strip())
        except Exception:
            tipo = "KG"
        if tipo != "CX":
            return 0
        try:
            return max(safe_int((self.ent_kg.get() or "").strip(), 0), 0)
        except Exception:
            return 0

    def _compute_total_caixas_tree(self) -> int:
        total = 0
        try:
            for iid in self.tree.get_children():
                vals = self._get_row_values(iid)
                total += safe_int(vals[4] if len(vals) > 4 else 0, 0)
        except Exception:
            logging.debug("Falha ignorada")
        return max(total, 0)

    def _refresh_total_caixas_field(self):
        if not hasattr(self, "ent_total_caixas_prog"):
            return
        total = self._compute_total_caixas_tree()
        if total <= 0:
            total = self._get_caixas_estimadas_header()
        try:
            self.ent_total_caixas_prog.config(state="normal")
            self.ent_total_caixas_prog.delete(0, "end")
            self.ent_total_caixas_prog.insert(0, str(total))
            self.ent_total_caixas_prog.config(state="readonly")
        except Exception:
            logging.debug("Falha ignorada")

    # -------------------------
    # Ações de itens
    # -------------------------
    def inserir_linha(self):
        tree_insert_aligned(self.tree, "", "end", ("", "", "", "", "1", "0.00", "0.00", "", "", ""))
        items = self.tree.get_children()
        if items:
            self.tree.see(items[-1])
        self._refresh_total_caixas_field()

    def remover_linha(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("ATENÇÃO", "Selecione uma linha para remover.")
            return
        for iid in sel:
            self.tree.delete(iid)
        self._refresh_total_caixas_field()

    def limpar_itens(self):
        self.tree.delete(*self.tree.get_children())
        self._loaded_venda_ids = []
        self._refresh_total_caixas_field()

    def _status_permite_edicao_programacao(self, status_raw: str) -> bool:
        st = upper(str(status_raw or "").strip())
        return st in {"", "ATIVA", "PENDENTE", "ABERTA", "PROGRAMADA"}

    def _status_permite_exclusao_programacao(self, status_raw: str) -> bool:
        st = upper(str(status_raw or "").strip())
        return st in {"", "ATIVA"}

    def _obter_status_programacao_para_edicao(self, codigo: str, status_local: str = ""):
        codigo = upper(str(codigo or "").strip())
        st_local = upper(str(status_local or "").strip())
        st_api = ""
        prest_api = ""
        prest_local = ""

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and codigo:
            try:
                resposta = _call_api(
                    "GET",
                    f"desktop/rotas/{codigo}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resposta.get("rota") if isinstance(resposta, dict) else None
                if isinstance(rota, dict):
                    st_api = upper(str(rota.get("status_operacional") or rota.get("status") or "").strip())
                    prest_api = upper(str(rota.get("prestacao_status") or "").strip())
            except Exception:
                logging.debug("Falha ao consultar status da programação na API central", exc_info=True)

        if codigo and not prest_api:
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    if "prestacao_status" in cols_prog:
                        cur.execute(
                            "SELECT COALESCE(prestacao_status,'') FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                            (codigo,),
                        )
                        row = cur.fetchone()
                        prest_local = upper((row[0] if row else "") or "")
            except Exception:
                logging.debug("Falha ao consultar prestacao_status local para programação", exc_info=True)

        if prest_api == "FECHADA" or prest_local == "FECHADA":
            st_ref = "PRESTACAO FECHADA"
        else:
            st_ref = st_api or st_local
        return st_ref, (prest_api if prest_api == "FECHADA" else st_api), (prest_local if prest_local == "FECHADA" else st_local)

    def _warn_bloqueio_edicao_status(self, codigo: str, status_ref: str):
        st = upper(str(status_ref or "").strip())
        if st in {"PRESTACAO FECHADA", "FECHADA"}:
            messagebox.showwarning(
                "BLOQUEADO",
                f"A programação {codigo} está com a prestação FECHADA.\n\n"
                "Não é permitido alterar ou reaproveitar essa programação no desktop."
            )
            return
        if st in {"EM_ROTA", "EM ROTA", "INICIADA"}:
            messagebox.showwarning(
                "BLOQUEADO",
                f"A programação {codigo} está em {st}.\n\n"
                "A partir desse status, a troca deve ser feita apenas no APK:\n"
                "- botão 'Substituir rota'; ou\n"
                "- botão 'Transferir caixas'."
            )
            return
        messagebox.showwarning(
            "ATENÇÃO",
            f"A programação {codigo} está com status {st or '-'}.\n"
            "Somente programação ATIVA pode ser alterada no desktop."
        )

    def _refresh_programacao_status_badge(self, codigo: str = None, status_local_hint: str = ""):
        codigo_ref = upper(str(codigo or (self.ent_codigo.get() if hasattr(self, "ent_codigo") else "") or "").strip())
        if not codigo_ref:
            self.lbl_prog_status_badge.config(text="Status atual: -", foreground="#6B7280")
            return

        status_ref, status_api, status_local = self._obter_status_programacao_para_edicao(codigo_ref, status_local_hint)
        if status_api:
            self.lbl_prog_status_badge.config(
                text=f"Status atual: {status_api} (API)",
                foreground=("#16A34A" if status_api == "ATIVA" else "#DC2626"),
            )
            return

        st_local = status_local or status_ref or "-"
        self.lbl_prog_status_badge.config(
            text=f"Status atual: {st_local} (LOCAL)",
            foreground=("#2563EB" if st_local in {"ATIVA", "PENDENTE", ""} else "#DC2626"),
        )

    def _set_motorista_combobox(self, motorista_nome: str, motorista_codigo: str):
        motorista_nome = upper(str(motorista_nome or "").strip())
        motorista_codigo = upper(str(motorista_codigo or "").strip())
        try:
            valores = [str(v or "") for v in list(self.cb_motorista["values"])]
        except Exception:
            valores = []

        alvo = self._motorista_display(motorista_nome, motorista_codigo) if motorista_nome else ""
        if alvo and alvo in valores:
            self.cb_motorista.set(alvo)
            return

        if motorista_codigo:
            tag = f"({motorista_codigo})"
            for v in valores:
                if tag in upper(v):
                    self.cb_motorista.set(v)
                    return

        if motorista_nome:
            for v in valores:
                if upper(v).startswith(motorista_nome):
                    self.cb_motorista.set(v)
                    return

        if alvo:
            self.cb_motorista.set(alvo)

    def carregar_programacao_para_edicao(self):
        codigo_atual = upper((self.ent_codigo.get() or "").strip())
        codigo = upper(
            simple_input(
                "Editar Programação",
                "Informe o código da programação para editar:",
                master=self.app,
                initial=codigo_atual,
                allow_empty=False,
            )
            or ""
        )
        if not codigo:
            return

        loaded_from_api = False
        itens_from_api = []
        motorista_nome = veiculo = equipe_raw = ""
        kg_estimado = 0.0
        status = ""
        local_rota = local_carreg = ""
        adiantamento = 0.0
        motorista_codigo = ""
        motorista_id = 0
        tipo_estimativa = "KG"
        caixas_estimado = 0

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resposta = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(codigo)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resposta.get("rota") if isinstance(resposta, dict) else None
                clientes_api = resposta.get("clientes") if isinstance(resposta, dict) else []
                if isinstance(rota, dict):
                    motorista_nome = str(rota.get("motorista") or "")
                    veiculo = str(rota.get("veiculo") or "")
                    equipe_raw = str(rota.get("equipe") or "")
                    kg_estimado = safe_float(rota.get("kg_estimado"), 0.0)
                    status = upper(str(rota.get("status") or rota.get("status_operacional") or ""))
                    local_rota = self._normalize_local_rota(rota.get("local_rota") or rota.get("tipo_rota") or "")
                    local_carreg = upper(
                        str(
                            rota.get("local_carregamento")
                            or rota.get("granja_carregada")
                            or rota.get("local_carregado")
                            or rota.get("local_carreg")
                            or ""
                        )
                    )
                    adiantamento = safe_float(rota.get("adiantamento"), safe_float(rota.get("adiantamento_rota"), 0.0))
                    motorista_codigo = upper(str(rota.get("motorista_codigo") or rota.get("codigo_motorista") or ""))
                    motorista_id = safe_int(rota.get("motorista_id"), 0)
                    tipo_estimativa = upper(str(rota.get("tipo_estimativa") or "KG").strip())
                    if tipo_estimativa not in {"KG", "CX"}:
                        tipo_estimativa = "KG"
                    caixas_estimado = safe_int(rota.get("caixas_estimado"), 0)
                    if isinstance(clientes_api, list):
                        itens_from_api = [r for r in clientes_api if isinstance(r, dict)]
                    loaded_from_api = True
            except Exception:
                logging.debug("Falha ao carregar programacao para edicao via API; usando fallback local.", exc_info=True)

        if not loaded_from_api:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}

                local_expr = "COALESCE(local_rota,'')" if "local_rota" in cols_prog else ("COALESCE(tipo_rota,'')" if "tipo_rota" in cols_prog else "''")
                carreg_expr = (
                    "COALESCE(local_carregamento,'')"
                    if "local_carregamento" in cols_prog
                    else ("COALESCE(granja_carregada,'')" if "granja_carregada" in cols_prog else ("COALESCE(local_carregado,'')" if "local_carregado" in cols_prog else "''"))
                )
                adiant_expr = (
                    "COALESCE(adiantamento, 0)"
                    if "adiantamento" in cols_prog
                    else ("COALESCE(adiantamento_rota, 0)" if "adiantamento_rota" in cols_prog else "0")
                )
                mot_cod_expr = (
                    "COALESCE(motorista_codigo, '')"
                    if "motorista_codigo" in cols_prog
                    else ("COALESCE(codigo_motorista, '')" if "codigo_motorista" in cols_prog else "''")
                )
                mot_id_expr = "COALESCE(motorista_id, 0)" if "motorista_id" in cols_prog else "0"
                tipo_estim_expr = "COALESCE(tipo_estimativa, 'KG')" if "tipo_estimativa" in cols_prog else "'KG'"
                caixas_estim_expr = "COALESCE(caixas_estimado, 0)" if "caixas_estimado" in cols_prog else "0"

                cur.execute(
                    f"""
                    SELECT
                        COALESCE(motorista, ''),
                        COALESCE(veiculo, ''),
                        COALESCE(equipe, ''),
                        COALESCE(kg_estimado, 0),
                        COALESCE(status, ''),
                        {local_expr} as local_rota,
                        {carreg_expr} as local_carreg,
                        {adiant_expr} as adiantamento,
                        {mot_cod_expr} as motorista_codigo,
                        {mot_id_expr} as motorista_id,
                        {tipo_estim_expr} as tipo_estimativa,
                        {caixas_estim_expr} as caixas_estimado
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                    """,
                    (codigo,),
                )
                row = cur.fetchone()

            if not row:
                messagebox.showwarning("ATENÇÃO", f"Programação não encontrada: {codigo}")
                return

            motorista_nome = row[0] or ""
            veiculo = row[1] or ""
            equipe_raw = row[2] or ""
            kg_estimado = safe_float(row[3], 0.0)
            status = upper(row[4] or "")
            local_rota = self._normalize_local_rota(row[5] or "")
            local_carreg = upper(row[6] or "")
            adiantamento = safe_float(row[7], 0.0)
            motorista_codigo = upper(row[8] or "")
            motorista_id = safe_int(row[9], 0)
            tipo_estimativa = upper((row[10] or "KG").strip())
            if tipo_estimativa not in {"KG", "CX"}:
                tipo_estimativa = "KG"
            caixas_estimado = safe_int(row[11], 0)

        status_ref, status_api, status_local = self._obter_status_programacao_para_edicao(codigo, status)
        if not self._status_permite_edicao_programacao(status_ref):
            self._warn_bloqueio_edicao_status(codigo, status_ref)
            return

        if not motorista_codigo and motorista_id > 0:
            try:
                resolved = False
                desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                if desktop_secret and is_desktop_api_sync_enabled():
                    try:
                        mot_resp = _call_api(
                            "GET",
                            "desktop/cadastros/motoristas",
                            extra_headers={"X-Desktop-Secret": desktop_secret},
                        )
                        for r in (mot_resp or []):
                            if not isinstance(r, dict):
                                continue
                            if safe_int(r.get("id"), 0) == safe_int(motorista_id, 0):
                                motorista_codigo = upper(str(r.get("codigo") or ""))
                                resolved = True
                                break
                    except Exception:
                        logging.debug("Falha ao resolver codigo do motorista via API; usando fallback local.", exc_info=True)

                if not resolved:
                    with get_db() as conn:
                        cur = conn.cursor()
                        cur.execute("SELECT COALESCE(codigo,'') FROM motoristas WHERE id=? LIMIT 1", (motorista_id,))
                        rr = cur.fetchone()
                        motorista_codigo = upper(rr[0] if rr else "")
            except Exception:
                logging.debug("Falha ignorada")

        self.refresh_comboboxes()
        self._set_motorista_combobox(motorista_nome, motorista_codigo)
        self.cb_veiculo.set(upper(veiculo))
        self.cb_local_rota.set(local_rota)
        self.cb_estimativa_tipo.set(tipo_estimativa)
        self.ent_kg.delete(0, "end")
        if tipo_estimativa == "CX":
            self.ent_kg.insert(0, str(max(caixas_estimado, 0)))
        else:
            self.ent_kg.insert(0, f"{kg_estimado:.2f}")
        self._on_estimativa_tipo_change()
        self.ent_carregamento_prog.delete(0, "end")
        self.ent_carregamento_prog.insert(0, local_carreg)
        self.ent_adiantamento_prog.delete(0, "end")
        self.ent_adiantamento_prog.insert(0, f"{adiantamento:.2f}".replace(".", ","))

        keys = [upper(p.strip()) for p in re.split(r"[|,;/]+", str(equipe_raw or "")) if p.strip()]
        valid_keys = {str(r.get("key", "")).strip() for r in (self._ajudantes_rows or [])}
        self._ajudantes_selected_keys = [k for k in keys if k in valid_keys]
        self._enforce_ajudantes_limit()
        self._sync_tree_selection_from_state()
        self._refresh_ajudantes_selected_label()

        itens = []
        if loaded_from_api and itens_from_api:
            itens = [
                {
                    "cod_cliente": str(it.get("cod_cliente") or ""),
                    "nome_cliente": str(it.get("nome_cliente") or ""),
                    "produto": str(it.get("produto") or ""),
                    "endereco": str(it.get("endereco") or ""),
                    "qnt_caixas": safe_int(it.get("qnt_caixas"), 0),
                    "kg": safe_float(it.get("kg"), 0.0),
                    "preco": safe_float(it.get("preco"), 0.0),
                    "vendedor": str(it.get("vendedor") or ""),
                    "pedido": str(it.get("pedido") or ""),
                    "obs": str(it.get("obs") or it.get("observacao") or ""),
                }
                for it in (itens_from_api or [])
            ]
        linked_ids = []
        if not itens:
            _itens_res = fetch_programacao_itens_contract(codigo)
            itens = (
                (_itens_res.get("data") or [])
                if isinstance(_itens_res, dict) and bool(_itens_res.get("ok", False))
                else []
            )
        if not itens:
            itens, linked_ids = self._load_linked_vendas_importadas(codigo)
        self.limpar_itens()
        self._loaded_venda_ids = list(linked_ids or [])
        for it in (itens or []):
            tree_insert_aligned(
                self.tree,
                "",
                "end",
                (
                    upper(it.get("cod_cliente", "")),
                    upper(it.get("nome_cliente", "")),
                    upper(it.get("produto", "")),
                    upper(it.get("endereco", "")),
                    str(safe_int(it.get("qnt_caixas"), 0)),
                    f"{safe_float(it.get('kg'), 0.0):.2f}",
                    f"{safe_float(it.get('preco'), 0.0):.2f}",
                    upper(it.get("vendedor", "")),
                    upper(it.get("pedido", "")),
                    upper(it.get("obs", "")),
                ),
            )

        self._refresh_total_caixas_field()
        self.ent_codigo.config(state="normal")
        self.ent_codigo.delete(0, "end")
        self.ent_codigo.insert(0, codigo)
        self.ent_codigo.config(state="readonly")
        self._editing_programacao_codigo = codigo
        self._refresh_programacao_status_badge(codigo, status_local_hint=status_local or status_ref)
        if status_api:
            self.set_status(
                f"STATUS: Programação {codigo} carregada para edição. "
                f"Status validado na API: {status_api}."
            )
        else:
            self.set_status(
                f"STATUS: Programação {codigo} carregada para edição. "
                f"Status local: {status_local or status_ref or '-'}."
            )

    def _load_linked_vendas_importadas(self, codigo_programacao: str):
        codigo_programacao = upper(str(codigo_programacao or "").strip())
        if not codigo_programacao:
            return [], []
        rows = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/vendas-importadas?codigo_programacao={urllib.parse.quote(codigo_programacao)}&limit=5000",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rows = (resp or {}).get("rows") if isinstance(resp, dict) else []
            except Exception:
                logging.debug("Falha ao carregar vendas vinculadas via API; usando fallback local.", exc_info=True)
        if not rows:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    SELECT id, pedido, cliente, nome_cliente, produto, qnt, valor_unitario, vendedor
                    FROM vendas_importadas
                    WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)
                    ORDER BY id
                    """,
                    (codigo_programacao,),
                )
                rows = [
                    {
                        "id": safe_int(r[0], 0),
                        "pedido": r[1],
                        "cliente": r[2],
                        "nome_cliente": r[3],
                        "produto": r[4],
                        "qnt": safe_float(r[5], 0.0),
                        "valor_unitario": safe_float(r[6], 0.0),
                        "vendedor": r[7],
                    }
                    for r in (cur.fetchall() or [])
                ]
        itens = []
        ids = []
        for r in (rows or []):
            if not isinstance(r, dict):
                continue
            rid = safe_int(r.get("id"), 0)
            if rid > 0:
                ids.append(rid)
            caixas = max(safe_int(r.get("qnt"), 0), 0)
            itens.append(
                {
                    "cod_cliente": str(r.get("cliente") or ""),
                    "nome_cliente": str(r.get("nome_cliente") or ""),
                    "produto": str(r.get("produto") or ""),
                    "endereco": "",
                    "qnt_caixas": caixas,
                    "kg": 0.0,
                    "preco": safe_float(r.get("valor_unitario"), 0.0),
                    "vendedor": str(r.get("vendedor") or ""),
                    "pedido": str(r.get("pedido") or ""),
                    "obs": "",
                }
            )
        return itens, ids

    def excluir_programacao(self):
        codigo = upper((self.ent_codigo.get() or "").strip())
        if not codigo:
            codigo = upper(
                simple_input(
                    "Excluir Programação",
                    "Informe o código da programação para excluir:",
                    master=self.app,
                    allow_empty=False,
                )
                or ""
            )
        if not codigo:
            return

        status_ref, status_api, status_local = self._obter_status_programacao_para_edicao(codigo)
        if not self._status_permite_exclusao_programacao(status_ref):
            messagebox.showwarning(
                "ATENÇÃO",
                f"A programação {codigo} está com status {status_api or status_local or status_ref or '-'}.\n"
                "Somente programações ATIVAS podem ser excluídas.",
            )
            return

        if not messagebox.askyesno(
            "Excluir Programação",
            f"Tem certeza que deseja excluir a programação {codigo}?\n\nEssa ação não pode ser desfeita.",
        ):
            return

        escolha_vendas = messagebox.askyesnocancel(
            "Vendas Vinculadas",
            "O que deseja fazer com as vendas desta programação?\n\n"
            "Sim: voltar as vendas para a tela Importar Vendas.\n"
            "Não: excluir as vendas também.\n"
            "Cancelar: abortar exclusão.",
        )
        if escolha_vendas is None:
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_delete_error = None
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                _call_api(
                    "DELETE",
                    f"desktop/rotas/{urllib.parse.quote(codigo)}?delete_vendas={1 if not escolha_vendas else 0}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
            except SyncError as exc:
                api_delete_error = exc
                msg = str(exc)
                if msg.startswith("409 "):
                    messagebox.showwarning("ATENÇÃO", msg.split(":", 1)[-1].strip() or msg)
                else:
                    messagebox.showerror(
                        "ERRO",
                        "Nao foi possivel excluir a programacao na API central.\n\n"
                        "A exclusao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                        f"Detalhe: {msg}",
                    )
                return
            except Exception as exc:
                api_delete_error = exc
                logging.debug("Falha ao excluir programacao na API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Nao foi possivel excluir a programacao na API central.\n\n"
                    "A exclusao local foi bloqueada para evitar divergencia entre servidor e desktop.",
                )
                return

        try:
            with get_db() as conn:
                cur = conn.cursor()

                if escolha_vendas:
                    cur.execute(
                        """
                        UPDATE vendas_importadas
                        SET usada=0,
                            usada_em='',
                            codigo_programacao='',
                            selecionada=0
                        WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)
                        """,
                        (codigo,),
                    )
                else:
                    cur.execute(
                        "DELETE FROM vendas_importadas WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)",
                        (codigo,),
                    )

                for tabela in (
                    "programacao_itens_log",
                    "programacao_itens_controle",
                    "programacao_itens",
                    "recebimentos",
                    "despesas",
                    "rota_gps_pings",
                    "rota_substituicoes",
                    "cliente_localizacao_amostras",
                ):
                    try:
                        cur.execute(f"DELETE FROM {tabela} WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)", (codigo,))
                    except Exception:
                        logging.debug("Falha ao excluir registros da tabela %s para a programacao %s", tabela, codigo, exc_info=True)

                try:
                    cur.execute(
                        """
                        DELETE FROM transferencias
                        WHERE UPPER(COALESCE(codigo_origem,''))=UPPER(?)
                           OR UPPER(COALESCE(codigo_destino,''))=UPPER(?)
                        """,
                        (codigo, codigo),
                    )
                except Exception:
                    logging.debug("Falha ao excluir transferencias da programacao %s", codigo, exc_info=True)

                cur.execute(
                    "DELETE FROM programacoes WHERE UPPER(COALESCE(codigo_programacao,''))=UPPER(?)",
                    (codigo,),
                )

            self._reset_form_after_save()
            self._editing_programacao_codigo = ""
            self._refresh_programacao_status_badge()
            if hasattr(self.app, "refresh_programacao_comboboxes"):
                self.app.refresh_programacao_comboboxes()
            try:
                page_imp = self.app.pages.get("ImportarVendas") if hasattr(self.app, "pages") else None
                if page_imp and hasattr(page_imp, "carregar"):
                    page_imp.carregar()
            except Exception:
                logging.debug("Falha ao atualizar tela ImportarVendas apos exclusao da programacao.", exc_info=True)

            if api_delete_error:
                messagebox.showwarning(
                    "Exclusão Parcial",
                    f"A programação {codigo} foi excluída localmente, mas houve falha ao remover na API.\n\nDetalhe: {api_delete_error}",
                )
            else:
                messagebox.showinfo("OK", f"Programação {codigo} excluída com sucesso.")
            self.set_status(f"STATUS: Programação {codigo} excluída.")
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao excluir programação {codigo}: {e}")

    def carregar_vendas_selecionadas(self):
        """
        Carrega todas as vendas selecionadas.
        API-first com fallback local.
        """
        def _is_bad(v):
            return upper(str(v or "").strip()) in {"", "NAN", "NAT", "NONE", "NULL", "<NA>"}

        def _clean_pedido_local(v):
            s = str(v or "").strip()
            if _is_bad(s):
                return ""
            try:
                f = float(str(s).replace(",", "."))
                if abs(f - int(f)) < 1e-9:
                    return str(int(f))
                return str(f).rstrip("0").rstrip(".")
            except Exception:
                return upper(s)

        vendas = []
        codigo_vinculo = upper((self.ent_codigo.get() or "").strip())
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp_v = _call_api(
                    "GET",
                    "desktop/vendas-importadas?limit=5000",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rows_v = resp_v.get("rows") if isinstance(resp_v, dict) else []

                resp_c = _call_api(
                    "GET",
                    "desktop/clientes/base?ordem=codigo&limit=1000",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rows_c = resp_c if isinstance(resp_c, list) else []
                end_map = {
                    upper(str((r or {}).get("cod_cliente") or "")): upper(str((r or {}).get("endereco") or ""))
                    for r in rows_c
                    if isinstance(r, dict)
                }

                for r in (rows_v or []):
                    if not isinstance(r, dict):
                        continue
                    if safe_int(r.get("selecionada"), 0) != 1:
                        continue
                    vendas.append(
                        (
                            safe_int(r.get("id"), 0),
                            r.get("pedido"),
                            r.get("data_venda"),
                            r.get("cliente"),
                            r.get("nome_cliente"),
                            r.get("vendedor"),
                            r.get("produto"),
                            safe_float(r.get("vr_total"), 0.0),
                            safe_float(r.get("qnt"), 0.0),
                            0.0,
                            "",
                            r.get("cidade"),
                            end_map.get(upper(str(r.get("cliente") or "")), ""),
                        )
                    )
            except Exception:
                logging.debug("Falha ao carregar vendas selecionadas via API; usando fallback local.", exc_info=True)

        if not vendas:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT
                        v.id,
                        v.pedido,
                        v.data_venda,
                        v.cliente AS cod_cliente,
                        v.nome_cliente,
                        v.vendedor,
                        v.produto,
                        v.vr_total,
                        v.qnt,
                        v.valor_unitario,
                        v.observacao,
                        v.cidade,
                        COALESCE(c.endereco, '') AS endereco
                    FROM vendas_importadas v
                    LEFT JOIN clientes c
                           ON c.cod_cliente = v.cliente
                    WHERE v.selecionada = 1
                      AND IFNULL(v.usada, 0) = 0
                    ORDER BY v.id ASC
                """)
                vendas = cur.fetchall() or []

        self.limpar_itens()
        self._loaded_venda_ids = []

        seen = set()
        ignorados_invalidos = 0
        for (venda_id, pedido, data_venda, cod_cliente, nome_cliente, vendedor, produto, vr_total, qnt, valor_unit, obs, cidade, endereco) in vendas:
            cod_cliente = upper(str(cod_cliente or "").strip())
            pedido_u = _clean_pedido_local(pedido)
            nome_u = upper(str(nome_cliente or "").strip())
            produto_u = upper(str(produto or "").strip())

            if _is_bad(cod_cliente) or _is_bad(pedido_u) or _is_bad(nome_u) or _is_bad(produto_u):
                ignorados_invalidos += 1
                continue

            key = (pedido_u, cod_cliente, produto_u)
            if key in seen:
                continue
            seen.add(key)
            self._loaded_venda_ids.append(safe_int(venda_id, 0))

            caixas = 1
            kg = 0.0
            vr_total_f = safe_float(vr_total, 0.0)
            qnt_f = safe_float(qnt, 0.0)
            valor_unit_f = safe_float(valor_unit, 0.0)
            if qnt_f > 0:
                preco = vr_total_f / qnt_f
            else:
                preco = valor_unit_f
            cidade_txt = upper(cidade) if cidade else ""
            endereco_txt = upper(endereco) if endereco else ""
            endereco_final = cidade_txt if cidade_txt else endereco_txt
            observacao = upper(obs) if obs else ""

            if obs:
                low = str(obs).lower()
                m = re.search(r"(\d+[\,\.]?\d*)\s*kg", low)
                if m:
                    kg = safe_float(m.group(1), 0.0)

                m2 = re.search(r"(\d+)\s*cx", low)
                if m2:
                    caixas = safe_int(m2.group(1), 1)

            tree_insert_aligned(self.tree, "", "end", (
                cod_cliente,
                nome_u,
                produto_u,
                endereco_final,
                str(caixas),
                f"{float(kg):.2f}",
                f"{float(preco):.2f}",
                upper(vendedor),
                upper(pedido),
                observacao
            ))

        self._refresh_total_caixas_field()
        qtd_itens = len(self.tree.get_children())
        if codigo_vinculo:
            msg = (
                f"STATUS: {qtd_itens} vendas carregadas para a programacao {codigo_vinculo}. "
                f"Clique em SALVAR PROGRAMACAO para vincular os itens."
            )
        else:
            msg = f"STATUS: Itens carregados: {qtd_itens} vendas selecionadas (nao usadas). (edite antes de salvar)"
        if ignorados_invalidos:
            msg += f" Ignoradas invalidas: {ignorados_invalidos}."
        self.set_status(msg)

    def _start_edit_cell(self, event=None):
        if self._editing:
            self._commit_edit()

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        col_index = int(col_id.replace("#", "")) - 1
        cols = list(self.tree["columns"])

        if col_index < 0 or col_index >= len(cols):
            return

        col_name = cols[col_index]

        # âÅâ€œââ‚¬¦ Regra: só permite editar essas colunas
        editable = {"ENDERECO", "CAIXAS", "KG", "PRECO", "VENDEDOR", "PEDIDO", "OBS"}
        if col_name not in editable:
            return

        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return

        vals = self._get_row_values(row_id)
        value = vals[col_index]

        x, y, w, h = bbox
        entry = ttk.Entry(self.tree, style="Field.TEntry")
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.focus_set()
        entry.selection_range(0, "end")

        entry.bind("<Return>", lambda e: self._commit_edit())
        entry.bind("<Escape>", lambda e: self._cancel_edit())
        entry.bind("<FocusOut>", lambda e: self._commit_edit())

        self._editing = (row_id, col_index, col_name, entry)

    def _commit_edit(self):
        if not self._editing:
            return

        row_id, col_index, col_name, entry = self._editing
        new_value = str(entry.get() or "").strip()

        # âÅâ€œââ‚¬¦ Validação por tipo
        if col_name == "CAIXAS":
            v = safe_int(new_value, 0)
            if v < 0:
                v = 0
            new_value = str(v)

        elif col_name in {"KG", "PRECO"}:
            v = safe_float(new_value, 0.0)
            if v < 0:
                v = 0.0
            new_value = f"{v:.2f}"

        vals = self._get_row_values(row_id)
        vals[col_index] = new_value
        self.tree.item(row_id, values=tuple(vals))
        if col_name == "CAIXAS":
            self._refresh_total_caixas_field()

        try:
            entry.destroy()
        except Exception:
            logging.debug("Falha ignorada")
        self._editing = None

    def _cancel_edit(self):
        if not self._editing:
            return
        _, _, _, entry = self._editing
        try:
            entry.destroy()
        except Exception:
            logging.debug("Falha ignorada")
        self._editing = None

    # -------------------------
    # Compatibilidade API / DB
    # -------------------------
    def _ensure_prog_columns_for_api(self, cur):
        # Garante colunas que a API usa (sem quebrar bases antigas)
        try:
            cur.execute("PRAGMA table_info(programacoes)")
            cols = [str(r[1]).lower() for r in cur.fetchall()]
        except Exception:
            cols = []

        def add_col(name: str, coltype: str):
            if name.lower() not in cols:
                try:
                    cur.execute(f"ALTER TABLE programacoes ADD COLUMN {name} {coltype}")
                    cols.append(name.lower())
                except Exception as e:
                    logging.exception("Falha ao adicionar coluna programacoes.%s: %s", name, e)

        add_col("motorista_id", "INTEGER")
        add_col("motorista_codigo", "TEXT")
        add_col("codigo_motorista", "TEXT")
        add_col("codigo", "TEXT")
        add_col("data", "TEXT")
        add_col("total_caixas", "INTEGER DEFAULT 0")
        add_col("quilos", "REAL DEFAULT 0")
        add_col("saida_dt", "TEXT")
        add_col("chegada_dt", "TEXT")
        add_col("tipo_estimativa", "TEXT DEFAULT 'KG'")
        add_col("caixas_estimado", "INTEGER DEFAULT 0")
        add_col("usuario_criacao", "TEXT")
        add_col("usuario_ultima_edicao", "TEXT")
        add_col("status_operacional", "TEXT")
        add_col("finalizada_no_app", "INTEGER DEFAULT 0")

    def _ensure_vendas_usada_cols(self, cur):
        """Garante colunas para bloquear reutilização (compatÃÂÂvel com bases antigas)."""
        try:
            cur.execute("PRAGMA table_info(vendas_importadas)")
            cols = [str(r[1]).lower() for r in cur.fetchall()]
        except Exception:
            cols = []

        def add_col(name: str, coltype: str):
            if name.lower() not in cols:
                try:
                    cur.execute(f"ALTER TABLE vendas_importadas ADD COLUMN {name} {coltype}")
                    cols.append(name.lower())
                except Exception as e:
                    logging.exception("Falha ao adicionar coluna vendas_importadas.%s: %s", name, e)

        add_col("usada", "INTEGER DEFAULT 0")
        add_col("usada_em", "TEXT")
        add_col("codigo_programacao", "TEXT")

    def salvar_programacao(self):
        if self._editing:
            self._commit_edit()
        if not ensure_system_api_binding(context="Salvar programacao", parent=self):
            return

        motorista_sel = upper(self.cb_motorista.get()).strip()
        motorista_nome, motorista_codigo = self._parse_motorista_display(motorista_sel)

        veiculo = upper(self.cb_veiculo.get()).strip()
        ajudantes_sel = self._get_selected_ajudantes()
        ajudante1 = ajudantes_sel[0] if len(ajudantes_sel) >= 1 else ""
        ajudante2 = ajudantes_sel[1] if len(ajudantes_sel) >= 2 else ""
        if self._ajudantes_mode == "equipes":
            equipe = ajudante1
        else:
            equipe = f"{ajudante1}|{ajudante2}" if (ajudante1 and ajudante2) else (ajudante1 or ajudante2)
        local_rota = self._normalize_local_rota(self.cb_local_rota.get())
        local_carreg = upper((self.ent_carregamento_prog.get() or "").strip())
        tipo_estimativa = upper((self.cb_estimativa_tipo.get() or "KG").strip())
        if tipo_estimativa not in {"KG", "CX"}:
            tipo_estimativa = "KG"
        estimativa_raw = (self.ent_kg.get() or "").strip()
        if tipo_estimativa == "CX":
            caixas_estimado = safe_int(estimativa_raw, 0)
            kg_estimado = 0.0
        else:
            kg_estimado = safe_float(estimativa_raw, 0.0)
            caixas_estimado = 0
        usuario_logado = upper(str((getattr(self.app, "user", {}) or {}).get("nome", "")).strip()) or "ADMIN"
        adiantamento_val = safe_money(self.ent_adiantamento_prog.get(), 0.0)
        if adiantamento_val < 0:
            messagebox.showwarning("ATENÇÃO", "Adiantamento não pode ser negativo.")
            return
        if tipo_estimativa == "CX":
            if caixas_estimado <= 0:
                messagebox.showwarning("ATENÇÃO", "Informe a estimativa em caixas (CX) para FOB.")
                return
        else:
            if kg_estimado <= 0:
                messagebox.showwarning("ATENÇÃO", "Informe o KG estimado para CIF.")
                return

        if not motorista_nome or not veiculo:
            messagebox.showwarning("ATENCAO", "Selecione Motorista e Veiculo.")
            return
        if self._ajudantes_mode == "equipes":
            if not ajudante1:
                messagebox.showwarning("ATENÇÃO", "Selecione a equipe da programação.")
                return
        else:
            if not ajudante1 or not ajudante2:
                messagebox.showwarning("ATENÇÃO", "Selecione exatamente 2 ajudantes da programação.")
                return
            if upper(ajudante1) == upper(ajudante2):
                messagebox.showwarning("ATENÇÃO", "Os ajudantes selecionados devem ser diferentes.")
                return
        if local_rota not in {"SERRA", "SERTAO"}:
            messagebox.showwarning("ATENÇÃO", "Selecione o Local da Rota (SERRA ou SERTAO).")
            return
        if not local_carreg:
            messagebox.showwarning("ATENCAO", "Informe o local de Carregamento.")
            return

        codigo = None
        codigo_atual = upper((self.ent_codigo.get() or "").strip())
        explicit_edit_mode = upper(str(self._editing_programacao_codigo or "").strip()) == codigo_atual

        itens = []
        for iid in self.tree.get_children():
            itens.append(self._get_row_values(iid))

        itens_salvar = list(itens)
        preserving_existing_items = False
        if not itens_salvar and codigo_atual and explicit_edit_mode:
            itens_existentes = self._load_existing_programacao_item_rows(codigo_atual)
            if itens_existentes:
                itens_salvar = itens_existentes
                preserving_existing_items = True

        # âÅâ€œââ‚¬¦ validação mÃÂÂnima por linha (segurança)
        for v in itens_salvar:
            cod_cliente = v[0]
            nome_cliente = v[1]
            if not str(cod_cliente).strip() or not str(nome_cliente).strip():
                messagebox.showwarning("ATENÇÃO", "Há linhas sem COD CLIENTE ou NOME CLIENTE. Corrija antes de salvar.")
                return

        is_update_existing = False
        data_criacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        creating_base_programacao = not itens and not preserving_existing_items and not codigo_atual

        # Totais (compatibilidade API)
        try:
            total_caixas = sum(safe_int(v[4], 0) for v in itens_salvar)
        except Exception:
            total_caixas = 0
        if total_caixas <= 0:
            total_caixas = self._get_caixas_estimadas_header()
        if total_caixas <= 0 and tipo_estimativa == "CX":
            total_caixas = max(safe_int(caixas_estimado, 0), 0)
        try:
            total_quilos = round(sum(safe_float(v[5], 0.0) for v in itens_salvar), 2)
        except Exception:
            total_quilos = 0.0
        if total_quilos <= 0 and tipo_estimativa == "KG":
            total_quilos = round(max(safe_float(kg_estimado, 0.0), 0.0), 2)

        caixas_programadas_sync = max(safe_int(total_caixas, 0), 0)
        kg_programado_sync = round(max(safe_float(total_quilos, 0.0), 0.0), 2)
        nf_kg_sync = kg_programado_sync if tipo_estimativa == "KG" else 0.0
        nf_caixas_sync = caixas_programadas_sync if caixas_programadas_sync > 0 else max(safe_int(caixas_estimado, 0), 0)

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                vei_resp = _call_api(
                    "GET",
                    "desktop/cadastros/veiculos",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                capacidade_cx = -1
                for r in (vei_resp or []):
                    if not isinstance(r, dict):
                        continue
                    if upper(r.get("placa") or "") == upper(veiculo):
                        capacidade_cx = safe_int(r.get("capacidade_cx"), -1)
                        break
                if capacidade_cx >= 0:
                    caixas_para_validar = caixas_estimado if tipo_estimativa == "CX" else total_caixas
                    if caixas_para_validar > capacidade_cx:
                        messagebox.showwarning(
                            "ATENÇÃO",
                            f"Capacidade excedida para o veiculo {veiculo}.\n\n"
                            f"Caixas na programação: {caixas_para_validar}\n"
                            f"Capacidade do veiculo: {capacidade_cx}"
                        )
                        return
            except Exception:
                logging.debug("Falha ao validar capacidade via API; usando validacao local/fallback.", exc_info=True)

        motorista_id = None
        api_saved = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        conn = None

        try:
            if desktop_secret and is_desktop_api_sync_enabled():
                itens_payload = []
                for (cod_cliente, nome_cliente, produto, endereco, caixas, kg, preco, vendedor, pedido, obs) in itens_salvar:
                    itens_payload.append({
                        "cod_cliente": upper(cod_cliente),
                        "nome_cliente": upper(nome_cliente),
                        "qnt_caixas": safe_int(caixas, 0),
                        "kg": safe_float(kg, 0.0),
                        "preco": safe_float(preco, 0.0),
                        "endereco": upper(endereco),
                        "vendedor": upper(vendedor),
                        "pedido": upper(pedido),
                        "produto": upper(produto),
                        "obs": upper(obs),
                    })

                try:
                    mot_resp = _call_api(
                        "GET",
                        "desktop/cadastros/motoristas",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    for r in (mot_resp or []):
                        if not isinstance(r, dict):
                            continue
                        cod_r = upper(r.get("codigo") or "")
                        nome_r = upper(r.get("nome") or "")
                        if (motorista_codigo and cod_r == upper(motorista_codigo)) or (nome_r == upper(motorista_nome)):
                            motorista_id = safe_int(r.get("id"), 0)
                            break
                except Exception:
                    motorista_id = None

                explicit_edit_mode = upper(str(self._editing_programacao_codigo or "").strip()) == codigo_atual
                is_update_existing = bool(codigo_atual and explicit_edit_mode)

                max_tries = 5 if not is_update_existing else 1
                last_api_error = None
                next_codigo_try = codigo_atual if is_update_existing else self._next_programacao_codigo()
                for _ in range(max_tries):
                    codigo_try = next_codigo_try
                    payload_sync = {
                        "codigo_programacao": codigo_try,
                        "data_criacao": data_criacao,
                        "motorista": motorista_nome,
                        "motorista_id": safe_int(motorista_id, 0),
                        "motorista_codigo": motorista_codigo,
                        "codigo_motorista": motorista_codigo,
                        "veiculo": veiculo,
                        "equipe": equipe,
                        "kg_estimado": safe_float(kg_estimado, 0.0),
                        "tipo_estimativa": tipo_estimativa,
                        "caixas_estimado": safe_int(caixas_estimado, 0),
                        "status": "ATIVA",
                        "local_rota": local_rota,
                        "tipo_rota": local_rota,
                        "local_carregamento": local_carreg,
                        "local_carregado": local_carreg,
                        "granja_carregada": local_carreg,
                        "local_carreg": local_carreg,
                        "adiantamento": safe_float(adiantamento_val, 0.0),
                        "total_caixas": caixas_programadas_sync,
                        "quilos": kg_programado_sync,
                        "nf_kg": nf_kg_sync,
                        "nf_caixas": nf_caixas_sync,
                        "caixas_carregadas": nf_caixas_sync,
                        "usuario_criacao": usuario_logado,
                        "usuario_ultima_edicao": usuario_logado,
                        "linked_venda_ids": [safe_int(x, 0) for x in (self._loaded_venda_ids or []) if safe_int(x, 0) > 0],
                        "vendas_usada_em": data_criacao,
                        "itens": itens_payload,
                    }
                    try:
                        _call_api(
                            "POST",
                            "desktop/rotas/upsert",
                            payload=payload_sync,
                            extra_headers={"X-Desktop-Secret": desktop_secret},
                        )
                        codigo = codigo_try
                        api_saved = True
                        break
                    except Exception as exc_try:
                        last_api_error = exc_try
                        if is_update_existing:
                            break
                        next_codigo_try = self._increment_programacao_codigo(codigo_try)
                if (not api_saved) and last_api_error:
                    logging.warning(
                        "Falha ao salvar programacao via API. Erro: %s",
                        last_api_error,
                    )
                    messagebox.showerror(
                        "ERRO",
                        "Nao foi possivel salvar a programacao na API central.\n\n"
                        "A gravacao local foi bloqueada para evitar divergencia entre servidor e desktop.\n\n"
                        f"Detalhe: {last_api_error}",
                    )
                    return

            if not api_saved:
                conn = db_connect()
                conn.row_factory = sqlite3.Row
                cur = conn.cursor()

                if not self._prog_cols_checked:
                    self._ensure_prog_columns_for_api(cur)
                    self._prog_cols_checked = True

                if not self._vendas_cols_checked:
                    self._ensure_vendas_usada_cols(cur)
                    self._vendas_cols_checked = True

                # Valida capacidade do veiculo (CX) antes de salvar programacao.
                cap_col = "capacidade_cx"
                try:
                    cur.execute("PRAGMA table_info(veiculos)")
                    cols_vei = [str(r[1]).lower() for r in cur.fetchall()]
                    if "capacidade_cx" not in cols_vei and "capacidade_c" in cols_vei:
                        cap_col = "capacidade_c"
                except Exception:
                    cap_col = "capacidade_cx"

                veiculo_info = self._veiculos_lookup.get(upper(veiculo)) or {}
                if not veiculo_info and veiculo:
                    veiculo_info = {
                        "placa": upper(veiculo),
                        "capacidade_cx": -1,
                    }
                try:
                    cur.execute(
                        f"SELECT {cap_col} FROM veiculos WHERE UPPER(placa)=UPPER(?) LIMIT 1",
                        (veiculo,),
                    )
                    vrow = cur.fetchone()
                except Exception:
                    vrow = None

                if not vrow:
                    capacidade_cx = safe_int(veiculo_info.get("capacidade_cx"), -1)
                    if not veiculo_info:
                        messagebox.showwarning(
                            "ATENÇÃO",
                            f"Veiculo nao encontrado no cadastro: {veiculo}.",
                        )
                        return
                else:
                    capacidade_cx = safe_int(vrow[0], -1)
                if capacidade_cx < 0:
                    logging.warning(
                        "Capacidade (CX) indisponivel para o veiculo %s; seguindo sem validacao de capacidade.",
                        veiculo,
                    )
                    capacidade_cx = -1

                caixas_para_validar = caixas_estimado if tipo_estimativa == "CX" else total_caixas
                if capacidade_cx >= 0 and caixas_para_validar > capacidade_cx:
                    messagebox.showwarning(
                        "ATENÇÃO",
                        f"Capacidade excedida para o veiculo {veiculo}.\n\n"
                        f"Caixas na programação: {caixas_para_validar}\n"
                        f"Capacidade do veiculo: {capacidade_cx}"
                    )
                    return

                # Resolve motorista_id priorizando codigo, depois nome
                try:
                    if motorista_codigo:
                        cur.execute("SELECT id FROM motoristas WHERE UPPER(codigo)=UPPER(?) LIMIT 1", (motorista_codigo,))
                    else:
                        cur.execute("SELECT id FROM motoristas WHERE UPPER(nome)=UPPER(?) LIMIT 1", (motorista_nome,))
                    r = cur.fetchone()
                    if r:
                        motorista_id = safe_int(r[0], 0)
                except Exception:
                    motorista_id = None

                motorista_id_db = safe_int(motorista_id, 0)
                if motorista_id_db <= 0:
                    motorista_id_db = None

                row_exist = None
                status_exist = ""
                explicit_edit_mode = False
                if codigo_atual:
                    try:
                        cur.execute(
                            "SELECT COALESCE(status, '') FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                            (codigo_atual,),
                        )
                        row_exist = cur.fetchone()
                        status_exist = upper((row_exist[0] if row_exist else "") or "")
                        explicit_edit_mode = upper(str(self._editing_programacao_codigo or "").strip()) == codigo_atual
                    except Exception:
                        row_exist = None
                        status_exist = ""
                        explicit_edit_mode = False

                # Regra: cada programacao deve ser unica.
                # So permite atualizar codigo existente quando veio do fluxo explicito de "Editar Programacao".
                if row_exist and not explicit_edit_mode:
                    row_exist = None

                if row_exist:
                    status_ref, _, _ = self._obter_status_programacao_para_edicao(codigo_atual, status_exist)
                    if not self._status_permite_edicao_programacao(status_ref):
                        self._warn_bloqueio_edicao_status(codigo_atual, status_ref)
                        return

                    codigo = codigo_atual
                    is_update_existing = True
                    cur.execute(
                        """
                        UPDATE programacoes
                        SET
                            motorista=?,
                            motorista_id=?,
                            veiculo=?,
                            equipe=?,
                            kg_estimado=?,
                            tipo_estimativa=?,
                            caixas_estimado=?,
                            usuario_ultima_edicao=?,
                            total_caixas=?,
                            quilos=?,
                            status='ATIVA',
                            status_operacional=NULL,
                            finalizada_no_app=0
                        WHERE codigo_programacao=?
                        """,
                        (
                            motorista_nome,
                            motorista_id_db,
                            veiculo,
                            equipe,
                            kg_estimado,
                            tipo_estimativa,
                            caixas_estimado,
                            usuario_logado,
                            total_caixas,
                            total_quilos,
                            codigo,
                        ),
                    )

                # Insert compatÃÂÂvel (tenta gerar código único)
                if not row_exist:
                    inserted = False
                    codigo = self._next_programacao_codigo()
                    for _ in range(5):
                        try:
                            cur.execute(
                                "SELECT 1 FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                                (codigo,)
                            )
                            if cur.fetchone():
                                codigo = self._increment_programacao_codigo(codigo)
                                continue
                        except Exception:
                            logging.debug("Falha ignorada")
                        try:
                            cur.execute("""
                                INSERT INTO programacoes (
                                    codigo_programacao, codigo, data, data_criacao,
                                    motorista, motorista_id, veiculo, equipe,
                                    kg_estimado, tipo_estimativa, caixas_estimado,
                                    usuario_criacao, usuario_ultima_edicao,
                                    status, status_operacional, finalizada_no_app, total_caixas, quilos
                                )
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ATIVA', NULL, 0, ?, ?)
                            """, (
                                codigo, codigo, data_criacao, data_criacao,
                                motorista_nome, motorista_id_db, veiculo, equipe,
                                kg_estimado, tipo_estimativa, caixas_estimado,
                                usuario_logado, usuario_logado,
                                total_caixas, total_quilos
                            ))
                            inserted = True
                            break
                        except sqlite3.IntegrityError:
                            codigo = self._increment_programacao_codigo(codigo)
                            continue
                        except Exception:
                            try:
                                cur.execute("""
                                    INSERT INTO programacoes (codigo_programacao, data_criacao, motorista, veiculo, equipe, kg_estimado, status)
                                    VALUES (?, ?, ?, ?, ?, ?, 'ATIVA')
                                """, (codigo, data_criacao, motorista_nome, veiculo, equipe, kg_estimado))
                                inserted = True
                                break
                            except sqlite3.IntegrityError:
                                codigo = self._increment_programacao_codigo(codigo)
                                continue

                    if not inserted:
                        raise sqlite3.IntegrityError("Falha ao gerar código único para programação.")

                # Salva local da rota com compatibilidade de bases.
                try:
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols_prog = [str(r[1]).lower() for r in cur.fetchall()]
                    if "tipo_rota" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET tipo_rota=? WHERE codigo_programacao=?",
                            (local_rota, codigo),
                        )
                    if "local_rota" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET local_rota=? WHERE codigo_programacao=?",
                            (local_rota, codigo),
                        )
                    if "local_carregamento" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET local_carregamento=? WHERE codigo_programacao=?",
                            (local_carreg, codigo),
                        )
                    if "granja_carregada" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET granja_carregada=? WHERE codigo_programacao=?",
                            (local_carreg, codigo),
                        )
                    if "local_carregado" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET local_carregado=? WHERE codigo_programacao=?",
                            (local_carreg, codigo),
                        )
                    if "local_carreg" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET local_carreg=? WHERE codigo_programacao=?",
                            (local_carreg, codigo),
                        )
                    if "tipo_estimativa" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET tipo_estimativa=? WHERE codigo_programacao=?",
                            (tipo_estimativa, codigo),
                        )
                    if "caixas_estimado" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET caixas_estimado=? WHERE codigo_programacao=?",
                            (caixas_estimado, codigo),
                        )
                    if "usuario_criacao" in cols_prog:
                        cur.execute(
                            """
                            UPDATE programacoes
                               SET usuario_criacao = COALESCE(NULLIF(TRIM(usuario_criacao), ''), ?)
                             WHERE codigo_programacao=?
                            """,
                            (usuario_logado, codigo),
                        )
                    if "usuario_ultima_edicao" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET usuario_ultima_edicao=? WHERE codigo_programacao=?",
                            (usuario_logado, codigo),
                        )
                except Exception:
                    logging.debug("Falha ignorada")

                # Persiste tambem o codigo do motorista para filtro estavel no app/API.
                if motorista_codigo:
                    try:
                        cur.execute("PRAGMA table_info(programacoes)")
                        cols_prog = [str(r[1]).lower() for r in cur.fetchall()]
                    except Exception:
                        cols_prog = []
                    try:
                        if "motorista_codigo" in cols_prog:
                            cur.execute(
                                "UPDATE programacoes SET motorista_codigo=? WHERE codigo_programacao=?",
                                (motorista_codigo, codigo),
                            )
                        if "codigo_motorista" in cols_prog:
                            cur.execute(
                                "UPDATE programacoes SET codigo_motorista=? WHERE codigo_programacao=?",
                                (motorista_codigo, codigo),
                            )
                    except Exception:
                        logging.debug("Falha ignorada")

                # Salva adiantamento com compatibilidade entre colunas novas/legadas.
                try:
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols_prog = [str(r[1]).lower() for r in cur.fetchall()]
                    if "adiantamento" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET adiantamento=? WHERE codigo_programacao=?",
                            (adiantamento_val, codigo),
                        )
                    if "adiantamento_rota" in cols_prog:
                        cur.execute(
                            "UPDATE programacoes SET adiantamento_rota=? WHERE codigo_programacao=?",
                            (adiantamento_val, codigo),
                        )
                except Exception:
                    logging.debug("Falha ignorada")

                # garante prestação pendente quando coluna existir
                try:
                    if db_has_column(cur, "programacoes", "prestacao_status"):
                        cur.execute(
                            "UPDATE programacoes SET prestacao_status=COALESCE(NULLIF(TRIM(prestacao_status), ''), 'PENDENTE') WHERE codigo_programacao=?",
                            (codigo,)
                        )
                except Exception:
                    logging.debug("Falha ignorada")

                # Itens
                replace_existing_items = bool(itens)
                if is_update_existing and replace_existing_items:
                    cur.execute("DELETE FROM programacao_itens WHERE codigo_programacao=?", (codigo,))
                    # Limpa estado operacional antigo para evitar herdar "finalizada/entregue"
                    # de execucoes anteriores ao replanejar a mesma programacao.
                    for tbl in ("programacao_itens_controle", "recebimentos", "despesas", "rota_gps_pings", "rota_substituicoes"):
                        try:
                            cur.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1", (tbl,))
                            if cur.fetchone():
                                cur.execute(f"DELETE FROM {tbl} WHERE codigo_programacao=?", (codigo,))
                        except Exception:
                            logging.debug("Falha ignorada")
                if (not is_update_existing) or replace_existing_items:
                    for (cod_cliente, nome_cliente, produto, endereco, caixas, kg, preco, vendedor, pedido, obs) in itens_salvar:
                        cur.execute("""
                            INSERT INTO programacao_itens (
                                codigo_programacao, cod_cliente, nome_cliente,
                                qnt_caixas, kg, preco, endereco, vendedor, pedido, produto, observacao
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            codigo,
                            upper(cod_cliente),
                            upper(nome_cliente),
                            safe_int(caixas, 0),
                            safe_float(kg, 0.0),
                            safe_float(preco, 0.0),
                            upper(endereco),
                            upper(vendedor),
                            upper(pedido),
                            upper(produto),
                            upper(obs)
                        ))

                        # Mantém clientes atualizados (compatÃÂÂvel com bases antigas)
                        if cod_cliente and nome_cliente:
                            try:
                                cur.execute("""
                                    INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone, vendedor)
                                    VALUES (?, ?, ?, '', ?)
                                    ON CONFLICT(cod_cliente) DO UPDATE SET
                                        nome_cliente=excluded.nome_cliente,
                                        endereco=COALESCE(NULLIF(excluded.endereco,''), clientes.endereco),
                                        vendedor=COALESCE(NULLIF(excluded.vendedor,''), clientes.vendedor)
                                """, (upper(cod_cliente), upper(nome_cliente), upper(endereco), upper(vendedor)))
                            except Exception:
                                cur.execute("""
                                    INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone)
                                    VALUES (?, ?, ?, '')
                                    ON CONFLICT(cod_cliente) DO UPDATE SET
                                        nome_cliente=excluded.nome_cliente,
                                        endereco=COALESCE(NULLIF(excluded.endereco,''), clientes.endereco)
                                """, (upper(cod_cliente), upper(nome_cliente), upper(endereco)))

                # =========================================================
                # âÅâ€œââ‚¬¦ REGRA NOVA: Vendas selecionadas viram "usadas" e somem
                # =========================================================
                try:
                    ids_carregados = [safe_int(x, 0) for x in (self._loaded_venda_ids or []) if safe_int(x, 0) > 0]
                    if ids_carregados:
                        qmarks = ",".join(["?"] * len(ids_carregados))
                        cur.execute(
                            f"""
                            UPDATE vendas_importadas
                            SET
                                usada = 1,
                                usada_em = ?,
                                codigo_programacao = ?,
                                selecionada = 0
                            WHERE id IN ({qmarks})
                              AND selecionada = 1
                              AND IFNULL(usada, 0) = 0
                            """,
                            tuple([data_criacao, codigo] + ids_carregados),
                        )
                except Exception:
                    logging.debug("Falha ao marcar vendas importadas como usadas.", exc_info=True)

                conn.commit()
                conn.close()
                conn = None

        except Exception as e:
            if conn is not None:
                try:
                    conn.rollback()
                    conn.close()
                except Exception:
                    logging.debug("Falha ao encerrar conexao de salvar_programacao", exc_info=True)
            messagebox.showerror("ERRO", f"Erro ao salvar programação: {str(e)}")
            return

        self.ent_codigo.config(state="normal")
        self.ent_codigo.delete(0, "end")
        self.ent_codigo.insert(0, codigo)
        self.ent_codigo.config(state="readonly")
        self._editing_programacao_codigo = codigo
        self._refresh_programacao_status_badge(codigo, status_local_hint="ATIVA")

        if is_update_existing:
            messagebox.showinfo("OK", f"Programação atualizada: {codigo}")
            self.set_status(f"STATUS: Programação atualizada: {codigo}")
        else:
            if creating_base_programacao:
                messagebox.showinfo(
                    "OK",
                    f"Programação-base salva: {codigo}.\n\n"
                    "O vínculo com o celular já pode ser feito.\n"
                    "Depois, importe as vendas, carregue as selecionadas e salve novamente para anexar os clientes."
                )
                self.set_status(
                    f"STATUS: Programação-base {codigo} salva sem clientes. "
                    "Importe as vendas e salve novamente para vincular os itens."
                )
            else:
                messagebox.showinfo("OK", f"Programação salva: {codigo} (ABERTA/ATIVA)")
                self.set_status(f"STATUS: Programação salva: {codigo} (ABERTA/ATIVA)")

        # Sincroniza programação com a API central (para aparecer no app mobile).
        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            if (not api_saved) and desktop_secret:
                itens_payload = []
                for (cod_cliente, nome_cliente, produto, endereco, caixas, kg, preco, vendedor, pedido, obs) in itens:
                    itens_payload.append({
                        "cod_cliente": upper(cod_cliente),
                        "nome_cliente": upper(nome_cliente),
                        "qnt_caixas": safe_int(caixas, 0),
                        "kg": safe_float(kg, 0.0),
                        "preco": safe_float(preco, 0.0),
                        "endereco": upper(endereco),
                        "vendedor": upper(vendedor),
                        "pedido": upper(pedido),
                        "produto": upper(produto),
                        "obs": upper(obs),
                    })

                payload_sync = {
                    "codigo_programacao": codigo,
                    "data_criacao": data_criacao,
                    "motorista": motorista_nome,
                    "motorista_id": safe_int(motorista_id, 0),
                    "motorista_codigo": motorista_codigo,
                    "codigo_motorista": motorista_codigo,
                    "veiculo": veiculo,
                    "equipe": equipe,
                    "kg_estimado": safe_float(kg_estimado, 0.0),
                    "tipo_estimativa": tipo_estimativa,
                    "caixas_estimado": safe_int(caixas_estimado, 0),
                    "status": "ATIVA",
                    "local_rota": local_rota,
                    "tipo_rota": local_rota,
                    "local_carregamento": local_carreg,
                    "local_carregado": local_carreg,
                    "granja_carregada": local_carreg,
                    "local_carreg": local_carreg,
                    "adiantamento": safe_float(adiantamento_val, 0.0),
                    "total_caixas": safe_int(total_caixas, 0),
                    "quilos": safe_float(total_quilos, 0.0),
                    "usuario_criacao": usuario_logado,
                    "usuario_ultima_edicao": usuario_logado,
                    "linked_venda_ids": [safe_int(x, 0) for x in (self._loaded_venda_ids or []) if safe_int(x, 0) > 0],
                    "vendas_usada_em": data_criacao,
                    "itens": itens_payload,
                }
                _call_api(
                    "POST",
                    "desktop/rotas/upsert",
                    payload=payload_sync,
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
        except Exception as exc:
            logging.warning("Falha ao sincronizar programacao %s com API: %s", codigo, exc)

        self.app.refresh_programacao_comboboxes()
        try:
            page_imp = self.app.pages.get("ImportarVendas") if hasattr(self.app, "pages") else None
            if page_imp and hasattr(page_imp, "carregar"):
                page_imp.carregar()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            if hasattr(self.app, "pages"):
                page_home = self.app.pages.get("Home")
                if page_home and hasattr(page_home, "on_show"):
                    page_home.on_show()
                page_rotas = self.app.pages.get("Rotas")
                if page_rotas and hasattr(page_rotas, "_load_data"):
                    page_rotas._load_data()
        except Exception:
            logging.debug("Falha ignorada")

        if creating_base_programacao:
            self._loaded_venda_ids = []
            return

        if messagebox.askyesno("PDF", "Deseja gerar o PDF da programação agora?\n\n(Pronto para impressão A4)"):
            self._abrir_previsualizacao_programacao_salva(
                codigo, motorista_nome, veiculo, equipe, kg_estimado, tipo_estimativa, caixas_estimado, usuario_logado
            )


        self._reset_form_after_save()

    def _preview_sheet_mode_label(self, mode_value: str) -> str:
        return {
            "fit_page": "Inteira",
            "fit_width": "Largura",
            "actual": "Real",
        }.get(str(mode_value or "").strip(), "Inteira")

    def _get_default_windows_printer(self) -> dict:
        if os.name != "nt":
            return {}
        try:
            winspool = ctypes.windll.winspool.drv
            needed = ctypes.c_uint(0)
            winspool.GetDefaultPrinterW(None, ctypes.byref(needed))
            if needed.value <= 0:
                return {}
            buf = ctypes.create_unicode_buffer(needed.value)
            if not winspool.GetDefaultPrinterW(buf, ctypes.byref(needed)):
                return {}
            name = str(buf.value or "").strip()
            if name:
                return {"name": name, "driver": "", "port": "", "default": True}
        except Exception:
            logging.debug("Falha ao obter impressora padrao do Windows.", exc_info=True)
        return {}

    def _list_windows_printers(self, force: bool = False):
        cached = tuple(getattr(self, "_windows_printers_cache", ()) or ())
        cached_at = safe_float(getattr(self, "_windows_printers_cache_at", 0.0), 0.0)
        if cached and not force and (time.time() - cached_at) <= 20.0:
            return [dict(item) for item in cached]

        printers = []
        if os.name == "nt":
            try:
                ps_exe = os.path.join(
                    os.environ.get("WINDIR", r"C:\Windows"),
                    "System32",
                    "WindowsPowerShell",
                    "v1.0",
                    "powershell.exe",
                )
                cmd = [
                    ps_exe if os.path.exists(ps_exe) else "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    "Get-CimInstance Win32_Printer | "
                    "Select-Object Name,DriverName,PortName,Default | "
                    "ConvertTo-Json -Compress",
                ]
                resp = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=12,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                raw = str(resp.stdout or "").strip()
                if resp.returncode == 0 and raw and raw.lower() != "null":
                    data = json.loads(raw)
                    if isinstance(data, dict):
                        data = [data]
                    for item in data or []:
                        if not isinstance(item, dict):
                            continue
                        name = str(item.get("Name") or item.get("name") or "").strip()
                        if not name:
                            continue
                        printers.append(
                            {
                                "name": name,
                                "driver": str(item.get("DriverName") or item.get("driver") or "").strip(),
                                "port": str(item.get("PortName") or item.get("port") or "").strip(),
                                "default": bool(item.get("Default") or item.get("default")),
                            }
                        )
                elif resp.returncode != 0:
                    logging.debug(
                        "Falha ao consultar impressoras via PowerShell: %s",
                        (resp.stderr or resp.stdout or "").strip(),
                    )
            except Exception:
                logging.debug("Falha ao listar impressoras instaladas.", exc_info=True)

        if not printers:
            default = self._get_default_windows_printer()
            if default:
                printers = [default]

        dedup = {}
        for item in printers:
            key = upper(item.get("name") or "")
            if key and key not in dedup:
                dedup[key] = item
        printers = sorted(dedup.values(), key=lambda item: (not bool(item.get("default")), upper(item.get("name") or "")))

        self._windows_printers_cache = tuple(dict(item) for item in printers)
        self._windows_printers_cache_at = time.time()
        return [dict(item) for item in printers]

    def _select_windows_printer(
        self,
        parent,
        document_label: str,
        sheet_mode: str = "",
        zoom_percent=None,
        extra_hint: str = "",
    ):
        printers = self._list_windows_printers()
        if not printers:
            messagebox.showerror(
                "Impressoras",
                "Nenhuma impressora instalada foi encontrada no Windows.\n\n"
                "Verifique o cadastro da impressora e tente novamente.",
                parent=parent or self.app,
            )
            return None

        win = tk.Toplevel(parent or self.app)
        win.title("Selecionar Impressora")
        win.geometry("640x360")
        win.resizable(False, False)
        win.transient(parent or self.app)
        win.grab_set()

        outer = ttk.Frame(win, style="Content.TFrame", padding=14)
        outer.pack(fill="both", expand=True)
        card = ttk.Frame(outer, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True)
        card.grid_columnconfigure(0, weight=1)

        ttk.Label(card, text="Imprimir", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            card,
            text=f"Documento: {document_label}",
            style="CardLabel.TLabel",
            wraplength=500,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 4))

        if sheet_mode or zoom_percent not in (None, ""):
            ajuste_txt = f"Ajuste de folha no preview: {sheet_mode or '-'}"
            if zoom_percent not in (None, ""):
                ajuste_txt += f" | Zoom: {safe_int(zoom_percent, 100)}%"
            ttk.Label(card, text=ajuste_txt, style="CardLabel.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 10))

        if extra_hint:
            ttk.Label(
                card,
                text=extra_hint,
                style="CardLabel.TLabel",
                wraplength=500,
                justify="left",
            ).grid(row=3, column=0, sticky="w", pady=(0, 10))
            combo_row = 4
        else:
            combo_row = 3

        ttk.Label(card, text="Impressora instalada", style="CardLabel.TLabel").grid(row=combo_row, column=0, sticky="w")
        cb = ttk.Combobox(card, state="readonly", width=68)
        cb.grid(row=combo_row + 1, column=0, sticky="ew", pady=(4, 0))

        labels = []
        label_to_printer = {}
        default_index = 0
        for idx, item in enumerate(printers):
            suffix = " (PADRAO)" if item.get("default") else ""
            label = f"{item.get('name') or '-'}{suffix}"
            labels.append(label)
            label_to_printer[label] = item
            if item.get("default"):
                default_index = idx
        cb["values"] = labels
        cb.current(default_index)

        detail_var = tk.StringVar()
        ttk.Label(card, textvariable=detail_var, style="CardLabel.TLabel", wraplength=500, justify="left").grid(
            row=combo_row + 2, column=0, sticky="w", pady=(10, 0)
        )

        result = {"printer": None}

        def _refresh_detail(_e=None):
            printer = label_to_printer.get(cb.get()) or {}
            tipo = "Padrao do Windows" if printer.get("default") else "Impressora instalada"
            detail_var.set(
                f"{tipo} | Driver: {printer.get('driver') or '-'} | Porta: {printer.get('port') or '-'}"
            )

        def _confirm():
            result["printer"] = label_to_printer.get(cb.get())
            win.destroy()

        def _cancel():
            result["printer"] = None
            win.destroy()

        cb.bind("<<ComboboxSelected>>", _refresh_detail)
        win.bind("<Return>", lambda _e: _confirm())
        win.bind("<Escape>", lambda _e: _cancel())
        win.protocol("WM_DELETE_WINDOW", _cancel)
        _refresh_detail()

        btns = ttk.Frame(card, style="Card.TFrame")
        btns.grid(row=combo_row + 3, column=0, sticky="e", pady=(16, 0))
        ttk.Button(btns, text="CANCELAR", style="Ghost.TButton", command=_cancel).pack(side="right")
        ttk.Button(btns, text="IMPRIMIR", style="Primary.TButton", command=_confirm).pack(side="right", padx=(0, 8))

        try:
            win.update_idletasks()
            screen_w = max(int(win.winfo_screenwidth() or 0), 900)
            screen_h = max(int(win.winfo_screenheight() or 0), 680)
            req_w = max(640, min(int(outer.winfo_reqwidth() + 44), screen_w - 40))
            req_h = max(360, min(int(outer.winfo_reqheight() + 44), screen_h - 60))

            owner = parent or self.app
            try:
                owner.update_idletasks()
                base_x = int(owner.winfo_rootx() or 0)
                base_y = int(owner.winfo_rooty() or 0)
                base_w = max(int(owner.winfo_width() or 0), req_w)
                base_h = max(int(owner.winfo_height() or 0), req_h)
                pos_x = base_x + max(int((base_w - req_w) / 2), 16)
                pos_y = base_y + max(int((base_h - req_h) / 2), 16)
            except Exception:
                pos_x = max(int((screen_w - req_w) / 2), 16)
                pos_y = max(int((screen_h - req_h) / 2), 16)

            pos_x = max(8, min(pos_x, max(screen_w - req_w - 8, 8)))
            pos_y = max(8, min(pos_y, max(screen_h - req_h - 32, 8)))
            win.geometry(f"{req_w}x{req_h}+{pos_x}+{pos_y}")
        except Exception:
            logging.debug("Falha ao ajustar dimensoes da janela de impressora.", exc_info=True)

        win.wait_window()
        return result.get("printer")

    def _shell_print_error_text(self, code: int) -> str:
        return {
            0: "Memoria insuficiente.",
            2: "Arquivo nao encontrado.",
            3: "Caminho nao encontrado.",
            5: "Acesso negado.",
            8: "Memoria insuficiente.",
            26: "Nao foi possivel compartilhar o arquivo para impressao.",
            27: "Associacao de impressao incompleta no Windows.",
            28: "Tempo esgotado ao iniciar o app de impressao.",
            29: "Falha ao executar o app associado ao PDF.",
            30: "A impressora esta ocupada ou sem resposta.",
            31: "Nenhum aplicativo associado ao PDF com suporte de impressao.",
        }.get(int(code or 0), f"Codigo do Windows: {int(code or 0)}")

    def _quote_windows_print_arg(self, value: str) -> str:
        return '"' + str(value or "").replace('"', '""') + '"'

    def _send_file_to_windows_printer(self, file_path: str, printer: dict):
        if os.name != "nt":
            raise RuntimeError("Impressao direta disponivel apenas no Windows.")

        file_path = os.path.abspath(file_path)
        printer_name = str((printer or {}).get("name") or "").strip()
        driver_name = str((printer or {}).get("driver") or "").strip()
        port_name = str((printer or {}).get("port") or "").strip()

        shell32 = ctypes.windll.shell32
        result = 0

        if printer_name:
            params = " ".join(
                self._quote_windows_print_arg(part)
                for part in (printer_name, driver_name, port_name)
                if str(part or "").strip()
            )
            if params:
                result = int(shell32.ShellExecuteW(None, "printto", file_path, params, None, 0))
            if result <= 32:
                result = int(
                    shell32.ShellExecuteW(
                        None,
                        "printto",
                        file_path,
                        self._quote_windows_print_arg(printer_name),
                        None,
                        0,
                    )
                )

        if result <= 32 and (not printer_name or printer.get("default")):
            result = int(shell32.ShellExecuteW(None, "print", file_path, None, None, 0))

        if result <= 32:
            raise RuntimeError(self._shell_print_error_text(result))

    def _build_temp_pdf_path(self, prefix: str, codigo: str) -> str:
        safe_code = re.sub(r"[^A-Z0-9_-]+", "_", upper(str(codigo or "").strip()) or "DOC")
        file_name = f"{prefix}_{safe_code}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.pdf"
        return os.path.join(tempfile.gettempdir(), file_name)

    def _gerar_pdf_programacao_salva_em_path(
        self,
        path,
        codigo,
        motorista,
        veiculo,
        equipe,
        kg_estimado,
        tipo_estimativa="KG",
        caixas_estimado=0,
        usuario_criacao="",
        itens_override=None,
        usuario_edicao_override="",
        data_emissao_override="",
    ):
        itens = list(itens_override or [])
        if not itens:
            for iid in self.tree.get_children():
                itens.append(self._get_row_values(iid))

        if not itens:
            raise ValueError("Sem itens na programacao.")

        usuario_edicao = upper(str(usuario_edicao_override or "").strip())
        try:
            meta_prog = self._buscar_meta_programacao(codigo)
            usuario_criacao = upper(str(meta_prog.get("usuario_criacao") or usuario_criacao or "").strip())
            usuario_edicao = upper(str(meta_prog.get("usuario_ultima_edicao") or usuario_edicao or "").strip())
        except Exception:
            usuario_criacao = upper((usuario_criacao or "").strip())
            usuario_edicao = upper((usuario_edicao or "").strip())

        c = canvas.Canvas(path, pagesize=A4)
        w, h = A4
        to_txt = lambda v: fix_mojibake_text(str(v or ""))

        y = h - 60
        c.setFont("Helvetica-Bold", 14)
        c.drawString(40, y, f"PROGRAMACAO: {to_txt(codigo)}")
        y -= 22

        c.setFont("Helvetica", 10)
        data_emissao = str(data_emissao_override or datetime.now().strftime('%d/%m/%Y %H:%M'))
        c.drawString(40, y, f"Data: {data_emissao}")
        y -= 16
        equipe_txt = self._resolve_equipe_ajudantes(equipe)
        c.drawString(40, y, f"Motorista: {to_txt(motorista)}  |  Veiculo: {to_txt(veiculo)}  |  Equipe: {to_txt(equipe_txt)}")
        y -= 16
        if upper(tipo_estimativa) == "CX":
            c.drawString(40, y, f"Estimado (FOB): {safe_int(caixas_estimado, 0)} CX")
        else:
            c.drawString(40, y, f"Estimado (CIF): {safe_float(kg_estimado, 0.0):.2f} KG")
        y -= 16
        c.drawString(
            40,
            y,
            f"Criado por: {to_txt(usuario_criacao or '-')}  |  Ultima edicao: {to_txt(usuario_edicao or '-')}",
        )
        y -= 24

        c.setFont("Helvetica-Bold", 9)
        c.drawString(40, y, "CLIENTE / ENDERECO")
        c.drawString(320, y, "CX")
        c.drawString(370, y, "PRECO")
        c.drawString(430, y, "VENDEDOR")
        c.drawString(520, y, "PEDIDO")
        y -= 10
        c.line(40, y, w - 40, y)
        y -= 14

        c.setFont("Helvetica", 8)

        def _extrair_cidade(endereco_raw):
            txt = upper(str(endereco_raw or "").strip())
            if not txt:
                return ""
            for sep in (" - ", ",", "/", ";"):
                if sep in txt:
                    partes = [p.strip() for p in txt.split(sep) if p.strip()]
                    if partes:
                        return partes[-1]
            return txt

        for (cod_cliente, nome_cliente, produto, endereco, caixas, kg, preco, vendedor, pedido, obs) in itens:
            if y < 95:
                c.showPage()
                y = h - 60
                c.setFont("Helvetica", 8)

            cidade = _extrair_cidade(endereco)
            linha_cliente = f"{cidade} - {cod_cliente} - {nome_cliente}" if cidade else f"{cod_cliente} - {nome_cliente}"
            if len(linha_cliente) > 78:
                linha_cliente = linha_cliente[:78] + "..."

            c.drawString(40, y, to_txt(linha_cliente))
            c.drawRightString(340, y, str(caixas))
            c.drawRightString(410, y, str(preco))
            c.drawString(430, y, to_txt(vendedor)[:12])
            c.drawString(520, y, to_txt(pedido)[:18])
            y -= 12

            if obs:
                obs_line = f"OBS: {to_txt(obs)}"
                if len(obs_line) > 110:
                    obs_line = obs_line[:110] + "..."
            else:
                obs_line = "OBS: ________________________________________________"
            c.setFont("Helvetica-Oblique", 8)
            c.drawString(50, y, obs_line)
            c.setFont("Helvetica", 8)
            y -= 10
            y -= 4

        if y < 40:
            c.showPage()
            y = h - 60

        c.setFont("Helvetica-Oblique", 9)
        c.drawCentredString(w / 2, 26, '"Tudo posso naquele que me fortalece." (Filipenses 4:13)')
        c.save()
        return path

    def _imprimir_programacao_salva(
        self,
        codigo,
        motorista,
        veiculo,
        equipe,
        kg_estimado,
        tipo_estimativa="KG",
        caixas_estimado=0,
        usuario_criacao="",
        itens_override=None,
        usuario_edicao_override="",
        data_emissao_override="",
        parent=None,
        preview_mode="fit_page",
        preview_zoom=100,
    ):
        if not require_reportlab():
            return None

        printer = self._select_windows_printer(
            parent=parent or self.app,
            document_label=f"Folha da Programacao {codigo}",
            sheet_mode=self._preview_sheet_mode_label(preview_mode),
            zoom_percent=preview_zoom,
            extra_hint=(
                "O preview continua com os ajustes de folha/zoom. "
                "A regulagem final de papel depende do driver da impressora selecionada."
            ),
        )
        if not printer:
            return None

        path = self._build_temp_pdf_path("PROGRAMACAO", codigo)
        try:
            self._gerar_pdf_programacao_salva_em_path(
                path,
                codigo,
                motorista,
                veiculo,
                equipe,
                kg_estimado,
                tipo_estimativa,
                caixas_estimado,
                usuario_criacao,
                itens_override=itens_override,
                usuario_edicao_override=usuario_edicao_override,
                data_emissao_override=data_emissao_override,
            )
            self._send_file_to_windows_printer(path, printer)
            messagebox.showinfo(
                "Impressao enviada",
                f"Folha da programacao {codigo} enviada para:\n{printer.get('name') or 'IMPRESSORA PADRAO'}",
                parent=parent or self.app,
            )
            return path
        except Exception as exc:
            messagebox.showerror(
                "ERRO",
                "Nao foi possivel enviar a folha da programacao para a impressora.\n\n"
                f"Detalhe: {exc}",
                parent=parent or self.app,
            )
            return None

    def gerar_pdf_programacao_salva(
        self,
        codigo,
        motorista,
        veiculo,
        equipe,
        kg_estimado,
        tipo_estimativa="KG",
        caixas_estimado=0,
        usuario_criacao="",
        perguntar_romaneios=False,
        itens_override=None,
        usuario_edicao_override="",
        data_emissao_override="",
    ):
        if not require_reportlab():
            return None
        path = filedialog.asksaveasfilename(
            title="Salvar PDF da Programação",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"PROGRAMACAO_{codigo}.pdf"
        )
        if not path:
            return None

        try:
            self._gerar_pdf_programacao_salva_em_path(
                path,
                codigo,
                motorista,
                veiculo,
                equipe,
                kg_estimado,
                tipo_estimativa,
                caixas_estimado,
                usuario_criacao,
                itens_override=itens_override,
                usuario_edicao_override=usuario_edicao_override,
                data_emissao_override=data_emissao_override,
            )
            messagebox.showinfo("OK", "PDF gerado com sucesso! (A4 pronto para impressão)")
            if perguntar_romaneios and messagebox.askyesno("Romaneios", "Deseja gerar os romaneios de entrega desta programacao agora?"):
                self.imprimir_romaneios_programacao(codigo_override=codigo)
            return path

        except Exception as e:
            messagebox.showerror("ERRO", str(e))
            return None

    def _build_preview_programacao_salva_text(
        self, codigo, motorista, veiculo, equipe, kg_estimado, tipo_estimativa="KG", caixas_estimado=0, usuario_criacao=""
    ) -> str:
        itens = [self._get_row_values(iid) for iid in self.tree.get_children()]
        equipe_txt = self._resolve_equipe_ajudantes(equipe)
        linhas = [
            f"FOLHA DE PROGRAMACAO - {upper(codigo)}",
            "=" * 110,
            "",
            f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            f"Motorista: {fix_mojibake_text(str(motorista or ''))}",
            f"Veiculo: {fix_mojibake_text(str(veiculo or ''))}",
            f"Equipe: {fix_mojibake_text(str(equipe_txt or '-'))}",
        ]
        if upper(tipo_estimativa) == "CX":
            linhas.append(f"Estimado (FOB): {safe_int(caixas_estimado, 0)} CX")
        else:
            linhas.append(f"Estimado (CIF): {safe_float(kg_estimado, 0.0):.2f} KG")
        linhas.append(f"Criado por: {fix_mojibake_text(str(usuario_criacao or '-'))}")
        linhas.extend(["", "[CLIENTES]", "CLIENTE / ENDERECO | CX | PRECO | VENDEDOR | PEDIDO | OBS", "-" * 110])

        for cod_cliente, nome_cliente, _produto, endereco, caixas, _kg, preco, vendedor, pedido, obs in itens:
            cidade = self._extrair_cidade_do_endereco(endereco)
            linha_cliente = f"{cidade} - {cod_cliente} - {nome_cliente}" if cidade else f"{cod_cliente} - {nome_cliente}"
            linhas.append(
                f"{fix_mojibake_text(str(linha_cliente)[:58])} | {caixas} | {preco} | {fix_mojibake_text(str(vendedor)[:12])} | {fix_mojibake_text(str(pedido)[:18])} | {fix_mojibake_text(str(obs or '')[:24])}"
            )

        if not itens:
            linhas.append("Sem itens na programacao.")

        return "\n".join(linhas)

    def _collect_programacao_preview_payload(
        self, codigo, motorista, veiculo, equipe, kg_estimado, tipo_estimativa="KG", caixas_estimado=0, usuario_criacao=""
    ):
        itens = []
        for iid in self.tree.get_children():
            cod_cliente, nome_cliente, _produto, endereco, caixas, _kg, preco, vendedor, pedido, obs = self._get_row_values(iid)
            itens.append(
                {
                    "cod_cliente": upper(cod_cliente),
                    "nome_cliente": upper(nome_cliente),
                    "endereco": upper(endereco),
                    "qnt_caixas": safe_int(caixas, 0),
                    "preco": self._normalizar_preco_item(preco),
                    "vendedor": upper(vendedor),
                    "pedido": upper(pedido),
                    "obs": str(obs or "").strip(),
                }
            )

        meta = {
            "codigo": upper(codigo or "-"),
            "motorista": upper(motorista or "-"),
            "veiculo": upper(veiculo or "-"),
            "equipe": self._resolve_equipe_ajudantes(equipe),
            "usuario_criacao": upper(usuario_criacao or "-"),
            "usuario_ultima_edicao": upper(usuario_criacao or "-"),
            "tipo_estimativa": upper(tipo_estimativa or "KG"),
            "kg_estimado": safe_float(kg_estimado, 0.0),
            "caixas_estimado": safe_int(caixas_estimado, 0),
            "data_emissao": datetime.now().strftime('%d/%m/%Y %H:%M'),
        }
        try:
            meta_prog = self._buscar_meta_programacao(codigo)
            if isinstance(meta_prog, dict):
                meta["usuario_criacao"] = upper(str(meta_prog.get("usuario_criacao") or meta["usuario_criacao"] or "-").strip())
                meta["usuario_ultima_edicao"] = upper(str(meta_prog.get("usuario_ultima_edicao") or meta["usuario_ultima_edicao"] or "-").strip())
        except Exception:
            logging.debug("Falha ao buscar metadados da programacao para preview local.", exc_info=True)

        return meta, itens

    def _create_programacao_canvas_preview_tab_local(self, notebook, payload, zoom_var=None, mode_var=None):
        tab = ttk.Frame(notebook)
        notebook.add(tab, text="Folha Programacao")
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        container = ttk.Frame(tab, style="Content.TFrame")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(container, bg="#ECEFF4", highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        meta = payload.get("meta") or {}
        itens = payload.get("itens") or []

        def _fit_font(base_size, scale_factor):
            return max(8, int(round(base_size * max(0.8, scale_factor))))

        def _paginate_preview_rows():
            pages = []
            rows = []
            y_pdf = 842.0 - 60.0
            y_pdf -= 22
            y_pdf -= 16
            y_pdf -= 16
            y_pdf -= 16
            y_pdf -= 24
            y_pdf -= 10
            y_pdf -= 14

            for item in itens:
                if y_pdf < 95:
                    pages.append(rows)
                    rows = []
                    y_pdf = 842.0 - 60.0
                cidade = self._extrair_cidade_do_endereco(item.get("endereco"))
                linha_cliente = (
                    f"{cidade} - {item.get('cod_cliente')} - {item.get('nome_cliente')}"
                    if cidade
                    else f"{item.get('cod_cliente')} - {item.get('nome_cliente')}"
                )
                if len(linha_cliente) > 78:
                    linha_cliente = linha_cliente[:78] + "..."
                obs = str(item.get("obs") or item.get("observacao") or "").strip()
                obs_line = f"OBS: {fix_mojibake_text(obs)}" if obs else "OBS: ________________________________________________"
                if len(obs_line) > 110:
                    obs_line = obs_line[:110] + "..."
                rows.append(
                    {
                        "cliente": fix_mojibake_text(linha_cliente),
                        "caixas": str(item.get("qnt_caixas") or 0),
                        "preco": f"{safe_float(item.get('preco'), 0.0):.2f}",
                        "vendedor": fix_mojibake_text(str(item.get("vendedor") or "")[:12]),
                        "pedido": fix_mojibake_text(str(item.get("pedido") or "")[:18]),
                        "obs": obs_line,
                    }
                )
                y_pdf -= 12
                y_pdf -= 10
                y_pdf -= 4
            pages.append(rows)
            return pages

        def _render(_e=None):
            canvas.delete("all")
            cw = max(canvas.winfo_width(), 980)
            ch = max(canvas.winfo_height(), 720)
            pad = 24
            ratio = 210.0 / 297.0
            zoom = 1.0
            try:
                if zoom_var is not None:
                    zoom = max(0.5, min(2.0, float(zoom_var.get()) / 100.0))
            except Exception:
                zoom = 1.0
            mode = "fit_page"
            try:
                if mode_var is not None:
                    mode = str(mode_var.get() or "fit_page").strip()
            except Exception:
                mode = "fit_page"

            avail_w = max(400, cw - (2 * pad))
            avail_h = max(500, ch - (2 * pad))
            if mode == "fit_width":
                page_w = avail_w
                page_h = page_w / ratio
            elif mode == "actual":
                page_w = 595.0 * zoom
                page_h = 842.0 * zoom
            else:
                scale = min(avail_w / 595.0, avail_h / 842.0)
                page_w = 595.0 * scale * zoom
                page_h = 842.0 * scale * zoom

            page_gap = max(26, int(page_h * 0.04))
            pages = _paginate_preview_rows()
            total_h = 20 + (len(pages) * page_h) + (max(0, len(pages) - 1) * page_gap) + 24
            canvas.configure(scrollregion=(0, 0, max(cw, page_w + 80), max(ch, total_h)))

            for page_index, page_rows in enumerate(pages):
                x0 = max(20, (cw - page_w) / 2)
                y0 = 20 + (page_index * (page_h + page_gap))
                x1 = x0 + page_w
                y1 = y0 + page_h
                canvas.create_rectangle(x0, y0, x1, y1, fill="white", outline="#C7CDD4", width=1)

                def ppdf(x_pdf):
                    return x0 + (x_pdf * (page_w / 595.0))

                def qpdf(y_pdf):
                    return y0 + (y_pdf * (page_h / 842.0))

                scale_factor = page_w / 595.0
                title_font = ("Segoe UI", _fit_font(14, scale_factor), "bold")
                body_font = ("Segoe UI", _fit_font(10, scale_factor))
                header_font = ("Segoe UI", _fit_font(9, scale_factor), "bold")
                item_font = ("Consolas", _fit_font(8, scale_factor))
                obs_font = ("Segoe UI", _fit_font(8, scale_factor), "italic")
                body_width = max(120, ppdf(555) - ppdf(40))

                y_pdf = 60
                canvas.create_text(
                    ppdf(40),
                    qpdf(y_pdf),
                    text=f"PROGRAMACAO: {fix_mojibake_text(str(meta.get('codigo') or '-'))}",
                    anchor="w",
                    font=title_font,
                )
                y_pdf += 22
                canvas.create_text(
                    ppdf(40),
                    qpdf(y_pdf),
                    text=f"Data: {meta.get('data_emissao') or datetime.now().strftime('%d/%m/%Y %H:%M')}",
                    anchor="w",
                    font=body_font,
                    width=body_width,
                )
                y_pdf += 16
                canvas.create_text(
                    ppdf(40),
                    qpdf(y_pdf),
                    text=f"Motorista: {fix_mojibake_text(str(meta.get('motorista') or '-'))}  |  Veiculo: {fix_mojibake_text(str(meta.get('veiculo') or '-'))}",
                    anchor="w",
                    font=body_font,
                    width=body_width,
                )
                y_pdf += 16
                canvas.create_text(
                    ppdf(40),
                    qpdf(y_pdf),
                    text=f"Equipe: {fix_mojibake_text(str(meta.get('equipe') or '-'))}",
                    anchor="w",
                    font=body_font,
                    width=body_width,
                )
                y_pdf += 16
                estimado_txt = (
                    f"Estimado (FOB): {safe_int(meta.get('caixas_estimado'), 0)} CX"
                    if upper(meta.get("tipo_estimativa") or "KG") == "CX"
                    else f"Estimado (CIF): {safe_float(meta.get('kg_estimado'), 0.0):.2f} KG"
                )
                canvas.create_text(ppdf(40), qpdf(y_pdf), text=estimado_txt, anchor="w", font=body_font, width=body_width)
                y_pdf += 16
                canvas.create_text(
                    ppdf(40),
                    qpdf(y_pdf),
                    text=f"Criado por: {fix_mojibake_text(str(meta.get('usuario_criacao') or '-'))}  |  Ultima edicao: {fix_mojibake_text(str(meta.get('usuario_ultima_edicao') or '-'))}",
                    anchor="w",
                    font=body_font,
                    width=body_width,
                )
                y_pdf += 22

                canvas.create_text(ppdf(40), qpdf(y_pdf), text="CLIENTE / ENDERECO", anchor="w", font=header_font)
                canvas.create_text(ppdf(320), qpdf(y_pdf), text="CX", anchor="w", font=header_font)
                canvas.create_text(ppdf(370), qpdf(y_pdf), text="PRECO", anchor="w", font=header_font)
                canvas.create_text(ppdf(430), qpdf(y_pdf), text="VENDEDOR", anchor="w", font=header_font)
                canvas.create_text(ppdf(520), qpdf(y_pdf), text="PEDIDO", anchor="w", font=header_font)
                y_pdf += 10
                canvas.create_line(ppdf(40), qpdf(y_pdf), ppdf(555), qpdf(y_pdf), fill="#222")
                y_pdf += 14

                for row in page_rows:
                    cliente_txt = row["cliente"][:52]
                    vendedor_txt = row["vendedor"][:11]
                    pedido_txt = row["pedido"][:10]
                    canvas.create_text(ppdf(40), qpdf(y_pdf), text=cliente_txt, anchor="w", font=item_font)
                    canvas.create_text(ppdf(340), qpdf(y_pdf), text=row["caixas"], anchor="e", font=item_font)
                    canvas.create_text(ppdf(410), qpdf(y_pdf), text=row["preco"], anchor="e", font=item_font)
                    canvas.create_text(ppdf(430), qpdf(y_pdf), text=vendedor_txt, anchor="w", font=item_font)
                    canvas.create_text(ppdf(520), qpdf(y_pdf), text=pedido_txt, anchor="w", font=item_font)
                    y_pdf += 12
                    canvas.create_text(
                        ppdf(50),
                        qpdf(y_pdf),
                        text=row["obs"],
                        anchor="w",
                        font=obs_font,
                        width=max(100, ppdf(285) - ppdf(50)),
                    )
                    y_pdf += 14

                canvas.create_text(
                    ppdf(297.5),
                    qpdf(816),
                    text='"Tudo posso naquele que me fortalece." (Filipenses 4:13)',
                    anchor="center",
                    font=("Segoe UI", _fit_font(9, scale_factor), "italic"),
                )

        canvas.bind("<Configure>", _render)
        if zoom_var is not None:
            zoom_var.trace_add("write", lambda *_: _render())
        if mode_var is not None:
            mode_var.trace_add("write", lambda *_: _render())
        _render()

    def _abrir_previsualizacao_programacao_salva(
        self, codigo, motorista, veiculo, equipe, kg_estimado, tipo_estimativa="KG", caixas_estimado=0, usuario_criacao=""
    ):
        itens = [self._get_row_values(iid) for iid in self.tree.get_children()]
        if not itens:
            messagebox.showwarning("ATENCAO", "Sem itens na programacao.")
            return

        top = tk.Toplevel(self.app)
        top.title(f"Pre-visualizacao da Programacao - {codigo}")
        top.geometry("1180x760")
        top.minsize(900, 600)
        top.transient(self.app)
        top.grab_set()
        toolbar = ttk.Frame(top, padding=(8, 8, 8, 0))
        toolbar.pack(fill="x")
        nb = ttk.Notebook(top)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        footer = ttk.Frame(top, padding=(8, 0, 8, 8))
        footer.pack(fill="x")
        preview_zoom = tk.IntVar(value=100)
        preview_mode = tk.StringVar(value="fit_page")

        def _abrir_mesma_preview():
            try:
                top.destroy()
            except Exception:
                logging.debug("Falha ignorada")
            self._abrir_previsualizacao_programacao_salva(
                codigo, motorista, veiculo, equipe, kg_estimado, tipo_estimativa, caixas_estimado, usuario_criacao
            )

        meta_preview, itens_preview = self._collect_programacao_preview_payload(
            codigo, motorista, veiculo, equipe, kg_estimado, tipo_estimativa, caixas_estimado, usuario_criacao
        )

        def _gerar_pdf():
            return self.gerar_pdf_programacao_salva(
                codigo,
                motorista,
                veiculo,
                equipe,
                kg_estimado,
                tipo_estimativa,
                caixas_estimado,
                usuario_criacao,
                perguntar_romaneios=True,
                itens_override=[
                    (
                        item.get("cod_cliente"),
                        item.get("nome_cliente"),
                        "",
                        item.get("endereco"),
                        item.get("qnt_caixas"),
                        0,
                        item.get("preco"),
                        item.get("vendedor"),
                        item.get("pedido"),
                        item.get("obs") or item.get("observacao"),
                    )
                    for item in itens_preview
                ],
                usuario_edicao_override=meta_preview.get("usuario_ultima_edicao") or "",
                data_emissao_override=meta_preview.get("data_emissao") or "",
            )

        def _imprimir_programacao():
            return self._imprimir_programacao_salva(
                codigo,
                motorista,
                veiculo,
                equipe,
                kg_estimado,
                tipo_estimativa,
                caixas_estimado,
                usuario_criacao,
                itens_override=[
                    (
                        item.get("cod_cliente"),
                        item.get("nome_cliente"),
                        "",
                        item.get("endereco"),
                        item.get("qnt_caixas"),
                        0,
                        item.get("preco"),
                        item.get("vendedor"),
                        item.get("pedido"),
                        item.get("obs") or item.get("observacao"),
                    )
                    for item in itens_preview
                ],
                usuario_edicao_override=meta_preview.get("usuario_ultima_edicao") or "",
                data_emissao_override=meta_preview.get("data_emissao") or "",
                parent=top,
                preview_mode=preview_mode.get(),
                preview_zoom=preview_zoom.get(),
            )

        ttk.Button(toolbar, text="Atualizar", style="Ghost.TButton", command=_abrir_mesma_preview).pack(side="left")
        ttk.Button(toolbar, text="\U0001F5A8 Imprimir", style="Primary.TButton", command=_imprimir_programacao).pack(side="left", padx=(8, 0))
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Label(toolbar, text="Folha:", style="CardLabel.TLabel").pack(side="left", padx=(0, 4))
        ttk.Radiobutton(toolbar, text="Inteira", value="fit_page", variable=preview_mode).pack(side="left")
        ttk.Radiobutton(toolbar, text="Largura", value="fit_width", variable=preview_mode).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(toolbar, text="Real", value="actual", variable=preview_mode).pack(side="left", padx=(6, 0))
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Label(toolbar, text="Zoom:", style="CardLabel.TLabel").pack(side="left", padx=(12, 4))

        def _set_zoom(v):
            try:
                novo = max(50, min(200, int(v)))
            except Exception:
                novo = 100
            preview_zoom.set(novo)
            try:
                cb_zoom.set(str(novo))
            except Exception:
                logging.debug("Falha ignorada")

        ttk.Button(toolbar, text="-", width=3, command=lambda: _set_zoom(preview_zoom.get() - 10)).pack(side="left")
        cb_zoom = ttk.Combobox(toolbar, state="readonly", width=6, values=["75", "90", "100", "110", "125", "150"])
        cb_zoom.pack(side="left", padx=4)
        cb_zoom.set("100")
        cb_zoom.bind("<<ComboboxSelected>>", lambda _e: _set_zoom(cb_zoom.get()))
        ttk.Button(toolbar, text="+", width=3, command=lambda: _set_zoom(preview_zoom.get() + 10)).pack(side="left")
        ttk.Label(toolbar, text="%", style="CardLabel.TLabel").pack(side="left", padx=(2, 0))
        ttk.Button(toolbar, text="Fechar", style="Ghost.TButton", command=top.destroy).pack(side="right")

        self._create_programacao_canvas_preview_tab_local(
            nb,
            {"meta": meta_preview, "itens": itens_preview},
            preview_zoom,
            preview_mode,
        )

        tab_resumo = ttk.Frame(nb)
        nb.add(tab_resumo, text="Resumo")
        text = tk.Text(tab_resumo, wrap="word")
        text.pack(fill="both", expand=True)
        text.insert(
            "1.0",
            self._build_preview_programacao_salva_text(
                codigo, motorista, veiculo, equipe, kg_estimado, tipo_estimativa, caixas_estimado, usuario_criacao
            ),
        )
        text.configure(state="disabled")

        ttk.Button(footer, text="Fechar", style="Ghost.TButton", command=top.destroy).pack(side="right")
        ttk.Button(footer, text="GERAR PDF", style="Primary.TButton", command=_gerar_pdf).pack(side="right", padx=(0, 8))
        ttk.Button(
            footer,
            text="IMPRIMIR ROMANEIOS",
            style="Ghost.TButton",
            command=lambda: self.imprimir_romaneios_programacao(codigo_override=codigo),
        ).pack(side="right", padx=(0, 8))


    def _normalizar_preco_item(self, valor):
        p = safe_float(valor, 0.0)
        if abs(p) >= 100 and abs(p - round(p)) < 1e-9:
            p = p / 100.0
        return p

    def _extrair_cidade_do_endereco(self, endereco_raw: str) -> str:
        txt = upper(str(endereco_raw or "").strip())
        if not txt:
            return ""
        for sep in (" - ", ",", "/", ";"):
            if sep in txt:
                partes = [p.strip() for p in txt.split(sep) if p.strip()]
                if partes:
                    return partes[-1]
        return txt

    def _buscar_meta_programacao(self, codigo: str) -> dict:
        meta = {
            "motorista": "",
            "veiculo": "",
            "equipe": "",
            "data_criacao": "",
            "kg_estimado": 0.0,
            "tipo_estimativa": "KG",
            "caixas_estimado": 0,
            "aves_por_caixa": 6,
            "local_rota": "",
            "local_carregamento": "",
            "usuario_criacao": "",
            "usuario_ultima_edicao": "",
        }
        try:
            bundle = self._api_bundle_relatorio(upper(codigo or "").strip()) if hasattr(self, "_api_bundle_relatorio") else None
            rota = bundle.get("rota") if isinstance(bundle, dict) else None
            if isinstance(rota, dict):
                    meta["motorista"] = upper(str(rota.get("motorista") or ""))
                    meta["veiculo"] = upper(str(rota.get("veiculo") or ""))
                    meta["equipe"] = upper(str(rota.get("equipe") or ""))
                    meta["data_criacao"] = str(rota.get("data_criacao") or rota.get("data") or "")
                    meta["kg_estimado"] = safe_float(rota.get("kg_estimado"), 0.0)
                    meta["tipo_estimativa"] = upper(str(rota.get("tipo_estimativa") or "KG").strip()) or "KG"
                    meta["caixas_estimado"] = safe_int(rota.get("caixas_estimado"), 0)
                    meta["usuario_criacao"] = upper(
                        str(
                            rota.get("usuario_criacao")
                            or rota.get("usuario")
                            or rota.get("criado_por")
                            or rota.get("created_by")
                            or rota.get("autor")
                            or ""
                        ).strip()
                    )
                    meta["usuario_ultima_edicao"] = upper(str(rota.get("usuario_ultima_edicao") or "").strip())
                    meta["local_rota"] = upper(str(rota.get("local_rota") or rota.get("tipo_rota") or ""))
                    meta["local_carregamento"] = upper(
                        str(
                            rota.get("local_carregamento")
                            or rota.get("granja_carregada")
                            or rota.get("local_carregado")
                            or rota.get("local_carreg")
                            or ""
                        )
                    )
                    for cand in ("qnt_aves_por_cx", "aves_por_caixa", "qnt_aves_por_caixa"):
                        apc = safe_int(rota.get(cand), 0)
                        if apc > 0:
                            meta["aves_por_caixa"] = apc
                            break
                    return meta
        except Exception:
            logging.debug("Falha ao buscar metadados da programacao via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                data_expr = "COALESCE(data_criacao,'')" if "data_criacao" in cols_prog else ("COALESCE(data,'')" if "data" in cols_prog else "''")
                kg_expr = "COALESCE(kg_estimado,0)" if "kg_estimado" in cols_prog else "0"
                tipo_estim_expr = "COALESCE(tipo_estimativa,'KG')" if "tipo_estimativa" in cols_prog else "'KG'"
                caixas_estim_expr = "COALESCE(caixas_estimado,0)" if "caixas_estimado" in cols_prog else "0"
                usuario_criacao_expr = "COALESCE(usuario_criacao,'')" if "usuario_criacao" in cols_prog else "''"
                usuario_edicao_expr = "COALESCE(usuario_ultima_edicao,'')" if "usuario_ultima_edicao" in cols_prog else "''"
                cur.execute(
                    f"""
                    SELECT COALESCE(motorista,''), COALESCE(veiculo,''), COALESCE(equipe,''),
                           {data_expr}, {kg_expr},
                           {tipo_estim_expr}, {caixas_estim_expr},
                           {usuario_criacao_expr}, {usuario_edicao_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                    """,
                    (codigo,),
                )
                row = cur.fetchone()
                if row:
                    meta["motorista"] = upper(row[0])
                    meta["veiculo"] = upper(row[1])
                    meta["equipe"] = upper(row[2])
                    meta["data_criacao"] = row[3] or ""
                    meta["kg_estimado"] = safe_float(row[4], 0.0)
                    meta["tipo_estimativa"] = upper((row[5] or "KG").strip()) or "KG"
                    meta["caixas_estimado"] = safe_int(row[6], 0)
                    meta["usuario_criacao"] = upper((row[7] or "").strip())
                    meta["usuario_ultima_edicao"] = upper((row[8] or "").strip())

                for cand in ("qnt_aves_por_cx", "aves_por_caixa", "qnt_aves_por_caixa"):
                    if db_has_column(cur, "programacoes", cand):
                        cur.execute(f"SELECT COALESCE({cand},0) FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1", (codigo,))
                        r = cur.fetchone()
                        apc = safe_int((r[0] if r else 0), 0)
                        if apc > 0:
                            meta["aves_por_caixa"] = apc
                            break

                for cand in ("local_rota", "tipo_rota"):
                    if db_has_column(cur, "programacoes", cand):
                        cur.execute(f"SELECT COALESCE({cand},'') FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1", (codigo,))
                        r = cur.fetchone()
                        v = upper((r[0] if r else "") or "")
                        if v:
                            meta["local_rota"] = v
                            break

                for cand in ("local_carregamento", "granja_carregada", "local_carregado", "local_carreg"):
                    if db_has_column(cur, "programacoes", cand):
                        cur.execute(f"SELECT COALESCE({cand},'') FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1", (codigo,))
                        r = cur.fetchone()
                        v = upper((r[0] if r else "") or "")
                        if v:
                            meta["local_carregamento"] = v
                            break
        except Exception:
            logging.debug("Falha ignorada")
        return meta

    def _desenhar_bloco_romaneio(self, c, x, y_top, largura, altura, codigo_prog: str, meta: dict, item: dict, via_label: str):
        from reportlab.lib.units import mm

        y_base = y_top - altura
        c.setLineWidth(0.55)
        c.rect(x, y_base, largura, altura)

        pad = 2.4 * mm
        y = y_top - pad

        def money(v):
            return f"{safe_float(v, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        def fmt_pedido(v):
            s = str(v or "").strip()
            if not s:
                return ""
            try:
                n = float(s.replace(",", "."))
                if abs(n - int(n)) < 1e-9:
                    return str(int(n))
            except Exception:
                pass
            return s

        # Header
        logo_w = 28 * mm
        logo_h = 14.5 * mm
        c.rect(x + pad, y - logo_h, logo_w, logo_h)
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(x + pad + logo_w / 2, y - 8 * mm, "LOGO")

        c.setFont("Helvetica-Bold", 10.0)
        c.drawCentredString(x + largura / 2, y - 1.2 * mm, "Romaneio de Entrega")
        c.setFont("Helvetica", 6.6)
        c.drawRightString(x + largura - pad, y - 1.2 * mm, via_label)
        y -= 8.6 * mm

        cod_cli = upper(item.get("cod_cliente", ""))
        pedido = upper(fmt_pedido(item.get("pedido", "")))
        nome_cli = upper(item.get("nome_cliente", ""))
        endereco = upper(item.get("endereco", ""))
        cidade = self._extrair_cidade_do_endereco(endereco)
        produto = upper(item.get("produto") or "FRANGO VIVO")
        caixas = safe_int(item.get("qnt_caixas"), 0)
        aves_cx = safe_int(meta.get("aves_por_caixa"), 6)
        if aves_cx <= 0:
            aves_cx = 6
        total_aves = caixas * aves_cx
        kg = safe_float(item.get("kg"), 0.0)
        preco = self._normalizar_preco_item(item.get("preco"))
        valor_venda = kg * preco if (kg > 0 and preco > 0) else 0.0
        mort_aves = safe_int(item.get("mortalidade_aves"), 0)
        media_kg_ave = (kg / total_aves) if total_aves > 0 else 0.0
        mort_kg = mort_aves * media_kg_ave
        valor_mortalidade = mort_kg * preco if preco > 0 else 0.0
        valor_total_venda = valor_venda - valor_mortalidade
        if valor_total_venda < 0:
            valor_total_venda = 0.0
        vendedor = upper(item.get("vendedor", ""))
        if vendedor in {"NAN", "NONE", "NULL"}:
            vendedor = ""
        local_rota = meta.get("local_rota", "") or "-"
        local_carreg = meta.get("local_carregamento", "") or "-"
        tipo_estimativa = upper(meta.get("tipo_estimativa", "KG") or "KG")
        if tipo_estimativa == "CX":
            estimativa_txt = f"CIF / CX EST: {safe_int(meta.get('caixas_estimado'), 0)}"
        else:
            estimativa_txt = f"FOB / KG EST: {safe_float(meta.get('kg_estimado'), 0.0):.2f}"

        data_ref = datetime.now().strftime("%d/%m/%Y")

        c.setFont("Helvetica", 6.9)
        c.drawString(x + pad + logo_w + 4 * mm, y, f"COD CLIENTE: {cod_cli}")
        c.drawString(x + 90 * mm, y, f"PEDIDO: {pedido}")
        c.drawRightString(x + largura - pad, y, f"DATA: {data_ref}")
        y -= 4.2 * mm

        c.drawString(x + pad + logo_w + 4 * mm, y, f"RAZAO SOCIAL: {nome_cli[:75]}")
        y -= 4.0 * mm
        c.drawString(x + pad + logo_w + 4 * mm, y, f"NOME FANTASIA: {nome_cli[:73]}")
        y -= 4.0 * mm
        c.drawString(x + pad + logo_w + 4 * mm, y, f"ENDERECO: {endereco[:75]}")
        y -= 4.0 * mm
        c.drawString(x + pad + logo_w + 4 * mm, y, f"BAIRRO: {cidade if cidade else '-'}    CIDADE: {cidade if cidade else '-'}")

        y -= 4.6 * mm
        c.line(x + pad, y, x + largura - pad, y)
        y -= 3.9 * mm

        c.setFont("Helvetica-Bold", 7.7)
        c.drawString(x + pad, y, f"PRODUTO: {produto[:26]}")
        c.drawString(x + 62 * mm, y, f"PRECO DO KG: R$ {money(preco)}")
        c.drawString(x + 108 * mm, y, f"PESO MEDIO (KG): {money(media_kg_ave)}")
        c.drawRightString(x + largura - pad, y, f"LOCAL: {local_rota[:18]}")

        y -= 7.0 * mm

        c.setFont("Helvetica", 6.8)

        def box(label, value, bx, by, bw=30*mm, bh=4.7*mm):
            c.drawString(bx, by + bh + 0.7*mm, label)
            c.rect(bx, by, bw, bh, stroke=1, fill=0)
            if value not in (None, ""):
                c.drawRightString(bx + bw - 1.2*mm, by + 1.2*mm, str(value))

        # Grade principal no padrao do modelo Excel (3 colunas, 4 linhas)
        left_x = x + pad
        left_w = 35 * mm
        col_gap = 4 * mm
        mid_x = left_x + left_w + col_gap
        mid_w = 48 * mm
        right_x = mid_x + mid_w + col_gap
        right_w = largura - pad - right_x
        row_y = y

        step = 6.0 * mm
        box_h = 4.7 * mm
        box("Qtd. de Caixas:", caixas, left_x, row_y, bw=left_w, bh=box_h)
        box("Aves por Caixa:", aves_cx, left_x, row_y - step, bw=left_w, bh=box_h)
        box("Total de Aves:", total_aves, left_x, row_y - (2 * step), bw=left_w, bh=box_h)
        box("Peso Total:", money(kg), left_x, row_y - (3 * step), bw=left_w, bh=box_h)

        box("Valor da Venda:", money(valor_venda), mid_x, row_y, bw=mid_w, bh=box_h)
        box("Aves Mortas (und):", mort_aves, mid_x, row_y - step, bw=mid_w, bh=box_h)
        box("Desc. Mort (R$):", money(valor_mortalidade), mid_x, row_y - (2 * step), bw=mid_w, bh=box_h)
        box("Valor Final da venda:", money(valor_total_venda), mid_x, row_y - (3 * step), bw=mid_w, bh=box_h)

        box("Deb. Anterior do Cliente:", "", right_x, row_y, bw=right_w, bh=box_h)
        box("Valor recebido:", "", right_x, row_y - step, bw=right_w, bh=box_h)
        box("Valor recebido:", "", right_x, row_y - (2 * step), bw=right_w, bh=box_h)
        box("Recebido Total:", "", right_x, row_y - (3 * step), bw=right_w, bh=box_h)

        y_footer = y_base + 9.0 * mm
        c.setFont("Helvetica-Bold", 6.5)
        c.drawCentredString(
            x + largura / 2,
            y_footer,
            "CONTA PARA DEPOSITO BANCO DO BRASIL AGENCIA 0532-0 CONTA CORRENTE 25.852-0",
        )
        c.drawCentredString(
            x + largura / 2,
            y_footer - 3.8 * mm,
            "CHAVE PIX: 37.752.738/0001-15 (CNPJ)",
        )

        c.setFont("Helvetica", 6.6)
        c.drawString(x + pad, y_base + 2.2 * mm, f"PROGRAMACAO: {codigo_prog}   CARREGOU EM: {local_carreg}")
        c.drawRightString(x + largura - pad, y_base + 2.2 * mm, f"MOTORISTA: {meta.get('motorista', '-')}")
        c.setFont("Helvetica", 6.2)
        c.drawString(
            x + pad,
            y_base + 0.1 * mm,
            f"{estimativa_txt}   CRIADO POR: {meta.get('usuario_criacao', '-') or '-'}",
        )
        c.line(x + pad, y_base + 5.6 * mm, x + 85 * mm, y_base + 5.6 * mm)
        c.drawString(x + pad, y_base + 6.2 * mm, nome_cli[:38])

    def _gerar_pdf_romaneios(self, path: str, codigo: str, itens: list, meta: dict):
        from reportlab.lib.units import mm

        # Formulario continuo 5 1/2 (duas vias):
        # cada via: 240mm x 139,7mm | pagina final: 240mm x 279,4mm
        via_w = 240.0 * mm
        via_h = 139.7 * mm
        pagina_w = via_w
        pagina_h = via_h * 2.0

        c = canvas.Canvas(path, pagesize=(pagina_w, pagina_h))

        for item in itens:
            self._desenhar_bloco_romaneio(
                c, 0, pagina_h, via_w, via_h,
                codigo, meta, item, "VIA CLIENTE"
            )
            self._desenhar_bloco_romaneio(
                c, 0, via_h, via_w, via_h,
                codigo, meta, item, "VIA EMPRESA"
            )
            c.showPage()

        c.save()

    def _imprimir_romaneios(self, codigo: str, itens: list, meta: dict, parent=None):
        if not require_reportlab():
            return None

        printer = self._select_windows_printer(
            parent=parent or self.app,
            document_label=f"Romaneios de Entrega {codigo}",
            sheet_mode="Formulario continuo 240 x 279,4 mm",
            extra_hint=(
                "Os romaneios serao enviados com as duas vias (cliente e empresa) "
                "para a impressora selecionada."
            ),
        )
        if not printer:
            return None

        path = self._build_temp_pdf_path("ROMANEIOS", codigo)
        try:
            self._gerar_pdf_romaneios(path, codigo, itens, meta)
            self._send_file_to_windows_printer(path, printer)
            messagebox.showinfo(
                "Impressao enviada",
                f"Romaneios da programacao {codigo} enviados para:\n{printer.get('name') or 'IMPRESSORA PADRAO'}",
                parent=parent or self.app,
            )
            return path
        except Exception as exc:
            messagebox.showerror(
                "ERRO",
                "Nao foi possivel enviar os romaneios para a impressora.\n\n"
                f"Detalhe: {exc}",
                parent=parent or self.app,
            )
            return None

    def _draw_romaneio_preview_on_canvas(self, cv, w: int, h: int, codigo: str, meta: dict, item: dict, via_label: str):
        cv.delete("all")
        # Mantem preview com proporcao real da via: 240mm x 139,7mm
        ratio = 240.0 / 139.7
        m = 16
        avail_w = max(200, w - (2 * m))
        avail_h = max(120, h - (2 * m))
        target_w = avail_w
        target_h = int(target_w / ratio)
        if target_h > avail_h:
            target_h = avail_h
            target_w = int(target_h * ratio)
        x0 = (w - target_w) // 2
        y0 = (h - target_h) // 2
        x1 = x0 + target_w
        y1 = y0 + target_h
        cv.create_rectangle(x0, y0, x1, y1, outline="#111", width=1)

        pad = 10
        lx = x0 + pad
        ty = y0 + pad

        # Header
        logo_w, logo_h = 112, 64
        cv.create_rectangle(lx, ty, lx + logo_w, ty + logo_h, outline="#222", width=1)
        cv.create_text(lx + logo_w / 2, ty + logo_h / 2, text="LOGO", font=("Segoe UI", 10, "bold"))

        cv.create_text((x0 + x1) / 2, ty + 2, text="Romaneio de Entrega", font=("Segoe UI", 16, "bold"), anchor="n")
        cv.create_text(x1 - pad, ty + 8, text=via_label, font=("Segoe UI", 10, "normal"), anchor="ne")

        def fmt_pedido(v):
            s = str(v or "").strip()
            if not s:
                return ""
            try:
                n = float(s.replace(",", "."))
                if abs(n - int(n)) < 1e-9:
                    return str(int(n))
            except Exception:
                pass
            return s

        cod_cli = upper(item.get("cod_cliente", ""))
        pedido = upper(fmt_pedido(item.get("pedido", "")))
        nome_cli = upper(item.get("nome_cliente", ""))
        endereco = upper(item.get("endereco", ""))
        cidade = self._extrair_cidade_do_endereco(endereco)
        produto = upper(item.get("produto") or "FRANGO VIVO")
        caixas = safe_int(item.get("qnt_caixas"), 0)
        aves_cx = safe_int(meta.get("aves_por_caixa"), 6)
        if aves_cx <= 0:
            aves_cx = 6
        total_aves = caixas * aves_cx
        kg = safe_float(item.get("kg"), 0.0)
        preco = self._normalizar_preco_item(item.get("preco"))
        valor_venda = kg * preco if (kg > 0 and preco > 0) else 0.0
        mort_aves = safe_int(item.get("mortalidade_aves"), 0)
        media_kg_ave = (kg / total_aves) if total_aves > 0 else 0.0
        mort_kg = mort_aves * media_kg_ave
        valor_mortalidade = mort_kg * preco if preco > 0 else 0.0
        valor_total_venda = valor_venda - valor_mortalidade
        if valor_total_venda < 0:
            valor_total_venda = 0.0
        vendedor = upper(item.get("vendedor", ""))
        if vendedor in {"NAN", "NONE", "NULL"}:
            vendedor = ""
        local_rota = meta.get("local_rota", "") or "-"
        local_carreg = meta.get("local_carregamento", "") or "-"
        data_ref = datetime.now().strftime("%d/%m/%Y")

        def money(v):
            return f"{safe_float(v, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        tx = lx + logo_w + 14
        y = ty + 20
        cv.create_text(tx, y + 14, text=f"COD CLIENTE: {cod_cli}", anchor="w", font=("Segoe UI", 10))
        cv.create_text(tx + 260, y + 14, text=f"PEDIDO: {pedido}", anchor="w", font=("Segoe UI", 10))
        cv.create_text(x1 - pad, y + 14, text=f"DATA: {data_ref}", anchor="e", font=("Segoe UI", 10))
        cv.create_text(tx, y + 38, text=f"RAZAO SOCIAL: {nome_cli[:42]}", anchor="w", font=("Segoe UI", 10))
        cv.create_text(tx, y + 60, text=f"NOME FANTASIA: {nome_cli[:40]}", anchor="w", font=("Segoe UI", 10))
        cv.create_text(tx, y + 82, text=f"ENDERECO: {endereco[:44]}", anchor="w", font=("Segoe UI", 10))
        cv.create_text(tx, y + 104, text=f"BAIRRO: {cidade if cidade else '-'}    CIDADE: {cidade if cidade else '-'}", anchor="w", font=("Segoe UI", 10))

        sep_y = ty + 138
        cv.create_line(lx, sep_y, x1 - pad, sep_y, fill="#222", width=1)
        cv.create_text(lx, sep_y + 12, text=f"PRODUTO: {produto[:24]}", anchor="w", font=("Segoe UI", 10, "bold"))
        cv.create_text(lx + 300, sep_y + 12, text=f"PRECO DO KG: R$ {money(preco)}", anchor="w", font=("Segoe UI", 10, "bold"))
        cv.create_text(lx + 480, sep_y + 12, text=f"PESO MEDIO (KG): {money(media_kg_ave)}", anchor="w", font=("Segoe UI", 10))
        cv.create_text(x1 - pad - 6, sep_y + 12, text=f"LOCAL: {local_rota[:16]}", anchor="e", font=("Segoe UI", 10, "bold"))

        # Boxes
        by = sep_y + 48
        content_w = (x1 - pad) - lx
        gap = 12
        left_w = max(165, int((content_w - 2 * gap) * 0.30))
        mid_w = max(220, int((content_w - 2 * gap) * 0.38))
        right_w = (content_w - 2 * gap) - left_w - mid_w
        if right_w < 175:
            delta = 175 - right_w
            mid_w = max(190, mid_w - delta)
            right_w = 175
        left_x = lx
        mid_x = left_x + left_w + gap
        right_x = mid_x + mid_w + gap
        row_h = 44
        box_h = 28

        def draw_box(bx, byy, bw, label, value):
            cv.create_text(bx, byy - 12, text=label, anchor="w", font=("Segoe UI", 9))
            cv.create_rectangle(bx, byy, bx + bw, byy + box_h, outline="#222", width=1)
            if value is not None and str(value) != "":
                cv.create_text(bx + bw - 6, byy + (box_h // 2), text=str(value), anchor="e", font=("Segoe UI", 11))

        draw_box(left_x, by, left_w, "Qtd. de Caixas:", caixas)
        draw_box(left_x, by + row_h, left_w, "Aves por Caixa:", aves_cx)
        draw_box(left_x, by + row_h * 2, left_w, "Total de Aves:", total_aves)
        draw_box(left_x, by + row_h * 3, left_w, "Peso Total:", money(kg))

        draw_box(mid_x, by, mid_w, "Valor da Venda:", money(valor_venda))
        draw_box(mid_x, by + row_h, mid_w, "Aves Mortas (und):", mort_aves)
        draw_box(mid_x, by + row_h * 2, mid_w, "Desc. Mort (R$):", money(valor_mortalidade))
        draw_box(mid_x, by + row_h * 3, mid_w, "Valor Final da venda:", money(valor_total_venda))

        draw_box(right_x, by, right_w, "Deb. Anterior do Cliente:", "")
        draw_box(right_x, by + row_h, right_w, "Valor recebido:", "")
        draw_box(right_x, by + row_h * 2, right_w, "Valor recebido:", "")
        draw_box(right_x, by + row_h * 3, right_w, "Recebido Total:", "")

        # Footer
        cv.create_text((x0 + x1) / 2, y1 - 52, text="CONTA PARA DEPOSITO BANCO DO BRASIL AGENCIA 0532-0 CONTA CORRENTE 25.852-0", font=("Segoe UI", 10, "bold"))
        cv.create_text((x0 + x1) / 2, y1 - 30, text="CHAVE PIX: 37.752.738/0001-15 (CNPJ)", font=("Segoe UI", 10, "bold"))
        cv.create_text(lx, y1 - 12, text=f"PROGRAMACAO: {codigo}   CARREGOU EM: {local_carreg}", anchor="w", font=("Segoe UI", 10))
        cv.create_text(x1 - pad, y1 - 12, text=f"MOTORISTA: {meta.get('motorista', '-')}", anchor="e", font=("Segoe UI", 10))

    def _abrir_previsualizacao_romaneios(self, codigo: str, itens: list, meta: dict):
        if not itens:
            messagebox.showwarning("ATENCAO", "Sem itens para pre-visualizar.")
            return

        def _fmt_pedido(v):
            s = str(v or "").strip()
            if not s:
                return ""
            try:
                n = float(s.replace(",", "."))
                if abs(n - int(n)) < 1e-9:
                    return str(int(n))
            except Exception:
                pass
            return s

        win = tk.Toplevel(self.app)
        win.title(f"Pre-visualizacao Romaneios - {codigo}")
        win.geometry("1250x820")
        win.transient(self.app)
        win.grab_set()

        top = ttk.Frame(win, padding=8)
        top.pack(fill="x")
        center = ttk.Frame(win, padding=8)
        center.pack(fill="both", expand=True)
        bottom = ttk.Frame(win, padding=8)
        bottom.pack(fill="x")

        ttk.Label(top, text="Cliente:", style="CardLabel.TLabel").pack(side="left")
        cb = ttk.Combobox(top, state="readonly", width=60)
        cb.pack(side="left", padx=6)
        vias_var = tk.StringVar(value="VIA CLIENTE")
        ttk.Radiobutton(top, text="Via Cliente", value="VIA CLIENTE", variable=vias_var).pack(side="left", padx=8)
        ttk.Radiobutton(top, text="Via Empresa", value="VIA EMPRESA", variable=vias_var).pack(side="left")

        preview = tk.Canvas(center, bg="white", highlightthickness=1, highlightbackground="#D1D5DB")
        preview.pack(fill="both", expand=True)

        options = []
        for idx, it in enumerate(itens):
            cod = upper(it.get("cod_cliente", ""))
            nome = upper(it.get("nome_cliente", ""))
            pedido = upper(_fmt_pedido(it.get("pedido", "")))
            options.append((idx, f"{cod} - {nome} (PED {pedido})"))
        cb["values"] = [o[1] for o in options]
        cb.current(0)

        state = {"idx": 0}

        def _render():
            idx = state["idx"]
            item = itens[idx]
            w = max(preview.winfo_width(), 900)
            h = max(preview.winfo_height(), 620)
            self._draw_romaneio_preview_on_canvas(preview, w, h, codigo, meta, item, vias_var.get())

        def _on_sel(_e=None):
            pos = cb.current()
            if pos < 0:
                return
            state["idx"] = options[pos][0]
            _render()

        def _prev():
            if state["idx"] <= 0:
                return
            state["idx"] -= 1
            cb.current(state["idx"])
            _render()

        def _next():
            if state["idx"] >= len(itens) - 1:
                return
            state["idx"] += 1
            cb.current(state["idx"])
            _render()

        def _export_pdf():
            path = filedialog.asksaveasfilename(
                title="Salvar Romaneios de Entrega",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"ROMANEIOS_{codigo}.pdf",
            )
            if not path:
                return
            try:
                self._gerar_pdf_romaneios(path, codigo, itens, meta)
                messagebox.showinfo("OK", f"Romaneios gerados com sucesso!\n\nArquivo: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("ERRO", f"Erro ao gerar romaneios: {str(e)}")

        def _imprimir():
            self._imprimir_romaneios(codigo, itens, meta, parent=win)

        ttk.Label(top, text="Folha: formulário contínuo 240 x 279,4 mm", style="CardLabel.TLabel").pack(side="left", padx=(18, 0))
        ttk.Button(top, text="\U0001F5A8 Imprimir", style="Primary.TButton", command=_imprimir).pack(side="right")

        ttk.Button(bottom, text="Anterior", style="Ghost.TButton", command=_prev).pack(side="left")
        ttk.Button(bottom, text="Proximo", style="Ghost.TButton", command=_next).pack(side="left", padx=8)
        ttk.Button(bottom, text="GERAR PDF", style="Primary.TButton", command=_export_pdf).pack(side="right")
        ttk.Button(bottom, text="IMPRIMIR", style="Ghost.TButton", command=_imprimir).pack(side="right", padx=(0, 8))
        ttk.Button(bottom, text="Fechar", style="Danger.TButton", command=win.destroy).pack(side="right", padx=8)

        cb.bind("<<ComboboxSelected>>", _on_sel)
        vias_var.trace_add("write", lambda *_: _render())
        preview.bind("<Configure>", lambda _e: _render())
        _render()

    def imprimir_romaneios_programacao(self, codigo_override=None):
        if not require_reportlab():
            return

        codigo = upper(str(codigo_override or "").strip()) or upper((self.ent_codigo.get() or "").strip())
        if not codigo:
            codigo = upper(simple_input("Romaneios", "Informe o codigo da programacao:", master=self.app, allow_empty=False) or "")
        if not codigo:
            return

        itens_res = fetch_programacao_itens_contract(codigo)
        itens = itens_res.get("data") if isinstance(itens_res, dict) and bool(itens_res.get("ok", False)) else []
        if not itens:
            messagebox.showwarning("ATENCAO", f"Sem itens para gerar romaneios na programacao {codigo}.")
            return

        meta = self._buscar_meta_programacao(codigo)
        self._abrir_previsualizacao_romaneios(codigo, itens, meta)


# ==========================
# ===== FIM DA PARTE 5 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 6 (ATUALIZADA) =====
# ==========================

# =========================================================
# 6.0 FUNÇÃO SIMPLE INPUT (antes de RecebimentosPage)
# =========================================================
def simple_input(title, prompt, master=None, initial="", allow_empty=True, max_len=200):
    """
    Janela de diálogo simples para entrada de texto.

    CompatÃÂÂvel com seu uso atual:
        simple_input("TÃÂÂtulo", "Pergunta?")

    Extras opcionais (não quebram):
        master=app  -> mantém a janela presa ao app
        initial="..." -> valor inicial
        allow_empty=False -> obriga preencher
        max_len -> limita tamanho (segurança)
    """
    # Se não passar master, tenta usar a janela root atual
    if master is None:
        try:
            master = tk._default_root
        except Exception:
            master = None

    win = tk.Toplevel(master) if master else tk.Toplevel()
    win.title(title)
    win.geometry("540x220")
    win.resizable(False, False)
    win.grab_set()

    # Melhor UX: janela modal ligada ao app
    if master:
        try:
            win.transient(master)
        except Exception:
            logging.debug("Falha ignorada")

    # Layout moderno (card)
    outer = ttk.Frame(win, style="Content.TFrame", padding=14)
    outer.pack(fill="both", expand=True)

    card = ttk.Frame(outer, style="Card.TFrame", padding=14)
    card.pack(fill="both", expand=True)
    card.grid_columnconfigure(0, weight=1)

    ttk.Label(card, text=title, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(card, text=prompt, style="CardLabel.TLabel", justify="left", wraplength=470).grid(row=1, column=0, sticky="w", pady=(8, 8))

    ent = ttk.Entry(card, style="Field.TEntry")
    ent.grid(row=2, column=0, sticky="ew")
    if initial:
        ent.insert(0, str(initial)[:max_len])
        ent.selection_range(0, "end")
    ent.focus_set()

    result = {"value": ""}

    def _get_value():
        v = str(ent.get() or "")
        v = v.strip()
        if max_len and len(v) > max_len:
            v = v[:max_len]
        return v

    def ok():
        v = _get_value()
        if (not allow_empty) and (not v):
            messagebox.showwarning("ATENÇÃO", "Preencha o campo antes de confirmar.")
            ent.focus_set()
            return
        result["value"] = v
        try:
            win.destroy()
        except Exception:
            logging.debug("Falha ignorada")

    def cancel():
        result["value"] = ""
        try:
            win.destroy()
        except Exception:
            logging.debug("Falha ignorada")

    # Se fechar no X, sempre cancela
    win.protocol("WM_DELETE_WINDOW", cancel)

    btns = ttk.Frame(card, style="Card.TFrame")
    btns.grid(row=3, column=0, sticky="ew", pady=(14, 0))
    btns.grid_columnconfigure(0, weight=1)
    btns.grid_columnconfigure(1, weight=1)

    ttk.Button(btns, text="CONFIRMAR", style="Primary.TButton", command=ok).grid(row=0, column=0, sticky="ew", padx=(0, 6))
    ttk.Button(btns, text="CANCELAR", style="Ghost.TButton", command=cancel).grid(row=0, column=1, sticky="ew", padx=(6, 0))

    # Atalhos
    win.bind("<Return>", lambda e: ok())
    win.bind("<Escape>", lambda e: cancel())

    # Centraliza na tela (simples e sem depender de libs)
    try:
        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = int((sw / 2) - (w / 2))
        y = int((sh / 2) - (h / 2))
        win.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        logging.debug("Falha ignorada")

    win.wait_window()
    return result["value"]

# ==========================
# ===== FIM DA PARTE 6 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INICIO DA PARTE 6B (ATUALIZADA) =====
# ==========================

class RecebimentosPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Recebimentos")

        self._current_prog = ""
        self._selected_cliente = None
        self._sort_reverse = False
        self._current_sort_column = None
        self._is_collapsed_after_close = False  # modo oculto apos finalizar/salvar
        self._rota_atual = ""
        self._equipe_raw = ""

        # -------------------------
        # Card superior (programacao + dados de rota)
        # -------------------------
        self.card = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        self.card.grid(row=0, column=0, sticky="ew")
        self.body.grid_columnconfigure(0, weight=1)

        for i in range(0, 4):
            self.card.grid_columnconfigure(i, weight=0)
        self.card.grid_columnconfigure(2, weight=1)  # info textual
        self.card.grid_columnconfigure(3, weight=0)  # bloco fixo de datas/horas

        ttk.Label(self.card, text="Programacao (pendente)", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.cb_prog = ttk.Combobox(self.card, state="readonly", width=24)
        self.cb_prog.grid(row=1, column=0, sticky="w", padx=(0, 10))

        ttk.Button(self.card, text="\U0001F4E6 CARREGAR", style="Ghost.TButton", command=self.carregar_programacao)\
            .grid(row=1, column=1, padx=(0, 14))

        info_frame = ttk.Frame(self.card, style="Card.TFrame")
        info_frame.grid(row=1, column=2, sticky="ew", padx=(8, 8))
        for i in range(2):
            info_frame.grid_columnconfigure(i, weight=1)

        self.lbl_motorista_info = ttk.Label(
            info_frame,
            text="Motorista: -",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
            wraplength=260,
            justify="left"
        )
        self.lbl_motorista_info.grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.lbl_veiculo_info = ttk.Label(
            info_frame,
            text="Veiculo: -",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
            wraplength=260,
            justify="left"
        )
        self.lbl_veiculo_info.grid(row=1, column=0, sticky="w", padx=(0, 8))

        self.lbl_equipe_info = ttk.Label(
            info_frame,
            text="Equipe: -",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
            wraplength=320,
            justify="left"
        )
        self.lbl_equipe_info.grid(row=0, column=1, sticky="w", padx=(0, 8))

        self.lbl_rota_info = ttk.Label(
            info_frame,
            text="Rota: -",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
            wraplength=320,
            justify="left"
        )
        self.lbl_rota_info.grid(row=1, column=1, sticky="w")

        horarios_frame = ttk.Frame(self.card, style="Card.TFrame")
        horarios_frame.grid(row=0, column=3, rowspan=2, sticky="e")

        ttk.Label(horarios_frame, text="Diaria Motorista (R$)", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_diaria_motorista = ttk.Entry(horarios_frame, style="Field.TEntry", width=12)
        self.ent_diaria_motorista.grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.ent_diaria_motorista.insert(0, "0,00")
        self._bind_money_entry(self.ent_diaria_motorista)

        ttk.Label(horarios_frame, text="Saida (data)", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.ent_data_saida = ttk.Entry(horarios_frame, style="Field.TEntry", width=12)
        self.ent_data_saida.grid(row=1, column=1, sticky="w", padx=(0, 8))
        self._bind_date_entry(self.ent_data_saida)

        ttk.Label(horarios_frame, text="Saida (hora)", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.ent_hora_saida = ttk.Entry(horarios_frame, style="Field.TEntry", width=9)
        self.ent_hora_saida.grid(row=1, column=2, sticky="w", padx=(0, 8))
        self._bind_time_entry(self.ent_hora_saida)

        ttk.Label(horarios_frame, text="Chegada (data)", style="CardLabel.TLabel").grid(row=0, column=3, sticky="w")
        self.ent_data_chegada = ttk.Entry(horarios_frame, style="Field.TEntry", width=12)
        self.ent_data_chegada.grid(row=1, column=3, sticky="w", padx=(0, 8))
        self._bind_date_entry(self.ent_data_chegada)

        ttk.Label(horarios_frame, text="Chegada (hora)", style="CardLabel.TLabel").grid(row=0, column=4, sticky="w")
        self.ent_hora_chegada = ttk.Entry(horarios_frame, style="Field.TEntry", width=9)
        self.ent_hora_chegada.grid(row=1, column=4, sticky="w")
        self._bind_time_entry(self.ent_hora_chegada)


        resumo = ttk.Frame(self.card, style="Card.TFrame")
        resumo.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        for i in range(3):
            resumo.grid_columnconfigure(i, weight=1)

        self.lbl_resumo_diaria = ttk.Label(
            resumo, text="VALOR DIARIA: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_diaria.grid(row=0, column=0, sticky="w", padx=(0, 16), pady=(0, 4))

        self.lbl_resumo_qtd = ttk.Label(
            resumo, text="QTD DIARIAS: 0", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_qtd.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=(0, 4))

        self.lbl_resumo_mot = ttk.Label(
            resumo, text="MOTORISTA: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_mot.grid(row=0, column=2, sticky="w", padx=(0, 16), pady=(0, 4))

        self.lbl_resumo_eqp = ttk.Label(
            resumo, text="EQUIPE: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_eqp.grid(row=1, column=0, sticky="w", padx=(0, 16))

        self.lbl_resumo_total = ttk.Label(
            resumo, text="TOTAL: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_total.grid(row=1, column=1, sticky="w", padx=(0, 16))

        self.lbl_resumo_pagar = ttk.Label(
            resumo, text="VALOR A PAGAR: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_pagar.grid(row=1, column=2, sticky="w")

        self._bind_diarias_preview()

        # -------------------------
        # Card principal (tabela + filtros + formulario)
        # -------------------------
        self.card2 = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        self.card2.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.body.grid_rowconfigure(1, weight=1)

        self.card2.grid_columnconfigure(0, weight=1)
        self.card2.grid_rowconfigure(2, weight=1)

        top2 = ttk.Frame(self.card2, style="Card.TFrame")
        top2.grid(row=0, column=0, sticky="ew")
        top2.grid_columnconfigure(30, weight=1)

        ttk.Button(top2, text="\U0001F464 INSERIR CLIENTE MANUAL", style="Warn.TButton", command=self.inserir_cliente_manual)\
            .grid(row=0, column=0, padx=6)

        ttk.Button(top2, text="\U0001F9FD ZERAR RECEBIMENTO", style="Danger.TButton", command=self.zerar_recebimento)\
            .grid(row=0, column=1, padx=6)

        ttk.Button(top2, text="\u27A1 IR PARA DESPESAS", style="Primary.TButton", command=self._ir_para_despesas)\
            .grid(row=0, column=2, padx=6)

        self.lbl_total = ttk.Label(
            top2,
            text="TOTAL RECEBIDO: R$ 0,00",
            background="white",
            foreground="#111827",
            font=("Segoe UI", 10, "bold")
        )
        self.lbl_total.grid(row=0, column=30, sticky="e", padx=6)

        ttk.Separator(self.card2).grid(row=1, column=0, sticky="ew", pady=10)

        cols = ["COD", "CLIENTE", "VALOR", "FORMA", "OBS", "DATA REGISTRO"]

        table_wrap = ttk.Frame(self.card2, style="Card.TFrame")
        table_wrap.grid(row=2, column=0, sticky="nsew")
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150, minwidth=90, anchor="w")

        self.tree.column("COD", width=90, minwidth=80, anchor="center")
        self.tree.column("CLIENTE", width=320, minwidth=220)
        self.tree.column("VALOR", width=120, minwidth=90, anchor="e")
        self.tree.column("FORMA", width=110, minwidth=90, anchor="center")
        self.tree.column("OBS", width=260, minwidth=140)
        self.tree.column("DATA REGISTRO", width=160, minwidth=130, anchor="center")

        for col in cols:
            self.tree.heading(col, command=lambda c=col: self.sort_by_column(c))

        try:
            enable_treeview_sorting(
                self.tree,
                numeric_cols=set(),
                money_cols={"VALOR"},
                date_cols={"DATA REGISTRO"}
            )
        except Exception:
            logging.debug("Falha ignorada")

        self.tree.bind("<<TreeviewSelect>>", self._on_select_row)
        self.tree.bind("<Double-1>", self._on_tree_double_click_value, add=True)

        # -------------------------
        # Formulário de recebimento
        # -------------------------
        frm = ttk.Frame(self.card2, style="Card.TFrame")
        frm.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        frm.grid_columnconfigure(4, weight=1)

        ttk.Label(frm, text="Cód Cliente", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_cod = ttk.Entry(frm, style="Field.TEntry", width=12)
        self.ent_cod.grid(row=1, column=0, sticky="w", padx=6)
        bind_entry_smart(self.ent_cod, "int")

        ttk.Label(frm, text="Nome", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.ent_nome = ttk.Entry(frm, style="Field.TEntry", width=28)
        self.ent_nome.grid(row=1, column=1, sticky="w", padx=6)
        bind_entry_smart(self.ent_nome, "text")

        ttk.Label(frm, text="Valor", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.ent_valor = ttk.Entry(frm, style="Field.TEntry", width=12)
        self.ent_valor.grid(row=1, column=2, sticky="w", padx=6)
        bind_entry_smart(self.ent_valor, "money")
        self.ent_valor.bind("<Return>", lambda e: self.salvar_recebimento())
        self._bind_money_entry(self.ent_valor)

        ttk.Label(frm, text="Forma", style="CardLabel.TLabel").grid(row=0, column=3, sticky="w")
        self.cb_forma = ttk.Combobox(
            frm,
            state="readonly",
            width=14,
            values=["DINHEIRO", "PIX", "CARTAO", "BOLETO", "OUTRO"]
        )
        self.cb_forma.grid(row=1, column=3, sticky="w", padx=6)
        self.cb_forma.set("DINHEIRO")

        ttk.Label(frm, text="Observação", style="CardLabel.TLabel").grid(row=0, column=4, sticky="w")
        self.ent_obs = ttk.Entry(frm, style="Field.TEntry", width=40)
        self.ent_obs.grid(row=1, column=4, sticky="ew", padx=6)
        bind_entry_smart(self.ent_obs, "text")

        ttk.Button(frm, text="\U0001F4BE SALVAR RECEBIMENTOS", style="Primary.TButton", command=self.salvar_recebimento)\
            .grid(row=1, column=5, sticky="e", padx=(12, 0))

        # âÅâ€œââ‚¬¦ TROCA: no lugar do Excel, botão IMPRIMIR PDF
        ttk.Button(frm, text="\U0001F5A8 IMPRIMIR PDF", style="Warn.TButton", command=self.imprimir_pdf)\
            .grid(row=1, column=6, sticky="e", padx=6)

        # Por padrão: COD/NOME não editáveis
        self._set_cliente_fields_readonly(True)

        # -------------------------
        # Painel âââââ€š¬Å¡¬Åâââ€š¬Åâ€œmodo ocultoâââââ€š¬Å¡¬Â (apÃÆââ‚¬â„¢³s finalizar prestaÃÆââ‚¬â„¢§ÃÆââ‚¬â„¢£o)
        # -------------------------
        self._wrap_collapsed = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        self._wrap_collapsed.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        self._wrap_collapsed.grid_remove()

        self._lbl_collapsed = ttk.Label(
            self._wrap_collapsed,
            text="PRESTAÇÃO FECHADA / SALVA.\nCabeçalhos e tabela foram ocultados.",
            background="white",
            foreground="#111827",
            font=("Segoe UI", 10, "bold"),
            justify="left"
        )
        self._lbl_collapsed.grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(self._wrap_collapsed, style="Card.TFrame")
        btns.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        btns.grid_columnconfigure(10, weight=1)

        ttk.Button(btns, text="\U0001F5A8 IMPRIMIR PDF", style="Warn.TButton", command=self.imprimir_pdf)\
            .grid(row=0, column=0, padx=6, sticky="w")

        ttk.Button(btns, text="\U0001F441 MOSTRAR DADOS (CONSULTA)", style="Ghost.TButton", command=self._expand_view)\
            .grid(row=0, column=1, padx=6, sticky="w")

        ttk.Button(btns, text="LIMPAR / NOVA PROGRAMAÇÃO", style="Primary.TButton", command=self._reset_view)\
            .grid(row=0, column=2, padx=6, sticky="w")

        self.refresh_comboboxes()
        self.carregar_tabela_vazia()
        self._refresh_diarias_preview()

    # =========================
    # Modos de visualização (ocultar/mostrar)
    # =========================
    def _collapse_view(self):
        """Oculta cabeçalhos e dados de tabela (mantém imprimir/limpar)."""
        self._is_collapsed_after_close = True
        try:
            self.card.grid_remove()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self.card2.grid_remove()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self._wrap_collapsed.grid()
        except Exception:
            logging.debug("Falha ignorada")

    def _expand_view(self):
        """Mostra novamente cabeçalho e tabela."""
        self._is_collapsed_after_close = False  # modo oculto apos finalizar/salvar
        try:
            self._wrap_collapsed.grid_remove()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self.card.grid()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self.card2.grid()
        except Exception:
            logging.debug("Falha ignorada")

        if self._is_prestacao_fechada(self._current_prog):
            self.set_status(f"STATUS: Programacao {self._current_prog} (PRESTACAO FECHADA - somente consulta)")
        else:
            self.set_status(f"STATUS: Programacao {self._current_prog}")

    def _reset_view(self):
        """Limpa seleção e volta para estado inicial."""
        self._expand_view()
        self._current_prog = ""
        self.cb_prog.set("")
        self._nf_current = ""
        self.lbl_motorista_info.config(text="Motorista: -")
        self.lbl_veiculo_info.config(text="Veiculo: -")
        self.lbl_equipe_info.config(text="Equipe: -")
        self.lbl_rota_info.config(text="Rota: -")
        self._rota_atual = ""
        self._equipe_raw = ""
        self._safe_set_entry(self.ent_diaria_motorista, "0,00", readonly_back=False)
        self.ent_data_saida.delete(0, "end")
        self.ent_hora_saida.delete(0, "end")
        self.ent_data_chegada.delete(0, "end")
        self.ent_hora_chegada.delete(0, "end")
        self._refresh_diarias_preview()
        self._selected_cliente = None
        self._clear_form_recebimento()
        self.carregar_tabela_vazia()
        try:
            self.cb_filtro_forma.set("TODAS")
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self.ent_filtro_valor_min.delete(0, "end")
        except Exception:
            logging.debug("Falha ignorada")

        self.refresh_comboboxes()
        if hasattr(self.app, "refresh_programacao_comboboxes"):
            try:
                self.app.refresh_programacao_comboboxes()
            except Exception:
                logging.debug("Falha ignorada")

        self.set_status("STATUS: Selecione uma programacao para Recebimentos. Rotas em aberto serao consolidadas ao finalizar.")

    # -------------------------
    # Helpers de segurança UI (readonly sem quebrar inserts)
    # -------------------------
    def _set_cliente_fields_readonly(self, readonly: bool):
        try:
            self.ent_cod.configure(state=("readonly" if readonly else "normal"))
            self.ent_nome.configure(state=("readonly" if readonly else "normal"))
        except Exception:
            logging.debug("Falha ignorada")

    def _safe_set_entry(self, entry, value: str, readonly_back: bool = None):
        try:
            prev = entry.cget("state")
        except Exception:
            prev = None

        try:
            try:
                entry.configure(state="normal")
            except Exception:
                logging.debug("Falha ignorada")
            entry.delete(0, "end")
            entry.insert(0, value if value is not None else "")
        finally:
            try:
                if readonly_back is None:
                    if prev is not None:
                        entry.configure(state=prev)
                else:
                    entry.configure(state=("readonly" if readonly_back else "normal"))
            except Exception:
                logging.debug("Falha ignorada")

    # -------------------------
    # Resolução de nomes (motorista/equipe)
    # -------------------------
    def _resolve_motorista_nome(self, motorista_raw: str) -> str:
        m = (motorista_raw or "").strip()
        if not m:
            return ""
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/cadastros/motoristas",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                alvo = upper(m)
                for r in (resp or []):
                    if not isinstance(r, dict):
                        continue
                    cod = upper(str(r.get("codigo") or "").strip())
                    if cod and cod == alvo:
                        return upper(str(r.get("nome") or "").strip())
            except Exception:
                logging.debug("Falha ao resolver motorista via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                # Evita OperationalError em bases legadas sem tabela/coluna esperada.
                candidatos = []
                for tabela in ("motoristas", "cadastro_motoristas"):
                    try:
                        cur.execute(
                            "SELECT name FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
                            (tabela,)
                        )
                        existe_tabela = cur.fetchone() is not None
                    except Exception:
                        existe_tabela = False
                    if not existe_tabela:
                        continue
                    if not db_has_column(cur, tabela, "nome"):
                        continue
                    for coluna in ("codigo", "cod_motorista"):
                        if db_has_column(cur, tabela, coluna):
                            candidatos.append((tabela, coluna))

                for tabela, coluna in candidatos:
                    try:
                        cur.execute(
                            f"SELECT nome FROM {_safe_ident(tabela)} WHERE {_safe_ident(coluna)}=? LIMIT 1",
                            (m,)
                        )
                        r = cur.fetchone()
                        if r and r[0]:
                            return upper(r[0])
                    except Exception:
                        continue
        except Exception:
            logging.debug("Falha ignorada")
        return upper(m)

    def _resolve_equipe_integrantes(self, equipe_raw: str) -> str:
        return resolve_equipe_nomes(equipe_raw)

    # -------------------------
    # Ordenação (padronizada / robusta)
    # -------------------------
    def sort_by_column(self, col):
        if self._current_sort_column == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._current_sort_column = col
            self._sort_reverse = False

        reverse = self._sort_reverse

        def _to_float(v):
            try:
                s = str(v or "").strip()
                if not s or s in {"", "-", "None"}:
                    return 0.0
                s = s.replace("R$", "").replace(" ", "")
                s = s.replace(".", "").replace(",", ".")
                return float(s)
            except Exception:
                return 0.0

        def _to_date_key(v):
            s = str(v or "").strip()
            if not s:
                return (0, 0, 0, 0, 0, 0)

            if "-" in s and len(s) >= 10:
                try:
                    y, m, d = s[:10].split("-")
                    hh = mm = ss = 0
                    if len(s) >= 19 and ":" in s[11:19]:
                        hh, mm, ss = s[11:19].split(":")
                    return (int(y), int(m), int(d), int(hh), int(mm), int(ss))
                except Exception:
                    logging.debug("Falha ignorada")

            if "/" in s and len(s) >= 10:
                try:
                    d, m, y = s[:10].split("/")
                    hh = mm = ss = 0
                    if len(s) >= 19 and ":" in s[11:19]:
                        hh, mm, ss = s[11:19].split(":")
                    return (int(y), int(m), int(d), int(hh), int(mm), int(ss))
                except Exception:
                    logging.debug("Falha ignorada")

            return (0, 0, 0, 0, 0, 0)

        data = [(self.tree.set(child, col), child) for child in self.tree.get_children("")]

        if col == "VALOR":
            data.sort(key=lambda x: _to_float(x[0]), reverse=reverse)
        elif col == "DATA REGISTRO":
            data.sort(key=lambda x: _to_date_key(x[0]), reverse=reverse)
        else:
            data.sort(key=lambda x: str(x[0] or "").upper(), reverse=reverse)

        for idx, (_, child) in enumerate(data):
            self.tree.move(child, "", idx)

        arrow = " ââââ€š¬ âââ€š¬Åâ€œ" if reverse else " ââââ€š¬ âââ€š¬ËÅ“"
        for c in self.tree["columns"]:
            current_text = self.tree.heading(c, "text")
            if current_text.endswith(" ââââ€š¬ âââ€š¬ËÅ“") or current_text.endswith(" ââââ€š¬ âââ€š¬Åâ€œ"):
                current_text = current_text[:-2]
            self.tree.heading(c, text=current_text)

        current_text = self.tree.heading(col, "text")
        if current_text.endswith(" ââââ€š¬ âââ€š¬ËÅ“") or current_text.endswith(" ââââ€š¬ âââ€š¬Åâ€œ"):
            current_text = current_text[:-2]
        self.tree.heading(col, text=current_text + arrow)

    # -------------------------
    # Combobox programação (pendentes)
    # -------------------------
    def _rota_apt_para_recebimentos(self, prog: str) -> bool:
        prog = upper(str(prog or "").strip())
        if not prog:
            return False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(prog)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
                if isinstance(rota, dict):
                    st_op = upper(str(rota.get("status_operacional") or rota.get("status") or ""))
                    fin_app = safe_int(rota.get("finalizada_no_app"), 0)
                    prest = upper(str(rota.get("prestacao_status") or "PENDENTE"))
                    if st_op:
                        if fin_app != 1:
                            return False
                        return st_op in {"FINALIZADA", "FINALIZADO"} and prest != "FECHADA"
                    return False
            except Exception:
                logging.debug("Falha ao validar rota apta via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}

                has_status_op = "status_operacional" in cols_prog
                has_finalizada_app = "finalizada_no_app" in cols_prog
                has_prest = "prestacao_status" in cols_prog
                status_op_expr = "COALESCE(status_operacional,'')" if has_status_op else "''"
                finalizada_app_expr = "COALESCE(finalizada_no_app,0)" if has_finalizada_app else "0"
                prest_expr = "COALESCE(prestacao_status,'PENDENTE')" if has_prest else "'PENDENTE'"
                cur.execute(
                    f"""
                    SELECT {status_op_expr} AS status_operacional,
                           {finalizada_app_expr} AS finalizada_no_app,
                           COALESCE(status,'') AS status,
                           {prest_expr} AS prestacao_status
                    FROM programacoes
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (prog,),
                )
                row = cur.fetchone()
                if not row:
                    return False
                st_op = upper(row[0] if row[0] is not None else "")
                fin_app = safe_int(row[1], 0)
                st = upper(row[2] if row[2] is not None else "")
                prest = upper(row[3] if row[3] is not None else "PENDENTE")
        except Exception:
            return False

        # Mantido como indicador de conclusao da rota.
        # A programacao pode existir/vincular Recebimentos antes da finalizacao;
        # a finalizacao no app apenas consolida os dados operacionais.
        if st_op:
            if fin_app != 1:
                return False
            return st_op in {"FINALIZADA", "FINALIZADO"} and prest != "FECHADA"
        return False

    def _sync_status_finalizacao_from_api(self, max_rows: int = 80):
        """Mantido por compatibilidade: valida conectividade API sem persistencia local."""
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret or not is_desktop_api_sync_enabled():
            return
        try:
            _call_api(
                "GET",
                f"desktop/programacoes?modo=finalizadas_pendentes&limit={max(int(max_rows or 0), 1)}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
        except Exception:
            logging.debug("Falha ao consultar finalizacoes na API (recebimentos)", exc_info=True)

    def refresh_comboboxes(self):
        try:
            # Mesmo com sync global desligado, recebimentos precisa enxergar finalização vinda do app.
            self._sync_status_finalizacao_from_api(max_rows=80)
        except Exception:
            logging.debug("Falha ignorada")
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/programacoes?modo=todas&limit=300",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                valores = []
                for r in (arr or []):
                    if not isinstance(r, dict):
                        continue
                    codigo = upper((r or {}).get("codigo_programacao") or "")
                    status_ref = upper((r or {}).get("status_operacional") or (r or {}).get("status") or "")
                    if not codigo or status_ref in {"CANCELADA", "CANCELADO"}:
                        continue
                    valores.append(codigo)
                self.cb_prog["values"] = valores
                atual = upper(str(self.cb_prog.get() or "").strip())
                if atual and atual not in {upper(str(v or "").strip()) for v in valores}:
                    self.cb_prog.set("")
                return
            except Exception:
                logging.debug("Falha ao carregar combobox Recebimentos via API", exc_info=True)
        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                has_status_op = "status_operacional" in cols_prog
                has_finalizada_app = "finalizada_no_app" in cols_prog
                has_prest = "prestacao_status" in cols_prog

                if has_status_op:
                    status_where = "UPPER(TRIM(COALESCE(status_operacional,''))) NOT IN ('CANCELADA','CANCELADO')"
                else:
                    status_where = "UPPER(TRIM(COALESCE(status,''))) NOT IN ('CANCELADA','CANCELADO')"
                if has_finalizada_app:
                    status_where += " AND COALESCE(finalizada_no_app,0)=1"
                prest_where = " AND COALESCE(prestacao_status,'PENDENTE') IN ('PENDENTE','FECHADA')" if has_prest else ""

                cur.execute(
                    f"""
                    SELECT codigo_programacao
                    FROM programacoes
                    WHERE {status_where}{prest_where}
                    ORDER BY id DESC
                    LIMIT 300
                    """
                )
            except Exception:
                cur.execute("""
                    SELECT codigo_programacao
                    FROM programacoes
                    WHERE UPPER(TRIM(COALESCE(status,''))) NOT IN ('CANCELADA','CANCELADO')
                    ORDER BY id DESC
                    LIMIT 300
                """)
            valores = [r[0] for r in cur.fetchall()]
            self.cb_prog["values"] = valores
            atual = upper(str(self.cb_prog.get() or "").strip())
            if atual and atual not in {upper(str(v or "").strip()) for v in valores}:
                self.cb_prog.set("")

    def _sync_horarios_from_programacao(self, prog: str) -> bool:
        """Sincroniza campos de saida/chegada com o que veio do app mobile via API."""
        prog = upper(prog)
        if not prog:
            return False
        row = None
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(prog)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
                if isinstance(rota, dict):
                    row = (
                        str(rota.get("data_saida") or ""),
                        str(rota.get("hora_saida") or ""),
                        str(rota.get("data_chegada") or ""),
                        str(rota.get("hora_chegada") or ""),
                    )
            except Exception:
                logging.debug("Falha ao sincronizar horarios via API; usando fallback local.", exc_info=True)

        if row is None:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                data_saida_expr = "COALESCE(data_saida,'')" if "data_saida" in cols_prog else "''"
                hora_saida_expr = "COALESCE(hora_saida,'')" if "hora_saida" in cols_prog else "''"
                data_chegada_expr = "COALESCE(data_chegada,'')" if "data_chegada" in cols_prog else "''"
                hora_chegada_expr = "COALESCE(hora_chegada,'')" if "hora_chegada" in cols_prog else "''"
                cur.execute(
                    f"""
                    SELECT {data_saida_expr}, {hora_saida_expr}, {data_chegada_expr}, {hora_chegada_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (prog,),
                )
                row = cur.fetchone()

        if not row:
            return False

        data_saida_n, hora_saida_n = normalize_date_time_components(row[0], row[1])
        data_chegada_n, hora_chegada_n = normalize_date_time_components(row[2], row[3])

        # Nao apaga dados digitados no desktop; apenas preenche/atualiza quando existe valor no banco.
        changed = False
        if data_saida_n:
            if (self.ent_data_saida.get() or "").strip() != data_saida_n:
                self._safe_set_entry(self.ent_data_saida, data_saida_n, readonly_back=False)
                changed = True
        if hora_saida_n:
            if (self.ent_hora_saida.get() or "").strip() != hora_saida_n:
                self._safe_set_entry(self.ent_hora_saida, hora_saida_n, readonly_back=False)
                changed = True
        if data_chegada_n:
            if (self.ent_data_chegada.get() or "").strip() != data_chegada_n:
                self._safe_set_entry(self.ent_data_chegada, data_chegada_n, readonly_back=False)
                changed = True
        if hora_chegada_n:
            if (self.ent_hora_chegada.get() or "").strip() != hora_chegada_n:
                self._safe_set_entry(self.ent_hora_chegada, hora_chegada_n, readonly_back=False)
                changed = True

        if changed:
            self._refresh_diarias_preview()
        return changed

    def on_show(self):
        self.refresh_comboboxes()
        try:
            valores = [upper(str(v or "").strip()) for v in list(self.cb_prog["values"])]
        except Exception:
            valores = []
        if self._current_prog and upper(self._current_prog) not in valores:
            # Programacao saiu da lista (ex.: prestacao fechada/finalizada em Despesas).
            self._reset_view()
            return
        if self._current_prog:
            synced = False
            try:
                synced = self._sync_horarios_from_programacao(self._current_prog)
            except Exception:
                logging.debug("Falha ignorada")
            if synced:
                self.set_status(f"STATUS: Horarios de saida/chegada sincronizados do app para {self._current_prog}.")
                return
        self.set_status("STATUS: Selecione uma programacao FINALIZADA e lance os pagantes (prestacao de contas).")

    def _sync_programacao_from_api(self, prog: str, silent: bool = True) -> bool:
        """Valida disponibilidade da programacao na API central (sem persistencia local)."""
        prog = upper(prog)
        if not prog:
            return False

        try:
            resposta = self._api_bundle_prestacao(prog) if hasattr(self, "_api_bundle_prestacao") else None
            rota = resposta.get("rota") if isinstance(resposta, dict) else None
            return isinstance(rota, dict)
        except Exception:
            logging.debug("Falha ao sincronizar recebimentos/despesas pela API central", exc_info=True)
            if not silent:
                messagebox.showwarning(
                    "Sincronizacao",
                    "Nao foi possivel sincronizar agora com a API central.\n"
                    "A tela sera carregada com os dados locais.",
                )
            return False

    def _rehydrate_programacao_missing_on_server(self, prog: str) -> bool:
        prog = upper(str(prog or "").strip())
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not prog or not desktop_secret or not is_desktop_api_sync_enabled():
            return False

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacoes'")
                if not cur.fetchone():
                    return False

                cur.execute(
                    """
                    SELECT *
                    FROM programacoes
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (prog,),
                )
                row = cur.fetchone()
                if not row:
                    return False
                rota = dict(row) if hasattr(row, "keys") else {}

                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programacao_itens'")
                itens = []
                if cur.fetchone():
                    cur.execute(
                        """
                        SELECT *
                        FROM programacao_itens
                        WHERE codigo_programacao=?
                        ORDER BY id ASC
                        """,
                        (prog,),
                    )
                    for rr in (cur.fetchall() or []):
                        itens.append(dict(rr) if hasattr(rr, "keys") else {})

        except Exception:
            logging.debug("Falha ao ler programação local para reidratar no servidor.", exc_info=True)
            return False

        def _pick(src: dict, *keys, default=None):
            for key in keys:
                if key in src and src.get(key) not in (None, ""):
                    return src.get(key)
            return default

        try:
            itens_payload = []
            for it in (itens or []):
                cod_cliente = upper(_pick(it, "cod_cliente"))
                nome_cliente = upper(_pick(it, "nome_cliente"))
                if not cod_cliente or not nome_cliente:
                    continue
                itens_payload.append(
                    {
                        "cod_cliente": cod_cliente,
                        "nome_cliente": nome_cliente,
                        "qnt_caixas": safe_int(_pick(it, "qnt_caixas", "caixas_atual"), 0),
                        "kg": safe_float(_pick(it, "kg"), 0.0),
                        "preco": safe_float(_pick(it, "preco"), 0.0),
                        "endereco": upper(_pick(it, "endereco")),
                        "vendedor": upper(_pick(it, "vendedor")),
                        "pedido": upper(_pick(it, "pedido")),
                        "produto": upper(_pick(it, "produto")),
                        "obs": upper(_pick(it, "obs", "observacao")),
                    }
                )

            status_local = upper(str(_pick(rota, "status") or "").strip())
            status_operacional_local = upper(str(_pick(rota, "status_operacional") or "").strip())
            prestacao_local = upper(str(_pick(rota, "prestacao_status", default="PENDENTE") or "PENDENTE").strip())
            finalizada_local = safe_int(_pick(rota, "finalizada_no_app"), 0)
            status_exec = status_operacional_local or status_local
            if status_exec in {"FINALIZADA", "FINALIZADO"} and finalizada_local != 1:
                finalizada_local = 1

            payload_sync = {
                "codigo_programacao": prog,
                "data_criacao": str(_pick(rota, "data_criacao", "data") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                "motorista": upper(_pick(rota, "motorista")),
                "motorista_id": safe_int(_pick(rota, "motorista_id"), 0),
                "motorista_codigo": upper(_pick(rota, "motorista_codigo", "codigo_motorista")),
                "codigo_motorista": upper(_pick(rota, "codigo_motorista", "motorista_codigo")),
                "veiculo": upper(_pick(rota, "veiculo")),
                "equipe": upper(_pick(rota, "equipe")),
                "kg_estimado": safe_float(_pick(rota, "kg_estimado"), 0.0),
                "tipo_estimativa": upper(str(_pick(rota, "tipo_estimativa", default="KG") or "KG")),
                "caixas_estimado": safe_int(_pick(rota, "caixas_estimado"), 0),
                "status": status_local or "ATIVA",
                "local_rota": upper(_pick(rota, "local_rota", "tipo_rota", "local")),
                "tipo_rota": upper(_pick(rota, "tipo_rota", "local_rota", "local")),
                "local_carregamento": upper(_pick(rota, "local_carregamento", "local_carregado", "granja_carregada", "local_carreg")),
                "local_carregado": upper(_pick(rota, "local_carregado", "local_carregamento", "granja_carregada", "local_carreg")),
                "granja_carregada": upper(_pick(rota, "granja_carregada", "local_carregamento", "local_carregado", "local_carreg")),
                "local_carreg": upper(_pick(rota, "local_carreg", "local_carregamento", "local_carregado", "granja_carregada")),
                "adiantamento": safe_float(_pick(rota, "adiantamento", "adiantamento_rota"), 0.0),
                "total_caixas": safe_int(_pick(rota, "total_caixas", "nf_caixas", "caixas_carregadas"), 0),
                "quilos": safe_float(_pick(rota, "quilos", "nf_kg", "kg_carregado"), 0.0),
                "nf_kg": safe_float(_pick(rota, "nf_kg", "kg_nf"), 0.0),
                "nf_preco": safe_float(_pick(rota, "nf_preco", "preco_nf"), 0.0),
                "nf_caixas": safe_int(_pick(rota, "nf_caixas", "caixas_carregadas", "qnt_cx_carregada"), 0),
                "caixas_carregadas": safe_int(_pick(rota, "caixas_carregadas", "nf_caixas", "qnt_cx_carregada"), 0),
                "usuario_criacao": upper(_pick(rota, "usuario_criacao")),
                "usuario_ultima_edicao": upper(_pick(rota, "usuario_ultima_edicao", "usuario_criacao")),
                "linked_venda_ids": [],
                "vendas_usada_em": str(_pick(rota, "data_criacao", "data") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                "itens": itens_payload,
            }
            _call_api(
                "POST",
                "desktop/rotas/upsert",
                payload=payload_sync,
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            payload_status = {
                "status": status_local or None,
                "prestacao_status": prestacao_local or "PENDENTE",
                "status_operacional": status_operacional_local or None,
                "finalizada_no_app": int(finalizada_local or 0),
            }
            if any(v not in (None, "", 0) for v in payload_status.values()) or payload_status.get("finalizada_no_app") == 0:
                _call_api(
                    "PUT",
                    f"desktop/rotas/{urllib.parse.quote(prog)}/status",
                    payload=payload_status,
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )

            data_saida_n = normalize_date(_pick(rota, "data_saida"))
            hora_saida_n = normalize_time(_pick(rota, "hora_saida"))
            data_chegada_n = normalize_date(_pick(rota, "data_chegada"))
            hora_chegada_n = normalize_time(_pick(rota, "hora_chegada"))
            diaria_motorista = _pick(rota, "diaria_motorista_valor")
            if any([data_saida_n, hora_saida_n, data_chegada_n, hora_chegada_n, diaria_motorista not in (None, "")]):
                _call_api(
                    "PUT",
                    f"desktop/rotas/{urllib.parse.quote(prog)}/cabecalho",
                    payload={
                        "data_saida": data_saida_n or "",
                        "hora_saida": hora_saida_n or "",
                        "data_chegada": data_chegada_n or "",
                        "hora_chegada": hora_chegada_n or "",
                        "diaria_motorista_valor": safe_float(diaria_motorista, 0.0) if diaria_motorista not in (None, "") else None,
                    },
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
            return True
        except Exception:
            logging.debug("Falha ao reidratar programação ausente no servidor.", exc_info=True)
            return False

    # -------------------------
    # Regras: bloqueio quando FECHADA
    # -------------------------
    def _is_prestacao_fechada(self, prog: str) -> bool:
        if not prog:
            return False
        try:
            bundle = self._api_bundle_prestacao(prog) if hasattr(self, "_api_bundle_prestacao") else None
            rota = bundle.get("rota") if isinstance(bundle, dict) else None
            if isinstance(rota, dict):
                return upper(str(rota.get("prestacao_status") or "")) == "FECHADA"
        except Exception:
            logging.debug("Falha ao consultar prestacao via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT prestacao_status FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1", (prog,))
                r = cur.fetchone()
                st = upper(r[0]) if r and r[0] is not None else ""
                return st == "FECHADA"
        except Exception:
            return False

    def _warn_if_fechada(self) -> bool:
        if self._is_prestacao_fechada(self._current_prog):
            messagebox.showwarning("ATENCAO", "Esta prestacao ja esta FECHADA. Nao e possivel alterar recebimentos.")
            return True
        return False

    def _clean_cod_cliente(self, cod):
        s = str(cod or "").strip()
        if s.endswith(".0"):
            base = s[:-2]
            if base.isdigit():
                return base
        return s

    def _format_money_from_digits(self, digits: str) -> str:
        digits = re.sub(r"\D", "", str(digits or ""))
        if not digits:
            return "0,00"
        if len(digits) == 1:
            int_part = "0"
            cents = "0" + digits
        elif len(digits) == 2:
            int_part = "0"
            cents = digits
        else:
            int_part = digits[:-2]
            cents = digits[-2:]
        int_part = int_part.lstrip("0") or "0"

        parts = []
        while len(int_part) > 3:
            parts.insert(0, int_part[-3:])
            int_part = int_part[:-3]
        if int_part:
            parts.insert(0, int_part)
        int_part = ".".join(parts) if parts else "0"
        return f"{int_part},{cents}"

    def _bind_money_entry(self, ent: tk.Entry):
        def _on_focus_in(_e=None):
            try:
                v = str(ent.get() or "").strip()
                if v in {"0", "0,00", "0.00", "R$ 0,00", "R$0,00"}:
                    ent.delete(0, "end")
                else:
                    ent.selection_range(0, "end")
            except Exception:
                logging.debug("Falha ignorada")

        def _on_key_release(_e=None):
            try:
                v = str(ent.get() or "")
                digits = re.sub(r"\D", "", v)
                if digits:
                    ent.delete(0, "end")
                    ent.insert(0, self._format_money_from_digits(digits))
                    ent.icursor("end")
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out(_e=None):
            try:
                v = str(ent.get() or "").strip()
                if not v:
                    ent.insert(0, "0,00")
                else:
                    ent.delete(0, "end")
                    ent.insert(0, self._format_money_from_digits(v))
            except Exception:
                logging.debug("Falha ignorada")

        ent.bind("<FocusIn>", _on_focus_in)
        ent.bind("<KeyRelease>", _on_key_release)
        ent.bind("<FocusOut>", _on_focus_out)

    def _bind_date_entry(self, ent: tk.Entry):
        digits_state = {"value": ""}

        def _is_partial_date_digits_valid(digits: str) -> bool:
            n = len(digits)
            if n == 0:
                return True
            if n >= 1 and int(digits[0]) > 3:
                return False
            if n >= 2:
                dd = int(digits[:2])
                if dd < 1 or dd > 31:
                    return False
            if n >= 3 and int(digits[2]) > 1:
                return False
            if n >= 4:
                mm = int(digits[2:4])
                if mm < 1 or mm > 12:
                    return False
            if n in (6, 8):
                try:
                    dd = int(digits[:2])
                    mm = int(digits[2:4])
                    yy = int(digits[4:]) if n == 6 else int(digits[4:8])
                    if n == 6:
                        yy = (2000 + yy) if yy <= 69 else (1900 + yy)
                    datetime(yy, mm, dd)
                except Exception:
                    return False
            return True

        def _to_masked(digits: str) -> str:
            if len(digits) <= 2:
                return digits
            if len(digits) <= 4:
                return f"{digits[:2]}/{digits[2:]}"
            return f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"

        def _apply_mask():
            try:
                raw = str(ent.get() or "")
                digits = re.sub(r"\D", "", raw)[:8]

                # Bloqueia valores inválidos removendo o último dÃÂÂgito inválido.
                while digits and (not _is_partial_date_digits_valid(digits)):
                    digits = digits[:-1]

                masked = _to_masked(digits)
                ent.delete(0, "end")
                ent.insert(0, masked)
                ent.icursor("end")
                digits_state["value"] = digits
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out():
            try:
                digits = digits_state["value"]
                if digits and len(digits) not in (6, 8):
                    ent.delete(0, "end")
                    digits_state["value"] = ""
                    return
                nd = normalize_date(ent.get())
                if nd:
                    ent.delete(0, "end")
                    ent.insert(0, format_date_br_short(nd))
            except Exception:
                logging.debug("Falha ignorada")

        ent.bind("<KeyRelease>", lambda _e: _apply_mask())
        ent.bind("<FocusOut>", lambda _e: (_apply_mask(), _on_focus_out()))

    def _bind_time_entry(self, ent: tk.Entry):
        digits_state = {"value": ""}

        def _is_partial_time_digits_valid(digits: str) -> bool:
            n = len(digits)
            if n == 0:
                return True
            if n >= 1 and int(digits[0]) > 2:
                return False
            if n >= 2:
                hh = int(digits[:2])
                if hh < 0 or hh > 23:
                    return False
            if n >= 3 and int(digits[2]) > 5:
                return False
            if n >= 4:
                mm = int(digits[2:4])
                if mm < 0 or mm > 59:
                    return False
            if n >= 5 and int(digits[4]) > 5:
                return False
            if n == 6:
                ss = int(digits[4:6])
                if ss < 0 or ss > 59:
                    return False
            return True

        def _apply_mask():
            try:
                raw = str(ent.get() or "")
                digits = re.sub(r"\D", "", raw)[:6]

                # Bloqueia valores inválidos removendo o último dÃÂÂgito inválido.
                while digits and (not _is_partial_time_digits_valid(digits)):
                    digits = digits[:-1]

                if len(digits) <= 2:
                    masked = digits
                elif len(digits) <= 4:
                    masked = f"{digits[:2]}:{digits[2:]}"
                else:
                    masked = f"{digits[:2]}:{digits[2:4]}:{digits[4:]}"
                ent.delete(0, "end")
                ent.insert(0, masked)
                ent.icursor("end")
                digits_state["value"] = digits
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out():
            try:
                digits = digits_state["value"]
                if digits and len(digits) not in (4, 6):
                    ent.delete(0, "end")
                    digits_state["value"] = ""
                    return
                nt = normalize_time(ent.get())
                ent.delete(0, "end")
                if nt:
                    ent.insert(0, nt)
            except Exception:
                logging.debug("Falha ignorada")

        ent.bind("<KeyRelease>", lambda _e: _apply_mask())
        ent.bind("<FocusOut>", lambda _e: (_apply_mask(), _on_focus_out()))

    def _bind_diarias_preview(self):
        for ent in (
            self.ent_diaria_motorista,
            self.ent_data_saida,
            self.ent_hora_saida,
            self.ent_data_chegada,
            self.ent_hora_chegada,
        ):
            ent.bind("<KeyRelease>", lambda _e: self._refresh_diarias_preview(), add="+")
            ent.bind("<FocusOut>", lambda _e: self._refresh_diarias_preview(), add="+")

    def _bind_km_auto_calc(self):
        for ent in (
            getattr(self, "ent_km_inicial", None),
            getattr(self, "ent_km_final", None),
            getattr(self, "ent_litros", None),
        ):
            if not ent:
                continue
            ent.bind("<KeyRelease>", lambda _e: self._calcular_km_media(), add="+")
            ent.bind("<FocusOut>", lambda _e: self._calcular_km_media(), add="+")

    def _refresh_diarias_preview(self):
        try:
            data_saida = normalize_date(self.ent_data_saida.get())
            hora_saida = normalize_time(self.ent_hora_saida.get())
            data_chegada = normalize_date(self.ent_data_chegada.get())
            hora_chegada = normalize_time(self.ent_hora_chegada.get())
            diaria_motorista = safe_money(self.ent_diaria_motorista.get(), 0.0)

            if (
                data_saida in (None, "")
                or hora_saida in (None, "")
                or data_chegada in (None, "")
                or hora_chegada in (None, "")
            ):
                qtd = 0.0
            else:
                qtd = self._calc_qtd_diarias_regra(data_saida, hora_saida, data_chegada, hora_chegada)

            qtd_ajudantes = self._count_ajudantes(self._equipe_raw)
            diaria_ajudante = max(diaria_motorista - 10.0, 0.0)
            total_mot = round(qtd * diaria_motorista, 2)
            total_eqp = round(qtd * diaria_ajudante * qtd_ajudantes, 2)
            total = round(total_mot + total_eqp, 2)

            if hasattr(self, "lbl_resumo_diaria"):
                self.lbl_resumo_diaria.config(text=f"BASE MOTORISTA: {fmt_money(diaria_motorista)}")
            self.lbl_resumo_qtd.config(text=f"QTD DIARIAS: {qtd:g} | AJUDANTES: {qtd_ajudantes}")
            self.lbl_resumo_mot.config(text=f"MOTORISTA: {qtd:g} x {fmt_money(diaria_motorista)} = {fmt_money(total_mot)}")
            self.lbl_resumo_eqp.config(
                text=f"AJUDANTES: {qtd_ajudantes} x {fmt_money(diaria_ajudante)} = {fmt_money(total_eqp)}"
            )
            self.lbl_resumo_total.config(text=f"TOTAL DIARIAS: {fmt_money(total)}")
            if hasattr(self, "lbl_resumo_pagar"):
                self.lbl_resumo_pagar.config(text=f"VALOR A PAGAR: {fmt_money(total)}")
        except Exception:
            logging.debug("Falha ignorada")

    def _abrir_despesas_com_programacao(self):
        if not self._current_prog:
            messagebox.showwarning("ATENÇÃO", "Carregue uma programação primeiro.")
            return False
        prog = self._current_prog
        # Tenta persistir o cabeçalho sem bloquear a navegação.
        try:
            self.salvar_dados_rota(silent=True)
        except Exception:
            logging.debug("Falha ignorada")
        synced = False
        try:
            synced = self._sync_programacao_from_api(prog, silent=True)
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self.app.show_page("Despesas")
            page = self.app.pages.get("Despesas") if hasattr(self.app, "pages") else None
            if page and hasattr(page, "set_programacao"):
                page.set_programacao(prog)
                try:
                    if hasattr(page, "_load_by_programacao"):
                        page._load_by_programacao()
                except Exception:
                    logging.debug("Falha ignorada")
            if synced:
                self.set_status(f"STATUS: Indo para DESPESAS - programação {prog} (dados sincronizados da API).")
            else:
                self.set_status(f"STATUS: Indo para DESPESAS - programação {prog}.")
            return True
        except Exception:
            logging.debug("Falha ignorada")
            return False

    def _ir_para_despesas(self):
        self._abrir_despesas_com_programacao()

    # -------------------------
    # Carregar programação
    # -------------------------
    def carregar_programacao(self):
        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Selecione uma programação pendente.")
            return

        rota_finalizada = self._rota_apt_para_recebimentos(prog)

        synced_api = False
        try:
            synced_api = self._sync_programacao_from_api_desktop(prog, silent=True)
            if not synced_api:
                synced_api = self._sync_programacao_from_api(prog, silent=True)
        except Exception:
            logging.debug("Falha ignorada")

        # Revalida após sincronizar com API para impedir abertura indevida.
        rota_finalizada = self._rota_apt_para_recebimentos(prog)

        self._current_prog = prog
        motorista = veiculo = equipe = nf = ""
        data_saida = hora_saida = data_chegada = hora_chegada = ""
        rota = ""
        diaria_motorista = 0.0
        found_prog = False

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
        api_failed = False
        api_checked = False
        if api_enabled:
            try:
                api_checked = True
                resp = self._api_bundle_prestacao(prog) if hasattr(self, "_api_bundle_prestacao") else None
                rota_obj = resp.get("rota") if isinstance(resp, dict) else None
                if isinstance(rota_obj, dict):
                    motorista = str(rota_obj.get("motorista") or "")
                    veiculo = str(rota_obj.get("veiculo") or "")
                    equipe = str(rota_obj.get("equipe") or "")
                    nf = str(rota_obj.get("num_nf") or rota_obj.get("nf_numero") or "")
                    data_saida = str(rota_obj.get("data_saida") or "")
                    hora_saida = str(rota_obj.get("hora_saida") or "")
                    data_chegada = str(rota_obj.get("data_chegada") or "")
                    hora_chegada = str(rota_obj.get("hora_chegada") or "")
                    rota = str(rota_obj.get("local_rota") or rota_obj.get("tipo_rota") or rota_obj.get("local") or "")
                    diaria_motorista = safe_float(rota_obj.get("diaria_motorista_valor"), 0.0)
                    found_prog = True
            except Exception:
                api_failed = True
                logging.debug("Falha ao carregar programacao em recebimentos via API; usando fallback local.", exc_info=True)

        if api_enabled and api_checked and (not api_failed) and (not found_prog):
            recovered = False
            try:
                recovered = self._rehydrate_programacao_missing_on_server(prog)
            except Exception:
                logging.debug("Falha ao tentar reidratar programação ausente no servidor.", exc_info=True)
            if recovered:
                try:
                    resp = self._api_bundle_prestacao(prog) if hasattr(self, "_api_bundle_prestacao") else None
                    rota_obj = resp.get("rota") if isinstance(resp, dict) else None
                    if isinstance(rota_obj, dict):
                        motorista = str(rota_obj.get("motorista") or "")
                        veiculo = str(rota_obj.get("veiculo") or "")
                        equipe = str(rota_obj.get("equipe") or "")
                        nf = str(rota_obj.get("num_nf") or rota_obj.get("nf_numero") or "")
                        data_saida = str(rota_obj.get("data_saida") or "")
                        hora_saida = str(rota_obj.get("hora_saida") or "")
                        data_chegada = str(rota_obj.get("data_chegada") or "")
                        hora_chegada = str(rota_obj.get("hora_chegada") or "")
                        rota = str(rota_obj.get("local_rota") or rota_obj.get("tipo_rota") or rota_obj.get("local") or "")
                        diaria_motorista = safe_float(rota_obj.get("diaria_motorista_valor"), 0.0)
                        found_prog = True
                        self.set_status(f"STATUS: Programação {prog} reidratada no servidor para Recebimentos.")
                except Exception:
                    logging.debug("Falha ao recarregar programação após reidratar no servidor.", exc_info=True)
            if not found_prog:
                messagebox.showwarning("ATENÇÃO", f"Programação não encontrada no servidor: {prog}")
                self._reset_view()
                return

        if not found_prog:
            with get_db() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols_prog = [str(r[1]).lower() for r in cur.fetchall()]
                except Exception:
                    cols_prog = []
                col_rota = (
                    "local_rota"
                    if "local_rota" in cols_prog
                    else ("tipo_rota" if "tipo_rota" in cols_prog else ("local" if "local" in cols_prog else "''"))
                )
                col_diaria = "diaria_motorista_valor" if "diaria_motorista_valor" in cols_prog else "0"
                col_nf = (
                    "COALESCE(NULLIF(num_nf,''), NULLIF(nf_numero,''), '')"
                    if ("num_nf" in cols_prog and "nf_numero" in cols_prog)
                    else ("num_nf" if "num_nf" in cols_prog else ("nf_numero" if "nf_numero" in cols_prog else "''"))
                )
                cur.execute("""
                    SELECT motorista, veiculo, equipe, {col_nf} as num_nf,
                           data_saida, hora_saida, data_chegada, hora_chegada,
                           {col_rota} as rota, COALESCE({col_diaria}, 0) as diaria_motorista
                    FROM programacoes
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 1
                """.format(col_nf=col_nf, col_rota=col_rota, col_diaria=col_diaria), (prog,))
                row = cur.fetchone()

            if not row:
                messagebox.showwarning("ATENÇÃO", f"Programação não encontrada: {prog}")
                self._reset_view()
                return
            motorista, veiculo, equipe, nf, data_saida, hora_saida, data_chegada, hora_chegada, rota, diaria_motorista = row

        motorista_nome = self._resolve_motorista_nome(motorista)
        equipe_nomes = self._resolve_equipe_integrantes(equipe)
        self._equipe_raw = upper(equipe or "")
        self.lbl_motorista_info.config(text=f"Motorista: {motorista_nome}")
        self.lbl_veiculo_info.config(text=f"Veiculo: {upper(veiculo)}")
        self.lbl_equipe_info.config(text=f"Equipe: {equipe_nomes}")
        self._rota_atual = upper(rota or "")
        self.lbl_rota_info.config(text=f"Rota: {self._rota_atual or '-'}")
        self._safe_set_entry(self.ent_diaria_motorista, f"{safe_float(diaria_motorista, 0.0):.2f}".replace(".", ","), readonly_back=False)

        self._nf_current = nf or ""

        data_saida_n, hora_saida_n = normalize_date_time_components(data_saida, hora_saida)
        data_chegada_n, hora_chegada_n = normalize_date_time_components(data_chegada, hora_chegada)

        self.ent_data_saida.delete(0, "end"); self.ent_data_saida.insert(0, data_saida_n)
        self.ent_hora_saida.delete(0, "end"); self.ent_hora_saida.insert(0, hora_saida_n)
        self.ent_data_chegada.delete(0, "end"); self.ent_data_chegada.insert(0, data_chegada_n)
        self.ent_hora_chegada.delete(0, "end"); self.ent_hora_chegada.insert(0, hora_chegada_n)
        self._refresh_diarias_preview()

        self._selected_cliente = None
        self._clear_form_recebimento()

        if self._is_collapsed_after_close:
            self._expand_view()

        self.carregar_clientes_e_recebimentos()

        if self._is_prestacao_fechada(prog):
            self.set_status(f"STATUS: Programacao {self._current_prog} (PRESTACAO FECHADA - somente consulta)")
        else:
            if rota_finalizada and synced_api:
                self.set_status(f"STATUS: Programacao finalizada e sincronizada da API: {prog}")
            elif rota_finalizada:
                self.set_status(f"STATUS: Programacao finalizada pronta para recebimentos: {prog}")
            elif synced_api:
                self.set_status(
                    f"STATUS: Programacao carregada em aberto: {prog}. "
                    "Recebimentos serao consolidados ao finalizar no app ou no encerramento manual."
                )
            else:
                self.set_status(
                    f"STATUS: Programacao carregada em aberto: {prog}. "
                    "Ela ja esta vinculada a Recebimentos."
                )

    # -------------------------
    # Validação leve de data/hora
    # -------------------------
    def _validate_date(self, s: str) -> bool:
        s = (s or "").strip()
        if not s:
            return True
        return normalize_date(s) is not None

    def _validate_time(self, s: str) -> bool:
        s = (s or "").strip()
        if not s:
            return True
        return normalize_time(s) is not None

    def salvar_dados_rota(self, silent: bool = False):
        if not self._current_prog:
            if not silent:
                messagebox.showwarning("ATENCAO", "Carregue uma programacao primeiro.")
            return False
        if not ensure_system_api_binding(context=f"Salvar cabecalho rota ({self._current_prog})", parent=self):
            return False
        if self._warn_if_fechada():
            return False

        diaria_motorista = safe_money(self.ent_diaria_motorista.get(), 0.0)
        if diaria_motorista < 0:
            if not silent:
                messagebox.showwarning("ATENCAO", "Diaria do motorista nao pode ser negativa.")
            return False

        data_saida = normalize_date(self.ent_data_saida.get())
        hora_saida = normalize_time(self.ent_hora_saida.get())
        data_chegada = normalize_date(self.ent_data_chegada.get())
        hora_chegada = normalize_time(self.ent_hora_chegada.get())

        if data_saida is None or data_chegada is None:
            if not silent:
                messagebox.showwarning("ATENCAO", "Formato de data invalido. Use DD/MM/AA.")
            return False
        if hora_saida is None or hora_chegada is None:
            if not silent:
                messagebox.showwarning("ATENCAO", "Formato de hora invalido. Use HH:MM:SS (ex.: 07:30:00).")
            return False

        # atualiza campos com formato normalizado
        self._safe_set_entry(self.ent_data_saida, format_date_br_short(data_saida), readonly_back=False)
        self._safe_set_entry(self.ent_hora_saida, hora_saida, readonly_back=False)
        self._safe_set_entry(self.ent_data_chegada, format_date_br_short(data_chegada), readonly_back=False)
        self._safe_set_entry(self.ent_hora_chegada, hora_chegada, readonly_back=False)
        self._safe_set_entry(self.ent_diaria_motorista, f"{diaria_motorista:.2f}".replace(".", ","), readonly_back=False)
        self._refresh_diarias_preview()

        try:
            resumo_diarias = self._sync_diarias_despesas(
                prog=self._current_prog,
                rota=self._rota_atual,
                equipe_raw=self._equipe_raw,
                data_saida=data_saida,
                hora_saida=hora_saida,
                data_chegada=data_chegada,
                hora_chegada=hora_chegada,
                diaria_motorista=diaria_motorista,
                persist_remote=False,
            )
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "PUT",
                f"desktop/rotas/{urllib.parse.quote(upper(self._current_prog))}/cabecalho",
                payload={
                    "data_saida": data_saida,
                    "hora_saida": hora_saida,
                    "data_chegada": data_chegada,
                    "hora_chegada": hora_chegada,
                    "diaria_motorista_valor": float(diaria_motorista),
                    "qtd_diarias": float(resumo_diarias.get("qtd_diarias", 0.0)),
                    "qtd_ajudantes": int(resumo_diarias.get("qtd_ajudantes", 0)),
                    "total_motorista": float(resumo_diarias.get("total_motorista", 0.0)),
                    "total_ajudantes": float(resumo_diarias.get("total_ajudantes", 0.0)),
                    "observacao_motorista": str(resumo_diarias.get("observacao_motorista") or ""),
                    "observacao_ajudantes": str(resumo_diarias.get("observacao_ajudantes") or ""),
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            if not silent:
                messagebox.showinfo("OK", "Dados da rota atualizados!")
            if resumo_diarias:
                self.set_status(
                    "STATUS: Diarias atualizadas - "
                    f"QTD: {resumo_diarias['qtd_diarias']:g} | "
                    f"Motorista: {fmt_money(resumo_diarias['total_motorista'])} | "
                    f"Equipe ({resumo_diarias['qtd_ajudantes']}): {fmt_money(resumo_diarias['total_ajudantes'])} | "
                    f"Total: {fmt_money(resumo_diarias['total_geral'])}"
                )
            else:
                self.set_status("STATUS: Dados do cabecalho atualizados.")
            return True
        except Exception as e:
            if not silent:
                messagebox.showerror(
                    "ERRO",
                    "Não foi possível salvar o cabeçalho da rota.\n\n"
                    f"{_friendly_sync_error(e)}",
                )
            return False
    

    def _parse_dt_diaria(self, data_s: str, hora_s: str):
        data_s = (data_s or "").strip()
        hora_s = (hora_s or "").strip()
        if not data_s:
            return None
        try:
            if "-" in data_s:
                y, m, d = data_s[:10].split("-")
            elif "/" in data_s:
                d, m, y = data_s[:10].split("/")
            else:
                return None
            hh, mm = "00", "00"
            if hora_s and ":" in hora_s:
                hh, mm = hora_s.split(":")[:2]
            y = int(y)
            if y < 100:
                y = (2000 + y) if y <= 69 else (1900 + y)
            return datetime(y, int(m), int(d), int(hh), int(mm))
        except Exception:
            return None

    def _calc_qtd_diarias_regra(self, data_saida: str, hora_saida: str, data_chegada: str, hora_chegada: str) -> float:
        dt_saida = self._parse_dt_diaria(data_saida, hora_saida)
        dt_chegada = self._parse_dt_diaria(data_chegada, hora_chegada)
        if not dt_saida:
            return 0.0
        if not dt_chegada or dt_chegada <= dt_saida:
            return 1.0
        horas = (dt_chegada - dt_saida).total_seconds() / 3600.0
        if horas <= 24.0:
            return 1.0
        rem = horas - 24.0
        full = int(rem // 24.0)
        half = 0.5 if (rem - (full * 24.0)) > 0 else 0.0
        return 1.0 + full + half

    def _count_ajudantes(self, equipe_raw: str) -> int:
        raw = str(equipe_raw or "").strip()
        if not raw:
            return 0

        # Preferencial: estrutura bruta salva em programacoes (ex.: ID1|ID2).
        parts_raw = [p.strip() for p in re.split(r"[|,;/]+", raw) if p.strip()]
        if len(parts_raw) >= 2:
            return len(parts_raw)

        # Fallback: nomes resolvidos da equipe.
        nomes = str(resolve_equipe_nomes(raw) or "").strip()
        if not nomes:
            return 2
        parts_nome = [p.strip() for p in re.split(r"[|,;/]+", nomes) if p.strip()]
        if len(parts_nome) >= 2:
            return len(parts_nome)
        return 2

    def _sync_diarias_despesas(self, prog: str, rota: str, equipe_raw: str, data_saida: str, hora_saida: str, data_chegada: str, hora_chegada: str, diaria_motorista: float, persist_remote: bool = True):
        qtd = self._calc_qtd_diarias_regra(data_saida, hora_saida, data_chegada, hora_chegada)
        diaria_motorista = safe_float(diaria_motorista, 0.0)
        qtd_ajudantes = self._count_ajudantes(equipe_raw)
        diaria_ajudante = max(diaria_motorista - 10.0, 0.0)
        total_mot = round(qtd * diaria_motorista, 2)
        total_ajud = round(qtd * (diaria_ajudante * qtd_ajudantes), 2)
        total_geral = round(total_mot + total_ajud, 2)

        motorista_nome = ""
        ajudantes_nome = ""
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        try:
            bundle = self._api_bundle_prestacao(prog) if hasattr(self, "_api_bundle_prestacao") else None
            rota_obj = bundle.get("rota") if isinstance(bundle, dict) else None
            if isinstance(rota_obj, dict):
                motorista_nome = self._resolve_motorista_nome(str(rota_obj.get("motorista") or ""))
                ajudantes_nome = self._resolve_equipe_integrantes(str(rota_obj.get("equipe") or equipe_raw or ""))
        except Exception:
            logging.debug("Falha ao carregar equipe/motorista via API para diarias; usando fallback local.", exc_info=True)
        if not motorista_nome and not ajudantes_nome:
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute(
                        "SELECT COALESCE(motorista,''), COALESCE(equipe,'') FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                        (prog,),
                    )
                    rr = cur.fetchone() or ("", "")
                    motorista_nome = self._resolve_motorista_nome(rr[0] or "")
                    ajudantes_nome = self._resolve_equipe_integrantes(rr[1] or equipe_raw or "")
            except Exception:
                motorista_nome = self._resolve_motorista_nome("")
                ajudantes_nome = self._resolve_equipe_integrantes(equipe_raw or "")
        if not motorista_nome:
            motorista_nome = "-"
        if not ajudantes_nome:
            ajudantes_nome = "-"

        obs_motorista = f"QTD DIARIAS: {qtd:g} | MOTORISTA: {motorista_nome}"
        obs_ajudantes = f"QTD DIARIAS: {qtd:g} | AJUDANTES: {ajudantes_nome}"
        resumo = {
            "qtd_diarias": qtd,
            "qtd_ajudantes": qtd_ajudantes,
            "diaria_motorista": diaria_motorista,
            "diaria_ajudante": diaria_ajudante,
            "total_motorista": total_mot,
            "total_ajudantes": total_ajud,
            "total_geral": total_geral,
            "observacao_motorista": obs_motorista,
            "observacao_ajudantes": obs_ajudantes,
        }
        if persist_remote:
            _call_api(
                "POST",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/diarias/sync",
                payload={
                    "qtd_diarias": float(qtd),
                    "qtd_ajudantes": int(qtd_ajudantes),
                    "total_motorista": float(total_mot),
                    "total_ajudantes": float(total_ajud),
                    "observacao_motorista": obs_motorista,
                    "observacao_ajudantes": obs_ajudantes,
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

        return resumo
    

# -------------------------
    # Carregar clientes + recebimentos (resumo)
    # -------------------------
    def carregar_clientes_e_recebimentos(self):
        prog = self._current_prog
        if not prog:
            return

        filtro_forma = ""
        filtro_valor_min = 0.0
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
        api_loaded = False

        if api_enabled:
            try:
                bundle_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/bundle",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = bundle_resp.get("rota") if isinstance(bundle_resp, dict) else None
                clientes_api = bundle_resp.get("clientes") if isinstance(bundle_resp, dict) else []
                rec_items = bundle_resp.get("recebimentos") if isinstance(bundle_resp, dict) else []
                if isinstance(rota, dict):
                    clientes = [
                        (
                            str((r or {}).get("cod_cliente") or ""),
                            str((r or {}).get("nome_cliente") or ""),
                        )
                        for r in (clientes_api or [])
                        if isinstance(r, dict)
                    ]
                    ctrl_map = {}
                    for r in (clientes_api or []):
                        if not isinstance(r, dict):
                            continue
                        cod_u = upper(self._clean_cod_cliente((r or {}).get("cod_cliente")))
                        vrec = safe_float((r or {}).get("valor_recebido"), safe_float((r or {}).get("recebido_valor"), 0.0))
                        if not cod_u or vrec <= 0:
                            continue
                        ctrl_map[cod_u] = {
                            "valor": vrec,
                            "forma": upper((r or {}).get("forma_recebimento") or (r or {}).get("recebido_forma") or ""),
                            "obs": str((r or {}).get("obs_recebimento") or (r or {}).get("recebido_obs") or ""),
                            "data": str((r or {}).get("alterado_em") or (r or {}).get("updated_at") or ""),
                        }

                    recs = []
                    for r in (rec_items or []):
                        if not isinstance(r, dict):
                            continue
                        cod = str((r or {}).get("cod_cliente") or "").strip()
                        nome = str((r or {}).get("nome_cliente") or "").strip()
                        valor = safe_float((r or {}).get("valor"), 0.0)
                        forma = str((r or {}).get("forma_pagamento") or "").strip()
                        obs = str((r or {}).get("observacao") or "").strip()
                        data_registro = str((r or {}).get("data_registro") or "").strip()
                        if filtro_forma and filtro_forma != "TODAS" and upper(forma) != upper(filtro_forma):
                            continue
                        if filtro_valor_min > 0 and valor < filtro_valor_min:
                            continue
                        recs.append((cod, nome, valor, forma, obs, data_registro))
                    api_loaded = True
            except Exception:
                logging.debug("Falha ao carregar clientes/recebimentos via API; usando fallback local.", exc_info=True)

        if not api_loaded:
            with get_db() as conn:
                cur = conn.cursor()

                cur.execute("""
                    SELECT DISTINCT cod_cliente, nome_cliente
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                    ORDER BY nome_cliente ASC
                """, (prog,))
                clientes = cur.fetchall() or []

                # dados sincronizados do app (controle)
                ctrl_map = {}
                try:
                    cur.execute("PRAGMA table_info(programacao_itens_controle)")
                    cols_ctrl = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    col_valor = "valor_recebido" if "valor_recebido" in cols_ctrl else ("recebido_valor" if "recebido_valor" in cols_ctrl else "0")
                    col_forma = "forma_recebimento" if "forma_recebimento" in cols_ctrl else ("recebido_forma" if "recebido_forma" in cols_ctrl else "''")
                    col_obs = "obs_recebimento" if "obs_recebimento" in cols_ctrl else ("recebido_obs" if "recebido_obs" in cols_ctrl else "''")
                    col_data = "alterado_em" if "alterado_em" in cols_ctrl else ("updated_at" if "updated_at" in cols_ctrl else "''")
                    cur.execute("""
                        SELECT cod_cliente,
                               COALESCE({col_valor}, 0),
                               COALESCE({col_forma}, ''),
                               COALESCE({col_obs}, ''),
                               COALESCE({col_data}, '')
                        FROM programacao_itens_controle
                        WHERE codigo_programacao=?
                    """.format(
                        col_valor=col_valor,
                        col_forma=col_forma,
                        col_obs=col_obs,
                        col_data=col_data,
                    ), (prog,))
                    for cod, vrec, forma, obs, alterado_em in (cur.fetchall() or []):
                        cod_u = upper(self._clean_cod_cliente(cod))
                        ctrl_map[cod_u] = {
                            "valor": safe_float(vrec, 0.0),
                            "forma": upper(forma),
                            "obs": obs or "",
                            "data": alterado_em or ""
                        }
                except Exception:
                    ctrl_map = {}

                recs = []
                if api_enabled:
                    # Se API estava habilitada e houve falha técnica, usa o local como contingência.
                    query = """
                        SELECT cod_cliente, nome_cliente, valor, forma_pagamento, observacao, data_registro
                        FROM recebimentos
                        WHERE codigo_programacao=?
                    """
                    params = [prog]

                    if filtro_forma and filtro_forma != "TODAS":
                        query += " AND forma_pagamento = ?"
                        params.append(filtro_forma)

                    if filtro_valor_min > 0:
                        query += " AND valor >= ?"
                        params.append(filtro_valor_min)

                    query += " ORDER BY id DESC"
                    cur.execute(query, params)
                    recs = cur.fetchall() or []
                else:
                    query = """
                        SELECT cod_cliente, nome_cliente, valor, forma_pagamento, observacao, data_registro
                        FROM recebimentos
                        WHERE codigo_programacao=?
                    """
                    params = [prog]

                    if filtro_forma and filtro_forma != "TODAS":
                        query += " AND forma_pagamento = ?"
                        params.append(filtro_forma)

                    if filtro_valor_min > 0:
                        query += " AND valor >= ?"
                        params.append(filtro_valor_min)

                    query += " ORDER BY id DESC"
                    cur.execute(query, params)
                    recs = cur.fetchall() or []

        mapa = {}
        for cod, nome, valor, forma, obs, data_registro in recs:
            cod_u = upper(self._clean_cod_cliente(cod))
            if cod_u not in mapa:
                mapa[cod_u] = {
                    "nome": upper(nome),
                    "valor": 0.0,
                    "forma": upper(forma or ""),
                    "obs": (obs or ""),
                    "ultima_data": (data_registro or "")
                }
            mapa[cod_u]["valor"] += safe_float(valor, 0.0)
            if forma:
                mapa[cod_u]["forma"] = upper(forma)
            if obs:
                mapa[cod_u]["obs"] = obs
            if data_registro:
                mapa[cod_u]["ultima_data"] = data_registro

        self.tree.delete(*self.tree.get_children())

        total = 0.0
        for cod, nome in clientes:
            cod_u = upper(self._clean_cod_cliente(cod))
            nome_u = upper(nome)
            info = mapa.get(cod_u, {"valor": 0.0, "forma": "", "obs": "", "ultima_data": ""})
            ctrl = ctrl_map.get(cod_u)
            if ctrl and ctrl.get("valor", 0.0) > 0:
                info = {
                    "valor": ctrl.get("valor", 0.0),
                    "forma": ctrl.get("forma", ""),
                    "obs": ctrl.get("obs", ""),
                    "ultima_data": ctrl.get("data", "")
                }

            total += safe_float(info["valor"], 0.0)
            tree_insert_aligned(
                self.tree,
                "",
                "end",
                (
                    cod_u,
                    nome_u,
                    fmt_money(info["valor"]),
                    upper(info["forma"]),
                    info["obs"],
                    info["ultima_data"][:19] if info["ultima_data"] else "",
                ),
            )

        self._update_total(total)
        self.set_status(f"STATUS: {len(clientes)} clientes carregados. Total: {fmt_money(total)}")

    def carregar_tabela_vazia(self):
        self.tree.delete(*self.tree.get_children())
        self._update_total(0.0)

    def _update_total(self, total):
        self.lbl_total.config(text=f"TOTAL RECEBIDO: {fmt_money(total)}")

    # -------------------------
    # Seleção / formulário
    # -------------------------
    def _on_select_row(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values") or ()
        if len(vals) < 6:
            return

        cod, nome, valor, forma, obs, data_registro = vals
        cod = self._clean_cod_cliente(cod)
        self._selected_cliente = upper(cod)

        self._safe_set_entry(self.ent_cod, cod, readonly_back=True)
        self._safe_set_entry(self.ent_nome, nome, readonly_back=True)

        try:
            vv = str(valor).replace("R$", "").strip()
        except Exception:
            vv = ""
        self.ent_valor.delete(0, "end")
        self.ent_valor.insert(0, vv)

        forma_u = upper(forma)
        try:
            allowed = set(self.cb_forma["values"])
        except Exception:
            allowed = {"DINHEIRO", "PIX", "CARTAO", "BOLETO", "OUTRO"}

        self.cb_forma.set(forma_u if forma_u in allowed else "DINHEIRO")

        self.ent_obs.delete(0, "end")
        self.ent_obs.insert(0, obs or "")

    def _on_tree_double_click_value(self, event=None):
        try:
            if self._warn_if_fechada():
                return
            if not self._current_prog:
                return

            region = self.tree.identify("region", event.x, event.y)
            if region != "cell":
                return

            row_id = self.tree.identify_row(event.y)
            col_id = self.tree.identify_column(event.x)
            if not row_id or not col_id:
                return

            cols = list(self.tree["columns"] or [])
            try:
                col_idx = int(str(col_id).replace("#", "")) - 1
            except Exception:
                return
            if col_idx < 0 or col_idx >= len(cols):
                return

            col_name = cols[col_idx]
            if col_name != "VALOR":
                return

            vals = self.tree.item(row_id, "values") or ()
            if len(vals) < 6:
                return

            cod, nome, valor_atual, forma_atual, obs_atual, _data = vals
            cod = upper(self._clean_cod_cliente(cod))
            nome = upper(nome)
            if not cod or not nome:
                return

            valor_sugerido = self._format_money_from_digits(str(valor_atual or ""))
            novo_valor_txt = simple_input(
                "Recebimento",
                f"Cliente: {nome}\nInforme o valor recebido:",
                initial=valor_sugerido,
                master=self.app if hasattr(self, "app") else None,
                allow_empty=False,
            )
            if novo_valor_txt is None or str(novo_valor_txt).strip() == "":
                return

            valor = safe_money(novo_valor_txt, 0.0)
            if valor <= 0:
                messagebox.showwarning("ATENÇÃO", "Informe um valor válido (maior que zero).")
                return

            forma = upper(forma_atual or self.cb_forma.get() or "DINHEIRO")
            obs = upper(obs_atual or "")
            self._salvar_recebimento_item(cod=cod, nome=nome, valor=valor, forma=forma, obs=obs)
            self.carregar_clientes_e_recebimentos()
            self.set_status(f"STATUS: Recebimento de {fmt_money(valor)} salvo para {nome}.")
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao lançar recebimento na tabela: {str(e)}")

    def _clear_form_recebimento(self):
        self._safe_set_entry(self.ent_cod, "", readonly_back=True)
        self._safe_set_entry(self.ent_nome, "", readonly_back=True)
        self.ent_valor.delete(0, "end")
        self.cb_forma.set("DINHEIRO")
        self.ent_obs.delete(0, "end")

    def _salvar_recebimento_item(self, cod: str, nome: str, valor: float, forma: str, obs: str):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Carregue uma programação primeiro.")
            return
        if not ensure_system_api_binding(context=f"Salvar recebimento ({prog})", parent=self):
            return
        if self._warn_if_fechada():
            return
        if not cod or not nome:
            messagebox.showwarning("ATENÇÃO", "Selecione um cliente na tabela ou insira manualmente.")
            return
        if safe_float(valor, 0.0) <= 0:
            messagebox.showwarning("ATENÇÃO", "Informe um valor válido (maior que zero).")
            return
        forma = upper(str(forma or "").strip())
        formas_validas = {"DINHEIRO", "PIX", "CARTAO", "BOLETO", "OUTRO"}
        if not forma:
            messagebox.showwarning("ATENÇÃO", "Informe a forma de pagamento.")
            return
        if forma not in formas_validas:
            messagebox.showwarning("ATENÇÃO", "Forma de pagamento inválida.")
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        payload = {
            "cod_cliente": upper(cod),
            "nome_cliente": upper(nome),
            "valor": float(valor),
            "forma_pagamento": upper(forma),
            "observacao": upper(obs),
            "num_nf": upper(self._nf_current),
        }
        _call_api(
            "POST",
            f"desktop/rotas/{urllib.parse.quote(upper(prog))}/recebimentos",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )

    # -------------------------
    # Salvar recebimento (com bloqueio se FECHADA)
    # -------------------------
    def salvar_recebimento(self):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Carregue uma programação primeiro.")
            return
        if self._warn_if_fechada():
            return

        cod = upper(self.ent_cod.get())
        nome = upper(self.ent_nome.get())
        valor = safe_money(self.ent_valor.get(), 0.0)
        forma = upper(self.cb_forma.get())
        obs = upper(self.ent_obs.get())

        if not cod or not nome:
            messagebox.showwarning("ATENÇÃO", "Selecione um cliente na tabela ou insira manualmente.")
            return
        if valor <= 0:
            messagebox.showwarning("ATENÇÃO", "Informe um valor válido (maior que zero).")
            return
        if not forma:
            messagebox.showwarning("ATENÇÃO", "Informe a forma de pagamento.")
            return

        try:
            self._salvar_recebimento_item(cod=cod, nome=nome, valor=valor, forma=forma, obs=obs)
            self._clear_form_recebimento()
            self.carregar_clientes_e_recebimentos()
            self.set_status(f"STATUS: Recebimento de {fmt_money(valor)} salvo para {nome}.")
        except Exception as e:
            messagebox.showerror(
                "ERRO",
                "Não foi possível salvar o recebimento.\n\n"
                f"{_friendly_sync_error(e)}",
            )

    def zerar_recebimento(self):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Carregue uma programação primeiro.")
            return
        if not ensure_system_api_binding(context=f"Zerar recebimentos ({prog})", parent=self):
            return
        if self._warn_if_fechada():
            return

        cod = upper(self.ent_cod.get().strip())
        if not cod:
            messagebox.showwarning("ATENÇÃO", "Selecione um cliente na tabela.")
            return

        if not messagebox.askyesno("Confirmar", f"Zerar recebimentos do cliente {cod} nessa programação?"):
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "DELETE",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/recebimentos/{urllib.parse.quote(upper(cod))}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            self._clear_form_recebimento()
            self.carregar_clientes_e_recebimentos()
            self.set_status(f"STATUS: Recebimentos zerados para cliente {cod}.")
        except Exception as e:
            messagebox.showerror(
                "ERRO",
                "Não foi possível zerar os recebimentos do cliente.\n\n"
                f"{_friendly_sync_error(e)}",
            )

    def inserir_cliente_manual(self):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENCAO", "Carregue uma programacao primeiro.")
            return
        if not ensure_system_api_binding(context=f"Inserir cliente manual ({prog})", parent=self):
            return
        if self._warn_if_fechada():
            return

        cod = upper(simple_input(
            "Cliente Manual", "Digite o CODIGO do cliente:",
            master=self.app if hasattr(self, "app") else None,
            allow_empty=False
        ))
        if not cod:
            return

        nome = upper(simple_input(
            "Cliente Manual", "Digite o NOME do cliente:",
            master=self.app if hasattr(self, "app") else None,
            allow_empty=False
        ))
        if not nome:
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "POST",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/clientes/manual",
                payload={
                    "cod_cliente": cod,
                    "nome_cliente": nome,
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            self._safe_set_entry(self.ent_cod, cod, readonly_back=True)
            self._safe_set_entry(self.ent_nome, nome, readonly_back=True)
            self._selected_cliente = {"cod_cliente": cod, "nome_cliente": nome}

            messagebox.showinfo("OK", "Cliente inserido/atualizado na programacao (manual). Agora pode lancar o recebimento.")
            self.carregar_clientes_e_recebimentos()

            try:
                self.ent_valor.focus_set()
                self.ent_valor.selection_range(0, "end")
            except Exception:
                logging.debug("Falha ignorada")

        except Exception as e:
            messagebox.showerror(
                "ERRO",
                "Não foi possível inserir o cliente manual na programação.\n\n"
                f"{_friendly_sync_error(e)}",
            )
    

# =========================
    # âÅâââ€š¬Åâ€œââââ‚¬Å¡¬¦ IMPRESSÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢O PDF (A4) âââââ€š¬Å¡¬ââââ‚¬Å¡¬ alinhado e com assinaturas
    # =========================
    def _get_dados_para_impressao(self, prog: str):
        """
        Retorna:
          - header dict (motorista/equipe/veiculo/nf/saida/chegada)
          - rows: lista de (cod, nome, valor_total, forma_mais_recente)
          - total_geral
        Apenas clientes pagantes (valor_total > 0).
        """
        header = {
            "prog": prog,
            "motorista_nome": "",
            "equipe_nomes": "",
            "veiculo": "",
            "nf": "",
            "data_saida": "",
            "hora_saida": "",
            "data_chegada": "",
            "hora_chegada": "",
        }

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                bundle_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/bundle",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = bundle_resp.get("rota") if isinstance(bundle_resp, dict) else None
                recebimentos = bundle_resp.get("recebimentos") if isinstance(bundle_resp, dict) else []
                if isinstance(rota, dict):
                    motorista_raw = str(rota.get("motorista") or "")
                    equipe_raw = str(rota.get("equipe") or "")
                    header["motorista_nome"] = self._resolve_motorista_nome(motorista_raw)
                    header["equipe_nomes"] = self._resolve_equipe_integrantes(equipe_raw)
                    header["veiculo"] = upper(str(rota.get("veiculo") or ""))
                    header["nf"] = upper(str(rota.get("num_nf") or rota.get("nf_numero") or ""))
                    dsa_n, hsa_n = normalize_date_time_components(
                        str(rota.get("data_saida") or ""),
                        str(rota.get("hora_saida") or ""),
                    )
                    dch_n, hch_n = normalize_date_time_components(
                        str(rota.get("data_chegada") or ""),
                        str(rota.get("hora_chegada") or ""),
                    )
                    header["data_saida"] = dsa_n
                    header["hora_saida"] = hsa_n
                    header["data_chegada"] = dch_n
                    header["hora_chegada"] = hch_n

                agg = {}
                for rr in (recebimentos or []):
                    if not isinstance(rr, dict):
                        continue
                    cod = upper(str(rr.get("cod_cliente") or "").strip())
                    nome = upper(str(rr.get("nome_cliente") or "").strip())
                    if not cod and not nome:
                        continue
                    key = cod or nome
                    item = agg.setdefault(
                        key,
                        {"cod": cod, "nome": nome, "total": 0.0, "forma": ""},
                    )
                    item["total"] += safe_float(rr.get("valor"), 0.0)
                    if not item["forma"]:
                        item["forma"] = upper(str(rr.get("forma_pagamento") or ""))

                parsed_rows = []
                total_geral = 0.0
                for item in sorted(agg.values(), key=lambda it: upper(it.get("nome") or "")):
                    v = safe_float(item.get("total"), 0.0)
                    if v <= 0:
                        continue
                    total_geral += v
                    parsed_rows.append(
                        (
                            upper(item.get("cod") or ""),
                            upper(item.get("nome") or ""),
                            float(v),
                            upper(item.get("forma") or ""),
                        )
                    )
                return header, parsed_rows, float(total_geral)
            except Exception:
                logging.debug("Falha ao montar dados de impressao via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(programacoes)")
            cols_prog = {str(c[1]).lower() for c in (cur.fetchall() or [])}
            nf_expr = "COALESCE(num_nf,'')" if "num_nf" in cols_prog else ("COALESCE(nf_numero,'')" if "nf_numero" in cols_prog else "''")
            data_saida_expr = "COALESCE(data_saida,'')" if "data_saida" in cols_prog else "''"
            hora_saida_expr = "COALESCE(hora_saida,'')" if "hora_saida" in cols_prog else "''"
            data_chegada_expr = "COALESCE(data_chegada,'')" if "data_chegada" in cols_prog else "''"
            hora_chegada_expr = "COALESCE(hora_chegada,'')" if "hora_chegada" in cols_prog else "''"
            cur.execute(
                f"""
                SELECT
                    COALESCE(motorista,''),
                    COALESCE(veiculo,''),
                    COALESCE(equipe,''),
                    {nf_expr},
                    {data_saida_expr},
                    {hora_saida_expr},
                    {data_chegada_expr},
                    {hora_chegada_expr}
                FROM programacoes
                WHERE codigo_programacao=?
                LIMIT 1
                """,
                (prog,),
            )
            r = cur.fetchone()

            if r:
                motorista_raw, veiculo, equipe_raw, nf, dsa, hsa, dch, hch = r
                header["motorista_nome"] = self._resolve_motorista_nome(motorista_raw)
                header["equipe_nomes"] = self._resolve_equipe_integrantes(equipe_raw)
                header["veiculo"] = upper(veiculo or "")
                header["nf"] = upper(nf or "")
                dsa_n, hsa_n = normalize_date_time_components(dsa, hsa)
                dch_n, hch_n = normalize_date_time_components(dch, hch)
                header["data_saida"] = dsa_n
                header["hora_saida"] = hsa_n
                header["data_chegada"] = dch_n
                header["hora_chegada"] = hch_n

            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='recebimentos'")
            if cur.fetchone():
                cur.execute("PRAGMA table_info(recebimentos)")
                cols_rec = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                cod_expr = "COALESCE(cod_cliente,'')" if "cod_cliente" in cols_rec else "''"
                nome_expr = "COALESCE(nome_cliente,'')" if "nome_cliente" in cols_rec else "''"
                valor_expr = "COALESCE(valor,0)" if "valor" in cols_rec else "0"
                forma_expr = "COALESCE(forma_pagamento,'')" if "forma_pagamento" in cols_rec else "''"
                data_expr = "COALESCE(data_registro,'')" if "data_registro" in cols_rec else "''"
                cliente_where = (
                    "upper(COALESCE(r2.cod_cliente,'')) = upper(COALESCE(r.cod_cliente,''))"
                    if "cod_cliente" in cols_rec
                    else "1=1"
                )
                group_cod = "upper(COALESCE(cod_cliente,''))" if "cod_cliente" in cols_rec else "''"
                order_id_expr = "r2.id DESC" if "id" in cols_rec else "1 DESC"
                cur.execute(
                    f"""
                    SELECT
                        {cod_expr},
                        {nome_expr},
                        SUM({valor_expr}) AS total_valor,
                        (
                            SELECT {forma_expr}
                            FROM recebimentos r2
                            WHERE r2.codigo_programacao = r.codigo_programacao
                              AND {cliente_where}
                            ORDER BY datetime({data_expr}) DESC, {order_id_expr}
                            LIMIT 1
                        ) AS forma_recente
                    FROM recebimentos r
                    WHERE codigo_programacao=?
                    GROUP BY {group_cod}, upper(COALESCE(nome_cliente,''))
                    HAVING SUM({valor_expr}) > 0
                    ORDER BY upper(COALESCE(nome_cliente,'')) ASC
                    """,
                    (prog,),
                )
                rows = cur.fetchall() or []
            else:
                rows = []

        parsed_rows = []
        total_geral = 0.0
        for cod, nome, total_valor, forma_recente in rows:
            v = safe_float(total_valor, 0.0)
            total_geral += v
            parsed_rows.append((upper(cod), upper(nome), float(v), upper(forma_recente or "")))

        return header, parsed_rows, float(total_geral)


    def imprimir_pdf(self):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Carregue uma programação primeiro.")
            return
        if not require_reportlab():
            return

        header, rows, total_geral = self._get_dados_para_impressao(prog)
        if not rows:
            messagebox.showwarning("Impressão", "Não há clientes pagantes (valor > 0) para imprimir.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar PDF - Recebimentos",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"RECEBIMENTOS_{prog}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        )
        if not path:
            return

        try:
            # ReportLab (já instalado no seu ambiente)
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm

            c = canvas.Canvas(path, pagesize=A4)
            width, height = A4

            # Margens
            left = 12 * mm
            right = 12 * mm
            top = 12 * mm
            bottom = 12 * mm

            y = height - top

            # ===== TTULO =====
            c.setFont("Helvetica-Bold", 14)
            c.drawString(left, y, f"RECIBO / RELATÓRIO DE RECEBIMENTOS - PROGRAMAÇÃO {header['prog']}")
            y -= 10 * mm

            # ===== DADOS DO MOTORISTA / ROTA (BLOCO ALINHADO) =====
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "DADOS DO MOTORISTA / ROTA")
            y -= 6 * mm

            c.setFont("Helvetica", 9)

            # Layout em duas colunas (2 linhas)
            col1_x = left
            col2_x = left + (width - left - right) * 0.55

            line_h = 5.2 * mm

            def draw_kv(x, y0, k, v):
                c.setFont("Helvetica-Bold", 9)
                c.drawString(x, y0, f"{k}:")
                c.setFont("Helvetica", 9)
                c.drawString(x + 28 * mm, y0, v or "")

            draw_kv(col1_x, y, "Motorista", header["motorista_nome"])
            draw_kv(col2_x, y, "Veiculo", header["veiculo"])
            y -= line_h

            draw_kv(col1_x, y, "Equipe", header["equipe_nomes"])
            draw_kv(col2_x, y, "NF", header["nf"])
            y -= line_h

            draw_kv(col1_x, y, "SaÃÂÂda", f"{header['data_saida']} {header['hora_saida']}".strip())
            draw_kv(col2_x, y, "Chegada", f"{header['data_chegada']} {header['hora_chegada']}".strip())
            y -= 8 * mm

            # Linha separadora
            c.setLineWidth(0.8)
            c.line(left, y, width - right, y)
            y -= 6 * mm

            # ===== TABELA (apenas pagantes) =====
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "RECEBIMENTOS (APENAS CLIENTES PAGANTES)")
            y -= 6 * mm

            # Cabeçalho da tabela
            table_x = left
            table_w = width - left - right

            # Colunas: COD | CLIENTE | FORMA | VALOR
            col_cod = table_w * 0.16
            col_nome = table_w * 0.50
            col_forma = table_w * 0.18
            col_valor = table_w * 0.16

            x_cod = table_x
            x_nome = x_cod + col_cod
            x_forma = x_nome + col_nome
            x_valor = x_forma + col_forma

            row_h = 6.2 * mm

            # função segura p/ quebrar texto
            def clip_text(s, max_chars):
                s = (s or "")
                return s if len(s) <= max_chars else s[:max_chars - 1] + "âââ€š¬¦"

            # desenha header
            c.setFont("Helvetica-Bold", 9)
            c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            c.drawString(x_cod + 2, y - row_h + 3, "CÓD CLIENTE")
            c.drawString(x_nome + 2, y - row_h + 3, "NOME CLIENTE")
            c.drawString(x_forma + 2, y - row_h + 3, "FORMA PGTO")
            c.drawRightString(x_valor + col_valor - 2, y - row_h + 3, "VALOR (R$)")
            y -= row_h

            c.setFont("Helvetica", 9)

            # desenha linhas
            for cod, nome, valor, forma in rows:
                # quebra de página se necessário
                if y < bottom + 55 * mm:
                    c.showPage()
                    width, height = A4
                    y = height - top
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(left, y, f"RECIBO / RELATÓRIO DE RECEBIMENTOS - PROGRAMAÇÃO {header['prog']}")
                    y -= 10 * mm

                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(left, y, "RECEBIMENTOS (CONTINUAÇÃO)")
                    y -= 6 * mm

                    c.setFont("Helvetica-Bold", 9)
                    c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
                    c.drawString(x_cod + 2, y - row_h + 3, "CÓD CLIENTE")
                    c.drawString(x_nome + 2, y - row_h + 3, "NOME CLIENTE")
                    c.drawString(x_forma + 2, y - row_h + 3, "FORMA PGTO")
                    c.drawRightString(x_valor + col_valor - 2, y - row_h + 3, "VALOR (R$)")
                    y -= row_h
                    c.setFont("Helvetica", 9)

                c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
                c.drawString(x_cod + 2, y - row_h + 3, clip_text(cod, 18))
                c.drawString(x_nome + 2, y - row_h + 3, clip_text(nome, 52))
                c.drawString(x_forma + 2, y - row_h + 3, clip_text(forma, 14))
                c.drawRightString(x_valor + col_valor - 2, y - row_h + 3, fmt_money(valor).replace("R$", "").strip())
                y -= row_h

            # total
            y -= 4 * mm
            c.setFont("Helvetica-Bold", 10)
            c.drawRightString(width - right, y, f"TOTAL RECEBIDO: {fmt_money(total_geral)}")
            y -= 10 * mm

            # Linha separadora
            c.setLineWidth(0.8)
            c.line(left, y, width - right, y)
            y -= 10 * mm

            # ===== REA DE ASSINATURAS =====
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "ASSINATURAS / CONFERÊNCIA")
            y -= 10 * mm

            # 2 colunas x 2 linhas (4 setores)
            block_w = (width - left - right - 10 * mm) / 2.0
            block_h = 18 * mm
            gap_x = 10 * mm
            gap_y = 10 * mm

            def assinatura_block(x, y_top, titulo):
                # caixa
                c.setLineWidth(0.8)
                c.rect(x, y_top - block_h, block_w, block_h, stroke=1, fill=0)
                c.setFont("Helvetica-Bold", 9)
                c.drawString(x + 3 * mm, y_top - 5 * mm, titulo)
                # linha assinatura
                c.setFont("Helvetica", 9)
                c.line(x + 3 * mm, y_top - 13 * mm, x + block_w - 3 * mm, y_top - 13 * mm)
                c.drawString(x + 3 * mm, y_top - 16.5 * mm, "Assinatura / Carimbo")

            x1 = left
            x2 = left + block_w + gap_x

            # linha 1
            assinatura_block(x1, y, "SETOR FATURAMENTO")
            assinatura_block(x2, y, "SETOR FINANCEIRO")
            y -= (block_h + gap_y)

            # linha 2
            assinatura_block(x1, y, "SETOR DE CAIXA")
            assinatura_block(x2, y, "SETOR DE CONFERÊNCIA")

            # rodapé
            c.setFont("Helvetica", 8)
            c.drawRightString(width - right, bottom - 2 * mm, f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")

            c.save()

            messagebox.showinfo("OK", f"PDF gerado com sucesso!\n\nArquivo:\n{os.path.basename(path)}")
            self.set_status(f"STATUS: PDF impresso/gerado: {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao gerar PDF: {str(e)}")

# ==========================
# ===== FIM DA PARTE 6C (ATUALIZADA) =====
# ==========================


# ==========================
# ===== INCIO DA PARTE 7 (ATUALIZADA) =====
# ==========================

# =========================================================
# 7.0 CONFIGURAÇÃO DE LOGGING (DespesasPage) - mais seguro
# =========================================================
def setup_despesas_logger():
    """
    Logger com rotação de arquivo para não crescer infinito.
    Não altera regra do sistema, só evita travar com log grande.
    """
    logger = logging.getLogger(__name__ + ".DespesasPage")
    if not logger.handlers:
        # Em app instalado, Program Files e somente leitura.
        # Usa pasta de dados do app (mesma base do DB_PATH) para garantir escrita.
        log_dir = os.path.dirname(DB_PATH) or APP_DIR
        try:
            os.makedirs(log_dir, exist_ok=True)
        except Exception:
            logging.debug("Falha ignorada")
            log_dir = APP_DIR
        log_path = os.path.join(log_dir, "despesas_audit.log")

        try:
            from logging.handlers import RotatingFileHandler
            handler = RotatingFileHandler(
                log_path,
                maxBytes=2_000_000,   # ~2MB
                backupCount=5,
                encoding="utf-8"
            )
        except Exception:
            handler = logging.FileHandler(log_path, encoding="utf-8")

        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.propagate = False
    return logger


# =========================================================
# 7.1 DESPESAS PAGE (PARTE 1/3) âÅâ€œââ‚¬¦ layout mais robusto + scroll seguro
# =========================================================
class DespesasPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Despesas")

        self.selected_despesa_id = None
        self.logger = setup_despesas_logger()
        self._current_programacao = None

        # =========================================================
        # âÅâ€œââ‚¬¦ Remove cinza / body branco (sem mexer no PageBase)
        # =========================================================
        try:
            self.body.configure(padding=0)
            self.body.configure(style="Card.TFrame")
        except Exception:
            logging.debug("Falha ignorada")

        self.body.grid_columnconfigure(0, weight=1)
        self.body.grid_rowconfigure(0, weight=1)

        page_card = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        page_card.grid(row=0, column=0, sticky="nsew")
        page_card.grid_columnconfigure(0, weight=1)

        page_card.grid_rowconfigure(0, weight=0)
        page_card.grid_rowconfigure(1, weight=0)
        page_card.grid_rowconfigure(2, weight=1)
        page_card.grid_rowconfigure(3, weight=0)
        page_card.grid_rowconfigure(4, weight=0)

        # =========================================================
        # ACOES (FILTROS / BOTOES)
        # =========================================================
        actions_frame = ttk.Frame(page_card, style="Card.TFrame")
        actions_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        actions_frame.grid_columnconfigure(6, weight=1)

        ttk.Button(
            actions_frame,
            text="\U0001F504 ATUALIZAR",
            style="Ghost.TButton",
            command=self._refresh_all,
            width=14
        ).grid(row=0, column=2, padx=6)

        self.lbl_stats = ttk.Label(
            actions_frame,
            text="Despesas: 0 | Total: R$ 0,00",
            font=("Segoe UI", 10, "bold"),
            background="white",
            foreground="#2c3e50"
        )
        self.lbl_stats.grid(row=0, column=6, sticky="e", padx=5)

        # =========================================================
        # TOPO (DADOS DA PROGRAMACAO)
        # =========================================================
        top_frame = ttk.Frame(page_card, style="Card.TFrame")
        top_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for i in range(0, 4):
            top_frame.grid_columnconfigure(i, weight=0)
        top_frame.grid_columnconfigure(2, weight=1)

        ttk.Label(
            top_frame,
            text="PROGRAMACAO",
            style="CardTitle.TLabel",
            font=("Segoe UI", 12, "bold")
        ).grid(row=0, column=0, sticky="w", padx=5)

        ttk.Label(top_frame, text="Codigo da Programacao:", style="CardLabel.TLabel").grid(
            row=1, column=0, sticky="w", padx=5, pady=(6, 0)
        )

        self.cb_prog = ttk.Combobox(top_frame, state="readonly", width=30)
        self.cb_prog.grid(row=1, column=1, sticky="w", padx=5, pady=(6, 0))
        self.cb_prog.bind("<<ComboboxSelected>>", lambda e: self._load_by_programacao())

        self.lbl_motorista = ttk.Label(
            top_frame,
            text="Motorista: - | Veiculo: - | Equipe: -",
            font=("Segoe UI", 10),
            background="white",
            wraplength=520,
            justify="left"
        )
        self.lbl_motorista.grid(row=2, column=0, columnspan=4, sticky="w", padx=5, pady=(6, 0))

        # =========================================================
        # SUB-MENUS (Notebook principal da tela de Despesas)
        # =========================================================
        sections_nb = ttk.Notebook(page_card)
        sections_nb.grid(row=2, column=0, sticky="nsew")

        tab_despesas = ttk.Frame(sections_nb, style="Card.TFrame", padding=10)
        sections_nb.add(tab_despesas, text="Despesas Registradas")
        tab_despesas.grid_columnconfigure(0, weight=1)
        tab_despesas.grid_rowconfigure(0, weight=1)

        tab_rota = ttk.Frame(sections_nb, style="Card.TFrame", padding=10)
        sections_nb.add(tab_rota, text="Dados da Rota")
        tab_rota.grid_columnconfigure(1, weight=1)

        tab_nf = ttk.Frame(sections_nb, style="Card.TFrame", padding=10)
        sections_nb.add(tab_nf, text="Nota Fiscal")
        tab_nf.grid_columnconfigure(0, weight=0)
        tab_nf.grid_columnconfigure(1, weight=0)
        tab_nf.grid_columnconfigure(2, weight=1)
        tab_nf.grid_columnconfigure(3, weight=1)

        tab_cedulas = ttk.Frame(sections_nb, style="Card.TFrame", padding=6)
        sections_nb.add(tab_cedulas, text="Contagem + Resumo")
        tab_cedulas.grid_columnconfigure(0, weight=0, minsize=640)
        tab_cedulas.grid_columnconfigure(1, weight=1)

        tab_resumos = tab_cedulas
        tab_resumos.grid_columnconfigure(0, weight=1)

        # =========================================================
        # TABELA (DESPESAS)
        # =========================================================
        table_frame = ttk.Frame(tab_despesas, style="Card.TFrame")
        table_frame.grid(row=0, column=0, sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(1, weight=1)

        header_frame = ttk.Frame(table_frame, style="Card.TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header_frame.grid_columnconfigure(2, weight=1)

        ttk.Label(
            header_frame,
            text="DESPESAS REGISTRADAS",
            style="CardTitle.TLabel",
            font=("Segoe UI", 12, "bold")
        ).grid(row=0, column=0, sticky="w")

        ttk.Button(
            header_frame,
            text="ADICIONAR DESPESA",
            style="Primary.TButton",
            command=self._open_registrar_rapido
        ).grid(row=0, column=1, padx=10)

        columns = ["ID", "DESCRICAO", "VALOR", "DATA", "CATEGORIA", "OBSERVACAO"]
        self.tree_desp = ttk.Treeview(table_frame, columns=columns, show="headings")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree_desp.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree_desp.xview)
        self.tree_desp.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree_desp.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew", columnspan=2)
        self.tree_desp.tag_configure("odd", background="#f6f7f9")

        col_widths = {
            "ID": 55,
            "DESCRICAO": 300,
            "VALOR": 95,
            "DATA": 95,
            "CATEGORIA": 120,
            "OBSERVACAO": 200
        }

        for col in columns:
            self.tree_desp.heading(col, text=col, command=lambda c=col: self._sort_despesas_by_column(c))
            self.tree_desp.column(
                col,
                width=col_widths.get(col, 100),
                minwidth=60,
                anchor="center" if col in ["ID", "VALOR", "DATA"] else "w"
            )

        self.tree_desp.bind("<<TreeviewSelect>>", self._on_select_despesa)
        self.tree_desp.bind("<Double-1>", lambda e: self._editar_linha_selecionada())
        # =========================================================
        # ---- ABA 1: NOTA FISCAL
        # =========================================================

        calc_frame = ttk.Frame(tab_nf)
        calc_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Button(
            calc_frame,
            text="Calcular Saldo Automático",
            style="Ghost.TButton",
            width=24,
            command=self._calcular_saldo_auto
        ).pack(anchor="w")

        tab_nf.grid_columnconfigure(0, weight=5)
        tab_nf.grid_columnconfigure(1, weight=4)
        tab_nf.grid_rowconfigure(1, weight=1)

        nf_form_wrap = ttk.Frame(tab_nf, style="Card.TFrame")
        nf_form_wrap.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        nf_form_wrap.grid_columnconfigure(0, weight=1)
        nf_form_wrap.grid_columnconfigure(1, weight=1)

        nf_base = ttk.LabelFrame(nf_form_wrap, text="NF Base", padding=8)
        nf_base.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
        nf_base.grid_columnconfigure(1, weight=1)

        row_nf = 0
        self.ent_nf_numero = self._create_field(nf_base, "Nº NOTA FISCAL:", row_nf, 12); row_nf += 1
        self.ent_nf_kg = self._create_field(nf_base, "KG NOTA FISCAL:", row_nf, 12); row_nf += 1
        self.ent_nf_preco = self._create_field(nf_base, "PRECO NF (R$/KG):", row_nf, 12); row_nf += 1
        self.ent_nf_caixas = self._create_field(nf_base, "CAIXAS:", row_nf, 12)

        nf_mov = ttk.LabelFrame(nf_form_wrap, text="Movimento", padding=8)
        nf_mov.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6))
        nf_mov.grid_columnconfigure(1, weight=1)

        row_nf = 0
        self.ent_nf_kg_carregado = self._create_field(nf_mov, "KG CARREGADO:", row_nf, 12); row_nf += 1
        self.ent_nf_kg_vendido = self._create_field(nf_mov, "KG VENDIDO:", row_nf, 12); row_nf += 1
        self.ent_nf_saldo = self._create_field(nf_mov, "SALDO (KG):", row_nf, 12)

        nf_params = ttk.LabelFrame(nf_form_wrap, text="Parametros e Sincronismo", padding=8)
        nf_params.grid(row=1, column=0, columnspan=2, sticky="ew")
        for idx, minsize in enumerate((120, 120, 105, 120, 180)):
            nf_params.grid_columnconfigure(idx, weight=1 if idx in (1, 3) else 0, minsize=minsize)

        ttk.Label(nf_params, text="MEDIA CARREGADA:", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", pady=2)
        self.ent_nf_media_carregada = ttk.Entry(nf_params, style="Field.TEntry", width=12)
        self.ent_nf_media_carregada.grid(row=0, column=1, sticky="ew", padx=(10, 12), pady=2)
        bind_entry_smart(self.ent_nf_media_carregada, "decimal", precision=2)
        self._bind_focus_scroll(self.ent_nf_media_carregada)

        ttk.Label(nf_params, text="CAIXA FINAL:", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w", pady=2)
        self.ent_nf_caixa_final = ttk.Entry(nf_params, style="Field.TEntry", width=12)
        self.ent_nf_caixa_final.grid(row=0, column=3, sticky="ew", padx=(10, 12), pady=2)
        bind_entry_smart(self.ent_nf_caixa_final, "int")
        self._bind_focus_scroll(self.ent_nf_caixa_final)
        try:
            self.ent_nf_caixas.configure(state="readonly")
            self.ent_nf_kg_carregado.configure(state="readonly")
            self.ent_nf_kg_vendido.configure(state="readonly")
        except Exception:
            logging.debug("Falha ignorada")

        ttk.Button(
            nf_params,
            text="Sincronizar com App",
            style="Warn.TButton",
            width=20,
            command=self._sincronizar_com_app
        ).grid(row=0, column=4, sticky="e", padx=(10, 0), pady=2)

        nf_resumo = ttk.LabelFrame(tab_nf, text="Resumo Compra x Venda", padding=12)
        nf_resumo.grid(row=1, column=1, sticky="nsew")
        nf_resumo.grid_columnconfigure(0, weight=1)

        ttk.Label(nf_resumo, text="COMPRA", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.lbl_nf_compra_kg = ttk.Label(nf_resumo, text="KG NF: 0,00")
        self.lbl_nf_compra_kg.grid(row=1, column=0, sticky="w")
        self.lbl_nf_compra_preco = ttk.Label(nf_resumo, text="Preco NF: R$ 0,00")
        self.lbl_nf_compra_preco.grid(row=2, column=0, sticky="w")
        self.lbl_nf_compra_total = ttk.Label(nf_resumo, text="Total compra: R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_nf_compra_total.grid(row=3, column=0, sticky="w", pady=(0, 8))

        ttk.Separator(nf_resumo, orient="horizontal").grid(row=4, column=0, sticky="ew", pady=(2, 8))

        ttk.Label(nf_resumo, text="VENDA", font=("Segoe UI", 10, "bold")).grid(row=5, column=0, sticky="w", pady=(0, 6))
        self.lbl_nf_venda_preco = ttk.Label(nf_resumo, text="Preco medio venda: R$ 0,00")
        self.lbl_nf_venda_preco.grid(row=6, column=0, sticky="w")
        self.lbl_nf_venda_kg = ttk.Label(nf_resumo, text="KG vendido: 0,00")
        self.lbl_nf_venda_kg.grid(row=7, column=0, sticky="w")
        self.lbl_nf_venda_total = ttk.Label(nf_resumo, text="Receita estimada: R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_nf_venda_total.grid(row=8, column=0, sticky="w", pady=(0, 8))

        ttk.Separator(nf_resumo, orient="horizontal").grid(row=9, column=0, sticky="ew", pady=(2, 8))

        ttk.Label(nf_resumo, text="RESULTADO", font=("Segoe UI", 10, "bold")).grid(row=10, column=0, sticky="w", pady=(0, 6))
        self.lbl_nf_desp_total = ttk.Label(nf_resumo, text="Despesas da rota: R$ 0,00")
        self.lbl_nf_desp_total.grid(row=11, column=0, sticky="w")
        self.lbl_nf_lucro_bruto = ttk.Label(nf_resumo, text="Lucro bruto: R$ 0,00")
        self.lbl_nf_lucro_bruto.grid(row=12, column=0, sticky="w")
        self.lbl_nf_lucro_liquido = ttk.Label(nf_resumo, text="Lucro liquido: R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_nf_lucro_liquido.grid(row=13, column=0, sticky="w")
        self.lbl_nf_margem = ttk.Label(nf_resumo, text="Margem liquida: 0,00%")
        self.lbl_nf_margem.grid(row=14, column=0, sticky="w", pady=(0, 4))
        ttk.Separator(nf_resumo, orient="horizontal").grid(row=15, column=0, sticky="ew", pady=(2, 8))
        ocorr_wrap = ttk.Frame(nf_resumo, style="CardInset.TFrame", padding=(8, 6))
        ocorr_wrap.grid(row=16, column=0, sticky="ew")
        ocorr_wrap.grid_columnconfigure(0, weight=1)
        ocorr_wrap.grid_columnconfigure(1, weight=0)
        ttk.Label(ocorr_wrap, text="OCORRÊNCIAS (TRANSBORDO)", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.btn_toggle_nf_ocorrencias = ttk.Button(
            ocorr_wrap,
            text="Expandir",
            style="Toggle.TButton",
            command=self._toggle_nf_ocorrencias,
        )
        self.btn_toggle_nf_ocorrencias.grid(row=0, column=1, sticky="e")
        self.nf_ocorrencias_content = ttk.Frame(ocorr_wrap, style="CardInset.TFrame")
        self.nf_ocorrencias_content.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.nf_ocorrencias_content.grid_columnconfigure(0, weight=1)
        self.lbl_nf_mort_aves = ttk.Label(self.nf_ocorrencias_content, text="Mortalidade (aves): 0")
        self.lbl_nf_mort_aves.grid(row=0, column=0, sticky="w")
        self.lbl_nf_mort_kg = ttk.Label(self.nf_ocorrencias_content, text="Mortalidade (KG): 0,00")
        self.lbl_nf_mort_kg.grid(row=1, column=0, sticky="w")
        self.lbl_nf_kg_util = ttk.Label(self.nf_ocorrencias_content, text="KG útil NF (NF - mortalidade): 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_nf_kg_util.grid(row=2, column=0, sticky="w")
        self.lbl_nf_obs_transb = ttk.Label(self.nf_ocorrencias_content, text="Obs transbordo: -", wraplength=280, justify="left")
        self.lbl_nf_obs_transb.grid(row=3, column=0, sticky="w", pady=(2, 0))
        self._nf_ocorrencias_collapsed = True
        self.nf_ocorrencias_content.grid_remove()

        self._bind_nf_summary_auto_calc()
        self._refresh_nf_trade_summary()

        # =========================================================
        # ---- ABA 2: DADOS DA ROTA
        # =========================================================

        tab_rota.grid_columnconfigure(0, weight=0)
        tab_rota.grid_columnconfigure(1, weight=1)

        km_form = ttk.Frame(tab_rota, style="Card.TFrame")
        km_form.grid(row=0, column=0, sticky="nw", padx=(0, 12))

        calc_km_frame = ttk.Frame(km_form, style="Card.TFrame")
        calc_km_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Button(
            calc_km_frame,
            text="Calcular KM/Média",
            style="Ghost.TButton",
            width=20,
            command=self._calcular_km_media
        ).grid(row=0, column=0, sticky="w")

        row_km = 1
        self.ent_km_inicial = self._create_field(km_form, "KM INICIAL:", row_km, 12); row_km += 1
        self.ent_km_final = self._create_field(km_form, "KM FINAL:", row_km, 12); row_km += 1
        self.ent_litros = self._create_field(km_form, "LITROS:", row_km, 12); row_km += 1
        self.ent_km_rodado = self._create_field(km_form, "KM RODADO:", row_km, 12); row_km += 1
        self.ent_media = self._create_field(km_form, "MÉDIA (KM/L):", row_km, 12); row_km += 1
        self.ent_custo_km = self._create_field(km_form, "CUSTO P/KM:", row_km, 12); row_km += 1
        try:
            self.ent_km_rodado.configure(state="readonly")
            self.ent_media.configure(state="readonly")
            self.ent_custo_km.configure(state="readonly")
        except Exception:
            logging.debug("Falha ignorada")
        self._bind_km_auto_calc()
        self.lbl_km_anterior = ttk.Label(
            km_form,
            text="Ultimo KM do veiculo: -",
            style="CardLabel.TLabel",
            foreground="#1D4ED8",
        )
        self.lbl_km_anterior.grid(row=row_km, column=0, columnspan=2, sticky="w", pady=(4, 2))
        row_km += 1

        ttk.Label(km_form, text="OBSERVAÇÃO:", style="CardLabel.TLabel").grid(row=row_km, column=0, sticky="nw", pady=(8, 2))
        self.txt_rota_obs = tk.Text(km_form, height=4, width=28, font=("Segoe UI", 9))
        self.txt_rota_obs.grid(row=row_km, column=1, sticky="w", pady=(8, 2), padx=(10, 0))

        km_side = ttk.Frame(tab_rota, style="Card.TFrame")
        km_side.grid(row=0, column=1, sticky="nsew")
        km_side.grid_columnconfigure(0, weight=1)
        km_side.grid_rowconfigure(1, weight=1)

        ttk.Label(km_side, text="MELHORES ÚLTIMAS MÉDIAS", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        chart_wrap = ttk.Frame(km_side, style="Card.TFrame")
        chart_wrap.grid(row=1, column=0, sticky="nsew", pady=(6, 10))
        chart_wrap.grid_columnconfigure(0, weight=1)

        self.canvas_km_pie = tk.Canvas(chart_wrap, width=300, height=180, bg="white", highlightthickness=1, highlightbackground="#E5E7EB")
        self.canvas_km_pie.grid(row=0, column=0, sticky="w")

        ttk.Label(km_side, text="ULTIMOS KM POR VEICULO", style="CardTitle.TLabel").grid(row=2, column=0, sticky="w", pady=(4, 4))
        self.tree_km_veiculos = ttk.Treeview(
            km_side,
            columns=["VEICULO", "KM", "MÉDIA", "DATA"],
            show="headings",
            height=6
        )
        for col, w in [("VEICULO", 130), ("KM", 80), ("MÉDIA", 80), ("DATA", 110)]:
            self.tree_km_veiculos.heading(col, text=col)
            self.tree_km_veiculos.column(col, width=w, anchor="center" if col != "VEICULO" else "w")
        self.tree_km_veiculos.grid(row=3, column=0, sticky="ew")

        # =========================================================
        # ---- ABA 3: CONTAGEM DE CÉDULAS
        # =========================================================

        contagem_wrap = ttk.LabelFrame(tab_cedulas, text="CONTAGEM DE CEDULAS", padding=6)
        contagem_wrap.grid(row=0, column=0, sticky="nw", pady=(0, 6), padx=(0, 8))
        contagem_wrap.grid_columnconfigure(0, weight=0, minsize=220)
        contagem_wrap.grid_columnconfigure(1, weight=0, minsize=220)
        contagem_wrap.grid_columnconfigure(2, weight=0, minsize=220)

        calc_ced_frame = ttk.Frame(contagem_wrap)
        calc_ced_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 4))

        ttk.Label(calc_ced_frame, text="Valor Total:", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.ent_total_dinheiro = ttk.Entry(calc_ced_frame, style="Field.TEntry", width=12)
        self.ent_total_dinheiro.grid(row=0, column=1, sticky="w")
        self._bind_money_entry(self.ent_total_dinheiro)
        self._bind_focus_scroll(self.ent_total_dinheiro)
        ttk.Button(
            calc_ced_frame,
            text="\U0001F9EE Distribuir",
            style="Ghost.TButton",
            command=self._distribuir_cedulas,
            width=11,
        ).grid(row=0, column=2, sticky="w", padx=(10, 0))

        header_ced = ttk.Frame(contagem_wrap)
        header_ced.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(2, 3))
        for _c in range(3):
            header_ced.grid_columnconfigure(_c, weight=1, uniform="ced_head")

        ttk.Label(header_ced, text="QTD", width=8, style="CardLabel.TLabel", anchor="center", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="ew")
        ttk.Label(header_ced, text="CEDULA", width=15, style="CardLabel.TLabel", anchor="center").grid(row=0, column=1, sticky="ew")
        ttk.Label(header_ced, text="TOTAL", width=15, style="CardLabel.TLabel", anchor="center", font=("Segoe UI", 8)).grid(row=0, column=2, sticky="ew")

        self.ced_entries = {}
        self.ced_totals = {}
        ced_list = [200, 100, 50, 20, 10, 5, 2]
        ced_colors = {
            200: "#166534",  # verde escuro
            100: "#15803D",
            50: "#B45309",   # laranja escuro
            20: "#9A3412",
            10: "#1D4ED8",   # azul
            5: "#7C3AED",    # roxo
            2: "#BE123C",    # vermelho
        }

        for i, ced in enumerate(ced_list, start=2):
            ent = ttk.Entry(contagem_wrap, width=12, style="Field.TEntry", justify="center")
            ent.grid(row=i, column=0, sticky="ew", pady=0, padx=4)
            ent.bind("<KeyRelease>", lambda e: self._calc_valor_dinheiro())
            self._bind_focus_scroll(ent)

            lbl_ced = ttk.Label(
                contagem_wrap,
                text=f"R$ {ced:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                style="CardLabel.TLabel",
                anchor="center",
                foreground=ced_colors.get(ced, "#111827")
            )
            lbl_ced.grid(row=i, column=1, sticky="ew", pady=0, padx=4)

            lbl_total = ttk.Label(contagem_wrap, text="R$ 0,00",
                                  font=("Segoe UI", 8, "bold"), anchor="center",
                                  foreground=ced_colors.get(ced, "#111827"))
            lbl_total.grid(row=i, column=2, sticky="ew", pady=0, padx=4)

            self.ced_entries[ced] = ent
            self.ced_totals[ced] = lbl_total

        ttk.Separator(contagem_wrap, orient="horizontal")\
            .grid(row=9, column=0, columnspan=3, sticky="ew", pady=4)

        ttk.Label(contagem_wrap, text="TOTAL DINHEIRO:", font=("Segoe UI", 9, "bold"))\
            .grid(row=10, column=0, columnspan=2, sticky="e", pady=2)

        self.lbl_valor_dinheiro = ttk.Label(contagem_wrap, text="R$ 0,00", font=("Segoe UI", 11, "bold"))
        self.lbl_valor_dinheiro.grid(row=10, column=2, sticky="w", pady=2, padx=4)
        ttk.Separator(contagem_wrap, orient="horizontal").grid(row=11, column=0, columnspan=3, sticky="ew", pady=(4, 6))

        # força atualizar scroll no primeiro render
        # =========================================================
        # ? RESUMO + BOT?ES (como voc? j? tinha)
        # =========================================================
        # =========================================================
        # RESUMO + BOTOES
        # =========================================================
        # =========================================================
        # RESUMO + BOTOES
        # =========================================================
        resumo_frame = ttk.Frame(tab_resumos, style="Card.TFrame", padding=6)
        resumo_frame.grid(row=0, column=1, sticky="nsew")
        resumo_frame.grid_columnconfigure(0, weight=1)
        resumo_frame.grid_rowconfigure(2, weight=1)

        ttk.Label(
            resumo_frame,
            text="RESUMO FINANCEIRO INTELIGENTE",
            style="CardTitle.TLabel",
            font=("Segoe UI", 11, "bold")
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        kpi_row = ttk.Frame(resumo_frame, style="Card.TFrame")
        kpi_row.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        for i in range(4):
            kpi_row.grid_columnconfigure(i, weight=1, uniform="kpi")

        kpi_receb = ttk.LabelFrame(kpi_row, text=" RECEBIMENTOS ", padding=5)
        kpi_receb.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.lbl_kpi_receb_total = ttk.Label(kpi_receb, text="R$ 0,00", font=("Segoe UI", 11, "bold"), foreground="#1D4ED8")
        self.lbl_kpi_receb_total.grid(row=0, column=0, sticky="w")

        kpi_desp = ttk.LabelFrame(kpi_row, text=" DESPESAS ", padding=5)
        kpi_desp.grid(row=0, column=1, sticky="ew", padx=(0, 4))
        self.lbl_kpi_desp_total = ttk.Label(kpi_desp, text="R$ 0,00", font=("Segoe UI", 11, "bold"), foreground="#C62828")
        self.lbl_kpi_desp_total.grid(row=0, column=0, sticky="w")

        kpi_ced = ttk.LabelFrame(kpi_row, text=" CÉDULAS ", padding=5)
        kpi_ced.grid(row=0, column=2, sticky="ew", padx=(0, 4))
        self.lbl_kpi_cedulas_total = ttk.Label(kpi_ced, text="R$ 0,00", font=("Segoe UI", 11, "bold"), foreground="#7C3AED")
        self.lbl_kpi_cedulas_total.grid(row=0, column=0, sticky="w")

        kpi_res = ttk.LabelFrame(kpi_row, text=" RESULTADO LIQUIDO ", padding=5)
        kpi_res.grid(row=0, column=3, sticky="ew")
        self.lbl_kpi_resultado_liquido = ttk.Label(kpi_res, text="R$ 0,00", font=("Segoe UI", 11, "bold"), foreground="#2E7D32")
        self.lbl_kpi_resultado_liquido.grid(row=0, column=0, sticky="w")

        details_wrap = ttk.Frame(resumo_frame, style="Card.TFrame")
        details_wrap.grid(row=2, column=0, sticky="nsew")
        for i in range(2):
            details_wrap.grid_columnconfigure(i, weight=1, uniform="res_cards")
        details_wrap.grid_rowconfigure(0, weight=1)
        details_wrap.grid_rowconfigure(1, weight=1)

        entrada_frame = ttk.LabelFrame(details_wrap, text=" ENTRADAS ", padding=5)
        entrada_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 3))
        entrada_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(entrada_frame, text="Recebimentos Total:", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=1)
        self.lbl_receb_total = ttk.Label(entrada_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_receb_total.grid(row=0, column=1, sticky="e", pady=1)

        ttk.Label(entrada_frame, text="Adiantamento p/ Rota:", font=("Segoe UI", 8)).grid(row=1, column=0, sticky="w", pady=1)
        self.ent_adiantamento = ttk.Entry(entrada_frame, style="Field.TEntry", width=10)
        self.ent_adiantamento.grid(row=1, column=1, sticky="e", pady=1)
        self._bind_money_entry(self.ent_adiantamento)
        self.ent_adiantamento.bind("<KeyRelease>", lambda e: self._update_resumo_financeiro(), add="+")
        self.ent_adiantamento.bind("<FocusOut>", lambda e: self._update_resumo_financeiro(), add="+")

        ttk.Label(entrada_frame, text="Total Entradas:", font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 1))
        self.lbl_total_entradas = ttk.Label(entrada_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"), foreground="#2E7D32")
        self.lbl_total_entradas.grid(row=2, column=1, sticky="e", pady=(4, 1))

        saida_frame = ttk.LabelFrame(details_wrap, text=" SAIDAS ", padding=5)
        saida_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(3, 0))
        saida_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(saida_frame, text="Despesas Total:", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=1)
        self.lbl_desp_total = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_desp_total.grid(row=0, column=1, sticky="e", pady=1)

        ttk.Label(saida_frame, text="Contagem Cedulas:", font=("Segoe UI", 8)).grid(row=1, column=0, sticky="w", pady=1)
        self.lbl_cedulas_total = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_cedulas_total.grid(row=1, column=1, sticky="e", pady=1)

        ttk.Label(saida_frame, text="Total Saidas:", font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 1))
        self.lbl_total_saidas = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"), foreground="#C62828")
        self.lbl_total_saidas.grid(row=2, column=1, sticky="e", pady=(4, 1))

        resultado_frame = ttk.LabelFrame(details_wrap, text=" RESULTADOS ", padding=5)
        resultado_frame.grid(row=0, column=1, rowspan=2, sticky="nsew")
        resultado_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(resultado_frame, text="Valor p/ Caixa:", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=1)
        self.lbl_valor_final_caixa = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_valor_final_caixa.grid(row=0, column=1, sticky="e", pady=1)

        ttk.Label(resultado_frame, text="Diferenca (Caixa - Ced):", font=("Segoe UI", 8)).grid(row=1, column=0, sticky="w", pady=1)
        self.lbl_diferenca = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_diferenca.grid(row=1, column=1, sticky="e", pady=1)

        ttk.Label(resultado_frame, text="Resultado Liquido:", font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 1))
        self.lbl_resultado_liquido = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 11, "bold"))
        self.lbl_resultado_liquido.grid(row=2, column=1, sticky="e", pady=(4, 1))

        botoes_frame = ttk.Frame(page_card, style="Card.TFrame")
        botoes_frame.grid(row=4, column=0, sticky="ew")

        for i in range(3):
            botoes_frame.grid_columnconfigure(i, weight=1, uniform="desp_btns")

        botoes = [
            ("\u2B05 VOLTAR", "Ghost.TButton", self._voltar_recebimentos),
            ("\u2795 ADICIONAR DESPESA", "Warn.TButton", self._open_registrar_rapido),
            ("\U0001F4D1 GERAR PDF", "Ghost.TButton", self.abrir_previsualizacao_prestacao),
            ("\U0001F4BE SALVAR", "Primary.TButton", self.salvar_tudo),
            ("\u270F\ufe0f EDITAR", "Warn.TButton", self._editar_linha_selecionada),
            ("\U0001F3C1 FINALIZAR", "Danger.TButton", self.finalizar_prestacao_despesas),
        ]

        for i, (texto, estilo, comando) in enumerate(botoes):
            row = i // 3
            col = i % 3
            btn = ttk.Button(botoes_frame, text=texto, style=estilo, command=comando)
            btn.grid(row=row, column=col, sticky="ew", padx=2, pady=2, ipady=1)

        # carrega dados
        self.refresh_comboboxes()
        self._refresh_all()

    def _toggle_nf_ocorrencias(self):
        self._nf_ocorrencias_collapsed = not bool(getattr(self, "_nf_ocorrencias_collapsed", True))
        if self._nf_ocorrencias_collapsed:
            self.nf_ocorrencias_content.grid_remove()
            self.btn_toggle_nf_ocorrencias.configure(text="Expandir")
        else:
            self.nf_ocorrencias_content.grid()
            self.btn_toggle_nf_ocorrencias.configure(text="Ocultar")


# ==========================
# ===== FIM DA PARTE 7 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INICIO DA PARTE 7B (ATUALIZADA) =====
# ==========================

    # =========================================================
    # 7.2 UTILITRIOS / LOAD / FILTROS / CLCULOS
    # =========================================================
    def _scroll_to_widget(self, widget):
        try:
            canvas = getattr(self, "_nb_canvas", None)
            if not canvas:
                return
            self.update_idletasks()
            bbox = canvas.bbox("all")
            if not bbox:
                return
            view_top = canvas.canvasy(0)
            view_bottom = view_top + canvas.winfo_height()
            y = widget.winfo_rooty() - canvas.winfo_rooty() + view_top
            h = widget.winfo_height()
            total_h = (bbox[3] - bbox[1]) or 1
            if y < view_top + 20:
                canvas.yview_moveto(max(0, (y - 20) / total_h))
            elif (y + h) > (view_bottom - 20):
                canvas.yview_moveto(max(0, (y - 20) / total_h))
        except Exception:
            logging.debug("Falha ignorada")

    def _bind_focus_scroll(self, widget):
        try:
            widget.bind("<FocusIn>", lambda e: self._scroll_to_widget(widget), add=True)
        except Exception:
            logging.debug("Falha ignorada")

    def _bind_focus_scroll_tree(self, parent):
        try:
            for child in parent.winfo_children():
                try:
                    child.bind("<FocusIn>", lambda e, w=child: self._scroll_to_widget(w), add=True)
                except Exception:
                    logging.debug("Falha ignorada")
                self._bind_focus_scroll_tree(child)
        except Exception:
            logging.debug("Falha ignorada")

    def _create_field(self, parent, label, row, width=20):
        ttk.Label(parent, text=label, style="CardLabel.TLabel").grid(row=row, column=0, sticky="w", pady=2)
        ent = ttk.Entry(parent, style="Field.TEntry", width=width)
        ent.grid(row=row, column=1, sticky="w", pady=2, padx=(10, 0))
        self._bind_focus_scroll(ent)
        lbl = upper(label)
        if "HORA" in lbl:
            bind_entry_smart(ent, "time")
        elif "DATA" in lbl:
            bind_entry_smart(ent, "date")
        elif "VALOR" in lbl or "R$" in lbl:
            bind_entry_smart(ent, "money")
        elif "CAIXA" in lbl or "QTD" in lbl or "Nº" in lbl or "NOTA FISCAL" in lbl:
            bind_entry_smart(ent, "int")
        elif "KM" in lbl or "KG" in lbl or "LITRO" in lbl or "MÉDIA" in lbl or "MEDIA" in lbl or "CUSTO" in lbl:
            bind_entry_smart(ent, "decimal", precision=2)
        else:
            bind_entry_smart(ent, "text")
        return ent

    def _parse_dt(self, data_str: str, hora_str: str):
        data_str = (data_str or "").strip()
        hora_str = (hora_str or "").strip()
        if not data_str:
            return None
        # aceita yyyy-mm-dd ou dd/mm/yyyy
        try:
            if "-" in data_str and len(data_str) >= 10:
                y, m, d = data_str[:10].split("-")
            elif "/" in data_str and len(data_str) >= 8:
                d, m, y = data_str[:10].split("/")
            else:
                return None
            hh, mm = "00", "00"
            if hora_str and ":" in hora_str:
                hh, mm = hora_str.split(":")[:2]
            y = int(y)
            if y < 100:
                y = (2000 + y) if y <= 69 else (1900 + y)
            return datetime(y, int(m), int(d), int(hh), int(mm))
        except Exception:
            return None

    def _calc_qtd_diarias(self):
        dt_saida = self._parse_dt(self._data_saida, self._hora_saida)
        dt_chegada = self._parse_dt(self._data_chegada, self._hora_chegada)
        if not dt_saida:
            return 0.0
        if not dt_chegada or dt_chegada <= dt_saida:
            return 1.0
        horas = (dt_chegada - dt_saida).total_seconds() / 3600.0
        if horas <= 24.0:
            return 1.0
        rem = horas - 24.0
        full = int(rem // 24.0)
        half = 0.5 if (rem - (full * 24.0)) > 0 else 0.0
        return 1.0 + full + half

    def _update_diarias(self):
        try:
            val_ajud = self._parse_money_local(self.ent_diaria_ajudante.get())
        except Exception:
            val_ajud = 0.0
        val_mot = val_ajud + 10.0 if val_ajud > 0 else 0.0
        qtd = self._calc_qtd_diarias()

        self.lbl_diaria_motorista.config(text=f"Diaria Motorista: {self._fmt_money(val_mot)}")
        self.lbl_total_diarias.config(text=f"Qtd Diarias: {qtd:g}")

        total = (val_mot + (val_ajud * 2)) * qtd
        self.lbl_total_pagto.config(text=f"Total Pagamento (Motorista + 2 Ajudantes): {self._fmt_money(total)}")

    # ---- helpers seguros (não mudam regra, só evitam bug)
    def _parse_money(self, v):
        return safe_money(v, 0.0)

    def _fmt_money(self, v):
        try:
            valor = float(v or 0.0)
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return "R$ 0,00"

    def _api_bundle_prestacao(self, prog: str):
        prog = upper(str(prog or "").strip())
        if not prog:
            return None
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret or not is_desktop_api_sync_enabled():
            return None
        try:
            resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(prog)}/bundle",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            rota = resp.get("rota") if isinstance(resp, dict) else None
            if not isinstance(rota, dict):
                return None
            return resp
        except Exception:
            logging.debug("Falha ao montar bundle de prestacao via API.", exc_info=True)
            return None

    def _draw_km_pie(self, items):
        canvas = getattr(self, "canvas_km_pie", None)
        if not canvas:
            return
        canvas.delete("all")

        if not items:
            canvas.create_text(150, 90, text="Sem dados de média para exibir", fill="#6B7280", font=("Segoe UI", 9, "bold"))
            return

        colors = ["#1D4ED8", "#16A34A", "#EA580C", "#7C3AED", "#0891B2", "#BE123C"]
        total = sum(max(0.0, safe_float(v, 0.0)) for _, v in items)
        if total <= 0:
            canvas.create_text(150, 90, text="Sem dados de média para exibir", fill="#6B7280", font=("Segoe UI", 9, "bold"))
            return

        x0, y0, x1, y1 = 10, 10, 170, 170
        start = 0.0
        for idx, (label, value) in enumerate(items):
            v = max(0.0, safe_float(value, 0.0))
            extent = 360.0 * (v / total)
            color = colors[idx % len(colors)]
            canvas.create_arc(x0, y0, x1, y1, start=start, extent=extent, fill=color, outline="white")
            start += extent

        lx, ly = 185, 18
        for idx, (label, value) in enumerate(items):
            color = colors[idx % len(colors)]
            canvas.create_rectangle(lx, ly + idx * 24, lx + 10, ly + 10 + idx * 24, fill=color, outline=color)
            txt = f"{str(label)[:12]}: {safe_float(value, 0.0):.1f}"
            canvas.create_text(lx + 16, ly + 5 + idx * 24, anchor="w", text=txt, fill="#111827", font=("Segoe UI", 8, "bold"))

    def _refresh_km_insights(self):
        if not hasattr(self, "tree_km_veiculos"):
            return

        for iid in self.tree_km_veiculos.get_children():
            self.tree_km_veiculos.delete(iid)

        medias_top = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        loaded_api = False
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/relatorios/km-veiculos",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rows_api = (resp.get("rows") if isinstance(resp, dict) else []) or []
                parsed = [
                    (
                        str((r or {}).get("veiculo") or "-"),
                        safe_int((r or {}).get("viagens"), 0),
                        safe_float((r or {}).get("km_rodado"), 0.0),
                        safe_float((r or {}).get("media_km_l"), 0.0),
                    )
                    for r in rows_api
                    if isinstance(r, dict)
                ]
                medias_top = sorted(parsed, key=lambda t: (-t[3], t[0]))[:5]
                medias_top = [(p[0], p[3]) for p in medias_top if safe_float(p[3], 0.0) > 0]
                for veic, _viagens, km, med in sorted(parsed, key=lambda t: upper(t[0]))[:20]:
                    tree_insert_aligned(self.tree_km_veiculos, "", "end", (veic, f"{km:.1f}", f"{med:.1f}", "-"))
                loaded_api = True
            except Exception:
                logging.debug("Falha ao carregar insights KM via API; usando fallback local.", exc_info=True)

        if not loaded_api:
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    media_expr = "COALESCE(media_km_l, 0)" if "media_km_l" in cols_prog else "0"
                    cur.execute(f"""
                        SELECT COALESCE(veiculo, '-'), {media_expr}
                        FROM programacoes
                        WHERE {media_expr} > 0
                        ORDER BY {media_expr} DESC, id DESC
                        LIMIT 5
                    """)
                    medias_top = [(str(r[0] or "-"), safe_float(r[1], 0.0)) for r in (cur.fetchall() or [])]

                    cur.execute("""
                        SELECT p.veiculo, p.km_rodado, p.media_km_l, p.data_criacao
                        FROM programacoes p
                        INNER JOIN (
                            SELECT veiculo, MAX(id) AS max_id
                            FROM programacoes
                            WHERE COALESCE(veiculo, '') <> ''
                            GROUP BY veiculo
                        ) u ON u.max_id = p.id
                        ORDER BY p.veiculo ASC
                        LIMIT 20
                    """)
                    for row in cur.fetchall() or []:
                        veic = str(row[0] or "-")
                        km = safe_float(row[1], 0.0)
                        med = safe_float(row[2], 0.0)
                        data = str(row[3] or "")[:10]
                        tree_insert_aligned(self.tree_km_veiculos, "", "end", (veic, f"{km:.1f}", f"{med:.1f}", data))
            except Exception:
                logging.debug("Falha ignorada")

        self._draw_km_pie(medias_top)

    def _load_despesas(self, prog):
        busca = upper(self.ent_busca_despesa.get().strip()) if hasattr(self, "ent_busca_despesa") else ""
        categoria = upper(self.cb_filtro_categoria.get().strip()) if hasattr(self, "cb_filtro_categoria") else "TODAS"
        ordem = upper(self.cb_ordem_despesa.get().strip()) if hasattr(self, "cb_ordem_despesa") else "DATA DESC"

        order_sql_map = {
            "DATA DESC": "data_registro DESC",
            "DATA ASC": "data_registro ASC",
            "VALOR DESC": "valor DESC",
            "VALOR ASC": "valor ASC",
        }
        order_sql = order_sql_map.get(ordem, "data_registro DESC")

        rows = []
        self._despesas_cache = {}
        try:
            bundle = self._api_bundle_prestacao(prog)
            itens = (bundle.get("despesas") if isinstance(bundle, dict) else []) or []
            for r in itens:
                rid = safe_int((r or {}).get("id"), 0)
                desc = str((r or {}).get("descricao") or "")
                val = safe_float((r or {}).get("valor"), 0.0)
                data_reg = str((r or {}).get("data_registro") or "")
                cat_row = upper(str((r or {}).get("categoria") or "OUTROS"))
                observacao = str((r or {}).get("observacao") or "")

                if busca:
                    like_txt = f"{upper(desc)} {upper(cat_row)} {upper(observacao)}"
                    if upper(busca) not in like_txt:
                        continue
                if categoria and categoria != "TODAS" and upper(cat_row) != upper(categoria):
                    continue
                rows.append((rid, desc, val, data_reg, cat_row, observacao))
                self._despesas_cache[str(rid)] = (rid, upper(prog), desc, val, cat_row, observacao, data_reg)
        except Exception:
            sql = """
                SELECT id, descricao, valor, data_registro,
                       COALESCE(categoria, 'OUTROS') as categoria,
                       COALESCE(observacao, '') as observacao
                FROM despesas
                WHERE codigo_programacao=?
            """
            params = [prog]

            if busca:
                sql += """
                  AND (
                        UPPER(COALESCE(descricao,'')) LIKE ?
                     OR UPPER(COALESCE(categoria,'')) LIKE ?
                     OR UPPER(COALESCE(observacao,'')) LIKE ?
                  )
                """
                like = f"%{busca}%"
                params.extend([like, like, like])

            if categoria and categoria != "TODAS":
                sql += " AND UPPER(COALESCE(categoria,'OUTROS')) = ? "
                params.append(categoria)

            sql += f" ORDER BY {order_sql}"

            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(sql, tuple(params))
                rows = cur.fetchall()
            for rid, desc, val, data_reg, cat_row, observacao in rows:
                self._despesas_cache[str(rid)] = (rid, upper(prog), desc, val, cat_row, observacao, data_reg)

        if rows:
            if ordem == "DATA ASC":
                rows.sort(key=lambda r: (str(r[3] or ""), safe_int(r[0], 0)))
            elif ordem == "VALOR DESC":
                rows.sort(key=lambda r: (safe_float(r[2], 0.0), safe_int(r[0], 0)), reverse=True)
            elif ordem == "VALOR ASC":
                rows.sort(key=lambda r: (safe_float(r[2], 0.0), safe_int(r[0], 0)))
            else:
                rows.sort(key=lambda r: (str(r[3] or ""), safe_int(r[0], 0)), reverse=True)

        for i in self.tree_desp.get_children():
            self.tree_desp.delete(i)

        total_despesas = 0.0
        for idx, (rid, desc, val, data_reg, cat_row, observacao) in enumerate(rows):
            data_formatada = data_reg[:10] if data_reg else ""
            tag = "odd" if idx % 2 else ""
            tree_insert_aligned(
                self.tree_desp,
                "",
                "end",
                (
                    rid,
                    desc or "",
                    self._fmt_money(val),
                    data_formatada,
                    cat_row or "OUTROS",
                    (observacao or "")[:60],
                ),
                tags=(tag,) if tag else (),
            )
            total_despesas += safe_float(val, 0.0)

        self._ajustar_largura_colunas()
        self.lbl_stats.config(text=f"Despesas: {len(rows)} | Total: {self._fmt_money(total_despesas)}")
        if hasattr(self, "lbl_total_filtrado"):
            self.lbl_total_filtrado.config(text=f"Total filtrado: {self._fmt_money(total_despesas)}")
        if hasattr(self, "lbl_kpi_despesas"):
            self.lbl_kpi_despesas.config(text=f"Despesas: {len(rows)}")
        if hasattr(self, "lbl_kpi_total"):
            self.lbl_kpi_total.config(text=f"Total: {self._fmt_money(total_despesas)}")
        self.set_status(f"STATUS: {len(rows)} despesas carregadas")

    def _ajustar_largura_colunas(self):
        for col in self.tree_desp["columns"]:
            if col == "ID":
                self.tree_desp.column(col, width=55, minwidth=55)
            elif col == "VALOR":
                self.tree_desp.column(col, width=95, minwidth=80)
            elif col == "DATA":
                self.tree_desp.column(col, width=95, minwidth=80)
            elif col == "CATEGORIA":
                self.tree_desp.column(col, width=140, minwidth=110)
            elif col in {"OBSERVACAO"}:
                self.tree_desp.column(col, width=240, minwidth=160)
            elif col in {"DESCRICAO"}:
                self.tree_desp.column(col, width=320, minwidth=200)

    def _filtrar_despesas(self):
        if not self._current_programacao:
            return
        self._load_despesas(self._current_programacao)

    def _refresh_all(self):
        # evita refazer trabalho excessivo se nada selecionado
        self.refresh_comboboxes()
        self._load_by_programacao()
        if hasattr(self, "lbl_data_ref"):
            self.lbl_data_ref.config(text=datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.set_status("STATUS: Todos os dados atualizados")

    def _voltar_recebimentos(self):
        try:
            self.app.show_page("Recebimentos")
        except Exception:
            logging.debug("Falha ignorada")

    def set_programacao(self, prog: str):
        prog = upper(str(prog or "").strip())
        if not prog:
            return
        try:
            if hasattr(self, "refresh_comboboxes"):
                self.refresh_comboboxes()
            try:
                vals = list(self.cb_prog["values"])
            except Exception:
                vals = []
            cods_existentes = {upper(str(v or "").split(" ")[0].strip()) for v in vals}
            if prog not in cods_existentes:
                vals = [prog] + vals
                self.cb_prog["values"] = vals
            alvo = None
            for v in vals:
                if upper(str(v or "").split(" ")[0].strip()) == prog:
                    alvo = v
                    break
            self.cb_prog.set(alvo or prog)
            self._load_by_programacao()
        except Exception:
            logging.debug("Falha ignorada")

    def _load_by_programacao(self):
        selecionado = self.cb_prog.get()
        if not selecionado:
            for i in self.tree_desp.get_children():
                self.tree_desp.delete(i)
            self.lbl_motorista.config(text="Motorista: - | Veiculo: - | Equipe: -")
            if hasattr(self, "lbl_rota_chip"):
                self.lbl_rota_chip.config(text="---")
            self.lbl_stats.config(text="Despesas: 0 | Total: R$ 0,00")
            if hasattr(self, "lbl_total_filtrado"):
                self.lbl_total_filtrado.config(text="Total filtrado: R$ 0,00")
            if hasattr(self, "lbl_kpi_despesas"):
                self.lbl_kpi_despesas.config(text="Despesas: 0")
            if hasattr(self, "lbl_kpi_total"):
                self.lbl_kpi_total.config(text="Total: R$ 0,00")
            if hasattr(self, "lbl_kpi_status"):
                self.lbl_kpi_status.config(text="Status: Em aberto")
            self._reset_campos()
            self._update_resumo_financeiro()
            return

        prog = selecionado.split(" ")[0]
        self._current_programacao = prog
        if hasattr(self, "lbl_rota_chip"):
            self.lbl_rota_chip.config(text=prog)

        try:
            self._sync_programacao_from_api_desktop(prog, silent=True)
        except Exception:
            logging.debug("Falha ignorada")

        self._load_despesas(prog)
        self._load_programacao_extras(prog)
        try:
            self._calcular_km_media()
        except Exception:
            logging.debug("Falha ignorada")
        self._calc_valor_dinheiro()
        self._update_resumo_financeiro()
        status = self._get_prestacao_status(prog)
        if hasattr(self, "lbl_kpi_status"):
            self.lbl_kpi_status.config(text=f"Status: {'Finalizado' if status == 'FECHADA' else 'Em aberto'}")
        self.logger.info(f"Programacao carregada: {prog}")

    def _reset_campos(self):
        # NF
        for attr in ("ent_nf_numero", "ent_nf_kg", "ent_nf_preco", "ent_nf_caixas",
                     "ent_nf_kg_carregado", "ent_nf_kg_vendido", "ent_nf_saldo",
                     "ent_nf_media_carregada", "ent_nf_caixa_final"):
            field = getattr(self, attr, None)
            if field:
                self._safe_set_entry(field, "")
        self._refresh_nf_trade_summary()

        # Rota / KM
        for attr in ("ent_km_inicial", "ent_km_final", "ent_litros",
                     "ent_km_rodado", "ent_media", "ent_custo_km"):
            field = getattr(self, attr, None)
            if field:
                self._safe_set_entry(field, "")

        if hasattr(self, "txt_rota_obs") and self.txt_rota_obs:
            try:
                self.txt_rota_obs.delete("1.0", "end")
            except Exception:
                logging.debug("Falha ignorada")

        # Cédulas
        for ced in self.ced_entries.values():
            ced.delete(0, "end")
            ced.insert(0, "0")

        # Adiantamento
        self.ent_adiantamento.delete(0, "end")
        self.ent_adiantamento.insert(0, "0,00")

        # Total dinheiro
        self.ent_total_dinheiro.delete(0, "end")
        self.ent_total_dinheiro.insert(0, "0,00")

        self.lbl_valor_dinheiro.config(text="R$ 0,00")

        for lbl in self.ced_totals.values():
            lbl.config(text="R$ 0,00")
        if hasattr(self, "lbl_km_anterior"):
            self.lbl_km_anterior.config(text="Ultimo KM do veiculo: -")

        self._refresh_km_insights()

    def _load_programacao_extras(self, prog):
        row = None
        cols_all = []
        col_adiant = None
        has_data_saida = False
        has_hora_saida = False
        has_data_chegada = False
        has_hora_chegada = False
        has_rota_obs = False
        has_media_carregada = False
        has_caixa_final = False
        has_nf_preco = False
        has_mort_aves = False
        has_mort_kg = False
        has_obs_transbordo = False

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        try:
            resp = self._api_bundle_prestacao(prog)
            rota_api = resp.get("rota") if isinstance(resp, dict) else None
            if isinstance(rota_api, dict):
                    motorista = str(rota_api.get("motorista") or "")
                    veiculo = str(rota_api.get("veiculo") or "")
                    equipe = str(rota_api.get("equipe") or "")
                    nf_numero = str(rota_api.get("num_nf") or rota_api.get("nf_numero") or "")
                    nf_kg = safe_float(rota_api.get("nf_kg"), 0.0)
                    nf_caixas = safe_int(rota_api.get("nf_caixas"), 0)
                    nf_kg_carregado = safe_float(rota_api.get("nf_kg_carregado"), 0.0)
                    nf_kg_vendido = safe_float(rota_api.get("nf_kg_vendido"), 0.0)
                    nf_saldo = safe_float(rota_api.get("nf_saldo"), 0.0)
                    km_inicial = safe_float(rota_api.get("km_inicial"), 0.0)
                    km_final = safe_float(rota_api.get("km_final"), 0.0)
                    litros = safe_float(rota_api.get("litros"), 0.0)
                    km_rodado = safe_float(rota_api.get("km_rodado"), 0.0)
                    media_km_l = safe_float(rota_api.get("media_km_l"), 0.0)
                    custo_km = safe_float(rota_api.get("custo_km"), 0.0)
                    nf_preco = safe_float(rota_api.get("nf_preco"), 0.0)
                    adiant_val = safe_float(rota_api.get("adiantamento"), safe_float(rota_api.get("adiantamento_rota"), 0.0))
                    self._data_saida = str(rota_api.get("data_saida") or "")
                    self._hora_saida = str(rota_api.get("hora_saida") or "")
                    self._data_chegada = str(rota_api.get("data_chegada") or "")
                    self._hora_chegada = str(rota_api.get("hora_chegada") or "")
                    media_carregada = self._normalize_media_kg_ave(
                        safe_float(rota_api.get("media"), safe_float(rota_api.get("media_carregada"), 0.0))
                    )
                    caixa_final = safe_int(
                        rota_api.get("aves_caixa_final"),
                        safe_int(rota_api.get("qnt_aves_caixa_final"), 0),
                    )
                    rota_obs = str(rota_api.get("rota_observacao") or "")
                    mort_aves_transb = safe_int(rota_api.get("mortalidade_transbordo_aves"), 0)
                    mort_kg_transb = safe_float(rota_api.get("mortalidade_transbordo_kg"), 0.0)
                    obs_transb = str(rota_api.get("obs_transbordo") or "")

                    self._nf_mort_aves = mort_aves_transb
                    self._nf_mort_kg = mort_kg_transb
                    self._nf_obs_transbordo = obs_transb

                    self._data_saida, self._hora_saida = normalize_date_time_components(self._data_saida, self._hora_saida)
                    self._data_chegada, self._hora_chegada = normalize_date_time_components(self._data_chegada, self._hora_chegada)

                    equipe_nomes = resolve_equipe_nomes(equipe)
                    self.lbl_motorista.config(
                        text=fix_mojibake_text(
                            f"Motorista: {motorista or '-'} | Veiculo: {veiculo or '-'} | Equipe: {equipe_nomes or '-'}"
                        )
                    )

                    self.ent_adiantamento.delete(0, "end")
                    self.ent_adiantamento.insert(0, f"{safe_float(adiant_val, 0.0):.2f}".replace(".", ","))

                    self._set_ent(self.ent_nf_numero, nf_numero)
                    self._set_ent(self.ent_nf_kg, nf_kg)
                    if hasattr(self, "ent_nf_preco"):
                        self._set_ent(self.ent_nf_preco, nf_preco)
                    self._set_ent(self.ent_nf_caixas, nf_caixas)
                    self._set_ent(self.ent_nf_kg_carregado, nf_kg_carregado)
                    self._set_ent(self.ent_nf_kg_vendido, nf_kg_vendido)
                    self._set_ent(self.ent_nf_saldo, nf_saldo)
                    if hasattr(self, "ent_nf_media_carregada"):
                        self._set_ent(self.ent_nf_media_carregada, f"{self._normalize_media_kg_ave(media_carregada):.3f}")
                    if hasattr(self, "ent_nf_caixa_final"):
                        self._set_ent(self.ent_nf_caixa_final, safe_int(caixa_final, 0))

                    km_anterior = self._buscar_ultimo_km_final_veiculo(veiculo, prog)
                    if hasattr(self, "lbl_km_anterior"):
                        if km_anterior > 0:
                            self.lbl_km_anterior.config(text=f"Ultimo KM do veiculo: {km_anterior:.1f}")
                        else:
                            self.lbl_km_anterior.config(text="Ultimo KM do veiculo: sem historico")

                    km_inicial_eff = safe_float(km_inicial, 0.0)
                    if km_inicial_eff <= 0 and km_anterior > 0:
                        km_inicial_eff = km_anterior
                        self.set_status(
                            f"STATUS: KM inicial sugerido pelo ultimo KM final do veiculo {veiculo}: {km_anterior:.1f}"
                        )

                    self._set_ent(self.ent_km_inicial, f"{km_inicial_eff:.1f}" if km_inicial_eff > 0 else "")
                    self._set_ent(self.ent_km_final, km_final)
                    self._set_ent(self.ent_litros, litros)
                    self._set_ent(self.ent_km_rodado, km_rodado)
                    self._set_ent(self.ent_media, media_km_l)
                    self._set_ent(self.ent_custo_km, custo_km)

                    if hasattr(self, "txt_rota_obs") and self.txt_rota_obs:
                        try:
                            self.txt_rota_obs.delete("1.0", "end")
                            self.txt_rota_obs.insert("1.0", str(rota_obs or ""))
                        except Exception:
                            logging.debug("Falha ignorada")

                    try:
                        self._calcular_km_media()
                    except Exception:
                        logging.debug("Falha ignorada")
                    try:
                        self._calcular_saldo_auto()
                    except Exception:
                        logging.debug("Falha ignorada")

                    self._refresh_km_insights()
                    return
        except Exception:
            logging.debug("Falha ao carregar extras via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()

            # coluna de adiantamento com fallback seguro
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols_all = [c[1] for c in cur.fetchall()]
            except Exception:
                cols_all = []

            col_adiant = "adiantamento" if "adiantamento" in cols_all else ("adiantamento_rota" if "adiantamento_rota" in cols_all else None)

            has_data_saida = "data_saida" in cols_all
            has_hora_saida = "hora_saida" in cols_all
            has_data_chegada = "data_chegada" in cols_all
            has_hora_chegada = "hora_chegada" in cols_all
            has_rota_obs = "rota_observacao" in cols_all
            has_media_carregada = "media" in cols_all
            has_nf_preco = "nf_preco" in cols_all
            has_aves_caixa_final = "aves_caixa_final" in cols_all
            has_qnt_aves_caixa_final = "qnt_aves_caixa_final" in cols_all
            has_caixa_final = has_aves_caixa_final or has_qnt_aves_caixa_final
            has_mort_aves = "mortalidade_transbordo_aves" in cols_all
            has_mort_kg = "mortalidade_transbordo_kg" in cols_all
            has_obs_transbordo = "obs_transbordo" in cols_all

            select_cols = [
                "motorista",
                "veiculo",
                "equipe",
                "nf_numero", "nf_kg", "nf_caixas", "nf_kg_carregado", "nf_kg_vendido", "nf_saldo",
                "km_inicial", "km_final", "litros", "km_rodado", "media_km_l", "custo_km",
                "ced_200_qtd", "ced_100_qtd", "ced_50_qtd", "ced_20_qtd", "ced_10_qtd", "ced_5_qtd", "ced_2_qtd",
                "valor_dinheiro",
            ]
            if has_nf_preco:
                select_cols.append("nf_preco")
            if col_adiant:
                select_cols.append(col_adiant)
            if has_data_saida:
                select_cols.append("data_saida")
            if has_hora_saida:
                select_cols.append("hora_saida")
            if has_data_chegada:
                select_cols.append("data_chegada")
            if has_hora_chegada:
                select_cols.append("hora_chegada")
            if has_media_carregada:
                select_cols.append("media")
            if has_aves_caixa_final and has_qnt_aves_caixa_final:
                select_cols.append("COALESCE(aves_caixa_final, qnt_aves_caixa_final, 0) as caixa_final")
            elif has_aves_caixa_final:
                select_cols.append("COALESCE(aves_caixa_final, 0) as caixa_final")
            elif has_qnt_aves_caixa_final:
                select_cols.append("COALESCE(qnt_aves_caixa_final, 0) as caixa_final")
            if has_rota_obs:
                select_cols.append("rota_observacao")
            if has_mort_aves:
                select_cols.append("COALESCE(mortalidade_transbordo_aves,0) as mortalidade_transbordo_aves")
            if has_mort_kg:
                select_cols.append("COALESCE(mortalidade_transbordo_kg,0) as mortalidade_transbordo_kg")
            if has_obs_transbordo:
                select_cols.append("COALESCE(obs_transbordo,'') as obs_transbordo")

            try:
                cur.execute(f"""
                    SELECT {', '.join(select_cols)}
                    FROM programacoes
                    WHERE codigo_programacao=?
                """, (prog,))
                row = cur.fetchone()
            except Exception as e:
                self.logger.error(f"Erro ao carregar extras: {e}")
                row = None

        if not row:
            self.lbl_motorista.config(text="Motorista: - | Veiculo: - | Equipe: -")
            if hasattr(self, "lbl_km_anterior"):
                self.lbl_km_anterior.config(text="Ultimo KM do veiculo: -")
            return

        idx = 0
        motorista = row[idx]; idx += 1
        veiculo = row[idx]; idx += 1
        equipe = row[idx]; idx += 1
        nf_numero = row[idx]; idx += 1
        nf_kg = row[idx]; idx += 1
        nf_caixas = row[idx]; idx += 1
        nf_kg_carregado = row[idx]; idx += 1
        nf_kg_vendido = row[idx]; idx += 1
        nf_saldo = row[idx]; idx += 1
        km_inicial = row[idx]; idx += 1
        km_final = row[idx]; idx += 1
        litros = row[idx]; idx += 1
        km_rodado = row[idx]; idx += 1
        media_km_l = row[idx]; idx += 1
        custo_km = row[idx]; idx += 1
        q200 = row[idx]; idx += 1
        q100 = row[idx]; idx += 1
        q50 = row[idx]; idx += 1
        q20 = row[idx]; idx += 1
        q10 = row[idx]; idx += 1
        q5 = row[idx]; idx += 1
        q2 = row[idx]; idx += 1
        valor_dinheiro = row[idx]; idx += 1
        nf_preco = 0.0
        if has_nf_preco:
            nf_preco = row[idx]; idx += 1

        if col_adiant:
            adiant_val = row[idx]; idx += 1
        else:
            adiant_val = 0

        self._data_saida = ""
        self._hora_saida = ""
        self._data_chegada = ""
        self._hora_chegada = ""
        media_carregada = 0.0
        caixa_final = 0
        if has_data_saida:
            self._data_saida = row[idx] or ""; idx += 1
        if has_hora_saida:
            self._hora_saida = row[idx] or ""; idx += 1
        if has_data_chegada:
            self._data_chegada = row[idx] or ""; idx += 1
        if has_hora_chegada:
            self._hora_chegada = row[idx] or ""; idx += 1
        if has_media_carregada:
            media_carregada = self._normalize_media_kg_ave(row[idx]); idx += 1
        if has_caixa_final:
            caixa_final = safe_int(row[idx], 0); idx += 1
        rota_obs = ""
        if has_rota_obs:
            rota_obs = row[idx] or ""; idx += 1
        mort_aves_transb = 0
        mort_kg_transb = 0.0
        obs_transb = ""
        if has_mort_aves:
            mort_aves_transb = safe_int(row[idx], 0); idx += 1
        if has_mort_kg:
            mort_kg_transb = safe_float(row[idx], 0.0); idx += 1
        if has_obs_transbordo:
            obs_transb = str(row[idx] or ""); idx += 1

        self._nf_mort_aves = mort_aves_transb
        self._nf_mort_kg = mort_kg_transb
        self._nf_obs_transbordo = obs_transb

        self._data_saida, self._hora_saida = normalize_date_time_components(self._data_saida, self._hora_saida)
        self._data_chegada, self._hora_chegada = normalize_date_time_components(self._data_chegada, self._hora_chegada)

        equipe_nomes = resolve_equipe_nomes(equipe)
        self.lbl_motorista.config(
            text=fix_mojibake_text(
                f"Motorista: {motorista or '-'} | Veiculo: {veiculo or '-'} | Equipe: {equipe_nomes or '-'}"
            )
        )

        self.ent_adiantamento.delete(0, "end")
        self.ent_adiantamento.insert(0, f"{safe_float(adiant_val, 0.0):.2f}".replace(".", ","))

        # Preenche campos de NF / KM da programacao carregada
        self._set_ent(self.ent_nf_numero, nf_numero)
        self._set_ent(self.ent_nf_kg, nf_kg)
        if hasattr(self, "ent_nf_preco"):
            self._set_ent(self.ent_nf_preco, nf_preco)
        self._set_ent(self.ent_nf_caixas, nf_caixas)
        self._set_ent(self.ent_nf_kg_carregado, nf_kg_carregado)
        self._set_ent(self.ent_nf_kg_vendido, nf_kg_vendido)
        self._set_ent(self.ent_nf_saldo, nf_saldo)
        if hasattr(self, "ent_nf_media_carregada"):
            self._set_ent(self.ent_nf_media_carregada, f"{self._normalize_media_kg_ave(media_carregada):.3f}")
        if hasattr(self, "ent_nf_caixa_final"):
            self._set_ent(self.ent_nf_caixa_final, safe_int(caixa_final, 0))

        # KM inicial: usa o valor da programacao; se vazio/0, puxa ultimo KM final do veiculo
        km_anterior = self._buscar_ultimo_km_final_veiculo(veiculo, prog)
        if hasattr(self, "lbl_km_anterior"):
            if km_anterior > 0:
                self.lbl_km_anterior.config(text=f"Ultimo KM do veiculo: {km_anterior:.1f}")
            else:
                self.lbl_km_anterior.config(text="Ultimo KM do veiculo: sem historico")

        km_inicial_eff = safe_float(km_inicial, 0.0)
        if km_inicial_eff <= 0:
            if km_anterior > 0:
                km_inicial_eff = km_anterior
                self.set_status(
                    f"STATUS: KM inicial sugerido pelo ultimo KM final do veiculo {veiculo}: {km_anterior:.1f}"
                )

        self._set_ent(self.ent_km_inicial, f"{km_inicial_eff:.1f}" if km_inicial_eff > 0 else "")
        self._set_ent(self.ent_km_final, km_final)
        self._set_ent(self.ent_litros, litros)
        self._set_ent(self.ent_km_rodado, km_rodado)
        self._set_ent(self.ent_media, media_km_l)
        self._set_ent(self.ent_custo_km, custo_km)

        if hasattr(self, "txt_rota_obs") and self.txt_rota_obs:
            try:
                self.txt_rota_obs.delete("1.0", "end")
                self.txt_rota_obs.insert("1.0", str(rota_obs or ""))
            except Exception:
                logging.debug("Falha ignorada")

        try:
            self._calcular_km_media()
        except Exception:
            logging.debug("Falha ignorada")
        try:
            self._calcular_saldo_auto()
        except Exception:
            logging.debug("Falha ignorada")

        self._refresh_km_insights()

    def _safe_set_entry(self, entry, value: str, readonly_back: bool = None):
        try:
            prev = entry.cget("state")
        except Exception:
            prev = None

        try:
            try:
                entry.configure(state="normal")
            except Exception:
                logging.debug("Falha ignorada")
            entry.delete(0, "end")
            entry.insert(0, value if value is not None else "")
        finally:
            try:
                if readonly_back is None:
                    if prev is not None:
                        entry.configure(state=prev)
                else:
                    entry.configure(state=("readonly" if readonly_back else "normal"))
            except Exception:
                logging.debug("Falha ignorada")

    def _set_ent(self, ent, val):
        self._safe_set_entry(ent, str(val if val is not None else ""))

    def _normalize_media_kg_ave(self, v) -> float:
        m = safe_float(v, 0.0)
        # Corrige escalas indevidas (ex.: 3455000 -> 3.455, 3455 -> 3.455).
        while m > 50:
            m = m / 1000.0
        return m

    def _buscar_ultimo_km_final_veiculo(self, veiculo: str, prog_atual: str = "") -> float:
        v = upper(veiculo or "").strip()
        if not v:
            return 0.0
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/veiculos/{urllib.parse.quote(v)}/ultimo-km-final?exclude_programacao={urllib.parse.quote(upper(prog_atual or ''))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    return safe_float(resp.get("km_final"), 0.0)
            except Exception:
                logging.debug("Falha ao buscar ultimo km via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    SELECT COALESCE(km_final, 0)
                    FROM programacoes
                    WHERE UPPER(COALESCE(veiculo,'')) = UPPER(?)
                      AND COALESCE(km_final, 0) > 0
                      AND UPPER(COALESCE(codigo_programacao,'')) <> UPPER(?)
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (v, upper(prog_atual or "")),
                )
                row = cur.fetchone()
                return safe_float((row[0] if row else 0.0), 0.0)
        except Exception:
            logging.debug("Falha ignorada")
            return 0.0

    def _calc_valor_dinheiro(self):
        total = 0.0
        for ced, ent in self.ced_entries.items():
            try:
                qtd = int(str(ent.get() or "0").strip())
            except Exception:
                qtd = 0
            linha_total = qtd * float(ced)
            total += linha_total
            self.ced_totals[ced].config(text=self._fmt_money(linha_total))

        self.lbl_valor_dinheiro.config(text=self._fmt_money(total))
        self._update_resumo_financeiro()

    def _update_resumo_financeiro(self):
        prog = self._current_programacao
        if not prog:
            self._reset_resumo()
            return

        total_desp = 0.0
        total_receb = 0.0
        loaded_api = False
        bundle = self._api_bundle_prestacao(prog)
        if bundle:
            rec_rows = bundle.get("recebimentos") if isinstance(bundle, dict) else []
            desp_rows = bundle.get("despesas") if isinstance(bundle, dict) else []
            total_receb = sum(
                safe_float((r or {}).get("valor"), 0.0)
                for r in (rec_rows or [])
                if isinstance(r, dict)
            )
            total_desp = sum(
                safe_float((r or {}).get("valor"), 0.0)
                for r in (desp_rows or [])
                if isinstance(r, dict)
            )
            loaded_api = True

        if not loaded_api:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT SUM(valor) FROM despesas WHERE codigo_programacao=?", (prog,))
                row = cur.fetchone()
                total_desp = float(row[0]) if row and row[0] else 0.0

                try:
                    cur.execute("SELECT SUM(valor) FROM recebimentos WHERE codigo_programacao=?", (prog,))
                    row = cur.fetchone()
                    total_receb = float(row[0]) if row and row[0] else 0.0
                except Exception:
                    total_receb = 0.0

        ced_total = 0.0
        for ced, ent in self.ced_entries.items():
            try:
                qtd = int(str(ent.get() or "0").strip())
            except Exception:
                qtd = 0
            ced_total += qtd * float(ced)

        try:
            adiant = self._parse_money(self.ent_adiantamento.get())
        except Exception:
            adiant = 0.0

        total_entradas = total_receb + adiant
        total_saidas = total_desp + ced_total
        valor_final_caixa = total_entradas - total_desp
        diferenca = valor_final_caixa - ced_total
        resultado_liquido = total_entradas - total_saidas

        self.lbl_receb_total.config(text=self._fmt_money(total_receb))
        self.lbl_desp_total.config(text=self._fmt_money(total_desp))
        self.lbl_cedulas_total.config(text=self._fmt_money(ced_total))
        self.lbl_total_entradas.config(text=self._fmt_money(total_entradas))
        self.lbl_total_saidas.config(text=self._fmt_money(total_saidas))
        self.lbl_valor_final_caixa.config(text=self._fmt_money(valor_final_caixa))
        self.lbl_diferenca.config(text=self._fmt_money(diferenca))
        self.lbl_resultado_liquido.config(text=self._fmt_money(resultado_liquido))
        if hasattr(self, "lbl_kpi_receb_total"):
            self.lbl_kpi_receb_total.config(text=self._fmt_money(total_receb))
        if hasattr(self, "lbl_kpi_desp_total"):
            self.lbl_kpi_desp_total.config(text=self._fmt_money(total_desp))
        if hasattr(self, "lbl_kpi_cedulas_total"):
            self.lbl_kpi_cedulas_total.config(text=self._fmt_money(ced_total))
        if hasattr(self, "lbl_kpi_resultado_liquido"):
            self.lbl_kpi_resultado_liquido.config(text=self._fmt_money(resultado_liquido))

        self.lbl_resultado_liquido.config(foreground="#2E7D32" if resultado_liquido >= 0 else "#C62828")
        if hasattr(self, "lbl_kpi_resultado_liquido"):
            self.lbl_kpi_resultado_liquido.config(foreground="#2E7D32" if resultado_liquido >= 0 else "#C62828")

    def _reset_resumo(self):
        self.lbl_receb_total.config(text="R$ 0,00")
        self.lbl_desp_total.config(text="R$ 0,00")
        self.lbl_cedulas_total.config(text="R$ 0,00")
        self.lbl_total_entradas.config(text="R$ 0,00")
        self.lbl_total_saidas.config(text="R$ 0,00")
        self.lbl_valor_final_caixa.config(text="R$ 0,00")
        self.lbl_diferenca.config(text="R$ 0,00")
        self.lbl_resultado_liquido.config(text="R$ 0,00")
        if hasattr(self, "lbl_kpi_receb_total"):
            self.lbl_kpi_receb_total.config(text="R$ 0,00")
        if hasattr(self, "lbl_kpi_desp_total"):
            self.lbl_kpi_desp_total.config(text="R$ 0,00")
        if hasattr(self, "lbl_kpi_cedulas_total"):
            self.lbl_kpi_cedulas_total.config(text="R$ 0,00")
        if hasattr(self, "lbl_kpi_resultado_liquido"):
            self.lbl_kpi_resultado_liquido.config(text="R$ 0,00", foreground="#2E7D32")

    # =========================================================
    # 7.3 EVENTOS / FILTROS / ORDENAÇÃO
    # =========================================================
    def _on_select_despesa(self, event=None):
        sel = self.tree_desp.selection()
        if not sel:
            self.selected_despesa_id = None
            return
        vals = self.tree_desp.item(sel[0], "values")
        self.selected_despesa_id = vals[0] if vals else None

    def _aplicar_filtro_mes(self):
        return

    def _sort_despesas_by_column(self, col):
        # estado simples (mantém comportamento previsÃÂÂvel)
        if not hasattr(self, "_sort_state"):
            self._sort_state = {"col": None, "reverse": False}

        if self._sort_state["col"] == col:
            self._sort_state["reverse"] = not self._sort_state["reverse"]
        else:
            self._sort_state["col"] = col
            self._sort_state["reverse"] = False

        reverse = self._sort_state["reverse"]
        data = [(self.tree_desp.set(child, col), child) for child in self.tree_desp.get_children("")]

        if col == "VALOR":
            data.sort(key=lambda x: self._parse_money(x[0]), reverse=reverse)
        elif col == "DATA":
            data.sort(key=lambda x: str(x[0] or ""), reverse=reverse)
            data.sort(key=lambda x: safe_int(x[0], 0), reverse=reverse)
        else:
            data.sort(key=lambda x: str(x[0] or "").upper(), reverse=reverse)

        for idx, (_, child) in enumerate(data):
            self.tree_desp.move(child, "", idx)

        # seta setinha no cabeçalho (igual seu padrão em outras telas)
        arrow = " ââââ€š¬ âââ€š¬Åâ€œ" if reverse else " ââââ€š¬ âââ€š¬ËÅ“"
        for c in self.tree_desp["columns"]:
            txt = self.tree_desp.heading(c)["text"]
            if txt.endswith(" ââââ€š¬ âââ€š¬ËÅ“") or txt.endswith(" ââââ€š¬ âââ€š¬Åâ€œ"):
                txt = txt[:-2]
            self.tree_desp.heading(c, text=txt)

        txt = self.tree_desp.heading(col)["text"]
        if txt.endswith(" ââââ€š¬ âââ€š¬ËÅ“") or txt.endswith(" ââââ€š¬ âââ€š¬Åâ€œ"):
            txt = txt[:-2]
        self.tree_desp.heading(col, text=txt + arrow)

    def _limpar_busca_despesas(self):
        if not hasattr(self, "ent_busca_despesa"):
            return
        self.ent_busca_despesa.delete(0, "end")
        if hasattr(self, "cb_filtro_categoria"):
            self.cb_filtro_categoria.set("TODAS")
        if hasattr(self, "cb_ordem_despesa"):
            self.cb_ordem_despesa.set("DATA DESC")
        if self._current_programacao:
            self._load_despesas(self._current_programacao)
        self.set_status("STATUS: Busca limpa")

    # =========================================================
    # 7.4 CÃÆââ‚¬â„¢ÂLCULOS AUTOMÃÆââ‚¬â„¢ÂTICOS / SINCRONIZAÃÆââ‚¬â„¢ââââ‚¬Å¡¬¡ÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢O SIMULADA
    # =========================================================
    def _calcular_saldo_auto(self):
        try:
            prog = self._current_programacao or ""
            if not prog:
                return

            def _prog_cols(cur):
                try:
                    cur.execute("PRAGMA table_info(programacoes)")
                    return {str(r[1]).lower() for r in (cur.fetchall() or [])}
                except Exception:
                    return set()

            def _fetch_prog_first_numeric(cur, cols_set, candidates, default=0.0):
                for c in candidates:
                    if c.lower() not in cols_set:
                        continue
                    try:
                        cur.execute(
                            f"SELECT COALESCE({c}, 0) FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                            (prog,),
                        )
                        row = cur.fetchone()
                        if row:
                            v = safe_float(row[0], 0.0)
                            if v > 0:
                                return v
                    except Exception:
                        logging.debug("Falha ignorada")
                return default

            caixas = safe_int(self.ent_nf_caixas.get(), 0)
            media = self._normalize_media_kg_ave(getattr(self, "ent_nf_media_carregada", self.ent_nf_saldo).get())
            caixa_final = safe_int(getattr(self, "ent_nf_caixa_final", self.ent_nf_caixas).get(), 0)

            aves_por_caixa = 6
            kg_vendido = safe_float(self.ent_nf_kg_vendido.get(), 0.0)
            used_api = False
            try:
                bundle = self._api_bundle_prestacao(prog)
                rota = bundle.get("rota") if isinstance(bundle, dict) else None
                clientes = bundle.get("clientes") if isinstance(bundle, dict) else []
                if bundle and isinstance(rota, dict):
                    if caixas <= 0:
                        caixas = safe_int(
                            rota.get("nf_caixas"),
                            safe_int(rota.get("caixas_carregadas"), safe_int(rota.get("qnt_cx_carregada"), safe_int(rota.get("total_caixas"), 0))),
                        )
                    if media <= 0:
                        media = self._normalize_media_kg_ave(
                            safe_float(rota.get("media"), safe_float(rota.get("media_carregada"), safe_float(rota.get("media_base"), 0.0)))
                        )
                    if caixa_final <= 0:
                        caixa_final = safe_int(
                            rota.get("aves_caixa_final"),
                            safe_int(rota.get("qnt_aves_caixa_final"), 0),
                        )
                    aves_por_caixa = safe_int(
                        rota.get("qnt_aves_por_cx"),
                        safe_int(rota.get("aves_por_caixa"), safe_int(rota.get("qnt_aves_por_caixa"), 6)),
                    )
                    if aves_por_caixa <= 0:
                        aves_por_caixa = 6

                    kg_vendido_calc = 0.0
                    for it in (clientes or []):
                        if not isinstance(it, dict):
                            continue
                        st = upper(str(it.get("status_pedido") or ""))
                        if st != "ENTREGUE":
                            continue
                        peso = safe_float(it.get("peso_previsto"), 0.0)
                        if peso > 0:
                            kg_vendido_calc += peso
                            continue
                        cx = safe_float(it.get("caixas_atual"), safe_float(it.get("qnt_caixas"), 0.0))
                        if cx > 0:
                            kg_vendido_calc += cx * aves_por_caixa * media
                    if kg_vendido_calc > 0:
                        kg_vendido = kg_vendido_calc
                    used_api = True
            except Exception:
                logging.debug("Falha ao calcular saldo via API; usando fallback local.", exc_info=True)

            if not used_api:
                with get_db() as conn:
                    cur = conn.cursor()
                    cols_set = _prog_cols(cur)

                    if caixas <= 0:
                        caixas = safe_int(
                            _fetch_prog_first_numeric(
                                cur,
                                cols_set,
                                ["nf_caixas", "caixas_carregadas", "qnt_cx_carregada", "total_caixas"],
                                0.0,
                            ),
                            0,
                        )
                    if media <= 0:
                        media = self._normalize_media_kg_ave(
                            _fetch_prog_first_numeric(cur, cols_set, ["media", "media_carregada", "media_base"], 0.0)
                        )
                    if caixa_final <= 0:
                        caixa_final = safe_int(
                            _fetch_prog_first_numeric(cur, cols_set, ["aves_caixa_final", "qnt_aves_caixa_final"], 0.0),
                            0,
                        )
                    aves_por_caixa = safe_int(
                        _fetch_prog_first_numeric(cur, cols_set, ["qnt_aves_por_cx", "aves_por_caixa", "qnt_aves_por_caixa"], 6.0),
                        6,
                    )
                    if aves_por_caixa <= 0:
                        aves_por_caixa = 6

                    # KG vendido = soma dos pedidos ENTREGUE (peso_previsto quando existir)
                    try:
                        cur.execute("PRAGMA table_info(programacao_itens_controle)")
                        ctrl_cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    except Exception:
                        ctrl_cols = set()
                    try:
                        cur.execute("PRAGMA table_info(programacao_itens)")
                        itens_cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    except Exception:
                        itens_cols = set()

                    if "codigo_programacao" in ctrl_cols and "status_pedido" in ctrl_cols:
                        peso_expr = "COALESCE(peso_previsto,0)" if "peso_previsto" in ctrl_cols else "0"
                        caixas_expr = "COALESCE(caixas_atual,0)" if "caixas_atual" in ctrl_cols else "0"
                        cur.execute(
                            f"""
                            SELECT COALESCE(SUM(
                                CASE
                                    WHEN UPPER(COALESCE(status_pedido,''))='ENTREGUE' THEN
                                        CASE
                                            WHEN {peso_expr} > 0 THEN {peso_expr}
                                            WHEN {caixas_expr} > 0 THEN ({caixas_expr} * ? * ?)
                                            ELSE 0
                                        END
                                    ELSE 0
                                END
                            ),0)
                            FROM programacao_itens_controle
                            WHERE codigo_programacao=?
                            """,
                            (aves_por_caixa, media, prog),
                        )
                        row = cur.fetchone()
                        kg_vendido = safe_float((row[0] if row else 0), 0.0)

                    # fallback: sem controle de peso, estima por itens entregues da programação
                    if kg_vendido <= 0 and "codigo_programacao" in itens_cols:
                        col_status = "status_pedido" if "status_pedido" in itens_cols else None
                        col_cx = None
                        for c in ("qnt_caixas", "qtd_caixas", "caixas", "cx"):
                            if c in itens_cols:
                                col_cx = c
                                break
                        if col_status and col_cx:
                            cur.execute(
                                f"""
                                SELECT COALESCE(SUM(
                                    CASE WHEN UPPER(COALESCE({col_status},''))='ENTREGUE'
                                         THEN COALESCE({col_cx},0) ELSE 0 END
                                ),0)
                                FROM programacao_itens
                                WHERE codigo_programacao=?
                                """,
                                (prog,),
                            )
                            row = cur.fetchone()
                            cx_entregues = safe_float((row[0] if row else 0), 0.0)
                            kg_vendido = cx_entregues * aves_por_caixa * media

            # KG carregado = caixas/aves/media (considera caixa final)
            total_aves = 0
            if caixas > 0 and aves_por_caixa > 0:
                if caixa_final > 0:
                    cx_cheias = max(caixas - 1, 0)
                    total_aves = (cx_cheias * aves_por_caixa) + max(min(caixa_final, aves_por_caixa), 0)
                else:
                    total_aves = caixas * aves_por_caixa
            carregado = (total_aves * media) if (total_aves > 0 and media > 0) else safe_float(self.ent_nf_kg_carregado.get(), 0.0)

            # fallback prático: quando não há entregas registradas ainda, vendido tende ao carregado no fechamento
            if kg_vendido <= 0 and carregado > 0:
                kg_vendido = carregado

            saldo = carregado - kg_vendido

            self._safe_set_entry(self.ent_nf_caixas, str(caixas))
            self._safe_set_entry(self.ent_nf_kg_carregado, f"{carregado:.2f}")
            self._safe_set_entry(self.ent_nf_kg_vendido, f"{kg_vendido:.2f}")
            self._safe_set_entry(self.ent_nf_saldo, f"{saldo:.2f}")
            if hasattr(self, "ent_nf_media_carregada"):
                self._safe_set_entry(self.ent_nf_media_carregada, f"{media:.3f}")
            if hasattr(self, "ent_nf_caixa_final"):
                self._safe_set_entry(self.ent_nf_caixa_final, str(safe_int(caixa_final, 0)))
            self._refresh_nf_trade_summary()

            self.set_status(
                f"STATUS: NF recalculada | Caixas: {caixas} | KG carregado: {carregado:.2f} | "
                f"KG vendido: {kg_vendido:.2f} | Saldo: {saldo:.2f}"
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao calcular saldo: {str(e)}")

    def _bind_km_auto_calc(self):
        for ent in (
            getattr(self, "ent_km_inicial", None),
            getattr(self, "ent_km_final", None),
            getattr(self, "ent_litros", None),
        ):
            if not ent:
                continue
            ent.bind("<KeyRelease>", lambda _e: self._calcular_km_media(), add="+")
            ent.bind("<FocusOut>", lambda _e: self._calcular_km_media(), add="+")

    def _bind_nf_summary_auto_calc(self):
        for ent in (
            getattr(self, "ent_nf_kg", None),
            getattr(self, "ent_nf_preco", None),
            getattr(self, "ent_nf_kg_vendido", None),
            getattr(self, "ent_nf_media_carregada", None),
            getattr(self, "ent_nf_caixa_final", None),
        ):
            if not ent:
                continue
            ent.bind("<KeyRelease>", lambda _e: self._refresh_nf_trade_summary(), add="+")
            ent.bind("<FocusOut>", lambda _e: self._refresh_nf_trade_summary(), add="+")

    def _calc_preco_medio_venda(self, prog: str) -> float:
        if not prog:
            return 0.0
        try:
            bundle = self._api_bundle_prestacao(prog)
            if bundle:
                clientes = bundle.get("clientes") if isinstance(bundle, dict) else []
                precos = []
                for r in (clientes or []):
                    if not isinstance(r, dict):
                        continue
                    p = safe_float((r or {}).get("preco_atual"), safe_float((r or {}).get("preco"), 0.0))
                    if p > 0:
                        precos.append(p)
                if precos:
                    return sum(precos) / float(len(precos))
                return 0.0
        except Exception:
            logging.debug("Falha ao calcular preco medio via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacao_itens)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                col_preco = None
                for c in ("preco", "preco_unitario", "preco_venda", "valor_unitario"):
                    if c in cols:
                        col_preco = c
                        break
                if not col_preco:
                    return 0.0
                cur.execute(
                    f"""
                    SELECT COALESCE(AVG(COALESCE({col_preco}, 0)), 0)
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                      AND COALESCE({col_preco}, 0) > 0
                    """,
                    (prog,),
                )
                row = cur.fetchone()
                return safe_float((row[0] if row else 0.0), 0.0)
        except Exception:
            logging.debug("Falha ignorada")
            return 0.0

    def _refresh_nf_trade_summary(self):
        labels = (
            "lbl_nf_compra_kg",
            "lbl_nf_compra_preco",
            "lbl_nf_compra_total",
            "lbl_nf_venda_preco",
            "lbl_nf_venda_kg",
            "lbl_nf_venda_total",
            "lbl_nf_desp_total",
            "lbl_nf_lucro_bruto",
            "lbl_nf_lucro_liquido",
            "lbl_nf_margem",
        )
        if not all(hasattr(self, x) for x in labels):
            return

        nf_kg = safe_float(getattr(self, "ent_nf_kg", None).get() if getattr(self, "ent_nf_kg", None) else 0, 0.0)
        nf_preco = safe_float(getattr(self, "ent_nf_preco", None).get() if getattr(self, "ent_nf_preco", None) else 0, 0.0)
        kg_vendido = safe_float(getattr(self, "ent_nf_kg_vendido", None).get() if getattr(self, "ent_nf_kg_vendido", None) else 0, 0.0)
        mort_aves = safe_int(getattr(self, "_nf_mort_aves", 0), 0)
        mort_kg = safe_float(getattr(self, "_nf_mort_kg", 0.0), 0.0)
        obs_transb = str(getattr(self, "_nf_obs_transbordo", "") or "").strip()
        kg_nf_util = max(nf_kg - mort_kg, 0.0)

        compra_total = nf_kg * nf_preco
        preco_medio_venda = self._calc_preco_medio_venda(self._current_programacao or "")
        receita_estimada = kg_vendido * preco_medio_venda

        despesas_rota = 0.0
        try:
            prog = self._current_programacao or ""
            if prog:
                loaded_api = False
                try:
                    bundle = self._api_bundle_prestacao(prog)
                    arr = (bundle.get("despesas") if isinstance(bundle, dict) else []) or []
                    if bundle:
                        despesas_rota = sum(
                            safe_float((r or {}).get("valor"), 0.0)
                            for r in arr
                            if isinstance(r, dict)
                        )
                        loaded_api = True
                except Exception:
                    logging.debug("Falha ao calcular despesas da rota via API; usando fallback local.", exc_info=True)
                if not loaded_api:
                    with get_db() as conn:
                        cur = conn.cursor()
                        cur.execute(
                            "SELECT COALESCE(SUM(valor), 0) FROM despesas WHERE codigo_programacao=?",
                            (prog,),
                        )
                        row = cur.fetchone()
                        despesas_rota = safe_float((row[0] if row else 0.0), 0.0)
        except Exception:
            logging.debug("Falha ignorada")

        lucro_bruto = receita_estimada - compra_total
        lucro_liquido = lucro_bruto - despesas_rota
        margem = (lucro_liquido / receita_estimada * 100.0) if receita_estimada > 0 else 0.0

        self.lbl_nf_compra_kg.config(text=f"KG NF: {nf_kg:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        self.lbl_nf_compra_preco.config(text=f"Preco NF: {self._fmt_money(nf_preco)}")
        self.lbl_nf_compra_total.config(text=f"Total compra: {self._fmt_money(compra_total)}")

        self.lbl_nf_venda_preco.config(text=f"Preco medio venda: {self._fmt_money(preco_medio_venda)}")
        self.lbl_nf_venda_kg.config(text=f"KG vendido: {kg_vendido:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        self.lbl_nf_venda_total.config(text=f"Receita estimada: {self._fmt_money(receita_estimada)}")

        self.lbl_nf_desp_total.config(text=f"Despesas da rota: {self._fmt_money(despesas_rota)}")
        self.lbl_nf_lucro_bruto.config(text=f"Lucro bruto: {self._fmt_money(lucro_bruto)}")
        self.lbl_nf_lucro_liquido.config(text=f"Lucro liquido: {self._fmt_money(lucro_liquido)}")
        if hasattr(self, "lbl_nf_mort_aves"):
            self.lbl_nf_mort_aves.config(text=f"Mortalidade (aves): {mort_aves}")
        if hasattr(self, "lbl_nf_mort_kg"):
            self.lbl_nf_mort_kg.config(
                text=f"Mortalidade (KG): {mort_kg:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
        if hasattr(self, "lbl_nf_kg_util"):
            self.lbl_nf_kg_util.config(
                text=f"KG útil NF (NF - mortalidade): {kg_nf_util:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
        if hasattr(self, "lbl_nf_obs_transb"):
            self.lbl_nf_obs_transb.config(text=f"Obs transbordo: {obs_transb or '-'}")

        if receita_estimada <= 0:
            cor_margem = "#6B7280"   # cinza (sem base de receita)
            faixa = "SEM BASE"
        elif margem < 0:
            cor_margem = "#C62828"   # vermelho (prejuizo)
            faixa = "CRITICA"
        elif margem < 8:
            cor_margem = "#EA580C"   # laranja (margem baixa)
            faixa = "BAIXA"
        elif margem < 15:
            cor_margem = "#B45309"   # amarelo escuro (atenção)
            faixa = "ATENCAO"
        else:
            cor_margem = "#2E7D32"   # verde (margem saudavel)
            faixa = "BOA"

        self.lbl_nf_margem.config(
            text=f"Margem liquida: {margem:.2f}% ({faixa})".replace(".", ",")
        )
        cor_lucro = "#2E7D32" if lucro_liquido >= 0 else "#C62828"
        try:
            self.lbl_nf_lucro_bruto.config(foreground=cor_lucro)
            self.lbl_nf_lucro_liquido.config(foreground=cor_lucro)
            self.lbl_nf_margem.config(foreground=cor_margem)
        except Exception:
            logging.debug("Falha ignorada")

    def _calcular_km_media(self):
        try:
            km_inicial = safe_float(self.ent_km_inicial.get(), 0.0)
            km_final = safe_float(self.ent_km_final.get(), 0.0)
            litros = safe_float(self.ent_litros.get(), 0.0)

            km_rodado = km_final - km_inicial
            if km_rodado < 0:
                km_rodado = 0.0

            media = (km_rodado / litros) if litros > 0 else 0.0

            total_despesas = 0.0
            try:
                if self._current_programacao:
                    loaded_api = False
                    try:
                        bundle = self._api_bundle_prestacao(self._current_programacao)
                        arr = (bundle.get("despesas") if isinstance(bundle, dict) else []) or []
                        if bundle:
                            total_despesas = sum(
                                safe_float((r or {}).get("valor"), 0.0)
                                for r in arr
                                if isinstance(r, dict)
                            )
                            loaded_api = True
                    except Exception:
                        logging.debug("Falha ao calcular total despesas via API; usando fallback local.", exc_info=True)
                    if not loaded_api:
                        with get_db() as conn:
                            cur = conn.cursor()
                            cur.execute(
                                "SELECT COALESCE(SUM(valor), 0) FROM despesas WHERE codigo_programacao=?",
                                (self._current_programacao,),
                            )
                            row = cur.fetchone()
                            total_despesas = safe_float((row[0] if row else 0), 0.0)
            except Exception:
                logging.debug("Falha ignorada")

            # custo por km baseado nas despesas da viagem / km rodado
            custo_km = (total_despesas / km_rodado) if km_rodado > 0 else 0.0

            self._safe_set_entry(self.ent_km_rodado, f"{km_rodado:.1f}")
            self._safe_set_entry(self.ent_media, f"{media:.1f}")
            self._safe_set_entry(self.ent_custo_km, f"{custo_km:.2f}")

            self.set_status(
                f"STATUS: KM rodado: {km_rodado:.1f} km | Média: {media:.1f} km/l | "
                f"Custo/km: {custo_km:.2f}"
            )
            self._refresh_km_insights()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao calcular KM/média: {str(e)}")

    def _distribuir_cedulas(self):
        try:
            valor_total = self._parse_money(self.ent_total_dinheiro.get())
            if valor_total <= 0:
                messagebox.showwarning("Atenção", "Digite um valor total válido.")
                return

            cedulas = [200, 100, 50, 20, 10, 5, 2]
            restante = float(valor_total)

            for ced in cedulas:
                self.ced_entries[ced].delete(0, "end")
                self.ced_entries[ced].insert(0, "0")

            for ced in cedulas:
                if restante >= ced:
                    quantidade = int(restante // ced)
                    self.ced_entries[ced].delete(0, "end")
                    self.ced_entries[ced].insert(0, str(quantidade))
                    restante -= quantidade * ced

            # arredondamento: se sobrar qualquer centavo, ajusta na menor cédula
            if restante > 0:
                atual = safe_int(self.ced_entries[2].get(), 0)
                self.ced_entries[2].delete(0, "end")
                self.ced_entries[2].insert(0, str(atual + 1))

            self._calc_valor_dinheiro()
            self.set_status(f"STATUS: Valor {self._fmt_money(valor_total)} distribuÃÂÂdo automaticamente")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao distribuir cédulas: {str(e)}")

    def _sincronizar_com_app(self):
        if not self._current_programacao:
            messagebox.showwarning("Sincronização", "Selecione a programação primeiro.")
            return

        prog = self._current_programacao
        motor_hint = self._fetch_motorista_codigo_hint(prog)
        creds = self._prompt_sync_credentials(motor_hint)
        if not creds:
            return
        codigo, senha = creds

        try:
            token = self._api_login(codigo, senha)
            resposta = _call_api("GET", f"rotas/{prog}", token=token)
            rota = resposta.get("rota")
            if not rota:
                raise SyncError("A API não retornou os dados da rota.")
            clientes = resposta.get("clientes") or []

            with get_db() as conn:
                cur = conn.cursor()
                self._apply_api_programacao(cur, prog, rota)
                self._apply_api_clientes(cur, prog, clientes)
                conn.commit()

            self._populate_route_entries(rota)
            self._load_by_programacao()
            self._calcular_km_media()
            self._calcular_saldo_auto()

            self.logger.info("Sincronização com App Mobile concluÃÂÂda para %s", prog)
            messagebox.showinfo("Sincronização", "Dados sincronizados com sucesso.")
        except SyncError as exc:
            self.logger.error("Falha ao sincronizar com o App: %s", exc)
            messagebox.showerror("Sincronização", f"Falha ao sincronizar: {exc}")
        except Exception as exc:
            logging.exception("Erro inesperado na sincronização", exc_info=exc)
            messagebox.showerror("Sincronização", f"Erro inesperado: {exc}")

    def _prompt_sync_credentials(self, default_code: str):
        title = "Sincronização App Mobile"
        codigo = simpledialog.askstring(
            title,
            "Código do motorista utilizado no App Mobile:",
            initialvalue=(default_code or ""),
            parent=self,
        )
        if codigo is None:
            return None
        codigo = codigo.strip()
        if not codigo:
            messagebox.showwarning("Sincronização", "Código do motorista é obrigatório.")
            return None

        senha = simpledialog.askstring(
            title,
            f"Senha do motorista {codigo}:",
            show="*",
            parent=self,
        )
        if senha is None:
            return None
        senha = senha.strip()
        if not senha:
            messagebox.showwarning("Sincronização", "Senha do motorista é obrigatória.")
            return None

        return codigo, senha

    def _sync_programacao_from_api_desktop(self, prog: str, silent: bool = True) -> bool:
        """
        Sincroniza rota + clientes usando endpoint desktop protegido por ROTA_SECRET.
        Não exige login de motorista e mantém o fluxo estação<->servidor.
        """
        prog = upper(str(prog or "").strip())
        if not prog:
            return False

        try:
            resposta = self._api_bundle_prestacao(prog) if hasattr(self, "_api_bundle_prestacao") else None
            rota = resposta.get("rota") if isinstance(resposta, dict) else None
            clientes = resposta.get("clientes") if isinstance(resposta, dict) else []
            if not isinstance(rota, dict):
                return False

            with get_db() as conn:
                cur = conn.cursor()
                self._apply_api_programacao(cur, prog, rota)
                self._apply_api_clientes(cur, prog, clientes or [])
                conn.commit()
            return True
        except Exception:
            logging.debug("Falha ao sincronizar programação via endpoint desktop", exc_info=True)
            if not silent:
                messagebox.showwarning(
                    "Sincronizacao",
                    "Nao foi possivel sincronizar agora com a API central.\n"
                    "A tela sera carregada com os dados locais.",
                )
            return False

    def _fetch_motorista_codigo_hint(self, prog: str) -> str:
        try:
            bundle = self._api_bundle_prestacao(upper(prog or "").strip())
            rota = bundle.get("rota") if isinstance(bundle, dict) else None
            if isinstance(rota, dict):
                cod = str(rota.get("motorista_codigo") or rota.get("codigo_motorista") or "").strip()
                if cod:
                    return cod
                return str(rota.get("motorista") or "").strip()
        except Exception:
            logging.debug("Falha ao obter hint de motorista via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    "SELECT motorista_id, motorista FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                    (prog,),
                )
                row = cur.fetchone()
                if not row:
                    return ""
                motorista_id = row["motorista_id"]
                if motorista_id:
                    cur.execute("SELECT codigo FROM motoristas WHERE id=? LIMIT 1", (motorista_id,))
                    mrow = cur.fetchone()
                    if mrow and mrow["codigo"]:
                        return mrow["codigo"]
                return (row["motorista"] or "").strip()
        except Exception:
            return ""

    def _api_login(self, codigo: str, senha: str) -> str:
        if not API_BASE_URL:
            raise SyncError("ROTA_SERVER_URL não configurada.")
        resposta = _call_api("POST", "auth/motorista/login", payload={"codigo": codigo, "senha": senha})
        token = resposta.get("token")
        if not token:
            raise SyncError("Nao foi possivel obter token da API de sincronizacao.")
        return token

    def _apply_api_programacao(self, cur, prog: str, rota: dict):
        def _pick(*keys):
            for k in keys:
                if k in rota and rota.get(k) not in (None, ""):
                    return rota.get(k)
            return None

        nf_numero = _pick("nf_numero", "num_nf", "nf")
        nf_kg = _pick("nf_kg", "kg_nota_fiscal")
        nf_preco = _pick("nf_preco", "preco_nf", "preco_nota_fiscal")
        nf_caixas = _pick("nf_caixas", "caixas_carregadas", "total_caixas")
        nf_kg_carregado = _pick("nf_kg_carregado", "kg_carregado")
        nf_kg_vendido = _pick("nf_kg_vendido", "kg_vendido")
        nf_saldo = _pick("nf_saldo")
        media_carregada = self._normalize_media_kg_ave(_pick("media", "media_carregada", "media_base"))
        caixa_final = _pick("caixa_final", "qnt_aves_caixa_final", "aves_caixa_final")
        data_saida_raw = _pick("data_saida", "saida_data")
        hora_saida_raw = _pick("hora_saida", "saida_hora")
        data_chegada_raw = _pick("data_chegada", "chegada_data")
        hora_chegada_raw = _pick("hora_chegada", "chegada_hora")
        data_saida_n, hora_saida_n = normalize_date_time_components(data_saida_raw, hora_saida_raw)
        data_chegada_n, hora_chegada_n = normalize_date_time_components(data_chegada_raw, hora_chegada_raw)

        # Se a API nao trouxer saldo, calcula com base no carregado.
        if nf_saldo in (None, ""):
            try:
                nf_saldo = round(max(safe_float(nf_kg, 0.0) - safe_float(nf_kg_carregado, 0.0), 0.0), 2)
            except Exception:
                nf_saldo = 0

        valores_por_campo = {
            # status operacional da rota
            "status": _pick("status_operacional", "status"),
            "local_rota": _pick("local_rota", "tipo_rota", "local"),
            "tipo_rota": _pick("tipo_rota", "local_rota", "local"),
            "local": _pick("local_rota", "tipo_rota", "local"),
            "num_nf": nf_numero,
            "nf_numero": nf_numero,
            "nf_kg": nf_kg,
            "nf_preco": nf_preco,
            "nf_caixas": nf_caixas,
            "nf_kg_carregado": nf_kg_carregado,
            "nf_kg_vendido": nf_kg_vendido,
            "nf_saldo": nf_saldo,
            "media": media_carregada,
            "aves_caixa_final": caixa_final,
            "qnt_aves_caixa_final": caixa_final,
            "km_inicial": _pick("km_inicial"),
            "km_final": _pick("km_final"),
            "litros": _pick("litros"),
            "km_rodado": _pick("km_rodado"),
            "media_km_l": _pick("media_km_l"),
            "custo_km": _pick("custo_km"),
            # datas/horas: API mobile usa saida_*/chegada_* e o desktop usa data_*/hora_*
            "data_saida": data_saida_n or "",
            "hora_saida": hora_saida_n or "",
            "data_chegada": data_chegada_n or "",
            "hora_chegada": hora_chegada_n or "",
            # local de carregamento (compatibilidade entre esquemas)
            "local_carregamento": _pick("local_carregado", "local_carregamento", "granja_carregada", "local_carreg"),
            "granja_carregada": _pick("local_carregado", "local_carregamento", "granja_carregada", "local_carreg"),
            "local_carregado": _pick("local_carregado", "local_carregamento", "granja_carregada", "local_carreg"),
            "local_carreg": _pick("local_carregado", "local_carregamento", "granja_carregada", "local_carreg"),
            # espelha dados de carregamento bruto quando vierem da API
            "kg_carregado": _pick("kg_carregado", "nf_kg_carregado"),
            "caixas_carregadas": _pick("caixas_carregadas", "nf_caixas"),
            # FOB/CIF + auditoria
            "tipo_estimativa": _pick("tipo_estimativa"),
            "caixas_estimado": _pick("caixas_estimado"),
            "usuario_criacao": _pick("usuario_criacao", "criado_por", "created_by"),
            "usuario_ultima_edicao": _pick("usuario_ultima_edicao", "alterado_por", "updated_by"),
            # Ocorrências / mortalidade no transbordo (origem APK)
            "mortalidade_transbordo_aves": _pick("aves_mortas_transbordo", "mortalidade_transbordo_aves"),
            "mortalidade_transbordo_kg": _pick("mortalidade_transbordo_kg"),
            "obs_transbordo": _pick("obs_transbordo", "mortalidade_transbordo_obs"),
        }
        status_operacional_api = upper(str(rota.get("status_operacional") or "").strip())
        if status_operacional_api:
            valores_por_campo["status_operacional"] = status_operacional_api
            valores_por_campo["finalizada_no_app"] = 1 if status_operacional_api in {"FINALIZADA", "FINALIZADO"} else 0

        try:
            cur.execute("PRAGMA table_info(programacoes)")
            cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
        except Exception:
            cols_prog = set()

        updates = []
        valores = []
        for campo, valor in valores_por_campo.items():
            if campo.lower() in cols_prog and valor not in (None, ""):
                updates.append(f"{campo}=?")
                valores.append(valor)
        if not updates:
            return
        cur.execute(
            f"UPDATE programacoes SET {', '.join(updates)} WHERE codigo_programacao=?",
            (*valores, prog),
        )

    def _apply_api_clientes(self, cur, prog: str, clientes):
        if not clientes:
            return
        def _norm_pedido(v):
            s = (str(v or "")).strip()
            if not s:
                return ""
            s_num = s.replace(" ", "").replace(",", ".")
            try:
                f = float(s_num)
                if f.is_integer():
                    return str(int(f))
                return ("%f" % f).rstrip("0").rstrip(".")
            except Exception:
                if s.endswith(".0"):
                    base = s[:-2].strip()
                    if base:
                        return base
                return s

        def _table_cols(tbl: str):
            try:
                cur.execute(f"PRAGMA table_info({tbl})")
                return {str(r[1]).lower() for r in (cur.fetchall() or [])}
            except Exception:
                return set()

        cols_ctrl = _table_cols("programacao_itens_controle")
        cols_itens = _table_cols("programacao_itens")
        cols_receb = _table_cols("recebimentos")

        def _registrar_amostra_local(cliente_row: dict, pedido_norm_local: str, cod_cliente_local: str):
            cols_amostras = _table_cols("cliente_localizacao_amostras")
            if not cols_amostras:
                return
            lat_evento = cliente_row.get("lat_evento")
            lon_evento = cliente_row.get("lon_evento")
            endereco_evento = str(cliente_row.get("endereco_evento") or "").strip()
            cidade_evento = str(cliente_row.get("cidade_evento") or "").strip()
            bairro_evento = str(cliente_row.get("bairro_evento") or "").strip()
            has_geo = lat_evento not in (None, "") and lon_evento not in (None, "")
            has_addr = any((endereco_evento, cidade_evento, bairro_evento))
            if not has_geo and not has_addr:
                return
            try:
                cur.execute(
                    """
                    SELECT latitude, longitude, endereco, cidade, bairro, status_pedido
                    FROM cliente_localizacao_amostras
                    WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)
                      AND UPPER(COALESCE(codigo_programacao,''))=UPPER(?)
                      AND COALESCE(TRIM(pedido), '')=COALESCE(TRIM(?), '')
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (cod_cliente_local, prog, pedido_norm_local),
                )
                last = cur.fetchone()
                if last:
                    same_geo = str(last["latitude"] or "") == str(lat_evento or "") and str(last["longitude"] or "") == str(lon_evento or "")
                    same_addr = (
                        str(last["endereco"] or "").strip() == endereco_evento
                        and str(last["cidade"] or "").strip() == cidade_evento
                        and str(last["bairro"] or "").strip() == bairro_evento
                    )
                    same_status = str(last["status_pedido"] or "").strip() == str(cliente_row.get("status_pedido") or "").strip()
                    if same_geo and same_addr and same_status:
                        return
                cur.execute(
                    """
                    INSERT INTO cliente_localizacao_amostras
                        (cod_cliente, codigo_programacao, pedido, latitude, longitude, endereco, cidade, bairro,
                         status_pedido, motorista, origem, registrado_em)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        cod_cliente_local.upper(),
                        upper(prog),
                        pedido_norm_local,
                        float(lat_evento) if lat_evento not in (None, "") else None,
                        float(lon_evento) if lon_evento not in (None, "") else None,
                        endereco_evento,
                        cidade_evento,
                        bairro_evento,
                        str(cliente_row.get("status_pedido") or "").strip(),
                        str(cliente_row.get("alterado_por") or "").strip(),
                        "API",
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ),
                )
            except Exception:
                logging.debug("Falha ao registrar amostra localizacao do cliente", exc_info=True)

        for cliente in clientes:
            cod_cliente = str(cliente.get("cod_cliente") or "").strip()
            if not cod_cliente:
                continue
            pedido_norm = _norm_pedido(cliente.get("pedido"))

            ctrl_map = {
                "mortalidade_aves": cliente.get("mortalidade_aves"),
                "media_aplicada": cliente.get("media_aplicada"),
                "peso_previsto": cliente.get("peso_previsto"),
                "valor_recebido": cliente.get("recebido_valor", cliente.get("valor_recebido")),
                "forma_recebimento": cliente.get("recebido_forma", cliente.get("forma_recebimento")),
                "obs_recebimento": cliente.get("recebido_obs", cliente.get("obs_recebimento")),
                "status_pedido": cliente.get("status_pedido"),
                "alteracao_tipo": cliente.get("alteracao_tipo"),
                "alteracao_detalhe": cliente.get("alteracao_detalhe"),
                "pedido": pedido_norm,
                "caixas_atual": cliente.get("caixas_atual"),
                "preco_atual": cliente.get("preco_atual"),
                "alterado_em": cliente.get("alterado_em"),
                "alterado_por": cliente.get("alterado_por"),
                "lat_evento": cliente.get("lat_evento"),
                "lon_evento": cliente.get("lon_evento"),
                "endereco_evento": cliente.get("endereco_evento"),
                "cidade_evento": cliente.get("cidade_evento"),
                "bairro_evento": cliente.get("bairro_evento"),
            }

            # UPSERT robusto no controle (compatÃÂvel com bancos antigos e novos)
            if cols_ctrl:
                key_cols = [("codigo_programacao", prog), ("cod_cliente", cod_cliente)]
                if "pedido" in cols_ctrl and pedido_norm:
                    key_cols.append(("pedido", pedido_norm))

                set_cols = [(k, v) for k, v in ctrl_map.items() if k.lower() in cols_ctrl and k.lower() not in {kc[0] for kc in key_cols}]
                if "updated_at" in cols_ctrl:
                    set_cols.append(("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

                if set_cols:
                    set_sql = ", ".join(f"{k}=?" for k, _ in set_cols)
                    where_sql = " AND ".join(f"{k}=?" for k, _ in key_cols)
                    params = [v for _, v in set_cols] + [v for _, v in key_cols]
                    cur.execute(f"UPDATE programacao_itens_controle SET {set_sql} WHERE {where_sql}", params)

                if cur.rowcount == 0:
                    ins_cols = [k for k, _ in key_cols]
                    ins_vals = [v for _, v in key_cols]
                    for k, v in ctrl_map.items():
                        if k.lower() in cols_ctrl and k not in ins_cols:
                            ins_cols.append(k)
                            ins_vals.append(v)
                    if "updated_at" in cols_ctrl and "updated_at" not in ins_cols:
                        ins_cols.append("updated_at")
                        ins_vals.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    ph = ", ".join("?" for _ in ins_cols)
                    cur.execute(
                        f"INSERT INTO programacao_itens_controle ({', '.join(ins_cols)}) VALUES ({ph})",
                        ins_vals,
                    )
                _registrar_amostra_local(ctrl_map, pedido_norm, cod_cliente)

            # Espelha também em programacao_itens para a UI do desktop refletir imediatamente.
            if cols_itens:
                sets = []
                vals = []
                if "status_pedido" in cols_itens and ctrl_map.get("status_pedido") not in (None, ""):
                    sets.append("status_pedido=?")
                    vals.append(ctrl_map.get("status_pedido"))
                if "caixas_atual" in cols_itens and ctrl_map.get("caixas_atual") not in (None, ""):
                    sets.append("caixas_atual=?")
                    vals.append(ctrl_map.get("caixas_atual"))
                if "preco_atual" in cols_itens and ctrl_map.get("preco_atual") not in (None, ""):
                    sets.append("preco_atual=?")
                    vals.append(ctrl_map.get("preco_atual"))
                if "kg" in cols_itens and ctrl_map.get("peso_previsto") not in (None, ""):
                    sets.append("kg=?")
                    vals.append(ctrl_map.get("peso_previsto"))
                if sets:
                    where = "codigo_programacao=? AND UPPER(TRIM(cod_cliente))=UPPER(TRIM(?))"
                    wvals = [prog, cod_cliente]
                    if "pedido" in cols_itens and pedido_norm:
                        where += " AND TRIM(COALESCE(pedido,''))=TRIM(?)"
                        wvals.append(pedido_norm)
                    cur.execute(f"UPDATE programacao_itens SET {', '.join(sets)} WHERE {where}", (*vals, *wvals))

            # Espelha recebimentos para a tela Recebimentos do desktop.
            if cols_receb:
                valor = safe_float(ctrl_map.get("valor_recebido"), 0.0)
                if valor > 0:
                    nome = str(cliente.get("nome_cliente") or cliente.get("nome") or "").strip()
                    forma = str(ctrl_map.get("forma_recebimento") or "DINHEIRO").strip() or "DINHEIRO"
                    obs = str(ctrl_map.get("obs_recebimento") or "").strip()
                    if pedido_norm:
                        obs = (f"[PED {pedido_norm}] " + obs).strip()
                    now_s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Mantém 1 registro vigente por cliente (compatÃÂvel com esquema atual do desktop).
                    cur.execute(
                        "DELETE FROM recebimentos WHERE codigo_programacao=? AND cod_cliente=?",
                        (prog, cod_cliente),
                    )
                    cur.execute(
                        """
                        INSERT INTO recebimentos
                            (codigo_programacao, cod_cliente, nome_cliente, valor, forma_pagamento, observacao, data_registro)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (prog, cod_cliente, nome, valor, forma, obs, now_s),
                    )

    def _populate_route_entries(self, rota: dict):
        def _pick(*keys):
            for k in keys:
                if k in rota and rota.get(k) not in (None, ""):
                    return rota.get(k)
            return ""

        nf_kg = _pick("nf_kg", "kg_nota_fiscal")
        nf_preco = _pick("nf_preco", "preco_nf", "preco_nota_fiscal")
        nf_kg_carregado = _pick("nf_kg_carregado", "kg_carregado")
        nf_caixas = _pick("nf_caixas", "caixas_carregadas", "total_caixas")
        nf_kg_vendido = _pick("nf_kg_vendido", "kg_vendido")
        mort_aves_transb = _pick("aves_mortas_transbordo", "mortalidade_transbordo_aves")
        mort_kg_transb = _pick("mortalidade_transbordo_kg")
        obs_transb = _pick("obs_transbordo", "mortalidade_transbordo_obs")
        nf_saldo = _pick("nf_saldo")
        media_carregada = self._normalize_media_kg_ave(_pick("media", "media_carregada", "media_base"))
        caixa_final = _pick("caixa_final", "qnt_aves_caixa_final", "aves_caixa_final")
        if nf_saldo in (None, ""):
            nf_saldo = round(max(safe_float(nf_kg, 0.0) - safe_float(nf_kg_carregado, 0.0), 0.0), 2)

        campos = [
            (_pick("nf_numero", "num_nf", "nf"), self.ent_nf_numero),
            (nf_kg, self.ent_nf_kg),
            (nf_preco, getattr(self, "ent_nf_preco", None)),
            (nf_caixas, self.ent_nf_caixas),
            (nf_kg_carregado, self.ent_nf_kg_carregado),
            (nf_kg_vendido, self.ent_nf_kg_vendido),
            (nf_saldo, self.ent_nf_saldo),
            (media_carregada, getattr(self, "ent_nf_media_carregada", None)),
            (caixa_final, getattr(self, "ent_nf_caixa_final", None)),
            (_pick("km_inicial"), self.ent_km_inicial),
            (_pick("km_final"), self.ent_km_final),
            (_pick("litros"), self.ent_litros),
            (_pick("km_rodado"), self.ent_km_rodado),
            (_pick("media_km_l"), self.ent_media),
            (_pick("custo_km"), self.ent_custo_km),
        ]
        self._nf_mort_aves = safe_int(mort_aves_transb, 0)
        self._nf_mort_kg = safe_float(mort_kg_transb, 0.0)
        self._nf_obs_transbordo = str(obs_transb or "")
        for valor, entry in campos:
            if not entry:
                continue
            self._set_ent(entry, valor)
        self._refresh_nf_trade_summary()

    # =========================================================
    # 7.5 PageBase (refresh / on_show / status)
    # =========================================================
    def refresh_comboboxes(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/programacoes?modo=todas&limit=300",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                programas = []
                for r in (arr or []):
                    if not isinstance(r, dict):
                        continue
                    codigo = upper((r or {}).get("codigo_programacao") or "")
                    status_ref = upper((r or {}).get("status_operacional") or (r or {}).get("status") or "")
                    data = str((r or {}).get("data_criacao") or "")[:10] or "Sem data"
                    if codigo and status_ref not in {"CANCELADA", "CANCELADO"}:
                        programas.append(f"{codigo} ({data})")
                self.cb_prog["values"] = programas
                atual = str(self.cb_prog.get() or "").strip()
                if atual and atual not in set(programas):
                    self.cb_prog.set("")
                return
            except Exception:
                logging.debug("Falha ao carregar combobox Despesas via API", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                has_status = "status" in cols
                has_status_op = "status_operacional" in cols
                has_finalizada_app = "finalizada_no_app" in cols
                has_prest = "prestacao_status" in cols

                status_base_expr = "UPPER(TRIM(COALESCE(status,'')))" if has_status else "''"
                status_op_expr = "UPPER(TRIM(COALESCE(status_operacional,'')))" if has_status_op else "''"
                finalizada_expr = "COALESCE(finalizada_no_app,0)" if has_finalizada_app else "0"
                prest_expr = "UPPER(TRIM(COALESCE(prestacao_status,'PENDENTE')))" if has_prest else "'PENDENTE'"

                where_parts = []
                if has_status_op:
                    where_parts.append(f"{status_op_expr} NOT IN ('CANCELADA','CANCELADO')")
                if has_status:
                    where_parts.append(f"{status_base_expr} NOT IN ('CANCELADA','CANCELADO')")
                where_sql = " OR ".join(where_parts) if where_parts else "1=1"

                prest_where = ""
                if has_prest:
                    prest_where = f" AND {prest_expr} IN ('PENDENTE','FECHADA')"
                data_expr = "COALESCE(data_criacao,'')" if "data_criacao" in cols else ("COALESCE(data,'')" if "data" in cols else ("COALESCE(data_saida,'')" if "data_saida" in cols else "''"))

                cur.execute(
                    f"""
                    SELECT codigo_programacao, {data_expr} AS data_ref
                    FROM programacoes
                    WHERE ({where_sql}){prest_where}
                    ORDER BY id DESC
                    LIMIT 300
                    """
                )
            except Exception:
                cur.execute("SELECT codigo_programacao, '' AS data_ref FROM programacoes ORDER BY id DESC LIMIT 300")

            programas = []
            for row in cur.fetchall():
                codigo = row[0]
                data = row[1][:10] if row[1] else "Sem data"
                programas.append(f"{codigo} ({data})")

            self.cb_prog["values"] = programas

    def on_show(self):
        self.refresh_comboboxes()
        self._refresh_all()
        self.set_status("STATUS: Despesas da rota e controle completo (NF / KM / Dinheiro).")
        self.logger.info("Página Despesas carregada")

    def set_status(self, text):
        try:
            PageBase.set_status(self, text)
        except Exception:
            try:
                self.lbl_status.config(text=text)
            except Exception:
                print(f"STATUS: {text}")

# ==========================
# ===== FIM DA PARTE 7B (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INICIO DA PARTE 7C (ATUALIZADA) =====
# ==========================

    # =========================================================
    # 7.6 CRUD DESPESAS (REGISTRAR / EDITAR / EXCLUIR)
    # =========================================================

    # --------- REGRAS DE SEGURANÇA (não altera sistema; só bloqueia quando necessário)
    def _get_prestacao_status(self, prog: str) -> str:
        try:
            bundle = self._api_bundle_prestacao(prog)
            rota = bundle.get("rota") if isinstance(bundle, dict) else None
            if isinstance(rota, dict):
                return upper(str(rota.get("prestacao_status") or ""))
        except Exception:
            logging.debug("Falha ao consultar status da prestacao via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT COALESCE(prestacao_status,'') FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1", (prog,))
                row = cur.fetchone()
            return upper(row[0] if row else "")
        except Exception:
            return ""

    def _can_edit_current_prog(self) -> bool:
        prog = self._current_programacao
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Selecione a Programação primeiro.")
            return False
        if not ensure_system_api_binding(context=f"Editar despesas ({prog})", parent=self):
            return False

        status = self._get_prestacao_status(prog)
        if status == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"Esta programação ({prog}) está com a prestação FECHADA.\n\n"
                "Por segurança, não é permitido registrar/editar/excluir despesas."
            )
            return False
        return True

    def _collect_logistica_rastreio(self, prog: str) -> dict:
        """
        Consolida rastreabilidade de substituições e transferências para
        validar fechamento sem caixas "soltas".
        """
        out = {
            "pend_substituicao": 0,
            "pend_transferencia": 0,
            "transf_out": 0,
            "transf_in": 0,
            "base_cx": 0,
            "atual_cx": 0,
            "esperado_cx": 0,
            "delta_cx": 0,
            "itens_ok": True,
            "resumo": [],
        }
        prog = upper(str(prog or "").strip())
        if not prog:
            return out

        try:
            bundle = self._api_bundle_prestacao(prog)
            info = bundle.get("logistica") if isinstance(bundle, dict) else None
            if isinstance(info, dict):
                out.update({
                    "pend_substituicao": safe_int(info.get("pend_substituicao"), 0),
                    "pend_transferencia": safe_int(info.get("pend_transferencia"), 0),
                    "transf_out": safe_int(info.get("transf_out"), 0),
                    "transf_in": safe_int(info.get("transf_in"), 0),
                    "base_cx": safe_int(info.get("base_cx"), 0),
                    "atual_cx": safe_int(info.get("atual_cx"), 0),
                    "esperado_cx": safe_int(info.get("esperado_cx"), 0),
                    "delta_cx": safe_int(info.get("delta_cx"), 0),
                    "itens_ok": bool(info.get("itens_ok", True)),
                    "resumo": list(info.get("resumo") or []),
                })
                return out
        except Exception:
            logging.debug("Falha ao consolidar rastreabilidade via API; usando fallback local.", exc_info=True)

        try:
            with get_db() as conn:
                cur = conn.cursor()

                def _has_table(tn: str) -> bool:
                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tn,))
                    return bool(cur.fetchone())

                # Substituição pendente bloqueia fechamento.
                if _has_table("rota_substituicoes"):
                    cur.execute(
                        """
                        SELECT COUNT(1)
                        FROM rota_substituicoes
                        WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
                          AND UPPER(TRIM(COALESCE(status,'')))='PENDENTE'
                        """,
                        (prog,),
                    )
                    out["pend_substituicao"] = safe_int((cur.fetchone() or [0])[0], 0)

                # Transferências pendentes e totais ativos (origem/destino).
                if _has_table("transferencias"):
                    cur.execute(
                        """
                        SELECT
                            SUM(CASE
                                    WHEN UPPER(TRIM(COALESCE(status,'')))='PENDENTE'
                                     AND (UPPER(TRIM(COALESCE(codigo_origem,'')))=UPPER(TRIM(?))
                                       OR UPPER(TRIM(COALESCE(codigo_destino,'')))=UPPER(TRIM(?)))
                                    THEN 1 ELSE 0
                                END) AS pend,
                            SUM(CASE
                                    WHEN UPPER(TRIM(COALESCE(codigo_origem,'')))=UPPER(TRIM(?))
                                     AND UPPER(TRIM(COALESCE(status,''))) IN ('PENDENTE','ACEITA','CONVERTIDA')
                                    THEN COALESCE(qtd_caixas,0) ELSE 0
                                END) AS out_cx,
                            SUM(CASE
                                    WHEN UPPER(TRIM(COALESCE(codigo_destino,'')))=UPPER(TRIM(?))
                                     AND UPPER(TRIM(COALESCE(status,''))) IN ('PENDENTE','ACEITA','CONVERTIDA')
                                    THEN COALESCE(qtd_caixas,0) ELSE 0
                                END) AS in_cx
                        FROM transferencias
                        """,
                        (prog, prog, prog, prog),
                    )
                    row = cur.fetchone() or [0, 0, 0]
                    out["pend_transferencia"] = safe_int(row[0], 0)
                    out["transf_out"] = safe_int(row[1], 0)
                    out["transf_in"] = safe_int(row[2], 0)

                # Conciliação de caixas da programação.
                if _has_table("programacao_itens"):
                    cur.execute("PRAGMA table_info(programacao_itens)")
                    cols_it = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    has_cx_atual = "caixas_atual" in cols_it

                    cur.execute(
                        f"""
                        SELECT
                            COALESCE(SUM(COALESCE(qnt_caixas,0)),0) AS base_cx,
                            COALESCE(SUM(COALESCE({"caixas_atual" if has_cx_atual else "qnt_caixas"},0)),0) AS atual_cx
                        FROM programacao_itens
                        WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
                        """,
                        (prog,),
                    )
                    row = cur.fetchone() or [0, 0]
                    out["base_cx"] = safe_int(row[0], 0)
                    atual_calc = safe_int(row[1], 0)

                    # Fechamento de prestação: saldo da rota deve ser zero.
                    # Em bases legadas, tenta obter caixas_atual do controle.
                    if not has_cx_atual and _has_table("programacao_itens_controle"):
                        try:
                            cur.execute("PRAGMA table_info(programacao_itens_controle)")
                            cols_ctl = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                            if "caixas_atual" in cols_ctl:
                                cur.execute(
                                    """
                                    SELECT COALESCE(SUM(COALESCE(caixas_atual,0)),0)
                                    FROM programacao_itens_controle
                                    WHERE UPPER(TRIM(COALESCE(codigo_programacao,'')))=UPPER(TRIM(?))
                                    """,
                                    (prog,),
                                )
                                atual_calc = safe_int((cur.fetchone() or [0])[0], 0)
                        except Exception:
                            logging.debug("Falha ao calcular caixas_atual via controle", exc_info=True)

                    out["atual_cx"] = max(atual_calc, 0)
                    out["esperado_cx"] = 0
                    out["delta_cx"] = out["atual_cx"] - out["esperado_cx"]
                    out["itens_ok"] = (out["atual_cx"] == 0)

                out["resumo"] = [
                    f"Substituições pendentes: {out['pend_substituicao']}",
                    f"Transferências pendentes: {out['pend_transferencia']}",
                    f"Transferência caixas (origem): {out['transf_out']} cx",
                    f"Transferência caixas (destino): {out['transf_in']} cx",
                    f"Caixas base: {out['base_cx']} cx",
                    f"Caixas atuais: {out['atual_cx']} cx",
                    f"Caixas esperadas no fechamento: {out['esperado_cx']} cx",
                    f"Delta caixas: {out['delta_cx']} cx",
                ]
        except Exception:
            logging.debug("Falha ao consolidar rastreabilidade logÃÂstica", exc_info=True)
        return out

    def _validar_fechamento_logistico(self, prog: str) -> tuple[bool, str]:
        info = self._collect_logistica_rastreio(prog)
        if safe_int(info.get("pend_substituicao"), 0) > 0:
            return (
                False,
                "Nao e possivel finalizar a prestacao:\n"
                "existe substituição de rota pendente de aceite/recusa.",
            )
        if safe_int(info.get("pend_transferencia"), 0) > 0:
            return (
                False,
                "Nao e possivel finalizar a prestacao:\n"
                "existem transferências de caixas pendentes.",
            )
        if not bool(info.get("itens_ok", True)):
            return (
                False,
                "Inconsistência de caixas detectada para fechamento:\n"
                + "\n".join(info.get("resumo") or []),
            )
        return True, ""

    def finalizar_prestacao_despesas(self):
        prog = self._current_programacao
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione a programacao primeiro.")
            return
        if not ensure_system_api_binding(context=f"Finalizar prestacao ({prog})", parent=self):
            return

        status = self._get_prestacao_status(prog)
        if status == "FECHADA":
            messagebox.showinfo("Prestacao", "Esta prestacao ja esta FECHADA.")
            return

        ok_log, msg_log = self._validar_fechamento_logistico(prog)
        if not ok_log:
            messagebox.showerror("Fechamento bloqueado", msg_log)
            return

        if not messagebox.askyesno(
            "Confirmar",
            f"Finalizar prestacao de contas da rota {prog}?\n\n"
            "A programacao sera finalizada."
        ):
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "PUT",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/status",
                payload={
                    "prestacao_status": "FECHADA",
                    "status": "FINALIZADA",
                    "status_operacional": "FINALIZADA",
                    "finalizada_no_app": 1,
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            messagebox.showinfo("OK", f"Prestacao finalizada: {prog}")
            try:
                if messagebox.askyesno(
                    "Imprimir prestacao",
                    f"Deseja imprimir agora a prestacao da rota {prog}?\n\n"
                    "Sera gerada a folha completa com retorno operacional, recebimentos e despesas."
                ):
                    self._current_programacao = prog
                    self.imprimir_resumo()
            except Exception:
                logging.exception("Falha ao imprimir prestacao apos finalizar")
            self.refresh_comboboxes()
            if hasattr(self.app, "refresh_programacao_comboboxes"):
                try:
                    self.app.refresh_programacao_comboboxes()
                except Exception:
                    logging.debug("Falha ignorada")
            try:
                if hasattr(self, "app") and hasattr(self.app, "pages"):
                    rec_page = self.app.pages.get("Recebimentos")
                    if rec_page and hasattr(rec_page, "_reset_view"):
                        rec_page._reset_view()
            except Exception:
                logging.debug("Falha ignorada")
            try:
                if hasattr(self, "app") and hasattr(self.app, "pages"):
                    home_page = self.app.pages.get("Home")
                    if home_page and hasattr(home_page, "on_show"):
                        home_page.on_show()
            except Exception:
                logging.debug("Falha ignorada")
            self._current_programacao = None
            try:
                self.cb_prog.set("")
            except Exception:
                logging.debug("Falha ignorada")
            try:
                if hasattr(self, "_limpar_busca_despesas"):
                    self._limpar_busca_despesas()
            except Exception:
                logging.debug("Falha ignorada")
            self._load_by_programacao()
            self.set_status("STATUS: Prestacao finalizada. Selecione uma nova programacao para continuar.")
        except Exception as e:
            messagebox.showerror(
                "ERRO",
                "Não foi possível finalizar a prestação.\n\n"
                f"{_friendly_sync_error(e)}",
            )
    

    def _parse_money_local(self, s):
        # garante parse de 'R$ 1.234,56' / '1234,56' / '1234.56'
        return safe_money(s, 0.0)

    def _format_money_from_digits(self, digits: str) -> str:
        digits = re.sub(r"\D", "", str(digits or ""))
        if not digits:
            return "0,00"
        if len(digits) == 1:
            int_part = "0"
            cents = "0" + digits
        elif len(digits) == 2:
            int_part = "0"
            cents = digits
        else:
            int_part = digits[:-2]
            cents = digits[-2:]
        int_part = int_part.lstrip("0") or "0"

        parts = []
        while len(int_part) > 3:
            parts.insert(0, int_part[-3:])
            int_part = int_part[:-3]
        if int_part:
            parts.insert(0, int_part)
        int_part = ".".join(parts) if parts else "0"
        return f"{int_part},{cents}"

    def _bind_money_entry(self, ent: tk.Entry):
        def _on_focus_in(_e=None):
            try:
                v = str(ent.get() or "").strip()
                if v in {"0", "0,00", "0.00", "R$ 0,00", "R$0,00"}:
                    ent.delete(0, "end")
                else:
                    ent.selection_range(0, "end")
            except Exception:
                logging.debug("Falha ignorada")

        def _on_key_release(_e=None):
            try:
                v = str(ent.get() or "")
                digits = re.sub(r"\D", "", v)
                if digits:
                    ent.delete(0, "end")
                    ent.insert(0, self._format_money_from_digits(digits))
                    ent.icursor("end")
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out(_e=None):
            try:
                v = str(ent.get() or "").strip()
                if not v:
                    ent.insert(0, "0,00")
                else:
                    ent.delete(0, "end")
                    ent.insert(0, self._format_money_from_digits(v))
            except Exception:
                logging.debug("Falha ignorada")

        ent.bind("<FocusIn>", _on_focus_in)
        ent.bind("<KeyRelease>", _on_key_release)
        ent.bind("<FocusOut>", _on_focus_out)

    def _get_despesa_info(self, did):
        # retorna (id, codigo_programacao, descricao, valor, categoria)
        try:
            cached = (getattr(self, "_despesas_cache", {}) or {}).get(str(did))
            if cached:
                return (cached[0], cached[1], cached[2], cached[3], cached[4])
        except Exception:
            logging.debug("Falha ignorada")
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/despesas/{safe_int(did, 0)}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                desp = resp.get("despesa") if isinstance(resp, dict) else None
                if isinstance(desp, dict):
                    return (
                        safe_int(desp.get("id"), 0),
                        upper(str(desp.get("codigo_programacao") or "")),
                        str(desp.get("descricao") or ""),
                        safe_float(desp.get("valor"), 0.0),
                        upper(str(desp.get("categoria") or "OUTROS")),
                    )
            except Exception:
                logging.debug("Falha ao obter despesa via API; usando fallback local.", exc_info=True)
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT id, codigo_programacao, descricao, valor, COALESCE(categoria,'GERAL')
                    FROM despesas
                    WHERE id=?
                    LIMIT 1
                """, (did,))
                return cur.fetchone()
        except Exception:
            return None

    def _open_registrar_rapido(self):
        if not self._can_edit_current_prog():
            return

        prog = self._current_programacao

        win = tk.Toplevel(self)
        win.title("Registrar Despesa (Rápido)")
        win.geometry("520x360")
        win.grab_set()
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=20)
        frm.pack(fill="both", expand=True)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="NOVA DESPESA", font=("Segoe UI", 14, "bold"))\
            .grid(row=0, column=0, columnspan=2, pady=(0, 16), sticky="w")

        ttk.Label(frm, text="Descrição:", font=("Segoe UI", 10, "bold"))\
            .grid(row=1, column=0, sticky="w", pady=5)
        ent_desc = ttk.Entry(frm, style="Field.TEntry")
        ent_desc.grid(row=1, column=1, sticky="ew", pady=5, padx=(10, 0))

        ttk.Label(frm, text="Valor (R$):", font=("Segoe UI", 10, "bold"))\
            .grid(row=2, column=0, sticky="w", pady=5)
        ent_val = ttk.Entry(frm, style="Field.TEntry")
        ent_val.grid(row=2, column=1, sticky="ew", pady=5, padx=(10, 0))
        ent_val.insert(0, "0,00")
        self._bind_money_entry(ent_val)

        ttk.Label(frm, text="Categoria:", font=("Segoe UI", 10, "bold"))\
            .grid(row=3, column=0, sticky="w", pady=5)
        cb_categoria = ttk.Combobox(
            frm, state="readonly", width=22,
            values=["DIARIAS", "COMBUSTIVEL", "SERVICOS NO VEICULOS", "DIARIA EXTRA", "ESTRADA", "BANHO", "OUTROS"]
        )
        cb_categoria.set("OUTROS")
        cb_categoria.grid(row=3, column=1, sticky="w", pady=5, padx=(10, 0))

        ttk.Label(frm, text="Observação:", font=("Segoe UI", 10, "bold"))\
            .grid(row=4, column=0, sticky="w", pady=5)
        ent_obs = tk.Text(frm, height=4, width=30, font=("Segoe UI", 9))
        ent_obs.grid(row=4, column=1, sticky="ew", pady=5, padx=(10, 0))

        # Enter salva / Esc fecha
        win.bind("<Escape>", lambda e: win.destroy())

        def salvar():
            if not self._can_edit_current_prog():
                return

            desc = upper(ent_desc.get().strip())
            val = self._parse_money_local(ent_val.get().strip())
            categoria = upper(cb_categoria.get())
            obs = upper(ent_obs.get("1.0", "end-1c").strip())

            if not desc:
                messagebox.showwarning("ATENÇÃO", "Informe a descrição.")
                return
            if val <= 0:
                messagebox.showwarning("ATENÇÃO", "O valor deve ser maior que zero.")
                return
            if not categoria:
                categoria = "OUTROS"

            try:
                desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                _call_api(
                    "POST",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/despesas",
                    payload={
                        "descricao": desc,
                        "valor": float(val),
                        "categoria": categoria,
                        "observacao": obs,
                    },
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )

                self.logger.info(f"Despesa registrada: {prog} - {desc} - R${val:.2f} - {categoria}")
                win.destroy()
                self._refresh_all()

            except Exception as e:
                self.logger.error(f"Erro ao salvar despesa: {str(e)}")
                messagebox.showerror(
                    "ERRO",
                    "Não foi possível registrar a despesa.\n\n"
                    f"{_friendly_sync_error(e)}",
                )

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        ttk.Button(btn_frame, text="\U0001F4BE SALVAR", style="Primary.TButton", command=salvar)\
            .grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btn_frame, text="\u274C CANCELAR", style="Ghost.TButton", command=win.destroy)\
            .grid(row=0, column=1, sticky="ew", padx=(6, 0))

        win.bind("<Return>", lambda e: salvar())
        ent_desc.focus_set()

    def _editar_linha_selecionada(self):
        if not self._can_edit_current_prog():
            return

        prog = self._current_programacao

        if not self.selected_despesa_id:
            messagebox.showwarning("ATENÇÃO", "Selecione uma linha na tabela.")
            return

        sel = self.tree_desp.selection()
        if not sel:
            return

        vals = self.tree_desp.item(sel[0], "values")
        if len(vals) < 6:
            messagebox.showerror("ERRO", "Dados incompletos na linha selecionada.")
            return

        did, desc, val_str, data_reg, categoria, observacao = vals

        # âÅâ€œââ‚¬¦ valida que essa despesa pertence à programação atual
        info = self._get_despesa_info(did)
        if not info:
            messagebox.showerror("ERRO", "Não encontrei essa despesa no banco.")
            return
        _, codigo_prog_db, _, _, _ = info
        if upper(codigo_prog_db) != upper(prog):
            messagebox.showwarning(
                "ATENÇÃO",
                "Por segurança, não é permitido editar uma despesa que não pertence à programação atual."
            )
            return

        win = tk.Toplevel(self)
        win.title(f"Editar Despesa - ID {did}")
        win.geometry("520x380")
        win.grab_set()
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=20)
        frm.pack(fill="both", expand=True)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="EDITAR DESPESA", font=("Segoe UI", 14, "bold"))\
            .grid(row=0, column=0, columnspan=2, pady=(0, 16), sticky="w")

        row = 1

        ttk.Label(frm, text="Descrição:", font=("Segoe UI", 10, "bold"))\
            .grid(row=row, column=0, sticky="w", pady=5)
        ent_desc = ttk.Entry(frm, style="Field.TEntry")
        ent_desc.grid(row=row, column=1, sticky="ew", pady=5, padx=(10, 0))
        ent_desc.insert(0, desc or "")
        row += 1

        ttk.Label(frm, text="Valor (R$):", font=("Segoe UI", 10, "bold"))\
            .grid(row=row, column=0, sticky="w", pady=5)
        ent_val = ttk.Entry(frm, style="Field.TEntry")
        ent_val.grid(row=row, column=1, sticky="ew", pady=5, padx=(10, 0))
        ent_val.insert(0, str(val_str).replace("R$", "").strip())
        self._bind_money_entry(ent_val)
        row += 1

        ttk.Label(frm, text="Categoria:", font=("Segoe UI", 10, "bold"))\
            .grid(row=row, column=0, sticky="w", pady=5)
        cb_categoria = ttk.Combobox(
            frm, state="readonly", width=22,
            values=["DIARIAS", "COMBUSTIVEL", "SERVICOS NO VEICULOS", "DIARIA EXTRA", "ESTRADA", "BANHO", "OUTROS"]
        )
        cb_categoria.set(upper(categoria) if categoria else "OUTROS")
        cb_categoria.grid(row=row, column=1, sticky="w", pady=5, padx=(10, 0))
        row += 1

        ttk.Label(frm, text="Observação:", font=("Segoe UI", 10, "bold"))\
            .grid(row=row, column=0, sticky="w", pady=5)
        ent_obs = tk.Text(frm, height=4, width=30, font=("Segoe UI", 9))
        ent_obs.grid(row=row, column=1, sticky="ew", pady=5, padx=(10, 0))
        ent_obs.insert("1.0", observacao or "")
        row += 1

        ttk.Label(frm, text="Data Registro:", font=("Segoe UI", 10, "bold"))\
            .grid(row=row, column=0, sticky="w", pady=5)
        lbl_data = ttk.Label(frm, text=data_reg or "", font=("Segoe UI", 9))
        lbl_data.grid(row=row, column=1, sticky="w", pady=5, padx=(10, 0))
        row += 1

        win.bind("<Escape>", lambda e: win.destroy())

        def salvar():
            if not self._can_edit_current_prog():
                return

            ndesc = upper(ent_desc.get().strip())
            nval = self._parse_money_local(ent_val.get().strip())
            ncategoria = upper(cb_categoria.get())
            nobs = upper(ent_obs.get("1.0", "end-1c").strip())

            if not ndesc:
                messagebox.showwarning("ATENÇÃO", "Informe a descrição.")
                return
            if nval <= 0:
                messagebox.showwarning("ATENÇÃO", "O valor deve ser maior que zero.")
                return
            if not ncategoria:
                ncategoria = "OUTROS"

            # âÅâ€œââ‚¬¦ valida novamente (segurança)
            info2 = self._get_despesa_info(did)
            if not info2 or upper(info2[1]) != upper(prog):
                messagebox.showwarning("ATENÇÃO", "Despesa não pertence à programação atual (bloqueado).")
                return

            try:
                desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                _call_api(
                    "PUT",
                    f"desktop/despesas/{did}",
                    payload={
                        "descricao": ndesc,
                        "valor": float(nval),
                        "categoria": ncategoria,
                        "observacao": nobs,
                    },
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )

                self.logger.info(f"Despesa atualizada: ID {did} - {ndesc} - R${nval:.2f}")
                win.destroy()
                self._load_by_programacao()

            except Exception as e:
                self.logger.error(f"Erro ao atualizar despesa: {str(e)}")
                messagebox.showerror(
                    "ERRO",
                    "Não foi possível atualizar a despesa.\n\n"
                    f"{_friendly_sync_error(e)}",
                )

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        ttk.Button(btn_frame, text="\U0001F4BE SALVAR", style="Primary.TButton", command=salvar)\
            .grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btn_frame, text="\u274C CANCELAR", style="Ghost.TButton", command=win.destroy)\
            .grid(row=0, column=1, sticky="ew", padx=(6, 0))

        win.bind("<Return>", lambda e: salvar())
        ent_desc.focus_set()

    def _excluir_linha_selecionada(self):
        if not self._can_edit_current_prog():
            return

        prog = self._current_programacao
        if not self.selected_despesa_id:
            messagebox.showwarning("ATENÇÃO", "Selecione uma linha na tabela.")
            return

        did = self.selected_despesa_id

        # âÅâ€œââ‚¬¦ pega dados reais do banco para confirmar e validar pertencimento
        info = self._get_despesa_info(did)
        if not info:
            messagebox.showerror("ERRO", "Não encontrei essa despesa no banco.")
            return

        _, codigo_prog_db, desc_db, val_db, cat_db = info
        if upper(codigo_prog_db) != upper(prog):
            messagebox.showwarning(
                "ATENÇÃO",
                "Por segurança, não é permitido excluir uma despesa que não pertence à programação atual."
            )
            return

        val_fmt = self._fmt_money(val_db) if hasattr(self, "_fmt_money") else f"R$ {float(val_db or 0):.2f}"
        resposta = messagebox.askyesno(
            "CONFIRMAR EXCLUSÃO",
            "Deseja realmente excluir esta despesa?\n\n"
            f"ID: {did}\n"
            f"Descrição: {desc_db}\n"
            f"Valor: {val_fmt}\n"
            f"Categoria: {cat_db}\n\n"
            "Esta ação não pode ser desfeita."
        )
        if not resposta:
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "DELETE",
                f"desktop/despesas/{did}?codigo_programacao={urllib.parse.quote(upper(prog))}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            self.logger.warning(f"Despesa excluÃÂÂda: {prog} - ID {did} - {desc_db} - {val_fmt}")
            self.selected_despesa_id = None
            self._refresh_all()
            self.set_status(f"STATUS: Despesa ID {did} excluÃÂÂda com sucesso")

        except Exception as e:
            self.logger.error(f"Erro ao excluir despesa: {str(e)}")
            messagebox.showerror(
                "ERRO",
                "Não foi possível excluir a despesa.\n\n"
                f"{_friendly_sync_error(e)}",
            )

    # =========================================================
    # 7.7 RELATÓRIO EM TELA + IMPRESSÃO SIMULADA
    # =========================================================
    def abrir_previsualizacao_prestacao(self):
        prog = self._current_programacao
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione a Programacao primeiro.")
            return

        relatorios = None
        if hasattr(self, "app") and hasattr(self.app, "pages"):
            relatorios = self.app.pages.get("Relatorios")

        if not relatorios or not hasattr(relatorios, "abrir_previsualizacao_relatorio"):
            self.imprimir_resumo()
            return

        prev_tipo = relatorios.cb_tipo_rel.get().strip() if hasattr(relatorios, "cb_tipo_rel") else ""
        prev_prog = relatorios.cb_prog.get().strip() if hasattr(relatorios, "cb_prog") else ""
        try:
            if hasattr(relatorios, "cb_tipo_rel"):
                relatorios.cb_tipo_rel.set("Prestacao de Contas")
            if hasattr(relatorios, "cb_prog"):
                relatorios.cb_prog.set(prog)
            relatorios.abrir_previsualizacao_relatorio()
        finally:
            try:
                if hasattr(relatorios, "cb_tipo_rel"):
                    relatorios.cb_tipo_rel.set(prev_tipo)
                if hasattr(relatorios, "cb_prog"):
                    relatorios.cb_prog.set(prev_prog)
            except Exception:
                logging.debug("Falha ignorada")

    def imprimir_resumo(self):
        if not require_reportlab():
            return
        try:
            prog = self._current_programacao
            if not prog:
                messagebox.showwarning("ATENCAO", "Selecione a Programacao primeiro.")
                return

            path_pdf = filedialog.asksaveasfilename(
                title="Salvar PDF - Despesas",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"DESPESAS_{prog}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            )
            if not path_pdf:
                return

            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm

            # ---- coleta dados ----
            motorista = veiculo = equipe = ""
            equipe_txt = ""
            nf = local_rota = local_carreg = ""
            data_saida = hora_saida = data_chegada = hora_chegada = ""
            nf_kg = nf_caixas = nf_kg_carregado = nf_kg_vendido = nf_saldo = 0.0
            nf_preco = 0.0
            nf_media_carregada = 0.0
            nf_caixa_final = 0
            km_inicial = km_final = litros = km_rodado = media_km_l = custo_km = 0.0
            ced_qtd = {200: 0, 100: 0, 50: 0, 20: 0, 10: 0, 5: 0, 2: 0}
            valor_dinheiro = 0.0
            adiantamento = 0.0
            recebimentos = []
            despesas = []
            total_receb = 0.0
            total_desp = 0.0

            used_api = False
            bundle = self._api_bundle_prestacao(prog)
            if bundle:
                rota = bundle.get("rota") if isinstance(bundle, dict) else None
                recebimentos_api = bundle.get("recebimentos") if isinstance(bundle, dict) else []
                despesas_api = bundle.get("despesas") if isinstance(bundle, dict) else []

                if isinstance(rota, dict):
                    motorista = str(rota.get("motorista") or "")
                    veiculo = str(rota.get("veiculo") or "")
                    equipe = str(rota.get("equipe") or "")
                    equipe_txt = resolve_equipe_nomes(equipe)

                    nf = str(rota.get("num_nf") or rota.get("nf_numero") or "")
                    local_rota = str(rota.get("local_rota") or rota.get("tipo_rota") or "")
                    local_carreg = str(rota.get("local_carregamento") or rota.get("granja_carregada") or "")

                    data_saida, hora_saida = normalize_date_time_components(
                        rota.get("data_saida"), rota.get("hora_saida")
                    )
                    data_chegada, hora_chegada = normalize_date_time_components(
                        rota.get("data_chegada"), rota.get("hora_chegada")
                    )

                    nf_kg = safe_float(rota.get("nf_kg"), 0.0)
                    nf_caixas = safe_int(rota.get("nf_caixas"), 0)
                    nf_kg_carregado = safe_float(rota.get("nf_kg_carregado"), 0.0)
                    nf_kg_vendido = safe_float(rota.get("nf_kg_vendido"), 0.0)
                    nf_saldo = safe_float(rota.get("nf_saldo"), 0.0)
                    nf_preco = safe_float(rota.get("nf_preco") or rota.get("preco_nf"), 0.0)
                    nf_media_carregada = safe_float(rota.get("media") or rota.get("media_carregada"), 0.0)
                    nf_caixa_final = safe_int(
                        rota.get("aves_caixa_final")
                        if rota.get("aves_caixa_final") not in (None, "")
                        else rota.get("qnt_aves_caixa_final"),
                        0,
                    )

                    km_inicial = safe_float(rota.get("km_inicial"), 0.0)
                    km_final = safe_float(rota.get("km_final"), 0.0)
                    litros = safe_float(rota.get("litros"), 0.0)
                    km_rodado = safe_float(rota.get("km_rodado"), 0.0)
                    media_km_l = safe_float(rota.get("media_km_l"), 0.0)
                    custo_km = safe_float(rota.get("custo_km"), 0.0)

                    ced_qtd[200] = safe_int(rota.get("ced_200_qtd"), 0)
                    ced_qtd[100] = safe_int(rota.get("ced_100_qtd"), 0)
                    ced_qtd[50] = safe_int(rota.get("ced_50_qtd"), 0)
                    ced_qtd[20] = safe_int(rota.get("ced_20_qtd"), 0)
                    ced_qtd[10] = safe_int(rota.get("ced_10_qtd"), 0)
                    ced_qtd[5] = safe_int(rota.get("ced_5_qtd"), 0)
                    ced_qtd[2] = safe_int(rota.get("ced_2_qtd"), 0)
                    valor_dinheiro = safe_float(rota.get("valor_dinheiro"), 0.0)
                    adiantamento = safe_float(
                        rota.get("adiantamento")
                        if rota.get("adiantamento") not in (None, "")
                        else rota.get("adiantamento_rota"),
                        0.0,
                    )

                    recebimentos = [
                        (
                            r.get("cod_cliente"),
                            r.get("nome_cliente"),
                            safe_float(r.get("valor"), 0.0),
                            r.get("forma_pagamento"),
                            r.get("observacao"),
                            r.get("data_registro"),
                        )
                        for r in (recebimentos_api or [])
                        if isinstance(r, dict)
                    ]
                    despesas = [
                        (
                            d.get("descricao"),
                            safe_float(d.get("valor"), 0.0),
                            d.get("categoria"),
                            d.get("observacao"),
                            d.get("data_registro"),
                        )
                        for d in (despesas_api or [])
                        if isinstance(d, dict)
                    ]

                    total_receb = sum(safe_float(r[2], 0.0) for r in recebimentos)
                    total_desp = sum(safe_float(d[1], 0.0) for d in despesas)
                    used_api = True

            if not used_api:
                with get_db() as conn:
                    cur = conn.cursor()
                    def has_col(col):
                        try:
                            return db_has_column(cur, "programacoes", col)
                        except Exception:
                            return False

                    nf_col = "num_nf" if has_col("num_nf") else ("nf_numero" if has_col("nf_numero") else "'' as num_nf")
                    local_rota_col = "local_rota" if has_col("local_rota") else ("tipo_rota" if has_col("tipo_rota") else "'' as local_rota")
                    local_carreg_col = "local_carregamento" if has_col("local_carregamento") else ("granja_carregada" if has_col("granja_carregada") else "'' as local_carregamento")
                    data_saida_col = "COALESCE(data_saida,'')" if has_col("data_saida") else "''"
                    hora_saida_col = "COALESCE(hora_saida,'')" if has_col("hora_saida") else "''"
                    data_chegada_col = "COALESCE(data_chegada,'')" if has_col("data_chegada") else "''"
                    hora_chegada_col = "COALESCE(hora_chegada,'')" if has_col("hora_chegada") else "''"
                    nf_kg_col = "COALESCE(nf_kg,0)" if has_col("nf_kg") else "0"
                    nf_caixas_col = "COALESCE(nf_caixas,0)" if has_col("nf_caixas") else "0"
                    nf_kg_carregado_col = "COALESCE(nf_kg_carregado,0)" if has_col("nf_kg_carregado") else ("COALESCE(kg_carregado,0)" if has_col("kg_carregado") else "0")
                    nf_kg_vendido_col = "COALESCE(nf_kg_vendido,0)" if has_col("nf_kg_vendido") else ("COALESCE(kg_vendido,0)" if has_col("kg_vendido") else "0")
                    nf_saldo_col = "COALESCE(nf_saldo,0)" if has_col("nf_saldo") else "0"
                    nf_preco_col = "COALESCE(nf_preco,0) as nf_preco" if has_col("nf_preco") else "0 as nf_preco"
                    media_carreg_col = "COALESCE(media,0) as media_carregada" if has_col("media") else "0 as media_carregada"
                    if has_col("aves_caixa_final") and has_col("qnt_aves_caixa_final"):
                        caixa_final_col = "COALESCE(aves_caixa_final, qnt_aves_caixa_final, 0) as caixa_final"
                    elif has_col("aves_caixa_final"):
                        caixa_final_col = "COALESCE(aves_caixa_final, 0) as caixa_final"
                    elif has_col("qnt_aves_caixa_final"):
                        caixa_final_col = "COALESCE(qnt_aves_caixa_final, 0) as caixa_final"
                    else:
                        caixa_final_col = "0 as caixa_final"
                    km_inicial_col = "COALESCE(km_inicial,0)" if has_col("km_inicial") else "0"
                    km_final_col = "COALESCE(km_final,0)" if has_col("km_final") else "0"
                    litros_col = "COALESCE(litros,0)" if has_col("litros") else "0"
                    km_rodado_col = "COALESCE(km_rodado,0)" if has_col("km_rodado") else "0"
                    media_km_l_col = "COALESCE(media_km_l,0)" if has_col("media_km_l") else "0"
                    custo_km_col = "COALESCE(custo_km,0)" if has_col("custo_km") else "0"
                    ced_200_col = "COALESCE(ced_200_qtd,0)" if has_col("ced_200_qtd") else "0"
                    ced_100_col = "COALESCE(ced_100_qtd,0)" if has_col("ced_100_qtd") else "0"
                    ced_50_col = "COALESCE(ced_50_qtd,0)" if has_col("ced_50_qtd") else "0"
                    ced_20_col = "COALESCE(ced_20_qtd,0)" if has_col("ced_20_qtd") else "0"
                    ced_10_col = "COALESCE(ced_10_qtd,0)" if has_col("ced_10_qtd") else "0"
                    ced_5_col = "COALESCE(ced_5_qtd,0)" if has_col("ced_5_qtd") else "0"
                    ced_2_col = "COALESCE(ced_2_qtd,0)" if has_col("ced_2_qtd") else "0"
                    valor_dinheiro_col = "COALESCE(valor_dinheiro,0)" if has_col("valor_dinheiro") else "0"

                    if has_col("adiantamento"):
                        adiant_col = "adiantamento"
                    elif has_col("adiantamento_rota"):
                        adiant_col = "adiantamento_rota as adiantamento"
                    else:
                        adiant_col = "0 as adiantamento"

                    cur.execute(f"""
                        SELECT
                            motorista,
                            veiculo,
                            equipe,
                            {nf_col} as num_nf,
                            {local_rota_col} as local_rota,
                            {local_carreg_col} as local_carreg,
                            {data_saida_col}, {hora_saida_col},
                            {data_chegada_col}, {hora_chegada_col},
                            {nf_kg_col}, {nf_caixas_col}, {nf_kg_carregado_col},
                            {nf_kg_vendido_col}, {nf_saldo_col},
                            {nf_preco_col},
                            {media_carreg_col},
                            {caixa_final_col},
                            {km_inicial_col}, {km_final_col}, {litros_col},
                            {km_rodado_col}, {media_km_l_col}, {custo_km_col},
                            {ced_200_col}, {ced_100_col}, {ced_50_col},
                            {ced_20_col}, {ced_10_col}, {ced_5_col}, {ced_2_col},
                            {valor_dinheiro_col},
                            {adiant_col}
                        FROM programacoes
                        WHERE codigo_programacao=?
                    """, (prog,))
                    row = cur.fetchone() or []

                    if row:
                        idx = 0
                        motorista, veiculo, equipe = row[idx] or "", row[idx + 1] or "", row[idx + 2] or ""
                        idx += 3
                        nf, local_rota, local_carreg = row[idx] or "", row[idx + 1] or "", row[idx + 2] or ""
                        idx += 3
                        data_saida, hora_saida = row[idx] or "", row[idx + 1] or ""
                        idx += 2
                        data_chegada, hora_chegada = row[idx] or "", row[idx + 1] or ""
                        idx += 2
                        nf_kg, nf_caixas, nf_kg_carregado = safe_float(row[idx], 0.0), safe_int(row[idx + 1], 0), safe_float(row[idx + 2], 0.0)
                        idx += 3
                        nf_kg_vendido, nf_saldo = safe_float(row[idx], 0.0), safe_float(row[idx + 1], 0.0)
                        idx += 2
                        nf_preco = safe_float(row[idx], 0.0)
                        nf_media_carregada = safe_float(row[idx + 1], 0.0)
                        nf_caixa_final = safe_int(row[idx + 2], 0)
                        idx += 3
                        km_inicial, km_final = safe_float(row[idx], 0.0), safe_float(row[idx + 1], 0.0)
                        idx += 2
                        litros, km_rodado = safe_float(row[idx], 0.0), safe_float(row[idx + 1], 0.0)
                        idx += 2
                        media_km_l, custo_km = safe_float(row[idx], 0.0), safe_float(row[idx + 1], 0.0)
                        idx += 2
                        ced_qtd[200], ced_qtd[100], ced_qtd[50] = safe_int(row[idx], 0), safe_int(row[idx + 1], 0), safe_int(row[idx + 2], 0)
                        idx += 3
                        ced_qtd[20], ced_qtd[10], ced_qtd[5], ced_qtd[2] = safe_int(row[idx], 0), safe_int(row[idx + 1], 0), safe_int(row[idx + 2], 0), safe_int(row[idx + 3], 0)
                        idx += 4
                        valor_dinheiro = safe_float(row[idx], 0.0)
                        idx += 1
                        adiantamento = safe_float(row[idx], 0.0)

                        equipe_txt = resolve_equipe_nomes(equipe)

                        data_saida, hora_saida = normalize_date_time_components(data_saida, hora_saida)
                        data_chegada, hora_chegada = normalize_date_time_components(data_chegada, hora_chegada)

                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='despesas'")
                    if cur.fetchone():
                        cur.execute("PRAGMA table_info(despesas)")
                        cols_desp = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                        descricao_desp_col = "COALESCE(descricao,'')" if "descricao" in cols_desp else "''"
                        valor_desp_col = "COALESCE(valor,0)" if "valor" in cols_desp else "0"
                        categoria_desp_col = "COALESCE(categoria,'OUTROS')" if "categoria" in cols_desp else "'OUTROS'"
                        obs_desp_col = "COALESCE(observacao,'')" if "observacao" in cols_desp else "''"
                        data_desp_col = "COALESCE(data_registro,'')" if "data_registro" in cols_desp else "''"
                        order_desp_col = "data_registro DESC" if "data_registro" in cols_desp else "1 DESC"
                        cur.execute(
                            f"""
                            SELECT {descricao_desp_col}, {valor_desp_col}, {categoria_desp_col}, {obs_desp_col}, {data_desp_col}
                            FROM despesas
                            WHERE codigo_programacao=?
                            ORDER BY {order_desp_col}
                            """,
                            (prog,),
                        )
                        despesas = cur.fetchall() or []
                        cur.execute("SELECT SUM(valor) FROM despesas WHERE codigo_programacao=?", (prog,))
                        total_desp = safe_float((cur.fetchone() or [0])[0], 0.0)
                    else:
                        despesas = []
                        total_desp = 0.0

                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='recebimentos'")
                    if cur.fetchone():
                        cur.execute("PRAGMA table_info(recebimentos)")
                        cols_rec = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                        cod_rec_col = "COALESCE(cod_cliente,'')" if "cod_cliente" in cols_rec else "''"
                        nome_rec_col = "COALESCE(nome_cliente,'')" if "nome_cliente" in cols_rec else "''"
                        valor_rec_col = "COALESCE(valor,0)" if "valor" in cols_rec else "0"
                        forma_rec_col = "COALESCE(forma_pagamento,'')" if "forma_pagamento" in cols_rec else "''"
                        obs_rec_col = "COALESCE(observacao,'')" if "observacao" in cols_rec else "''"
                        data_rec_col = "COALESCE(data_registro,'')" if "data_registro" in cols_rec else "''"
                        order_rec_col = "data_registro DESC, id DESC" if "data_registro" in cols_rec and "id" in cols_rec else ("id DESC" if "id" in cols_rec else "1 DESC")
                        cur.execute(
                            f"""
                            SELECT {cod_rec_col}, {nome_rec_col}, {valor_rec_col}, {forma_rec_col}, {obs_rec_col}, {data_rec_col}
                            FROM recebimentos
                            WHERE codigo_programacao=?
                            ORDER BY {order_rec_col}
                            """,
                            (prog,),
                        )
                        recebimentos = cur.fetchall() or []
                        cur.execute("SELECT SUM(valor) FROM recebimentos WHERE codigo_programacao=?", (prog,))
                        total_receb = safe_float((cur.fetchone() or [0])[0], 0.0)
                    else:
                        recebimentos = []
                        total_receb = 0.0

            # totais
            ced_total = sum(float(ced) * safe_int(qtd, 0) for ced, qtd in ced_qtd.items())
            total_entradas = total_receb + adiantamento
            total_saidas = total_desp + ced_total
            valor_final_caixa = total_entradas - total_desp
            diferenca = valor_final_caixa - ced_total
            resultado = total_entradas - total_saidas

            # ---- PDF ----
            c = canvas.Canvas(path_pdf, pagesize=A4)
            width, height = A4
            left, right, top, bottom = 12 * mm, 12 * mm, 12 * mm, 12 * mm
            y = height - top

            def new_page(title=None):
                nonlocal y
                c.showPage()
                y = height - top
                if title:
                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(left, y, title)
                    y -= 8 * mm

            def ensure_space(mm_needed):
                nonlocal y
                if y < bottom + (mm_needed * mm):
                    new_page(f"RELATORIO DE DESPESAS - PROGRAMACAO {prog}")

            def _clean(v):
                return fix_mojibake_text(str(v or ""))

            def _clip_width(txt, max_width, font_name="Helvetica", font_size=9):
                txt = _clean(txt)
                if c.stringWidth(txt, font_name, font_size) <= max_width:
                    return txt
                ell = "..."
                while txt and c.stringWidth(txt + ell, font_name, font_size) > max_width:
                    txt = txt[:-1]
                return (txt + ell) if txt else ell

            col_total_w = (width - left - right)
            col1_w = col_total_w * 0.50
            col_gap = 6 * mm
            col2_x = left + col1_w + col_gap
            col2_w = (width - right) - col2_x

            def draw_kv(x, y0, k, v, col_w):
                key_text = f"{_clean(k)}:"
                c.setFont("Helvetica-Bold", 9)
                c.drawString(x, y0, key_text)
                c.setFont("Helvetica", 9)
                key_w = c.stringWidth(key_text, "Helvetica-Bold", 9)
                offset = min(max(22 * mm, key_w + 2.5 * mm), col_w * 0.55)
                avail = max(col_w - offset - 1.5 * mm, 10 * mm)
                c.drawString(x + offset, y0, _clip_width(v, avail, "Helvetica", 9))

            def _fmt_pdf_datetime(v):
                raw = str(v or "").strip()
                if not raw:
                    return ""
                try:
                    dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
                    return dt.strftime("%d/%m/%y %H:%M:%S")
                except Exception:
                    pass
                try:
                    parts = raw.split()
                    d_part = parts[0] if parts else raw
                    h_part = parts[1] if len(parts) > 1 else ""
                    d_fmt = format_date_br_short(d_part)
                    h_fmt = normalize_time(h_part) if h_part else ""
                    h_fmt = h_fmt if h_fmt is not None else h_part
                    return f"{d_fmt} {h_fmt}".strip()
                except Exception:
                    return raw

            def _extract_qtd_diarias(raw_obs):
                obs_up = upper(str(raw_obs or "").strip())
                if not obs_up:
                    return "-"
                for pattern in (
                    r"QTD\s*DIARIAS?\s*[:=-]\s*([0-9]+)",
                    r"\b([0-9]+)\s*DIARIAS?\b",
                    r"\b([0-9]+)\s*AJUDANTES?\b",
                ):
                    m = re.search(pattern, obs_up, flags=re.IGNORECASE)
                    if m:
                        return m.group(1).strip()
                return "-"

            receb_assinaturas_desenhadas = False
            receb_reserva_assinaturas_mm = 70
            receb_reserva_total_mm = 16

            def _draw_receb_assinaturas():
                nonlocal y, receb_assinaturas_desenhadas
                c.setFont("Helvetica-Bold", 10)
                c.drawString(left, y, "ASSINATURAS / CONFERENCIA")
                y -= 10 * mm

                block_w = (width - left - right - 10 * mm) / 2.0
                block_h = 18 * mm
                gap_x = 10 * mm
                gap_y = 10 * mm

                def assinatura_block(x, y_top, titulo):
                    c.setLineWidth(0.8)
                    c.rect(x, y_top - block_h, block_w, block_h, stroke=1, fill=0)
                    c.setFont("Helvetica-Bold", 9)
                    c.drawString(x + 3 * mm, y_top - 5 * mm, titulo)
                    c.setFont("Helvetica", 9)
                    c.line(x + 3 * mm, y_top - 13 * mm, x + block_w - 3 * mm, y_top - 13 * mm)
                    c.drawString(x + 3 * mm, y_top - 16.5 * mm, "Assinatura / Carimbo")

                x1 = left
                x2 = left + block_w + gap_x

                assinatura_block(x1, y, "SETOR FATURAMENTO")
                assinatura_block(x2, y, "SETOR FINANCEIRO")
                y -= (block_h + gap_y)
                assinatura_block(x1, y, "SETOR DE CAIXA")
                assinatura_block(x2, y, "SETOR DE CONFERENCIA")
                receb_assinaturas_desenhadas = True

            def _draw_retorno_sheet():
                nonlocal y
                meta_ret = fetch_programacao_meta_relatorio(prog)
                _itens_ret_res = fetch_programacao_itens_contract(prog)
                itens_ret = (
                    (_itens_ret_res.get("data") or [])
                    if isinstance(_itens_ret_res, dict) and bool(_itens_ret_res.get("ok", False))
                    else []
                )

                media_carreg = safe_float(meta_ret.get("media"), 0.0)
                if media_carreg > 20:
                    media_carreg = media_carreg / 1000.0
                aves_cx = safe_int(meta_ret.get("aves_caixa_final"), 0)
                if aves_cx <= 0:
                    aves_cx = 6

                status_entregue = {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}
                status_cancelado = {"CANCELADO", "CANCELADA"}

                tot_ent = tot_cancel = tot_alt = 0
                tot_cx_prog = tot_cx_ent = 0
                tot_kg_ent = tot_val = 0.0
                tot_mort_aves = safe_int(meta_ret.get("mortalidade_transbordo_aves"), 0)
                tot_mort_kg = safe_float(meta_ret.get("mortalidade_transbordo_kg"), 0.0)
                divergencias = 0
                linhas_ret = []
                ocorrencias_ret = []

                for item in itens_ret:
                    status_item = upper(str(item.get("status_pedido") or "PENDENTE").strip()) or "PENDENTE"
                    cx_prog = max(safe_int(item.get("qnt_caixas"), 0), 0)
                    cx_alt = max(safe_int(item.get("caixas_atual"), 0), 0)
                    cx_ent = 0 if status_item in status_cancelado else (cx_alt if cx_alt > 0 else (cx_prog if status_item in status_entregue else 0))
                    preco_orig = safe_float(item.get("preco"), 0.0)
                    preco_final = safe_float(item.get("preco_atual"), 0.0) or preco_orig
                    kg_orig = safe_float(item.get("kg"), 0.0)
                    kg_base = safe_float(item.get("peso_previsto"), 0.0) or kg_orig
                    if kg_base <= 0 and media_carreg > 0:
                        base_cx = cx_ent if cx_ent > 0 else cx_prog
                        kg_base = float(base_cx) * float(max(aves_cx, 1)) * media_carreg
                    aves_total = max((cx_ent if cx_ent > 0 else cx_prog) * max(aves_cx, 1), 0)
                    media_cli = (kg_base / aves_total) if aves_total > 0 and kg_base > 0 else media_carreg
                    mort_aves = max(safe_int(item.get("mortalidade_aves"), 0), 0)
                    mort_kg = mort_aves * media_cli if media_cli > 0 else 0.0
                    kg_ent = kg_base if status_item not in status_cancelado else 0.0
                    valor_total = max(kg_ent * preco_final, 0.0) if status_item not in status_cancelado else 0.0
                    alterado = bool(
                        str(item.get("alterado_em") or "").strip()
                        or str(item.get("alterado_por") or "").strip()
                        or str(item.get("alteracao_tipo") or "").strip()
                        or str(item.get("alteracao_detalhe") or "").strip()
                        or (cx_alt > 0 and cx_alt != cx_prog)
                        or abs(preco_final - preco_orig) > 0.0001
                    )
                    delta_kg = kg_ent - kg_orig if kg_orig > 0 else 0.0

                    tot_cx_prog += cx_prog
                    tot_cx_ent += cx_ent
                    tot_kg_ent += kg_ent
                    tot_val += valor_total
                    tot_mort_aves += mort_aves
                    tot_mort_kg += mort_kg
                    if status_item in status_entregue:
                        tot_ent += 1
                    if status_item in status_cancelado:
                        tot_cancel += 1
                    if alterado:
                        tot_alt += 1
                    if abs(delta_kg) >= 0.01:
                        divergencias += 1

                    linhas_ret.append(
                        (
                            str(item.get("cod_cliente") or "")[:6],
                            upper(str(item.get("nome_cliente") or "-"))[:24],
                            status_item[:10],
                            f"{cx_prog}/{cx_ent}",
                            f"{safe_float(media_cli, 0.0):.3f}",
                            f"{safe_float(kg_ent, 0.0):.2f}",
                            f"{safe_float(preco_final, 0.0):.2f}",
                            f"{safe_float(valor_total, 0.0):.2f}",
                            f"{safe_int(mort_aves, 0)}/{safe_float(mort_kg, 0.0):.2f}",
                            f"{delta_kg:+.2f}" if abs(delta_kg) >= 0.01 else "-",
                        )
                    )

                    if status_item in status_cancelado or alterado:
                        para_quem, motivo, autorizado = _parse_alteracao_campos(item)
                        ocorrencias_ret.append(
                            (
                                upper(str(item.get("nome_cliente") or "-"))[:20],
                                status_item[:10],
                                "SIM" if status_item in status_cancelado else "NAO",
                                "SIM" if alterado else "NAO",
                                para_quem[:14],
                                motivo[:28],
                                autorizado[:14],
                            )
                        )

                kg_carregado_ret = safe_float(meta_ret.get("nf_kg_carregado"), 0.0)
                if kg_carregado_ret <= 0:
                    kg_carregado_ret = safe_float(meta_ret.get("nf_kg"), 0.0)
                caixas_carreg_ret = safe_int(meta_ret.get("nf_caixas"), 0)
                if caixas_carreg_ret <= 0:
                    caixas_carreg_ret = tot_cx_prog

                data_saida_ret, hora_saida_ret = normalize_date_time_components(meta_ret.get("data_saida"), meta_ret.get("hora_saida"))
                data_chegada_ret, hora_chegada_ret = normalize_date_time_components(meta_ret.get("data_chegada"), meta_ret.get("hora_chegada"))
                c.showPage()
                y = height - top
                c.setFont("Helvetica-Bold", 14)
                c.drawString(left, y, f"FOLHA DE RETORNO OPERACIONAL - PROGRAMACAO {prog}")
                y -= 8 * mm

                c.setFont("Helvetica", 9)
                draw_kv(col1_x, y, "Motorista", str(meta_ret.get("motorista") or motorista), col1_w)
                draw_kv(col2_x, y, "Veiculo", str(meta_ret.get("veiculo") or veiculo), col2_w)
                y -= line_h
                draw_kv(col1_x, y, "Equipe", resolve_equipe_nomes(str(meta_ret.get("equipe") or equipe)), col1_w)
                draw_kv(col2_x, y, "NF", str(meta_ret.get("num_nf") or nf), col2_w)
                y -= line_h
                draw_kv(col1_x, y, "Local Rota", str(meta_ret.get("local_rota") or local_rota), col1_w)
                draw_kv(col2_x, y, "Carregou em", str(meta_ret.get("local_carregamento") or local_carreg), col2_w)
                y -= line_h
                draw_kv(col1_x, y, "Saida", f"{data_saida_ret} {hora_saida_ret}".strip(), col1_w)
                draw_kv(col2_x, y, "Chegada", f"{data_chegada_ret} {hora_chegada_ret}".strip(), col2_w)
                y -= 8 * mm

                box_gap = 4 * mm
                box_w = (table_w - (box_gap * 3)) / 4.0
                box_h = 13 * mm

                def _summary_box(x, y_top, titulo, valor):
                    c.rect(x, y_top - box_h, box_w, box_h, stroke=1, fill=0)
                    c.setFont("Helvetica-Bold", 8)
                    c.drawString(x + 2 * mm, y_top - 4.5 * mm, titulo)
                    c.setFont("Helvetica", 11)
                    c.drawRightString(x + box_w - 2 * mm, y_top - 9.5 * mm, valor)

                _summary_box(left, y, "Clientes", str(len(itens_ret)))
                _summary_box(left + box_w + box_gap, y, "Entregues/Cancel.", f"{tot_ent}/{tot_cancel}")
                _summary_box(left + ((box_w + box_gap) * 2), y, "CX Carreg./Entreg.", f"{caixas_carreg_ret}/{tot_cx_ent}")
                _summary_box(left + ((box_w + box_gap) * 3), y, "KG Carreg./Entreg.", f"{kg_carregado_ret:.2f}/{tot_kg_ent:.2f}")
                y -= (box_h + 6 * mm)

                _summary_box(left, y, "Media Carreg.", f"{media_carreg:.3f}")
                _summary_box(left + box_w + box_gap, y, "Mortalidade", f"{tot_mort_aves} / {tot_mort_kg:.2f}kg")
                _summary_box(left + ((box_w + box_gap) * 2), y, "Valor Entregue", self._fmt_money(tot_val))
                _summary_box(left + ((box_w + box_gap) * 3), y, "Diverg. Peso", str(divergencias))
                y -= (box_h + 7 * mm)

                c.setFont("Helvetica-Bold", 10)
                c.drawString(left, y, "ENTREGAS POR CLIENTE")
                y -= 5.5 * mm

                row_h_ret = 5.8 * mm
                col_cod_ret = table_w * 0.08
                col_cli_ret = table_w * 0.21
                col_st_ret = table_w * 0.11
                col_cx_ret = table_w * 0.08
                col_med_ret = table_w * 0.08
                col_kg_ret = table_w * 0.10
                col_preco_ret = table_w * 0.10
                col_val_ret = table_w * 0.12
                col_mort_ret = table_w * 0.08
                col_div_ret = table_w - sum([col_cod_ret, col_cli_ret, col_st_ret, col_cx_ret, col_med_ret, col_kg_ret, col_preco_ret, col_val_ret, col_mort_ret])

                x_cod_ret = table_x
                x_cli_ret = x_cod_ret + col_cod_ret
                x_st_ret = x_cli_ret + col_cli_ret
                x_cx_ret = x_st_ret + col_st_ret
                x_med_ret = x_cx_ret + col_cx_ret
                x_kg_ret = x_med_ret + col_med_ret
                x_preco_ret = x_kg_ret + col_kg_ret
                x_val_ret = x_preco_ret + col_preco_ret
                x_mort_ret = x_val_ret + col_val_ret
                x_div_ret = x_mort_ret + col_mort_ret

                def _draw_ret_header(title_suffix=""):
                    nonlocal y
                    c.setFont("Helvetica-Bold", 7)
                    c.rect(table_x, y - row_h_ret + 1, table_w, row_h_ret, stroke=1, fill=0)
                    c.drawString(x_cod_ret + 2, y - row_h_ret + 3, "COD")
                    c.drawString(x_cli_ret + 2, y - row_h_ret + 3, "CLIENTE")
                    c.drawString(x_st_ret + 2, y - row_h_ret + 3, "STATUS")
                    c.drawRightString(x_cx_ret + col_cx_ret - 2, y - row_h_ret + 3, "CX")
                    c.drawRightString(x_med_ret + col_med_ret - 2, y - row_h_ret + 3, "MEDIA")
                    c.drawRightString(x_kg_ret + col_kg_ret - 2, y - row_h_ret + 3, "KG")
                    c.drawRightString(x_preco_ret + col_preco_ret - 2, y - row_h_ret + 3, "PRECO")
                    c.drawRightString(x_val_ret + col_val_ret - 2, y - row_h_ret + 3, "VALOR")
                    c.drawRightString(x_mort_ret + col_mort_ret - 2, y - row_h_ret + 3, "MORT")
                    c.drawRightString(x_div_ret + col_div_ret - 2, y - row_h_ret + 3, "DIV")
                    y -= row_h_ret
                    c.setFont("Helvetica", 7)
                    if title_suffix:
                        c.setFont("Helvetica-Bold", 12)
                        c.drawString(left, height - top, title_suffix)
                        c.setFont("Helvetica", 7)

                _draw_ret_header()
                if not linhas_ret:
                    c.rect(table_x, y - row_h_ret + 1, table_w, row_h_ret, stroke=1, fill=0)
                    c.drawString(x_cod_ret + 2, y - row_h_ret + 3, "SEM ENTREGAS REGISTRADAS")
                    y -= row_h_ret
                else:
                    for row in linhas_ret:
                        if y < bottom + 58 * mm:
                            c.showPage()
                            y = height - top - 8 * mm
                            c.setFont("Helvetica-Bold", 12)
                            c.drawString(left, height - top, f"FOLHA DE RETORNO OPERACIONAL - PROGRAMACAO {prog} (CONT.)")
                            _draw_ret_header()
                        c.rect(table_x, y - row_h_ret + 1, table_w, row_h_ret, stroke=1, fill=0)
                        c.drawString(x_cod_ret + 2, y - row_h_ret + 3, _clip_width(row[0], col_cod_ret - 4, "Helvetica", 7))
                        c.drawString(x_cli_ret + 2, y - row_h_ret + 3, _clip_width(row[1], col_cli_ret - 4, "Helvetica", 7))
                        c.drawString(x_st_ret + 2, y - row_h_ret + 3, _clip_width(row[2], col_st_ret - 4, "Helvetica", 7))
                        c.drawRightString(x_cx_ret + col_cx_ret - 2, y - row_h_ret + 3, row[3])
                        c.drawRightString(x_med_ret + col_med_ret - 2, y - row_h_ret + 3, row[4].replace(".", ","))
                        c.drawRightString(x_kg_ret + col_kg_ret - 2, y - row_h_ret + 3, row[5].replace(".", ","))
                        c.drawRightString(x_preco_ret + col_preco_ret - 2, y - row_h_ret + 3, row[6].replace(".", ","))
                        c.drawRightString(x_val_ret + col_val_ret - 2, y - row_h_ret + 3, row[7].replace(".", ","))
                        c.drawRightString(x_mort_ret + col_mort_ret - 2, y - row_h_ret + 3, row[8].replace(".", ","))
                        c.drawRightString(x_div_ret + col_div_ret - 2, y - row_h_ret + 3, row[9].replace(".", ","))
                        y -= row_h_ret

                y -= 6 * mm
                c.setFont("Helvetica-Bold", 10)
                c.drawString(left, y, "OCORRENCIAS / AJUSTES")
                y -= 5.5 * mm

                row_h_occ = 5.8 * mm
                occ_cols = [0.21, 0.10, 0.06, 0.06, 0.15, 0.26, 0.16]
                occ_ws = [table_w * v for v in occ_cols]
                occ_ws[-1] = table_w - sum(occ_ws[:-1])
                x_occ = [table_x]
                for wv in occ_ws[:-1]:
                    x_occ.append(x_occ[-1] + wv)

                def _draw_occ_header():
                    nonlocal y
                    c.setFont("Helvetica-Bold", 7)
                    c.rect(table_x, y - row_h_occ + 1, table_w, row_h_occ, stroke=1, fill=0)
                    headers = ["CLIENTE", "STATUS", "CANC", "ALT", "PARA", "MOTIVO", "AUTORIZOU"]
                    for idx_h, head in enumerate(headers):
                        c.drawString(x_occ[idx_h] + 2, y - row_h_occ + 3, head)
                    y -= row_h_occ
                    c.setFont("Helvetica", 7)

                _draw_occ_header()
                if not ocorrencias_ret:
                    c.rect(table_x, y - row_h_occ + 1, table_w, row_h_occ, stroke=1, fill=0)
                    c.drawString(table_x + 2, y - row_h_occ + 3, "SEM OCORRENCIAS DE CANCELAMENTO/ALTERACAO")
                    y -= row_h_occ
                else:
                    for row in ocorrencias_ret:
                        if y < bottom + 20 * mm:
                            c.showPage()
                            y = height - top - 8 * mm
                            c.setFont("Helvetica-Bold", 12)
                            c.drawString(left, height - top, f"OCORRENCIAS - RETORNO OPERACIONAL {prog} (CONT.)")
                            _draw_occ_header()
                        c.rect(table_x, y - row_h_occ + 1, table_w, row_h_occ, stroke=1, fill=0)
                        for idx_v, value in enumerate(row):
                            txt = _clip_width(str(value), occ_ws[idx_v] - 4, "Helvetica", 7)
                            if idx_v in {2, 3}:
                                c.drawCentredString(x_occ[idx_v] + (occ_ws[idx_v] / 2.0), y - row_h_occ + 3, txt)
                            else:
                                c.drawString(x_occ[idx_v] + 2, y - row_h_occ + 3, txt)
                        y -= row_h_occ

            # ===== FOLHA 1: RECEBIMENTOS =====
            c.setFont("Helvetica-Bold", 14)
            c.drawString(left, y, f"FOLHA DE RECEBIMENTOS - PROGRAMACAO {prog}")
            y -= 8 * mm

            c.setFont("Helvetica", 9)
            col1_x = left
            line_h = 5.2 * mm

            draw_kv(col1_x, y, "Motorista", motorista, col1_w)
            draw_kv(col2_x, y, "Veiculo", veiculo, col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Equipe", equipe_txt, col1_w)
            draw_kv(col2_x, y, "NF", nf, col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Saida", f"{data_saida} {hora_saida}".strip(), col1_w)
            draw_kv(col2_x, y, "Chegada", f"{data_chegada} {hora_chegada}".strip(), col2_w)
            y -= 8 * mm

            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "RECEBIMENTOS REGISTRADOS")
            y -= 6 * mm

            table_x = left
            table_w = width - left - right
            col_cod = table_w * 0.12
            col_nome = table_w * 0.28
            col_forma = table_w * 0.14
            col_valor = table_w * 0.12
            col_obs = table_w * 0.18
            col_data = table_w * 0.16

            x_cod = table_x
            x_nome = x_cod + col_cod
            x_forma = x_nome + col_nome
            x_valor = x_forma + col_forma
            x_obs = x_valor + col_valor
            x_data = x_obs + col_obs
            row_h_rec = 6.2 * mm

            def _draw_receb_header():
                nonlocal y
                c.setFont("Helvetica-Bold", 8)
                c.rect(table_x, y - row_h_rec + 1, table_w, row_h_rec, stroke=1, fill=0)
                c.drawString(x_cod + 2, y - row_h_rec + 3, "COD")
                c.drawString(x_nome + 2, y - row_h_rec + 3, "CLIENTE")
                c.drawString(x_forma + 2, y - row_h_rec + 3, "FORMA")
                c.drawRightString(x_valor + col_valor - 2, y - row_h_rec + 3, "VALOR")
                c.drawString(x_obs + 2, y - row_h_rec + 3, "OBS")
                c.drawRightString(x_data + col_data - 2, y - row_h_rec + 3, "DATA")
                y -= row_h_rec
                c.setFont("Helvetica", 8)

            def _fmt_data_rec(v):
                return _fmt_pdf_datetime(v)

            _draw_receb_header()

            total_receb_detalhe = 0.0
            if not recebimentos:
                c.rect(table_x, y - row_h_rec + 1, table_w, row_h_rec, stroke=1, fill=0)
                c.drawString(x_cod + 2, y - row_h_rec + 3, "SEM RECEBIMENTOS REGISTRADOS")
                y -= row_h_rec
            else:
                for cod_cli, nome_cli, valor_rec, forma_rec, obs_rec, data_rec in recebimentos:
                    reserva_mm = receb_reserva_assinaturas_mm if not receb_assinaturas_desenhadas else receb_reserva_total_mm
                    if (y - row_h_rec) < bottom + (reserva_mm * mm):
                        if not receb_assinaturas_desenhadas:
                            _draw_receb_assinaturas()
                        new_page(f"FOLHA DE RECEBIMENTOS - PROGRAMACAO {prog} (CONT.)")
                        _draw_receb_header()

                    c.rect(table_x, y - row_h_rec + 1, table_w, row_h_rec, stroke=1, fill=0)
                    cod_txt = _clip_width(str(cod_cli or ""), col_cod - 4, "Helvetica", 8)
                    nome_txt = _clip_width(str(nome_cli or ""), col_nome - 4, "Helvetica", 8)
                    forma_txt = _clip_width(str(forma_rec or ""), col_forma - 4, "Helvetica", 8)
                    obs_txt = _clip_width(str(obs_rec or ""), col_obs - 4, "Helvetica", 8)
                    c.drawString(x_cod + 2, y - row_h_rec + 3, cod_txt)
                    c.drawString(x_nome + 2, y - row_h_rec + 3, nome_txt)
                    c.drawString(x_forma + 2, y - row_h_rec + 3, forma_txt)
                    c.drawRightString(
                        x_valor + col_valor - 2,
                        y - row_h_rec + 3,
                        f"{float(valor_rec or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    )
                    c.drawString(x_obs + 2, y - row_h_rec + 3, obs_txt)
                    data_rec_txt = _clip_width(_fmt_data_rec(data_rec), col_data - 4, "Helvetica", 7)
                    c.drawRightString(x_data + col_data - 2, y - row_h_rec + 3, data_rec_txt)
                    y -= row_h_rec
                    total_receb_detalhe += safe_float(valor_rec, 0.0)

            y -= 4 * mm
            c.setFont("Helvetica-Bold", 10)
            c.drawRightString(width - right, y, f"TOTAL RECEBIMENTOS: {self._fmt_money(total_receb_detalhe)}")
            y -= 8 * mm

            if not receb_assinaturas_desenhadas:
                _draw_receb_assinaturas()

            _draw_retorno_sheet()

            meta_desp_pdf = fetch_programacao_meta_relatorio(prog)
            _itens_desp_res = fetch_programacao_itens_contract(prog)
            itens_desp_pdf = (
                (_itens_desp_res.get("data") or [])
                if isinstance(_itens_desp_res, dict) and bool(_itens_desp_res.get("ok", False))
                else []
            )
            media_desp_pdf = self._normalize_media_kg_ave(safe_float(meta_desp_pdf.get("media"), 0.0))
            if media_desp_pdf <= 0:
                media_desp_pdf = self._normalize_media_kg_ave(nf_media_carregada)
            data_saida_pdf, hora_saida_pdf = normalize_date_time_components(
                meta_desp_pdf.get("data_saida"),
                meta_desp_pdf.get("hora_saida"),
            )
            data_chegada_pdf, hora_chegada_pdf = normalize_date_time_components(
                meta_desp_pdf.get("data_chegada"),
                meta_desp_pdf.get("hora_chegada"),
            )
            aves_cx_desp_pdf = safe_int(meta_desp_pdf.get("aves_caixa_final"), 0)
            if aves_cx_desp_pdf <= 0:
                aves_cx_desp_pdf = safe_int(nf_caixa_final, 0)
            if aves_cx_desp_pdf <= 0:
                aves_cx_desp_pdf = 6

            tot_cx_prog_pdf = 0
            tot_kg_ent_pdf = 0.0
            status_entregue_pdf = {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÍDO"}
            status_cancelado_pdf = {"CANCELADO", "CANCELADA"}
            for item_pdf in itens_desp_pdf:
                status_item_pdf = upper(str(item_pdf.get("status_pedido") or "PENDENTE").strip()) or "PENDENTE"
                cx_prog_pdf = max(safe_int(item_pdf.get("qnt_caixas"), 0), 0)
                cx_alt_pdf = max(safe_int(item_pdf.get("caixas_atual"), 0), 0)
                cx_ent_pdf = 0 if status_item_pdf in status_cancelado_pdf else (
                    cx_alt_pdf if cx_alt_pdf > 0 else (cx_prog_pdf if status_item_pdf in status_entregue_pdf else 0)
                )
                kg_base_pdf = safe_float(item_pdf.get("peso_previsto"), 0.0) or safe_float(item_pdf.get("kg"), 0.0)
                if kg_base_pdf <= 0 and media_desp_pdf > 0:
                    base_cx_pdf = cx_ent_pdf if cx_ent_pdf > 0 else cx_prog_pdf
                    kg_base_pdf = float(base_cx_pdf) * float(max(aves_cx_desp_pdf, 1)) * media_desp_pdf
                aves_total_pdf = max((cx_ent_pdf if cx_ent_pdf > 0 else cx_prog_pdf) * max(aves_cx_desp_pdf, 1), 0)
                media_cli_pdf = (kg_base_pdf / aves_total_pdf) if aves_total_pdf > 0 and kg_base_pdf > 0 else media_desp_pdf
                mort_aves_pdf = max(safe_int(item_pdf.get("mortalidade_aves"), 0), 0)
                mort_kg_pdf = mort_aves_pdf * media_cli_pdf if media_cli_pdf > 0 else 0.0
                kg_ent_pdf = kg_base_pdf if status_item_pdf not in status_cancelado_pdf else 0.0
                tot_cx_prog_pdf += cx_prog_pdf
                tot_kg_ent_pdf += kg_ent_pdf

            if not motorista:
                motorista = str(meta_desp_pdf.get("motorista") or "")
            if not veiculo:
                veiculo = str(meta_desp_pdf.get("veiculo") or "")
            if not equipe_txt:
                equipe_txt = resolve_equipe_nomes(str(meta_desp_pdf.get("equipe") or equipe))
            if not nf:
                nf = str(meta_desp_pdf.get("num_nf") or "")
            if not local_rota:
                local_rota = str(meta_desp_pdf.get("local_rota") or "")
            if not local_carreg:
                local_carreg = str(meta_desp_pdf.get("local_carregamento") or "")
            if nf_kg <= 0:
                nf_kg = safe_float(meta_desp_pdf.get("nf_kg"), 0.0)
            if nf_preco <= 0:
                nf_preco = safe_float(meta_desp_pdf.get("nf_preco"), 0.0)
            if not data_saida:
                data_saida = data_saida_pdf
            if not hora_saida:
                hora_saida = hora_saida_pdf
            if not data_chegada:
                data_chegada = data_chegada_pdf
            if not hora_chegada:
                hora_chegada = hora_chegada_pdf
            if nf_media_carregada <= 0:
                nf_media_carregada = media_desp_pdf
            if nf_caixas <= 0:
                nf_caixas_meta = safe_int(meta_desp_pdf.get("nf_caixas"), 0)
                nf_caixas = nf_caixas_meta if nf_caixas_meta > 0 else tot_cx_prog_pdf
            if nf_caixa_final <= 0:
                nf_caixa_final_meta = safe_int(meta_desp_pdf.get("aves_caixa_final"), 0)
                if nf_caixa_final_meta > 0:
                    nf_caixa_final = nf_caixa_final_meta
            if nf_kg_carregado <= 0:
                nf_kg_carregado = safe_float(meta_desp_pdf.get("nf_kg_carregado"), 0.0)
                if nf_kg_carregado <= 0:
                    nf_kg_carregado = safe_float(meta_desp_pdf.get("nf_kg"), 0.0)
            if nf_kg_vendido <= 0:
                nf_kg_vendido = tot_kg_ent_pdf
            if abs(nf_saldo) < 0.0001 and (nf_kg_carregado > 0 or nf_kg_vendido > 0):
                nf_saldo = nf_kg_carregado - nf_kg_vendido
            if km_rodado <= 0 and km_final > 0 and km_inicial > 0:
                km_rodado = max(km_final - km_inicial, 0.0)
            if media_km_l <= 0 and litros > 0 and km_rodado > 0:
                media_km_l = km_rodado / litros
            if custo_km <= 0 and km_rodado > 0:
                custo_km = total_desp / km_rodado

            # ===== FOLHA 2: DESPESAS =====
            new_page(f"FOLHA DE DESPESAS - PROGRAMACAO {prog}")

            # Dados da rota
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "DADOS DA ROTA")
            y -= 6 * mm

            c.setFont("Helvetica", 9)

            draw_kv(col1_x, y, "Motorista", motorista, col1_w)
            draw_kv(col2_x, y, "Veiculo", veiculo, col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Equipe", equipe_txt, col1_w)
            draw_kv(col2_x, y, "NF", nf, col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Local da Rota", local_rota, col1_w)
            draw_kv(col2_x, y, "Local Carregamento", local_carreg, col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Saida", f"{data_saida} {hora_saida}".strip(), col1_w)
            draw_kv(col2_x, y, "Chegada", f"{data_chegada} {hora_chegada}".strip(), col2_w)
            y -= 8 * mm

            # Carregamentos / NF
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "CARREGAMENTO / NOTA FISCAL")
            y -= 6 * mm

            draw_kv(col1_x, y, "NF KG", f"{nf_kg:.2f}".replace(".", ","), col1_w)
            draw_kv(col2_x, y, "NF Caixas", str(nf_caixas), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "KG Carregado", f"{nf_kg_carregado:.2f}".replace(".", ","), col1_w)
            draw_kv(col2_x, y, "KG Vendido", f"{nf_kg_vendido:.2f}".replace(".", ","), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Saldo (KG)", f"{nf_saldo:.2f}".replace(".", ","), col1_w)
            draw_kv(col2_x, y, "Preco NF", self._fmt_money(nf_preco), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Media carregada", f"{nf_media_carregada:.3f}".replace(".", ","), col1_w)
            draw_kv(col2_x, y, "Caixa final", str(nf_caixa_final), col2_w)
            y -= 8 * mm

            # Dados de rota (KM)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "DADOS DE ROTA (KM)")
            y -= 6 * mm

            draw_kv(col1_x, y, "KM Inicial", f"{km_inicial:.2f}".replace(".", ","), col1_w)
            draw_kv(col2_x, y, "KM Final", f"{km_final:.2f}".replace(".", ","), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Litros", f"{litros:.2f}".replace(".", ","), col1_w)
            draw_kv(col2_x, y, "KM Rodado", f"{km_rodado:.2f}".replace(".", ","), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Media", f"{media_km_l:.2f}".replace(".", ","), col1_w)
            draw_kv(col2_x, y, "Custo KM", f"{custo_km:.2f}".replace(".", ","), col2_w)
            y -= 8 * mm

            # Contagem de cedulas
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "CONTAGEM DE CEDULAS")
            y -= 6 * mm

            c.setFont("Helvetica-Bold", 9)
            c.drawString(left, y, "CEDULA")
            c.drawString(left + 25*mm, y, "QTD")
            c.drawString(left + 45*mm, y, "TOTAL")
            y -= 5 * mm
            c.setFont("Helvetica", 9)
            for ced in [200, 100, 50, 20, 10, 5, 2]:
                total_ced = ced_qtd[ced] * ced
                c.drawString(left, y, f"R$ {ced:.2f}".replace(".", ","))
                c.drawString(left + 25*mm, y, str(ced_qtd[ced]))
                c.drawString(left + 45*mm, y, f"R$ {total_ced:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                y -= 5 * mm

            y -= 2 * mm
            c.setFont("Helvetica-Bold", 9)
            c.drawString(left, y, f"TOTAL DINHEIRO: {self._fmt_money(valor_dinheiro or ced_total)}")
            y -= 8 * mm

            # Despesas
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "DESPESAS")
            y -= 6 * mm

            table_x = left
            table_w = width - left - right
            col_desc = table_w * 0.25
            col_cat = table_w * 0.13
            col_val = table_w * 0.12
            col_obs = table_w * 0.32
            col_data = table_w - (col_desc + col_cat + col_val + col_obs)

            x_desc = table_x
            x_cat = x_desc + col_desc
            x_val = x_cat + col_cat
            x_obs = x_val + col_val
            x_data = x_obs + col_obs

            row_h = 6.5 * mm
            row_line_h = 4.2 * mm

            def _draw_desp_header():
                nonlocal y
                c.setFont("Helvetica-Bold", 8)
                c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
                c.drawString(x_desc + 2, y - row_h + 3, "DESCRICAO")
                c.drawString(x_cat + 2, y - row_h + 3, "CATEGORIA")
                c.drawRightString(x_val + col_val - 2, y - row_h + 3, "VALOR")
                c.drawString(x_obs + 2, y - row_h + 3, "OBS")
                c.drawRightString(x_data + col_data - 2, y - row_h + 3, "DATA")
                y -= row_h
                c.setFont("Helvetica", 8)

            _draw_desp_header()

            for desc, val, cat, obs, data_reg in despesas:
                desc_up = upper(str(desc or "").strip())
                obs_txt = str(obs or "").strip()
                if desc_up in {"DIARIAS MOTORISTA", "DIARIA MOTORISTA"}:
                    m_nome = motorista or "-"
                    qtd_d = _extract_qtd_diarias(obs_txt)
                    obs_txt = f"QTD DIARIAS: {qtd_d} | MOTORISTA: {m_nome}"
                elif desc_up in {"DIARIAS AJUDANTES", "DIARIA AJUDANTE", "DIARIA AJUDANTES"}:
                    a_nome = equipe_txt or "-"
                    qtd_d = _extract_qtd_diarias(obs_txt)
                    obs_txt = f"QTD DIARIAS: {qtd_d} | AJUDANTES: {a_nome}"
                desc_lines = self._wrap_pdf_lines(str(desc or ""), col_desc - 4, font_name="Helvetica", font_size=8) or [""]
                cat_lines = self._wrap_pdf_lines(str(cat or ""), col_cat - 4, font_name="Helvetica", font_size=8) or [""]
                obs_lines = self._wrap_pdf_lines(obs_txt, col_obs - 4, font_name="Helvetica", font_size=8) or [""]
                data_lines = self._wrap_pdf_lines(_fmt_pdf_datetime(data_reg), col_data - 4, font_name="Helvetica", font_size=8) or [""]
                line_count = max(len(desc_lines), len(cat_lines), len(obs_lines), len(data_lines), 1)
                row_h_curr = max(row_h, (line_count * row_line_h) + 4)

                if y < bottom + row_h_curr + (24 * mm):
                    new_page(f"FOLHA DE DESPESAS - PROGRAMACAO {prog} (CONT.)")
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(left, y, "DESPESAS")
                    y -= 6 * mm
                    _draw_desp_header()

                c.rect(table_x, y - row_h_curr + 1, table_w, row_h_curr, stroke=1, fill=0)
                text_y = y - 10
                for idx_line, line in enumerate(desc_lines):
                    c.drawString(x_desc + 2, text_y - (idx_line * row_line_h), line)
                for idx_line, line in enumerate(cat_lines):
                    c.drawString(x_cat + 2, text_y - (idx_line * row_line_h), line)
                c.drawRightString(
                    x_val + col_val - 2,
                    text_y,
                    f"{float(val or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                )
                for idx_line, line in enumerate(obs_lines):
                    c.drawString(x_obs + 2, text_y - (idx_line * row_line_h), line)
                for idx_line, line in enumerate(data_lines):
                    c.drawRightString(x_data + col_data - 2, text_y - (idx_line * row_line_h), line)
                y -= row_h_curr

            y -= 6 * mm

            if y < bottom + (90 * mm):
                new_page(f"FOLHA DE DESPESAS - PROGRAMACAO {prog} (CONT.)")

            # Resumo financeiro
            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "RESUMO FINANCEIRO (CAIXA)")
            y -= 6 * mm

            c.setFont("Helvetica", 9)
            draw_kv(col1_x, y, "Recebimentos", self._fmt_money(total_receb), col1_w)
            draw_kv(col2_x, y, "Adiantamento", self._fmt_money(adiantamento), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Despesas", self._fmt_money(total_desp), col1_w)
            draw_kv(col2_x, y, "Cedulas", self._fmt_money(ced_total), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Total Entradas", self._fmt_money(total_entradas), col1_w)
            draw_kv(col2_x, y, "Total Saidas", self._fmt_money(total_saidas), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Caixa", self._fmt_money(valor_final_caixa), col1_w)
            draw_kv(col2_x, y, "Diferenca", self._fmt_money(diferenca), col2_w)
            y -= line_h
            draw_kv(col1_x, y, "Resultado", self._fmt_money(resultado), col1_w)
            y -= 10 * mm

            c.setFont("Helvetica", 8)
            c.drawRightString(width - right, bottom - 2 * mm, f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")

            c.save()
            messagebox.showinfo("OK", f"PDF gerado com sucesso!\n\nArquivo:\n{os.path.basename(path_pdf)}")
            self.set_status(f"STATUS: PDF impresso/gerado: {os.path.basename(path_pdf)}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao gerar PDF: {str(e)}")

    def _imprimir_tela(self, window):
        messagebox.showinfo(
            "IMPRESSÃO",
            "Função de impressão será implementada na próxima versão.\n\n"
            "Por enquanto, você pode:\n"
            "1. Tirar um print screen desta tela\n"
            "2. Usar o botão 'Exportar Excel' para gerar arquivo\n"
            "3. Salvar como PDF usando Ctrl+P"
        )
        try:
            self.logger.info("Solicitada impressão de relatório (DespesasPage)")
        except Exception:
            logging.debug("Falha ignorada")

    # =========================================================
    # 7.8 SALVAR TUDO (NF / KM / CÉDULAS / ADIANTAMENTO)
    # =========================================================
    def _wrap_pdf_lines(self, text, max_width, font_name="Helvetica", font_size=9):
        from reportlab.pdfbase.pdfmetrics import stringWidth

        lines = []
        for raw in str(text or "").splitlines():
            words = raw.split()
            if not words:
                lines.append("")
                continue
            current = ""
            for word in words:
                candidate = word if not current else f"{current} {word}"
                if stringWidth(candidate, font_name, font_size) <= max_width:
                    current = candidate
                else:
                    if current:
                        lines.append(current)
                    current = word
            if current:
                lines.append(current)
        return lines

    def salvar_tudo(self):
        if not self._can_edit_current_prog():
            return

        prog = self._current_programacao

        nf_numero = self.ent_nf_numero.get().strip()
        nf_kg = safe_float(self.ent_nf_kg.get(), 0.0)
        nf_preco = safe_float(getattr(self, "ent_nf_preco", self.ent_nf_kg).get(), 0.0)
        nf_caixas = safe_int(self.ent_nf_caixas.get(), 0)
        nf_kg_carregado = safe_float(self.ent_nf_kg_carregado.get(), 0.0)
        nf_kg_vendido = safe_float(self.ent_nf_kg_vendido.get(), 0.0)
        nf_saldo = safe_float(self.ent_nf_saldo.get(), 0.0)
        nf_media_carregada = safe_float(getattr(self, "ent_nf_media_carregada", self.ent_nf_saldo).get(), 0.0)
        nf_caixa_final = safe_int(getattr(self, "ent_nf_caixa_final", self.ent_nf_caixas).get(), 0)

        km_inicial = safe_float(self.ent_km_inicial.get(), 0.0)
        km_final = safe_float(self.ent_km_final.get(), 0.0)
        litros = safe_float(self.ent_litros.get(), 0.0)
        km_rodado = km_final - km_inicial
        if km_rodado < 0:
            km_rodado = 0.0
        media_km_l = (km_rodado / litros) if litros > 0 else 0.0
        total_despesas_prog = 0.0
        try:
            if prog:
                used_api = False
                bundle = self._api_bundle_prestacao(prog)
                if bundle:
                    despesas_api = bundle.get("despesas") if isinstance(bundle, dict) else []
                    total_despesas_prog = sum(
                        safe_float((d.get("valor") if isinstance(d, dict) else 0), 0.0)
                        for d in (despesas_api or [])
                    )
                    used_api = True

                if not used_api:
                    with get_db() as conn_tmp:
                        cur_tmp = conn_tmp.cursor()
                        cur_tmp.execute(
                            "SELECT COALESCE(SUM(valor), 0) FROM despesas WHERE codigo_programacao=?",
                            (prog,),
                        )
                        row_tmp = cur_tmp.fetchone()
                        total_despesas_prog = safe_float((row_tmp[0] if row_tmp else 0), 0.0)
        except Exception:
            logging.debug("Falha ignorada")
        custo_km = (total_despesas_prog / km_rodado) if km_rodado > 0 else 0.0

        try:
            self._safe_set_entry(self.ent_km_rodado, f"{km_rodado:.1f}")
            self._safe_set_entry(self.ent_media, f"{media_km_l:.1f}")
            self._safe_set_entry(self.ent_custo_km, f"{custo_km:.2f}")
        except Exception:
            logging.debug("Falha ignorada")
        rota_observacao = ""
        if hasattr(self, "txt_rota_obs") and self.txt_rota_obs:
            try:
                rota_observacao = self.txt_rota_obs.get("1.0", "end-1c").strip()
            except Exception:
                rota_observacao = ""

        cedulas_data = {f"ced_{ced}_qtd": safe_int(ent.get(), 0) for ced, ent in self.ced_entries.items()}

        valor_dinheiro = 0.0
        for ced, ent in self.ced_entries.items():
            valor_dinheiro += safe_int(ent.get(), 0) * float(ced)

        # âÅâ€œââ‚¬¦ aceita "10,00" ou "R$ 10,00"
        adiantamento_val = self._parse_money_local(self.ent_adiantamento.get())

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                _call_api(
                    "PUT",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/financeiro",
                    payload={
                        "nf_numero": nf_numero,
                        "nf_kg": nf_kg,
                        "nf_caixas": nf_caixas,
                        "nf_kg_carregado": nf_kg_carregado,
                        "nf_kg_vendido": nf_kg_vendido,
                        "nf_saldo": nf_saldo,
                        "nf_preco": nf_preco,
                        "media": nf_media_carregada,
                        "nf_caixa_final": nf_caixa_final,
                        "km_inicial": km_inicial,
                        "km_final": km_final,
                        "litros": litros,
                        "km_rodado": km_rodado,
                        "media_km_l": media_km_l,
                        "custo_km": custo_km,
                        "ced_200_qtd": cedulas_data.get("ced_200_qtd", 0),
                        "ced_100_qtd": cedulas_data.get("ced_100_qtd", 0),
                        "ced_50_qtd": cedulas_data.get("ced_50_qtd", 0),
                        "ced_20_qtd": cedulas_data.get("ced_20_qtd", 0),
                        "ced_10_qtd": cedulas_data.get("ced_10_qtd", 0),
                        "ced_5_qtd": cedulas_data.get("ced_5_qtd", 0),
                        "ced_2_qtd": cedulas_data.get("ced_2_qtd", 0),
                        "valor_dinheiro": valor_dinheiro,
                        "adiantamento": adiantamento_val,
                        "rota_observacao": rota_observacao,
                    },
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                self.logger.info(f"Dados salvos via API para programação {prog}")
                messagebox.showinfo("SUCESSO", "Todos os dados foram salvos com sucesso!")
                self.set_status(f"STATUS: Dados salvos para {prog}")
                self._refresh_km_insights()
                self._refresh_nf_trade_summary()
                return
            except Exception:
                logging.debug("Falha ao salvar dados financeiros via API.", exc_info=True)
                messagebox.showerror(
                    "ERRO",
                    "Não foi possível salvar os dados financeiros na API central.\n\n"
                    "A gravação local foi bloqueada para evitar divergência entre servidor e desktop.",
                )
                return

        try:
            with get_db() as conn:
                cur = conn.cursor()

                cur.execute("PRAGMA table_info(programacoes)")
                columns = [col[1] for col in cur.fetchall()]

                if "adiantamento" in columns:
                    cur.execute("""
                        UPDATE programacoes SET
                            nf_numero=?, nf_kg=?, nf_caixas=?, nf_kg_carregado=?, nf_kg_vendido=?, nf_saldo=?,
                            km_inicial=?, km_final=?, litros=?, km_rodado=?, media_km_l=?, custo_km=?,
                            ced_200_qtd=?, ced_100_qtd=?, ced_50_qtd=?, ced_20_qtd=?, ced_10_qtd=?, ced_5_qtd=?, ced_2_qtd=?,
                            valor_dinheiro=?,
                            adiantamento=?
                        WHERE codigo_programacao=?
                    """, (
                        nf_numero, nf_kg, nf_caixas, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                        km_inicial, km_final, litros, km_rodado, media_km_l, custo_km,
                        cedulas_data.get("ced_200_qtd", 0), cedulas_data.get("ced_100_qtd", 0),
                        cedulas_data.get("ced_50_qtd", 0), cedulas_data.get("ced_20_qtd", 0),
                        cedulas_data.get("ced_10_qtd", 0), cedulas_data.get("ced_5_qtd", 0),
                        cedulas_data.get("ced_2_qtd", 0),
                        valor_dinheiro,
                        adiantamento_val,
                        prog
                    ))
                elif "adiantamento_rota" in columns:
                    cur.execute("""
                        UPDATE programacoes SET
                            nf_numero=?, nf_kg=?, nf_caixas=?, nf_kg_carregado=?, nf_kg_vendido=?, nf_saldo=?,
                            km_inicial=?, km_final=?, litros=?, km_rodado=?, media_km_l=?, custo_km=?,
                            ced_200_qtd=?, ced_100_qtd=?, ced_50_qtd=?, ced_20_qtd=?, ced_10_qtd=?, ced_5_qtd=?, ced_2_qtd=?,
                            valor_dinheiro=?,
                            adiantamento_rota=?
                        WHERE codigo_programacao=?
                    """, (
                        nf_numero, nf_kg, nf_caixas, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                        km_inicial, km_final, litros, km_rodado, media_km_l, custo_km,
                        cedulas_data.get("ced_200_qtd", 0), cedulas_data.get("ced_100_qtd", 0),
                        cedulas_data.get("ced_50_qtd", 0), cedulas_data.get("ced_20_qtd", 0),
                        cedulas_data.get("ced_10_qtd", 0), cedulas_data.get("ced_5_qtd", 0),
                        cedulas_data.get("ced_2_qtd", 0),
                        valor_dinheiro,
                        adiantamento_val,
                        prog
                    ))
                else:
                    cur.execute("""
                        UPDATE programacoes SET
                            nf_numero=?, nf_kg=?, nf_caixas=?, nf_kg_carregado=?, nf_kg_vendido=?, nf_saldo=?,
                            km_inicial=?, km_final=?, litros=?, km_rodado=?, media_km_l=?, custo_km=?,
                            ced_200_qtd=?, ced_100_qtd=?, ced_50_qtd=?, ced_20_qtd=?, ced_10_qtd=?, ced_5_qtd=?, ced_2_qtd=?,
                            valor_dinheiro=?
                        WHERE codigo_programacao=?
                    """, (
                        nf_numero, nf_kg, nf_caixas, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                        km_inicial, km_final, litros, km_rodado, media_km_l, custo_km,
                        cedulas_data.get("ced_200_qtd", 0), cedulas_data.get("ced_100_qtd", 0),
                        cedulas_data.get("ced_50_qtd", 0), cedulas_data.get("ced_20_qtd", 0),
                        cedulas_data.get("ced_10_qtd", 0), cedulas_data.get("ced_5_qtd", 0),
                        cedulas_data.get("ced_2_qtd", 0),
                        valor_dinheiro,
                        prog
                    ))

                if "rota_observacao" in columns:
                    cur.execute(
                        "UPDATE programacoes SET rota_observacao=? WHERE codigo_programacao=?",
                        (rota_observacao, prog),
                    )
                if "nf_preco" in columns:
                    cur.execute(
                        "UPDATE programacoes SET nf_preco=? WHERE codigo_programacao=?",
                        (nf_preco, prog),
                    )
                if "media" in columns:
                    cur.execute(
                        "UPDATE programacoes SET media=? WHERE codigo_programacao=?",
                        (nf_media_carregada, prog),
                    )
                if "aves_caixa_final" in columns:
                    cur.execute(
                        "UPDATE programacoes SET aves_caixa_final=? WHERE codigo_programacao=?",
                        (nf_caixa_final, prog),
                    )
                if "qnt_aves_caixa_final" in columns:
                    cur.execute(
                        "UPDATE programacoes SET qnt_aves_caixa_final=? WHERE codigo_programacao=?",
                        (nf_caixa_final, prog),
                    )

            self.logger.info(f"Dados salvos para programação {prog}")
            messagebox.showinfo("SUCESSO", "Todos os dados foram salvos com sucesso!")
            self.set_status(f"STATUS: Dados salvos para {prog}")
            self._refresh_km_insights()
            self._refresh_nf_trade_summary()

        except Exception as e:
            self.logger.error(f"Erro ao salvar dados: {str(e)}")
            messagebox.showerror("ERRO", f"Erro ao salvar dados: {str(e)}")

    # =========================================================
    # 7.9 EXPORTAR EXCEL (RELATÓRIO COMPLETO)
    # =========================================================
    def exportar_excel(self):
        # exportar pode ser permitido mesmo com FECHADA (é leitura). Então NÃO bloqueio.
        prog = self._current_programacao
        if not (require_pandas() and require_openpyxl()):
            return
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Selecione a Programação primeiro.")
            return

        path = filedialog.asksaveasfilename(
            title="Exportar Relatório Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"RELATORIO_DESPESAS_{prog}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        if not path:
            return


        try:
            used_api = False
            bundle = self._api_bundle_prestacao(prog)
            if bundle:
                rota = bundle.get("rota") if isinstance(bundle, dict) else None
                recebimentos = bundle.get("recebimentos") if isinstance(bundle, dict) else []
                despesas = bundle.get("despesas") if isinstance(bundle, dict) else []
                if isinstance(rota, dict):
                    df_prog = pd.DataFrame([{
                        "codigo_programacao": upper(prog),
                        "motorista": rota.get("motorista"),
                        "veiculo": rota.get("veiculo"),
                        "data_criacao": rota.get("data_criacao") or rota.get("data"),
                        "nf_numero": rota.get("num_nf") or rota.get("nf_numero"),
                        "nf_kg": rota.get("nf_kg"),
                        "nf_caixas": rota.get("nf_caixas"),
                        "nf_kg_carregado": rota.get("nf_kg_carregado"),
                        "nf_kg_vendido": rota.get("nf_kg_vendido"),
                        "nf_saldo": rota.get("nf_saldo"),
                        "km_inicial": rota.get("km_inicial"),
                        "km_final": rota.get("km_final"),
                        "litros": rota.get("litros"),
                        "km_rodado": rota.get("km_rodado"),
                        "media_km_l": rota.get("media_km_l"),
                        "custo_km": rota.get("custo_km"),
                        "valor_dinheiro": rota.get("valor_dinheiro"),
                        "adiantamento": rota.get("adiantamento"),
                    }])
                    df_despesas = pd.DataFrame(
                        [
                            {
                                "id": d.get("id"),
                                "descricao": d.get("descricao"),
                                "valor": d.get("valor"),
                                "categoria": d.get("categoria"),
                                "observacao": d.get("observacao"),
                                "data_registro": d.get("data_registro"),
                            }
                            for d in (despesas or [])
                            if isinstance(d, dict)
                        ]
                    )
                    df_receb = pd.DataFrame(
                        [
                            {
                                "cod_cliente": r.get("cod_cliente"),
                                "nome_cliente": r.get("nome_cliente"),
                                "valor": r.get("valor"),
                                "forma_pagamento": r.get("forma_pagamento"),
                                "observacao": r.get("observacao"),
                                "num_nf": r.get("num_nf"),
                                "data_registro": r.get("data_registro"),
                            }
                            for r in (recebimentos or [])
                            if isinstance(r, dict)
                        ]
                    )
                    used_api = True

            if not used_api:
                with get_db() as conn:
                    cur = conn.cursor()

                    def _table_exists(name: str) -> bool:
                        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,))
                        return bool(cur.fetchone())

                    def _prog_expr(name: str, *, alias: str | None = None, fallback: str = "''") -> str:
                        return f"{name} AS {alias or name}" if db_has_column(cur, "programacoes", name) else f"{fallback} AS {alias or name}"

                    data_expr = _prog_expr("data_criacao", fallback=("data" if db_has_column(cur, "programacoes", "data") else "''"))
                    nf_numero_expr = _prog_expr("nf_numero", fallback=("num_nf" if db_has_column(cur, "programacoes", "num_nf") else "''"))
                    nf_kg_carregado_expr = _prog_expr(
                        "nf_kg_carregado",
                        fallback=("COALESCE(kg_carregado,0)" if db_has_column(cur, "programacoes", "kg_carregado") else "0"),
                    )
                    nf_kg_vendido_expr = _prog_expr(
                        "nf_kg_vendido",
                        fallback=("COALESCE(kg_vendido,0)" if db_has_column(cur, "programacoes", "kg_vendido") else "0"),
                    )
                    adiant_expr = _prog_expr(
                        "adiantamento",
                        fallback=("COALESCE(adiantamento_rota,0)" if db_has_column(cur, "programacoes", "adiantamento_rota") else "0"),
                    )
                    sql_prog = f"""
                        SELECT
                            {_prog_expr('codigo_programacao')},
                            {_prog_expr('motorista')},
                            {_prog_expr('veiculo')},
                            {data_expr},
                            {nf_numero_expr},
                            {_prog_expr('nf_kg', fallback='0')},
                            {_prog_expr('nf_caixas', fallback='0')},
                            {nf_kg_carregado_expr},
                            {nf_kg_vendido_expr},
                            {_prog_expr('nf_saldo', fallback='0')},
                            {_prog_expr('km_inicial', fallback='0')},
                            {_prog_expr('km_final', fallback='0')},
                            {_prog_expr('litros', fallback='0')},
                            {_prog_expr('km_rodado', fallback='0')},
                            {_prog_expr('media_km_l', fallback='0')},
                            {_prog_expr('custo_km', fallback='0')},
                            {_prog_expr('valor_dinheiro', fallback='0')},
                            {adiant_expr}
                        FROM programacoes
                        WHERE codigo_programacao=?
                    """
                    df_prog = pd.read_sql_query(sql_prog, conn, params=(prog,))

                    if _table_exists("despesas"):
                        cols_desp = {str(r[1]).lower() for r in (cur.execute("PRAGMA table_info(despesas)").fetchall() or [])}
                        id_desp_expr = "id" if "id" in cols_desp else "0 AS id"
                        descricao_desp_expr = "COALESCE(descricao,'') AS descricao" if "descricao" in cols_desp else "'' AS descricao"
                        valor_desp_expr = "COALESCE(valor,0) AS valor" if "valor" in cols_desp else "0 AS valor"
                        categoria_desp_expr = "COALESCE(categoria,'') AS categoria" if "categoria" in cols_desp else "'' AS categoria"
                        observacao_desp_expr = "COALESCE(observacao,'') AS observacao" if "observacao" in cols_desp else "'' AS observacao"
                        data_desp_expr = "COALESCE(data_registro,'') AS data_registro" if "data_registro" in cols_desp else "'' AS data_registro"
                        df_despesas = pd.read_sql_query(
                            f"""
                            SELECT
                                {id_desp_expr},
                                {descricao_desp_expr},
                                {valor_desp_expr},
                                {categoria_desp_expr},
                                {observacao_desp_expr},
                                {data_desp_expr}
                            FROM despesas
                            WHERE codigo_programacao=?
                            ORDER BY data_registro DESC
                            """,
                            conn,
                            params=(prog,),
                        )
                    else:
                        df_despesas = pd.DataFrame(columns=["id", "descricao", "valor", "categoria", "observacao", "data_registro"])

                    if _table_exists("recebimentos"):
                        cols_rec = {str(r[1]).lower() for r in (cur.execute("PRAGMA table_info(recebimentos)").fetchall() or [])}
                        cod_cliente_rec_expr = "COALESCE(cod_cliente,'') AS cod_cliente" if "cod_cliente" in cols_rec else "'' AS cod_cliente"
                        nome_cliente_rec_expr = "COALESCE(nome_cliente,'') AS nome_cliente" if "nome_cliente" in cols_rec else "'' AS nome_cliente"
                        valor_rec_expr = "COALESCE(valor,0) AS valor" if "valor" in cols_rec else "0 AS valor"
                        forma_pag_rec_expr = "COALESCE(forma_pagamento,'') AS forma_pagamento" if "forma_pagamento" in cols_rec else "'' AS forma_pagamento"
                        observacao_rec_expr = "COALESCE(observacao,'') AS observacao" if "observacao" in cols_rec else "'' AS observacao"
                        num_nf_rec_expr = "COALESCE(num_nf,'') AS num_nf" if "num_nf" in cols_rec else "'' AS num_nf"
                        data_rec_expr = "COALESCE(data_registro,'') AS data_registro" if "data_registro" in cols_rec else "'' AS data_registro"
                        df_receb = pd.read_sql_query(
                            f"""
                            SELECT
                                {cod_cliente_rec_expr},
                                {nome_cliente_rec_expr},
                                {valor_rec_expr},
                                {forma_pag_rec_expr},
                                {observacao_rec_expr},
                                {num_nf_rec_expr},
                                {data_rec_expr}
                            FROM recebimentos
                            WHERE codigo_programacao=?
                            ORDER BY data_registro DESC
                            """,
                            conn,
                            params=(prog,),
                        )
                    else:
                        df_receb = pd.DataFrame(columns=[
                            "cod_cliente", "nome_cliente", "valor", "forma_pagamento", "observacao", "num_nf", "data_registro"
                        ])

            if "data_criacao" in df_prog.columns:
                df_prog["data_criacao"] = normalize_datetime_column(df_prog["data_criacao"])
            if "data_registro" in df_despesas.columns:
                df_despesas["data_registro"] = normalize_datetime_column(df_despesas["data_registro"])
            if "data_registro" in df_receb.columns:
                df_receb["data_registro"] = normalize_datetime_column(df_receb["data_registro"])

            cedulas_data = []
            for ced in [200, 100, 50, 20, 10, 5, 2]:
                qtd = safe_int(self.ced_entries[ced].get(), 0)
                cedulas_data.append({"CEDULA": ced, "QUANTIDADE": qtd, "TOTAL": qtd * ced})
            df_cedulas = pd.DataFrame(cedulas_data)

            total_desp = float(df_despesas["valor"].sum()) if not df_despesas.empty else 0.0
            total_receb = float(df_receb["valor"].sum()) if not df_receb.empty else 0.0
            total_ced = float(df_cedulas["TOTAL"].sum()) if not df_cedulas.empty else 0.0

            adiant = self._parse_money_local(self.ent_adiantamento.get())
            total_entradas = total_receb + adiant
            total_saidas = total_desp + total_ced
            resultado_liquido = total_entradas - total_saidas

            df_resumo = pd.DataFrame([
                ["PROGRAMAÇÃO", prog],
                ["TOTAL RECEBIMENTOS", total_receb],
                ["TOTAL ADIANTAMENTO", adiant],
                ["TOTAL ENTRADAS", total_entradas],
                ["TOTAL DESPESAS", total_desp],
                ["TOTAL CÉDULAS", total_ced],
                ["TOTAL SADAS", total_saidas],
                ["RESULTADO LQUIDO", resultado_liquido],
                ["DATA EXPORTAÇÃO", datetime.now().strftime("%d/%m/%Y %H:%M")]
            ], columns=["ITEM", "VALOR"])

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df_resumo.to_excel(writer, sheet_name="RESUMO", index=False)
                df_prog.to_excel(writer, sheet_name="PROGRAMAÇÃO", index=False)

                if not df_despesas.empty:
                    df_despesas.to_excel(writer, sheet_name="DESPESAS", index=False)
                else:
                    pd.DataFrame([["SEM DESPESAS REGISTRADAS"]]).to_excel(writer, sheet_name="DESPESAS", index=False, header=False)

                if not df_receb.empty:
                    df_receb.to_excel(writer, sheet_name="RECEBIMENTOS", index=False)
                else:
                    pd.DataFrame([["SEM RECEBIMENTOS REGISTRADOS"]]).to_excel(writer, sheet_name="RECEBIMENTOS", index=False, header=False)

                df_cedulas.to_excel(writer, sheet_name="CEDULAS", index=False)

            try:
                self.logger.info(f"Exportação Excel concluÃÂÂda: {path}")
            except Exception:
                logging.debug("Falha ignorada")

            messagebox.showinfo(
                "OK",
                "Exportação concluÃÂÂda!\n\n"
                f"Arquivo: {os.path.basename(path)}\n"
                f"Despesas: {len(df_despesas)}\n"
                f"Recebimentos: {len(df_receb)}\n"
                f"Resultado Liquido: {self._fmt_money(resultado_liquido) if hasattr(self, '_fmt_money') else resultado_liquido}"
            )

        except Exception as e:
            try:
                self.logger.error(f"Erro ao exportar Excel: {str(e)}")
            except Exception:
                logging.debug("Falha ignorada")
            messagebox.showerror("ERRO", f"Erro ao exportar Excel: {str(e)}")

# ==========================
# ===== FIM DA PARTE 7C (ATUALIZADA) =====
# ==========================

# =========================
# ===== FIM DA PARTE 7 =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 8 (ATUALIZADA) =====
# ==========================

class EscalaPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Escala")

        self.body.grid_rowconfigure(2, weight=1)
        self.body.grid_columnconfigure(0, weight=1)

        filtros = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        filtros.grid(row=0, column=0, sticky="ew")
        for c in range(7):
            filtros.grid_columnconfigure(c, weight=0)
        filtros.grid_columnconfigure(6, weight=1)

        ttk.Label(filtros, text="Periodo", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.var_periodo = tk.StringVar(value="30")
        self.cb_periodo = ttk.Combobox(
            filtros,
            textvariable=self.var_periodo,
            values=["7", "15", "30", "60", "90", "180", "TODAS"],
            state="readonly",
            width=10,
        )
        self.cb_periodo.grid(row=1, column=0, sticky="w", padx=(0, 10))

        ttk.Label(filtros, text="Status", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.var_status = tk.StringVar(value="ATIVAS")
        self.cb_status = ttk.Combobox(
            filtros,
            textvariable=self.var_status,
            values=["ATIVAS", "TODOS", "ATIVA", "EM_ROTA", "CARREGADA", "FINALIZADA", "CANCELADA"],
            state="readonly",
            width=14,
        )
        self.cb_status.grid(row=1, column=1, sticky="w", padx=(0, 10))

        ttk.Button(filtros, text="\U0001F504 ATUALIZAR", style="Ghost.TButton", command=self.refresh_data).grid(
            row=1, column=2, sticky="w"
        )
        ttk.Button(
            filtros,
            text="\U0001F4D1 GERAR PDF",
            style="Primary.TButton",
            command=self.gerar_pdf_escala,
        ).grid(row=1, column=3, sticky="w", padx=(8, 0))

        resumo = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        resumo.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        resumo.grid_columnconfigure(0, weight=1)
        ttk.Label(resumo, text="Resumo de distribuicao", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        kpi_row = ttk.Frame(resumo, style="Card.TFrame")
        kpi_row.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        for c in range(4):
            kpi_row.grid_columnconfigure(c, weight=1)

        kpi_1 = ttk.LabelFrame(kpi_row, text=" Rotas no Periodo ", padding=8)
        kpi_1.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.lbl_esc_kpi_rotas = ttk.Label(kpi_1, text="0", font=("Segoe UI", 13, "bold"), foreground="#1D4ED8")
        self.lbl_esc_kpi_rotas.grid(row=0, column=0, sticky="w")

        kpi_2 = ttk.LabelFrame(kpi_row, text=" Motoristas ", padding=8)
        kpi_2.grid(row=0, column=1, sticky="ew", padx=6)
        self.lbl_esc_kpi_mot = ttk.Label(kpi_2, text="0", font=("Segoe UI", 13, "bold"), foreground="#0F766E")
        self.lbl_esc_kpi_mot.grid(row=0, column=0, sticky="w")

        kpi_3 = ttk.LabelFrame(kpi_row, text=" KM Medio/Motorista ", padding=8)
        kpi_3.grid(row=0, column=2, sticky="ew", padx=6)
        self.lbl_esc_kpi_km = ttk.Label(kpi_3, text="0,0", font=("Segoe UI", 13, "bold"), foreground="#B45309")
        self.lbl_esc_kpi_km.grid(row=0, column=0, sticky="w")

        kpi_4 = ttk.LabelFrame(kpi_row, text=" Horas Medias/Motorista ", padding=8)
        kpi_4.grid(row=0, column=3, sticky="ew", padx=(6, 0))
        self.lbl_esc_kpi_horas = ttk.Label(kpi_4, text="0,00", font=("Segoe UI", 13, "bold"), foreground="#7C3AED")
        self.lbl_esc_kpi_horas.grid(row=0, column=0, sticky="w")

        txt_row = ttk.Frame(resumo, style="Card.TFrame")
        txt_row.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        txt_row.grid_columnconfigure(0, weight=1)
        txt_row.grid_columnconfigure(1, weight=1)

        box_resumo = ttk.Frame(txt_row, style="CardInset.TFrame", padding=(10, 8))
        box_resumo.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        box_resumo.grid_columnconfigure(0, weight=1)
        box_resumo.grid_columnconfigure(1, weight=0)
        self.box_resumo = box_resumo

        ttk.Label(box_resumo, text="Resumo Operacional", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.btn_toggle_resumo_escala = ttk.Button(
            box_resumo,
            text="Expandir",
            style="Toggle.TButton",
            command=self._toggle_resumo_operacional,
        )
        self.btn_toggle_resumo_escala.grid(row=0, column=1, sticky="e")

        self.resumo_content = ttk.Frame(box_resumo, style="CardInset.TFrame")
        self.resumo_content.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.resumo_content.grid_columnconfigure(0, weight=1)

        self.lbl_resumo = ttk.Label(
            self.resumo_content,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=480,
        )
        self.lbl_resumo.grid(row=0, column=0, sticky="nw")

        box_reco = ttk.Frame(txt_row, style="CardInset.TFrame", padding=(10, 8))
        box_reco.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        box_reco.grid_columnconfigure(0, weight=1)
        box_reco.grid_columnconfigure(1, weight=0)
        self.box_reco = box_reco

        ttk.Label(box_reco, text="Recomendacoes", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.btn_toggle_reco_escala = ttk.Button(
            box_reco,
            text="Expandir",
            style="Toggle.TButton",
            command=self._toggle_recomendacoes_escala,
        )
        self.btn_toggle_reco_escala.grid(row=0, column=1, sticky="e")

        self.reco_content = ttk.Frame(box_reco, style="CardInset.TFrame")
        self.reco_content.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.reco_content.grid_columnconfigure(0, weight=1)

        self.lbl_recomendacoes = ttk.Label(
            self.reco_content,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=480,
            foreground="#1E3A8A",
        )
        self.lbl_recomendacoes.grid(row=0, column=0, sticky="nw")
        self._escala_resumo_collapsed = True
        self._escala_reco_collapsed = True
        self.resumo_content.grid_remove()
        self.reco_content.grid_remove()

        chart_wrap = ttk.Frame(resumo, style="Card.TFrame")
        chart_wrap.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        chart_wrap.grid_columnconfigure(0, weight=1)
        self.lbl_esc_chart_title = ttk.Label(
            chart_wrap,
            text="Carga de Horas por Motorista (top 8)",
            style="CardLabel.TLabel",
        )
        self.lbl_esc_chart_title.grid(row=0, column=0, sticky="w")
        self.cv_escala = tk.Canvas(
            chart_wrap,
            height=230,
            bg="white",
            highlightthickness=1,
            highlightbackground="#E5E7EB",
        )
        self.cv_escala.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.cv_escala.bind("<Configure>", lambda _e: self._draw_escala_chart())
        self.lbl_esc_chart_meta = ttk.Label(
            chart_wrap,
            text="-",
            style="CardLabel.TLabel",
            foreground="#4B5563",
        )
        self.lbl_esc_chart_meta.grid(row=2, column=0, sticky="w", pady=(4, 0))
        self._esc_chart_data = []
        resumo.bind("<Configure>", self._on_escala_resize)

        tabs = ttk.Notebook(self.body)
        tabs.grid(row=2, column=0, sticky="nsew")

        tab_m = ttk.Frame(tabs, style="Content.TFrame")
        tab_m.grid_columnconfigure(0, weight=1)
        tab_m.grid_rowconfigure(0, weight=1)
        tabs.add(tab_m, text="Motoristas")

        tab_a = ttk.Frame(tabs, style="Content.TFrame")
        tab_a.grid_columnconfigure(0, weight=1)
        tab_a.grid_rowconfigure(0, weight=1)
        tabs.add(tab_a, text="Ajudantes")

        cols = ("NOME", "ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "LOCAL", "KM_RODADO", "HORAS_TRAB")
        self.tree_m = ttk.Treeview(tab_m, columns=cols, show="headings")
        self.tree_m.grid(row=0, column=0, sticky="nsew")
        vsb_m = ttk.Scrollbar(tab_m, orient="vertical", command=self.tree_m.yview)
        self.tree_m.configure(yscrollcommand=vsb_m.set)
        vsb_m.grid(row=0, column=1, sticky="ns")

        cols_a = ("NOME", "ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "KM_RODADO")
        self.tree_a = ttk.Treeview(tab_a, columns=cols_a, show="headings")
        self.tree_a.grid(row=0, column=0, sticky="nsew")
        vsb_a = ttk.Scrollbar(tab_a, orient="vertical", command=self.tree_a.yview)
        self.tree_a.configure(yscrollcommand=vsb_a.set)
        vsb_a.grid(row=0, column=1, sticky="ns")

        for c in cols:
            self.tree_m.heading(c, text=c)
            if c == "NOME":
                w = 220
            elif c == "LOCAL":
                w = 120
            elif c == "KM_RODADO":
                w = 110
            elif c == "HORAS_TRAB":
                w = 105
            else:
                w = 95
            self.tree_m.column(c, anchor="w" if c in ("NOME", "LOCAL") else "center", width=w)
        enable_treeview_sorting(
            self.tree_m,
            numeric_cols={"ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "KM_RODADO", "HORAS_TRAB"},
        )

        cols_a = ("NOME", "ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "KM_RODADO", "HORAS_TRAB")
        self.tree_a.configure(columns=cols_a)
        for c in cols_a:
            self.tree_a.heading(c, text=c)
            if c == "NOME":
                w = 280
            elif c == "KM_RODADO":
                w = 120
            elif c == "HORAS_TRAB":
                w = 110
            else:
                w = 115
            self.tree_a.column(c, anchor="w" if c == "NOME" else "center", width=w)
        enable_treeview_sorting(
            self.tree_a,
            numeric_cols={"ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "KM_RODADO", "HORAS_TRAB"},
        )

        # Indicadores visuais de carga (sobrecarga)
        for tree in (self.tree_m, self.tree_a):
            tree.tag_configure(
                "carga_sobrecarga",
                foreground="#7F1D1D",
                background="#FEE2E2",
            )
            tree.tag_configure(
                "carga_alerta",
                foreground="#7C2D12",
                background="#FFEDD5",
            )
            tree.tag_configure(
                "carga_equilibrada",
                foreground="#14532D",
                background="#ECFDF3",
            )

        self.cb_periodo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())
        self.cb_status.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())
        self.tree_m.bind("<Double-1>", lambda _e: self._mostrar_detalhe_carga(self.tree_m, "MOTORISTA"))
        self.tree_a.bind("<Double-1>", lambda _e: self._mostrar_detalhe_carga(self.tree_a, "AJUDANTE"))

    def on_show(self):
        self.refresh_data()

    def _set_kpi_escala(self, rotas: int, motoristas: int, km_medio: float, horas_medias: float):
        try:
            self.lbl_esc_kpi_rotas.config(text=str(max(safe_int(rotas, 0), 0)))
            self.lbl_esc_kpi_mot.config(text=str(max(safe_int(motoristas, 0), 0)))
            self.lbl_esc_kpi_km.config(text=f"{safe_float(km_medio, 0.0):.1f}".replace(".", ","))
            self.lbl_esc_kpi_horas.config(text=f"{safe_float(horas_medias, 0.0):.2f}".replace(".", ","))
        except Exception:
            logging.debug("Falha ignorada")

    def _toggle_resumo_operacional(self):
        self._escala_resumo_collapsed = not bool(getattr(self, "_escala_resumo_collapsed", True))
        if self._escala_resumo_collapsed:
            self.resumo_content.grid_remove()
            self.btn_toggle_resumo_escala.configure(text="Expandir")
        else:
            self.resumo_content.grid()
            self.btn_toggle_resumo_escala.configure(text="Ocultar")

    def _toggle_recomendacoes_escala(self):
        self._escala_reco_collapsed = not bool(getattr(self, "_escala_reco_collapsed", True))
        if self._escala_reco_collapsed:
            self.reco_content.grid_remove()
            self.btn_toggle_reco_escala.configure(text="Expandir")
        else:
            self.reco_content.grid()
            self.btn_toggle_reco_escala.configure(text="Ocultar")

    def _on_escala_resize(self, event=None):
        try:
            w = max(int(getattr(event, "width", 0) or 0), 700)
            half_wrap = max(320, (w - 70) // 2)
            self.lbl_resumo.configure(wraplength=half_wrap)
            self.lbl_recomendacoes.configure(wraplength=half_wrap)
            self._draw_escala_chart()
        except Exception:
            logging.debug("Falha ignorada")

    def _draw_escala_chart(self):
        cv = getattr(self, "cv_escala", None)
        if not cv:
            return
        cv.delete("all")
        data = list(getattr(self, "_esc_chart_data", []) or [])
        if not data:
            cv.create_text(12, 12, anchor="nw", text="Sem dados para o periodo selecionado.", fill="#6B7280", font=("Segoe UI", 9))
            return

        w = max(cv.winfo_width(), 360)
        h = max(cv.winfo_height(), 160)
        left = 10
        right = w - 10
        top = 10
        bottom = h - 10
        count = max(min(len(data), 8), 1)
        usable_h = max(bottom - top, 80)
        # Ajusta altura de linha automaticamente para não cortar barras.
        line_h = max(14, min(22, int((usable_h - (count - 1) * 2) / count)))
        row_gap = 2
        max_horas = max((safe_float(item.get("horas", 0.0), 0.0) for item in data), default=0.0)
        max_horas = max(max_horas, 1.0)
        name_w = 145
        val_w = 52
        bar_x0 = left + name_w
        bar_x1 = right - val_w
        usable_w = max(bar_x1 - bar_x0, 48)

        for idx, item in enumerate(data[:8]):
            y = top + idx * (line_h + row_gap)
            if y + line_h > bottom:
                break
            nome = str(item.get("nome", "-"))[:22]
            horas = safe_float(item.get("horas", 0.0), 0.0)
            dias = safe_int(item.get("dias", 0), 0)
            ratio = min(max(horas / max_horas, 0.0), 1.0)
            fill_w = int(usable_w * ratio)
            color = "#7C3AED" if ratio >= 0.8 else ("#2563EB" if ratio >= 0.5 else "#16A34A")
            y_mid = y + (line_h // 2)

            cv.create_text(left, y_mid, anchor="w", text=nome, fill="#111827", font=("Segoe UI", 8, "bold"))
            cv.create_rectangle(bar_x0, y + 2, bar_x1, y + line_h - 2, fill="#F3F4F6", outline="#E5E7EB")
            cv.create_rectangle(bar_x0, y + 2, bar_x0 + fill_w, y + line_h - 2, fill=color, outline=color)
            cv.create_text(bar_x1 - 6, y_mid, anchor="e", text=f"{dias}d", fill="#374151", font=("Segoe UI", 8))
            cv.create_text(right - 2, y_mid, anchor="e", text=f"{horas:.1f}h", fill="#111827", font=("Segoe UI", 8))

    def _mostrar_detalhe_carga(self, tree: ttk.Treeview, perfil: str):
        try:
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], "values") or ()
            if not vals:
                return
            nome = str(vals[0] if len(vals) > 0 else "-")
            rotas = safe_int(vals[1] if len(vals) > 1 else 0, 0)
            em_rota = safe_int(vals[2] if len(vals) > 2 else 0, 0)
            ativas = safe_int(vals[3] if len(vals) > 3 else 0, 0)
            finalizadas = safe_int(vals[4] if len(vals) > 4 else 0, 0)
            canceladas = safe_int(vals[5] if len(vals) > 5 else 0, 0)
            if perfil == "MOTORISTA":
                local = str(vals[6] if len(vals) > 6 else "-")
                km = safe_float(vals[7] if len(vals) > 7 else 0.0, 0.0)
                horas = safe_float(vals[8] if len(vals) > 8 else 0.0, 0.0)
                detalhe = (
                    f"{perfil}: {nome}\n\n"
                    f"Rotas: {rotas}\nEm rota: {em_rota}\nAtivas: {ativas}\n"
                    f"Finalizadas: {finalizadas}\nCanceladas: {canceladas}\n"
                    f"Local predominante: {local or '-'}\nKM rodado: {km:.1f}\nHoras trabalhadas: {horas:.2f}"
                )
            else:
                km = safe_float(vals[6] if len(vals) > 6 else 0.0, 0.0)
                horas = safe_float(vals[7] if len(vals) > 7 else 0.0, 0.0)
                detalhe = (
                    f"{perfil}: {nome}\n\n"
                    f"Rotas: {rotas}\nEm rota: {em_rota}\nAtivas: {ativas}\n"
                    f"Finalizadas: {finalizadas}\nCanceladas: {canceladas}\n"
                    f"KM rodado: {km:.1f}\nHoras trabalhadas: {horas:.2f}"
                )
            messagebox.showinfo("Detalhe de Carga", detalhe)
        except Exception:
            logging.debug("Falha ignorada")

    def _status_normalizado(self, v) -> str:
        return upper(str(v or "").strip())

    def _status_match_filter(self, status: str, filtro: str) -> bool:
        st = self._status_normalizado(status)
        ff = self._status_normalizado(filtro)
        if ff == "TODOS":
            return True
        if ff == "ATIVAS":
            return self._is_ativa(st)
        if ff in {"FINALIZADA", "FINALIZADO"}:
            return self._is_finalizada(st)
        if ff in {"CANCELADA", "CANCELADO"}:
            return self._is_cancelada(st)
        if ff in {"EM_ROTA", "EM ROTA"}:
            return self._is_em_rota(st)
        return st == ff

    def _is_em_rota(self, s: str) -> bool:
        return s in ("EM_ROTA", "EM ROTA", "INICIADA")

    def _is_ativa(self, s: str) -> bool:
        return s in ("ATIVA", "EM_ROTA", "EM ROTA", "INICIADA", "CARREGADA")

    def _is_finalizada(self, s: str) -> bool:
        return s in ("FINALIZADA", "FINALIZADO")

    def _is_cancelada(self, s: str) -> bool:
        return s in ("CANCELADA", "CANCELADO")

    def _parse_data_programacao(self, raw: str):
        txt = str(raw or "").strip()
        if not txt:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%d/%m/%y %H:%M:%S", "%d/%m/%y"):
            try:
                return datetime.strptime(txt[:19], fmt) if "H" in fmt else datetime.strptime(txt[:10], fmt)
            except Exception:
                pass
        try:
            return datetime.fromisoformat(txt.replace(" ", "T"))
        except Exception:
            return None

    def _split_ajudantes(self, equipe_raw: str):
        raw = str(equipe_raw or "").strip()
        if not raw:
            return []
        nomes = resolve_equipe_nomes(raw) or raw
        parts = [p.strip() for p in re.split(r"[|,;/]+", nomes) if p.strip()]
        out = []
        seen = set()
        for p in parts:
            up = upper(p)
            if up in ("NAN", "NONE", "-", "SEM EQUIPE"):
                continue
            if up not in seen:
                seen.add(up)
                out.append(up)
        return out

    def _parse_data_hora(self, data_raw: str, hora_raw: str):
        data_txt = str(data_raw or "").strip()
        hora_txt = str(hora_raw or "").strip()
        if not data_txt:
            return None
        dt_data = self._parse_data_programacao(data_txt)
        if dt_data is None:
            return None
        hh = 0
        mm = 0
        ss = 0
        if hora_txt:
            nt = normalize_time(hora_txt)
            if nt:
                try:
                    p = (nt + ":00:00").split(":")
                    hh = safe_int(p[0], 0)
                    mm = safe_int(p[1], 0)
                    ss = safe_int(p[2], 0)
                except Exception:
                    hh = mm = ss = 0
        return dt_data.replace(hour=hh, minute=mm, second=ss, microsecond=0)

    def _calc_horas_trabalhadas(self, data_saida: str, hora_saida: str, data_chegada: str, hora_chegada: str) -> float:
        dt_saida = self._parse_data_hora(data_saida, hora_saida)
        dt_chegada = self._parse_data_hora(data_chegada, hora_chegada)
        if not dt_saida or not dt_chegada:
            return 0.0
        diff = (dt_chegada - dt_saida).total_seconds() / 3600.0
        if diff <= 0:
            return 0.0
        # Evita valores absurdos por dado corrompido.
        return min(round(diff, 2), 72.0)

    def _tag_por_carga(
        self,
        rotas: int,
        media_rotas: float,
        km_rodado: float = 0.0,
        media_km: float = 0.0,
        horas_trab: float = 0.0,
        media_horas: float = 0.0,
    ) -> str:
        idx_rotas = (float(rotas) / float(media_rotas)) if media_rotas and media_rotas > 0 else 0.0
        idx_km = (float(km_rodado) / float(media_km)) if media_km and media_km > 0 else 0.0
        idx_horas = (float(horas_trab) / float(media_horas)) if media_horas and media_horas > 0 else 0.0
        idx = max(idx_rotas, idx_km, idx_horas)
        if idx > 1.25:
            return "carga_sobrecarga"
        if idx > 1.05:
            return "carga_alerta"
        return "carga_equilibrada"

    def _recomendacoes_distribuicao(self, qtd_rotas: int, mot_rows, aj_rows):
        if qtd_rotas <= 0:
            return "Recomendações: sem dados no filtro para sugerir distribuição."

        now = datetime.now()
        recs = []
        local_alvo = "-"
        if mot_rows:
            local_counts = {}
            for _n, d in mot_rows:
                for loc, qtd in (d.get("local_counts") or {}).items():
                    local_counts[loc] = int(local_counts.get(loc, 0)) + int(qtd or 0)
            if local_counts:
                local_alvo = sorted(local_counts.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]

        if mot_rows:
            media_rotas = max(qtd_rotas / float(len(mot_rows)), 1.0)
            media_kg = max(sum(float(d.get("kg", 0.0) or 0.0) for _n, d in mot_rows) / float(len(mot_rows)), 1.0)
            media_km = max(sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in mot_rows) / float(len(mot_rows)), 1.0)
            media_horas = max(sum(float(d.get("horas_trab", 0.0) or 0.0) for _n, d in mot_rows) / float(len(mot_rows)), 1.0)

            def _score_motorista(d):
                rotas = float(d.get("rotas", 0) or 0)
                em_rota = float(d.get("em_rota", 0) or 0)
                kg = float(d.get("kg", 0.0) or 0.0)
                km = float(d.get("km_rodado", 0.0) or 0.0)
                horas = float(d.get("horas_trab", 0.0) or 0.0)
                score = (
                    (rotas / media_rotas)
                    + (em_rota * 0.85)
                    + (kg / media_kg) * 0.65
                    + (km / media_km) * 0.55
                    + (horas / media_horas) * 0.70
                )

                last_dt = d.get("last_dt")
                if isinstance(last_dt, datetime):
                    horas = max((now - last_dt).total_seconds() / 3600.0, 0.0)
                    if horas < 12:
                        score += 0.80
                    elif horas < 24:
                        score += 0.35

                if local_alvo != "-":
                    total_loc = sum(int(v or 0) for v in (d.get("local_counts") or {}).values())
                    if total_loc > 0:
                        aderencia = float((d.get("local_counts") or {}).get(local_alvo, 0) or 0) / float(total_loc)
                        score -= aderencia * 0.40
                return score

            mot_sorted = sorted(
                mot_rows,
                key=lambda kv: (
                    _score_motorista(kv[1]),
                    kv[0],
                ),
            )
            prox_mot = mot_sorted[0][0]
            recs.append(f"Próxima rota sugerida (motorista): {prox_mot}")

            media_mot = qtd_rotas / float(len(mot_rows))
            sob_mot = [
                n for n, d in mot_rows
                if (
                    int(d.get("rotas", 0)) > media_mot * 1.25
                    or float(d.get("km_rodado", 0.0) or 0.0) > media_km * 1.25
                    or float(d.get("horas_trab", 0.0) or 0.0) > media_horas * 1.20
                )
            ]
            if sob_mot:
                recs.append("Evitar novas rotas para motoristas sobrecarregados: " + ", ".join(sob_mot[:4]))

        if aj_rows:
            media_aj = max(qtd_rotas / float(len(aj_rows)), 1.0)
            media_km_aj = max(sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in aj_rows) / float(len(aj_rows)), 1.0)
            media_horas_aj = max(sum(float(d.get("horas_trab", 0.0) or 0.0) for _n, d in aj_rows) / float(len(aj_rows)), 1.0)

            def _score_ajudante(d):
                rotas = float(d.get("rotas", 0) or 0)
                em_rota = float(d.get("em_rota", 0) or 0)
                km = float(d.get("km_rodado", 0.0) or 0.0)
                horas = float(d.get("horas_trab", 0.0) or 0.0)
                score = (rotas / media_aj) + (em_rota * 0.90) + (km / media_km_aj) * 0.55 + (horas / media_horas_aj) * 0.65
                last_dt = d.get("last_dt")
                if isinstance(last_dt, datetime):
                    horas = max((now - last_dt).total_seconds() / 3600.0, 0.0)
                    if horas < 12:
                        score += 0.75
                    elif horas < 24:
                        score += 0.30
                return score

            aj_sorted = sorted(
                aj_rows,
                key=lambda kv: (
                    _score_ajudante(kv[1]),
                    kv[0],
                ),
            )
            prox_aj = [n for n, _d in aj_sorted[:2]]
            if prox_aj:
                recs.append("Ajudantes sugeridos para próxima equipe: " + " / ".join(prox_aj))

            media_aj = qtd_rotas / float(len(aj_rows))
            sob_aj = [
                n for n, d in aj_rows
                if (
                    int(d.get("rotas", 0)) > media_aj * 1.25
                    or float(d.get("km_rodado", 0.0) or 0.0) > media_km_aj * 1.25
                    or float(d.get("horas_trab", 0.0) or 0.0) > media_horas_aj * 1.20
                )
            ]
            if sob_aj:
                recs.append("Evitar escalar ajudantes sobrecarregados: " + ", ".join(sob_aj[:4]))

        if not recs:
            return "Recomendações: distribuição está equilibrada no filtro atual."
        return f"Recomendacoes (local-alvo: {local_alvo}):\n- " + "\n- ".join(recs)

    def _listar_programacoes_filtradas(self):
        status_filtro = self._status_normalizado(self.var_status.get())
        periodo = upper(self.var_periodo.get().strip())
        cutoff = None
        if periodo != "TODAS":
            try:
                cutoff = datetime.now() - timedelta(days=int(periodo))
            except Exception:
                cutoff = None

        rows = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and can_read_from_api():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/escala/rows?limit=5000",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [r for r in (resp.get("rows") or []) if isinstance(r, dict)]
            except Exception:
                logging.debug("Falha ao listar programacoes da escala via API; usando fallback local.", exc_info=True)

        if not rows:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                if "local_rota" in cols:
                    local_expr = "COALESCE(local_rota,'')"
                elif "local" in cols:
                    local_expr = "COALESCE(local,'')"
                else:
                    local_expr = "''"
                km_rodado_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
                status_op_expr = "COALESCE(status_operacional,'')" if "status_operacional" in cols else "''"
                finalizada_app_expr = "COALESCE(finalizada_no_app,0)" if "finalizada_no_app" in cols else "0"
                kg_estimado_expr = "COALESCE(kg_estimado,0)" if "kg_estimado" in cols else "0"
                data_saida_expr = "COALESCE(data_saida,'')" if "data_saida" in cols else "''"
                hora_saida_expr = "COALESCE(hora_saida,'')" if "hora_saida" in cols else "''"
                data_chegada_expr = "COALESCE(data_chegada,'')" if "data_chegada" in cols else "''"
                hora_chegada_expr = "COALESCE(hora_chegada,'')" if "hora_chegada" in cols else "''"
                if "data_saida" in cols and "data_criacao" in cols and "data" in cols:
                    data_ref_expr = "COALESCE(data_saida,data_criacao,data,'')"
                elif "data_saida" in cols and "data_criacao" in cols:
                    data_ref_expr = "COALESCE(data_saida,data_criacao,'')"
                elif "data_saida" in cols and "data" in cols:
                    data_ref_expr = "COALESCE(data_saida,data,'')"
                elif "data_criacao" in cols and "data" in cols:
                    data_ref_expr = "COALESCE(data_criacao,data,'')"
                elif "data_saida" in cols:
                    data_ref_expr = "COALESCE(data_saida,'')"
                elif "data_criacao" in cols:
                    data_ref_expr = "COALESCE(data_criacao,'')"
                elif "data" in cols:
                    data_ref_expr = "COALESCE(data,'')"
                else:
                    data_ref_expr = "''"
                cur.execute(
                    f"""
                    SELECT codigo_programacao,
                           {data_ref_expr} AS data_ref,
                           COALESCE(motorista,'') AS motorista,
                           COALESCE(equipe,'') AS equipe,
                           COALESCE(status,'') AS status,
                           {status_op_expr} AS status_operacional,
                           {finalizada_app_expr} AS finalizada_no_app,
                           {kg_estimado_expr} AS kg_estimado,
                           {data_saida_expr} AS data_saida,
                           {hora_saida_expr} AS hora_saida,
                           {data_chegada_expr} AS data_chegada,
                           {hora_chegada_expr} AS hora_chegada,
                           {local_expr} AS local_rota,
                           {km_rodado_expr} AS km_rodado
                      FROM programacoes
                     ORDER BY id DESC
                    """
                )
                rows = cur.fetchall() or []

        out = []
        for r in rows:
            st_local = self._status_normalizado(r["status"] if hasattr(r, "keys") else r[4])
            st_op = self._status_normalizado(r["status_operacional"] if hasattr(r, "keys") else r[5])
            fin_app = safe_int(r["finalizada_no_app"] if hasattr(r, "keys") else r[6], 0)
            st = st_op or st_local
            if not st and fin_app == 1:
                st = "FINALIZADA"
            if not self._status_match_filter(st, status_filtro):
                continue

            if cutoff is not None:
                dt = self._parse_data_programacao(r["data_ref"] if hasattr(r, "keys") else r[1])
                if dt is not None and dt < cutoff:
                    continue

            out.append(r)
        return out

    def refresh_data(self):
        try:
            rows = self._listar_programacoes_filtradas()
            qtd_rotas = len(rows)

            mot = {}
            aju = {}

            for r in rows:
                motorista = upper(str((r["motorista"] if hasattr(r, "keys") else r[2]) or "").strip()) or "SEM MOTORISTA"
                st_local = self._status_normalizado(r["status"] if hasattr(r, "keys") else r[4])
                st_op = self._status_normalizado(r["status_operacional"] if hasattr(r, "keys") else r[5])
                fin_app = safe_int(r["finalizada_no_app"] if hasattr(r, "keys") else r[6], 0)
                st = st_op or st_local
                if not st and fin_app == 1:
                    st = "FINALIZADA"
                kg = safe_float(r["kg_estimado"] if hasattr(r, "keys") else r[7], 0.0)
                data_saida = (r["data_saida"] if hasattr(r, "keys") else r[8]) or ""
                hora_saida = (r["hora_saida"] if hasattr(r, "keys") else r[9]) or ""
                data_chegada = (r["data_chegada"] if hasattr(r, "keys") else r[10]) or ""
                hora_chegada = (r["hora_chegada"] if hasattr(r, "keys") else r[11]) or ""
                horas_trab = self._calc_horas_trabalhadas(data_saida, hora_saida, data_chegada, hora_chegada)
                dt_prog = self._parse_data_programacao(r["data_ref"] if hasattr(r, "keys") else r[1])
                local_rota = upper(str((r["local_rota"] if hasattr(r, "keys") else r[12]) or "").strip())
                km_rodado = safe_float(r["km_rodado"] if hasattr(r, "keys") else r[13], 0.0)

                m = mot.setdefault(
                    motorista,
                    {
                        "rotas": 0,
                        "em_rota": 0,
                        "ativas": 0,
                        "finalizadas": 0,
                        "canceladas": 0,
                        "kg": 0.0,
                        "km_rodado": 0.0,
                        "horas_trab": 0.0,
                        "local_ref": "-",
                        "last_dt": None,
                        "local_counts": {},
                        "dias_set": set(),
                    },
                )
                m["rotas"] += 1
                m["kg"] += kg
                m["km_rodado"] += km_rodado
                m["horas_trab"] += horas_trab
                if isinstance(dt_prog, datetime):
                    if m.get("last_dt") is None or dt_prog > m["last_dt"]:
                        m["last_dt"] = dt_prog
                    try:
                        m["dias_set"].add(dt_prog.date().isoformat())
                    except Exception:
                        logging.debug("Falha ignorada")
                dt_saida_ref = self._parse_data_hora(data_saida, hora_saida)
                dt_cheg_ref = self._parse_data_hora(data_chegada, hora_chegada)
                if isinstance(dt_saida_ref, datetime):
                    try:
                        m["dias_set"].add(dt_saida_ref.date().isoformat())
                    except Exception:
                        logging.debug("Falha ignorada")
                if isinstance(dt_cheg_ref, datetime):
                    try:
                        m["dias_set"].add(dt_cheg_ref.date().isoformat())
                    except Exception:
                        logging.debug("Falha ignorada")
                if local_rota:
                    lc = m["local_counts"]
                    lc[local_rota] = int(lc.get(local_rota, 0)) + 1
                    m["local_ref"] = sorted(lc.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
                if self._is_em_rota(st):
                    m["em_rota"] += 1
                if self._is_ativa(st):
                    m["ativas"] += 1
                if self._is_finalizada(st):
                    m["finalizadas"] += 1
                if self._is_cancelada(st):
                    m["canceladas"] += 1

                equipe = r["equipe"] if hasattr(r, "keys") else r[3]
                ajudantes = self._split_ajudantes(equipe)
                for nm in ajudantes:
                    a = aju.setdefault(
                        nm,
                        {
                            "rotas": 0,
                            "em_rota": 0,
                            "ativas": 0,
                            "finalizadas": 0,
                            "canceladas": 0,
                            "km_rodado": 0.0,
                            "horas_trab": 0.0,
                            "last_dt": None,
                        },
                    )
                    a["rotas"] += 1
                    a["km_rodado"] += km_rodado
                    a["horas_trab"] += horas_trab
                    if isinstance(dt_prog, datetime):
                        if a.get("last_dt") is None or dt_prog > a["last_dt"]:
                            a["last_dt"] = dt_prog
                    if self._is_em_rota(st):
                        a["em_rota"] += 1
                    if self._is_ativa(st):
                        a["ativas"] += 1
                    if self._is_finalizada(st):
                        a["finalizadas"] += 1
                    if self._is_cancelada(st):
                        a["canceladas"] += 1

            for t in (self.tree_m, self.tree_a):
                try:
                    t.delete(*t.get_children())
                except Exception:
                    logging.debug("Falha ignorada")

            mot_rows = sorted(mot.items(), key=lambda kv: (-kv[1]["rotas"], kv[0]))
            media_mot = (qtd_rotas / float(len(mot_rows))) if mot_rows else 0.0
            media_km_mot = (
                sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in mot_rows) / float(len(mot_rows))
            ) if mot_rows else 0.0
            media_horas_mot = (
                sum(float(d.get("horas_trab", 0.0) or 0.0) for _n, d in mot_rows) / float(len(mot_rows))
            ) if mot_rows else 0.0
            for nome, d in mot_rows:
                tag = self._tag_por_carga(
                    int(d["rotas"]),
                    media_mot,
                    float(d.get("km_rodado", 0.0) or 0.0),
                    media_km_mot,
                    float(d.get("horas_trab", 0.0) or 0.0),
                    media_horas_mot,
                )
                tree_insert_aligned(
                    self.tree_m,
                    "",
                    "end",
                    (
                        nome,
                        d["rotas"],
                        d["em_rota"],
                        d["ativas"],
                        d["finalizadas"],
                        d["canceladas"],
                        d.get("local_ref", "-"),
                        round(float(d.get("km_rodado", 0.0) or 0.0), 1),
                        round(float(d.get("horas_trab", 0.0) or 0.0), 2),
                    ),
                    tags=(tag,),
                )

            aj_rows = sorted(aju.items(), key=lambda kv: (-kv[1]["rotas"], kv[0]))
            media_aju = (qtd_rotas / float(len(aj_rows))) if aj_rows else 0.0
            media_km_aju = (
                sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in aj_rows) / float(len(aj_rows))
            ) if aj_rows else 0.0
            media_horas_aju = (
                sum(float(d.get("horas_trab", 0.0) or 0.0) for _n, d in aj_rows) / float(len(aj_rows))
            ) if aj_rows else 0.0
            for nome, d in aj_rows:
                tag = self._tag_por_carga(
                    int(d["rotas"]),
                    media_aju,
                    float(d.get("km_rodado", 0.0) or 0.0),
                    media_km_aju,
                    float(d.get("horas_trab", 0.0) or 0.0),
                    media_horas_aju,
                )
                tree_insert_aligned(
                    self.tree_a,
                    "",
                    "end",
                    (
                        nome,
                        d["rotas"],
                        d["em_rota"],
                        d["ativas"],
                        d["finalizadas"],
                        d["canceladas"],
                        round(float(d.get("km_rodado", 0.0) or 0.0), 1),
                        round(float(d.get("horas_trab", 0.0) or 0.0), 2),
                    ),
                    tags=(tag,),
                )

            qtd_motoristas = len(mot_rows)
            qtd_ajudantes = len(aj_rows)
            total_km = sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in mot_rows)
            total_horas = sum(float(d.get("horas_trab", 0.0) or 0.0) for _n, d in mot_rows)
            total_dias_trab = sum(len((d.get("dias_set") or set())) for _n, d in mot_rows)
            if qtd_motoristas > 0:
                media = qtd_rotas / float(qtd_motoristas)
                media_km = total_km / float(qtd_motoristas)
                media_horas = total_horas / float(qtd_motoristas)
                media_dias = total_dias_trab / float(qtd_motoristas)
                mais_sobrecarregado = mot_rows[0][0] if mot_rows else "-"
                carga_max = float(mot_rows[0][1].get("rotas", 0) or 0) if mot_rows else 0.0
                nivel = "EQUILIBRADA"
                cor_nivel = "#14532D"
                if media > 0 and carga_max > media * 1.25:
                    nivel = "SOBRECARGA"
                    cor_nivel = "#B91C1C"
                elif media > 0 and carga_max > media * 1.05:
                    nivel = "ALERTA"
                    cor_nivel = "#9A3412"
                msg = (
                    f"Nivel da escala: {nivel} | Maior carga atual: {mais_sobrecarregado}\n"
                    f"Rotas no filtro: {qtd_rotas} | Motoristas: {qtd_motoristas} | Ajudantes: {qtd_ajudantes}\n"
                    f"Média por motorista: {media:.2f} | KM total: {total_km:.1f} | KM médio/motorista: {media_km:.1f}\n"
                    f"Horas totais: {total_horas:.2f} | Horas médias/motorista: {media_horas:.2f}\n"
                    f"Dias trabalhados (motoristas): {total_dias_trab} | Média de dias/motorista: {media_dias:.2f}\n"
                    "Legenda visual: verde=equilibrado | laranja=alerta | vermelho=sobrecarga"
                )
                self.lbl_resumo.config(foreground=cor_nivel)
                self._set_kpi_escala(qtd_rotas, qtd_motoristas, media_km, media_horas)
            else:
                msg = f"Rotas no filtro: {qtd_rotas} | Sem motoristas no periodo/filtro selecionado."
                self.lbl_resumo.config(foreground="#6B7280")
                self._set_kpi_escala(qtd_rotas, 0, 0.0, 0.0)

            self.lbl_resumo.config(text=msg)
            chart_rows = sorted(
                [
                    {
                        "nome": n,
                        "horas": float(d.get("horas_trab", 0.0) or 0.0),
                        "dias": len((d.get("dias_set") or set())),
                    }
                    for n, d in mot_rows
                ],
                key=lambda x: (-x["horas"], -x["dias"], x["nome"]),
            )[:8]
            self._esc_chart_data = chart_rows
            self._draw_escala_chart()
            if qtd_motoristas > 0:
                dias_media_txt = f"{(total_dias_trab / float(qtd_motoristas)):.2f}".replace(".", ",")
            else:
                dias_media_txt = "0,00"
            self.lbl_esc_chart_meta.config(
                text=f"Leitura de produtividade: horas acumuladas + dias ativos por motorista | Média de dias: {dias_media_txt}"
            )
            self.lbl_recomendacoes.config(
                text=self._recomendacoes_distribuicao(qtd_rotas, mot_rows, aj_rows) + "\n\nDuplo clique em uma linha para abrir detalhe da carga."
            )
            self.set_status(
                f"STATUS: Escala atualizada ({self.var_periodo.get()} dias / {self.var_status.get()}) - "
                f"Rotas: {qtd_rotas}."
            )
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao atualizar escala:\n\n{e}")

    def _escala_tree_rows(self, tree):
        rows = []
        try:
            for iid in tree.get_children():
                vals = tuple(tree.item(iid, "values") or ())
                if vals:
                    rows.append(vals)
        except Exception:
            logging.debug("Falha ignorada")
        return rows

    def _wrap_pdf_lines(self, text, max_width, font_name="Helvetica", font_size=9):
        from reportlab.pdfbase.pdfmetrics import stringWidth

        lines = []
        for raw in str(text or "").splitlines():
            words = raw.split()
            if not words:
                lines.append("")
                continue
            current = ""
            for word in words:
                candidate = word if not current else f"{current} {word}"
                if stringWidth(candidate, font_name, font_size) <= max_width:
                    current = candidate
                else:
                    if current:
                        lines.append(current)
                    current = word
            if current:
                lines.append(current)
        return lines

    def _fit_pdf_lines(self, text, max_width, max_height, line_height=4.6, font_name="Helvetica", font_size=9):
        from reportlab.lib.units import mm

        lines = self._wrap_pdf_lines(text, max_width, font_name=font_name, font_size=font_size)
        max_count = max(1, int(max_height // (line_height * mm)))
        if len(lines) > max_count:
            lines = lines[:max_count]
            if lines:
                last = lines[-1].rstrip()
                lines[-1] = (last[: max(0, len(last) - 3)] + "...") if len(last) > 3 else "..."
        return lines

    def _draw_pdf_wrapped_text(self, c, text, x, y, max_width, line_height=4.6, max_lines=None, font_name="Helvetica", font_size=9):
        from reportlab.lib.units import mm
        lines = self._wrap_pdf_lines(text, max_width, font_name=font_name, font_size=font_size)

        if max_lines is not None and len(lines) > max_lines:
            lines = lines[:max_lines]
            if lines:
                last = lines[-1]
                lines[-1] = (last[: max(0, len(last) - 3)] + "...") if len(last) > 3 else "..."

        c.setFont(font_name, font_size)
        for line in lines:
            c.drawString(x, y, line)
            y -= line_height * mm
        return y

    def _compact_pdf_text(self, text, max_lines=6):
        cleaned = []
        for raw in str(text or "").splitlines():
            line = " ".join(str(raw).strip().split())
            if not line:
                continue
            if line.startswith("- "):
                line = "• " + line[2:]
            cleaned.append(line)
        if len(cleaned) > max_lines:
            cleaned = cleaned[:max_lines]
            cleaned[-1] = cleaned[-1][: max(0, len(cleaned[-1]) - 3)] + "..."
        return "\n".join(cleaned)

    def _draw_escala_pdf_table(self, c, title, headers, rows, x, y, col_widths, page_width, page_height, bottom_margin):
        from reportlab.lib.units import mm

        row_h = 5.6 * mm

        def _new_page():
            c.showPage()
            c.setFont("Helvetica-Bold", 16)
            c.drawString(14 * mm, page_height - 16 * mm, "RELATORIO DE ESCALA")
            c.setFont("Helvetica", 8)
            c.drawRightString(page_width - 14 * mm, page_height - 16 * mm, datetime.now().strftime("%d/%m/%Y %H:%M"))
            return page_height - 25 * mm

        c.setFont("Helvetica-Bold", 10)
        c.drawString(x, y, title)
        y -= 5.5 * mm

        if y < bottom_margin + 25 * mm:
            y = _new_page()

        def _draw_header(current_y):
            c.setFont("Helvetica-Bold", 7)
            cx = x
            for head, w in zip(headers, col_widths):
                c.rect(cx, current_y - row_h, w, row_h, stroke=1, fill=0)
                c.drawString(cx + 2, current_y - row_h + 4, str(head))
                cx += w
            return current_y - row_h

        y = _draw_header(y)
        c.setFont("Helvetica", 7)

        for row in rows:
            if y < bottom_margin + 16 * mm:
                y = _new_page()
                y = _draw_header(y)
                c.setFont("Helvetica", 7)
            cx = x
            for value, w in zip(row, col_widths):
                c.rect(cx, y - row_h, w, row_h, stroke=1, fill=0)
                txt = str(value or "")
                limit = max(8, int((w / mm) // 2.7))
                if len(txt) > limit:
                    txt = txt[: max(0, limit - 3)] + "..."
                c.drawString(cx + 2, y - row_h + 4, txt)
                cx += w
            y -= row_h

        return y - 3 * mm

    def gerar_pdf_escala(self):
        if not require_reportlab():
            return

        path = filedialog.asksaveasfilename(
            title="Salvar PDF da Escala",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"Escala_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        )
        if not path:
            return

        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm

            mot_rows = self._escala_tree_rows(self.tree_m)
            aj_rows = self._escala_tree_rows(self.tree_a)
            periodo = str(self.var_periodo.get() or "-")
            status = str(self.var_status.get() or "-")

            page_w, page_h = A4
            left = 14 * mm
            right = 14 * mm
            top = page_h - 16 * mm
            bottom = 14 * mm
            usable_w = page_w - left - right

            c = canvas.Canvas(path, pagesize=A4)
            c.setTitle("Relatorio de Escala")
            y = top

            c.setFont("Helvetica-Bold", 16)
            c.drawString(left, y, "RELATORIO DE ESCALA")
            y -= 8 * mm

            c.setFont("Helvetica", 9)
            c.drawString(left, y, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            c.drawRightString(page_w - right, y, f"Periodo: {periodo} dia(s)  |  Status: {status}")
            y -= 8 * mm

            c.setFont("Helvetica-Bold", 10)
            c.drawString(left, y, "Resumo Executivo")
            y -= 6 * mm

            card_gap = 4 * mm
            card_w = (usable_w - card_gap) / 2.0
            card_h = 12 * mm
            kpi_cards = [
                ("Rotas no periodo", self.lbl_esc_kpi_rotas.cget("text")),
                ("Motoristas", self.lbl_esc_kpi_mot.cget("text")),
                ("KM medio/motorista", self.lbl_esc_kpi_km.cget("text")),
                ("Horas medias/motorista", self.lbl_esc_kpi_horas.cget("text")),
            ]
            for idx, (label, value) in enumerate(kpi_cards):
                row = idx // 2
                col = idx % 2
                bx = left + col * (card_w + card_gap)
                by = y - row * (card_h + 3 * mm)
                c.rect(bx, by - card_h, card_w, card_h, stroke=1, fill=0)
                c.setFont("Helvetica", 8)
                c.drawString(bx + 3, by - 4 * mm, label)
                c.setFont("Helvetica-Bold", 11)
                c.drawString(bx + 3, by - 9 * mm, str(value))
            y -= (2 * card_h) + (3 * mm) + 2 * mm

            half_w = (usable_w - 6 * mm) / 2.0
            box_h = 40 * mm
            text_top_gap = 11 * mm
            text_bottom_gap = 4 * mm
            text_line_h = 4.3
            text_font = 8
            usable_box_h = box_h - text_top_gap - text_bottom_gap

            resumo_txt = self._compact_pdf_text(self.lbl_resumo.cget("text"), max_lines=8)
            reco_txt = self._compact_pdf_text(self.lbl_recomendacoes.cget("text"), max_lines=8)
            resumo_lines = self._fit_pdf_lines(
                resumo_txt,
                half_w - 8,
                usable_box_h,
                line_height=text_line_h,
                font_size=text_font,
            )
            reco_lines = self._fit_pdf_lines(
                reco_txt,
                half_w - 8,
                usable_box_h,
                line_height=text_line_h,
                font_size=text_font,
            )
            resumo_txt = "\n".join(resumo_lines)
            reco_txt = "\n".join(reco_lines)

            c.rect(left, y - box_h, half_w, box_h, stroke=1, fill=0)
            c.rect(left + half_w + 6 * mm, y - box_h, half_w, box_h, stroke=1, fill=0)

            c.setFont("Helvetica-Bold", 10)
            c.drawString(left + 3, y - 5 * mm, "Resumo Operacional")
            c.drawString(left + half_w + 6 * mm + 3, y - 5 * mm, "Recomendacoes")

            self._draw_pdf_wrapped_text(
                c,
                resumo_txt,
                left + 3,
                y - text_top_gap,
                half_w - 8,
                line_height=text_line_h,
                font_size=text_font,
            )
            self._draw_pdf_wrapped_text(
                c,
                reco_txt,
                left + half_w + 6 * mm + 3,
                y - text_top_gap,
                half_w - 8,
                line_height=text_line_h,
                font_size=text_font,
            )
            y = y - box_h - 4 * mm

            headers_m = ["Motorista", "Rotas", "Em rota", "Ativas", "Final.", "Canc.", "Local", "KM", "Horas"]
            rows_m = [tuple(r[:9]) for r in mot_rows[:12]]
            widths_m = [40 * mm, 13 * mm, 13 * mm, 13 * mm, 13 * mm, 13 * mm, 22 * mm, 17 * mm, 17 * mm]
            y = self._draw_escala_pdf_table(c, "Distribuicao por Motorista", headers_m, rows_m, left, y, widths_m, page_w, page_h, bottom)
            if len(mot_rows) > 12:
                c.setFont("Helvetica-Oblique", 8)
                c.drawString(left, y, f"* Exibindo os 12 primeiros motoristas de {len(mot_rows)} registros.")
                y -= 5 * mm

            headers_a = ["Ajudante", "Rotas", "Em rota", "Ativas", "Final.", "Canc.", "KM", "Horas"]
            rows_a = [tuple(r[:8]) for r in aj_rows[:12]]
            widths_a = [54 * mm, 15 * mm, 15 * mm, 15 * mm, 15 * mm, 15 * mm, 20 * mm, 20 * mm]
            y = self._draw_escala_pdf_table(c, "Distribuicao por Ajudante", headers_a, rows_a, left, y, widths_a, page_w, page_h, bottom)
            if len(aj_rows) > 12:
                c.setFont("Helvetica-Oblique", 8)
                c.drawString(left, y, f"* Exibindo os 12 primeiros ajudantes de {len(aj_rows)} registros.")
                y -= 5 * mm

            c.setFont("Helvetica", 8)
            c.drawRightString(page_w - right, bottom - 2 * mm, "RotaHub Desktop - Escala")
            c.save()

            messagebox.showinfo("OK", f"PDF da escala gerado com sucesso!\n\nArquivo:\n{os.path.basename(path)}")
            self.set_status(f"STATUS: PDF da escala gerado: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao gerar PDF da escala:\n\n{e}")


class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE_DESKTOP)
        apply_window_icon(self)
        self.geometry("440x280")  # +20px p/ não apertar botões
        self.resizable(False, False)

        # ---- Aplica estilo
        try:
            apply_style(self)
        except Exception:
            logging.debug("Falha ignorada")

        # ---- Centraliza
        self.update_idletasks()
        w, h = 520, 320
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

        # ---- Fallback de estilo (evita botão "sumir" por estilo invisÃÂÂvel)
        self._ensure_button_styles()

        # ---- Controle de tentativas (segurança leve sem quebrar fluxo)
        self._attempts = 0
        self._blocked_until = 0  # epoch seconds
        self._max_attempts = 5
        self._block_seconds = 15

        card = ttk.Frame(self, style="Card.TFrame", padding=18)
        card.pack(fill="both", expand=True, padx=12, pady=12)

        # âÅâ€œââ‚¬¦ garante espaço do grid dentro do card
        card.grid_columnconfigure(0, weight=1)

        ttk.Label(
            card, text="ACESSO AO SISTEMA", style="CardTitle.TLabel"
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        frm = ttk.Frame(card, style="Card.TFrame")
        frm.grid(row=1, column=0, sticky="ew")
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="Nome:", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", pady=5)
        self.ent_codigo = ttk.Entry(frm, style="Field.TEntry")
        self.ent_codigo.grid(row=0, column=1, sticky="ew", pady=5)
        bind_entry_smart(self.ent_codigo, "text")

        ttk.Label(frm, text="Senha:", style="CardLabel.TLabel").grid(row=1, column=0, sticky="w", pady=5)
        self.ent_senha = ttk.Entry(frm, style="Field.TEntry", show="*")
        self.ent_senha.grid(row=1, column=1, sticky="ew", pady=5)
        bind_entry_smart(self.ent_senha, "text")

        # Mostrar/ocultar senha
        self._show_pass = tk.BooleanVar(value=False)

        def _toggle_pass():
            self.ent_senha.config(show="" if self._show_pass.get() else "*")

        chk = ttk.Checkbutton(frm, text="Mostrar senha", variable=self._show_pass, command=_toggle_pass)
        chk.grid(row=2, column=1, sticky="w", pady=(2, 0))

        self.lbl_status = ttk.Label(card, text="", background="white", foreground="#555")
        self.lbl_status.grid(row=2, column=0, sticky="w", pady=(6, 0))

        btns = ttk.Frame(card, style="Card.TFrame")
        btns.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)

        # âÅâ€œââ‚¬¦ padding e sticky garantem que apareçam e tenham tamanho
        self.btn_entrar = ttk.Button(btns, text="🔐 ENTRAR", style="Primary.TButton", command=self.try_login)
        self.btn_entrar.grid(row=0, column=0, sticky="ew", padx=(0, 6), ipady=8)

        self.btn_sair = ttk.Button(btns, text="SAIR", style="Danger.TButton", command=self._request_close)
        self.btn_sair.grid(row=0, column=1, sticky="ew", padx=(6, 0), ipady=8)

        self.ent_codigo.focus_set()
        self.bind("<Return>", lambda e: self.try_login())
        self.bind("<Escape>", lambda e: self._request_close())
        self.protocol("WM_DELETE_WINDOW", self._request_close)

        self.user = None  # será preenchido quando logar

    def _request_close(self):
        try:
            ok = messagebox.askyesno("Confirmar saida", "Deseja realmente fechar o sistema?")
        except Exception:
            ok = True
        if ok:
            self.destroy()

    def _ensure_button_styles(self):
        """
        Se o tema não criou Primary.TButton / Danger.TButton, cria um fallback
        para evitar o efeito de "botão invisÃÂÂvel".
        Não altera seu tema se já existir.
        """
        try:
            st = ttk.Style(self)
            existing = set(st.theme_names())  # só pra não dar erro em alguns temas
            _ = existing  # quiet

            # Alguns temas não suportam lookup; então testamos com lookup e fallback
            def _style_exists(name: str) -> bool:
                try:
                    # se retornar algo sem dar exception, consideramos "existe"
                    st.lookup(name, "foreground")
                    return True
                except Exception:
                    return False

            # Fallback simples (sem brigar com o apply_style)
            if not _style_exists("Primary.TButton"):
                st.configure("Primary.TButton", padding=10)

            if not _style_exists("Danger.TButton"):
                st.configure("Danger.TButton", padding=10)

        except Exception:
            logging.debug("Falha ignorada")

    def _is_blocked(self) -> bool:
        try:
            import time
            return time.time() < self._blocked_until
        except Exception:
            return False

    def _block_now(self):
        try:
            import time
            self._blocked_until = time.time() + self._block_seconds
        except Exception:
            self._blocked_until = 0

    def try_login(self):
        if self._is_blocked():
            self.lbl_status.config(text=f"Muitas tentativas. Aguarde {self._block_seconds}s e tente novamente.")
            return

        codigo = upper(self.ent_codigo.get().strip())
        senha = self.ent_senha.get().strip()

        if not codigo or not senha:
            messagebox.showwarning("ATENÇÃO", "Informe código e senha.")
            return

        # âÅâ€œââ‚¬¦ login real: usa sua função existente (mantém o sistema)
        try:
            auth_res = autenticar_usuario(codigo, senha)
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao autenticar: {str(e)}")
            return

        user = auth_res.get("data") if isinstance(auth_res, dict) and bool(auth_res.get("ok", False)) else None

        if not user:
            self._attempts += 1
            if self._attempts >= self._max_attempts:
                self._block_now()
                self._attempts = 0
                self.lbl_status.config(text="Muitas tentativas. Login bloqueado temporariamente.")
            else:
                self.lbl_status.config(
                    text=f"Código ou senha inválidos. Tentativas: {self._attempts}/{self._max_attempts}"
                )
            return

        self.user = user
        self.destroy()

# Compartilha helpers de sincronizacao desktop com Recebimentos.
for _shared_method in (
    "_api_bundle_prestacao",
    "_sync_programacao_from_api_desktop",
    "_apply_api_programacao",
    "_apply_api_clientes",
    "_normalize_media_kg_ave",
):
    if not hasattr(RecebimentosPage, _shared_method) and hasattr(DespesasPage, _shared_method):
        setattr(RecebimentosPage, _shared_method, getattr(DespesasPage, _shared_method))

def abrir_login():
    win = LoginWindow()
    win.mainloop()

    if not getattr(win, "user", None):
        return  # usuário saiu / cancelou

    app = App(user=win.user)

    if not hasattr(app, "close_db_connections"):
        def _close_db_connections_best_effort():
            return
        app.close_db_connections = _close_db_connections_best_effort

    app.mainloop()


if __name__ == "__main__":
    db_init()  # garante migrações
    abrir_login()

# ==========================
# ===== FIM DA PARTE 10 (ATUALIZADA) =====
# ==========================

