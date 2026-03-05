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
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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
from contextlib import contextmanager
import random
import string
import base64
import hashlib
import hmac
import secrets
import ctypes
import time

# =========================================================
# CONSTANTES E CONFIGURAÇÕES
# =========================================================
APP_W, APP_H = 1360, 780
DB_NAME = "rota_granja"
APP_TITLE_DESKTOP = "ROTAHUB DESKTOP"

IS_FROZEN = getattr(sys, "frozen", False)
APP_DIR = os.path.dirname(os.path.abspath(__file__))
RESOURCE_DIR = getattr(sys, "_MEIPASS", APP_DIR)

# Em dev (python main.py), mantem o banco local da pasta do projeto.
# Em app instalado (PyInstaller), usa pasta persistente do usuario para nao perder dados em updates.
if IS_FROZEN:
    USER_DATA_DIR = os.path.join(
        os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
        "RotaHubDesktop",
    )
    os.makedirs(USER_DATA_DIR, exist_ok=True)
    DEFAULT_DB_PATH = os.path.join(USER_DATA_DIR, "rota_granja.db")
else:
    DEFAULT_DB_PATH = os.path.join(APP_DIR, "rota_granja.db")

DB_PATH = os.environ.get("ROTA_DB", DEFAULT_DB_PATH)
os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)

# Primeira execucao do app instalado: copia o banco seed embutido, se existir.
if IS_FROZEN and not os.path.exists(DB_PATH):
    seed_candidates = [
        os.path.join(RESOURCE_DIR, "rota_granja.db"),
        os.path.join(os.path.dirname(sys.executable), "rota_granja.db"),
    ]
    for seed in seed_candidates:
        try:
            if os.path.exists(seed):
                shutil.copy2(seed, DB_PATH)
                break
        except OSError:
            logging.debug("Falha ignorada")

_default_api_url = "https://rotahub-api.onrender.com" if IS_FROZEN else "http://127.0.0.1:8000"
API_BASE_URL = os.environ.get("ROTA_SERVER_URL", _default_api_url).strip().rstrip("/")
if not API_BASE_URL:
    API_BASE_URL = "http://127.0.0.1:8000"
try:
    # Render free pode levar mais de 15s no cold-start.
    API_SYNC_TIMEOUT = float(os.environ.get("ROTA_SYNC_TIMEOUT", "60"))
except Exception:
    API_SYNC_TIMEOUT = 60.0


def is_desktop_api_sync_enabled() -> bool:
    """Controla sincronizacao automatica Desktop <-> API central.
    Padrao: ligado, para manter vinculo continuo entre estacoes e app mobile.
    Para testes locais isolados, desative explicitamente com ROTA_DESKTOP_SYNC_API=0.
    """
    raw = str(os.environ.get("ROTA_DESKTOP_SYNC_API", "1") or "").strip().lower()
    return raw in {"1", "true", "yes", "y", "sim", "on"}


_API_BINDING_CACHE = {"ok": False, "checked_at": 0.0, "error": ""}


def ensure_system_api_binding(context: str = "Operacao", parent=None, force_probe: bool = False) -> bool:
    """Garante vinculo obrigatorio Desktop <-> API para operacoes criticas."""
    if not is_desktop_api_sync_enabled():
        messagebox.showerror(
            "INTEGRACAO OBRIGATORIA",
            "A integracao Desktop<->Servidor esta desativada (ROTA_DESKTOP_SYNC_API=0).\n\n"
            f"Operacao bloqueada: {context}\n"
            "Ative ROTA_DESKTOP_SYNC_API=1 para continuar.",
            parent=parent,
        )
        return False

    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if not desktop_secret:
        messagebox.showerror(
            "INTEGRACAO OBRIGATORIA",
            "ROTA_SECRET nao configurada.\n\n"
            f"Operacao bloqueada: {context}\n"
            "Defina a chave da estacao para manter o fluxo unico Desktop/Servidor/Dispositivo.",
            parent=parent,
        )
        return False

    now = time.time()
    if (not force_probe) and _API_BINDING_CACHE.get("ok") and (now - float(_API_BINDING_CACHE.get("checked_at") or 0.0) <= 20.0):
        return True

    try:
        _call_api(
            "GET",
            "admin/motoristas/acesso",
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )
        _API_BINDING_CACHE["ok"] = True
        _API_BINDING_CACHE["checked_at"] = now
        _API_BINDING_CACHE["error"] = ""
        return True
    except Exception as exc:
        _API_BINDING_CACHE["ok"] = False
        _API_BINDING_CACHE["checked_at"] = now
        _API_BINDING_CACHE["error"] = str(exc or "")
        messagebox.showerror(
            "INTEGRACAO OBRIGATORIA",
            "Nao foi possivel validar conexao com a API central.\n\n"
            f"Operacao bloqueada: {context}\n"
            f"URL base: {API_BASE_URL}\n"
            f"Detalhe: {str(exc or 'Falha de conectividade')}",
            parent=parent,
        )
        return False


if is_desktop_api_sync_enabled() and not os.environ.get("ROTA_SECRET", "").strip():
    logging.warning(
        "Sincronizacao Desktop<->API esta ativa, mas ROTA_SECRET nao foi definido. "
        "Operacoes protegidas da API podem falhar ate a configuracao da chave."
    )

# Versao/update/suporte (customizavel via variaveis de ambiente no servidor/estacao)
APP_VERSION = "1.1.4"
UPDATE_MANIFEST_URL = os.environ.get("ROTA_UPDATE_MANIFEST_URL", "").strip()
SETUP_DOWNLOAD_URL = os.environ.get("ROTA_SETUP_URL", "").strip()
CHANGELOG_URL = os.environ.get("ROTA_CHANGELOG_URL", "").strip()
SUPPORT_WHATSAPP = os.environ.get("ROTA_SUPPORT_WHATSAPP", "").strip()
SUPPORT_EMAIL = os.environ.get("ROTA_SUPPORT_EMAIL", "").strip()

# Log global em DEBUG (pedido do usuario)
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
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

# =========================================================
# REPARO GLOBAL DE TEXTO (MOJIBAKE)
# =========================================================
_MOJIBAKE_MARKERS = (
    "\u00c3", "\u00c2", "\u00e2", "\u0192", "\u00c6", "\u00a2",
    "\u20ac", "\u2122", "\u0153", "\u017e", "\ufffd",
    "\u0102", "\u0103", "\u0161", "\u2030", "\x89", "\x9a",
)


def _mojibake_score(text: str) -> int:
    if not text:
        return 0
    return sum(text.count(ch) for ch in _MOJIBAKE_MARKERS)


def fix_mojibake_text(value):
    if not isinstance(value, str) or not value:
        return value

    best = value
    best_score = _mojibake_score(best)
    if best_score == 0:
        return value

    for _ in range(4):
        improved = False
        # latin-1/cp1252 cobrem o mojibake classico (ex.: "Programa\u00c3\u00a7\u00c3\u00a3o")
        # cp1250/iso8859_2 cobrem variacoes do tipo "\u0102\u0161" / "\u0102\u2030".
        for enc in ("latin-1", "cp1252", "cp1250", "iso8859_2"):
            try:
                candidate = best.encode(enc).decode("utf-8")
            except Exception:
                continue
            cand_score = _mojibake_score(candidate)
            if cand_score < best_score:
                best, best_score = candidate, cand_score
                improved = True
        if not improved:
            break

    # fallback leve para sequencias restantes comuns
    best = best.replace("\u00e2\u20ac\u201d", "-").replace("\u00e2\u20ac\u201c", "-")
    best = best.replace("\u00e2\u20ac\u02dc", "\u2191").replace("\u00e2\u20ac\u0153", "\u2193")
    return best


def _align_values_to_tree_columns(tree: ttk.Treeview, values):
    try:
        cols = list(tree.cget("columns") or [])
    except Exception:
        cols = []
    ncols = len(cols)

    if isinstance(values, (list, tuple)):
        vals = list(values)
    elif isinstance(values, sqlite3.Row):
        # sqlite3.Row precisa virar lista de valores; caso contrario cai como objeto unico.
        vals = list(values)
    elif values is None:
        vals = []
    else:
        vals = [values]

    if ncols > 0:
        if len(vals) < ncols:
            vals.extend([""] * (ncols - len(vals)))
        elif len(vals) > ncols:
            vals = vals[:ncols]
    return tuple(vals)


def tree_insert_aligned(tree: ttk.Treeview, parent: str, index: str, values, **kwargs):
    aligned = _align_values_to_tree_columns(tree, values)
    # Zebra leve + melhor leitura visual entre linhas
    try:
        if not getattr(tree, "_zebra_ready", False):
            tree.tag_configure("row_even", background="#FFFFFF")
            tree.tag_configure("row_odd", background="#F8FAFC")
            tree._zebra_ready = True
        existing_tags = list(kwargs.get("tags", ()))
        pos = len(tree.get_children(parent or ""))
        zebra_tag = "row_even" if (pos % 2 == 0) else "row_odd"
        if zebra_tag not in existing_tags:
            existing_tags.append(zebra_tag)
        kwargs["tags"] = tuple(existing_tags)
    except Exception:
        logging.debug("Falha ao aplicar estilo zebrado na tabela")
    return tree.insert(parent, index, values=aligned, **kwargs)


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
# GERENCIADOR DE BANCO DE DADOS (CONTEXT MANAGER)
# =========================================================
def _configure_sqlite(conn: sqlite3.Connection) -> sqlite3.Connection:
    """Aplica pragmas de desempenho/concorrencia."""
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.execute("PRAGMA busy_timeout = 5000")
    except Exception:
        logging.debug("Falha ignorada")
    return conn


def _is_mutating_sql(sql: str) -> bool:
    s = str(sql or "").lstrip().upper()
    return s.startswith("INSERT ") or s.startswith("UPDATE ") or s.startswith("DELETE ") or s.startswith("REPLACE ")


def _normalize_sql_scalar(v):
    if v is None:
        return None
    if isinstance(v, (str, int, float, bool)):
        return v
    if isinstance(v, (bytes, bytearray)):
        try:
            return bytes(v).decode("utf-8", errors="replace")
        except Exception:
            return str(v)
    return str(v)


def _normalize_sql_params(params):
    if params is None:
        return []
    if isinstance(params, dict):
        out = {}
        for k, v in params.items():
            out[str(k)] = _normalize_sql_scalar(v)
        return out
    if not isinstance(params, (list, tuple)):
        return [_normalize_sql_scalar(params)]
    return [_normalize_sql_scalar(v) for v in params]


def _is_sql_mirror_enabled() -> bool:
    raw = str(os.environ.get("ROTA_SQL_MIRROR_API", "1") or "").strip().lower()
    return is_desktop_api_sync_enabled() and raw in {"1", "true", "yes", "y", "sim", "on"}


def _push_sql_mutations_to_api(statements):
    if not statements:
        return
    if not _is_sql_mirror_enabled():
        return
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if not desktop_secret:
        raise RuntimeError("ROTA_SECRET nao configurada para espelhamento SQL.")
    _call_api(
        "POST",
        "desktop/sql/mutate",
        payload={"statements": statements},
        extra_headers={"X-Desktop-Secret": desktop_secret},
    )


class SyncedCursor(sqlite3.Cursor):
    def execute(self, sql, parameters=()):
        result = super().execute(sql, parameters)
        try:
            conn = getattr(self, "connection", None)
            if conn is not None and hasattr(conn, "_track_sql_mutation"):
                conn._track_sql_mutation(sql, parameters)
        except Exception:
            logging.debug("Falha ao rastrear mutacao SQL", exc_info=True)
        return result

    def executemany(self, sql, seq_of_parameters):
        seq = list(seq_of_parameters or [])
        result = super().executemany(sql, seq)
        try:
            conn = getattr(self, "connection", None)
            if conn is not None and hasattr(conn, "_track_sql_mutation"):
                for p in seq:
                    conn._track_sql_mutation(sql, p)
        except Exception:
            logging.debug("Falha ao rastrear mutacoes SQL (executemany)", exc_info=True)
        return result


class SyncedConnection(sqlite3.Connection):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._pending_sql_mutations = []
        self._suspend_sql_mirror = False

    def cursor(self, *args, **kwargs):
        kwargs.setdefault("factory", SyncedCursor)
        return super().cursor(*args, **kwargs)

    def _track_sql_mutation(self, sql, params):
        if self._suspend_sql_mirror:
            return
        if not _is_mutating_sql(sql):
            return
        self._pending_sql_mutations.append(
            {
                "sql": str(sql),
                "params": _normalize_sql_params(params),
            }
        )

    def commit(self):
        if self._pending_sql_mutations and not self._suspend_sql_mirror:
            _push_sql_mutations_to_api(self._pending_sql_mutations)
            self._pending_sql_mutations = []
        return super().commit()

    def rollback(self):
        self._pending_sql_mutations = []
        return super().rollback()


@contextmanager
def get_db():
    """Gerenciador de contexto para conexões com o banco"""
    conn = sqlite3.connect(DB_PATH, factory=SyncedConnection)
    _configure_sqlite(conn)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def db_connect():
    """Função compatÃÂÂvel para código existente"""
    conn = sqlite3.connect(DB_PATH, factory=SyncedConnection)
    _configure_sqlite(conn)
    return conn

# =========================================================
# FUNÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢âââââ€š¬Å¡¬¡ÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢âââââ€š¬Å¡¬¢ES UTILITÃÆââ‚¬â„¢Æâââ€š¬ââ€ž¢Ãâââ€š¬Å¡ÂÂÂÂRIAS
# =========================================================
def upper(s):
    return str(s).strip().upper() if s is not None else ""

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


def _format_money_from_digits_global(raw: str) -> str:
    digits = re.sub(r"\D", "", str(raw or ""))
    if not digits:
        return "0,00"
    if len(digits) == 1:
        int_part, cents = "0", "0" + digits
    elif len(digits) == 2:
        int_part, cents = "0", digits
    else:
        int_part, cents = digits[:-2], digits[-2:]
    int_part = int_part.lstrip("0") or "0"
    parts = []
    while len(int_part) > 3:
        parts.insert(0, int_part[-3:])
        int_part = int_part[:-3]
    parts.insert(0, int_part)
    return f"{'.'.join(parts)},{cents}"


def bind_entry_smart(entry: tk.Entry, kind: str = "text", precision: int = 2):
    """
    Comportamento padrao para Entry:
      - limpa/seleciona ao focar;
      - mascara/normaliza conforme tipo (money/int/decimal/date/time/text).
    """
    kind = (kind or "text").strip().lower()

    def _focus_in(_e=None):
        try:
            v = str(entry.get() or "").strip()
            clear_values = {"", "0", "0,0", "0,00", "0.0", "0.00", "R$ 0,00", "R$0,00", "__/__/__", "__/__/____", "__:__", "__:__:__"}
            if v in clear_values:
                entry.delete(0, "end")
            else:
                entry.selection_range(0, "end")
        except Exception:
            logging.debug("Falha ignorada")

    def _sanitize_decimal(txt: str) -> str:
        txt = str(txt or "").replace(".", ",")
        txt = re.sub(r"[^0-9,]", "", txt)
        if txt.count(",") > 1:
            head, tail = txt.split(",", 1)
            tail = tail.replace(",", "")
            txt = f"{head},{tail}"
        return txt

    def _mask_cpf(digits: str) -> str:
        d = re.sub(r"\D", "", digits)[:11]
        if len(d) <= 3:
            return d
        if len(d) <= 6:
            return f"{d[:3]}.{d[3:]}"
        if len(d) <= 9:
            return f"{d[:3]}.{d[3:6]}.{d[6:]}"
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"

    def _mask_phone(digits: str) -> str:
        d = re.sub(r"\D", "", digits)[:11]
        if len(d) <= 2:
            return d
        if len(d) <= 6:
            return f"({d[:2]}) {d[2:]}"
        if len(d) <= 10:
            return f"({d[:2]}) {d[2:6]}-{d[6:]}"
        return f"({d[:2]}) {d[2:7]}-{d[7:]}"

    def _key_release(_e=None):
        try:
            raw = str(entry.get() or "")
            if kind == "money":
                masked = _format_money_from_digits_global(raw)
                entry.delete(0, "end")
                if re.sub(r"\D", "", raw):
                    entry.insert(0, masked)
                entry.icursor("end")
                return
            if kind == "int":
                digits = re.sub(r"\D", "", raw)
                entry.delete(0, "end")
                entry.insert(0, digits)
                entry.icursor("end")
                return
            if kind == "cpf":
                masked = _mask_cpf(raw)
                entry.delete(0, "end")
                entry.insert(0, masked)
                entry.icursor("end")
                return
            if kind == "phone":
                masked = _mask_phone(raw)
                entry.delete(0, "end")
                entry.insert(0, masked)
                entry.icursor("end")
                return
            if kind == "decimal":
                masked = _sanitize_decimal(raw)
                entry.delete(0, "end")
                entry.insert(0, masked)
                entry.icursor("end")
                return
            if kind == "date":
                digits = re.sub(r"\D", "", raw)[:8]
                if len(digits) <= 2:
                    masked = digits
                elif len(digits) <= 4:
                    masked = f"{digits[:2]}/{digits[2:]}"
                else:
                    masked = f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"
                entry.delete(0, "end")
                entry.insert(0, masked)
                entry.icursor("end")
                return
            if kind == "time":
                digits = re.sub(r"\D", "", raw)[:6]
                if len(digits) <= 2:
                    masked = digits
                elif len(digits) <= 4:
                    masked = f"{digits[:2]}:{digits[2:]}"
                else:
                    masked = f"{digits[:2]}:{digits[2:4]}:{digits[4:]}"
                entry.delete(0, "end")
                entry.insert(0, masked)
                entry.icursor("end")
                return
        except Exception:
            logging.debug("Falha ignorada")

    def _focus_out(_e=None):
        try:
            raw = str(entry.get() or "").strip()
            if kind == "money":
                entry.delete(0, "end")
                entry.insert(0, _format_money_from_digits_global(raw))
                return
            if kind == "int":
                if not raw:
                    entry.insert(0, "0")
                return
            if kind == "cpf":
                entry.delete(0, "end")
                entry.insert(0, _mask_cpf(raw))
                return
            if kind == "phone":
                entry.delete(0, "end")
                entry.insert(0, _mask_phone(raw))
                return
            if kind == "decimal":
                if not raw:
                    entry.insert(0, "0")
                    return
                s = raw.replace(".", ",")
                try:
                    num = float(s.replace(",", "."))
                    entry.delete(0, "end")
                    entry.insert(0, f"{num:.{max(0, precision)}f}".replace(".", ","))
                except Exception:
                    pass
                return
            if kind == "date":
                nd = normalize_date(raw)
                if nd is None:
                    entry.delete(0, "end")
                    return
                if nd:
                    entry.delete(0, "end")
                    entry.insert(0, format_date_br_short(nd))
                return
            if kind == "time":
                nt = normalize_time(raw)
                entry.delete(0, "end")
                if nt:
                    entry.insert(0, nt)
                return
        except Exception:
            logging.debug("Falha ignorada")

    entry.bind("<FocusIn>", _focus_in, add="+")
    entry.bind("<KeyRelease>", _key_release, add="+")
    entry.bind("<FocusOut>", _focus_out, add="+")

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

def safe_float(v, default=0.0):
    try:
        if v is None:
            return default
        if isinstance(v, str):
            s = v.replace("R$", "").replace(" ", "").strip()
            if not s:
                return default

            has_dot = "." in s
            has_comma = "," in s

            if has_dot and has_comma:
                # Decide decimal separator by rightmost symbol.
                if s.rfind(",") > s.rfind("."):
                    # 1.234,56
                    s = s.replace(".", "").replace(",", ".")
                else:
                    # 1,234.56
                    s = s.replace(",", "")
            elif has_comma:
                # 1234,56 (pt-BR decimal comma)
                s = s.replace(".", "").replace(",", ".")
            elif has_dot:
                # Dot-only numbers: prefer decimal interpretation (Flutter envia 3.455 / 6499.44).
                # If there are multiple dots, keep only the last as decimal separator.
                if s.count(".") > 1:
                    parts = s.split(".")
                    s = "".join(parts[:-1]) + "." + parts[-1]
            v = s
        return float(v)
    except Exception:
        return default

def safe_money(v, default=0.0):
    """Converte valor monetário para float com 2 casas (Decimal)."""
    try:
        if v is None:
            return default
        s = str(v).replace("R$", "").strip()
        if not s:
            return default
        s = s.replace(".", "").replace(",", ".")
        d = Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return float(d)
    except (InvalidOperation, ValueError):
        return default

def normalize_date(s: str):
    """Normaliza para YYYY-MM-DD. Retorna '' se vazio, None se inválido."""
    s = (s or "").strip()
    if not s:
        return ""
    part = s.split()[0]
    try:
        toks = [t for t in re.split(r"[^0-9]", part) if t != ""]
        if len(toks) == 3:
            if len(toks[0]) == 4:
                y, m, d = int(toks[0]), int(toks[1]), int(toks[2])
            elif len(toks[2]) == 4:
                d, m, y = int(toks[0]), int(toks[1]), int(toks[2])
            elif len(toks[2]) == 2:
                d, m, y2 = int(toks[0]), int(toks[1]), int(toks[2])
                y = (2000 + y2) if y2 <= 69 else (1900 + y2)
            else:
                return None
            if 1 <= m <= 12 and 1 <= d <= 31:
                return f"{y:04d}-{m:02d}-{d:02d}"
            return None
    except Exception:
        return None
    digits = re.sub(r"\D", "", part)
    if len(digits) == 8:
        try:
            # Aceita YYYYMMDD e DDMMYYYY.
            y1, m1, d1 = int(digits[:4]), int(digits[4:6]), int(digits[6:8])
            if 1900 <= y1 <= 2199 and 1 <= m1 <= 12 and 1 <= d1 <= 31:
                return f"{y1:04d}-{m1:02d}-{d1:02d}"
            d2, m2, y2 = int(digits[:2]), int(digits[2:4]), int(digits[4:8])
            if 1 <= m2 <= 12 and 1 <= d2 <= 31 and 1900 <= y2 <= 2199:
                return f"{y2:04d}-{m2:02d}-{d2:02d}"
        except Exception:
            return None
    if len(digits) == 6:
        try:
            d, m, y2 = int(digits[:2]), int(digits[2:4]), int(digits[4:6])
            y = (2000 + y2) if y2 <= 69 else (1900 + y2)
            if 1 <= m <= 12 and 1 <= d <= 31:
                return f"{y:04d}-{m:02d}-{d:02d}"
        except Exception:
            return None
    return None

def normalize_time(s: str):
    """Normaliza para HH:MM:SS. Retorna '' se vazio, None se inválido."""
    s = (s or "").strip()
    if not s:
        return ""
    part = s.split()[0]
    digits = re.sub(r"\D", "", part)
    if not digits:
        return None
    if len(digits) == 3:
        digits = "0" + digits
    if len(digits) == 4:
        digits += "00"
    if len(digits) != 6:
        return None
    try:
        hh = int(digits[:2])
        mm = int(digits[2:4])
        ss = int(digits[4:6])
        if 0 <= hh <= 23 and 0 <= mm <= 59 and 0 <= ss <= 59:
            return f"{hh:02d}:{mm:02d}:{ss:02d}"
    except Exception:
        return None
    return None

def format_date_br_short(value: str) -> str:
    nd = normalize_date(value)
    if nd is None:
        return str(value or "")
    if not nd:
        return ""
    try:
        y, m, d = nd.split("-")
        return f"{int(d):02d}/{int(m):02d}/{(int(y) % 100):02d}"
    except Exception:
        return str(value or "")


class SyncError(Exception):
    """Erro especÃÂÂfico ao tentar sincronizar com a API mobile."""


def _build_api_url(path: str) -> str:
    path = (path or "").strip().lstrip("/")
    if path:
        return f"{API_BASE_URL}/{path}"
    return API_BASE_URL


def _call_api(method: str, path: str, payload=None, token: str = None, extra_headers: dict = None):
    url = _build_api_url(path)
    headers = {"Accept": "application/json"}
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if extra_headers:
        for k, v in (extra_headers or {}).items():
            if k and v is not None:
                headers[str(k)] = str(v)

    req = urllib.request.Request(url, data=data, headers=headers, method=method.upper())
    try:
        with urllib.request.urlopen(req, timeout=API_SYNC_TIMEOUT) as resp:
            body = resp.read()
            if not body:
                return {}
            text = body.decode("utf-8")
            return json.loads(text)
    except urllib.error.HTTPError as exc:
        body = exc.read()
        detail = ""
        if body:
            try:
                payload = json.loads(body.decode("utf-8"))
                if isinstance(payload, dict):
                    detail = payload.get("detail") or payload.get("message") or str(payload)
                else:
                    detail = str(payload)
            except Exception:
                detail = body.decode("utf-8", errors="ignore")
        raise SyncError(f"{exc.code} {exc.reason}: {detail or 'Sem detalhes'}")
    except urllib.error.URLError as exc:
        raise SyncError(f"Falha ao conectar-se a {url}: {exc.reason}")
    except Exception as exc:
        raise SyncError(f"Erro inesperado ao chamar a API de sincronização: {exc}")
    return None

def normalize_date_time_components(data: str, hora: str):
    """Normaliza data/hora e retorna tupla (data, hora) preservando valor original se inválido."""
    nd = normalize_date(data)
    nt = normalize_time(hora)
    out_data = format_date_br_short(nd if nd is not None else (data or ""))
    out_hora = nt if nt is not None else (hora or "")
    return out_data, out_hora

def format_date_time(data: str, hora: str) -> str:
    d, h = normalize_date_time_components(data, hora)
    return f"{d} {h}".strip()

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

def safe_int(v, default=0):
    try:
        if v is None:
            return default
        if isinstance(v, str):
            v = re.sub(r"[^\d\-]", "", v.strip())
        return int(float(v))
    except Exception:
        return default

def fmt_money(v):
    """Formata valor monetário"""
    return f"R$ {safe_float(v,0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def require_pandas() -> bool:
    if pd is None:
        try:
            messagebox.showerror(
                "ERRO",
                "Esta funcionalidade requer o pacote 'pandas'.\n\n"
                "Instale com: pip install pandas"
            )
        except Exception:
            logging.debug("Falha ignorada")
        return False
    return True

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

def excel_engine_for(path: str):
    p = str(path or "").lower()
    if p.endswith(".xls") and not p.endswith(".xlsx"):
        return "xlrd"
    if p.endswith(".xlsx"):
        return "openpyxl"
    return None

def require_excel_support(path: str) -> bool:
    p = str(path or "").lower()
    if p.endswith(".xls") and not p.endswith(".xlsx"):
        return require_xlrd()
    if p.endswith(".xlsx"):
        return require_openpyxl()
    try:
        messagebox.showerror(
            "ERRO",
            "Formato de arquivo não suportado.\n\n"
            "Use .xlsx ou .xls."
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

def now_str():
    """Retorna data/hora atual formatada"""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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

def guess_col(cols, candidates):
    """Tenta adivinhar o nome da coluna no Excel"""
    def _norm(s: str) -> str:
        s = str(s or "").strip().lower()
        repl = {
            "á": "a", "à": "a", "â": "a", "ã": "a",
            "é": "e", "è": "e", "ê": "e",
            "ÃÂ": "i", "ì": "i", "î": "i",
            "ó": "o", "ò": "o", "ô": "o", "õ": "o",
            "ú": "u", "ù": "u", "û": "u",
            "ç": "c",
        }
        for k, v in repl.items():
            s = s.replace(k, v)
        s = re.sub(r"[^a-z0-9]+", " ", s).strip()
        return s

    cols_lower = {_norm(c): c for c in cols}
    for cand in candidates:
        cand_low = _norm(cand)
        for col_low, original in cols_lower.items():
            if cand_low and (cand_low == col_low or cand_low in col_low):
                return original
    return None

def validate_required(value: str, field_name: str, min_len=1, max_len=120):
    v = (value or "").strip()
    if len(v) < min_len:
        return False, f"{field_name} é obrigatório."
    if len(v) > max_len:
        return False, f"{field_name} muito longo (máx {max_len})."
    return True, ""


def validate_codigo(value: str, field_name="Código", min_len=2, max_len=20):
    v = upper((value or "").strip())
    ok, msg = validate_required(v, field_name, min_len=min_len, max_len=max_len)
    if not ok:
        return False, msg
    if not re.match(r"^[A-Z0-9_-]+$", v):
        return False, f"{field_name} deve conter apenas A-Z, 0-9, '_' ou '-'."
    return True, ""


def validate_placa(value: str):
    v = upper((value or "").strip())
    ok, msg = validate_required(v, "Placa", min_len=6, max_len=8)
    if not ok:
        return False, msg
    if not re.match(r"^[A-Z0-9-]+$", v):
        return False, "Placa inválida."
    return True, ""


def validate_money(value: str, field_name="Valor"):
    v = (value or "").strip().replace(".", "").replace(",", ".")
    try:
        f = float(v)
        if f < 0:
            return False, f"{field_name} não pode ser negativo."
        return True, ""
    except Exception:
        return False, f"{field_name} inválido."


# =========================================================
# AUTENTICAÇÃO + ADMIN PADRÃO
# =========================================================

import os
import base64
import hashlib
import hmac

def hash_password_pbkdf2(password: str, *, iterations: int = 200_000) -> str:
    password = str(password or "")
    if password == "":
        raise ValueError("Senha vazia.")
    salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=32)
    return "pbkdf2_sha256${}${}${}".format(
        iterations,
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(dk).decode("ascii"),
    )

def verify_password_pbkdf2(password: str, stored: str) -> bool:
    try:
        password = str(password or "")
        stored = str(stored or "")

        if not stored.startswith("pbkdf2_sha256$"):
            return False

        _, iters_s, salt_b64, hash_b64 = stored.split("$", 3)
        iterations = int(iters_s)

        salt = base64.b64decode(salt_b64.encode("ascii"))
        expected = base64.b64decode(hash_b64.encode("ascii"))

        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=len(expected))
        return hmac.compare_digest(dk, expected)
    except Exception:
        return False


def autenticar_usuario(login: str, senha: str):
    """
    âÅâ€œââ‚¬¦ Login por NOME + SENHA
    âÅâ€œââ‚¬¦ Migração automática:
       - Se senha no DB já for HASH: valida HASH
       - Se senha no DB for PURA: valida pura e, se OK, converte e salva HASH
    Retorna dict do usuário se ok, senão None.
    """
    login = (login or "").strip()
    senha = (senha or "").strip()
    if not login or not senha:
        return None

    with get_db() as conn:
        cur = conn.cursor()

        # Descobre colunas disponÃÂÂveis
        cur.execute("PRAGMA table_info(usuarios)")
        cols = [str(r[1] or "").lower() for r in cur.fetchall()]

        has_permissoes = "permissoes" in cols
        has_cpf = "cpf" in cols
        has_telefone = "telefone" in cols
        has_senha = "senha" in cols

        if not has_senha:
            return None

        # Puxa o usuário pelo nome (case-insensitive)
        select_parts = ["id", "nome", "senha"]
        select_parts.append("permissoes" if has_permissoes else "'' as permissoes")
        select_parts.append("cpf" if has_cpf else "'' as cpf")
        select_parts.append("telefone" if has_telefone else "'' as telefone")

        cur.execute(f"""
            SELECT {", ".join(select_parts)}
            FROM usuarios
            WHERE UPPER(nome)=UPPER(?)
            LIMIT 1
        """, (login,))
        row = cur.fetchone()

        if not row:
            return None

        user_id = row[0]
        nome = row[1] or ""
        senha_db = row[2] or ""
        permissoes = row[3] if len(row) > 3 else ""
        cpf = row[4] if len(row) > 4 else ""
        telefone = row[5] if len(row) > 5 else ""

        # 1) Se já é hash, valida por hash
        if str(senha_db).startswith("pbkdf2_sha256$"):
            if not verify_password_pbkdf2(senha, senha_db):
                return None

        # 2) Senha pura: valida pura e MIGRA para hash automaticamente
        else:
            if str(senha_db) != senha:
                return None

            try:
                novo_hash = hash_password_pbkdf2(senha)
                cur.execute("UPDATE usuarios SET senha=? WHERE id=?", (novo_hash, user_id))
                # conn.commit() é automático no seu contexto get_db() se ele commita ao sair;
                # se não, descomente a linha abaixo:
                # conn.commit()
            except Exception as e:
                # Se falhar a migração, pelo menos deixa logar (já validou a senha pura)
                logging.exception("Falha ao migrar senha do usuario id=%s: %s", user_id, e)

    is_admin = ("ADMIN" in (permissoes or "").upper()) or (nome.strip().upper() == "ADMIN")

    return {
        "id": user_id,
        "nome": nome,
        "permissoes": permissoes,
        "cpf": cpf,
        "telefone": telefone,
        "is_admin": is_admin,
    }

def _notify_admin_seed(password: str):
    """Notifica senha temporária do ADMIN sem gravar em arquivo."""
    msg = (
        "ADMIN criado automaticamente.\n"
        f"Senha temporaria: {password}\n\n"
        "Troque a senha no primeiro acesso."
    )
    try:
        if tk._default_root is not None:
            messagebox.showwarning("ADMIN CRIADO", msg)
            return
    except Exception:
        logging.debug("Falha ignorada")
    print(msg)


def ensure_admin_user():
    """Garante que exista ADMIN (sem sobrescrever senha existente)."""
    with get_db() as conn:
        cur = conn.cursor()

        cur.execute("PRAGMA table_info(usuarios)")
        cols = [str(r[1] or "").lower() for r in cur.fetchall()]
        if "senha" not in cols:
            return

        has_permissoes = "permissoes" in cols

        cur.execute("SELECT id, senha FROM usuarios WHERE UPPER(nome)=?", ("ADMIN",))
        row = cur.fetchone()

        if not row:
            senha_plana = (
                os.environ.get("ROTA_ADMIN_PASS")
                or os.environ.get("ROTA_ADMIN_PASSWORD")
                or ""
            ).strip()
            if not senha_plana:
                senha_plana = secrets.token_urlsafe(8)
            senha_hash = hash_password_pbkdf2(senha_plana)

            if has_permissoes:
                cur.execute(
                    "INSERT INTO usuarios (nome, senha, permissoes) VALUES (?, ?, ?)",
                    ("ADMIN", senha_hash, "ADMIN")
                )
            else:
                cur.execute(
                    "INSERT INTO usuarios (nome, senha) VALUES (?, ?)",
                    ("ADMIN", senha_hash)
                )
            _notify_admin_seed(senha_plana)
        else:
            admin_id = row[0]
            senha_db = row[1] or ""
            # se admin ainda estiver com senha pura, converte para hash mantendo o mesmo valor
            if senha_db and not str(senha_db).startswith("pbkdf2_sha256$"):
                try:
                    novo_hash = hash_password_pbkdf2(str(senha_db))
                    if has_permissoes:
                        cur.execute(
                            "UPDATE usuarios SET senha=?, permissoes=? WHERE id=?",
                            (novo_hash, "ADMIN", admin_id)
                        )
                    else:
                        cur.execute("UPDATE usuarios SET senha=? WHERE id=?", (novo_hash, admin_id))
                except Exception as e:
                    logging.exception("Falha ao migrar senha do ADMIN id=%s: %s", admin_id, e)

# =========================================================
# MIGRAÇÃO DE BANCO DE DADOS
# =========================================================
def table_has_column(cur, table, col):
    """Verifica se coluna existe na tabela"""
    cur.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    return col in cols

def safe_add_column(cur, table, col, coltype):
    """Adiciona coluna apenas se não existir"""
    if not table_has_column(cur, table, col):
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")

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

        try:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_clientes_cod ON clientes(cod_cliente)")
        except Exception as e:
            logging.exception("Falha ao criar indice de clientes (cod_cliente): %s", e)

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
        # Regra FOB/CIF: estimativa pode ser por KG (FOB) ou por CX (CIF)
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
                obs_recebimento TEXT
            )
        """)
        safe_add_column(cur, "programacao_itens_controle", "peso_previsto", "REAL DEFAULT 0")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS programacao_itens_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_programacao TEXT,
                cod_cliente TEXT,
                payload_json TEXT,
                registrado_em TEXT
            )
        """)

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

# =========================================================
# ESTILOS DA INTERFACE (MODERNIZADO)
# =========================================================
def apply_style(root):
    """Aplica estilos Tkinter (tema moderno, sem afetar a lógica do sistema)"""
    style = ttk.Style(root)

    # Em alguns Windows, "clam" funciona melhor e é mais consistente
    try:
        style.theme_use("clam")
    except Exception:
        logging.debug("Falha ignorada")

    # Paleta moderna
    PRIMARY = "#2B2F8F"
    PRIMARY_DARK = "#1F246F"
    BG = "#F4F6FB"
    CARD = "#FFFFFF"
    TEXT = "#1F2937"
    MUTED = "#6B7280"
    BORDER = "#E5E7EB"
    DANGER = "#B42318"
    DANGER_HOVER = "#D92D20"
    WARN = "#F79009"
    WARN_HOVER = "#FDB022"
    GHOST = "#EEF2F7"
    GHOST_HOVER = "#E5EAF3"

    # Tabelas (Treeview): linhas/colunas mais visÃÂÂveis
    try:
        style.configure(
            "Treeview",
            rowheight=26,
            borderwidth=1,
            relief="solid",
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            lightcolor=BORDER,
            darkcolor=BORDER,
            bordercolor=BORDER
        )
        style.map(
            "Treeview",
            background=[("selected", "#DDE6F8")],
            foreground=[("selected", TEXT)]
        )
        style.configure(
            "Treeview.Heading",
            borderwidth=1,
            relief="solid",
            background=BG,
            foreground=TEXT,
            font=("Segoe UI", 8, "bold")
        )
    except Exception:
        logging.debug("Falha ignorada")

    # Base
    style.configure(".", font=("Segoe UI", 10))
    style.configure("Sidebar.TFrame", background=PRIMARY)
    style.configure("Content.TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD, relief="flat", borderwidth=0)

    # Labels
    style.configure("CardTitle.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 15, "bold"))
    style.configure("CardLabel.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 8, "bold"))
    style.configure("SidebarLogo.TLabel", background=PRIMARY, foreground="white", font=("Segoe UI", 16, "bold"))
    style.configure("SidebarSmall.TLabel", background=PRIMARY, foreground="#DDE3FF", font=("Segoe UI", 9))

    # Entradas
    style.configure("Field.TEntry", padding=(10, 8))
    style.map(
        "Field.TEntry",
        bordercolor=[("focus", PRIMARY), ("!focus", BORDER)],
        lightcolor=[("focus", PRIMARY), ("!focus", BORDER)],
        darkcolor=[("focus", PRIMARY), ("!focus", BORDER)],
    )

    # Botões (mantendo seus nomes de style)
    style.configure(
        "Side.TButton",
        background=PRIMARY,
        foreground="white",
        borderwidth=0,
        font=("Segoe UI", 10, "bold"),
        anchor="w",
        padding=(12, 10),
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Side.TButton", background=[("active", PRIMARY_DARK)])

    style.configure(
        "SideActive.TButton",
        background=PRIMARY_DARK,
        foreground="white",
        borderwidth=0,
        font=("Segoe UI", 10, "bold"),
        anchor="w",
        padding=(12, 10),
        focusthickness=0,
        focuscolor="none",
    )

    style.configure(
        "SideSub.TButton",
        background=PRIMARY,
        foreground="#EEF2FF",
        borderwidth=0,
        font=("Segoe UI", 8, "bold"),
        anchor="w",
        padding=(12, 8),
        focusthickness=0,
        focuscolor="none",
    )
    style.map("SideSub.TButton", background=[("active", PRIMARY_DARK)])

    style.configure(
        "SideSubActive.TButton",
        background=PRIMARY_DARK,
        foreground="white",
        borderwidth=0,
        font=("Segoe UI", 8, "bold"),
        anchor="w",
        padding=(12, 8),
        focusthickness=0,
        focuscolor="none",
    )

    style.configure(
        "Primary.TButton",
        background=PRIMARY,
        foreground="white",
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Primary.TButton", background=[("active", PRIMARY_DARK)])

    style.configure(
        "Ghost.TButton",
        background=GHOST,
        foreground=TEXT,
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Ghost.TButton", background=[("active", GHOST_HOVER)])

    style.configure(
        "Warn.TButton",
        background=WARN,
        foreground="black",
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Warn.TButton", background=[("active", WARN_HOVER)])

    style.configure(
        "Danger.TButton",
        background=DANGER,
        foreground="white",
        font=("Segoe UI", 10, "bold"),
        padding=(14, 10),
        borderwidth=0,
        focusthickness=0,
        focuscolor="none",
    )
    style.map("Danger.TButton", background=[("active", DANGER_HOVER)])

    # Separators
    style.configure("TSeparator", background=BORDER)

    # Treeview modernizado
    style.configure(
        "Treeview",
        background="white",
        foreground=TEXT,
        fieldbackground="white",
        rowheight=28,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        borderwidth=1,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Treeview.Heading",
        background="#F3F4F6",
        foreground=TEXT,
        relief="flat",
        font=("Segoe UI", 8, "bold"),
        padding=(8, 6),
    )
    style.map("Treeview", background=[("selected", "#DDE7FF")], foreground=[("selected", TEXT)])


# =========================================================
# BASE PARA PÃÆââ‚¬â„¢ÂGINAS (mesma lÃÆââ‚¬â„¢³gica, sÃÆââ‚¬â„¢³ ajuste visual/organizaÃÆââ‚¬â„¢§ÃÆââ‚¬â„¢£o)
# =========================================================
class PageBase(ttk.Frame):
    """Classe base para todas as páginas da aplicação"""
    def __init__(self, parent, app, title):
        super().__init__(parent, style="Content.TFrame")
        self.app = app

        # Estrutura
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(self, style="Content.TFrame", padding=(18, 16))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)
        self.header = header

        ttk.Label(
            header,
            text=title,
            font=("Segoe UI", 18, "bold"),
            background="#F4F6FB",
            foreground="#111827"
        ).grid(row=0, column=0, sticky="w")

        self.lbl_status = ttk.Label(
            header,
            text="STATUS: âââââ€š¬Å¡¬ââââ‚¬Å¡¬",
            background="#F4F6FB",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_status.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.header_right = ttk.Frame(header, style="Content.TFrame")
        self.header_right.grid(row=0, column=1, rowspan=2, sticky="e")

        ttk.Separator(self).grid(row=1, column=0, sticky="ew")

        self.body = ttk.Frame(self, style="Content.TFrame", padding=(18, 14))
        self.body.grid(row=2, column=0, sticky="nsew")
        self.body.grid_columnconfigure(0, weight=1)
        self.body.grid_rowconfigure(1, weight=1)

        ttk.Separator(self).grid(row=3, column=0, sticky="ew")

        footer = ttk.Frame(self, style="Content.TFrame", padding=(18, 12))
        footer.grid(row=4, column=0, sticky="ew")
        footer.grid_columnconfigure(0, weight=1)

        self.footer_left = ttk.Frame(footer, style="Content.TFrame")
        self.footer_left.grid(row=0, column=0, sticky="w")

        self.footer_right = ttk.Frame(footer, style="Content.TFrame")
        self.footer_right.grid(row=0, column=1, sticky="e")

    def set_status(self, txt):
        """Atualiza texto de status na página"""
        self.lbl_status.config(text=txt)

    def on_show(self):
        """Método chamado quando a página é exibida (pode ser sobrescrito)"""
        pass


# =========================================================
# COMPONENTE CRUD GENÉRICO (mesma lógica, layout mais robusto)
# =========================================================
# ==========================
# ===== CADASTRO CRUD (ATUALIZADO) =====
# ==========================

class CadastroCRUD(ttk.Frame):
    """Componente CRUD reutilizável para cadastros (com validações por tabela)"""
    def __init__(self, parent, titulo, table, fields, app=None):
        super().__init__(parent, style="Content.TFrame")
        self.table = table
        self.fields = fields
        self.app = app
        self.selected_id = None
        self._is_admin = bool(getattr(self.app, "user", {}).get("is_admin")) if self.app else False

        # Card
        card = ttk.Frame(self, style="Card.TFrame", padding=16)
        card.pack(fill="both", expand=True)

        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(6, weight=1)

        ttk.Label(card, text=titulo, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        # FORM responsivo
        form = ttk.Frame(card, style="Card.TFrame")
        form.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        form.grid_columnconfigure(0, weight=1)

        self.entries = {}

        max_cols = 4
        for idx, (col, label) in enumerate(fields):
            r = (idx // max_cols) * 2
            c = idx % max_cols

            form.grid_columnconfigure(c, weight=1, uniform="formcols")

            ttk.Label(form, text=label, style="CardLabel.TLabel").grid(
                row=r, column=c, sticky="w", padx=6, pady=(0, 4)
            )

            if col == "status" and self.table in {"ajudantes", "motoristas"}:
                if self.table == "motoristas":
                    ent = ttk.Combobox(form, state="readonly", values=["ATIVO", "INATIVO"])
                else:
                    ent = ttk.Combobox(form, state="readonly", values=["ATIVO", "DESATIVADO"])
                ent.set("ATIVO")
            else:
                ent = ttk.Entry(form, style="Field.TEntry")
            ent.grid(row=r + 1, column=c, sticky="ew", padx=6, pady=(0, 10))
            self.entries[col] = ent
            if isinstance(ent, ttk.Entry):
                kind, precision = self._infer_entry_kind(col, label)
                bind_entry_smart(ent, kind, precision=precision)

        # BOTÕES
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(4, 12))
        actions.grid_columnconfigure(20, weight=1)

        ttk.Button(actions, text="ðÅ¸â€™¾ SALVAR", style="Primary.TButton", command=self.salvar).grid(row=0, column=0, padx=6)
        ttk.Button(actions, text="âÅ“Âï¸Â EDITAR", style="Ghost.TButton", command=self.editar).grid(row=0, column=1, padx=6)
        ttk.Button(actions, text="ðÅ¸â€”â€˜ï¸Â EXCLUIR", style="Danger.TButton", command=self.excluir).grid(row=0, column=2, padx=6)
        ttk.Button(actions, text="ðÅ¸§¹ LIMPAR", style="Ghost.TButton", command=self.limpar).grid(row=0, column=3, padx=6)
        ttk.Button(actions, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton", command=self.carregar).grid(row=0, column=4, padx=6)

        if self.table == "ajudantes":
            self._ajudantes_status_filter = tk.StringVar(value="TODOS")
            ttk.Label(actions, text="FILTRO STATUS:", style="CardLabel.TLabel").grid(row=0, column=10, padx=(18, 6), sticky="e")
            cb_status_filter = ttk.Combobox(
                actions,
                state="readonly",
                width=14,
                textvariable=self._ajudantes_status_filter,
                values=["TODOS", "ATIVO", "DESATIVADO"],
            )
            cb_status_filter.grid(row=0, column=11, padx=6, sticky="w")
            cb_status_filter.bind("<<ComboboxSelected>>", lambda _e: self.carregar())

        # Importar clientes (se você ainda usa esse CRUD para clientes em algum lugar)
        if self.table == "clientes":
            ttk.Button(
                actions,
                text="ðÅ¸â€œ¥ IMPORTAR CLIENTES (EXCEL)",
                style="Warn.TButton",
                command=self.importar_clientes_excel
            ).grid(row=0, column=5, padx=6)

        if self.table == "motoristas":
            self.btn_senha = ttk.Button(
                actions,
                text="ALTERAR SENHA",
                style="Warn.TButton",
                command=self.alterar_senha_motorista
            )
            self.btn_senha.grid(row=0, column=6, padx=6)
            self.btn_liberar_app = ttk.Button(
                actions,
                text="LIBERAR APP",
                style="Primary.TButton",
                command=lambda: self.definir_acesso_app_motorista(True),
            )
            self.btn_liberar_app.grid(row=0, column=7, padx=6)
            self.btn_bloquear_app = ttk.Button(
                actions,
                text="BLOQUEAR APP",
                style="Danger.TButton",
                command=lambda: self.definir_acesso_app_motorista(False),
            )
            self.btn_bloquear_app.grid(row=0, column=8, padx=6)

        if self.table == "usuarios":
            self.btn_senha = ttk.Button(
                actions,
                text="ALTERAR SENHA",
                style="Warn.TButton",
                command=self.alterar_senha
            )
            self.btn_senha.grid(row=0, column=6, padx=6)

        
        # Tabela
        cols = ["ID"] + [label for _, label in self.fields]
        self.tree = ttk.Treeview(card, columns=cols, show="headings", height=12)
        self.tree.grid(row=3, column=0, sticky="nsew", pady=(0, 6))
        card.grid_rowconfigure(3, weight=1)

        vsb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=3, column=1, sticky="ns", pady=(0, 6))

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=160, anchor="w")
        self.tree.column("ID", width=70, anchor="center")

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"ID"},
            money_cols=set(),
            date_cols=set()
        )

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.carregar()
        self._update_password_controls()

    # -------------------------
    # Helpers
    # -------------------------
    def _norm(self, v):
        return upper(str(v or "").strip())

    def _infer_entry_kind(self, col: str, label: str):
        c = str(col or "").strip().lower()
        l = upper(label or "")

        if any(k in c for k in ("hora",)) or "HORA" in l:
            return "time", 2
        if any(k in c for k in ("data",)) or "DATA" in l:
            return "date", 2
        if any(k in c for k in ("valor", "preco", "custo", "adiantamento", "diaria")):
            return "money", 2
        if "cpf" in c:
            return "cpf", 2
        if "telefone" in c:
            return "phone", 2
        if any(k in c for k in ("capacidade", "qtd", "quantidade", "caixa", "caixas")):
            return "int", 2
        if any(k in c for k in ("kg", "km", "media", "litros")):
            return "decimal", 2
        if c in {"senha", "codigo", "nome", "sobrenome", "status", "veiculo", "motorista", "equipe"}:
            return "text", 2
        if any(k in l for k in ("VALOR", "PREÇO", "CUSTO", "ADIANTAMENTO", "DIÃÂÂRIA")):
            return "money", 2
        if "CPF" in l:
            return "cpf", 2
        if "TELEFONE" in l:
            return "phone", 2
        if any(k in l for k in ("CAPACIDADE", "CAIXA", "QTD")):
            return "int", 2
        if any(k in l for k in ("KG", "KM", "MÉDIA", "MEDIA", "LITROS")):
            return "decimal", 2
        return "text", 2

    def _get(self, key):
        ent = self.entries.get(key)
        return ent.get().strip() if ent else ""

    def _set(self, key, value):
        ent = self.entries.get(key)
        if not ent:
            return
        txt = "" if value is None else str(value)
        try:
            if isinstance(ent, ttk.Combobox):
                ent.set(txt)
            else:
                ent.delete(0, "end")
                ent.insert(0, txt)
        except Exception:
            logging.debug("Falha ignorada")

    def _user_can_change_password(self, user_id):
        if not self.app:
            return False
        if self._is_admin:
            return True
        return str(user_id) == str(self.app.user.get("id"))

    def _update_password_controls(self):
        if self.table not in {"usuarios", "motoristas"}:
            return
        if self.table == "usuarios":
            can_change = self.selected_id and self._user_can_change_password(self.selected_id)
        else:
            can_change = bool(self.selected_id) and self._is_admin
        if hasattr(self, "btn_senha"):
            self.btn_senha.config(state="normal" if can_change else "disabled")
        senha_ent = self.entries.get("senha")
        if senha_ent:
            if self.table == "usuarios":
                if self._is_admin and not self.selected_id:
                    senha_ent.config(state="normal")
                else:
                    senha_ent.config(state="disabled")
            else:
                if self._is_admin and not self.selected_id:
                    senha_ent.config(state="normal")
                else:
                    senha_ent.config(state="disabled")

    def _require_fields(self, fields_required):
        for f in fields_required:
            if not str(self._get(f)).strip():
                return False, f
        return True, ""

    def _dup_exists(self, cur, colname, value_norm, ignore_id=None):
        """
        Verifica duplicidade case-insensitive:
        SELECT id FROM table WHERE UPPER(col)=UPPER(?) AND id<>?
        """
        if not value_norm:
            return False
        try:
            if ignore_id:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE UPPER({colname})=UPPER(?) AND id<>? LIMIT 1",
                    (value_norm, ignore_id)
                )
            else:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE UPPER({colname})=UPPER(?) LIMIT 1",
                    (value_norm,)
                )
            return cur.fetchone() is not None
        except Exception:
            # Se por algum motivo a coluna não existir no DB, não trava o sistema.
            return False

    def _dup_exists_exact(self, cur, colname, value, ignore_id=None):
        """
        Verifica duplicidade exata (útil para CPF/telefone quando você normaliza para dÃÂÂgitos).
        """
        if not value:
            return False
        try:
            if ignore_id:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE {colname}=? AND id<>? LIMIT 1",
                    (value, ignore_id)
                )
            else:
                cur.execute(
                    f"SELECT id FROM {self.table} WHERE {colname}=? LIMIT 1",
                    (value,)
                )
            return cur.fetchone() is not None
        except Exception:
            return False

    def _sync_motorista_upsert_api(self, data: dict):
        if self.table != "motoristas":
            return
        if not is_desktop_api_sync_enabled():
            return
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return
        codigo = self._norm(data.get("codigo"))
        nome = self._norm(data.get("nome"))
        if not codigo or not nome:
            return
        acesso_liberado_raw = data.get("acesso_liberado", None)
        acesso_liberado_por = data.get("acesso_liberado_por", None)
        acesso_obs = data.get("acesso_obs", None)

        # Regra de negócio:
        # - se vier explÃÂcito no data, usa esse valor
        # - se NÃO vier (caso comum no formulário), preserva estado atual no banco
        #   para não desbloquear motorista sem clicar em "LIBERAR APP".
        if acesso_liberado_raw in (None, ""):
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("PRAGMA table_info(motoristas)")
                    cols_m = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                    if "acesso_liberado" in cols_m:
                        cur.execute(
                            """
                            SELECT
                                COALESCE(acesso_liberado,1) AS acesso_liberado,
                                COALESCE(acesso_liberado_por,'') AS acesso_liberado_por,
                                COALESCE(acesso_obs,'') AS acesso_obs
                            FROM motoristas
                            WHERE UPPER(COALESCE(codigo,''))=UPPER(?)
                            ORDER BY id DESC
                            LIMIT 1
                            """,
                            (codigo,),
                        )
                        row = cur.fetchone()
                        if row:
                            acesso_liberado_raw = int(row[0] or 0)
                            if not acesso_liberado_por:
                                acesso_liberado_por = str(row[1] or "").strip()
                            if not acesso_obs:
                                acesso_obs = str(row[2] or "").strip()
            except Exception:
                logging.debug("Falha ao preservar acesso_liberado do motorista para sync", exc_info=True)

        if acesso_liberado_raw in (None, ""):
            acesso_liberado_payload = None
        else:
            try:
                acesso_liberado_payload = bool(int(acesso_liberado_raw or 0))
            except Exception:
                acesso_liberado_payload = bool(acesso_liberado_raw)

        payload = {
            "codigo": codigo,
            "nome": nome,
            "telefone": self._norm(data.get("telefone")),
            "cpf": self._norm(data.get("cpf")),
            "status": self._norm(data.get("status") or "ATIVO"),
            "senha": data.get("senha") or None,
            "acesso_liberado": acesso_liberado_payload,
            "acesso_liberado_por": self._norm(acesso_liberado_por or "DESKTOP_SYNC"),
            "acesso_obs": self._norm(acesso_obs or "Sincronizado via Desktop"),
        }
        _call_api(
            "POST",
            "desktop/cadastros/motoristas/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )

    def _sync_veiculo_upsert_api(self, data: dict):
        if self.table != "veiculos":
            return
        if not is_desktop_api_sync_enabled():
            return
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return
        placa = self._norm(data.get("placa"))
        modelo = self._norm(data.get("modelo"))
        if not placa or not modelo:
            return
        payload = {
            "placa": placa,
            "modelo": modelo,
            "capacidade_cx": int(safe_int(data.get("capacidade_cx"), 0)),
        }
        _call_api(
            "POST",
            "desktop/cadastros/veiculos/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )

    def _sync_ajudante_upsert_api(self, data: dict):
        if self.table != "ajudantes":
            return
        if not is_desktop_api_sync_enabled():
            return
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return
        nome = self._norm(data.get("nome"))
        sobrenome = self._norm(data.get("sobrenome"))
        if not nome or not sobrenome:
            return
        payload = {
            "nome": nome,
            "sobrenome": sobrenome,
            "telefone": self._norm(data.get("telefone")),
            "status": self._norm(data.get("status") or "ATIVO"),
        }
        _call_api(
            "POST",
            "desktop/cadastros/ajudantes/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )

    def _has_delete_references(self, cur) -> str:
        if not self.selected_id:
            return ""
        try:
            if self.table == "motoristas":
                cod = self._norm(self._get("codigo"))
                nome = self._norm(self._get("nome"))
                if not cod and not nome:
                    return ""
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                wh = []
                params = []
                if "motorista_codigo" in cols and cod:
                    wh.append("UPPER(COALESCE(motorista_codigo,''))=UPPER(?)")
                    params.append(cod)
                if "codigo_motorista" in cols and cod:
                    wh.append("UPPER(COALESCE(codigo_motorista,''))=UPPER(?)")
                    params.append(cod)
                if "motorista" in cols:
                    if cod:
                        wh.append("UPPER(COALESCE(motorista,''))=UPPER(?)")
                        params.append(cod)
                    if nome:
                        wh.append("UPPER(COALESCE(motorista,''))=UPPER(?)")
                        params.append(nome)
                if wh:
                    cur.execute(
                        "SELECT COUNT(*) FROM programacoes WHERE " + " OR ".join(wh),
                        tuple(params),
                    )
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Motorista vinculado a programação/rota."
            elif self.table == "veiculos":
                placa = self._norm(self._get("placa"))
                if placa:
                    cur.execute("SELECT COUNT(*) FROM programacoes WHERE UPPER(COALESCE(veiculo,''))=UPPER(?)", (placa,))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "VeÃÂculo vinculado a programação/rota."
            elif self.table == "clientes":
                cod = self._norm(self._get("cod_cliente"))
                if cod:
                    cur.execute("SELECT COUNT(*) FROM programacao_itens WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)", (cod,))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Cliente vinculado a programação."
                    try:
                        cur.execute("SELECT COUNT(*) FROM recebimentos WHERE UPPER(COALESCE(cod_cliente,''))=UPPER(?)", (cod,))
                        if int((cur.fetchone() or [0])[0] or 0) > 0:
                            return "Cliente vinculado a recebimentos."
                    except Exception:
                        logging.debug("Falha ignorada")
            elif self.table == "ajudantes":
                aid = str(self.selected_id)
                nome = self._norm(self._get("nome"))
                sobrenome = self._norm(self._get("sobrenome"))
                alvo = f"{nome} {sobrenome}".strip()
                cur.execute("PRAGMA table_info(equipes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                conds = []
                params = []
                for c in ("ajudante1", "ajudante2", "ajudante_1", "ajudante_2"):
                    if c in cols:
                        conds.append(f"UPPER(COALESCE({c},''))=UPPER(?)")
                        params.append(aid)
                        if nome:
                            conds.append(f"UPPER(COALESCE({c},''))=UPPER(?)")
                            params.append(nome)
                        if alvo:
                            conds.append(f"UPPER(COALESCE({c},''))=UPPER(?)")
                            params.append(alvo)
                if conds:
                    cur.execute("SELECT COUNT(*) FROM equipes WHERE " + " OR ".join(conds), tuple(params))
                    if int((cur.fetchone() or [0])[0] or 0) > 0:
                        return "Ajudante vinculado a equipe."
        except Exception:
            logging.debug("Falha na verificação de vÃÂnculos para exclusão", exc_info=True)
        return ""

    # -------------------------
    # Ações padrão
    # -------------------------
    def limpar(self):
        self.selected_id = None
        for col, _ in self.fields:
            self._set(col, "")
        if self.table == "ajudantes" and "status" in self.entries:
            self._set("status", "ATIVO")
        if self.table == "motoristas" and "status" in self.entries:
            self._set("status", "ATIVO")
        self._update_password_controls()

    def carregar(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                api_rows = None
                if self.table == "motoristas":
                    api_rows = _call_api(
                        "GET",
                        "desktop/cadastros/motoristas",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                elif self.table == "veiculos":
                    api_rows = _call_api(
                        "GET",
                        "desktop/cadastros/veiculos",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                elif self.table == "ajudantes":
                    api_rows = _call_api(
                        "GET",
                        "desktop/cadastros/ajudantes",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )

                if isinstance(api_rows, list):
                    rows = []
                    for r in api_rows:
                        if not isinstance(r, dict):
                            continue
                        row_map = {}
                        if self.table == "motoristas":
                            row_map = {
                                "nome": str(r.get("nome") or ""),
                                "codigo": str(r.get("codigo") or ""),
                                "senha": "",
                                "cpf": "",
                                "telefone": "",
                                "status": str(r.get("status") or "ATIVO"),
                            }
                        elif self.table == "veiculos":
                            row_map = {
                                "placa": str(r.get("placa") or ""),
                                "modelo": str(r.get("modelo") or ""),
                                "capacidade_cx": str(r.get("capacidade_cx") or 0),
                            }
                        elif self.table == "ajudantes":
                            row_map = {
                                "nome": str(r.get("nome_base") or r.get("nome") or ""),
                                "sobrenome": str(r.get("sobrenome") or ""),
                                "telefone": str(r.get("telefone") or ""),
                                "status": str(r.get("status") or "ATIVO"),
                            }
                        row = [int(safe_int(r.get("id"), 0))]
                        for c, _ in self.fields:
                            row.append(row_map.get(c, ""))
                        rows.append(tuple(row))

                    self.tree.delete(*self.tree.get_children())
                    senha_pos = None
                    if self.table in {"usuarios", "motoristas"}:
                        try:
                            senha_pos = 1 + [c for c, _ in self.fields].index("senha")
                        except Exception:
                            senha_pos = None
                    for row in rows:
                        if senha_pos is not None:
                            row = list(row)
                            row[senha_pos] = "******" if row[senha_pos] else ""
                            row = tuple(row)
                        tree_insert_aligned(self.tree, "", "end", row)
                    return
            except Exception:
                logging.debug("Falha ao carregar cadastro via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute(f"PRAGMA table_info({self.table})")
                cols_exist = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            except Exception:
                cols_exist = set()

            select_cols = []
            for c, _ in self.fields:
                if str(c).lower() in cols_exist:
                    select_cols.append(c)
                else:
                    select_cols.append(f"'' AS {c}")
            cols_db = ", ".join(select_cols)
            if self.table == "ajudantes":
                status_filter = "TODOS"
                try:
                    status_filter = upper(self._ajudantes_status_filter.get())
                except Exception:
                    status_filter = "TODOS"

                if status_filter in {"ATIVO", "DESATIVADO"}:
                    cur.execute(
                        f"SELECT id, {cols_db} FROM {self.table} "
                        "WHERE UPPER(COALESCE(status, 'ATIVO'))=? ORDER BY id DESC",
                        (status_filter,),
                    )
                else:
                    cur.execute(f"SELECT id, {cols_db} FROM {self.table} ORDER BY id DESC")
            else:
                cur.execute(f"SELECT id, {cols_db} FROM {self.table} ORDER BY id DESC")
            rows = cur.fetchall() or []

        self.tree.delete(*self.tree.get_children())
        senha_pos = None
        if self.table in {"usuarios", "motoristas"}:
            try:
                senha_pos = 1 + [c for c, _ in self.fields].index("senha")
            except Exception:
                senha_pos = None
        for row in rows:
            if senha_pos is not None:
                row = list(row)
                row[senha_pos] = "******" if row[senha_pos] else ""
                row = tuple(row)
            tree_insert_aligned(self.tree, "", "end", row)

    def salvar(self):
        """
        Salva/atualiza com validações por cadastro:
        - motoristas: valida NOME/CPF/TELEFONE; exige codigo e senha (manual); codigo único; CPF único (se houver coluna)
        - usuarios: exige nome, senha; nome único; sem codigo
        - veiculos: exige placa, modelo, capacidade_cx; placa única; capacidade int >= 0

        Obs: ID é a chave primária autogerada (sequência do SQLite). Você já vê o ID na tabela.
        """
        if not ensure_system_api_binding(context=f"Salvar cadastro ({self.table})", parent=self):
            return

        data = {col: self.entries[col].get().strip() for col, _ in self.fields}

        # Normalizações por tabela
        if self.table in {"motoristas", "usuarios", "veiculos", "equipes", "ajudantes", "clientes"}:
            for k in list(data.keys()):
                data[k] = self._norm(data[k])

        saved_via_api = False
        # Validar obrigatórios + duplicidade + tipos
        try:
            with get_db() as conn:
                cur = conn.cursor()

                # -------------------------
                # MOTORISTAS
                # -------------------------
                if self.table == "motoristas":
                    # Exige NOME, CÓDIGO e SENHA (manuais) + TELEFONE
                    # CPF é opcional, porém deve ser válido e único quando informado.
                    ok, f = self._require_fields(["nome", "codigo", "senha", "telefone"])
                    if not ok:
                        messagebox.showwarning("ATENÇÃO", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    # Nome (básico)
                    nome = self._norm(data.get("nome"))
                    if len(nome) < 3:
                        messagebox.showwarning("ATENÇÃO", "NOME deve ter pelo menos 3 caracteres.")
                        return
                    data["nome"] = nome

                    # Código (login flutter)  manual, mas pode usar GERAR
                    cod = self._norm(data.get("codigo"))
                    if not is_valid_motorista_codigo(cod):
                        messagebox.showwarning(
                            "ATENÇÃO",
                            "CÓDIGO inválido. Use apenas letras/números/._- e 3 a 24 caracteres."
                        )
                        return
                    if self._dup_exists(cur, "codigo", cod, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE MOTORISTA COM ESTE CÓDIGO: {cod}")
                        return
                    data["codigo"] = cod

                    # Senha (login flutter)  manual
                    senha = (self.entries.get("senha").get().strip() if self.entries.get("senha") else "").strip()
                    if self.selected_id:
                        if senha:
                            if not self._is_admin:
                                messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar senha do motorista.")
                                return
                            if not is_valid_motorista_senha(senha):
                                messagebox.showwarning(
                                    "ATENÇÃO",
                                    "SENHA inválida. Use 4 a 24 caracteres (não pode ser vazia)."
                                )
                                return
                            data["senha"] = hash_password_pbkdf2(senha)
                        else:
                            data.pop("senha", None)
                    else:
                        if not is_valid_motorista_senha(senha):
                            messagebox.showwarning(
                                "ATENÇÃO",
                                "SENHA inválida. Use 4 a 24 caracteres (não pode ser vazia)."
                            )
                            return
                        data["senha"] = hash_password_pbkdf2(senha)

                    # CPF (opcional)
                    cpf_raw = self.entries.get("cpf").get().strip() if self.entries.get("cpf") else ""
                    cpf = normalize_cpf(cpf_raw)
                    if cpf and not is_valid_cpf(cpf):
                        messagebox.showwarning("ATENÇÃO", "CPF inválido.")
                        return
                    data["cpf"] = cpf  # salva só dÃÂÂgitos

                    # Telefone
                    tel_raw = self.entries.get("telefone").get().strip() if self.entries.get("telefone") else ""
                    tel = normalize_phone(tel_raw)
                    if not is_valid_phone(tel):
                        messagebox.showwarning(
                            "ATENÇÃO",
                            "TELEFONE inválido. Informe DDD+Número (10 ou 11 dÃÂÂgitos)."
                        )
                        return
                    data["telefone"] = tel  # salva só dÃÂÂgitos

                    # Status do motorista
                    status_m = self._norm(data.get("status") or "ATIVO")
                    if status_m not in {"ATIVO", "DESATIVADO"}:
                        status_m = "ATIVO"
                    data["status"] = status_m

                    # CPF único quando informado (se a coluna existir)
                    if cpf and db_has_column(cur, "motoristas", "cpf"):
                        if self._dup_exists_exact(cur, "cpf", cpf, ignore_id=self.selected_id):
                            messagebox.showerror("ERRO", "J EXISTE MOTORISTA COM ESTE CPF.")
                            return

                    # Fluxo natural: cadastro novo de motorista já nasce apto ao app mobile.
                    # Mantém compatibilidade com bases que ainda não possuem essas colunas.
                    if not self.selected_id:
                        try:
                            cur.execute("PRAGMA table_info(motoristas)")
                            cols_m = {str(r[1]).lower() for r in (cur.fetchall() or [])}
                        except Exception:
                            cols_m = set()
                        if "acesso_liberado" in cols_m and "acesso_liberado" not in data:
                            data["acesso_liberado"] = 1
                        if "acesso_liberado_por" in cols_m and not data.get("acesso_liberado_por"):
                            data["acesso_liberado_por"] = "AUTO_DESKTOP"
                        if "acesso_liberado_em" in cols_m and not data.get("acesso_liberado_em"):
                            data["acesso_liberado_em"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        if "acesso_obs" in cols_m and not data.get("acesso_obs"):
                            data["acesso_obs"] = "Liberacao automatica no cadastro desktop"

                # -------------------------
                # USURIOS (login por nome + senha)
                # -------------------------
                if self.table == "usuarios":
                    nome = self._norm(data.get("nome"))
                    senha_plana =""
                    try:
                        senha_plana = (self.entries.get("senha").get() or "").strip()
                    except Exception:
                        senha_plana = str(data.get("senha") or "").strip()

                    if not nome:
                        messagebox.showwarning("ANTENÇÃO", "PREENCHA O CAMPO: NOME.")
                        return
                    if self._dup_exists(cur, "nome", nome, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE USURIO COM ESTE NOME:{nome}")
                        return

                    data["nome"] = nome
                    if self.selected_id:
                        if senha_plana:
                            if not self._is_admin:
                                messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar senha aqui.")
                                return
                            if len(senha_plana) < 6:
                                messagebox.showwarning("ANTENÇÃO", "SENHA deve ter pelo menos 6 caracteres.")
                                return
                            data["senha"] = hash_password_pbkdf2(senha_plana)
                        else:
                            data.pop("senha", None)
                    else:
                        if not senha_plana:
                            messagebox.showwarning("ANTENÇÃO", "PREENCHA O CAMPO: SENHA.")
                            return
                        if len(senha_plana) < 6:
                            messagebox.showwarning("ANTENÇÃO", "SENHA deve ter pelo menos 6 caracteres.")
                            return
                        data["senha"] = hash_password_pbkdf2(senha_plana)

                # -------------------------
                # AJUDANTES (nome, sobrenome, telefone)
                # -------------------------
                if self.table == "ajudantes":
                    ok, f = self._require_fields(["nome", "sobrenome", "telefone", "status"])
                    if not ok:
                        messagebox.showwarning("ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    nome = self._norm(data.get("nome"))
                    sobrenome = self._norm(data.get("sobrenome"))
                    telefone = normalize_phone(data.get("telefone"))
                    status = self._norm(data.get("status"))

                    if len(nome) < 2:
                        messagebox.showwarning("ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O", "NOME deve ter pelo menos 2 caracteres.")
                        return
                    if len(sobrenome) < 2:
                        messagebox.showwarning("ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O", "SOBRENOME deve ter pelo menos 2 caracteres.")
                        return
                    if not is_valid_phone(telefone):
                        messagebox.showwarning(
                            "ATENÃâââ€š¬¡ÃÆââ‚¬â„¢O",
                            "TELEFONE inválido. Informe DDD+Número (10 ou 11 dÃÂÂÂgitos)."
                        )
                        return
                    if status not in {"ATIVO", "DESATIVADO"}:
                        messagebox.showwarning("ATENÇÃO", "STATUS inválido para ajudante.")
                        return

                    # Evita duplicidade de pessoa com mesmo nome completo.
                    try:
                        if self.selected_id:
                            cur.execute(
                                """
                                SELECT id FROM ajudantes
                                WHERE UPPER(nome)=UPPER(?) AND UPPER(sobrenome)=UPPER(?) AND id<>?
                                LIMIT 1
                                """,
                                (nome, sobrenome, self.selected_id),
                            )
                        else:
                            cur.execute(
                                """
                                SELECT id FROM ajudantes
                                WHERE UPPER(nome)=UPPER(?) AND UPPER(sobrenome)=UPPER(?)
                                LIMIT 1
                                """,
                                (nome, sobrenome),
                            )
                        if cur.fetchone():
                            messagebox.showerror("ERRO", "JÃ EXISTE AJUDANTE COM ESTE NOME/SOBRENOME.")
                            return
                    except Exception:
                        logging.debug("Falha ignorada")

                    data["nome"] = nome
                    data["sobrenome"] = sobrenome
                    data["telefone"] = telefone
                    data["status"] = status

                # -------------------------
                # VECULOS (placa, modelo, capacidade_cx)
                # -------------------------
                if self.table == "veiculos":
                    ok, f = self._require_fields(["placa", "modelo", "capacidade_cx"])
                    if not ok:
                        messagebox.showwarning("ATENÇÃO", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    placa = self._norm(data.get("placa"))
                    if self._dup_exists(cur, "placa", placa, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE VECULO COM ESTA PLACA: {placa}")
                        return

                    cap = safe_int(data.get("capacidade_cx"), -1)
                    if cap < 0:
                        messagebox.showwarning("ATENÇÃO", "CAPACIDADE (CX) deve ser um número inteiro >= 0.")
                        return
                    data["capacidade_cx"] = cap

                # -------------------------
                # SALVAR (UPDATE/INSERT)
                # -------------------------
                desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                if self.table in {"motoristas", "veiculos", "ajudantes"} and desktop_secret and is_desktop_api_sync_enabled():
                    try:
                        if self.table == "motoristas":
                            self._sync_motorista_upsert_api(data)
                        elif self.table == "veiculos":
                            self._sync_veiculo_upsert_api(data)
                        elif self.table == "ajudantes":
                            self._sync_ajudante_upsert_api(data)
                        saved_via_api = True
                    except Exception:
                        logging.debug("Falha ao salvar cadastro via API; usando fallback local.", exc_info=True)

                if not saved_via_api:
                    if self.selected_id:
                        sets = ", ".join([f"{c}=?" for c in data.keys()])
                        values = list(data.values()) + [self.selected_id]
                        cur.execute(f"UPDATE {self.table} SET {sets} WHERE id=?", values)
                    else:
                        cols = ", ".join(data.keys())
                        qs = ", ".join(["?"] * len(data))
                        cur.execute(f"INSERT INTO {self.table} ({cols}) VALUES ({qs})", list(data.values()))

        except sqlite3.IntegrityError as e:
            messagebox.showerror("ERRO", f"REGISTRO DUPLICADO OU INVLIDO.\n\n{e}")
            return
        except Exception as e:
            messagebox.showerror("ERRO", str(e))
            return

        if (not saved_via_api) and self.table == "motoristas":
            try:
                self._sync_motorista_upsert_api(data)
            except Exception as e:
                messagebox.showwarning(
                    "Sincronização",
                    "Motorista salvo localmente, mas falhou ao sincronizar com a API.\n"
                    f"Detalhe: {e}",
                )
        elif (not saved_via_api) and self.table == "veiculos":
            try:
                self._sync_veiculo_upsert_api(data)
            except Exception as e:
                messagebox.showwarning(
                    "Sincronização",
                    "VeÃÂculo salvo localmente, mas falhou ao sincronizar com a API.\n"
                    f"Detalhe: {e}",
                )
        elif (not saved_via_api) and self.table == "ajudantes":
            try:
                self._sync_ajudante_upsert_api(data)
            except Exception as e:
                messagebox.showwarning(
                    "Sincronização",
                    "Ajudante salvo localmente, mas falhou ao sincronizar com a API.\n"
                    f"Detalhe: {e}",
                )

        self.carregar()
        self.limpar()

        if self.app:
            self.app.refresh_programacao_comboboxes()

    def excluir(self):
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM ITEM NA TABELA ANTES DE EXCLUIR.")
            return

        if not messagebox.askyesno("CONFIRMAR", "DESEJA EXCLUIR ESTE REGISTRO?"):
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                endpoint = ""
                if self.table == "motoristas":
                    cod = self._norm(self._get("codigo"))
                    if cod:
                        endpoint = f"desktop/cadastros/motoristas/{urllib.parse.quote(cod)}"
                elif self.table == "veiculos":
                    placa = self._norm(self._get("placa"))
                    if placa:
                        endpoint = f"desktop/cadastros/veiculos/{urllib.parse.quote(placa)}"
                elif self.table == "ajudantes":
                    endpoint = f"desktop/cadastros/ajudantes/{int(safe_int(self.selected_id, 0))}"
                elif self.table == "clientes":
                    cod_cli = self._norm(self._get("cod_cliente"))
                    if cod_cli:
                        endpoint = f"desktop/cadastros/clientes/{urllib.parse.quote(cod_cli)}"

                if endpoint:
                    _call_api(
                        "DELETE",
                        endpoint,
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    self.carregar()
                    self.limpar()
                    if self.app:
                        self.app.refresh_programacao_comboboxes()
                    return
            except Exception:
                logging.debug("Falha ao excluir cadastro via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            bloqueio = self._has_delete_references(cur)
            if bloqueio:
                messagebox.showwarning("ATENÇÃO", f"Exclusão bloqueada.\n\n{bloqueio}\nUse status DESATIVADO.")
                return
            cur.execute(f"DELETE FROM {self.table} WHERE id=?", (self.selected_id,))

        self.carregar()
        self.limpar()

        if self.app:
            self.app.refresh_programacao_comboboxes()

    def editar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("ATENÇÃO", "Selecione um item na tabela para editar.")
            return
        self._on_select()
        try:
            first_col = self.fields[0][0] if self.fields else ""
            ent = self.entries.get(first_col)
            if ent:
                ent.focus_set()
        except Exception:
            logging.debug("Falha ignorada")

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        row = self.tree.item(sel[0], "values")
        self.selected_id = row[0]
        for i, (col, _) in enumerate(self.fields, start=1):
            if self.table in {"usuarios", "motoristas"} and col == "senha":
                self._set(col, "")
                continue
            self._set(col, row[i] if i < len(row) else "")
        self._update_password_controls()

        # Mantive seu importar_clientes_excel como estava (se você ainda usa em algum ponto)
    def importar_clientes_excel(self):
        if not ensure_system_api_binding(context="Importar clientes (cadastros)", parent=self):
            return
        path = filedialog.askopenfilename(
            title="IMPORTAR CLIENTES (EXCEL)",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if not path:
            return

        if not (require_pandas() and require_excel_support(path)):
            return

        try:
            df = pd.read_excel(path, engine=excel_engine_for(path))

            col_cod = guess_col(df.columns, ["cod", "cód", "codigo", "cliente", "cod cliente"])
            col_nome = guess_col(df.columns, ["nome", "cliente"])
            col_end = guess_col(df.columns, ["endereco", "endereço", "rua", "logradouro"])
            col_bairro = guess_col(df.columns, ["bairro"])
            col_cidade = guess_col(df.columns, ["cidade"])
            col_uf = guess_col(df.columns, ["uf", "estado"])
            col_tel = guess_col(df.columns, ["telefone", "fone", "celular", "contato"])
            col_rota = guess_col(df.columns, ["rota"])

            if not col_cod or not col_nome:
                cols = list(df.columns or [])
                if len(cols) >= 2:
                    col_cod = col_cod or cols[0]
                    col_nome = col_nome or cols[1]
                else:
                    messagebox.showerror("ERRO", "NÃO IDENTIFIQUEI AS COLUNAS DE CÓDIGO E NOME DO CLIENTE NO EXCEL.")
                    return

            total = 0
            with get_db() as conn:
                cur = conn.cursor()
                for _, r in df.iterrows():
                    cod = str(r.get(col_cod, "")).strip()
                    nome = str(r.get(col_nome, "")).strip()
                    if not cod or not nome:
                        continue

                    endereco = str(r.get(col_end, "")).strip() if col_end else ""
                    bairro = str(r.get(col_bairro, "")).strip() if col_bairro else ""
                    cidade = str(r.get(col_cidade, "")).strip() if col_cidade else ""
                    uf = str(r.get(col_uf, "")).strip() if col_uf else ""
                    telefone = str(r.get(col_tel, "")).strip() if col_tel else ""
                    rota = str(r.get(col_rota, "")).strip() if col_rota else ""

                    cur.execute("""
                        INSERT INTO clientes (cod_cliente, nome_cliente, endereco, bairro, cidade, uf, telefone, rota)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(cod_cliente) DO UPDATE SET
                            nome_cliente=excluded.nome_cliente,
                            endereco=excluded.endereco,
                            bairro=excluded.bairro,
                            cidade=excluded.cidade,
                            uf=excluded.uf,
                            telefone=excluded.telefone,
                            rota=excluded.rota
                    """, (
                        upper(cod),
                        upper(nome),
                        upper(endereco),
                        upper(bairro),
                        upper(cidade),
                        upper(uf),
                        upper(telefone),
                        upper(rota)
                    ))
                    total += 1

            self.carregar()
            messagebox.showinfo("OK", f"CLIENTES IMPORTADOS/ATUALIZADOS: {total}")

        except Exception as e:
            messagebox.showerror("ERRO", str(e))

    def alterar_senha(self):
        if self.table != "usuarios":
            return
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM USURIO NA TABELA.")
            return
        if not self._user_can_change_password(self.selected_id):
            messagebox.showwarning("ATENÇÃO", "Você não tem permissão para alterar esta senha.")
            return

        require_current = not self._is_admin

        win = tk.Toplevel(self)
        win.title("Alterar Senha")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        row = 0
        ent_atual = None
        if require_current:
            ttk.Label(frm, text="Senha atual:").grid(row=row, column=0, sticky="w", pady=2)
            ent_atual = ttk.Entry(frm, show="*")
            ent_atual.grid(row=row, column=1, sticky="ew", pady=2)
            row += 1

        ttk.Label(frm, text="Nova senha:").grid(row=row, column=0, sticky="w", pady=2)
        ent_nova = ttk.Entry(frm, show="*")
        ent_nova.grid(row=row, column=1, sticky="ew", pady=2)
        row += 1

        ttk.Label(frm, text="Confirmar:").grid(row=row, column=0, sticky="w", pady=2)
        ent_conf = ttk.Entry(frm, show="*")
        ent_conf.grid(row=row, column=1, sticky="ew", pady=2)
        row += 1

        frm.grid_columnconfigure(1, weight=1)

        def _salvar():
            senha_atual = ent_atual.get().strip() if ent_atual else ""
            senha_nova = ent_nova.get().strip()
            senha_conf = ent_conf.get().strip()

            if not senha_nova:
                messagebox.showwarning("ATENÇÃO", "Informe a nova senha.")
                return
            if len(senha_nova) < 6:
                messagebox.showwarning("ATENÇÃO", "A nova senha deve ter pelo menos 6 caracteres.")
                return
            if senha_nova != senha_conf:
                messagebox.showwarning("ATENÇÃO", "Confirmação não confere.")
                return

            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT senha FROM usuarios WHERE id=?", (self.selected_id,))
                row_db = cur.fetchone()
                if not row_db:
                    messagebox.showerror("ERRO", "Usuário não encontrado.")
                    return
                senha_db = row_db[0] or ""

                if require_current:
                    ok = False
                    if str(senha_db).startswith("pbkdf2_sha256$"):
                        ok = verify_password_pbkdf2(senha_atual, senha_db)
                    else:
                        ok = (senha_db == senha_atual)
                    if not ok:
                        messagebox.showwarning("ATENÇÃO", "Senha atual inválida.")
                        return

                nova_hash = hash_password_pbkdf2(senha_nova)
                cur.execute("UPDATE usuarios SET senha=? WHERE id=?", (nova_hash, self.selected_id))

            messagebox.showinfo("OK", "Senha atualizada com sucesso.")
            try:
                win.destroy()
            except Exception:
                logging.debug("Falha ignorada")
            self.carregar()
            self.limpar()

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="ðÅ¸â€™¾ SALVAR", style="Primary.TButton", command=_salvar).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="âÅ“â€“ CANCELAR", style="Ghost.TButton", command=win.destroy).grid(row=0, column=1, padx=6)

    def alterar_senha_motorista(self):
        if self.table != "motoristas":
            return
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM MOTORISTA NA TABELA.")
            return
        if not self._is_admin:
            messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar senha do motorista.")
            return

        win = tk.Toplevel(self)
        win.title("Alterar Senha do Motorista")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Nova senha:").grid(row=0, column=0, sticky="w", pady=2)
        ent_nova = ttk.Entry(frm, show="*")
        ent_nova.grid(row=0, column=1, sticky="ew", pady=2)

        ttk.Label(frm, text="Confirmar:").grid(row=1, column=0, sticky="w", pady=2)
        ent_conf = ttk.Entry(frm, show="*")
        ent_conf.grid(row=1, column=1, sticky="ew", pady=2)

        frm.grid_columnconfigure(1, weight=1)

        def _salvar():
            senha_nova = ent_nova.get().strip()
            senha_conf = ent_conf.get().strip()
            if not senha_nova:
                messagebox.showwarning("ATENÇÃO", "Informe a nova senha.")
                return
            if not is_valid_motorista_senha(senha_nova):
                messagebox.showwarning(
                    "ATENÇÃO",
                    "SENHA inválida. Use 4 a 24 caracteres (não pode ser vazia)."
                )
                return
            if senha_nova != senha_conf:
                messagebox.showwarning("ATENÇÃO", "Confirmação não confere.")
                return

            codigo = self._norm(self._get("codigo"))
            if not codigo:
                messagebox.showwarning("ATENÇÃO", "Código do motorista não encontrado.")
                return

            if not ensure_system_api_binding(context=f"Alterar senha do motorista {codigo}", parent=self):
                return

            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            admin_nome = upper((getattr(self.app, "user", {}) or {}).get("nome", "")) or "ADMIN_DESKTOP"
            payload = {
                "nova_senha": senha_nova,
                "admin": admin_nome,
                "motivo": "Alteracao de senha no cadastro de motoristas (desktop)",
            }
            try:
                _call_api(
                    "POST",
                    f"admin/motoristas/senha/{urllib.parse.quote(codigo)}",
                    payload=payload,
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
            except Exception as exc:
                messagebox.showerror("ERRO", f"Falha ao sincronizar senha do motorista na API:\n{exc}")
                return

            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("UPDATE motoristas SET senha=? WHERE id=?", (self._norm(senha_nova), self.selected_id))

            messagebox.showinfo("OK", "Senha do motorista atualizada com sucesso (LOCAL + API).")
            try:
                win.destroy()
            except Exception:
                logging.debug("Falha ignorada")
            self.carregar()
            self.limpar()

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="ðÅ¸â€™¾ SALVAR", style="Primary.TButton", command=_salvar).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="âÅ“â€“ CANCELAR", style="Ghost.TButton", command=win.destroy).grid(row=0, column=1, padx=6)

    def definir_acesso_app_motorista(self, liberar: bool):
        if self.table != "motoristas":
            return
        if not self._is_admin:
            messagebox.showwarning("ATENÇÃO", "Somente ADMIN pode alterar acesso ao app.")
            return
        if not self.selected_id:
            messagebox.showwarning("ATENÇÃO", "SELECIONE UM MOTORISTA NA TABELA.")
            return

        codigo = self._norm(self._get("codigo"))
        if not codigo:
            messagebox.showwarning("ATENÇÃO", "Código do motorista não encontrado no item selecionado.")
            return

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            messagebox.showerror(
                "ERRO",
                "ROTA_SECRET não configurada no Desktop.\nDefina a variável de ambiente e tente novamente.",
            )
            return

        acao = "LIBERAR" if liberar else "BLOQUEAR"
        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja {acao} o acesso do motorista {codigo} ao app?",
        ):
            return

        admin_nome = upper((getattr(self.app, "user", {}) or {}).get("nome", "")) or "ADMIN_DESKTOP"
        payload = {
            "liberado": bool(liberar),
            "admin": admin_nome,
            "motivo": "Definido no cadastro de motoristas (desktop)",
        }
        if not ensure_system_api_binding(context=f"Alterar acesso do motorista {codigo}", parent=self):
            return
        try:
            _call_api(
                "POST",
                f"admin/motoristas/acesso/{urllib.parse.quote(codigo)}",
                payload=payload,
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            messagebox.showinfo("OK", f"Acesso ao app atualizado para {codigo}.")
            return
        except Exception as exc:
            messagebox.showerror("ERRO", f"Falha ao atualizar acesso do motorista:\n{exc}")
            return


# ==========================
# ===== CADASTROS PAGE (ATUALIZADA) =====
# ==========================

class CadastrosPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Cadastros")

        card = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        card.grid(row=0, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        nb = ttk.Notebook(card)
        nb.grid(row=0, column=0, sticky="nsew")

        # -------------------------
        # MOTORISTAS
        # - ID é gerado pelo SQLite (sequência/autoincrement)
        # - valida CPF/NOME/TELEFONE
        # - código e senha são MANUAIS (botão GERAR é opcional)
        # -------------------------
        frm_motoristas = ttk.Frame(nb, style="Content.TFrame")
        crud_motoristas = CadastroCRUD(
            frm_motoristas,
            "Motoristas",
            "motoristas",
            [
                ("nome", "NOME"),
                ("codigo", "CÓDIGO"),
                ("senha", "SENHA"),
                ("cpf", "CPF"),
                ("telefone", "TELEFONE"),
                ("status", "STATUS"),
            ],
            app=app
        )
        crud_motoristas.pack(fill="both", expand=True)
        nb.add(frm_motoristas, text="Motoristas")

        # -------------------------
        # USURIOS (login por nome + senha, sem idade e sem código)
        # -------------------------
        frm_usuarios = ttk.Frame(nb, style="Content.TFrame")
        crud_usuarios = CadastroCRUD(
            frm_usuarios,
            "Usuários",
            "usuarios",
            [
                ("nome", "NOME"),
                ("senha", "SENHA"),
                ("permissoes", "PERMISSÕES"),
                ("cpf", "CPF"),
                ("telefone", "TELEFONE"),
            ],
            app=app
        )
        crud_usuarios.pack(fill="both", expand=True)
        nb.add(frm_usuarios, text="Usuários")

        # -------------------------
        # VECULOS (placa, modelo, capacidade_cx)
        # -------------------------
        frm_veiculos = ttk.Frame(nb, style="Content.TFrame")
        crud_veiculos = CadastroCRUD(
            frm_veiculos,
            "VeÃÂÂculos",
            "veiculos",
            [
                ("placa", "PLACA"),
                ("modelo", "MODELO"),
                ("capacidade_cx", "CAPACIDADE (CX)"),
            ],
            app=app
        )
        crud_veiculos.pack(fill="both", expand=True)
        nb.add(frm_veiculos, text="VeÃÂÂculos")

        # -------------------------
        # EQUIPES (mantém)
        # -------------------------
        frm_ajudantes = ttk.Frame(nb, style="Content.TFrame")
        crud_ajudantes = CadastroCRUD(
            frm_ajudantes,
            "Ajudantes",
            "ajudantes",
            [
                ("nome", "NOME"),
                ("sobrenome", "SOBRENOME"),
                ("telefone", "TELEFONE"),
                ("status", "STATUS"),
            ],
            app=app
        )
        crud_ajudantes.pack(fill="both", expand=True)
        nb.add(frm_ajudantes, text="Ajudantes")

        # -------------------------
        # CLIENTES (mantém sua ClientesImportPage)
        # -------------------------
        frm_clientes = ttk.Frame(nb, style="Content.TFrame")
        clientes_page = ClientesImportPage(frm_clientes, app=app)
        clientes_page.pack(fill="both", expand=True)
        nb.add(frm_clientes, text="Clientes")

    def on_show(self):
        self.set_status("STATUS: Cadastros (CRUD).")

# ==========================
# ===== FIM DA PARTE 1 =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 2 (ATUALIZADA) =====
# ==========================

class Sidebar(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, style="Sidebar.TFrame", width=260)
        self.app = app
        self.pack_propagate(False)

        self.buttons = {}

        # Topo
        top = ttk.Frame(self, style="Sidebar.TFrame", padding=(14, 16))
        top.pack(fill="x")

        ttk.Label(top, text="ROTAHUB DESKTOP", style="SidebarLogo.TLabel").pack(anchor="w")
        ttk.Label(top, text="Centralizando sua operacao do inicio ao fim.", style="SidebarSmall.TLabel").pack(anchor="w", pady=(2, 0))

        ttk.Separator(self).pack(fill="x", padx=10, pady=(6, 10))

        # âÅâ€œââ‚¬¦ Menu com scroll (evita quebra/sumir item em telas pequenas)
        wrap = ttk.Frame(self, style="Sidebar.TFrame")
        wrap.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(wrap, highlightthickness=0, bd=0, bg="#2B2F8F")
        self.canvas.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        vsb.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=vsb.set)

        self.menu = ttk.Frame(self.canvas, style="Sidebar.TFrame", padding=(10, 6))
        self.menu_id = self.canvas.create_window((0, 0), window=self.menu, anchor="nw")

        def _on_frame_configure(event=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas_configure(event):
            self.canvas.itemconfig(self.menu_id, width=event.width)

        self.menu.bind("<Configure>", _on_frame_configure)
        self.canvas.bind("<Configure>", _on_canvas_configure)

        # Itens do menu
        self._add_btn("home", "ðÅ¸Â  Home", lambda: app.show_page("Home"))
        self._add_btn("cadastros", "ðÅ¸â€œÅ¡ Cadastros", lambda: app.show_page("Cadastros"))
        self._add_btn("rotas", "ðÅ¸â€”ºï¸Â Rotas", lambda: app.show_page("Rotas"))
        self._add_btn("vendas", "ðÅ¸â€œ¥ Importar Vendas", lambda: app.show_page("ImportarVendas"))
        self._add_btn("programacao", "ðÅ¸â€”â€œï¸Â Programacao", lambda: app.show_page("Programacao"))
        self._add_btn("recebimentos", "ðÅ¸â€™µ Recebimentos", lambda: app.show_page("Recebimentos"))
        self._add_btn("despesas", "ðÅ¸â€™¸ Despesas", lambda: app.show_page("Despesas"))
        self._add_btn("escala", "ðÅ¸â€œÅ  Escala", lambda: app.show_page("Escala"))
        self._add_btn("centro_custos", "ðÅ¸Å¡â€º Centro de Custos", lambda: app.show_page("CentroCustos"))
        self._add_btn("relatorios", "ðÅ¸â€œâ€ž Relatorios", lambda: app.show_page("Relatorios"))
        self._add_btn("backup", "ðÅ¸â€”â€žï¸Â Backup / Exportar", lambda: app.show_page("BackupExportar"))

        # Rodapé
        bottom = ttk.Frame(self, style="Sidebar.TFrame", padding=(10, 12))
        bottom.pack(fill="x")

        ttk.Button(bottom, text="âÂ» SAIR", style="Danger.TButton", command=self._safe_quit).pack(fill="x")

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
        self.geometry(f"{APP_W}x{APP_H}")
        self.minsize(1200, 700)

        try:
            self.state("zoomed")
        except Exception:
            logging.debug("Falha ignorada")

        apply_style(self)
        db_init()

        if not self.ensure_integracao_sistema("Inicializacao do sistema", force_probe=True):
            self.after(80, self.destroy)
            return

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.sidebar = Sidebar(self, self)
        self.sidebar.grid(row=0, column=0, sticky="ns")

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
            ok = messagebox.askyesno("Confirmar saÃÂda", "Deseja realmente fechar o sistema?")
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
        }

    def _create_page_if_needed(self, name):
        p = self.pages.get(name)
        if p:
            return p
        factory = self._page_factories.get(name)
        if not factory:
            return None
        p = factory(self.container, self)
        p.grid(row=0, column=0, sticky="nsew")
        self.pages[name] = p
        return p

    def _build_initial_pages(self):
        """Cria apenas páginas essenciais na inicialização para evitar travamento."""
        self._create_page_if_needed("Home")

    def show_page(self, name):
        """Exibe página e atualiza menu lateral"""
        mapping = {
            "Home": "home",
            "Cadastros": "cadastros",
            "Rotas": "rotas",
            "ImportarVendas": "vendas",
            "Programacao": "programacao",
            "Recebimentos": "recebimentos",
            "Despesas": "despesas",
            "Escala": "escala",
            "CentroCustos": "centro_custos",
            "Relatorios": "relatorios",
            "BackupExportar": "backup",
        }

        if name in mapping:
            self.sidebar.set_active(mapping[name])

        page = self._create_page_if_needed(name)
        if not page:
            messagebox.showwarning("ATENÇÃO", f"Página '{name}' não encontrada.")
            return

        page.tkraise()
        self.current_page_name = name
        try:
            page.on_show()
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao abrir página '{name}':\n\n{e}")

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


class HomePage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Home")
        self._api_job = None
        self._update_notice_shown = False
        self._remote_version = "-"
        self._remote_setup_url = SETUP_DOWNLOAD_URL
        self._remote_changelog_url = CHANGELOG_URL
        self._alerts_text = ""

        card = ttk.Frame(self.body, style="Card.TFrame", padding=18)
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=0)
        card.grid_columnconfigure(2, weight=0)

        ttk.Label(card, text="Bem-vindo!", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.lbl_clock = tk.Label(
            card,
            text="--/--/---- --:--:--",
            bg="white",
            fg="#111827",
            font=("Segoe UI", 11, "bold")
        )
        self.lbl_clock.grid(row=0, column=1, sticky="e")
        self._build_system_info_panel(card)

        ttk.Label(
            card,
            text=(
                "âââ€š¬¢ Cadastre Motoristas, VeÃÂÂculos, Equipes e Clientes.\n"
                "âââ€š¬¢ Importe Vendas via Excel.\n"
                "âââ€š¬¢ Gere Programações (códigos automáticos) e vincule pedidos/entregas.\n"
                "âââ€š¬¢ Registre Recebimentos e Despesas.\n"
                "âââ€š¬¢ Emita Relatórios e PDF.\n"
            ),
            style="CardLabel.TLabel",
            justify="left"
        ).grid(row=1, column=0, sticky="w", pady=(6, 0), columnspan=2)

        self.card_stats = ttk.Frame(self.body, style="Content.TFrame")
        self.card_stats.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.card_stats.grid_columnconfigure((0, 1, 2), weight=1)

        self.lbl_total_prog = self._build_stat(self.card_stats, 0, "Programações Ativas", "")
        self.lbl_total_vendas = self._build_stat(self.card_stats, 1, "Vendas Importadas", "")
        self.lbl_total_clientes_ativos = self._build_stat(self.card_stats, 2, "Clientes (Ativos)", "")

        dash = ttk.Frame(self.body, style="Content.TFrame")
        dash.grid(row=2, column=0, sticky="nsew", pady=(14, 0))
        dash.grid_columnconfigure(0, weight=1)
        dash.grid_rowconfigure(0, weight=1)

        card_rotas = ttk.Frame(dash, style="Card.TFrame", padding=12)
        card_rotas.grid(row=0, column=0, sticky="nsew")
        card_rotas.grid_rowconfigure(1, weight=1)
        card_rotas.grid_columnconfigure(0, weight=1)

        ttk.Label(card_rotas, text="Rotas Ativas", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        cols = ["COD", "MOTORISTA", "VEICULO", "DATA"]
        self.tree_rotas = ttk.Treeview(card_rotas, columns=cols, show="headings", height=10)
        self.tree_rotas.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        vsb = ttk.Scrollbar(card_rotas, orient="vertical", command=self.tree_rotas.yview)
        self.tree_rotas.configure(yscrollcommand=vsb.set)
        vsb.grid(row=1, column=1, sticky="ns", pady=(6, 0))

        for c in cols:
            self.tree_rotas.heading(c, text=c)
            self.tree_rotas.column(c, width=160, anchor="w")

        self.tree_rotas.column("COD", width=170)
        self.tree_rotas.column("DATA", width=160, anchor="center")

        self.tree_rotas.bind("<Double-1>", self._open_rota_preview)

        enable_treeview_sorting(
            self.tree_rotas,
            numeric_cols=set(),
            money_cols=set(),
            date_cols={"DATA"}
        )

        self._clock_job = None
        self._update_clock()
        self._build_api_status_badge()
        self._update_api_status()
        self._check_updates(silent=True)
        self.after(300, self._maybe_show_post_update_notifications)

    def _build_system_info_panel(self, parent):
        panel = ttk.Frame(parent, style="Card.TFrame", padding=10)
        panel.grid(row=0, column=2, rowspan=2, sticky="ne", padx=(14, 0))
        panel.grid_columnconfigure(0, weight=1)

        ttk.Label(panel, text="Informações do Sistema", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.lbl_version_local = ttk.Label(
            panel,
            text=f"Versão local: {APP_VERSION}",
            background="white",
            foreground="#111827",
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_version_local.grid(row=1, column=0, sticky="w", pady=(4, 0))

        self.lbl_version_remote = ttk.Label(
            panel,
            text="Versão disponÃÂvel: -",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9),
        )
        self.lbl_version_remote.grid(row=2, column=0, sticky="w")

        actions = ttk.Frame(panel, style="Card.TFrame")
        actions.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        actions.grid_columnconfigure((0, 1), weight=1)

        ttk.Button(actions, text="Histórico", style="Ghost.TButton", command=self._open_changelog).grid(row=0, column=0, sticky="ew")
        ttk.Button(actions, text="Baixar Setup", style="Ghost.TButton", command=self._open_setup).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Button(actions, text="Atualizar versão", style="Primary.TButton", command=self._check_updates).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(actions, text="Publicar versao", style="Ghost.TButton", command=self._publish_version).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        support = ttk.Frame(panel, style="Card.TFrame")
        support.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(support, text="Suporte Técnico", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")

        wpp_txt = f"WhatsApp: {SUPPORT_WHATSAPP}" if SUPPORT_WHATSAPP else "WhatsApp: não definido"
        mail_txt = SUPPORT_EMAIL if SUPPORT_EMAIL else "E-mail: não definido"
        ttk.Label(support, text=wpp_txt, background="white", foreground="#2563EB", font=("Segoe UI", 9, "underline")).grid(row=1, column=0, sticky="w")
        ttk.Label(support, text=mail_txt, background="white", foreground="#111827", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w")

        alerts = ttk.Frame(panel, style="Card.TFrame")
        alerts.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(alerts, text="Alertas", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.lbl_alerts = ttk.Label(
            alerts,
            text="Sem alertas.",
            background="white",
            foreground="#6B7280",
            justify="left",
            wraplength=280,
        )
        self.lbl_alerts.grid(row=1, column=0, sticky="w", pady=(2, 0))

    def _open_url_safe(self, url: str, label: str):
        if not url:
            messagebox.showwarning("Aviso", f"URL de {label} não configurada.")
            return
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao abrir {label}:\n{e}")

    def _open_setup(self):
        url = self._remote_setup_url or SETUP_DOWNLOAD_URL
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

    def _update_now(self):
        url = self._remote_setup_url or SETUP_DOWNLOAD_URL
        if not url:
            messagebox.showwarning(
                "Atualização",
                "Setup não configurado no manifesto.\nUse 'Baixar Setup' com um link válido para o .exe.",
            )
            return

        if not self._can_auto_update_from_url(url):
            messagebox.showinfo(
                "Atualização",
                "Não foi possÃÂvel atualizar automaticamente porque o link atual não aponta para um arquivo .exe.\n\n"
                "Use o botão 'Baixar Setup' para baixar manualmente.",
            )
            self._open_setup()
            return

        try:
            target_v = self._remote_version or "latest"
            self.set_status(f"STATUS: Baixando atualização {target_v}...")
            setup_path = self._download_setup_to_temp(url, target_v)
            self._run_setup_installer(setup_path)
            messagebox.showinfo(
                "Atualização",
                "Instalador aberto com sucesso.\nConclua a instalação para atualizar o Desktop.",
            )
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao atualizar automaticamente:\n{e}")
        finally:
            self.set_status("STATUS: Atualização finalizada (manual/automática).")

    def _open_changelog(self):
        url = self._remote_changelog_url or CHANGELOG_URL
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
        if IS_FROZEN:
            messagebox.showinfo(
                "Publicar versao",
                "Use esta funcao no ambiente de desenvolvimento (python main.py),\n"
                "onde o repositorio Git esta disponivel.",
            )
            return

        updates_dir = os.path.join(APP_DIR, "updates")
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

        current_v = str(manifest.get("version") or self._remote_version or APP_VERSION).strip()
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
            manifest.get("setup_url") or SETUP_DOWNLOAD_URL or "https://github.com/andersoncantenas-glitch/rotahub-api/releases/latest"
        ).strip()
        changelog_url = str(
            manifest.get("changelog_url")
            or CHANGELOG_URL
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

        local_v = self._version_tuple(APP_VERSION)
        remote_v = self._version_tuple(remote_version)
        is_candidate = (remote_v > local_v) and self._is_same_release_line(local_v, remote_v)
        remote_display = remote_version if is_candidate else "0.0.0"

        self.lbl_version_remote.config(text=f"Versão disponÃÂvel: {remote_display}")
        self.lbl_alerts.config(text=alert_txt or "Sem alertas.")

        if is_candidate:
            self.lbl_version_remote.config(foreground="#B45309")
        else:
            self.lbl_version_remote.config(foreground="#166534")

    def _check_updates(self, silent=False):
        if not UPDATE_MANIFEST_URL:
            if not silent:
                messagebox.showinfo(
                    "Atualizações",
                    "Manifesto de atualização não configurado.\n"
                    "Defina ROTA_UPDATE_MANIFEST_URL no ambiente para habilitar checagem automática.",
                )
            return

        try:
            req = urllib.request.Request(UPDATE_MANIFEST_URL, method="GET")
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = resp.read().decode("utf-8", errors="replace")
            manifest = json.loads(data)
            if not isinstance(manifest, dict):
                raise ValueError("Manifesto inválido")
            self._apply_manifest(manifest)

            if not silent:
                local_v = self._version_tuple(APP_VERSION)
                remote_v = self._version_tuple(self._remote_version)
                is_candidate = (remote_v > local_v) and self._is_same_release_line(local_v, remote_v)
                if is_candidate:
                    ask = messagebox.askyesno(
                        "Atualização disponÃÂvel",
                        f"Versão local: {APP_VERSION}\nVersão disponÃÂvel: {self._remote_version}\n\n"
                        "Deseja atualizar agora?",
                    )
                    if ask:
                        self._update_now()
                else:
                    messagebox.showinfo("Atualizações", f"Você já está na versão mais recente ({APP_VERSION}).")
        except Exception as e:
            if not silent:
                messagebox.showerror("ERRO", f"Falha ao verificar atualizações:\n{e}")

    def _update_state_path(self):
        base_dir = USER_DATA_DIR if IS_FROZEN else APP_DIR
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
            os.path.join(APP_DIR, "updates", "changelog.txt"),
            os.path.join(RESOURCE_DIR, "updates", "changelog.txt"),
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
        url = self._remote_changelog_url or CHANGELOG_URL
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

        return "Sem detalhes de changelog disponÃÂveis."

    def _maybe_show_post_update_notifications(self):
        if self._update_notice_shown:
            return
        self._update_notice_shown = True

        st = self._load_update_state()
        last_seen = str(st.get("last_seen_version") or "").strip()

        if self._version_tuple(APP_VERSION) <= self._version_tuple(last_seen):
            return

        alert_block = (self._alerts_text or "").strip()
        changelog = self._changelog_preview()
        notes = self._extract_latest_notes(changelog, APP_VERSION)
        lines = [
            f"Versão atual: {APP_VERSION}",
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
        apply_window_icon(win)
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
            st["last_seen_version"] = APP_VERSION
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
        online = False
        api_host = "-"
        try:
            url = _build_api_url("openapi.json")
            try:
                api_host = urllib.parse.urlparse(url).netloc or url
            except Exception:
                api_host = url
            req = urllib.request.Request(url, method="GET")
            with urllib.request.urlopen(req, timeout=min(max(API_SYNC_TIMEOUT, 3.0), 10.0)) as resp:
                online = int(getattr(resp, "status", 0) or 0) == 200
        except Exception:
            online = False
            if api_host == "-":
                try:
                    api_host = urllib.parse.urlparse(API_BASE_URL).netloc or API_BASE_URL
                except Exception:
                    api_host = API_BASE_URL

        if online:
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
        if not desktop_secret or not is_desktop_api_sync_enabled():
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
        if not desktop_secret or not is_desktop_api_sync_enabled():
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

    def on_show(self):
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
            for i in self.tree_rotas.get_children():
                self.tree_rotas.delete(i)
            for r in rows:
                if isinstance(r, dict):
                    codigo = upper(r.get("codigo_programacao") or "")
                    motorista = upper(r.get("motorista") or "")
                    veiculo = upper(r.get("veiculo") or "")
                    data_criacao = str(r.get("data_criacao") or "")
                else:
                    codigo = upper(r[0] if len(r) > 0 else "")
                    motorista = upper(r[1] if len(r) > 1 else "")
                    veiculo = upper(r[2] if len(r) > 2 else "")
                    data_criacao = str(r[3] if len(r) > 3 else "")
                data_fmt = self._format_home_data(data_criacao)
                self.tree_rotas.insert("", "end", values=(codigo, motorista, veiculo, data_fmt))
            self.lbl_total_prog.config(text=str(total_prog))
            self.lbl_total_vendas.config(text=str(total_vendas))
            self.lbl_total_clientes_ativos.config(text=str(total_clientes_ativos))
            self.set_status(f"STATUS: Home carregada ({source}).")
            return

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = {str(r[1]).lower() for r in (cur.fetchall() or [])}
            except Exception:
                cols_prog = set()
            where_ativas = self._home_local_not_finalized_where(cols_prog)

            try:
                cur.execute("""
                    SELECT COUNT(*)
                    FROM programacoes
                    WHERE """ + where_ativas)
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

            for i in self.tree_rotas.get_children():
                self.tree_rotas.delete(i)

            api_rows = self._home_api_rows()
            if api_rows:
                rows = api_rows
                total_prog = len(api_rows)
                source = "API CENTRAL"
            else:
                try:
                    cur.execute("""
                        SELECT codigo_programacao, motorista, veiculo, data_criacao
                        FROM programacoes
                        WHERE """ + where_ativas + """
                        ORDER BY id DESC
                        LIMIT 80
                    """)
                    rows = cur.fetchall() or []
                except Exception:
                    try:
                        cur.execute("""
                            SELECT codigo_programacao, motorista, veiculo, data_criacao
                            FROM programacoes
                            ORDER BY id DESC
                            LIMIT 80
                        """)
                        rows = cur.fetchall() or []
                    except Exception:
                        rows = []

            for r in rows:
                tree_insert_aligned(self.tree_rotas, "", "end", (r[0], r[1], r[2], r[3]))

        self.lbl_total_prog.config(text=str(total_prog))
        self.lbl_total_vendas.config(text=str(total_vendas))
        self.lbl_total_clientes_ativos.config(text=str(total_clientes_ativos))

        nome = self.app.user.get("nome", "")
        is_admin = bool(self.app.user.get("is_admin", False))
        self.set_status(f"STATUS: Logado como {nome} (ADMIN: {is_admin}). Fonte Home: {source}.")

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
            loc = {"latitude": "", "longitude": "", "endereco": "", "cidade": "", "bairro": ""}
            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cols = _table_cols(cur, "clientes")
                    if not cols:
                        return loc
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
                        for k in loc.keys():
                            loc[k] = str(rr.get(k) or "")
            except Exception:
                logging.debug("Falha ao coletar localização do cliente", exc_info=True)
            return loc

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

            linhas = []
            linhas.append(f"PROGRAMAÇÃO: {cabecalho_ctx.get('codigo') or codigo}")
            linhas.append(f"NOTA FISCAL: {cabecalho_ctx.get('nf_numero') or '-'}")
            linhas.append(
                f"MOTORISTA: {cabecalho_ctx.get('motorista') or '-'}    VEÃÂCULO: {cabecalho_ctx.get('veiculo') or '-'}"
            )
            eq_raw = cabecalho_ctx.get("equipe") or ""
            linhas.append(f"EQUIPE: {resolve_equipe_nomes(eq_raw) or eq_raw or '-'}")
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
            linhas.append("[LOCALIZAÇÃO CLIENTE]")
            linhas.append(f"LATITUDE: {local.get('latitude') or '-'}")
            linhas.append(f"LONGITUDE: {local.get('longitude') or '-'}")
            linhas.append(f"ENDEREÇO: {local.get('endereco') or '-'}")
            linhas.append(f"CIDADE: {local.get('cidade') or '-'}    BAIRRO: {local.get('bairro') or '-'}")
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
            st_entregue = {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÃÂDO"}
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
            ttk.Label(head, text=f"VeÃÂculo: {cabecalho_ctx.get('veiculo') or '-'}").grid(row=1, column=3, sticky="w")

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
            grp_geo.grid(row=1, column=0, columnspan=2, sticky="nsew")

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
                f"Equipe: {resolve_equipe_nomes(cabecalho_ctx.get('equipe') or '') or (cabecalho_ctx.get('equipe') or '-')}",
                f"Hora da entrega/alteração: {alterado_em or '-'}",
                f"Alterado por: {alterado_por or '-'}",
                f"Tipo de alteração: {alteracao_tipo or '-'}",
                f"Detalhe da alteração: {alteracao_detalhe or '-'}",
                f"Observação recebimento: {obs_receb or '-'}",
                f"Transferências encontradas: {len(transferencias)}",
                f"Ocorrências/logs: {len(logs_item)}",
            ]
            for i, t in enumerate(l_rst):
                ttk.Label(grp_rst, text=t).grid(row=i, column=0, sticky="w", pady=1)

            ttk.Label(grp_geo, text=f"Latitude: {local.get('latitude') or '-'}").grid(row=0, column=0, sticky="w", padx=(0, 20))
            ttk.Label(grp_geo, text=f"Longitude: {local.get('longitude') or '-'}").grid(row=0, column=1, sticky="w")
            ttk.Label(grp_geo, text=f"Endereço: {local.get('endereco') or '-'}").grid(row=1, column=0, columnspan=2, sticky="w")
            ttk.Label(grp_geo, text=f"Cidade: {local.get('cidade') or '-'}").grid(row=2, column=0, sticky="w")
            ttk.Label(grp_geo, text=f"Bairro: {local.get('bairro') or '-'}").grid(row=2, column=1, sticky="w")

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
                    messagebox.showwarning("Mapa", "Latitude/Longitude não disponÃÂveis para este cliente.")
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
                        return db_has_column(cur, table, col)
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

            # 1.2) opcional: fonte central (API) para refletir status/carregamento do app
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
                    itens_local = fetch_programacao_itens(codigo) or []
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
                if st_norm in {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÃÂÂDO"}:
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
                }

            # 3) labels
            lbl_status.config(text=f"Status: {status or ''}")
            lbl_motorista.config(text=f"Motorista: {motorista or ''}")
            equipe_txt = resolve_equipe_nomes(equipe)
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
        ttk.Button(btns, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton", command=_load_preview).pack(side="right")


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

        ttk.Button(self.footer_right, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton", command=self.carregar).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(self.footer_right, text="ðÅ¸â€”º MAPA SELECIONADO", style="Primary.TButton", command=self._abrir_mapa_selecionado).grid(
            row=0, column=1, padx=4
        )
        ttk.Button(self.footer_right, text="ðÅ¸Å’Â MAPA DE TODAS", style="Primary.TButton", command=self._abrir_mapa_todas).grid(
            row=0, column=2, padx=4
        )

        self._rows_cache = []
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
        if desktop_secret and is_desktop_api_sync_enabled():
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

    def carregar(self):
        self._rows_cache = self._fetch_rows()
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

import re

# =========================================================
# ORDENAÇÃO UNIVERSAL PARA TREEVIEW (CLIQUE NO CABEÇALHO)
# =========================================================
def enable_treeview_sorting(tree: ttk.Treeview, numeric_cols=None, money_cols=None, date_cols=None):
    """
    Habilita recursos de tabela no Treeview:
    - Ordenacao por cabecalho
    - Indicador de filtro ao passar mouse no cabe?alho
    - Filtro por coluna (valores especificos)

    Uso do filtro:
    - Passe o mouse no cabe?alho para ver o indicador "\u23F7"
    - Clique no indicador (lado direito do cabecalho) ou botao direito no cabecalho
    """
    numeric_cols = set(numeric_cols or [])
    money_cols = set(money_cols or [])
    date_cols = set(date_cols or [])

    if not hasattr(tree, "_sort_state"):
        tree._sort_state = {"col": None, "reverse": False}

    if not hasattr(tree, "_filter_state"):
        tree._filter_state = {}

    if not hasattr(tree, "_filter_all_iids"):
        tree._filter_all_iids = list(tree.get_children(""))

    if not hasattr(tree, "_base_heading_text"):
        tree._base_heading_text = {
            c: (tree.heading(c, "text") or c)
            for c in tree["columns"]
        }

    tree._hover_filter_col = None

    def _clean_header_text(t: str) -> str:
        t = (t or "").strip()
        for suffix in (" \u2191", " \u2193", " \u23F7", " \u23F7*"):
            if t.endswith(suffix):
                t = t[: -len(suffix)].strip()
        return t

    def _to_float(v):
        try:
            if v is None:
                return 0.0
            s = str(v).strip()
            if not s or s in {"", "-", "None"}:
                return 0.0

            s = s.replace("R$", "").replace(" ", "")

            neg = False
            if s.startswith("(") and s.endswith(")"):
                neg = True
                s = s[1:-1].strip()

            s = s.replace(".", "").replace(",", ".")
            val = float(s)
            return -val if neg else val
        except Exception:
            return 0.0

    def _to_date_key(v):
        if v is None:
            return (0, 0, 0)

        s = str(v).strip()
        if not s or s in {"", "-", "None"}:
            return (0, 0, 0)

        if "-" in s and len(s) >= 10:
            try:
                y, m, d = s[:10].split("-")
                return (int(y), int(m), int(d))
            except Exception:
                logging.debug("Falha ignorada")

        if "/" in s and len(s) >= 10:
            try:
                d, m, y = s[:10].split("/")
                return (int(y), int(m), int(d))
            except Exception:
                logging.debug("Falha ignorada")

        return (0, 0, 0)

    def _format_header(col):
        base = _clean_header_text(tree._base_heading_text.get(col) or tree.heading(col, "text") or col)
        suffix = ""

        if tree._sort_state.get("col") == col:
            suffix += " \u2193" if tree._sort_state.get("reverse") else " \u2191"

        if tree._hover_filter_col == col:
            suffix += " \u23F7*" if col in tree._filter_state else " \u23F7"
        elif col in tree._filter_state:
            suffix += " \u23F7*"

        return f"{base}{suffix}"

    def _refresh_headers():
        for c in tree["columns"]:
            tree.heading(c, text=_format_header(c))

    def _value_for_compare(v):
        return str(v or "").strip()

    def _row_matches_filters(iid):
        for col, allowed in tree._filter_state.items():
            current = _value_for_compare(tree.set(iid, col))
            if current not in allowed:
                return False
        return True

    def _apply_filters():
        all_iids = [iid for iid in tree._filter_all_iids if tree.exists(iid)]
        if not tree._filter_state:
            for idx, iid in enumerate(all_iids):
                tree.reattach(iid, "", idx)
            _refresh_headers()
            return

        pos = 0
        for iid in all_iids:
            if _row_matches_filters(iid):
                tree.reattach(iid, "", pos)
                pos += 1
            else:
                tree.detach(iid)
        _refresh_headers()

    def _iter_values_for_col(col):
        vals = []
        for iid in tree._filter_all_iids:
            if tree.exists(iid):
                vals.append(_value_for_compare(tree.set(iid, col)))
        return vals

    def _open_filter_popup(col):
        top = tk.Toplevel(tree)
        top.title(f"Filtrar: {col}")
        top.transient(tree.winfo_toplevel())
        top.resizable(False, False)
        top.grab_set()

        frm = ttk.Frame(top, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")
        frm.grid_columnconfigure(0, weight=1)

        ttk.Label(frm, text=f"Coluna: {col}", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")

        search_var = tk.StringVar(value="")
        ent = ttk.Entry(frm, textvariable=search_var)
        ent.grid(row=1, column=0, sticky="ew", pady=(6, 6))

        listbox = tk.Listbox(frm, selectmode="multiple", exportselection=False, height=12)
        listbox.grid(row=2, column=0, sticky="nsew")

        unique_vals = sorted(set(_iter_values_for_col(col)), key=lambda s: s.upper())
        current_allowed = set(tree._filter_state.get(col, set()))

        def _fill_list():
            needle = search_var.get().strip().upper()
            listbox.delete(0, "end")
            for v in unique_vals:
                show = "(vazio)" if v == "" else v
                if needle and needle not in show.upper():
                    continue
                listbox.insert("end", show)
                if (v in current_allowed) or (not current_allowed):
                    listbox.selection_set("end")

        def _selected_values():
            selected = set()
            for idx in listbox.curselection():
                txt = listbox.get(idx)
                selected.add("" if txt == "(vazio)" else txt)
            return selected

        def _apply_and_close():
            vals = _selected_values()
            if vals and len(vals) < len(unique_vals):
                tree._filter_state[col] = vals
            else:
                tree._filter_state.pop(col, None)
            _apply_filters()
            top.destroy()

        def _clear_col_filter():
            tree._filter_state.pop(col, None)
            _apply_filters()
            top.destroy()

        def _clear_all_filters():
            tree._filter_state.clear()
            _apply_filters()
            top.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="âœ” Aplicar", style="Primary.TButton", command=_apply_and_close).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(btns, text="ðÅ¸§¹ Limpar coluna", style="Ghost.TButton", command=_clear_col_filter).grid(row=0, column=1, padx=6)
        ttk.Button(btns, text="ðÅ¸§¹ LIMPAR TUDO", style="Ghost.TButton", command=_clear_all_filters).grid(row=0, column=2, padx=6)

        search_var.trace_add("write", lambda *_: _fill_list())
        _fill_list()
        ent.focus_set()

    def sort_by(col):
        if tree._sort_state["col"] == col:
            tree._sort_state["reverse"] = not tree._sort_state["reverse"]
        else:
            tree._sort_state["col"] = col
            tree._sort_state["reverse"] = False

        reverse = tree._sort_state["reverse"]
        data = [(tree.set(iid, col), iid) for iid in tree.get_children("")]

        if col in money_cols or col in numeric_cols:
            data.sort(key=lambda x: _to_float(x[0]), reverse=reverse)
        elif col in date_cols:
            data.sort(key=lambda x: _to_date_key(x[0]), reverse=reverse)
        else:
            data.sort(key=lambda x: str(x[0]).strip().upper(), reverse=reverse)

        for idx, (_, iid) in enumerate(data):
            tree.move(iid, "", idx)

        _refresh_headers()

    def _column_right_edge(col_name):
        x = 0
        for c in tree["columns"]:
            w = int(tree.column(c, "width") or 0)
            x += w
            if c == col_name:
                return x
        return None

    def _on_motion(event):
        region = tree.identify("region", event.x, event.y)
        col_id = tree.identify_column(event.x)
        hover_col = None
        if region == "heading" and col_id and col_id.startswith("#"):
            try:
                idx = int(col_id[1:]) - 1
                if 0 <= idx < len(tree["columns"]):
                    hover_col = tree["columns"][idx]
            except Exception:
                hover_col = None

        if tree._hover_filter_col != hover_col:
            tree._hover_filter_col = hover_col
            _refresh_headers()

    def _on_leave(_event):
        if tree._hover_filter_col is not None:
            tree._hover_filter_col = None
            _refresh_headers()

    def _on_right_click(event):
        region = tree.identify("region", event.x, event.y)
        if region != "heading":
            return
        col_id = tree.identify_column(event.x)
        if not col_id.startswith("#"):
            return
        idx = int(col_id[1:]) - 1
        if idx < 0 or idx >= len(tree["columns"]):
            return
        col = tree["columns"][idx]
        _open_filter_popup(col)
        return "break"

    def _on_left_click(event):
        region = tree.identify("region", event.x, event.y)
        if region != "heading":
            return
        col_id = tree.identify_column(event.x)
        if not col_id.startswith("#"):
            return
        idx = int(col_id[1:]) - 1
        if idx < 0 or idx >= len(tree["columns"]):
            return
        col = tree["columns"][idx]

        col_right = _column_right_edge(col)
        if col_right is None:
            return

        if event.x >= (col_right - 20):
            _open_filter_popup(col)
            return "break"

    if not hasattr(tree, "_filter_wrapped_io"):
        tree._filter_wrapped_io = True
        _orig_insert = tree.insert
        _orig_delete = tree.delete

        def _insert_wrapped(parent, index, iid=None, **kw):
            new_iid = _orig_insert(parent, index, iid=iid, **kw)
            if parent == "" and new_iid not in tree._filter_all_iids:
                tree._filter_all_iids.append(new_iid)
            return new_iid

        def _delete_wrapped(*items):
            for iid in items:
                try:
                    tree._filter_all_iids.remove(iid)
                except Exception:
                    pass
            return _orig_delete(*items)

        tree.insert = _insert_wrapped
        tree.delete = _delete_wrapped

    for c in tree["columns"]:
        tree.heading(c, command=lambda col=c: sort_by(col))

    tree.bind("<Motion>", _on_motion, add="+")
    tree.bind("<Leave>", _on_leave, add="+")
    tree.bind("<Button-3>", _on_right_click, add="+")
    tree.bind("<Button-1>", _on_left_click, add="+")

    _refresh_headers()

_CPF_DIGITS = re.compile(r"\D+")
_PHONE_DIGITS = re.compile(r"\D+")
_MOTORISTA_COD_RE = re.compile(r"^[A-Z0-9._-]{3,24}$")  # já normaliza em upper()

def normalize_cpf(v: str) -> str:
    return _CPF_DIGITS.sub("", str(v or "").strip())

def is_valid_cpf(cpf_digits: str) -> bool:
    """
    Validador de CPF (Brasil).
    Aceita apenas 11 dÃÂÂgitos e verifica dÃÂÂgitos verificadores.
    """
    cpf = normalize_cpf(cpf_digits)
    if len(cpf) != 11:
        return False
    if cpf == cpf[0] * 11:
        return False

    try:
        nums = [int(x) for x in cpf]

        # 1º dÃÂÂgito
        s1 = sum(nums[i] * (10 - i) for i in range(9))
        d1 = (s1 * 10) % 11
        d1 = 0 if d1 == 10 else d1
        if d1 != nums[9]:
            return False

        # 2º dÃÂÂgito
        s2 = sum(nums[i] * (11 - i) for i in range(10))
        d2 = (s2 * 10) % 11
        d2 = 0 if d2 == 10 else d2
        if d2 != nums[10]:
            return False

        return True
    except Exception:
        return False

def normalize_phone(v: str) -> str:
    """
    Normaliza telefone para dÃÂÂgitos.
    Se vier com DDI 55 (Brasil), remove quando fizer sentido.
    """
    s = _PHONE_DIGITS.sub("", str(v or "").strip())
    # remove 55 se vier com DDI e sobrar 10/11 dÃÂÂgitos
    if len(s) in (12, 13) and s.startswith("55"):
        s2 = s[2:]
        if len(s2) in (10, 11):
            return s2
    return s

def is_valid_phone(phone_digits: str) -> bool:
    tel = normalize_phone(phone_digits)
    return len(tel) in (10, 11) and tel.isdigit()

def is_valid_motorista_codigo(cod: str) -> bool:
    cod = upper(str(cod or "").strip())
    return bool(_MOTORISTA_COD_RE.match(cod))

def is_valid_motorista_senha(senha: str) -> bool:
    senha = str(senha or "").strip()
    return 4 <= len(senha) <= 24


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
    try:
        table_name = _safe_ident(table_name)
        column_name = str(column_name or "").strip().lower()
        cur.execute(f"PRAGMA table_info({table_name})")
        cols = [str(r[1]).lower() for r in cur.fetchall()]
        return column_name in cols
    except Exception:
        return False


def fetch_programacoes_ativas(limit: int = 400):
    """
    Retorna lista de programações ATIVAS (compatÃÂÂvel com bases antigas e novas).
    SaÃÂÂda: lista de dicts: {codigo, motorista, veiculo, equipe, data_criacao, status}
    """
    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if desktop_secret and is_desktop_api_sync_enabled():
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

            has_status = db_has_column(cur, "programacoes", "status")
            has_data = db_has_column(cur, "programacoes", "data_criacao")

            if has_status:
                if has_data:
                    cur.execute("""
                        SELECT codigo_programacao, motorista, veiculo, equipe, data_criacao, status
                        FROM programacoes
                        WHERE status='ATIVA'
                        ORDER BY id DESC
                        LIMIT ?
                    """, (limit,))
                else:
                    cur.execute("""
                        SELECT codigo_programacao, motorista, veiculo, equipe, '' as data_criacao, status
                        FROM programacoes
                        WHERE status='ATIVA'
                        ORDER BY id DESC
                        LIMIT ?
                    """, (limit,))
            else:
                # base antiga sem status: pega as últimas como "ativas"
                if has_data:
                    cur.execute("""
                        SELECT codigo_programacao, motorista, veiculo, equipe, data_criacao, 'ATIVA' as status
                        FROM programacoes
                        ORDER BY id DESC
                        LIMIT ?
                    """, (limit,))
                else:
                    cur.execute("""
                        SELECT codigo_programacao, motorista, veiculo, equipe, '' as data_criacao, 'ATIVA' as status
                        FROM programacoes
                        ORDER BY id DESC
                        LIMIT ?
                    """, (limit,))

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


def fetch_programacao_itens(codigo_programacao: str, limit: int = 8000):
    """
    Retorna itens de uma programacao.
    Saida: lista de dicts:
      {cod_cliente, nome_cliente, produto, endereco, qnt_caixas, kg, preco, vendedor, pedido, obs,
       status_pedido, caixas_atual, preco_atual, alterado_em, alterado_por,
       mortalidade_aves, valor_recebido, forma_recebimento, obs_recebimento}
    Compativel com schemas diferentes.
    """
    codigo_programacao = upper(str(codigo_programacao or "").strip())
    if not codigo_programacao:
        return []

    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
    if desktop_secret and is_desktop_api_sync_enabled():
        try:
            resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(codigo_programacao)}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            clientes = resp.get("clientes") if isinstance(resp, dict) else []
            out_api = []
            for d in (clientes or [])[: max(int(limit or 0), 1)]:
                if not isinstance(d, dict):
                    continue
                out_api.append(
                    {
                        "cod_cliente": str(d.get("cod_cliente") or "").strip().upper(),
                        "nome_cliente": str(d.get("nome_cliente") or "").strip().upper(),
                        "endereco": str(d.get("endereco") or "").strip().upper(),
                        "produto": str(d.get("produto") or "").strip().upper(),
                        "qnt_caixas": safe_int(d.get("qnt_caixas"), 0),
                        "kg": safe_float(d.get("kg"), 0.0),
                        "preco": safe_float(d.get("preco"), 0.0),
                        "vendedor": str(d.get("vendedor") or "").strip().upper(),
                        "pedido": str(d.get("pedido") or "").strip().upper(),
                        "obs": str(d.get("obs") or d.get("observacao") or "").strip(),
                        "status_pedido": str(d.get("status_pedido") or "PENDENTE").strip().upper(),
                        "caixas_atual": safe_int(d.get("caixas_atual"), 0),
                        "preco_atual": safe_float(d.get("preco_atual"), 0.0),
                        "alterado_em": str(d.get("alterado_em") or "").strip(),
                        "alterado_por": str(d.get("alterado_por") or "").strip().upper(),
                        "mortalidade_aves": safe_int(d.get("mortalidade_aves"), 0),
                        "peso_previsto": safe_float(d.get("peso_previsto"), 0.0),
                        "valor_recebido": safe_float(d.get("valor_recebido"), 0.0),
                        "forma_recebimento": str(d.get("forma_recebimento") or "").strip().upper(),
                        "obs_recebimento": str(d.get("obs_recebimento") or "").strip(),
                        "alteracao_tipo": str(d.get("alteracao_tipo") or "").strip().upper(),
                        "alteracao_detalhe": str(d.get("alteracao_detalhe") or "").strip(),
                    }
                )
            if out_api:
                return out_api
        except Exception:
            logging.debug("Falha ao buscar itens da programacao na API", exc_info=True)

    try:
        with get_db() as conn:
            cur = conn.cursor()

            has_obs = db_has_column(cur, "programacao_itens", "obs") or db_has_column(cur, "programacao_itens", "observacao")
            has_vendedor = db_has_column(cur, "programacao_itens", "vendedor")
            has_pedido = db_has_column(cur, "programacao_itens", "pedido")
            has_end = db_has_column(cur, "programacao_itens", "endereco") or db_has_column(cur, "programacao_itens", "endereco")
            has_produto = db_has_column(cur, "programacao_itens", "produto")

            has_status = db_has_column(cur, "programacao_itens", "status_pedido")
            has_caixas_atual = db_has_column(cur, "programacao_itens", "caixas_atual")
            has_preco_atual = db_has_column(cur, "programacao_itens", "preco_atual")
            has_alt_em = db_has_column(cur, "programacao_itens", "alterado_em")
            has_alt_por = db_has_column(cur, "programacao_itens", "alterado_por")

            # controle
            has_ctrl = db_has_column(cur, "programacao_itens_controle", "codigo_programacao")
            has_ctrl_pedido = has_ctrl and db_has_column(cur, "programacao_itens_controle", "pedido")
            has_ctrl_alt_tipo = has_ctrl and db_has_column(cur, "programacao_itens_controle", "alteracao_tipo")
            has_ctrl_alt_detalhe = has_ctrl and db_has_column(cur, "programacao_itens_controle", "alteracao_detalhe")

            status_expr = ("pi.status_pedido" if has_status else "''")
            caixas_atual_expr = ("pi.caixas_atual" if has_caixas_atual else "0")
            preco_atual_expr = ("pi.preco_atual" if has_preco_atual else "0")
            alterado_em_expr = ("pi.alterado_em" if has_alt_em else "''")
            alterado_por_expr = ("pi.alterado_por" if has_alt_por else "''")

            if has_ctrl:
                status_expr = "COALESCE(NULLIF(" + status_expr + ", ''), NULLIF(pc.status_pedido, ''), 'PENDENTE')"
                caixas_atual_expr = (
                    "CASE "
                    "WHEN COALESCE(pc.caixas_atual, 0) > 0 THEN pc.caixas_atual "
                    "WHEN COALESCE(" + caixas_atual_expr + ", 0) > 0 THEN " + caixas_atual_expr + " "
                    "ELSE COALESCE(pi.qnt_caixas, 0) "
                    "END"
                )
                preco_atual_expr = "COALESCE(" + preco_atual_expr + ", pc.preco_atual, 0)"
                alterado_em_expr = "COALESCE(" + alterado_em_expr + ", pc.alterado_em, '')"
                alterado_por_expr = "COALESCE(" + alterado_por_expr + ", pc.alterado_por, '')"
            else:
                status_expr = "COALESCE(NULLIF(" + status_expr + ", ''), 'PENDENTE')"

            select_cols = [
                "pi.cod_cliente",
                "pi.nome_cliente",
                ("pi.endereco" if has_end else "'' as endereco"),
                ("pi.produto" if has_produto else "'' as produto"),
                "pi.qnt_caixas",
                "pi.kg",
                "pi.preco",
                ("pi.vendedor" if has_vendedor else "'' as vendedor"),
                ("pi.pedido" if has_pedido else "'' as pedido"),
                ("pi.obs" if db_has_column(cur, "programacao_itens", "obs") else ("pi.observacao" if db_has_column(cur, "programacao_itens", "observacao") else "'' as obs")),
                status_expr + " as status_pedido",
                caixas_atual_expr + " as caixas_atual",
                preco_atual_expr + " as preco_atual",
                alterado_em_expr + " as alterado_em",
                alterado_por_expr + " as alterado_por",
            ]

            join_ctrl = ""
            if has_ctrl:
                select_cols.extend([
                    "COALESCE(pc.mortalidade_aves, 0) as mortalidade_aves",
                    "COALESCE(pc.peso_previsto, 0) as peso_previsto",
                    "COALESCE(pc.valor_recebido, 0) as valor_recebido",
                    "COALESCE(pc.forma_recebimento, '') as forma_recebimento",
                    "COALESCE(pc.obs_recebimento, '') as obs_recebimento",
                    ("COALESCE(pc.alteracao_tipo, '') as alteracao_tipo" if has_ctrl_alt_tipo else "'' as alteracao_tipo"),
                    ("COALESCE(pc.alteracao_detalhe, '') as alteracao_detalhe" if has_ctrl_alt_detalhe else "'' as alteracao_detalhe"),
                ])
                join_ctrl = (
                    "LEFT JOIN programacao_itens_controle pc "
                    "ON pc.codigo_programacao = pi.codigo_programacao "
                    "AND UPPER(pc.cod_cliente)=UPPER(pi.cod_cliente)"
                )
                if has_pedido and has_ctrl_pedido:
                    join_ctrl += " AND COALESCE(TRIM(pc.pedido),'') = COALESCE(TRIM(pi.pedido),'')"
            else:
                select_cols.extend([
                    "0 as mortalidade_aves",
                    "0 as peso_previsto",
                    "0 as valor_recebido",
                    "'' as forma_recebimento",
                    "'' as obs_recebimento",
                    "'' as alteracao_tipo",
                    "'' as alteracao_detalhe",
                ])

            cur.execute(f"""
                SELECT {", ".join(select_cols)}
                FROM programacao_itens pi
                {join_ctrl}
                WHERE pi.codigo_programacao=?
                ORDER BY pi.id ASC
                LIMIT ?
            """, (codigo_programacao, limit))

            rows = cur.fetchall()

        out = []
        for r in rows:
            out.append({
                "cod_cliente": r[0] or "",
                "nome_cliente": r[1] or "",
                "endereco": r[2] or "",
                "produto": r[3] or "",
                "qnt_caixas": safe_int(r[4], 0),
                "kg": safe_float(r[5], 0.0),
                "preco": safe_float(r[6], 0.0),
                "vendedor": r[7] or "",
                "pedido": r[8] or "",
                "obs": (r[9] or "") if has_obs else (r[9] or ""),
                "status_pedido": r[10] or "",
                "caixas_atual": safe_int(r[11], 0),
                "preco_atual": safe_float(r[12], 0.0),
                "alterado_em": r[13] or "",
                "alterado_por": r[14] or "",
                "mortalidade_aves": safe_int(r[15], 0),
                "peso_previsto": safe_float(r[16], 0.0),
                "valor_recebido": safe_float(r[17], 0.0),
                "forma_recebimento": r[18] or "",
                "obs_recebimento": r[19] or "",
                "alteracao_tipo": r[20] or "",
                "alteracao_detalhe": r[21] or "",
            })
        return out

    except Exception:
        return []


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

# Compatibilidade: o fluxo "Programacao" usa a mesma tela de "Rotas".
ProgramacaoPage = RotasPage

# ==========================
# ===== INCIO DA PARTE 3 (FINAL / SEM DUPLICIDADE) =====
# ==========================

# =========================================================
# 3.0.1 CLIENTES (IMPORTAÇÃO + EDIÇÃO DIRETA NA TABELA)
# =========================================================
class ClientesImportPage(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, style="Content.TFrame")
        self.app = app
        self._editing = None

        card = ttk.Frame(self, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        ttk.Label(card, text="Clientes (Base do Wibi)", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        actions.grid_columnconfigure(10, weight=1)

        ttk.Button(actions, text="ðÅ¸â€œ¥ IMPORTAR CLIENTES (EXCEL)", style="Warn.TButton",
                   command=self.importar_clientes_excel).grid(row=0, column=0, padx=6)

        ttk.Button(actions, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton",
                   command=self.carregar).grid(row=0, column=1, padx=6)

        ttk.Button(actions, text="âÅ¾â€¢ INSERIR LINHA", style="Ghost.TButton",
                   command=self.inserir_linha).grid(row=0, column=2, padx=6)

        ttk.Button(actions, text="SALVAR ALTERAÇÕES", style="Primary.TButton",
                   command=self.salvar_alteracoes).grid(row=0, column=3, padx=6)

        self.lbl_info = ttk.Label(
            actions,
            text="Dica: duplo clique na célula para editar. ENTER salva a célula. ESC cancela.",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_info.grid(row=0, column=4, padx=12, sticky="w")

        cols = ["CÓD CLIENTE", "NOME CLIENTE", "ENDEREÇO", "TELEFONE", "VENDEDOR"]
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
            self.tree.column(c, width=190, minwidth=120, anchor="w")

        self.tree.column("ENDEREÇO", width=420, minwidth=200)
        self.tree.bind("<Double-1>", self._start_edit_cell)

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"CÓD CLIENTE"},
            money_cols=set(),
            date_cols=set()
        )

        # Se rolar o tree durante edição, fecha/commita corretamente
        self.tree.bind("<MouseWheel>", self._on_tree_scroll, add=True)
        self.tree.bind("<Button-4>", self._on_tree_scroll, add=True)  # linux
        self.tree.bind("<Button-5>", self._on_tree_scroll, add=True)  # linux

        self.carregar()

    # -------------------------
    # Helpers (reduz duplicidade)
    # -------------------------
    def _get_row_values(self, iid):
        vals = self.tree.item(iid, "values") or ("", "", "", "", "")
        vals = list(vals) + [""] * (5 - len(vals))
        return [str(v or "").strip() for v in vals[:5]]

    def _is_blank_row(self, cod, nome, endereco, telefone, vendedor):
        return not (cod or nome or endereco or telefone or vendedor)

    def _normalize_cod(self, cod):
        return str(cod or "").strip()

    def _sync_cliente_upsert_api(self, cod: str, nome: str, endereco: str, telefone: str, vendedor: str):
        if not is_desktop_api_sync_enabled():
            return
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return
        cod_n = upper(str(cod or "").strip())
        nome_n = upper(str(nome or "").strip())
        if not cod_n or not nome_n:
            return
        payload = {
            "cod_cliente": cod_n,
            "nome_cliente": nome_n,
            "endereco": upper(str(endereco or "").strip()),
            "telefone": upper(str(telefone or "").strip()),
            "vendedor": upper(str(vendedor or "").strip()),
        }
        _call_api(
            "POST",
            "desktop/cadastros/clientes/upsert",
            payload=payload,
            extra_headers={"X-Desktop-Secret": desktop_secret},
        )

    def carregar(self):
        rows = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/clientes/base?ordem=nome&limit=5000",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, list):
                    rows = [
                        (
                            str(r.get("cod_cliente") or ""),
                            str(r.get("nome_cliente") or ""),
                            str(r.get("endereco") or ""),
                            str(r.get("telefone") or ""),
                            str(r.get("vendedor") or ""),
                        )
                        for r in resp
                        if isinstance(r, dict)
                    ]
            except Exception:
                logging.debug("Falha ao carregar clientes via API; usando fallback local.", exc_info=True)

        if not rows:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT cod_cliente, nome_cliente, endereco, telefone, vendedor
                    FROM clientes
                    ORDER BY nome_cliente ASC
                    LIMIT 5000
                """)
                rows = cur.fetchall()

        self.tree.delete(*self.tree.get_children())

        for r in rows:
            tree_insert_aligned(self.tree, "", "end", (r[0] or "", r[1] or "", r[2] or "", r[3] or "", r[4] or ""))

        if self.app and hasattr(self.app, "refresh_programacao_comboboxes"):
            self.app.refresh_programacao_comboboxes()

    def inserir_linha(self):
        tree_insert_aligned(self.tree, "", "end", ("", "", "", "", ""))
        items = self.tree.get_children()
        if items:
            self.tree.see(items[-1])

    def _on_tree_scroll(self, event=None):
        if self._editing:
            self._commit_edit()

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
        vals = self._get_row_values(row_id)
        if col_index < 0 or col_index >= len(vals):
            return

        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return

        x, y, w, h = bbox
        value = vals[col_index]

        entry = ttk.Entry(self.tree, style="Field.TEntry")
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.focus_set()
        entry.selection_range(0, "end")

        entry.bind("<Return>", lambda e: self._commit_edit())
        entry.bind("<Escape>", lambda e: self._cancel_edit())
        entry.bind("<FocusOut>", lambda e: self._commit_edit())

        self._editing = (row_id, col_index, entry)

    def _commit_edit(self):
        if not self._editing:
            return

        row_id, col_index, entry = self._editing
        try:
            new_value = entry.get()
        except Exception:
            new_value = ""

        vals = self._get_row_values(row_id)
        if 0 <= col_index < len(vals):
            vals[col_index] = str(new_value).strip()
            self.tree.item(row_id, values=tuple(vals))

        try:
            entry.destroy()
        except Exception:
            logging.debug("Falha ignorada")
        self._editing = None

    def _cancel_edit(self):
        if not self._editing:
            return
        _, _, entry = self._editing
        try:
            entry.destroy()
        except Exception:
            logging.debug("Falha ignorada")
        self._editing = None

    def salvar_alteracoes(self):
        if not ensure_system_api_binding(context="Salvar alteracoes de clientes", parent=self):
            return
        if self._editing:
            self._commit_edit()

        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("ATENÇÃO", "Não há dados para salvar.")
            return

        linhas = []
        cods_seen = set()

        for iid in items:
            cod, nome, endereco, telefone, vendedor = self._get_row_values(iid)

            if self._is_blank_row(cod, nome, endereco, telefone, vendedor):
                continue

            cod_norm = self._normalize_cod(cod)
            nome_norm = str(nome or "").strip()

            if not cod_norm or not nome_norm:
                messagebox.showwarning("ATENÇÃO", "Todas as linhas precisam ter pelo menos CÓD CLIENTE e NOME CLIENTE.")
                return

            cod_key = upper(cod_norm)
            if cod_key in cods_seen:
                messagebox.showwarning("ATENÇÃO", f"CÓD CLIENTE duplicado na tabela: {cod_norm}\n\nCorrija antes de salvar.")
                return
            cods_seen.add(cod_key)

            linhas.append((upper(cod_norm), upper(nome_norm), upper(endereco), upper(telefone), upper(vendedor)))

        if not linhas:
            messagebox.showwarning("ATENÇÃO", "Nenhuma linha válida para salvar.")
            return

        total = 0
        sync_falhas = 0
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_mode = bool(desktop_secret and is_desktop_api_sync_enabled())
        try:
            if api_mode:
                for cod, nome, endereco, telefone, vendedor in linhas:
                    try:
                        self._sync_cliente_upsert_api(cod, nome, endereco, telefone, vendedor)
                        total += 1
                    except Exception:
                        sync_falhas += 1
            else:
                with get_db() as conn:
                    cur = conn.cursor()
                    for cod, nome, endereco, telefone, vendedor in linhas:
                        cur.execute("""
                            INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone, vendedor)
                            VALUES (?, ?, ?, ?, ?)
                            ON CONFLICT(cod_cliente) DO UPDATE SET
                                nome_cliente=excluded.nome_cliente,
                                endereco=excluded.endereco,
                                telefone=excluded.telefone,
                                vendedor=excluded.vendedor
                        """, (cod, nome, endereco, telefone, vendedor))
                        total += 1
                        try:
                            self._sync_cliente_upsert_api(cod, nome, endereco, telefone, vendedor)
                        except Exception:
                            sync_falhas += 1

            msg = f"Clientes salvos/atualizados: {total}"
            if sync_falhas:
                msg += f"\nFalhas de sincronização API: {sync_falhas}"
            messagebox.showinfo("OK", msg)

        except Exception as e:
            messagebox.showerror("ERRO", str(e))

        self.carregar()

    def importar_clientes_excel(self):
        if not ensure_system_api_binding(context="Importar clientes via planilha", parent=self):
            return
        path = filedialog.askopenfilename(
            title="IMPORTAR CLIENTES (EXCEL)",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if not path:
            return

        if not (require_pandas() and require_excel_support(path)):
            return

        try:
            df = pd.read_excel(path, engine=excel_engine_for(path))

            col_cod = guess_col(df.columns, ["cod", "cód", "codigo", "cliente", "cod cliente"])
            col_nome = guess_col(df.columns, ["nome", "cliente"])
            col_end = guess_col(df.columns, ["endereco", "endereço", "rua", "logradouro"])
            col_tel = guess_col(df.columns, ["telefone", "fone", "celular", "contato"])
            col_vendedor = guess_col(df.columns, ["vendedor", "vend", "representante"])

            if not col_cod or not col_nome:
                # Fallback: usa as duas primeiras colunas como codigo/nome quando o cabecalho vier fora do padrao.
                cols = list(df.columns or [])
                if len(cols) >= 2:
                    col_cod = col_cod or cols[0]
                    col_nome = col_nome or cols[1]
                else:
                    messagebox.showerror("ERRO", "NÃO IDENTIFIQUEI AS COLUNAS DE CÓDIGO E NOME DO CLIENTE NO EXCEL.")
                    return

            total = 0
            sync_falhas = 0
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            api_mode = bool(desktop_secret and is_desktop_api_sync_enabled())
            if api_mode:
                for _, r in df.iterrows():
                    cod = str(r.get(col_cod, "")).strip()
                    nome = str(r.get(col_nome, "")).strip()
                    if not cod or not nome:
                        continue

                    endereco = str(r.get(col_end, "")).strip() if col_end else ""
                    telefone = str(r.get(col_tel, "")).strip() if col_tel else ""
                    vendedor = str(r.get(col_vendedor, "")).strip() if col_vendedor else ""
                    try:
                        self._sync_cliente_upsert_api(cod, nome, endereco, telefone, vendedor)
                        total += 1
                    except Exception:
                        sync_falhas += 1
            else:
                with get_db() as conn:
                    cur = conn.cursor()
                    for _, r in df.iterrows():
                        cod = str(r.get(col_cod, "")).strip()
                        nome = str(r.get(col_nome, "")).strip()
                        if not cod or not nome:
                            continue

                        endereco = str(r.get(col_end, "")).strip() if col_end else ""
                        telefone = str(r.get(col_tel, "")).strip() if col_tel else ""
                        vendedor = str(r.get(col_vendedor, "")).strip() if col_vendedor else ""

                        cur.execute("""
                            INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone, vendedor)
                            VALUES (?, ?, ?, ?, ?)
                            ON CONFLICT(cod_cliente) DO UPDATE SET
                                nome_cliente=excluded.nome_cliente,
                                endereco=excluded.endereco,
                                telefone=excluded.telefone,
                                vendedor=excluded.vendedor
                        """, (upper(cod), upper(nome), upper(endereco), upper(telefone), upper(vendedor)))
                        total += 1
                        try:
                            self._sync_cliente_upsert_api(cod, nome, endereco, telefone, vendedor)
                        except Exception:
                            sync_falhas += 1

            msg = f"CLIENTES IMPORTADOS/ATUALIZADOS: {total}"
            if sync_falhas:
                msg += f"\nFalhas de sincronização API: {sync_falhas}"
            messagebox.showinfo("OK", msg)
            self.carregar()

        except Exception as e:
            messagebox.showerror("ERRO", str(e))


class CadastrosPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Cadastros")

        card = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        card.grid(row=0, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        nb = ttk.Notebook(card)
        nb.grid(row=0, column=0, sticky="nsew")

        frm_motoristas = ttk.Frame(nb, style="Content.TFrame")
        crud_motoristas = CadastroCRUD(
            frm_motoristas,
            "Motoristas",
            "motoristas",
            [
                ("nome", "NOME"),
                ("codigo", "CODIGO"),
                ("senha", "SENHA"),
                
                ("cpf", "CPF"),
                
                ("telefone", "TELEFONE"),
                ("status", "STATUS"),
            ],
            app=app
        )
        crud_motoristas.pack(fill="both", expand=True)
        nb.add(frm_motoristas, text="Motoristas")

        frm_usuarios = ttk.Frame(nb, style="Content.TFrame")
        crud_usuarios = CadastroCRUD(
            frm_usuarios,
            "Usuários",
            "usuarios",
            [
                ("nome", "NOME"),
                ("codigo", "CODIGO"),
                ("senha", "SENHA"),
                ("permissoes", "PERMISSÕES"),
                ("cpf", "CPF"),
                
                ("telefone", "TELEFONE"),
            ],
            app=app
        )
        crud_usuarios.pack(fill="both", expand=True)
        nb.add(frm_usuarios, text="Usuários")

        frm_veiculos = ttk.Frame(nb, style="Content.TFrame")
        crud_veiculos = CadastroCRUD(
            frm_veiculos,
            "VeÃÂÂculos",
            "veiculos",
            [
                ("placa", "PLACA"),
                ("modelo", "MODELO"),
                
                ("capacidade_cx", "CAPACIDADE (CX)"),
            ],
            app=app
        )
        crud_veiculos.pack(fill="both", expand=True)
        nb.add(frm_veiculos, text="VeÃÂÂculos")

        frm_ajudantes = ttk.Frame(nb, style="Content.TFrame")
        crud_ajudantes = CadastroCRUD(
            frm_ajudantes,
            "Ajudantes",
            "ajudantes",
            [
                ("nome", "NOME"),
                ("sobrenome", "SOBRENOME"),
                ("telefone", "TELEFONE"),
                ("status", "STATUS"),
            ],
            app=app
        )
        crud_ajudantes.pack(fill="both", expand=True)
        nb.add(frm_ajudantes, text="Ajudantes")

        
        frm_clientes = ttk.Frame(nb, style="Content.TFrame")
        clientes_page = ClientesImportPage(frm_clientes, app=app)
        clientes_page.pack(fill="both", expand=True)
        nb.add(frm_clientes, text="Clientes")

    def on_show(self):
        self.set_status("STATUS: Cadastros (CRUD + Base de Clientes via Wibi).")

# ==========================
# ===== FIM DA PARTE 3 (FINAL) =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 4 (ATUALIZADA) =====
# ==========================

# =========================================================
# 4.0 IMPORTAÇÃO DE VENDAS
# =========================================================
class ImportarVendasPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Importar Vendas (Excel)")

        top = ttk.Frame(self.body, style="Content.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(3, weight=1)

        ttk.Button(top, text="ðÅ¸â€œ¥ IMPORTAR EXCEL", style="Primary.TButton", command=self.importar_excel).grid(row=0, column=0, padx=6)
        ttk.Button(top, text="ðÅ¸§¹ LIMPAR TUDO", style="Danger.TButton", command=self.limpar_tudo).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(top, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton", command=self.carregar).grid(row=0, column=2, padx=6)

        self.lbl_info = ttk.Label(
            top,
            text="Selecione as vendas que irão para Programação (duplo clique marca/desmarca).",
            background="#F4F6FB",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_info.grid(row=0, column=3, sticky="w", padx=10)

        card = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        card.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        # Filtros
        filt = ttk.Frame(card, style="Card.TFrame")
        filt.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        filt.grid_columnconfigure(1, weight=1)
        filt.grid_columnconfigure(9, weight=1)

        ttk.Label(filt, text="Buscar:", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.ent_busca = ttk.Entry(filt, style="Field.TEntry")
        self.ent_busca.grid(row=0, column=1, sticky="ew", padx=6)
        self.ent_busca.bind("<Return>", lambda e: self.carregar())

        ttk.Button(filt, text="ðŸ”Ž FILTRAR", style="Ghost.TButton", command=self.carregar).grid(row=0, column=2, padx=6)
        ttk.Button(filt, text="âÅ“â€¦ MARCAR", style="Primary.TButton", command=self.marcar_selecionadas).grid(row=0, column=3, padx=6)

        ttk.Button(filt, text="âÅ“â€¦ MARCAR TODOS", style="Warn.TButton", command=lambda: self.set_all_selected(1)).grid(row=0, column=4, padx=6)
        ttk.Button(filt, text="âËœâ€˜ DESMARCAR TODOS", style="Ghost.TButton", command=lambda: self.set_all_selected(0)).grid(row=0, column=5, padx=6)

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

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"ID", "QNT"},
            money_cols={"VR TOTAL"},
            date_cols={"DATA"}
        )

        self.carregar()

    def on_show(self):
        self.set_status("STATUS: Importação e seleção de vendas para programação.")
        self.carregar()

    # -------------------------
    # Helpers segurança/normalização
    # -------------------------
    def _norm(self, v):
        return upper(str(v or "").strip())

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
        Normaliza data da venda para chave estável (YYYY-MM-DD quando possÃÂvel).
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
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            return
        try:
            with get_db() as conn:
                cur = conn.cursor()
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

    def importar_excel(self):
        path = filedialog.askopenfilename(
            title="IMPORTAR VENDAS (EXCEL)",
            filetypes=[("Excel", "*.xls *.xlsx")]
        )
        if not path:
            return

        if not (require_pandas() and require_excel_support(path)):
            return

        try:
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
                messagebox.showerror("ERRO", "Nao identifiquei as colunas: " + ", ".join(missing))
                return

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
                    for row in payload_rows:
                        data_venda = str(row.get("data_venda") or "")
                        pedido_u = str(row.get("pedido") or "")
                        cliente_u = str(row.get("cliente") or "")
                        produto_u = str(row.get("produto") or "")
                        try:
                            cur.execute("""
                                SELECT 1
                                FROM vendas_importadas
                                WHERE UPPER(TRIM(COALESCE(pedido,'')))=UPPER(TRIM(?))
                                  AND UPPER(TRIM(COALESCE(cliente,'')))=UPPER(TRIM(?))
                                  AND UPPER(TRIM(COALESCE(produto,'')))=UPPER(TRIM(?))
                                  AND COALESCE(TRIM(data_venda),'')=COALESCE(TRIM(?),'')
                                LIMIT 1
                            """, (pedido_u, cliente_u, produto_u, data_venda))
                            exists = cur.fetchone()
                        except Exception:
                            exists = None
                        if exists:
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
                        total += 1

            msg = f"Vendas importadas: {total}"
            if ignoradas:
                msg += f"\nIgnoradas (duplicadas/invalidas): {ignoradas}"
            if ignoradas_invalidas:
                msg += f"\nIgnoradas (linhas invalidas/NaN): {ignoradas_invalidas}"

            if hasattr(self, "ent_busca"):
                try:
                    self.ent_busca.delete(0, "end")
                except Exception:
                    logging.debug("Falha ao limpar busca ap?s importa??o")

            if opcionais_ausentes:
                msg += "\nCampos opcionais nao encontrados (preenchidos em branco): " + ", ".join(opcionais_ausentes)

            messagebox.showinfo("OK", msg)
            self.carregar()

        except Exception as e:
            messagebox.showerror("ERRO", str(e))

    def carregar(self):
        self._ensure_vendas_usada_cols()

        busca = self._norm(self.ent_busca.get()) if hasattr(self, "ent_busca") else ""
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
                        )
                        for r in rows_api
                        if isinstance(r, dict)
                    ]
            except Exception:
                logging.debug("Falha ao carregar vendas importadas via API; usando fallback local.", exc_info=True)

        if not rows:
            with get_db() as conn:
                cur = conn.cursor()
                if busca:
                    cur.execute("""
                        SELECT id, selecionada, pedido, data_venda, cliente, nome_cliente, produto, vr_total, qnt, cidade, vendedor
                        FROM vendas_importadas
                        WHERE (IFNULL(usada,0)=0)
                          AND (
                            pedido LIKE ? OR cliente LIKE ? OR nome_cliente LIKE ? OR vendedor LIKE ? OR produto LIKE ?
                          )
                        ORDER BY id DESC
                    """, (f"%{busca}%", f"%{busca}%", f"%{busca}%", f"%{busca}%", f"%{busca}%"))
                else:
                    cur.execute("""
                        SELECT id, selecionada, pedido, data_venda, cliente, nome_cliente, produto, vr_total, qnt, cidade, vendedor
                        FROM vendas_importadas
                        WHERE (IFNULL(usada,0)=0)
                        ORDER BY id DESC
                    """)
                rows = cur.fetchall() or []

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

        self.set_status(f"STATUS: {len(rows)} registros carregados (N?O usadas)  Selecionadas: {selected_count}.")

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
            except Exception:
                logging.debug("Falha ao alternar selecao via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE vendas_importadas
                SET selecionada = CASE WHEN selecionada=1 THEN 0 ELSE 1 END
                WHERE id=? AND IFNULL(usada,0)=0
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
            except Exception:
                logging.debug("Falha ao marcar/desmarcar todas via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE vendas_importadas SET selecionada=? WHERE IFNULL(usada,0)=0", (int(val),))
        self.carregar()

    def marcar_selecionadas(self):
        """
        Marca apenas as vendas selecionadas na tabela (suporta sele??o m?ltipla com Shift/Ctrl).
        """
        self._ensure_vendas_usada_cols()
        itens = self.tree.selection() or ()
        if not itens:
            messagebox.showwarning("ATEN??O", "Selecione uma ou mais vendas para marcar.")
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
            messagebox.showwarning("ATEN??O", "N?o foi poss?vel identificar as vendas selecionadas.")
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
                self.set_status(f"STATUS: {len(ids)} venda(s) marcada(s) a partir da sele??o.")
                return
            except Exception:
                logging.debug("Falha ao marcar ids via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.executemany(
                "UPDATE vendas_importadas SET selecionada=1 WHERE id=? AND IFNULL(usada,0)=0",
                [(rid,) for rid in ids],
            )

        self.carregar()
        self.set_status(f"STATUS: {len(ids)} venda(s) marcada(s) a partir da sele??o.")

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
            except Exception:
                logging.debug("Falha ao limpar vendas importadas via API; usando fallback local.", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM vendas_importadas")

        self.carregar()

    def __init__(self, parent, app):
        super().__init__(parent, app, "Programação")

        self._editing = None
        self._prog_cols_checked = False  # evita PRAGMA/ALTER toda hora
        self._vendas_cols_checked = False  # evita PRAGMA/ALTER toda hora
        self._equipe_display_map = {}
        self._editing_programacao_codigo = ""
        self._loaded_venda_ids = []

        # -------------------------
        # Cabeçalho (dados da programação)
        # -------------------------
        card = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        card.grid(row=0, column=0, sticky="ew")
        # Distribui o espa?o entre os campos para evitar "sumir" o ?ltimo campo (C?digo)
        for col in range(0, 9):
            card.grid_columnconfigure(col, weight=1, uniform="prog_head")

        ttk.Label(card, text="Motorista", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.cb_motorista = ttk.Combobox(card, state="readonly", width=16)
        self.cb_motorista.grid(row=1, column=0, sticky="ew", padx=6)

        ttk.Label(card, text="VeÃÂÂculo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.cb_veiculo = ttk.Combobox(card, state="readonly", width=12)
        self.cb_veiculo.grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Ajudantes (multipla escolha)", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.btn_ajudantes = ttk.Button(card, text="ðÅ¸â€˜¥ Selecionar ajudantes", style="Ghost.TButton", command=self._open_ajudantes_selector)
        self.btn_ajudantes.grid(row=1, column=2, columnspan=2, sticky="ew", padx=6)
        self.lbl_ajudantes_sel = ttk.Label(card, text="Nenhum selecionado", style="CardLabel.TLabel")
        self.lbl_ajudantes_sel.grid(row=2, column=2, columnspan=2, sticky="w", padx=6, pady=(2, 0))
        self._ajudantes_rows = []
        self._ajudantes_selected_keys = []
        self._ajudantes_mode = "ajudantes"
        self._ajudantes_popup = None
        self._ajudantes_tree = None
        self._ajudantes_tree_key_by_iid = {}
        self._ajudantes_tree_iid_by_key = {}
        self._ajudantes_filter_var = tk.StringVar(value="")

        ttk.Label(card, text="Local da Rota", style="CardLabel.TLabel").grid(row=0, column=4, sticky="w")
        self.cb_local_rota = ttk.Combobox(card, state="readonly", values=["SERRA", "SERTAO"], width=10)
        self.cb_local_rota.grid(row=1, column=4, sticky="ew", padx=6)

        ttk.Label(card, text="Estimativa (KG/CX)", style="CardLabel.TLabel").grid(row=0, column=5, sticky="w")
        frm_estim = ttk.Frame(card, style="Card.TFrame")
        frm_estim.grid(row=1, column=5, sticky="ew", padx=6)
        frm_estim.grid_columnconfigure(1, weight=1)
        self.cb_estimativa_tipo = ttk.Combobox(frm_estim, state="readonly", values=["KG", "CX"], width=5)
        self.cb_estimativa_tipo.grid(row=0, column=0, sticky="w")
        self.cb_estimativa_tipo.set("KG")
        self.cb_estimativa_tipo.bind("<<ComboboxSelected>>", lambda _e: self._on_estimativa_tipo_change())
        self.ent_kg = ttk.Entry(frm_estim, style="Field.TEntry", width=10)
        self.ent_kg.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.lbl_estimado_hint = ttk.Label(card, text="KG estimado", style="CardLabel.TLabel")
        self.lbl_estimado_hint.grid(row=2, column=5, sticky="w", padx=6, pady=(2, 0))

        ttk.Label(card, text="Carregamento", style="CardLabel.TLabel").grid(row=0, column=6, sticky="w")
        self.ent_carregamento_prog = ttk.Entry(card, style="Field.TEntry", width=12)
        self.ent_carregamento_prog.grid(row=1, column=6, sticky="ew", padx=6)
        bind_entry_smart(self.ent_carregamento_prog, "text")

        ttk.Label(card, text="Adiantamento (R$)", style="CardLabel.TLabel").grid(row=0, column=7, sticky="w")
        self.ent_adiantamento_prog = ttk.Entry(card, style="Field.TEntry", width=12)
        self.ent_adiantamento_prog.grid(row=1, column=7, sticky="ew", padx=6)
        self.ent_adiantamento_prog.insert(0, "0,00")
        self._bind_money_entry(self.ent_adiantamento_prog)

        ttk.Label(card, text="Código", style="CardLabel.TLabel").grid(row=0, column=8, sticky="w")
        ttk.Label(card, text="Total Caixas", style="CardLabel.TLabel").grid(row=2, column=6, sticky="w")
        self.ent_total_caixas_prog = ttk.Entry(card, style="Field.TEntry", state="readonly", width=12)
        self.ent_total_caixas_prog.grid(row=2, column=7, sticky="ew", padx=6, pady=(2, 0))

        self.ent_codigo = ttk.Entry(card, style="Field.TEntry", state="readonly", width=10)
        self.ent_codigo.grid(row=1, column=8, sticky="ew", padx=6)
        self.lbl_prog_status_badge = ttk.Label(
            card,
            text="Status atual: -",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_prog_status_badge.grid(row=2, column=8, sticky="w", padx=6, pady=(2, 0))

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
            text="CARREGAR VENDAS SELECIONADAS",
            style="Warn.TButton",
            command=self.carregar_vendas_selecionadas
        ).grid(row=0, column=0, padx=4, sticky="ew")

        ttk.Button(acoes_linha_1, text="âÅ¾â€¢ INSERIR LINHA", style="Ghost.TButton",
                   command=self.inserir_linha).grid(row=0, column=1, padx=4, sticky="ew")

        ttk.Button(acoes_linha_1, text="âÅ¾â€“ REMOVER LINHA", style="Danger.TButton",
                   command=self.remover_linha).grid(row=0, column=2, padx=4, sticky="ew")

        ttk.Button(acoes_linha_1, text="ðÅ¸§¹ LIMPAR ITENS", style="Danger.TButton",
                   command=self.limpar_itens).grid(row=0, column=3, padx=4, sticky="ew")

        acoes_linha_2 = ttk.Frame(top2, style="Card.TFrame")
        acoes_linha_2.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        for i in range(3):
            acoes_linha_2.grid_columnconfigure(i, weight=1, uniform="prog_actions_2")

        ttk.Button(
            acoes_linha_2,
            text="SALVAR PROGRAMAÇÃO",
            style="Primary.TButton",
            command=self.salvar_programacao
        ).grid(row=0, column=0, padx=4, sticky="ew")

        ttk.Button(
            acoes_linha_2,
            text="âÅ“Âï¸Â EDITAR PROGRAMAÇÃO",
            style="Ghost.TButton",
            command=self.carregar_programacao_para_edicao
        ).grid(row=0, column=1, padx=4, sticky="ew")

        ttk.Button(
            acoes_linha_2,
            text="IMPRIMIR ROMANEIOS",
            style="Ghost.TButton",
            command=self.imprimir_romaneios_programacao
        ).grid(row=0, column=2, padx=4, sticky="ew")

        ttk.Label(
            top2,
            text="Dica: duplo clique para editar Endereço/Caixas/Preço/Vendedor/Pedido/Obs. ENTER confirma, ESC cancela.",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        ).grid(row=2, column=0, padx=6, pady=(8, 0), sticky="w")

        cols = ["COD CLIENTE", "NOME CLIENTE", "PRODUTO", "ENDEREÇO", "CAIXAS", "KG", "PREÇO", "VENDEDOR", "PEDIDO", "OBS"]

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
                "ENDEREÇO",
                "CAIXAS",
                "PREÇO",
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
            self.tree.column(c, width=160, minwidth=90, anchor="w")

        self.tree.column("ENDEREÇO", width=260, minwidth=180)
        self.tree.column("NOME CLIENTE", width=260, minwidth=160)
        self.tree.column("PRODUTO", width=160, minwidth=120)
        self.tree.column("PEDIDO", width=160, minwidth=120)
        self.tree.column("OBS", width=260, minwidth=140)
        self.tree.column("CAIXAS", width=90, anchor="center")
        self.tree.column("KG", width=90, anchor="center")
        self.tree.column("PREÇO", width=110, anchor="e")

        self.tree.bind("<Double-1>", self._start_edit_cell)
        self.tree.bind("<MouseWheel>", self._on_tree_scroll, add=True)
        self.tree.bind("<Button-4>", self._on_tree_scroll, add=True)  # linux
        self.tree.bind("<Button-5>", self._on_tree_scroll, add=True)  # linux

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"CAIXAS", "KG"},
            money_cols={"PREÇO"},
            date_cols=set()
        )

        self.refresh_comboboxes()
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
            self.btn_ajudantes.configure(text="ðÅ¸â€˜¥ Selecionar ajudantes")
            self.lbl_ajudantes_sel.configure(text="Nenhum selecionado")
            return
        self.btn_ajudantes.configure(text=f"{qtd} selecionado(s)")
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
        ttk.Label(search_row, text="ðŸ”Ž BUSCAR", style="CardLabel.TLabel").pack(side="left", padx=(0, 8))
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
        ttk.Button(footer, text="ðÅ¸§¹ LIMPAR", style="Ghost.TButton", command=self._clear_ajudantes_selection).pack(side="left")
        ttk.Button(footer, text="âœ” CONFIRMAR", style="Primary.TButton", command=self._confirm_ajudantes_popup).pack(side="right")

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

    def _on_tree_scroll(self, event=None):
        if self._editing:
            self._commit_edit()

    def refresh_comboboxes(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
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
                self.cb_motorista["values"] = valores_motoristas

                self.cb_veiculo["values"] = [upper((r or {}).get("placa") or "") for r in (vei_resp or []) if isinstance(r, dict)]

                self._equipe_display_map = {}
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
                        self._equipe_display_map[upper(display)] = ajudante_id
                self._set_ajudantes_options(rows_ajudantes, mode="ajudantes")
                return
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

            self.cb_motorista["values"] = valores_motoristas

            cur.execute("SELECT placa FROM veiculos ORDER BY placa")
            self.cb_veiculo["values"] = [r[0] for r in cur.fetchall()]

            self._equipe_display_map = {}
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
                            self._equipe_display_map[upper(display)] = ajudante_id
                self._set_ajudantes_options(rows_ajudantes, mode="ajudantes")
            except Exception:
                # fallback de base antiga
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
                                self._equipe_display_map[upper(display)] = upper(codigo)
                    self._set_ajudantes_options(rows_equipes, mode="equipes")
                except Exception:
                    self._set_ajudantes_options([], mode="ajudantes")

    def on_show(self):
        self.set_status("STATUS: Carregue vendas e ajuste dados antes de salvar a programação.")
        self.refresh_comboboxes()
        self._on_estimativa_tipo_change()
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
            self._refresh_programacao_status_badge()
        except Exception:
            logging.debug("Falha ignorada")

    def _on_estimativa_tipo_change(self):
        tipo = upper((self.cb_estimativa_tipo.get() if hasattr(self, "cb_estimativa_tipo") else "KG") or "KG")
        if tipo == "CX":
            self.lbl_estimado_hint.config(text="Modo CIF (Caixas estimadas)")
        else:
            self.lbl_estimado_hint.config(text="Modo FOB (KG estimado)")

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

    def _obter_status_programacao_para_edicao(self, codigo: str, status_local: str = ""):
        codigo = upper(str(codigo or "").strip())
        st_local = upper(str(status_local or "").strip())
        st_api = ""

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
            except Exception:
                logging.debug("Falha ao consultar status da programação na API central", exc_info=True)

        st_ref = st_api or st_local
        return st_ref, st_api, st_local

    def _warn_bloqueio_edicao_status(self, codigo: str, status_ref: str):
        st = upper(str(status_ref or "").strip())
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
                carreg_cols = [c for c in ("local_carregamento", "granja_carregada", "local_carregado", "local_carreg") if c in cols_prog]
                if carreg_cols:
                    carreg_expr = "COALESCE(" + ", ".join(carreg_cols) + ", '')"
                else:
                    carreg_expr = "''"
                adiant_cols = [c for c in ("adiantamento", "adiantamento_rota") if c in cols_prog]
                if adiant_cols:
                    adiant_expr = "COALESCE(" + ", ".join(adiant_cols) + ", 0)"
                else:
                    adiant_expr = "0"
                mot_cod_cols = [c for c in ("motorista_codigo", "codigo_motorista") if c in cols_prog]
                if mot_cod_cols:
                    mot_cod_expr = "COALESCE(" + ", ".join(mot_cod_cols) + ", '')"
                else:
                    mot_cod_expr = "''"
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
        if not itens:
            itens = fetch_programacao_itens(codigo)
        self.limpar_itens()
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
                    "desktop/clientes/base?ordem=codigo&limit=5000",
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
        msg = f"STATUS: Itens carregados: {len(self.tree.get_children())} vendas selecionadas (nao usadas). (edite antes de salvar)"
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
        editable = {"ENDEREÇO", "CAIXAS", "KG", "PREÇO", "VENDEDOR", "PEDIDO", "OBS"}
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

        elif col_name in {"KG", "PREÇO"}:
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
            cols = [upper(r[1]).lower() for r in cur.fetchall()]
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
                messagebox.showwarning("ATENÇÃO", "Informe a estimativa em caixas (CX) para CIF.")
                return
        else:
            if kg_estimado <= 0:
                messagebox.showwarning("ATENÇÃO", "Informe o KG estimado para FOB.")
                return

        if not motorista_nome or not veiculo:
            messagebox.showwarning("ATENÇÃO", "Selecione Motorista e VeÃÂÂculo.")
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

        itens = []
        for iid in self.tree.get_children():
            itens.append(self._get_row_values(iid))

        if not itens:
            messagebox.showwarning("ATENÇÃO", "Carregue itens (vendas selecionadas) antes de salvar.")
            return

        # âÅâ€œââ‚¬¦ validação mÃÂÂnima por linha (segurança)
        for v in itens:
            cod_cliente = v[0]
            nome_cliente = v[1]
            if not str(cod_cliente).strip() or not str(nome_cliente).strip():
                messagebox.showwarning("ATENÇÃO", "Há linhas sem COD CLIENTE ou NOME CLIENTE. Corrija antes de salvar.")
                return

        codigo = None
        codigo_atual = upper((self.ent_codigo.get() or "").strip())
        is_update_existing = False
        data_criacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Totais (compatibilidade API)
        try:
            total_caixas = sum(safe_int(v[4], 0) for v in itens)
        except Exception:
            total_caixas = 0
        try:
            total_quilos = round(sum(safe_float(v[5], 0.0) for v in itens), 2)
        except Exception:
            total_quilos = 0.0

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
                            f"Capacidade excedida para o veículo {veiculo}.\n\n"
                            f"Caixas na programação: {caixas_para_validar}\n"
                            f"Capacidade do veículo: {capacidade_cx}"
                        )
                        return
            except Exception:
                logging.debug("Falha ao validar capacidade via API; usando validacao local/fallback.", exc_info=True)

        motorista_id = None
        api_saved = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()

        try:
            if desktop_secret and is_desktop_api_sync_enabled():
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
                for _ in range(max_tries):
                    codigo_try = codigo_atual if is_update_existing else generate_program_code()
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
                        "local_carregamento": local_carreg,
                        "adiantamento": safe_float(adiantamento_val, 0.0),
                        "total_caixas": safe_int(total_caixas, 0),
                        "quilos": safe_float(total_quilos, 0.0),
                        "usuario_criacao": usuario_logado,
                        "usuario_ultima_edicao": usuario_logado,
                        "itens": itens_payload,
                    }
                    try:
                        _call_api(
                            "POST",
                            "desktop/rotas/upsert",
                            payload=payload_sync,
                            extra_headers={"X-Desktop-Secret": desktop_secret},
                        )
                        ids_carregados_api = [safe_int(x, 0) for x in (self._loaded_venda_ids or []) if safe_int(x, 0) > 0]
                        if ids_carregados_api:
                            _call_api(
                                "POST",
                                "desktop/vendas-importadas/consumir",
                                payload={
                                    "ids": ids_carregados_api,
                                    "codigo_programacao": upper(codigo_try),
                                    "usada_em": data_criacao,
                                },
                                extra_headers={"X-Desktop-Secret": desktop_secret},
                            )
                        codigo = codigo_try
                        api_saved = True
                        break
                    except Exception as exc_try:
                        last_api_error = exc_try
                        if is_update_existing:
                            break
                if (not api_saved) and last_api_error:
                    logging.warning(
                        "Falha ao salvar programacao via API; aplicando fallback local. Erro: %s",
                        last_api_error,
                    )

            if not api_saved:
                with get_db() as conn:
                    cur = conn.cursor()

                if not self._prog_cols_checked:
                    self._ensure_prog_columns_for_api(cur)
                    self._prog_cols_checked = True

                if not self._vendas_cols_checked:
                    self._ensure_vendas_usada_cols(cur)
                    self._vendas_cols_checked = True

                # Valida capacidade do veÃÂÂculo (CX) antes de salvar programação.
                cap_col = "capacidade_cx"
                try:
                    cur.execute("PRAGMA table_info(veiculos)")
                    cols_vei = [str(r[1]).lower() for r in cur.fetchall()]
                    if "capacidade_cx" not in cols_vei and "capacidade_c" in cols_vei:
                        cap_col = "capacidade_c"
                except Exception:
                    cap_col = "capacidade_cx"

                try:
                    cur.execute(
                        f"SELECT {cap_col} FROM veiculos WHERE UPPER(placa)=UPPER(?) LIMIT 1",
                        (veiculo,),
                    )
                    vrow = cur.fetchone()
                except Exception:
                    vrow = None

                if not vrow:
                    messagebox.showwarning(
                        "ATENÇÃO",
                        f"VeÃÂÂculo não encontrado no cadastro: {veiculo}."
                    )
                    return

                capacidade_cx = safe_int(vrow[0], -1)
                if capacidade_cx < 0:
                    messagebox.showwarning(
                        "ATENÇÃO",
                        f"Capacidade (CX) inválida para o veÃÂÂculo {veiculo}. Ajuste no cadastro de veÃÂÂculos."
                    )
                    return

                caixas_para_validar = caixas_estimado if tipo_estimativa == "CX" else total_caixas
                if caixas_para_validar > capacidade_cx:
                    messagebox.showwarning(
                        "ATENÇÃO",
                        f"Capacidade excedida para o veÃÂÂculo {veiculo}.\n\n"
                        f"Caixas na programação: {caixas_para_validar}\n"
                        f"Capacidade do veÃÂÂculo: {capacidade_cx}"
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
                    for _ in range(5):
                        codigo = generate_program_code()
                        try:
                            cur.execute(
                                "SELECT 1 FROM programacoes WHERE codigo_programacao=? ORDER BY id DESC LIMIT 1",
                                (codigo,)
                            )
                            if cur.fetchone():
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
                            "UPDATE programacoes SET prestacao_status='PENDENTE' WHERE codigo_programacao=?",
                            (codigo,)
                        )
                except Exception:
                    logging.debug("Falha ignorada")

                # Itens
                if is_update_existing:
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
                for (cod_cliente, nome_cliente, produto, endereco, caixas, kg, preco, vendedor, pedido, obs) in itens:
                    cur.execute("""
                        INSERT INTO programacao_itens (
                            codigo_programacao, cod_cliente, nome_cliente,
                            qnt_caixas, kg, preco, endereco, vendedor, pedido, produto
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                        upper(produto)
                    ))

                    # Mantém clientes atualizados (compatÃÂÂvel com bases antigas)
                    if cod_cliente and nome_cliente:
                        try:
                            cur.execute("""
                                INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone, vendedor)
                                VALUES (?, ?, ?, '', ?)
                                ON CONFLICT(cod_cliente) DO UPDATE SET
                                    nome_cliente=excluded.nome_cliente,
                                    endereco=excluded.endereco,
                                    vendedor=excluded.vendedor
                            """, (upper(cod_cliente), upper(nome_cliente), upper(endereco), upper(vendedor)))
                        except Exception:
                            cur.execute("""
                                INSERT INTO clientes (cod_cliente, nome_cliente, endereco, telefone)
                                VALUES (?, ?, ?, '')
                                ON CONFLICT(cod_cliente) DO UPDATE SET
                                    nome_cliente=excluded.nome_cliente,
                                    endereco=excluded.endereco
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

        except Exception as e:
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
                    "local_carregamento": local_carreg,
                    "adiantamento": safe_float(adiantamento_val, 0.0),
                    "total_caixas": safe_int(total_caixas, 0),
                    "quilos": safe_float(total_quilos, 0.0),
                    "usuario_criacao": usuario_logado,
                    "usuario_ultima_edicao": usuario_logado,
                    "itens": itens_payload,
                }
                _call_api(
                    "POST",
                    "desktop/rotas/upsert",
                    payload=payload_sync,
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                ids_carregados_api = [safe_int(x, 0) for x in (self._loaded_venda_ids or []) if safe_int(x, 0) > 0]
                if ids_carregados_api:
                    _call_api(
                        "POST",
                        "desktop/vendas-importadas/consumir",
                        payload={
                            "ids": ids_carregados_api,
                            "codigo_programacao": upper(codigo),
                            "usada_em": data_criacao,
                        },
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

        if messagebox.askyesno("PDF", "Deseja gerar o PDF da programação agora?\n\n(Pronto para impressão A4)"):
            self.gerar_pdf_programacao_salva(
                codigo, motorista_nome, veiculo, equipe, kg_estimado, tipo_estimativa, caixas_estimado, usuario_logado
            )

        if messagebox.askyesno("Romaneios", "Deseja gerar os romaneios de entrega desta programacao agora?"):
            self.imprimir_romaneios_programacao()

        self._reset_form_after_save()

    def gerar_pdf_programacao_salva(
        self, codigo, motorista, veiculo, equipe, kg_estimado, tipo_estimativa="KG", caixas_estimado=0, usuario_criacao=""
    ):
        if not require_reportlab():
            return
        path = filedialog.asksaveasfilename(
            title="Salvar PDF da Programação",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"PROGRAMACAO_{codigo}.pdf"
        )
        if not path:
            return

        itens = []
        for iid in self.tree.get_children():
            itens.append(self._get_row_values(iid))

        if not itens:
            messagebox.showwarning("ATENÇÃO", "Sem itens na programação.")
            return

        try:
            usuario_edicao = ""
            try:
                meta_prog = self._buscar_meta_programacao(codigo)
                usuario_criacao = upper(str(meta_prog.get("usuario_criacao") or usuario_criacao or "").strip())
                usuario_edicao = upper(str(meta_prog.get("usuario_ultima_edicao") or "").strip())
            except Exception:
                usuario_criacao = upper((usuario_criacao or "").strip())

            c = canvas.Canvas(path, pagesize=A4)
            w, h = A4
            to_txt = lambda v: fix_mojibake_text(str(v or ""))

            y = h - 60
            c.setFont("Helvetica-Bold", 14)
            c.drawString(40, y, f"PROGRAMACAO: {to_txt(codigo)}")
            y -= 22

            c.setFont("Helvetica", 10)
            c.drawString(40, y, f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            y -= 16
            equipe_txt = self._resolve_equipe_ajudantes(equipe)
            c.drawString(40, y, f"Motorista: {to_txt(motorista)}  |  Veiculo: {to_txt(veiculo)}  |  Equipe: {to_txt(equipe_txt)}")
            y -= 16
            if upper(tipo_estimativa) == "CX":
                c.drawString(40, y, f"Estimado (CIF): {safe_int(caixas_estimado, 0)} CX")
            else:
                c.drawString(40, y, f"Estimado (FOB): {safe_float(kg_estimado, 0.0):.2f} KG")
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
                # Espaco minimo para 2 linhas (cliente + observacao) + espaco extra
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

                # Campo de observacao no espaco abaixo do cliente
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

                # Espaco extra entre clientes
                y -= 4

            if y < 40:
                c.showPage()
                y = h - 60

            c.setFont("Helvetica-Oblique", 9)
            c.drawCentredString(w / 2, 26, '"Tudo posso naquele que me fortalece." (Filipenses 4:13)')

            c.save()
            messagebox.showinfo("OK", "PDF gerado com sucesso! (A4 pronto para impressão)")

        except Exception as e:
            messagebox.showerror("ERRO", str(e))


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
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(codigo or '').strip())}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
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
                cur.execute(
                    """
                    SELECT COALESCE(motorista,''), COALESCE(veiculo,''), COALESCE(equipe,''),
                           COALESCE(data_criacao,''), COALESCE(kg_estimado,0),
                           COALESCE(tipo_estimativa,'KG'), COALESCE(caixas_estimado,0),
                           COALESCE(usuario_criacao,''), COALESCE(usuario_ultima_edicao,'')
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

        ttk.Button(bottom, text="â¬â€¦ Anterior", style="Ghost.TButton", command=_prev).pack(side="left")
        ttk.Button(bottom, text="Próximo âÅ¾¡", style="Ghost.TButton", command=_next).pack(side="left", padx=8)
        ttk.Button(bottom, text="ðÅ¸â€œâ€ž GERAR PDF", style="Primary.TButton", command=_export_pdf).pack(side="right")
        ttk.Button(bottom, text="âÅ“â€“ Fechar", style="Danger.TButton", command=win.destroy).pack(side="right", padx=8)

        cb.bind("<<ComboboxSelected>>", _on_sel)
        vias_var.trace_add("write", lambda *_: _render())
        preview.bind("<Configure>", lambda _e: _render())
        _render()

    def imprimir_romaneios_programacao(self):
        if not require_reportlab():
            return

        codigo = upper((self.ent_codigo.get() or "").strip())
        if not codigo:
            codigo = upper(simple_input("Romaneios", "Informe o codigo da programacao:", master=self.app, allow_empty=False) or "")
        if not codigo:
            return

        itens = fetch_programacao_itens(codigo)
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
    win.geometry("460x190")
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
    ttk.Label(card, text=prompt, style="CardLabel.TLabel", justify="left").grid(row=1, column=0, sticky="w", pady=(8, 8))

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

    left = ttk.Frame(btns, style="Card.TFrame")
    left.grid(row=0, column=0, sticky="w")

    ttk.Button(left, text="âœ” CONFIRMAR", style="Primary.TButton", command=ok).pack(side="left")
    ttk.Button(left, text="âÅ“â€“ CANCELAR", style="Ghost.TButton", command=cancel).pack(side="left", padx=8)

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

        ttk.Button(self.card, text="ðÅ¸â€œâ€š CARREGAR", style="Ghost.TButton", command=self.carregar_programacao)\
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
        for i in range(6):
            resumo.grid_columnconfigure(i, weight=1)

        self.lbl_resumo_diaria = ttk.Label(
            resumo, text="VALOR DIARIA: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_diaria.grid(row=0, column=0, sticky="w", padx=(0, 12))

        self.lbl_resumo_qtd = ttk.Label(
            resumo, text="QTD DIARIAS: 0", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_qtd.grid(row=0, column=1, sticky="w", padx=(0, 12))

        self.lbl_resumo_mot = ttk.Label(
            resumo, text="MOTORISTA: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_mot.grid(row=0, column=2, sticky="w", padx=(0, 12))

        self.lbl_resumo_eqp = ttk.Label(
            resumo, text="EQUIPE: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_eqp.grid(row=0, column=3, sticky="w", padx=(0, 12))

        self.lbl_resumo_total = ttk.Label(
            resumo, text="TOTAL: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_total.grid(row=0, column=4, sticky="w", padx=(0, 12))

        self.lbl_resumo_pagar = ttk.Label(
            resumo, text="VALOR A PAGAR: R$ 0,00", background="white", foreground="#111827", font=("Segoe UI", 9, "bold")
        )
        self.lbl_resumo_pagar.grid(row=0, column=5, sticky="w")

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

        ttk.Button(top2, text="ðÅ¸â€˜¤ INSERIR CLIENTE MANUAL", style="Warn.TButton", command=self.inserir_cliente_manual)\
            .grid(row=0, column=0, padx=6)

        ttk.Button(top2, text="ðÅ¸§½ ZERAR RECEBIMENTO", style="Danger.TButton", command=self.zerar_recebimento)\
            .grid(row=0, column=1, padx=6)

        ttk.Button(top2, text="âÅ¾¡ IR PARA DESPESAS", style="Primary.TButton", command=self._ir_para_despesas)\
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

        ttk.Button(frm, text="ðÅ¸â€™¾ SALVAR RECEBIMENTOS", style="Primary.TButton", command=self.salvar_recebimento)\
            .grid(row=1, column=5, sticky="e", padx=(12, 0))

        # âÅâ€œââ‚¬¦ TROCA: no lugar do Excel, botão IMPRIMIR PDF
        ttk.Button(frm, text="ðÅ¸â€“¨ IMPRIMIR PDF", style="Warn.TButton", command=self.imprimir_pdf)\
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

        ttk.Button(btns, text="ðÅ¸â€“¨ IMPRIMIR PDF", style="Warn.TButton", command=self.imprimir_pdf)\
            .grid(row=0, column=0, padx=6, sticky="w")

        ttk.Button(btns, text="ðÅ¸â€˜Â MOSTRAR DADOS (CONSULTA)", style="Ghost.TButton", command=self._expand_view)\
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

        self.set_status("STATUS: Selecione uma programacao FINALIZADA e lance os pagantes (prestacao de contas).")

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

        # Regra estrita: somente quando motorista finalizou no app (status operacional).
        # Se a coluna existir, status local nao libera recebimentos.
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
                    "desktop/programacoes?modo=finalizadas_pendentes&limit=300",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                valores = [upper((r or {}).get("codigo_programacao") or "") for r in (arr or []) if isinstance(r, dict)]
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
                    status_where = "UPPER(TRIM(COALESCE(status_operacional,''))) IN ('FINALIZADA','FINALIZADO')"
                else:
                    status_where = "UPPER(TRIM(COALESCE(status,''))) IN ('FINALIZADA','FINALIZADO')"
                if has_finalizada_app:
                    status_where += " AND COALESCE(finalizada_no_app,0)=1"
                prest_where = " AND COALESCE(prestacao_status,'PENDENTE')='PENDENTE'" if has_prest else ""

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
                    WHERE UPPER(TRIM(COALESCE(status,''))) IN ('FINALIZADA','FINALIZADO')
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
                cur.execute(
                    """
                    SELECT data_saida, hora_saida, data_chegada, hora_chegada
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

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return False

        try:
            resposta = _call_api(
                "GET",
                f"desktop/rotas/{prog}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
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

    # -------------------------
    # Regras: bloqueio quando FECHADA
    # -------------------------
    def _is_prestacao_fechada(self, prog: str) -> bool:
        if not prog:
            return False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
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
            messagebox.showwarning("ATENÇÃO", "Esta prestação já está FECHADA. Não é possÃÂÂvel alterar recebimentos.")
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
                self.lbl_resumo_diaria.config(text=f"VALOR DIARIA: {fmt_money(diaria_motorista)}")
            self.lbl_resumo_qtd.config(text=f"QTD DIARIAS: {qtd:g}")
            self.lbl_resumo_mot.config(text=f"MOTORISTA: {fmt_money(total_mot)}")
            self.lbl_resumo_eqp.config(text=f"EQUIPE: {fmt_money(total_eqp)}")
            self.lbl_resumo_total.config(text=f"TOTAL: {fmt_money(total)}")
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

        if not self._rota_apt_para_recebimentos(prog):
            messagebox.showwarning(
                "ATENCAO",
                "Esse codigo ainda nao esta liberado para Recebimentos.\n\n"
                "Regra: so libera quando a rota estiver FINALIZADA no celular pelo motorista.",
            )
            self.refresh_comboboxes()
            return

        synced_api = False
        try:
            synced_api = self._sync_programacao_from_api(prog, silent=True)
        except Exception:
            logging.debug("Falha ignorada")

        # Revalida após sincronizar com API para impedir abertura indevida.
        if not self._rota_apt_para_recebimentos(prog):
            messagebox.showwarning(
                "ATENCAO",
                "Esse codigo nao esta mais liberado para Recebimentos apos sincronizacao.\n\n"
                "Regra: so libera quando a rota estiver FINALIZADA no celular e pendente de prestacao.",
            )
            self.refresh_comboboxes()
            return

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
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
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
                    rota = str(rota_obj.get("local_rota") or rota_obj.get("tipo_rota") or "")
                    diaria_motorista = safe_float(rota_obj.get("diaria_motorista_valor"), 0.0)
                    found_prog = True
            except Exception:
                api_failed = True
                logging.debug("Falha ao carregar programacao em recebimentos via API; usando fallback local.", exc_info=True)

        if api_enabled and api_checked and (not api_failed) and (not found_prog):
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
                col_rota = "local_rota" if "local_rota" in cols_prog else ("tipo_rota" if "tipo_rota" in cols_prog else "''")
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
            if synced_api:
                self.set_status(f"STATUS: Programacao carregada e sincronizada da API: {prog}")
            else:
                self.set_status(f"STATUS: Programacao carregada: {prog}")

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
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            resumo_diarias = self._sync_diarias_despesas(
                prog=self._current_prog,
                rota=self._rota_atual,
                equipe_raw=self._equipe_raw,
                data_saida=data_saida,
                hora_saida=hora_saida,
                data_chegada=data_chegada,
                hora_chegada=hora_chegada,
                diaria_motorista=diaria_motorista,
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
                messagebox.showerror("ERRO", f"Erro ao salvar dados: {str(e)}")
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
            return 1
        parts_nome = [p.strip() for p in re.split(r"[|,;/]+", nomes) if p.strip()]
        if parts_nome:
            return len(parts_nome)
        return 1

    def _sync_diarias_despesas(self, prog: str, rota: str, equipe_raw: str, data_saida: str, hora_saida: str, data_chegada: str, hora_chegada: str, diaria_motorista: float):
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
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota_obj = resp.get("rota") if isinstance(resp, dict) else None
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

        return {
            "qtd_diarias": qtd,
            "qtd_ajudantes": qtd_ajudantes,
            "diaria_motorista": diaria_motorista,
            "diaria_ajudante": diaria_ajudante,
            "total_motorista": total_mot,
            "total_ajudantes": total_ajud,
            "total_geral": total_geral,
        }
    

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
                rota_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rec_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/recebimentos",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = rota_resp.get("rota") if isinstance(rota_resp, dict) else None
                clientes_api = rota_resp.get("clientes") if isinstance(rota_resp, dict) else []
                rec_items = rec_resp.get("recebimentos") if isinstance(rec_resp, dict) else []
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
        if not forma:
            messagebox.showwarning("ATENÇÃO", "Informe a forma de pagamento.")
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
            messagebox.showerror("ERRO", f"Erro ao salvar recebimento: {str(e)}")

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
            messagebox.showerror("ERRO", f"Erro ao zerar recebimentos: {str(e)}")

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
            "Cliente Manual", "Digite o C?DIGO do cliente:",
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
            messagebox.showerror("ERRO", f"Erro ao inserir cliente: {str(e)}")
    

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
                rota_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rec_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/recebimentos",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = rota_resp.get("rota") if isinstance(rota_resp, dict) else None
                recebimentos = rec_resp.get("recebimentos") if isinstance(rec_resp, dict) else []
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
            cur.execute("""
                SELECT motorista, veiculo, equipe, num_nf, data_saida, hora_saida, data_chegada, hora_chegada
                FROM programacoes
                WHERE codigo_programacao=?
                LIMIT 1
            """, (prog,))
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

            cur.execute("""
                SELECT cod_cliente, nome_cliente, SUM(valor) AS total_valor,
                       (SELECT forma_pagamento
                          FROM recebimentos r2
                         WHERE r2.codigo_programacao = r.codigo_programacao
                           AND upper(r2.cod_cliente) = upper(r.cod_cliente)
                         ORDER BY datetime(r2.data_registro) DESC, r2.id DESC
                         LIMIT 1) AS forma_recente
                FROM recebimentos r
                WHERE codigo_programacao=?
                GROUP BY upper(cod_cliente), upper(nome_cliente)
                HAVING SUM(valor) > 0
                ORDER BY upper(nome_cliente) ASC
            """, (prog,))
            rows = cur.fetchall() or []

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
            draw_kv(col2_x, y, "VeÃÂÂculo", header["veiculo"])
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
            text="ðŸ”„ ATUALIZAR",
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
        calc_frame.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 10))
        ttk.Button(
            calc_frame,
            text="Calcular Saldo Automático",
            style="Ghost.TButton",
            width=24,
            command=self._calcular_saldo_auto
        ).pack(anchor="w")

        tab_nf.grid_columnconfigure(0, weight=3)
        tab_nf.grid_columnconfigure(1, weight=2)
        tab_nf.grid_rowconfigure(1, weight=1)

        nf_form_wrap = ttk.Frame(tab_nf, style="Card.TFrame")
        nf_form_wrap.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        nf_form_wrap.grid_columnconfigure(0, weight=1)
        nf_form_wrap.grid_columnconfigure(1, weight=1)

        nf_base = ttk.LabelFrame(nf_form_wrap, text="NF Base", padding=8)
        nf_base.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))

        row_nf = 0
        self.ent_nf_numero = self._create_field(nf_base, "Nº NOTA FISCAL:", row_nf, 14); row_nf += 1
        self.ent_nf_kg = self._create_field(nf_base, "KG NOTA FISCAL:", row_nf, 14); row_nf += 1
        self.ent_nf_preco = self._create_field(nf_base, "PRECO NF (R$/KG):", row_nf, 14); row_nf += 1
        self.ent_nf_caixas = self._create_field(nf_base, "CAIXAS:", row_nf, 14)

        nf_mov = ttk.LabelFrame(nf_form_wrap, text="Movimento", padding=8)
        nf_mov.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6))

        row_nf = 0
        self.ent_nf_kg_carregado = self._create_field(nf_mov, "KG CARREGADO:", row_nf, 14); row_nf += 1
        self.ent_nf_kg_vendido = self._create_field(nf_mov, "KG VENDIDO:", row_nf, 14); row_nf += 1
        self.ent_nf_saldo = self._create_field(nf_mov, "SALDO (KG):", row_nf, 14)

        nf_params = ttk.LabelFrame(nf_form_wrap, text="Parametros e Sincronismo", padding=8)
        nf_params.grid(row=1, column=0, columnspan=2, sticky="ew")

        row_nf = 0
        self.ent_nf_media_carregada = self._create_field(nf_params, "MEDIA CARREGADA:", row_nf, 14); row_nf += 1
        self.ent_nf_caixa_final = self._create_field(nf_params, "CAIXA FINAL:", row_nf, 14)
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
            width=24,
            command=self._sincronizar_com_app
        ).grid(row=row_nf + 1, column=0, columnspan=2, sticky="w", pady=(10, 0))

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
        ttk.Label(nf_resumo, text="OCORRÊNCIAS (TRANSBORDO)", font=("Segoe UI", 10, "bold")).grid(row=16, column=0, sticky="w", pady=(0, 6))
        self.lbl_nf_mort_aves = ttk.Label(nf_resumo, text="Mortalidade (aves): 0")
        self.lbl_nf_mort_aves.grid(row=17, column=0, sticky="w")
        self.lbl_nf_mort_kg = ttk.Label(nf_resumo, text="Mortalidade (KG): 0,00")
        self.lbl_nf_mort_kg.grid(row=18, column=0, sticky="w")
        self.lbl_nf_kg_util = ttk.Label(nf_resumo, text="KG útil NF (NF - mortalidade): 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_nf_kg_util.grid(row=19, column=0, sticky="w")
        self.lbl_nf_obs_transb = ttk.Label(nf_resumo, text="Obs transbordo: -", wraplength=320, justify="left")
        self.lbl_nf_obs_transb.grid(row=20, column=0, sticky="w", pady=(2, 0))

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

        ttk.Label(km_side, text="ÚLTIMOS KM POR VEÃÂÂCULO", style="CardTitle.TLabel").grid(row=2, column=0, sticky="w", pady=(4, 4))
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
            text="ðÅ¸§® Distribuir",
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

        kpi_res = ttk.LabelFrame(kpi_row, text=" RESULTADO LÃÂÂQUIDO ", padding=5)
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
        self.ent_adiantamento.bind("<KeyRelease>", lambda e: self._update_resumo_financeiro())
        self._bind_money_entry(self.ent_adiantamento)

        ttk.Label(entrada_frame, text="Total Entradas:", font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 1))
        self.lbl_total_entradas = ttk.Label(entrada_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"), foreground="#2E7D32")
        self.lbl_total_entradas.grid(row=2, column=1, sticky="e", pady=(4, 1))

        saida_frame = ttk.LabelFrame(details_wrap, text=" SAÃÂÂDAS ", padding=5)
        saida_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(3, 0))
        saida_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(saida_frame, text="Despesas Total:", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=1)
        self.lbl_desp_total = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_desp_total.grid(row=0, column=1, sticky="e", pady=1)

        ttk.Label(saida_frame, text="Contagem C?dulas:", font=("Segoe UI", 8)).grid(row=1, column=0, sticky="w", pady=1)
        self.lbl_cedulas_total = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_cedulas_total.grid(row=1, column=1, sticky="e", pady=1)

        ttk.Label(saida_frame, text="Total Sa?das:", font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 1))
        self.lbl_total_saidas = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"), foreground="#C62828")
        self.lbl_total_saidas.grid(row=2, column=1, sticky="e", pady=(4, 1))

        resultado_frame = ttk.LabelFrame(details_wrap, text=" RESULTADOS ", padding=5)
        resultado_frame.grid(row=0, column=1, rowspan=2, sticky="nsew")
        resultado_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(resultado_frame, text="Valor p/ Caixa:", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="w", pady=1)
        self.lbl_valor_final_caixa = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_valor_final_caixa.grid(row=0, column=1, sticky="e", pady=1)

        ttk.Label(resultado_frame, text="Diferen?a (Caixa - C?d):", font=("Segoe UI", 8)).grid(row=1, column=0, sticky="w", pady=1)
        self.lbl_diferenca = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 9, "bold"))
        self.lbl_diferenca.grid(row=1, column=1, sticky="e", pady=1)

        ttk.Label(resultado_frame, text="Resultado L?quido:", font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 1))
        self.lbl_resultado_liquido = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 11, "bold"))
        self.lbl_resultado_liquido.grid(row=2, column=1, sticky="e", pady=(4, 1))

        botoes_frame = ttk.Frame(page_card, style="Card.TFrame")
        botoes_frame.grid(row=4, column=0, sticky="ew")

        for i in range(6):
            botoes_frame.grid_columnconfigure(i, weight=1)

        botoes = [
            ("â¬â€¦ VOLTAR", "Ghost.TButton", self._voltar_recebimentos),
            ("âÅ¾â€¢ ADICIONAR DESPESA", "Warn.TButton", self._open_registrar_rapido),
            ("ðÅ¸â€“¨ IMPRIMIR PDF", "Ghost.TButton", self.imprimir_resumo),
            ("ðÅ¸â€™¾ SALVAR", "Primary.TButton", self.salvar_tudo),
            ("âÅ“Âï¸Â EDITAR", "Warn.TButton", self._editar_linha_selecionada),
            ("ðÅ¸ÂÂ FINALIZAR", "Danger.TButton", self.finalizar_prestacao_despesas),
        ]

        for i, (texto, estilo, comando) in enumerate(botoes):
            btn = ttk.Button(botoes_frame, text=texto, style=estilo, command=comando, width=12)
            btn.grid(row=0, column=i, sticky="ew", padx=1, ipady=2)

        # carrega dados
        self.refresh_comboboxes()
        self._refresh_all()

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
                    cur.execute("""
                        SELECT COALESCE(veiculo, '-'), COALESCE(media_km_l, 0)
                        FROM programacoes
                        WHERE COALESCE(media_km_l, 0) > 0
                        ORDER BY media_km_l DESC, id DESC
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
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/despesas",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            itens = (resp.get("despesas") if isinstance(resp, dict) else []) or []
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
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
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
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        loaded_api = False
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                rec_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/recebimentos",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                desp_resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}/despesas",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rec_rows = (rec_resp.get("recebimentos") if isinstance(rec_resp, dict) else []) or []
                desp_rows = (desp_resp.get("despesas") if isinstance(desp_resp, dict) else []) or []
                total_receb = sum(
                    safe_float((r or {}).get("valor"), 0.0)
                    for r in rec_rows
                    if isinstance(r, dict)
                )
                total_desp = sum(
                    safe_float((r or {}).get("valor"), 0.0)
                    for r in desp_rows
                    if isinstance(r, dict)
                )
                loaded_api = True
            except Exception:
                logging.debug("Falha ao calcular resumo financeiro via API; usando fallback local.", exc_info=True)

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
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            if desktop_secret and is_desktop_api_sync_enabled():
                try:
                    resp = _call_api(
                        "GET",
                        f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    rota = resp.get("rota") if isinstance(resp, dict) else None
                    clientes = resp.get("clientes") if isinstance(resp, dict) else []
                    if isinstance(rota, dict):
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
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                clientes = resp.get("clientes") if isinstance(resp, dict) else []
                precos = []
                for r in (clientes or []):
                    if not isinstance(r, dict):
                        continue
                    p = safe_float((r or {}).get("preco_atual"), safe_float((r or {}).get("preco"), 0.0))
                    if p > 0:
                        precos.append(p)
                if precos:
                    return sum(precos) / float(len(precos))
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
                desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                loaded_api = False
                if desktop_secret and is_desktop_api_sync_enabled():
                    try:
                        resp = _call_api(
                            "GET",
                            f"desktop/rotas/{urllib.parse.quote(upper(prog))}/despesas",
                            extra_headers={"X-Desktop-Secret": desktop_secret},
                        )
                        arr = (resp.get("despesas") if isinstance(resp, dict) else []) or []
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
                    desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                    loaded_api = False
                    if desktop_secret and is_desktop_api_sync_enabled():
                        try:
                            resp = _call_api(
                                "GET",
                                f"desktop/rotas/{urllib.parse.quote(upper(self._current_programacao))}/despesas",
                                extra_headers={"X-Desktop-Secret": desktop_secret},
                            )
                            arr = (resp.get("despesas") if isinstance(resp, dict) else []) or []
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

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret:
            return False

        try:
            resposta = _call_api(
                "GET",
                f"desktop/rotas/{prog}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
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
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog or '').strip())}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
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
            raise SyncError("Não foi possÃÂÂvel obter token da API de sincronização.")
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
                    "desktop/programacoes?modo=finalizadas_prestacao&limit=300",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                programas = []
                for r in (arr or []):
                    if not isinstance(r, dict):
                        continue
                    codigo = upper((r or {}).get("codigo_programacao") or "")
                    data = str((r or {}).get("data_criacao") or "")[:10] or "Sem data"
                    if codigo:
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
                    part = f"{status_op_expr} IN ('FINALIZADA','FINALIZADO')"
                    if has_finalizada_app:
                        part = f"({part} AND {finalizada_expr}=1)"
                    where_parts.append(part)
                if has_status:
                    where_parts.append(f"{status_base_expr} IN ('FINALIZADA','FINALIZADO')")
                where_sql = " OR ".join(where_parts) if where_parts else "1=0"

                prest_where = ""
                if has_prest:
                    prest_where = f" AND {prest_expr} IN ('PENDENTE','FECHADA')"

                cur.execute(
                    f"""
                    SELECT codigo_programacao, data_criacao
                    FROM programacoes
                    WHERE ({where_sql}){prest_where}
                    ORDER BY id DESC
                    LIMIT 300
                    """
                )
            except Exception:
                cur.execute("SELECT codigo_programacao, data_criacao FROM programacoes ORDER BY id DESC LIMIT 300")

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
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                rota = resp.get("rota") if isinstance(resp, dict) else None
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

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    f"desktop/rotas/{urllib.parse.quote(prog)}/logistica",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                info = resp.get("logistica") if isinstance(resp, dict) else None
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
                "Não é possÃÂvel finalizar a prestação:\n"
                "existe substituição de rota pendente de aceite/recusa.",
            )
        if safe_int(info.get("pend_transferencia"), 0) > 0:
            return (
                False,
                "Não é possÃÂvel finalizar a prestação:\n"
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
                    "Sera gerada a folha completa com recebimentos e despesas."
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
            messagebox.showerror("ERRO", f"Erro ao finalizar prestacao: {str(e)}")
    

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
                messagebox.showerror("ERRO", f"Erro ao salvar despesa: {str(e)}")

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        ttk.Button(btn_frame, text="ðÅ¸â€™¾ SALVAR", style="Primary.TButton", command=salvar)\
            .grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btn_frame, text="âÅ“â€“ CANCELAR", style="Ghost.TButton", command=win.destroy)\
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
                messagebox.showerror("ERRO", f"Erro ao atualizar despesa: {str(e)}")

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        ttk.Button(btn_frame, text="ðÅ¸â€™¾ SALVAR", style="Primary.TButton", command=salvar)\
            .grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btn_frame, text="âÅ“â€“ CANCELAR", style="Ghost.TButton", command=win.destroy)\
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
            messagebox.showerror("ERRO", f"Erro ao excluir despesa: {str(e)}")

    # =========================================================
    # 7.7 RELATÓRIO EM TELA + IMPRESSÃO SIMULADA
    # =========================================================
    def imprimir_resumo(self):
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
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            if prog and desktop_secret and is_desktop_api_sync_enabled():
                try:
                    rota_resp = _call_api(
                        "GET",
                        f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    rec_resp = _call_api(
                        "GET",
                        f"desktop/rotas/{urllib.parse.quote(upper(prog))}/recebimentos",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    desp_resp = _call_api(
                        "GET",
                        f"desktop/rotas/{urllib.parse.quote(upper(prog))}/despesas",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )

                    rota = rota_resp.get("rota") if isinstance(rota_resp, dict) else None
                    recebimentos_api = rec_resp.get("recebimentos") if isinstance(rec_resp, dict) else []
                    despesas_api = desp_resp.get("despesas") if isinstance(desp_resp, dict) else []

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
                except Exception:
                    logging.debug("Falha ao carregar resumo via API; usando fallback local.", exc_info=True)

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
                        COALESCE(data_saida,''), COALESCE(hora_saida,''),
                        COALESCE(data_chegada,''), COALESCE(hora_chegada,''),
                        COALESCE(nf_kg,0), COALESCE(nf_caixas,0), COALESCE(nf_kg_carregado,0),
                        COALESCE(nf_kg_vendido,0), COALESCE(nf_saldo,0),
                        {nf_preco_col},
                        {media_carreg_col},
                        {caixa_final_col},
                        COALESCE(km_inicial,0), COALESCE(km_final,0), COALESCE(litros,0),
                        COALESCE(km_rodado,0), COALESCE(media_km_l,0), COALESCE(custo_km,0),
                        COALESCE(ced_200_qtd,0), COALESCE(ced_100_qtd,0), COALESCE(ced_50_qtd,0),
                        COALESCE(ced_20_qtd,0), COALESCE(ced_10_qtd,0), COALESCE(ced_5_qtd,0), COALESCE(ced_2_qtd,0),
                        COALESCE(valor_dinheiro,0),
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

                cur.execute("""
                    SELECT descricao, valor, COALESCE(categoria,'OUTROS'), COALESCE(observacao,''), data_registro
                    FROM despesas
                    WHERE codigo_programacao=?
                    ORDER BY data_registro DESC
                """, (prog,))
                despesas = cur.fetchall() or []

                cur.execute("""
                    SELECT
                        COALESCE(cod_cliente,''),
                        COALESCE(nome_cliente,''),
                        COALESCE(valor,0),
                        COALESCE(forma_pagamento,''),
                        COALESCE(observacao,''),
                        COALESCE(data_registro,'')
                    FROM recebimentos
                    WHERE codigo_programacao=?
                    ORDER BY data_registro DESC, id DESC
                """, (prog,))
                recebimentos = cur.fetchall() or []

                cur.execute("SELECT SUM(valor) FROM recebimentos WHERE codigo_programacao=?", (prog,))
                total_receb = safe_float((cur.fetchone() or [0])[0], 0.0)

                cur.execute("SELECT SUM(valor) FROM despesas WHERE codigo_programacao=?", (prog,))
                total_desp = safe_float((cur.fetchone() or [0])[0], 0.0)

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

            _draw_receb_header()

            total_receb_detalhe = 0.0
            if not recebimentos:
                c.rect(table_x, y - row_h_rec + 1, table_w, row_h_rec, stroke=1, fill=0)
                c.drawString(x_cod + 2, y - row_h_rec + 3, "SEM RECEBIMENTOS REGISTRADOS")
                y -= row_h_rec
            else:
                for cod_cli, nome_cli, valor_rec, forma_rec, obs_rec, data_rec in recebimentos:
                    if y < bottom + 45 * mm:
                        c.showPage()
                        y = height - top
                        c.setFont("Helvetica-Bold", 12)
                        c.drawString(left, y, f"FOLHA DE RECEBIMENTOS - PROGRAMACAO {prog} (CONT.)")
                        y -= 8 * mm
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
            col_desc = table_w * 0.38
            col_cat = table_w * 0.18
            col_val = table_w * 0.14
            col_obs = table_w * 0.30

            x_desc = table_x
            x_cat = x_desc + col_desc
            x_val = x_cat + col_cat
            x_obs = x_val + col_val

            row_h = 6.5 * mm

            c.setFont("Helvetica-Bold", 9)
            c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            c.drawString(x_desc + 2, y - row_h + 3, "DESCRICAO")
            c.drawString(x_cat + 2, y - row_h + 3, "CATEGORIA")
            c.drawRightString(x_val + col_val - 2, y - row_h + 3, "VALOR")
            c.drawString(x_obs + 2, y - row_h + 3, "OBS")
            y -= row_h

            c.setFont("Helvetica", 9)
            for desc, val, cat, obs, data_reg in despesas:
                ensure_space(15)
                desc_up = upper(str(desc or "").strip())
                obs_txt = str(obs or "").strip()
                if desc_up == "DIARIAS MOTORISTA":
                    m_nome = motorista or "-"
                    qtd_d = "-"
                    m = re.search(r"QTD\s*DIARIAS\s*:\s*([^|]+)", obs_txt, flags=re.IGNORECASE)
                    if m:
                        qtd_d = m.group(1).strip()
                    obs_txt = f"QTD DIARIAS: {qtd_d} | MOTORISTA: {m_nome}"
                elif desc_up == "DIARIAS AJUDANTES":
                    a_nome = equipe_txt or "-"
                    qtd_d = "-"
                    m = re.search(r"QTD\s*DIARIAS\s*:\s*([^|]+)", obs_txt, flags=re.IGNORECASE)
                    if m:
                        qtd_d = m.group(1).strip()
                    obs_txt = f"QTD DIARIAS: {qtd_d} | AJUDANTES: {a_nome}"
                c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
                c.drawString(x_desc + 2, y - row_h + 3, str(desc or "")[:40])
                c.drawString(x_cat + 2, y - row_h + 3, str(cat or "")[:14])
                c.drawRightString(x_val + col_val - 2, y - row_h + 3, f"{float(val or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c.drawString(x_obs + 2, y - row_h + 3, _clip_width(obs_txt, col_obs - 4, "Helvetica", 9))
                y -= row_h

            y -= 6 * mm

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

            # Assinaturas
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
                desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
                if desktop_secret and is_desktop_api_sync_enabled():
                    try:
                        desp_resp = _call_api(
                            "GET",
                            f"desktop/rotas/{urllib.parse.quote(upper(prog))}/despesas",
                            extra_headers={"X-Desktop-Secret": desktop_secret},
                        )
                        despesas_api = desp_resp.get("despesas") if isinstance(desp_resp, dict) else []
                        total_despesas_prog = sum(
                            safe_float((d.get("valor") if isinstance(d, dict) else 0), 0.0)
                            for d in (despesas_api or [])
                        )
                        used_api = True
                    except Exception:
                        logging.debug("Falha ao calcular total de despesas via API; usando fallback local.", exc_info=True)

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
                logging.debug("Falha ao salvar dados financeiros via API; usando fallback local.", exc_info=True)

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
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            if desktop_secret and is_desktop_api_sync_enabled():
                try:
                    rota_resp = _call_api(
                        "GET",
                        f"desktop/rotas/{urllib.parse.quote(upper(prog))}",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    rec_resp = _call_api(
                        "GET",
                        f"desktop/rotas/{urllib.parse.quote(upper(prog))}/recebimentos",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    desp_resp = _call_api(
                        "GET",
                        f"desktop/rotas/{urllib.parse.quote(upper(prog))}/despesas",
                        extra_headers={"X-Desktop-Secret": desktop_secret},
                    )
                    rota = rota_resp.get("rota") if isinstance(rota_resp, dict) else None
                    recebimentos = rec_resp.get("recebimentos") if isinstance(rec_resp, dict) else []
                    despesas = desp_resp.get("despesas") if isinstance(desp_resp, dict) else []
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
                except Exception:
                    logging.debug("Falha ao carregar dados para exportacao via API; usando fallback local.", exc_info=True)

            if not used_api:
                with get_db() as conn:
                    try:
                        df_prog = pd.read_sql_query("""
                            SELECT codigo_programacao, motorista, veiculo, data_criacao,
                                   nf_numero, nf_kg, nf_caixas, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                                   km_inicial, km_final, litros, km_rodado, media_km_l, custo_km,
                                   valor_dinheiro,
                                   adiantamento
                            FROM programacoes
                            WHERE codigo_programacao=?
                        """, conn, params=(prog,))
                    except Exception:
                        df_prog = pd.read_sql_query("""
                            SELECT codigo_programacao, motorista, veiculo, data_criacao,
                                   nf_numero, nf_kg, nf_caixas, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                                   km_inicial, km_final, litros, km_rodado, media_km_l, custo_km,
                                   valor_dinheiro,
                                   adiantamento_rota
                            FROM programacoes
                            WHERE codigo_programacao=?
                        """, conn, params=(prog,))

                    if df_prog.empty:
                        df_prog = pd.read_sql_query("""
                            SELECT codigo_programacao, motorista, veiculo, data_criacao,
                                   nf_numero, nf_kg, nf_caixas, nf_kg_carregado, nf_kg_vendido, nf_saldo,
                                   km_inicial, km_final, litros, km_rodado, media_km_l, custo_km,
                                   valor_dinheiro
                            FROM programacoes
                            WHERE codigo_programacao=?
                        """, conn, params=(prog,))

                    df_despesas = pd.read_sql_query("""
                        SELECT id, descricao, valor, categoria, observacao, data_registro
                        FROM despesas
                        WHERE codigo_programacao=?
                        ORDER BY data_registro DESC
                    """, conn, params=(prog,))

                    try:
                        df_receb = pd.read_sql_query("""
                            SELECT cod_cliente, nome_cliente, valor, forma_pagamento, observacao, num_nf, data_registro
                            FROM recebimentos
                            WHERE codigo_programacao=?
                            ORDER BY data_registro DESC
                        """, conn, params=(prog,))
                    except Exception:
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
                f"Resultado LÃÂÂquido: {self._fmt_money(resultado_liquido) if hasattr(self, '_fmt_money') else resultado_liquido}"
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

        ttk.Button(filtros, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton", command=self.refresh_data).grid(
            row=1, column=2, sticky="w"
        )

        resumo = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        resumo.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        resumo.grid_columnconfigure(0, weight=1)
        ttk.Label(resumo, text="Resumo de distribuicao", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        kpi_row = ttk.Frame(resumo, style="Card.TFrame")
        kpi_row.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        for c in range(4):
            kpi_row.grid_columnconfigure(c, weight=1)

        kpi_1 = ttk.LabelFrame(kpi_row, text=" Rotas no PerÃÂodo ", padding=8)
        kpi_1.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.lbl_esc_kpi_rotas = ttk.Label(kpi_1, text="0", font=("Segoe UI", 13, "bold"), foreground="#1D4ED8")
        self.lbl_esc_kpi_rotas.grid(row=0, column=0, sticky="w")

        kpi_2 = ttk.LabelFrame(kpi_row, text=" Motoristas ", padding=8)
        kpi_2.grid(row=0, column=1, sticky="ew", padx=6)
        self.lbl_esc_kpi_mot = ttk.Label(kpi_2, text="0", font=("Segoe UI", 13, "bold"), foreground="#0F766E")
        self.lbl_esc_kpi_mot.grid(row=0, column=0, sticky="w")

        kpi_3 = ttk.LabelFrame(kpi_row, text=" KM Médio/Motorista ", padding=8)
        kpi_3.grid(row=0, column=2, sticky="ew", padx=6)
        self.lbl_esc_kpi_km = ttk.Label(kpi_3, text="0,0", font=("Segoe UI", 13, "bold"), foreground="#B45309")
        self.lbl_esc_kpi_km.grid(row=0, column=0, sticky="w")

        kpi_4 = ttk.LabelFrame(kpi_row, text=" Horas Médias/Motorista ", padding=8)
        kpi_4.grid(row=0, column=3, sticky="ew", padx=(6, 0))
        self.lbl_esc_kpi_horas = ttk.Label(kpi_4, text="0,00", font=("Segoe UI", 13, "bold"), foreground="#7C3AED")
        self.lbl_esc_kpi_horas.grid(row=0, column=0, sticky="w")

        txt_row = ttk.Frame(resumo, style="Card.TFrame")
        txt_row.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        txt_row.grid_columnconfigure(0, weight=1)
        txt_row.grid_columnconfigure(1, weight=1)

        box_resumo = ttk.LabelFrame(txt_row, text=" Resumo Operacional ", padding=(10, 8))
        box_resumo.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        box_resumo.grid_columnconfigure(0, weight=1)

        box_reco = ttk.LabelFrame(txt_row, text=" Recomendações ", padding=(10, 8))
        box_reco.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        box_reco.grid_columnconfigure(0, weight=1)

        self.lbl_resumo = ttk.Label(
            box_resumo,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=480,
        )
        self.lbl_resumo.grid(row=0, column=0, sticky="nw")
        self.lbl_recomendacoes = ttk.Label(
            box_reco,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=480,
            foreground="#1E3A8A",
        )
        self.lbl_recomendacoes.grid(row=0, column=0, sticky="nw")

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
            cv.create_text(12, 12, anchor="nw", text="Sem dados para o perÃÂodo selecionado.", fill="#6B7280", font=("Segoe UI", 9))
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
        return f"Recomendações (local-alvo: {local_alvo}):\nââ‚¬¢ " + "\nââ‚¬¢ ".join(recs)

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
        if desktop_secret and is_desktop_api_sync_enabled():
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
                data_saida_expr = "COALESCE(data_saida,'')" if "data_saida" in cols else "''"
                hora_saida_expr = "COALESCE(hora_saida,'')" if "hora_saida" in cols else "''"
                data_chegada_expr = "COALESCE(data_chegada,'')" if "data_chegada" in cols else "''"
                hora_chegada_expr = "COALESCE(hora_chegada,'')" if "hora_chegada" in cols else "''"
                data_ref_expr = (
                    "COALESCE(data_saida,data_criacao,data,'')"
                    if "data_saida" in cols or "data_criacao" in cols or "data" in cols
                    else "''"
                )
                cur.execute(
                    f"""
                    SELECT codigo_programacao,
                           {data_ref_expr} AS data_ref,
                           motorista,
                           equipe,
                           COALESCE(status,'') AS status,
                           {status_op_expr} AS status_operacional,
                           {finalizada_app_expr} AS finalizada_no_app,
                           kg_estimado,
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
                    f"NÃÂvel da escala: {nivel} | Maior carga atual: {mais_sobrecarregado}\n"
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


class CentroCustosPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Centro de Custos")
        self.body.grid_rowconfigure(3, weight=1)
        self.body.grid_columnconfigure(0, weight=1)
        self._chart_labels = []
        self._chart_values = []

        filtros = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        filtros.grid(row=0, column=0, sticky="ew")
        filtros.grid_columnconfigure(8, weight=1)

        ttk.Label(filtros, text="PerÃÂodo (dias)", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.var_periodo = tk.StringVar(value="30")
        self.cb_periodo = ttk.Combobox(
            filtros, textvariable=self.var_periodo, state="readonly", width=10,
            values=["7", "15", "30", "60", "90", "180", "TODAS"]
        )
        self.cb_periodo.grid(row=1, column=0, sticky="w", padx=(0, 10))

        ttk.Label(filtros, text="VeÃÂculo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.var_veiculo = tk.StringVar(value="TODOS")
        self.cb_veiculo = ttk.Combobox(
            filtros, textvariable=self.var_veiculo, state="readonly", width=16, values=["TODOS"]
        )
        self.cb_veiculo.grid(row=1, column=1, sticky="w", padx=(0, 10))

        ttk.Button(filtros, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton", command=self.refresh_data).grid(
            row=1, column=2, sticky="w", padx=(0, 6)
        )

        ttk.Label(filtros, text="Métrica do gráfico", style="CardLabel.TLabel").grid(row=0, column=3, sticky="w")
        self.var_chart_metric = tk.StringVar(value="CUSTO_KM")
        self.cb_chart_metric = ttk.Combobox(
            filtros,
            textvariable=self.var_chart_metric,
            state="readonly",
            width=16,
            values=["CUSTO_KM", "CUSTO_KG", "DESPESA_TOTAL"],
        )
        self.cb_chart_metric.grid(row=1, column=3, sticky="w", padx=(0, 10))

        self.lbl_resumo = ttk.Label(
            self.body,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=980,
        )
        self.lbl_resumo.grid(row=1, column=0, sticky="ew", pady=(8, 8))

        chart_wrap = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        chart_wrap.grid(row=2, column=0, sticky="ew")
        chart_wrap.grid_columnconfigure(0, weight=1)
        self.lbl_chart_title = ttk.Label(chart_wrap, text="Custo por VeÃÂculo (Custo/KM)", style="CardTitle.TLabel")
        self.lbl_chart_title.grid(row=0, column=0, sticky="w")
        self.cv_chart = tk.Canvas(chart_wrap, height=220, bg="white", highlightthickness=1, highlightbackground="#E5E7EB")
        self.cv_chart.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.cv_chart.bind("<Configure>", lambda _e: self._draw_chart(self._chart_labels, self._chart_values))

        table_wrap = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        table_wrap.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        cols = ("VEICULO", "ROTAS", "KM_RODADO", "KG_CARREGADO", "DESPESAS", "CUSTO_KM", "CUSTO_KG", "TICKET_ROTA")
        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

        for c in cols:
            self.tree.heading(c, text=c)
            if c == "VEICULO":
                w = 160
            elif c in ("DESPESAS", "CUSTO_KM", "CUSTO_KG", "TICKET_ROTA"):
                w = 130
            else:
                w = 110
            self.tree.column(c, anchor="w" if c == "VEICULO" else "center", width=w)
        enable_treeview_sorting(
            self.tree,
            numeric_cols={"ROTAS", "KM_RODADO", "KG_CARREGADO", "DESPESAS", "CUSTO_KM", "CUSTO_KG", "TICKET_ROTA"},
        )

        self.tree.tag_configure("bom", foreground="#14532D", background="#ECFDF3")
        self.tree.tag_configure("medio", foreground="#7C2D12", background="#FFEDD5")
        self.tree.tag_configure("alto", foreground="#7F1D1D", background="#FEE2E2")

        self.cb_periodo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())
        self.cb_veiculo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())
        self.cb_chart_metric.bind("<<ComboboxSelected>>", lambda _e: self.refresh_data())

    def on_show(self):
        self.refresh_comboboxes()
        self.refresh_data()

    def _parse_data(self, raw: str):
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

    def _draw_chart(self, labels, values):
        try:
            cv = self.cv_chart
            cv.delete("all")
            w = max(cv.winfo_width(), 10)
            h = max(cv.winfo_height(), 10)
            if (not labels) or (not values):
                cv.create_text(w / 2, h / 2, text="Sem dados para gráfico", fill="#6B7280", font=("Segoe UI", 10))
                return

            n = min(len(labels), len(values), 10)
            labels = labels[:n]
            vals = [max(safe_float(v, 0.0), 0.0) for v in values[:n]]
            vmax = max(vals) if vals else 1.0
            if vmax <= 0:
                vmax = 1.0

            metric = "CUSTO_KM"
            try:
                metric = upper(str(self.var_chart_metric.get() or "").strip()) or "CUSTO_KM"
            except Exception:
                metric = "CUSTO_KM"

            def _fmt_axis(v: float) -> str:
                vv = safe_float(v, 0.0)
                if metric == "DESPESA_TOTAL":
                    return f"R$ {vv:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{vv:.3f}"

            def _fmt_bar(v: float) -> str:
                vv = safe_float(v, 0.0)
                if metric == "DESPESA_TOTAL":
                    return f"R$ {vv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{vv:.3f}"

            palette = ["#F68D56", "#F2C96D", "#D7E64A", "#63D04A", "#67D08D", "#58C6C4", "#5E9AC8", "#A78BFA", "#F472B6", "#22C55E"]
            left, right, top, bottom = 50, 16, 30, 28
            plot_w = max(w - left - right, 40)
            plot_h = max(h - top - bottom, 30)
            x0, y0 = left, top
            x1, y1 = left + plot_w, top + plot_h

            # Grade
            ticks = 5
            for i in range(ticks + 1):
                v = (vmax / ticks) * i
                yy = y1 - (v / vmax) * plot_h
                cv.create_line(x0, yy, x1, yy, fill="#E5E7EB")
                cv.create_text(x0 - 6, yy, text=_fmt_axis(v), anchor="e", fill="#4B5563", font=("Segoe UI", 8, "bold"))

            cv.create_line(x0, y0, x0, y1, fill="#374151")
            cv.create_line(x0, y1, x1, y1, fill="#374151", width=2)

            gap = 10
            bw = max((plot_w - gap * (n + 1)) / n, 10)
            for i, (lb, vv) in enumerate(zip(labels, vals)):
                c = palette[i % len(palette)]
                bx1 = x0 + gap + i * (bw + gap)
                bx2 = bx1 + bw
                bh = (vv / vmax) * plot_h
                by1 = y1 - bh
                cv.create_rectangle(bx1, by1, bx2, y1, fill=c, outline="")
                cv.create_text((bx1 + bx2) / 2, y1 + 11, text=str(lb)[:10], fill="#111827", font=("Segoe UI", 8))
                if bh > 16:
                    cv.create_text((bx1 + bx2) / 2, by1 - 8, text=_fmt_bar(vv), fill="#111827", font=("Segoe UI", 7, "bold"))
        except Exception:
            logging.debug("Falha ignorada")

    def refresh_comboboxes(self):
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/programacoes?modo=todas&limit=1200",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                vals = sorted(
                    {
                        upper((r or {}).get("veiculo") or "")
                        for r in (arr or [])
                        if isinstance(r, dict) and str((r or {}).get("veiculo") or "").strip()
                    }
                )
                values = ["TODOS"] + vals
                self.cb_veiculo.configure(values=values)
                if self.var_veiculo.get() not in values:
                    self.var_veiculo.set("TODOS")
                return
            except Exception:
                logging.debug("Falha ao carregar filtro de veiculos via API", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("""
                    SELECT DISTINCT upper(trim(COALESCE(veiculo,'')))
                    FROM programacoes
                    WHERE trim(COALESCE(veiculo,'')) <> ''
                    ORDER BY 1
                """)
                vals = [r[0] for r in (cur.fetchall() or []) if str(r[0] or "").strip()]
            except Exception:
                vals = []
        values = ["TODOS"] + vals
        self.cb_veiculo.configure(values=values)
        if self.var_veiculo.get() not in values:
            self.var_veiculo.set("TODOS")

    def refresh_data(self):
        periodo = upper(self.var_periodo.get().strip())
        veiculo_sel = upper(self.var_veiculo.get().strip())
        cutoff = None
        if periodo != "TODAS":
            try:
                cutoff = datetime.now() - timedelta(days=int(periodo))
            except Exception:
                cutoff = None

        rows = []
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/centro-custos/rows?limit=5000",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("rows") if isinstance(resp, dict) else []
                rows = [
                    (
                        (r or {}).get("codigo_programacao", ""),
                        (r or {}).get("veiculo", ""),
                        (r or {}).get("data_ref", ""),
                        (r or {}).get("km_rodado", 0),
                        (r or {}).get("kg_carregado", 0),
                        (r or {}).get("total_desp", 0),
                    )
                    for r in (arr or [])
                    if isinstance(r, dict)
                ]
            except Exception:
                logging.debug("Falha ao carregar centro de custos via API", exc_info=True)

        if not rows:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(r[1]).lower() for r in (cur.fetchall() or [])}

                data_expr = "COALESCE(data_saida,data_criacao,'')"
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
                kg_expr = "COALESCE(nf_kg_carregado, kg_carregado, 0)" if "nf_kg_carregado" in cols else ("COALESCE(kg_carregado,0)" if "kg_carregado" in cols else "0")

                cur.execute(
                    f"""
                    SELECT p.codigo_programacao,
                           upper(trim(COALESCE(p.veiculo,''))) as veiculo,
                           {data_expr} as data_ref,
                           {km_expr} as km_rodado,
                           {kg_expr} as kg_carregado,
                           COALESCE(d.total_desp,0) as total_desp
                      FROM programacoes p
                 LEFT JOIN (
                        SELECT codigo_programacao, COALESCE(SUM(valor),0) as total_desp
                          FROM despesas
                         GROUP BY codigo_programacao
                    ) d ON d.codigo_programacao = p.codigo_programacao
                     WHERE trim(COALESCE(p.veiculo,'')) <> ''
                    """
                )
                rows = cur.fetchall() or []

        agg = {}
        for r in rows:
            codigo = str(r[0] or "").strip()
            veiculo = upper(str(r[1] or "").strip())
            dt_ref = self._parse_data(r[2])
            km = safe_float(r[3], 0.0)
            kg = safe_float(r[4], 0.0)
            desp = safe_float(r[5], 0.0)

            if cutoff is not None and dt_ref is not None and dt_ref < cutoff:
                continue
            if veiculo_sel and veiculo_sel != "TODOS" and veiculo != veiculo_sel:
                continue
            if not veiculo:
                continue

            d = agg.setdefault(veiculo, {"rotas": 0, "km": 0.0, "kg": 0.0, "desp": 0.0})
            d["rotas"] += 1
            d["km"] += km
            d["kg"] += kg
            d["desp"] += desp

        self.tree.delete(*self.tree.get_children())
        rows_out = []
        for veic, d in sorted(agg.items(), key=lambda kv: (-safe_float(kv[1].get("desp", 0), 0.0), kv[0])):
            rotas = safe_int(d["rotas"], 0)
            km = safe_float(d["km"], 0.0)
            kg = safe_float(d["kg"], 0.0)
            desp = safe_float(d["desp"], 0.0)
            custo_km = (desp / km) if km > 0 else 0.0
            custo_kg = (desp / kg) if kg > 0 else 0.0
            ticket = (desp / rotas) if rotas > 0 else 0.0
            rows_out.append((veic, rotas, km, kg, desp, custo_km, custo_kg, ticket))

        media_custo_km = (
            sum(r[5] for r in rows_out) / float(len(rows_out))
            if rows_out else 0.0
        )
        for veic, rotas, km, kg, desp, custo_km, custo_kg, ticket in rows_out:
            if media_custo_km <= 0:
                tag = "medio"
            elif custo_km > media_custo_km * 1.2:
                tag = "alto"
            elif custo_km < media_custo_km * 0.9:
                tag = "bom"
            else:
                tag = "medio"
            tree_insert_aligned(
                self.tree, "", "end",
                (
                    veic,
                    rotas,
                    round(km, 1),
                    round(kg, 2),
                    round(desp, 2),
                    round(custo_km, 3),
                    round(custo_kg, 3),
                    round(ticket, 2),
                ),
                tags=(tag,),
            )

        total_rotas = sum(r[1] for r in rows_out)
        total_km = sum(r[2] for r in rows_out)
        total_kg = sum(r[3] for r in rows_out)
        total_desp = sum(r[4] for r in rows_out)
        custo_km_global = (total_desp / total_km) if total_km > 0 else 0.0
        custo_kg_global = (total_desp / total_kg) if total_kg > 0 else 0.0
        metric = upper(self.var_chart_metric.get().strip())
        if metric == "CUSTO_KG":
            idx = 6
            title = "Custo por VeÃÂculo (Custo/KG)"
        elif metric == "DESPESA_TOTAL":
            idx = 4
            title = "Custo por VeÃÂculo (Despesa Total)"
        else:
            idx = 5
            title = "Custo por VeÃÂculo (Custo/KM)"
        self.lbl_chart_title.config(text=title)
        chart_rows = sorted(rows_out, key=lambda r: safe_float(r[idx], 0.0), reverse=True)
        self._chart_labels = [r[0] for r in chart_rows[:10]]
        self._chart_values = [r[idx] for r in chart_rows[:10]]
        self._draw_chart(self._chart_labels, self._chart_values)

        self.lbl_resumo.config(
            text=(
                f"VeÃÂculos: {len(rows_out)} | Rotas: {total_rotas} | "
                f"KM: {total_km:.1f} | KG carregado: {total_kg:.2f} | "
                f"Despesas: R$ {total_desp:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                + f" | Custo/KM global: {custo_km_global:.3f} | Custo/KG global: {custo_kg_global:.3f}"
            )
        )
        self.set_status(f"STATUS: Centro de Custos atualizado ({self.var_periodo.get()} dias / veÃÂculo {self.var_veiculo.get()}).")


class RelatoriosPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Relatorios")
        self.body.grid_rowconfigure(2, weight=1)
        self.body.grid_columnconfigure(0, weight=1)

        card = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(12, weight=1)

        ttk.Label(card, text="Tipo de Relatorio", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.cb_tipo_rel = ttk.Combobox(
            card,
            state="readonly",
            values=[
                "Programacoes",
                "Prestacao de Contas",
                "Mortalidade Motorista",
                "Rotina Motorista/Ajudantes",
                "KM de Veiculos",
                "Despesas",
            ],
            width=28,
        )
        self.cb_tipo_rel.grid(row=1, column=0, sticky="ew", padx=6)
        self.cb_tipo_rel.set("Programacoes")

        ttk.Label(card, text="Codigo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.ent_filtro_codigo = ttk.Entry(card, style="Field.TEntry", width=16)
        self.ent_filtro_codigo.grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Motorista", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.ent_filtro_motorista = ttk.Entry(card, style="Field.TEntry", width=24)
        self.ent_filtro_motorista.grid(row=1, column=2, sticky="ew", padx=6)

        ttk.Label(card, text="Data", style="CardLabel.TLabel").grid(row=0, column=3, sticky="w")
        self.ent_filtro_data = ttk.Entry(card, style="Field.TEntry", width=12)
        self.ent_filtro_data.grid(row=1, column=3, sticky="ew", padx=6)
        self._bind_date_mask_relatorio(self.ent_filtro_data)

        ttk.Button(card, text="ðŸ”Ž BUSCAR", style="Primary.TButton", command=self._buscar_programacoes_relatorio).grid(
            row=1, column=4, padx=6
        )
        ttk.Button(card, text="ðÅ¸§¹ LIMPAR", style="Ghost.TButton", command=self._limpar_filtros_relatorio).grid(
            row=1, column=5, padx=6
        )

        ttk.Label(card, text="Programacao", style="CardLabel.TLabel").grid(row=2, column=0, sticky="w")
        self.cb_prog = ttk.Combobox(card, state="readonly")
        self.cb_prog.grid(row=3, column=0, sticky="ew", padx=6)

        ttk.Button(card, text="ðÅ¸§¾ GERAR RESUMO", style="Primary.TButton", command=self.gerar_resumo).grid(
            row=3, column=1, padx=6
        )
        ttk.Button(card, text="ðÅ¸â€œ¤ EXPORTAR EXCEL", style="Warn.TButton", command=self.exportar_excel).grid(
            row=3, column=2, padx=6
        )
        ttk.Button(card, text="ðÅ¸â€œâ€ž GERAR PDF", style="Primary.TButton", command=self.gerar_pdf).grid(
            row=3, column=3, padx=6
        )
        ttk.Button(card, text="ðÅ¸â€˜Â PREVIEW", style="Ghost.TButton", command=self.abrir_previsualizacao_relatorio).grid(
            row=3, column=4, padx=6
        )
        ttk.Button(card, text="ðŸ”„ ATUALIZAR", style="Ghost.TButton", command=self.refresh_comboboxes).grid(
            row=3, column=5, padx=6
        )

        ttk.Button(card, text="ðÅ¸ÂÂ FINALIZAR ROTA", style="Danger.TButton", command=self.finalizar_rota).grid(
            row=3, column=6, padx=6
        )
        ttk.Button(card, text="ââ€ © REABRIR ROTA", style="Warn.TButton", command=self.reabrir_rota).grid(
            row=3, column=7, padx=6
        )

        self.var_show_receb_detalhe = tk.BooleanVar(value=True)
        self.var_show_desp_detalhe = tk.BooleanVar(value=True)

        details_frame = ttk.Frame(card, style="Card.TFrame")
        details_frame.grid(row=4, column=0, columnspan=8, sticky="w", pady=(8, 0))
        ttk.Label(details_frame, text="Blocos do Resumo:", style="CardLabel.TLabel").pack(side="left", padx=(0, 8))
        ttk.Checkbutton(
            details_frame,
            text="Recebimentos detalhados",
            variable=self.var_show_receb_detalhe,
            command=self._refresh_resumo_if_ready,
        ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(
            details_frame,
            text="Despesas detalhadas",
            variable=self.var_show_desp_detalhe,
            command=self._refresh_resumo_if_ready,
        ).pack(side="left")

        dash = ttk.Frame(self.body, style="Card.TFrame", padding=10)
        dash.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        dash.grid_columnconfigure(0, weight=1)
        dash.grid_columnconfigure(1, weight=1)

        self.lbl_kpi_1 = ttk.Label(dash, text="Registros: 0", style="CardLabel.TLabel")
        self.lbl_kpi_1.grid(row=0, column=0, sticky="w")
        self.lbl_kpi_2 = ttk.Label(dash, text="Total: R$ 0,00", style="CardLabel.TLabel")
        self.lbl_kpi_2.grid(row=0, column=1, sticky="w")
        self.lbl_kpi_3 = ttk.Label(dash, text="Média: 0,00", style="CardLabel.TLabel")
        self.lbl_kpi_3.grid(row=1, column=0, sticky="w")
        self.lbl_kpi_4 = ttk.Label(dash, text="Destaque: -", style="CardLabel.TLabel")
        self.lbl_kpi_4.grid(row=1, column=1, sticky="w")

        self.cv_chart = tk.Canvas(dash, height=140, bg="white", highlightthickness=1, highlightbackground="#E5E7EB")
        self.cv_chart.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self._chart_bars = []
        self._chart_tip_items = ()
        self.cv_chart.bind("<Motion>", self._on_chart_motion)
        self.cv_chart.bind("<Leave>", lambda _e: self._hide_chart_tooltip())

        self.txt = tk.Text(self.body, height=18)
        self.txt.grid(row=2, column=0, sticky="nsew", pady=(10, 0))

        self.cb_tipo_rel.bind("<<ComboboxSelected>>", lambda _e: self._buscar_programacoes_relatorio())
        self.cb_prog.bind("<<ComboboxSelected>>", lambda _e: self._refresh_resumo_if_ready())
        self.ent_filtro_codigo.bind("<Return>", lambda _e: self._buscar_programacoes_relatorio())
        self.ent_filtro_motorista.bind("<Return>", lambda _e: self._buscar_programacoes_relatorio())
        self.ent_filtro_data.bind("<Return>", lambda _e: self._buscar_programacoes_relatorio())

        self.refresh_comboboxes()

    # ----------------------------
    # Helpers de regra/segurança
    # ----------------------------
    def _get_prog_status_info(self, prog: str):
        """
        Retorna dict com status/prestacao_status quando existir.
        CompatÃÂÂvel com bases antigas (sem coluna prestacao_status).
        """
        info = {"status": "", "prestacao_status": ""}
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA table_info(programacoes)")
            cols = [c[1] for c in cur.fetchall()]
            has_prest = "prestacao_status" in cols

            if has_prest:
                cur.execute("""
                    SELECT COALESCE(status,''), COALESCE(prestacao_status,'')
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                row = cur.fetchone()
                if row:
                    info["status"] = upper(row[0])
                    info["prestacao_status"] = upper(row[1])
            else:
                cur.execute("""
                    SELECT COALESCE(status,'')
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                row = cur.fetchone()
                if row:
                    info["status"] = upper(row[0])
                    info["prestacao_status"] = ""
        return info

    def _is_prestacao_fechada(self, prog: str) -> bool:
        info = self._get_prog_status_info(prog)
        return info.get("prestacao_status", "") == "FECHADA"

    def _tipo_exige_programacao(self) -> bool:
        tipo = upper(self.cb_tipo_rel.get().strip())
        return tipo in ("PROGRAMACOES", "PRESTACAO DE CONTAS")

    def refresh_comboboxes(self):
        self._buscar_programacoes_relatorio()

    def _limpar_filtros_relatorio(self):
        self.cb_tipo_rel.set("Programacoes")
        self.ent_filtro_codigo.delete(0, "end")
        self.ent_filtro_motorista.delete(0, "end")
        self.ent_filtro_data.delete(0, "end")
        self._buscar_programacoes_relatorio()

    def _buscar_programacoes_relatorio(self):
        tipo_rel = upper(self.cb_tipo_rel.get().strip()) if hasattr(self, "cb_tipo_rel") else "PROGRAMACOES"
        filtro_cod = upper(self.ent_filtro_codigo.get().strip()) if hasattr(self, "ent_filtro_codigo") else ""
        filtro_mot = upper(self.ent_filtro_motorista.get().strip()) if hasattr(self, "ent_filtro_motorista") else ""
        filtro_data_raw = str(self.ent_filtro_data.get().strip()) if hasattr(self, "ent_filtro_data") else ""

        data_patterns = []
        if filtro_data_raw:
            n = normalize_date(filtro_data_raw)
            if n:
                data_patterns.append(n)
                data_patterns.append(f"{n[8:10]}/{n[5:7]}/{n[0:4]}")
            data_patterns.append(filtro_data_raw)
            data_patterns = [p for p in dict.fromkeys(data_patterns) if p]

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                modo = "finalizadas_prestacao" if "PRESTACAO" in tipo_rel else "todas"
                resp = _call_api(
                    "GET",
                    f"desktop/programacoes?modo={urllib.parse.quote(modo)}&limit=400",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                arr = resp.get("programacoes") if isinstance(resp, dict) else []
                encontrados = []
                for r in (arr or []):
                    if not isinstance(r, dict):
                        continue
                    codigo = upper((r or {}).get("codigo_programacao") or "")
                    motorista = upper((r or {}).get("motorista") or "")
                    data_ref = str((r or {}).get("data_referencia") or (r or {}).get("data_criacao") or "")
                    if filtro_cod and filtro_cod not in codigo:
                        continue
                    if filtro_mot and filtro_mot not in motorista:
                        continue
                    if data_patterns:
                        txt = upper(data_ref)
                        ok_data = any(upper(p) in txt for p in data_patterns)
                        if not ok_data:
                            continue
                    if codigo:
                        encontrados.append(codigo)
                atual = upper(self.cb_prog.get().strip())
                self.cb_prog["values"] = encontrados
                if atual not in encontrados:
                    self.cb_prog.set("")
                self.set_status(f"STATUS: {len(encontrados)} programacoes encontradas para {self.cb_tipo_rel.get()}.")
                return
            except Exception:
                logging.debug("Falha ao buscar programacoes de relatorio via API", exc_info=True)

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols = [str(c[1]).lower() for c in cur.fetchall()]

                sql = "SELECT codigo_programacao FROM programacoes WHERE 1=1"
                params = []

                if filtro_cod:
                    sql += " AND UPPER(COALESCE(codigo_programacao,'')) LIKE ?"
                    params.append(f"%{filtro_cod}%")

                if filtro_mot and "motorista" in cols:
                    sql += " AND UPPER(COALESCE(motorista,'')) LIKE ?"
                    params.append(f"%{filtro_mot}%")

                if data_patterns:
                    date_cols = [c for c in ("data_criacao", "data", "data_saida") if c in cols]
                    if date_cols:
                        clauses = []
                        for dc in date_cols:
                            for pat in data_patterns:
                                clauses.append(f"COALESCE({dc},'') LIKE ?")
                                params.append(f"%{pat}%")
                        sql += " AND (" + " OR ".join(clauses) + ")"

                if "PRESTACAO" in tipo_rel:
                    if "prestacao_status" in cols:
                        sql += " AND (UPPER(COALESCE(prestacao_status,'')) IN ('PENDENTE','FECHADA') OR UPPER(COALESCE(status,''))='FINALIZADA')"
                    elif "status" in cols:
                        sql += " AND UPPER(COALESCE(status,''))='FINALIZADA'"

                if "status" in cols and "id" in cols:
                    sql += " ORDER BY CASE WHEN UPPER(COALESCE(status,''))='ATIVA' THEN 0 ELSE 1 END, id DESC LIMIT 400"
                elif "id" in cols:
                    sql += " ORDER BY id DESC LIMIT 400"
                else:
                    sql += " LIMIT 400"

                cur.execute(sql, tuple(params))
            except Exception:
                cur.execute("SELECT codigo_programacao FROM programacoes ORDER BY id DESC LIMIT 400")

            encontrados = [r[0] for r in cur.fetchall()]
            atual = upper(self.cb_prog.get().strip())
            self.cb_prog["values"] = encontrados
            if atual not in encontrados:
                self.cb_prog.set("")

        self.set_status(f"STATUS: {len(encontrados)} programacoes encontradas para {self.cb_tipo_rel.get()}.")

    def _refresh_resumo_if_ready(self):
        try:
            if (not self._tipo_exige_programacao()) or upper(self.cb_prog.get().strip()):
                self.gerar_resumo()
        except Exception:
            logging.debug("Falha ignorada")

    def _api_bundle_relatorio(self, prog: str):
        prog = upper(str(prog or "").strip())
        self._last_bundle_api_state = "disabled"
        if not prog:
            self._last_bundle_api_state = "invalid"
            return None
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if not desktop_secret or not is_desktop_api_sync_enabled():
            self._last_bundle_api_state = "disabled"
            return None
        try:
            rota_resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(prog)}",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            rec_resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(prog)}/recebimentos",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            desp_resp = _call_api(
                "GET",
                f"desktop/rotas/{urllib.parse.quote(prog)}/despesas",
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )
            rota = rota_resp.get("rota") if isinstance(rota_resp, dict) else None
            clientes = rota_resp.get("clientes") if isinstance(rota_resp, dict) else []
            receb = rec_resp.get("recebimentos") if isinstance(rec_resp, dict) else []
            desp = desp_resp.get("despesas") if isinstance(desp_resp, dict) else []
            if not isinstance(rota, dict):
                self._last_bundle_api_state = "not_found"
                return None
            self._last_bundle_api_state = "ok"
            return {
                "rota": rota,
                "clientes": clientes if isinstance(clientes, list) else [],
                "recebimentos": receb if isinstance(receb, list) else [],
                "despesas": desp if isinstance(desp, list) else [],
            }
        except Exception:
            self._last_bundle_api_state = "failed"
            logging.debug("Falha ao montar bundle de relatorio via API", exc_info=True)
            return None

    def _fmt_rel_money(self, v):
        return f"R$ {safe_float(v, 0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _draw_chart(self, labels, values, color="#1D4ED8"):
        try:
            cv = self.cv_chart
            cv.delete("all")
            self._chart_bars = []
            self._chart_tip_items = ()
            w = max(cv.winfo_width(), 10)
            h = max(cv.winfo_height(), 10)
            if (not labels) or (not values):
                cv.create_text(w / 2, h / 2, text="Sem dados para gráfico", fill="#6B7280", font=("Segoe UI", 10))
                return

            n = min(len(labels), len(values), 10)
            labels = labels[:n]
            vals = [max(safe_float(v, 0.0), 0.0) for v in values[:n]]
            vmax = max(vals) if vals else 1.0
            if vmax <= 0:
                vmax = 1.0

            # Paleta no estilo da referência (barras coloridas + legenda)
            palette = [
                "#F68D56", "#F2C96D", "#D7E64A", "#63D04A", "#67D08D",
                "#58C6C4", "#5E9AC8", "#A78BFA", "#F472B6", "#22C55E"
            ]

            # ÃÂrea do gráfico
            left = 52
            right = 16
            top = 46
            bottom = 26
            plot_w = max(w - left - right, 40)
            plot_h = max(h - top - bottom, 30)
            x0 = left
            y0 = top
            x1 = left + plot_w
            y1 = top + plot_h

            # Fundo
            cv.create_rectangle(0, 0, w, h, fill="#FFFFFF", outline="")

            # Grade horizontal + escala Y
            def _nice_step(v):
                if v <= 5:
                    return 1.0
                if v <= 10:
                    return 2.0
                if v <= 25:
                    return 5.0
                if v <= 50:
                    return 10.0
                if v <= 100:
                    return 20.0
                return max(round(v / 5.0), 1.0)

            step = _nice_step(vmax)
            y_top = step * max(int((vmax / step) + 0.999), 1)
            ticks = 5
            for i in range(ticks + 1):
                v = (y_top / ticks) * i
                yy = y1 - (v / y_top) * plot_h
                cv.create_line(x0, yy, x1, yy, fill="#D1D5DB", width=1)
                txt = f"{v:.0f}" if abs(v - round(v)) < 1e-9 else f"{v:.1f}"
                cv.create_text(x0 - 8, yy, text=txt, anchor="e", fill="#4B5563", font=("Segoe UI", 8, "bold"))

            # Eixos
            cv.create_line(x0, y0, x0, y1, fill="#374151", width=1)
            cv.create_line(x0, y1, x1, y1, fill="#374151", width=2)

            # Barras
            gap = 10
            bw = max((plot_w - gap * (n + 1)) / n, 10)
            for i, (lb, vv) in enumerate(zip(labels, vals)):
                c = palette[i % len(palette)] if n > 1 else color
                bx1 = x0 + gap + i * (bw + gap)
                bx2 = bx1 + bw
                bh = (vv / y_top) * plot_h
                by1 = y1 - bh
                cv.create_rectangle(bx1, by1, bx2, y1, fill=c, outline="")
                cv.create_text((bx1 + bx2) / 2, y1 + 11, text=str(lb)[:10], fill="#111827", font=("Segoe UI", 8))
                self._chart_bars.append({
                    "x1": bx1, "y1": by1, "x2": bx2, "y2": y1,
                    "label": str(lb), "value": vv, "color": c
                })

            # Legenda superior (duas linhas quando necessário)
            legend_x = x0 + 4
            legend_y = 10
            max_w = x1 - x0 - 8
            cx = legend_x
            cy = legend_y
            for i, lb in enumerate(labels):
                c = palette[i % len(palette)] if n > 1 else color
                entry_text = str(lb)[:16]
                txt_w = max(len(entry_text) * 6, 32)
                block_w = 14 + 6 + txt_w + 12
                if (cx - legend_x + block_w) > max_w:
                    cx = legend_x
                    cy += 18
                cv.create_rectangle(cx, cy + 3, cx + 12, cy + 11, fill=c, outline="")
                cv.create_text(cx + 16, cy + 7, text=entry_text, anchor="w", fill="#374151", font=("Segoe UI", 8, "bold"))
                cx += block_w
        except Exception:
            logging.debug("Falha ignorada")

    def _hide_chart_tooltip(self):
        try:
            cv = self.cv_chart
            for iid in self._chart_tip_items:
                try:
                    cv.delete(iid)
                except Exception:
                    logging.debug("Falha ignorada")
            self._chart_tip_items = ()
        except Exception:
            logging.debug("Falha ignorada")

    def _on_chart_motion(self, event=None):
        try:
            if event is None:
                return
            hit = None
            for bar in (self._chart_bars or []):
                if bar["x1"] <= event.x <= bar["x2"] and bar["y1"] <= event.y <= bar["y2"]:
                    hit = bar
                    break
            if not hit:
                self._hide_chart_tooltip()
                return

            self._hide_chart_tooltip()
            cv = self.cv_chart
            txt = f"{hit['label']}: {safe_float(hit['value'], 0.0):.2f}"
            tx = min(max(event.x + 10, 10), max(cv.winfo_width() - 150, 10))
            ty = max(event.y - 24, 8)
            rect = cv.create_rectangle(tx, ty, tx + 140, ty + 20, fill="#111827", outline="")
            label = cv.create_text(tx + 6, ty + 10, text=txt, anchor="w", fill="white", font=("Segoe UI", 8, "bold"))
            self._chart_tip_items = (rect, label)
        except Exception:
            logging.debug("Falha ignorada")

    def _set_dashboard(self, k1, k2, k3, k4, labels=None, values=None, color="#1D4ED8"):
        self.lbl_kpi_1.config(text=k1)
        self.lbl_kpi_2.config(text=k2)
        self.lbl_kpi_3.config(text=k3)
        self.lbl_kpi_4.config(text=k4)
        self._draw_chart(labels or [], values or [], color=color)

    def _bind_date_mask_relatorio(self, ent: tk.Entry):
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
            if n == 8:
                try:
                    dd = int(digits[:2])
                    mm = int(digits[2:4])
                    yy = int(digits[4:8])
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
                while digits and (not _is_partial_date_digits_valid(digits)):
                    digits = digits[:-1]
                ent.delete(0, "end")
                ent.insert(0, _to_masked(digits))
                ent.icursor("end")
                digits_state["value"] = digits
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out():
            try:
                digits = digits_state["value"]
                if digits and len(digits) != 8:
                    ent.delete(0, "end")
                    digits_state["value"] = ""
            except Exception:
                logging.debug("Falha ignorada")

        ent.bind("<KeyRelease>", lambda _e: _apply_mask())
        ent.bind("<FocusOut>", lambda _e: (_apply_mask(), _on_focus_out()))

    def on_show(self):
        self.refresh_comboboxes()
        self.set_status("STATUS: Relatórios e exportação.")

    def abrir_previsualizacao_relatorio(self):
        tipo_rel = upper(self.cb_tipo_rel.get().strip())
        prog = upper(self.cb_prog.get().strip())
        if self._tipo_exige_programacao() and not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao para visualizar.")
            return
        self.gerar_resumo()
        txt = self.txt.get("1.0", "end").strip()
        top = tk.Toplevel(self)
        top.title(f"Pré-visualização - {self.cb_tipo_rel.get()}")
        top.geometry("1180x760")
        top.minsize(900, 600)
        nb = ttk.Notebook(top)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        if "PROGRAMACOES" in tipo_rel and prog:
            self._create_a4_preview_tab(nb, "Folha Programação", self._build_preview_folha_programacao(prog))

        if "PRESTACAO" in tipo_rel and prog:
            self._create_a4_preview_tab(nb, "Folha Prestação", self._build_preview_folha_prestacao(prog))

        tab_resumo = ttk.Frame(nb)
        nb.add(tab_resumo, text="Resumo")
        t = tk.Text(tab_resumo, wrap="word")
        t.pack(fill="both", expand=True)
        t.insert("1.0", txt)
        t.configure(state="disabled")

        if "PRESTACAO" in tipo_rel and prog:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
            bundle = self._api_bundle_relatorio(prog)
            api_state = str(getattr(self, "_last_bundle_api_state", ""))
            rec_lines = []
            if bundle:
                rows = [
                    (
                        str((r or {}).get("cod_cliente") or ""),
                        str((r or {}).get("nome_cliente") or ""),
                        safe_float((r or {}).get("valor"), 0.0),
                        str((r or {}).get("forma_pagamento") or ""),
                        str((r or {}).get("observacao") or ""),
                    )
                    for r in (bundle.get("recebimentos") or [])
                    if isinstance(r, dict)
                ]
            elif api_enabled and api_state == "not_found":
                rows = []
            else:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        SELECT COALESCE(cod_cliente,''), COALESCE(nome_cliente,''), COALESCE(valor,0), COALESCE(forma_pagamento,''), COALESCE(observacao,'')
                        FROM recebimentos WHERE codigo_programacao=? ORDER BY id DESC
                    """, (prog,))
                    rows = cur.fetchall() or []
            rec_lines.append(f"FOLHA DE RECEBIMENTOS - {prog}")
            rec_lines.append("=" * 90)
            rec_lines.append("")
            if not rows:
                rec_lines.append("Sem recebimentos.")
            else:
                rec_lines.append("COD | CLIENTE | VALOR | FORMA | OBS")
                rec_lines.append("-" * 90)
                for r in rows:
                    rec_lines.append(f"{r[0]} | {r[1]} | {self._fmt_rel_money(r[2])} | {r[3] or '-'} | {r[4] or '-'}")
            self._create_a4_preview_tab(nb, "Folha Recebimentos", "\n".join(rec_lines))

            desp_lines = []
            if bundle:
                rows = [
                    (
                        str((r or {}).get("categoria") or "OUTROS"),
                        str((r or {}).get("descricao") or ""),
                        safe_float((r or {}).get("valor"), 0.0),
                        str((r or {}).get("observacao") or ""),
                    )
                    for r in (bundle.get("despesas") or [])
                    if isinstance(r, dict)
                ]
            elif api_enabled and api_state == "not_found":
                rows = []
            else:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        SELECT COALESCE(categoria,'OUTROS'), COALESCE(descricao,''), COALESCE(valor,0), COALESCE(observacao,'')
                        FROM despesas WHERE codigo_programacao=? ORDER BY id DESC
                    """, (prog,))
                    rows = cur.fetchall() or []
            desp_lines.append(f"FOLHA DE DESPESAS - {prog}")
            desp_lines.append("=" * 90)
            desp_lines.append("")
            if not rows:
                desp_lines.append("Sem despesas.")
            else:
                desp_lines.append("CATEGORIA | DESCRICAO | VALOR | OBS")
                desp_lines.append("-" * 90)
                for r in rows:
                    desp_lines.append(f"{r[0] or 'OUTROS'} | {r[1]} | {self._fmt_rel_money(r[2])} | {r[3] or '-'}")
            self._create_a4_preview_tab(nb, "Folha Despesas", "\n".join(desp_lines))

    def _create_a4_preview_tab(self, notebook, title: str, content: str):
        """
        Cria aba de preview em formato de folha A4 (largura fixa), com rolagem.
        """
        tab = ttk.Frame(notebook)
        notebook.add(tab, text=title)
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

        # A4 aproximado em tela: largura fixa e aspecto de folha.
        sheet_width_px = 860
        sheet = ttk.Frame(canvas, style="Card.TFrame", padding=18)
        text = tk.Text(
            sheet,
            wrap="none",
            font=("Consolas", 10),
            width=116,
            height=60,
            relief="flat",
            borderwidth=0,
            background="white",
            foreground="#111827",
        )
        text.pack(fill="both", expand=True)
        text.insert("1.0", content or "")
        text.configure(state="disabled")

        win_id = canvas.create_window((20, 20), window=sheet, anchor="nw", width=sheet_width_px)

        def _on_configure(_e=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                logging.debug("Falha ignorada")

        def _on_resize(event=None):
            try:
                cw = max(int(canvas.winfo_width()), 0)
                x = max(int((cw - sheet_width_px) / 2), 20)
                canvas.coords(win_id, x, 20)
                _on_configure()
            except Exception:
                logging.debug("Falha ignorada")

        sheet.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", _on_resize)
        _on_resize()

    def _build_preview_folha_programacao(self, prog: str) -> str:
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
        bundle = self._api_bundle_relatorio(prog)
        api_state = str(getattr(self, "_last_bundle_api_state", ""))
        if bundle:
            try:
                rota = bundle.get("rota") or {}
                clientes_api = [r for r in (bundle.get("clientes") or []) if isinstance(r, dict)]
                data_criacao = str(rota.get("data_criacao") or rota.get("data") or "")
                motorista = str(rota.get("motorista") or "")
                veiculo = str(rota.get("veiculo") or "")
                equipe = str(rota.get("equipe") or "")
                kg_estimado = safe_float(rota.get("kg_estimado"), 0.0)
                status = str(rota.get("status") or "")
                local = str(rota.get("local_rota") or rota.get("local") or "")
                tipo_estimativa = str(rota.get("tipo_estimativa") or "KG")
                caixas_estimado = safe_int(rota.get("caixas_estimado"), 0)
                usuario_criacao = str(
                    rota.get("usuario_criacao")
                    or rota.get("usuario")
                    or rota.get("criado_por")
                    or rota.get("created_by")
                    or rota.get("autor")
                    or ""
                )
                usuario_edicao = str(rota.get("usuario_ultima_edicao") or "")
                itens = [
                    (
                        str((r or {}).get("cod_cliente") or ""),
                        str((r or {}).get("nome_cliente") or ""),
                        safe_float((r or {}).get("preco_atual"), safe_float((r or {}).get("preco"), 0.0)),
                        str((r or {}).get("vendedor") or ""),
                        str((r or {}).get("cidade") or ""),
                        str((r or {}).get("pedido") or ""),
                        safe_int((r or {}).get("caixas_atual"), safe_int((r or {}).get("qnt_caixas"), 0)),
                    )
                    for r in clientes_api
                ]
                itens.sort(key=lambda r: (upper(r[1] or ""), upper(r[0] or "")))
                equipe_txt = resolve_equipe_nomes(equipe)
                total_prev = sum(safe_float(r[2], 0.0) for r in itens)
                tipo_estimativa = upper(tipo_estimativa or "KG")
                if tipo_estimativa == "CX":
                    estimativa_txt = f"CIF / CX ESTIMADO: {safe_int(caixas_estimado, 0)}"
                else:
                    estimativa_txt = f"FOB / KG ESTIMADO: {safe_float(kg_estimado, 0.0):.2f}"
                lines = []
                lines.append("=" * 118)
                lines.append(f"{'FOLHA DE PROGRAMAÃ‡ÃƒO':^118}")
                lines.append("=" * 118)
                lines.append(f"CÃ“DIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}")
                lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEÃƒÃ‚ÂCULO: {upper(veiculo or '-')}   LOCAL: {upper(local or '-')}")
                lines.append(f"EQUIPE: {equipe_txt or '-'}")
                lines.append(
                    f"{estimativa_txt}   CLIENTES: {len(itens)}   TOTAL ESTIMADO: {self._fmt_rel_money(total_prev)}"
                )
                lines.append(
                    f"CRIADO POR: {upper(usuario_criacao or '-')}" +
                    f"   ÃšLTIMA EDIÃ‡ÃƒO: {upper(usuario_edicao or '-')}"
                )
                lines.append("-" * 118)
                lines.append(f"{'COD':<6} {'CLIENTE / CIDADE':<48} {'CX':>4} {'PREÃ‡O':>12} {'VENDEDOR':<22} {'PEDIDO':>12}")
                lines.append("-" * 118)
                for cod, nome, preco, vendedor, cidade, pedido, caixas in itens:
                    cli = f"{upper(nome or '-')}"
                    if cidade:
                        cli += f" - {upper(cidade)}"
                    lines.append(
                        f"{str(cod)[:6]:<6} {cli[:48]:<48} {safe_int(caixas,0):>4} {self._fmt_rel_money(preco):>12} "
                        f"{upper(vendedor or '-')[:22]:<22} {str(pedido or '-')[:12]:>12}"
                    )
                lines.append("-" * 118)
                lines.append("ObservaÃ§Ã£o: _______________________________________________________________________________________________")
                lines.append("Assinatura responsÃ¡vel: _________________________________________")
                return "\n".join(lines)
            except Exception:
                logging.debug("Falha ao montar folha de programacao via API bundle", exc_info=True)
                if api_enabled:
                    return (
                        f"FOLHA DE PROGRAMACAO - {prog}\n"
                        + "=" * 90
                        + "\n\nFalha ao processar dados retornados pelo servidor."
                    )

        if api_enabled and api_state == "not_found":
            return (
                f"FOLHA DE PROGRAMACAO - {prog}\n"
                + "=" * 90
                + "\n\nProgramacao nao encontrada no servidor."
            )

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_p = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                local_expr = "COALESCE(local_rota,'')" if "local_rota" in cols_p else ("COALESCE(local,'')" if "local" in cols_p else "''")
                kg_expr = "COALESCE(kg_estimado,0)" if "kg_estimado" in cols_p else "0"
                status_expr = "COALESCE(status,'')" if "status" in cols_p else "''"
                tipo_estim_expr = "COALESCE(tipo_estimativa,'KG')" if "tipo_estimativa" in cols_p else "'KG'"
                caixas_estim_expr = "COALESCE(caixas_estimado,0)" if "caixas_estimado" in cols_p else "0"
                user_criacao_expr = "COALESCE(usuario_criacao,'')" if "usuario_criacao" in cols_p else "''"
                user_edicao_expr = "COALESCE(usuario_ultima_edicao,'')" if "usuario_ultima_edicao" in cols_p else "''"

                cur.execute(f"""
                    SELECT COALESCE(data_criacao,''), COALESCE(motorista,''), COALESCE(veiculo,''),
                           COALESCE(equipe,''), {kg_expr}, {status_expr}, {local_expr},
                           {tipo_estim_expr}, {caixas_estim_expr}, {user_criacao_expr}, {user_edicao_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                meta = cur.fetchone() or ("", "", "", "", 0, "", "", "KG", 0, "", "")

                cur.execute("PRAGMA table_info(programacao_itens)")
                cols_i = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                cidade_expr = "COALESCE(cidade,'')" if "cidade" in cols_i else "''"
                vendedor_expr = "COALESCE(vendedor,'')" if "vendedor" in cols_i else "''"
                pedido_expr = "COALESCE(pedido,'')" if "pedido" in cols_i else "''"
                caixas_expr = "COALESCE(qnt_caixas,0)" if "qnt_caixas" in cols_i else "0"
                preco_expr = "COALESCE(preco,0)" if "preco" in cols_i else "0"

                cur.execute(f"""
                    SELECT COALESCE(cod_cliente,''), COALESCE(nome_cliente,''), {preco_expr},
                           {vendedor_expr}, {cidade_expr}, {pedido_expr}, {caixas_expr}
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                    ORDER BY nome_cliente ASC, cod_cliente ASC
                """, (prog,))
                itens = cur.fetchall() or []
        except Exception as e:
            return (
                f"ERRO AO MONTAR FOLHA DE PROGRAMAÇÃO ({prog})\n"
                + "=" * 90
                + f"\n\n{str(e)}\n\n"
                + "Verifique a estrutura do banco e tente novamente."
            )

        data_criacao, motorista, veiculo, equipe, kg_estimado, status, local, tipo_estimativa, caixas_estimado, usuario_criacao, usuario_edicao = meta
        equipe_txt = resolve_equipe_nomes(equipe)
        total_prev = sum(safe_float(r[2], 0.0) for r in itens)
        tipo_estimativa = upper(tipo_estimativa or "KG")
        if tipo_estimativa == "CX":
            estimativa_txt = f"CIF / CX ESTIMADO: {safe_int(caixas_estimado, 0)}"
        else:
            estimativa_txt = f"FOB / KG ESTIMADO: {safe_float(kg_estimado, 0.0):.2f}"
        lines = []
        lines.append("=" * 118)
        lines.append(f"{'FOLHA DE PROGRAMAÇÃO':^118}")
        lines.append("=" * 118)
        lines.append(f"CÓDIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}")
        lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEÃÂCULO: {upper(veiculo or '-')}   LOCAL: {upper(local or '-')}")
        lines.append(f"EQUIPE: {equipe_txt or '-'}")
        lines.append(
            f"{estimativa_txt}   CLIENTES: {len(itens)}   TOTAL ESTIMADO: {self._fmt_rel_money(total_prev)}"
        )
        lines.append(
            f"CRIADO POR: {upper(usuario_criacao or '-')}" +
            f"   ÚLTIMA EDIÇÃO: {upper(usuario_edicao or '-')}"
        )
        lines.append("-" * 118)
        lines.append(f"{'COD':<6} {'CLIENTE / CIDADE':<48} {'CX':>4} {'PREÇO':>12} {'VENDEDOR':<22} {'PEDIDO':>12}")
        lines.append("-" * 118)
        for cod, nome, preco, vendedor, cidade, pedido, caixas in itens:
            cli = f"{upper(nome or '-')}"
            if cidade:
                cli += f" - {upper(cidade)}"
            lines.append(
                f"{str(cod)[:6]:<6} {cli[:48]:<48} {safe_int(caixas,0):>4} {self._fmt_rel_money(preco):>12} "
                f"{upper(vendedor or '-')[:22]:<22} {str(pedido or '-')[:12]:>12}"
            )
        lines.append("-" * 118)
        lines.append("Observação: _______________________________________________________________________________________________")
        lines.append("Assinatura responsável: _________________________________________")
        return "\n".join(lines)

    def _build_preview_folha_prestacao(self, prog: str) -> str:
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
        bundle = self._api_bundle_relatorio(prog)
        api_state = str(getattr(self, "_last_bundle_api_state", ""))
        if bundle:
            try:
                rota = bundle.get("rota") or {}
                receb_api = [r for r in (bundle.get("recebimentos") or []) if isinstance(r, dict)]
                desp_api = [r for r in (bundle.get("despesas") or []) if isinstance(r, dict)]

                data_criacao = str(rota.get("data_criacao") or rota.get("data") or "")
                motorista = str(rota.get("motorista") or "")
                veiculo = str(rota.get("veiculo") or "")
                equipe = str(rota.get("equipe") or "")
                status = str(rota.get("status") or "")
                prest = str(rota.get("prestacao_status") or "PENDENTE")
                km_rodado = safe_float(rota.get("km_rodado"), 0.0)
                media_km_l = safe_float(rota.get("media_km_l"), 0.0)
                custo_km = safe_float(rota.get("custo_km"), 0.0)
                adiantamento = safe_float(rota.get("adiantamento"), 0.0)
                mort_aves = safe_int(rota.get("mortalidade_transbordo_aves"), 0)
                mort_kg = safe_float(rota.get("mortalidade_transbordo_kg"), 0.0)
                obs_transb = str(rota.get("obs_transbordo") or "")
                nf_kg_base = safe_float(rota.get("nf_kg"), 0.0)

                total_receb = sum(safe_float(r.get("valor"), 0.0) for r in receb_api)
                total_desp = sum(safe_float(r.get("valor"), 0.0) for r in desp_api)
                receb = [
                    (
                        str(r.get("cod_cliente") or ""),
                        str(r.get("nome_cliente") or ""),
                        safe_float(r.get("valor"), 0.0),
                        str(r.get("forma_pagamento") or ""),
                    )
                    for r in receb_api[:12]
                ]
                desp = [
                    (
                        str(r.get("categoria") or "OUTROS"),
                        str(r.get("descricao") or ""),
                        safe_float(r.get("valor"), 0.0),
                    )
                    for r in desp_api[:12]
                ]

                equipe_txt = resolve_equipe_nomes(equipe)
                entradas = total_receb + safe_float(adiantamento, 0.0)
                saidas = total_desp
                resultado = entradas - saidas
                log_info = self._collect_logistica_rastreio(prog)
                kg_nf_util = max(nf_kg_base - safe_float(mort_kg, 0.0), 0.0)

                lines = []
                lines.append("=" * 118)
                lines.append(f"{'FOLHA DE PRESTACAO DE CONTAS':^118}")
                lines.append("=" * 118)
                lines.append(f"CODIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}   PRESTACAO: {upper(prest or '-')}")
                lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEICULO: {upper(veiculo or '-')}   EQUIPE: {equipe_txt or '-'}")
                lines.append("-" * 118)
                lines.append(
                    f"ENTRADAS (RECEB+ADIANT.): {self._fmt_rel_money(entradas)}   "
                    f"SAIDAS (DESPESAS): {self._fmt_rel_money(saidas)}   "
                    f"RESULTADO: {self._fmt_rel_money(resultado)}"
                )
                lines.append(
                    f"KM RODADO: {safe_float(km_rodado,0.0):.2f}   "
                    f"MEDIA KM/L: {safe_float(media_km_l,0.0):.2f}   "
                    f"CUSTO/KM: {safe_float(custo_km,0.0):.2f}"
                )
                lines.append("-" * 118)
                lines.append("[RECEBIMENTOS - ULTIMOS LANCAMENTOS]")
                lines.append(f"{'COD':<6} {'CLIENTE':<52} {'VALOR':>14} {'FORMA':<16}")
                for cod, nome, valor, forma in receb:
                    lines.append(f"{str(cod)[:6]:<6} {upper(nome or '-')[:52]:<52} {self._fmt_rel_money(valor):>14} {upper(forma or '-')[:16]:<16}")
                if not receb:
                    lines.append("Sem recebimentos.")
                lines.append("-" * 118)
                lines.append("[DESPESAS - ULTIMOS LANCAMENTOS]")
                lines.append(f"{'CATEGORIA':<20} {'DESCRICAO':<72} {'VALOR':>14}")
                for cat, desc, valor in desp:
                    lines.append(f"{upper(cat or 'OUTROS')[:20]:<20} {upper(desc or '-')[:72]:<72} {self._fmt_rel_money(valor):>14}")
                if not desp:
                    lines.append("Sem despesas.")
                lines.append("-" * 118)
                lines.append("[RASTREABILIDADE LOGISTICA]")
                for ln in (log_info.get("resumo") or []):
                    lines.append(str(ln))
                lines.append(
                    "Conciliacao de caixas: "
                    + ("OK" if bool(log_info.get("itens_ok", True)) else "DIVERGENTE")
                )
                lines.append("-" * 118)
                lines.append("[OCORRENCIAS DE TRANSBORDO / MORTALIDADE]")
                lines.append(f"Mortalidade (aves): {safe_int(mort_aves, 0)}")
                lines.append(
                    f"Mortalidade (KG): {safe_float(mort_kg, 0.0):.2f}".replace(".", ",")
                )
                lines.append(
                    f"KG util NF (NF - mortalidade): {safe_float(kg_nf_util, 0.0):.2f}".replace(".", ",")
                )
                lines.append(f"Obs transbordo: {str(obs_transb or '-').strip() or '-'}")
                lines.append("-" * 118)
                lines.append("Conferido por: ____________________________________   Data: ____/____/________")
                return "\n".join(lines)
            except Exception:
                logging.debug("Falha ao montar folha de prestacao via API bundle", exc_info=True)
                if api_enabled:
                    return (
                        f"FOLHA DE PRESTACAO - {prog}\n"
                        + "=" * 90
                        + "\n\nFalha ao processar dados retornados pelo servidor."
                    )

        if api_enabled and api_state == "not_found":
            return (
                f"FOLHA DE PRESTACAO - {prog}\n"
                + "=" * 90
                + "\n\nProgramacao nao encontrada no servidor."
            )

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_p = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                status_expr = "COALESCE(status,'')" if "status" in cols_p else "''"
                prest_expr = "COALESCE(prestacao_status,'')" if "prestacao_status" in cols_p else "''"
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols_p else "0"
                media_expr = "COALESCE(media_km_l,0)" if "media_km_l" in cols_p else "0"
                custo_expr = "COALESCE(custo_km,0)" if "custo_km" in cols_p else "0"
                adiant_expr = "COALESCE(adiantamento,0)" if "adiantamento" in cols_p else "0"
                mort_aves_expr = "COALESCE(mortalidade_transbordo_aves,0)" if "mortalidade_transbordo_aves" in cols_p else "0"
                mort_kg_expr = "COALESCE(mortalidade_transbordo_kg,0)" if "mortalidade_transbordo_kg" in cols_p else "0"
                obs_transb_expr = "COALESCE(obs_transbordo,'')" if "obs_transbordo" in cols_p else "''"
                cur.execute(f"""
                    SELECT COALESCE(data_criacao,''), COALESCE(motorista,''), COALESCE(veiculo,''), COALESCE(equipe,''),
                           {status_expr}, {prest_expr}, {km_expr}, {media_expr}, {custo_expr}, {adiant_expr},
                           {mort_aves_expr}, {mort_kg_expr}, {obs_transb_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                meta = cur.fetchone() or ("", "", "", "", "", "", 0, 0, 0, 0, 0, 0.0, "")

                cur.execute("SELECT COALESCE(SUM(valor),0) FROM recebimentos WHERE codigo_programacao=?", (prog,))
                total_receb = safe_float((cur.fetchone() or [0])[0], 0.0)
                cur.execute("SELECT COALESCE(SUM(valor),0) FROM despesas WHERE codigo_programacao=?", (prog,))
                total_desp = safe_float((cur.fetchone() or [0])[0], 0.0)
                cur.execute("""
                    SELECT COALESCE(cod_cliente,''), COALESCE(nome_cliente,''), COALESCE(valor,0), COALESCE(forma_pagamento,'')
                    FROM recebimentos
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 12
                """, (prog,))
                receb = cur.fetchall() or []
                cur.execute("""
                    SELECT COALESCE(categoria,'OUTROS'), COALESCE(descricao,''), COALESCE(valor,0)
                    FROM despesas
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 12
                """, (prog,))
                desp = cur.fetchall() or []
        except Exception as e:
            return (
                f"ERRO AO MONTAR FOLHA DE PRESTAÇÃO ({prog})\n"
                + "=" * 90
                + f"\n\n{str(e)}\n\n"
                + "Verifique a estrutura do banco e tente novamente."
            )

        (
            data_criacao,
            motorista,
            veiculo,
            equipe,
            status,
            prest,
            km_rodado,
            media_km_l,
            custo_km,
            adiantamento,
            mort_aves,
            mort_kg,
            obs_transb,
        ) = meta
        equipe_txt = resolve_equipe_nomes(equipe)
        entradas = total_receb + safe_float(adiantamento, 0.0)
        saidas = total_desp
        resultado = entradas - saidas
        log_info = self._collect_logistica_rastreio(prog)
        kg_nf_util = 0.0
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    SELECT COALESCE(nf_kg,0)
                    FROM programacoes
                    WHERE codigo_programacao=?
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (prog,),
                )
                row_nf = cur.fetchone()
                nf_kg_base = safe_float((row_nf[0] if row_nf else 0.0), 0.0)
                kg_nf_util = max(nf_kg_base - safe_float(mort_kg, 0.0), 0.0)
        except Exception:
            logging.debug("Falha ignorada")

        lines = []
        lines.append("=" * 118)
        lines.append(f"{'FOLHA DE PRESTAÇÃO DE CONTAS':^118}")
        lines.append("=" * 118)
        lines.append(f"CÓDIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}   PRESTAÇÃO: {upper(prest or '-')}")
        lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEÃÂCULO: {upper(veiculo or '-')}   EQUIPE: {equipe_txt or '-'}")
        lines.append("-" * 118)
        lines.append(
            f"ENTRADAS (RECEB+ADIANT.): {self._fmt_rel_money(entradas)}   "
            f"SAÃÂDAS (DESPESAS): {self._fmt_rel_money(saidas)}   "
            f"RESULTADO: {self._fmt_rel_money(resultado)}"
        )
        lines.append(
            f"KM RODADO: {safe_float(km_rodado,0.0):.2f}   "
            f"MÉDIA KM/L: {safe_float(media_km_l,0.0):.2f}   "
            f"CUSTO/KM: {safe_float(custo_km,0.0):.2f}"
        )
        lines.append("-" * 118)
        lines.append("[RECEBIMENTOS - ÚLTIMOS LANÇAMENTOS]")
        lines.append(f"{'COD':<6} {'CLIENTE':<52} {'VALOR':>14} {'FORMA':<16}")
        for cod, nome, valor, forma in receb:
            lines.append(f"{str(cod)[:6]:<6} {upper(nome or '-')[:52]:<52} {self._fmt_rel_money(valor):>14} {upper(forma or '-')[:16]:<16}")
        if not receb:
            lines.append("Sem recebimentos.")
        lines.append("-" * 118)
        lines.append("[DESPESAS - ÚLTIMOS LANÇAMENTOS]")
        lines.append(f"{'CATEGORIA':<20} {'DESCRIÇÃO':<72} {'VALOR':>14}")
        for cat, desc, valor in desp:
            lines.append(f"{upper(cat or 'OUTROS')[:20]:<20} {upper(desc or '-')[:72]:<72} {self._fmt_rel_money(valor):>14}")
        if not desp:
            lines.append("Sem despesas.")
        lines.append("-" * 118)
        lines.append("[RASTREABILIDADE LOGÃÂSTICA]")
        for ln in (log_info.get("resumo") or []):
            lines.append(str(ln))
        lines.append(
            "Conciliação de caixas: "
            + ("OK" if bool(log_info.get("itens_ok", True)) else "DIVERGENTE")
        )
        lines.append("-" * 118)
        lines.append("[OCORRÊNCIAS DE TRANSBORDO / MORTALIDADE]")
        lines.append(f"Mortalidade (aves): {safe_int(mort_aves, 0)}")
        lines.append(
            f"Mortalidade (KG): {safe_float(mort_kg, 0.0):.2f}".replace(".", ",")
        )
        lines.append(
            f"KG útil NF (NF - mortalidade): {safe_float(kg_nf_util, 0.0):.2f}".replace(".", ",")
        )
        lines.append(f"Obs transbordo: {str(obs_transb or '-').strip() or '-'}")
        lines.append("-" * 118)
        lines.append("Conferido por: ____________________________________   Data: ____/____/________")
        return "\n".join(lines)

    def _gerar_relatorio_rotina_motoristas_ajudantes(self):
        filtro_mot = upper(self.ent_filtro_motorista.get().strip()) if hasattr(self, "ent_filtro_motorista") else ""
        rows = []
        api_succeeded = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                q = urllib.parse.quote(filtro_mot, safe="")
                resp = _call_api(
                    "GET",
                    f"desktop/relatorios/rotina-motoristas?motorista_like={q}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("codigo_programacao") or ""),
                            str((r or {}).get("motorista") or ""),
                            str((r or {}).get("equipe") or ""),
                            str((r or {}).get("status") or ""),
                            safe_float((r or {}).get("kg_vendido"), 0.0),
                            safe_float((r or {}).get("km_rodado"), 0.0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    api_succeeded = True
            except Exception:
                logging.debug("Falha no relatorio de rotina via API; usando fallback local.", exc_info=True)

        if not api_succeeded:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
                cur.execute(f"""
                    SELECT COALESCE(codigo_programacao,''), COALESCE(motorista,''), COALESCE(equipe,''), COALESCE(status,''), COALESCE(kg_vendido,0), {km_expr}
                    FROM programacoes
                    ORDER BY id DESC
                """)
                rows = cur.fetchall() or []

        mot = {}
        aju = {}
        for codigo, motorista, equipe, status, kg_vendido, km_rodado in rows:
            m = upper(motorista or "SEM MOTORISTA")
            if filtro_mot and filtro_mot not in m:
                continue
            d = mot.setdefault(m, {"viagens": 0, "kg": 0.0, "km": 0.0, "em_rota": 0})
            d["viagens"] += 1
            d["kg"] += safe_float(kg_vendido, 0.0)
            d["km"] += safe_float(km_rodado, 0.0)
            if upper(status) in ("EM_ROTA", "ATIVA", "CARREGADA", "INICIADA"):
                d["em_rota"] += 1
            for nm in re.split(r"[|,;/]+", resolve_equipe_nomes(equipe) or ""):
                an = upper((nm or "").strip())
                if not an or an in ("-", "NAN", "NONE", "SEM EQUIPE"):
                    continue
                a = aju.setdefault(an, {"viagens": 0, "km": 0.0})
                a["viagens"] += 1
                a["km"] += safe_float(km_rodado, 0.0)

        rank_m = sorted(mot.items(), key=lambda kv: (-kv[1]["viagens"], kv[0]))
        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO DE ROTINA - MOTORISTAS E AJUDANTES\n")
        self.txt.insert("end", "=" * 100 + "\n\n")
        self.txt.insert("end", "[MOTORISTAS]\n")
        self.txt.insert("end", "MOTORISTA | VIAGENS | KG ENTREGUES | KM RODADO | MEDIA KM/VIAGEM | EM ROTA\n")
        self.txt.insert("end", "-" * 100 + "\n")
        for nome, d in rank_m:
            media_km = d["km"] / max(d["viagens"], 1)
            self.txt.insert("end", f"{nome} | {d['viagens']} | {d['kg']:.2f} | {d['km']:.2f} | {media_km:.2f} | {d['em_rota']}\n")
        self.txt.insert("end", "\n[AJUDANTES]\n")
        self.txt.insert("end", "AJUDANTE | VIAGENS | KM RODADO | MEDIA KM/VIAGEM\n")
        self.txt.insert("end", "-" * 100 + "\n")
        for nome, d in sorted(aju.items(), key=lambda kv: (-kv[1]["viagens"], kv[0])):
            media_km = d["km"] / max(d["viagens"], 1)
            self.txt.insert("end", f"{nome} | {d['viagens']} | {d['km']:.2f} | {media_km:.2f}\n")

        top_nome = rank_m[0][0] if rank_m else "-"
        top_viagens = rank_m[0][1]["viagens"] if rank_m else 0
        self._set_dashboard(
            f"Registros: {len(rank_m)} motoristas",
            f"Total viagens: {sum(d['viagens'] for _n, d in rank_m)}",
            f"Total KG: {sum(d['kg'] for _n, d in rank_m):.2f}",
            f"Destaque: {top_nome} ({top_viagens} viagens)",
            [n for n, _d in rank_m[:8]],
            [d["viagens"] for _n, d in rank_m[:8]],
            color="#0EA5E9",
        )
        self.set_status(f"STATUS: Relatório de rotina gerado ({len(rank_m)} motorista(s)).")

    def _gerar_relatorio_km_veiculos(self):
        rows = []
        api_succeeded = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/relatorios/km-veiculos",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("veiculo") or ""),
                            safe_int((r or {}).get("viagens"), 0),
                            safe_float((r or {}).get("km_rodado"), 0.0),
                            safe_float((r or {}).get("media_km_l"), 0.0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    api_succeeded = True
            except Exception:
                logging.debug("Falha no relatorio KM via API; usando fallback local.", exc_info=True)

        if not api_succeeded:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                km_expr = "COALESCE(km_rodado,0)" if "km_rodado" in cols else "0"
                media_expr = "COALESCE(media_km_l,0)" if "media_km_l" in cols else "0"
                cur.execute(f"""
                    SELECT COALESCE(veiculo,''), COUNT(1), COALESCE(SUM({km_expr}),0), COALESCE(AVG(NULLIF({media_expr},0)),0)
                    FROM programacoes
                    GROUP BY COALESCE(veiculo,'')
                    ORDER BY 3 DESC, 1 ASC
                """)
                rows = cur.fetchall() or []
        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO DE KM POR VEICULO\n")
        self.txt.insert("end", "=" * 90 + "\n\n")
        self.txt.insert("end", "VEICULO | VIAGENS | KM RODADO | MELHOR MEDIA KM/L\n")
        self.txt.insert("end", "-" * 90 + "\n")
        for veic, viagens, km, media in rows:
            self.txt.insert("end", f"{upper(veic or '-')} | {safe_int(viagens, 0)} | {safe_float(km, 0.0):.2f} | {safe_float(media, 0.0):.2f}\n")
        total_km = sum(safe_float(r[2], 0.0) for r in rows)
        top = rows[0][0] if rows else "-"
        self._set_dashboard(
            f"VeÃÂculos: {len(rows)}",
            f"KM total: {total_km:.2f}",
            f"Média KM/veÃÂculo: {(total_km / max(len(rows), 1)):.2f}",
            f"Destaque: {upper(top or '-')}",
            [upper((r[0] or "-")) for r in rows[:8]],
            [safe_float(r[2], 0.0) for r in rows[:8]],
            color="#2563EB",
        )
        self.set_status(f"STATUS: Relatório de KM por veÃÂculo gerado ({len(rows)} veÃÂculo(s)).")

    def _gerar_relatorio_despesas_geral(self):
        rows = []
        api_succeeded = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        if desktop_secret and is_desktop_api_sync_enabled():
            try:
                resp = _call_api(
                    "GET",
                    "desktop/relatorios/despesas-categorias",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("categoria") or "OUTROS"),
                            safe_int((r or {}).get("qtd"), 0),
                            safe_float((r or {}).get("total"), 0.0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    api_succeeded = True
            except Exception:
                logging.debug("Falha no relatorio de despesas via API; usando fallback local.", exc_info=True)

        if not api_succeeded:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT COALESCE(categoria,'OUTROS') AS categoria, COUNT(1) AS qtd, COALESCE(SUM(valor),0) AS total
                    FROM despesas
                    GROUP BY COALESCE(categoria,'OUTROS')
                    ORDER BY total DESC, categoria ASC
                """)
                rows = cur.fetchall() or []
        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO GERAL DE DESPESAS\n")
        self.txt.insert("end", "=" * 90 + "\n\n")
        self.txt.insert("end", "CATEGORIA | QTD LANÇAMENTOS | TOTAL\n")
        self.txt.insert("end", "-" * 90 + "\n")
        for cat, qtd, total in rows:
            self.txt.insert("end", f"{upper(cat or 'OUTROS')} | {safe_int(qtd, 0)} | {self._fmt_rel_money(total)}\n")
        total_geral = sum(safe_float(r[2], 0.0) for r in rows)
        top = upper(rows[0][0]) if rows else "-"
        self._set_dashboard(
            f"Categorias: {len(rows)}",
            f"Total despesas: {self._fmt_rel_money(total_geral)}",
            f"Média/categoria: {self._fmt_rel_money(total_geral / max(len(rows), 1))}",
            f"Maior categoria: {top}",
            [upper((r[0] or "OUTROS")) for r in rows[:8]],
            [safe_float(r[2], 0.0) for r in rows[:8]],
            color="#DC2626",
        )
        self.set_status(f"STATUS: Relatório de despesas gerado ({len(rows)} categoria(s)).")

    def gerar_resumo(self):
        tipo_rel = self.cb_tipo_rel.get().strip() if hasattr(self, "cb_tipo_rel") else "Programacoes"
        is_mortalidade = "MORTALIDADE" in upper(tipo_rel)
        is_rotina = "ROTINA" in upper(tipo_rel)
        is_km_veic = "KM DE VEICULOS" in upper(tipo_rel)
        is_despesas = upper(tipo_rel) == "DESPESAS"
        if is_mortalidade:
            self._gerar_resumo_mortalidade_motorista()
            return
        if is_rotina:
            self._gerar_relatorio_rotina_motoristas_ajudantes()
            return
        if is_km_veic:
            self._gerar_relatorio_km_veiculos()
            return
        if is_despesas:
            self._gerar_relatorio_despesas_geral()
            return

        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return

        is_prestacao = "PRESTACAO" in upper(tipo_rel)

        def fmt_money(v):
            vv = safe_float(v, 0.0)
            return f"R$ {vv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        def normalize_unit_price(v):
            vv = safe_float(v, 0.0)
            if abs(vv) >= 100 and abs(vv - round(vv)) < 1e-9:
                return vv / 100.0
            return vv

        motorista = veiculo = equipe = ""
        usuario_criacao = ""
        status = prestacao = ""
        data_criacao = ""
        nf = local_rota = local_carreg = ""
        data_saida = hora_saida = data_chegada = hora_chegada = ""
        kg_estimado = 0.0
        nf_kg = nf_caixas = nf_kg_carregado = nf_kg_vendido = nf_saldo = 0.0
        km_inicial = km_final = litros = km_rodado = media_km_l = custo_km = 0.0
        ced_qtd = {200: 0, 100: 0, 50: 0, 20: 0, 10: 0, 5: 0, 2: 0}
        valor_dinheiro = 0.0
        adiantamento = 0.0
        total_entregas = 0
        total_receb = 0.0
        total_desp = 0.0
        recebimentos = []
        despesas = []
        clientes_programacao = []

        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
        bundle = self._api_bundle_relatorio(prog)
        api_state = str(getattr(self, "_last_bundle_api_state", ""))
        if api_enabled and api_state == "not_found":
            messagebox.showwarning("ATENCAO", f"Programacao nao encontrada no servidor: {prog}")
            return
        if bundle:
            try:
                rota = bundle.get("rota") or {}
                clientes_api = [r for r in (bundle.get("clientes") or []) if isinstance(r, dict)]
                receb_api = [r for r in (bundle.get("recebimentos") or []) if isinstance(r, dict)]
                desp_api = [r for r in (bundle.get("despesas") or []) if isinstance(r, dict)]

                motorista = str(rota.get("motorista") or "")
                veiculo = str(rota.get("veiculo") or "")
                equipe = str(rota.get("equipe") or "")
                status = upper(str(rota.get("status") or ""))
                prestacao = upper(str(rota.get("prestacao_status") or ""))
                data_criacao = str(rota.get("data_criacao") or rota.get("data") or "")
                usuario_criacao = str(
                    rota.get("usuario")
                    or rota.get("criado_por")
                    or rota.get("usuario_criacao")
                    or rota.get("created_by")
                    or rota.get("autor")
                    or rota.get("responsavel")
                    or ""
                )
                kg_estimado = safe_float(rota.get("kg_estimado"), 0.0)
                nf = str(rota.get("num_nf") or rota.get("nf_numero") or "")
                local_rota = str(rota.get("local_rota") or rota.get("tipo_rota") or "")
                local_carreg = str(rota.get("local_carregamento") or rota.get("granja_carregada") or "")
                data_saida = str(rota.get("data_saida") or "")
                hora_saida = str(rota.get("hora_saida") or "")
                data_chegada = str(rota.get("data_chegada") or "")
                hora_chegada = str(rota.get("hora_chegada") or "")
                nf_kg = safe_float(rota.get("nf_kg"), 0.0)
                nf_caixas = safe_int(rota.get("nf_caixas"), 0)
                nf_kg_carregado = safe_float(rota.get("nf_kg_carregado"), 0.0)
                nf_kg_vendido = safe_float(rota.get("nf_kg_vendido"), 0.0)
                nf_saldo = safe_float(rota.get("nf_saldo"), 0.0)
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
                adiantamento = safe_float(rota.get("adiantamento"), 0.0)

                data_saida, hora_saida = normalize_date_time_components(data_saida, hora_saida)
                data_chegada, hora_chegada = normalize_date_time_components(data_chegada, hora_chegada)

                clientes_programacao = [
                    (
                        str((r or {}).get("cod_cliente") or ""),
                        str((r or {}).get("nome_cliente") or ""),
                        safe_float((r or {}).get("preco_atual"), safe_float((r or {}).get("preco"), 0.0)),
                        str((r or {}).get("vendedor") or ""),
                    )
                    for r in clientes_api
                ]
                total_entregas = len(clientes_programacao)

                if is_prestacao:
                    recebimentos = [
                        (
                            str((r or {}).get("cod_cliente") or ""),
                            str((r or {}).get("nome_cliente") or ""),
                            safe_float((r or {}).get("valor"), 0.0),
                            str((r or {}).get("forma_pagamento") or ""),
                            str((r or {}).get("observacao") or ""),
                            str((r or {}).get("data_registro") or ""),
                        )
                        for r in receb_api
                    ]
                    despesas = [
                        (
                            str((r or {}).get("descricao") or ""),
                            safe_float((r or {}).get("valor"), 0.0),
                            str((r or {}).get("categoria") or "OUTROS"),
                            str((r or {}).get("observacao") or ""),
                            str((r or {}).get("data_registro") or ""),
                        )
                        for r in desp_api
                    ]
                    total_receb = sum(safe_float(r[2], 0.0) for r in recebimentos)
                    total_desp = sum(safe_float(r[1], 0.0) for r in despesas)
            except Exception:
                logging.debug("Falha ao carregar resumo via API.", exc_info=True)
                if api_enabled:
                    messagebox.showwarning(
                        "Relatorios",
                        "Falha ao processar dados retornados pelo servidor para esta programacao.\n"
                        "Tente novamente apos atualizar o servidor.",
                    )
                    return
                bundle = None

        if not bundle:
            with get_db() as conn:
                cur = conn.cursor()
                cols = []

                def has_col(col):
                    try:
                        return db_has_column(cur, "programacoes", col)
                    except Exception:
                        return False

                try:
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols = [str(c[1]).lower() for c in (cur.fetchall() or [])]
                except Exception:
                    cols = []

                usuario_col_name = ""
                for cand in ("usuario", "criado_por", "usuario_criacao", "user", "created_by", "autor", "responsavel"):
                    if cand in cols:
                        usuario_col_name = cand
                        break

                nf_col = "num_nf" if has_col("num_nf") else ("nf_numero" if has_col("nf_numero") else "'' as num_nf")
                local_rota_col = "local_rota" if has_col("local_rota") else ("tipo_rota" if has_col("tipo_rota") else "'' as local_rota")
                local_carreg_col = "local_carregamento" if has_col("local_carregamento") else ("granja_carregada" if has_col("granja_carregada") else "'' as local_carregamento")

                if has_col("adiantamento") and has_col("adiantamento_rota"):
                    adiant_col = (
                        "CASE WHEN COALESCE(adiantamento, 0) <> 0 THEN adiantamento "
                        "ELSE COALESCE(adiantamento_rota, 0) END as adiantamento"
                    )
                elif has_col("adiantamento"):
                    adiant_col = "adiantamento"
                elif has_col("adiantamento_rota"):
                    adiant_col = "adiantamento_rota as adiantamento"
                else:
                    adiant_col = "0 as adiantamento"

                status_col = "COALESCE(status,'') as status" if has_col("status") else "'' as status"
                prest_col = "COALESCE(prestacao_status,'') as prestacao_status" if has_col("prestacao_status") else "'' as prestacao_status"
                data_col = "COALESCE(data_criacao,'') as data_criacao" if has_col("data_criacao") else ("COALESCE(data,'') as data_criacao" if has_col("data") else "'' as data_criacao")
                usuario_col = f"COALESCE({usuario_col_name},'') as usuario_criacao" if usuario_col_name else "'' as usuario_criacao"

                cur.execute(f"""
                    SELECT
                        motorista, veiculo, equipe,
                        {status_col}, {prest_col}, {data_col},
                        {usuario_col},
                        COALESCE(kg_estimado,0),
                        {nf_col} as num_nf,
                        {local_rota_col} as local_rota,
                        {local_carreg_col} as local_carreg,
                        COALESCE(data_saida,''), COALESCE(hora_saida,''),
                        COALESCE(data_chegada,''), COALESCE(hora_chegada,''),
                        COALESCE(nf_kg,0), COALESCE(nf_caixas,0), COALESCE(nf_kg_carregado,0),
                        COALESCE(nf_kg_vendido,0), COALESCE(nf_saldo,0),
                        COALESCE(km_inicial,0), COALESCE(km_final,0), COALESCE(litros,0),
                        COALESCE(km_rodado,0), COALESCE(media_km_l,0), COALESCE(custo_km,0),
                        COALESCE(ced_200_qtd,0), COALESCE(ced_100_qtd,0), COALESCE(ced_50_qtd,0),
                        COALESCE(ced_20_qtd,0), COALESCE(ced_10_qtd,0), COALESCE(ced_5_qtd,0), COALESCE(ced_2_qtd,0),
                        COALESCE(valor_dinheiro,0),
                        {adiant_col}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                row = cur.fetchone()
                if not row:
                    messagebox.showwarning("ATENCAO", f"Programacao nao encontrada: {prog}")
                    return

                motorista, veiculo, equipe = row[0] or "", row[1] or "", row[2] or ""
                status, prestacao = upper(row[3] or ""), upper(row[4] or "")
                data_criacao = row[5] or ""
                usuario_criacao = row[6] or ""
                kg_estimado = safe_float(row[7], 0.0)
                nf, local_rota, local_carreg = row[8] or "", row[9] or "", row[10] or ""
                data_saida, hora_saida = row[11] or "", row[12] or ""
                data_chegada, hora_chegada = row[13] or "", row[14] or ""
                nf_kg, nf_caixas, nf_kg_carregado = safe_float(row[15], 0.0), safe_int(row[16], 0), safe_float(row[17], 0.0)
                nf_kg_vendido, nf_saldo = safe_float(row[18], 0.0), safe_float(row[19], 0.0)
                km_inicial, km_final = safe_float(row[20], 0.0), safe_float(row[21], 0.0)
                litros, km_rodado = safe_float(row[22], 0.0), safe_float(row[23], 0.0)
                media_km_l, custo_km = safe_float(row[24], 0.0), safe_float(row[25], 0.0)
                ced_qtd[200], ced_qtd[100], ced_qtd[50] = safe_int(row[26], 0), safe_int(row[27], 0), safe_int(row[28], 0)
                ced_qtd[20], ced_qtd[10], ced_qtd[5], ced_qtd[2] = safe_int(row[29], 0), safe_int(row[30], 0), safe_int(row[31], 0), safe_int(row[32], 0)
                valor_dinheiro = safe_float(row[33], 0.0)
                adiantamento = safe_float(row[34], 0.0)

                data_saida, hora_saida = normalize_date_time_components(data_saida, hora_saida)
                data_chegada, hora_chegada = normalize_date_time_components(data_chegada, hora_chegada)

                try:
                    cur.execute("SELECT COUNT(*) FROM programacao_itens WHERE codigo_programacao=?", (prog,))
                    total_entregas = safe_int((cur.fetchone() or [0])[0], 0)
                except Exception:
                    total_entregas = 0

                cur.execute("""
                    SELECT COALESCE(cod_cliente,''), COALESCE(nome_cliente,''), COALESCE(preco,0), COALESCE(vendedor,'')
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                    ORDER BY nome_cliente ASC, cod_cliente ASC
                """, (prog,))
                clientes_programacao = cur.fetchall() or []

                if is_prestacao:
                    cur.execute("SELECT COALESCE(SUM(valor),0) FROM recebimentos WHERE codigo_programacao=?", (prog,))
                    total_receb = safe_float((cur.fetchone() or [0])[0], 0.0)

                    cur.execute("SELECT COALESCE(SUM(valor),0) FROM despesas WHERE codigo_programacao=?", (prog,))
                    total_desp = safe_float((cur.fetchone() or [0])[0], 0.0)

                    cur.execute("""
                        SELECT COALESCE(cod_cliente,''), COALESCE(nome_cliente,''), COALESCE(valor,0),
                               COALESCE(forma_pagamento,''), COALESCE(observacao,''), COALESCE(data_registro,'')
                        FROM recebimentos
                        WHERE codigo_programacao=?
                        ORDER BY data_registro DESC, id DESC
                    """, (prog,))
                    recebimentos = cur.fetchall() or []

                    cur.execute("""
                        SELECT COALESCE(descricao,''), COALESCE(valor,0), COALESCE(categoria,'OUTROS'),
                               COALESCE(observacao,''), COALESCE(data_registro,'')
                        FROM despesas
                        WHERE codigo_programacao=?
                        ORDER BY data_registro DESC, id DESC
                    """, (prog,))
                    despesas = cur.fetchall() or []

        equipe_txt = resolve_equipe_nomes(equipe)
        data_criacao_n = normalize_date(data_criacao)
        data_criacao_show = data_criacao_n if data_criacao_n is not None else (data_criacao or "")

        self.txt.delete("1.0", "end")

        if not is_prestacao:
            valor_total_clientes = sum(normalize_unit_price(r[2]) for r in clientes_programacao)

            self.txt.insert("end", f"RELATORIO DE PROGRAMACAO - {prog}\n")
            self.txt.insert("end", "=" * 90 + "\n\n")

            self.txt.insert("end", "[IDENTIFICACAO]\n")
            self.txt.insert("end", f"Codigo: {prog}\n")
            self.txt.insert("end", f"Status: {status or '-'}\n")
            self.txt.insert("end", f"Data: {data_criacao_show or '-'}\n")
            self.txt.insert("end", f"Usuario criacao: {usuario_criacao or '-'}\n")
            self.txt.insert("end", f"Motorista: {motorista or '-'}\n")
            self.txt.insert("end", f"Equipe: {equipe_txt or '-'}\n")
            self.txt.insert("end", f"Veiculo: {veiculo or '-'}\n")
            self.txt.insert("end", f"KG estimado: {kg_estimado:.2f}\n")
            self.txt.insert("end", f"Clientes na programacao: {len(clientes_programacao)}\n")
            self.txt.insert("end", f"Total estimado (clientes): {fmt_money(valor_total_clientes)}\n\n")

            self.txt.insert("end", "[CLIENTES / PRECO / VENDEDOR]\n")
            if not clientes_programacao:
                self.txt.insert("end", "Sem clientes cadastrados na programacao.\n")
            else:
                self.txt.insert("end", "COD | CLIENTE | PRECO | VENDEDOR\n")
                self.txt.insert("end", "-" * 90 + "\n")
                for cod_cli, nome_cli, preco_cli, vendedor in clientes_programacao:
                    self.txt.insert(
                        "end",
                        f"{(cod_cli or '-')} | {(nome_cli or '-')} | {fmt_money(normalize_unit_price(preco_cli))} | {(vendedor or '-')}\n"
                    )

            self.set_status("STATUS: Resumo de programacao gerado.")
            self._set_dashboard(
                f"Clientes: {len(clientes_programacao)}",
                f"Total estimado: {fmt_money(valor_total_clientes)}",
                f"Preço médio: {fmt_money(valor_total_clientes / max(len(clientes_programacao), 1))}",
                f"Motorista: {motorista or '-'}",
                [str(r[0]) for r in clientes_programacao[:8]],
                [normalize_unit_price(r[2]) for r in clientes_programacao[:8]],
                color="#1D4ED8",
            )
            return

        ced_total = sum(float(ced) * safe_int(qtd, 0) for ced, qtd in ced_qtd.items())
        dinheiro_total = valor_dinheiro if safe_float(valor_dinheiro, 0.0) > 0 else ced_total
        total_entradas = total_receb + adiantamento
        total_saidas = total_desp + ced_total
        valor_final_caixa = total_entradas - total_desp
        diferenca = valor_final_caixa - ced_total
        resultado = total_entradas - total_saidas

        self.txt.insert("end", f"RELATORIO DE PRESTACAO DE CONTAS - PROGRAMACAO {prog}\n")
        self.txt.insert("end", f"Tipo: {tipo_rel}\n")
        self.txt.insert("end", "=" * 90 + "\n\n")

        self.txt.insert("end", "[IDENTIFICACAO]\n")
        self.txt.insert("end", f"Status: {status or '-'}\n")
        self.txt.insert("end", f"Prestacao: {prestacao or '-'}\n")
        self.txt.insert("end", f"Data criacao: {data_criacao_show or '-'}\n")
        self.txt.insert("end", f"Motorista: {motorista or '-'}\n")
        self.txt.insert("end", f"Veiculo: {veiculo or '-'}\n")
        self.txt.insert("end", f"Equipe: {equipe_txt or '-'}\n")
        self.txt.insert("end", f"Entregas (itens): {total_entregas}\n")
        self.txt.insert("end", f"KG estimado: {kg_estimado:.2f}\n\n")

        self.txt.insert("end", "[DADOS DA ROTA]\n")
        self.txt.insert("end", f"NF: {nf or '-'}\n")
        self.txt.insert("end", f"Local da rota: {local_rota or '-'}\n")
        self.txt.insert("end", f"Local carregamento: {local_carreg or '-'}\n")
        self.txt.insert("end", f"Saida: {(data_saida + ' ' + hora_saida).strip() or '-'}\n")
        self.txt.insert("end", f"Chegada: {(data_chegada + ' ' + hora_chegada).strip() or '-'}\n\n")

        self.txt.insert("end", "[NOTA FISCAL / CARREGAMENTO]\n")
        self.txt.insert("end", f"NF KG: {nf_kg:.2f}\n")
        self.txt.insert("end", f"NF caixas: {safe_int(nf_caixas, 0)}\n")
        self.txt.insert("end", f"KG carregado: {nf_kg_carregado:.2f}\n")
        self.txt.insert("end", f"KG vendido: {nf_kg_vendido:.2f}\n")
        self.txt.insert("end", f"Saldo (KG): {nf_saldo:.2f}\n\n")

        self.txt.insert("end", "[ROTA / KM]\n")
        self.txt.insert("end", f"KM inicial: {km_inicial:.2f}\n")
        self.txt.insert("end", f"KM final: {km_final:.2f}\n")
        self.txt.insert("end", f"Litros: {litros:.2f}\n")
        self.txt.insert("end", f"KM rodado: {km_rodado:.2f}\n")
        self.txt.insert("end", f"Media km/l: {media_km_l:.2f}\n")
        self.txt.insert("end", f"Custo por KM: {custo_km:.2f}\n\n")

        self.txt.insert("end", "[CONTAGEM DE CEDULAS]\n")
        for ced in [200, 100, 50, 20, 10, 5, 2]:
            qtd = safe_int(ced_qtd.get(ced, 0), 0)
            self.txt.insert("end", f"R$ {ced:>3},00 -> QTD {qtd:>4} -> TOTAL {fmt_money(qtd * ced)}\n")
        self.txt.insert("end", f"Total cedulas: {fmt_money(ced_total)}\n")
        self.txt.insert("end", f"Total dinheiro (campo): {fmt_money(dinheiro_total)}\n\n")

        self.txt.insert("end", "[RESUMO FINANCEIRO]\n")
        self.txt.insert("end", f"Recebimentos: {fmt_money(total_receb)}\n")
        self.txt.insert("end", f"Adiantamento: {fmt_money(adiantamento)}\n")
        self.txt.insert("end", f"Despesas: {fmt_money(total_desp)}\n")
        self.txt.insert("end", f"Cedulas: {fmt_money(ced_total)}\n")
        self.txt.insert("end", f"Total entradas: {fmt_money(total_entradas)}\n")
        self.txt.insert("end", f"Total saidas: {fmt_money(total_saidas)}\n")
        self.txt.insert("end", f"Valor final caixa: {fmt_money(valor_final_caixa)}\n")
        self.txt.insert("end", f"Diferenca caixa x cedulas: {fmt_money(diferenca)}\n")
        self.txt.insert("end", f"Resultado liquido: {fmt_money(resultado)}\n\n")

        if bool(self.var_show_receb_detalhe.get()):
            self.txt.insert("end", "[RECEBIMENTOS DETALHADOS]\n")
            if not recebimentos:
                self.txt.insert("end", "Sem recebimentos registrados.\n\n")
            else:
                for cod_cli, nome_cli, valor, forma, obs, data_reg in recebimentos:
                    data_show = (data_reg or "")[:19]
                    self.txt.insert("end", f"{data_show} | {cod_cli} | {nome_cli} | {fmt_money(valor)} | {forma or '-'} | {obs or '-'}\n")
                self.txt.insert("end", "\n")

        if bool(self.var_show_desp_detalhe.get()):
            self.txt.insert("end", "[DESPESAS DETALHADAS]\n")
            if not despesas:
                self.txt.insert("end", "Sem despesas registradas.\n")
            else:
                for desc, valor, cat, obs, data_reg in despesas:
                    data_show = (data_reg or "")[:19]
                    self.txt.insert("end", f"{data_show} | {cat or 'OUTROS'} | {desc or '-'} | {fmt_money(valor)} | {obs or '-'}\n")

        if prestacao == "FECHADA":
            self.txt.insert("end", "\n[ALERTA] Prestacao FECHADA: alteracoes financeiras estao bloqueadas.\n")

        self.set_status("STATUS: Resumo detalhado gerado.")
        self._set_dashboard(
            f"Recebimentos: {fmt_money(total_receb)}",
            f"Despesas: {fmt_money(total_desp)}",
            f"Resultado: {fmt_money(resultado)}",
            f"Prestação: {prestacao or '-'}",
            ["Entradas", "SaÃÂdas", "Resultado"],
            [total_entradas, total_saidas, abs(resultado)],
            color="#0EA5E9",
        )

    def _gerar_resumo_mortalidade_motorista(self):
        tipo_rel = self.cb_tipo_rel.get().strip() if hasattr(self, "cb_tipo_rel") else "Mortalidade Motorista"
        filtro_cod = upper(self.ent_filtro_codigo.get().strip()) if hasattr(self, "ent_filtro_codigo") else ""
        filtro_mot = upper(self.ent_filtro_motorista.get().strip()) if hasattr(self, "ent_filtro_motorista") else ""
        filtro_data_raw = str(self.ent_filtro_data.get().strip()) if hasattr(self, "ent_filtro_data") else ""

        data_patterns = []
        if filtro_data_raw:
            n = normalize_date(filtro_data_raw)
            if n:
                data_patterns.append(n)
                data_patterns.append(f"{n[8:10]}/{n[5:7]}/{n[0:4]}")
            data_patterns.append(filtro_data_raw)
            data_patterns = [p for p in dict.fromkeys(data_patterns) if p]

        rows = []
        used_api = False
        api_failed = False
        desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
        api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
        if api_enabled:
            try:
                q_cod = urllib.parse.quote(filtro_cod, safe="")
                q_mot = urllib.parse.quote(filtro_mot, safe="")
                q_data = urllib.parse.quote("|".join(data_patterns), safe="")
                resp = _call_api(
                    "GET",
                    f"desktop/relatorios/mortalidade-motorista?codigo_like={q_cod}&motorista_like={q_mot}&data_like={q_data}",
                    extra_headers={"X-Desktop-Secret": desktop_secret},
                )
                if isinstance(resp, dict):
                    rows = [
                        (
                            str((r or {}).get("codigo_programacao") or ""),
                            str((r or {}).get("motorista") or ""),
                            str((r or {}).get("data_ref") or ""),
                            str((r or {}).get("status_ref") or ""),
                            safe_int((r or {}).get("mortalidade_total"), 0),
                            safe_int((r or {}).get("clientes_com_mortalidade"), 0),
                        )
                        for r in (resp.get("rows") or [])
                        if isinstance(r, dict)
                    ]
                    used_api = True
                else:
                    messagebox.showwarning(
                        "Relatorios",
                        "Resposta invalida do servidor para o relatorio de mortalidade.",
                    )
                    return
            except Exception:
                api_failed = True
                logging.debug("Falha no relatorio de mortalidade via API; usando fallback local.", exc_info=True)

        if not used_api:
            if api_enabled and (not api_failed):
                messagebox.showwarning(
                    "Relatorios",
                    "Nao foi possivel validar os dados do relatorio de mortalidade no servidor.",
                )
                return
            with get_db() as conn:
                cur = conn.cursor()

                cur.execute("PRAGMA table_info(programacoes)")
                cols_p = [str(c[1]).lower() for c in (cur.fetchall() or [])]
                has_status = "status" in cols_p
                has_data_criacao = "data_criacao" in cols_p
                has_data = "data" in cols_p
                data_expr = (
                    "COALESCE(p.data_criacao,'')"
                    if has_data_criacao
                    else ("COALESCE(p.data,'')" if has_data else "''")
                )
                status_expr = "COALESCE(p.status,'')" if has_status else "''"

                sql = f"""
                    SELECT
                        COALESCE(p.codigo_programacao,'') as codigo_programacao,
                        COALESCE(p.motorista,'') as motorista,
                        {data_expr} as data_ref,
                        {status_expr} as status_ref,
                        COALESCE(SUM(COALESCE(pc.mortalidade_aves, 0)), 0) as mortalidade_total,
                        COUNT(CASE WHEN COALESCE(pc.mortalidade_aves,0) > 0 THEN 1 END) as clientes_com_mortalidade
                    FROM programacoes p
                    LEFT JOIN programacao_itens_controle pc
                      ON UPPER(COALESCE(pc.codigo_programacao,'')) = UPPER(COALESCE(p.codigo_programacao,''))
                    WHERE 1=1
                """
                params = []

                if filtro_cod:
                    sql += " AND UPPER(COALESCE(p.codigo_programacao,'')) LIKE ?"
                    params.append(f"%{filtro_cod}%")
                if filtro_mot:
                    sql += " AND UPPER(COALESCE(p.motorista,'')) LIKE ?"
                    params.append(f"%{filtro_mot}%")
                if data_patterns:
                    clauses = []
                    for pat in data_patterns:
                        clauses.append(f"{data_expr} LIKE ?")
                        params.append(f"%{pat}%")
                    sql += " AND (" + " OR ".join(clauses) + ")"

                sql += """
                    GROUP BY p.codigo_programacao, p.motorista, data_ref, status_ref
                    ORDER BY mortalidade_total ASC, p.codigo_programacao DESC
                """
                cur.execute(sql, tuple(params))
                rows = cur.fetchall() or []

        self.txt.delete("1.0", "end")
        self.txt.insert("end", "RELATORIO DE MORTALIDADE POR MOTORISTA\n")
        self.txt.insert("end", f"Tipo: {tipo_rel}\n")
        self.txt.insert("end", "=" * 95 + "\n\n")

        if not rows:
            self.txt.insert("end", "Nenhum dado encontrado para os filtros informados.\n")
            self.set_status("STATUS: Nenhum dado de mortalidade encontrado.")
            return

        # Ranking por rota (menor mortalidade primeiro)
        self.txt.insert("end", "[RANKING POR ROTA - MENOR MORTALIDADE]\n")
        self.txt.insert("end", "POS | PROGRAMAÇÃO | MOTORISTA | MORTALIDADE | CLIENTES C/ MORT. | DATA | STATUS\n")
        self.txt.insert("end", "-" * 95 + "\n")
        for i, (codigo, motorista, data_ref, status_ref, mort_total, cli_mort) in enumerate(rows, start=1):
            self.txt.insert(
                "end",
                f"{i:>3} | {codigo or '-'} | {(motorista or '-')} | {safe_int(mort_total, 0):>11} | {safe_int(cli_mort, 0):>16} | {(data_ref or '-')[:19]} | {upper(status_ref or '-')}\n",
            )

        # Consolidado por motorista
        resumo_mot = {}
        for codigo, motorista, data_ref, status_ref, mort_total, cli_mort in rows:
            mot = upper(motorista or "SEM MOTORISTA")
            item = resumo_mot.setdefault(
                mot,
                {"rotas": 0, "mort_total": 0, "melhor": None, "pior": 0},
            )
            mort = safe_int(mort_total, 0)
            item["rotas"] += 1
            item["mort_total"] += mort
            item["pior"] = max(item["pior"], mort)
            if item["melhor"] is None or mort < item["melhor"]:
                item["melhor"] = mort

        self.txt.insert("end", "\n[CONSOLIDADO POR MOTORISTA]\n")
        self.txt.insert("end", "MOTORISTA | ROTAS | MORTALIDADE TOTAL | MEDIA/ROTA | MELHOR ROTA | PIOR ROTA\n")
        self.txt.insert("end", "-" * 95 + "\n")
        ranking_motorista = sorted(
            resumo_mot.items(),
            key=lambda kv: (
                (kv[1]["mort_total"] / max(kv[1]["rotas"], 1)),
                kv[1]["mort_total"],
                kv[0],
            ),
        )
        for mot, d in ranking_motorista:
            media = d["mort_total"] / max(d["rotas"], 1)
            self.txt.insert(
                "end",
                f"{mot} | {d['rotas']} | {d['mort_total']} | {media:.2f} | {safe_int(d['melhor'], 0)} | {safe_int(d['pior'], 0)}\n",
            )

        melhor_mot, melhor_data = ranking_motorista[0]
        melhor_media = melhor_data["mort_total"] / max(melhor_data["rotas"], 1)
        self.txt.insert("end", "\n[DESTAQUE]\n")
        self.txt.insert(
            "end",
            f"Motorista com menor média de mortalidade por rota: {melhor_mot} "
            f"(média {melhor_media:.2f} aves/rota em {melhor_data['rotas']} rota(s)).\n",
        )
        self._set_dashboard(
            f"Rotas analisadas: {len(rows)}",
            f"Motoristas: {len(resumo_mot)}",
            f"Melhor média: {melhor_media:.2f}",
            f"Destaque: {melhor_mot}",
            [m for m, _d in ranking_motorista[:8]],
            [(_d["mort_total"] / max(_d["rotas"], 1)) for _m, _d in ranking_motorista[:8]],
            color="#7C3AED",
        )
        self.set_status(f"STATUS: Relatório de mortalidade gerado ({len(rows)} rota(s)).")

    def exportar_excel(self):
        prog = upper(self.cb_prog.get())
        if not (require_pandas() and require_openpyxl()):
            return
        if not prog:
            messagebox.showwarning("ATENÇÃO", "Selecione uma programação.")
            return

        path = filedialog.asksaveasfilename(
            title="Exportar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"RELATORIO_{prog}.xlsx"
        )
        if not path:
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            api_enabled = bool(desktop_secret and is_desktop_api_sync_enabled())
            bundle = self._api_bundle_relatorio(prog)
            api_state = str(getattr(self, "_last_bundle_api_state", ""))
            if api_enabled and api_state == "not_found":
                messagebox.showwarning("ATENCAO", f"Programacao nao encontrada no servidor: {prog}")
                return
            if bundle:
                itens = [dict(r) for r in (bundle.get("clientes") or []) if isinstance(r, dict)]
                rec = [dict(r) for r in (bundle.get("recebimentos") or []) if isinstance(r, dict)]
                desp = [dict(r) for r in (bundle.get("despesas") or []) if isinstance(r, dict)]
                df_itens = pd.DataFrame(itens)
                df_rec = pd.DataFrame(rec)
                df_desp = pd.DataFrame(desp)
            else:
                if api_enabled and api_state not in ("failed", "disabled"):
                    messagebox.showwarning(
                        "ATENCAO",
                        "Falha ao processar dados da programacao no servidor. Tente novamente.",
                    )
                    return
                with get_db() as conn:
                    cur = conn.cursor()

                    cur.execute("SELECT * FROM programacao_itens WHERE codigo_programacao=?", (prog,))
                    itens = cur.fetchall()
                    cols_itens = [d[0] for d in cur.description]

                    cur.execute("SELECT * FROM recebimentos WHERE codigo_programacao=?", (prog,))
                    rec = cur.fetchall()
                    cols_rec = [d[0] for d in cur.description]

                    cur.execute("SELECT * FROM despesas WHERE codigo_programacao=?", (prog,))
                    desp = cur.fetchall()
                    cols_desp = [d[0] for d in cur.description]

                df_itens = pd.DataFrame(itens, columns=cols_itens)
                df_rec = pd.DataFrame(rec, columns=cols_rec)
                df_desp = pd.DataFrame(desp, columns=cols_desp)

            with pd.ExcelWriter(path, engine="openpyxl") as writer:

                if "data_registro" in df_rec.columns:
                    df_rec["data_registro"] = normalize_datetime_column(df_rec["data_registro"])
                if "data_registro" in df_desp.columns:
                    df_desp["data_registro"] = normalize_datetime_column(df_desp["data_registro"])

                # normaliza datas no ITENS, se existirem
                if "data_venda" in df_itens.columns:
                    df_itens["data_venda"] = normalize_date_column(df_itens["data_venda"])
                if "data_saida" in df_itens.columns:
                    df_itens["data_saida"] = normalize_date_column(df_itens["data_saida"])
                if "data_chegada" in df_itens.columns:
                    df_itens["data_chegada"] = normalize_date_column(df_itens["data_chegada"])
                if "saida_dt" in df_itens.columns:
                    df_itens["saida_dt"] = normalize_datetime_column(df_itens["saida_dt"])
                if "chegada_dt" in df_itens.columns:
                    df_itens["chegada_dt"] = normalize_datetime_column(df_itens["chegada_dt"])

                df_itens.to_excel(writer, index=False, sheet_name="ITENS")
                df_rec.to_excel(writer, index=False, sheet_name="RECEBIMENTOS")
                df_desp.to_excel(writer, index=False, sheet_name="DESPESAS")

            messagebox.showinfo("OK", "Excel exportado com sucesso!")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao exportar Excel: {str(e)}")

    def gerar_pdf(self):
        if not require_reportlab():
            return
        tipo_rel = self.cb_tipo_rel.get().strip() if hasattr(self, "cb_tipo_rel") else "Programacoes"
        is_mortalidade = "MORTALIDADE" in upper(tipo_rel)
        is_prestacao = "PRESTACAO" in upper(tipo_rel)
        prog = upper(self.cb_prog.get())

        # Prestacao de contas: reutiliza exatamente o mesmo gerador de PDF da tela de Despesas.
        if is_prestacao:
            if not prog:
                messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
                return
            try:
                despesas_page = None
                if hasattr(self, "app") and hasattr(self.app, "pages"):
                    despesas_page = self.app.pages.get("Despesas")
                if not despesas_page or not hasattr(despesas_page, "imprimir_resumo"):
                    messagebox.showerror("ERRO", "Tela de Despesas indisponÃÂvel para gerar PDF da prestação.")
                    return

                prev_prog = getattr(despesas_page, "_current_programacao", "")
                despesas_page._current_programacao = prog
                despesas_page.imprimir_resumo()
                despesas_page._current_programacao = prev_prog
            except Exception as e:
                messagebox.showerror("ERRO", f"Erro ao gerar PDF da prestação: {str(e)}")
            return

        if (not is_mortalidade) and (not prog):
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return

        nome_base = f"RELATORIO_{prog}" if prog else "RELATORIO_MORTALIDADE_MOTORISTA"

        path = filedialog.asksaveasfilename(
            title="Salvar PDF do Relatorio",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"{nome_base}.pdf"
        )
        if not path:
            return

        self.gerar_resumo()
        txt = self.txt.get("1.0", "end").strip().splitlines()

        try:
            c = canvas.Canvas(path, pagesize=A4)
            w, h = A4
            y = h - 60

            c.setFont("Helvetica-Bold", 14)
            titulo_pdf = f"RELATORIO - {prog}" if prog else "RELATORIO - MORTALIDADE POR MOTORISTA"
            c.drawString(40, y, titulo_pdf)
            y -= 24

            c.setFont("Helvetica", 10)
            for line in txt:
                c.drawString(40, y, line[:110])
                y -= 14
                if y < 60:
                    c.showPage()
                    y = h - 60
                    c.setFont("Helvetica", 10)

            c.save()
            messagebox.showinfo("OK", "PDF gerado com sucesso!")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao gerar PDF: {str(e)}")

    def finalizar_rota(self):
        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return
        if not ensure_system_api_binding(context=f"Finalizar rota ({prog})", parent=self):
            return

        info = self._get_prog_status_info(prog)
        st = info.get("status", "")
        prest = info.get("prestacao_status", "")

        if prest == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"A rota {prog} esta com a prestacao FECHADA.\n\n"
                "Ela ja esta travada para alteracoes."
            )
            return

        if st == "FINALIZADA":
            messagebox.showinfo("Info", f"A rota {prog} ja esta FINALIZADA.")
            return

        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja FINALIZAR a rota {prog}?\n\n"
            "Ela deixara de aparecer nas Rotas Ativas."
        ):
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "PUT",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/status",
                payload={
                    "status": "FINALIZADA",
                    "finalizada_no_app": 0,
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            messagebox.showinfo("OK", f"Rota FINALIZADA: {prog}")

            self.refresh_comboboxes()
            if hasattr(self.app, "refresh_programacao_comboboxes"):
                self.app.refresh_programacao_comboboxes()

            self.set_status(f"STATUS: Rota finalizada: {prog}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao finalizar rota: {str(e)}")

    def reabrir_rota(self):
        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENCAO", "Selecione uma programacao.")
            return
        if not ensure_system_api_binding(context=f"Reabrir rota ({prog})", parent=self):
            return

        info = self._get_prog_status_info(prog)
        st = info.get("status", "")
        prest = info.get("prestacao_status", "")

        if prest == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"Nao e permitido REABRIR a rota {prog} pois a prestacao esta FECHADA.\n\n"
                "Se precisar reabrir, primeiro reabra a prestacao."
            )
            return

        if st == "ATIVA":
            messagebox.showinfo("Info", f"A rota {prog} ja esta ATIVA.")
            return

        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja REABRIR a rota {prog}?\n\n"
            "Ela voltara a aparecer nas Rotas Ativas."
        ):
            return

        try:
            desktop_secret = os.environ.get("ROTA_SECRET", "").strip()
            _call_api(
                "PUT",
                f"desktop/rotas/{urllib.parse.quote(upper(prog))}/status",
                payload={
                    "status": "ATIVA",
                    "status_operacional": "",
                    "finalizada_no_app": 0,
                },
                extra_headers={"X-Desktop-Secret": desktop_secret},
            )

            messagebox.showinfo("OK", f"Rota REABERTA: {prog}")

            self.refresh_comboboxes()
            if hasattr(self.app, "refresh_programacao_comboboxes"):
                self.app.refresh_programacao_comboboxes()

            self.set_status(f"STATUS: Rota reaberta: {prog}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao reabrir rota: {str(e)}")
    

# ==========================
# ===== FIM DA PARTE 8 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 9 (ATUALIZADA) =====
# ==========================

class BackupExportarPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Backup / Exportar")

        self.app = app

        card = ttk.Frame(self.body, style="Card.TFrame", padding=18)
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        ttk.Label(card, text="Ferramentas", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))

        ttk.Button(card, text="ðÅ¸â€”â€ž FAZER BACKUP DO BANCO", style="Primary.TButton", command=self.backup_db)\
            .grid(row=1, column=0, sticky="ew", pady=6)
        ttk.Button(card, text="ââ„¢» RESTAURAR BANCO (IMPORTAR .DB)", style="Warn.TButton", command=self.restore_db)\
            .grid(row=2, column=0, sticky="ew", pady=6)
        ttk.Button(card, text="ðÅ¸â€œ¤ EXPORTAR VENDAS IMPORTADAS (EXCEL)", style="Ghost.TButton", command=self.exportar_vendas)\
            .grid(row=3, column=0, sticky="ew", pady=6)

        self.lbl = ttk.Label(card, text="Dica: Faça backup diariamente.", background="white", foreground="#444")
        self.lbl.grid(row=4, column=0, sticky="w", pady=(12, 0))

    def on_show(self):
        self.set_status("STATUS: Backup e exportações.")

    # ----------------------------
    # Helpers de segurança
    # ----------------------------
    def _is_sqlite_db_file(self, path: str) -> bool:
        """Valida assinatura do arquivo SQLite (primeiros 16 bytes)."""
        try:
            with open(path, "rb") as f:
                header = f.read(16)
            return header == b"SQLite format 3\x00"
        except Exception:
            return False

    def _make_safe_copy(self, src: str, dst: str):
        """Cópia binária simples (mantém compatibilidade)."""
        with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
            shutil.copyfileobj(fsrc, fdst)

    def _cleanup_sqlite_wal(self, db_path: str):
        for suffix in ("-wal", "-shm"):
            p = f"{db_path}{suffix}"
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    logging.debug("Falha ignorada")

    def _backup_current_db_automatic(self) -> str:
        """Cria um backup automático no mesmo diretório do DB (antes do restore)."""
        if not os.path.exists(DB_PATH):
            return ""
        try:
            base_dir = os.path.dirname(DB_PATH) or "."
            auto_name = f"auto_backup_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            auto_path = os.path.join(base_dir, auto_name)
            self._make_safe_copy(DB_PATH, auto_path)
            return auto_path
        except Exception:
            return ""

    def backup_db(self):
        if not os.path.exists(DB_PATH):
            messagebox.showerror("ERRO", "Banco não encontrado.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar backup do banco",
            defaultextension=".db",
            filetypes=[("DB", "*.db")],
            initialfile=f"backup_rota_granja_{datetime.now().strftime('%Y%m%d_%H%M')}.db"
        )
        if not path:
            return

        try:
            # âÅâ€œââ‚¬¦ melhor prática: usar SQLite backup API quando possÃÂÂvel
            try:
                import sqlite3
                src = sqlite3.connect(DB_PATH)
                dst = sqlite3.connect(path)
                with dst:
                    src.backup(dst)
                dst.close()
                src.close()
            except Exception:
                # fallback: cópia binária (mantém funcionamento se der algo no backup API)
                self._make_safe_copy(DB_PATH, path)

            messagebox.showinfo("OK", "Backup criado com sucesso!")
            self.set_status(f"STATUS: Backup criado: {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao criar backup: {str(e)}")

    def restore_db(self):
        path = filedialog.askopenfilename(
            title="Selecionar banco (.db)",
            filetypes=[("DB", "*.db")]
        )
        if not path:
            return

        # âÅâ€œââ‚¬¦ valida se é sqlite de verdade
        if not self._is_sqlite_db_file(path):
            messagebox.showerror(
                "ERRO",
                "O arquivo selecionado não parece ser um banco SQLite válido.\n\n"
                "Selecione um .db gerado pelo sistema (backup)."
            )
            return

        # confirmação mais clara (alto risco)
        if not messagebox.askyesno(
            "CONFIRMAR RESTAURAÇÃO",
            "Isso vai SUBSTITUIR seu banco atual.\n\n"
            "Recomendado: fechar telas e não estar com operações em andamento.\n\n"
            "Deseja continuar?"
        ):
            return

        try:
            # âÅâ€œââ‚¬¦ cria backup automático do DB atual
            auto_backup = self._backup_current_db_automatic()

            # âÅâââ€š¬Åâ€œââââ‚¬Å¡¬¦ tenta fechar conexÃÆââ‚¬â„¢µes âââââ€š¬Å¡¬Åâââ€š¬Åâ€œconhecidasâââââ€š¬Å¡¬Â (melhor esforÃÆââ‚¬â„¢§o, sem quebrar)
            try:
                # se existir algum método no app para reabrir/fechar conexões, chamamos
                if hasattr(self.app, "close_db_connections"):
                    self.app.close_db_connections()
            except Exception:
                logging.debug("Falha ignorada")

            # limpa WAL/SHM antigos (melhor esforço)
            try:
                self._cleanup_sqlite_wal(DB_PATH)
            except Exception:
                logging.debug("Falha ignorada")

            # âÅâ€œââ‚¬¦ restaura por cópia binária (simples e compatÃÂÂvel)
            # Observação: se houver conexão aberta, pode falhar no Windows.
            self._make_safe_copy(path, DB_PATH)

            msg = "Banco restaurado! Reinicie o sistema."
            if auto_backup:
                msg += f"\n\nBackup automático do banco anterior:\n{os.path.basename(auto_backup)}"

            messagebox.showinfo("OK", msg)
            self.set_status("STATUS: Banco restaurado. Reinicie o sistema.")

        except PermissionError:
            messagebox.showerror(
                "ERRO",
                "Não foi possÃÂÂvel substituir o banco (arquivo em uso).\n\n"
                "Feche o sistema e tente novamente, ou reinicie o computador."
            )
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao restaurar banco: {str(e)}")

    def exportar_vendas(self):
        if not (require_pandas() and require_openpyxl()):
            return
        path = filedialog.asksaveasfilename(
            title="Exportar vendas importadas",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"VENDAS_IMPORTADAS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        if not path:
            return

        try:
            with get_db() as conn:
                df = pd.read_sql_query("SELECT * FROM vendas_importadas ORDER BY id DESC", conn)

            if df.empty:
                messagebox.showwarning("ATENÇÃO", "Não há vendas importadas para exportar.")
                return

            try:
                if "data_venda" in df.columns:
                    df["data_venda"] = normalize_date_column(df["data_venda"])
                df.to_excel(path, index=False)
            except Exception as e:
                messagebox.showerror(
                    "ERRO",
                    "Falha ao exportar para Excel.\n\n"
                    "Verifique se o pacote 'openpyxl' está instalado.\n\n"
                    f"Detalhes: {str(e)}"
                )
                return

            messagebox.showinfo("OK", "Exportação feita com sucesso!")
            self.set_status(f"STATUS: Vendas exportadas: {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao exportar vendas: {str(e)}")

# ==========================
# ===== FIM DA PARTE 9 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 10 (ATUALIZADA) =====
# ==========================

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
        w, h = 440, 280
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
        self.btn_entrar.grid(row=0, column=0, sticky="ew", padx=(0, 6), ipady=6)

        self.btn_sair = ttk.Button(btns, text="âÂ» SAIR", style="Danger.TButton", command=self._request_close)
        self.btn_sair.grid(row=0, column=1, sticky="ew", padx=(6, 0), ipady=6)

        self.ent_codigo.focus_set()
        self.bind("<Return>", lambda e: self.try_login())
        self.bind("<Escape>", lambda e: self._request_close())
        self.protocol("WM_DELETE_WINDOW", self._request_close)

        self.user = None  # será preenchido quando logar

    def _request_close(self):
        try:
            ok = messagebox.askyesno("Confirmar saÃÂda", "Deseja realmente fechar o sistema?")
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
            user = autenticar_usuario(codigo, senha)
        except Exception as e:
            messagebox.showerror("ERRO", f"Falha ao autenticar: {str(e)}")
            return

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


