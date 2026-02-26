# -*- coding: utf-8 -*-
# ==========================
# ===== INCIO DA PARTE 1 =====
# ==========================
import json
import logging
import os
import re
import sqlite3
import tempfile
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from tkinter import simpledialog, ttk, messagebox, filedialog
import shutil
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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

# =========================================================
# CONSTANTES E CONFIGURAÃ‡Ã•ES
# =========================================================
APP_W, APP_H = 1360, 780
DB_PATH = os.path.join(os.path.dirname(__file__), "rota_granja.db")
DB_NAME = "rota_granja"
APP_TITLE_DESKTOP = "ROTAHUB DESKTOP"

API_BASE_URL = os.environ.get("ROTA_SERVER_URL", "http://127.0.0.1:8000").strip().rstrip("/")
if not API_BASE_URL:
    API_BASE_URL = "http://127.0.0.1:8000"
try:
    API_SYNC_TIMEOUT = float(os.environ.get("ROTA_SYNC_TIMEOUT", "15"))
except Exception:
    API_SYNC_TIMEOUT = 15.0

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

    base = os.path.dirname(__file__)
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
    """Aplica pragmas de desempenho/concorrÃªncia."""
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.execute("PRAGMA busy_timeout = 5000")
    except Exception:
        logging.debug("Falha ignorada")
    return conn

@contextmanager
def get_db():
    """Gerenciador de contexto para conexÃµes com o banco"""
    conn = sqlite3.connect(DB_PATH)
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
    """FunÃ§Ã£o compatÃ­vel para cÃ³digo existente"""
    conn = sqlite3.connect(DB_PATH)
    _configure_sqlite(conn)
    return conn

# =========================================================
# FUNÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¢ES UTILITÃƒÆ’Ã†â€™Ãƒâ€šÃ‚ÂRIAS
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
            clear_values = {"", "0", "0,0", "0,00", "0.0", "0.00", "R$ 0,00", "R$0,00", "__/__/____", "__:__"}
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
                digits = re.sub(r"\D", "", raw)[:4]
                masked = digits if len(digits) <= 2 else f"{digits[:2]}:{digits[2:]}"
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
                    try:
                        y, m, d = nd.split("-")
                        entry.delete(0, "end")
                        entry.insert(0, f"{d}/{m}/{y}")
                    except Exception:
                        entry.delete(0, "end")
                        entry.insert(0, nd)
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
    """Converte valor monetÃ¡rio para float com 2 casas (Decimal)."""
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
    """Normaliza para YYYY-MM-DD. Retorna '' se vazio, None se invÃ¡lido."""
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
            y, m, d = int(digits[:4]), int(digits[4:6]), int(digits[6:8])
            if 1 <= m <= 12 and 1 <= d <= 31:
                return f"{y:04d}-{m:02d}-{d:02d}"
        except Exception:
            return None
    return None

def normalize_time(s: str):
    """Normaliza para HH:MM. Retorna '' se vazio, None se invÃ¡lido."""
    s = (s or "").strip()
    if not s:
        return ""
    part = s.split()[0]
    digits = re.sub(r"\D", "", part)
    if not digits:
        return None
    if len(digits) == 3:
        digits = "0" + digits
    if len(digits) != 4:
        return None
    try:
        hh = int(digits[:2])
        mm = int(digits[2:])
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return f"{hh:02d}:{mm:02d}"
    except Exception:
        return None
    return None


class SyncError(Exception):
    """Erro especÃ­fico ao tentar sincronizar com a API mobile."""


def _build_api_url(path: str) -> str:
    path = (path or "").strip().lstrip("/")
    if path:
        return f"{API_BASE_URL}/{path}"
    return API_BASE_URL


def _call_api(method: str, path: str, payload=None, token: str = None):
    url = _build_api_url(path)
    headers = {"Accept": "application/json"}
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
    if token:
        headers["Authorization"] = f"Bearer {token}"

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
        raise SyncError(f"Erro inesperado ao chamar a API de sincronizaÃ§Ã£o: {exc}")
    return None

def normalize_date_time_components(data: str, hora: str):
    """Normaliza data/hora e retorna tupla (data, hora) preservando valor original se invÃ¡lido."""
    nd = normalize_date(data)
    nt = normalize_time(hora)
    out_data = nd if nd is not None else (data or "")
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
    """Formata valor monetÃ¡rio"""
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
                "ExportaÃ§Ã£o Excel requer o pacote 'openpyxl'.\n\n"
                "Instale com: pip install openpyxl"
            )
        except Exception:
            logging.debug("Falha ignorada")
        return False

def require_xlrd() -> bool:
    try:
        import xlrd  # noqa: F401
        return True
    except Exception:
        try:
            messagebox.showerror(
                "ERRO",
                "ImportaÃ§Ã£o de .xls requer o pacote 'xlrd'.\n\n"
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
            "Formato de arquivo nÃ£o suportado.\n\n"
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
                "GeraÃ§Ã£o de PDF requer o pacote 'reportlab'.\n\n"
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
    """Gera um cÃ³digo curto para programaÃ§Ã£o (PG + YYYY + sequÃªncia hÃ¡ 2 dÃ­gitos)."""
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
        logging.debug("Falha ao calcular prÃ³ximo cÃ³digo de programaÃ§Ã£o")

    return f"{prefix}{suffix:02d}"

def generate_motorista_code(motorista_id: int = None) -> str:
    """Gera um cÃ³digo simples e Ãºnico pro motorista (ex.: MOT000123 / MOTAB12CD)."""
    base = "MOT"
    if motorista_id is not None:
        try:
            return f"{base}{int(motorista_id):06d}"
        except Exception:
            logging.debug("Falha ignorada")
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"{base}{suffix}"

def generate_usuario_code(usuario_id: int = None) -> str:
    """Gera um cÃ³digo simples e Ãºnico pro usuÃ¡rio (ex.: USR000123 / USRAB12CD)."""
    base = "USR"
    if usuario_id is not None:
        try:
            return f"{base}{int(usuario_id):06d}"
        except Exception:
            logging.debug("Falha ignorada")
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=6))
    return f"{base}{suffix}"

def generate_motorista_password(length: int = 6) -> str:
    """Senha numÃ©rica simples (pra bater com a API atual)."""
    try:
        length = int(length)
    except Exception:
        length = 6
    if length < 4:
        length = 4
    return "".join(random.choices("0123456789", k=length))

def generate_usuario_password(length: int = 6) -> str:
    """Senha numÃ©rica simples pro usuÃ¡rio."""
    return generate_motorista_password(length)

def guess_col(cols, candidates):
    """Tenta adivinhar o nome da coluna no Excel"""
    cols_lower = {str(c).strip().lower(): c for c in cols}
    for cand in candidates:
        cand_low = cand.strip().lower()
        for col_low, original in cols_lower.items():
            if cand_low in col_low:
                return original
    return None

def validate_required(value: str, field_name: str, min_len=1, max_len=120):
    v = (value or "").strip()
    if len(v) < min_len:
        return False, f"{field_name} Ã© obrigatÃ³rio."
    if len(v) > max_len:
        return False, f"{field_name} muito longo (mÃ¡x {max_len})."
    return True, ""


def validate_codigo(value: str, field_name="CÃ³digo", min_len=2, max_len=20):
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
        return False, "Placa invÃ¡lida."
    return True, ""


def validate_money(value: str, field_name="Valor"):
    v = (value or "").strip().replace(".", "").replace(",", ".")
    try:
        f = float(v)
        if f < 0:
            return False, f"{field_name} nÃ£o pode ser negativo."
        return True, ""
    except Exception:
        return False, f"{field_name} invÃ¡lido."


# =========================================================
# AUTENTICAÃ‡ÃƒO + ADMIN PADRÃƒO
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
    âœ… Login por NOME + SENHA
    âœ… MigraÃ§Ã£o automÃ¡tica:
       - Se senha no DB jÃ¡ for HASH: valida HASH
       - Se senha no DB for PURA: valida pura e, se OK, converte e salva HASH
    Retorna dict do usuÃ¡rio se ok, senÃ£o None.
    """
    login = (login or "").strip()
    senha = (senha or "").strip()
    if not login or not senha:
        return None

    with get_db() as conn:
        cur = conn.cursor()

        # Descobre colunas disponÃ­veis
        cur.execute("PRAGMA table_info(usuarios)")
        cols = [str(r[1] or "").lower() for r in cur.fetchall()]

        has_permissoes = "permissoes" in cols
        has_cpf = "cpf" in cols
        has_telefone = "telefone" in cols
        has_senha = "senha" in cols

        if not has_senha:
            return None

        # Puxa o usuÃ¡rio pelo nome (case-insensitive)
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

        # 1) Se jÃ¡ Ã© hash, valida por hash
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
                # conn.commit() Ã© automÃ¡tico no seu contexto get_db() se ele commita ao sair;
                # se nÃ£o, descomente a linha abaixo:
                # conn.commit()
            except Exception as e:
                # Se falhar a migraÃ§Ã£o, pelo menos deixa logar (jÃ¡ validou a senha pura)
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
    """Notifica senha temporÃ¡ria do ADMIN sem gravar em arquivo."""
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
# MIGRAÃ‡ÃƒO DE BANCO DE DADOS
# =========================================================
def table_has_column(cur, table, col):
    """Verifica se coluna existe na tabela"""
    cur.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    return col in cols

def safe_add_column(cur, table, col, coltype):
    """Adiciona coluna apenas se nÃ£o existir"""
    if not table_has_column(cur, table, col):
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")

def db_init():
    """Inicializa/atualiza banco de dados com migraÃ§Ãµes"""
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
        # âœ… ADICIONADO (pra login + gerar senha/codigo)
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
        # MigraÃ§Ã£o: se banco antigo usava capacidade_c, copiar para capacidade_cx
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

        # Migra base legada de equipes -> ajudantes (melhor esforÃ§o)
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

        # PROGRAMAÃ‡Ã•ES
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

        try:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_programacoes_codigo ON programacoes(codigo_programacao)")
        except Exception as e:
            logging.exception("Falha ao criar indice de programacoes (codigo_programacao): %s", e)

        # NOVAS COLUNAS PARA STATUS E ROTAS ATIVAS
        safe_add_column(cur, "programacoes", "status", "TEXT DEFAULT 'ATIVA'")
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

        # âœ… SUA TELA DESPESAS USA "adiantamento_rota" (ADICIONADO)
        safe_add_column(cur, "programacoes", "adiantamento_rota", "REAL DEFAULT 0")

        # NOVAS MIGRAÃ‡Ã•ES (NF / KM / CÃ‰DULAS)
        safe_add_column(cur, "programacoes", "nf_numero", "TEXT")
        safe_add_column(cur, "programacoes", "nf_kg", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_preco", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_caixas", "INTEGER DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_kg_carregado", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_kg_vendido", "REAL DEFAULT 0")
        safe_add_column(cur, "programacoes", "nf_saldo", "REAL DEFAULT 0")

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

        # garante prestaÃ§Ã£o pendente em bases antigas (nÃ£o sobrescreve FECHADA)
        try:
            if table_has_column(cur, "programacoes", "prestacao_status"):
                cur.execute("""
                    UPDATE programacoes
                    SET prestacao_status='PENDENTE'
                    WHERE (prestacao_status IS NULL OR prestacao_status='')
                """)
        except Exception as e:
            logging.exception("Falha ao garantir prestacao_status pendente: %s", e)

        # ITENS DA PROGRAMAÃ‡ÃƒO
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

        # CONTROLE/LOG (sincronizaÃ§Ã£o app)
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

        # MIGRAÃ‡Ã•ES PARA RELATÃ“RIOS
        safe_add_column(cur, "despesas", "tipo_despesa", "TEXT DEFAULT 'ROTA'")
        safe_add_column(cur, "despesas", "categoria", "TEXT")
        safe_add_column(cur, "despesas", "motorista", "TEXT")
        safe_add_column(cur, "despesas", "veiculo", "TEXT")
        safe_add_column(cur, "despesas", "observacao", "TEXT")  # âœ… sua tela usa isso

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

    # âœ… garante ADMIN (senha segura e temporaria)
    ensure_admin_user()

# =========================================================
# ESTILOS DA INTERFACE (MODERNIZADO)
# =========================================================
def apply_style(root):
    """Aplica estilos Tkinter (tema moderno, sem afetar a lÃ³gica do sistema)"""
    style = ttk.Style(root)

    # Em alguns Windows, "clam" funciona melhor e Ã© mais consistente
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

    # Tabelas (Treeview): linhas/colunas mais visÃ­veis
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

    # BotÃµes (mantendo seus nomes de style)
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
# BASE PARA PÃƒÆ’Ã‚GINAS (mesma lÃƒÆ’Ã‚Â³gica, sÃƒÆ’Ã‚Â³ ajuste visual/organizaÃƒÆ’Ã‚Â§ÃƒÆ’Ã‚Â£o)
# =========================================================
class PageBase(ttk.Frame):
    """Classe base para todas as pÃ¡ginas da aplicaÃ§Ã£o"""
    def __init__(self, parent, app, title):
        super().__init__(parent, style="Content.TFrame")
        self.app = app

        # Estrutura
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(self, style="Content.TFrame", padding=(18, 16))
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ttk.Label(
            header,
            text=title,
            font=("Segoe UI", 18, "bold"),
            background="#F4F6FB",
            foreground="#111827"
        ).grid(row=0, column=0, sticky="w")

        self.lbl_status = ttk.Label(
            header,
            text="STATUS: ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬",
            background="#F4F6FB",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_status.grid(row=1, column=0, sticky="w", pady=(6, 0))

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
        """Atualiza texto de status na pÃ¡gina"""
        self.lbl_status.config(text=txt)

    def on_show(self):
        """MÃ©todo chamado quando a pÃ¡gina Ã© exibida (pode ser sobrescrito)"""
        pass


# =========================================================
# COMPONENTE CRUD GENÃ‰RICO (mesma lÃ³gica, layout mais robusto)
# =========================================================
# ==========================
# ===== CADASTRO CRUD (ATUALIZADO) =====
# ==========================

class CadastroCRUD(ttk.Frame):
    """Componente CRUD reutilizÃ¡vel para cadastros (com validaÃ§Ãµes por tabela)"""
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

            if self.table == "ajudantes" and col == "status":
                ent = ttk.Combobox(form, state="readonly", values=["ATIVO", "DESATIVADO"])
                ent.set("ATIVO")
            else:
                ent = ttk.Entry(form, style="Field.TEntry")
            ent.grid(row=r + 1, column=c, sticky="ew", padx=6, pady=(0, 10))
            self.entries[col] = ent
            if isinstance(ent, ttk.Entry):
                kind, precision = self._infer_entry_kind(col, label)
                bind_entry_smart(ent, kind, precision=precision)

        # BOTÃ•ES
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(4, 12))
        actions.grid_columnconfigure(20, weight=1)

        ttk.Button(actions, text="💾 SALVAR", style="Primary.TButton", command=self.salvar).grid(row=0, column=0, padx=6)
        ttk.Button(actions, text="✏️ EDITAR", style="Ghost.TButton", command=self.editar).grid(row=0, column=1, padx=6)
        ttk.Button(actions, text="🗑️ EXCLUIR", style="Danger.TButton", command=self.excluir).grid(row=0, column=2, padx=6)
        ttk.Button(actions, text="🧹 LIMPAR", style="Ghost.TButton", command=self.limpar).grid(row=0, column=3, padx=6)
        ttk.Button(actions, text="🔄 ATUALIZAR", style="Ghost.TButton", command=self.carregar).grid(row=0, column=4, padx=6)

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

        # Importar clientes (se vocÃª ainda usa esse CRUD para clientes em algum lugar)
        if self.table == "clientes":
            ttk.Button(
                actions,
                text="📥 IMPORTAR CLIENTES (EXCEL)",
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
        if any(k in l for k in ("VALOR", "PREÃ‡O", "CUSTO", "ADIANTAMENTO", "DIÃRIA")):
            return "money", 2
        if "CPF" in l:
            return "cpf", 2
        if "TELEFONE" in l:
            return "phone", 2
        if any(k in l for k in ("CAPACIDADE", "CAIXA", "QTD")):
            return "int", 2
        if any(k in l for k in ("KG", "KM", "MÃ‰DIA", "MEDIA", "LITROS")):
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
            # Se por algum motivo a coluna nÃ£o existir no DB, nÃ£o trava o sistema.
            return False

    def _dup_exists_exact(self, cur, colname, value, ignore_id=None):
        """
        Verifica duplicidade exata (Ãºtil para CPF/telefone quando vocÃª normaliza para dÃ­gitos).
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

    # -------------------------
    # AÃ§Ãµes padrÃ£o
    # -------------------------
    def limpar(self):
        self.selected_id = None
        for col, _ in self.fields:
            self._set(col, "")
        if self.table == "ajudantes" and "status" in self.entries:
            self._set("status", "ATIVO")
        self._update_password_controls()

    def carregar(self):
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
        Salva/atualiza com validaÃ§Ãµes por cadastro:
        - motoristas: valida NOME/CPF/TELEFONE; exige codigo e senha (manual); codigo Ãºnico; CPF Ãºnico (se houver coluna)
        - usuarios: exige nome, senha; nome Ãºnico; sem codigo
        - veiculos: exige placa, modelo, capacidade_cx; placa Ãºnica; capacidade int >= 0

        Obs: ID Ã© a chave primÃ¡ria autogerada (sequÃªncia do SQLite). VocÃª jÃ¡ vÃª o ID na tabela.
        """
        data = {col: self.entries[col].get().strip() for col, _ in self.fields}

        # NormalizaÃ§Ãµes por tabela
        if self.table in {"motoristas", "usuarios", "veiculos", "equipes", "ajudantes", "clientes"}:
            for k in list(data.keys()):
                data[k] = self._norm(data[k])

        # Validar obrigatÃ³rios + duplicidade + tipos
        try:
            with get_db() as conn:
                cur = conn.cursor()

                # -------------------------
                # MOTORISTAS
                # -------------------------
                if self.table == "motoristas":
                    # Exige NOME, CÃ“DIGO e SENHA (manuais) + TELEFONE
                    # CPF Ã© opcional, porÃ©m deve ser vÃ¡lido e Ãºnico quando informado.
                    ok, f = self._require_fields(["nome", "codigo", "senha", "telefone"])
                    if not ok:
                        messagebox.showwarning("ATENÃ‡ÃƒO", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    # Nome (bÃ¡sico)
                    nome = self._norm(data.get("nome"))
                    if len(nome) < 3:
                        messagebox.showwarning("ATENÃ‡ÃƒO", "NOME deve ter pelo menos 3 caracteres.")
                        return
                    data["nome"] = nome

                    # CÃ³digo (login flutter)  manual, mas pode usar GERAR
                    cod = self._norm(data.get("codigo"))
                    if not is_valid_motorista_codigo(cod):
                        messagebox.showwarning(
                            "ATENÃ‡ÃƒO",
                            "CÃ“DIGO invÃ¡lido. Use apenas letras/nÃºmeros/._- e 3 a 24 caracteres."
                        )
                        return
                    if self._dup_exists(cur, "codigo", cod, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE MOTORISTA COM ESTE CÃ“DIGO: {cod}")
                        return
                    data["codigo"] = cod

                    # Senha (login flutter)  manual
                    senha = (self.entries.get("senha").get().strip() if self.entries.get("senha") else "").strip()
                    if self.selected_id:
                        if senha:
                            if not self._is_admin:
                                messagebox.showwarning("ATENÃ‡ÃƒO", "Somente ADMIN pode alterar senha do motorista.")
                                return
                            if not is_valid_motorista_senha(senha):
                                messagebox.showwarning(
                                    "ATENÃ‡ÃƒO",
                                    "SENHA invÃ¡lida. Use 4 a 24 caracteres (nÃ£o pode ser vazia)."
                                )
                                return
                            data["senha"] = hash_password_pbkdf2(senha)
                        else:
                            data.pop("senha", None)
                    else:
                        if not is_valid_motorista_senha(senha):
                            messagebox.showwarning(
                                "ATENÃ‡ÃƒO",
                                "SENHA invÃ¡lida. Use 4 a 24 caracteres (nÃ£o pode ser vazia)."
                            )
                            return
                        data["senha"] = hash_password_pbkdf2(senha)

                    # CPF (opcional)
                    cpf_raw = self.entries.get("cpf").get().strip() if self.entries.get("cpf") else ""
                    cpf = normalize_cpf(cpf_raw)
                    if cpf and not is_valid_cpf(cpf):
                        messagebox.showwarning("ATENÃ‡ÃƒO", "CPF invÃ¡lido.")
                        return
                    data["cpf"] = cpf  # salva sÃ³ dÃ­gitos

                    # Telefone
                    tel_raw = self.entries.get("telefone").get().strip() if self.entries.get("telefone") else ""
                    tel = normalize_phone(tel_raw)
                    if not is_valid_phone(tel):
                        messagebox.showwarning(
                            "ATENÃ‡ÃƒO",
                            "TELEFONE invÃ¡lido. Informe DDD+NÃºmero (10 ou 11 dÃ­gitos)."
                        )
                        return
                    data["telefone"] = tel  # salva sÃ³ dÃ­gitos

                    # CPF Ãºnico quando informado (se a coluna existir)
                    if cpf and db_has_column(cur, "motoristas", "cpf"):
                        if self._dup_exists_exact(cur, "cpf", cpf, ignore_id=self.selected_id):
                            messagebox.showerror("ERRO", "J EXISTE MOTORISTA COM ESTE CPF.")
                            return

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
                        messagebox.showwarning("ANTENÃ‡ÃƒO", "PREENCHA O CAMPO: NOME.")
                        return
                    if self._dup_exists(cur, "nome", nome, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE USURIO COM ESTE NOME:{nome}")
                        return

                    data["nome"] = nome
                    if self.selected_id:
                        if senha_plana:
                            if not self._is_admin:
                                messagebox.showwarning("ATENÃ‡ÃƒO", "Somente ADMIN pode alterar senha aqui.")
                                return
                            if len(senha_plana) < 6:
                                messagebox.showwarning("ANTENÃ‡ÃƒO", "SENHA deve ter pelo menos 6 caracteres.")
                                return
                            data["senha"] = hash_password_pbkdf2(senha_plana)
                        else:
                            data.pop("senha", None)
                    else:
                        if not senha_plana:
                            messagebox.showwarning("ANTENÃ‡ÃƒO", "PREENCHA O CAMPO: SENHA.")
                            return
                        if len(senha_plana) < 6:
                            messagebox.showwarning("ANTENÃ‡ÃƒO", "SENHA deve ter pelo menos 6 caracteres.")
                            return
                        data["senha"] = hash_password_pbkdf2(senha_plana)

                # -------------------------
                # AJUDANTES (nome, sobrenome, telefone)
                # -------------------------
                if self.table == "ajudantes":
                    ok, f = self._require_fields(["nome", "sobrenome", "telefone", "status"])
                    if not ok:
                        messagebox.showwarning("ATENÃƒâ€¡ÃƒÆ’O", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    nome = self._norm(data.get("nome"))
                    sobrenome = self._norm(data.get("sobrenome"))
                    telefone = normalize_phone(data.get("telefone"))
                    status = self._norm(data.get("status"))

                    if len(nome) < 2:
                        messagebox.showwarning("ATENÃƒâ€¡ÃƒÆ’O", "NOME deve ter pelo menos 2 caracteres.")
                        return
                    if len(sobrenome) < 2:
                        messagebox.showwarning("ATENÃƒâ€¡ÃƒÆ’O", "SOBRENOME deve ter pelo menos 2 caracteres.")
                        return
                    if not is_valid_phone(telefone):
                        messagebox.showwarning(
                            "ATENÃƒâ€¡ÃƒÆ’O",
                            "TELEFONE invÃƒÂ¡lido. Informe DDD+NÃƒÂºmero (10 ou 11 dÃƒÂ­gitos)."
                        )
                        return
                    if status not in {"ATIVO", "DESATIVADO"}:
                        messagebox.showwarning("ATENÃ‡ÃƒO", "STATUS invÃ¡lido para ajudante.")
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
                            messagebox.showerror("ERRO", "JÃƒ EXISTE AJUDANTE COM ESTE NOME/SOBRENOME.")
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
                        messagebox.showwarning("ATENÃ‡ÃƒO", f"PREENCHA O CAMPO: {f.upper()}.")
                        return

                    placa = self._norm(data.get("placa"))
                    if self._dup_exists(cur, "placa", placa, ignore_id=self.selected_id):
                        messagebox.showerror("ERRO", f"J EXISTE VECULO COM ESTA PLACA: {placa}")
                        return

                    cap = safe_int(data.get("capacidade_cx"), -1)
                    if cap < 0:
                        messagebox.showwarning("ATENÃ‡ÃƒO", "CAPACIDADE (CX) deve ser um nÃºmero inteiro >= 0.")
                        return
                    data["capacidade_cx"] = cap

                # -------------------------
                # SALVAR (UPDATE/INSERT)
                # -------------------------
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

        self.carregar()
        self.limpar()

        if self.app:
            self.app.refresh_programacao_comboboxes()

    def excluir(self):
        if not self.selected_id:
            messagebox.showwarning("ATENÃ‡ÃƒO", "SELECIONE UM ITEM NA TABELA ANTES DE EXCLUIR.")
            return

        if not messagebox.askyesno("CONFIRMAR", "DESEJA EXCLUIR ESTE REGISTRO?"):
            return

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute(f"DELETE FROM {self.table} WHERE id=?", (self.selected_id,))

        self.carregar()
        self.limpar()

        if self.app:
            self.app.refresh_programacao_comboboxes()

    def editar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione um item na tabela para editar.")
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

        # Mantive seu importar_clientes_excel como estava (se vocÃª ainda usa em algum ponto)
    def importar_clientes_excel(self):
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

            col_cod = guess_col(df.columns, ["cod", "cÃ³d", "codigo", "cliente", "cod cliente"])
            col_nome = guess_col(df.columns, ["nome", "cliente"])
            col_end = guess_col(df.columns, ["endereco", "endereÃ§o", "rua", "logradouro"])
            col_bairro = guess_col(df.columns, ["bairro"])
            col_cidade = guess_col(df.columns, ["cidade"])
            col_uf = guess_col(df.columns, ["uf", "estado"])
            col_tel = guess_col(df.columns, ["telefone", "fone", "celular", "contato"])
            col_rota = guess_col(df.columns, ["rota"])

            if not col_cod or not col_nome:
                messagebox.showerror("ERRO", "NÃƒO IDENTIFIQUEI AS COLUNAS DE CÃ“DIGO E NOME DO CLIENTE NO EXCEL.")
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "SELECIONE UM USURIO NA TABELA.")
            return
        if not self._user_can_change_password(self.selected_id):
            messagebox.showwarning("ATENÃ‡ÃƒO", "VocÃª nÃ£o tem permissÃ£o para alterar esta senha.")
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
                messagebox.showwarning("ATENÃ‡ÃƒO", "Informe a nova senha.")
                return
            if len(senha_nova) < 6:
                messagebox.showwarning("ATENÃ‡ÃƒO", "A nova senha deve ter pelo menos 6 caracteres.")
                return
            if senha_nova != senha_conf:
                messagebox.showwarning("ATENÃ‡ÃƒO", "ConfirmaÃ§Ã£o nÃ£o confere.")
                return

            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT senha FROM usuarios WHERE id=?", (self.selected_id,))
                row_db = cur.fetchone()
                if not row_db:
                    messagebox.showerror("ERRO", "UsuÃ¡rio nÃ£o encontrado.")
                    return
                senha_db = row_db[0] or ""

                if require_current:
                    ok = False
                    if str(senha_db).startswith("pbkdf2_sha256$"):
                        ok = verify_password_pbkdf2(senha_atual, senha_db)
                    else:
                        ok = (senha_db == senha_atual)
                    if not ok:
                        messagebox.showwarning("ATENÃ‡ÃƒO", "Senha atual invÃ¡lida.")
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
        ttk.Button(btns, text="💾 SALVAR", style="Primary.TButton", command=_salvar).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="✖ CANCELAR", style="Ghost.TButton", command=win.destroy).grid(row=0, column=1, padx=6)

    def alterar_senha_motorista(self):
        if self.table != "motoristas":
            return
        if not self.selected_id:
            messagebox.showwarning("ATENÃ‡ÃƒO", "SELECIONE UM MOTORISTA NA TABELA.")
            return
        if not self._is_admin:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Somente ADMIN pode alterar senha do motorista.")
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
                messagebox.showwarning("ATENÃ‡ÃƒO", "Informe a nova senha.")
                return
            if not is_valid_motorista_senha(senha_nova):
                messagebox.showwarning(
                    "ATENÃ‡ÃƒO",
                    "SENHA invÃ¡lida. Use 4 a 24 caracteres (nÃ£o pode ser vazia)."
                )
                return
            if senha_nova != senha_conf:
                messagebox.showwarning("ATENÃ‡ÃƒO", "ConfirmaÃ§Ã£o nÃ£o confere.")
                return

            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("UPDATE motoristas SET senha=? WHERE id=?", (self._norm(senha_nova), self.selected_id))

            messagebox.showinfo("OK", "Senha do motorista atualizada com sucesso.")
            try:
                win.destroy()
            except Exception:
                logging.debug("Falha ignorada")
            self.carregar()
            self.limpar()

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="💾 SALVAR", style="Primary.TButton", command=_salvar).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="✖ CANCELAR", style="Ghost.TButton", command=win.destroy).grid(row=0, column=1, padx=6)


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
        # - ID Ã© gerado pelo SQLite (sequÃªncia/autoincrement)
        # - valida CPF/NOME/TELEFONE
        # - cÃ³digo e senha sÃ£o MANUAIS (botÃ£o GERAR Ã© opcional)
        # -------------------------
        frm_motoristas = ttk.Frame(nb, style="Content.TFrame")
        crud_motoristas = CadastroCRUD(
            frm_motoristas,
            "Motoristas",
            "motoristas",
            [
                ("nome", "NOME"),
                ("codigo", "CÃ“DIGO"),
                ("senha", "SENHA"),
                ("cpf", "CPF"),
                ("telefone", "TELEFONE"),
            ],
            app=app
        )
        crud_motoristas.pack(fill="both", expand=True)
        nb.add(frm_motoristas, text="Motoristas")

        # -------------------------
        # USURIOS (login por nome + senha, sem idade e sem cÃ³digo)
        # -------------------------
        frm_usuarios = ttk.Frame(nb, style="Content.TFrame")
        crud_usuarios = CadastroCRUD(
            frm_usuarios,
            "UsuÃ¡rios",
            "usuarios",
            [
                ("nome", "NOME"),
                ("senha", "SENHA"),
                ("permissoes", "PERMISSÃ•ES"),
                ("cpf", "CPF"),
                ("telefone", "TELEFONE"),
            ],
            app=app
        )
        crud_usuarios.pack(fill="both", expand=True)
        nb.add(frm_usuarios, text="UsuÃ¡rios")

        # -------------------------
        # VECULOS (placa, modelo, capacidade_cx)
        # -------------------------
        frm_veiculos = ttk.Frame(nb, style="Content.TFrame")
        crud_veiculos = CadastroCRUD(
            frm_veiculos,
            "VeÃ­culos",
            "veiculos",
            [
                ("placa", "PLACA"),
                ("modelo", "MODELO"),
                ("capacidade_cx", "CAPACIDADE (CX)"),
            ],
            app=app
        )
        crud_veiculos.pack(fill="both", expand=True)
        nb.add(frm_veiculos, text="VeÃ­culos")

        # -------------------------
        # EQUIPES (mantÃ©m)
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
        # CLIENTES (mantÃ©m sua ClientesImportPage)
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

        # âœ… Menu com scroll (evita quebra/sumir item em telas pequenas)
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
        self._add_btn("home", "🏠 Home", lambda: app.show_page("Home"))
        self._add_btn("cadastros", "📚 Cadastros", lambda: app.show_page("Cadastros"))
        self._add_btn("rotas", "🗺️ Rotas", lambda: app.show_page("Rotas"))
        self._add_btn("vendas", "📥 Importar Vendas", lambda: app.show_page("ImportarVendas"))
        self._add_btn("programacao", "🗓️ Programacao", lambda: app.show_page("Programacao"))
        self._add_btn("recebimentos", "💵 Recebimentos", lambda: app.show_page("Recebimentos"))
        self._add_btn("despesas", "💸 Despesas", lambda: app.show_page("Despesas"))
        self._add_btn("escala", "📊 Escala", lambda: app.show_page("Escala"))
        self._add_btn("centro_custos", "🚛 Centro de Custos", lambda: app.show_page("CentroCustos"))
        self._add_btn("relatorios", "📄 Relatorios", lambda: app.show_page("Relatorios"))
        self._add_btn("backup", "🗄️ Backup / Exportar", lambda: app.show_page("BackupExportar"))

        # RodapÃ©
        bottom = ttk.Frame(self, style="Sidebar.TFrame", padding=(10, 12))
        bottom.pack(fill="x")

        ttk.Button(bottom, text="⏻ SAIR", style="Danger.TButton", command=self._safe_quit).pack(fill="x")

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
        """Adiciona botÃ£o ao menu lateral"""
        b = ttk.Button(self.menu, text=text, style="Side.TButton", command=cmd)
        b.pack(fill="x", pady=2)
        self.buttons[key] = b

    def set_active(self, page_key):
        """Destaca botÃ£o ativo"""
        for k, b in self.buttons.items():
            b.config(style="SideActive.TButton" if k == page_key else "Side.TButton")


class App(tk.Tk):
    def __init__(self, user=None):
        super().__init__()

        # âœ… SeguranÃ§a bÃ¡sica: sempre garante chaves esperadas
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

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.sidebar = Sidebar(self, self)
        self.sidebar.grid(row=0, column=0, sticky="ns")

        self.container = ttk.Frame(self, style="Content.TFrame")
        self.container.grid(row=0, column=1, sticky="nsew")
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.pages = {}
        self.current_page_name = None
        self._build_pages()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self.request_close)

        self.show_page("Home")

    def request_close(self):
        """Confirma fechamento da aplicação."""
        try:
            ok = messagebox.askyesno("Confirmar saída", "Deseja realmente fechar o sistema?")
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

    def _build_pages(self):
        """ConstrÃ³i todas as pÃ¡ginas da aplicaÃ§Ã£o"""
        self.pages["Home"] = HomePage(self.container, self)
        self.pages["Cadastros"] = CadastrosPage(self.container, self)
        self.pages["Rotas"] = RotasPage(self.container, self)
        self.pages["ImportarVendas"] = ImportarVendasPage(self.container, self)
        self.pages["Programacao"] = ProgramacaoPage(self.container, self)
        self.pages["Recebimentos"] = RecebimentosPage(self.container, self)
        self.pages["Despesas"] = DespesasPage(self.container, self)
        self.pages["Escala"] = EscalaPage(self.container, self)
        self.pages["CentroCustos"] = CentroCustosPage(self.container, self)
        self.pages["Relatorios"] = RelatoriosPage(self.container, self)
        self.pages["BackupExportar"] = BackupExportarPage(self.container, self)

        for p in self.pages.values():
            p.grid(row=0, column=0, sticky="nsew")

    def show_page(self, name):
        """Exibe pÃ¡gina e atualiza menu lateral"""
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

        page = self.pages.get(name)
        if not page:
            messagebox.showwarning("ATENÃ‡ÃƒO", f"PÃ¡gina '{name}' nÃ£o encontrada.")
            return

        page.tkraise()
        self.current_page_name = name
        try:
            page.on_show()
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao abrir pÃ¡gina '{name}':\n\n{e}")

    def refresh_programacao_comboboxes(self):
        """Atualiza comboboxes em pÃ¡ginas relevantes"""
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
        # ESC: fecha modal atual; se nÃ£o houver, volta para Home.
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
        # DEL: nÃ£o intercepta quando estiver digitando em campo.
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
        # ENTER global: executa aÃ§Ã£o principal da tela quando nÃ£o estÃ¡ digitando em campo.
        if self._is_text_input_focus():
            return None
        page = self._active_page()
        if self._try_call(page, ["salvar_programacao", "salvar_recebimento", "salvar_tudo", "gerar_resumo", "refresh_data", "carregar"]):
            return "break"
        return None


class HomePage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Home")

        card = ttk.Frame(self.body, style="Card.TFrame", padding=18)
        card.grid(row=0, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=0)

        ttk.Label(card, text="Bem-vindo!", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        self.lbl_clock = ttk.Label(
            card,
            text="--/--/---- --:--:--",
            background="white",
            foreground="#111827",
            font=("Segoe UI", 11, "bold")
        )
        self.lbl_clock.grid(row=0, column=1, sticky="e")

        ttk.Label(
            card,
            text=(
                "â€¢ Cadastre Motoristas, VeÃ­culos, Equipes e Clientes.\n"
                "â€¢ Importe Vendas via Excel.\n"
                "â€¢ Gere ProgramaÃ§Ãµes (cÃ³digos automÃ¡ticos) e vincule pedidos/entregas.\n"
                "â€¢ Registre Recebimentos e Despesas.\n"
                "â€¢ Emita RelatÃ³rios e PDF.\n"
            ),
            style="CardLabel.TLabel",
            justify="left"
        ).grid(row=1, column=0, sticky="w", pady=(6, 0), columnspan=2)

        self.card_stats = ttk.Frame(self.body, style="Content.TFrame")
        self.card_stats.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.card_stats.grid_columnconfigure((0, 1, 2), weight=1)

        self.lbl_total_prog = self._build_stat(self.card_stats, 0, "ProgramaÃ§Ãµes Ativas", "")
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

        cols = ["COD", "MOTORISTA", "VECULO", "DATA"]
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

    def on_show(self):
        total_prog = 0
        total_vendas = 0
        total_clientes_ativos = 0
        rows = []

        with get_db() as conn:
            cur = conn.cursor()

            try:
                cur.execute("""
                    SELECT COUNT(*)
                    FROM programacoes
                    WHERE UPPER(TRIM(COALESCE(status, ''))) NOT IN
                          ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')
                """)
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
                        WHERE UPPER(TRIM(COALESCE(status, ''))) NOT IN
                              ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')
                    )
                """)
                r = cur.fetchone()
                total_clientes_ativos = (r[0] if r else 0) or 0
            except Exception:
                total_clientes_ativos = 0

            for i in self.tree_rotas.get_children():
                self.tree_rotas.delete(i)

            try:
                cur.execute("""
                    SELECT codigo_programacao, motorista, veiculo, data_criacao
                    FROM programacoes
                    WHERE UPPER(TRIM(COALESCE(status, ''))) NOT IN
                          ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')
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
        self.set_status(f"STATUS: Logado como {nome} (ADMIN: {is_admin}).")

    # =========================================================
    # PREVIEW ROTA ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬ PUXA DADOS REAIS DA PROGRAMAÃƒÆ’Ã¢â‚¬Â¡ÃƒÆ’Ã†â€™O ATIVA
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
            Compatibilidade: alguns bancos guardam preÃ§o unitÃ¡rio em centavos
            (ex.: 630 para R$ 6,30). Ajusta apenas para exibiÃ§Ã£o no preview.
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

        def _load_preview():
            try:
                tree.delete(*tree.get_children())
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

            # âœ… helper local pra nÃ£o quebrar se nÃ£o existir ainda no arquivo
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

            # 2) ITENS (fallback se helper nÃ£o existir ainda)
            try:
                if "fetch_programacao_itens" in globals():
                    itens = fetch_programacao_itens(codigo)
                else:
                    itens = []
            except Exception:
                itens = []

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
                if st_norm in {"ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO", "CONCLUÃDO"}:
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

                tree_insert_aligned(tree, "", "end", (
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

            # 3) labels
            lbl_status.config(text=f"Status: {status or ''}")
            lbl_motorista.config(text=f"Motorista: {motorista or ''}")
            equipe_txt = resolve_equipe_nomes(equipe)
            lbl_equipe.config(text=f"Equipe: {equipe_txt or ''}")
            lbl_nf.config(text=f"NF: {nf_numero or ''}")

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
            if abs(nf_saldo) < 0.0001:
                nf_saldo = round(max(nf_kg - nf_kg_carregado, 0.0), 2)

            lbl_kg_carregado.config(text=f"KG Carregado: {nf_kg_carregado:.2f}".replace(".", ","))
            lbl_caixas_carreg.config(text=f"Caixas Carregadas: {nf_caixas}")
            lbl_caixas_entreg.config(text=f"Caixas Entregues: {total_cx_entregue}")
            lbl_saldo.config(text=f"Saldo (NF - Granja): {nf_saldo:.2f}".replace(".", ","))
            lbl_media_carreg.config(text=f"Media Carregada: {media_carregada:.3f}".replace(".", ","))
            lbl_caixa_final.config(text=f"Caixa Final: {caixa_final if caixa_final > 0 else '-'}")
            lbl_subst.config(text=subst_txt)
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
        ttk.Button(btns, text="🔄 ATUALIZAR", style="Ghost.TButton", command=_load_preview).pack(side="right")


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

        ttk.Button(self.footer_right, text="🔄 ATUALIZAR", style="Ghost.TButton", command=self.carregar).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(self.footer_right, text="🗺 MAPA SELECIONADO", style="Primary.TButton", command=self._abrir_mapa_selecionado).grid(
            row=0, column=1, padx=4
        )
        ttk.Button(self.footer_right, text="🌐 MAPA DE TODAS", style="Primary.TButton", command=self._abrir_mapa_todas).grid(
            row=0, column=2, padx=4
        )

        self._rows_cache = []
        self._refresh_job = None
        self._refresh_ms = 10000
        self._refresh_var = tk.StringVar(value="10s")
        ttk.Label(self.footer_left, text="AtualizaÃ§Ã£o automÃ¡tica:", style="CardLabel.TLabel").grid(
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
        with get_db() as conn:
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()

            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='rota_gps_pings'")
            has_gps = cur.fetchone() is not None

            if has_gps:
                sql = """
                    SELECT
                        p.codigo_programacao,
                        COALESCE(p.motorista, '') AS motorista,
                        COALESCE(p.veiculo, '') AS veiculo,
                        COALESCE(p.status, '') AS status,
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
                    WHERE UPPER(TRIM(COALESCE(p.status, ''))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')
                    ORDER BY p.id DESC
                    LIMIT 300
                """
            else:
                sql = """
                    SELECT
                        p.codigo_programacao,
                        COALESCE(p.motorista, '') AS motorista,
                        COALESCE(p.veiculo, '') AS veiculo,
                        COALESCE(p.status, '') AS status,
                        NULL AS lat,
                        NULL AS lon,
                        NULL AS speed,
                        NULL AS accuracy,
                        NULL AS recorded_at
                    FROM programacoes p
                    WHERE UPPER(TRIM(COALESCE(p.status, ''))) NOT IN ('FINALIZADA', 'FINALIZADO', 'CANCELADA', 'CANCELADO')
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
        self.set_status(f"STATUS: {len(self._rows_cache)} rota(s) ativa(s). GPS em {com_gps} rota(s).")

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
            f"STATUS: atualizaÃ§Ã£o automÃ¡tica ajustada para {sel or '10s'}."
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
# ORDENAÃ‡ÃƒO UNIVERSAL PARA TREEVIEW (CLIQUE NO CABEÃ‡ALHO)
# =========================================================
def enable_treeview_sorting(tree: ttk.Treeview, numeric_cols=None, money_cols=None, date_cols=None):
    """
    Habilita recursos de tabela no Treeview:
    - Ordena??o por cabe?alho
    - Indicador de filtro ao passar mouse no cabe?alho
    - Filtro por coluna (valores espec?ficos)

    Uso do filtro:
    - Passe o mouse no cabe?alho para ver o indicador "\u23F7"
    - Clique no indicador (lado direito do cabe?alho) ou bot?o direito no cabe?alho
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
        ttk.Button(btns, text="✔ Aplicar", style="Primary.TButton", command=_apply_and_close).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(btns, text="🧹 Limpar coluna", style="Ghost.TButton", command=_clear_col_filter).grid(row=0, column=1, padx=6)
        ttk.Button(btns, text="🧹 LIMPAR TUDO", style="Ghost.TButton", command=_clear_all_filters).grid(row=0, column=2, padx=6)

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
_MOTORISTA_COD_RE = re.compile(r"^[A-Z0-9._-]{3,24}$")  # jÃ¡ normaliza em upper()

def normalize_cpf(v: str) -> str:
    return _CPF_DIGITS.sub("", str(v or "").strip())

def is_valid_cpf(cpf_digits: str) -> bool:
    """
    Validador de CPF (Brasil).
    Aceita apenas 11 dÃ­gitos e verifica dÃ­gitos verificadores.
    """
    cpf = normalize_cpf(cpf_digits)
    if len(cpf) != 11:
        return False
    if cpf == cpf[0] * 11:
        return False

    try:
        nums = [int(x) for x in cpf]

        # 1Âº dÃ­gito
        s1 = sum(nums[i] * (10 - i) for i in range(9))
        d1 = (s1 * 10) % 11
        d1 = 0 if d1 == 10 else d1
        if d1 != nums[9]:
            return False

        # 2Âº dÃ­gito
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
    Normaliza telefone para dÃ­gitos.
    Se vier com DDI 55 (Brasil), remove quando fizer sentido.
    """
    s = _PHONE_DIGITS.sub("", str(v or "").strip())
    # remove 55 se vier com DDI e sobrar 10/11 dÃ­gitos
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
    SQLite nÃ£o aceita parametrizar PRAGMA table_info(?)  entÃ£o validamos.
    """
    name = str(name or "").strip()
    if not _SAFE_SQL_IDENT.match(name):
        raise ValueError(f"Identificador SQL invÃ¡lido: {name!r}")
    return name


def db_has_column(cur, table_name: str, column_name: str) -> bool:
    """
    Verifica se uma coluna existe numa tabela (SQLite).
    SeguranÃ§a: nÃ£o altera nada, sÃ³ consulta PRAGMA.
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
    Retorna lista de programaÃ§Ãµes ATIVAS (compatÃ­vel com bases antigas e novas).
    SaÃ­da: lista de dicts: {codigo, motorista, veiculo, equipe, data_criacao, status}
    """
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
                # base antiga sem status: pega as Ãºltimas como "ativas"
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
                ])
                join_ctrl = "LEFT JOIN programacao_itens_controle pc ON pc.codigo_programacao = pi.codigo_programacao AND UPPER(pc.cod_cliente)=UPPER(pi.cod_cliente)"
            else:
                select_cols.extend([
                    "0 as mortalidade_aves",
                    "0 as peso_previsto",
                    "0 as valor_recebido",
                    "'' as forma_recebimento",
                    "'' as obs_recebimento",
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
            })
        return out

    except Exception:
        return []


def format_prog_display(p: dict) -> str:
    """
    Formata texto amigÃ¡vel p/ lista/combobox sem quebrar.
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

# =========================================================
# 3.0.1 CLIENTES (IMPORTAÃ‡ÃƒO + EDIÃ‡ÃƒO DIRETA NA TABELA)
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

        ttk.Button(actions, text="📥 IMPORTAR CLIENTES (EXCEL)", style="Warn.TButton",
                   command=self.importar_clientes_excel).grid(row=0, column=0, padx=6)

        ttk.Button(actions, text="🔄 ATUALIZAR", style="Ghost.TButton",
                   command=self.carregar).grid(row=0, column=1, padx=6)

        ttk.Button(actions, text="➕ INSERIR LINHA", style="Ghost.TButton",
                   command=self.inserir_linha).grid(row=0, column=2, padx=6)

        ttk.Button(actions, text="SALVAR ALTERAÃ‡Ã•ES", style="Primary.TButton",
                   command=self.salvar_alteracoes).grid(row=0, column=3, padx=6)

        self.lbl_info = ttk.Label(
            actions,
            text="Dica: duplo clique na cÃ©lula para editar. ENTER salva a cÃ©lula. ESC cancela.",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        )
        self.lbl_info.grid(row=0, column=4, padx=12, sticky="w")

        cols = ["CÃ“D CLIENTE", "NOME CLIENTE", "ENDEREÃ‡O", "TELEFONE", "VENDEDOR"]
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

        self.tree.column("ENDEREÃ‡O", width=420, minwidth=200)
        self.tree.bind("<Double-1>", self._start_edit_cell)

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"CÃ“D CLIENTE"},
            money_cols=set(),
            date_cols=set()
        )

        # Se rolar o tree durante ediÃ§Ã£o, fecha/commita corretamente
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

    def carregar(self):
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
        if self._editing:
            self._commit_edit()

        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("ATENÃ‡ÃƒO", "NÃ£o hÃ¡ dados para salvar.")
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
                messagebox.showwarning("ATENÃ‡ÃƒO", "Todas as linhas precisam ter pelo menos CÃ“D CLIENTE e NOME CLIENTE.")
                return

            cod_key = upper(cod_norm)
            if cod_key in cods_seen:
                messagebox.showwarning("ATENÃ‡ÃƒO", f"CÃ“D CLIENTE duplicado na tabela: {cod_norm}\n\nCorrija antes de salvar.")
                return
            cods_seen.add(cod_key)

            linhas.append((upper(cod_norm), upper(nome_norm), upper(endereco), upper(telefone), upper(vendedor)))

        if not linhas:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Nenhuma linha vÃ¡lida para salvar.")
            return

        total = 0
        try:
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

            messagebox.showinfo("OK", f"Clientes salvos/atualizados: {total}")

        except Exception as e:
            messagebox.showerror("ERRO", str(e))

        self.carregar()

    def importar_clientes_excel(self):
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

            col_cod = guess_col(df.columns, ["cod", "cÃ³d", "codigo", "cliente", "cod cliente"])
            col_nome = guess_col(df.columns, ["nome", "cliente"])
            col_end = guess_col(df.columns, ["endereco", "endereÃ§o", "rua", "logradouro"])
            col_tel = guess_col(df.columns, ["telefone", "fone", "celular", "contato"])
            col_vendedor = guess_col(df.columns, ["vendedor", "vend", "representante"])

            if not col_cod or not col_nome:
                messagebox.showerror("ERRO", "NÃƒO IDENTIFIQUEI AS COLUNAS DE CÃ“DIGO E NOME DO CLIENTE NO EXCEL.")
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

            messagebox.showinfo("OK", f"CLIENTES IMPORTADOS/ATUALIZADOS: {total}")
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
            ],
            app=app
        )
        crud_motoristas.pack(fill="both", expand=True)
        nb.add(frm_motoristas, text="Motoristas")

        frm_usuarios = ttk.Frame(nb, style="Content.TFrame")
        crud_usuarios = CadastroCRUD(
            frm_usuarios,
            "UsuÃ¡rios",
            "usuarios",
            [
                ("nome", "NOME"),
                ("codigo", "CODIGO"),
                ("senha", "SENHA"),
                ("permissoes", "PERMISSÃ•ES"),
                ("cpf", "CPF"),
                
                ("telefone", "TELEFONE"),
            ],
            app=app
        )
        crud_usuarios.pack(fill="both", expand=True)
        nb.add(frm_usuarios, text="UsuÃ¡rios")

        frm_veiculos = ttk.Frame(nb, style="Content.TFrame")
        crud_veiculos = CadastroCRUD(
            frm_veiculos,
            "VeÃ­culos",
            "veiculos",
            [
                ("placa", "PLACA"),
                ("modelo", "MODELO"),
                
                ("capacidade_cx", "CAPACIDADE (CX)"),
            ],
            app=app
        )
        crud_veiculos.pack(fill="both", expand=True)
        nb.add(frm_veiculos, text="VeÃ­culos")

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
# 4.0 IMPORTAÃ‡ÃƒO DE VENDAS
# =========================================================
class ImportarVendasPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Importar Vendas (Excel)")

        top = ttk.Frame(self.body, style="Content.TFrame")
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(3, weight=1)

        ttk.Button(top, text="📥 IMPORTAR EXCEL", style="Primary.TButton", command=self.importar_excel).grid(row=0, column=0, padx=6)
        ttk.Button(top, text="🧹 LIMPAR TUDO", style="Danger.TButton", command=self.limpar_tudo).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(top, text="🔄 ATUALIZAR", style="Ghost.TButton", command=self.carregar).grid(row=0, column=2, padx=6)

        self.lbl_info = ttk.Label(
            top,
            text="Selecione as vendas que irÃ£o para ProgramaÃ§Ã£o (duplo clique marca/desmarca).",
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

        ttk.Button(filt, text="🔎 FILTRAR", style="Ghost.TButton", command=self.carregar).grid(row=0, column=2, padx=6)
        ttk.Button(filt, text="✅ MARCAR", style="Primary.TButton", command=self.marcar_selecionadas).grid(row=0, column=3, padx=6)

        ttk.Button(filt, text="✅ MARCAR TODOS", style="Warn.TButton", command=lambda: self.set_all_selected(1)).grid(row=0, column=4, padx=6)
        ttk.Button(filt, text="☑ DESMARCAR TODOS", style="Ghost.TButton", command=lambda: self.set_all_selected(0)).grid(row=0, column=5, padx=6)

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
        self.set_status("STATUS: ImportaÃ§Ã£o e seleÃ§Ã£o de vendas para programaÃ§Ã£o.")
        self.carregar()

    # -------------------------
    # Helpers seguranÃ§a/normalizaÃ§Ã£o
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

    def _ensure_vendas_usada_cols(self):
        """Garante colunas para evitar reutilizaÃ§Ã£o (nÃ£o quebra bases antigas)."""
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

            # Exige apenas o nucleo obrigatorio; demais campos sao opcionais.
            # Isso permite importar arquivos com colunas extras/faltantes sem quebrar.
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

            self._ensure_vendas_usada_cols()

            with get_db() as conn:
                cur = conn.cursor()

                for _, r in df.iterrows():
                    pedido = self._clean_pedido(r.get(col_pedido, ""))
                    if self._is_invalid_token(pedido):
                        ignoradas_invalidas += 1
                        continue

                    # Campos opcionais com fallback seguro
                    data_venda = self._excel_text(r.get(col_data, "")) if col_data else ""
                    cliente = self._excel_text(r.get(col_cliente, ""))

                    nome_cliente = self._excel_text(r.get(col_nome, "")) if col_nome else ""
                    vendedor = self._excel_text(r.get(col_vend, "")) if col_vend else ""
                    produto = self._excel_text(r.get(col_prod, "")) if col_prod else ""
                    vr_total = safe_float(r.get(col_vr_total, 0)) if col_vr_total else 0.0
                    qnt = safe_float(r.get(col_qnt, 0)) if col_qnt else 0.0
                    cidade = self._excel_text(r.get(col_cidade, "")) if col_cidade else ""
                    valor_unit = (vr_total / qnt) if qnt else 0.0
                    obs = self._excel_text(r.get(col_obs, "")) if col_obs else ""

                    # âœ… Chave natural para evitar duplicidade (sem mexer no banco)
                    pedido_u = self._norm(pedido)
                    cliente_u = self._norm(cliente)
                    produto_u = self._norm(produto)
                    nome_u = self._norm(nome_cliente)

                    # Proteção extra: bloqueia importação de linhas corrompidas (NaN/NaT)
                    if (not pedido_u) or (not cliente_u) or (not nome_u) or (not produto_u):
                        ignoradas_invalidas += 1
                        continue

                    try:
                        cur.execute("""
                            SELECT 1
                            FROM vendas_importadas
                            WHERE pedido=? AND cliente=? AND produto=? AND data_venda=?
                            LIMIT 1
                        """, (pedido_u, cliente_u, produto_u, data_venda))
                        exists = cur.fetchone()
                    except Exception:
                        exists = None  # se schema mudar, nÃ£o bloqueia importaÃ§Ã£o

                    if exists:
                        ignoradas += 1
                        continue

                    # âœ… Insere como NÃƒO usada por padrÃ£o
                    cur.execute("""
                        INSERT INTO vendas_importadas
                        (pedido, data_venda, cliente, nome_cliente, vendedor, produto, vr_total, qnt, cidade, valor_unitario, observacao, selecionada, usada, usada_em, codigo_programacao)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, '', '')
                    """, (
                        pedido_u,
                        data_venda,
                        cliente_u,
                        nome_u,
                        self._norm(vendedor),
                        produto_u,
                        float(vr_total or 0),
                        float(qnt or 0),
                        self._norm(cidade),
                        float(valor_unit or 0),
                        self._norm(obs),
                    ))
                    total += 1

            msg = f"Vendas importadas: {total}"
            if ignoradas:
                msg += f"\nIgnoradas (duplicadas/invÃ¡lidas): {ignoradas}"
            if ignoradas_invalidas:
                msg += f"\nIgnoradas (linhas inválidas/NaN): {ignoradas_invalidas}"

            if hasattr(self, "ent_busca"):
                try:
                    self.ent_busca.delete(0, "end")
                except Exception:
                    logging.debug("Falha ao limpar busca apÃ³s importaÃ§Ã£o")

            if opcionais_ausentes:
                msg += "\nCampos opcionais nao encontrados (preenchidos em branco): " + ", ".join(opcionais_ausentes)

            messagebox.showinfo("OK", msg)
            self.carregar()

        except Exception as e:
            messagebox.showerror("ERRO", str(e))

    def carregar(self):
        self._ensure_vendas_usada_cols()

        busca = self._norm(self.ent_busca.get()) if hasattr(self, "ent_busca") else ""

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

            # Quadradinhos de seleção (ASCII), evitando problemas de codificação.
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

        self.set_status(f"STATUS: {len(rows)} registros carregados (NÃƒO usadas)  Selecionadas: {selected_count}.")

    def toggle_selected(self, event=None):
        # âœ… sÃ³ alterna se clicar em uma cÃ©lula (evita bug ao clicar em cabeÃ§alho)
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

        with get_db() as conn:
            cur = conn.cursor()
            # âœ… sÃ³ marca/desmarca as que nÃ£o foram usadas
            cur.execute("UPDATE vendas_importadas SET selecionada=? WHERE IFNULL(usada,0)=0", (int(val),))
        self.carregar()

    def marcar_selecionadas(self):
        """
        Marca apenas as vendas selecionadas na tabela (suporta seleÃ§Ã£o mÃºltipla com Shift/Ctrl).
        """
        self._ensure_vendas_usada_cols()
        itens = self.tree.selection() or ()
        if not itens:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma ou mais vendas para marcar.")
            return

        ids = []
        for iid in itens:
            vals = self.tree.item(iid, "values") or ()
            if not vals:
                continue
            rid = safe_int(vals[0], 0)
            if rid > 0:
                ids.append(rid)

        if not ids:
            messagebox.showwarning("ATENÃ‡ÃƒO", "NÃ£o foi possÃ­vel identificar as vendas selecionadas.")
            return

        with get_db() as conn:
            cur = conn.cursor()
            cur.executemany(
                "UPDATE vendas_importadas SET selecionada=1 WHERE id=? AND IFNULL(usada,0)=0",
                [(rid,) for rid in ids],
            )

        self.carregar()
        self.set_status(f"STATUS: {len(ids)} venda(s) marcada(s) a partir da seleÃ§Ã£o.")

    def limpar_tudo(self):
        if not messagebox.askyesno("CONFIRMAR", "Deseja apagar TODAS as vendas importadas?"):
            return

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM vendas_importadas")

        self.carregar()


# ==========================
# ===== FIM DA PARTE 4 (ATUALIZADA) =====
# ==========================

# ==========================
# ===== INCIO DA PARTE 5 (ATUALIZADA) =====
# ==========================

# =========================================================
# 5.0 PROGRAMACAO PAGE
# =========================================================
class ProgramacaoPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "ProgramaÃ§Ã£o")

        self._editing = None
        self._prog_cols_checked = False  # evita PRAGMA/ALTER toda hora
        self._vendas_cols_checked = False  # evita PRAGMA/ALTER toda hora
        self._equipe_display_map = {}

        # -------------------------
        # CabeÃ§alho (dados da programaÃ§Ã£o)
        # -------------------------
        card = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        card.grid(row=0, column=0, sticky="ew")
        # Distribui o espa?o entre os campos para evitar "sumir" o ?ltimo campo (C?digo)
        for col in range(0, 9):
            card.grid_columnconfigure(col, weight=1, uniform="prog_head")

        ttk.Label(card, text="Motorista", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.cb_motorista = ttk.Combobox(card, state="readonly", width=16)
        self.cb_motorista.grid(row=1, column=0, sticky="ew", padx=6)

        ttk.Label(card, text="VeÃ­culo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.cb_veiculo = ttk.Combobox(card, state="readonly", width=12)
        self.cb_veiculo.grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(card, text="Ajudantes (multipla escolha)", style="CardLabel.TLabel").grid(row=0, column=2, sticky="w")
        self.btn_ajudantes = ttk.Button(card, text="👥 Selecionar ajudantes", style="Ghost.TButton", command=self._open_ajudantes_selector)
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

        ttk.Label(card, text="KG Estimado", style="CardLabel.TLabel").grid(row=0, column=5, sticky="w")
        self.ent_kg = ttk.Entry(card, style="Field.TEntry", width=10)
        self.ent_kg.grid(row=1, column=5, sticky="ew", padx=6)

        ttk.Label(card, text="Carregamento", style="CardLabel.TLabel").grid(row=0, column=6, sticky="w")
        self.ent_carregamento_prog = ttk.Entry(card, style="Field.TEntry", width=12)
        self.ent_carregamento_prog.grid(row=1, column=6, sticky="ew", padx=6)
        bind_entry_smart(self.ent_carregamento_prog, "text")

        ttk.Label(card, text="Adiantamento (R$)", style="CardLabel.TLabel").grid(row=0, column=7, sticky="w")
        self.ent_adiantamento_prog = ttk.Entry(card, style="Field.TEntry", width=12)
        self.ent_adiantamento_prog.grid(row=1, column=7, sticky="ew", padx=6)
        self.ent_adiantamento_prog.insert(0, "0,00")
        self._bind_money_entry(self.ent_adiantamento_prog)

        ttk.Label(card, text="CÃ³digo", style="CardLabel.TLabel").grid(row=0, column=8, sticky="w")
        self.ent_codigo = ttk.Entry(card, style="Field.TEntry", state="readonly", width=10)
        self.ent_codigo.grid(row=1, column=8, sticky="ew", padx=6)

        # -------------------------
        # Itens (vendas / ediÃ§Ã£o)
        # -------------------------
        card2 = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        card2.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.body.grid_rowconfigure(1, weight=1)

        card2.grid_columnconfigure(0, weight=1)
        card2.grid_rowconfigure(1, weight=1)

        top2 = ttk.Frame(card2, style="Card.TFrame")
        top2.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        top2.grid_columnconfigure(10, weight=1)

        ttk.Button(
            top2,
            text="CARREGAR VENDAS SELECIONADAS",
            style="Warn.TButton",
            command=self.carregar_vendas_selecionadas
        ).grid(row=0, column=0, padx=6)

        ttk.Button(top2, text="➕ INSERIR LINHA", style="Ghost.TButton",
                   command=self.inserir_linha).grid(row=0, column=1, padx=6)

        ttk.Button(top2, text="➖ REMOVER LINHA", style="Danger.TButton",
                   command=self.remover_linha).grid(row=0, column=2, padx=6)

        ttk.Button(top2, text="🧹 LIMPAR ITENS", style="Danger.TButton",
                   command=self.limpar_itens).grid(row=0, column=3, padx=6)

        ttk.Button(
            top2,
            text="SALVAR PROGRAMAÃ‡ÃƒO",
            style="Primary.TButton",
            command=self.salvar_programacao
        ).grid(row=0, column=4, padx=6)

        ttk.Button(
            top2,
            text="IMPRIMIR ROMANEIOS",
            style="Ghost.TButton",
            command=self.imprimir_romaneios_programacao
        ).grid(row=0, column=5, padx=6)

        ttk.Label(
            top2,
            text="Dica: duplo clique para editar EndereÃ§o/Caixas/KG/PreÃ§o/Vendedor/Pedido/Obs. ENTER confirma, ESC cancela.",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 8, "bold")
        ).grid(row=0, column=6, padx=12, sticky="w")

        cols = ["COD CLIENTE", "NOME CLIENTE", "PRODUTO", "ENDEREÃ‡O", "CAIXAS", "KG", "PREÃ‡O", "VENDEDOR", "PEDIDO", "OBS"]

        table_wrap = ttk.Frame(card2, style="Card.TFrame")
        table_wrap.grid(row=1, column=0, sticky="nsew")
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
            self.tree.column(c, width=160, minwidth=90, anchor="w")

        self.tree.column("ENDEREÃ‡O", width=260, minwidth=180)
        self.tree.column("NOME CLIENTE", width=260, minwidth=160)
        self.tree.column("PRODUTO", width=160, minwidth=120)
        self.tree.column("PEDIDO", width=160, minwidth=120)
        self.tree.column("OBS", width=260, minwidth=140)
        self.tree.column("CAIXAS", width=90, anchor="center")
        self.tree.column("KG", width=90, anchor="center")
        self.tree.column("PREÃ‡O", width=110, anchor="e")

        self.tree.bind("<Double-1>", self._start_edit_cell)
        self.tree.bind("<MouseWheel>", self._on_tree_scroll, add=True)
        self.tree.bind("<Button-4>", self._on_tree_scroll, add=True)  # linux
        self.tree.bind("<Button-5>", self._on_tree_scroll, add=True)  # linux

        enable_treeview_sorting(
            self.tree,
            numeric_cols={"CAIXAS", "KG"},
            money_cols={"PREÃ‡O"},
            date_cols=set()
        )

        self.refresh_comboboxes()

    # -------------------------
    # UtilitÃ¡rios
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
        Tolerante a acentos/mojibake (ex.: SERTÃO, SERTÃƒO).
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
            self.btn_ajudantes.configure(text="👥 Selecionar ajudantes")
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
        ttk.Label(search_row, text="🔎 BUSCAR", style="CardLabel.TLabel").pack(side="left", padx=(0, 8))
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
        ttk.Button(footer, text="🧹 LIMPAR", style="Ghost.TButton", command=self._clear_ajudantes_selection).pack(side="left")
        ttk.Button(footer, text="✔ CONFIRMAR", style="Primary.TButton", command=self._confirm_ajudantes_popup).pack(side="right")

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
        with get_db() as conn:
            cur = conn.cursor()

            valores_motoristas = []
            try:
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
        self.set_status("STATUS: Carregue vendas e ajuste dados antes de salvar a programaÃ§Ã£o.")
        self.refresh_comboboxes()

    # -------------------------
    # AÃ§Ãµes de itens
    # -------------------------
    def inserir_linha(self):
        tree_insert_aligned(self.tree, "", "end", ("", "", "", "", "1", "0.00", "0.00", "", "", ""))
        items = self.tree.get_children()
        if items:
            self.tree.see(items[-1])

    def remover_linha(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma linha para remover.")
            return
        for iid in sel:
            self.tree.delete(iid)

    def limpar_itens(self):
        self.tree.delete(*self.tree.get_children())

    def carregar_vendas_selecionadas(self):
        """
        âœ… Melhorado: carrega tudo em 1 query (sem N+1).
        TambÃ©m evita duplicar por (pedido + cod_cliente).
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

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT
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
                ORDER BY v.id ASC
            """)
            vendas = cur.fetchall() or []

        self.limpar_itens()

        seen = set()
        ignorados_invalidos = 0
        for (pedido, data_venda, cod_cliente, nome_cliente, vendedor, produto, vr_total, qnt, valor_unit, obs, cidade, endereco) in vendas:
            cod_cliente = upper(str(cod_cliente or "").strip())
            pedido_u = _clean_pedido_local(pedido)
            nome_u = upper(str(nome_cliente or "").strip())
            produto_u = upper(str(produto or "").strip())

            if _is_bad(cod_cliente) or _is_bad(pedido_u) or _is_bad(nome_u) or _is_bad(produto_u):
                ignorados_invalidos += 1
                continue

            key = (pedido_u, cod_cliente)
            if key in seen:
                continue
            seen.add(key)

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

        msg = f"STATUS: Itens carregados: {len(self.tree.get_children())} vendas selecionadas. (edite antes de salvar)"
        if ignorados_invalidos:
            msg += f" Ignoradas inválidas: {ignorados_invalidos}."
        self.set_status(msg)

    # -------------------------
    # EdiÃ§Ã£o de cÃ©lula (com regras)
    # -------------------------
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

        # âœ… Regra: sÃ³ permite editar essas colunas
        editable = {"ENDEREÃ‡O", "CAIXAS", "KG", "PREÃ‡O", "VENDEDOR", "PEDIDO", "OBS"}
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

        # âœ… ValidaÃ§Ã£o por tipo
        if col_name == "CAIXAS":
            v = safe_int(new_value, 0)
            if v < 0:
                v = 0
            new_value = str(v)

        elif col_name in {"KG", "PREÃ‡O"}:
            v = safe_float(new_value, 0.0)
            if v < 0:
                v = 0.0
            new_value = f"{v:.2f}"

        vals = self._get_row_values(row_id)
        vals[col_index] = new_value
        self.tree.item(row_id, values=tuple(vals))

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
        add_col("codigo", "TEXT")
        add_col("data", "TEXT")
        add_col("total_caixas", "INTEGER DEFAULT 0")
        add_col("quilos", "REAL DEFAULT 0")
        add_col("saida_dt", "TEXT")
        add_col("chegada_dt", "TEXT")

    def _ensure_vendas_usada_cols(self, cur):
        """Garante colunas para bloquear reutilizaÃ§Ã£o (compatÃ­vel com bases antigas)."""
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
        kg_estimado = safe_float(self.ent_kg.get(), 0.0)
        adiantamento_val = safe_money(self.ent_adiantamento_prog.get(), 0.0)
        if adiantamento_val < 0:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Adiantamento nÃ£o pode ser negativo.")
            return

        if not motorista_nome or not veiculo:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione Motorista e VeÃ­culo.")
            return
        if self._ajudantes_mode == "equipes":
            if not ajudante1:
                messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione a equipe da programaÃ§Ã£o.")
                return
        else:
            if not ajudante1 or not ajudante2:
                messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione exatamente 2 ajudantes da programaÃ§Ã£o.")
                return
            if upper(ajudante1) == upper(ajudante2):
                messagebox.showwarning("ATENÃ‡ÃƒO", "Os ajudantes selecionados devem ser diferentes.")
                return
        if local_rota not in {"SERRA", "SERTAO"}:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione o Local da Rota (SERRA ou SERTAO).")
            return
        if not local_carreg:
            messagebox.showwarning("ATENCAO", "Informe o local de Carregamento.")
            return

        itens = []
        for iid in self.tree.get_children():
            itens.append(self._get_row_values(iid))

        if not itens:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue itens (vendas selecionadas) antes de salvar.")
            return

        # âœ… validaÃ§Ã£o mÃ­nima por linha (seguranÃ§a)
        for v in itens:
            cod_cliente = v[0]
            nome_cliente = v[1]
            if not str(cod_cliente).strip() or not str(nome_cliente).strip():
                messagebox.showwarning("ATENÃ‡ÃƒO", "HÃ¡ linhas sem COD CLIENTE ou NOME CLIENTE. Corrija antes de salvar.")
                return

        codigo = None
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

        motorista_id = None

        try:
            with get_db() as conn:
                cur = conn.cursor()

                if not self._prog_cols_checked:
                    self._ensure_prog_columns_for_api(cur)
                    self._prog_cols_checked = True

                if not self._vendas_cols_checked:
                    self._ensure_vendas_usada_cols(cur)
                    self._vendas_cols_checked = True

                # Valida capacidade do veÃ­culo (CX) antes de salvar programaÃ§Ã£o.
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
                        "ATENÃ‡ÃƒO",
                        f"VeÃ­culo nÃ£o encontrado no cadastro: {veiculo}."
                    )
                    return

                capacidade_cx = safe_int(vrow[0], -1)
                if capacidade_cx < 0:
                    messagebox.showwarning(
                        "ATENÃ‡ÃƒO",
                        f"Capacidade (CX) invÃ¡lida para o veÃ­culo {veiculo}. Ajuste no cadastro de veÃ­culos."
                    )
                    return

                if total_caixas > capacidade_cx:
                    messagebox.showwarning(
                        "ATENÃ‡ÃƒO",
                        f"Capacidade excedida para o veÃ­culo {veiculo}.\n\n"
                        f"Caixas na programaÃ§Ã£o: {total_caixas}\n"
                        f"Capacidade do veÃ­culo: {capacidade_cx}"
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

                # Insert compatÃ­vel (tenta gerar cÃ³digo Ãºnico)
                inserted = False
                for _ in range(5):
                    codigo = generate_program_code()
                    try:
                        cur.execute(
                            "SELECT 1 FROM programacoes WHERE codigo_programacao=? LIMIT 1",
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
                                kg_estimado, status, total_caixas, quilos
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'ATIVA', ?, ?)
                        """, (
                            codigo, codigo, data_criacao, data_criacao,
                            motorista_nome, safe_int(motorista_id, 0), veiculo, equipe,
                            kg_estimado, total_caixas, total_quilos
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
                    raise sqlite3.IntegrityError("Falha ao gerar cÃ³digo Ãºnico para programaÃ§Ã£o.")

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

                # garante prestaÃ§Ã£o pendente quando coluna existir
                try:
                    if db_has_column(cur, "programacoes", "prestacao_status"):
                        cur.execute(
                            "UPDATE programacoes SET prestacao_status='PENDENTE' WHERE codigo_programacao=?",
                            (codigo,)
                        )
                except Exception:
                    logging.debug("Falha ignorada")

                # Itens
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

                    # MantÃ©m clientes atualizados (compatÃ­vel com bases antigas)
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
                # âœ… REGRA NOVA: Vendas selecionadas viram "usadas" e somem
                # =========================================================
                try:
                    cur.execute("""
                        UPDATE vendas_importadas
                        SET
                            usada = 1,
                            usada_em = ?,
                            codigo_programacao = ?,
                            selecionada = 0
                        WHERE selecionada = 1
                    """, (data_criacao, codigo))
                except Exception:
                    # fallback: pelo menos zera selecionada se algo der ruim
                    cur.execute("UPDATE vendas_importadas SET selecionada=0 WHERE selecionada=1")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao salvar programaÃ§Ã£o: {str(e)}")
            return

        self.ent_codigo.config(state="normal")
        self.ent_codigo.delete(0, "end")
        self.ent_codigo.insert(0, codigo)
        self.ent_codigo.config(state="readonly")

        messagebox.showinfo("OK", f"ProgramaÃ§Ã£o salva: {codigo} (ABERTA/ATIVA)")
        self.set_status(f"STATUS: ProgramaÃ§Ã£o salva: {codigo} (ABERTA/ATIVA)")

        self.app.refresh_programacao_comboboxes()

        if messagebox.askyesno("PDF", "Deseja gerar o PDF da programaÃ§Ã£o agora?\n\n(Pronto para impressÃ£o A4)"):
            self.gerar_pdf_programacao_salva(codigo, motorista_nome, veiculo, equipe, kg_estimado)

        if messagebox.askyesno("Romaneios", "Deseja gerar os romaneios de entrega desta programacao agora?"):
            self.imprimir_romaneios_programacao()

        self.limpar_itens()

    def gerar_pdf_programacao_salva(self, codigo, motorista, veiculo, equipe, kg_estimado):
        if not require_reportlab():
            return
        path = filedialog.asksaveasfilename(
            title="Salvar PDF da ProgramaÃ§Ã£o",
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Sem itens na programaÃ§Ã£o.")
            return

        try:
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
            c.drawString(40, y, f"KG Estimado: {kg_estimado:.2f}")
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
            messagebox.showinfo("OK", "PDF gerado com sucesso! (A4 pronto para impressÃ£o)")

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
            "aves_por_caixa": 6,
            "local_rota": "",
            "local_carregamento": "",
        }
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    SELECT COALESCE(motorista,''), COALESCE(veiculo,''), COALESCE(equipe,''),
                           COALESCE(data_criacao,''), COALESCE(kg_estimado,0)
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

                for cand in ("qnt_aves_por_cx", "aves_por_caixa", "qnt_aves_por_caixa"):
                    if db_has_column(cur, "programacoes", cand):
                        cur.execute(f"SELECT COALESCE({cand},0) FROM programacoes WHERE codigo_programacao=? LIMIT 1", (codigo,))
                        r = cur.fetchone()
                        apc = safe_int((r[0] if r else 0), 0)
                        if apc > 0:
                            meta["aves_por_caixa"] = apc
                            break

                for cand in ("local_rota", "tipo_rota"):
                    if db_has_column(cur, "programacoes", cand):
                        cur.execute(f"SELECT COALESCE({cand},'') FROM programacoes WHERE codigo_programacao=? LIMIT 1", (codigo,))
                        r = cur.fetchone()
                        v = upper((r[0] if r else "") or "")
                        if v:
                            meta["local_rota"] = v
                            break

                for cand in ("local_carregamento", "granja_carregada", "local_carregado", "local_carreg"):
                    if db_has_column(cur, "programacoes", cand):
                        cur.execute(f"SELECT COALESCE({cand},'') FROM programacoes WHERE codigo_programacao=? LIMIT 1", (codigo,))
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

        ttk.Button(bottom, text="⬅ Anterior", style="Ghost.TButton", command=_prev).pack(side="left")
        ttk.Button(bottom, text="Próximo ➡", style="Ghost.TButton", command=_next).pack(side="left", padx=8)
        ttk.Button(bottom, text="📄 GERAR PDF", style="Primary.TButton", command=_export_pdf).pack(side="right")
        ttk.Button(bottom, text="✖ Fechar", style="Danger.TButton", command=win.destroy).pack(side="right", padx=8)

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
# 6.0 FUNÃ‡ÃƒO SIMPLE INPUT (antes de RecebimentosPage)
# =========================================================
def simple_input(title, prompt, master=None, initial="", allow_empty=True, max_len=200):
    """
    Janela de diÃ¡logo simples para entrada de texto.

    CompatÃ­vel com seu uso atual:
        simple_input("TÃ­tulo", "Pergunta?")

    Extras opcionais (nÃ£o quebram):
        master=app  -> mantÃ©m a janela presa ao app
        initial="..." -> valor inicial
        allow_empty=False -> obriga preencher
        max_len -> limita tamanho (seguranÃ§a)
    """
    # Se nÃ£o passar master, tenta usar a janela root atual
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Preencha o campo antes de confirmar.")
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

    ttk.Button(left, text="✔ CONFIRMAR", style="Primary.TButton", command=ok).pack(side="left")
    ttk.Button(left, text="✖ CANCELAR", style="Ghost.TButton", command=cancel).pack(side="left", padx=8)

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

        ttk.Button(self.card, text="📂 CARREGAR", style="Ghost.TButton", command=self.carregar_programacao)\
            .grid(row=1, column=1, padx=(0, 14))

        info_frame = ttk.Frame(self.card, style="Card.TFrame")
        info_frame.grid(row=1, column=2, sticky="ew", padx=(8, 8))
        for i in range(2):
            info_frame.grid_columnconfigure(i, weight=1)

        self.lbl_motorista_info = ttk.Label(
            info_frame,
            text="Motorista: ?",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
            wraplength=260,
            justify="left"
        )
        self.lbl_motorista_info.grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.lbl_veiculo_info = ttk.Label(
            info_frame,
            text="Veiculo: ?",
            background="white",
            foreground="#6B7280",
            font=("Segoe UI", 9, "bold"),
            wraplength=260,
            justify="left"
        )
        self.lbl_veiculo_info.grid(row=1, column=0, sticky="w", padx=(0, 8))

        self.lbl_equipe_info = ttk.Label(
            info_frame,
            text="Equipe: ?",
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

        ttk.Button(top2, text="👤 INSERIR CLIENTE MANUAL", style="Warn.TButton", command=self.inserir_cliente_manual)\
            .grid(row=0, column=0, padx=6)

        ttk.Button(top2, text="🧽 ZERAR RECEBIMENTO", style="Danger.TButton", command=self.zerar_recebimento)\
            .grid(row=0, column=1, padx=6)

        ttk.Button(top2, text="➡ IR PARA DESPESAS", style="Primary.TButton", command=self._ir_para_despesas)\
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
        # FormulÃ¡rio de recebimento
        # -------------------------
        frm = ttk.Frame(self.card2, style="Card.TFrame")
        frm.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        frm.grid_columnconfigure(4, weight=1)

        ttk.Label(frm, text="CÃ³d Cliente", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
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

        ttk.Label(frm, text="ObservaÃ§Ã£o", style="CardLabel.TLabel").grid(row=0, column=4, sticky="w")
        self.ent_obs = ttk.Entry(frm, style="Field.TEntry", width=40)
        self.ent_obs.grid(row=1, column=4, sticky="ew", padx=6)
        bind_entry_smart(self.ent_obs, "text")

        ttk.Button(frm, text="💾 SALVAR RECEBIMENTOS", style="Primary.TButton", command=self.salvar_recebimento)\
            .grid(row=1, column=5, sticky="e", padx=(12, 0))

        # âœ… TROCA: no lugar do Excel, botÃ£o IMPRIMIR PDF
        ttk.Button(frm, text="🖨 IMPRIMIR PDF", style="Warn.TButton", command=self.imprimir_pdf)\
            .grid(row=1, column=6, sticky="e", padx=6)

        # Por padrÃ£o: COD/NOME nÃ£o editÃ¡veis
        self._set_cliente_fields_readonly(True)

        # -------------------------
        # Painel ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œmodo ocultoÃƒÂ¢Ã¢â€šÂ¬Ã‚ (apÃƒÆ’Ã‚Â³s finalizar prestaÃƒÆ’Ã‚Â§ÃƒÆ’Ã‚Â£o)
        # -------------------------
        self._wrap_collapsed = ttk.Frame(self.body, style="Card.TFrame", padding=14)
        self._wrap_collapsed.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        self._wrap_collapsed.grid_remove()

        self._lbl_collapsed = ttk.Label(
            self._wrap_collapsed,
            text="PRESTAÃ‡ÃƒO FECHADA / SALVA.\nCabeÃ§alhos e tabela foram ocultados.",
            background="white",
            foreground="#111827",
            font=("Segoe UI", 10, "bold"),
            justify="left"
        )
        self._lbl_collapsed.grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(self._wrap_collapsed, style="Card.TFrame")
        btns.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        btns.grid_columnconfigure(10, weight=1)

        ttk.Button(btns, text="🖨 IMPRIMIR PDF", style="Warn.TButton", command=self.imprimir_pdf)\
            .grid(row=0, column=0, padx=6, sticky="w")

        ttk.Button(btns, text="👁 MOSTRAR DADOS (CONSULTA)", style="Ghost.TButton", command=self._expand_view)\
            .grid(row=0, column=1, padx=6, sticky="w")

        ttk.Button(btns, text="LIMPAR / NOVA PROGRAMAÃ‡ÃƒO", style="Primary.TButton", command=self._reset_view)\
            .grid(row=0, column=2, padx=6, sticky="w")

        self.refresh_comboboxes()
        self.carregar_tabela_vazia()
        self._refresh_diarias_preview()

    # =========================
    # Modos de visualizaÃ§Ã£o (ocultar/mostrar)
    # =========================
    def _collapse_view(self):
        """Oculta cabeÃ§alhos e dados de tabela (mantÃ©m imprimir/limpar)."""
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
        """Mostra novamente cabeÃ§alho e tabela."""
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
        """Limpa seleÃ§Ã£o e volta para estado inicial."""
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
    # Helpers de seguranÃ§a UI (readonly sem quebrar inserts)
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
    # ResoluÃ§Ã£o de nomes (motorista/equipe)
    # -------------------------
    def _resolve_motorista_nome(self, motorista_raw: str) -> str:
        m = (motorista_raw or "").strip()
        if not m:
            return ""
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
    # OrdenaÃ§Ã£o (padronizada / robusta)
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

        arrow = " Ã¢â€ â€œ" if reverse else " Ã¢â€ â€˜"
        for c in self.tree["columns"]:
            current_text = self.tree.heading(c, "text")
            if current_text.endswith(" Ã¢â€ â€˜") or current_text.endswith(" Ã¢â€ â€œ"):
                current_text = current_text[:-2]
            self.tree.heading(c, text=current_text)

        current_text = self.tree.heading(col, "text")
        if current_text.endswith(" Ã¢â€ â€˜") or current_text.endswith(" Ã¢â€ â€œ"):
            current_text = current_text[:-2]
        self.tree.heading(col, text=current_text + arrow)

    # -------------------------
    # Combobox programaÃ§Ã£o (pendentes)
    # -------------------------
    def refresh_comboboxes(self):
        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("""
                    SELECT codigo_programacao
                    FROM programacoes
                    WHERE (
                        status='ATIVA'
                        OR status='EM_ROTA'
                        OR status='INICIADA'
                        OR status='CARREGADA'
                        OR status='FINALIZADA'
                    )
                      AND (prestacao_status IS NULL OR prestacao_status='PENDENTE')
                    ORDER BY id DESC
                    LIMIT 300
                """)
            except Exception:
                cur.execute("""
                    SELECT codigo_programacao
                    FROM programacoes
                    WHERE (
                        status='ATIVA'
                        OR status='EM_ROTA'
                        OR status='INICIADA'
                        OR status='CARREGADA'
                        OR status='FINALIZADA'
                    )
                    ORDER BY id DESC
                    LIMIT 300
                """)
            self.cb_prog["values"] = [r[0] for r in cur.fetchall()]

    def _sync_horarios_from_programacao(self, prog: str) -> bool:
        """Sincroniza campos de saida/chegada com o que veio do app mobile via API."""
        prog = upper(prog)
        if not prog:
            return False

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT data_saida, hora_saida, data_chegada, hora_chegada
                FROM programacoes
                WHERE codigo_programacao=?
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

    # -------------------------
    # Regras: bloqueio quando FECHADA
    # -------------------------
    def _is_prestacao_fechada(self, prog: str) -> bool:
        if not prog:
            return False
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT prestacao_status FROM programacoes WHERE codigo_programacao=? LIMIT 1", (prog,))
                r = cur.fetchone()
                st = upper(r[0]) if r and r[0] is not None else ""
                return st == "FECHADA"
        except Exception:
            return False

    def _warn_if_fechada(self) -> bool:
        if self._is_prestacao_fechada(self._current_prog):
            messagebox.showwarning("ATENÃ‡ÃƒO", "Esta prestaÃ§Ã£o jÃ¡ estÃ¡ FECHADA. NÃ£o Ã© possÃ­vel alterar recebimentos.")
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

                # Bloqueia valores invÃ¡lidos removendo o Ãºltimo dÃ­gito invÃ¡lido.
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
                if digits and len(digits) not in (8,):
                    ent.delete(0, "end")
                    digits_state["value"] = ""
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
            if n == 4:
                mm = int(digits[2:4])
                if mm < 0 or mm > 59:
                    return False
            return True

        def _apply_mask():
            try:
                raw = str(ent.get() or "")
                digits = re.sub(r"\D", "", raw)[:4]

                # Bloqueia valores invÃ¡lidos removendo o Ãºltimo dÃ­gito invÃ¡lido.
                while digits and (not _is_partial_time_digits_valid(digits)):
                    digits = digits[:-1]

                if len(digits) <= 2:
                    masked = digits
                else:
                    masked = f"{digits[:2]}:{digits[2:]}"
                ent.delete(0, "end")
                ent.insert(0, masked)
                ent.icursor("end")
                digits_state["value"] = digits
            except Exception:
                logging.debug("Falha ignorada")

        def _on_focus_out():
            try:
                digits = digits_state["value"]
                if digits and len(digits) not in (4,):
                    ent.delete(0, "end")
                    digits_state["value"] = ""
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue uma programaÃ§Ã£o primeiro.")
            return False
        prog = self._current_prog
        # Tenta persistir o cabeÃ§alho sem bloquear a navegaÃ§Ã£o.
        try:
            self.salvar_dados_rota(silent=True)
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
            self.set_status(f"STATUS: Indo para DESPESAS - programaÃ§Ã£o {prog}.")
            return True
        except Exception:
            logging.debug("Falha ignorada")
            return False

    def _ir_para_despesas(self):
        self._abrir_despesas_com_programacao()

    # -------------------------
    # Carregar programaÃ§Ã£o
    # -------------------------
    def carregar_programacao(self):
        prog = upper(self.cb_prog.get())
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma programaÃ§Ã£o pendente.")
            return

        self._current_prog = prog

        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("PRAGMA table_info(programacoes)")
                cols_prog = [str(r[1]).lower() for r in cur.fetchall()]
            except Exception:
                cols_prog = []
            col_rota = "local_rota" if "local_rota" in cols_prog else ("tipo_rota" if "tipo_rota" in cols_prog else "''")
            col_diaria = "diaria_motorista_valor" if "diaria_motorista_valor" in cols_prog else "0"
            cur.execute("""
                SELECT motorista, veiculo, equipe, num_nf, data_saida, hora_saida, data_chegada, hora_chegada,
                       {col_rota} as rota, COALESCE({col_diaria}, 0) as diaria_motorista
                FROM programacoes
                WHERE codigo_programacao=?
            """.format(col_rota=col_rota, col_diaria=col_diaria), (prog,))
            row = cur.fetchone()

        if not row:
            messagebox.showwarning("ATENÃ‡ÃƒO", f"ProgramaÃ§Ã£o nÃ£o encontrada: {prog}")
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
            self.set_status(f"STATUS: Programacao carregada: {prog}")

    # -------------------------
    # ValidaÃ§Ã£o leve de data/hora
    # -------------------------
    def _validate_date(self, s: str) -> bool:
        s = (s or "").strip()
        if not s:
            return True
        if "-" in s and len(s) >= 10:
            try:
                y, m, d = s[:10].split("-")
                int(y); int(m); int(d)
                return True
            except Exception:
                return False
        if "/" in s and len(s) >= 10:
            try:
                d, m, y = s[:10].split("/")
                int(y); int(m); int(d)
                return True
            except Exception:
                return False
        return False

    def _validate_time(self, s: str) -> bool:
        s = (s or "").strip()
        if not s:
            return True
        try:
            hh, mm = s.split(":")[:2]
            hh = int(hh); mm = int(mm)
            return 0 <= hh <= 23 and 0 <= mm <= 59
        except Exception:
            return False

    def salvar_dados_rota(self, silent: bool = False):
        if not self._current_prog:
            if not silent:
                messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue uma programaÃ§Ã£o primeiro.")
            return False
        if self._warn_if_fechada():
            return False

        diaria_motorista = safe_money(self.ent_diaria_motorista.get(), 0.0)
        if diaria_motorista < 0:
            if not silent:
                messagebox.showwarning("ATENÃ‡ÃƒO", "DiÃ¡ria do motorista nÃ£o pode ser negativa.")
            return False

        data_saida = normalize_date(self.ent_data_saida.get())
        hora_saida = normalize_time(self.ent_hora_saida.get())
        data_chegada = normalize_date(self.ent_data_chegada.get())
        hora_chegada = normalize_time(self.ent_hora_chegada.get())

        if data_saida is None or data_chegada is None:
            if not silent:
                messagebox.showwarning("ATENÃ‡ÃƒO", "Formato de data invÃ¡lido. Use dd/mm/aaaa ou yyyy-mm-dd.")
            return False
        if hora_saida is None or hora_chegada is None:
            if not silent:
                messagebox.showwarning("ATENÃ‡ÃƒO", "Formato de hora invÃ¡lido. Use HH:MM (ex.: 07:30).")
            return False

        # atualiza campos com formato normalizado
        self._safe_set_entry(self.ent_data_saida, data_saida, readonly_back=False)
        self._safe_set_entry(self.ent_hora_saida, hora_saida, readonly_back=False)
        self._safe_set_entry(self.ent_data_chegada, data_chegada, readonly_back=False)
        self._safe_set_entry(self.ent_hora_chegada, hora_chegada, readonly_back=False)
        self._safe_set_entry(self.ent_diaria_motorista, f"{diaria_motorista:.2f}".replace(".", ","), readonly_back=False)
        self._refresh_diarias_preview()

        try:
            resumo_diarias = None
            with get_db() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("PRAGMA table_info(programacoes)")
                    cols_prog = [str(r[1]).lower() for r in cur.fetchall()]
                except Exception:
                    cols_prog = []
                cur.execute("""
                    UPDATE programacoes
                       SET data_saida=?,
                           hora_saida=?,
                           data_chegada=?,
                           hora_chegada=?
                     WHERE codigo_programacao=?
                """, (data_saida, hora_saida, data_chegada, hora_chegada, self._current_prog))
                if "diaria_motorista_valor" in cols_prog:
                    cur.execute(
                        "UPDATE programacoes SET diaria_motorista_valor=? WHERE codigo_programacao=?",
                        (diaria_motorista, self._current_prog)
                    )
                resumo_diarias = self._sync_diarias_despesas(
                    cur=cur,
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
                    "STATUS: DiÃ¡rias atualizadas - "
                    f"QTD: {resumo_diarias['qtd_diarias']:g} | "
                    f"Motorista: {fmt_money(resumo_diarias['total_motorista'])} | "
                    f"Equipe ({resumo_diarias['qtd_ajudantes']}): {fmt_money(resumo_diarias['total_ajudantes'])} | "
                    f"Total: {fmt_money(resumo_diarias['total_geral'])}"
                )
            else:
                self.set_status("STATUS: Dados do cabeÃ§alho atualizados.")
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
            return datetime(int(y), int(m), int(d), int(hh), int(mm))
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

    def _sync_diarias_despesas(self, cur, prog: str, rota: str, equipe_raw: str, data_saida: str, hora_saida: str, data_chegada: str, hora_chegada: str, diaria_motorista: float):
        qtd = self._calc_qtd_diarias_regra(data_saida, hora_saida, data_chegada, hora_chegada)
        diaria_motorista = safe_float(diaria_motorista, 0.0)
        qtd_ajudantes = self._count_ajudantes(equipe_raw)
        diaria_ajudante = max(diaria_motorista - 10.0, 0.0)
        total_mot = round(qtd * diaria_motorista, 2)
        total_ajud = round(qtd * (diaria_ajudante * qtd_ajudantes), 2)
        total_geral = round(total_mot + total_ajud, 2)

        obs_base = (
            f"AUTO_DIARIA|ROTA={upper(rota)}|QTD={qtd:g}|"
            f"AJUDANTES={qtd_ajudantes}|DIARIA_MOTORISTA={diaria_motorista:.2f}|DIARIA_AJUDANTE={diaria_ajudante:.2f}"
        )
        cur.execute("""
            DELETE FROM despesas
             WHERE codigo_programacao=?
               AND descricao IN ('DIARIAS MOTORISTA', 'DIARIAS AJUDANTES')
               AND COALESCE(observacao, '') LIKE 'AUTO_DIARIA|%'
        """, (prog,))

        now_s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute("""
            INSERT INTO despesas (codigo_programacao, descricao, valor, categoria, observacao, data_registro)
            VALUES (?, ?, ?, 'DIARIAS', ?, ?)
        """, (prog, "DIARIAS MOTORISTA", total_mot, obs_base, now_s))
        cur.execute("""
            INSERT INTO despesas (codigo_programacao, descricao, valor, categoria, observacao, data_registro)
            VALUES (?, ?, ?, 'DIARIAS', ?, ?)
        """, (prog, "DIARIAS AJUDANTES", total_ajud, obs_base, now_s))

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
    # SeleÃ§Ã£o / formulÃ¡rio
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
                messagebox.showwarning("ATENÃ‡ÃƒO", "Informe um valor vÃ¡lido (maior que zero).")
                return

            forma = upper(forma_atual or self.cb_forma.get() or "DINHEIRO")
            obs = upper(obs_atual or "")
            self._salvar_recebimento_item(cod=cod, nome=nome, valor=valor, forma=forma, obs=obs)
            self.carregar_clientes_e_recebimentos()
            self.set_status(f"STATUS: Recebimento de {fmt_money(valor)} salvo para {nome}.")
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao lanÃ§ar recebimento na tabela: {str(e)}")

    def _clear_form_recebimento(self):
        self._safe_set_entry(self.ent_cod, "", readonly_back=True)
        self._safe_set_entry(self.ent_nome, "", readonly_back=True)
        self.ent_valor.delete(0, "end")
        self.cb_forma.set("DINHEIRO")
        self.ent_obs.delete(0, "end")

    def _salvar_recebimento_item(self, cod: str, nome: str, valor: float, forma: str, obs: str):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue uma programaÃ§Ã£o primeiro.")
            return
        if self._warn_if_fechada():
            return
        if not cod or not nome:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione um cliente na tabela ou insira manualmente.")
            return
        if safe_float(valor, 0.0) <= 0:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Informe um valor vÃ¡lido (maior que zero).")
            return
        if not forma:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Informe a forma de pagamento.")
            return

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO recebimentos
                    (codigo_programacao, cod_cliente, nome_cliente, valor, forma_pagamento, observacao, num_nf, data_registro)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                prog,
                upper(cod),
                upper(nome),
                float(valor),
                upper(forma),
                upper(obs),
                upper(self._nf_current),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))

    # -------------------------
    # Salvar recebimento (com bloqueio se FECHADA)
    # -------------------------
    def salvar_recebimento(self):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue uma programaÃ§Ã£o primeiro.")
            return
        if self._warn_if_fechada():
            return

        cod = upper(self.ent_cod.get())
        nome = upper(self.ent_nome.get())
        valor = safe_money(self.ent_valor.get(), 0.0)
        forma = upper(self.cb_forma.get())
        obs = upper(self.ent_obs.get())

        if not cod or not nome:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione um cliente na tabela ou insira manualmente.")
            return
        if valor <= 0:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Informe um valor vÃ¡lido (maior que zero).")
            return
        if not forma:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Informe a forma de pagamento.")
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue uma programaÃ§Ã£o primeiro.")
            return
        if self._warn_if_fechada():
            return

        cod = upper(self.ent_cod.get().strip())
        if not cod:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione um cliente na tabela.")
            return

        if not messagebox.askyesno("Confirmar", f"Zerar recebimentos do cliente {cod} nessa programaÃ§Ã£o?"):
            return

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("DELETE FROM recebimentos WHERE codigo_programacao=? AND cod_cliente=?", (prog, cod))

            self._clear_form_recebimento()
            self.carregar_clientes_e_recebimentos()
            self.set_status(f"STATUS: Recebimentos zerados para cliente {cod}.")
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao zerar recebimentos: {str(e)}")

    def inserir_cliente_manual(self):
        prog = self._current_prog
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue uma programaÃ§Ã£o primeiro.")
            return
        if self._warn_if_fechada():
            return

        cod = upper(simple_input(
            "Cliente Manual", "Digite o CÃ“DIGO do cliente:",
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
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    SELECT COUNT(1)
                    FROM programacao_itens
                    WHERE codigo_programacao=?
                      AND cod_cliente=?
                """, (prog, cod))
                exists = int((cur.fetchone() or [0])[0] or 0)

                if not exists:
                    cur.execute("""
                        INSERT INTO programacao_itens
                            (codigo_programacao, cod_cliente, nome_cliente, qnt_caixas, kg, preco, endereco, vendedor, pedido)
                        VALUES (?, ?, ?, 0, 0, 0, '', '', 'MANUAL')
                    """, (prog, cod, nome))
                else:
                    cur.execute("""
                        UPDATE programacao_itens
                           SET nome_cliente=?
                         WHERE codigo_programacao=?
                           AND cod_cliente=?
                    """, (nome, prog, cod))

            self._safe_set_entry(self.ent_cod, cod, readonly_back=True)
            self._safe_set_entry(self.ent_nome, nome, readonly_back=True)
            self._selected_cliente = {"cod_cliente": cod, "nome_cliente": nome}

            messagebox.showinfo("OK", "Cliente inserido/atualizado na programaÃ§Ã£o (manual). Agora pode lanÃ§ar o recebimento.")
            self.carregar_clientes_e_recebimentos()

            try:
                self.ent_valor.focus_set()
                self.ent_valor.selection_range(0, "end")
            except Exception:
                logging.debug("Falha ignorada")

        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao inserir cliente: {str(e)}")

    # =========================
    # ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ IMPRESSÃƒÆ’Ã†â€™O PDF (A4) ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬ alinhado e com assinaturas
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

            # total por cliente + forma mais recente (pega a do Ãºltimo registro)
            # OBS: aqui prioriza â€œforma do registro mais recente, e soma valores.
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Carregue uma programaÃ§Ã£o primeiro.")
            return

        header, rows, total_geral = self._get_dados_para_impressao(prog)
        if not rows:
            messagebox.showwarning("ImpressÃ£o", "NÃ£o hÃ¡ clientes pagantes (valor > 0) para imprimir.")
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
            # ReportLab (jÃ¡ instalado no seu ambiente)
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
            c.drawString(left, y, f"RECIBO / RELATÃ“RIO DE RECEBIMENTOS - PROGRAMAÃ‡ÃƒO {header['prog']}")
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
            draw_kv(col2_x, y, "VeÃ­culo", header["veiculo"])
            y -= line_h

            draw_kv(col1_x, y, "Equipe", header["equipe_nomes"])
            draw_kv(col2_x, y, "NF", header["nf"])
            y -= line_h

            draw_kv(col1_x, y, "SaÃ­da", f"{header['data_saida']} {header['hora_saida']}".strip())
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

            # CabeÃ§alho da tabela
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

            # funÃ§Ã£o segura p/ quebrar texto
            def clip_text(s, max_chars):
                s = (s or "")
                return s if len(s) <= max_chars else s[:max_chars - 1] + "â€¦"

            # desenha header
            c.setFont("Helvetica-Bold", 9)
            c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
            c.drawString(x_cod + 2, y - row_h + 3, "CÃ“D CLIENTE")
            c.drawString(x_nome + 2, y - row_h + 3, "NOME CLIENTE")
            c.drawString(x_forma + 2, y - row_h + 3, "FORMA PGTO")
            c.drawRightString(x_valor + col_valor - 2, y - row_h + 3, "VALOR (R$)")
            y -= row_h

            c.setFont("Helvetica", 9)

            # desenha linhas
            for cod, nome, valor, forma in rows:
                # quebra de pÃ¡gina se necessÃ¡rio
                if y < bottom + 55 * mm:
                    c.showPage()
                    width, height = A4
                    y = height - top
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(left, y, f"RECIBO / RELATÃ“RIO DE RECEBIMENTOS - PROGRAMAÃ‡ÃƒO {header['prog']}")
                    y -= 10 * mm

                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(left, y, "RECEBIMENTOS (CONTINUAÃ‡ÃƒO)")
                    y -= 6 * mm

                    c.setFont("Helvetica-Bold", 9)
                    c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
                    c.drawString(x_cod + 2, y - row_h + 3, "CÃ“D CLIENTE")
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
            c.drawString(left, y, "ASSINATURAS / CONFERÃŠNCIA")
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
            assinatura_block(x2, y, "SETOR DE CONFERÃŠNCIA")

            # rodapÃ©
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
# 7.0 CONFIGURAÃ‡ÃƒO DE LOGGING (DespesasPage) - mais seguro
# =========================================================
def setup_despesas_logger():
    """
    Logger com rotaÃ§Ã£o de arquivo para nÃ£o crescer infinito.
    NÃ£o altera regra do sistema, sÃ³ evita travar com log grande.
    """
    logger = logging.getLogger(__name__ + ".DespesasPage")
    if not logger.handlers:
        try:
            from logging.handlers import RotatingFileHandler
            handler = RotatingFileHandler(
                "despesas_audit.log",
                maxBytes=2_000_000,   # ~2MB
                backupCount=5,
                encoding="utf-8"
            )
        except Exception:
            handler = logging.FileHandler("despesas_audit.log", encoding="utf-8")

        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.propagate = False
    return logger


# =========================================================
# 7.1 DESPESAS PAGE (PARTE 1/3) âœ… layout mais robusto + scroll seguro
# =========================================================
class DespesasPage(PageBase):
    def __init__(self, parent, app):
        super().__init__(parent, app, "Despesas")

        self.selected_despesa_id = None
        self.logger = setup_despesas_logger()
        self._current_programacao = None

        # =========================================================
        # âœ… Remove cinza / body branco (sem mexer no PageBase)
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
            text="🔄 ATUALIZAR",
            style="Ghost.TButton",
            command=self._refresh_all,
            width=10
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
            text="Motorista: ? | Veiculo: ? | Equipe: ?",
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

        tab_cedulas = ttk.Frame(sections_nb, style="Card.TFrame", padding=10)
        sections_nb.add(tab_cedulas, text="Contagem de Cedulas")
        tab_cedulas.grid_columnconfigure(0, weight=1)
        tab_cedulas.grid_columnconfigure(1, weight=1)
        tab_cedulas.grid_columnconfigure(2, weight=1)

        tab_resumos = ttk.Frame(sections_nb, style="Card.TFrame", padding=10)
        sections_nb.add(tab_resumos, text="Resumos")
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
            text="Calcular Saldo AutomÃ¡tico",
            style="Ghost.TButton",
            width=24,
            command=self._calcular_saldo_auto
        ).pack(anchor="w")

        row_nf = 1
        self.ent_nf_numero = self._create_field(tab_nf, "NÂº NOTA FISCAL:", row_nf, 14); row_nf += 1
        self.ent_nf_kg = self._create_field(tab_nf, "KG NOTA FISCAL:", row_nf, 14); row_nf += 1
        self.ent_nf_preco = self._create_field(tab_nf, "PRECO NF (R$/KG):", row_nf, 14); row_nf += 1
        self.ent_nf_caixas = self._create_field(tab_nf, "CAIXAS:", row_nf, 14); row_nf += 1
        self.ent_nf_kg_carregado = self._create_field(tab_nf, "KG CARREGADO:", row_nf, 14); row_nf += 1
        self.ent_nf_kg_vendido = self._create_field(tab_nf, "KG VENDIDO:", row_nf, 14); row_nf += 1
        self.ent_nf_saldo = self._create_field(tab_nf, "SALDO (KG):", row_nf, 14)
        self.ent_nf_media_carregada = self._create_field(tab_nf, "MEDIA CARREGADA:", row_nf, 14); row_nf += 1
        self.ent_nf_caixa_final = self._create_field(tab_nf, "CAIXA FINAL:", row_nf, 14)
        try:
            self.ent_nf_caixas.configure(state="readonly")
            self.ent_nf_kg_carregado.configure(state="readonly")
            self.ent_nf_kg_vendido.configure(state="readonly")
        except Exception:
            logging.debug("Falha ignorada")

        ttk.Button(
            tab_nf,
            text="Sincronizar com App",
            style="Warn.TButton",
            width=24,
            command=self._sincronizar_com_app
        ).grid(row=row_nf + 1, column=0, columnspan=2, sticky="w", pady=(15, 0))

        nf_resumo = ttk.LabelFrame(tab_nf, text="Resumo Compra x Venda", padding=12)
        nf_resumo.grid(row=1, column=3, rowspan=max(row_nf + 2, 12), sticky="nsew", padx=(20, 0))
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
            text="Calcular KM/MÃ©dia",
            style="Ghost.TButton",
            width=20,
            command=self._calcular_km_media
        ).grid(row=0, column=0, sticky="w")

        row_km = 1
        self.ent_km_inicial = self._create_field(km_form, "KM INICIAL:", row_km, 12); row_km += 1
        self.ent_km_final = self._create_field(km_form, "KM FINAL:", row_km, 12); row_km += 1
        self.ent_litros = self._create_field(km_form, "LITROS:", row_km, 12); row_km += 1
        self.ent_km_rodado = self._create_field(km_form, "KM RODADO:", row_km, 12); row_km += 1
        self.ent_media = self._create_field(km_form, "MÃ‰DIA (KM/L):", row_km, 12); row_km += 1
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

        ttk.Label(km_form, text="OBSERVAÃ‡ÃƒO:", style="CardLabel.TLabel").grid(row=row_km, column=0, sticky="nw", pady=(8, 2))
        self.txt_rota_obs = tk.Text(km_form, height=4, width=28, font=("Segoe UI", 9))
        self.txt_rota_obs.grid(row=row_km, column=1, sticky="w", pady=(8, 2), padx=(10, 0))

        km_side = ttk.Frame(tab_rota, style="Card.TFrame")
        km_side.grid(row=0, column=1, sticky="nsew")
        km_side.grid_columnconfigure(0, weight=1)
        km_side.grid_rowconfigure(1, weight=1)

        ttk.Label(km_side, text="MELHORES ÃšLTIMAS MÃ‰DIAS", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")

        chart_wrap = ttk.Frame(km_side, style="Card.TFrame")
        chart_wrap.grid(row=1, column=0, sticky="nsew", pady=(6, 10))
        chart_wrap.grid_columnconfigure(0, weight=1)

        self.canvas_km_pie = tk.Canvas(chart_wrap, width=300, height=180, bg="white", highlightthickness=1, highlightbackground="#E5E7EB")
        self.canvas_km_pie.grid(row=0, column=0, sticky="w")

        ttk.Label(km_side, text="ÃšLTIMOS KM POR VEÃCULO", style="CardTitle.TLabel").grid(row=2, column=0, sticky="w", pady=(4, 4))
        self.tree_km_veiculos = ttk.Treeview(
            km_side,
            columns=["VEICULO", "KM", "MÃ‰DIA", "DATA"],
            show="headings",
            height=6
        )
        for col, w in [("VEICULO", 130), ("KM", 80), ("MÃ‰DIA", 80), ("DATA", 110)]:
            self.tree_km_veiculos.heading(col, text=col)
            self.tree_km_veiculos.column(col, width=w, anchor="center" if col != "VEICULO" else "w")
        self.tree_km_veiculos.grid(row=3, column=0, sticky="ew")

        # =========================================================
        # ---- ABA 3: CONTAGEM DE CÃ‰DULAS
        # =========================================================

        calc_ced_frame = ttk.Frame(tab_cedulas)
        calc_ced_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        ttk.Label(calc_ced_frame, text="Valor Total:", style="CardLabel.TLabel").pack(side="left", padx=(0, 5))
        self.ent_total_dinheiro = ttk.Entry(calc_ced_frame, style="Field.TEntry", width=15)
        self.ent_total_dinheiro.pack(side="left", padx=5)
        self._bind_money_entry(self.ent_total_dinheiro)
        self._bind_focus_scroll(self.ent_total_dinheiro)
        ttk.Button(calc_ced_frame, text="🧮 Distribuir", style="Ghost.TButton",
                   command=self._distribuir_cedulas, width=10).pack(side="left", padx=5)

        header_ced = ttk.Frame(tab_cedulas)
        header_ced.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 3))

        ttk.Label(header_ced, text="QTD", width=8, style="CardLabel.TLabel", anchor="center", font=("Segoe UI", 8)).grid(row=0, column=0, sticky="ew")
        ttk.Label(header_ced, text="CÃ‰DULA", width=15, style="CardLabel.TLabel", anchor="center").grid(row=0, column=1)
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
            ent = ttk.Entry(tab_cedulas, width=12, style="Field.TEntry", justify="center")
            ent.grid(row=i, column=0, sticky="ew", pady=1, padx=4)
            ent.bind("<KeyRelease>", lambda e: self._calc_valor_dinheiro())
            self._bind_focus_scroll(ent)

            lbl_ced = ttk.Label(
                tab_cedulas,
                text=f"R$ {ced:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                style="CardLabel.TLabel",
                anchor="center",
                foreground=ced_colors.get(ced, "#111827")
            )
            lbl_ced.grid(row=i, column=1, sticky="ew", pady=1, padx=4)

            lbl_total = ttk.Label(tab_cedulas, text="R$ 0,00",
                                  font=("Segoe UI", 8, "bold"), anchor="center",
                                  foreground=ced_colors.get(ced, "#111827"))
            lbl_total.grid(row=i, column=2, sticky="ew", pady=1, padx=4)

            self.ced_entries[ced] = ent
            self.ced_totals[ced] = lbl_total

        ttk.Separator(tab_cedulas, orient="horizontal")\
            .grid(row=9, column=0, columnspan=3, sticky="ew", pady=6)

        ttk.Label(tab_cedulas, text="TOTAL DINHEIRO:", font=("Segoe UI", 9, "bold"))\
            .grid(row=10, column=0, columnspan=2, sticky="e", pady=3)

        self.lbl_valor_dinheiro = ttk.Label(tab_cedulas, text="R$ 0,00", font=("Segoe UI", 11, "bold"))
        self.lbl_valor_dinheiro.grid(row=10, column=2, sticky="w", pady=3, padx=4)

        # forÃ§a atualizar scroll no primeiro render
        # =========================================================
        # ? RESUMO + BOT?ES (como voc? j? tinha)
        # =========================================================
        # =========================================================
        # RESUMO + BOTOES
        # =========================================================
        # =========================================================
        # RESUMO + BOTOES
        # =========================================================
        resumo_frame = ttk.Frame(tab_resumos, style="Card.TFrame", padding=10)
        resumo_frame.grid(row=0, column=0, sticky="nsew")
        resumo_frame.grid_columnconfigure(0, weight=1)
        resumo_frame.grid_rowconfigure(2, weight=1)

        ttk.Label(
            resumo_frame,
            text="RESUMO FINANCEIRO (PADRÃƒO ÃšNICO)",
            style="CardTitle.TLabel",
            font=("Segoe UI", 12, "bold")
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        kpi_row = ttk.Frame(resumo_frame, style="Card.TFrame")
        kpi_row.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for i in range(4):
            kpi_row.grid_columnconfigure(i, weight=1, uniform="kpi")

        kpi_receb = ttk.LabelFrame(kpi_row, text=" RECEBIMENTOS ", padding=8)
        kpi_receb.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.lbl_kpi_receb_total = ttk.Label(kpi_receb, text="R$ 0,00", font=("Segoe UI", 12, "bold"), foreground="#1D4ED8")
        self.lbl_kpi_receb_total.grid(row=0, column=0, sticky="w")

        kpi_desp = ttk.LabelFrame(kpi_row, text=" DESPESAS ", padding=8)
        kpi_desp.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self.lbl_kpi_desp_total = ttk.Label(kpi_desp, text="R$ 0,00", font=("Segoe UI", 12, "bold"), foreground="#C62828")
        self.lbl_kpi_desp_total.grid(row=0, column=0, sticky="w")

        kpi_ced = ttk.LabelFrame(kpi_row, text=" CÃ‰DULAS ", padding=8)
        kpi_ced.grid(row=0, column=2, sticky="ew", padx=(0, 6))
        self.lbl_kpi_cedulas_total = ttk.Label(kpi_ced, text="R$ 0,00", font=("Segoe UI", 12, "bold"), foreground="#7C3AED")
        self.lbl_kpi_cedulas_total.grid(row=0, column=0, sticky="w")

        kpi_res = ttk.LabelFrame(kpi_row, text=" RESULTADO LÃQUIDO ", padding=8)
        kpi_res.grid(row=0, column=3, sticky="ew")
        self.lbl_kpi_resultado_liquido = ttk.Label(kpi_res, text="R$ 0,00", font=("Segoe UI", 12, "bold"), foreground="#2E7D32")
        self.lbl_kpi_resultado_liquido.grid(row=0, column=0, sticky="w")

        details_wrap = ttk.Frame(resumo_frame, style="Card.TFrame")
        details_wrap.grid(row=2, column=0, sticky="nsew")
        for i in range(3):
            details_wrap.grid_columnconfigure(i, weight=1, uniform="res_cards")

        entrada_frame = ttk.LabelFrame(details_wrap, text=" ENTRADAS ", padding=8)
        entrada_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        entrada_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(entrada_frame, text="Recebimentos Total:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        self.lbl_receb_total = ttk.Label(entrada_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_receb_total.grid(row=0, column=1, sticky="e", pady=2)

        ttk.Label(entrada_frame, text="Adiantamento p/ Rota:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        self.ent_adiantamento = ttk.Entry(entrada_frame, style="Field.TEntry", width=12)
        self.ent_adiantamento.grid(row=1, column=1, sticky="e", pady=2)
        self.ent_adiantamento.bind("<KeyRelease>", lambda e: self._update_resumo_financeiro())
        self._bind_money_entry(self.ent_adiantamento)

        ttk.Label(entrada_frame, text="Total Entradas:", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.lbl_total_entradas = ttk.Label(entrada_frame, text="R$ 0,00", font=("Segoe UI", 11, "bold"), foreground="#2E7D32")
        self.lbl_total_entradas.grid(row=2, column=1, sticky="e", pady=(8, 2))

        saida_frame = ttk.LabelFrame(details_wrap, text=" SAÃDAS ", padding=8)
        saida_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 6))
        saida_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(saida_frame, text="Despesas Total:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        self.lbl_desp_total = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_desp_total.grid(row=0, column=1, sticky="e", pady=2)

        ttk.Label(saida_frame, text="Contagem CÃ©dulas:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        self.lbl_cedulas_total = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_cedulas_total.grid(row=1, column=1, sticky="e", pady=2)

        ttk.Label(saida_frame, text="Total SaÃ­das:", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.lbl_total_saidas = ttk.Label(saida_frame, text="R$ 0,00", font=("Segoe UI", 11, "bold"), foreground="#C62828")
        self.lbl_total_saidas.grid(row=2, column=1, sticky="e", pady=(8, 2))

        resultado_frame = ttk.LabelFrame(details_wrap, text=" RESULTADOS ", padding=8)
        resultado_frame.grid(row=0, column=2, sticky="nsew")
        resultado_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(resultado_frame, text="Valor p/ Caixa:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        self.lbl_valor_final_caixa = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_valor_final_caixa.grid(row=0, column=1, sticky="e", pady=2)

        ttk.Label(resultado_frame, text="DiferenÃ§a (Caixa - CÃ©d):", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        self.lbl_diferenca = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 10, "bold"))
        self.lbl_diferenca.grid(row=1, column=1, sticky="e", pady=2)

        ttk.Label(resultado_frame, text="Resultado LÃ­quido:", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.lbl_resultado_liquido = ttk.Label(resultado_frame, text="R$ 0,00", font=("Segoe UI", 12, "bold"))
        self.lbl_resultado_liquido.grid(row=2, column=1, sticky="e", pady=(8, 2))

        botoes_frame = ttk.Frame(page_card, style="Card.TFrame")
        botoes_frame.grid(row=4, column=0, sticky="ew")

        for i in range(6):
            botoes_frame.grid_columnconfigure(i, weight=1)

        botoes = [
            ("⬅ VOLTAR", "Ghost.TButton", self._voltar_recebimentos),
            ("➕ ADICIONAR DESPESA", "Warn.TButton", self._open_registrar_rapido),
            ("🖨 IMPRIMIR PDF", "Ghost.TButton", self.imprimir_resumo),
            ("💾 SALVAR", "Primary.TButton", self.salvar_tudo),
            ("✏️ EDITAR", "Warn.TButton", self._editar_linha_selecionada),
            ("🏁 FINALIZAR", "Danger.TButton", self.finalizar_prestacao_despesas),
        ]

        for i, (texto, estilo, comando) in enumerate(botoes):
            btn = ttk.Button(botoes_frame, text=texto, style=estilo, command=comando, width=14)
            btn.grid(row=0, column=i, sticky="ew", padx=2, ipady=4)

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
        elif "CAIXA" in lbl or "QTD" in lbl or "NÂº" in lbl or "NOTA FISCAL" in lbl:
            bind_entry_smart(ent, "int")
        elif "KM" in lbl or "KG" in lbl or "LITRO" in lbl or "MÃ‰DIA" in lbl or "MEDIA" in lbl or "CUSTO" in lbl:
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
            elif "/" in data_str and len(data_str) >= 10:
                d, m, y = data_str[:10].split("/")
            else:
                return None
            hh, mm = "00", "00"
            if hora_str and ":" in hora_str:
                hh, mm = hora_str.split(":")[:2]
            return datetime(int(y), int(m), int(d), int(hh), int(mm))
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

    # ---- helpers seguros (nÃ£o mudam regra, sÃ³ evitam bug)
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
            canvas.create_text(150, 90, text="Sem dados de mÃ©dia para exibir", fill="#6B7280", font=("Segoe UI", 9, "bold"))
            return

        colors = ["#1D4ED8", "#16A34A", "#EA580C", "#7C3AED", "#0891B2", "#BE123C"]
        total = sum(max(0.0, safe_float(v, 0.0)) for _, v in items)
        if total <= 0:
            canvas.create_text(150, 90, text="Sem dados de mÃ©dia para exibir", fill="#6B7280", font=("Segoe UI", 9, "bold"))
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
            elif col in {"OBSERVA??O", "OBSERVACAO"}:
                self.tree_desp.column(col, width=240, minwidth=160)
            elif col in {"DESCRI??O", "DESCRICAO"}:
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
            if prog not in vals:
                vals = [prog] + vals
                self.cb_prog["values"] = vals
            self.cb_prog.set(prog)
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

        self._load_despesas(prog)
        self._load_programacao_extras(prog)
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

        # CÃ©dulas
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

        with get_db() as conn:
            cur = conn.cursor()

            cur.execute("SELECT SUM(valor) FROM despesas WHERE codigo_programacao=?", (prog,))
            row = cur.fetchone()
            total_desp = float(row[0]) if row and row[0] else 0.0

            total_receb = 0.0
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
    # 7.3 EVENTOS / FILTROS / ORDENAÃ‡ÃƒO
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
        # estado simples (mantÃ©m comportamento previsÃ­vel)
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

        # seta setinha no cabeÃ§alho (igual seu padrÃ£o em outras telas)
        arrow = " Ã¢â€ â€œ" if reverse else " Ã¢â€ â€˜"
        for c in self.tree_desp["columns"]:
            txt = self.tree_desp.heading(c)["text"]
            if txt.endswith(" Ã¢â€ â€˜") or txt.endswith(" Ã¢â€ â€œ"):
                txt = txt[:-2]
            self.tree_desp.heading(c, text=txt)

        txt = self.tree_desp.heading(col)["text"]
        if txt.endswith(" Ã¢â€ â€˜") or txt.endswith(" Ã¢â€ â€œ"):
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
    # 7.4 CÃƒÆ’Ã‚LCULOS AUTOMÃƒÆ’Ã‚TICOS / SINCRONIZAÃƒÆ’Ã¢â‚¬Â¡ÃƒÆ’Ã†â€™O SIMULADA
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
                            f"SELECT COALESCE({c}, 0) FROM programacoes WHERE codigo_programacao=? LIMIT 1",
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

        compra_total = nf_kg * nf_preco
        preco_medio_venda = self._calc_preco_medio_venda(self._current_programacao or "")
        receita_estimada = kg_vendido * preco_medio_venda

        despesas_rota = 0.0
        try:
            prog = self._current_programacao or ""
            if prog:
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
                f"STATUS: KM rodado: {km_rodado:.1f} km | MÃ©dia: {media:.1f} km/l | "
                f"Custo/km: {custo_km:.2f}"
            )
            self._refresh_km_insights()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao calcular KM/mÃ©dia: {str(e)}")

    def _distribuir_cedulas(self):
        try:
            valor_total = self._parse_money(self.ent_total_dinheiro.get())
            if valor_total <= 0:
                messagebox.showwarning("AtenÃ§Ã£o", "Digite um valor total vÃ¡lido.")
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

            # arredondamento: se sobrar qualquer centavo, ajusta na menor cÃ©dula
            if restante > 0:
                atual = safe_int(self.ced_entries[2].get(), 0)
                self.ced_entries[2].delete(0, "end")
                self.ced_entries[2].insert(0, str(atual + 1))

            self._calc_valor_dinheiro()
            self.set_status(f"STATUS: Valor {self._fmt_money(valor_total)} distribuÃ­do automaticamente")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao distribuir cÃ©dulas: {str(e)}")

    def _sincronizar_com_app(self):
        if not self._current_programacao:
            messagebox.showwarning("SincronizaÃ§Ã£o", "Selecione a programaÃ§Ã£o primeiro.")
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
                raise SyncError("A API nÃ£o retornou os dados da rota.")
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

            self.logger.info("SincronizaÃ§Ã£o com App Mobile concluÃ­da para %s", prog)
            messagebox.showinfo("SincronizaÃ§Ã£o", "Dados sincronizados com sucesso.")
        except SyncError as exc:
            self.logger.error("Falha ao sincronizar com o App: %s", exc)
            messagebox.showerror("SincronizaÃ§Ã£o", f"Falha ao sincronizar: {exc}")
        except Exception as exc:
            logging.exception("Erro inesperado na sincronizaÃ§Ã£o", exc_info=exc)
            messagebox.showerror("SincronizaÃ§Ã£o", f"Erro inesperado: {exc}")

    def _prompt_sync_credentials(self, default_code: str):
        title = "SincronizaÃ§Ã£o App Mobile"
        codigo = simpledialog.askstring(
            title,
            "CÃ³digo do motorista utilizado no App Mobile:",
            initialvalue=(default_code or ""),
            parent=self,
        )
        if codigo is None:
            return None
        codigo = codigo.strip()
        if not codigo:
            messagebox.showwarning("SincronizaÃ§Ã£o", "CÃ³digo do motorista Ã© obrigatÃ³rio.")
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
            messagebox.showwarning("SincronizaÃ§Ã£o", "Senha do motorista Ã© obrigatÃ³ria.")
            return None

        return codigo, senha

    def _fetch_motorista_codigo_hint(self, prog: str) -> str:
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute(
                    "SELECT motorista_id, motorista FROM programacoes WHERE codigo_programacao=? LIMIT 1",
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
            raise SyncError("ROTA_SERVER_URL nÃ£o configurada.")
        resposta = _call_api("POST", "auth/motorista/login", payload={"codigo": codigo, "senha": senha})
        token = resposta.get("token")
        if not token:
            raise SyncError("NÃ£o foi possÃ­vel obter token da API de sincronizaÃ§Ã£o.")
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

        # Se a API nao trouxer saldo, calcula com base no carregado.
        if nf_saldo in (None, ""):
            try:
                nf_saldo = round(max(safe_float(nf_kg, 0.0) - safe_float(nf_kg_carregado, 0.0), 0.0), 2)
            except Exception:
                nf_saldo = 0

        valores_por_campo = {
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
            # espelha dados de carregamento bruto quando vierem da API
            "kg_carregado": _pick("kg_carregado", "nf_kg_carregado"),
            "caixas_carregadas": _pick("caixas_carregadas", "nf_caixas"),
        }

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
        ctrl_cols = [
            "mortalidade_aves",
            "media_aplicada",
            "peso_previsto",
            "valor_recebido",
            "forma_recebimento",
            "obs_recebimento",
            "status_pedido",
            "alteracao_tipo",
            "alteracao_detalhe",
            "pedido",
            "caixas_atual",
            "preco_atual",
            "alterado_em",
            "alterado_por",
        ]
        cols_sql = ", ".join(ctrl_cols)
        placeholders = ", ".join("?" for _ in ctrl_cols)
        update_sql = ", ".join(f"{col}=excluded.{col}" for col in ctrl_cols)
        sql = f"""
            INSERT INTO programacao_itens_controle
                (codigo_programacao, cod_cliente, {cols_sql}, updated_at)
            VALUES (?, ?, {placeholders}, datetime('now'))
            ON CONFLICT(codigo_programacao, cod_cliente)
            DO UPDATE SET
                {update_sql},
                updated_at=datetime('now')
        """
        for cliente in clientes:
            cod_cliente = (cliente.get("cod_cliente") or "").strip()
            if not cod_cliente:
                continue
            valores = [cliente.get(col) for col in ctrl_cols]
            cur.execute(sql, (prog, cod_cliente, *valores))

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
        for valor, entry in campos:
            if not entry:
                continue
            self._set_ent(entry, valor)
        self._refresh_nf_trade_summary()

    # =========================================================
    # 7.5 PageBase (refresh / on_show / status)
    # =========================================================
    def refresh_comboboxes(self):
        with get_db() as conn:
            cur = conn.cursor()
            try:
                cur.execute("""
                    SELECT codigo_programacao, data_criacao
                    FROM programacoes
                    WHERE (
                        status='ATIVA'
                        OR status='EM_ROTA'
                        OR status='INICIADA'
                        OR status='CARREGADA'
                        OR status='FINALIZADA'
                    )
                    ORDER BY id DESC
                    LIMIT 300
                """)
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
        self.logger.info("PÃ¡gina Despesas carregada")

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

    # --------- REGRAS DE SEGURANÃ‡A (nÃ£o altera sistema; sÃ³ bloqueia quando necessÃ¡rio)
    def _get_prestacao_status(self, prog: str) -> str:
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("SELECT COALESCE(prestacao_status,'') FROM programacoes WHERE codigo_programacao=? LIMIT 1", (prog,))
                row = cur.fetchone()
            return upper(row[0] if row else "")
        except Exception:
            return ""

    def _can_edit_current_prog(self) -> bool:
        prog = self._current_programacao
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione a ProgramaÃ§Ã£o primeiro.")
            return False

        status = self._get_prestacao_status(prog)
        if status == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"Esta programaÃ§Ã£o ({prog}) estÃ¡ com a prestaÃ§Ã£o FECHADA.\n\n"
                "Por seguranÃ§a, nÃ£o Ã© permitido registrar/editar/excluir despesas."
            )
            return False
        return True

    def finalizar_prestacao_despesas(self):
        prog = self._current_programacao
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione a ProgramaÃ§Ã£o primeiro.")
            return

        status = self._get_prestacao_status(prog)
        if status == "FECHADA":
            messagebox.showinfo("PrestaÃ§Ã£o", "Esta prestaÃ§Ã£o jÃ¡ estÃ¡ FECHADA.")
            return

        if not messagebox.askyesno(
            "Confirmar",
            f"Finalizar prestaÃ§Ã£o de contas da rota {prog}?\n\n"
            "A programaÃ§Ã£o serÃ¡ finalizada."
        ):
            return

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("""
                    UPDATE programacoes
                    SET prestacao_status='FECHADA', status='FINALIZADA'
                    WHERE codigo_programacao=?
                """, (prog,))

            messagebox.showinfo("OK", f"PrestaÃ§Ã£o finalizada: {prog}")
            try:
                if messagebox.askyesno(
                    "Imprimir prestação",
                    f"Deseja imprimir agora a prestação de contas da rota {prog}?\n\n"
                    "Será gerada a folha completa com recebimentos e despesas."
                ):
                    # Garante contexto da programação ao abrir a impressão.
                    self._current_programacao = prog
                    self.imprimir_resumo()
            except Exception:
                logging.exception("Falha ao imprimir prestação após finalizar")
            self.refresh_comboboxes()
            if hasattr(self.app, "refresh_programacao_comboboxes"):
                try:
                    self.app.refresh_programacao_comboboxes()
                except Exception:
                    logging.debug("Falha ignorada")
            # Ap?s finalizar, limpa a tela para iniciar nova programa??o.
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
            self.set_status("STATUS: Presta??o finalizada. Selecione uma nova programa??o para continuar.")
        except Exception as e:
            messagebox.showerror("ERRO", f"Erro ao finalizar prestaÃ§Ã£o: {str(e)}")

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
        win.title("Registrar Despesa (RÃ¡pido)")
        win.geometry("520x360")
        win.grab_set()
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=20)
        frm.pack(fill="both", expand=True)
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="NOVA DESPESA", font=("Segoe UI", 14, "bold"))\
            .grid(row=0, column=0, columnspan=2, pady=(0, 16), sticky="w")

        ttk.Label(frm, text="DescriÃ§Ã£o:", font=("Segoe UI", 10, "bold"))\
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

        ttk.Label(frm, text="ObservaÃ§Ã£o:", font=("Segoe UI", 10, "bold"))\
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
                messagebox.showwarning("ATENÃ‡ÃƒO", "Informe a descriÃ§Ã£o.")
                return
            if val <= 0:
                messagebox.showwarning("ATENÃ‡ÃƒO", "O valor deve ser maior que zero.")
                return
            if not categoria:
                categoria = "OUTROS"

            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        INSERT INTO despesas (codigo_programacao, descricao, valor, categoria, observacao, data_registro)
                        VALUES (?, ?, ?, ?, ?, datetime('now'))
                    """, (prog, desc, float(val), categoria, obs))

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

        ttk.Button(btn_frame, text="💾 SALVAR", style="Primary.TButton", command=salvar)\
            .grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btn_frame, text="✖ CANCELAR", style="Ghost.TButton", command=win.destroy)\
            .grid(row=0, column=1, sticky="ew", padx=(6, 0))

        win.bind("<Return>", lambda e: salvar())
        ent_desc.focus_set()

    def _editar_linha_selecionada(self):
        if not self._can_edit_current_prog():
            return

        prog = self._current_programacao

        if not self.selected_despesa_id:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma linha na tabela.")
            return

        sel = self.tree_desp.selection()
        if not sel:
            return

        vals = self.tree_desp.item(sel[0], "values")
        if len(vals) < 6:
            messagebox.showerror("ERRO", "Dados incompletos na linha selecionada.")
            return

        did, desc, val_str, data_reg, categoria, observacao = vals

        # âœ… valida que essa despesa pertence Ã  programaÃ§Ã£o atual
        info = self._get_despesa_info(did)
        if not info:
            messagebox.showerror("ERRO", "NÃ£o encontrei essa despesa no banco.")
            return
        _, codigo_prog_db, _, _, _ = info
        if upper(codigo_prog_db) != upper(prog):
            messagebox.showwarning(
                "ATENÃ‡ÃƒO",
                "Por seguranÃ§a, nÃ£o Ã© permitido editar uma despesa que nÃ£o pertence Ã  programaÃ§Ã£o atual."
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

        ttk.Label(frm, text="DescriÃ§Ã£o:", font=("Segoe UI", 10, "bold"))\
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

        ttk.Label(frm, text="ObservaÃ§Ã£o:", font=("Segoe UI", 10, "bold"))\
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
                messagebox.showwarning("ATENÃ‡ÃƒO", "Informe a descriÃ§Ã£o.")
                return
            if nval <= 0:
                messagebox.showwarning("ATENÃ‡ÃƒO", "O valor deve ser maior que zero.")
                return
            if not ncategoria:
                ncategoria = "OUTROS"

            # âœ… valida novamente (seguranÃ§a)
            info2 = self._get_despesa_info(did)
            if not info2 or upper(info2[1]) != upper(prog):
                messagebox.showwarning("ATENÃ‡ÃƒO", "Despesa nÃ£o pertence Ã  programaÃ§Ã£o atual (bloqueado).")
                return

            try:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        UPDATE despesas
                        SET descricao=?, valor=?, categoria=?, observacao=?
                        WHERE id=?
                    """, (ndesc, float(nval), ncategoria, nobs, did))

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

        ttk.Button(btn_frame, text="💾 SALVAR", style="Primary.TButton", command=salvar)\
            .grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(btn_frame, text="✖ CANCELAR", style="Ghost.TButton", command=win.destroy)\
            .grid(row=0, column=1, sticky="ew", padx=(6, 0))

        win.bind("<Return>", lambda e: salvar())
        ent_desc.focus_set()

    def _excluir_linha_selecionada(self):
        if not self._can_edit_current_prog():
            return

        prog = self._current_programacao
        if not self.selected_despesa_id:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma linha na tabela.")
            return

        did = self.selected_despesa_id

        # âœ… pega dados reais do banco para confirmar e validar pertencimento
        info = self._get_despesa_info(did)
        if not info:
            messagebox.showerror("ERRO", "NÃ£o encontrei essa despesa no banco.")
            return

        _, codigo_prog_db, desc_db, val_db, cat_db = info
        if upper(codigo_prog_db) != upper(prog):
            messagebox.showwarning(
                "ATENÃ‡ÃƒO",
                "Por seguranÃ§a, nÃ£o Ã© permitido excluir uma despesa que nÃ£o pertence Ã  programaÃ§Ã£o atual."
            )
            return

        val_fmt = self._fmt_money(val_db) if hasattr(self, "_fmt_money") else f"R$ {float(val_db or 0):.2f}"
        resposta = messagebox.askyesno(
            "CONFIRMAR EXCLUSÃƒO",
            "Deseja realmente excluir esta despesa?\n\n"
            f"ID: {did}\n"
            f"DescriÃ§Ã£o: {desc_db}\n"
            f"Valor: {val_fmt}\n"
            f"Categoria: {cat_db}\n\n"
            "Esta aÃ§Ã£o nÃ£o pode ser desfeita."
        )
        if not resposta:
            return

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("DELETE FROM despesas WHERE id=? AND codigo_programacao=?", (did, prog))

            self.logger.warning(f"Despesa excluÃ­da: {prog} - ID {did} - {desc_db} - {val_fmt}")
            self.selected_despesa_id = None
            self._refresh_all()
            self.set_status(f"STATUS: Despesa ID {did} excluÃ­da com sucesso")

        except Exception as e:
            self.logger.error(f"Erro ao excluir despesa: {str(e)}")
            messagebox.showerror("ERRO", f"Erro ao excluir despesa: {str(e)}")

    # =========================================================
    # 7.7 RELATÃ“RIO EM TELA + IMPRESSÃƒO SIMULADA
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
            col_nome = table_w * 0.30
            col_forma = table_w * 0.14
            col_valor = table_w * 0.12
            col_obs = table_w * 0.20
            col_data = table_w * 0.12

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
                c.drawString(x_data + 2, y - row_h_rec + 3, "DATA")
                y -= row_h_rec
                c.setFont("Helvetica", 8)

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
                    c.drawString(x_cod + 2, y - row_h_rec + 3, str(cod_cli or "")[:10])
                    c.drawString(x_nome + 2, y - row_h_rec + 3, str(nome_cli or "")[:34])
                    c.drawString(x_forma + 2, y - row_h_rec + 3, str(forma_rec or "")[:14])
                    c.drawRightString(
                        x_valor + col_valor - 2,
                        y - row_h_rec + 3,
                        f"{float(valor_rec or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    )
                    c.drawString(x_obs + 2, y - row_h_rec + 3, str(obs_rec or "")[:24])
                    c.drawString(x_data + 2, y - row_h_rec + 3, str(data_rec or "")[:16])
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
                c.rect(table_x, y - row_h + 1, table_w, row_h, stroke=1, fill=0)
                c.drawString(x_desc + 2, y - row_h + 3, str(desc or "")[:40])
                c.drawString(x_cat + 2, y - row_h + 3, str(cat or "")[:14])
                c.drawRightString(x_val + col_val - 2, y - row_h + 3, f"{float(val or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c.drawString(x_obs + 2, y - row_h + 3, str(obs or "")[:35])
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
            "IMPRESSÃƒO",
            "FunÃ§Ã£o de impressÃ£o serÃ¡ implementada na prÃ³xima versÃ£o.\n\n"
            "Por enquanto, vocÃª pode:\n"
            "1. Tirar um print screen desta tela\n"
            "2. Usar o botÃ£o 'Exportar Excel' para gerar arquivo\n"
            "3. Salvar como PDF usando Ctrl+P"
        )
        try:
            self.logger.info("Solicitada impressÃ£o de relatÃ³rio (DespesasPage)")
        except Exception:
            logging.debug("Falha ignorada")

    # =========================================================
    # 7.8 SALVAR TUDO (NF / KM / CÃ‰DULAS / ADIANTAMENTO)
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

        # âœ… aceita "10,00" ou "R$ 10,00"
        adiantamento_val = self._parse_money_local(self.ent_adiantamento.get())

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

            self.logger.info(f"Dados salvos para programaÃ§Ã£o {prog}")
            messagebox.showinfo("SUCESSO", "Todos os dados foram salvos com sucesso!")
            self.set_status(f"STATUS: Dados salvos para {prog}")
            self._refresh_km_insights()
            self._refresh_nf_trade_summary()

        except Exception as e:
            self.logger.error(f"Erro ao salvar dados: {str(e)}")
            messagebox.showerror("ERRO", f"Erro ao salvar dados: {str(e)}")

    # =========================================================
    # 7.9 EXPORTAR EXCEL (RELATÃ“RIO COMPLETO)
    # =========================================================
    def exportar_excel(self):
        # exportar pode ser permitido mesmo com FECHADA (Ã© leitura). EntÃ£o NÃƒO bloqueio.
        prog = self._current_programacao
        if not (require_pandas() and require_openpyxl()):
            return
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione a ProgramaÃ§Ã£o primeiro.")
            return

        path = filedialog.asksaveasfilename(
            title="Exportar RelatÃ³rio Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"RELATORIO_DESPESAS_{prog}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        if not path:
            return


        try:
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

                if "data_criacao" in df_prog.columns:
                    df_prog["data_criacao"] = normalize_datetime_column(df_prog["data_criacao"])

                df_despesas = pd.read_sql_query("""
                    SELECT id, descricao, valor, categoria, observacao, data_registro
                    FROM despesas
                    WHERE codigo_programacao=?
                    ORDER BY data_registro DESC
                """, conn, params=(prog,))
                if "data_registro" in df_despesas.columns:
                    df_despesas["data_registro"] = normalize_datetime_column(df_despesas["data_registro"])

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
                ["PROGRAMAÃ‡ÃƒO", prog],
                ["TOTAL RECEBIMENTOS", total_receb],
                ["TOTAL ADIANTAMENTO", adiant],
                ["TOTAL ENTRADAS", total_entradas],
                ["TOTAL DESPESAS", total_desp],
                ["TOTAL CÃ‰DULAS", total_ced],
                ["TOTAL SADAS", total_saidas],
                ["RESULTADO LQUIDO", resultado_liquido],
                ["DATA EXPORTAÃ‡ÃƒO", datetime.now().strftime("%d/%m/%Y %H:%M")]
            ], columns=["ITEM", "VALOR"])

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df_resumo.to_excel(writer, sheet_name="RESUMO", index=False)
                df_prog.to_excel(writer, sheet_name="PROGRAMAÃ‡ÃƒO", index=False)

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
                self.logger.info(f"ExportaÃ§Ã£o Excel concluÃ­da: {path}")
            except Exception:
                logging.debug("Falha ignorada")

            messagebox.showinfo(
                "OK",
                "ExportaÃ§Ã£o concluÃ­da!\n\n"
                f"Arquivo: {os.path.basename(path)}\n"
                f"Despesas: {len(df_despesas)}\n"
                f"Recebimentos: {len(df_receb)}\n"
                f"Resultado LÃ­quido: {self._fmt_money(resultado_liquido) if hasattr(self, '_fmt_money') else resultado_liquido}"
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

        ttk.Button(filtros, text="🔄 ATUALIZAR", style="Ghost.TButton", command=self.refresh_data).grid(
            row=1, column=2, sticky="w"
        )

        resumo = ttk.Frame(self.body, style="Card.TFrame", padding=12)
        resumo.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        resumo.grid_columnconfigure(0, weight=1)
        ttk.Label(resumo, text="Resumo de distribuicao", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.lbl_resumo = ttk.Label(
            resumo,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=980,
        )
        self.lbl_resumo.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.lbl_recomendacoes = ttk.Label(
            resumo,
            text="-",
            style="CardLabel.TLabel",
            justify="left",
            wraplength=980,
            foreground="#1E3A8A",
        )
        self.lbl_recomendacoes.grid(row=2, column=0, sticky="w", pady=(8, 0))

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

        cols = ("NOME", "ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "LOCAL", "KM_RODADO")
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
            else:
                w = 95
            self.tree_m.column(c, anchor="w" if c in ("NOME", "LOCAL") else "center", width=w)
        enable_treeview_sorting(
            self.tree_m,
            numeric_cols={"ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "KM_RODADO"},
        )

        for c in cols_a:
            self.tree_a.heading(c, text=c)
            if c == "NOME":
                w = 280
            elif c == "KM_RODADO":
                w = 120
            else:
                w = 115
            self.tree_a.column(c, anchor="w" if c == "NOME" else "center", width=w)
        enable_treeview_sorting(
            self.tree_a,
            numeric_cols={"ROTAS", "EM_ROTA", "ATIVAS", "FINALIZADAS", "CANCELADAS", "KM_RODADO"},
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

    def on_show(self):
        self.refresh_data()

    def _status_normalizado(self, v) -> str:
        return upper(str(v or "").strip())

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
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
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

    def _tag_por_carga(self, rotas: int, media_rotas: float, km_rodado: float = 0.0, media_km: float = 0.0) -> str:
        idx_rotas = (float(rotas) / float(media_rotas)) if media_rotas and media_rotas > 0 else 0.0
        idx_km = (float(km_rodado) / float(media_km)) if media_km and media_km > 0 else 0.0
        idx = max(idx_rotas, idx_km)
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

            def _score_motorista(d):
                rotas = float(d.get("rotas", 0) or 0)
                em_rota = float(d.get("em_rota", 0) or 0)
                kg = float(d.get("kg", 0.0) or 0.0)
                km = float(d.get("km_rodado", 0.0) or 0.0)
                score = (rotas / media_rotas) + (em_rota * 0.85) + (kg / media_kg) * 0.65 + (km / media_km) * 0.55

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
            sob_mot = [n for n, d in mot_rows if int(d.get("rotas", 0)) > media_mot * 1.25]
            if sob_mot:
                recs.append("Evitar novas rotas para motoristas sobrecarregados: " + ", ".join(sob_mot[:4]))

        if aj_rows:
            media_aj = max(qtd_rotas / float(len(aj_rows)), 1.0)
            media_km_aj = max(sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in aj_rows) / float(len(aj_rows)), 1.0)

            def _score_ajudante(d):
                rotas = float(d.get("rotas", 0) or 0)
                em_rota = float(d.get("em_rota", 0) or 0)
                km = float(d.get("km_rodado", 0.0) or 0.0)
                score = (rotas / media_aj) + (em_rota * 0.90) + (km / media_km_aj) * 0.55
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
            sob_aj = [n for n, d in aj_rows if int(d.get("rotas", 0)) > media_aj * 1.25]
            if sob_aj:
                recs.append("Evitar escalar ajudantes sobrecarregados: " + ", ".join(sob_aj[:4]))

        if not recs:
            return "Recomendações: distribuição está equilibrada no filtro atual."
        return f"Recomendações (local-alvo: {local_alvo}):\n• " + "\n• ".join(recs)

    def _listar_programacoes_filtradas(self):
        status_filtro = self._status_normalizado(self.var_status.get())
        periodo = upper(self.var_periodo.get().strip())
        cutoff = None
        if periodo != "TODAS":
            try:
                cutoff = datetime.now() - timedelta(days=int(periodo))
            except Exception:
                cutoff = None

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
            cur.execute(
                f"""
                SELECT codigo_programacao, data_criacao, motorista, equipe, status, kg_estimado,
                       {local_expr} AS local_rota, {km_rodado_expr} AS km_rodado
                  FROM programacoes
                 ORDER BY id DESC
                """
            )
            rows = cur.fetchall() or []

        out = []
        for r in rows:
            st = self._status_normalizado(r["status"] if hasattr(r, "keys") else r[4])
            if status_filtro == "ATIVAS":
                if not self._is_ativa(st):
                    continue
            elif status_filtro != "TODOS":
                if st != status_filtro:
                    continue

            if cutoff is not None:
                dt = self._parse_data_programacao(r["data_criacao"] if hasattr(r, "keys") else r[1])
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
                st = self._status_normalizado(r["status"] if hasattr(r, "keys") else r[4])
                kg = safe_float(r["kg_estimado"] if hasattr(r, "keys") else r[5], 0.0)
                dt_prog = self._parse_data_programacao(r["data_criacao"] if hasattr(r, "keys") else r[1])
                local_rota = upper(str((r["local_rota"] if hasattr(r, "keys") else r[6]) or "").strip())
                km_rodado = safe_float(r["km_rodado"] if hasattr(r, "keys") else r[7], 0.0)

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
                        "local_ref": "-",
                        "last_dt": None,
                        "local_counts": {},
                    },
                )
                m["rotas"] += 1
                m["kg"] += kg
                m["km_rodado"] += km_rodado
                if isinstance(dt_prog, datetime):
                    if m.get("last_dt") is None or dt_prog > m["last_dt"]:
                        m["last_dt"] = dt_prog
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
                            "last_dt": None,
                        },
                    )
                    a["rotas"] += 1
                    a["km_rodado"] += km_rodado
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
            for nome, d in mot_rows:
                tag = self._tag_por_carga(
                    int(d["rotas"]),
                    media_mot,
                    float(d.get("km_rodado", 0.0) or 0.0),
                    media_km_mot,
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
                    ),
                    tags=(tag,),
                )

            aj_rows = sorted(aju.items(), key=lambda kv: (-kv[1]["rotas"], kv[0]))
            media_aju = (qtd_rotas / float(len(aj_rows))) if aj_rows else 0.0
            media_km_aju = (
                sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in aj_rows) / float(len(aj_rows))
            ) if aj_rows else 0.0
            for nome, d in aj_rows:
                tag = self._tag_por_carga(
                    int(d["rotas"]),
                    media_aju,
                    float(d.get("km_rodado", 0.0) or 0.0),
                    media_km_aju,
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
                    ),
                    tags=(tag,),
                )

            qtd_motoristas = len(mot_rows)
            qtd_ajudantes = len(aj_rows)
            total_km = sum(float(d.get("km_rodado", 0.0) or 0.0) for _n, d in mot_rows)
            if qtd_motoristas > 0:
                media = qtd_rotas / float(qtd_motoristas)
                media_km = total_km / float(qtd_motoristas)
                mais_sobrecarregado = mot_rows[0][0] if mot_rows else "-"
                msg = (
                    f"Rotas no filtro: {qtd_rotas} | Motoristas: {qtd_motoristas} | Ajudantes: {qtd_ajudantes} | "
                    f"Media por motorista: {media:.2f} | KM total: {total_km:.1f} | KM médio/motorista: {media_km:.1f} | "
                    f"Maior carga atual: {mais_sobrecarregado} | "
                    "Legenda: verde=equilibrado, laranja=alerta, vermelho=sobrecarga (por rotas ou km)"
                )
            else:
                msg = f"Rotas no filtro: {qtd_rotas} | Sem motoristas no periodo/filtro selecionado."

            self.lbl_resumo.config(text=msg)
            self.lbl_recomendacoes.config(
                text=self._recomendacoes_distribuicao(qtd_rotas, mot_rows, aj_rows)
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

        ttk.Label(filtros, text="Período (dias)", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.var_periodo = tk.StringVar(value="30")
        self.cb_periodo = ttk.Combobox(
            filtros, textvariable=self.var_periodo, state="readonly", width=10,
            values=["7", "15", "30", "60", "90", "180", "TODAS"]
        )
        self.cb_periodo.grid(row=1, column=0, sticky="w", padx=(0, 10))

        ttk.Label(filtros, text="Veículo", style="CardLabel.TLabel").grid(row=0, column=1, sticky="w")
        self.var_veiculo = tk.StringVar(value="TODOS")
        self.cb_veiculo = ttk.Combobox(
            filtros, textvariable=self.var_veiculo, state="readonly", width=16, values=["TODOS"]
        )
        self.cb_veiculo.grid(row=1, column=1, sticky="w", padx=(0, 10))

        ttk.Button(filtros, text="🔄 ATUALIZAR", style="Ghost.TButton", command=self.refresh_data).grid(
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
        self.lbl_chart_title = ttk.Label(chart_wrap, text="Custo por Veículo (Custo/KM)", style="CardTitle.TLabel")
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
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
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
            title = "Custo por Veículo (Custo/KG)"
        elif metric == "DESPESA_TOTAL":
            idx = 4
            title = "Custo por Veículo (Despesa Total)"
        else:
            idx = 5
            title = "Custo por Veículo (Custo/KM)"
        self.lbl_chart_title.config(text=title)
        chart_rows = sorted(rows_out, key=lambda r: safe_float(r[idx], 0.0), reverse=True)
        self._chart_labels = [r[0] for r in chart_rows[:10]]
        self._chart_values = [r[idx] for r in chart_rows[:10]]
        self._draw_chart(self._chart_labels, self._chart_values)

        self.lbl_resumo.config(
            text=(
                f"Veículos: {len(rows_out)} | Rotas: {total_rotas} | "
                f"KM: {total_km:.1f} | KG carregado: {total_kg:.2f} | "
                f"Despesas: R$ {total_desp:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                + f" | Custo/KM global: {custo_km_global:.3f} | Custo/KG global: {custo_kg_global:.3f}"
            )
        )
        self.set_status(f"STATUS: Centro de Custos atualizado ({self.var_periodo.get()} dias / veículo {self.var_veiculo.get()}).")


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

        ttk.Button(card, text="🔎 BUSCAR", style="Primary.TButton", command=self._buscar_programacoes_relatorio).grid(
            row=1, column=4, padx=6
        )
        ttk.Button(card, text="🧹 LIMPAR", style="Ghost.TButton", command=self._limpar_filtros_relatorio).grid(
            row=1, column=5, padx=6
        )

        ttk.Label(card, text="Programacao", style="CardLabel.TLabel").grid(row=2, column=0, sticky="w")
        self.cb_prog = ttk.Combobox(card, state="readonly")
        self.cb_prog.grid(row=3, column=0, sticky="ew", padx=6)

        ttk.Button(card, text="🧾 GERAR RESUMO", style="Primary.TButton", command=self.gerar_resumo).grid(
            row=3, column=1, padx=6
        )
        ttk.Button(card, text="📤 EXPORTAR EXCEL", style="Warn.TButton", command=self.exportar_excel).grid(
            row=3, column=2, padx=6
        )
        ttk.Button(card, text="📄 GERAR PDF", style="Primary.TButton", command=self.gerar_pdf).grid(
            row=3, column=3, padx=6
        )
        ttk.Button(card, text="👁 PREVIEW", style="Ghost.TButton", command=self.abrir_previsualizacao_relatorio).grid(
            row=3, column=4, padx=6
        )
        ttk.Button(card, text="🔄 ATUALIZAR", style="Ghost.TButton", command=self.refresh_comboboxes).grid(
            row=3, column=5, padx=6
        )

        ttk.Button(card, text="🏁 FINALIZAR ROTA", style="Danger.TButton", command=self.finalizar_rota).grid(
            row=3, column=6, padx=6
        )
        ttk.Button(card, text="↩ REABRIR ROTA", style="Warn.TButton", command=self.reabrir_rota).grid(
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
    # Helpers de regra/seguranÃ§a
    # ----------------------------
    def _get_prog_status_info(self, prog: str):
        """
        Retorna dict com status/prestacao_status quando existir.
        CompatÃ­vel com bases antigas (sem coluna prestacao_status).
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

            # Área do gráfico
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
        self.set_status("STATUS: RelatÃ³rios e exportaÃ§Ã£o.")

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
            rec_lines = []
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
        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("PRAGMA table_info(programacoes)")
                cols_p = {str(c[1]).lower() for c in (cur.fetchall() or [])}
                local_expr = "COALESCE(local_rota,'')" if "local_rota" in cols_p else ("COALESCE(local,'')" if "local" in cols_p else "''")
                kg_expr = "COALESCE(kg_estimado,0)" if "kg_estimado" in cols_p else "0"
                status_expr = "COALESCE(status,'')" if "status" in cols_p else "''"

                cur.execute(f"""
                    SELECT COALESCE(data_criacao,''), COALESCE(motorista,''), COALESCE(veiculo,''),
                           COALESCE(equipe,''), {kg_expr}, {status_expr}, {local_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                meta = cur.fetchone() or ("", "", "", "", 0, "", "")

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

        data_criacao, motorista, veiculo, equipe, kg_estimado, status, local = meta
        equipe_txt = resolve_equipe_nomes(equipe)
        total_prev = sum(safe_float(r[2], 0.0) for r in itens)
        lines = []
        lines.append("=" * 118)
        lines.append(f"{'FOLHA DE PROGRAMAÇÃO':^118}")
        lines.append("=" * 118)
        lines.append(f"CÓDIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}")
        lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEÍCULO: {upper(veiculo or '-')}   LOCAL: {upper(local or '-')}")
        lines.append(f"EQUIPE: {equipe_txt or '-'}")
        lines.append(f"KG ESTIMADO: {safe_float(kg_estimado, 0.0):.2f}   CLIENTES: {len(itens)}   TOTAL ESTIMADO: {self._fmt_rel_money(total_prev)}")
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
                cur.execute(f"""
                    SELECT COALESCE(data_criacao,''), COALESCE(motorista,''), COALESCE(veiculo,''), COALESCE(equipe,''),
                           {status_expr}, {prest_expr}, {km_expr}, {media_expr}, {custo_expr}, {adiant_expr}
                    FROM programacoes
                    WHERE codigo_programacao=?
                    LIMIT 1
                """, (prog,))
                meta = cur.fetchone() or ("", "", "", "", "", "", 0, 0, 0, 0)

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

        data_criacao, motorista, veiculo, equipe, status, prest, km_rodado, media_km_l, custo_km, adiantamento = meta
        equipe_txt = resolve_equipe_nomes(equipe)
        entradas = total_receb + safe_float(adiantamento, 0.0)
        saidas = total_desp
        resultado = entradas - saidas

        lines = []
        lines.append("=" * 118)
        lines.append(f"{'FOLHA DE PRESTAÇÃO DE CONTAS':^118}")
        lines.append("=" * 118)
        lines.append(f"CÓDIGO: {prog}   DATA: {data_criacao or '-'}   STATUS: {upper(status or '-')}   PRESTAÇÃO: {upper(prest or '-')}")
        lines.append(f"MOTORISTA: {upper(motorista or '-')}   VEÍCULO: {upper(veiculo or '-')}   EQUIPE: {equipe_txt or '-'}")
        lines.append("-" * 118)
        lines.append(
            f"ENTRADAS (RECEB+ADIANT.): {self._fmt_rel_money(entradas)}   "
            f"SAÍDAS (DESPESAS): {self._fmt_rel_money(saidas)}   "
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
        lines.append("Conferido por: ____________________________________   Data: ____/____/________")
        return "\n".join(lines)

    def _gerar_relatorio_rotina_motoristas_ajudantes(self):
        filtro_mot = upper(self.ent_filtro_motorista.get().strip()) if hasattr(self, "ent_filtro_motorista") else ""
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
            f"Veículos: {len(rows)}",
            f"KM total: {total_km:.2f}",
            f"Média KM/veículo: {(total_km / max(len(rows), 1)):.2f}",
            f"Destaque: {upper(top or '-')}",
            [upper((r[0] or "-")) for r in rows[:8]],
            [safe_float(r[2], 0.0) for r in rows[:8]],
            color="#2563EB",
        )
        self.set_status(f"STATUS: Relatório de KM por veículo gerado ({len(rows)} veículo(s)).")

    def _gerar_relatorio_despesas_geral(self):
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
            ["Entradas", "Saídas", "Resultado"],
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
        self.txt.insert("end", "POS | PROGRAMAÃ‡ÃƒO | MOTORISTA | MORTALIDADE | CLIENTES C/ MORT. | DATA | STATUS\n")
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
            f"Motorista com menor mÃ©dia de mortalidade por rota: {melhor_mot} "
            f"(mÃ©dia {melhor_media:.2f} aves/rota em {melhor_data['rotas']} rota(s)).\n",
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
        self.set_status(f"STATUS: RelatÃ³rio de mortalidade gerado ({len(rows)} rota(s)).")

    def exportar_excel(self):
        prog = upper(self.cb_prog.get())
        if not (require_pandas() and require_openpyxl()):
            return
        if not prog:
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma programaÃ§Ã£o.")
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

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df_itens = pd.DataFrame(itens, columns=cols_itens)
                df_rec = pd.DataFrame(rec, columns=cols_rec)
                df_desp = pd.DataFrame(desp, columns=cols_desp)

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
                    messagebox.showerror("ERRO", "Tela de Despesas indisponível para gerar PDF da prestação.")
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma programaÃ§Ã£o.")
            return

        info = self._get_prog_status_info(prog)
        st = info.get("status", "")
        prest = info.get("prestacao_status", "")

        # âœ… Se prestaÃ§Ã£o fechada, nÃ£o faz sentido "finalizar" (jÃ¡ deveria estar travada)
        if prest == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"A rota {prog} estÃ¡ com a prestaÃ§Ã£o FECHADA.\n\n"
                "Ela jÃ¡ estÃ¡ travada para alteraÃ§Ãµes. (VocÃª pode apenas gerar relatÃ³rios/exportar.)"
            )
            return

        # âœ… Evita clicar 2x
        if st == "FINALIZADA":
            messagebox.showinfo("Info", f"A rota {prog} jÃ¡ estÃ¡ FINALIZADA.")
            return

        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja FINALIZAR a rota {prog}?\n\n"
            "Ela deixarÃ¡ de aparecer nas Rotas Ativas."
        ):
            return

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("UPDATE programacoes SET status='FINALIZADA' WHERE codigo_programacao=?", (prog,))

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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Selecione uma programaÃ§Ã£o.")
            return

        info = self._get_prog_status_info(prog)
        st = info.get("status", "")
        prest = info.get("prestacao_status", "")

        # âœ… Regra principal: prestaÃ§Ã£o FECHADA => nÃ£o reabrir rota (senÃ£o volta a mexer em despesas/adiantamento)
        if prest == "FECHADA":
            messagebox.showwarning(
                "BLOQUEADO",
                f"NÃ£o Ã© permitido REABRIR a rota {prog} pois a prestaÃ§Ã£o estÃ¡ FECHADA.\n\n"
                "Se precisar reabrir, primeiro reabra a prestaÃ§Ã£o (criar funÃ§Ã£o especÃ­fica) ou ajuste no administrativo."
            )
            return

        if st == "ATIVA":
            messagebox.showinfo("Info", f"A rota {prog} jÃ¡ estÃ¡ ATIVA.")
            return

        if not messagebox.askyesno(
            "CONFIRMAR",
            f"Deseja REABRIR a rota {prog}?\n\n"
            "Ela voltarÃ¡ a aparecer nas Rotas Ativas."
        ):
            return

        try:
            with get_db() as conn:
                cur = conn.cursor()
                cur.execute("UPDATE programacoes SET status='ATIVA' WHERE codigo_programacao=?", (prog,))

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

        ttk.Button(card, text="🗄 FAZER BACKUP DO BANCO", style="Primary.TButton", command=self.backup_db)\
            .grid(row=1, column=0, sticky="ew", pady=6)
        ttk.Button(card, text="♻ RESTAURAR BANCO (IMPORTAR .DB)", style="Warn.TButton", command=self.restore_db)\
            .grid(row=2, column=0, sticky="ew", pady=6)
        ttk.Button(card, text="📤 EXPORTAR VENDAS IMPORTADAS (EXCEL)", style="Ghost.TButton", command=self.exportar_vendas)\
            .grid(row=3, column=0, sticky="ew", pady=6)

        self.lbl = ttk.Label(card, text="Dica: FaÃ§a backup diariamente.", background="white", foreground="#444")
        self.lbl.grid(row=4, column=0, sticky="w", pady=(12, 0))

    def on_show(self):
        self.set_status("STATUS: Backup e exportaÃ§Ãµes.")

    # ----------------------------
    # Helpers de seguranÃ§a
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
        """CÃ³pia binÃ¡ria simples (mantÃ©m compatibilidade)."""
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
        """Cria um backup automÃ¡tico no mesmo diretÃ³rio do DB (antes do restore)."""
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
            messagebox.showerror("ERRO", "Banco nÃ£o encontrado.")
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
            # âœ… melhor prÃ¡tica: usar SQLite backup API quando possÃ­vel
            try:
                import sqlite3
                src = sqlite3.connect(DB_PATH)
                dst = sqlite3.connect(path)
                with dst:
                    src.backup(dst)
                dst.close()
                src.close()
            except Exception:
                # fallback: cÃ³pia binÃ¡ria (mantÃ©m funcionamento se der algo no backup API)
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

        # âœ… valida se Ã© sqlite de verdade
        if not self._is_sqlite_db_file(path):
            messagebox.showerror(
                "ERRO",
                "O arquivo selecionado nÃ£o parece ser um banco SQLite vÃ¡lido.\n\n"
                "Selecione um .db gerado pelo sistema (backup)."
            )
            return

        # confirmaÃ§Ã£o mais clara (alto risco)
        if not messagebox.askyesno(
            "CONFIRMAR RESTAURAÃ‡ÃƒO",
            "Isso vai SUBSTITUIR seu banco atual.\n\n"
            "Recomendado: fechar telas e nÃ£o estar com operaÃ§Ãµes em andamento.\n\n"
            "Deseja continuar?"
        ):
            return

        try:
            # âœ… cria backup automÃ¡tico do DB atual
            auto_backup = self._backup_current_db_automatic()

            # ÃƒÂ¢Ã…â€œÃ¢â‚¬Â¦ tenta fechar conexÃƒÆ’Ã‚Âµes ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œconhecidasÃƒÂ¢Ã¢â€šÂ¬Ã‚ (melhor esforÃƒÆ’Ã‚Â§o, sem quebrar)
            try:
                # se existir algum mÃ©todo no app para reabrir/fechar conexÃµes, chamamos
                if hasattr(self.app, "close_db_connections"):
                    self.app.close_db_connections()
            except Exception:
                logging.debug("Falha ignorada")

            # limpa WAL/SHM antigos (melhor esforÃ§o)
            try:
                self._cleanup_sqlite_wal(DB_PATH)
            except Exception:
                logging.debug("Falha ignorada")

            # âœ… restaura por cÃ³pia binÃ¡ria (simples e compatÃ­vel)
            # ObservaÃ§Ã£o: se houver conexÃ£o aberta, pode falhar no Windows.
            self._make_safe_copy(path, DB_PATH)

            msg = "Banco restaurado! Reinicie o sistema."
            if auto_backup:
                msg += f"\n\nBackup automÃ¡tico do banco anterior:\n{os.path.basename(auto_backup)}"

            messagebox.showinfo("OK", msg)
            self.set_status("STATUS: Banco restaurado. Reinicie o sistema.")

        except PermissionError:
            messagebox.showerror(
                "ERRO",
                "NÃ£o foi possÃ­vel substituir o banco (arquivo em uso).\n\n"
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
                messagebox.showwarning("ATENÃ‡ÃƒO", "NÃ£o hÃ¡ vendas importadas para exportar.")
                return

            try:
                if "data_venda" in df.columns:
                    df["data_venda"] = normalize_date_column(df["data_venda"])
                df.to_excel(path, index=False)
            except Exception as e:
                messagebox.showerror(
                    "ERRO",
                    "Falha ao exportar para Excel.\n\n"
                    "Verifique se o pacote 'openpyxl' estÃ¡ instalado.\n\n"
                    f"Detalhes: {str(e)}"
                )
                return

            messagebox.showinfo("OK", "ExportaÃ§Ã£o feita com sucesso!")
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
        self.geometry("440x280")  # +20px p/ nÃ£o apertar botÃµes
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

        # ---- Fallback de estilo (evita botÃ£o "sumir" por estilo invisÃ­vel)
        self._ensure_button_styles()

        # ---- Controle de tentativas (seguranÃ§a leve sem quebrar fluxo)
        self._attempts = 0
        self._blocked_until = 0  # epoch seconds
        self._max_attempts = 5
        self._block_seconds = 15

        card = ttk.Frame(self, style="Card.TFrame", padding=18)
        card.pack(fill="both", expand=True, padx=12, pady=12)

        # âœ… garante espaÃ§o do grid dentro do card
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

        # âœ… padding e sticky garantem que apareÃ§am e tenham tamanho
        self.btn_entrar = ttk.Button(btns, text="🔐 ENTRAR", style="Primary.TButton", command=self.try_login)
        self.btn_entrar.grid(row=0, column=0, sticky="ew", padx=(0, 6), ipady=6)

        self.btn_sair = ttk.Button(btns, text="⏻ SAIR", style="Danger.TButton", command=self._request_close)
        self.btn_sair.grid(row=0, column=1, sticky="ew", padx=(6, 0), ipady=6)

        self.ent_codigo.focus_set()
        self.bind("<Return>", lambda e: self.try_login())
        self.bind("<Escape>", lambda e: self._request_close())
        self.protocol("WM_DELETE_WINDOW", self._request_close)

        self.user = None  # serÃ¡ preenchido quando logar

    def _request_close(self):
        try:
            ok = messagebox.askyesno("Confirmar saída", "Deseja realmente fechar o sistema?")
        except Exception:
            ok = True
        if ok:
            self.destroy()

    def _ensure_button_styles(self):
        """
        Se o tema nÃ£o criou Primary.TButton / Danger.TButton, cria um fallback
        para evitar o efeito de "botÃ£o invisÃ­vel".
        NÃ£o altera seu tema se jÃ¡ existir.
        """
        try:
            st = ttk.Style(self)
            existing = set(st.theme_names())  # sÃ³ pra nÃ£o dar erro em alguns temas
            _ = existing  # quiet

            # Alguns temas nÃ£o suportam lookup; entÃ£o testamos com lookup e fallback
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
            messagebox.showwarning("ATENÃ‡ÃƒO", "Informe cÃ³digo e senha.")
            return

        # âœ… login real: usa sua funÃ§Ã£o existente (mantÃ©m o sistema)
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
                    text=f"CÃ³digo ou senha invÃ¡lidos. Tentativas: {self._attempts}/{self._max_attempts}"
                )
            return

        self.user = user
        self.destroy()


def abrir_login():
    win = LoginWindow()
    win.mainloop()

    if not getattr(win, "user", None):
        return  # usuÃ¡rio saiu / cancelou

    app = App(user=win.user)

    if not hasattr(app, "close_db_connections"):
        def _close_db_connections_best_effort():
            return
        app.close_db_connections = _close_db_connections_best_effort

    app.mainloop()


if __name__ == "__main__":
    db_init()  # garante migraÃ§Ãµes
    abrir_login()

# ==========================
# ===== FIM DA PARTE 10 (ATUALIZADA) =====
# ==========================


