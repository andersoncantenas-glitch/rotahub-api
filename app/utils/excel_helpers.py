# -*- coding: utf-8 -*-
import logging
import re
import unicodedata

try:
    from tkinter import messagebox
except Exception:
    messagebox = None

try:
    import pandas as pd
except Exception:
    pd = None


def _show_error(title: str, message: str) -> None:
    if messagebox is None:
        logging.warning("%s: %s", title, message.replace("\n", " "))
        return
    try:
        messagebox.showerror(title, message)
    except Exception:
        logging.debug("Falha ignorada")


def upper(s):
    return str(s).strip().upper() if s is not None else ""


def require_pandas() -> bool:
    if pd is None:
        _show_error(
            "ERRO",
            "Esta funcionalidade requer o pacote 'pandas'.\n\n"
            "Instale com: pip install pandas"
        )
        return False
    return True


def _require_openpyxl() -> bool:
    try:
        import openpyxl  # noqa: F401
        return True
    except Exception:
        _show_error(
            "ERRO",
            "Exportacao Excel requer o pacote 'openpyxl'.\n\n"
            "Instale com: pip install openpyxl"
        )
        return False


def _require_xlrd() -> bool:
    try:
        import xlrd  # noqa: F401  # type: ignore[import-not-found]
        return True
    except Exception:
        _show_error(
            "ERRO",
            "Importacao de .xls requer o pacote 'xlrd'.\n\n"
            "Instale com: pip install xlrd"
        )
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
        return _require_xlrd()
    if p.endswith(".xlsx"):
        return _require_openpyxl()
    _show_error(
        "ERRO",
        "Formato de arquivo nao suportado.\n\n"
        "Use .xlsx ou .xls."
    )
    return False


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
        mojibake_repl = {
            "ã¡": "a", "ã ": "a", "ã¢": "a", "ã£": "a",
            "ã©": "e", "ã¨": "e", "ãª": "e",
            "ã¬": "i", "ã®": "i",
            "ã³": "o", "ã²": "o", "ã´": "o", "ãµ": "o",
            "ãº": "u", "ã¹": "u", "ã»": "u",
            "ã§": "c",
        }
        for k, v in mojibake_repl.items():
            s = s.replace(k, v)
        s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
        s = re.sub(r"[^a-z0-9]+", " ", s).strip()
        return s

    cols_lower = {_norm(c): c for c in cols}
    for cand in candidates:
        cand_low = _norm(cand)
        for col_low, original in cols_lower.items():
            if cand_low and (cand_low == col_low or cand_low in col_low):
                return original
    return None


__all__ = [
    "upper",
    "require_pandas",
    "require_excel_support",
    "excel_engine_for",
    "guess_col",
]
