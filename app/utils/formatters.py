import re
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


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
    """Converte valor monetario para float com 2 casas (Decimal)."""
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
    """Normaliza para YYYY-MM-DD. Retorna '' se vazio, None se invalido."""
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
    """Normaliza para HH:MM:SS. Retorna '' se vazio, None se invalido."""
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


def normalize_date_time_components(data: str, hora: str):
    """Normaliza data/hora e retorna tupla (data, hora) preservando valor original se invalido."""
    nd = normalize_date(data)
    nt = normalize_time(hora)
    out_data = format_date_br_short(nd if nd is not None else (data or ""))
    out_hora = nt if nt is not None else (hora or "")
    return out_data, out_hora


def format_date_time(data: str, hora: str) -> str:
    d, h = normalize_date_time_components(data, hora)
    return f"{d} {h}".strip()


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
    """Formata valor monetario."""
    return f"R$ {safe_float(v,0.0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def now_str():
    """Retorna data/hora atual formatada."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
