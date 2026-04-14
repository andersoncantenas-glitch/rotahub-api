import re


_CPF_DIGITS = re.compile(r"\D+")
_PHONE_DIGITS = re.compile(r"\D+")
_MOTORISTA_COD_RE = re.compile(r"^[A-Z0-9._-]{3,24}$")


def validate_required(value: str, field_name: str, min_len=1, max_len=120):
    v = (value or "").strip()
    if len(v) < min_len:
        return False, f"{field_name} \u00e9 obrigat\u00f3rio."
    if len(v) > max_len:
        return False, f"{field_name} muito longo (m\u00e1x {max_len})."
    return True, ""


def validate_codigo(value: str, field_name="C\u00f3digo", min_len=2, max_len=20):
    v = str(value or "").strip().upper()
    ok, msg = validate_required(v, field_name, min_len=min_len, max_len=max_len)
    if not ok:
        return False, msg
    if not re.match(r"^[A-Z0-9_-]+$", v):
        return False, f"{field_name} deve conter apenas A-Z, 0-9, '_' ou '-'."
    return True, ""


def validate_placa(value: str):
    v = str(value or "").strip().upper()
    ok, msg = validate_required(v, "Placa", min_len=6, max_len=8)
    if not ok:
        return False, msg
    if not re.match(r"^[A-Z0-9-]+$", v):
        return False, "Placa inv\u00e1lida."
    return True, ""


def validate_money(value: str, field_name="Valor"):
    v = (value or "").strip().replace(".", "").replace(",", ".")
    try:
        f = float(v)
        if f < 0:
            return False, f"{field_name} n\u00e3o pode ser negativo."
        return True, ""
    except Exception:
        return False, f"{field_name} inv\u00e1lido."


def normalize_cpf(v: str) -> str:
    return _CPF_DIGITS.sub("", str(v or "").strip())


def normalize_phone(v: str) -> str:
    """
    Normaliza telefone para dígitos.
    Se vier com DDI 55 (Brasil), remove quando fizer sentido.
    """
    s = _PHONE_DIGITS.sub("", str(v or "").strip())
    if len(s) in (12, 13) and s.startswith("55"):
        s2 = s[2:]
        if len(s2) in (10, 11):
            return s2
    return s


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


def is_valid_phone(phone_digits: str) -> bool:
    tel = normalize_phone(phone_digits)
    return len(tel) in (10, 11) and tel.isdigit()


def is_valid_motorista_codigo(cod: str) -> bool:
    cod = str(cod or "").strip().upper()
    return bool(_MOTORISTA_COD_RE.match(cod))


def is_valid_motorista_senha(senha: str) -> bool:
    senha = str(senha or "").strip()
    return 4 <= len(senha) <= 24
