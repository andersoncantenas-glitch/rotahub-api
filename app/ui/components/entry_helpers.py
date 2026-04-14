# -*- coding: utf-8 -*-
import logging
import re
import tkinter as tk

from app.utils.formatters import format_date_br_short, normalize_date, normalize_time


def _format_money_from_digits(raw: str) -> str:
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
                masked = _format_money_from_digits(raw)
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
                entry.insert(0, _format_money_from_digits(raw))
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
