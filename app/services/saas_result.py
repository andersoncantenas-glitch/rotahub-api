from __future__ import annotations


def service_result(*, ok: bool, data=None, error: str | None = None, source: str = "local") -> dict:
    return {
        "ok": bool(ok),
        "data": data,
        "error": str(error) if error else None,
        "source": str(source or "local"),
    }


def error_message(exc: Exception, default_message: str) -> str:
    msg = str(exc or "").strip()
    return msg or str(default_message or "Falha inesperada.")
