import base64
import hashlib
import hmac
import os


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


__all__ = [
    "hash_password_pbkdf2",
    "verify_password_pbkdf2",
]
