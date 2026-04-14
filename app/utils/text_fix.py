_MOJIBAKE_MARKERS = (
    "\u00c3",
    "\u00c2",
    "\u00e2",
    "\u0192",
    "\u00c6",
    "\u00a2",
    "\u20ac",
    "\u2122",
    "\u0153",
    "\u017e",
    "\ufffd",
    "\u0102",
    "\u0103",
    "\u0161",
    "\u2030",
    "\x89",
    "\x9a",
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
