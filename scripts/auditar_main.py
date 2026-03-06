from __future__ import annotations

import ast
from collections import Counter
from pathlib import Path
from typing import Iterable


ROOT = Path(__file__).resolve().parents[1]
MAIN_PY = ROOT / "main.py"

REQUIRED_PAGES = {
    "HomePage",
    "CadastrosPage",
    "RotasPage",
    "ImportarVendasPage",
    "ProgramacaoPage",
    "RecebimentosPage",
    "DespesasPage",
    "EscalaPage",
    "CentroCustosPage",
    "RelatoriosPage",
    "BackupExportarPage",
}

REQUIRED_METHODS = {
    "ImportarVendasPage": {"importar_excel", "carregar"},
    "ProgramacaoPage": {"carregar_vendas_selecionadas", "salvar_programacao"},
    "RecebimentosPage": {"carregar_programacao", "salvar_recebimento"},
    "DespesasPage": {"salvar_tudo"},
    "HomePage": {"_update_api_status", "_open_rota_preview"},
}


def _find_dups(values: Iterable[str]) -> list[str]:
    c = Counter(values)
    return sorted([k for k, v in c.items() if v > 1])


def main() -> int:
    txt = MAIN_PY.read_text(encoding="utf-8", errors="replace")
    tree = ast.parse(txt)

    class_defs = [n for n in tree.body if isinstance(n, ast.ClassDef)]
    func_defs = [n for n in tree.body if isinstance(n, ast.FunctionDef)]

    class_names = [c.name for c in class_defs]
    func_names = [f.name for f in func_defs]

    dup_classes = _find_dups(class_names)
    dup_funcs = _find_dups(func_names)

    class_methods: dict[str, set[str]] = {}
    for c in class_defs:
        methods = {n.name for n in c.body if isinstance(n, ast.FunctionDef)}
        class_methods[c.name] = methods

    missing_pages = sorted(REQUIRED_PAGES - set(class_names))
    missing_methods = []
    for cls, methods in REQUIRED_METHODS.items():
        for m in sorted(methods):
            if m not in class_methods.get(cls, set()):
                missing_methods.append(f"{cls}.{m}")

    mojibake_tokens = [
        "Ã",
        "â€",
        "â€œ",
        "â€",
        "Â",
        "�",
    ]
    mojibake_hits = []
    for i, line in enumerate(txt.splitlines(), start=1):
        if any(tok in line for tok in mojibake_tokens):
            mojibake_hits.append((i, line.strip()))
            if len(mojibake_hits) >= 20:
                break

    def _safe(s: str) -> str:
        return str(s).encode("cp1252", errors="replace").decode("cp1252", errors="replace")

    print(_safe("=== AUDITORIA MAIN.PY ==="))
    print(_safe(f"Arquivo: {MAIN_PY}"))
    print(_safe(f"Classes totais: {len(class_defs)}"))
    print(_safe(f"Funcoes top-level totais: {len(func_defs)}"))
    print(_safe(f"Classes duplicadas: {dup_classes or 'nenhuma'}"))
    print(_safe(f"Funcoes top-level duplicadas: {dup_funcs or 'nenhuma'}"))
    print(_safe(f"Paginas obrigatorias ausentes: {missing_pages or 'nenhuma'}"))
    print(_safe(f"Metodos criticos ausentes: {missing_methods or 'nenhum'}"))

    if mojibake_hits:
        print(_safe("Possiveis mojibake (primeiros 20):"))
        for ln, content in mojibake_hits:
            print(_safe(f"  L{ln}: {content[:160]}"))
    else:
        print(_safe("Possiveis mojibake: nenhum detectado pelo token scanner."))

    has_error = bool(dup_classes or dup_funcs or missing_pages or missing_methods)
    return 1 if has_error else 0


if __name__ == "__main__":
    raise SystemExit(main())
