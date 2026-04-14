from __future__ import annotations

import ast
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MAIN_PY = ROOT / "main.py"
PAGES_DIR = ROOT / "app" / "ui" / "pages"

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

REQUIRED_FACTORY_BINDINGS = {
    "Home": "HomePage",
    "Cadastros": "CadastrosPage",
    "Rotas": "RotasPage",
    "ImportarVendas": "ImportarVendasPage",
    "Programacao": "ProgramacaoPage",
    "Recebimentos": "RecebimentosPage",
    "Despesas": "DespesasPage",
    "Escala": "EscalaPage",
    "CentroCustos": "CentroCustosPage",
    "Relatorios": "RelatoriosPage",
    "BackupExportar": "BackupExportarPage",
}

MOJIBAKE_TOKENS = (
    "Ãƒ",
    "Ã¢â‚¬",
    "Ã¢â‚¬Å“",
    "Ã¢â‚¬Â",
    "Ã‚",
    "ï¿½",
    "âÅ",
)


@dataclass(frozen=True)
class ParsedModule:
    path: Path
    class_methods: dict[str, set[str]]
    top_level_funcs: list[str]


def _safe(s: str) -> str:
    return str(s).encode("cp1252", errors="replace").decode("cp1252", errors="replace")


def _find_dups(values: list[str]) -> list[str]:
    counts = Counter(values)
    return sorted([key for key, value in counts.items() if value > 1])


def _source_files() -> list[Path]:
    files = [MAIN_PY]
    if PAGES_DIR.exists():
        files.extend(sorted(PAGES_DIR.glob("*.py")))
    return files


def _parse_module(path: Path) -> ParsedModule:
    txt = path.read_text(encoding="utf-8", errors="replace")
    tree = ast.parse(txt, filename=str(path))

    class_methods: dict[str, set[str]] = {}
    top_level_funcs: list[str] = []

    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            methods = {child.name for child in node.body if isinstance(child, ast.FunctionDef)}
            class_methods[node.name] = methods
        elif isinstance(node, ast.FunctionDef):
            top_level_funcs.append(node.name)

    return ParsedModule(path=path, class_methods=class_methods, top_level_funcs=top_level_funcs)


def _scan_mojibake(paths: list[Path], limit: int = 20) -> list[tuple[str, int, str]]:
    hits: list[tuple[str, int, str]] = []
    for path in paths:
        txt = path.read_text(encoding="utf-8", errors="replace")
        for lineno, line in enumerate(txt.splitlines(), start=1):
            if any(token in line for token in MOJIBAKE_TOKENS):
                hits.append((str(path.relative_to(ROOT)).replace("\\", "/"), lineno, line.strip()))
                if len(hits) >= limit:
                    return hits
    return hits


def _factory_bindings_status(main_text: str) -> list[str]:
    missing: list[str] = []
    for page_name, class_name in REQUIRED_FACTORY_BINDINGS.items():
        pattern = re.compile(rf'["\']{re.escape(page_name)}["\']\s*:\s*{re.escape(class_name)}\b')
        if not pattern.search(main_text):
            missing.append(f"{page_name}:{class_name}")
    return missing


def main() -> int:
    source_files = _source_files()
    parse_errors: list[str] = []
    modules: list[ParsedModule] = []
    for path in source_files:
        try:
            modules.append(_parse_module(path))
        except Exception as exc:
            parse_errors.append(f"{path.relative_to(ROOT)} -> {exc}")

    all_class_names: list[str] = []
    all_top_level_funcs: list[str] = []
    main_top_level_funcs: list[str] = []
    class_methods: dict[str, set[str]] = {}
    class_origins: dict[str, str] = {}
    for module in modules:
        all_top_level_funcs.extend(module.top_level_funcs)
        if module.path == MAIN_PY:
            main_top_level_funcs.extend(module.top_level_funcs)
        for class_name, methods in module.class_methods.items():
            all_class_names.append(class_name)
            class_methods[class_name] = methods
            class_origins.setdefault(
                class_name,
                str(module.path.relative_to(ROOT)).replace("\\", "/"),
            )

    dup_classes = _find_dups(all_class_names)
    dup_funcs = _find_dups(main_top_level_funcs)
    missing_pages = sorted(REQUIRED_PAGES - set(all_class_names))

    missing_methods: list[str] = []
    for class_name, required_methods in REQUIRED_METHODS.items():
        existing = class_methods.get(class_name, set())
        for method_name in sorted(required_methods):
            if method_name not in existing:
                missing_methods.append(f"{class_name}.{method_name}")

    main_text = MAIN_PY.read_text(encoding="utf-8", errors="replace")
    missing_factory_bindings = _factory_bindings_status(main_text)
    mojibake_hits = _scan_mojibake(source_files)

    print(_safe("=== AUDITORIA MAIN / SHELL DESKTOP ==="))
    print(_safe(f"Arquivo shell: {MAIN_PY}"))
    print(_safe(f"Modulos auditados: {len(source_files)}"))
    print(_safe(f"Classes totais: {len(all_class_names)}"))
    print(_safe(f"Funcoes top-level totais: {len(all_top_level_funcs)}"))
    print(_safe(f"Classes duplicadas: {dup_classes or 'nenhuma'}"))
    print(_safe(f"Funcoes top-level duplicadas: {dup_funcs or 'nenhuma'}"))
    print(_safe(f"Paginas obrigatorias ausentes: {missing_pages or 'nenhuma'}"))
    print(_safe(f"Bindings ausentes no _page_factories: {missing_factory_bindings or 'nenhum'}"))
    print(_safe(f"Metodos criticos ausentes: {missing_methods or 'nenhum'}"))
    if parse_errors:
        print(_safe("Erros de parse:"))
        for item in parse_errors:
            print(_safe(f"- {item}"))
    else:
        print(_safe("Erros de parse: nenhum"))

    print(_safe("Origem das paginas encontradas:"))
    for class_name in sorted(REQUIRED_PAGES):
        origin = class_origins.get(class_name, "-")
        print(_safe(f"- {class_name}: {origin}"))

    if mojibake_hits:
        print(_safe("Possiveis mojibake (primeiros 20):"))
        for rel_path, lineno, content in mojibake_hits:
            print(_safe(f"  {rel_path}:L{lineno}: {content[:160]}"))
    else:
        print(_safe("Possiveis mojibake: nenhum detectado pelo token scanner."))

    has_error = bool(parse_errors or dup_classes or dup_funcs or missing_pages or missing_methods or missing_factory_bindings)
    return 1 if has_error else 0


if __name__ == "__main__":
    raise SystemExit(main())
