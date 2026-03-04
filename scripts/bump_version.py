#!/usr/bin/env python3
"""
Bump de versao SemVer para o projeto.

Atualiza:
- main.py            -> APP_VERSION = "x.y.z"
- api_server.py      -> FastAPI(..., version="x.y.z")
- updates/version.json -> {"version": "x.y.z", ...}

Uso:
  python scripts/bump_version.py patch
  python scripts/bump_version.py minor
  python scripts/bump_version.py major
  python scripts/bump_version.py --set 1.1.2
  python scripts/bump_version.py patch --dry-run
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN_PY = ROOT / "main.py"
API_SERVER = ROOT / "api_server.py"
VERSION_JSON = ROOT / "updates" / "version.json"

SEMVER_RE = re.compile(r"^\d+\.\d+\.\d+$")
APP_VER_RE = re.compile(r'^(?P<prefix>\s*APP_VERSION\s*=\s*")(?P<ver>\d+\.\d+\.\d+)(".*)$', re.M)
FASTAPI_VER_RE = re.compile(r'(?P<prefix>FastAPI\([^\n]*\bversion\s*=\s*")(?P<ver>\d+\.\d+\.\d+)(")', re.M)


def parse_semver(v: str) -> tuple[int, int, int]:
    if not SEMVER_RE.match(v or ""):
        raise ValueError(f"Versao invalida: {v!r}. Use x.y.z")
    a, b, c = v.split(".")
    return int(a), int(b), int(c)


def bump(v: str, kind: str) -> str:
    major, minor, patch = parse_semver(v)
    if kind == "major":
        return f"{major + 1}.0.0"
    if kind == "minor":
        return f"{major}.{minor + 1}.0"
    if kind == "patch":
        return f"{major}.{minor}.{patch + 1}"
    raise ValueError(f"Tipo de bump invalido: {kind}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str, dry_run: bool) -> None:
    if dry_run:
        return
    path.write_text(content, encoding="utf-8")


def current_version_from_main() -> str:
    txt = read_text(MAIN_PY)
    m = APP_VER_RE.search(txt)
    if not m:
        raise RuntimeError("Nao achei APP_VERSION em main.py")
    return m.group("ver")


def replace_version_in_main(new_v: str, dry_run: bool) -> bool:
    txt = read_text(MAIN_PY)
    new_txt, n = APP_VER_RE.subn(rf"\g<prefix>{new_v}\3", txt, count=1)
    if n != 1:
        raise RuntimeError("Falha ao atualizar APP_VERSION em main.py")
    changed = new_txt != txt
    if changed:
        write_text(MAIN_PY, new_txt, dry_run=dry_run)
    return changed


def replace_version_in_api_server(new_v: str, dry_run: bool) -> bool:
    txt = read_text(API_SERVER)
    m = FASTAPI_VER_RE.search(txt)
    if not m:
        raise RuntimeError("Nao achei FastAPI(... version=...) em api_server.py")
    new_txt, n = FASTAPI_VER_RE.subn(rf"\g<prefix>{new_v}\3", txt, count=1)
    if n != 1:
        raise RuntimeError("Falha ao atualizar versao em api_server.py")
    changed = new_txt != txt
    if changed:
        write_text(API_SERVER, new_txt, dry_run=dry_run)
    return changed


def replace_version_in_manifest(new_v: str, dry_run: bool) -> bool:
    if not VERSION_JSON.exists():
        raise RuntimeError("Nao achei updates/version.json")
    raw = read_text(VERSION_JSON)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"JSON invalido em updates/version.json: {exc}") from exc
    old = str(data.get("version") or "").strip()
    data["version"] = new_v
    changed = old != new_v
    if changed and not dry_run:
        VERSION_JSON.write_text(
            json.dumps(data, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    return changed


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Bump SemVer do projeto (Desktop/API/manifest).")
    p.add_argument("kind", nargs="?", choices=("major", "minor", "patch"), help="Tipo de incremento")
    p.add_argument("--set", dest="set_version", help="Define versao exata (x.y.z)")
    p.add_argument("--dry-run", action="store_true", help="Nao grava arquivos; apenas mostra o que faria")
    return p


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if not args.kind and not args.set_version:
        parser.error("Informe kind (major/minor/patch) ou --set x.y.z")
    if args.kind and args.set_version:
        parser.error("Use apenas um: kind ou --set")

    current_v = current_version_from_main()
    target_v = args.set_version.strip() if args.set_version else bump(current_v, args.kind)
    parse_semver(target_v)

    print(f"Versao atual : {current_v}")
    print(f"Versao alvo  : {target_v}")
    if args.dry_run:
        print("Modo dry-run : ativo (sem gravacao)")

    changed_main = replace_version_in_main(target_v, dry_run=args.dry_run)
    changed_api = replace_version_in_api_server(target_v, dry_run=args.dry_run)
    changed_manifest = replace_version_in_manifest(target_v, dry_run=args.dry_run)

    print("Arquivos atualizados:")
    print(f"- main.py: {'OK' if changed_main else 'sem alteracao'}")
    print(f"- api_server.py: {'OK' if changed_api else 'sem alteracao'}")
    print(f"- updates/version.json: {'OK' if changed_manifest else 'sem alteracao'}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        raise SystemExit(1)

