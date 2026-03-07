#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
VERSION_PY = ROOT / "version.py"
API_SERVER = ROOT / "api_server.py"
MANIFEST_JSON = ROOT / "updates" / "version.json"
INSTALLER_ISS = ROOT / "installer" / "rotahub.iss"

SEMVER_RE = re.compile(r"^\d+\.\d+\.\d+$")
VERSION_PY_RE = re.compile(r'^(?P<prefix>\s*APP_VERSION\s*=\s*")(?P<ver>\d+\.\d+\.\d+)(".*)$', re.M)
FASTAPI_VER_RE = re.compile(r'(?P<prefix>version\s*=\s*")(?P<ver>\d+\.\d+\.\d+)(")')
INNO_VER_RE = re.compile(r'^(?P<prefix>\#define\s+MyAppVersionBase\s+")(?P<ver>\d+\.\d+\.\d+)(".*)$', re.M)


def parse_semver(value: str) -> tuple[int, int, int]:
    if not SEMVER_RE.fullmatch(str(value or "").strip()):
        raise ValueError(f"Versao invalida: {value!r}. Use MAJOR.MINOR.PATCH")
    major, minor, patch = str(value).split(".")
    return int(major), int(minor), int(patch)


def bump_semver(value: str, kind: str) -> str:
    major, minor, patch = parse_semver(value)
    if kind == "patch":
        return f"{major}.{minor}.{patch + 1}"
    if kind == "minor":
        return f"{major}.{minor + 1}.0"
    if kind == "major":
        return f"{major + 1}.0.0"
    raise ValueError(f"Tipo de bump invalido: {kind}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


def current_version() -> str:
    content = read_text(VERSION_PY)
    match = VERSION_PY_RE.search(content)
    if not match:
        raise RuntimeError("Nao foi possivel localizar APP_VERSION em version.py")
    return match.group("ver")


def replace_first(pattern: re.Pattern[str], content: str, new_version: str, label: str) -> str:
    new_content, count = pattern.subn(rf"\g<prefix>{new_version}\3", content, count=1)
    if count != 1:
        raise RuntimeError(f"Nao foi possivel atualizar a versao em {label}")
    return new_content


def update_version_py(new_version: str) -> None:
    content = read_text(VERSION_PY)
    write_text(VERSION_PY, replace_first(VERSION_PY_RE, content, new_version, "version.py"))


def update_api_server(new_version: str) -> None:
    content = read_text(API_SERVER)
    if "version=APP_VERSION" in content:
        return
    write_text(API_SERVER, replace_first(FASTAPI_VER_RE, content, new_version, "api_server.py"))


def update_manifest(new_version: str) -> None:
    if not MANIFEST_JSON.exists():
        return
    data = json.loads(read_text(MANIFEST_JSON))
    if not isinstance(data, dict):
        raise RuntimeError("updates/version.json invalido")
    data["version"] = new_version
    MANIFEST_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def update_installer(new_version: str) -> None:
    if not INSTALLER_ISS.exists():
        return
    content = read_text(INSTALLER_ISS)
    write_text(INSTALLER_ISS, replace_first(INNO_VER_RE, content, new_version, "installer/rotahub.iss"))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Incrementa a versao semantica central do projeto.")
    parser.add_argument("kind", choices=("patch", "minor", "major"), help="Tipo de incremento da versao")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    previous = current_version()
    new_version = bump_semver(previous, args.kind)

    update_version_py(new_version)
    update_api_server(new_version)
    update_manifest(new_version)
    update_installer(new_version)

    print(f"Versao anterior: {previous}")
    print(f"Nova versao: {new_version}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        raise SystemExit(1)
