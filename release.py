#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parent
VERSION_PY = ROOT / "version.py"
API_SERVICE = ROOT / "api_service.py"
MANIFEST_JSON = ROOT / "updates" / "version.json"
CHANGELOG_TXT = ROOT / "updates" / "changelog.txt"
INSTALLER_ISS = ROOT / "installer" / "rotahub.iss"
RENDER_YAML = ROOT / "render.yaml"

SEMVER_RE = re.compile(r"^\d+\.\d+\.\d+$")
VERSION_PY_RE = re.compile(r'^(?P<prefix>\s*APP_VERSION\s*=\s*")(?P<ver>\d+\.\d+\.\d+)(".*)$', re.M)
FASTAPI_VER_RE = re.compile(r'(?P<prefix>version\s*=\s*")(?P<ver>\d+\.\d+\.\d+)(")')
INNO_VER_RE = re.compile(r'^(?P<prefix>\#define\s+MyAppVersionBase\s+")(?P<ver>\d+\.\d+\.\d+)(".*)$', re.M)

DEFAULT_SETUP_URL = "https://github.com/andersoncantenas-glitch/rotahub-api/releases/download/v{version}/RotaHubDesktop_Setup_{version}.exe"
DEFAULT_CHANGELOG_URL = "https://raw.githubusercontent.com/andersoncantenas-glitch/rotahub-api/main/updates/changelog.txt"
RELEASE_KINDS = ("patch", "minor", "major")


@dataclass(frozen=True)
class VersionSource:
    label: str
    path: Path
    version: str


@dataclass
class ReleaseSummary:
    previous_version: str
    new_version: str
    release_kind: str
    changed_files: list[str]
    git_push_status: str
    render_status: str
    installer_status: str
    mismatches: list[str]
    branch: str
    dry_run: bool


def parse_semver(value: str) -> tuple[int, int, int]:
    raw = str(value or "").strip()
    if not SEMVER_RE.fullmatch(raw):
        raise ValueError(f"Versao invalida: {value!r}. Use MAJOR.MINOR.PATCH")
    major, minor, patch = raw.split(".")
    return int(major), int(minor), int(patch)


def bump_semver(value: str, kind: str) -> str:
    major, minor, patch = parse_semver(value)
    if kind == "patch":
        return f"{major}.{minor}.{patch + 1}"
    if kind == "minor":
        return f"{major}.{minor + 1}.0"
    if kind == "major":
        return f"{major + 1}.0.0"
    raise ValueError(f"Tipo de release invalido: {kind}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


def replace_first(pattern: re.Pattern[str], content: str, new_version: str, label: str) -> str:
    new_content, count = pattern.subn(rf"\g<prefix>{new_version}\3", content, count=1)
    if count != 1:
        raise RuntimeError(f"Nao foi possivel atualizar a versao em {label}")
    return new_content


def load_version_sources() -> list[VersionSource]:
    sources: list[VersionSource] = []

    if VERSION_PY.exists():
        match = VERSION_PY_RE.search(read_text(VERSION_PY))
        if not match:
            raise RuntimeError("Nao foi possivel localizar APP_VERSION em version.py")
        sources.append(VersionSource("version.py", VERSION_PY, match.group("ver")))

    if MANIFEST_JSON.exists():
        data = json.loads(read_text(MANIFEST_JSON))
        if not isinstance(data, dict):
            raise RuntimeError("updates/version.json invalido")
        version = str(data.get("version") or "").strip()
        if version:
            sources.append(VersionSource("updates/version.json", MANIFEST_JSON, version))

    if INSTALLER_ISS.exists():
        match = INNO_VER_RE.search(read_text(INSTALLER_ISS))
        if not match:
            raise RuntimeError("Nao foi possivel localizar MyAppVersionBase em installer/rotahub.iss")
        sources.append(VersionSource("installer/rotahub.iss", INSTALLER_ISS, match.group("ver")))

    if API_SERVICE.exists():
        match = FASTAPI_VER_RE.search(read_text(API_SERVICE))
        if match:
            sources.append(VersionSource("api_service.py", API_SERVICE, match.group("ver")))

    if not sources:
        raise RuntimeError("Nenhuma fonte de versao encontrada")
    return sources


def canonical_current_version(sources: Iterable[VersionSource]) -> tuple[str, list[str]]:
    src_list = list(sources)
    parsed = [(parse_semver(src.version), src) for src in src_list]
    parsed.sort(key=lambda item: item[0], reverse=True)
    current = parsed[0][1].version
    mismatches = []
    distinct = {src.version for src in src_list}
    if len(distinct) > 1:
        for src in src_list:
            if src.version != current:
                mismatches.append(f"{src.label}={src.version}")
    return current, mismatches


def run_git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args],
        cwd=ROOT,
        check=check,
        text=True,
        capture_output=True,
    )


def current_branch() -> str:
    result = run_git("branch", "--show-current")
    branch = result.stdout.strip()
    return branch or "main"


def changed_files_for_classification() -> list[str]:
    result = run_git("status", "--short", check=False)
    files = []
    for raw_line in result.stdout.splitlines():
        line = raw_line.rstrip()
        if not line:
            continue
        files.append(line[3:].strip())
    return files


def suggest_release_kind(files: Iterable[str]) -> str:
    changed = [str(f).replace("\\", "/").lower() for f in files]
    if not changed:
        return "patch"

    major_markers = (
        "runtime_config.py",
        "environment.py",
        "database_runtime.py",
        "server.py",
        "api_server.py",
        "render.yaml",
        "config/",
    )
    minor_markers = (
        "main.py",
        "api_service.py",
        "assets/",
        "installer/",
        "scripts/",
    )

    if any(any(marker in item for marker in major_markers) for item in changed):
        return "major"
    if any(any(marker in item for marker in minor_markers) for item in changed):
        return "minor"
    return "patch"


def resolve_release_kind(cli_kind: str | None) -> str:
    if cli_kind:
        return cli_kind

    suggested = suggest_release_kind(changed_files_for_classification())
    if sys.stdin.isatty():
        print(f"Tipo sugerido: {suggested}")
        answer = input("Tipo da release [patch/minor/major] (Enter para aceitar): ").strip().lower()
        if answer:
            if answer not in RELEASE_KINDS:
                raise ValueError(f"Tipo invalido: {answer}. Use patch, minor ou major.")
            return answer
    return suggested


def update_version_py(new_version: str) -> None:
    content = read_text(VERSION_PY)
    write_text(VERSION_PY, replace_first(VERSION_PY_RE, content, new_version, "version.py"))


def update_api_service(new_version: str) -> None:
    if not API_SERVICE.exists():
        return
    content = read_text(API_SERVICE)
    if 'version="' not in content:
        return
    write_text(API_SERVICE, replace_first(FASTAPI_VER_RE, content, new_version, "api_service.py"))


def derive_setup_url(current_url: str, new_version: str) -> str:
    url = str(current_url or "").strip()
    if not url:
        return DEFAULT_SETUP_URL.format(version=new_version)

    updated = re.sub(r"/download/v\d+\.\d+\.\d+/", f"/download/v{new_version}/", url)
    updated = re.sub(r"RotaHubDesktop_Setup_\d+\.\d+\.\d+\.exe", f"RotaHubDesktop_Setup_{new_version}.exe", updated)
    if updated != url:
        return updated

    parsed = urlparse(url)
    if parsed.scheme and parsed.netloc:
        base = url.rsplit("/", 1)[0]
        return f"{base}/RotaHubDesktop_Setup_{new_version}.exe"
    return DEFAULT_SETUP_URL.format(version=new_version)


def update_manifest(new_version: str, release_kind: str) -> None:
    data: dict[str, object] = {}
    if MANIFEST_JSON.exists():
        loaded = json.loads(read_text(MANIFEST_JSON))
        if not isinstance(loaded, dict):
            raise RuntimeError("updates/version.json invalido")
        data = loaded

    setup_url = derive_setup_url(str(data.get("setup_url") or ""), new_version)
    changelog_url = str(data.get("changelog_url") or DEFAULT_CHANGELOG_URL).strip() or DEFAULT_CHANGELOG_URL
    alerts = data.get("alerts")
    alert = data.get("alert")

    manifest = {
        "version": new_version,
        "setup_url": setup_url,
        "changelog_url": changelog_url,
        "release_kind": release_kind,
        "released_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "alert": alert if isinstance(alert, str) else "",
        "alerts": alerts if isinstance(alerts, list) else [],
    }
    MANIFEST_JSON.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def update_changelog(previous_version: str, new_version: str, release_kind: str, notes: str) -> None:
    old = ""
    if CHANGELOG_TXT.exists():
        old = read_text(CHANGELOG_TXT).strip()
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = (
        f"{stamp} - v{new_version}\n"
        f"- Release automatizada ({release_kind}).\n"
        f"- Versao anterior: v{previous_version}.\n"
        f"- {notes.strip() or 'Preparacao de deploy no Render e setup para Inno Setup.'}\n"
    ).strip()
    content = entry if not old else entry + "\n\n" + old + "\n"
    write_text(CHANGELOG_TXT, content if content.endswith("\n") else content + "\n")


def update_installer(new_version: str) -> None:
    if not INSTALLER_ISS.exists():
        return
    content = read_text(INSTALLER_ISS)
    write_text(INSTALLER_ISS, replace_first(INNO_VER_RE, content, new_version, "installer/rotahub.iss"))


def capture_file_state(paths: Iterable[Path]) -> dict[Path, str]:
    state: dict[Path, str] = {}
    for path in paths:
        if path.exists():
            state[path] = read_text(path)
        else:
            state[path] = ""
    return state


def changed_files_from_state(before: dict[Path, str], after: dict[Path, str]) -> list[str]:
    changed = []
    for path in sorted(set(before) | set(after), key=lambda p: str(p).lower()):
        if before.get(path, "") != after.get(path, ""):
            changed.append(str(path.relative_to(ROOT)).replace("\\", "/"))
    return changed


def git_commit_and_push(new_version: str, dry_run: bool) -> str:
    commit_msg = f"release: versão {new_version}"
    branch = current_branch()
    if dry_run:
        return f"DRY-RUN: git add . && git commit -m \"{commit_msg}\" && git push origin {branch}"

    run_git("add", ".")
    commit = run_git("commit", "-m", commit_msg, check=False)
    if commit.returncode != 0:
        stderr = (commit.stderr or commit.stdout or "").strip()
        raise RuntimeError(f"Falha no git commit: {stderr or 'sem detalhes'}")

    push = run_git("push", "origin", branch, check=False)
    if push.returncode != 0:
        stderr = (push.stderr or push.stdout or "").strip()
        return f"FALHA: {stderr or 'git push retornou erro'}"
    return f"OK: push enviado para origin/{branch}"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Fluxo automatizado de release com semver.")
    parser.add_argument("kind", nargs="?", choices=RELEASE_KINDS, help="Tipo de release: patch, minor ou major.")
    parser.add_argument("--notes", default="", help="Observacao adicionada no changelog automatizado.")
    parser.add_argument("--dry-run", action="store_true", help="Atualiza arquivos localmente, mas nao executa git add/commit/push.")
    return parser


def ensure_repo_ready() -> None:
    if not (ROOT / ".git").exists():
        raise RuntimeError("Este script precisa ser executado na raiz de um repositorio git.")


def main() -> int:
    ensure_repo_ready()
    args = build_parser().parse_args()

    sources = load_version_sources()
    previous_version, mismatches = canonical_current_version(sources)
    release_kind = resolve_release_kind(args.kind)
    new_version = bump_semver(previous_version, release_kind)

    tracked_files = [VERSION_PY, API_SERVICE, MANIFEST_JSON, CHANGELOG_TXT, INSTALLER_ISS]
    before = capture_file_state(tracked_files)

    update_version_py(new_version)
    update_api_service(new_version)
    update_manifest(new_version, release_kind)
    update_changelog(previous_version, new_version, release_kind, args.notes)
    update_installer(new_version)

    after = capture_file_state(tracked_files)
    changed_files = changed_files_from_state(before, after)
    push_status = git_commit_and_push(new_version, args.dry_run)

    render_status = (
        "Render pronto: render.yaml encontrado; git push no branch atual deve acionar deploy se o servico estiver conectado."
        if RENDER_YAML.exists()
        else "Render: render.yaml nao encontrado; revise a configuracao de deploy."
    )
    installer_status = (
        (
            f"Inno Setup pronto: primeiro rode "
            f"'powershell -ExecutionPolicy Bypass -File .\\scripts\\build_desktop.ps1', "
            f"depois gere dist_installer/RotaHubDesktop_Setup_{new_version}.exe em installer/rotahub.iss."
        )
        if INSTALLER_ISS.exists()
        else "Inno Setup: installer/rotahub.iss nao encontrado."
    )

    summary = ReleaseSummary(
        previous_version=previous_version,
        new_version=new_version,
        release_kind=release_kind,
        changed_files=changed_files,
        git_push_status=push_status,
        render_status=render_status,
        installer_status=installer_status,
        mismatches=mismatches,
        branch=current_branch(),
        dry_run=args.dry_run,
    )

    print(f"Versao anterior: {summary.previous_version}")
    print(f"Nova versao: {summary.new_version}")
    print(f"Tipo da release: {summary.release_kind}")
    if summary.mismatches:
        print("Divergencias corrigidas:")
        for item in summary.mismatches:
            print(f"- {item}")
    print("Arquivos alterados:")
    for path in summary.changed_files:
        print(f"- {path}")
    print(f"Status do git push: {summary.git_push_status}")
    print("Instrucoes finais:")
    print(f"- Render: {summary.render_status}")
    print(f"- Inno Setup: {summary.installer_status}")
    print(f"- Branch atual: {summary.branch}")
    if summary.dry_run:
        print("- Dry-run ativo: nenhum commit/push real foi executado.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        raise SystemExit(1)
