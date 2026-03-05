import argparse
import json
import os
import re
import sqlite3
import sys
import urllib.error
import urllib.request
from typing import Any


def _env(name: str, default: str = "") -> str:
    return str(os.environ.get(name, default) or "").strip()


def _upper(v: Any) -> str:
    return str(v or "").strip().upper()


def _safe_int(v: Any, default: int = 0) -> int:
    try:
        return int(float(str(v).replace(",", ".").strip()))
    except Exception:
        return int(default)


def _safe_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(str(v).replace(",", ".").strip())
    except Exception:
        return float(default)


def _pick(row: sqlite3.Row, keys: list[str], fallback: Any = "") -> Any:
    for k in keys:
        if k in row.keys():
            v = row[k]
            if v is not None and str(v).strip() != "":
                return v
    return fallback


def _post_json(url: str, payload: dict, desktop_secret: str, timeout: int = 30) -> dict:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        method="POST",
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "X-Desktop-Secret": desktop_secret,
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = (resp.read() or b"").decode("utf-8", errors="ignore")
        return json.loads(body) if body else {}


def _first_name(txt: str) -> str:
    txt = str(txt or "").strip().upper()
    if not txt:
        return ""
    return txt.split()[0]


def _build_ajudantes_map(cur: sqlite3.Cursor) -> dict[str, str]:
    out: dict[str, str] = {}
    try:
        cur.execute("SELECT id, COALESCE(nome,''), COALESCE(sobrenome,'') FROM ajudantes")
        for r in cur.fetchall() or []:
            nome = f"{str(r[1] or '').strip()} {str(r[2] or '').strip()}".strip().upper()
            out[str(int(r[0]))] = _first_name(nome)
    except Exception:
        pass
    return out


def _resolve_equipe_nomes(raw: str, ajudantes_map: dict[str, str]) -> str:
    txt = str(raw or "").strip()
    if not txt:
        return ""
    parts = [p.strip() for p in re.split(r"[|/,;]+", txt) if p.strip()]
    nomes: list[str] = []
    for p in parts:
        nome = ajudantes_map.get(p, _first_name(p))
        if nome and nome not in nomes:
            nomes.append(nome)
    return " / ".join(nomes) if nomes else _upper(txt)


def sync_motoristas(cur: sqlite3.Cursor, api_base: str, secret: str) -> tuple[int, int]:
    cur.execute(
        """
        SELECT
            COALESCE(codigo,'') AS codigo,
            COALESCE(nome,'') AS nome,
            COALESCE(telefone,'') AS telefone,
            COALESCE(cpf,'') AS cpf,
            COALESCE(status,'ATIVO') AS status,
            COALESCE(senha,'') AS senha,
            COALESCE(acesso_liberado,1) AS acesso_liberado
        FROM motoristas
        WHERE TRIM(COALESCE(codigo,'')) <> ''
        ORDER BY id
        """
    )
    rows = cur.fetchall() or []
    ok = 0
    fail = 0
    for r in rows:
        payload = {
            "codigo": _upper(r["codigo"]),
            "nome": _upper(r["nome"]),
            "telefone": str(r["telefone"] or "").strip(),
            "cpf": str(r["cpf"] or "").strip(),
            "status": _upper(r["status"] or "ATIVO"),
            "senha": str(r["senha"] or "").strip() or None,
            "acesso_liberado": bool(_safe_int(r["acesso_liberado"], 1)),
            "acesso_liberado_por": "SYNC_TUDO",
            "acesso_obs": "Sincronizado em lote (desktop -> nuvem)",
        }
        try:
            _post_json(f"{api_base}/desktop/cadastros/motoristas/upsert", payload, secret)
            ok += 1
        except Exception as exc:
            fail += 1
            print(f"[ERRO] motorista {payload['codigo']}: {exc}")
    return ok, fail


def sync_veiculos(cur: sqlite3.Cursor, api_base: str, secret: str) -> tuple[int, int]:
    cur.execute(
        """
        SELECT
            COALESCE(placa,'') AS placa,
            COALESCE(modelo,'') AS modelo,
            COALESCE(capacidade_cx,0) AS capacidade_cx
        FROM veiculos
        WHERE TRIM(COALESCE(placa,'')) <> ''
        ORDER BY id
        """
    )
    rows = cur.fetchall() or []
    ok = 0
    fail = 0
    for r in rows:
        payload = {
            "placa": _upper(r["placa"]),
            "modelo": _upper(r["modelo"]),
            "capacidade_cx": _safe_int(r["capacidade_cx"], 0),
        }
        try:
            _post_json(f"{api_base}/desktop/cadastros/veiculos/upsert", payload, secret)
            ok += 1
        except Exception as exc:
            fail += 1
            print(f"[ERRO] veiculo {payload['placa']}: {exc}")
    return ok, fail


def sync_ajudantes(cur: sqlite3.Cursor, api_base: str, secret: str) -> tuple[int, int]:
    cur.execute(
        """
        SELECT
            COALESCE(nome,'') AS nome,
            COALESCE(sobrenome,'') AS sobrenome,
            COALESCE(telefone,'') AS telefone,
            COALESCE(status,'ATIVO') AS status
        FROM ajudantes
        WHERE TRIM(COALESCE(nome,'')) <> ''
        ORDER BY id
        """
    )
    rows = cur.fetchall() or []
    ok = 0
    fail = 0
    for r in rows:
        payload = {
            "nome": _upper(r["nome"]),
            "sobrenome": _upper(r["sobrenome"]),
            "telefone": str(r["telefone"] or "").strip(),
            "status": _upper(r["status"] or "ATIVO"),
        }
        try:
            _post_json(f"{api_base}/desktop/cadastros/ajudantes/upsert", payload, secret)
            ok += 1
        except Exception as exc:
            fail += 1
            print(f"[ERRO] ajudante {payload['nome']} {payload['sobrenome']}: {exc}")
    return ok, fail


def sync_clientes(cur: sqlite3.Cursor, api_base: str, secret: str) -> tuple[int, int]:
    cur.execute(
        """
        SELECT
            COALESCE(cod_cliente,'') AS cod_cliente,
            COALESCE(nome_cliente,'') AS nome_cliente,
            COALESCE(endereco,'') AS endereco,
            COALESCE(telefone,'') AS telefone,
            COALESCE(vendedor,'') AS vendedor
        FROM clientes
        WHERE TRIM(COALESCE(cod_cliente,'')) <> ''
        ORDER BY id
        """
    )
    rows = cur.fetchall() or []
    ok = 0
    fail = 0
    for r in rows:
        payload = {
            "cod_cliente": _upper(r["cod_cliente"]),
            "nome_cliente": _upper(r["nome_cliente"]),
            "endereco": _upper(r["endereco"]),
            "telefone": str(r["telefone"] or "").strip(),
            "vendedor": _upper(r["vendedor"]),
        }
        try:
            _post_json(f"{api_base}/desktop/cadastros/clientes/upsert", payload, secret)
            ok += 1
        except Exception as exc:
            fail += 1
            print(f"[ERRO] cliente {payload['cod_cliente']}: {exc}")
    return ok, fail


def sync_programacoes(cur: sqlite3.Cursor, api_base: str, secret: str, include_all: bool = False) -> tuple[int, int, int]:
    where = ""
    if not include_all:
        where = "WHERE UPPER(TRIM(COALESCE(status,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')"
    cur.execute(
        f"""
        SELECT *
        FROM programacoes
        {where}
        ORDER BY id DESC
        """
    )
    rows = cur.fetchall() or []
    ajudantes_map = _build_ajudantes_map(cur)
    ok = 0
    fail = 0
    skipped = 0
    for r in rows:
        codigo = _upper(r["codigo_programacao"] if "codigo_programacao" in r.keys() else "")
        if not codigo:
            continue
        cur.execute(
            """
            SELECT cod_cliente, nome_cliente, qnt_caixas, kg, preco, endereco, vendedor, pedido, produto
            FROM programacao_itens
            WHERE codigo_programacao=?
            ORDER BY id
            """,
            (codigo,),
        )
        itens_rows = cur.fetchall() or []
        itens = []
        for it in itens_rows:
            itens.append(
                {
                    "cod_cliente": _upper(it["cod_cliente"]),
                    "nome_cliente": _upper(it["nome_cliente"]),
                    "qnt_caixas": _safe_int(it["qnt_caixas"], 0),
                    "kg": _safe_float(it["kg"], 0.0),
                    "preco": _safe_float(it["preco"], 0.0),
                    "endereco": _upper(it["endereco"]),
                    "vendedor": _upper(it["vendedor"]),
                    "pedido": _upper(it["pedido"]),
                    "produto": _upper(it["produto"]),
                    "obs": "",
                }
            )

        equipe_raw = str(r["equipe"] if "equipe" in r.keys() else "")
        equipe_nomes = _resolve_equipe_nomes(equipe_raw, ajudantes_map)
        payload = {
            "codigo_programacao": codigo,
            "data_criacao": str(r["data_criacao"] if "data_criacao" in r.keys() else ""),
            "motorista": _upper(r["motorista"] if "motorista" in r.keys() else ""),
            "motorista_id": _safe_int(r["motorista_id"] if "motorista_id" in r.keys() else 0, 0),
            "motorista_codigo": _upper(r["motorista_codigo"] if "motorista_codigo" in r.keys() else ""),
            "codigo_motorista": _upper(r["codigo_motorista"] if "codigo_motorista" in r.keys() else ""),
            "veiculo": _upper(r["veiculo"] if "veiculo" in r.keys() else ""),
            "equipe": equipe_nomes or _upper(equipe_raw),
            "kg_estimado": _safe_float(r["kg_estimado"] if "kg_estimado" in r.keys() else 0.0, 0.0),
            "tipo_estimativa": _upper(r["tipo_estimativa"] if "tipo_estimativa" in r.keys() else "KG") or "KG",
            "caixas_estimado": _safe_int(r["caixas_estimado"] if "caixas_estimado" in r.keys() else 0, 0),
            "status": _upper(r["status"] if "status" in r.keys() else "ATIVA") or "ATIVA",
            "local_rota": _upper(_pick(r, ["local_rota", "tipo_rota", "local"], "")),
            "local_carregamento": _upper(
                _pick(r, ["local_carregamento", "carregamento", "granja_carregada", "local_carregado"], "")
            ),
            "adiantamento": _safe_float(r["adiantamento"] if "adiantamento" in r.keys() else 0.0, 0.0),
            "total_caixas": _safe_int(
                _pick(r, ["total_caixas", "caixas_estimado", "qnt_cx_carregada", "caixas_carregadas"], 0), 0
            ),
            "quilos": _safe_float(r["quilos"] if "quilos" in r.keys() else 0.0, 0.0),
            "usuario_criacao": _upper(r["usuario_criacao"] if "usuario_criacao" in r.keys() else ""),
            "usuario_ultima_edicao": _upper(r["usuario_ultima_edicao"] if "usuario_ultima_edicao" in r.keys() else ""),
            "itens": itens,
        }
        try:
            _post_json(f"{api_base}/desktop/rotas/upsert", payload, secret)
            ok += 1
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="ignore")
            if exc.code == 409:
                skipped += 1
                print(f"[SKIP 409] {codigo}: {body}")
            else:
                fail += 1
                print(f"[ERRO] programacao {codigo}: HTTP {exc.code} {exc.reason} {body}")
        except Exception as exc:
            fail += 1
            print(f"[ERRO] programacao {codigo}: {exc}")
    try:
        out = _post_json(f"{api_base}/desktop/programacoes/reconciliar-vinculos", {}, secret)
        print(f"Reconciliar vínculos: {out}")
    except Exception as exc:
        print(f"Aviso: falha ao reconciliar vínculos: {exc}")
    return ok, fail, skipped


def main() -> int:
    parser = argparse.ArgumentParser(description="Sincronização completa desktop -> nuvem (cadastros + programações).")
    parser.add_argument("--all", action="store_true", help="Sincroniza todas as programações (inclui finalizadas/canceladas).")
    args = parser.parse_args()

    db_path = _env("ROTA_DB", r"C:\rotahub\rota_granja.db")
    api_base = _env("ROTA_SERVER_URL", "https://rotahub-api.onrender.com").rstrip("/")
    secret = _env("ROTA_SECRET")

    if not secret:
        print("ERRO: ROTA_SECRET não definido.")
        return 2
    if not os.path.exists(db_path):
        print(f"ERRO: banco não encontrado: {db_path}")
        return 2

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    print(f"DB: {db_path}")
    print(f"API: {api_base}")
    print("")

    m_ok, m_fail = sync_motoristas(cur, api_base, secret)
    v_ok, v_fail = sync_veiculos(cur, api_base, secret)
    a_ok, a_fail = sync_ajudantes(cur, api_base, secret)
    c_ok, c_fail = sync_clientes(cur, api_base, secret)
    p_ok, p_fail, p_skip = sync_programacoes(cur, api_base, secret, include_all=args.all)
    conn.close()

    print("")
    print("=== RESUMO ===")
    print(f"Motoristas: OK {m_ok} | Falhas {m_fail}")
    print(f"Veículos:   OK {v_ok} | Falhas {v_fail}")
    print(f"Ajudantes:  OK {a_ok} | Falhas {a_fail}")
    print(f"Clientes:   OK {c_ok} | Falhas {c_fail}")
    print(f"Programações: OK {p_ok} | Falhas {p_fail} | Skipped409 {p_skip}")

    total_fail = m_fail + v_fail + a_fail + c_fail + p_fail
    return 0 if total_fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())

