import json
import os
import sqlite3
import sys
import urllib.error
import urllib.request


def _env(name: str, default: str = "") -> str:
    return str(os.environ.get(name, default) or "").strip()


def _upper(v) -> str:
    return str(v or "").strip().upper()


def _safe_int(v, default=0) -> int:
    try:
        return int(float(str(v).replace(",", ".").strip()))
    except Exception:
        return int(default)


def _safe_float(v, default=0.0) -> float:
    try:
        return float(str(v).replace(",", ".").strip())
    except Exception:
        return float(default)


def _post_json(url: str, payload: dict, desktop_secret: str, timeout: int = 20) -> dict:
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


def main() -> int:
    db_path = _env("ROTA_DB", r"C:\pdc_rota\rota_granja.db")
    api_base = _env("ROTA_SERVER_URL", "https://rotahub-api.onrender.com").rstrip("/")
    secret = _env("ROTA_SECRET")
    if not secret:
        print("ERRO: ROTA_SECRET não definido.")
        return 2
    if not os.path.exists(db_path):
        print(f"ERRO: Banco local não encontrado: {db_path}")
        return 2

    upsert_url = f"{api_base}/desktop/rotas/upsert"
    reconcile_url = f"{api_base}/desktop/programacoes/reconciliar-vinculos"

    sent = 0
    failed = 0

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(
        """
        SELECT *
        FROM programacoes
        WHERE UPPER(TRIM(COALESCE(status,''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')
        ORDER BY id DESC
        """
    )
    rows = cur.fetchall() or []
    print(f"Programações abertas encontradas: {len(rows)}")

    for r in rows:
        codigo = _upper(r["codigo_programacao"])
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

        payload = {
            "codigo_programacao": codigo,
            "data_criacao": str(r["data_criacao"] or ""),
            "motorista": _upper(r["motorista"]),
            "motorista_id": _safe_int(r["motorista_id"] if "motorista_id" in r.keys() else 0, 0),
            "motorista_codigo": _upper(r["motorista_codigo"] if "motorista_codigo" in r.keys() else ""),
            "codigo_motorista": _upper(r["codigo_motorista"] if "codigo_motorista" in r.keys() else ""),
            "veiculo": _upper(r["veiculo"]),
            "equipe": _upper(r["equipe"]),
            "kg_estimado": _safe_float(r["kg_estimado"], 0.0),
            "tipo_estimativa": _upper(r["tipo_estimativa"] if "tipo_estimativa" in r.keys() else "KG") or "KG",
            "caixas_estimado": _safe_int(r["caixas_estimado"] if "caixas_estimado" in r.keys() else 0, 0),
            "status": _upper(r["status"]) or "ATIVA",
            "local_rota": _upper(r["local_rota"] if "local_rota" in r.keys() else ""),
            "local_carregamento": _upper(
                (r["local_carregamento"] if "local_carregamento" in r.keys() else "")
                or (r["granja_carregada"] if "granja_carregada" in r.keys() else "")
            ),
            "adiantamento": _safe_float((r["adiantamento"] if "adiantamento" in r.keys() else 0), 0.0),
            "total_caixas": _safe_int((r["total_caixas"] if "total_caixas" in r.keys() else 0), 0),
            "quilos": _safe_float((r["quilos"] if "quilos" in r.keys() else 0.0), 0.0),
            "usuario_criacao": _upper(r["usuario_criacao"] if "usuario_criacao" in r.keys() else ""),
            "usuario_ultima_edicao": _upper(r["usuario_ultima_edicao"] if "usuario_ultima_edicao" in r.keys() else ""),
            "itens": itens,
        }

        try:
            _post_json(upsert_url, payload, secret)
            sent += 1
            print(f"OK  {codigo} ({len(itens)} itens)")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            failed += 1
            print(f"ERRO {codigo}: HTTP {exc.code} {exc.reason} {detail}")
        except Exception as exc:
            failed += 1
            print(f"ERRO {codigo}: {exc}")

    conn.close()

    try:
        out = _post_json(reconcile_url, {}, secret)
        print(f"Reconciliar vínculos: {out}")
    except Exception as exc:
        print(f"Aviso: falha ao reconciliar vínculos no servidor: {exc}")

    print(f"Concluído. Enviadas: {sent} | Falhas: {failed}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())

