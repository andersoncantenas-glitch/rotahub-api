import os
from typing import Dict, List

import httpx
from fastapi import FastAPI, HTTPException, Request, Response


TARGET_SERVER = os.environ.get("ROTA_SERVER_URL", "http://127.0.0.1:8000").rstrip("/")
if not TARGET_SERVER:
    raise RuntimeError(
        "ROTA_SERVER_URL não configurado. Defina a variável de ambiente com a URL completa da API "
        "(ex: http://localhost:8000)."
    )

PROXY_METHODS: List[str] = ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"]

app = FastAPI(
    title="Rota Granja API (shim)",
    version="1.0.0",
    description=f"Proxy temporário para {TARGET_SERVER}",
)

client = httpx.AsyncClient(timeout=20.0)


def _sanitize_headers(headers: Dict[str, str]) -> Dict[str, str]:
    forbidden = {"host", "content-length", "transfer-encoding", "connection"}
    return {k: v for k, v in headers.items() if k.lower() not in forbidden}


async def _proxy_request(path: str, request: Request) -> Response:
    target_path = path.strip("/")
    target_url = f"{TARGET_SERVER}/{target_path}" if target_path else TARGET_SERVER

    try:
        response = await client.request(
            request.method,
            target_url,
            params=request.query_params,
            headers=_sanitize_headers(dict(request.headers)),
            content=await request.body(),
        )
    except httpx.RequestError as exc:
        raise HTTPException(status_code=502, detail=f"Erro ao encaminhar para {TARGET_SERVER}: {exc}")

    response_headers = _sanitize_headers(dict(response.headers))
    return Response(
        content=response.content,
        status_code=response.status_code,
        headers=response_headers,
        media_type=response.headers.get("content-type"),
    )


@app.api_route("/", methods=PROXY_METHODS)
@app.api_route("/{full_path:path}", methods=PROXY_METHODS)
async def proxy(request: Request, full_path: str = ""):
    return await _proxy_request(full_path, request)


@app.on_event("shutdown")
async def _shutdown_client():
    await client.aclose()
