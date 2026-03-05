# -*- coding: utf-8 -*-
"""Servidor FastAPI para o RotaHub"""
import os
import sqlite3
import logging
from fastapi import FastAPI, HTTPException, Depends
from fastapi.security import HTTPBasic, HTTPBasicCredentials
import secrets

# Configuração do logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Caminho do banco de dados
def get_db_path():
    """Retorna o caminho do banco de dados"""
    return os.environ.get("ROTA_DB", os.path.join(os.path.dirname(__file__), "rota_granja.db"))

# Inicialização do FastAPI
app = FastAPI(
    title="RotaHub API",
    version="1.0.0",
    description="API para o sistema RotaHub",
)

# Autenticação para endpoint de admin
security = HTTPBasic()
RESET_USER = "admin"
RESET_PASSWORD = os.environ.get("ROTA_RESET_PASSWORD", "super-senha-secreta-123")

def get_current_username(credentials: HTTPBasicCredentials = Depends(security)):
    """Verifica as credenciais do usuário"""
    correct_username = secrets.compare_digest(credentials.username, RESET_USER)
    correct_password = secrets.compare_digest(credentials.password, RESET_PASSWORD)
    if not (correct_username and correct_password):
        raise HTTPException(
            status_code=401,
            detail="Credenciais incorretas",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username

@app.post("/admin/reset-database", tags=["Admin"])
def reset_database_endpoint(username: str = Depends(get_current_username)):
    """
    Endpoint para apagar e recriar o banco de dados.
    Requer autenticação Basic Auth.
    """
    db_path = get_db_path()
    logging.warning(f"Requisição de reset para o banco de dados: {db_path}")
    
    try:
        if os.path.exists(db_path):
            os.remove(db_path)
            logging.info(f"Banco de dados '{db_path}' apagado com sucesso.")
            return {"message": f"Banco de dados '{os.path.basename(db_path)}' foi resetado com sucesso."}
        else:
            return {"message": "Banco de dados não encontrado. Nada a fazer."}
    except Exception as e:
        logging.error(f"Erro ao tentar resetar o banco de dados: {e}")
        raise HTTPException(status_code=500, detail=f"Erro no servidor ao resetar o banco: {e}")

# Endpoint de saúde
@app.get("/health", tags=["Health"])
def health_check():
    """Verifica se a API está funcionando"""
    return {"status": "ok"}

# Inicialização do servidor
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    print(f"Iniciando servidor FastAPI na porta {port}...")
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=False)