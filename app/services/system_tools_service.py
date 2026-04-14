"""
Serviço de ferramentas do sistema.
Backup, logs, limpeza de dados e informações do sistema.
"""

import os
import shutil
import sqlite3
from datetime import datetime
from app.db.connection import get_db


def _service_result(*, ok: bool, data=None, error: str = None, source: str = "local"):
    return {
        "ok": bool(ok),
        "data": data,
        "error": str(error) if error else None,
        "source": str(source or "local"),
    }


def registrar_acao_sistema(tipo_acao: str, descricao: str, usuario: str = "SISTEMA", 
                          status: str = "OK", resultado: str = None):
    """Registra uma ação de ferramentas do sistema no log."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO sistema_logs 
                (tipo_acao, descricao, usuario, status, resultado_texto)
                VALUES (?, ?, ?, ?, ?)
            """, (tipo_acao, descricao, usuario, status, resultado))
            conn.commit()
            return _service_result(ok=True, data={"logged": True})
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def listar_logs_sistema(limite: int = 100):
    """Retorna os últimos logs de ações do sistema."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, tipo_acao, descricao, usuario, status, resultado_texto, 
                       COALESCE(executado_em, '') as executado_em
                FROM sistema_logs
                ORDER BY id DESC
                LIMIT ?
            """, (max(1, min(limite, 500)),))
            
            logs = [
                {
                    "id": row[0],
                    "tipo_acao": row[1],
                    "descricao": row[2],
                    "usuario": row[3],
                    "status": row[4],
                    "resultado": row[5],
                    "executado_em": row[6]
                }
                for row in cur.fetchall()
            ]
            return _service_result(ok=True, data=logs)
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def limpar_logs_sistema(dias: int = 30):
    """Limpa logs antigos (mais antigos que N dias)."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute(f"""
                DELETE FROM sistema_logs
                WHERE executado_em < datetime('now', '-{max(1, dias)} days')
            """)
            linhas_afetadas = cur.rowcount
            conn.commit()
            
            registrar_acao_sistema(
                "LIMPEZA_LOGS", 
                f"Removidos {linhas_afetadas} logs com mais de {dias} dias",
                status="OK"
            )
            
            return _service_result(ok=True, data={"linhas_deletadas": linhas_afetadas})
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def fazer_backup_banco(usuario: str = "SISTEMA"):
    """Faz backup do banco de dados para a pasta de backups."""
    try:
        # Definir caminho
        pasta_backup = os.path.join(os.getcwd(), "backup")
        if not os.path.exists(pasta_backup):
            os.makedirs(pasta_backup)
        
        # Nome do arquivo com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo_backup = os.path.join(pasta_backup, f"banco_de_dados_{timestamp}.db")
        
        # Copiar arquivo
        db_principal = "banco.db"
        if not os.path.exists(db_principal):
            return _service_result(ok=False, error="Banco de dados principal não encontrado")
        
        shutil.copy2(db_principal, arquivo_backup)
        
        # Registrar ação
        registrar_acao_sistema(
            "BACKUP",
            f"Backup criado: {os.path.basename(arquivo_backup)}",
            usuario=usuario,
            status="OK",
            resultado=arquivo_backup
        )
        
        return _service_result(ok=True, data={
            "arquivo": arquivo_backup,
            "tamanho_kb": round(os.path.getsize(arquivo_backup) / 1024, 2),
            "timestamp": timestamp
        })
    except Exception as e:
        registrar_acao_sistema("BACKUP", f"Erro ao fazer backup", usuario=usuario, 
                              status="ERRO", resultado=str(e))
        return _service_result(ok=False, error=str(e))


def listar_backups():
    """Lista todos os backups disponíveis."""
    try:
        pasta_backup = os.path.join(os.getcwd(), "backup")
        if not os.path.exists(pasta_backup):
            return _service_result(ok=True, data=[])
        
        backups = []
        for arquivo in sorted(os.listdir(pasta_backup), reverse=True):
            if arquivo.startswith("banco_de_dados_") and arquivo.endswith(".db"):
                caminho = os.path.join(pasta_backup, arquivo)
                tamanho_kb = round(os.path.getsize(caminho) / 1024, 2)
                data_mod = datetime.fromtimestamp(os.path.getmtime(caminho))
                
                backups.append({
                    "arquivo": arquivo,
                    "caminho": caminho,
                    "tamanho_kb": tamanho_kb,
                    "data_criacao": data_mod.isoformat()
                })
        
        return _service_result(ok=True, data=backups)
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def restaurar_backup(arquivo_backup: str, usuario: str = "SISTEMA"):
    """Restaura um backup do banco de dados."""
    try:
        if not arquivo_backup or not os.path.exists(arquivo_backup):
            return _service_result(ok=False, error="Arquivo de backup não encontrado")
        
        db_principal = "banco.db"
        
        # Fazer backup do estado atual antes de restaurar
        timestamp_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
        db_backup_anterior = f"{db_principal}.pre_restore_{timestamp_atual}"
        if os.path.exists(db_principal):
            shutil.copy2(db_principal, db_backup_anterior)
        
        # Restaurar
        shutil.copy2(arquivo_backup, db_principal)
        
        # Registrar ação
        registrar_acao_sistema(
            "RESTAURACAO",
            f"Banco restaurado de: {os.path.basename(arquivo_backup)}",
            usuario=usuario,
            status="OK",
            resultado=f"Backup anterior guardado como: {db_backup_anterior}"
        )
        
        return _service_result(ok=True, data={
            "arquivo_restaurado": arquivo_backup,
            "backup_anterior": db_backup_anterior
        })
    except Exception as e:
        registrar_acao_sistema("RESTAURACAO", f"Erro ao restaurar backup", 
                              usuario=usuario, status="ERRO", resultado=str(e))
        return _service_result(ok=False, error=str(e))


def informacoes_sistema():
    """Retorna informações gerais do sistema."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            
            # Contar registros nas principais tabelas
            stats = {}
            tabelas = ["usuarios", "motoristas", "vendedores", "clientes", 
                      "veiculos", "programacoes", "programacao_itens"]
            
            for tabela in tabelas:
                try:
                    cur.execute(f"SELECT COUNT(*) FROM {tabela}")
                    stats[tabela] = cur.fetchone()[0]
                except:
                    pass
            
            # Tamanho do banco
            db_size_kb = round(os.path.getsize("banco.db") / 1024, 2) if os.path.exists("banco.db") else 0
            
            # Última ação registrada
            cur.execute("SELECT executado_em FROM sistema_logs ORDER BY id DESC LIMIT 1")
            ultima_acao = cur.fetchone()
            ultima_acao_em = ultima_acao[0] if ultima_acao else None
            
            return _service_result(ok=True, data={
                "registros_por_tabela": stats,
                "tamanho_banco_kb": db_size_kb,
                "ultima_acao_em": ultima_acao_em,
                "data_hora_atual": datetime.now().isoformat()
            })
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def verificar_integridade_banco():
    """Verifica a integridade do banco de dados SQLite."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA integrity_check")
            resultado = cur.fetchone()
            
            if resultado and resultado[0] == "ok":
                registrar_acao_sistema(
                    "VERIFICACAO",
                    "Integridade do banco verificada com sucesso",
                    status="OK"
                )
                return _service_result(ok=True, data={"integridade": "OK"})
            else:
                erro = resultado[0] if resultado else "Erro desconhecido"
                registrar_acao_sistema(
                    "VERIFICACAO",
                    f"Problemas encontrados na integridade: {erro}",
                    status="AVISO"
                )
                return _service_result(ok=False, error=f"Integridade comprometida: {erro}")
    except Exception as e:
        registrar_acao_sistema(
            "VERIFICACAO",
            f"Erro ao verificar integridade",
            status="ERRO"
        )
        return _service_result(ok=False, error=str(e))


__all__ = [
    "registrar_acao_sistema",
    "listar_logs_sistema",
    "limpar_logs_sistema",
    "fazer_backup_banco",
    "listar_backups",
    "restaurar_backup",
    "informacoes_sistema",
    "verificar_integridade_banco",
]
