"""
Serviço de gerenciamento de permissões dos usuários.
Permite conferir, da e revogar permissões de usuários no sistema.
"""

from app.db.connection import get_db


def _service_result(*, ok: bool, data=None, error: str = None, source: str = "local"):
    return {
        "ok": bool(ok),
        "data": data,
        "error": str(error) if error else None,
        "source": str(source or "local"),
    }


def listar_permissoes_disponiveis():
    """Lista todas as permissões disponíveis no sistema, organizadas por módulo."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, modulo, nome_permissao, descricao, ativo
                FROM permissoes
                ORDER BY modulo, nome_permissao
            """)
            permissoes = [
                {
                    "id": row[0],
                    "modulo": row[1],
                    "nome": row[2],
                    "descricao": row[3],
                    "ativo": bool(row[4])
                }
                for row in cur.fetchall()
            ]
            return _service_result(ok=True, data=permissoes)
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def listar_permissoes_usuario(usuario_id: int):
    """Lista as permissões concedidas a um usuário específico."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT p.id, p.modulo, p.nome_permissao, p.descricao, up.concedida_em, up.concedida_por
                FROM usuario_permissoes up
                JOIN permissoes p ON up.permissao_id = p.id
                WHERE up.usuario_id = ?
                ORDER BY p.modulo, p.nome_permissao
            """, (usuario_id,))
            
            permissoes = [
                {
                    "id": row[0],
                    "modulo": row[1],
                    "nome": row[2],
                    "descricao": row[3],
                    "concedida_em": row[4],
                    "concedida_por": row[5]
                }
                for row in cur.fetchall()
            ]
            return _service_result(ok=True, data=permissoes)
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def conceder_permissao(usuario_id: int, permissao_id: int, concedida_por: str = "ADMIN"):
    """Concede uma permissão a um usuário."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            
            # Verificar se usuário existe
            cur.execute("SELECT id FROM usuarios WHERE id=? LIMIT 1", (usuario_id,))
            if not cur.fetchone():
                return _service_result(ok=False, error="Usuário não encontrado")
            
            # Verificar se permissão existe
            cur.execute("SELECT id FROM permissoes WHERE id=? LIMIT 1", (permissao_id,))
            if not cur.fetchone():
                return _service_result(ok=False, error="Permissão não encontrada")
            
            # Inserir associação
            cur.execute("""
                INSERT OR IGNORE INTO usuario_permissoes (usuario_id, permissao_id, concedida_por)
                VALUES (?, ?, ?)
            """, (usuario_id, permissao_id, concedida_por))
            conn.commit()
            
            return _service_result(ok=True, data={"usuario_id": usuario_id, "permissao_id": permissao_id})
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def revogar_permissao(usuario_id: int, permissao_id: int):
    """Revoga uma permissão de um usuário."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                DELETE FROM usuario_permissoes
                WHERE usuario_id=? AND permissao_id=?
            """, (usuario_id, permissao_id))
            conn.commit()
            
            linhas_afetadas = cur.rowcount
            if linhas_afetadas == 0:
                return _service_result(ok=False, error="Permissão não encontrada para este usuário")
            
            return _service_result(ok=True, data={"linhas_afetadas": linhas_afetadas})
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def usuario_tem_permissao(usuario_id: int, modulo: str, nome_permissao: str) -> bool:
    """Verifica se um usuário tem uma permissão específica."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT 1 FROM usuario_permissoes up
                JOIN permissoes p ON up.permissao_id = p.id
                WHERE up.usuario_id = ? AND p.modulo = ? AND p.nome_permissao = ? AND p.ativo = 1
                LIMIT 1
            """, (usuario_id, modulo, nome_permissao))
            return cur.fetchone() is not None
    except Exception:
        return False


def listar_usuarios_com_modulo(modulo: str):
    """Lista usuários que têm pelo menos uma permissão em um módulo."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT DISTINCT u.id, u.nome, COUNT(DISTINCT up.permissao_id) as qtd_permissoes
                FROM usuarios u
                JOIN usuario_permissoes up ON u.id = up.usuario_id
                JOIN permissoes p ON up.permissao_id = p.id
                WHERE p.modulo = ?
                GROUP BY u.id, u.nome
                ORDER BY u.nome
            """, (modulo,))
            
            usuarios = [
                {
                    "id": row[0],
                    "nome": row[1],
                    "qtd_permissoes": row[2]
                }
                for row in cur.fetchall()
            ]
            return _service_result(ok=True, data=usuarios)
    except Exception as e:
        return _service_result(ok=False, error=str(e))


def atribuir_permissoes_por_perfil(usuario_id: int, perfil: str, usuario_admin: str = "ADMIN"):
    """Atribui permissões a um usuário baseado no perfil selecionado.
    
    Perfis disponíveis:
    - ADMIN: Todas as permissões
    - GERENTE: Permissões de gerenciamento (criar, editar, finalizar)
    - OPERADOR: Permissões básicas (visualizar, criar)
    - VISUALIZADOR: Apenas visualização
    """
    try:
        perfil = str(perfil).upper().strip()
        
        # Mapear perfis para permissões por módulo
        perfil_perms = {
            "ADMIN": {
                "programacoes": ["visualizar", "criar", "editar", "deletar", "finalizar"],
                "prestacao": ["gerar", "editar", "fechar"],
                "cadastros": ["gerenciar_clientes", "gerenciar_motoristas", "gerenciar_vendedores"],
                "relatorios": ["gerar", "exportar_dados"],
                "sistema": ["gerenciar_usuarios", "acessar_ferramentas", "fazer_backup", "restaurar_backup", "limpar_logs", "ver_configuracoes", "editar_configuracoes"],
            },
            "GERENTE": {
                "programacoes": ["visualizar", "criar", "editar", "finalizar"],
                "prestacao": ["gerar", "editar"],
                "cadastros": ["gerenciar_clientes", "gerenciar_motoristas"],
                "relatorios": ["gerar", "exportar_dados"],
                "sistema": ["ver_configuracoes"],
            },
            "OPERADOR": {
                "programacoes": ["visualizar", "criar"],
                "prestacao": ["gerar"],
                "cadastros": ["gerenciar_clientes"],
                "relatorios": ["gerar"],
            },
            "VISUALIZADOR": {
                "programacoes": ["visualizar"],
                "prestacao": [],
                "relatorios": ["gerar"],
            },
        }
        
        perms = perfil_perms.get(perfil, {})
        if not perms:
            return _service_result(ok=False, error=f"Perfil '{perfil}' não reconhecido")
        
        with get_db() as conn:
            cur = conn.cursor()
            
            # Primeiro, revogar todas as permissões do usuário
            cur.execute("DELETE FROM usuario_permissoes WHERE usuario_id=?", (usuario_id,))
            
            # Depois, conceder as permissões do perfil
            for modulo, nome_perms in perms.items():
                for nome_perm in nome_perms:
                    cur.execute("""
                        SELECT id FROM permissoes
                        WHERE LOWER(modulo)=LOWER(?) AND LOWER(nome_permissao)=LOWER(?)
                    """, (modulo, nome_perm))
                    perm_id = cur.fetchone()
                    if perm_id:
                        perm_id = perm_id[0]
                        cur.execute("""
                            INSERT OR IGNORE INTO usuario_permissoes (usuario_id, permissao_id, concedida_em, concedida_por)
                            VALUES (?, ?, datetime('now'), ?)
                        """, (usuario_id, perm_id, usuario_admin))
            
            conn.commit()
        
        return _service_result(ok=True, data={"usuario_id": usuario_id, "perfil": perfil, "permissoes_atribuidas": len(perms)})
    except Exception as e:
        return _service_result(ok=False, error=str(e))


__all__ = [
    "listar_permissoes_disponiveis",
    "listar_permissoes_usuario",
    "conceder_permissao",
    "revogar_permissao",
    "usuario_tem_permissao",
    "listar_usuarios_com_modulo",
    "atribuir_permissoes_por_perfil",
]
