# backend/config/database.py
"""
Database configuration and connection management
"""
import logging
import tempfile
from pathlib import Path
from typing import AsyncGenerator
from sqlalchemy import event
from sqlalchemy import inspect, text
from sqlalchemy.ext.asyncio import AsyncSession, create_async_engine, async_sessionmaker
from sqlalchemy.orm import DeclarativeBase, Session, with_loader_criteria
from backend.config.settings import settings

logger = logging.getLogger(__name__)


def _prepare_sqlite_database_url(database_url: str) -> str:
    if "sqlite" not in database_url:
        return database_url
    prefix = "sqlite+aiosqlite:///"
    if not database_url.startswith(prefix):
        return database_url
    raw_path = database_url.removeprefix(prefix)
    if raw_path in {"", ":memory:"}:
        return database_url
    path = Path(raw_path)
    if not path.is_absolute():
        path = path.resolve()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        probe = path.parent / ".rotahub_write_probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return f"{prefix}{path.as_posix()}"
    except Exception as exc:
        fallback = Path(tempfile.gettempdir()) / "rotahub" / (path.name or "rotadb.db")
        fallback.parent.mkdir(parents=True, exist_ok=True)
        logger.warning("SQLite path %s is not writable (%s); using %s", path, exc, fallback)
        return f"{prefix}{fallback.as_posix()}"


# Async engine for SQLite/PostgreSQL
DATABASE_URL = _prepare_sqlite_database_url(settings.DATABASE_URL)
if "sqlite" in DATABASE_URL:
    engine = create_async_engine(
        DATABASE_URL,
        echo=settings.DEBUG,
    )
else:
    engine = create_async_engine(
        DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://"),
        echo=settings.DEBUG,
        pool_size=10,
        max_overflow=20,
    )

# Async session factory
async_session = async_sessionmaker(engine, expire_on_commit=False)


class Base(DeclarativeBase):
    """Base class for all database models"""
    pass


def _tenant_mapped_classes() -> list[type]:
    classes: list[type] = []
    for mapper in list(Base.registry.mappers):
        if "company_id" in mapper.columns:
            classes.append(mapper.class_)
    return classes


@event.listens_for(AsyncSession.sync_session_class, "do_orm_execute")
def _add_tenant_filter(execute_state):
    if not execute_state.is_select:
        return
    session = execute_state.session
    if session.info.get("bypass_tenant"):
        return
    company_id = session.info.get("company_id")
    if not company_id:
        return
    options = [
        with_loader_criteria(
            model,
            model.company_id == int(company_id),
            include_aliases=True,
        )
        for model in _tenant_mapped_classes()
    ]
    if options:
        execute_state.statement = execute_state.statement.options(*options)


@event.listens_for(AsyncSession.sync_session_class, "before_flush")
def _assign_tenant_on_new_objects(session, flush_context, instances):
    if session.info.get("bypass_tenant"):
        return
    company_id = session.info.get("company_id")
    if not company_id:
        return
    for obj in session.new:
        if hasattr(obj, "company_id") and not getattr(obj, "company_id", None):
            setattr(obj, "company_id", int(company_id))


def _ensure_backend_columns(sync_conn):
    inspector = inspect(sync_conn)
    table_names = set(inspector.get_table_names())

    def add_missing_columns(table_name, expected_columns):
        if table_name not in table_names:
            return
        columns = {column["name"] for column in inspector.get_columns(table_name)}
        for column, definition in expected_columns.items():
            if column not in columns:
                sync_conn.execute(text(f"ALTER TABLE {table_name} ADD COLUMN {column} {definition}"))

    def backfill_company_id(table_name, codigo_column="codigo_programacao"):
        if table_name not in table_names:
            return
        local_columns = {column["name"] for column in inspect(sync_conn).get_columns(table_name)}
        if "company_id" not in local_columns:
            return
        if codigo_column in local_columns and "programacoes" in table_names:
            sync_conn.execute(
                text(
                    f"""
                    UPDATE {table_name}
                       SET company_id=COALESCE((
                           SELECT p.company_id
                             FROM programacoes p
                            WHERE UPPER(TRIM(COALESCE(p.codigo_programacao, ''))) =
                                  UPPER(TRIM(COALESCE({table_name}.{codigo_column}, '')))
                              AND p.company_id IS NOT NULL
                            LIMIT 1
                       ), 1)
                     WHERE company_id IS NULL OR company_id=0
                    """
                )
            )
        else:
            sync_conn.execute(text(f"UPDATE {table_name} SET company_id=1 WHERE company_id IS NULL OR company_id=0"))

    def repair_clientes_table():
        nonlocal table_names
        if "clientes" not in table_names:
            return

        expected_columns = {
            "nome": "TEXT",
            "cod_cliente": "TEXT",
            "nome_cliente": "TEXT",
            "endereco": "TEXT",
            "bairro": "TEXT",
            "cidade": "TEXT",
            "uf": "TEXT",
            "telefone": "TEXT",
            "rota": "TEXT",
            "vendedor": "TEXT",
        }

        def column_map(table_name: str) -> dict[str, str]:
            local_inspector = inspect(sync_conn)
            return {str(column["name"]).lower(): str(column["name"]) for column in local_inspector.get_columns(table_name)}

        columns = column_map("clientes")

        if "id" not in columns and sync_conn.dialect.name == "sqlite":
            existing_names = set(inspect(sync_conn).get_table_names())
            backup_name = "clientes_legacy_backup"
            suffix = 1
            while backup_name in existing_names:
                suffix += 1
                backup_name = f"clientes_legacy_backup_{suffix}"

            sync_conn.execute(text(f"ALTER TABLE clientes RENAME TO {backup_name}"))
            sync_conn.execute(
                text(
                    """
                    CREATE TABLE clientes (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nome TEXT,
                        cod_cliente TEXT,
                        nome_cliente TEXT NOT NULL DEFAULT '',
                        endereco TEXT,
                        bairro TEXT,
                        cidade TEXT,
                        uf TEXT,
                        telefone TEXT,
                        rota TEXT,
                        vendedor TEXT
                    )
                    """
                )
            )

            legacy_rows = sync_conn.execute(text(f"SELECT rowid AS __rowid__, * FROM {backup_name}")).mappings().all()
            seen_codes: set[str] = set()

            def clean(value) -> str:
                return str(value or "").strip().upper()

            for row in legacy_rows:
                source = {str(key).lower(): value for key, value in dict(row).items()}

                def pick(*names: str) -> str:
                    for name in names:
                        value = source.get(name.lower())
                        if value is not None and str(value).strip():
                            return clean(value)
                    return ""

                rowid = source.get("__rowid__") or len(seen_codes) + 1
                cod_cliente = pick("cod_cliente", "codigo", "cod", "cliente")
                nome_cliente = pick("nome_cliente", "nome", "razao", "cliente_nome")
                if not cod_cliente and not nome_cliente:
                    continue
                if not cod_cliente:
                    cod_cliente = f"LEGADO-{rowid}"
                if cod_cliente in seen_codes:
                    continue
                if not nome_cliente:
                    nome_cliente = cod_cliente
                seen_codes.add(cod_cliente)

                sync_conn.execute(
                    text(
                        """
                        INSERT INTO clientes (
                            nome, cod_cliente, nome_cliente, endereco, bairro, cidade, uf, telefone, rota, vendedor
                        ) VALUES (
                            :nome, :cod_cliente, :nome_cliente, :endereco, :bairro, :cidade, :uf, :telefone, :rota, :vendedor
                        )
                        """
                    ),
                    {
                        "nome": nome_cliente,
                        "cod_cliente": cod_cliente,
                        "nome_cliente": nome_cliente,
                        "endereco": pick("endereco", "logradouro", "rua"),
                        "bairro": pick("bairro"),
                        "cidade": pick("cidade", "municipio"),
                        "uf": pick("uf", "estado"),
                        "telefone": pick("telefone", "fone", "celular", "contato"),
                        "rota": pick("rota"),
                        "vendedor": pick("vendedor", "vend", "representante"),
                    },
                )

            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_clientes_cod ON clientes(cod_cliente)"))
            table_names = set(inspect(sync_conn).get_table_names())
            return

        add_missing_columns("clientes", expected_columns)
        columns = column_map("clientes")
        if "nome_cliente" in columns and "nome" in columns:
            sync_conn.execute(
                text(
                    """
                    UPDATE clientes
                       SET nome_cliente=COALESCE(NULLIF(TRIM(nome_cliente), ''), NULLIF(TRIM(nome), ''), cod_cliente, 'CLIENTE SEM NOME')
                     WHERE nome_cliente IS NULL OR TRIM(nome_cliente)=''
                    """
                )
            )
            sync_conn.execute(
                text(
                    """
                    UPDATE clientes
                       SET nome=COALESCE(NULLIF(TRIM(nome), ''), NULLIF(TRIM(nome_cliente), ''), cod_cliente, 'CLIENTE SEM NOME')
                     WHERE nome IS NULL OR TRIM(nome)=''
                    """
                )
            )
        elif "nome_cliente" in columns:
            sync_conn.execute(
                text(
                    """
                    UPDATE clientes
                       SET nome_cliente=COALESCE(NULLIF(TRIM(nome_cliente), ''), cod_cliente, 'CLIENTE SEM NOME')
                     WHERE nome_cliente IS NULL OR TRIM(nome_cliente)=''
                    """
                )
            )
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_clientes_cod ON clientes(cod_cliente)"))

    def repair_cadastros_tables():
        cadastro_columns = {
            "motoristas": {
                "codigo": "TEXT",
                "senha": "TEXT",
                "perfil_app": "TEXT DEFAULT 'MOTORISTA'",
                "cpf": "TEXT",
                "telefone": "TEXT",
                "status": "TEXT DEFAULT 'ATIVO'",
            },
            "vendedores": {
                "codigo": "TEXT",
                "nome": "TEXT",
                "senha": "TEXT",
                "telefone": "TEXT",
                "cidade_base": "TEXT",
                "status": "TEXT DEFAULT 'ATIVO'",
            },
            "veiculos": {
                "placa": "TEXT",
                "modelo": "TEXT",
                "capacidade_cx": "INTEGER DEFAULT 0",
                "status": "TEXT DEFAULT 'ATIVO'",
            },
            "caixas": {
                "codigo": "TEXT",
                "lote": "TEXT",
                "cor": "TEXT",
                "veiculo_placa": "TEXT",
                "status": "TEXT DEFAULT 'EM_ESTOQUE'",
                "data_compra": "TEXT",
                "observacao": "TEXT",
                "company_id": "INTEGER",
            },
            "caixas_movimentos": {
                "caixa_id": "INTEGER",
                "codigo": "TEXT",
                "movimento": "TEXT",
                "veiculo_origem": "TEXT",
                "veiculo_destino": "TEXT",
                "status_origem": "TEXT",
                "status_destino": "TEXT",
                "observacao": "TEXT",
                "criado_em": "TEXT",
                "company_id": "INTEGER",
            },
            "ajudantes": {
                "nome": "TEXT",
                "sobrenome": "TEXT",
                "telefone": "TEXT",
                "status": "TEXT DEFAULT 'ATIVO'",
            },
            "fornecedores": {
                "razao_social": "TEXT",
                "nome_fantasia": "TEXT",
                "documento": "TEXT",
                "tipo_pessoa": "TEXT DEFAULT 'CNPJ'",
                "perfil_fornecedor": "TEXT DEFAULT 'OUTROS'",
                "telefone": "TEXT",
                "email": "TEXT",
                "cidade": "TEXT",
                "uf": "TEXT",
                "status": "TEXT DEFAULT 'ATIVO'",
                "observacao": "TEXT",
                "certificado_nome": "TEXT",
                "certificado_path": "TEXT",
                "certificado_status": "TEXT DEFAULT 'NAO_INSTALADO'",
                "certificado_instalado_em": "TEXT",
            },
            "produtos": {
                "codigo": "TEXT",
                "nome": "TEXT",
                "descricao": "TEXT",
                "categoria": "TEXT DEFAULT 'AVES'",
                "unidade": "TEXT DEFAULT 'KG'",
                "unidade_estoque": "TEXT DEFAULT 'KG'",
                "controla_estoque_fisico": "INTEGER DEFAULT 1",
                "controla_estoque_fiscal": "INTEGER DEFAULT 1",
                "estoque_min_kg": "REAL DEFAULT 0",
                "estoque_min_caixas": "INTEGER DEFAULT 0",
                "ncm": "TEXT",
                "cest": "TEXT",
                "cfop_entrada": "TEXT",
                "cfop_saida": "TEXT",
                "ean": "TEXT",
                "custo_padrao": "REAL DEFAULT 0",
                "preco_padrao": "REAL DEFAULT 0",
                "status": "TEXT DEFAULT 'ATIVO'",
                "company_id": "INTEGER",
            },
        }
        for table_name, columns in cadastro_columns.items():
            add_missing_columns(table_name, columns)

        sync_conn.execute(
            text(
                """
                CREATE TABLE IF NOT EXISTS fornecedor_perfis (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT NOT NULL UNIQUE,
                    nome TEXT NOT NULL,
                    categoria TEXT DEFAULT 'OUTROS',
                    status TEXT DEFAULT 'ATIVO',
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
        )
        for codigo, nome, categoria in (
            ("FRANGO_VIVO", "Frango vivo", "FRANGO"),
            ("PRESTADOR_SERVICO", "Prestador de servico", "SERVICO"),
            ("PECAS", "Pecas", "MANUTENCAO"),
            ("MECANICO", "Mecanico", "MANUTENCAO"),
            ("BORRACHEIRO", "Borracheiro", "MANUTENCAO"),
            ("LAVADOR_CAIXAS", "Lavador de caixas", "OPERACIONAL"),
            ("PNEUS", "Pneus", "MANUTENCAO"),
            ("OLEO_LUBRIFICANTES", "Oleo e lubrificantes", "MANUTENCAO"),
            ("COMBUSTIVEL", "Combustivel", "OPERACIONAL"),
            ("SERVICO_SEM_NF", "Servico sem NF", "SERVICO"),
            ("OUTROS", "Outros", "OUTROS"),
        ):
            sync_conn.execute(
                text(
                    """
                    INSERT OR IGNORE INTO fornecedor_perfis (codigo, nome, categoria, status)
                    VALUES (:codigo, :nome, :categoria, 'ATIVO')
                    """
                ),
                {"codigo": codigo, "nome": nome, "categoria": categoria},
            )
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_fornecedor_perfis_status ON fornecedor_perfis(status)"))

        local_inspector = inspect(sync_conn)
        if "motoristas" in table_names:
            columns = {column["name"] for column in local_inspector.get_columns("motoristas")}
            if "codigo" in columns:
                sync_conn.execute(
                    text(
                        """
                        UPDATE motoristas
                           SET codigo='MOT-' || printf('%02d', id)
                         WHERE codigo IS NULL OR TRIM(codigo)=''
                        """
                    )
                )
            if "perfil_app" in columns:
                sync_conn.execute(
                    text(
                        """
                        UPDATE motoristas
                           SET perfil_app='MOTORISTA'
                         WHERE perfil_app IS NULL OR TRIM(perfil_app)=''
                        """
                    )
                )
            if "status" in columns:
                sync_conn.execute(
                    text(
                        """
                        UPDATE motoristas
                           SET status='ATIVO'
                         WHERE status IS NULL OR TRIM(status)=''
                        """
                    )
                )
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_motoristas_codigo ON motoristas(codigo)"))

        if "vendedores" in table_names:
            columns = {column["name"] for column in local_inspector.get_columns("vendedores")}
            if "codigo" in columns:
                sync_conn.execute(
                    text(
                        """
                        UPDATE vendedores
                           SET codigo='VEN-' || printf('%02d', id)
                         WHERE codigo IS NULL OR TRIM(codigo)=''
                        """
                    )
                )
            if "status" in columns:
                sync_conn.execute(
                    text(
                        """
                        UPDATE vendedores
                           SET status='ATIVO'
                         WHERE status IS NULL OR TRIM(status)=''
                        """
                    )
                )
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_vendedores_codigo ON vendedores(codigo)"))

        if "veiculos" in table_names:
            columns = {column["name"] for column in local_inspector.get_columns("veiculos")}
            if "capacidade_cx" in columns:
                if "capacidade_kg" in columns:
                    sync_conn.execute(
                        text(
                            """
                            UPDATE veiculos
                               SET capacidade_cx=COALESCE(capacidade_cx, capacidade_kg, 0)
                             WHERE capacidade_cx IS NULL
                            """
                        )
                    )
                sync_conn.execute(
                    text(
                        """
                        UPDATE veiculos
                           SET capacidade_cx=0
                         WHERE capacidade_cx IS NULL
                        """
                    )
                )
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_veiculos_placa ON veiculos(placa)"))

        if "caixas" in table_names:
            sync_conn.execute(
                text(
                    """
                    UPDATE caixas
                       SET codigo=UPPER(TRIM(COALESCE(NULLIF(codigo, ''), 'CX-' || printf('%05d', id)))),
                           lote=UPPER(TRIM(COALESCE(lote, ''))),
                           cor=UPPER(TRIM(COALESCE(cor, ''))),
                           veiculo_placa=UPPER(TRIM(COALESCE(veiculo_placa, ''))),
                           status=UPPER(TRIM(COALESCE(NULLIF(status, ''), 'EM_ESTOQUE')))
                    """
                )
            )
            sync_conn.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS idx_backend_caixas_codigo ON caixas(codigo)"))
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_caixas_lote ON caixas(lote)"))
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_caixas_veiculo ON caixas(veiculo_placa)"))
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_caixas_status ON caixas(status)"))

        if "caixas_movimentos" in table_names:
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_caixas_movimentos_caixa ON caixas_movimentos(caixa_id, criado_em)"))
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_caixas_movimentos_codigo ON caixas_movimentos(codigo, criado_em)"))

        if "fornecedores" in table_names:
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_fornecedores_documento ON fornecedores(documento)"))
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_fornecedores_perfil ON fornecedores(perfil_fornecedor)"))

        if "produtos" in table_names:
            sync_conn.execute(
                text(
                    """
                    UPDATE produtos
                       SET categoria=COALESCE(NULLIF(TRIM(categoria), ''), 'AVES'),
                           unidade=COALESCE(NULLIF(TRIM(unidade), ''), 'KG'),
                           unidade_estoque=COALESCE(NULLIF(TRIM(unidade_estoque), ''), 'KG'),
                           controla_estoque_fisico=COALESCE(controla_estoque_fisico, 1),
                           controla_estoque_fiscal=COALESCE(controla_estoque_fiscal, 1),
                           estoque_min_kg=COALESCE(estoque_min_kg, 0),
                           estoque_min_caixas=COALESCE(estoque_min_caixas, 0),
                           status=COALESCE(NULLIF(TRIM(status), ''), 'ATIVO')
                     WHERE categoria IS NULL OR TRIM(categoria)=''
                        OR unidade IS NULL OR TRIM(unidade)=''
                        OR unidade_estoque IS NULL OR TRIM(unidade_estoque)=''
                        OR controla_estoque_fisico IS NULL
                        OR controla_estoque_fiscal IS NULL
                        OR estoque_min_kg IS NULL
                        OR estoque_min_caixas IS NULL
                        OR status IS NULL OR TRIM(status)=''
                    """
                )
            )
            sync_conn.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS idx_backend_produtos_codigo ON produtos(codigo)"))
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_produtos_nome ON produtos(nome)"))
            sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_produtos_status ON produtos(status)"))

    if "usuarios" in table_names:
        columns = {column["name"] for column in inspector.get_columns("usuarios")}
        if "username" not in columns:
            sync_conn.execute(text("ALTER TABLE usuarios ADD COLUMN username TEXT"))
            columns.add("username")
        if "is_active" not in columns:
            default_value = "TRUE" if sync_conn.dialect.name.startswith("postgres") else "1"
            sync_conn.execute(
                text(f"ALTER TABLE usuarios ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT {default_value}")
            )
            columns.add("is_active")
        if "idade" not in columns:
            sync_conn.execute(text("ALTER TABLE usuarios ADD COLUMN idade INTEGER"))
            columns.add("idade")
        if "nome" in columns and "username" in columns:
            codigo_expr = "NULLIF(TRIM(codigo), '')" if "codigo" in columns else "NULL"
            sync_conn.execute(
                text(
                    f"""
                    UPDATE usuarios
                       SET username=COALESCE(NULLIF(TRIM(username), ''), {codigo_expr}, NULLIF(TRIM(nome), ''), 'ADMIN')
                     WHERE username IS NULL OR TRIM(username)=''
                    """
                )
            )
        if "nome" in columns:
            sync_conn.execute(
                text(
                    """
                    UPDATE usuarios
                       SET nome=COALESCE(NULLIF(TRIM(nome), ''), NULLIF(TRIM(username), ''), 'ADMIN')
                     WHERE nome IS NULL OR TRIM(nome)=''
                    """
                )
            )

    if "audit_logs" in table_names:
        columns = {column["name"] for column in inspector.get_columns("audit_logs")}
        expected_columns = {
            "company_id": "INTEGER",
            "user_id": "INTEGER",
            "actor_type": "TEXT",
            "action": "TEXT",
            "entity_type": "TEXT",
            "entity_id": "TEXT",
            "severity": "TEXT NOT NULL DEFAULT 'info'",
            "ip_address": "TEXT",
            "metadata_json": "TEXT NOT NULL DEFAULT '{}'",
            "created_at": "TEXT",
        }
        for column, definition in expected_columns.items():
            if column not in columns:
                sync_conn.execute(text(f"ALTER TABLE audit_logs ADD COLUMN {column} {definition}"))
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_audit_logs_action ON audit_logs(action)"))
        sync_conn.execute(
            text("CREATE INDEX IF NOT EXISTS idx_audit_logs_company_created ON audit_logs(company_id, created_at)")
        )

    repair_clientes_table()
    repair_cadastros_tables()

    add_missing_columns(
        "permissoes",
        {
            "modulo": "TEXT",
            "nome_permissao": "TEXT",
            "descricao": "TEXT",
            "ativo": "INTEGER DEFAULT 1",
        },
    )
    if "permissoes" in table_names:
        sync_conn.execute(
            text("CREATE INDEX IF NOT EXISTS idx_backend_permissoes_modulo ON permissoes(modulo, nome_permissao)")
        )

    add_missing_columns(
        "usuario_permissoes",
        {
            "usuario_id": "INTEGER",
            "permissao_id": "INTEGER",
            "concedida_em": "TEXT",
            "concedida_por": "TEXT",
        },
    )
    if "usuario_permissoes" in table_names:
        sync_conn.execute(
            text(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_backend_usuario_permissoes_unique "
                "ON usuario_permissoes(usuario_id, permissao_id)"
            )
        )
        sync_conn.execute(
            text("CREATE INDEX IF NOT EXISTS idx_backend_usuario_permissoes_usuario ON usuario_permissoes(usuario_id)")
        )

    add_missing_columns(
        "sistema_logs",
        {
            "tipo_acao": "TEXT",
            "descricao": "TEXT",
            "usuario": "TEXT",
            "status": "TEXT DEFAULT 'OK'",
            "resultado_texto": "TEXT",
            "executado_em": "TEXT",
        },
    )
    if "sistema_logs" in table_names:
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_sistema_logs_exec ON sistema_logs(executado_em)"))

    add_missing_columns(
        "programacoes",
        {
            "codigo_programacao": "TEXT",
            "codigo": "TEXT",
            "data": "TEXT",
            "data_criacao": "TEXT",
            "motorista": "TEXT",
            "motorista_id": "INTEGER",
            "motorista_codigo": "TEXT",
            "codigo_motorista": "TEXT",
            "veiculo": "TEXT",
            "equipe": "TEXT",
            "kg_estimado": "REAL DEFAULT 0",
            "tipo_estimativa": "TEXT DEFAULT 'KG'",
            "caixas_estimado": "INTEGER DEFAULT 0",
            "operacao_tipo": "TEXT DEFAULT 'VENDA'",
            "transbordo_modalidade": "TEXT",
            "transbordo_observacao": "TEXT",
            "transbordo_grupo": "TEXT",
            "status": "TEXT DEFAULT 'ATIVA'",
            "status_operacional": "TEXT",
            "status_operacional_obs": "TEXT",
            "status_operacional_em": "TEXT",
            "status_operacional_por": "TEXT",
            "foto_doa_path": "TEXT",
            "doa_foto_path": "TEXT",
            "mortalidade_transbordo_foto_path": "TEXT",
            "foto_doa_ref_json": "TEXT",
            "ajudantes_alteracao_motivo": "TEXT",
            "ajudantes_alterado_em": "TEXT",
            "historico_ajudantes": "TEXT",
            "finalizada_no_app": "INTEGER DEFAULT 0",
            "prestacao_status": "TEXT DEFAULT 'PENDENTE'",
            "local_rota": "TEXT",
            "tipo_rota": "TEXT",
            "local_carregamento": "TEXT",
            "granja_carregada": "TEXT",
            "local_carregado": "TEXT",
            "local_carreg": "TEXT",
            "num_nf": "TEXT",
            "nf_numero": "TEXT",
            "saida_data": "TEXT",
            "saida_hora": "TEXT",
            "data_saida": "TEXT",
            "hora_saida": "TEXT",
            "saida_dt": "TEXT",
            "inicio_carregamento": "TEXT",
            "fim_carregamento": "TEXT",
            "carregamento_fechado": "INTEGER DEFAULT 0",
            "carregamento_salvo_em": "TEXT",
            "data_chegada": "TEXT",
            "hora_chegada": "TEXT",
            "chegada_dt": "TEXT",
            "diaria_motorista_valor": "REAL DEFAULT 0",
            "adiantamento": "REAL DEFAULT 0",
            "adiantamento_rota": "REAL DEFAULT 0",
            "adiantamento_origem": "TEXT",
            "kg_carregado": "REAL DEFAULT 0",
            "caixas_carregadas": "INTEGER DEFAULT 0",
            "qnt_cx_carregada": "INTEGER DEFAULT 0",
            "media": "REAL DEFAULT 0",
            "media_1": "REAL",
            "media_2": "REAL",
            "media_3": "REAL",
            "qnt_aves_por_cx": "INTEGER DEFAULT 0",
            "qnt_aves_caixa_final": "INTEGER DEFAULT 0",
            "aves_caixa_final": "INTEGER DEFAULT 0",
            "km_inicial": "REAL DEFAULT 0",
            "km_final": "REAL DEFAULT 0",
            "km_rodado": "REAL DEFAULT 0",
            "litros": "REAL DEFAULT 0",
            "media_km_l": "REAL DEFAULT 0",
            "custo_km": "REAL DEFAULT 0",
            "mortalidade_transbordo_aves": "INTEGER DEFAULT 0",
            "mortalidade_transbordo_kg": "REAL DEFAULT 0",
            "obs_transbordo": "TEXT",
            "rota_observacao": "TEXT",
            "ced_200_qtd": "INTEGER DEFAULT 0",
            "ced_100_qtd": "INTEGER DEFAULT 0",
            "ced_50_qtd": "INTEGER DEFAULT 0",
            "ced_20_qtd": "INTEGER DEFAULT 0",
            "ced_10_qtd": "INTEGER DEFAULT 0",
            "ced_5_qtd": "INTEGER DEFAULT 0",
            "ced_2_qtd": "INTEGER DEFAULT 0",
            "valor_dinheiro": "REAL DEFAULT 0",
            "pix_motorista": "REAL DEFAULT 0",
            "usuario_criacao": "TEXT",
            "usuario_ultima_edicao": "TEXT",
            "total_caixas": "INTEGER DEFAULT 0",
            "quilos": "REAL DEFAULT 0",
            "nf_kg": "REAL DEFAULT 0",
            "kg_nf": "REAL DEFAULT 0",
            "nf_preco": "REAL DEFAULT 0",
            "preco_nf": "REAL DEFAULT 0",
            "nf_caixas": "INTEGER DEFAULT 0",
            "nf_kg_carregado": "REAL DEFAULT 0",
            "nf_kg_vendido": "REAL DEFAULT 0",
            "nf_saldo": "REAL DEFAULT 0",
        },
    )
    if "programacoes" in table_names:
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_programacoes_codigo ON programacoes(codigo_programacao)"))
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET status=COALESCE(NULLIF(TRIM(status), ''), 'ATIVA'),
                       prestacao_status=COALESCE(NULLIF(TRIM(prestacao_status), ''), 'PENDENTE')
                 WHERE status IS NULL OR TRIM(status)=''
                    OR prestacao_status IS NULL OR TRIM(prestacao_status)=''
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET codigo_programacao=TRIM(COALESCE(codigo, ''))
                 WHERE TRIM(COALESCE(codigo_programacao, '')) = ''
                   AND TRIM(COALESCE(codigo, '')) <> ''
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET codigo=TRIM(COALESCE(codigo_programacao, ''))
                 WHERE TRIM(COALESCE(codigo, '')) = ''
                   AND TRIM(COALESCE(codigo_programacao, '')) <> ''
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET data_criacao=COALESCE(NULLIF(TRIM(data_criacao), ''), NULLIF(TRIM(data), ''), date('now')),
                       data=COALESCE(NULLIF(TRIM(data), ''), NULLIF(TRIM(data_criacao), ''), date('now')),
                       tipo_estimativa=UPPER(COALESCE(NULLIF(TRIM(tipo_estimativa), ''), 'KG')),
                       local_rota=COALESCE(NULLIF(TRIM(local_rota), ''), NULLIF(TRIM(tipo_rota), '')),
                       tipo_rota=COALESCE(NULLIF(TRIM(tipo_rota), ''), NULLIF(TRIM(local_rota), '')),
                       local_carregamento=COALESCE(NULLIF(TRIM(local_carregamento), ''), NULLIF(TRIM(granja_carregada), ''), NULLIF(TRIM(local_carregado), ''), NULLIF(TRIM(local_carreg), '')),
                       granja_carregada=COALESCE(NULLIF(TRIM(granja_carregada), ''), NULLIF(TRIM(local_carregamento), '')),
                       local_carregado=COALESCE(NULLIF(TRIM(local_carregado), ''), NULLIF(TRIM(local_carregamento), '')),
                       local_carreg=COALESCE(NULLIF(TRIM(local_carreg), ''), NULLIF(TRIM(local_carregamento), '')),
                       nf_numero=COALESCE(NULLIF(TRIM(nf_numero), ''), NULLIF(TRIM(num_nf), '')),
                       num_nf=COALESCE(NULLIF(TRIM(num_nf), ''), NULLIF(TRIM(nf_numero), '')),
                       saida_dt=COALESCE(NULLIF(TRIM(saida_dt), ''), TRIM(COALESCE(data_saida, '') || ' ' || COALESCE(hora_saida, ''))),
                       chegada_dt=COALESCE(NULLIF(TRIM(chegada_dt), ''), TRIM(COALESCE(data_chegada, '') || ' ' || COALESCE(hora_chegada, '')))
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET operacao_tipo=CASE
                           WHEN UPPER(TRIM(COALESCE(operacao_tipo, '')))='TRANSBORDO'
                             OR UPPER(TRIM(COALESCE(tipo_estimativa, '')))='CX'
                             OR TRIM(COALESCE(transbordo_grupo, ''))<>''
                           THEN 'TRANSBORDO'
                           ELSE 'VENDA'
                       END
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET transbordo_modalidade=CASE
                           WHEN UPPER(TRIM(COALESCE(transbordo_modalidade, '')))='FOB' THEN 'EMPRESA_BUSCA'
                           ELSE COALESCE(NULLIF(TRIM(transbordo_modalidade), ''), 'EMPRESA_BUSCA')
                       END,
                       transbordo_grupo=COALESCE(NULLIF(TRIM(transbordo_grupo), ''), NULLIF(TRIM(codigo_programacao), ''), NULLIF(TRIM(codigo), ''))
                 WHERE UPPER(TRIM(COALESCE(operacao_tipo, '')))='TRANSBORDO'
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET status='FINALIZADA',
                       status_operacional='FINALIZADA',
                       finalizada_no_app=1
                 WHERE UPPER(TRIM(COALESCE(prestacao_status, '')))='FECHADA'
                    OR UPPER(TRIM(COALESCE(status, ''))) IN ('FINALIZADA','FINALIZADO')
                    OR UPPER(TRIM(COALESCE(status_operacional, ''))) IN ('FINALIZADA','FINALIZADO')
                    OR COALESCE(finalizada_no_app, 0)=1
                    OR TRIM(COALESCE(data_chegada, ''))<>''
                    OR TRIM(COALESCE(hora_chegada, ''))<>''
                    OR COALESCE(km_final, 0)>0
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET status='CANCELADA',
                       status_operacional='CANCELADA',
                       finalizada_no_app=1
                 WHERE UPPER(TRIM(COALESCE(status, ''))) IN ('CANCELADA','CANCELADO')
                    OR UPPER(TRIM(COALESCE(status_operacional, ''))) IN ('CANCELADA','CANCELADO')
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET status_operacional=NULL
                 WHERE UPPER(TRIM(COALESCE(status, ''))) NOT IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')
                   AND UPPER(TRIM(COALESCE(status_operacional, ''))) IN ('FINALIZADA','FINALIZADO','CANCELADA','CANCELADO')
                   AND UPPER(TRIM(COALESCE(prestacao_status, 'PENDENTE'))) <> 'FECHADA'
                   AND COALESCE(finalizada_no_app, 0)=0
                   AND TRIM(COALESCE(data_chegada, ''))=''
                   AND TRIM(COALESCE(hora_chegada, ''))=''
                   AND COALESCE(km_final, 0)=0
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET motorista_codigo=COALESCE(NULLIF(TRIM(motorista_codigo), ''), NULLIF(TRIM(codigo_motorista), '')),
                       codigo_motorista=COALESCE(NULLIF(TRIM(codigo_motorista), ''), NULLIF(TRIM(motorista_codigo), '')),
                       foto_doa_path=COALESCE(NULLIF(TRIM(foto_doa_path), ''), NULLIF(TRIM(doa_foto_path), ''), NULLIF(TRIM(mortalidade_transbordo_foto_path), '')),
                       doa_foto_path=COALESCE(NULLIF(TRIM(doa_foto_path), ''), NULLIF(TRIM(foto_doa_path), ''), NULLIF(TRIM(mortalidade_transbordo_foto_path), '')),
                       mortalidade_transbordo_foto_path=COALESCE(NULLIF(TRIM(mortalidade_transbordo_foto_path), ''), NULLIF(TRIM(foto_doa_path), ''), NULLIF(TRIM(doa_foto_path), ''))
                """
            )
        )
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET nf_saldo=ROUND(MAX(COALESCE(nf_kg, 0) - COALESCE(NULLIF(nf_kg_carregado, 0), kg_carregado, 0), 0), 2)
                 WHERE COALESCE(nf_kg, 0) > 0
                   AND COALESCE(NULLIF(nf_kg_carregado, 0), kg_carregado, 0) > 0
                   AND COALESCE(nf_saldo, 0) <= 0
                """
            )
        )

    add_missing_columns(
        "programacao_itens",
        {
            "codigo_programacao": "TEXT",
            "cod_cliente": "TEXT",
            "pedido": "TEXT",
            "nome_cliente": "TEXT",
            "qnt_caixas": "INTEGER DEFAULT 0",
            "kg": "REAL DEFAULT 0",
            "preco": "REAL DEFAULT 0",
            "endereco": "TEXT",
            "vendedor": "TEXT",
            "pedido": "TEXT",
            "produto": "TEXT",
            "produto_id": "INTEGER",
            "observacao": "TEXT",
            "status_pedido": "TEXT",
            "caixas_atual": "INTEGER",
            "preco_atual": "REAL",
            "alterado_em": "TEXT",
            "alterado_por": "TEXT",
            "ordem_sugerida": "INTEGER",
            "eta": "TEXT",
            "distancia": "REAL",
            "confianca_localizacao": "REAL",
            "carga_raiz_programacao": "TEXT",
            "carga_origem_imediata": "TEXT",
            "transferencia_origem_id": "TEXT",
        },
    )
    if "programacao_itens" in table_names:
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_prog_itens_programacao ON programacao_itens(codigo_programacao)"))
        sync_conn.execute(
            text(
                """
                UPDATE programacoes
                   SET total_caixas=COALESCE((
                           SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                             FROM programacao_itens pi
                            WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) =
                                  UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                       ), total_caixas, 0),
                       nf_caixas=COALESCE((
                           SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                             FROM programacao_itens pi
                            WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) =
                                  UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                       ), nf_caixas, 0),
                       caixas_carregadas=COALESCE((
                           SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                             FROM programacao_itens pi
                            WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) =
                                  UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                       ), caixas_carregadas, 0),
                       qnt_cx_carregada=COALESCE((
                           SELECT SUM(COALESCE(pi.qnt_caixas, 0))
                             FROM programacao_itens pi
                            WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) =
                                  UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                       ), qnt_cx_carregada, 0)
                 WHERE TRIM(COALESCE(codigo_programacao, '')) <> ''
                   AND EXISTS (
                       SELECT 1
                         FROM programacao_itens pi
                        WHERE UPPER(TRIM(COALESCE(pi.codigo_programacao, ''))) =
                              UPPER(TRIM(COALESCE(programacoes.codigo_programacao, '')))
                   )
                   AND (
                       COALESCE(total_caixas, 0)=0
                    OR COALESCE(nf_caixas, 0)=0
                    OR COALESCE(caixas_carregadas, 0)=0
                    OR COALESCE(qnt_cx_carregada, 0)=0
                   )
                """
            )
        )

    add_missing_columns(
        "recebimentos",
        {
            "codigo_programacao": "TEXT",
            "cod_cliente": "TEXT",
            "pedido": "TEXT",
            "nome_cliente": "TEXT",
            "valor": "REAL DEFAULT 0",
            "forma_pagamento": "TEXT",
            "observacao": "TEXT",
            "num_nf": "TEXT",
            "data_registro": "TEXT",
            "company_id": "INTEGER",
        },
    )
    if "recebimentos" in table_names:
        backfill_company_id("recebimentos")
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_receb_programacao ON recebimentos(codigo_programacao)"))
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_receb_cliente ON recebimentos(cod_cliente)"))

    add_missing_columns(
        "despesas",
        {
            "codigo_programacao": "TEXT",
            "descricao": "TEXT",
            "valor": "REAL DEFAULT 0",
            "data_registro": "TEXT",
            "tipo_despesa": "TEXT DEFAULT 'ROTA'",
            "categoria": "TEXT",
            "motorista": "TEXT",
            "veiculo": "TEXT",
            "observacao": "TEXT",
            "id_local": "TEXT",
            "forma_pagamento": "TEXT",
            "comprovante_path": "TEXT",
            "estabelecimento": "TEXT",
            "documento": "TEXT",
            "litros": "REAL",
            "valor_litro": "REAL",
            "desconto": "REAL",
            "combustivel": "TEXT",
            "odometro": "REAL",
            "lat": "REAL",
            "lon": "REAL",
            "accuracy": "REAL",
            "registrado_em": "TEXT",
            "motorista_codigo": "TEXT",
            "motorista_nome": "TEXT",
            "sync_key": "TEXT",
            "status_sync": "TEXT",
            "origem": "TEXT",
            "vinculo_prestacao_json": "TEXT",
            "desktop_web_json": "TEXT",
            "foto_despesa_ref_json": "TEXT",
            "company_id": "INTEGER",
        },
    )
    if "despesas" in table_names:
        backfill_company_id("despesas")
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_despesas_programacao ON despesas(codigo_programacao)"))
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_despesas_prog_id_local ON despesas(codigo_programacao, id_local)"))

    add_missing_columns(
        "programacao_itens_controle",
        {
            "codigo_programacao": "TEXT",
            "cod_cliente": "TEXT",
            "pedido": "TEXT",
            "status_pedido": "TEXT",
            "alteracao_tipo": "TEXT",
            "alteracao_detalhe": "TEXT",
            "caixas_atual": "INTEGER",
            "preco_atual": "REAL",
            "alterado_em": "TEXT",
            "alterado_por": "TEXT",
            "mortalidade_aves": "INTEGER DEFAULT 0",
            "media_aplicada": "REAL",
            "peso_previsto": "REAL DEFAULT 0",
            "valor_recebido": "REAL DEFAULT 0",
            "forma_recebimento": "TEXT",
            "obs_recebimento": "TEXT",
            "lat_evento": "REAL",
            "lon_evento": "REAL",
            "lat_entrega": "REAL",
            "lon_entrega": "REAL",
            "accuracy_entrega": "REAL",
            "timestamp_entrega": "TEXT",
            "endereco_evento": "TEXT",
            "cidade_evento": "TEXT",
            "bairro_evento": "TEXT",
            "foto_mortalidade_path": "TEXT",
            "mortalidade_foto_path": "TEXT",
            "foto_mortalidade_ref_json": "TEXT",
            "ordem_sugerida": "INTEGER",
            "eta": "TEXT",
            "distancia": "REAL",
            "confianca_localizacao": "REAL",
            "updated_at": "TEXT",
            "company_id": "INTEGER",
        },
    )
    if "programacao_itens_controle" in table_names:
        backfill_company_id("programacao_itens_controle")
        sync_conn.execute(
            text(
                "CREATE INDEX IF NOT EXISTS idx_backend_prog_itens_ctrl_prog_cliente "
                "ON programacao_itens_controle(codigo_programacao, cod_cliente)"
            )
        )

    if "rota_fotos" not in table_names:
        sync_conn.execute(
            text(
                """
                CREATE TABLE IF NOT EXISTS rota_fotos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    id_foto TEXT UNIQUE,
                    codigo_programacao TEXT NOT NULL,
                    categoria TEXT NOT NULL,
                    tipo_registro TEXT NOT NULL,
                    cod_cliente TEXT,
                    cliente_nome TEXT,
                    pedido TEXT,
                    id_vinculo TEXT,
                    path_local TEXT,
                    storage_path TEXT,
                    arquivo_nome TEXT,
                    mime_type TEXT,
                    tamanho_bytes INTEGER DEFAULT 0,
                    motorista_codigo TEXT,
                    motorista_nome TEXT,
                    registrado_em TEXT DEFAULT (datetime('now')),
                    payload_json TEXT DEFAULT '{}',
                    company_id INTEGER
                )
                """
            )
        )
        table_names = set(inspect(sync_conn).get_table_names())
    add_missing_columns(
        "rota_fotos",
        {
            "id_foto": "TEXT",
            "codigo_programacao": "TEXT",
            "categoria": "TEXT",
            "tipo_registro": "TEXT",
            "cod_cliente": "TEXT",
            "cliente_nome": "TEXT",
            "pedido": "TEXT",
            "id_vinculo": "TEXT",
            "path_local": "TEXT",
            "storage_path": "TEXT",
            "arquivo_nome": "TEXT",
            "mime_type": "TEXT",
            "tamanho_bytes": "INTEGER DEFAULT 0",
            "motorista_codigo": "TEXT",
            "motorista_nome": "TEXT",
            "registrado_em": "TEXT",
            "payload_json": "TEXT DEFAULT '{}'",
            "company_id": "INTEGER",
        },
    )
    if "rota_fotos" in table_names:
        backfill_company_id("rota_fotos")
        sync_conn.execute(text("CREATE UNIQUE INDEX IF NOT EXISTS idx_backend_rota_fotos_id ON rota_fotos(id_foto)"))
        sync_conn.execute(
            text("CREATE INDEX IF NOT EXISTS idx_backend_rota_fotos_prog_categoria ON rota_fotos(codigo_programacao, categoria)")
        )

    add_missing_columns(
        "vendas_importadas",
        {
            "pedido": "TEXT",
            "data_venda": "TEXT",
            "cliente": "TEXT",
            "nome_cliente": "TEXT",
            "vendedor": "TEXT",
            "produto": "TEXT",
            "produto_id": "INTEGER",
            "vr_total": "REAL DEFAULT 0",
            "qnt": "REAL DEFAULT 0",
            "qnt_caixas": "INTEGER DEFAULT 0",
            "cidade": "TEXT",
            "valor_unitario": "REAL DEFAULT 0",
            "observacao": "TEXT",
            "selecionada": "INTEGER DEFAULT 0",
            "usada": "INTEGER DEFAULT 0",
            "usada_em": "TEXT",
            "codigo_programacao": "TEXT",
        },
    )
    if "vendas_importadas" in table_names:
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_vendas_importadas_livre ON vendas_importadas(usada, codigo_programacao)"))
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_vendas_importadas_chave ON vendas_importadas(pedido, cliente, produto, data_venda)"))

    add_missing_columns(
        "rota_gps_pings",
        {
            "codigo_programacao": "TEXT",
            "motorista": "TEXT",
            "lat": "REAL",
            "lon": "REAL",
            "speed": "REAL",
            "accuracy": "REAL",
            "recorded_at": "TEXT",
            "company_id": "INTEGER",
        },
    )
    if "rota_gps_pings" in table_names:
        backfill_company_id("rota_gps_pings")
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_rota_gps_prog_id ON rota_gps_pings(codigo_programacao, id)"))

    add_missing_columns(
        "programacao_itens_log",
        {
            "codigo_programacao": "TEXT",
            "cod_cliente": "TEXT",
            "pedido": "TEXT",
            "evento": "TEXT DEFAULT 'cliente_controle'",
            "payload_json": "TEXT DEFAULT '{}'",
            "registrado_em": "TEXT",
            "created_at": "TEXT",
            "company_id": "INTEGER",
        },
    )
    if "programacao_itens_log" in table_names:
        backfill_company_id("programacao_itens_log")
        sync_conn.execute(
            text(
                """
                UPDATE programacao_itens_log
                   SET evento=COALESCE(NULLIF(TRIM(evento), ''), 'cliente_controle'),
                       created_at=COALESCE(NULLIF(TRIM(created_at), ''), NULLIF(TRIM(registrado_em), ''), datetime('now'))
                 WHERE evento IS NULL OR TRIM(evento)=''
                    OR created_at IS NULL OR TRIM(created_at)=''
                """
            )
        )
        sync_conn.execute(
            text(
                "CREATE INDEX IF NOT EXISTS idx_backend_prog_itens_log_prog_cliente "
                "ON programacao_itens_log(codigo_programacao, cod_cliente)"
            )
        )

    if "roteiro_operacional" not in table_names:
        sync_conn.execute(
            text(
                """
                CREATE TABLE IF NOT EXISTS roteiro_operacional (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo_evento TEXT NOT NULL,
                    codigo_programacao TEXT NOT NULL,
                    origem TEXT,
                    destino TEXT,
                    motorista_codigo TEXT,
                    motorista_nome TEXT,
                    pedido TEXT,
                    cod_cliente TEXT,
                    cliente_nome TEXT,
                    caixas INTEGER DEFAULT 0,
                    kg REAL DEFAULT 0,
                    media REAL DEFAULT 0,
                    aves_por_caixa INTEGER DEFAULT 0,
                    nf_numero TEXT,
                    nf_preco REAL DEFAULT 0,
                    lotes TEXT,
                    data_hora TEXT,
                    observacao TEXT,
                    payload_json TEXT DEFAULT '{}',
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
        )
        table_names = set(inspect(sync_conn).get_table_names())

    if "escala_folgas" not in table_names:
        sync_conn.execute(
            text(
                """
                CREATE TABLE IF NOT EXISTS escala_folgas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo TEXT NOT NULL,
                    pessoa_id TEXT,
                    pessoa_codigo TEXT,
                    pessoa_nome TEXT NOT NULL,
                    data_inicio TEXT NOT NULL,
                    data_fim TEXT NOT NULL,
                    motivo TEXT,
                    status TEXT DEFAULT 'ATIVA',
                    criado_em TEXT,
                    atualizado_em TEXT,
                    company_id INTEGER
                )
                """
            )
        )
        table_names = set(inspect(sync_conn).get_table_names())
    add_missing_columns(
        "escala_folgas",
        {
            "tipo": "TEXT",
            "pessoa_id": "TEXT",
            "pessoa_codigo": "TEXT",
            "pessoa_nome": "TEXT",
            "data_inicio": "TEXT",
            "data_fim": "TEXT",
            "motivo": "TEXT",
            "status": "TEXT DEFAULT 'ATIVA'",
            "criado_em": "TEXT",
            "atualizado_em": "TEXT",
            "company_id": "INTEGER",
        },
    )
    if "escala_folgas" in table_names:
        sync_conn.execute(text("UPDATE escala_folgas SET company_id=1 WHERE company_id IS NULL"))
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_escala_folgas_status ON escala_folgas(status)"))
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_escala_folgas_pessoa ON escala_folgas(tipo, pessoa_nome)"))
        sync_conn.execute(
            text("CREATE INDEX IF NOT EXISTS idx_backend_escala_folgas_company_status ON escala_folgas(company_id, status)")
        )

    add_missing_columns(
        "roteiro_operacional",
        {
            "tipo_evento": "TEXT",
            "codigo_programacao": "TEXT",
            "origem": "TEXT",
            "destino": "TEXT",
            "motorista_codigo": "TEXT",
            "motorista_nome": "TEXT",
            "pedido": "TEXT",
            "cod_cliente": "TEXT",
            "cliente_nome": "TEXT",
            "caixas": "INTEGER DEFAULT 0",
            "kg": "REAL DEFAULT 0",
            "media": "REAL DEFAULT 0",
            "aves_por_caixa": "INTEGER DEFAULT 0",
            "nf_numero": "TEXT",
            "nf_preco": "REAL DEFAULT 0",
            "lotes": "TEXT",
            "data_hora": "TEXT",
            "observacao": "TEXT",
            "payload_json": "TEXT DEFAULT '{}'",
            "created_at": "TEXT",
        },
    )
    if "roteiro_operacional" in table_names:
        sync_conn.execute(
            text("CREATE INDEX IF NOT EXISTS idx_backend_roteiro_operacional_prog_data ON roteiro_operacional(codigo_programacao, data_hora, id)")
        )
        sync_conn.execute(text("CREATE INDEX IF NOT EXISTS idx_backend_roteiro_operacional_tipo ON roteiro_operacional(tipo_evento)"))


async def get_db() -> AsyncGenerator[AsyncSession, None]:
    """Dependency for database sessions"""
    async with async_session() as session:
        try:
            yield session
        finally:
            await session.close()


async def create_tables():
    """Create all database tables"""
    try:
        async with engine.begin() as conn:
            await conn.run_sync(Base.metadata.create_all)
            await conn.run_sync(_ensure_backend_columns)
        logger.info("Database tables created successfully")
    except Exception as e:
        logger.error(f"Failed to create database tables: {e}")
        raise


async def drop_tables():
    """Drop all database tables (for testing/reset)"""
    try:
        async with engine.begin() as conn:
            await conn.run_sync(Base.metadata.drop_all)
        logger.info("Database tables dropped successfully")
    except Exception as e:
        logger.error(f"Failed to drop database tables: {e}")
        raise
