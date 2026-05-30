# backend/models/cadastro.py
"""
Operational cadastro models
"""
from sqlalchemy import Column, Float, Integer, String

from backend.config.database import Base


class MotoristaDB(Base):
    __tablename__ = "motoristas"

    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String, nullable=False)
    codigo = Column(String, index=True)
    senha = Column(String)
    perfil_app = Column(String, default="MOTORISTA")
    cpf = Column(String)
    telefone = Column(String)
    status = Column(String, default="ATIVO")
    company_id = Column(Integer)


class VendedorDB(Base):
    __tablename__ = "vendedores"

    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String, unique=True, index=True)
    nome = Column(String, nullable=False)
    senha = Column(String)
    telefone = Column(String)
    cidade_base = Column(String)
    status = Column(String, default="ATIVO")
    company_id = Column(Integer)


class VeiculoDB(Base):
    __tablename__ = "veiculos"

    id = Column(Integer, primary_key=True, index=True)
    placa = Column(String, index=True)
    modelo = Column(String)
    capacidade_cx = Column(Integer)
    status = Column(String, default="ATIVO")
    company_id = Column(Integer)


class CaixaDB(Base):
    __tablename__ = "caixas"

    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String, unique=True, index=True)
    lote = Column(String, index=True)
    cor = Column(String)
    veiculo_placa = Column(String, index=True)
    status = Column(String, default="EM_ESTOQUE")
    data_compra = Column(String)
    observacao = Column(String)
    company_id = Column(Integer)


class CaixaMovimentoDB(Base):
    __tablename__ = "caixas_movimentos"

    id = Column(Integer, primary_key=True, index=True)
    caixa_id = Column(Integer, index=True)
    codigo = Column(String, index=True)
    movimento = Column(String)
    veiculo_origem = Column(String)
    veiculo_destino = Column(String)
    status_origem = Column(String)
    status_destino = Column(String)
    observacao = Column(String)
    criado_em = Column(String)
    company_id = Column(Integer)


class AjudanteDB(Base):
    __tablename__ = "ajudantes"

    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String, nullable=False)
    sobrenome = Column(String)
    telefone = Column(String)
    status = Column(String, default="ATIVO")
    company_id = Column(Integer)


class EscalaFolgaDB(Base):
    __tablename__ = "escala_folgas"

    id = Column(Integer, primary_key=True, index=True)
    tipo = Column(String, nullable=False)
    pessoa_id = Column(String)
    pessoa_codigo = Column(String)
    pessoa_nome = Column(String, nullable=False)
    data_inicio = Column(String, nullable=False)
    data_fim = Column(String, nullable=False)
    motivo = Column(String)
    status = Column(String, default="ATIVA")
    criado_em = Column(String)
    atualizado_em = Column(String)
    company_id = Column(Integer)


class ClienteDB(Base):
    __tablename__ = "clientes"

    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String)
    cod_cliente = Column(String, unique=True, index=True)
    nome_cliente = Column(String, nullable=False)
    endereco = Column(String)
    bairro = Column(String)
    cidade = Column(String)
    uf = Column(String)
    telefone = Column(String)
    rota = Column(String)
    vendedor = Column(String)
    company_id = Column(Integer)


class FornecedorDB(Base):
    __tablename__ = "fornecedores"

    id = Column(Integer, primary_key=True, index=True)
    razao_social = Column(String, nullable=False)
    nome_fantasia = Column(String)
    documento = Column(String, index=True)
    tipo_pessoa = Column(String, default="CNPJ")
    perfil_fornecedor = Column(String, default="OUTROS")
    telefone = Column(String)
    email = Column(String)
    cidade = Column(String)
    uf = Column(String)
    status = Column(String, default="ATIVO")
    observacao = Column(String)
    certificado_nome = Column(String)
    certificado_path = Column(String)
    certificado_status = Column(String, default="NAO_INSTALADO")
    certificado_instalado_em = Column(String)
    company_id = Column(Integer)


class ProdutoDB(Base):
    __tablename__ = "produtos"

    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String, unique=True, index=True)
    nome = Column(String, nullable=False, index=True)
    descricao = Column(String)
    categoria = Column(String, default="AVES")
    unidade = Column(String, default="KG")
    unidade_estoque = Column(String, default="KG")
    controla_estoque_fisico = Column(Integer, default=1)
    controla_estoque_fiscal = Column(Integer, default=1)
    estoque_min_kg = Column(Float, default=0)
    estoque_min_caixas = Column(Integer, default=0)
    ncm = Column(String)
    cest = Column(String)
    cfop_entrada = Column(String)
    cfop_saida = Column(String)
    ean = Column(String)
    custo_padrao = Column(Float, default=0)
    preco_padrao = Column(Float, default=0)
    status = Column(String, default="ATIVO")
    company_id = Column(Integer)
