# Arquitetura Atual do Sistema TMS RotaHub

## Visão Geral
O sistema TMS (Transport Management System) RotaHub é uma aplicação desktop desenvolvida em Python usando Tkinter para interface gráfica. Possui aproximadamente 22.000 linhas de código no arquivo principal `main.py`, indicando uma arquitetura monolítica com alto acoplamento. Há tentativas iniciais de migração para arquitetura cliente-servidor com APIs FastAPI.

## Estrutura de Arquivos
```
c:\pdc_rota\
├── main.py (22.000+ linhas - Interface desktop e lógica principal)
├── database.py (Definições de tabelas SQLite)
├── models.py (Vazio - Não utilizado)
├── api_server.py (Servidor FastAPI principal)
├── server.py (Servidor FastAPI secundário)
├── app/
│   ├── db/ (Conexão e migrações)
│   ├── services/ (Lógica de negócio)
│   ├── ui/
│   │   ├── components/ (Componentes reutilizáveis)
│   │   └── pages/ (Páginas da interface)
│   ├── middleware/ (Middlewares para tenant, billing, etc.)
│   ├── security/ (Autenticação e senhas)
│   └── utils/
├── flutter_vendedor_app/ (App mobile Flutter)
├── migrations/ (Scripts SQL)
├── scripts/ (Scripts utilitários)
├── tests/ (Testes - provavelmente vazios)
└── requirements.txt (Dependências Python)
```

## Arquitetura Técnica

### Frontend (Desktop)
- **Framework**: Tkinter + ttk (widgets nativos)
- **Padrão**: Páginas baseadas em classes `PageBase`
- **Estilos**: Tema customizado com cores e fontes
- **Responsividade**: Limitada ao desktop
- **Navegação**: Menu lateral com botões

### Backend (API)
- **Framework**: FastAPI (Python)
- **Autenticação**: JWT (parcialmente implementado)
- **Banco**: SQLite local (sincronização com API)
- **Middlewares**: Tenant (multiempresa), Billing, Feature Gates
- **Serviços**: Lógica de negócio separada em `app/services/`

### Banco de Dados
- **Tipo**: SQLite (local)
- **Tabelas Principais**:
  - usuarios (autenticação)
  - motoristas, veiculos, equipes
  - programacoes, programacao_itens
  - pdc_lancamentos (recebimentos/despesas)
  - fechamento_rotas, fechamento_despesas
- **Migrações**: Scripts manuais em `migrations/`

## Módulos do Sistema

### 1. Dashboard Operacional
- KPIs: Rotas, motoristas, ajudantes, mortalidade, KM, horas
- Gráficos: Canvas customizado para barras
- Localização: `HomePage` e `EscalaPage`

### 2. Cadastros
- Usuários, Motoristas, Veículos, Equipes, Vendedores
- Interface: CRUD com Treeview e formulários
- Localização: `CadastrosPage`

### 3. Rotas
- Gerenciamento de rotas de entrega
- Localização: `RotasPage`

### 4. Programação de Entrega
- Agendamento e controle de entregas
- Itens de programação com equipes
- Localização: `ProgramacaoPage`

### 5. Acompanhamento
- Status: Ativa, Em Rota, Finalizada, Cancelada
- Mapas: Integração futura (não implementada)

### 6. Recebimentos e Despesas
- Lançamentos financeiros
- Fechamentos por rota
- Localização: `RecebimentosPage`, `DespesasPage`

### 7. Prestação de Contas
- Relatórios de custos e performance
- Mortalidade de aves, KM/L, custos

### 8. Escala
- Distribuição de carga por motorista/ajudante
- Indicadores de sobrecarga
- Recomendações automáticas
- Localização: `EscalaPage`

### 9. Centro de Custos
- Controle orçamentário
- Localização: `CentroCustosPage`

### 10. Relatórios
- Geração de PDF com ReportLab
- Exportação Excel com Pandas
- Localização: `RelatoriosPage`

### 11. Permissões e Configurações
- Controle de acesso por rotina
- Configurações do sistema
- Localização: `PermissionsPage`

### 12. Backup e Atualização
- Exportação de dados
- Atualização automática/manual
- Localização: `BackupExportarPage`

## Regras de Negócio Identificadas

### Autenticação
- Usuários com permissões (OPERADOR, etc.)
- Hashing PBKDF2 para senhas
- Controle por empresa (tenant)

### Programação
- Cálculo de horas trabalhadas (saída/chegada)
- Validação de equipes e motoristas
- Status de rotas com transições

### Financeiro
- Recebimentos por rota
- Despesas categorizadas
- Fechamentos mensais
- Cálculo de custos por KM

### Escala
- Distribuição equilibrada de carga
- Indicadores de sobrecarga (>1.25x média)
- Recomendações baseadas em localização

### Relatórios
- Filtros por período e status
- Geração PDF com tabelas e gráficos
- Exportação para Excel

## Pontos Críticos
- **Monolítico**: main.py com 22k linhas
- **Acoplamento**: Lógica UI + negócio misturadas
- **Banco Local**: SQLite vulnerável a perda de dados
- **Testes**: Ausentes
- **Documentação**: Limitada
- **Escalabilidade**: Não preparada para multi-usuário simultâneo

## Dependências Externas
- Tkinter (interface)
- SQLite3 (banco)
- FastAPI, Uvicorn (API)
- Pandas (exportação)
- ReportLab (PDF)
- httpx (HTTP client)
- Flutter (app mobile)

## Estado de SaaS
- Middlewares tenant e billing existem
- Controle de planos básico
- API para sincronização
- Pronto para migração incremental

## Conclusão
O sistema possui funcionalidades completas mas arquitetura obsoleta. A migração para SaaS web + desktop cliente é viável com preservação de regras de negócio.