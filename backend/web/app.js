const API_BASE = window.location.protocol === "file:"
  ? "http://127.0.0.1:8012/api/v1"
  : "/api/v1";
const TOKEN_KEY = "rotahub_access_token";
const TOKEN_EXPIRES_KEY = "rotahub_access_token_expires_at";
const REMEMBER_USER_KEY = "rotahub_remember_username";

const state = {
  token: sessionStorage.getItem(TOKEN_KEY) || "",
  tokenExpiresAt: Number(sessionStorage.getItem(TOKEN_EXPIRES_KEY) || 0),
  currentUser: null,
  planContext: null,
  users: [],
  audit: [],
  home: null,
  homeRoutePreview: null,
  currentView: "dashboard",
  cadastroResource: "motoristas",
  cadastroItems: [],
  cadastroLoadSeq: 0,
  cadastroNextMotoristaCodigo: "",
  clientesImportRows: [],
  clientesDashboard: null,
  clientesLookup: [],
  clientesSection: "hub",
  cadastroSelectedId: null,
  cadastroMode: "view",
  cadastroStatusFilter: "TODOS",
  caixasVeiculos: [],
  fornecedorPerfis: ["DISTRIBUICAO_GERAL", "PRESTADOR_SERVICO", "PECAS", "MECANICO", "BORRACHEIRO", "LAVADOR_CAIXAS", "PNEUS", "OLEO_LUBRIFICANTES", "COMBUSTIVEL", "SERVICO_SEM_NF", "OUTROS"],
  logisticaConfig: null,
  programacaoOptions: {motoristas: [], veiculos: [], ajudantes: [], proximo_codigo: ""},
  programacoes: [],
  programacaoItems: [],
  programacaoSelectedCodigo: "",
  programacaoLoadedVendaIds: [],
  programacaoRankings: {periodo_dias: 30, motoristas: [], ajudantes: [], resumo_motoristas: "", resumo_ajudantes: ""},
  vendasImportadas: [],
  vendasProgramacoes: [],
  rotas: [],
  rotasSelectedCodigo: "",
  rotasRefreshMs: 10000,
  escala: null,
  escalaPeriodo: "30",
  escalaStatus: "ATIVAS",
  escalaTab: "motoristas",
  escalaFolgaPessoas: [],
  escalaFolgaLoadSeq: 0,
  escalaLoadSeq: 0,
  recebimentosProgramacoes: [],
  recebimentosBundle: null,
  recebimentosSelectedCodigo: "",
  recebimentosSelectedCodCliente: "",
  despesasProgramacoes: [],
  despesasBundle: null,
  despesasSelectedCodigo: "",
  despesasSelectedId: null,
  mortalidade: null,
  mortalidadeDoaDetails: [],
  mortalidadeDoaPhotoUrl: "",
  centroCustosOptions: {periodos: ["7", "15", "30", "60", "90", "180", "TODAS"], metricas: ["CUSTO_KM", "CUSTO_KG", "DESPESA_TOTAL"], veiculos: ["TODOS"]},
  centroCustosResumo: null,
  centroCustosDespesasRota: null,
  centroCustosDespesasRotaQuick: "",
  centroCustosDespesasRotaMesAuto: false,
  comprasNfe: {naturezas: [], rows: [], selectedId: null, total: 0, total_valor: 0},
  comprasEstoque: null,
  comprasFornecedores: [],
  relatoriosOptions: {tipos: ["Nota Fiscal / Transbordo", "Detalhe Completo da Rota", "Programacoes", "Prestacao de Contas", "Mortalidades", "Ocorrencias por Motorista", "Rotina Motorista/Ajudantes", "KM de Veiculos", "Abastecimentos", "Banhos", "Despesas"]},
  relatoriosProgramacoes: [],
  relatoriosResumo: null,
  systemTools: {backups: [], logs: [], info: null},
  permissoes: {usuarios: [], permissoes: [], modulos: [], selectedUserId: null, granted: []},
  billing: {company: null, subscription: null, usage: {}, plans: [], pending_requests: []},
  saasAdmin: {
    companies: [],
    company: null,
    subscription: null,
    usage: {},
    plans: [],
    payments: [],
    audit_logs: [],
    features: null,
    selectedCompanyId: null,
  },
};

let rotasRefreshTimer = null;
let authRefreshTimer = null;
let authRefreshPromise = null;
let tableEnhanceTimer = null;
let activeColumnFilterMenu = null;
const tableColumnFilters = new WeakMap();
const tableColumnVisibility = new WeakMap();
const tableColumnVisibilityButtons = new WeakMap();
const COLUMN_VISIBILITY_STORAGE_KEY = "rotahub_table_column_visibility";
const DESPESAS_CEDULAS = [200, 100, 50, 20, 10, 5, 2];
const VIEW_PLAN_FEATURES = {
  rotas: "rotas",
  escala: "escala",
  importarVendas: "importar_vendas",
  programacao: "programacao",
  recebimentos: "recebimentos",
  despesas: "despesas",
  mortalidade: "mortalidade",
  centroCustos: "centro_custos",
  compras: "despesas",
  relatorios: "relatorios",
  cadastros: "cadastros",
  users: "cadastros",
  permissoes: "cadastros",
  backup: "private_deployment",
  ferramentas: "private_deployment",
  audit: "relatorios",
};

const RELATORIOS_MODE_META = {
  "Nota Fiscal / Transbordo": {
    title: "NF / Transbordo",
    desc: "Consolida carga raiz, transferencias, entregas, fotos, custos e recebimentos por nota.",
    tag: "Carga",
  },
  "Prestacao de Contas": {
    title: "Fechamento",
    desc: "Confere recebimentos, despesas, cedulas, caixa, resultado e documentos.",
    tag: "Caixa",
  },
  "Mortalidades": {
    title: "Mortalidades",
    desc: "Analisa aves, kg, medias, motivos, fotos e ranking por motorista ou rota.",
    tag: "Ocorrencias",
  },
  "KM de Veiculos": {
    title: "KM / Consumo",
    desc: "Compara km rodado, litros, media km/l, custo total e custo por km.",
    tag: "Frota",
  },
  "Abastecimentos": {
    title: "Abastecimentos",
    desc: "Lista combustivel, litros, valor/litro, odometro, veiculo e motorista.",
    tag: "Despesas",
  },
  "Banhos": {
    title: "Banhos",
    desc: "Agrupa lavagens e higienizacoes lancadas nas despesas do app motorista.",
    tag: "Operacao",
  },
  "Detalhe Completo da Rota": {
    title: "Rota Completa",
    desc: "Dossie completo da programacao: compra, venda, clientes, custos e transferencias.",
    tag: "Dossie",
  },
  "Programacoes": {
    title: "Planejamentos",
    desc: "Consulta o planejamento original e reimprime folha e romaneios.",
    tag: "Rota",
  },
  "Rotina Motorista/Ajudantes": {
    title: "Rotina",
    desc: "Resume viagens, equipe, kg, km e rotas por motorista e ajudantes.",
    tag: "Equipe",
  },
  "Despesas": {
    title: "Custos",
    desc: "Consolida categorias de despesas, totais e rankings de custo.",
    tag: "Financeiro",
  },
};

const cadastroResources = {
  motoristas: {
    label: "Motoristas",
    endpoint: "motoristas",
    statusField: "status",
    activeStatus: "ATIVO",
    inactiveStatus: "INATIVO",
    password: true,
    defaults: {status: "ATIVO", perfil_app: "MOTORISTA"},
    fields: [
      ["nome", "NOME", "text", true],
      ["codigo", "CODIGO", "text", false],
      ["senha", "SENHA", "password", true],
      ["perfil_app", "PERFIL APP", "select", false, ["MOTORISTA", "ADMIN"]],
      ["cpf", "CPF", "text", false],
      ["telefone", "TELEFONE", "text", true],
      ["status", "STATUS", "select", false, ["ATIVO", "INATIVO"]],
    ],
  },
  vendedores: {
    label: "Vendedores",
    endpoint: "vendedores",
    statusField: "status",
    activeStatus: "ATIVO",
    inactiveStatus: "DESATIVADO",
    password: true,
    defaults: {status: "ATIVO"},
    fields: [
      ["codigo", "CODIGO", "text", true],
      ["nome", "NOME", "text", true],
      ["senha", "SENHA", "password", true],
      ["telefone", "TELEFONE", "text", false],
      ["cidade_base", "CIDADE BASE", "text", false],
      ["status", "STATUS", "select", true, ["ATIVO", "DESATIVADO"]],
    ],
  },
  usuarios: {
    label: "Usuarios",
    endpoint: "usuarios",
    password: true,
    defaults: {permissoes: "OPERADOR"},
    fields: [
      ["nome", "NOME", "text", true],
      ["senha", "SENHA", "password", true],
      ["permissoes", "PERMISSOES", "select", true, ["ADMIN", "GERENTE", "OPERADOR", "VISUALIZADOR"]],
      ["cpf", "CPF", "text", false],
      ["telefone", "TELEFONE", "text", false],
    ],
  },
  veiculos: {
    label: "Veiculos",
    endpoint: "veiculos",
    statusField: "status",
    activeStatus: "ATIVO",
    defaults: {status: "ATIVO"},
    fields: [
      ["placa", "PLACA", "text", true],
      ["modelo", "MODELO", "text", true],
      ["capacidade_cx", "CAPACIDADE (CX)", "number", true],
      ["status", "STATUS", "select", true, ["ATIVO", "DESATIVADO"]],
    ],
  },
  caixas: {
    label: "Caixas",
    endpoint: "caixas",
    statusField: "status",
    activeStatus: "EM_ESTOQUE",
    inactiveStatus: "QUEBRADA",
    defaults: {status: "EM_ESTOQUE"},
    fields: [
      ["codigo", "CODIGO RASTREIO", "text", true],
      ["lote", "LOTE", "text", true],
      ["cor", "COR", "text", true],
      ["veiculo_placa", "VEICULO VINCULADO", "text", false],
      ["status", "STATUS", "select", true, ["EM_ESTOQUE", "VINCULADA", "EM_USO", "QUEBRADA", "BAIXADA"]],
      ["data_compra", "DATA COMPRA", "date", false],
      ["observacao", "OBSERVACAO / QUEBRA", "text", false],
    ],
  },
  ajudantes: {
    label: "Ajudantes",
    endpoint: "ajudantes",
    statusField: "status",
    activeStatus: "ATIVO",
    inactiveStatus: "DESATIVADO",
    defaults: {status: "ATIVO"},
    fields: [
      ["nome", "NOME", "text", true],
      ["sobrenome", "SOBRENOME", "text", true],
      ["telefone", "TELEFONE", "text", true],
      ["status", "STATUS", "select", true, ["ATIVO", "DESATIVADO"]],
    ],
  },
  clientes: {
    label: "Clientes",
    endpoint: "clientes",
    fields: [
      ["cod_cliente", "COD CLIENTE", "text", true],
      ["nome_cliente", "NOME CLIENTE", "text", true],
      ["endereco", "ENDERECO", "text", false],
      ["bairro", "BAIRRO", "text", false],
      ["cidade", "CIDADE", "text", false],
      ["uf", "UF", "text", false],
      ["telefone", "TELEFONE", "text", false],
      ["rota", "ROTA", "text", false],
      ["vendedor", "VENDEDOR", "text", false],
    ],
  },
  fornecedores: {
    label: "Fornecedores",
    endpoint: "fornecedores",
    statusField: "status",
    activeStatus: "ATIVO",
    inactiveStatus: "INATIVO",
    defaults: {status: "ATIVO", perfil_fornecedor: "OUTROS"},
    fields: [
      ["razao_social", "RAZAO SOCIAL", "text", true],
      ["nome_fantasia", "NOME FANTASIA", "text", false],
      ["documento", "CPF/CNPJ", "text", true],
      ["tipo_pessoa", "TIPO", "select", false, ["CNPJ", "CPF"]],
      ["perfil_fornecedor", "PERFIL", "select", true, []],
      ["telefone", "TELEFONE", "text", false],
      ["email", "E-MAIL", "email", false],
      ["cidade", "CIDADE", "text", false],
      ["uf", "UF", "text", false],
      ["status", "STATUS", "select", true, ["ATIVO", "INATIVO"]],
      ["observacao", "OBSERVACAO", "text", false],
      ["certificado_status", "CERTIFICADO", "text", false],
      ["certificado_nome", "CERT. NOME", "text", false],
      ["certificado_instalado_em", "CERT. INSTALADO EM", "text", false],
    ],
  },
  produtos: {
    label: "Produtos",
    endpoint: "produtos",
    statusField: "status",
    activeStatus: "ATIVO",
    inactiveStatus: "INATIVO",
    defaults: {
      categoria: "AVES",
      unidade: "KG",
      unidade_estoque: "KG",
      controla_estoque_fisico: 1,
      controla_estoque_fiscal: 1,
      status: "ATIVO",
    },
    fields: [
      ["codigo", "CODIGO", "text", true],
      ["nome", "NOME", "text", true],
      ["descricao", "DESCRICAO", "text", false],
      ["categoria", "CATEGORIA", "select", true, ["AVES", "INSUMOS", "EMBALAGENS", "SERVICOS", "OUTROS"]],
      ["unidade", "UNIDADE", "select", true, ["KG", "CX", "UN", "PC", "LT"]],
      ["unidade_estoque", "UNID. ESTOQUE", "select", true, ["KG", "CX", "UN", "PC", "LT"]],
      ["controla_estoque_fisico", "EST. FISICO", "select", true, [1, 0]],
      ["controla_estoque_fiscal", "EST. FISCAL", "select", true, [1, 0]],
      ["estoque_min_kg", "MIN KG", "number", false],
      ["estoque_min_caixas", "MIN CX", "number", false],
      ["ncm", "NCM", "text", false],
      ["cest", "CEST", "text", false],
      ["cfop_entrada", "CFOP ENT.", "text", false],
      ["cfop_saida", "CFOP SAI.", "text", false],
      ["ean", "EAN", "text", false],
      ["custo_padrao", "CUSTO", "number", false],
      ["preco_padrao", "PRECO", "number", false],
      ["status", "STATUS", "select", true, ["ATIVO", "INATIVO"]],
    ],
  },
};

const moduleCatalog = {
  cadastros: {
    title: "Cadastros",
    subtitle: "Base operacional",
    description: "No desktop, concentra os cadastros usados pelas rotinas de planejamento, rotas e financeiro.",
    migration: "CRUD operacional portado para a web com o fluxo do desktop: NOVO, ALTERAR status, LIBERAR/BLOQUEAR, senha e bloqueios principais de exclusao.",
    features: [
      ["Clientes", "Cadastro, importacao e consulta de clientes."],
      ["Motoristas", "Cadastro de motorista, acesso e dados operacionais."],
      ["Veiculos", "Cadastro e controle da frota."],
      ["Ajudantes", "Cadastro e composicao de equipes."],
      ["Vendedores", "Cadastro e acesso do app vendedor."],
      ["Equipes", "Relacionamento entre motorista, veiculo e ajudantes."],
    ],
  },
  rotas: {
    title: "Rotas",
    subtitle: "Acompanhamento operacional",
    description: "No desktop, acompanha rotas ativas, status, GPS, carregamento, entrega e finalizacao.",
    migration: "Monitoramento de rotas ativo na web com status, ultimo GPS, atualizacao automatica e abertura em mapa.",
    features: [
      ["Rotas ativas", "Lista e acompanhamento das rotas em aberto."],
      ["Status operacional", "Inicio, carregamento, entrega, ocorrencias e finalizacao."],
      ["GPS", "Pings, rastreio e visualizacao de posicao."],
      ["Substituicoes", "Trocas de motorista/veiculo quando necessario."],
    ],
  },
  importarVendas: {
    title: "Importar Pedidos",
    subtitle: "Entrada de pedidos",
    description: "No desktop, importa planilhas e prepara pedidos para vinculacao com planejamentos.",
    migration: "Upload web, validacao, marcacao e vinculo com planejamento ativo migrados para o navegador.",
    features: [
      ["Importar Excel", "Leitura e validacao da planilha."],
      ["Marcar pedidos", "Selecao dos pedidos que entram no planejamento."],
      ["Vincular planejamento", "Associacao dos pedidos importados a um planejamento ativo."],
    ],
  },
  programacao: {
    title: "Planejamento de Rota",
    subtitle: "Montagem da operacao",
    description: "No desktop, cria e edita planejamentos, itens, equipes, rankings e impressao da folha.",
    migration: "Criacao, edicao, exclusao, itens, equipe, estimativas e vinculo com pedidos migrados com API propria.",
    features: [
      ["Criar planejamento", "Dados base, data, rota, motorista, veiculo e equipe."],
      ["Itens do planejamento", "Inserir, remover e limpar clientes/linhas."],
      ["Ajudantes", "Selecionar composicao da equipe."],
      ["Editar/excluir", "Carregar planejamentos existentes para manutencao."],
      ["Impressao/PDF", "Folha do planejamento e pre-visualizacao."],
    ],
  },
  recebimentos: {
    title: "Recebimentos",
    subtitle: "Prestacao de contas",
    description: "No desktop, registra recebimentos por planejamento e fecha informacoes de retorno.",
    migration: "Lancamento de recebimentos, clientes manuais, zerar cliente e cabecalho de rota/diarias migrados para a web.",
    features: [
      ["Carregar planejamento", "Selecionar planejamento pendente."],
      ["Recebimentos", "Informar valores e situacao dos pedidos."],
      ["Imprimir PDF", "Geracao do documento de prestacao."],
      ["Finalizar", "Fechamento da prestacao conforme regras atuais."],
    ],
  },
  despesas: {
    title: "Custos e Despesas",
    subtitle: "Custos operacionais",
    description: "No desktop, registra custos por planejamento, rota, categoria e fechamento.",
    migration: "Lancamentos de custos, KM/media, NF, cedulas, PIX motorista e finalizacao do fechamento migrados para a web.",
    features: [
      ["Lancamentos", "Inserir custos por tipo/categoria."],
      ["NF e ocorrencias", "Dados fiscais e observacoes."],
      ["Resumo", "Totais por planejamento e motorista."],
      ["Finalizar rotina", "Encerramento financeiro do planejamento."],
    ],
  },
  escala: {
    title: "Escala",
    subtitle: "Planejamento de equipes",
    description: "No desktop, monta escala, resumo e recomendacoes para operacao.",
    migration: "Resumo operacional, KPIs, ranking de motoristas/ajudantes e recomendacoes migrados para consulta web.",
    features: [
      ["Grade de escala", "Visualizacao e organizacao das equipes."],
      ["Resumo", "Totais e disponibilidade."],
      ["Recomendacoes", "Apoio para tomada de decisao."],
    ],
  },
  centroCustos: {
    title: "Analise de Custos",
    subtitle: "Analise financeira",
    description: "No desktop, agrupa custos por veiculo, rota, motorista e periodo.",
    migration: "Resumo por veiculo, filtros de periodo/veiculo e grafico por custo/KM, custo/KG ou despesa total migrados para a web.",
    features: [
      ["Filtros", "Periodo, veiculo, rota e demais dimensoes."],
      ["Indicadores", "Totais e medias por centro de custo."],
      ["Grafico", "Visualizacao dos custos consolidados."],
    ],
  },
  compras: {
    title: "Entradas e Compras",
    subtitle: "Manifestador e entrada fiscal",
    description: "Recebe XML de NF-e, organiza compras por natureza de operacao e alimenta a base de fornecedores.",
    migration: "Base inicial do manifestador criada na web com importacao de XML e cadastro automatico do fornecedor pelo emitente da nota.",
    features: [
      ["Importar XML", "Registrar a NF-e recebida e manter o arquivo para consulta."],
      ["Manifestador", "Preparado para conexao SEFAZ via certificado digital."],
      ["Fornecedores", "Criacao automatica quando o emitente do XML ainda nao existe no cadastro."],
    ],
  },
  relatorios: {
    title: "Relatorios Operacionais",
    subtitle: "Consultas e exportacoes",
    description: "No desktop, concentra relatorios de rotina, KM, despesas, ocorrencias, folhas e Excel/PDF.",
    migration: "Resumo de planejamento, fechamento, ocorrencias, rotina, KM, custos, Excel, PDF e controle de status migrados para a web.",
    features: [
      ["Rotina", "Relatorio operacional por planejamento."],
      ["KM veiculos", "Analise de quilometragem."],
      ["Custos", "Relatorios financeiros por categoria."],
      ["Ocorrencias", "Indicadores por motorista/rota."],
      ["Exportacao", "PDF e Excel."],
    ],
  },
  backup: {
    title: "Backup / Exportar",
    subtitle: "Dados e restauracao",
    description: "No desktop, exporta dados, cria backup e restaura arquivos quando necessario.",
    migration: "Backup do banco, restauracao com confirmacao, download de backups e exportacao de pedidos importados migrados para a web.",
    features: [
      ["Backup", "Gerar copia do banco atual."],
      ["Restaurar", "Restauracao controlada com backup automatico."],
      ["Exportar", "Gerar arquivos de dados."],
    ],
  },
  permissoes: {
    title: "Gerenciar Permissoes",
    subtitle: "Controle de acesso",
    description: "No desktop, gerencia permissoes por usuario e modulo.",
    migration: "Catalogo de permissoes, filtro por modulo, concessao, revogacao e aplicacao de perfil migrados para a web.",
    features: [
      ["Usuarios", "Selecionar usuario alvo."],
      ["Permissoes", "Conceder ou remover acesso por modulo."],
      ["Auditoria", "Registrar alteracoes de acesso."],
    ],
  },
  saasAdmin: {
    title: "Admin SaaS",
    subtitle: "Planos, cobranca e auditoria",
    description: "No desktop, administra empresas, planos, pagamentos, assinaturas e logs SaaS.",
    migration: "Dashboard SaaS, troca de plano, status da empresa, cobrancas, baixa manual, suspensao por atraso, recursos e auditoria migrados para a web.",
    features: [
      ["Empresas", "Cadastro e status das empresas."],
      ["Planos", "Limites, recursos e precos."],
      ["Pagamentos", "Criacao e baixa manual."],
      ["Assinaturas", "Status, vencimento e suspensao."],
    ],
  },
  billing: {
    title: "Plano e Assinatura",
    subtitle: "Pacote, limites e solicitação de alteração",
    description: "Consulta o plano atual da empresa e permite solicitar upgrade ou downgrade para aprovação do administrador RotaHub.",
    migration: "Controle de pacote, limites de veículos, usuários, demonstração e solicitações de alteração disponíveis no navegador.",
    features: [
      ["Plano atual", "Status, vencimento e limites contratados."],
      ["Uso", "Veículos e usuários consumidos no pacote."],
      ["Solicitação", "Pedido formal de upgrade ou downgrade para o Owner."],
    ],
  },
  ferramentas: {
    title: "Ferramentas do Sistema",
    subtitle: "Manutencao",
    description: "No desktop, reune diagnostico, logs, backups e operacoes de manutencao.",
    migration: "Backups, logs do sistema, limpeza, verificacao de integridade e informacoes do banco migrados para a web.",
    features: [
      ["Diagnostico", "Informacoes do ambiente e banco."],
      ["Logs", "Consulta e limpeza de logs."],
      ["Integridade", "Verificacoes de consistencia."],
      ["Backup", "Apoio operacional para manutencao."],
    ],
  },
};

const saasFeatureLabels = {
  cadastros: "Cadastros",
  importar_vendas: "Importar pedidos",
  programacao: "Planejamento de rota",
  recebimentos: "Recebimentos",
  despesas: "Custos e Despesas",
  mortalidade: "Ocorrencias operacionais",
  centro_custos: "Centro de custos",
  relatorios: "Relatorios Operacionais",
  rotas: "Rotas",
  escala: "Escala",
  app_motorista: "App motorista",
  realtime_tracking: "Rastreamento",
  financial_reports: "Financeiro",
  advanced_reports: "Relatorios avancados",
  api_access: "API",
  custom_contract: "Contrato customizado",
  priority_support: "Suporte prioritario",
  private_deployment: "Implantacao privada",
};

const el = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  installTableEnhancer();
  restoreRememberedLogin();
  if (state.token) {
    loadSession();
  } else {
    showLogin();
  }
});

function bindEvents() {
  el("loginForm").addEventListener("submit", onLogin);
  el("loginRememberUser").addEventListener("change", onRememberLoginChange);
  el("logoutButton").addEventListener("click", logout);
  el("createUserForm").addEventListener("submit", onCreateUser);
  el("editUserForm").addEventListener("submit", onSaveUser);
  el("cadastroForm").addEventListener("submit", onSaveCadastroItem);
  el("newCadastroButton").addEventListener("click", startNewCadastro);
  el("changeCadastroButton").addEventListener("click", startCadastroEdit);
  el("saveCadastroButton").addEventListener("click", () => el("cadastroForm").requestSubmit());
  el("liberarCadastroButton").addEventListener("click", () => quickCadastroStatus(true));
  el("bloquearCadastroButton").addEventListener("click", () => quickCadastroStatus(false));
  el("passwordCadastroButton").addEventListener("click", openCadastroPasswordDialog);
  el("perfilFornecedorButton").addEventListener("click", createFornecedorPerfil);
  el("caixasBulkButton").addEventListener("click", openCaixasBulkDialog);
  el("caixasBulkForm").addEventListener("submit", createCaixasBulk);
  el("caixasMoveButton").addEventListener("click", openCaixasMoveDialog);
  el("caixasMoveForm").addEventListener("submit", movimentarCaixas);
  el("caixasHistoricoButton").addEventListener("click", openCaixaHistorico);
  el("caixasExportButton").addEventListener("click", exportCaixasCsv);
  el("caixasLabelsButton").addEventListener("click", printCaixasLabels);
  el("logisticaConfigSaveButton").addEventListener("click", saveLogisticaConfig);
  el("logisticaNovoPerfilButton").addEventListener("click", createLogisticaPerfil);
  el("logisticaNovaUnidadeButton").addEventListener("click", createLogisticaUnidade);
  el("logisticaNovaOcorrenciaButton").addEventListener("click", createLogisticaOcorrencia);
  el("logisticaPerfilCodigo").addEventListener("change", applySelectedLogisticaPerfil);
  el("fornecedorCertificadoCloseButton").addEventListener("click", closeFornecedorCertificadoDialog);
  el("fornecedorCertificadoForm").addEventListener("submit", onInstallFornecedorCertificado);
  el("fornecedorCertificadoFornecedor").addEventListener("change", onFornecedorCertificadoChange);
  el("deleteCadastroButton").addEventListener("click", deleteSelectedCadastroItem);
  el("cancelCadastroEdit").addEventListener("click", clearCadastroSelection);
  el("cadastroStatusFilter").addEventListener("change", () => {
    state.cadastroStatusFilter = el("cadastroStatusFilter").value;
    state.cadastroSelectedId = null;
    state.cadastroMode = "view";
    state.cadastroItems = [];
    renderCadastroForm();
    renderCadastroTable();
    loadCadastros();
  });
  el("clientesImportUploadForm").addEventListener("submit", onClientesImportUpload);
  el("clientesImportTemplateButton").addEventListener("click", downloadClientesImportTemplate);
  el("clientesImportRefreshButton").addEventListener("click", loadClientesImportRows);
  el("clientesImportInsertButton").addEventListener("click", insertClientesImportRow);
  el("clientesImportSaveButton").addEventListener("click", saveClientesImportRows);
  el("clientesOpenListButton").addEventListener("click", () => showClientesSection("lista"));
  el("clientesOpenHistoricoButton").addEventListener("click", () => showClientesSection("historico"));
  el("clientesOpenLocalizacaoButton").addEventListener("click", () => showClientesSection("localizacao"));
  el("clientesHistoricoLoadButton").addEventListener("click", loadClienteHistoricoSelecionado);
  el("clientesLocalizacaoLoadButton").addEventListener("click", loadClienteLocalizacaoSelecionado);
  document.querySelectorAll("[data-clientes-section]").forEach((button) => {
    button.addEventListener("click", () => showClientesSection(button.dataset.clientesSection || "hub"));
  });
  el("passwordCadastroForm").addEventListener("submit", onChangeCadastroPassword);
  el("importarVendasUploadForm").addEventListener("submit", onImportarVendasUpload);
  el("importarVendasSearchForm").addEventListener("submit", (event) => {
    event.preventDefault();
    loadVendasImportadas();
  });
  el("importarVendasRefreshButton").addEventListener("click", () => refreshImportarVendasData());
  el("importarVendasMarkButton").addEventListener("click", markSelectedVendasImportadas);
  el("importarVendasMarkAllButton").addEventListener("click", () => setAllVendasImportadas(true));
  el("importarVendasUnmarkAllButton").addEventListener("click", () => setAllVendasImportadas(false));
  el("importarVendasDeleteButton").addEventListener("click", deleteSelectedVendasImportadas);
  el("importarVendasClearButton").addEventListener("click", clearAllVendasImportadas);
  el("importarVendasLinkButton").addEventListener("click", linkSelectedVendasToProgramacao);
  el("importarVendasRows").addEventListener("change", updateImportarVendasTotals);
  el("rotasRefreshButton").addEventListener("click", loadRotasMonitoramento);
  el("rotasMapAllButton").addEventListener("click", openRotasMapAll);
  el("rotasMapSelectedButton").addEventListener("click", openRotasMapSelected);
  el("rotasRefreshInterval").addEventListener("change", onRotasRefreshIntervalChange);
  el("escalaRefreshButton").addEventListener("click", refreshEscalaCompleta);
  el("escalaPdfButton").addEventListener("click", downloadEscalaPdf);
  el("escalaPeriodo").addEventListener("change", onEscalaFilterChange);
  el("escalaStatus").addEventListener("change", onEscalaFilterChange);
  el("escalaFolgaTipo").addEventListener("change", loadEscalaFolgaPessoas);
  el("escalaFolgaInicio").addEventListener("change", onEscalaFolgaDateChange);
  el("escalaFolgaFim").addEventListener("change", onEscalaFolgaDateChange);
  el("escalaFolgaAplicarButton").addEventListener("click", aplicarEscalaFolga);
  document.querySelectorAll("[data-escala-tab]").forEach((button) => {
    button.addEventListener("click", () => setEscalaTab(button.dataset.escalaTab));
  });
  el("recebimentosRefreshButton").addEventListener("click", loadRecebimentosProgramacoes);
  el("recebimentosLoadButton").addEventListener("click", loadSelectedRecebimentos);
  el("recebimentosProgramacaoSelect").addEventListener("change", loadSelectedRecebimentos);
  el("recebimentosCabecalhoForm").addEventListener("submit", onSaveRecebimentosCabecalho);
  el("recebimentoForm").addEventListener("submit", onSaveRecebimento);
  el("recebimentoCodCliente").addEventListener("change", syncRecebimentoClienteFromCodigo);
  el("recebimentoCodCliente").addEventListener("input", syncRecebimentoClienteFromCodigo);
  el("recebimentoNomeCliente").addEventListener("change", syncRecebimentoClienteFromNome);
  el("recebimentoZerarButton").addEventListener("click", onZerarRecebimento);
  el("recebimentosPrintButton").addEventListener("click", downloadRecebimentosPdf);
  el("recebimentosDespesasButton").addEventListener("click", goRecebimentosToDespesas);
  el("despesasRefreshButton").addEventListener("click", loadDespesasProgramacoes);
  el("despesasLoadButton").addEventListener("click", loadSelectedDespesas);
  el("despesasProgramacaoSelect").addEventListener("change", loadSelectedDespesas);
  el("despesasAppMotoristaButton").addEventListener("click", openDespesasAppMotoristaDialog);
  el("despesaNovaButton").addEventListener("click", clearDespesaForm);
  el("despesaForm").addEventListener("submit", onSaveDespesa);
  el("despesaDeleteButton").addEventListener("click", onDeleteDespesa);
  el("despesasRotaForm").addEventListener("submit", onSaveDespesasRota);
  el("despesasNfForm").addEventListener("submit", onSaveDespesasNf);
  el("despesasFinanceiroForm").addEventListener("submit", onSaveDespesasFinanceiro);
  ["despAdiantamento", "despPixMotorista", ...DESPESAS_CEDULAS.map((ced) => `despCed${ced}`)].forEach((id) => {
    el(id).addEventListener("input", updateDespesasFinanceiroPreview);
    el(id).addEventListener("change", normalizeDespesasFinanceiroInput);
  });
  document.getElementById("despesasPrintButton")?.addEventListener("click", downloadDespesasPdf);
  el("despesasBackButton").addEventListener("click", goDespesasToRecebimentos);
  el("despesasFinalizarButton").addEventListener("click", onFinalizarDespesas);
  el("mortalidadeRefreshButton").addEventListener("click", loadMortalidade);
  el("mortalidadeFilterForm").addEventListener("submit", (event) => {
    event.preventDefault();
    loadMortalidade();
  });
  el("mortalidadeClearFiltersButton").addEventListener("click", clearMortalidadeFilters);
  el("mortalidadeManualOpenButton").addEventListener("click", openMortalidadeManualDialog);
  el("mortalidadeManualCloseButton").addEventListener("click", closeMortalidadeManualDialog);
  el("mortalidadeManualForm").addEventListener("submit", onSaveMortalidadeManual);
  el("mortalidadeManualClearButton").addEventListener("click", clearMortalidadeManualForm);
  ["manualMortPreco", "manualMortMedia", "manualMortAves"].forEach((id) => {
    el(id).addEventListener("input", updateMortalidadeManualPreview);
  });
  ["despKmInicial", "despKmFinal", "despLitros"].forEach((id) => {
    el(id).addEventListener("input", updateDespesasRotaPreview);
  });
  el("centroCustosRefreshButton").addEventListener("click", loadCentroCustos);
  el("centroCustosPeriodo").addEventListener("change", loadCentroCustosResumo);
  el("centroCustosVeiculo").addEventListener("change", onCentroCustosVeiculoFilterChange);
  el("centroCustosMetric").addEventListener("change", loadCentroCustosResumo);
  el("centroCustosDespesasRotaExportButton").addEventListener("click", exportCentroCustosDespesasRotaCsv);
  el("centroCustosDespesasRotaBusca").addEventListener("input", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaFiltroRota").addEventListener("change", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaFiltroTipo").addEventListener("change", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaMes").addEventListener("change", onCentroCustosDespesasRotaMesChange);
  el("centroCustosDespesasRotaDataIni").addEventListener("change", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaDataFim").addEventListener("change", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaValorMin").addEventListener("input", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaValorMax").addEventListener("input", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaOrdenacao").addEventListener("change", () => renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {}));
  el("centroCustosDespesasRotaLimparButton").addEventListener("click", clearCentroCustosDespesasRotaFilters);
  document.querySelectorAll("[data-rota-quick]").forEach((button) => {
    button.addEventListener("click", () => toggleCentroCustosDespesasRotaQuick(button.dataset.rotaQuick || ""));
  });
  document.querySelectorAll("[data-rota-tab]").forEach((button) => {
    button.addEventListener("click", () => setCentroCustosDespesasRotaTab(button.dataset.rotaTab || "programacoes"));
  });
  el("centroCustosDespesaVeiculoForm").addEventListener("submit", onSaveCentroCustosDespesaVeiculo);
  el("centroDespesaLimparButton").addEventListener("click", clearCentroCustosDespesaVeiculoForm);
  el("centroDespesaControleTipo").addEventListener("change", updateCentroDespesaControleFields);
  el("comprasRefreshButton").addEventListener("click", loadComprasNfe);
  el("comprasCarregarButton").addEventListener("click", loadComprasNfe);
  el("comprasNatureza").addEventListener("change", loadComprasNfe);
  el("comprasBusca").addEventListener("input", renderComprasNfe);
  el("comprasImportXmlButton").addEventListener("click", () => el("comprasXmlFile").click());
  el("comprasXmlFile").addEventListener("change", onComprasImportXml);
  el("comprasDownloadXmlButton").addEventListener("click", downloadComprasXml);
  el("comprasManifestadorButton").addEventListener("click", () => notify("Manifestador SEFAZ preparado para ativacao com certificado digital do fornecedor/empresa."));
  el("comprasCertificadoButton").addEventListener("click", openFornecedorCertificadoDialog);
  el("comprasPorChaveButton").addEventListener("click", () => notify("Consulta por chave sera ligada ao servico fiscal na proxima etapa."));
  el("comprasManualButton").addEventListener("click", () => notify("Entrada manual fiscal sera a proxima rotina desta tela."));
  el("comprasAcoesButton").addEventListener("click", () => notify("Selecione uma nota para baixar XML ou confirmar entrada."));
  el("comprasDownloadPdfButton").addEventListener("click", () => notify("Download de DANFE/PDF depende da integracao fiscal ou arquivo anexado."));
  el("comprasSyncProgramacoesButton").addEventListener("click", syncComprasSaidasProgramacoes);
  el("comprasAjusteButton").addEventListener("click", ajustarComprasEstoqueFisico);
  el("comprasZerarFisicoButton").addEventListener("click", zerarComprasEstoqueFisico);
  el("comprasConfigNfeButton").addEventListener("click", () => switchView("cadastros"));
  el("comprasConfirmarButton").addEventListener("click", confirmarComprasEntradaEstoque);
  el("comprasVoltarButton").addEventListener("click", () => switchView("dashboard"));
  el("relatoriosRefreshButton").addEventListener("click", loadRelatorios);
  el("relatoriosFilterForm").addEventListener("submit", (event) => {
    event.preventDefault();
    loadRelatoriosProgramacoes();
    if (relatorioTipoPorNotaFiscal() && clean(el("relatoriosNf").value)) {
      loadRelatoriosResumo();
    }
  });
  el("relatoriosTipo").addEventListener("change", onRelatoriosTipoChange);
  el("relatoriosNf").addEventListener("input", updateRelatoriosActionState);
  el("relatoriosClearButton").addEventListener("click", clearRelatoriosFilters);
  el("relatoriosProgramacaoSelect").addEventListener("change", () => {
    updateRelatoriosActionState();
    renderRelatoriosContextPanel(state.relatoriosResumo, selectedRelatorioProgramacaoInfo());
  });
  el("relatoriosResumoButton").addEventListener("click", loadRelatoriosResumo);
  el("relatoriosExcelButton").addEventListener("click", downloadRelatoriosExcel);
  el("relatoriosPdfButton").addEventListener("click", downloadRelatoriosPdf);
  el("relatoriosPrintProgramacaoButton").addEventListener("click", () => downloadRelatoriosDocumento("programacao"));
  el("relatoriosPrintPrestacaoButton").addEventListener("click", () => downloadRelatoriosDocumento("prestacao"));
  el("relatoriosPrintRomaneiosButton").addEventListener("click", () => downloadRelatoriosDocumento("romaneios"));
  el("relatoriosFinalizarButton").addEventListener("click", finalizarRelatoriosRota);
  el("relatoriosReabrirButton").addEventListener("click", reabrirRelatoriosRota);
  el("relatoriosShowRecebimentos").addEventListener("change", refreshRelatoriosResumoIfLoaded);
  el("relatoriosShowDespesas").addEventListener("change", refreshRelatoriosResumoIfLoaded);
  el("relatoriosTableSearch").addEventListener("input", renderRelatoriosResumo);
  el("backupRefreshButton").addEventListener("click", loadSystemTools);
  el("backupCreateButton").addEventListener("click", createSystemBackup);
  el("backupMigrationButton").addEventListener("click", downloadMigrationPackage);
  el("backupExportVendasButton").addEventListener("click", exportBackupVendas);
  el("backupRestoreForm").addEventListener("submit", restoreBackupUpload);
  el("ferramentasRefreshButton").addEventListener("click", loadSystemTools);
  el("ferramentasBackupButton").addEventListener("click", createSystemBackup);
  el("ferramentasIntegrityButton").addEventListener("click", checkSystemIntegrity);
  el("ferramentasClearLogsButton").addEventListener("click", clearSystemLogs);
  el("diariasConfigForm").addEventListener("submit", saveDiariasConfig);
  el("permissoesRefreshButton").addEventListener("click", loadPermissoes);
  el("permissoesModuloFilter").addEventListener("change", renderPermissoesView);
  el("permissoesApplyPerfilButton").addEventListener("click", applyPermissionProfile);
  el("billingRefreshButton").addEventListener("click", loadBilling);
  el("trialNoticeBillingButton").addEventListener("click", () => switchView("billing"));
  el("saasAdminRefreshButton").addEventListener("click", loadSaasAdmin);
  el("saasAdminCompanySelect").addEventListener("change", () => {
    state.saasAdmin.selectedCompanyId = Number(el("saasAdminCompanySelect").value || 0) || null;
    loadSaasAdmin();
  });
  el("saasAdminStatusActiveButton").addEventListener("click", () => updateSaasCompanyStatus("active"));
  el("saasAdminStatusSuspendedButton").addEventListener("click", () => updateSaasCompanyStatus("suspended"));
  el("saasAdminPlanForm").addEventListener("submit", changeSaasCompanyPlan);
  el("saasAdminPaymentForm").addEventListener("submit", createSaasPayment);
  el("saasAdminOverdueButton").addEventListener("click", runSaasOverdueCheck);
  el("programacaoForm").addEventListener("submit", onSaveProgramacao);
  el("progRefreshOptions").addEventListener("click", editarPlanejamentoAtual);
  el("progNewButton").addEventListener("click", clearProgramacaoForm);
  document.getElementById("progPdfButton")?.addEventListener("click", downloadProgramacaoPdf);
  document.getElementById("progAdiantamentoPdfButton")?.addEventListener("click", downloadProgramacaoAdiantamentoRecibo);
  el("progRomaneiosButton").addEventListener("click", downloadProgramacaoRomaneios);
  el("progDeleteButton").addEventListener("click", deleteSelectedProgramacao);
  el("progLoadVendasButton").addEventListener("click", loadSelectedVendasIntoProgramacao);
  el("progSuggestButton").addEventListener("click", suggestProgramacaoRoute);
  el("progAddItemButton").addEventListener("click", () => addProgramacaoItemRow());
  el("progRemoveItemButton").addEventListener("click", removeSelectedProgramacaoRows);
  el("progClearItemsButton").addEventListener("click", clearProgramacaoItems);
  el("progItemsRows").addEventListener("input", updateProgramacaoItemTotals);
  el("progTipoEstimativa").addEventListener("change", () => updateProgramacaoEstimateLabel({ syncFrete: true }));
  el("progTransbordoModalidade").addEventListener("change", () => updateProgramacaoEstimateLabel({ syncEstimativa: true }));
  el("progAjudantesButton").addEventListener("click", openProgramacaoAjudantesDialog);
  el("progAjudantesApplyButton").addEventListener("click", applyProgramacaoAjudantesDialog);
  el("progRankingPeriodo").addEventListener("change", loadProgramacaoRankings);
  el("progRankingRefreshButton").addEventListener("click", loadProgramacaoRankings);
  el("includeInactive").addEventListener("change", loadUsers);
  el("auditFilterForm").addEventListener("submit", (event) => {
    event.preventDefault();
    loadAudit();
  });

  document.querySelectorAll(".nav-item[data-view]").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.view));
  });

  document.querySelectorAll("[data-home-nav]").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.homeNav));
  });

  document.addEventListener("click", (event) => {
    const action = event.target.dataset.action;
    if (!action) return;

    if (action === "refresh-dashboard") loadDashboard();
    if (action === "refresh-users") loadUsers();
    if (action === "refresh-audit") loadAudit();
    if (action === "refresh-cadastros") loadCadastros();
    if (action === "refresh-rotas") loadRotasMonitoramento();
    if (action === "refresh-escala") refreshEscalaCompleta();
    if (action === "refresh-recebimentos") loadSelectedRecebimentos();
    if (action === "refresh-despesas") loadSelectedDespesas();
    if (action === "refresh-centro-custos") loadCentroCustosResumo();
    if (action === "close-centro-veiculo-dialog") closeCentroCustosVeiculoDetalhe();
    if (action === "close-centro-despesas-rota-dialog") closeCentroCustosDespesasRotaDialog();
    if (action === "refresh-relatorios") loadRelatoriosResumo();
    if (action === "refresh-system-tools") loadSystemTools();
    if (action === "download-installer") downloadSystemInstaller();
    if (action === "refresh-permissoes") loadPermissoes();
    if (action === "refresh-saas-admin") loadSaasAdmin();
    if (action === "refresh-importar-vendas") refreshImportarVendasData();
    if (action === "refresh-programacao") loadProgramacoes();
    if (action === "close-home-route-dialog") closeHomeRouteDialog();
    if (action === "close-home-trace-dialog") closeHomeTraceDialog();
    if (action === "close-despesas-app-dialog") closeDespesasAppMotoristaDialog();
    if (action === "close-mortalidade-manual-dialog") closeMortalidadeManualDialog();
    if (action === "close-mortalidade-doa-detail") closeMortalidadeDoaDetailDialog();
    if (action === "close-prog-ajudantes-dialog") closeProgramacaoAjudantesDialog();
    if (action === "close-user-dialog") closeUserDialog();
    if (action === "close-password-dialog") closeCadastroPasswordDialog();
    if (action === "close-caixas-bulk-dialog") closeCaixasBulkDialog();
    if (action === "close-caixas-move-dialog") closeCaixasMoveDialog();
    if (action === "close-caixas-historico-dialog") closeCaixaHistoricoDialog();
  });
  document.addEventListener("keydown", handleGlobalShortcuts);
}

function isTypingTarget(target) {
  if (!target) return false;
  const tag = String(target.tagName || "").toUpperCase();
  return tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || target.isContentEditable;
}

function clickIfEnabled(id) {
  const button = el(id);
  if (button && !button.disabled && !button.classList.contains("hidden")) {
    button.click();
    return true;
  }
  return false;
}

function handleGlobalShortcuts(event) {
  const key = event.key;
  const typing = isTypingTarget(event.target);
  if (event.ctrlKey && key.toLowerCase() === "s") {
    event.preventDefault();
    saveCurrentView();
    return;
  }
  if (event.ctrlKey && ["c", "v", "x", "a"].includes(key.toLowerCase()) && typing) return;
  if (event.altKey && !event.ctrlKey && !event.shiftKey) {
    const nav = {
      d: "dashboard",
      c: "cadastros",
      p: "programacao",
      o: "compras",
      r: "recebimentos",
      i: "importarVendas",
      e: "escala",
      l: "relatorios",
    }[key.toLowerCase()];
    if (nav) {
      event.preventDefault();
      switchView(nav);
    }
    return;
  }
  if (key === "Escape") {
    const openDialog = document.querySelector("dialog[open]");
    if (openDialog) {
      event.preventDefault();
      openDialog.close();
      return;
    }
    if (state.currentView === "cadastros" && state.cadastroMode !== "view") {
      event.preventDefault();
      clearCadastroSelection();
    }
    return;
  }
  if (key === "F2") {
    event.preventDefault();
    if (state.currentView === "cadastros") {
      currentCadastroItem() ? startCadastroEdit() : startNewCadastro();
    } else if (state.currentView === "programacao") {
      clearProgramacaoForm();
    } else if (state.currentView === "despesas") {
      clearDespesaForm();
    }
    return;
  }
  if (key === "F3") {
    event.preventDefault();
    focusVisibleSearch();
    return;
  }
  if ((key === "Enter" || key === " ") && !typing) {
    const row = event.target && event.target.closest && event.target.closest("tr[tabindex]");
    if (row) {
      event.preventDefault();
      row.click();
    }
  }
}

function saveCurrentView() {
  if (state.currentView === "cadastros") {
    if (state.cadastroMode !== "view") el("cadastroForm").requestSubmit();
    return;
  }
  if (state.currentView === "programacao") {
    el("programacaoForm").requestSubmit();
    return;
  }
  if (state.currentView === "recebimentos") {
    el("recebimentoForm").requestSubmit();
    return;
  }
  if (state.currentView === "despesas") {
    el("despesaForm").requestSubmit();
  }
}

function focusVisibleSearch() {
  const view = document.querySelector(".view:not(.hidden)");
  const target = view && (
    view.querySelector("input[type='search']:not(:disabled)") ||
    view.querySelector("input[placeholder*='Busca']:not(:disabled), input[placeholder*='busca']:not(:disabled)")
  );
  if (target) {
    target.focus();
    if (target.select) target.select();
  }
}

async function onLogin(event) {
  event.preventDefault();
  setLoginError("");
  const button = el("loginButton");
  button.disabled = true;

  const form = new FormData(event.currentTarget);
  const username = clean(form.get("username"));
  const body = new URLSearchParams();
  body.set("username", username);
  body.set("password", form.get("password"));

  try {
    const response = await fetch(`${API_BASE}/auth/login`, {
      method: "POST",
      headers: {"Content-Type": "application/x-www-form-urlencoded"},
      body,
    });
    if (!response.ok) {
      throw new Error("Usuario ou senha invalidos.");
    }
    const data = await response.json();
    persistRememberedLogin(username);
    storeAuthToken(data);
    await loadSession();
  } catch (error) {
    setLoginError(error.message || "Falha no login.");
  } finally {
    button.disabled = false;
  }
}

function restoreRememberedLogin() {
  const username = localStorage.getItem(REMEMBER_USER_KEY) || "";
  if (!username) return;
  el("loginUsername").value = username;
  el("loginRememberUser").checked = true;
  el("loginPassword").focus();
}

function persistRememberedLogin(username) {
  if (el("loginRememberUser").checked && username) {
    localStorage.setItem(REMEMBER_USER_KEY, username);
    return;
  }
  localStorage.removeItem(REMEMBER_USER_KEY);
}

function onRememberLoginChange(event) {
  if (!event.currentTarget.checked) {
    localStorage.removeItem(REMEMBER_USER_KEY);
  }
}

function storeAuthToken(data) {
  state.token = data.access_token || "";
  const expiresIn = Math.max(Number(data.expires_in || 0), 60);
  state.tokenExpiresAt = Date.now() + (expiresIn * 1000);
  sessionStorage.setItem(TOKEN_KEY, state.token);
  sessionStorage.setItem(TOKEN_EXPIRES_KEY, String(state.tokenExpiresAt));
}

async function loadSession() {
  try {
    state.currentUser = await apiRequest("/users/me");
    state.planContext = await apiRequest("/auth/plan-context");
    showApp();
    renderTrialNotice();
    scheduleAuthRefresh();
    await loadLogisticaConfig({silent: true});
    await loadDashboard();
  } catch (error) {
    logout();
  }
}

function showLogin() {
  el("loginScreen").classList.remove("hidden");
  el("appShell").classList.add("hidden");
}

function showApp() {
  el("loginScreen").classList.add("hidden");
  el("appShell").classList.remove("hidden");
  el("currentUserName").textContent = `${state.currentUser.nome || state.currentUser.username} (${state.currentUser.permissoes})`;
  applyPlanNavigation();
  switchView(state.currentView || "dashboard", {skipLoad: true});
}

function logout(event) {
  if (event && typeof event.preventDefault === "function") {
    event.preventDefault();
  }
  clearAuthRefreshTimer();
  const token = state.token;
  if (token) {
    fetch(`${API_BASE}/auth/logout`, {
      method: "POST",
      headers: {Authorization: `Bearer ${token}`},
    }).catch(() => {});
  }
  state.token = "";
  state.tokenExpiresAt = 0;
  state.currentUser = null;
  state.planContext = null;
  state.billing = {company: null, subscription: null, usage: {}, plans: [], pending_requests: []};
  sessionStorage.removeItem(TOKEN_KEY);
  sessionStorage.removeItem(TOKEN_EXPIRES_KEY);
  showLogin();
}

function clearAuthRefreshTimer() {
  if (authRefreshTimer) {
    window.clearTimeout(authRefreshTimer);
    authRefreshTimer = null;
  }
}

function scheduleAuthRefresh() {
  clearAuthRefreshTimer();
  if (!state.token || !state.tokenExpiresAt) return;
  const msToExpire = state.tokenExpiresAt - Date.now();
  if (msToExpire <= 0) {
    logout();
    return;
  }
  const delay = Math.max(30000, msToExpire - 120000);
  authRefreshTimer = window.setTimeout(() => {
    refreshAuthToken().catch(() => logout());
  }, delay);
}

async function refreshAuthToken() {
  if (!state.token) return;
  const response = await fetch(`${API_BASE}/auth/refresh`, {
    method: "POST",
    headers: {Authorization: `Bearer ${state.token}`},
  });
  if (!response.ok) {
    throw new Error("Sessao expirada.");
  }
  storeAuthToken(await response.json());
  scheduleAuthRefresh();
}

async function ensureFreshToken() {
  if (!state.token || !state.tokenExpiresAt) return;
  if (state.tokenExpiresAt - Date.now() > 60000) return;
  authRefreshPromise = authRefreshPromise || refreshAuthToken().finally(() => {
    authRefreshPromise = null;
  });
  await authRefreshPromise;
}

function setLoginError(message) {
  const node = el("loginError");
  node.textContent = message;
  node.classList.toggle("hidden", !message);
}

async function apiRequest(path, options = {}) {
  await ensureFreshToken();
  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${state.token}`);
  if (options.body && !(options.body instanceof URLSearchParams) && !(options.body instanceof FormData)) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers,
  });

  if (response.status === 401) {
    logout();
    throw new Error("Sessao expirada.");
  }

  let data = null;
  const text = await response.text();
  if (text) {
    data = JSON.parse(text);
  }

  if (!response.ok) {
    const detail = data && data.detail ? data.detail : "Requisicao recusada.";
    throw new Error(Array.isArray(detail) ? detail.map((item) => item.msg).join("; ") : detail);
  }

  return data;
}

async function apiBlobRequest(path, options = {}) {
  await ensureFreshToken();
  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${state.token}`);
  const response = await fetch(`${API_BASE}${path}`, {...options, headers});
  if (response.status === 401) {
    logout();
    throw new Error("Sessao expirada.");
  }
  if (!response.ok) {
    const text = await response.text();
    let detail = "Requisicao recusada.";
    try {
      const data = text ? JSON.parse(text) : null;
      detail = data && data.detail ? data.detail : detail;
    } catch (_error) {
      detail = text || detail;
    }
    throw new Error(Array.isArray(detail) ? detail.map((item) => item.msg).join("; ") : detail);
  }
  const blob = await response.blob();
  const disposition = response.headers.get("content-disposition") || "";
  const match = disposition.match(/filename="?([^"]+)"?/i);
  return {blob, filename: match ? match[1] : ""};
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename || "download";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function openPdfPreviewWindow(documents, title = "Previsualizacao") {
  const docs = (documents || []).filter((doc) => doc && doc.blob);
  if (!docs.length) return;
  const popup = window.open("", "_blank", "width=1180,height=860");
  if (!popup) {
    docs.forEach((doc) => downloadBlob(doc.blob, doc.filename));
    notify("O navegador bloqueou a pre-visualizacao. Os PDFs foram baixados para impressao.", true);
    return;
  }
  const safeDocs = docs.map((doc) => ({
    label: doc.label || "PDF",
    filename: doc.filename || "documento.pdf",
    url: URL.createObjectURL(doc.blob),
  }));
  popup.document.open();
  popup.document.write(pdfPreviewHtml(safeDocs, title));
  popup.document.close();
  popup.addEventListener("beforeunload", () => {
    safeDocs.forEach((doc) => URL.revokeObjectURL(doc.url));
  });
}

function pdfPreviewHtml(documents, title) {
  const docsJson = JSON.stringify(documents).replace(/</g, "\\u003c");
  return `<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(title)}</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: Arial, sans-serif; color: #122033; background: #eef3f8; }
    header { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 12px 16px; background: #142235; color: #fff; }
    h1 { margin: 0; font-size: 18px; }
    .tabs, .actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    button, a { min-height: 34px; padding: 0 12px; border-radius: 6px; border: 1px solid #2dd4bf; background: #17324b; color: #fff; font-weight: 700; text-decoration: none; cursor: pointer; }
    button.active { background: #0f766e; }
    main { height: calc(100vh - 58px); }
    iframe { width: 100%; height: 100%; border: 0; background: #fff; }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>${escapeHtml(title)}</h1>
      <div id="tabs" class="tabs"></div>
    </div>
    <div class="actions">
      <a id="saveLink" download>Salvar PDF</a>
      <button id="printButton" type="button">Imprimir</button>
    </div>
  </header>
  <main><iframe id="viewer" title="PDF"></iframe></main>
  <script>
    const docs = ${docsJson};
    let current = 0;
    const tabs = document.getElementById("tabs");
    const viewer = document.getElementById("viewer");
    const saveLink = document.getElementById("saveLink");
    function render() {
      const doc = docs[current];
      viewer.src = doc.url;
      saveLink.href = doc.url;
      saveLink.download = doc.filename;
      [...tabs.children].forEach((button, index) => button.classList.toggle("active", index === current));
    }
    docs.forEach((doc, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = doc.label;
      button.addEventListener("click", () => { current = index; render(); });
      tabs.appendChild(button);
    });
    document.getElementById("printButton").addEventListener("click", () => {
      viewer.contentWindow.focus();
      viewer.contentWindow.print();
    });
    render();
  </script>
</body>
</html>`;
}

function switchView(view, options = {}) {
  if (!canAccessView(view)) {
    notify("Este recurso não está disponível no plano atual. Solicite um upgrade para liberar o módulo.", true);
    return;
  }
  if (view !== "rotas") stopRotasAutoRefresh();
  state.currentView = view;
  const titles = {
    dashboard: ["Painel", "Operacao e seguranca"],
    cadastros: ["Cadastros", "Base operacional"],
    rotas: ["Rotas", "Acompanhamento operacional"],
    escala: ["Escala", "Planejamento de equipes"],
    importarVendas: ["Importar Pedidos", "Entrada de pedidos"],
    programacao: ["Planejamento de Rota", "Montagem da operacao"],
    recebimentos: ["Recebimentos", "Fechamento operacional"],
    despesas: ["Custos e Despesas", "Custos operacionais"],
    mortalidade: ["Ocorrencias operacionais", "Registros por rota"],
    centroCustos: ["Analise de Custos", "Analise financeira"],
    compras: ["Entradas e Compras", "Manifestador e entrada fiscal"],
    relatorios: ["Relatorios Operacionais", "Consultas e exportacoes"],
    backup: ["Backup / Exportar", "Dados e restauracao"],
    billing: ["Plano e Assinatura", "Pacote, limites e solicitacoes"],
    permissoes: ["Gerenciar Permissoes", "Controle de acesso"],
    saasAdmin: ["Admin SaaS", "Planos, cobranca e auditoria"],
    ferramentas: ["Ferramentas do Sistema", "Manutencao"],
    users: ["Usuarios", "Cadastro e acesso"],
    audit: ["Auditoria", "Eventos administrativos"],
  };
  const module = moduleCatalog[view];

  document.querySelectorAll(".nav-item").forEach((button) => {
    button.classList.toggle("active", button.dataset.view === view);
  });
  syncNavGroups(view);
  document.querySelectorAll(".view").forEach((section) => {
    section.classList.add("hidden");
  });

  if (view === "cadastros") {
    el("cadastrosView").classList.remove("hidden");
    el("viewTitle").textContent = titles.cadastros[0];
    el("viewSubtitle").textContent = titles.cadastros[1];
    renderCadastroTabs();
    renderCadastroForm();
    if (!options.skipLoad) {
      loadCadastros();
    }
    return;
  }

  if (view === "programacao") {
    el("programacaoView").classList.remove("hidden");
    el("viewTitle").textContent = titles.programacao[0];
    el("viewSubtitle").textContent = titles.programacao[1];
    renderProgramacaoItems();
    if (!options.skipLoad) {
      loadProgramacaoOptions();
      loadProgramacaoRankings();
      loadProgramacoes();
    }
    return;
  }

  if (view === "rotas") {
    el("rotasView").classList.remove("hidden");
    el("viewTitle").textContent = titles.rotas[0];
    el("viewSubtitle").textContent = titles.rotas[1];
    if (!options.skipLoad) {
      loadRotasMonitoramento();
      startRotasAutoRefresh();
    }
    return;
  }

  if (view === "escala") {
    el("escalaView").classList.remove("hidden");
    el("viewTitle").textContent = titles.escala[0];
    el("viewSubtitle").textContent = titles.escala[1];
    if (!options.skipLoad) {
      initEscalaFolgaDates();
      refreshEscalaCompleta();
    }
    return;
  }

  if (view === "recebimentos") {
    el("recebimentosView").classList.remove("hidden");
    el("viewTitle").textContent = titles.recebimentos[0];
    el("viewSubtitle").textContent = titles.recebimentos[1];
    if (!options.skipLoad) loadRecebimentosProgramacoes();
    return;
  }

  if (view === "despesas") {
    el("despesasView").classList.remove("hidden");
    el("viewTitle").textContent = titles.despesas[0];
    el("viewSubtitle").textContent = titles.despesas[1];
    if (!options.skipLoad) loadDespesasProgramacoes();
    return;
  }

  if (view === "mortalidade") {
    el("mortalidadeView").classList.remove("hidden");
    el("viewTitle").textContent = currentLogisticaConfig().perda_label || titles.mortalidade[0];
    el("viewSubtitle").textContent = titles.mortalidade[1];
    applyLogisticaLabels();
    if (!options.skipLoad) loadMortalidade();
    return;
  }

  if (view === "centroCustos") {
    el("centroCustosView").classList.remove("hidden");
    el("viewTitle").textContent = titles.centroCustos[0];
    el("viewSubtitle").textContent = titles.centroCustos[1];
    if (!options.skipLoad) loadCentroCustos();
    return;
  }

  if (view === "compras") {
    el("comprasView").classList.remove("hidden");
    el("viewTitle").textContent = titles.compras[0];
    el("viewSubtitle").textContent = titles.compras[1];
    applyLogisticaLabels();
    if (!options.skipLoad) loadComprasNfe();
    return;
  }

  if (view === "relatorios") {
    el("relatoriosView").classList.remove("hidden");
    el("viewTitle").textContent = titles.relatorios[0];
    el("viewSubtitle").textContent = titles.relatorios[1];
    if (!options.skipLoad) loadRelatorios();
    return;
  }

  if (view === "backup") {
    el("backupView").classList.remove("hidden");
    el("viewTitle").textContent = titles.backup[0];
    el("viewSubtitle").textContent = titles.backup[1];
    if (!options.skipLoad) loadSystemTools();
    return;
  }

  if (view === "ferramentas") {
    el("ferramentasView").classList.remove("hidden");
    el("viewTitle").textContent = titles.ferramentas[0];
    el("viewSubtitle").textContent = titles.ferramentas[1];
    if (!options.skipLoad) {
      loadSystemTools();
      loadLogisticaConfig();
    }
    return;
  }

  if (view === "permissoes") {
    el("permissoesView").classList.remove("hidden");
    el("viewTitle").textContent = titles.permissoes[0];
    el("viewSubtitle").textContent = titles.permissoes[1];
    if (!options.skipLoad) loadPermissoes();
    return;
  }

  if (view === "billing") {
    el("billingView").classList.remove("hidden");
    el("viewTitle").textContent = titles.billing[0];
    el("viewSubtitle").textContent = titles.billing[1];
    if (!options.skipLoad) loadBilling();
    return;
  }

  if (view === "saasAdmin") {
    el("saasAdminView").classList.remove("hidden");
    el("viewTitle").textContent = titles.saasAdmin[0];
    el("viewSubtitle").textContent = titles.saasAdmin[1];
    if (!options.skipLoad) loadSaasAdmin();
    return;
  }

  if (view === "importarVendas") {
    el("importarVendasView").classList.remove("hidden");
    el("viewTitle").textContent = titles.importarVendas[0];
    el("viewSubtitle").textContent = titles.importarVendas[1];
    if (!options.skipLoad) refreshImportarVendasData();
    return;
  }

  if (module) {
    el("moduleView").classList.remove("hidden");
    el("viewTitle").textContent = module.title;
    el("viewSubtitle").textContent = module.subtitle;
    renderModuleView(module);
    return;
  }

  el(`${view}View`).classList.remove("hidden");
  el("viewTitle").textContent = titles[view][0];
  el("viewSubtitle").textContent = titles[view][1];

  if (options.skipLoad) return;
  if (view === "dashboard") loadDashboard();
  if (view === "users") loadUsers();
  if (view === "audit") loadAudit();
}

function canAccessView(view) {
  const feature = VIEW_PLAN_FEATURES[view];
  if (!feature) return true;
  const features = (state.planContext || {}).features || {};
  return Boolean(features[feature]);
}

function billingLimitLabel(limit) {
  return limit === null || limit === undefined || limit === "" ? "Sem limite" : `${Number(limit || 0).toLocaleString("pt-BR")}`;
}

function billingStatusLabel(value) {
  const labels = {
    active: "Ativa",
    trialing: "Demonstração",
    past_due: "Em atraso",
    suspended: "Suspensa",
    trial_expired: "Demonstração vencida",
    pending: "Pendente",
    approved: "Aprovada",
    rejected: "Recusada",
    upgrade: "Upgrade",
    downgrade: "Downgrade",
    change: "Troca",
  };
  const key = clean(value).toLowerCase();
  return labels[key] || clean(value) || "-";
}

function dateOnly(value) {
  const raw = clean(value);
  if (!raw) return null;
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function trialDaysRemaining(value) {
  const due = dateOnly(value);
  if (!due) return null;
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return Math.ceil((due.getTime() - start.getTime()) / 86400000);
}

function trialMessage(subscription) {
  if (clean(subscription?.status).toLowerCase() !== "trialing") return null;
  const days = trialDaysRemaining(subscription.next_due_date);
  const planName = subscription.plan_name || subscription.plan_code || "melhor plano";
  if (days === null) {
    return {
      title: "Periodo de teste ativo",
      text: `Sua demonstracao esta ativa no ${planName}. Acesse Plano e Assinatura para definir o pacote definitivo.`,
      daysLabel: "-",
    };
  }
  if (days < 0) {
    return {
      title: "Periodo de teste vencido",
      text: `Seu periodo de teste venceu. Escolha a assinatura para manter o acesso ao sistema.`,
      daysLabel: "Vencido",
    };
  }
  const dayText = days === 1 ? "1 dia restante" : `${days} dias restantes`;
  return {
    title: `Teste gratis: ${dayText}`,
    text: `Durante a demonstracao voce esta usando o ${planName}, com todos os recursos do melhor plano liberados.`,
    daysLabel: dayText,
  };
}

function currentSubscriptionForTrialNotice() {
  const billingSubscription = state.billing?.subscription || {};
  if (billingSubscription.status) return billingSubscription;
  const context = state.planContext || {};
  return {
    status: context.subscription_status,
    next_due_date: context.next_due_date,
    plan_name: context.plan_name,
    plan_code: context.plan_code,
  };
}

function renderTrialNotice() {
  const notice = el("trialNotice");
  if (!notice) return;
  const message = trialMessage(currentSubscriptionForTrialNotice());
  notice.classList.toggle("hidden", !message);
  if (!message) return;
  el("trialNoticeTitle").textContent = message.title;
  el("trialNoticeText").textContent = message.text;
}

async function loadBilling() {
  try {
    const data = await apiRequest("/billing/my-plan");
    state.billing = data || {};
    if (state.billing.subscription) {
      state.planContext = {
        ...(state.planContext || {}),
        plan_code: state.billing.subscription.plan_code,
        plan_name: state.billing.subscription.plan_name,
        subscription_status: state.billing.subscription.status,
        next_due_date: state.billing.subscription.next_due_date,
      };
    }
    renderBilling();
    renderTrialNotice();
  } catch (error) {
    notify(error.message, true);
  }
}

function billingUsageData() {
  const usage = state.billing.usage || {};
  const vehicles = usage.vehicles || {};
  return {
    vehicles: Number(vehicles.vehicle_count || 0),
    users: Number(usage.users || 0),
  };
}

function renderBilling() {
  const data = state.billing || {};
  const subscription = data.subscription || {};
  const usage = billingUsageData();
  const planName = subscription.plan_name || subscription.plan_code || "Sem plano";
  const vehicleLimit = subscription.plan_vehicle_limit;
  const userLimit = subscription.plan_user_limit;
  renderBillingTrialBanner(subscription);
  el("billingInfo").textContent = `${planName} / ${billingStatusLabel(subscription.status)} / vencimento ${subscription.next_due_date || "-"}`;
  el("billingSummary").innerHTML = [
    ["Plano atual", planName, billingStatusLabel(subscription.status)],
    ["Veículos", `${usage.vehicles} de ${billingLimitLabel(vehicleLimit)}`, "ativos no pacote"],
    ["Usuários", `${usage.users} de ${billingLimitLabel(userLimit)}`, "contas ativas"],
    ["Vencimento", subscription.next_due_date || "-", subscription.status === "trialing" ? "fim da demonstração" : "próxima cobrança"],
  ].map(([label, value, detail]) => `
    <article class="metric-card">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <small>${escapeHtml(detail || "")}</small>
    </article>
  `).join("");
  renderBillingPlans();
  renderBillingRequests();
}

function renderBillingTrialBanner(subscription) {
  const banner = el("billingTrialBanner");
  if (!banner) return;
  const message = trialMessage(subscription || {});
  banner.classList.toggle("hidden", !message);
  if (!message) {
    banner.innerHTML = "";
    return;
  }
  banner.innerHTML = `
    <div>
      <span>Periodo de teste</span>
      <strong>${escapeHtml(message.title)}</strong>
      <p>${escapeHtml(message.text)}</p>
    </div>
    <em>${escapeHtml(message.daysLabel)}</em>
  `;
}

function renderBillingPlans() {
  const currentCode = clean((state.billing.subscription || {}).plan_code).toLowerCase();
  const trialing = clean((state.billing.subscription || {}).status).toLowerCase() === "trialing";
  const pending = (state.billing.pending_requests || []).some((item) => clean(item.status).toLowerCase() === "pending");
  const plans = state.billing.plans || [];
  el("billingPlans").innerHTML = plans.map((plan) => {
    const code = clean(plan.code).toLowerCase();
    const current = code === currentCode && !trialing;
    const disabled = pending || (current && !trialing);
    const actionLabel = pending
      ? "Solicitacao pendente"
      : current && !trialing
        ? "Plano atual"
        : trialing
          ? "Escolher este plano"
          : "Solicitar alteracao";
    const features = Object.entries(plan.features || {}).filter(([, enabled]) => enabled).slice(0, 7);
    return `
      <article class="billing-plan-card ${current ? "current" : ""}">
        <div>
          <span>${current ? (trialing ? "Liberado no teste" : "Plano atual") : "Pacote"}</span>
          <h3>${escapeHtml(plan.name || plan.code)}</h3>
          <strong>${escapeHtml(formatCurrencyBR(plan.monthly_price || 0))}/mês</strong>
        </div>
        <p>${escapeHtml(plan.description || "")}</p>
        <div class="billing-plan-limits">
          <span>${escapeHtml(billingLimitLabel(plan.vehicle_limit))} veículos</span>
          <span>${escapeHtml(billingLimitLabel(plan.user_limit))} usuários</span>
        </div>
        <div class="billing-plan-features">
          ${features.map(([key]) => `<span>${escapeHtml(saasFeatureLabels[key] || key)}</span>`).join("")}
        </div>
        <button type="button" class="secondary" data-request-plan="${escapeHtml(code)}" ${disabled ? "disabled" : ""}>
          ${pending ? "Solicitação pendente" : current ? "Plano atual" : "Solicitar alteração"}
        </button>
      </article>
    `;
  }).join("");
  el("billingPlans").querySelectorAll("[data-request-plan]").forEach((button) => {
    button.addEventListener("click", () => requestBillingPlan(button.dataset.requestPlan));
  });
}

function renderBillingRequests() {
  const rows = el("billingRequestRows");
  const requests = state.billing.pending_requests || [];
  if (!requests.length) {
    rows.innerHTML = '<tr><td colspan="6">Nenhuma solicitação enviada.</td></tr>';
    return;
  }
  rows.innerHTML = requests.map((request) => `
    <tr>
      <td>${escapeHtml(formatDate(request.created_at))}</td>
      <td>${escapeHtml(billingStatusLabel(request.request_type))}</td>
      <td>${escapeHtml(request.current_plan_name || request.current_plan_code || "-")}</td>
      <td>${escapeHtml(request.requested_plan_name || request.requested_plan_code || "-")}</td>
      <td>${escapeHtml(billingStatusLabel(request.status))}</td>
      <td>${escapeHtml(request.message || "-")}</td>
    </tr>
  `).join("");
}

async function requestBillingPlan(planCode) {
  const plan = (state.billing.plans || []).find((item) => clean(item.code).toLowerCase() === clean(planCode).toLowerCase()) || {};
  const message = window.prompt(`Observação para solicitar ${plan.name || planCode}`, "");
  if (message === null) return;
  try {
    await apiRequest("/billing/plan-change-requests", {
      method: "POST",
      body: JSON.stringify({requested_plan_code: planCode, message: clean(message) || null}),
    });
    notify("Solicitação enviada ao administrador RotaHub.");
    await loadBilling();
  } catch (error) {
    notify(error.message, true);
  }
}

function applyPlanNavigation() {
  document.querySelectorAll(".nav-item[data-view]").forEach((button) => {
    const allowed = canAccessView(button.dataset.view);
    button.classList.toggle("plan-locked", !allowed);
    button.setAttribute("aria-disabled", allowed ? "false" : "true");
    button.title = allowed ? "" : "Disponível mediante upgrade de plano";
  });
  document.querySelectorAll("[data-home-nav]").forEach((button) => {
    const allowed = canAccessView(button.dataset.homeNav);
    button.classList.toggle("plan-locked", !allowed);
    button.disabled = !allowed;
    button.title = allowed ? "" : "Disponível mediante upgrade de plano";
  });
}

function syncNavGroups(view) {
  document.querySelectorAll(".nav-group").forEach((group) => {
    if (!(group instanceof HTMLDetailsElement)) return;
    const hasActiveView = Boolean(group.querySelector(`.nav-item[data-view="${view}"]`));
    if (hasActiveView) group.open = true;
  });
}

function renderCadastroTabs() {
  const tabs = el("cadastroTabs");
  tabs.innerHTML = "";
  Object.entries(cadastroResources).forEach(([key, resource]) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `tab-button ${state.cadastroResource === key ? "active" : ""}`;
    button.textContent = resource.label;
    button.addEventListener("click", () => {
      if (state.cadastroResource === key) return;
      state.cadastroResource = key;
      state.cadastroSelectedId = null;
      state.cadastroMode = "view";
      state.cadastroStatusFilter = "TODOS";
      state.cadastroItems = [];
      state.clientesSection = "hub";
      state.cadastroLoadSeq += 1;
      renderCadastroTabs();
      renderCadastroForm();
      renderCadastroTable();
      loadCadastros();
    });
    tabs.appendChild(button);
  });
  renderCadastroLayout();
}

function currentCadastroConfig() {
  const resource = cadastroResources[state.cadastroResource];
  if (state.cadastroResource !== "produtos") return resource;
  const config = currentLogisticaConfig();
  const unidade = config.unidade_padrao || "KG";
  const embalagem = unidade === "CX" ? "UN" : "CX";
  const unidades = logisticaUnitOptions();
  const categoriaPadrao = logisticaProductCategory();
  return {
    ...resource,
    defaults: {
      ...(resource.defaults || {}),
      categoria: categoriaPadrao,
      unidade,
      unidade_estoque: unidade,
    },
    fields: resource.fields.map((field) => {
      const [name, label, type, required] = field;
      if (name === "categoria") return [name, label, type, required, productCategoryOptions(categoriaPadrao)];
      if (name === "unidade" || name === "unidade_estoque") return [name, label, type, required, productUnitOptions(unidades, embalagem)];
      if (name === "estoque_min_kg") return [name, `MIN ${unidade}`, type, required];
      if (name === "estoque_min_caixas") return [name, `MIN ${config.embalagem_label || "CX"}`, type, required];
      return field;
    }),
  };
}

function renderCadastroLayout() {
  const clientesMode = state.cadastroResource === "clientes";
  el("cadastroActionbar").classList.toggle("hidden", clientesMode);
  el("cadastroFormPanel").classList.toggle("hidden", clientesMode);
  el("cadastroTablePanel").classList.toggle("hidden", clientesMode);
  el("caixasResumoPanel").classList.toggle("hidden", state.cadastroResource !== "caixas");
  if (state.cadastroResource === "caixas") renderCaixasResumoPanel();
  ["clientesHubPanel", "clientesHistoricoPanel", "clientesLocalizacaoPanel", "clientesImportPanel"].forEach((id) => {
    const node = el(id);
    if (node) node.classList.add("hidden");
  });
  if (!clientesMode) return;
  const section = state.clientesSection || "hub";
  const panelId = section === "lista"
    ? "clientesImportPanel"
    : section === "historico"
      ? "clientesHistoricoPanel"
      : section === "localizacao"
        ? "clientesLocalizacaoPanel"
        : "clientesHubPanel";
  el(panelId).classList.remove("hidden");
}

function currentCadastroItem() {
  return state.cadastroItems.find((candidate) => candidate.id === state.cadastroSelectedId) || null;
}

function cadastroFieldEnabled(name) {
  const resource = currentCadastroConfig();
  if (state.cadastroResource === "fornecedores" && ["tipo_pessoa", "certificado_status", "certificado_nome", "certificado_instalado_em"].includes(name)) return false;
  if (state.cadastroResource === "motoristas" && name === "codigo") return false;
  if (name === "senha" && state.cadastroMode === "edit") return false;
  if (state.cadastroMode === "novo") return true;
  if (state.cadastroMode === "edit") return true;
  if (state.cadastroMode === "status") return name === resource.statusField;
  return false;
}

function cadastroFieldValue(name) {
  const resource = currentCadastroConfig();
  const item = currentCadastroItem();
  if (state.cadastroMode === "novo") {
    if (state.cadastroResource === "motoristas" && name === "codigo") {
      return state.cadastroNextMotoristaCodigo || "GERADO AO SALVAR";
    }
    return resource.defaults && Object.prototype.hasOwnProperty.call(resource.defaults, name) ? resource.defaults[name] : "";
  }
  if (!item) return resource.defaults && Object.prototype.hasOwnProperty.call(resource.defaults, name) ? resource.defaults[name] : "";
  if (name === "senha") return "";
  return item.data[name] ?? "";
}

function renderCadastroForm() {
  renderCadastroLayout();
  if (state.cadastroResource === "clientes") {
    renderClientesHub();
    return;
  }
  const resource = currentCadastroConfig();
  const item = currentCadastroItem();
  const modeTitle = state.cadastroMode === "novo"
    ? "Novo"
    : state.cadastroMode === "edit"
      ? "Alterar"
    : state.cadastroMode === "status"
      ? "Alterar status"
      : item
        ? "Selecionado"
        : "Consulta";

  el("cadastroFormTitle").textContent = `${modeTitle} ${resource.label}`;
  el("cadastroTableTitle").textContent = cadastroTableTitle(resource);
  el("cancelCadastroEdit").classList.toggle("hidden", !item && state.cadastroMode === "view");
  el("cadastroFormSubtitle").textContent = cadastroFormSubtitle(resource);

  const form = el("cadastroForm");
  form.innerHTML = "";
  resource.fields.forEach(([name, label, type, required, options]) => {
    if (state.cadastroResource === "fornecedores" && name === "perfil_fornecedor") {
      options = state.fornecedorPerfis || options || [];
    }
    const wrapper = document.createElement("label");
    wrapper.textContent = label;
    const input = type === "select" ? document.createElement("select") : document.createElement("input");
    input.name = name;
    if (type !== "select") input.type = type;
    input.required = Boolean(required && cadastroFieldEnabled(name));
    input.disabled = !cadastroFieldEnabled(name);
    input.value = cadastroFieldValue(name);
    input.autocomplete = name === "senha" ? "new-password" : "off";
    if (type === "number") {
      input.min = "0";
      input.step = ["estoque_min_kg", "custo_padrao", "preco_padrao"].includes(name) ? "0.01" : "1";
    }
    if (type === "select") {
      (options || []).forEach((optionValue) => {
        const option = document.createElement("option");
        option.value = optionValue;
        option.textContent = optionValue;
        input.appendChild(option);
      });
      input.value = cadastroFieldValue(name) || (options && options[0]) || "";
    }
    wrapper.appendChild(input);
    form.appendChild(wrapper);
  });

  renderCadastroActions();
}

function cadastroTableTitle(resource) {
  if (state.cadastroResource !== "caixas") return resource.label;
  const rows = Array.isArray(state.cadastroItems) ? state.cadastroItems : [];
  const count = (status) => rows.filter((item) => clean(item?.data?.status).toUpperCase() === status).length;
  return `Caixas - Total ${rows.length} | Estoque ${count("EM_ESTOQUE")} | Vinculadas ${count("VINCULADA") + count("EM_USO")} | Quebradas ${count("QUEBRADA")}`;
}

function caixaResumoRows() {
  const rows = Array.isArray(state.cadastroItems) ? state.cadastroItems : [];
  const statusTotals = {};
  const byVehicle = {};
  const byLote = {};
  rows.forEach((item) => {
    const data = item.data || {};
    const status = clean(data.status).toUpperCase() || "SEM_STATUS";
    const veiculo = clean(data.veiculo_placa).toUpperCase() || "SEM_VEICULO";
    const lote = clean(data.lote).toUpperCase() || "SEM_LOTE";
    statusTotals[status] = (statusTotals[status] || 0) + 1;
    byVehicle[veiculo] = (byVehicle[veiculo] || 0) + 1;
    byLote[lote] = (byLote[lote] || 0) + 1;
  });
  return {rows, statusTotals, byVehicle, byLote};
}

function resumoTopList(map, limit = 6) {
  const items = Object.entries(map || {}).sort((a, b) => b[1] - a[1]).slice(0, limit);
  if (!items.length) return '<span class="caixas-resumo-empty">Sem dados</span>';
  return items.map(([label, value]) => `<span><strong>${escapeHtml(label)}</strong>${escapeHtml(value)}</span>`).join("");
}

function renderCaixasResumoPanel() {
  const panel = el("caixasResumoPanel");
  if (!panel) return;
  const {rows, statusTotals, byVehicle, byLote} = caixaResumoRows();
  const capacityMap = {};
  (state.caixasVeiculos || []).forEach((item) => {
    const data = item.data || {};
    const placa = clean(data.placa).toUpperCase();
    if (!placa) return;
    capacityMap[placa] = Math.max(parseInt(clean(data.capacidade_cx) || "0", 10) || 0, 0);
  });
  const vehicleBalance = Object.keys({...capacityMap, ...byVehicle})
    .filter((placa) => placa !== "SEM_VEICULO")
    .map((placa) => {
      const vinculadas = byVehicle[placa] || 0;
      const capacidade = capacityMap[placa] || 0;
      return {placa, vinculadas, capacidade, saldo: capacidade - vinculadas};
    })
    .sort((a, b) => a.saldo - b.saldo || a.placa.localeCompare(b.placa));
  const disponiveis = statusTotals.EM_ESTOQUE || 0;
  const vinculadas = (statusTotals.VINCULADA || 0) + (statusTotals.EM_USO || 0);
  const quebradas = statusTotals.QUEBRADA || 0;
  const baixadas = statusTotals.BAIXADA || 0;
  const veiculosExcesso = vehicleBalance.filter((item) => item.capacidade > 0 && item.saldo < 0).length;
  const veiculosFalta = vehicleBalance.filter((item) => item.capacidade > 0 && item.saldo > 0).length;
  const vehicleCapacityList = vehicleBalance.length
    ? vehicleBalance.slice(0, 8).map((item) => {
      const cls = item.saldo < 0 ? "danger" : item.saldo > 0 ? "warning" : "ok";
      const capacidade = item.capacidade || "-";
      const saldo = item.capacidade ? ` | saldo ${item.saldo}` : " | sem capacidade";
      return `<span class="${cls}"><strong>${escapeHtml(item.placa)}</strong>${escapeHtml(item.vinculadas)}/${escapeHtml(capacidade)}${escapeHtml(saldo)}</span>`;
    }).join("")
    : '<span class="caixas-resumo-empty">Sem veiculos</span>';
  panel.innerHTML = `
    <div class="caixas-resumo-kpis">
      <span><small>Total</small><strong>${escapeHtml(rows.length)}</strong></span>
      <span><small>Estoque</small><strong>${escapeHtml(disponiveis)}</strong></span>
      <span><small>Vinculadas/uso</small><strong>${escapeHtml(vinculadas)}</strong></span>
      <span><small>Quebradas</small><strong>${escapeHtml(quebradas)}</strong></span>
      <span><small>Baixadas</small><strong>${escapeHtml(baixadas)}</strong></span>
      <span><small>Veic. excesso</small><strong>${escapeHtml(veiculosExcesso)}</strong></span>
      <span><small>Veic. falta</small><strong>${escapeHtml(veiculosFalta)}</strong></span>
    </div>
    <div class="caixas-resumo-lists">
      <div><small>Por veiculo</small>${resumoTopList(byVehicle)}</div>
      <div><small>Por lote</small>${resumoTopList(byLote)}</div>
      <div class="wide"><small>Capacidade x caixas</small>${vehicleCapacityList}</div>
    </div>
  `;
}

function cadastroFormSubtitle(resource) {
  if (state.cadastroMode === "novo" && state.cadastroResource === "motoristas") {
    return "Informe os dados do motorista. O codigo e gerado automaticamente na sequencia MOT-01, MOT-02...";
  }
  if (state.cadastroResource === "caixas") {
    if (state.cadastroMode === "novo") return "Cadastre uma caixa individual por codigo para rastrear lote, cor, veiculo e quebras.";
    if (state.cadastroMode === "edit") return "Altere placa/status para registrar rotatividade, quebra, baixa ou retorno ao estoque.";
  }
  if (state.cadastroMode === "novo") return "Preencha e clique em SALVAR.";
  if (state.cadastroMode === "edit") return "Altere os campos necessarios e clique em SALVAR.";
  if (state.cadastroMode === "status") return `Somente o campo ${resource.statusField.toUpperCase()} fica liberado, como no desktop.`;
  if (currentCadastroItem()) return "Registro selecionado. Use ALTERAR, LIBERAR, BLOQUEAR ou ALTERAR SENHA quando aplicavel.";
  return "Selecione um registro na tabela ou clique em NOVO.";
}

function renderCadastroActions() {
  el("perfilFornecedorButton").classList.toggle("hidden", state.cadastroResource !== "fornecedores");
  el("caixasBulkButton").classList.toggle("hidden", state.cadastroResource !== "caixas");
  el("caixasMoveButton").classList.toggle("hidden", state.cadastroResource !== "caixas");
  el("caixasHistoricoButton").classList.toggle("hidden", state.cadastroResource !== "caixas");
  el("caixasExportButton").classList.toggle("hidden", state.cadastroResource !== "caixas");
  el("caixasLabelsButton").classList.toggle("hidden", state.cadastroResource !== "caixas");
  if (state.cadastroResource === "clientes") return;
  const resource = currentCadastroConfig();
  const item = currentCadastroItem();
  const hasStatus = Boolean(resource.statusField);
  const hasPassword = Boolean(resource.password);
  el("changeCadastroButton").disabled = !item;
  el("saveCadastroButton").disabled = state.cadastroMode === "view";
  el("liberarCadastroButton").disabled = !item || !hasStatus;
  el("bloquearCadastroButton").disabled = !item || !hasStatus;
  el("passwordCadastroButton").disabled = !item || !hasPassword;
  el("perfilFornecedorButton").disabled = state.cadastroResource !== "fornecedores";
  el("deleteCadastroButton").disabled = !item;
  el("caixasBulkButton").disabled = state.cadastroResource !== "caixas";
  el("caixasMoveButton").disabled = state.cadastroResource !== "caixas";
  el("caixasHistoricoButton").disabled = state.cadastroResource !== "caixas" || !item;
  el("caixasExportButton").disabled = state.cadastroResource !== "caixas" || !state.cadastroItems.length;
  el("caixasLabelsButton").disabled = state.cadastroResource !== "caixas" || !state.cadastroItems.length;
  el("cadastroStatusFilterWrap").classList.toggle("hidden", !["ajudantes", "fornecedores", "produtos", "caixas"].includes(state.cadastroResource));
  const filterOptions = state.cadastroResource === "caixas"
    ? ["TODOS", "EM_ESTOQUE", "VINCULADA", "EM_USO", "QUEBRADA", "BAIXADA"]
    : state.cadastroResource === "ajudantes"
      ? ["TODOS", "ATIVO", "DESATIVADO"]
      : ["TODOS", "ATIVO", "INATIVO"];
  el("cadastroStatusFilter").innerHTML = filterOptions
    .map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
    .join("");
  if (!filterOptions.includes(state.cadastroStatusFilter)) {
    state.cadastroStatusFilter = "TODOS";
  }
  el("cadastroStatusFilter").value = state.cadastroStatusFilter;
}

async function startNewCadastro() {
  state.cadastroSelectedId = null;
  state.cadastroMode = "novo";
  if (state.cadastroResource === "motoristas") {
    state.cadastroNextMotoristaCodigo = "";
    try {
      const next = await apiRequest("/cadastros/motoristas/proximo-codigo");
      state.cadastroNextMotoristaCodigo = next.codigo || "";
    } catch (error) {
      state.cadastroNextMotoristaCodigo = "";
    }
  }
  renderCadastroForm();
  renderCadastroTable();
  const firstInput = el("cadastroForm").querySelector("input:not(:disabled), select:not(:disabled)");
  if (firstInput) firstInput.focus();
}

function startCadastroEdit() {
  if (!currentCadastroItem()) {
    notify("Selecione um cadastro na tabela.", true);
    return;
  }
  state.cadastroMode = "edit";
  renderCadastroForm();
  const firstInput = el("cadastroForm").querySelector("input:not(:disabled), select:not(:disabled)");
  if (firstInput) firstInput.focus();
}

function startCadastroStatusEdit() {
  const resource = currentCadastroConfig();
  if (!currentCadastroItem()) {
    notify("Selecione um cadastro na tabela.", true);
    return;
  }
  if (!resource.statusField) {
    notify("Este cadastro nao possui campo STATUS.", true);
    return;
  }
  state.cadastroMode = "status";
  renderCadastroForm();
  const statusInput = el("cadastroForm").elements[resource.statusField];
  if (statusInput) statusInput.focus();
}

function clearCadastroSelection() {
  state.cadastroSelectedId = null;
  state.cadastroMode = "view";
  renderCadastroForm();
  renderCadastroTable();
}

function selectCadastroItem(item) {
  state.cadastroSelectedId = item.id;
  state.cadastroMode = "view";
  renderCadastroForm();
  renderCadastroTable();
}

async function loadCadastros() {
  const resourceKey = state.cadastroResource;
  const loadSeq = ++state.cadastroLoadSeq;
  if (state.cadastroResource === "clientes") {
    await loadClientesDashboard();
    await loadClientesLookup();
    if (state.clientesSection === "lista") await loadClientesImportRows();
    return;
  }
  const resource = currentCadastroConfig();
  try {
    if (resourceKey === "fornecedores") {
      await loadFornecedorPerfis();
    }
    if (resourceKey === "caixas") {
      await loadCaixasVeiculosCapacidade();
    }
    const params = new URLSearchParams();
    if (["ajudantes", "fornecedores", "produtos", "caixas"].includes(resourceKey) && state.cadastroStatusFilter !== "TODOS") {
      params.set("status_filter", state.cadastroStatusFilter);
    }
    if (resourceKey === "caixas") {
      params.set("limit", "10000");
    }
    const query = params.toString() ? `?${params.toString()}` : "";
    const items = await apiRequest(`/cadastros/${resource.endpoint}${query}`);
    if (loadSeq !== state.cadastroLoadSeq || resourceKey !== state.cadastroResource) return;
    state.cadastroItems = Array.isArray(items) ? items : [];
    if (state.cadastroSelectedId && !currentCadastroItem()) {
      state.cadastroSelectedId = null;
      state.cadastroMode = "view";
    }
    renderCadastroTable();
    if (resourceKey === "caixas") renderCaixasResumoPanel();
    renderCadastroForm();
  } catch (error) {
    if (loadSeq !== state.cadastroLoadSeq || resourceKey !== state.cadastroResource) return;
    state.cadastroItems = [];
    renderCadastroTable();
    if (resourceKey === "caixas") renderCaixasResumoPanel();
    notify(error.message, true);
  }
}

async function loadCaixasVeiculosCapacidade() {
  try {
    const rows = await apiRequest("/cadastros/veiculos?limit=10000");
    state.caixasVeiculos = Array.isArray(rows) ? rows : [];
  } catch (_error) {
    state.caixasVeiculos = [];
  }
}

async function loadFornecedorPerfis() {
  const rows = await apiRequest("/cadastros/meta/fornecedor-perfis");
  const perfis = (Array.isArray(rows) ? rows : [])
    .filter((item) => clean(item.status || "ATIVO") === "ATIVO")
    .map((item) => clean(item.codigo))
    .filter(Boolean);
  if (perfis.length) {
    state.fornecedorPerfis = [...new Set(perfis)];
  }
}

async function loadLogisticaConfig(options = {}) {
  try {
    const data = await apiRequest("/logistica/config");
    state.logisticaConfig = data || null;
    renderLogisticaConfig();
    applyLogisticaLabels();
    if (state.currentView === "cadastros" && state.cadastroResource === "produtos") {
      renderCadastroForm();
    }
  } catch (error) {
    if (!options.silent) notify(error.message || "Falha ao carregar configuracao logistica.", true);
  }
}

function currentLogisticaConfig() {
  const fallback = {
    perfil_codigo: "DISTRIBUICAO_GERAL",
    produto_padrao: "CARGA",
    unidade_padrao: "KG",
    embalagem_label: "Volumes",
    perda_label: "Ocorrencias operacionais",
    quantidade_embalagem_label: "Qtd. por volume",
    usa_mortalidade: 1,
    usa_aves_por_caixa: 0,
    usa_nota_fiscal_motorista: 1,
    usa_estoque_fisico: 1,
    usa_estoque_fiscal: 1,
  };
  const config = (state.logisticaConfig && state.logisticaConfig.config) || fallback;
  return {...fallback, ...config, perda_label: normalizeLegacyPerdaLabel(config.perda_label)};
}

function normalizeLegacyPerdaLabel(value) {
  const label = clean(value);
  const key = label.toUpperCase();
  if (!label || key.includes("MORTAL") || ["PERDA", "PERDAS", "AVES MORTAS", "MORTES"].includes(key)) {
    return "Ocorrencias operacionais";
  }
  return label;
}

function logisticaUnitOptions() {
  const unidades = state.logisticaConfig && Array.isArray(state.logisticaConfig.unidades) ? state.logisticaConfig.unidades : [];
  const configured = unidades.map((item) => clean(item.codigo)).filter(Boolean);
  return [...new Set([currentLogisticaConfig().unidade_padrao || "KG", ...configured, "KG", "CX", "UN", "PC", "LT"])];
}

function productUnitOptions(unidades, embalagem) {
  return [...new Set([...(unidades || []), embalagem, "CX", "UN"])].filter(Boolean);
}

function logisticaProductCategory() {
  const config = currentLogisticaConfig();
  const produto = clean(config.produto_padrao).toUpperCase();
  if (produto.includes("OVO")) return "OVOS";
  if (produto.includes("POLPA")) return "CONGELADOS";
  if (produto.includes("PECA") || produto.includes("PECAS")) return "PECAS";
  if (produto.includes("SERVICO") || produto.includes("SERVICOS")) return "SERVICOS";
  return "GERAL";
}

function productCategoryOptions(categoriaPadrao) {
  return [...new Set([categoriaPadrao, "GERAL", "CARGA", "CONGELADOS", "OVOS", "PECAS", "INSUMOS", "EMBALAGENS", "SERVICOS", "OUTROS"])].filter(Boolean);
}

function setLabelTextFor(inputId, text) {
  const input = el(inputId);
  const label = input && input.closest("label");
  if (!label) return;
  const textNode = [...label.childNodes].find((node) => node.nodeType === Node.TEXT_NODE && clean(node.textContent));
  if (textNode) textNode.textContent = `\n              ${text}\n              `;
}

function setTextContent(selector, text) {
  document.querySelectorAll(selector).forEach((node) => {
    node.textContent = text;
  });
}

function setSelectOptionText(selectId, value, text) {
  const select = el(selectId);
  if (!select) return;
  const option = [...select.options].find((candidate) => candidate.value === value);
  if (option) option.textContent = text;
}

function applyLogisticaLabels() {
  const config = currentLogisticaConfig();
  const unidade = config.unidade_padrao || "KG";
  const embalagem = config.embalagem_label || "Volumes";
  const perda = normalizeLegacyPerdaLabel(config.perda_label);
  const produto = config.produto_padrao || "CARGA";
  const perdaUpper = perda.toUpperCase();
  const navMortalidade = document.querySelector(".nav-item[data-view='mortalidade']");
  if (navMortalidade) navMortalidade.textContent = "Ocorrencias";
  if (state.currentView === "mortalidade") {
    el("viewTitle").textContent = perda;
    el("viewSubtitle").textContent = "Registros por rota";
  }
  const manualHeading = document.querySelector("#mortalidadeManualForm .manual-form-heading h3");
  if (manualHeading) manualHeading.textContent = `${perda} manual`;
  setLabelTextFor("manualMortAves", `${perda} - unidades`);
  setLabelTextFor("manualMortMotivo", `Motivo da ocorrencia`);
  const saveManual = el("mortalidadeManualSaveButton");
  if (saveManual) saveManual.textContent = `REGISTRAR ${perdaUpper}`;
  setLabelTextFor("comprasEntradaKg", `${unidade} entrada`);
  setLabelTextFor("comprasEntradaCaixas", `${embalagem} entrada`);
  setLabelTextFor("comprasAjusteKg", `${unidade} ajuste`);
  setLabelTextFor("comprasAjusteCaixas", embalagem);
  setSelectOptionText("progTipoEstimativa", "KG", unidade);
  setSelectOptionText("progTipoEstimativa", "CX", embalagem);
  if (el("progEstimativa")) updateProgramacaoEstimateLabel();
  setTextContent("[data-logistica-label='unidade']", unidade);
  setTextContent("[data-logistica-label='unidade-upper']", unidade.toUpperCase());
  setTextContent("[data-logistica-label='embalagem']", embalagem);
  setTextContent("[data-logistica-label='embalagem-upper']", embalagem.toUpperCase());
  setTextContent("[data-logistica-label='perda']", perda);
  setTextContent("[data-logistica-label='perda-upper']", perdaUpper);
  setTextContent("[data-logistica-label='total-embalagem']", `Total de ${embalagem.toLowerCase()}`);
  setTextContent("[data-logistica-label='selecionadas-embalagem']", `${embalagem} selecionadas`);
  setTextContent("[data-logistica-label='fiscal-saldo-unidade']", `Fiscal saldo ${unidade}`);
  setTextContent("[data-logistica-label='fisico-saldo-unidade']", `Fisico saldo ${unidade}`);
  setTextContent("[data-logistica-label='saldo-unidade']", `Saldo ${unidade}`);
  const produtoOption = document.querySelector("#comprasAjusteProduto option[value='']");
  if (produtoOption) produtoOption.textContent = produto;
  ["comprasFiscalSaldo", "comprasFisicoSaldo"].forEach((id) => {
    const node = el(id);
    if (node) node.textContent = node.textContent.replace(/\s[A-Z0-9]+$/, ` ${unidade}`);
  });
}

function renderLogisticaConfig() {
  const data = state.logisticaConfig || {};
  const config = data.config || {};
  const perfis = data.perfis || [];
  const unidades = data.unidades || [];
  el("logisticaPerfilCodigo").innerHTML = perfis.map((perfil) => (
    `<option value="${escapeHtml(perfil.codigo || "")}">${escapeHtml(perfil.nome || perfil.codigo || "")}</option>`
  )).join("");
  el("logisticaUnidadePadrao").innerHTML = unidades.map((unidade) => (
    `<option value="${escapeHtml(unidade.codigo || "")}">${escapeHtml(unidade.codigo || "")} - ${escapeHtml(unidade.nome || "")}</option>`
  )).join("");
  el("logisticaPerfilCodigo").value = config.perfil_codigo || "DISTRIBUICAO_GERAL";
  el("logisticaProdutoPadrao").value = config.produto_padrao || "CARGA";
  el("logisticaUnidadePadrao").value = config.unidade_padrao || "KG";
  el("logisticaEmbalagemLabel").value = config.embalagem_label || "Volumes";
  el("logisticaPerdaLabel").value = config.perda_label || "Ocorrencias operacionais";
  el("logisticaQtdEmbalagemLabel").value = config.quantidade_embalagem_label || "Qtd. por volume";
  el("logisticaUsaMortalidade").value = String(config.usa_mortalidade ?? 1);
  el("logisticaUsaAvesPorCaixa").value = String(config.usa_aves_por_caixa ?? 0);
  el("logisticaUsaNfMotorista").value = String(config.usa_nota_fiscal_motorista ?? 1);
  el("logisticaUsaEstoqueFisico").value = String(config.usa_estoque_fisico ?? 1);
  el("logisticaUsaEstoqueFiscal").value = String(config.usa_estoque_fiscal ?? 1);
  el("logisticaConfigInfo").textContent = `Empresa ${data.company_id || "-"} | perfil ${config.perfil_codigo || "DISTRIBUICAO_GERAL"} | unidade ${config.unidade_padrao || "KG"}.`;
}

function applySelectedLogisticaPerfil() {
  const data = state.logisticaConfig || {};
  const codigo = clean(el("logisticaPerfilCodigo").value);
  const perfil = (data.perfis || []).find((item) => clean(item.codigo) === codigo);
  if (!perfil) return;
  el("logisticaProdutoPadrao").value = perfil.produto_padrao || el("logisticaProdutoPadrao").value;
  el("logisticaUnidadePadrao").value = perfil.unidade_padrao || el("logisticaUnidadePadrao").value;
  el("logisticaEmbalagemLabel").value = perfil.embalagem_label || el("logisticaEmbalagemLabel").value;
  el("logisticaPerdaLabel").value = perfil.perda_label || el("logisticaPerdaLabel").value;
  el("logisticaQtdEmbalagemLabel").value = perfil.quantidade_embalagem_label || el("logisticaQtdEmbalagemLabel").value;
  el("logisticaUsaMortalidade").value = String(perfil.usa_mortalidade ?? el("logisticaUsaMortalidade").value);
  el("logisticaUsaAvesPorCaixa").value = String(perfil.usa_aves_por_caixa ?? el("logisticaUsaAvesPorCaixa").value);
  el("logisticaUsaNfMotorista").value = String(perfil.usa_nota_fiscal_motorista ?? el("logisticaUsaNfMotorista").value);
  el("logisticaUsaEstoqueFisico").value = String(perfil.usa_estoque_fisico ?? el("logisticaUsaEstoqueFisico").value);
  el("logisticaUsaEstoqueFiscal").value = String(perfil.usa_estoque_fiscal ?? el("logisticaUsaEstoqueFiscal").value);
}

async function saveLogisticaConfig() {
  const data = state.logisticaConfig || {};
  const payload = {
    company_id: data.company_id || null,
    perfil_codigo: clean(el("logisticaPerfilCodigo").value) || "DISTRIBUICAO_GERAL",
    produto_padrao: clean(el("logisticaProdutoPadrao").value) || "PRODUTO",
    unidade_padrao: clean(el("logisticaUnidadePadrao").value) || "UN",
    embalagem_label: clean(el("logisticaEmbalagemLabel").value) || "Embalagens",
    perda_label: clean(el("logisticaPerdaLabel").value) || "Ocorrencias operacionais",
    quantidade_embalagem_label: clean(el("logisticaQtdEmbalagemLabel").value) || "Qtd. por volume",
    usa_mortalidade: Number(el("logisticaUsaMortalidade").value),
    usa_aves_por_caixa: Number(el("logisticaUsaAvesPorCaixa").value),
    usa_nota_fiscal_motorista: Number(el("logisticaUsaNfMotorista").value),
    usa_estoque_fisico: Number(el("logisticaUsaEstoqueFisico").value),
    usa_estoque_fiscal: Number(el("logisticaUsaEstoqueFiscal").value),
  };
  try {
    await apiRequest("/logistica/config", {
      method: "PUT",
      body: JSON.stringify(payload),
    });
    notify("Perfil logistico salvo.");
    await loadLogisticaConfig();
  } catch (error) {
    notify(error.message || "Falha ao salvar perfil logistico.", true);
  }
}

function normalizeCodeInput(value) {
  return clean(value).toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

async function createLogisticaPerfil() {
  const base = currentLogisticaConfig();
  const codigo = normalizeCodeInput(window.prompt("Codigo do novo perfil logistico (ex.: DISTRIBUICAO_GERAL, FRIOS, OVOS):") || "");
  if (!codigo) return;
  const nome = clean(window.prompt("Nome do perfil logistico:") || "");
  if (!nome) return;
  const produto = clean(window.prompt("Produto padrao:", base.produto_padrao || "PRODUTO") || "");
  const unidade = clean(window.prompt("Unidade padrao:", base.unidade_padrao || "UN") || "");
  const embalagem = clean(window.prompt("Nome da embalagem:", base.embalagem_label || "Embalagens") || "");
  const perda = clean(window.prompt("Nome da ocorrencia operacional:", base.perda_label || "Ocorrencias operacionais") || "");
  try {
    await apiRequest("/logistica/perfis", {
      method: "POST",
      body: JSON.stringify({
        codigo,
        nome,
        descricao: clean(window.prompt("Descricao do perfil:", "") || ""),
        produto_padrao: produto || "PRODUTO",
        unidade_padrao: unidade || "UN",
        embalagem_label: embalagem || "Embalagens",
        perda_label: perda || "Ocorrencias operacionais",
        quantidade_embalagem_label: base.quantidade_embalagem_label || "Qtd. por volume",
        usa_mortalidade: Number(base.usa_mortalidade || 0),
        usa_aves_por_caixa: Number(base.usa_aves_por_caixa || 0),
        usa_nota_fiscal_motorista: Number(base.usa_nota_fiscal_motorista ?? 1),
        usa_estoque_fisico: Number(base.usa_estoque_fisico ?? 1),
        usa_estoque_fiscal: Number(base.usa_estoque_fiscal ?? 1),
      }),
    });
    await loadLogisticaConfig();
    el("logisticaPerfilCodigo").value = codigo;
    notify("Perfil logistico criado.");
  } catch (error) {
    notify(error.message || "Falha ao criar perfil logistico.", true);
  }
}

async function createLogisticaUnidade() {
  const codigo = normalizeCodeInput(window.prompt("Codigo da unidade (ex.: BD, DZ, PALLET):") || "");
  if (!codigo) return;
  const nome = clean(window.prompt("Nome da unidade:") || "");
  if (!nome) return;
  try {
    await apiRequest("/logistica/unidades", {
      method: "POST",
      body: JSON.stringify({
        codigo,
        nome,
        tipo: clean(window.prompt("Tipo da unidade:", "UNIDADE") || "UNIDADE"),
      }),
    });
    await loadLogisticaConfig();
    el("logisticaUnidadePadrao").value = codigo;
    notify("Unidade logistica criada.");
  } catch (error) {
    notify(error.message || "Falha ao criar unidade logistica.", true);
  }
}

async function createLogisticaOcorrencia() {
  const codigo = normalizeCodeInput(window.prompt("Codigo da ocorrencia (ex.: AVARIA, FALTA, QUEBRA):") || "");
  if (!codigo) return;
  const nome = clean(window.prompt("Nome da ocorrencia:") || "");
  if (!nome) return;
  try {
    await apiRequest("/logistica/ocorrencias", {
      method: "POST",
      body: JSON.stringify({
        codigo,
        nome,
        categoria: clean(window.prompt("Categoria:", "PERDA") || "PERDA"),
        perfil_codigo: clean(el("logisticaPerfilCodigo").value) || currentLogisticaConfig().perfil_codigo || "DISTRIBUICAO_GERAL",
      }),
    });
    await loadLogisticaConfig();
    notify("Ocorrencia logistica criada.");
  } catch (error) {
    notify(error.message || "Falha ao criar ocorrencia logistica.", true);
  }
}

async function createFornecedorPerfil() {
  const codigo = clean(window.prompt("Codigo do perfil do fornecedor (ex.: LAVADOR_CAIXAS):") || "").toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  if (!codigo) return;
  const nome = clean(window.prompt("Nome do perfil do fornecedor:") || "");
  if (!nome) return;
  try {
    await apiRequest("/cadastros/meta/fornecedor-perfis", {
      method: "POST",
      body: JSON.stringify({codigo, nome, categoria: "OUTROS", status: "ATIVO"}),
    });
    await loadFornecedorPerfis();
    renderCadastroForm();
    notify("Perfil de fornecedor criado.");
  } catch (error) {
    notify(error.message || "Falha ao criar perfil de fornecedor.", true);
  }
}

function openCaixasBulkDialog() {
  if (state.cadastroResource !== "caixas") return;
  const form = el("caixasBulkForm");
  form.reset();
  form.elements.quantidade.value = "100";
  form.elements.prefixo.value = "CX";
  form.elements.numero_inicial.value = "1";
  form.elements.digitos.value = "4";
  form.elements.status.value = "EM_ESTOQUE";
  form.elements.observacao.value = "CRIACAO EM LOTE";
  el("caixasBulkDialog").showModal();
  form.elements.lote.focus();
}

function closeCaixasBulkDialog() {
  const dialog = el("caixasBulkDialog");
  if (dialog.open) dialog.close();
}

async function createCaixasBulk(event) {
  event.preventDefault();
  if (state.cadastroResource !== "caixas") return;
  const form = event.currentTarget;
  const data = Object.fromEntries(new FormData(form).entries());
  const veiculo = clean(data.veiculo_placa);
  try {
    const result = await apiRequest("/cadastros/caixas/bulk", {
      method: "POST",
      body: JSON.stringify({
        prefixo: clean(data.prefixo) || "CX",
        lote: clean(data.lote),
        cor: clean(data.cor),
        quantidade: Math.max(parseInt(clean(data.quantidade) || "0", 10) || 0, 0),
        numero_inicial: Math.max(parseInt(clean(data.numero_inicial) || "1", 10) || 1, 1),
        digitos: Math.max(parseInt(clean(data.digitos) || "4", 10) || 4, 1),
        veiculo_placa: veiculo || null,
        status: clean(data.status) || (veiculo ? "VINCULADA" : "EM_ESTOQUE"),
        data_compra: clean(data.data_compra) || null,
        observacao: clean(data.observacao) || "CRIACAO EM LOTE",
      }),
    });
    closeCaixasBulkDialog();
    notify(`Lote criado: ${result.criadas || 0} caixas (${result.primeiro_codigo} a ${result.ultimo_codigo}).`);
    await loadCadastros();
  } catch (error) {
    notify(error.message || "Falha ao gerar lote de caixas.", true);
  }
}

function openCaixasMoveDialog() {
  if (state.cadastroResource !== "caixas") return;
  const form = el("caixasMoveForm");
  form.reset();
  form.elements.quantidade.value = "1";
  form.elements.status_origem.value = state.cadastroStatusFilter !== "TODOS" ? state.cadastroStatusFilter : "TODOS";
  form.elements.status_destino.value = "VINCULADA";
  const item = currentCadastroItem();
  if (item && item.data) {
    form.elements.lote.value = item.data.lote || "";
    form.elements.cor.value = item.data.cor || "";
    form.elements.veiculo_origem.value = item.data.veiculo_placa || "";
  }
  el("caixasMoveDialog").showModal();
  form.elements.quantidade.focus();
}

function closeCaixasMoveDialog() {
  const dialog = el("caixasMoveDialog");
  if (dialog.open) dialog.close();
}

async function movimentarCaixas(event) {
  event.preventDefault();
  if (state.cadastroResource !== "caixas") return;
  const form = event.currentTarget;
  const data = Object.fromEntries(new FormData(form).entries());
  try {
    const result = await apiRequest("/cadastros/caixas/movimentar", {
      method: "POST",
      body: JSON.stringify({
        quantidade: Math.max(parseInt(clean(data.quantidade) || "0", 10) || 0, 0),
        lote: clean(data.lote) || null,
        cor: clean(data.cor) || null,
        veiculo_origem: clean(data.veiculo_origem) || null,
        status_origem: clean(data.status_origem) || "TODOS",
        veiculo_destino: clean(data.veiculo_destino) || null,
        status_destino: clean(data.status_destino) || "VINCULADA",
        observacao: clean(data.observacao) || null,
      }),
    });
    closeCaixasMoveDialog();
    notify(`${result.movimentadas || 0} caixa(s) movimentada(s).`);
    state.cadastroSelectedId = null;
    state.cadastroMode = "view";
    await loadCadastros();
  } catch (error) {
    notify(error.message || "Falha ao movimentar caixas.", true);
  }
}

async function openCaixaHistorico() {
  const item = currentCadastroItem();
  if (state.cadastroResource !== "caixas" || !item) {
    notify("Selecione uma caixa na tabela.", true);
    return;
  }
  try {
    const data = await apiRequest(`/cadastros/caixas/${item.id}/historico`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const caixa = data.caixa && data.caixa.data ? data.caixa.data : item.data || {};
    el("caixasHistoricoTitle").textContent = `Historico da caixa ${caixa.codigo || item.id}`;
    el("caixasHistoricoSubtitle").textContent = `${caixa.lote || "-"} | ${caixa.cor || "-"} | ${caixa.veiculo_placa || "ESTOQUE"} | ${caixa.status || "-"}`;
    el("caixasHistoricoRows").innerHTML = rows.length
      ? rows.map((row) => `
        <tr>
          <td>${escapeHtml(row.criado_em || "-")}</td>
          <td>${escapeHtml(row.movimento || "-")}</td>
          <td>${escapeHtml(row.veiculo_origem || "-")}</td>
          <td>${escapeHtml(row.veiculo_destino || "-")}</td>
          <td>${escapeHtml(row.status_origem || "-")}</td>
          <td>${escapeHtml(row.status_destino || "-")}</td>
          <td>${escapeHtml(row.observacao || "-")}</td>
        </tr>
      `).join("")
      : `<tr><td colspan="7">Sem historico registrado.</td></tr>`;
    el("caixasHistoricoDialog").showModal();
  } catch (error) {
    notify(error.message || "Falha ao carregar historico da caixa.", true);
  }
}

function closeCaixaHistoricoDialog() {
  const dialog = el("caixasHistoricoDialog");
  if (dialog.open) dialog.close();
}

function caixasRowsForOutput({selectedOnly = false} = {}) {
  if (state.cadastroResource !== "caixas") return [];
  const selected = currentCadastroItem();
  if (selectedOnly && selected) return [selected];
  return Array.isArray(state.cadastroItems) ? state.cadastroItems : [];
}

function csvCell(value) {
  const text = clean(value).replaceAll('"', '""');
  return `"${text}"`;
}

function downloadTextFile(filename, content, type = "text/plain;charset=utf-8") {
  const blob = new Blob([content], {type});
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportCaixasCsv() {
  const rows = caixasRowsForOutput();
  if (!rows.length) {
    notify("Sem caixas para exportar.", true);
    return;
  }
  const headers = ["ID", "CODIGO", "LOTE", "COR", "VEICULO", "STATUS", "DATA_COMPRA", "OBSERVACAO"];
  const lines = [
    headers.map(csvCell).join(";"),
    ...rows.map((item) => {
      const data = item.data || {};
      return [
        item.id,
        data.codigo,
        data.lote,
        data.cor,
        data.veiculo_placa,
        data.status,
        data.data_compra,
        data.observacao,
      ].map(csvCell).join(";");
    }),
  ];
  const stamp = new Date().toISOString().slice(0, 10);
  downloadTextFile(`caixas_${stamp}.csv`, `${lines.join("\r\n")}\r\n`, "text/csv;charset=utf-8");
  notify(`${rows.length} caixa(s) exportada(s).`);
}

function printCaixasLabels() {
  const selected = currentCadastroItem();
  const rows = caixasRowsForOutput({selectedOnly: Boolean(selected)});
  if (!rows.length) {
    notify("Sem caixas para imprimir.", true);
    return;
  }
  const popup = window.open("", "_blank", "width=980,height=720");
  if (!popup) {
    notify("Nao foi possivel abrir a janela de etiquetas.", true);
    return;
  }
  const labels = rows.map((item) => {
    const data = item.data || {};
    return `
      <article class="label">
        <strong>${escapeHtml(data.codigo || item.id)}</strong>
        <span>Lote: ${escapeHtml(data.lote || "-")}</span>
        <span>Cor: ${escapeHtml(data.cor || "-")}</span>
        <span>Veiculo: ${escapeHtml(data.veiculo_placa || "ESTOQUE")}</span>
        <small>Status: ${escapeHtml(data.status || "-")}</small>
      </article>
    `;
  }).join("");
  popup.document.write(`
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>Etiquetas de caixas</title>
        <style>
          @page { size: A4; margin: 10mm; }
          * { box-sizing: border-box; }
          body { margin: 0; font-family: Arial, sans-serif; color: #111827; }
          .sheet { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8mm; }
          .label { min-height: 36mm; padding: 5mm; border: 1px solid #111827; display: grid; gap: 2mm; break-inside: avoid; }
          .label strong { font-size: 18pt; letter-spacing: 0; }
          .label span { font-size: 10pt; }
          .label small { font-size: 8pt; color: #4b5563; }
          .toolbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8mm; }
          .toolbar button { padding: 8px 12px; border: 1px solid #d1d5db; background: #fff; border-radius: 6px; font-weight: 700; }
          @media print { .toolbar { display: none; } }
        </style>
      </head>
      <body>
        <div class="toolbar">
          <strong>${escapeHtml(rows.length)} etiqueta(s)</strong>
          <button onclick="window.print()">Imprimir</button>
        </div>
        <main class="sheet">${labels}</main>
      </body>
    </html>
  `);
  popup.document.close();
}

function renderCadastroTable() {
  if (state.cadastroResource === "clientes") {
    renderClientesImportRows();
    return;
  }
  const resource = currentCadastroConfig();
  const head = el("cadastroTableHead");
  const rows = el("cadastroRows");
  head.innerHTML = "";
  rows.innerHTML = "";

  const trHead = document.createElement("tr");
  trHead.innerHTML = `<th>ID</th>${resource.fields.map(([, label]) => `<th>${escapeHtml(label)}</th>`).join("")}<th class="actions-col">Acoes</th>`;
  head.appendChild(trHead);

  if (!state.cadastroItems.length) {
    rows.appendChild(emptyRow(resource.fields.length + 2, "Sem registros."));
    return;
  }

  state.cadastroItems.forEach((item) => {
    const tr = document.createElement("tr");
    tr.className = state.cadastroSelectedId === item.id ? "selected-row" : "";
    tr.dataset.cadastroId = String(item.id);
    const cells = resource.fields
      .map(([name]) => `<td>${escapeHtml(item.data[name] ?? "-")}</td>`)
      .join("");
    tr.innerHTML = `
      <td>${item.id}</td>
      ${cells}
      <td>
        <div class="actions">
          <button type="button" class="secondary" data-cadastro-action="select" data-cadastro-id="${item.id}">Selecionar</button>
        </div>
      </td>
    `;
    tr.addEventListener("click", (event) => {
      if (event.target.closest("button")) return;
      selectCadastroItem(item);
    });
    rows.appendChild(tr);
  });

  rows.querySelectorAll("[data-cadastro-action]").forEach((button) => {
    button.addEventListener("click", () => handleCadastroAction(button.dataset.cadastroAction, Number(button.dataset.cadastroId)));
  });
  renderCadastroActions();
}

function handleCadastroAction(action, itemId) {
  const item = state.cadastroItems.find((candidate) => candidate.id === itemId);
  if (!item) return;
  if (action === "select") selectCadastroItem(item);
}

function cadastroPayloadFromForm(options = {}) {
  const form = el("cadastroForm");
  const payload = {};
  currentCadastroConfig().fields.forEach(([name, , type]) => {
    if (state.cadastroResource === "motoristas" && name === "codigo") return;
    if (options.statusOnly && name !== currentCadastroConfig().statusField) return;
    const raw = clean(form.elements[name].value);
    if (type === "number") {
      payload[name] = raw === "" ? null : Number(raw);
    } else {
      payload[name] = raw || null;
    }
  });
  return payload;
}

async function onSaveCadastroItem(event) {
  event.preventDefault();
  const resource = currentCadastroConfig();
  if (state.cadastroMode === "view") {
    notify("Clique em NOVO ou ALTERAR antes de salvar.", true);
    return;
  }
  const editingId = ["status", "edit"].includes(state.cadastroMode) ? state.cadastroSelectedId : null;
  if (["status", "edit"].includes(state.cadastroMode) && !editingId) {
    notify("Selecione um cadastro para alterar o status.", true);
    return;
  }
  const payload = cadastroPayloadFromForm({statusOnly: state.cadastroMode === "status"});
  const path = editingId ? `/cadastros/${resource.endpoint}/${editingId}` : `/cadastros/${resource.endpoint}`;
  const method = editingId ? "PATCH" : "POST";

  try {
    await apiRequest(path, {
      method,
      body: JSON.stringify(payload),
    });
    notify(editingId ? "Cadastro atualizado." : "Cadastro criado.");
    state.cadastroSelectedId = null;
    state.cadastroMode = "view";
    await loadCadastros();
    await loadDashboard();
  } catch (error) {
    const select = el("escalaFolgaPessoa");
    if (select) {
      select.innerHTML = '<option value="">Falha ao carregar</option>';
    }
    notify(error.message, true);
  }
}

async function quickCadastroStatus(liberar) {
  const resource = currentCadastroConfig();
  const item = currentCadastroItem();
  if (!item) {
    notify("Selecione um cadastro na tabela.", true);
    return;
  }
  if (!resource.statusField) {
    notify("Este cadastro nao possui campo STATUS.", true);
    return;
  }
  const nextStatus = liberar ? resource.activeStatus : resource.inactiveStatus;
  try {
    await apiRequest(`/cadastros/${resource.endpoint}/${item.id}`, {
      method: "PATCH",
      body: JSON.stringify({[resource.statusField]: nextStatus}),
    });
    notify(`Status atualizado para ${nextStatus}.`);
    state.cadastroMode = "view";
    await loadCadastros();
    await loadDashboard();
  } catch (error) {
    notify(error.message, true);
  }
}

async function deleteSelectedCadastroItem() {
  const resource = currentCadastroConfig();
  const item = currentCadastroItem();
  if (!item) {
    notify("Selecione um cadastro na tabela.", true);
    return;
  }
  if (!confirm(`Excluir registro ${item.id} de ${resource.label}?`)) return;
  try {
    await apiRequest(`/cadastros/${resource.endpoint}/${item.id}`, {method: "DELETE"});
    notify("Cadastro excluido.");
    state.cadastroSelectedId = null;
    state.cadastroMode = "view";
    await loadCadastros();
    await loadDashboard();
  } catch (error) {
    notify(error.message, true);
  }
}

function openCadastroPasswordDialog() {
  const resource = currentCadastroConfig();
  const item = currentCadastroItem();
  if (!item || !resource.password) {
    notify("Selecione um cadastro com senha.", true);
    return;
  }
  const form = el("passwordCadastroForm");
  form.reset();
  el("passwordCadastroLabel").textContent = `${resource.label} #${item.id}`;
  el("passwordDialog").showModal();
}

function closeCadastroPasswordDialog() {
  el("passwordDialog").close();
}

async function onChangeCadastroPassword(event) {
  event.preventDefault();
  const resource = currentCadastroConfig();
  const item = currentCadastroItem();
  if (!item || !resource.password) {
    closeCadastroPasswordDialog();
    return;
  }
  const form = event.currentTarget;
  const novaSenha = clean(form.elements.newPassword.value);
  const confirmar = clean(form.elements.confirmPassword.value);
  if (novaSenha !== confirmar) {
    notify("Confirmacao nao confere.", true);
    return;
  }
  try {
    await apiRequest(`/cadastros/${resource.endpoint}/${item.id}/senha`, {
      method: "POST",
      body: JSON.stringify({nova_senha: novaSenha}),
    });
    closeCadastroPasswordDialog();
    notify("Senha atualizada.");
    await loadCadastros();
    await loadDashboard();
  } catch (error) {
    notify(error.message, true);
  }
}

async function openFornecedorCertificadoDialog() {
  el("fornecedorCertificadoForm").reset();
  try {
    await loadComprasFornecedores();
    const select = el("fornecedorCertificadoFornecedor");
    const selectedNfe = ((state.comprasNfe && state.comprasNfe.rows) || [])
      .find((row) => Number(row.id) === Number(state.comprasNfe.selectedId));
    const selectedFornecedorId = Number(selectedNfe && selectedNfe.fornecedor_id) || 0;
    select.innerHTML = state.comprasFornecedores.map((item) => {
      const data = item.data || {};
      const label = `${data.razao_social || data.nome_fantasia || "Fornecedor"} | ${data.documento || "sem documento"}`;
      return `<option value="${escapeHtml(item.id)}">${escapeHtml(label)}</option>`;
    }).join("");
    if (selectedFornecedorId && state.comprasFornecedores.some((item) => Number(item.id) === selectedFornecedorId)) {
      select.value = String(selectedFornecedorId);
    }
    const option = select.options[select.selectedIndex];
    el("fornecedorCertificadoInfo").textContent = option
      ? `Manifestador vinculado a ${option.textContent}`
      : "Cadastre ou importe uma NF-e para vincular um fornecedor ao manifestador.";
    if (!state.comprasFornecedores.length) {
      notify("Nenhum fornecedor ativo encontrado para vincular certificado.", true);
      return;
    }
    el("fornecedorCertificadoDialog").showModal();
  } catch (error) {
    notify(error.message || "Falha ao carregar fornecedores para certificado.", true);
  }
}

function closeFornecedorCertificadoDialog() {
  const dialog = el("fornecedorCertificadoDialog");
  if (dialog.open) dialog.close();
}

function onFornecedorCertificadoChange() {
  const select = el("fornecedorCertificadoFornecedor");
  const option = select.options[select.selectedIndex];
  el("fornecedorCertificadoInfo").textContent = option
    ? `Manifestador vinculado a ${option.textContent}`
    : "Selecione o fornecedor vinculado ao certificado.";
}

async function onInstallFornecedorCertificado(event) {
  event.preventDefault();
  const fornecedorId = Number(el("fornecedorCertificadoFornecedor").value || 0);
  if (!fornecedorId) {
    notify("Selecione o fornecedor vinculado ao certificado.", true);
    return;
  }
  const form = event.currentTarget;
  const payload = new FormData(form);
  try {
    await apiRequest(`/cadastros/fornecedores/${encodeURIComponent(fornecedorId)}/certificado`, {
      method: "POST",
      body: payload,
    });
    closeFornecedorCertificadoDialog();
    notify("Certificado instalado e vinculado ao manifestador.");
    await loadComprasNfe();
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadComprasFornecedores() {
  const rows = await apiRequest("/cadastros/fornecedores?status_filter=ATIVO");
  state.comprasFornecedores = Array.isArray(rows) ? rows : [];
  return state.comprasFornecedores;
}

async function loadClientesDashboard() {
  try {
    state.clientesDashboard = await apiRequest("/cadastros/clientes/dashboard");
    renderClientesHub();
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadClientesLookup() {
  try {
    state.clientesLookup = await apiRequest("/cadastros/clientes/lookup");
    renderClientesSelectors();
  } catch (error) {
    notify(error.message, true);
  }
}

function showClientesSection(section) {
  state.clientesSection = section || "hub";
  renderCadastroLayout();
  if (state.clientesSection === "hub") {
    loadClientesDashboard();
    return;
  }
  if (state.clientesSection === "lista") {
    loadClientesImportRows();
    return;
  }
  if (!state.clientesLookup.length) {
    loadClientesLookup().then(() => {
      if (state.clientesSection === "historico") loadClienteHistoricoSelecionado();
      if (state.clientesSection === "localizacao") loadClienteLocalizacaoSelecionado();
    });
    return;
  }
  renderClientesSelectors();
  if (state.clientesSection === "historico") loadClienteHistoricoSelecionado();
  if (state.clientesSection === "localizacao") loadClienteLocalizacaoSelecionado();
}

function renderClientesHub() {
  if (!el("clientesHubPanel")) return;
  const data = state.clientesDashboard || {};
  el("clientesHubTotal").textContent = `${formatPlainNumber(data.total_clientes || 0)} clientes`;
  el("clientesHubHistorico").textContent = `${formatPlainNumber(data.clientes_com_historico || 0)} com historico`;
  el("clientesHubLocalizacao").textContent = `${formatPlainNumber(data.clientes_com_localizacao || 0)} clientes / ${formatPlainNumber(data.amostras_localizacao || 0)} pontos`;
  el("clientesHubInfo").textContent = "Clique em um bloco para acessar lista, historico operacional ou localizacoes.";
}

function renderClientesSelectors() {
  const options = (state.clientesLookup || []).map((item) => {
    const cod = clean(item.cod_cliente);
    const nome = clean(item.nome_cliente);
    return `<option value="${escapeHtml(cod)}">${escapeHtml(cod)} - ${escapeHtml(nome)}</option>`;
  }).join("");
  ["clientesHistoricoSelect", "clientesLocalizacaoSelect"].forEach((id) => {
    const select = el(id);
    if (!select) return;
    const current = select.value;
    select.innerHTML = options || '<option value="">Sem clientes cadastrados</option>';
    if (current && [...select.options].some((option) => option.value === current)) select.value = current;
  });
}

async function loadClienteHistoricoSelecionado() {
  const select = el("clientesHistoricoSelect");
  const cod = clean(select && select.value);
  if (!cod) {
    renderClienteHistorico({resumo: {}, rows: []});
    return;
  }
  try {
    const data = await apiRequest(`/cadastros/clientes/${encodeURIComponent(cod)}/historico`);
    renderClienteHistorico(data || {resumo: {}, rows: []});
  } catch (error) {
    notify(error.message, true);
  }
}

function renderClienteHistorico(data) {
  const resumo = data.resumo || {};
  const kpis = [
    ["Programacoes", resumo.total_programacoes || 0],
    ["Entregues", resumo.entregues || 0],
    ["Canceladas", resumo.canceladas || 0],
    ["Alteradas", resumo.alteradas || 0],
    ["Mortalidade", resumo.mortalidade_aves || 0],
    ["KG recebidos", formatPlainNumber(resumo.kg_recebidos || 0)],
    ["KG descontados", formatPlainNumber(resumo.kg_descontados || 0)],
  ];
  el("clientesHistoricoResumo").innerHTML = kpis.map(([label, value]) => `
    <div class="client-kpi"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>
  `).join("");
  const tbody = el("clientesHistoricoRows");
  tbody.innerHTML = "";
  const rows = Array.isArray(data.rows) ? data.rows : [];
  if (!rows.length) {
    tbody.appendChild(emptyRow(10, "Sem historico para este cliente."));
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = [
      row.data_ref || "",
      row.codigo_programacao || "",
      row.pedido || "",
      row.status_pedido || "",
      `${formatPlainNumber(row.caixas_atuais || 0)}/${formatPlainNumber(row.caixas_programadas || 0)}`,
      formatPlainNumber(row.kg_recebido || 0),
      formatPlainNumber(row.kg_descontado || 0),
      formatPlainNumber(row.mortalidade_aves || 0),
      formatPlainNumber(row.valor_recebido || 0),
      row.motorista || "",
    ].map((value) => `<td>${escapeHtml(value)}</td>`).join("");
    tbody.appendChild(tr);
  });
}

async function loadClienteLocalizacaoSelecionado() {
  const select = el("clientesLocalizacaoSelect");
  const cod = clean(select && select.value);
  if (!cod) {
    renderClienteLocalizacoes({resumo: {}, rows: []});
    return;
  }
  try {
    const data = await apiRequest(`/cadastros/clientes/${encodeURIComponent(cod)}/localizacoes`);
    renderClienteLocalizacoes(data || {resumo: {}, rows: []});
  } catch (error) {
    notify(error.message, true);
  }
}

function renderClienteLocalizacoes(data) {
  const resumo = data.resumo || {};
  const ultima = resumo.ultima || {};
  el("clientesLocalizacaoInfo").textContent = `Amostras: ${formatPlainNumber(resumo.amostras || 0)} | Com coordenada: ${formatPlainNumber(resumo.com_coordenada || 0)} | Ultima: ${ultima.registrado_em || "-"}`;
  const tbody = el("clientesLocalizacaoRows");
  tbody.innerHTML = "";
  const rows = Array.isArray(data.rows) ? data.rows : [];
  if (!rows.length) {
    tbody.appendChild(emptyRow(9, "Sem localizacao registrada para este cliente."));
    return;
  }
  rows.forEach((row) => {
    const endereco = [row.endereco, row.bairro, row.cidade].map(clean).filter(Boolean).join(" - ");
    const tr = document.createElement("tr");
    tr.innerHTML = [
      row.registrado_em || "",
      row.codigo_programacao || "",
      row.pedido || "",
      row.status_pedido || "",
      row.latitude ?? "",
      row.longitude ?? "",
      endereco,
      row.motorista_nome || row.motorista_codigo || "",
      row.origem || "",
    ].map((value) => `<td>${escapeHtml(value)}</td>`).join("");
    tbody.appendChild(tr);
  });
}

async function loadClientesImportRows() {
  renderCadastroLayout();
  try {
    state.clientesImportRows = await apiRequest("/cadastros/clientes/importacao/rows");
    renderClientesImportRows();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderClientesImportRows() {
  if (!el("clientesImportRows")) return;
  const rows = el("clientesImportRows");
  rows.innerHTML = "";
  el("clientesImportInfo").textContent = `${state.clientesImportRows.length} cliente(s) carregado(s). Edite direto na grade e salve em lote.`;
  if (!state.clientesImportRows.length) {
    rows.appendChild(emptyRow(5, "Sem clientes. Use INSERIR LINHA ou IMPORTAR CLIENTES (EXCEL)."));
    return;
  }

  state.clientesImportRows.forEach((cliente, index) => {
    const tr = document.createElement("tr");
    tr.dataset.clienteIndex = String(index);
    tr.innerHTML = ["cod_cliente", "nome_cliente", "endereco", "telefone", "vendedor"]
      .map((field) => `
        <td>
          <input
            data-clientes-import-field="${field}"
            value="${escapeHtml(cliente[field] || "")}"
            autocomplete="off"
          >
        </td>
      `)
      .join("");
    rows.appendChild(tr);
  });
}

function insertClientesImportRow() {
  state.clientesImportRows.push({
    id: null,
    cod_cliente: "",
    nome_cliente: "",
    endereco: "",
    telefone: "",
    vendedor: "",
  });
  renderClientesImportRows();
  const lastRow = el("clientesImportRows").querySelector("tr:last-child input");
  if (lastRow) lastRow.focus();
}

function clientesImportRowsFromTable() {
  const payload = [];
  const seen = new Set();
  const tableRows = [...el("clientesImportRows").querySelectorAll("tr[data-cliente-index]")];
  for (const row of tableRows) {
    const data = {};
    row.querySelectorAll("[data-clientes-import-field]").forEach((input) => {
      data[input.dataset.clientesImportField] = clean(input.value);
    });
    const hasAny = Object.values(data).some((value) => clean(value));
    if (!hasAny) continue;
    if (!clean(data.cod_cliente) || !clean(data.nome_cliente)) {
      throw new Error("Todas as linhas precisam ter pelo menos COD CLIENTE e NOME CLIENTE.");
    }
    const key = clean(data.cod_cliente).toUpperCase();
    if (seen.has(key)) {
      throw new Error(`COD CLIENTE duplicado na tabela: ${key}`);
    }
    seen.add(key);
    payload.push(data);
  }
  return payload;
}

async function downloadClientesImportTemplate() {
  try {
    const result = await apiBlobRequest("/cadastros/clientes/importacao/modelo");
    downloadBlob(result.blob, result.filename || "MODELO_IMPORTACAO_CLIENTES.xlsx");
  } catch (error) {
    notify(error.message, true);
  }
}

async function saveClientesImportRows() {
  let rows;
  try {
    rows = clientesImportRowsFromTable();
  } catch (error) {
    notify(error.message, true);
    return;
  }
  if (!rows.length) {
    notify("Nenhuma linha valida para salvar.", true);
    return;
  }
  try {
    const result = await apiRequest("/cadastros/clientes/importacao/bulk-upsert", {
      method: "POST",
      body: JSON.stringify({rows}),
    });
    notify(clientesImportResultMessage(result));
    await Promise.all([loadClientesImportRows(), loadClientesDashboard(), loadClientesLookup(), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

async function onClientesImportUpload(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const fileInput = el("clientesImportFile");
  if (!fileInput.files.length) {
    notify("Selecione um arquivo Excel.", true);
    return;
  }
  try {
    const result = await apiRequest("/cadastros/clientes/importacao/upload", {
      method: "POST",
      body: new FormData(form),
    });
    form.reset();
    notify(clientesImportResultMessage(result));
    await Promise.all([loadClientesImportRows(), loadClientesDashboard(), loadClientesLookup(), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

function clientesImportResultMessage(result) {
  return [
    `Clientes salvos/atualizados: ${result.total || 0}`,
    `Inseridos: ${result.inseridos || 0}`,
    `Atualizados: ${result.atualizados || 0}`,
    `Ignorados: ${result.ignorados || 0}`,
  ].join(" | ");
}

async function loadRotasMonitoramento() {
  try {
    state.rotas = await apiRequest("/rotas/monitoramento");
    renderRotasMonitoramento();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderRotasMonitoramento() {
  const rows = el("rotasRows");
  rows.innerHTML = "";
  const gpsCount = state.rotas.filter((rota) => rotaHasGps(rota)).length;
  el("rotasInfo").textContent = `${state.rotas.length} rota(s) ativa(s). GPS em ${gpsCount} rota(s).`;

  if (!state.rotas.length) {
    rows.appendChild(emptyRow(10, "Sem rotas ativas."));
    return;
  }

  state.rotas.forEach((rota) => {
    const tr = document.createElement("tr");
    tr.dataset.rotaCodigo = rota.codigo_programacao || "";
    tr.className = state.rotasSelectedCodigo === rota.codigo_programacao ? "selected-row" : "";
    const speed = rota.speed === null || rota.speed === undefined ? "" : formatNumber(rota.speed);
    tr.innerHTML = `
      <td>${escapeHtml(rota.codigo_programacao || "")}</td>
      <td>${escapeHtml(rota.motorista || "")}</td>
      <td>${escapeHtml(rota.veiculo || "")}</td>
      <td>${escapeHtml(rota.status || "")}</td>
      <td>${escapeHtml(rota.recorded_at || "")}</td>
      <td>${escapeHtml(formatCoord(rota.lat))}</td>
      <td>${escapeHtml(formatCoord(rota.lon))}</td>
      <td class="${speedClass(rota.speed)}">${escapeHtml(speed)}</td>
      <td>${escapeHtml(rota.accuracy === null || rota.accuracy === undefined ? "" : formatNumber(rota.accuracy))}</td>
      <td>
        <div class="actions">
          <button type="button" class="secondary" data-rota-action="select" data-rota-code="${escapeHtml(rota.codigo_programacao || "")}">Selecionar</button>
          <button type="button" class="primary" data-rota-action="map" data-rota-code="${escapeHtml(rota.codigo_programacao || "")}">Mapa</button>
        </div>
      </td>
    `;
    tr.addEventListener("click", (event) => {
      if (event.target.closest("button")) return;
      selectRota(rota.codigo_programacao);
    });
    tr.addEventListener("dblclick", () => openRotaMap(rota));
    rows.appendChild(tr);
  });

  rows.querySelectorAll("[data-rota-action]").forEach((button) => {
    button.addEventListener("click", () => {
      const rota = state.rotas.find((item) => item.codigo_programacao === button.dataset.rotaCode);
      if (!rota) return;
      if (button.dataset.rotaAction === "select") selectRota(rota.codigo_programacao);
      if (button.dataset.rotaAction === "map") openRotaMap(rota);
    });
  });
}

function selectRota(codigo) {
  state.rotasSelectedCodigo = codigo || "";
  renderRotasMonitoramento();
}

function rotaHasGps(rota) {
  return Number.isFinite(Number(rota?.lat)) && Number.isFinite(Number(rota?.lon));
}

function speedClass(speed) {
  const value = Number(speed || 0);
  if (value >= 60) return "speed-high";
  if (value >= 30) return "speed-medium";
  if (value > 0) return "speed-low";
  return "";
}

function formatCoord(value) {
  if (value === null || value === undefined || value === "") return "";
  const number = Number(value);
  return Number.isFinite(number) ? number.toFixed(6).replace(".", ",") : "";
}

function openRotasMapAll() {
  const points = state.rotas.filter(rotaHasGps);
  if (!points.length) {
    notify("Nenhuma coordenada GPS disponivel para exibir no mapa.", true);
    return;
  }
  openTrackingMap({
    tipo: "todas",
    titulo: "Monitoramento de rotas",
    rotas: points,
    gerado_em: new Date().toLocaleString("pt-BR"),
  });
}

function openRotasMapSelected() {
  const rota = state.rotas.find((item) => item.codigo_programacao === state.rotasSelectedCodigo);
  if (!rota) {
    notify("Selecione uma rota na tabela.", true);
    return;
  }
  openRotaMap(rota);
}

async function openRotaMap(rota) {
  if (!rotaHasGps(rota)) {
    notify("Rota sem coordenada GPS.", true);
    return;
  }
  let tracking = null;
  try {
    tracking = await apiRequest(`/rotas/${encodeURIComponent(rota.codigo_programacao || "")}/rastreamento?limit=2000`);
  } catch (error) {
    notify(`Mapa aberto com ponto atual. Historico nao carregou: ${error.message}`, true);
  }
  openTrackingMap({
    tipo: "rota",
    titulo: `Rastreamento ${rota.codigo_programacao || ""}`,
    rota,
    rotas: [rota],
    tracking,
    gerado_em: new Date().toLocaleString("pt-BR"),
  });
}

function openTrackingMap(payload) {
  const popup = window.open("", "_blank", "width=1280,height=820");
  if (!popup) {
    notify("O navegador bloqueou a janela do mapa. Libere pop-ups para abrir o acompanhamento.", true);
    return;
  }
  popup.document.open();
  popup.document.write(trackingMapHtml(payload));
  popup.document.close();
}

function trackingMapHtml(payload) {
  const data = JSON.stringify({...payload, access_token: state.token}).replace(/</g, "\\u003c");
  return `<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(payload.titulo || "Monitoramento")}</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: Arial, sans-serif; color: #102033; background: #edf2f7; }
    .layout { display: grid; grid-template-columns: minmax(360px, 29vw) 1fr; min-height: 100vh; }
    aside { background: #fff; border-right: 1px solid #d5dee8; overflow: auto; }
    header { padding: 16px 18px 12px; border-bottom: 1px solid #d5dee8; }
    h1 { margin: 0 0 4px; font-size: 22px; }
    .muted { color: #5e6f82; font-size: 13px; }
    .section { padding: 14px 18px; border-bottom: 1px solid #d5dee8; }
    .grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
    .metric { background: #f7fafc; border: 1px solid #d9e2ec; border-radius: 6px; padding: 10px; min-height: 64px; }
    .metric span { display: block; color: #5e6f82; font-size: 12px; font-weight: 700; text-transform: uppercase; }
    .metric strong { display: block; margin-top: 6px; font-size: 18px; }
    #map { min-height: 100vh; width: 100%; }
    .list { display: grid; gap: 8px; }
    .item { border: 1px solid #d9e2ec; border-radius: 6px; padding: 10px; background: #fff; }
    .item strong { display: block; margin-bottom: 3px; }
    .actions { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 10px; }
    a.button { display: inline-flex; align-items: center; justify-content: center; min-height: 38px; padding: 0 12px; border: 1px solid #0f766e; border-radius: 6px; color: #0f766e; text-decoration: none; font-weight: 700; }
    .status { display: inline-flex; padding: 4px 8px; border-radius: 999px; background: #e7f5ef; color: #0f766e; font-weight: 700; }
    .refresh { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
    .refresh button { min-height: 34px; border: 1px solid #0f766e; border-radius: 6px; background: #0f766e; color: #fff; font-weight: 700; cursor: pointer; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 8px; border-bottom: 1px solid #d9e2ec; text-align: left; }
    th { background: #f7fafc; font-size: 12px; text-transform: uppercase; }
    @media (max-width: 900px) { .layout { grid-template-columns: 1fr; } #map { min-height: 62vh; } }
  </style>
</head>
<body>
  <div class="layout">
    <aside>
      <header>
        <h1 id="title">Monitoramento</h1>
        <div class="muted" id="subtitle"></div>
        <div class="refresh">
          <div class="muted" id="refreshInfo">Atualizacao automatica a cada 15s</div>
          <button type="button" id="refreshButton">Atualizar</button>
        </div>
      </header>
      <div class="section">
        <div class="grid" id="metrics"></div>
        <div class="actions" id="actions"></div>
      </div>
      <div class="section">
        <h2 style="font-size:16px;margin:0 0 10px;">Paradas detectadas</h2>
        <div class="list" id="stops"></div>
      </div>
      <div class="section">
        <h2 style="font-size:16px;margin:0 0 10px;">Rotas no mapa</h2>
        <div id="routes"></div>
      </div>
    </aside>
    <main><div id="map"></div></main>
  </div>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"><\/script>
  <script>
    const payload = ${data};
    let map = null;
    let layers = null;

    function text(value) {
      return String(value === null || value === undefined || value === "" ? "-" : value);
    }
    function html(value) {
      return text(value).replace(/[&<>"']/g, function (c) {
        return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;","'":"&#039;"}[c];
      });
    }
    function validPoint(point) {
      return Number.isFinite(Number(point && point.lat)) && Number.isFinite(Number(point && point.lon));
    }
    function num(value, digits) {
      const parsed = Number(value || 0);
      return Number.isFinite(parsed) ? parsed.toLocaleString("pt-BR", { maximumFractionDigits: digits || 0 }) : "0";
    }
    function when(value) {
      if (!value) return "-";
      const date = new Date(String(value).replace(" ", "T"));
      return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString("pt-BR");
    }
    function googlePoint(point) {
      return "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(point.lat + "," + point.lon);
    }
    function googleRoute(points) {
      const limited = points.slice(0, 10);
      return "https://www.google.com/maps/dir/" + limited.map(function (point) {
        return encodeURIComponent(point.lat + "," + point.lon);
      }).join("/");
    }
    function metric(label, value) {
      return '<div class="metric"><span>' + html(label) + '</span><strong>' + html(value) + '</strong></div>';
    }
    function authHeaders() {
      return payload.access_token ? { Authorization: "Bearer " + payload.access_token } : {};
    }

    function currentData() {
      const route = payload.rota || {};
      const tracking = payload.tracking || {};
      const routePoints = (tracking.pontos || []).filter(validPoint);
      const latestRoutes = (payload.rotas || []).filter(validPoint);
      const livePoints = routePoints.length ? routePoints : latestRoutes;
      return { route: route, tracking: tracking, routePoints: routePoints, latestRoutes: latestRoutes, livePoints: livePoints };
    }
    function renderPanel() {
      const data = currentData();
      const route = data.route;
      const tracking = data.tracking;
      const routePoints = data.routePoints;
      const latestRoutes = data.latestRoutes;
      const livePoints = data.livePoints;
      document.getElementById("title").textContent = payload.titulo || "Monitoramento";
      document.getElementById("subtitle").textContent = "Gerado em " + text(payload.gerado_em);
      const lastPoint = livePoints[livePoints.length - 1] || {};
      document.getElementById("metrics").innerHTML = [
        metric("Programacao", tracking.codigo_programacao || route.codigo_programacao || "-"),
        metric("Motorista", tracking.motorista || route.motorista || "-"),
        metric("Veiculo", tracking.veiculo || route.veiculo || "-"),
        metric("Status", tracking.status || route.status || "-"),
        metric("Km rodado", num(tracking.km_estimado, 2) + " km"),
        metric("Vel. media", num(tracking.velocidade_media, 1) + " km/h"),
        metric("Vel. maxima", num(tracking.velocidade_maxima, 1) + " km/h"),
        metric("Tempo rastreado", num(tracking.tempo_rastreado_min, 1) + " min"),
        metric("Pontos GPS", tracking.total_pontos || livePoints.length),
        metric("Ultima atualizacao", when(tracking.ultima_atualizacao || lastPoint.recorded_at || route.recorded_at)),
      ].join("");

      const actions = [];
      if (validPoint(lastPoint)) actions.push('<a class="button" target="_blank" rel="noopener" href="' + googlePoint(lastPoint) + '">Abrir ponto no Google</a>');
      if (routePoints.length > 1) actions.push('<a class="button" target="_blank" rel="noopener" href="' + googleRoute(routePoints) + '">Abrir trajeto no Google</a>');
      document.getElementById("actions").innerHTML = actions.join("");

      const stops = tracking.paradas || [];
      document.getElementById("stops").innerHTML = stops.length
        ? stops.map(function (stop, index) {
            return '<div class="item"><strong>Parada ' + (index + 1) + ' - ' + num(stop.minutos, 1) + ' min</strong><div class="muted">' + when(stop.inicio) + ' ate ' + when(stop.fim) + '</div><div class="muted">' + html(stop.lat) + ', ' + html(stop.lon) + ' | pontos: ' + html(stop.pontos) + '</div></div>';
          }).join("")
        : '<div class="muted">Nenhuma parada detectada pelo historico recebido.</div>';

      document.getElementById("routes").innerHTML = latestRoutes.length
        ? '<table><thead><tr><th>Prog.</th><th>Motorista</th><th>Vel.</th></tr></thead><tbody>' + latestRoutes.map(function (item) {
            return '<tr><td>' + html(item.codigo_programacao) + '</td><td>' + html(item.motorista) + '</td><td>' + num(item.speed, 1) + '</td></tr>';
          }).join("") + '</tbody></table>'
        : '<div class="muted">Nenhuma rota com GPS disponivel.</div>';
    }
    function ensureMap() {
      if (!window.L) {
        document.getElementById("map").innerHTML = '<div style="padding:24px;">Mapa indisponivel. Verifique a conexao com a internet para carregar o Leaflet/OpenStreetMap.</div>';
        return false;
      }
      if (map) return true;
      map = L.map("map", { zoomControl: true });
      L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        maxZoom: 19,
        attribution: "&copy; OpenStreetMap",
      }).addTo(map);
      layers = L.layerGroup().addTo(map);
      return true;
    }
    function renderMap() {
      const data = currentData();
      const routePoints = data.routePoints;
      const latestRoutes = data.latestRoutes;
      const livePoints = data.livePoints;
      if (!livePoints.length) {
        document.getElementById("map").innerHTML = '<div style="padding:24px;">Nenhum ponto GPS disponivel para esta rota.</div>';
        return;
      }
      if (!ensureMap()) return;
      layers.clearLayers();
      const bounds = [];
      function addToBounds(point) {
        const latLon = [Number(point.lat), Number(point.lon)];
        bounds.push(latLon);
        return latLon;
      }
      if (routePoints.length > 1) {
        L.polyline(routePoints.map(addToBounds), { color: "#0f766e", weight: 5, opacity: 0.85 }).addTo(layers);
      }
      latestRoutes.forEach(function (point) {
        const latLon = addToBounds(point);
        L.circleMarker(latLon, { radius: 8, color: "#0f4c81", weight: 3, fillColor: "#38bdf8", fillOpacity: 0.9 })
          .addTo(layers)
          .bindPopup("<strong>" + html(point.codigo_programacao) + "</strong><br>" + html(point.motorista) + "<br>Vel.: " + num(point.speed, 1) + " km/h<br>" + when(point.recorded_at));
      });
      if (routePoints.length) {
        const start = routePoints[0];
        const end = routePoints[routePoints.length - 1];
        L.circleMarker(addToBounds(start), { radius: 7, color: "#166534", fillColor: "#22c55e", fillOpacity: 1 }).addTo(layers).bindPopup("Inicio<br>" + when(start.recorded_at));
        L.circleMarker(addToBounds(end), { radius: 10, color: "#1d4ed8", fillColor: "#2563eb", fillOpacity: 1 }).addTo(layers).bindPopup("Atual<br>" + when(end.recorded_at) + "<br>Vel.: " + num(end.speed, 1) + " km/h");
      }
      (data.tracking.paradas || []).forEach(function (stop, index) {
        if (!validPoint(stop)) return;
        L.circleMarker(addToBounds(stop), { radius: 9, color: "#b45309", fillColor: "#f59e0b", fillOpacity: 0.95 })
          .addTo(layers)
          .bindPopup("Parada " + (index + 1) + "<br>" + num(stop.minutos, 1) + " min<br>" + when(stop.inicio) + " ate " + when(stop.fim));
      });
      if (bounds.length === 1) map.setView(bounds[0], 15);
      else map.fitBounds(bounds, { padding: [32, 32] });
    }
    function renderAll() {
      renderPanel();
      renderMap();
    }
    async function refreshTracking() {
      const code = (payload.rota && payload.rota.codigo_programacao) || (payload.tracking && payload.tracking.codigo_programacao);
      if (!code || payload.tipo !== "rota") return;
      try {
        const response = await fetch("/api/v1/rotas/" + encodeURIComponent(code) + "/rastreamento?limit=2000", { headers: authHeaders() });
        if (!response.ok) throw new Error(response.status + " " + response.statusText);
        payload.tracking = await response.json();
        document.getElementById("refreshInfo").textContent = "Atualizado em " + new Date().toLocaleTimeString("pt-BR") + " | auto 15s";
        renderAll();
      } catch (error) {
        document.getElementById("refreshInfo").textContent = "Falha ao atualizar: " + error.message;
      }
    }
    async function refreshAllRoutes() {
      if (payload.tipo !== "todas") return;
      try {
        const response = await fetch("/api/v1/rotas/monitoramento", { headers: authHeaders() });
        if (!response.ok) throw new Error(response.status + " " + response.statusText);
        payload.rotas = await response.json();
        document.getElementById("refreshInfo").textContent = "Atualizado em " + new Date().toLocaleTimeString("pt-BR") + " | auto 15s";
        renderAll();
      } catch (error) {
        document.getElementById("refreshInfo").textContent = "Falha ao atualizar: " + error.message;
      }
    }
    function refreshNow() {
      if (payload.tipo === "rota") refreshTracking();
      if (payload.tipo === "todas") refreshAllRoutes();
    }
    document.getElementById("refreshButton").addEventListener("click", refreshNow);
    renderAll();
    window.setInterval(refreshNow, 15000);
  <\/script>
</body>
</html>`;
}

function startRotasAutoRefresh() {
  stopRotasAutoRefresh();
  rotasRefreshTimer = window.setInterval(() => {
    if (state.currentView === "rotas") loadRotasMonitoramento();
  }, state.rotasRefreshMs);
}

function stopRotasAutoRefresh() {
  if (rotasRefreshTimer) {
    window.clearInterval(rotasRefreshTimer);
    rotasRefreshTimer = null;
  }
}

function onRotasRefreshIntervalChange() {
  state.rotasRefreshMs = Number(el("rotasRefreshInterval").value) || 10000;
  startRotasAutoRefresh();
  notify(`Atualizacao automatica ajustada para ${el("rotasRefreshInterval").selectedOptions[0].textContent}.`);
}

async function loadEscalaResumo() {
  const seq = ++state.escalaLoadSeq;
  const refreshButton = el("escalaRefreshButton");
  const pdfButton = el("escalaPdfButton");
  try {
    refreshButton.disabled = true;
    refreshButton.textContent = "CARREGANDO...";
    pdfButton.disabled = true;
    if (!state.escalaFolgaPessoas.length) {
      await loadEscalaFolgaPessoas();
    }
    state.escalaPeriodo = el("escalaPeriodo").value;
    state.escalaStatus = el("escalaStatus").value;
    const params = new URLSearchParams({
      periodo: state.escalaPeriodo,
      status: state.escalaStatus,
    });
    const resumo = await apiRequest(`/escala/resumo?${params.toString()}`);
    if (seq !== state.escalaLoadSeq) return;
    state.escala = resumo;
    renderEscalaResumo();
  } catch (error) {
    if (seq !== state.escalaLoadSeq) return;
    notify(error.message, true);
  } finally {
    if (seq === state.escalaLoadSeq) {
      refreshButton.disabled = false;
      refreshButton.textContent = "ATUALIZAR";
      pdfButton.disabled = false;
    }
  }
}

function onEscalaFilterChange() {
  loadEscalaResumo();
}

async function refreshEscalaCompleta() {
  await loadEscalaFolgaPessoas();
  await loadEscalaResumo();
}

function initEscalaFolgaDates() {
  const today = new Date().toISOString().slice(0, 10);
  if (!el("escalaFolgaInicio").value) el("escalaFolgaInicio").value = today;
  if (!el("escalaFolgaFim").value) el("escalaFolgaFim").value = today;
}

function onEscalaFolgaDateChange() {
  const inicio = el("escalaFolgaInicio").value;
  const fim = el("escalaFolgaFim").value;
  if (inicio && (!fim || fim < inicio)) {
    el("escalaFolgaFim").value = inicio;
  }
}

async function loadEscalaFolgaPessoas() {
  const seq = ++state.escalaFolgaLoadSeq;
  try {
    const tipo = el("escalaFolgaTipo").value || "MOTORISTA";
    const select = el("escalaFolgaPessoa");
    const previous = select.value;
    select.disabled = true;
    select.innerHTML = '<option value="">Carregando...</option>';
    const pessoas = await apiRequest(`/escala/pessoas?tipo=${encodeURIComponent(tipo)}`);
    if (seq !== state.escalaFolgaLoadSeq) return;
    state.escalaFolgaPessoas = Array.isArray(pessoas) ? pessoas : [];
    select.innerHTML = "";
    if (!state.escalaFolgaPessoas.length) {
      select.innerHTML = '<option value="">Nenhuma pessoa ativa</option>';
      select.disabled = true;
      return;
    }
    state.escalaFolgaPessoas.forEach((item, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = item.label || item.pessoa_nome || "-";
      select.appendChild(option);
    });
    if (previous && [...select.options].some((option) => option.value === previous)) {
      select.value = previous;
    }
    select.disabled = false;
  } catch (error) {
    state.escalaFolgaPessoas = [];
    const select = el("escalaFolgaPessoa");
    select.innerHTML = '<option value="">Erro ao carregar</option>';
    select.disabled = true;
    notify(error.message, true);
  }
}

async function aplicarEscalaFolga() {
  const button = el("escalaFolgaAplicarButton");
  try {
    const selectedIndex = Number(el("escalaFolgaPessoa").value);
    const pessoa = state.escalaFolgaPessoas[selectedIndex];
    if (!pessoa) {
      notify("Selecione motorista ou ajudante para aplicar folga.", true);
      return;
    }
    const inicio = el("escalaFolgaInicio").value;
    const fim = el("escalaFolgaFim").value || inicio;
    if (!inicio || !fim) {
      notify("Informe inicio e fim da folga.", true);
      return;
    }
    if (fim < inicio) {
      notify("A data final da folga nao pode ser menor que a inicial.", true);
      el("escalaFolgaFim").value = inicio;
      return;
    }
    button.disabled = true;
    button.textContent = "Aplicando...";
    await apiRequest("/escala/folgas", {
      method: "POST",
      body: JSON.stringify({
        tipo: pessoa.tipo,
        pessoa_id: pessoa.pessoa_id || "",
        pessoa_codigo: pessoa.pessoa_codigo || "",
        pessoa_nome: pessoa.pessoa_nome || "",
        data_inicio: inicio,
        data_fim: fim,
        motivo: el("escalaFolgaMotivo").value || "",
      }),
    });
    el("escalaFolgaMotivo").value = "";
    notify("Folga aplicada.");
    await refreshEscalaCompleta();
  } catch (error) {
    notify(error.message, true);
  } finally {
    button.disabled = false;
    button.textContent = "Aplicar folga";
  }
}

async function encerrarEscalaFolga(id, button = null) {
  try {
    if (!id) return;
    const ok = window.confirm("Encerrar esta folga?");
    if (!ok) return;
    if (button) {
      button.disabled = true;
      button.textContent = "Encerrando...";
    }
    await apiRequest(`/escala/folgas/${encodeURIComponent(id)}/encerrar`, {method: "PATCH"});
    notify("Folga encerrada.");
    await refreshEscalaCompleta();
  } catch (error) {
    notify(error.message, true);
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = "Encerrar";
    }
  }
}

async function downloadEscalaPdf() {
  try {
    state.escalaPeriodo = el("escalaPeriodo").value;
    state.escalaStatus = el("escalaStatus").value;
    const params = new URLSearchParams({
      periodo: state.escalaPeriodo,
      status: state.escalaStatus,
    });
    const result = await apiBlobRequest(`/escala/pdf?${params.toString()}`);
    downloadBlob(result.blob, result.filename || `ESCALA_${state.escalaPeriodo}_${state.escalaStatus}.pdf`);
  } catch (error) {
    notify(error.message, true);
  }
}

function setEscalaTab(tab) {
  state.escalaTab = tab === "ajudantes" ? "ajudantes" : "motoristas";
  renderEscalaTable();
}

function renderEscalaResumo() {
  const data = state.escala || {kpis: {}, motoristas: [], ajudantes: [], chart: [], resumo: "", recomendacoes: ""};
  const kpis = data.kpis || {};
  el("escalaInfo").textContent = `${kpis.rotas || 0} rota(s) no filtro ${escalaPeriodoLabel(data.periodo || state.escalaPeriodo)} / ${escalaStatusLabel(data.status || state.escalaStatus)}. Folgas: ${kpis.folgas_motoristas || 0} motorista(s), ${kpis.folgas_ajudantes || 0} ajudante(s).`;
  el("escalaResumoText").textContent = data.resumo || "-";
  el("escalaRecomendacoesText").textContent = data.recomendacoes || "-";

  const cards = [
    ["Rotas no periodo", kpis.rotas || 0],
    ["Motoristas", kpis.motoristas || 0],
    ["Ajudantes", kpis.ajudantes || 0],
    ["Folgas motoristas", kpis.folgas_motoristas || 0],
    ["Folgas ajudantes", kpis.folgas_ajudantes || 0],
    ["Ocorrencias/media", formatNumber(kpis.mortalidade_media || 0)],
    ["KM total", formatOne(kpis.km_total || 0)],
    ["KM medio/motorista", formatOne(kpis.km_medio_motorista || 0)],
    ["Media km/L", formatNumber(kpis.media_km_l || 0)],
    ["Horas medias/motorista", formatNumber(kpis.horas_medias_motorista || 0)],
  ];
  el("escalaKpis").innerHTML = cards.map(([label, value]) => `
    <article class="metric escala-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `).join("");
  renderEscalaChart();
  renderEscalaTable();
  renderEscalaFolgas();
}

function renderEscalaFolgas() {
  const wrap = el("escalaFolgasAtivas");
  const folgas = state.escala?.folgas || [];
  if (!folgas.length) {
    wrap.innerHTML = '<span>Folgas no periodo: nenhuma</span>';
    return;
  }
  wrap.innerHTML = folgas.map((item) => `
    <span class="escala-folga-chip">
      <strong>${escapeHtml(item.tipo || "")}</strong>
      ${escapeHtml(item.pessoa_nome || "")}
      <em>${escapeHtml(formatDateBR(item.data_inicio))} ate ${escapeHtml(formatDateBR(item.data_fim))}</em>
      ${item.motivo ? `<small>${escapeHtml(item.motivo)}</small>` : ""}
      <button type="button" data-escala-folga-close="${escapeHtml(item.id)}" title="Encerrar folga">Encerrar</button>
    </span>
  `).join("");
  wrap.querySelectorAll("[data-escala-folga-close]").forEach((button) => {
    button.addEventListener("click", () => encerrarEscalaFolga(button.dataset.escalaFolgaClose, button));
  });
}

function renderEscalaChart() {
  const chart = el("escalaChart");
  chart.innerHTML = "";
  const rows = (state.escala?.chart || []).slice(0, 8);
  if (!rows.length) {
    chart.innerHTML = '<p class="empty-row">Sem dados para o periodo selecionado.</p>';
    return;
  }
  const maxHoras = Math.max(...rows.map((item) => Number(item.horas || 0)), 1);
  rows.forEach((item) => {
    const horas = Number(item.horas || 0);
    const width = Math.max(3, Math.round((horas / maxHoras) * 100));
    const row = document.createElement("div");
    row.className = "escala-chart-row";
    row.innerHTML = `
      <span>${escapeHtml(item.nome || "-")}</span>
      <div class="escala-chart-track"><div class="escala-chart-bar" style="width:${width}%"></div></div>
      <strong>${escapeHtml(formatNumber(horas))}h</strong>
      <em>${escapeHtml(item.dias || 0)}d</em>
    `;
    chart.appendChild(row);
  });
}

function renderEscalaTable() {
  const activeTab = state.escalaTab === "ajudantes" ? "ajudantes" : "motoristas";
  document.querySelectorAll("[data-escala-tab]").forEach((button) => {
    button.classList.toggle("active", button.dataset.escalaTab === activeTab);
  });
  el("escalaTableTitle").textContent = activeTab === "ajudantes" ? "Ajudantes" : "Motoristas";

  const rows = el("escalaRows");
  rows.innerHTML = "";
  const dataRows = state.escala?.[activeTab] || [];
  if (!dataRows.length) {
    rows.appendChild(emptyRow(10, "Sem dados para o filtro selecionado."));
    return;
  }
  dataRows.forEach((item) => {
    const tr = document.createElement("tr");
    tr.className = `${cargaClass(item.carga)}${item.em_folga ? " escala-folga-row" : ""}`;
    tr.innerHTML = `
      <td>${escapeHtml(item.nome || "")}${item.em_folga ? ' <span class="status-pill escala-folga-pill">FOLGA</span>' : ""}</td>
      <td>${escapeHtml(item.rotas || 0)}</td>
      <td>${escapeHtml(item.em_rota || 0)}</td>
      <td>${escapeHtml(item.ativas || 0)}</td>
      <td>${escapeHtml(item.finalizadas || 0)}</td>
      <td>${escapeHtml(item.canceladas || 0)}</td>
      <td>${escapeHtml(activeTab === "motoristas" ? (item.local || "-") : "-")}</td>
      <td>${escapeHtml(formatOne(item.km_rodado || 0))}</td>
      <td>${escapeHtml(formatNumber(item.horas_trab || 0))}</td>
      <td><span class="status-pill ${cargaClass(item.carga)}">${escapeHtml(cargaLabel(item.carga))}</span></td>
    `;
    rows.appendChild(tr);
  });
}

function escalaPeriodoLabel(value) {
  const text = String(value || "").toUpperCase();
  return text === "TODAS" ? "Todas" : `${text} dias`;
}

function escalaStatusLabel(value) {
  return String(value || "").replace("_", " ");
}

function formatDateBR(value) {
  const text = String(value || "").slice(0, 10);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return text || "-";
  return `${match[3]}/${match[2]}/${match[1]}`;
}

function cargaClass(carga) {
  if (carga === "sobrecarga") return "carga-sobrecarga";
  if (carga === "alerta") return "carga-alerta";
  return "carga-equilibrada";
}

function cargaLabel(carga) {
  if (carga === "sobrecarga") return "SOBRECARGA";
  if (carga === "alerta") return "ALERTA";
  return "EQUILIBRADA";
}

function formatOne(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) ? number.toFixed(1) : "0.0";
}

async function loadRecebimentosProgramacoes() {
  try {
    const selected = state.recebimentosSelectedCodigo;
    state.recebimentosProgramacoes = await apiRequest("/recebimentos/programacoes");
    const selectedStillAvailable = selected && state.recebimentosProgramacoes.some((item) => item.codigo_programacao === selected);
    const selectedHasLoadedBundle = selected && state.recebimentosBundle?.cabecalho?.codigo_programacao === selected;
    if (!selectedStillAvailable && !selectedHasLoadedBundle) {
      state.recebimentosSelectedCodigo = "";
    }
    renderRecebimentosProgramacoes();
    if (selectedStillAvailable) {
      await loadRecebimentoBundle(selected);
    } else if (selectedHasLoadedBundle) {
      state.recebimentosSelectedCodigo = selected;
      renderRecebimentosBundle();
    } else {
      clearRecebimentosBundle();
    }
  } catch (error) {
    notify(error.message, true);
  }
}

function renderRecebimentosProgramacoes() {
  const select = el("recebimentosProgramacaoSelect");
  select.innerHTML = '<option value="">Selecione</option>';
  state.recebimentosProgramacoes.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.codigo_programacao;
    option.textContent = `${item.codigo_programacao} - ${item.motorista || "-"} - ${item.status || "-"}`;
    select.appendChild(option);
  });
  if (state.recebimentosSelectedCodigo) {
    ensureRecebimentosProgramacaoOption(state.recebimentosSelectedCodigo, state.recebimentosBundle?.cabecalho || {});
    select.value = state.recebimentosSelectedCodigo;
  }
  el("recebimentosInfo").textContent = `${state.recebimentosProgramacoes.length} planejamento(s) pendente(s) de fechamento.`;
}

function ensureRecebimentosProgramacaoOption(codigo, source = {}) {
  const cleanCodigo = clean(codigo);
  if (!cleanCodigo) return;
  const select = el("recebimentosProgramacaoSelect");
  const existing = Array.from(select.options).find((option) => option.value === cleanCodigo);
  const option = existing || document.createElement("option");
  const motorista = source.motorista_nome || source.motorista || "-";
  const status = source.prestacao_status || source.status || "PENDENTE";
  option.value = cleanCodigo;
  option.textContent = `${cleanCodigo} - ${motorista || "-"} - ${status || "PENDENTE"}`;
  if (!existing) select.appendChild(option);
  select.value = cleanCodigo;
}

async function loadSelectedRecebimentos() {
  const codigo = clean(el("recebimentosProgramacaoSelect").value);
  if (!codigo) {
    clearRecebimentosBundle();
    return;
  }
  await loadRecebimentoBundle(codigo);
}

async function loadRecebimentoBundle(codigo) {
  try {
    state.recebimentosSelectedCodigo = clean(codigo);
    state.recebimentosBundle = await apiRequest(`/recebimentos/${encodeURIComponent(state.recebimentosSelectedCodigo)}/bundle`);
    renderRecebimentosBundle();
  } catch (error) {
    notify(error.message, true);
  }
}

function clearRecebimentosBundle() {
  state.recebimentosBundle = null;
  state.recebimentosSelectedCodigo = "";
  state.recebimentosSelectedCodCliente = "";
  el("recebimentosInfo").textContent = `${state.recebimentosProgramacoes.length} planejamento(s) pendente(s) de fechamento.`;
  el("recebimentosMotorista").textContent = "Motorista: -";
  el("recebimentosVeiculo").textContent = "Veiculo: -";
  el("recebimentosEquipe").textContent = "Equipe: -";
  el("recebimentosRota").textContent = "Rota: -";
  el("recebimentosDiarias").innerHTML = "";
  el("recebimentosRows").innerHTML = "";
  el("recebimentosRows").appendChild(emptyRow(6, "Carregue um planejamento."));
  el("recebimentosTotal").textContent = "TOTAL RECEBIDO: R$ 0,00";
  renderRecebimentoClienteOptions([]);
  clearRecebimentoForm();
  setRecebimentosLocked(false);
}

function renderRecebimentosBundle() {
  const bundle = state.recebimentosBundle;
  if (!bundle) {
    clearRecebimentosBundle();
    return;
  }
  const cab = bundle.cabecalho || {};
  const diarias = bundle.diarias || {};
  const fechada = Boolean(cab.fechada);
  const codigo = cab.codigo_programacao || state.recebimentosSelectedCodigo;
  ensureRecebimentosProgramacaoOption(codigo, cab);
  el("recebimentosProgramacaoSelect").value = codigo;
  el("recebimentosInfo").textContent = `${cab.codigo_programacao || "-"} | ${cab.status || "-"} | Fechamento: ${cab.prestacao_status || "PENDENTE"}`;
  el("recebimentosMotorista").textContent = `Motorista: ${cab.motorista_nome || cab.motorista || "-"}`;
  el("recebimentosVeiculo").textContent = `Veiculo: ${cab.veiculo || "-"}`;
  el("recebimentosEquipe").textContent = `Equipe: ${cab.equipe_nomes || cab.equipe || "-"}`;
  el("recebimentosRota").textContent = `Rota: ${cab.rota || "-"}`;
  el("recebimentosDiariaMotorista").value = formatMoneyInput(cab.diaria_motorista_valor || 0);
  el("recebimentosDataSaida").value = cab.data_saida || "";
  el("recebimentosHoraSaida").value = cab.hora_saida || "";
  el("recebimentosDataChegada").value = cab.data_chegada || "";
  el("recebimentosHoraChegada").value = cab.hora_chegada || "";

  const cards = [
    ["Valor diaria", formatCurrencyBR(diarias.diaria_motorista || 0)],
    ["Qtd diarias", formatPlainNumber(diarias.qtd_diarias || 0)],
    ["Motorista", formatCurrencyBR(diarias.total_motorista || 0)],
    ["Equipe", formatCurrencyBR(diarias.total_ajudantes || 0)],
    ["Total diarias", formatCurrencyBR(diarias.total_geral || 0)],
    ["Ajudantes", diarias.qtd_ajudantes || 0],
  ];
  el("recebimentosDiarias").innerHTML = cards.map(([label, value]) => `
    <article class="metric recebimentos-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `).join("");

  const rows = el("recebimentosRows");
  rows.innerHTML = "";
  const clientes = bundle.clientes || [];
  renderRecebimentoClienteOptions(clientes);
  if (!clientes.length) {
    rows.appendChild(emptyRow(6, "Sem clientes neste planejamento."));
  } else {
    clientes.forEach((item) => {
      const tr = document.createElement("tr");
      tr.className = state.recebimentosSelectedCodCliente === item.cod_cliente ? "selected-row" : "";
      tr.dataset.codCliente = item.cod_cliente || "";
      tr.innerHTML = `
        <td>${escapeHtml(item.cod_cliente || "")}</td>
        <td>${escapeHtml(item.nome_cliente || "")}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor || 0))}</td>
        <td>${escapeHtml(item.forma_pagamento || "")}</td>
        <td>${escapeHtml(item.observacao || "")}</td>
        <td>${escapeHtml(item.data_registro || "")}</td>
      `;
      tr.addEventListener("click", () => selectRecebimentoCliente(item));
      rows.appendChild(tr);
    });
  }
  el("recebimentosTotal").textContent = `TOTAL RECEBIDO: ${formatCurrencyBR(bundle.total_recebido || 0)}`;
  setRecebimentosLocked(fechada);
}

function setRecebimentosLocked(locked) {
  ["recebimentosCabecalhoForm", "recebimentoForm"].forEach((formId) => {
    el(formId).querySelectorAll("input, select, button").forEach((node) => {
      node.disabled = locked;
    });
  });
}

function recebimentosLocked() {
  return Boolean(state.recebimentosBundle?.cabecalho?.fechada);
}

function renderRecebimentoClienteOptions(clientes) {
  const codigos = el("recebimentoClientesCodigos");
  const nomes = el("recebimentoClientesNomes");
  if (!codigos || !nomes) return;
  codigos.innerHTML = "";
  nomes.innerHTML = "";
  clientes.forEach((cliente) => {
    const cod = clean(cliente.cod_cliente);
    const nome = clean(cliente.nome_cliente);
    if (cod) {
      const option = document.createElement("option");
      option.value = cod;
      option.label = nome || cod;
      codigos.appendChild(option);
    }
    if (nome) {
      const option = document.createElement("option");
      option.value = nome;
      option.label = cod || nome;
      nomes.appendChild(option);
    }
  });
}

function syncRecebimentoClienteFromCodigo() {
  const cod = clean(el("recebimentoCodCliente").value).toUpperCase();
  if (!cod) return;
  const cliente = (state.recebimentosBundle?.clientes || []).find(
    (item) => clean(item.cod_cliente).toUpperCase() === cod,
  );
  if (!cliente) return;
  state.recebimentosSelectedCodCliente = cliente.cod_cliente || "";
  el("recebimentoNomeCliente").value = cliente.nome_cliente || "";
  if (!clean(el("recebimentoValor").value) && cliente.valor) {
    el("recebimentoValor").value = formatMoneyInput(cliente.valor);
  }
  if (cliente.forma_pagamento) el("recebimentoForma").value = cliente.forma_pagamento;
  if (!clean(el("recebimentoObs").value) && cliente.observacao) {
    el("recebimentoObs").value = cliente.observacao;
  }
  renderRecebimentosBundle();
}

function syncRecebimentoClienteFromNome() {
  const nome = clean(el("recebimentoNomeCliente").value).toUpperCase();
  if (!nome) return;
  const cliente = (state.recebimentosBundle?.clientes || []).find(
    (item) => clean(item.nome_cliente).toUpperCase() === nome,
  );
  if (!cliente) return;
  state.recebimentosSelectedCodCliente = cliente.cod_cliente || "";
  el("recebimentoCodCliente").value = cliente.cod_cliente || "";
  syncRecebimentoClienteFromCodigo();
}

function selectRecebimentoCliente(item) {
  state.recebimentosSelectedCodCliente = item.cod_cliente || "";
  el("recebimentoCodCliente").value = item.cod_cliente || "";
  el("recebimentoNomeCliente").value = item.nome_cliente || "";
  el("recebimentoValor").value = item.valor ? formatMoneyInput(item.valor) : "";
  el("recebimentoForma").value = item.forma_pagamento || "DINHEIRO";
  el("recebimentoObs").value = item.observacao || "";
  renderRecebimentosBundle();
}

function clearRecebimentoForm() {
  state.recebimentosSelectedCodCliente = "";
  el("recebimentoCodCliente").value = "";
  el("recebimentoNomeCliente").value = "";
  el("recebimentoValor").value = "";
  el("recebimentoForma").value = "DINHEIRO";
  el("recebimentoObs").value = "";
}

async function onSaveRecebimentosCabecalho(event) {
  event.preventDefault();
  if (recebimentosLocked()) {
    notify("Esta prestacao ja esta FECHADA.", true);
    return;
  }
  const codigo = state.recebimentosSelectedCodigo || clean(el("recebimentosProgramacaoSelect").value);
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  try {
    state.recebimentosBundle = await apiRequest(`/recebimentos/${encodeURIComponent(codigo)}/cabecalho`, {
      method: "PUT",
      body: JSON.stringify({
        data_saida: clean(el("recebimentosDataSaida").value),
        hora_saida: clean(el("recebimentosHoraSaida").value),
        data_chegada: clean(el("recebimentosDataChegada").value),
        hora_chegada: clean(el("recebimentosHoraChegada").value),
        diaria_motorista_valor: parseDecimal(el("recebimentosDiariaMotorista").value),
      }),
    });
    notify("Cabecalho de recebimentos salvo.");
    renderRecebimentosBundle();
    await loadRecebimentosProgramacoes();
  } catch (error) {
    notify(error.message, true);
  }
}

async function onSaveRecebimento(event) {
  event.preventDefault();
  if (recebimentosLocked()) {
    notify("Esta prestacao ja esta FECHADA.", true);
    return;
  }
  const codigo = state.recebimentosSelectedCodigo;
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  const codCliente = clean(el("recebimentoCodCliente").value);
  const nomeCliente = clean(el("recebimentoNomeCliente").value);
  if (!codCliente || !nomeCliente) {
    notify("Informe codigo e nome do cliente.", true);
    return;
  }
  try {
    await apiRequest(`/recebimentos/${encodeURIComponent(codigo)}/recebimentos`, {
      method: "POST",
      body: JSON.stringify({
        cod_cliente: codCliente,
        nome_cliente: nomeCliente,
        valor: parseDecimal(el("recebimentoValor").value),
        forma_pagamento: clean(el("recebimentoForma").value),
        observacao: clean(el("recebimentoObs").value),
      }),
    });
    notify("Recebimento salvo.");
    clearRecebimentoForm();
    await loadRecebimentoBundle(codigo);
    await loadRecebimentosProgramacoes();
  } catch (error) {
    notify(error.message, true);
  }
}

async function onZerarRecebimento() {
  if (recebimentosLocked()) {
    notify("Esta prestacao ja esta FECHADA.", true);
    return;
  }
  const codigo = state.recebimentosSelectedCodigo;
  const cod = clean(el("recebimentoCodCliente").value);
  if (!codigo || !cod) {
    notify("Selecione um cliente.", true);
    return;
  }
  if (!window.confirm(`Zerar recebimentos do cliente ${cod}?`)) return;
  try {
    state.recebimentosBundle = await apiRequest(
      `/recebimentos/${encodeURIComponent(codigo)}/recebimentos/${encodeURIComponent(cod)}`,
      {method: "DELETE"},
    );
    notify("Recebimento zerado.");
    clearRecebimentoForm();
    renderRecebimentosBundle();
    await loadRecebimentosProgramacoes();
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadRecebimentosPdf() {
  const codigo = state.recebimentosSelectedCodigo || clean(el("recebimentosProgramacaoSelect").value);
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  try {
    const result = await apiBlobRequest(`/recebimentos/${encodeURIComponent(codigo)}/pdf`);
    downloadBlob(result.blob, result.filename || `RECEBIMENTOS_${codigo}.pdf`);
  } catch (error) {
    notify(error.message, true);
  }
}

function getRecebimentosCodigoAtual() {
  const cab = state.recebimentosBundle?.cabecalho || {};
  return clean(cab.codigo_programacao || state.recebimentosSelectedCodigo || el("recebimentosProgramacaoSelect").value);
}

function ensureDespesasProgramacaoOption(codigo, source = {}) {
  const cleanCodigo = clean(codigo);
  if (!cleanCodigo) return;
  const select = el("despesasProgramacaoSelect");
  const existing = Array.from(select.options).find((option) => option.value === cleanCodigo);
  const option = existing || document.createElement("option");
  const motorista = source.motorista_nome || source.motorista || "-";
  const status = source.prestacao_status || source.status || "PENDENTE";
  option.value = cleanCodigo;
  option.textContent = `${cleanCodigo} - ${motorista || "-"} - ${status || "PENDENTE"}`;
  if (!existing) select.appendChild(option);
  select.value = cleanCodigo;
}

async function goRecebimentosToDespesas() {
  const codigo = getRecebimentosCodigoAtual();
  if (!codigo) {
    notify("Carregue uma programacao em recebimentos primeiro.", true);
    return;
  }
  const cab = state.recebimentosBundle?.cabecalho || {};
  state.despesasSelectedCodigo = codigo;
  switchView("despesas", {skipLoad: true});
  ensureDespesasProgramacaoOption(codigo, cab);
  el("despesasInfo").textContent = `${codigo} | Carregando custos...`;
  await loadDespesasBundle(codigo);
}

function getDespesasCodigoAtual() {
  const cab = state.despesasBundle?.cabecalho || {};
  return clean(cab.codigo_programacao || state.despesasSelectedCodigo || el("despesasProgramacaoSelect").value);
}

async function goDespesasToRecebimentos() {
  const codigo = getDespesasCodigoAtual();
  if (!codigo) {
    notify("Carregue uma programacao em custos primeiro.", true);
    return;
  }
  const cab = state.despesasBundle?.cabecalho || {};
  state.recebimentosSelectedCodigo = codigo;
  switchView("recebimentos", {skipLoad: true});
  ensureRecebimentosProgramacaoOption(codigo, cab);
  el("recebimentosInfo").textContent = `${codigo} | Carregando recebimentos...`;
  await loadRecebimentoBundle(codigo);
}

async function loadDespesasProgramacoes() {
  try {
    const selected = state.despesasSelectedCodigo;
    state.despesasProgramacoes = await apiRequest("/despesas/programacoes");
    const selectedStillAvailable = selected && state.despesasProgramacoes.some((item) => item.codigo_programacao === selected);
    const selectedHasLoadedBundle = selected && state.despesasBundle?.cabecalho?.codigo_programacao === selected;
    if (!selectedStillAvailable && !selectedHasLoadedBundle) {
      state.despesasSelectedCodigo = "";
    }
    renderDespesasProgramacoes();
    if (selectedStillAvailable) {
      await loadDespesasBundle(selected);
    } else if (selectedHasLoadedBundle) {
      state.despesasSelectedCodigo = selected;
      renderDespesasBundle();
    } else {
      clearDespesasBundle();
    }
  } catch (error) {
    notify(error.message, true);
  }
}

function renderDespesasProgramacoes() {
  const select = el("despesasProgramacaoSelect");
  select.innerHTML = '<option value="">Selecione</option>';
  state.despesasProgramacoes.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.codigo_programacao;
    option.textContent = `${item.codigo_programacao} - ${item.motorista || "-"} - ${item.prestacao_status || "PENDENTE"}`;
    select.appendChild(option);
  });
  if (state.despesasSelectedCodigo) {
    ensureDespesasProgramacaoOption(state.despesasSelectedCodigo, state.despesasBundle?.cabecalho || {});
    select.value = state.despesasSelectedCodigo;
  }
  el("despesasInfo").textContent = `${state.despesasProgramacoes.length} planejamento(s) pendente(s) de fechamento.`;
}

async function loadSelectedDespesas() {
  const codigo = clean(el("despesasProgramacaoSelect").value);
  if (!codigo) {
    clearDespesasBundle();
    return;
  }
  await loadDespesasBundle(codigo);
}

async function loadDespesasBundle(codigo) {
  try {
    state.despesasSelectedCodigo = clean(codigo);
    state.despesasBundle = await apiRequest(`/despesas/${encodeURIComponent(state.despesasSelectedCodigo)}/bundle`);
    renderDespesasBundle();
  } catch (error) {
    notify(error.message, true);
  }
}

function clearDespesasBundle() {
  state.despesasBundle = null;
  state.despesasSelectedCodigo = "";
  state.despesasSelectedId = null;
  el("despesasInfo").textContent = `${state.despesasProgramacoes.length} planejamento(s) pendente(s) de fechamento.`;
  el("despesasMotorista").textContent = "Motorista: -";
  el("despesasVeiculo").textContent = "Veiculo: -";
  el("despesasEquipe").textContent = "Equipe: -";
  el("despesasRota").textContent = "Rota: -";
  el("despesasMetrics").innerHTML = "";
  el("despesasFinanceiroResumo").innerHTML = "";
  el("despesasNfResumo").innerHTML = "";
  el("despesasRows").innerHTML = "";
  el("despesasRows").appendChild(emptyRow(6, "Carregue um planejamento."));
  el("despesasAppResumo").innerHTML = "";
  el("despesasAppEntregasRows").innerHTML = "";
  el("despesasAppEntregasRows").appendChild(emptyRow(12, "Carregue um planejamento."));
  el("despesasAppFotosRows").innerHTML = "";
  el("despesasAppFotosRows").appendChild(emptyRow(5, "Carregue um planejamento."));
  el("despesasAppDespesasRows").innerHTML = "";
  el("despesasAppDespesasRows").appendChild(emptyRow(6, "Carregue um planejamento."));
  el("despesasAppAjudantesRows").innerHTML = '<p class="empty-state">Carregue um planejamento.</p>';
  el("despesasAppMotoristaButton").disabled = true;
  el("despesasTotalLabel").textContent = "Custos: 0 | Total: R$ 0,00";
  fillDespesasForms({});
  clearDespesaForm({render: false});
  setDespesasLocked(false);
  el("despesasFinalizarButton").disabled = true;
}

function renderDespesasBundle() {
  const bundle = state.despesasBundle;
  if (!bundle) {
    clearDespesasBundle();
    return;
  }
  const cab = bundle.cabecalho || {};
  const financeiro = bundle.financeiro || {};
  const fechada = Boolean(cab.fechada);
  const codigo = cab.codigo_programacao || state.despesasSelectedCodigo;
  ensureDespesasProgramacaoOption(codigo, cab);
  el("despesasProgramacaoSelect").value = codigo;
  el("despesasInfo").textContent = `${cab.codigo_programacao || "-"} | ${cab.status || "-"} | Prestacao: ${cab.prestacao_status || "PENDENTE"}`;
  el("despesasMotorista").textContent = `Motorista: ${cab.motorista || "-"}`;
  el("despesasVeiculo").textContent = `Veiculo: ${cab.veiculo || "-"}`;
  el("despesasEquipe").textContent = `Equipe: ${cab.equipe || "-"}`;
  el("despesasRota").textContent = `Rota: ${cab.rota || "-"}`;

  renderDespesasMetrics(financeiro, bundle.operacional || {});
  renderDespesasTable();
  renderDespesasAppMotorista();
  fillDespesasForms(bundle);
  renderDespesasFinanceiroResumo(financeiro);
  renderDespesasNfResumo(bundle.nf || {});
  setDespesasLocked(fechada);
  el("despesasFinalizarButton").disabled = fechada;
  el("despesasAppMotoristaButton").disabled = false;
}

function renderDespesasMetrics(financeiro, operacional = {}) {
  const config = currentLogisticaConfig();
  const unidade = config.unidade_padrao || "KG";
  const perda = config.perda_label || "Ocorrencias operacionais";
  const metrics = [
    ["Recebimentos", formatCurrencyBR(financeiro.total_recebido || 0)],
    ["Despesas", formatCurrencyBR(financeiro.total_despesas || 0)],
    ["Entregues/Cancel.", `${operacional.pedidos_entregues || 0}/${operacional.pedidos_cancelados || 0}`],
    [`${unidade} entregue`, formatPlainNumber(operacional.kg_entregue || 0)],
    [perda, `${operacional.mortalidade_total_aves || 0} unid. / ${formatPlainNumber(operacional.mortalidade_total_kg || 0)} ${unidade}`],
    ["Saldo carga", formatPlainNumber(operacional.kg_saldo || 0)],
    ["Recebido app", formatCurrencyBR(operacional.valor_recebido_app || 0)],
    ["Resultado liquido", formatCurrencyBR(financeiro.resultado_liquido || 0)],
  ];
  el("despesasMetrics").innerHTML = metrics.map(([label, value]) => `
    <article class="metric mortalidade-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `).join("");
}

function formatCoordPair(lat, lon) {
  if (lat === null || lat === undefined || lat === "" || lon === null || lon === undefined || lon === "") return "-";
  const latNum = Number(lat);
  const lonNum = Number(lon);
  if (!Number.isFinite(latNum) || !Number.isFinite(lonNum)) return "-";
  return `${latNum.toFixed(6)}, ${lonNum.toFixed(6)}`;
}

function openDespesasAppMotoristaDialog() {
  if (!state.despesasBundle) {
    notify("Carregue um planejamento antes de abrir os dados do app.", true);
    return;
  }
  renderDespesasAppMotorista();
  const dialog = el("despesasAppMotoristaDialog");
  if (!dialog.open) dialog.showModal();
}

function closeDespesasAppMotoristaDialog() {
  const dialog = el("despesasAppMotoristaDialog");
  if (dialog.open) dialog.close();
}

function renderDespesasAppMotorista() {
  const config = currentLogisticaConfig();
  const unidade = config.unidade_padrao || "KG";
  const embalagem = config.embalagem_label || "CX";
  const bundle = state.despesasBundle || {};
  const entregas = Array.isArray(bundle.entregas) ? bundle.entregas : [];
  const fotos = Array.isArray(bundle.fotos) ? bundle.fotos : [];
  const operacional = bundle.operacional || {};
  const despesasApp = ((bundle.despesas || []).filter((item) => {
    const origem = clean(item.origem).toUpperCase();
    return origem.includes("APP") || clean(item.id_local) || clean(item.foto && item.foto.id_foto);
  }));
  const ajudantes = Array.isArray(bundle.ajudantes_historico) ? bundle.ajudantes_historico : [];
  const entregasComGps = entregas.filter((item) => formatCoordPair(item.latitude, item.longitude) !== "-").length;
  el("despesasAppResumo").innerHTML = [
    ["Entregas sincronizadas", operacional.pedidos_entregues || entregas.length],
    ["Entregas com GPS", operacional.gps_entregas || entregasComGps],
    ["Alterados/cancelados", `${operacional.pedidos_alterados || 0}/${operacional.pedidos_cancelados || 0}`],
    [`${embalagem} carreg./entreg.`, `${operacional.caixas_carregadas || 0}/${operacional.caixas_entregues || 0}`],
    [`${unidade} carreg./entreg.`, `${formatPlainNumber(operacional.kg_carregado || 0)}/${formatPlainNumber(operacional.kg_entregue || 0)}`],
    ["Media carreg./entreg.", `${formatPlainNumber(operacional.media_carregada || 0)}/${formatPlainNumber(operacional.media_entregue || 0)}`],
    ["Recebido app/manual", `${formatCurrencyBR(operacional.valor_recebido_app || 0)} / ${formatCurrencyBR(operacional.valor_recebido_manual || 0)}`],
    ["Fotos recebidas", fotos.length],
    ["Despesas do app", despesasApp.length],
    ["Alteracoes de ajudantes", ajudantes.length],
  ].map(([label, value]) => `
    <article class="metric despesas-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `).join("");

  const entregasRows = el("despesasAppEntregasRows");
  entregasRows.innerHTML = "";
  if (!entregas.length) {
    entregasRows.appendChild(emptyRow(12, "Sem entregas sincronizadas do app."));
  } else {
    const nf = bundle.nf || {};
    const avesCx = Number(nf.nf_caixa_final || 0) || 6;
    const mediaCarregada = Number(operacional.media_carregada || nf.nf_media_carregada || 0) || 0;
    entregas.forEach((item) => {
      const tr = document.createElement("tr");
      const cliente = item.nome_cliente || item.cliente_nome || item.cod_cliente || "";
      const caixas = Number(item.caixas_atual || 0) || 0;
      const mediaEntregue = Number(item.media_aplicada || operacional.media_entregue || 0) || 0;
      const kgCarregado = mediaCarregada > 0 && caixas > 0 ? caixas * avesCx * mediaCarregada : 0;
      const kgEntregue = Number(item.peso_previsto || 0) || (mediaEntregue > 0 && caixas > 0 ? caixas * avesCx * mediaEntregue : 0);
      tr.innerHTML = `
        <td>${escapeHtml(cliente)}</td>
        <td>${escapeHtml(item.pedido || "")}</td>
        <td>${escapeHtml(item.status_pedido || "")}</td>
        <td>${escapeHtml(caixas)}</td>
        <td>${escapeHtml(formatPlainNumber(mediaCarregada))}</td>
        <td>${escapeHtml(formatPlainNumber(mediaEntregue))}</td>
        <td>${escapeHtml(formatPlainNumber(kgCarregado))}</td>
        <td>${escapeHtml(formatPlainNumber(kgEntregue))}</td>
        <td>${escapeHtml(item.mortalidade_aves || 0)}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor_recebido || 0))}</td>
        <td>${escapeHtml(formatCoordPair(item.latitude, item.longitude))}</td>
        <td>${escapeHtml(item.timestamp_entrega || item.updated_at || "")}</td>
      `;
      entregasRows.appendChild(tr);
    });
  }

  const fotosRows = el("despesasAppFotosRows");
  fotosRows.innerHTML = "";
  if (!fotos.length) {
    fotosRows.appendChild(emptyRow(5, "Sem fotos recebidas para esta rota."));
  } else {
    fotos.forEach((item) => {
      const tr = document.createElement("tr");
      const vinculo = item.cliente_nome || item.cod_cliente || item.id_vinculo || item.pedido || "";
      const arquivo = item.storage_path || item.path_local || item.arquivo_nome || "";
      tr.innerHTML = `
        <td>${escapeHtml(item.categoria || "")}</td>
        <td>${escapeHtml(item.tipo_registro || "")}</td>
        <td>${escapeHtml(vinculo)}</td>
        <td class="path-cell" title="${escapeHtml(arquivo)}">${escapeHtml(arquivo)}</td>
        <td>${escapeHtml(item.registrado_em || "")}</td>
      `;
      fotosRows.appendChild(tr);
    });
  }

  const appRows = el("despesasAppDespesasRows");
  appRows.innerHTML = "";
  if (!despesasApp.length) {
    appRows.appendChild(emptyRow(6, "Sem despesas sincronizadas do app."));
  } else {
    despesasApp.forEach((item) => {
      const tr = document.createElement("tr");
      const comprovante = item.comprovante_path || (item.foto && (item.foto.storage_path || item.foto.path_local || item.foto.arquivo_nome)) || "";
      tr.innerHTML = `
        <td>${escapeHtml(item.categoria || item.descricao || "")}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor || 0))}</td>
        <td>${escapeHtml(item.estabelecimento || "")}</td>
        <td>${escapeHtml([item.combustivel, item.litros ? `${formatPlainNumber(item.litros)} L` : ""].filter(Boolean).join(" / "))}</td>
        <td>${escapeHtml(formatCoordPair(item.lat, item.lon))}</td>
        <td class="path-cell" title="${escapeHtml(comprovante)}">${escapeHtml(comprovante || "-")}</td>
      `;
      appRows.appendChild(tr);
    });
  }

  const ajudWrap = el("despesasAppAjudantesRows");
  if (!ajudantes.length) {
    ajudWrap.innerHTML = '<p class="empty-state">Sem alteracoes de ajudantes registradas pelo app.</p>';
  } else {
    ajudWrap.innerHTML = ajudantes.map((item) => `
      <article class="app-driver-event">
        <strong>${escapeHtml(item.alterado_em || "-")}</strong>
        <span>${escapeHtml(item.anteriores || "-")} -> ${escapeHtml(item.novos || "-")}</span>
        <small>${escapeHtml(item.motivo || item.origem || "")}</small>
      </article>
    `).join("");
  }
}

async function loadMortalidade() {
  try {
    state.mortalidade = await apiRequest(`/despesas/mortalidade/fotos?${mortalidadeFilterParams().toString()}`);
    renderMortalidade();
  } catch (error) {
    notify(error.message, true);
  }
}

function mortalidadeFilterParams() {
  const params = new URLSearchParams();
  const periodo = clean(el("mortalidadePeriodo").value) || "TODAS";
  const dataInicio = clean(el("mortalidadeDataInicio").value);
  const dataFim = clean(el("mortalidadeDataFim").value);
  const programacao = clean(el("mortalidadeProgramacao").value);
  const motorista = clean(el("mortalidadeMotorista").value);
  const nf = clean(el("mortalidadeNf").value);
  const tipo = clean(el("mortalidadeTipo").value) || "TODOS";
  const busca = clean(el("mortalidadeBusca").value);
  const limit = Math.max(1, Math.min(Number(el("mortalidadeLimit").value || 500), 1000));
  params.set("periodo", periodo);
  params.set("limit", String(limit));
  if (dataInicio) params.set("data_inicio", dataInicio);
  if (dataFim) params.set("data_fim", dataFim);
  if (programacao) params.set("codigo_programacao", programacao);
  if (motorista) params.set("motorista", motorista);
  if (nf) params.set("nf", nf);
  if (tipo && tipo !== "TODOS") params.set("escopo", tipo);
  if (busca) params.set("busca", busca);
  return params;
}

function clearMortalidadeFilters() {
  el("mortalidadeFilterForm").reset();
  el("mortalidadePeriodo").value = "30";
  el("mortalidadeTipo").value = "TODOS";
  el("mortalidadeLimit").value = "500";
  loadMortalidade();
}

function openMortalidadeManualDialog() {
  const dialog = el("mortalidadeManualDialog");
  const programacaoFiltro = clean(el("mortalidadeProgramacao").value);
  const nfFiltro = clean(el("mortalidadeNf").value);
  if (programacaoFiltro && !clean(el("manualMortProgramacao").value)) {
    el("manualMortProgramacao").value = programacaoFiltro;
  }
  if (nfFiltro && !clean(el("manualMortNf").value)) {
    el("manualMortNf").value = nfFiltro;
  }
  updateMortalidadeManualPreview();
  if (!dialog.open) dialog.showModal();
  setTimeout(() => el("manualMortProgramacao").focus(), 0);
}

function closeMortalidadeManualDialog() {
  const dialog = el("mortalidadeManualDialog");
  if (dialog.open) dialog.close();
}

function closeMortalidadeDoaDetailDialog() {
  const dialog = el("mortalidadeDoaDetailDialog");
  if (dialog.open) dialog.close();
  if (state.mortalidadeDoaPhotoUrl) {
    URL.revokeObjectURL(state.mortalidadeDoaPhotoUrl);
    state.mortalidadeDoaPhotoUrl = "";
  }
}

function updateMortalidadeManualPreview() {
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  const aves = Math.max(parseInt(clean(el("manualMortAves").value) || "0", 10) || 0, 0);
  const media = Math.max(parseDecimal(el("manualMortMedia").value), 0);
  const preco = Math.max(parseDecimal(el("manualMortPreco").value), 0);
  const kg = aves * media;
  const valor = kg * preco;
  el("mortalidadeManualPreview").textContent = `${unidade} ${formatMoneyInput(kg)} | Desconto ${formatCurrencyBR(valor)}`;
}

function clearMortalidadeManualForm() {
  el("mortalidadeManualForm").reset();
  updateMortalidadeManualPreview();
}

async function onSaveMortalidadeManual(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const body = new FormData(form);
  const perda = currentLogisticaConfig().perda_label || "Ocorrencias operacionais";
  const aves = Math.max(parseInt(clean(body.get("mortalidade_aves")) || "0", 10) || 0, 0);
  if (!clean(body.get("codigo_programacao")) || !clean(body.get("cliente")) || !clean(body.get("motivo")) || aves <= 0) {
    notify(`Informe programacao, cliente, unidades de ${perda.toLowerCase()} e motivo.`, true);
    return;
  }
  body.set("preco_cliente", String(parseDecimal(body.get("preco_cliente"))));
  body.set("media", String(parseDecimal(body.get("media"))));
  body.set("mortalidade_aves", String(aves));
  try {
    el("mortalidadeManualSaveButton").disabled = true;
    const saved = await apiRequest("/despesas/mortalidade/manual", {method: "POST", body});
    notify(`${perda} registrada: ${saved.mortalidade_aves || aves} unidade(s).`);
    const codigo = clean(body.get("codigo_programacao"));
    if (codigo) el("mortalidadeProgramacao").value = codigo;
    clearMortalidadeManualForm();
    closeMortalidadeManualDialog();
    await loadMortalidade();
  } catch (error) {
    notify(error.message, true);
  } finally {
    el("mortalidadeManualSaveButton").disabled = false;
  }
}

function renderMortalidade() {
  const config = currentLogisticaConfig();
  const perda = config.perda_label || "Ocorrencias operacionais";
  const unidade = config.unidade_padrao || "KG";
  const data = state.mortalidade || {};
  const kpis = data.kpis || {};
  const fotos = Array.isArray(data.fotos) ? data.fotos : [];
  const resumo = Array.isArray(data.por_programacao) ? data.por_programacao : [];
  const doa = fotos.filter((item) => clean(item.escopo).toUpperCase() !== "CLIENTE");
  const cliente = fotos.filter((item) => clean(item.escopo).toUpperCase() === "CLIENTE");
  const filtros = data.filtros || {};
  const periodoTxt = filtros.data_inicio || filtros.data_fim
    ? `${filtros.data_inicio || "..."} a ${filtros.data_fim || "..."}`
    : (filtros.periodo && filtros.periodo !== "TODAS" ? `${filtros.periodo} dia(s)` : "todas as datas");
  el("mortalidadeInfo").textContent = `${kpis.registros || fotos.length} registro(s) de ${perda.toLowerCase()} | Periodo: ${periodoTxt}.`;
  el("mortalidadeKpis").innerHTML = [
    ["Registros", kpis.registros || fotos.length],
    ["Programacoes", kpis.programacoes || 0],
    [`${perda} cliente`, kpis.mortalidade_cliente_aves || 0],
    [`${unidade} cliente`, formatPlainNumber(kpis.mortalidade_cliente_kg || 0)],
    ["Operacao unid.", kpis.mortalidade_doa_aves || 0],
    [`Operacao ${unidade}`, formatPlainNumber(kpis.mortalidade_doa_kg || 0)],
    [`${unidade} afetado`, formatPlainNumber(kpis.kg_afetado || 0)],
    ["Valor afetado", formatCurrencyBR(kpis.valor_afetado || 0)],
    ["Fotos", kpis.fotos || 0],
  ].map(([label, value]) => `
    <article class="metric despesas-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `).join("");

  renderMortalidadeVisuals({kpis, resumo, cliente, doa});

  const resumoRows = el("mortalidadeResumoRows");
  resumoRows.innerHTML = "";
  if (!resumo.length) {
    resumoRows.appendChild(emptyRow(9, `Sem registros de ${perda.toLowerCase()}.`));
  } else {
    resumo.forEach((item) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(item.codigo_programacao || "")}</td>
        <td>${escapeHtml(item.num_nf || "")}</td>
        <td>${escapeHtml(item.motorista || "")}</td>
        <td>${escapeHtml(item.mortalidade_cliente_aves || 0)}</td>
        <td>${escapeHtml(formatPlainNumber(item.mortalidade_cliente_kg || 0))}</td>
        <td>${escapeHtml(item.mortalidade_doa_aves || 0)}</td>
        <td>${escapeHtml(formatPlainNumber(item.mortalidade_doa_kg || 0))}</td>
        <td>${escapeHtml(formatPlainNumber(item.kg_afetado || 0))}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor_afetado || 0))}</td>
        <td>${escapeHtml(item.fotos || 0)}</td>
      `;
      resumoRows.appendChild(tr);
    });
  }

  const doaRows = el("mortalidadeDoaRows");
  doaRows.innerHTML = "";
  state.mortalidadeDoaDetails = doa;
  if (!doa.length) {
    doaRows.appendChild(emptyRow(12, "Sem registros de ocorrencias operacionais."));
  } else {
    doa.forEach((item, index) => {
      const arquivo = item.foto || item.storage_path || item.path_local || item.arquivo_nome || "";
      const tr = document.createElement("tr");
      tr.className = "mortalidade-clickable-row";
      tr.dataset.mortalidadeDoaIndex = String(index);
      tr.title = "Clique para abrir os detalhes da operacao/transbordo.";
      tr.innerHTML = `
        <td>${escapeHtml(item.codigo_programacao || "")}</td>
        <td>${escapeHtml(item.num_nf || "")}</td>
        <td>${escapeHtml(item.motorista || item.motorista_nome || "")}</td>
        <td>${escapeHtml(item.data || "")}</td>
        <td>${escapeHtml(item.hora || "")}</td>
        <td>${escapeHtml(item.mortalidade_doa_aves || 0)}</td>
        <td>${escapeHtml(formatPlainNumber(item.media_carregamento || 0))}</td>
        <td title="${escapeHtml(item.fonte_kg || "")}">${escapeHtml(formatPlainNumber(item.kg_afeta_carga || item.mortalidade_transbordo_kg || 0))}</td>
        <td title="${escapeHtml(item.fonte_preco || "")}">${escapeHtml(formatCurrencyBR(item.preco_compra || 0))}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor_afetado || 0))}</td>
        <td title="${escapeHtml((item.alertas || []).join(", "))}">${escapeHtml(item.obs || "")}</td>
        <td class="path-cell" title="${escapeHtml(arquivo)}">${escapeHtml(arquivo || "-")}</td>
      `;
      tr.addEventListener("click", () => openMortalidadeDoaDetail(index));
      doaRows.appendChild(tr);
    });
  }

  const clienteRows = el("mortalidadeClienteRows");
  clienteRows.innerHTML = "";
  if (!cliente.length) {
    clienteRows.appendChild(emptyRow(17, `Sem registros de ${perda.toLowerCase()} por cliente.`));
  } else {
    cliente.forEach((item) => {
      const arquivo = item.foto || item.storage_path || item.path_local || item.arquivo_nome || "";
      const clienteNome = item.nome_cliente || item.cliente_nome || item.cod_cliente || "";
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(item.codigo_programacao || "")}</td>
        <td>${escapeHtml(item.num_nf || "")}</td>
        <td>${escapeHtml(item.motorista || item.motorista_nome || "")}</td>
        <td>${escapeHtml(item.data || "")}</td>
        <td>${escapeHtml(item.hora || "")}</td>
        <td>${escapeHtml(clienteNome)}</td>
        <td>${escapeHtml(item.pedido || "")}</td>
        <td>${escapeHtml(item.vendedor || "")}</td>
        <td>${escapeHtml(item.mortalidade_cliente_aves || 0)}</td>
        <td>${escapeHtml(formatPlainNumber(item.media_cliente || item.media_carregamento || 0))}</td>
        <td title="${escapeHtml(item.fonte_kg || "")}">${escapeHtml(formatPlainNumber(item.kg_afeta_carga || item.mortalidade_cliente_kg || 0))}</td>
        <td title="${escapeHtml(item.fonte_preco || "")}">${escapeHtml(formatCurrencyBR(item.preco_cliente || item.preco_compra || 0))}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor_afetado || 0))}</td>
        <td>${escapeHtml(item.data_pedido || "")}</td>
        <td title="${escapeHtml((item.alertas || []).join(", "))}">${escapeHtml(item.motivo || "")}</td>
        <td>${escapeHtml(formatCoordPair(item.latitude, item.longitude))}</td>
        <td class="path-cell" title="${escapeHtml(arquivo)}">${escapeHtml(arquivo || "-")}</td>
      `;
      clienteRows.appendChild(tr);
    });
  }
}

function openMortalidadeDoaDetail(index) {
  const item = (state.mortalidadeDoaDetails || [])[Number(index)];
  if (!item) return;
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  const arquivo = item.foto || item.storage_path || item.path_local || item.arquivo_nome || "";
  const codigo = clean(item.codigo_programacao);
  const vinculados = ((state.mortalidade || {}).fotos || [])
    .filter((row) => clean(row.escopo).toUpperCase() === "CLIENTE" && clean(row.codigo_programacao).toUpperCase() === codigo.toUpperCase());
  const title = codigo || "Operacao / Transbordo";
  el("mortalidadeDoaDetailTitle").textContent = title;
  el("mortalidadeDoaDetailSubtitle").textContent = [
    item.num_nf ? `NF ${item.num_nf}` : "",
    item.motorista || item.motorista_nome || "",
    item.data || "",
  ].filter(Boolean).join(" | ") || "Detalhes da ocorrencia operacional.";
  const infoRows = [
    ["Planejamento", item.codigo_programacao || "-"],
    ["Nota fiscal", item.num_nf || item.nota_fiscal || "-"],
    ["Motorista", item.motorista || item.motorista_nome || "-"],
    ["Veiculo", item.veiculo || "-"],
    ["Data / hora", [item.data, item.hora].filter(Boolean).join(" ") || "-"],
    ["Rota / local", item.rota || item.local_rota || item.local || "-"],
    ["Local carregamento", item.local_carregamento || "-"],
    ["Unidades afetadas", formatPlainNumber(item.mortalidade_doa_aves || 0)],
    [`${unidade} afetado`, formatPlainNumber(item.kg_afeta_carga || item.mortalidade_doa_kg || 0)],
    ["Media", formatPlainNumber(item.media_carregamento || 0)],
    ["Preco compra", formatCurrencyBR(item.preco_compra || 0)],
    ["Valor afetado", formatCurrencyBR(item.valor_afetado || 0)],
    ["Fonte do KG", item.fonte_kg || "-"],
    ["Fonte do preco", item.fonte_preco || "-"],
    ["Observacao", item.obs || "-"],
    ["Foto", arquivo || "-"],
  ];
  const alertas = Array.isArray(item.alertas) ? item.alertas : [];
  el("mortalidadeDoaDetailBody").innerHTML = `
    <section class="mortalidade-detail-photo">
      <img id="mortalidadeDoaDetailPhoto" class="hidden" alt="Foto da ocorrencia operacional" loading="lazy">
      <div id="mortalidadeDoaDetailPhotoEmpty" class="mortalidade-detail-photo-empty">
        ${escapeHtml(item.id_foto ? "Carregando foto..." : (arquivo ? `Arquivo: ${arquivo}` : "Sem foto vinculada."))}
      </div>
    </section>
    <section class="mortalidade-detail-info">
      ${infoRows.map(([label, value]) => `
        <article>
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
        </article>
      `).join("")}
      ${alertas.length ? `
        <article class="wide">
          <span>Alertas</span>
          <strong>${escapeHtml(alertas.join(" | "))}</strong>
        </article>
      ` : ""}
      <article class="wide">
        <span>Vinculos da programacao</span>
        <strong>${escapeHtml(formatMortalidadeVinculos(vinculados, unidade))}</strong>
      </article>
    </section>
  `;
  const dialog = el("mortalidadeDoaDetailDialog");
  if (!dialog.open) dialog.showModal();
  loadMortalidadeDoaDetailPhoto(item.id_foto, arquivo);
}

function formatMortalidadeVinculos(vinculados, unidade) {
  if (!vinculados.length) return "Sem ocorrencias por cliente vinculadas no filtro atual.";
  return vinculados.slice(0, 8).map((row) => {
    const cliente = row.nome_cliente || row.cliente_nome || row.cod_cliente || "Cliente";
    const pedido = row.pedido ? `pedido ${row.pedido}` : "sem pedido";
    const kg = formatPlainNumber(row.kg_afeta_carga || row.mortalidade_cliente_kg || 0);
    const valor = formatCurrencyBR(row.valor_afetado || 0);
    return `${cliente} (${pedido}) - ${kg} ${unidade} - ${valor}`;
  }).join(" | ") + (vinculados.length > 8 ? ` | +${vinculados.length - 8} vinculo(s)` : "");
}

async function loadMortalidadeDoaDetailPhoto(idFoto, arquivo) {
  const img = el("mortalidadeDoaDetailPhoto");
  const empty = el("mortalidadeDoaDetailPhotoEmpty");
  if (state.mortalidadeDoaPhotoUrl) {
    URL.revokeObjectURL(state.mortalidadeDoaPhotoUrl);
    state.mortalidadeDoaPhotoUrl = "";
  }
  if (!idFoto) return;
  try {
    const {blob} = await apiBlobRequest(`/despesas/mortalidade/fotos/${encodeURIComponent(idFoto)}/arquivo`);
    const url = URL.createObjectURL(blob);
    state.mortalidadeDoaPhotoUrl = url;
    img.src = url;
    img.classList.remove("hidden");
    empty.classList.add("hidden");
  } catch (error) {
    img.classList.add("hidden");
    empty.classList.remove("hidden");
    empty.textContent = arquivo ? `Arquivo: ${arquivo}` : (error.message || "Foto nao encontrada no servidor.");
  }
}

function renderMortalidadeVisuals({kpis, resumo, cliente, doa}) {
  renderMortalidadeTipoChart(kpis || {});
  renderMortalidadeProgramacaoChart(resumo || []);
  renderMortalidadeInsights({kpis: kpis || {}, resumo: resumo || [], cliente: cliente || [], doa: doa || []});
}

function renderMortalidadeTipoChart(kpis) {
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  const rows = [
    {
      label: "Cliente",
      aves: Number(kpis.mortalidade_cliente_aves || 0),
      kg: Number(kpis.mortalidade_cliente_kg || 0),
      valor: Number(kpis.valor_cliente || 0),
    },
    {
      label: "Operacao / Transbordo",
      aves: Number(kpis.mortalidade_doa_aves || 0),
      kg: Number(kpis.mortalidade_doa_kg || 0),
      valor: Number(kpis.valor_operacao || 0),
    },
  ];
  const wrap = el("mortalidadeTipoChart");
  wrap.innerHTML = "";
  const maxKg = Math.max(...rows.map((item) => Math.abs(item.kg)), 1);
  rows.forEach((item, index) => {
    const pct = Math.max((Math.abs(item.kg) / maxKg) * 100, item.kg > 0 ? 4 : 0);
    const row = document.createElement("div");
    row.className = `mortalidade-chart-row ${index === 0 ? "cliente" : "doa"}`;
    row.innerHTML = `
      <div>
        <strong>${escapeHtml(item.label)}</strong>
        <span>${escapeHtml(`${formatPlainNumber(item.aves)} unid. | ${formatPlainNumber(item.kg)} ${unidade}`)}</span>
      </div>
      <div class="mortalidade-chart-track">
        <div class="mortalidade-chart-bar" style="width: ${pct}%"></div>
      </div>
      <em>${escapeHtml(formatCurrencyBR(item.valor || 0))}</em>
    `;
    wrap.appendChild(row);
  });
}

function renderMortalidadeProgramacaoChart(resumo) {
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  const wrap = el("mortalidadeProgramacaoChart");
  wrap.innerHTML = "";
  const top = [...resumo]
    .sort((a, b) => Number(b.kg_afetado || 0) - Number(a.kg_afetado || 0))
    .slice(0, 6);
  if (!top.length) {
    wrap.innerHTML = '<p class="empty-chart">Sem dados para grafico.</p>';
    return;
  }
  const maxKg = Math.max(...top.map((item) => Number(item.kg_afetado || 0)), 1);
  top.forEach((item) => {
    const kg = Number(item.kg_afetado || 0);
    const pct = Math.max((kg / maxKg) * 100, kg > 0 ? 4 : 0);
    const row = document.createElement("div");
    row.className = "mortalidade-chart-row ranking";
    row.innerHTML = `
      <div>
        <strong>${escapeHtml(item.codigo_programacao || "-")}</strong>
        <span>${escapeHtml(item.motorista || "-")} | NF ${escapeHtml(item.num_nf || "-")}</span>
      </div>
      <div class="mortalidade-chart-track">
        <div class="mortalidade-chart-bar" style="width: ${pct}%"></div>
      </div>
      <em>${escapeHtml(formatPlainNumber(kg))} ${escapeHtml(unidade)}</em>
    `;
    wrap.appendChild(row);
  });
}

function renderMortalidadeInsights({kpis, resumo, cliente, doa}) {
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  const maiorRota = [...resumo].sort((a, b) => Number(b.kg_afetado || 0) - Number(a.kg_afetado || 0))[0] || {};
  const maiorCliente = [...cliente].sort((a, b) => Number(b.kg_afeta_carga || b.mortalidade_cliente_kg || 0) - Number(a.kg_afeta_carga || a.mortalidade_cliente_kg || 0))[0] || {};
  const maiorDoa = [...doa].sort((a, b) => Number(b.kg_afeta_carga || b.mortalidade_doa_kg || 0) - Number(a.kg_afeta_carga || a.mortalidade_doa_kg || 0))[0] || {};
  const totalKg = Number(kpis.kg_afetado || 0);
  const doaKg = Number(kpis.mortalidade_doa_kg || 0);
  const clienteKg = Number(kpis.mortalidade_cliente_kg || 0);
  const doaShare = totalKg > 0 ? (doaKg / totalKg) * 100 : 0;
  const clienteShare = totalKg > 0 ? (clienteKg / totalKg) * 100 : 0;
  const clienteNome = maiorCliente.nome_cliente || maiorCliente.cliente_nome || maiorCliente.cod_cliente || "-";
  el("mortalidadeInsights").innerHTML = [
    ["Maior impacto", maiorRota.codigo_programacao ? `${maiorRota.codigo_programacao} | ${formatPlainNumber(maiorRota.kg_afetado || 0)} ${unidade}` : "-"],
    ["Operacao no total", `${formatPlainNumber(doaShare)}% do ${unidade} afetado`],
    ["Clientes no total", `${formatPlainNumber(clienteShare)}% do ${unidade} afetado`],
    ["Cliente critico", clienteNome],
    ["Maior ocorrencia operacional", maiorDoa.codigo_programacao ? `${maiorDoa.codigo_programacao} | ${formatPlainNumber(maiorDoa.kg_afeta_carga || maiorDoa.mortalidade_doa_kg || 0)} ${unidade}` : "-"],
  ].map(([label, value]) => `
    <div class="mortalidade-insight">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join("");
}

function renderDespesasTable() {
  const rows = el("despesasRows");
  rows.innerHTML = "";
  const despesas = (state.despesasBundle && state.despesasBundle.despesas) || [];
  const total = (state.despesasBundle && state.despesasBundle.financeiro && state.despesasBundle.financeiro.total_despesas) || 0;
  el("despesasTotalLabel").textContent = `Despesas: ${despesas.length} | Total: ${formatCurrencyBR(total)}`;
  if (!despesas.length) {
    rows.appendChild(emptyRow(7, "Sem custos neste planejamento."));
    return;
  }
  despesas.forEach((item) => {
    const tr = document.createElement("tr");
    tr.className = Number(state.despesasSelectedId) === Number(item.id) ? "selected-row" : "";
    tr.innerHTML = `
      <td>${escapeHtml(item.id)}</td>
      <td>${escapeHtml(item.descricao || "")}</td>
      <td>${escapeHtml(formatCurrencyBR(item.valor || 0))}</td>
      <td>${escapeHtml(item.data_registro || "")}</td>
      <td>${escapeHtml(item.tipo_despesa || "OUTRAS")}</td>
      <td>${escapeHtml(item.categoria || "")}</td>
      <td>${escapeHtml(item.observacao || "")}</td>
    `;
    tr.addEventListener("click", () => selectDespesa(item));
    rows.appendChild(tr);
  });
}

function fillDespesasForms(bundle) {
  const rota = bundle.rota || {};
  const nf = bundle.nf || {};
  const financeiro = bundle.financeiro || {};
  const cedulas = financeiro.cedulas || {};

  el("despKmInicial").value = formatMoneyInput(rota.km_inicial || 0);
  el("despKmFinal").value = formatMoneyInput(rota.km_final || 0);
  el("despLitros").value = formatMoneyInput(rota.litros || 0);
  el("despKmRodado").value = formatPlainNumber(rota.km_rodado || 0);
  el("despMediaKmL").value = formatPlainNumber(rota.media_km_l || 0);
  el("despCustoKm").value = formatMoneyInput(rota.custo_km || 0);
  el("despRotaObs").value = rota.rota_observacao || "";

  el("despNfNumero").value = nf.nf_numero || "";
  el("despNfKg").value = formatMoneyInput(nf.nf_kg || 0);
  el("despNfPreco").value = formatMoneyInput(nf.nf_preco || 0);
  el("despNfCaixas").value = nf.nf_caixas || 0;
  el("despNfKgCarregado").value = formatMoneyInput(nf.nf_kg_carregado || 0);
  el("despNfKgVendido").value = formatMoneyInput(nf.nf_kg_vendido || 0);
  el("despNfSaldo").value = formatMoneyInput(nf.nf_saldo || 0);
  el("despNfMedia").value = formatMoneyInput(nf.nf_media_carregada || 0);
  el("despNfCaixaFinal").value = nf.nf_caixa_final || 0;
  el("despMortAves").value = nf.mortalidade_transbordo_aves || 0;
  el("despMortKg").value = formatMoneyInput(nf.mortalidade_transbordo_kg || 0);
  el("despObsTransbordo").value = nf.obs_transbordo || "";

  el("despAdiantamento").value = formatMoneyInput(financeiro.adiantamento || 0);
  el("despPixMotorista").value = formatMoneyInput(financeiro.pix_motorista || 0);
  DESPESAS_CEDULAS.forEach((ced) => {
    el(`despCed${ced}`).value = cedulas[String(ced)] || 0;
  });
  updateDespesasFinanceiroPreview();
}

function renderDespesasFinanceiroResumo(financeiro) {
  const diferenca = Number(financeiro.diferenca || 0);
  const caixaTone = despesasDiferencaTone(diferenca);
  const caixaStatus = caixaTone === "success" ? "OK" : caixaTone === "danger" ? "FALTANDO" : "SOBRANDO";
  const rows = [
    ["Recebimentos total", financeiro.total_recebido, "neutral"],
    ["Adiantamento p/ rota", financeiro.adiantamento, "neutral"],
    ["Total entradas", financeiro.total_entradas, "neutral"],
    ["Despesas total", financeiro.total_despesas, "neutral"],
    ["Contagem cedulas", financeiro.valor_dinheiro, "neutral"],
    ["PIX motorista", financeiro.pix_motorista, "neutral"],
    ["Total saidas", financeiro.total_saidas, "neutral"],
    ["Valor p/ caixa", financeiro.valor_final_caixa, caixaTone],
    ["Total devolvido", financeiro.total_devolvido, caixaTone],
    [`Diferenca (${caixaStatus})`, financeiro.diferenca, caixaTone],
    ["Resultado liquido", financeiro.resultado_liquido, Number(financeiro.resultado_liquido || 0) < 0 ? "danger" : "success"],
  ];
  const totalCedulas = Number(financeiro.valor_dinheiro || 0);
  el("despesasCedulasResumo").textContent = `Total cedulas: ${formatCurrencyBR(totalCedulas)}`;
  el("despesasCedulasResumo").className = `cedulas-live-summary ${caixaTone}`;
  el("despesasFinanceiroResumo").innerHTML = rows.map(([label, value, tone]) => `
    <div class="despesas-summary-item ${escapeHtml(tone || "neutral")}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(formatCurrencyBR(value || 0))}</strong>
    </div>
  `).join("");
}

function despesasDiferencaTone(value) {
  const diferenca = Number(value || 0);
  if (Math.abs(diferenca) < 0.005) return "success";
  return diferenca < 0 ? "danger" : "warning";
}

function despesasCedulasTotalFromForm() {
  return DESPESAS_CEDULAS.reduce((total, ced) => {
    const qtd = Math.max(parseInt(clean(el(`despCed${ced}`).value) || "0", 10) || 0, 0);
    return total + (qtd * Number(ced));
  }, 0);
}

function despesasFinanceiroPreviewFromForm() {
  const base = (state.despesasBundle && state.despesasBundle.financeiro) || {};
  const totalRecebido = Number(base.total_recebido || 0);
  const totalDespesas = Number(base.total_despesas || 0);
  const adiantamento = parseDecimal(el("despAdiantamento").value);
  const pixMotorista = parseDecimal(el("despPixMotorista").value);
  const valorDinheiro = despesasCedulasTotalFromForm();
  const totalEntradas = totalRecebido + adiantamento;
  const totalDevolvido = valorDinheiro + pixMotorista;
  const totalSaidas = totalDespesas + totalDevolvido;
  const valorFinalCaixa = totalEntradas - totalDespesas;
  const diferenca = valorFinalCaixa - totalDevolvido;
  return {
    ...base,
    total_recebido: totalRecebido,
    total_despesas: totalDespesas,
    adiantamento,
    pix_motorista: pixMotorista,
    valor_dinheiro: valorDinheiro,
    total_entradas: totalEntradas,
    total_devolvido: totalDevolvido,
    total_saidas: totalSaidas,
    valor_final_caixa: valorFinalCaixa,
    diferenca,
    resultado_liquido: totalEntradas - totalSaidas,
  };
}

function updateDespesasFinanceiroPreview() {
  renderDespesasFinanceiroResumo(despesasFinanceiroPreviewFromForm());
}

function normalizeDespesasFinanceiroInput(event) {
  const input = event.currentTarget;
  if (!input) return;
  if (input.id === "despAdiantamento" || input.id === "despPixMotorista") {
    input.value = formatMoneyInput(parseDecimal(input.value));
  } else if (input.id && input.id.startsWith("despCed")) {
    input.value = String(Math.max(parseInt(clean(input.value) || "0", 10) || 0, 0));
  }
  updateDespesasFinanceiroPreview();
}

function renderDespesasNfResumo(nf) {
  const cabecalho = state.despesasBundle?.cabecalho || {};
  const isTransbordo = clean(cabecalho.operacao_tipo).toUpperCase() === "TRANSBORDO";
  const modalidade = clean(cabecalho.transbordo_modalidade).replaceAll("_", " ");
  const rows = [
    ...(isTransbordo ? [
      ["Operacao", "TRANSBORDO"],
      ["Modalidade", modalidade || "-"],
    ] : []),
    ["KG NF", formatPlainNumber(nf.nf_kg || 0)],
    ["Preco NF", formatCurrencyBR(nf.nf_preco || 0)],
    ["Total compra", formatCurrencyBR(nf.total_compra || 0)],
    ["Preco medio venda", formatCurrencyBR(nf.preco_medio_venda || 0)],
    ["KG vendido", formatPlainNumber(nf.nf_kg_vendido || 0)],
    ["Receita estimada", formatCurrencyBR(nf.receita_estimada || 0)],
    ["Despesas da rota", formatCurrencyBR(nf.despesas_rota || 0)],
    ["Lucro bruto", formatCurrencyBR(nf.lucro_bruto || 0)],
    ["Lucro liquido", formatCurrencyBR(nf.lucro_liquido || 0)],
    ["Margem liquida", `${formatPlainNumber(nf.margem_liquida || 0)}%`],
    ["KG util NF", formatPlainNumber(nf.kg_nf_util || 0)],
  ];
  el("despesasNfResumo").innerHTML = rows.map(([label, value]) => `
    <div class="despesas-summary-item">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join("");
}

function selectDespesa(item) {
  state.despesasSelectedId = item.id;
  el("despesaId").value = item.id || "";
  el("despesaDescricao").value = item.descricao || "";
  el("despesaValor").value = item.valor ? formatMoneyInput(item.valor) : "";
  el("despesaTipo").value = item.tipo_despesa || "OUTRAS";
  el("despesaCategoria").value = item.categoria || "ROTA";
  el("despesaObservacao").value = item.observacao || "";
  renderDespesasTable();
}

function clearDespesaForm(options = {}) {
  state.despesasSelectedId = null;
  el("despesaId").value = "";
  el("despesaDescricao").value = "";
  el("despesaValor").value = "";
  el("despesaTipo").value = "OUTRAS";
  el("despesaCategoria").value = "ROTA";
  el("despesaObservacao").value = "";
  if (options.render !== false && state.despesasBundle) {
    renderDespesasTable();
  }
}

async function onSaveDespesa(event) {
  event.preventDefault();
  const codigo = state.despesasSelectedCodigo;
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  const despesaId = clean(el("despesaId").value);
  const payload = {
    descricao: clean(el("despesaDescricao").value),
    valor: parseDecimal(el("despesaValor").value),
    categoria: clean(el("despesaCategoria").value),
    tipo_despesa: clean(el("despesaTipo").value) || "OUTRAS",
    observacao: clean(el("despesaObservacao").value),
  };
  try {
    const path = despesaId
      ? `/despesas/${encodeURIComponent(codigo)}/despesas/${encodeURIComponent(despesaId)}`
      : `/despesas/${encodeURIComponent(codigo)}/despesas`;
    await apiRequest(path, {
      method: despesaId ? "PATCH" : "POST",
      body: JSON.stringify(payload),
    });
    notify("Despesa salva.");
    clearDespesaForm({render: false});
    await loadDespesasBundle(codigo);
  } catch (error) {
    notify(error.message, true);
  }
}

async function onDeleteDespesa() {
  const codigo = state.despesasSelectedCodigo;
  const despesaId = clean(el("despesaId").value);
  if (!codigo || !despesaId) {
    notify("Selecione uma despesa.", true);
    return;
  }
  if (!window.confirm(`Excluir despesa ${despesaId}?`)) return;
  try {
    state.despesasBundle = await apiRequest(`/despesas/${encodeURIComponent(codigo)}/despesas/${encodeURIComponent(despesaId)}`, {
      method: "DELETE",
    });
    notify("Despesa excluida.");
    clearDespesaForm({render: false});
    renderDespesasBundle();
  } catch (error) {
    notify(error.message, true);
  }
}

async function onSaveDespesasRota(event) {
  event.preventDefault();
  const codigo = state.despesasSelectedCodigo;
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  try {
    state.despesasBundle = await apiRequest(`/despesas/${encodeURIComponent(codigo)}/rota`, {
      method: "PUT",
      body: JSON.stringify({
        km_inicial: parseDecimal(el("despKmInicial").value),
        km_final: parseDecimal(el("despKmFinal").value),
        litros: parseDecimal(el("despLitros").value),
        rota_observacao: clean(el("despRotaObs").value),
      }),
    });
    notify("Dados da rota salvos.");
    renderDespesasBundle();
  } catch (error) {
    notify(error.message, true);
  }
}

async function onSaveDespesasNf(event) {
  event.preventDefault();
  const codigo = state.despesasSelectedCodigo;
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  try {
    state.despesasBundle = await apiRequest(`/despesas/${encodeURIComponent(codigo)}/nf`, {
      method: "PUT",
      body: JSON.stringify({
        nf_numero: clean(el("despNfNumero").value),
        nf_kg: parseDecimal(el("despNfKg").value),
        nf_preco: parseDecimal(el("despNfPreco").value),
        nf_caixas: parseInt(clean(el("despNfCaixas").value) || "0", 10),
        nf_kg_carregado: parseDecimal(el("despNfKgCarregado").value),
        nf_kg_vendido: parseDecimal(el("despNfKgVendido").value),
        nf_saldo: parseDecimal(el("despNfSaldo").value),
        nf_media_carregada: parseDecimal(el("despNfMedia").value),
        nf_caixa_final: parseInt(clean(el("despNfCaixaFinal").value) || "0", 10),
        mortalidade_transbordo_aves: parseInt(clean(el("despMortAves").value) || "0", 10),
        mortalidade_transbordo_kg: parseDecimal(el("despMortKg").value),
        obs_transbordo: clean(el("despObsTransbordo").value),
      }),
    });
    notify("Nota fiscal salva.");
    renderDespesasBundle();
  } catch (error) {
    notify(error.message, true);
  }
}

async function onSaveDespesasFinanceiro(event) {
  event.preventDefault();
  const codigo = state.despesasSelectedCodigo;
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  try {
    state.despesasBundle = await apiRequest(`/despesas/${encodeURIComponent(codigo)}/financeiro`, {
      method: "PUT",
      body: JSON.stringify({
        adiantamento: parseDecimal(el("despAdiantamento").value),
        pix_motorista: parseDecimal(el("despPixMotorista").value),
        cedulas: despesasCedulasPayload(),
      }),
    });
    notify("Resumo financeiro salvo.");
    renderDespesasBundle();
  } catch (error) {
    notify(error.message, true);
  }
}

async function onFinalizarDespesas() {
  const codigo = state.despesasSelectedCodigo;
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  if (!window.confirm(`Finalizar fechamento operacional da rota ${codigo}?`)) return;
  try {
    state.despesasBundle = await apiRequest(`/despesas/${encodeURIComponent(codigo)}/finalizar`, {method: "POST"});
    notify("Fechamento finalizado.");
    renderDespesasBundle();
    await loadDespesasProgramacoes();
    const result = await apiBlobRequest(`/despesas/${encodeURIComponent(codigo)}/pdf`);
    openPdfPreviewWindow([
      {label: "Fechamento", blob: result.blob, filename: result.filename || `PRESTACAO_${codigo}.pdf`},
    ], `Previsualizacao do fechamento ${codigo}`);
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadDespesasPdf() {
  const codigo = state.despesasSelectedCodigo || clean(el("despesasProgramacaoSelect").value);
  if (!codigo) {
    notify("Carregue um planejamento primeiro.", true);
    return;
  }
  try {
    const result = await apiBlobRequest(`/despesas/${encodeURIComponent(codigo)}/pdf`);
    downloadBlob(result.blob, result.filename || `PRESTACAO_${codigo}.pdf`);
  } catch (error) {
    notify(error.message, true);
  }
}

function despesasCedulasPayload() {
  const payload = {};
  DESPESAS_CEDULAS.forEach((ced) => {
    payload[String(ced)] = Math.max(parseInt(clean(el(`despCed${ced}`).value) || "0", 10) || 0, 0);
  });
  return payload;
}

function updateDespesasRotaPreview() {
  const kmInicial = parseDecimal(el("despKmInicial").value);
  const kmFinal = parseDecimal(el("despKmFinal").value);
  const litros = parseDecimal(el("despLitros").value);
  const kmRodado = Math.max(kmFinal - kmInicial, 0);
  const media = litros > 0 ? kmRodado / litros : 0;
  const totalDespesas = state.despesasBundle && state.despesasBundle.financeiro
    ? Number(state.despesasBundle.financeiro.total_despesas || 0)
    : 0;
  const custoKm = kmRodado > 0 ? totalDespesas / kmRodado : 0;
  el("despKmRodado").value = formatPlainNumber(kmRodado);
  el("despMediaKmL").value = formatPlainNumber(media);
  el("despCustoKm").value = formatMoneyInput(custoKm);
}

function setDespesasLocked(locked) {
  ["despesaForm", "despesasRotaForm", "despesasNfForm", "despesasFinanceiroForm"].forEach((formId) => {
    el(formId).querySelectorAll("input, select, button").forEach((node) => {
      node.disabled = locked;
    });
  });
  el("despesaNovaButton").disabled = locked;
}

async function loadComprasNfe() {
  const natureza = clean(el("comprasNatureza").value) || "TODAS";
  try {
    const [data, estoque] = await Promise.all([
      apiRequest(`/compras/nfe?natureza=${encodeURIComponent(natureza)}`),
      apiRequest("/compras/estoque/resumo"),
    ]);
    state.comprasNfe = {
      naturezas: data.naturezas || [],
      rows: data.rows || [],
      selectedId: state.comprasNfe.selectedId,
      total: data.total || 0,
      total_valor: data.total_valor || 0,
    };
    state.comprasEstoque = estoque || null;
    renderComprasNaturezas(data.naturezas || []);
    renderComprasNfe();
    renderComprasEstoque();
  } catch (error) {
    notify(error.message || "Falha ao carregar notas de compras.", true);
  }
}

function renderComprasNaturezas(naturezas) {
  const select = el("comprasNatureza");
  const current = clean(select.value) || "TODAS";
  const options = ["TODAS", ...(naturezas || [])];
  select.innerHTML = options.map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("");
  select.value = options.includes(current) ? current : "TODAS";
}

function filteredComprasRows() {
  const busca = clean(el("comprasBusca").value).toUpperCase();
  const rows = (state.comprasNfe && state.comprasNfe.rows) || [];
  if (!busca) return rows;
  return rows.filter((item) => [
    item.chave_acesso,
    item.numero,
    item.serie,
    item.fornecedor_documento,
    item.fornecedor_razao,
    item.produto,
    item.natureza_operacao,
  ].some((value) => String(value || "").toUpperCase().includes(busca)));
}

function renderComprasNfe() {
  const rows = el("comprasNfeRows");
  const dataRows = filteredComprasRows();
  rows.innerHTML = "";
  if (!dataRows.length) {
    rows.appendChild(emptyRow(14, "Nenhuma NF-e encontrada. Importe um XML ou carregue o manifestador quando a integracao fiscal estiver ativa."));
  } else {
    dataRows.forEach((item) => {
      const tr = document.createElement("tr");
      tr.className = Number(state.comprasNfe.selectedId) === Number(item.id) ? "selected-row" : "";
      tr.tabIndex = 0;
      tr.addEventListener("click", () => selectComprasNfe(item.id));
      tr.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          selectComprasNfe(item.id);
        }
      });
      tr.innerHTML = `
        <td><span class="compras-status-dot ${item.xml_disponivel ? "ok" : ""}"></span></td>
        <td>${escapeHtml(item.chave_acesso || "-")}</td>
        <td>${escapeHtml(item.serie || "-")}</td>
        <td>${escapeHtml(item.numero || "-")}</td>
        <td>${escapeHtml(item.fornecedor_documento || "-")}</td>
        <td>${escapeHtml(item.fornecedor_razao || "-")}</td>
        <td>${escapeHtml(item.produto || "-")}</td>
        <td>${escapeHtml(formatDate(item.emissao) || item.emissao || "-")}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor_total || 0))}</td>
        <td>${escapeHtml(item.situacao_nfe || "-")}</td>
        <td>${escapeHtml(item.nsu || "-")}</td>
        <td>${escapeHtml(item.estoque_fiscal_status || "PENDENTE")}</td>
        <td>${escapeHtml(item.estoque_fisico_status || "PENDENTE")}</td>
        <td>${escapeHtml(formatPlainNumber(item.estoque_kg_saldo || item.estoque_kg_entrada || 0))}</td>
      `;
      rows.appendChild(tr);
    });
  }
  const totalValor = dataRows.reduce((sum, item) => sum + Number(item.valor_total || 0), 0);
  el("comprasResumo").textContent = `Total de Notas: ${dataRows.length} | Valor: ${formatCurrencyBR(totalValor)}`;
  el("comprasInfo").textContent = `${state.comprasNfe.total || 0} nota(s) carregada(s). XML importado cadastra fornecedor automaticamente quando necessario.`;
  el("comprasDownloadXmlButton").disabled = !state.comprasNfe.selectedId;
}

function selectComprasNfe(id) {
  state.comprasNfe.selectedId = Number(id) || null;
  const item = ((state.comprasNfe && state.comprasNfe.rows) || []).find((row) => Number(row.id) === Number(id));
  if (item) {
    el("comprasEntradaKg").value = item.estoque_kg_entrada ? formatMoneyInput(item.estoque_kg_entrada) : "";
    el("comprasEntradaCaixas").value = item.estoque_caixas_entrada || "";
  }
  renderComprasNfe();
}

function renderComprasEstoque() {
  const data = state.comprasEstoque || {};
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  el("comprasFiscalSaldo").textContent = `${formatPlainNumber(data.fiscal_saldo_kg || 0)} ${unidade}`;
  el("comprasFiscalDetalhe").textContent = `Entradas ${formatPlainNumber(data.fiscal_entrada_kg || 0)} ${unidade} | Saidas ${formatPlainNumber(data.fiscal_saida_kg || 0)} ${unidade}`;
  el("comprasFisicoSaldo").textContent = `${formatPlainNumber(data.fisico_saldo_kg || 0)} ${unidade}`;
  el("comprasFisicoDetalhe").textContent = `Entradas ${formatPlainNumber(data.fisico_entrada_kg || 0)} ${unidade} | Saidas ${formatPlainNumber(data.fisico_saida_kg || 0)} ${unidade}`;
  const pendencias = Number(data.notas_pendentes || 0) + Number(data.saidas_pendentes_sefaz || 0);
  el("comprasEstoquePendencias").textContent = formatPlainNumber(pendencias);
  el("comprasEstoquePendenciasDetalhe").textContent = `${formatPlainNumber(data.notas_pendentes || 0)} nota(s) sem entrada | ${formatPlainNumber(data.saidas_pendentes_sefaz || 0)} saida(s) pendente(s) SEFAZ`;
  renderComprasAjusteProdutos(data.produtos || []);
  renderComprasEstoqueProdutos(data.produtos || []);
}

function renderComprasAjusteProdutos(produtos) {
  const select = el("comprasAjusteProduto");
  if (!select) return;
  const current = select.value;
  const options = [{produto_id: "", nome: currentLogisticaConfig().produto_padrao || "CARGA"}, ...(produtos || [])];
  select.innerHTML = options.map((item) => {
    const id = item.produto_id || "";
    const nome = item.nome || item.codigo || currentLogisticaConfig().produto_padrao || "CARGA";
    return `<option value="${escapeHtml(String(id))}" data-produto="${escapeHtml(nome)}">${escapeHtml(nome)}</option>`;
  }).join("");
  select.value = [...select.options].some((option) => option.value === current) ? current : "";
}

function renderComprasEstoqueProdutos(produtos) {
  const rows = el("comprasEstoqueProdutosRows");
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  if (!rows) return;
  rows.innerHTML = "";
  if (!produtos.length) {
    rows.appendChild(emptyRow(5, "Sem movimentacao de estoque por produto."));
    return;
  }
  produtos.forEach((item) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.nome || item.codigo || "SEM PRODUTO")}</td>
      <td>${escapeHtml(formatPlainNumber(item.fiscal_saldo_kg || 0))} ${escapeHtml(unidade)}</td>
      <td>${escapeHtml(formatPlainNumber(item.fisico_saldo_kg || 0))} ${escapeHtml(unidade)}</td>
      <td>${escapeHtml(formatPlainNumber(item.fiscal_entrada_kg || 0))} ${escapeHtml(unidade)}</td>
      <td>${escapeHtml(formatPlainNumber(item.fiscal_saida_kg || 0))} ${escapeHtml(unidade)}</td>
    `;
    rows.appendChild(tr);
  });
}

async function onComprasImportXml(event) {
  const input = event.currentTarget;
  const file = input.files && input.files[0];
  if (!file) return;
  const selectedNatureza = clean(el("comprasNatureza").value);
  const natureza = selectedNatureza === "TODAS" ? "COMPRA/AQUISICAO" : selectedNatureza;
  const body = new FormData();
  body.append("file", file);
  try {
    await apiRequest(`/compras/nfe/importar-xml?natureza_operacao=${encodeURIComponent(natureza)}`, {
      method: "POST",
      body,
    });
    notify("XML importado. Fornecedor verificado no cadastro.");
    await loadComprasNfe();
  } catch (error) {
    notify(error.message || "Falha ao importar XML.", true);
  } finally {
    input.value = "";
  }
}

async function downloadComprasXml() {
  const id = state.comprasNfe && state.comprasNfe.selectedId;
  if (!id) {
    notify("Selecione uma NF-e para baixar o XML.", true);
    return;
  }
  try {
    const result = await apiBlobRequest(`/compras/nfe/${encodeURIComponent(id)}/xml`);
    downloadBlob(result.blob, result.filename || `NFE_${id}.xml`);
  } catch (error) {
    notify(error.message || "Falha ao baixar XML.", true);
  }
}

async function confirmarComprasEntradaEstoque() {
  const id = state.comprasNfe && state.comprasNfe.selectedId;
  if (!id) {
    notify("Selecione uma NF-e para confirmar entrada no estoque.", true);
    return;
  }
  const payload = {
    quantidade_kg: parseDecimal(el("comprasEntradaKg").value),
    quantidade_caixas: Math.max(parseInt(clean(el("comprasEntradaCaixas").value) || "0", 10) || 0, 0),
    observacao: "Entrada confirmada pela tela de compras",
  };
  try {
    await apiRequest(`/compras/nfe/${encodeURIComponent(id)}/confirmar-entrada`, {
      method: "POST",
      body: JSON.stringify(payload),
    });
    notify("Entrada fiscal e fisica confirmada no estoque.");
    await loadComprasNfe();
  } catch (error) {
    notify(error.message || "Falha ao confirmar entrada no estoque.", true);
  }
}

async function syncComprasSaidasProgramacoes() {
  try {
    const data = await apiRequest("/compras/estoque/sincronizar-programacoes", {method: "POST"});
    notify(`${data.programacoes_vinculadas || 0} planejamento(s) vinculado(s), ${data.saidas_criadas || 0} saida(s) criada(s).`);
    await loadComprasNfe();
  } catch (error) {
    notify(error.message || "Falha ao sincronizar saidas das programacoes.", true);
  }
}

async function ajustarComprasEstoqueFisico() {
  const produtoSelect = el("comprasAjusteProduto");
  const selected = produtoSelect.selectedOptions[0];
  const payload = {
    tipo_estoque: "FISICO",
    tipo_movimento: clean(el("comprasAjusteMovimento").value) || "ENTRADA",
    produto_id: Number(produtoSelect.value || 0) || null,
    produto: selected ? clean(selected.dataset.produto || selected.textContent) : currentLogisticaConfig().produto_padrao || "CARGA",
    quantidade_kg: parseDecimal(el("comprasAjusteKg").value),
    quantidade_caixas: Math.max(parseInt(clean(el("comprasAjusteCaixas").value) || "0", 10) || 0, 0),
    observacao: clean(el("comprasAjusteObs").value) || "Ajuste manual de estoque fisico",
  };
  if (payload.quantidade_kg <= 0 && payload.quantidade_caixas <= 0) {
    notify("Informe KG ou caixas para ajustar o estoque fisico.", true);
    return;
  }
  try {
    await apiRequest("/compras/estoque/ajuste", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    el("comprasAjusteKg").value = "";
    el("comprasAjusteCaixas").value = "";
    el("comprasAjusteObs").value = "";
    notify("Ajuste manual registrado no estoque fisico.");
    await loadComprasNfe();
  } catch (error) {
    notify(error.message || "Falha ao ajustar estoque fisico.", true);
  }
}

async function zerarComprasEstoqueFisico() {
  const saldo = state.comprasEstoque ? Number(state.comprasEstoque.fisico_saldo_kg || 0) : 0;
  const ok = window.confirm(`Zerar estoque fisico atual (${formatPlainNumber(saldo)} ${currentLogisticaConfig().unidade_padrao || "KG"})? A operacao cria movimentacao de ajuste e preserva historico.`);
  if (!ok) return;
  try {
    const data = await apiRequest("/compras/estoque/zerar-fisico", {method: "POST"});
    notify(`${data.movimentos_criados || 0} movimento(s) criado(s) para zerar estoque fisico.`);
    await loadComprasNfe();
  } catch (error) {
    notify(error.message || "Falha ao zerar estoque fisico.", true);
  }
}

async function loadCentroCustos() {
  try {
    state.centroCustosOptions = await apiRequest("/centro-custos/options");
    renderCentroCustosOptions();
    await loadCentroCustosResumo();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderCentroCustosOptions() {
  const options = state.centroCustosOptions || {};
  fillSimpleSelect("centroCustosPeriodo", options.periodos || ["7", "15", "30", "60", "90", "180", "TODAS"], clean(el("centroCustosPeriodo").value) || "30");
  fillSimpleSelect("centroCustosMetric", options.metricas || ["CUSTO_KM", "CUSTO_KG", "DESPESA_TOTAL"], clean(el("centroCustosMetric").value) || "CUSTO_KM");
  fillSimpleSelect("centroCustosVeiculo", options.veiculos || ["TODOS"], clean(el("centroCustosVeiculo").value) || "TODOS");
  const veiculosDespesa = (options.veiculos || []).filter((item) => item && item !== "TODOS");
  fillSimpleSelect("centroDespesaVeiculo", veiculosDespesa.length ? veiculosDespesa : [""], clean(el("centroDespesaVeiculo").value) || veiculosDespesa[0] || "");
  fillObjectSelect("centroDespesaPerfil", options.despesa_veiculo_perfis || [{codigo: "OUTROS", nome: "Outros"}], clean(el("centroDespesaPerfil").value) || "OUTROS");
  fillObjectSelect("centroDespesaControleTipo", options.despesa_controles || [{codigo: "SEM_CONTROLE", nome: "Sem controle"}], clean(el("centroDespesaControleTipo").value) || "SEM_CONTROLE");
  fillSimpleSelect("centroDespesaPrioridade", options.prioridades || ["BAIXA", "NORMAL", "ALTA", "CRITICA"], clean(el("centroDespesaPrioridade").value) || "NORMAL");
  if (!clean(el("centroDespesaData").value)) {
    el("centroDespesaData").value = new Date().toISOString().slice(0, 10);
  }
  updateCentroDespesaControleFields();
}

function onCentroCustosVeiculoFilterChange() {
  const veiculo = clean(el("centroCustosVeiculo").value);
  state.centroCustosVeiculoSelecionado = veiculo && veiculo !== "TODOS" ? veiculo : "";
  state.centroCustosVeiculoDetalhe = null;
  loadCentroCustosResumo();
}

function fillSimpleSelect(id, values, selected) {
  const select = el(id);
  const current = selected || select.value;
  select.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  select.value = values.includes(current) ? current : values[0] || "";
}

function fillObjectSelect(id, items, selected) {
  const select = el(id);
  const current = selected || select.value;
  select.innerHTML = "";
  items.forEach((item) => {
    const codigo = item.codigo || item.value || "";
    const nome = item.nome || item.label || codigo;
    const option = document.createElement("option");
    option.value = codigo;
    option.textContent = nome;
    select.appendChild(option);
  });
  const values = [...select.options].map((option) => option.value);
  select.value = values.includes(current) ? current : values[0] || "";
}

async function loadCentroCustosResumo() {
  const params = new URLSearchParams({
    periodo: clean(el("centroCustosPeriodo").value) || "30",
    veiculo: clean(el("centroCustosVeiculo").value) || "TODOS",
    metric: clean(el("centroCustosMetric").value) || "CUSTO_KM",
  });
  try {
    const financeiroParams = new URLSearchParams({
      periodo: params.get("periodo"),
      veiculo: params.get("veiculo"),
    });
    const resumo = await apiRequest(`/centro-custos/resumo?${params.toString()}`);
    let veiculos = null;
    try {
      veiculos = await apiRequest(`/centro-custos/veiculos?${financeiroParams.toString()}`);
    } catch (error) {
      console.warn("Centro de custos por veiculo indisponivel:", error);
    }
    let financeiro = null;
    try {
      financeiro = await apiRequest(`/centro-custos/financeiro?${financeiroParams.toString()}`);
    } catch (error) {
      console.warn("Resumo financeiro do centro de custos indisponivel:", error);
    }
    let despesasRota = null;
    try {
      despesasRota = await apiRequest(`/centro-custos/despesas-rota?${financeiroParams.toString()}`);
    } catch (error) {
      console.warn("Despesas por rota indisponiveis:", error);
    }
    state.centroCustosResumo = resumo;
    state.centroCustosFinanceiro = financeiro;
    state.centroCustosVeiculos = veiculos;
    state.centroCustosDespesasRota = despesasRota;
    renderCentroCustosResumo();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderCentroCustosResumo() {
  const unidade = currentLogisticaConfig().unidade_padrao || "KG";
  const data = state.centroCustosResumo;
  if (!data) {
    el("centroCustosInfo").textContent = "Sem dados carregados.";
    el("centroCustosMetrics").innerHTML = "";
    el("centroCustosFinanceiroMetrics").innerHTML = "";
    el("centroCustosAnalise").innerHTML = "";
    el("centroCustosChart").innerHTML = "";
    el("centroCustosResultadoChart").innerHTML = "";
    el("centroCustosDespesaChart").innerHTML = "";
    el("centroCustosVeiculosGrid").innerHTML = "";
    closeCentroCustosVeiculoDetalhe();
    el("centroCustosRows").innerHTML = "";
    el("centroCustosRows").appendChild(emptyRow(11, "Sem dados."));
    el("centroCustosFinanceiroRows").innerHTML = "";
    el("centroCustosFinanceiroRows").appendChild(emptyRow(21, "Sem dados."));
    renderCentroCustosDespesasRota({});
    return;
  }
  const kpis = data.kpis || {};
  el("centroCustosInfo").textContent = `${data.periodo} dias / veiculo ${data.veiculo || "TODOS"} / ${data.metric || "CUSTO_KM"}`;
  el("centroCustosResumo").textContent = data.resumo || "-";
  el("centroCustosChartTitle").textContent = data.chart_title || "Custo por Veiculo";
  const metrics = [
    ["Veiculos", kpis.veiculos || 0, "frota no filtro"],
    ["Rotas", kpis.rotas || 0, "viagens analisadas"],
    ["KM rodado", formatPlainNumber(kpis.km_total || 0), "quilometragem total"],
    [`${unidade} carregado`, formatPlainNumber(kpis.kg_carregado || 0), "volume movimentado"],
    ["Custo/KM", formatCurrencyBR(kpis.custo_km_global || 0), "media global"],
    [`Custo/${unidade}`, formatCurrencyBR(kpis.custo_kg_global || 0), "media global"],
  ];
  el("centroCustosMetrics").innerHTML = metrics.map(([label, value, detail]) => `
    <article class="metric centro-custos-metric centro-custos-kpi-card">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <em>${escapeHtml(detail || "")}</em>
    </article>
  `).join("");
  renderCentroCustosVeiculos(state.centroCustosVeiculos || {});
  renderCentroCustosFinanceiro(state.centroCustosFinanceiro || {});
  renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {});
  renderCentroCustosChart(data.chart || [], data.metric || "CUSTO_KM");
  renderCentroCustosRows(data.rows || []);
}

function renderCentroCustosChart(chart, metric) {
  const wrap = el("centroCustosChart");
  wrap.innerHTML = "";
  if (!chart.length) {
    wrap.innerHTML = '<p class="empty-chart">Sem dados para grafico.</p>';
    return;
  }
  const maxValue = Math.max(...chart.map((item) => Number(item.value || 0)), 1);
  chart.forEach((item) => {
    const value = Number(item.value || 0);
    const pct = Math.max((value / maxValue) * 100, 2);
    const row = document.createElement("div");
    row.className = "centro-custos-chart-row";
    row.innerHTML = `
      <span>${escapeHtml(item.label || "-")}</span>
      <div class="centro-custos-chart-track">
        <div class="centro-custos-chart-bar" style="width: ${pct}%"></div>
      </div>
      <strong>${escapeHtml(formatCentroCustosChartValue(value, metric))}</strong>
    `;
    wrap.appendChild(row);
  });
}

function renderCentroCustosFinanceiro(data) {
  const kpis = data.kpis || {};
  el("centroCustosFinanceiroResumo").textContent = `${data.periodo || "30"} dias / veiculo ${data.veiculo || "TODOS"}`;
  const metrics = [
    ["Venda", formatCurrencyBR(kpis.venda_total || 0), "receita registrada", Number(kpis.venda_total || 0) > 0 ? "success" : "neutral"],
    ["Compra", formatCurrencyBR(kpis.compra_total || 0), "custo de mercadoria", "neutral"],
    ["Despesas", formatCurrencyBR(kpis.despesas_total || 0), "veiculo, rota e diarias", Number(kpis.despesas_total || 0) > 0 ? "warning" : "neutral"],
    ["Lucro liquido", formatCurrencyBR(kpis.lucro_liquido || 0), "resultado final", Number(kpis.lucro_liquido || 0) < 0 ? "danger" : "success"],
    ["Margem liquida", `${formatPlainNumber(kpis.margem_liquida || 0)}%`, "percentual do periodo", Number(kpis.margem_liquida || 0) < 0 ? "danger" : "neutral"],
    ["Lucro/KM", formatCurrencyBR(kpis.lucro_km || 0), "resultado por km", Number(kpis.lucro_km || 0) < 0 ? "danger" : "success"],
    ["Lucro/KG", formatCurrencyBR(kpis.lucro_kg || 0), "resultado por kg", Number(kpis.lucro_kg || 0) < 0 ? "danger" : "success"],
    ["Custo/KM", formatCurrencyBR(kpis.custo_km || 0), "custo medio", "neutral"],
    ["Custo/KG", formatCurrencyBR(kpis.custo_kg || 0), "custo medio", "neutral"],
  ];
  el("centroCustosFinanceiroMetrics").innerHTML = metrics.map(([label, value, detail, tone]) => `
    <article class="metric centro-custos-metric centro-custos-kpi-card ${escapeHtml(tone || "neutral")}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <em>${escapeHtml(detail || "")}</em>
    </article>
  `).join("");
  renderCentroCustosAnalise(data);
  renderCentroCustosResultadoChart(kpis);
  renderCentroCustosDespesaChart(data.composicao || []);
  renderCentroCustosFinanceiroRows(data.rows || []);
}

function renderCentroCustosDespesasRota(data) {
  fillCentroCustosDespesasRotaOptions(data.rows || []);
  fillCentroCustosDespesasRotaTipoOptions(data.rows || []);
  applyCentroCustosDespesasRotaDefaultMonth(data.rows || []);
  renderCentroCustosDespesasRotaQuickButtons();
  const tipo = clean(el("centroCustosDespesasRotaFiltroTipo").value) || "TODAS";
  const tipoLabel = centroCustosDespesaRotaTipoLabel(tipo);
  const rows = centroCustosDespesasRotaFilteredRows(data.rows || []);
  const lancamentos = centroCustosDespesasRotaLancamentos(rows, tipo);
  const selectedTotal = rows.reduce((sum, item) => sum + centroCustosDespesaRotaSelectedValue(item, tipo), 0);
  const routeTotal = rows.reduce((sum, item) => sum + Number(item.total || 0), 0);
  const kpis = {
    total: selectedTotal,
    total_rota: routeTotal,
    media_rota: rows.length ? selectedTotal / rows.length : 0,
  };
  el("centroCustosDespesasRotaResumo").textContent = rows.length
    ? `${tipoLabel} / ${rows.length} programacao(oes) / periodo ${data.periodo || "30"} dias / veiculo ${data.veiculo || "TODOS"}`
    : `Nenhuma despesa de ${tipoLabel.toLowerCase()} encontrada nos filtros.`;
  el("centroCustosDespesasRotaValorHeader").textContent = tipo === "TODAS" ? "VALOR" : `VALOR ${tipoLabel}`;
  const metrics = [
    [`Total ${tipoLabel}`, formatCurrencyBR(kpis.total || 0)],
    ["Programacoes", formatPlainNumber(rows.length)],
    ["Lancamentos", formatPlainNumber(lancamentos.length)],
    ["Media / programacao", formatCurrencyBR(kpis.media_rota || 0)],
  ];
  el("centroCustosDespesasRotaMetrics").innerHTML = metrics.map(([label, value]) => `
    <div class="centro-custos-rota-summary-item">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join("");
  renderCentroCustosDespesasRotaTipoCards(rows, tipo);
  renderCentroCustosDespesasRotaLancamentos(rows, tipo, tipoLabel, lancamentos);

  const tbody = el("centroCustosDespesasRotaRows");
  tbody.innerHTML = "";
  if (!rows.length) {
    tbody.appendChild(emptyRow(10, "Sem despesas de rota no periodo."));
    return;
  }
  rows.forEach((item) => {
    const valorSelecionado = centroCustosDespesaRotaSelectedValue(item, tipo);
    const tr = document.createElement("tr");
    tr.className = "centro-custos-rota-row";
    tr.title = "Clique para ver os lancamentos desta rota.";
    tr.innerHTML = `
      <td>${escapeHtml(item.codigo_programacao || "-")}</td>
      <td>${escapeHtml(formatDate(item.data) || item.data || "-")}</td>
      <td>${escapeHtml(item.veiculo || "-")}</td>
      <td>${escapeHtml(item.motorista || "-")}</td>
      <td>${escapeHtml(item.rota || "-")}</td>
      <td><strong>${escapeHtml(formatCurrencyBR(valorSelecionado || 0))}</strong></td>
      <td>${escapeHtml(formatCurrencyBR(item.total || 0))}</td>
      <td>${escapeHtml(item.qtd_despesas || 0)}</td>
      <td>${escapeHtml(item.maior_despesa || "-")}</td>
      <td><button type="button" class="link-button" data-rota-action="abrir">Prestacao</button></td>
    `;
    tr.addEventListener("click", () => showCentroCustosDespesasRotaDialog(item));
    tr.querySelector("[data-rota-action='abrir']").addEventListener("click", (event) => {
      event.stopPropagation();
      openCentroCustosRotaDespesas(item);
    });
    tbody.appendChild(tr);
  });
}

function renderCentroCustosDespesasRotaTipoCards(rows, tipoAtual) {
  const types = centroCustosDespesasRotaTypes(rows);
  el("centroCustosDespesasRotaTipoCards").innerHTML = types.map((item) => {
    return `
      <button type="button" class="${tipoAtual === item.tipo ? "active" : ""}" data-rota-type-card="${escapeHtml(item.tipo)}">
        <span>${escapeHtml(item.label)}</span>
        <strong>${escapeHtml(formatCurrencyBR(item.total))}</strong>
        <em>${escapeHtml(formatPlainNumber(item.programacoes))} prog.</em>
      </button>
    `;
  }).join("");
  document.querySelectorAll("[data-rota-type-card]").forEach((button) => {
    button.addEventListener("click", () => {
      el("centroCustosDespesasRotaFiltroTipo").value = button.dataset.rotaTypeCard || "TODAS";
      renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {});
    });
  });
}

function renderCentroCustosDespesasRotaLancamentos(rows, tipo, tipoLabel, sourceLancamentos = null) {
  const lancamentos = sourceLancamentos || centroCustosDespesasRotaLancamentos(rows, tipo);
  el("centroCustosDespesasRotaLancamentosResumo").textContent = lancamentos.length
    ? `${lancamentos.length} lancamento(s) de ${tipoLabel.toLowerCase()} totalizando ${formatCurrencyBR(lancamentos.reduce((sum, item) => sum + Number(item.valor || 0), 0))}.`
    : `Sem lancamentos individuais de ${tipoLabel.toLowerCase()} nos filtros.`;
  const tbody = el("centroCustosDespesasRotaLancamentosRows");
  tbody.innerHTML = "";
  if (!lancamentos.length) {
    tbody.appendChild(emptyRow(9, "Sem lancamentos individuais para o filtro."));
    return;
  }
  lancamentos.forEach((item) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.codigo_programacao || "-")}</td>
      <td>${escapeHtml(item.data_registro || item.data || "-")}</td>
      <td>${escapeHtml(item.tipo || "-")}</td>
      <td>${escapeHtml(item.motorista || "-")}</td>
      <td>${escapeHtml(item.veiculo || "-")}</td>
      <td>${escapeHtml(item.descricao || "-")}</td>
      <td>${escapeHtml(item.categoria || "-")}</td>
      <td>${escapeHtml(formatCurrencyBR(item.valor || 0))}</td>
      <td>${escapeHtml(item.observacao || "-")}</td>
    `;
    tbody.appendChild(tr);
  });
}

function setCentroCustosDespesasRotaTab(tab) {
  const active = tab === "lancamentos" ? "lancamentos" : "programacoes";
  el("centroCustosDespesasRotaProgramacoesPanel").classList.toggle("hidden", active !== "programacoes");
  el("centroCustosDespesasRotaLancamentosPanel").classList.toggle("hidden", active !== "lancamentos");
  document.querySelectorAll("[data-rota-tab]").forEach((button) => {
    button.classList.toggle("active", button.dataset.rotaTab === active);
  });
}

function centroCustosDespesasRotaLancamentos(rows, tipo) {
  const tipoValue = tipo || clean(el("centroCustosDespesasRotaFiltroTipo").value) || "TODAS";
  return rows.flatMap((row) => (row.despesas || [])
    .filter((despesa) => tipoValue === "TODAS" || despesa.tipo === tipoValue)
    .map((despesa) => ({
      ...despesa,
      codigo_programacao: row.codigo_programacao,
      data: row.data,
      motorista: row.motorista,
      veiculo: row.veiculo,
      rota: row.rota,
    })))
    .sort((a, b) => Number(b.valor || 0) - Number(a.valor || 0));
}

function fillCentroCustosDespesasRotaTipoOptions(rows) {
  const select = el("centroCustosDespesasRotaFiltroTipo");
  const selected = clean(select.value) || "TODAS";
  const types = centroCustosDespesasRotaTypes(rows);
  select.innerHTML = '<option value="TODAS">Todos</option>';
  types.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.tipo;
    option.textContent = item.label;
    select.appendChild(option);
  });
  select.value = selected === "TODAS" || types.some((item) => item.tipo === selected) ? selected : "TODAS";
}

function centroCustosDespesasRotaTypes(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const tipoTotais = row.tipo_totais || {};
    Object.entries(tipoTotais).forEach(([tipo, valor]) => {
      const tipoValue = clean(tipo).toUpperCase();
      if (!tipoValue || Number(valor || 0) <= 0) return;
      const current = map.get(tipoValue) || {tipo: tipoValue, label: centroCustosDespesaRotaTipoLabel(tipoValue), total: 0, programacoes: 0};
      current.total += Number(valor || 0);
      current.programacoes += 1;
      map.set(tipoValue, current);
    });
  });
  return [...map.values()].sort((a, b) => Number(b.total || 0) - Number(a.total || 0));
}

function fillCentroCustosDespesasRotaOptions(rows) {
  const select = el("centroCustosDespesasRotaFiltroRota");
  const selected = clean(select.value) || "TODAS";
  const rotas = [...new Set(rows.map((item) => clean(item.rota)).filter(Boolean))].sort();
  select.innerHTML = '<option value="TODAS">Todas</option>';
  rotas.forEach((rota) => {
    const option = document.createElement("option");
    option.value = rota;
    option.textContent = rota;
    select.appendChild(option);
  });
  select.value = rotas.includes(selected) ? selected : "TODAS";
}

function applyCentroCustosDespesasRotaDefaultMonth(rows) {
  if (state.centroCustosDespesasRotaMesAuto) return;
  if (clean(el("centroCustosDespesasRotaMes").value) || clean(el("centroCustosDespesasRotaDataIni").value) || clean(el("centroCustosDespesasRotaDataFim").value)) {
    state.centroCustosDespesasRotaMesAuto = true;
    return;
  }
  const latestDate = rows
    .map((item) => String(item.data || "").slice(0, 10))
    .filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(value))
    .sort()
    .pop();
  if (!latestDate) return;
  const month = latestDate.slice(0, 7);
  el("centroCustosDespesasRotaMes").value = month;
  setCentroCustosDespesasRotaMonthRange(month);
  state.centroCustosDespesasRotaMesAuto = true;
}

function centroCustosDespesasRotaFilteredRows(sourceRows = null) {
  const rows = [...(sourceRows || state.centroCustosDespesasRota?.rows || [])];
  const baseTotal = rows.reduce((sum, item) => sum + Number(item.total || 0), 0);
  const baseMedia = rows.length ? baseTotal / rows.length : 0;
  const busca = clean(el("centroCustosDespesasRotaBusca").value).toUpperCase();
  const rota = clean(el("centroCustosDespesasRotaFiltroRota").value) || "TODAS";
  const tipo = clean(el("centroCustosDespesasRotaFiltroTipo").value) || "TODAS";
  const dataIni = clean(el("centroCustosDespesasRotaDataIni").value);
  const dataFim = clean(el("centroCustosDespesasRotaDataFim").value);
  const valorMin = parseDecimal(el("centroCustosDespesasRotaValorMin").value);
  const valorMax = parseDecimal(el("centroCustosDespesasRotaValorMax").value);
  const ordenacao = clean(el("centroCustosDespesasRotaOrdenacao").value) || "TOTAL_DESC";
  const filtered = rows.filter((item) => {
    const data = String(item.data || "").slice(0, 10);
    const selectedValue = centroCustosDespesaRotaSelectedValue(item, tipo);
    const text = [item.codigo_programacao, item.veiculo, item.motorista, item.rota, item.maior_despesa].join(" ").toUpperCase();
    return (!busca || text.includes(busca))
      && (rota === "TODAS" || item.rota === rota)
      && (tipo === "TODAS" || centroCustosDespesaRotaSelectedValue(item, tipo) > 0)
      && (!dataIni || data >= dataIni)
      && (!dataFim || data <= dataFim)
      && (!valorMin || selectedValue >= valorMin)
      && (!valorMax || selectedValue <= valorMax)
      && centroCustosDespesasRotaQuickPasses(item, state.centroCustosDespesasRotaQuick, baseMedia);
  });
  filtered.sort((a, b) => {
    if (ordenacao === "TOTAL_ASC") return centroCustosDespesaRotaSelectedValue(a, tipo) - centroCustosDespesaRotaSelectedValue(b, tipo);
    if (ordenacao === "DATA_DESC") return String(b.data || "").localeCompare(String(a.data || ""));
    if (ordenacao === "PROGRAMACAO_ASC") return String(a.codigo_programacao || "").localeCompare(String(b.codigo_programacao || ""));
    return centroCustosDespesaRotaSelectedValue(b, tipo) - centroCustosDespesaRotaSelectedValue(a, tipo);
  });
  return filtered;
}

function centroCustosDespesaRotaSelectedValue(item, tipo = null) {
  const tipoValue = tipo || clean(el("centroCustosDespesasRotaFiltroTipo").value) || "TODAS";
  const field = {DIARIAS: "diarias", BANHOS: "banhos", GUARDAS: "guardas", OUTRAS: "outras"}[tipoValue];
  if (field) return Number(item[field] || 0);
  if (tipoValue === "TODAS") return Number(item.total || 0);
  return Number((item.tipo_totais || {})[tipoValue] || 0);
}

function centroCustosDespesaRotaTipoLabel(tipo) {
  return {
    DIARIAS: "Diarias",
    BANHOS: "Banhos",
    GUARDAS: "Guardas",
    OUTRAS: "Outras",
  }[tipo] || clean(tipo).replaceAll("_", " ").toLowerCase().replace(/\b\w/g, (letter) => letter.toUpperCase()) || "Despesas";
}

function onCentroCustosDespesasRotaMesChange() {
  const month = clean(el("centroCustosDespesasRotaMes").value);
  if (month) {
    setCentroCustosDespesasRotaMonthRange(month);
  }
  renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {});
}

function setCentroCustosDespesasRotaMonthRange(month) {
  const [year, monthNumber] = month.split("-").map((part) => Number(part));
  if (!year || !monthNumber) return;
  const start = `${year}-${String(monthNumber).padStart(2, "0")}-01`;
  const endDate = new Date(year, monthNumber, 0);
  const end = `${year}-${String(monthNumber).padStart(2, "0")}-${String(endDate.getDate()).padStart(2, "0")}`;
  el("centroCustosDespesasRotaDataIni").value = start;
  el("centroCustosDespesasRotaDataFim").value = end;
}

function clearCentroCustosDespesasRotaFilters() {
  el("centroCustosDespesasRotaBusca").value = "";
  el("centroCustosDespesasRotaFiltroRota").value = "TODAS";
  el("centroCustosDespesasRotaFiltroTipo").value = "TODAS";
  el("centroCustosDespesasRotaMes").value = "";
  el("centroCustosDespesasRotaDataIni").value = "";
  el("centroCustosDespesasRotaDataFim").value = "";
  el("centroCustosDespesasRotaValorMin").value = "";
  el("centroCustosDespesasRotaValorMax").value = "";
  el("centroCustosDespesasRotaOrdenacao").value = "TOTAL_DESC";
  state.centroCustosDespesasRotaQuick = "";
  state.centroCustosDespesasRotaMesAuto = false;
  renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {});
}

function toggleCentroCustosDespesasRotaQuick(value) {
  state.centroCustosDespesasRotaQuick = state.centroCustosDespesasRotaQuick === value ? "" : value;
  renderCentroCustosDespesasRota(state.centroCustosDespesasRota || {});
}

function renderCentroCustosDespesasRotaQuickButtons() {
  document.querySelectorAll("[data-rota-quick]").forEach((button) => {
    button.classList.toggle("active", button.dataset.rotaQuick === state.centroCustosDespesasRotaQuick);
  });
}

function centroCustosDespesasRotaQuickPasses(item, quick, baseMedia) {
  if (!quick) return true;
  if (quick === "ACIMA_MEDIA") return Number(item.total || 0) > baseMedia;
  return true;
}

function showCentroCustosDespesasRotaDialog(item) {
  const despesas = item.despesas || [];
  el("centroCustosDespesasRotaDialogTitle").textContent = `Despesas ${item.codigo_programacao || "-"}`;
  el("centroCustosDespesasRotaDialogSubtitle").textContent = `${item.motorista || "-"} | ${item.veiculo || "-"} | ${item.rota || "-"} | Total ${formatCurrencyBR(item.total || 0)}`;
  const tbody = el("centroCustosDespesasRotaDialogRows");
  tbody.innerHTML = "";
  if (!despesas.length) {
    tbody.appendChild(emptyRow(7, "Sem lancamentos detalhados."));
  } else {
    despesas.forEach((despesa) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(despesa.id || "-")}</td>
        <td>${escapeHtml(despesa.data_registro || "-")}</td>
        <td>${escapeHtml(despesa.tipo || "OUTRAS")}</td>
        <td>${escapeHtml(despesa.categoria || "-")}</td>
        <td>${escapeHtml(despesa.descricao || "-")}</td>
        <td>${escapeHtml(formatCurrencyBR(despesa.valor || 0))}</td>
        <td>${escapeHtml(despesa.observacao || "-")}</td>
      `;
      tbody.appendChild(tr);
    });
  }
  const openButton = el("centroCustosDespesasRotaDialogOpenButton");
  openButton.onclick = () => {
    closeCentroCustosDespesasRotaDialog();
    openCentroCustosRotaDespesas(item);
  };
  const dialog = el("centroCustosDespesasRotaDialog");
  if (!dialog.open) dialog.showModal();
}

function closeCentroCustosDespesasRotaDialog() {
  const dialog = el("centroCustosDespesasRotaDialog");
  if (dialog.open) dialog.close();
}

async function openCentroCustosRotaDespesas(item) {
  const codigo = clean(item.codigo_programacao);
  if (!codigo) {
    notify("Programacao da rota nao encontrada.", true);
    return;
  }
  state.despesasSelectedCodigo = codigo;
  switchView("despesas", {skipLoad: true});
  ensureDespesasProgramacaoOption(codigo, item);
  el("despesasInfo").textContent = `${codigo} | Carregando custos...`;
  await loadDespesasBundle(codigo);
}

function exportCentroCustosDespesasRotaCsv() {
  const rows = centroCustosDespesasRotaFilteredRows();
  if (!rows.length) {
    notify("Sem despesas de rota para exportar.", true);
    return;
  }
  const tipo = clean(el("centroCustosDespesasRotaFiltroTipo").value) || "TODAS";
  const types = centroCustosDespesasRotaTypes(rows);
  const headers = ["PROGRAMACAO", "DATA", "VEICULO", "MOTORISTA", "ROTA", "VALOR_FILTRO", ...types.map((item) => item.label.toUpperCase()), "TOTAL_ROTA", "QTD_DESPESAS", "MAIOR_DESPESA"];
  const lines = [
    headers.map(csvCell).join(";"),
    ...rows.map((item) => [
      item.codigo_programacao,
      item.data,
      item.veiculo,
      item.motorista,
      item.rota,
      formatPlainNumber(centroCustosDespesaRotaSelectedValue(item, tipo) || 0),
      ...types.map((typeItem) => formatPlainNumber((item.tipo_totais || {})[typeItem.tipo] || 0)),
      formatPlainNumber(item.total || 0),
      item.qtd_despesas || 0,
      item.maior_despesa,
    ].map(csvCell).join(";")),
  ];
  const stamp = new Date().toISOString().slice(0, 10);
  downloadTextFile(`despesas_rotas_${stamp}.csv`, `\uFEFF${lines.join("\r\n")}\r\n`, "text/csv;charset=utf-8");
  notify(`${rows.length} rota(s) exportada(s).`);
}

function renderCentroCustosVeiculos(data) {
  const wrap = el("centroCustosVeiculosGrid");
  const veiculos = data.veiculos || [];
  if (!veiculos.length) {
    wrap.innerHTML = '<article class="centro-custos-veiculo-card empty">Nenhum veiculo cadastrado ou movimentado no filtro.</article>';
    return;
  }
  wrap.innerHTML = "";
  veiculos.forEach((item) => {
    const card = document.createElement("button");
    const resultado = Number(item.lucro_liquido || 0);
    card.type = "button";
    card.className = `centro-custos-veiculo-card ${resultado < 0 ? "negative" : resultado > 0 ? "positive" : ""} ${state.centroCustosVeiculoSelecionado === item.placa ? "active" : ""}`;
    card.innerHTML = `
      <span>Veiculo</span>
      <strong>${escapeHtml(item.placa || "-")}</strong>
      <em>${escapeHtml(item.modelo || "Modelo nao informado")}</em>
      <b>${escapeHtml(formatCurrencyBR(resultado))}</b>
      <em class="centro-custos-veiculo-result-label">${resultado < 0 ? "Prejuizo liquido" : resultado > 0 ? "Lucro liquido" : "Resultado zerado"}</em>
      <div class="centro-custos-veiculo-card-grid">
        <small><i>Compra</i>${escapeHtml(formatCurrencyBR(item.compra_total || 0))}</small>
        <small><i>Venda</i>${escapeHtml(formatCurrencyBR(item.venda_total || 0))}</small>
        <small><i>Despesas</i>${escapeHtml(formatCurrencyBR(item.despesas_total || 0))}</small>
        <small><i>Viagens</i>${escapeHtml(formatPlainNumber(item.rotas || 0))}</small>
      </div>
    `;
    card.addEventListener("click", () => onSelectCentroCustosVeiculo(item.placa));
    wrap.appendChild(card);
  });
}

async function onSelectCentroCustosVeiculo(placa) {
  const placaValue = clean(placa);
  if (!placaValue) return;
  state.centroCustosVeiculoSelecionado = placaValue;
  renderCentroCustosVeiculos(state.centroCustosVeiculos || {});
  await loadCentroCustosVeiculoDetalhe(placaValue);
}

async function loadCentroCustosVeiculoDetalhe(placa) {
  try {
    const params = new URLSearchParams({
      periodo: clean(el("centroCustosPeriodo").value) || "30",
    });
    state.centroCustosVeiculoDetalhe = await apiRequest(`/centro-custos/veiculos/${encodeURIComponent(placa)}?${params.toString()}`);
    renderCentroCustosVeiculoDetalhe();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderCentroCustosVeiculoDetalhe() {
  const data = state.centroCustosVeiculoDetalhe || {};
  const veiculo = data.veiculo || {};
  const panel = el("centroCustosVeiculoDetalhe");
  if (!veiculo.placa) {
    closeCentroCustosVeiculoDetalhe();
    return;
  }
  if (!panel.open) panel.showModal();
  el("centroVeiculoTitulo").textContent = `${veiculo.placa} ${veiculo.modelo ? `| ${veiculo.modelo}` : ""}`;
  el("centroVeiculoSubtitulo").textContent = `${data.periodo || "30"} dias | ${veiculo.rotas || 0} viagem(ns) | ultima movimentacao ${veiculo.ultima_data || "-"}`;
  const kpis = [
    ["Despesas", formatCurrencyBR(veiculo.despesas_total || 0)],
    ["Compra", formatCurrencyBR(veiculo.compra_total || 0)],
    ["Venda", formatCurrencyBR(veiculo.venda_total || 0)],
    ["Resultado", formatCurrencyBR(veiculo.lucro_liquido || 0)],
    ["Manutencao", formatCurrencyBR(veiculo.despesas_manutencao || 0)],
    ["KM rodado", formatPlainNumber(veiculo.km_rodado || 0)],
    ["Litros", formatPlainNumber(veiculo.litros || 0)],
    ["Media", `${formatMetricDecimal(veiculo.media_consumo || 0)} km/l`],
    ["Motoristas", veiculo.motoristas || 0],
  ];
  el("centroVeiculoKpis").innerHTML = kpis.map(([label, value]) => `
    <article class="metric centro-custos-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `).join("");
  el("centroVeiculoAlertas").innerHTML = (data.alertas || []).map((alerta) => `
    <article>${escapeHtml(alerta)}</article>
  `).join("");
  renderCentroCustosPecas(data.pecas || []);
  renderCentroCustosMotoristas(data.motoristas || []);
  renderCentroCustosVeiculoProgramacoes(data.programacoes || []);
  renderCentroCustosVeiculoDespesas(data.despesas || []);
}

function closeCentroCustosVeiculoDetalhe() {
  const dialog = el("centroCustosVeiculoDetalhe");
  if (dialog.open) dialog.close();
}

function renderCentroCustosPecas(pecas) {
  renderCentroCustosSimpleBars("centroVeiculoPecas", pecas.map((item) => ({
    label: `${item.grupo || "-"} (${item.eventos || 0})`,
    value: Number(item.valor || 0),
  })));
}

function renderCentroCustosMotoristas(motoristas) {
  const wrap = el("centroVeiculoMotoristas");
  wrap.innerHTML = motoristas.length
    ? motoristas.map((item) => `<span>${escapeHtml(item)}</span>`).join("")
    : '<p class="empty-state">Sem motorista no periodo.</p>';
}

function renderCentroCustosVeiculoProgramacoes(programacoes) {
  const rows = el("centroVeiculoProgramacoesRows");
  rows.innerHTML = "";
  if (!programacoes.length) {
    rows.appendChild(emptyRow(8, "Sem programacoes no periodo."));
    return;
  }
  programacoes.forEach((item) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.codigo_programacao || "")}</td>
      <td>${escapeHtml(item.data || "")}</td>
      <td>${escapeHtml(item.motorista || "")}</td>
      <td>${escapeHtml(item.rota || "")}</td>
      <td>${escapeHtml(formatPlainNumber(item.km_rodado || 0))}</td>
      <td>${escapeHtml(formatPlainNumber(item.litros || 0))}</td>
      <td>${escapeHtml(formatMetricDecimal(item.media_km_l || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.despesas_total || 0))}</td>
    `;
    rows.appendChild(tr);
  });
}

function renderCentroCustosVeiculoDespesas(despesas) {
  const rows = el("centroVeiculoDespesasRows");
  rows.innerHTML = "";
  if (!despesas.length) {
    rows.appendChild(emptyRow(8, "Sem despesas no periodo."));
    return;
  }
  despesas.forEach((item) => {
    const status = item.status_controle || "";
    const controle = status ? `${item.controle_tipo || "-"} | ${status}` : (item.controle_tipo || "-");
    const tr = document.createElement("tr");
    if (status === "VENCIDO") tr.className = "centro-custos-controle-vencido";
    if (status === "PROXIMO") tr.className = "centro-custos-controle-proximo";
    tr.innerHTML = `
      <td>${escapeHtml(item.data_registro || "")}</td>
      <td>${escapeHtml(item.grupo || item.categoria || "")}</td>
      <td>${escapeHtml(item.descricao || "")}</td>
      <td>${escapeHtml(controle)}</td>
      <td>${escapeHtml(item.data_vencimento || "-")}</td>
      <td>${escapeHtml(item.km_vencimento ? formatPlainNumber(item.km_vencimento) : "-")}</td>
      <td>${escapeHtml(item.documento || item.codigo_programacao || "")}</td>
      <td>${escapeHtml(formatCurrencyBR(item.valor || 0))}</td>
    `;
    rows.appendChild(tr);
  });
}

function renderCentroCustosAnalise(data) {
  const kpis = data.kpis || {};
  const rows = data.rows || [];
  const piorResultado = rows.length
    ? [...rows].sort((a, b) => Number(a.lucro_liquido || 0) - Number(b.lucro_liquido || 0))[0]
    : null;
  const maiorDespesa = rows.length
    ? [...rows].sort((a, b) => Number(b.despesas_total || 0) - Number(a.despesas_total || 0))[0]
    : null;
  const margem = Number(kpis.margem_liquida || 0);
  const statusResultado = Number(kpis.lucro_liquido || 0) < 0
    ? "Prejuizo no periodo"
    : margem < 5
      ? "Margem apertada"
      : "Resultado positivo";
  const despesasTotal = Number(kpis.despesas_total || 0);
  const compraTotal = Number(kpis.compra_total || 0);
  const vendaTotal = Number(kpis.venda_total || 0);
  const lucroLiquido = Number(kpis.lucro_liquido || 0);
  const cards = [
    {
      label: "Leitura do periodo",
      value: statusResultado,
      detail: `Resultado ${formatCurrencyBR(lucroLiquido)} com margem de ${formatPlainNumber(margem)}%.`,
      tone: lucroLiquido < 0 ? "danger" : margem < 5 ? "warning" : "success",
    },
    {
      label: "Maior impacto",
      value: piorResultado ? `${piorResultado.codigo_programacao || "-"} | ${piorResultado.veiculo || "-"}` : "-",
      detail: piorResultado ? `Resultado ${formatCurrencyBR(piorResultado.lucro_liquido || 0)}.` : "Sem rotas no filtro.",
      tone: piorResultado && Number(piorResultado.lucro_liquido || 0) < 0 ? "danger" : "neutral",
    },
    {
      label: "Peso das despesas",
      value: `${formatPlainNumber(compraTotal ? (despesasTotal / compraTotal) * 100 : 0)}% da compra`,
      detail: maiorDespesa
        ? `Maior despesa em ${maiorDespesa.codigo_programacao || "-"}: ${formatCurrencyBR(maiorDespesa.despesas_total || 0)}.`
        : `Despesas ${formatCurrencyBR(despesasTotal)} sobre compra ${formatCurrencyBR(compraTotal)}.`,
      tone: despesasTotal > 0 && !vendaTotal ? "warning" : "neutral",
    },
    {
      label: "Eficiencia",
      value: `${formatCurrencyBR(kpis.custo_km || 0)} / KM`,
      detail: `${formatCurrencyBR(kpis.custo_kg || 0)} por KG movimentado.`,
      tone: "neutral",
    },
  ];
  el("centroCustosAnalise").innerHTML = cards.map((card) => `
    <article class="centro-custos-insight ${escapeHtml(card.tone)}">
      <span>${escapeHtml(card.label)}</span>
      <strong>${escapeHtml(card.value)}</strong>
      <p>${escapeHtml(card.detail)}</p>
    </article>
  `).join("");
}

function renderCentroCustosResultadoChart(kpis) {
  renderCentroCustosSimpleBars("centroCustosResultadoChart", [
    {label: "Venda", value: Number(kpis.venda_total || 0)},
    {label: "Compra", value: Number(kpis.compra_total || 0)},
    {label: "Despesas", value: Number(kpis.despesas_total || 0)},
    {label: "Lucro liquido", value: Number(kpis.lucro_liquido || 0)},
  ]);
}

function renderCentroCustosDespesaChart(composicao) {
  renderCentroCustosSimpleBars("centroCustosDespesaChart", composicao.map((item) => ({
    label: `${item.grupo || "-"} (${formatPlainNumber(item.percentual || 0)}%)`,
    value: Number(item.valor || 0),
  })));
}

function renderCentroCustosSimpleBars(targetId, rows) {
  const wrap = el(targetId);
  wrap.innerHTML = "";
  if (!rows.length) {
    wrap.innerHTML = '<p class="empty-chart">Sem dados para grafico.</p>';
    return;
  }
  const maxValue = Math.max(...rows.map((item) => Math.abs(Number(item.value || 0))), 1);
  rows.forEach((item) => {
    const value = Number(item.value || 0);
    const pct = Math.max((Math.abs(value) / maxValue) * 100, value !== 0 ? 2 : 0);
    const row = document.createElement("div");
    row.className = `centro-custos-chart-row ${value < 0 ? "negative" : ""}`;
    row.innerHTML = `
      <span>${escapeHtml(item.label || "-")}</span>
      <div class="centro-custos-chart-track">
        <div class="centro-custos-chart-bar" style="width: ${pct}%"></div>
      </div>
      <strong>${escapeHtml(formatCurrencyBR(value))}</strong>
    `;
    wrap.appendChild(row);
  });
}

function renderCentroCustosFinanceiroRows(rowsData) {
  const rows = el("centroCustosFinanceiroRows");
  rows.innerHTML = "";
  if (!rowsData.length) {
    rows.appendChild(emptyRow(21, "Sem resultado financeiro no filtro selecionado."));
    return;
  }
  const renderRow = (item, options = {}) => {
    const filhos = Array.isArray(item.filhos) ? item.filhos : [];
    const hasChildren = filhos.length > 0 || item.has_children;
    const parentCodigo = clean(options.parentCodigo || item.parent_codigo || "");
    const isChild = Boolean(options.isChild || Number(item.nivel || 0) > 0);
    const tr = document.createElement("tr");
    tr.className = `${Number(item.lucro_liquido || 0) < 0 ? "centro-custos-alto" : "centro-custos-bom"} ${isChild ? "centro-custos-child-row hidden" : ""} ${hasChildren ? "centro-custos-parent-row" : ""}`;
    if (parentCodigo) tr.dataset.parentCodigo = parentCodigo;
    if (hasChildren) tr.dataset.codigo = item.codigo_programacao || "";
    const progCell = hasChildren
      ? `<button type="button" class="centro-custos-row-toggle" data-codigo="${escapeHtml(item.codigo_programacao || "")}" aria-expanded="false">+</button>${escapeHtml(item.codigo_programacao || "")}`
      : `${isChild ? '<span class="centro-custos-child-prefix">&rdsh;</span>' : ""}${escapeHtml(item.codigo_programacao || "")}`;
    const vendaTitle = [
      `Confirmada: ${formatCurrencyBR(item.venda_confirmada || 0)}`,
      `Prevista: ${formatCurrencyBR(item.venda_prevista || 0)}`,
      `Fonte: ${formatCentroCustosFonte(item.fonte_venda)}`,
    ].join(" | ");
    const alertas = Array.isArray(item.alertas) ? item.alertas : [];
    const alertasResumo = formatCentroCustosAlertas(alertas);
    tr.innerHTML = `
      <td>${progCell}</td>
      <td>${escapeHtml(item.data || "")}</td>
      <td>${escapeHtml(item.veiculo || "")}</td>
      <td>${escapeHtml(item.motorista || "")}</td>
      <td>${escapeHtml(item.rota || "")}</td>
      <td>${escapeHtml(formatPlainNumber(item.km_rodado || 0))}</td>
      <td>${escapeHtml(formatPlainNumber(item.kg || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.compra || 0))}</td>
      <td title="${escapeHtml(vendaTitle)}">${escapeHtml(formatCurrencyBR(item.venda || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.despesas_veiculo || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.despesas_rota || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.diarias || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.despesas_total || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.lucro_bruto || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.lucro_liquido || 0))}</td>
      <td>${escapeHtml(formatPlainNumber(item.margem_liquida || 0))}%</td>
      <td>${escapeHtml(formatCurrencyBR(item.custo_km || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.lucro_km || 0))}</td>
      <td>${escapeHtml(formatCentroCustosFonte(item.fonte_venda))}</td>
      <td><span class="centro-custos-confidence ${escapeHtml(formatCentroCustosConfidenceClass(item.confianca))}">${escapeHtml(formatCentroCustosConfianca(item.confianca))}</span></td>
      <td class="centro-custos-alert-cell" title="${escapeHtml(alertas.join(" | "))}">${escapeHtml(alertasResumo)}</td>
    `;
    rows.appendChild(tr);
    filhos.forEach((filho) => renderRow(filho, {isChild: true, parentCodigo: item.codigo_programacao || ""}));
  };
  rowsData.forEach((item) => renderRow(item));
  rows.querySelectorAll(".centro-custos-row-toggle").forEach((button) => {
    button.addEventListener("click", () => {
      const codigo = clean(button.dataset.codigo || "");
      const expanded = button.getAttribute("aria-expanded") === "true";
      rows.querySelectorAll(`tr[data-parent-codigo="${CSS.escape(codigo)}"]`).forEach((child) => {
        child.classList.toggle("hidden", expanded);
      });
      button.setAttribute("aria-expanded", expanded ? "false" : "true");
      button.textContent = expanded ? "+" : "-";
    });
  });
}

function formatCentroCustosFonte(value) {
  const fonte = clean(value).toUpperCase();
  const labels = {
    RECEBIMENTOS: "Recebido",
    CONTROLES: "Controle",
    ITENS: "Venda prevista",
    IMPORTACAO: "Importacao",
    ITENS_COM_RECEBIMENTO_PARCIAL: "Receb. parcial",
    IMPORTACAO_COM_RECEBIMENTO_PARCIAL: "Import. parcial",
    TRANSBORDO_CONSOLIDADO: "Transbordo consolidado",
    DESPESA_AVULSA: "Despesa avulsa",
    SEM_VENDA: "Sem venda",
  };
  return labels[fonte] || fonte || "-";
}

function formatCentroCustosConfianca(value) {
  const confianca = clean(value).toUpperCase();
  if (confianca === "ALTA") return "Alta";
  if (confianca === "MEDIA") return "Media";
  if (confianca === "BAIXA") return "Baixa";
  return confianca || "-";
}

function formatCentroCustosConfidenceClass(value) {
  const confianca = clean(value).toUpperCase();
  if (confianca === "ALTA") return "alta";
  if (confianca === "MEDIA") return "media";
  if (confianca === "BAIXA") return "baixa";
  return "neutra";
}

function formatCentroCustosAlertas(alertas) {
  if (!alertas.length) return "-";
  const principais = alertas.slice(0, 2).map((alerta) => {
    const texto = clean(alerta);
    if (texto.includes("Recebimento parcial")) return "Recebimento parcial";
    if (texto.includes("Venda confirmada difere")) return "Venda divergente";
    if (texto.includes("Resultado consolidado")) return "Transbordo consolidado";
    if (texto.includes("Sem venda")) return "Sem venda vinculada";
    if (texto.includes("sem KM")) return "Sem KM";
    if (texto.includes("sem KG")) return "Sem KG";
    if (texto.includes("sem despesas")) return "Sem despesas";
    return texto;
  });
  return principais.join(" | ") + (alertas.length > principais.length ? ` +${alertas.length - principais.length}` : "");
}

function renderCentroCustosRows(rowsData) {
  const rows = el("centroCustosRows");
  rows.innerHTML = "";
  if (!rowsData.length) {
    rows.appendChild(emptyRow(11, "Sem custos no filtro selecionado."));
    return;
  }
  rowsData.forEach((item) => {
    const tr = document.createElement("tr");
    tr.className = `centro-custos-${item.classificacao || "medio"}`;
    tr.innerHTML = `
      <td>${escapeHtml(item.veiculo || "")}</td>
      <td>${escapeHtml(item.rotas || 0)}</td>
      <td>${escapeHtml(formatPlainNumber(item.km_rodado || 0))}</td>
      <td>${escapeHtml(formatPlainNumber(item.kg_carregado || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.despesas || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.compra || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.venda || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.lucro_liquido || 0))}</td>
      <td>${escapeHtml(formatMetricDecimal(item.custo_km || 0))}</td>
      <td>${escapeHtml(formatMetricDecimal(item.custo_kg || 0))}</td>
      <td>${escapeHtml(formatCurrencyBR(item.ticket_rota || 0))}</td>
    `;
    rows.appendChild(tr);
  });
}

function formatMetricDecimal(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) ? number.toFixed(3) : "0.000";
}

function formatCentroCustosChartValue(value, metric) {
  if (metric === "DESPESA_TOTAL") {
    return formatCurrencyBR(value);
  }
  return formatMetricDecimal(value);
}

function clearCentroCustosDespesaVeiculoForm() {
  const form = el("centroCustosDespesaVeiculoForm");
  form.reset();
  const options = state.centroCustosOptions || {};
  const veiculos = (options.veiculos || []).filter((item) => item && item !== "TODOS");
  el("centroDespesaVeiculo").value = veiculos[0] || "";
  el("centroDespesaDocumentoTipo").value = "CUPOM FISCAL";
  el("centroDespesaPerfil").value = "OUTROS";
  el("centroDespesaControleTipo").value = "SEM_CONTROLE";
  el("centroDespesaPrioridade").value = "NORMAL";
  el("centroDespesaData").value = new Date().toISOString().slice(0, 10);
  updateCentroDespesaControleFields();
}

function updateCentroDespesaControleFields() {
  const controle = clean(el("centroDespesaControleTipo").value) || "SEM_CONTROLE";
  const usaData = ["DATA", "DATA_KM"].includes(controle);
  const usaKm = ["KM", "DATA_KM"].includes(controle);
  el("centroDespesaDataVencimento").disabled = !usaData;
  el("centroDespesaKmVencimento").disabled = !usaKm;
  if (!usaData) el("centroDespesaDataVencimento").value = "";
  if (!usaKm) el("centroDespesaKmVencimento").value = "";
}

async function onSaveCentroCustosDespesaVeiculo(event) {
  event.preventDefault();
  const payload = {
    veiculo: clean(el("centroDespesaVeiculo").value),
    documento_tipo: clean(el("centroDespesaDocumentoTipo").value) || "MANUAL",
    documento_numero: clean(el("centroDespesaDocumentoNumero").value),
    data_registro: clean(el("centroDespesaData").value),
    valor: parseDecimal(el("centroDespesaValor").value),
    fornecedor: clean(el("centroDespesaFornecedor").value),
    perfil: clean(el("centroDespesaPerfil").value) || "OUTROS",
    controle_tipo: clean(el("centroDespesaControleTipo").value) || "SEM_CONTROLE",
    data_vencimento: clean(el("centroDespesaDataVencimento").value),
    km_vencimento: parseDecimal(el("centroDespesaKmVencimento").value),
    odometro: parseDecimal(el("centroDespesaOdometro").value),
    prioridade: clean(el("centroDespesaPrioridade").value) || "NORMAL",
    descricao: clean(el("centroDespesaDescricao").value),
    observacao: clean(el("centroDespesaObservacao").value),
  };
  if (!payload.veiculo) {
    notify("Selecione o veiculo da despesa.", true);
    return;
  }
  if (!payload.descricao || payload.valor <= 0) {
    notify("Informe descricao e valor da despesa.", true);
    return;
  }
  if (["DATA", "DATA_KM"].includes(payload.controle_tipo) && !payload.data_vencimento) {
    notify("Informe a data de vencimento do controle.", true);
    return;
  }
  if (["KM", "DATA_KM"].includes(payload.controle_tipo) && payload.km_vencimento <= 0) {
    notify("Informe o KM limite do controle.", true);
    return;
  }
  try {
    await apiRequest("/centro-custos/despesas-veiculo", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    notify("Despesa por veiculo registrada.");
    clearCentroCustosDespesaVeiculoForm();
    await loadCentroCustosResumo();
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadRelatorios() {
  try {
    state.relatoriosOptions = await apiRequest("/relatorios/options");
    renderRelatoriosOptions();
    await loadRelatoriosProgramacoes();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderRelatoriosOptions() {
  const tipos = (state.relatoriosOptions && state.relatoriosOptions.tipos) || [
    "Detalhe Completo da Rota",
    "Nota Fiscal / Transbordo",
    "Programacoes",
    "Prestacao de Contas",
    "Mortalidades",
    "Ocorrencias por Motorista",
    "Rotina Motorista/Ajudantes",
    "KM de Veiculos",
    "Abastecimentos",
    "Banhos",
    "Despesas",
  ];
  const select = el("relatoriosTipo");
  const current = clean(select.value) || "Detalhe Completo da Rota";
  select.innerHTML = "";
  tipos.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = relatorioTipoDisplay(value);
    select.appendChild(option);
  });
  select.value = tipos.includes(current) ? current : tipos[0] || "";
  renderRelatoriosModeCards(tipos);
  normalizeRelatoriosButtonLabels();
  updateRelatoriosFilterHints();
  renderRelatoriosContextPanel(state.relatoriosResumo, selectedRelatorioProgramacaoInfo());
  updateRelatoriosActionState();
}

function renderRelatoriosModeCards(tipos) {
  const wrap = document.getElementById("relatoriosModeCards");
  if (!wrap) return;
  const current = clean(el("relatoriosTipo").value);
  const priority = [
    "Nota Fiscal / Transbordo",
    "Prestacao de Contas",
    "Mortalidades",
    "KM de Veiculos",
    "Abastecimentos",
    "Banhos",
    "Detalhe Completo da Rota",
    "Programacoes",
  ];
  const ordered = priority.filter((item) => tipos.includes(item));
  wrap.innerHTML = ordered.map((tipo) => {
    const meta = RELATORIOS_MODE_META[tipo] || {title: relatorioTipoDisplay(tipo), desc: "", tag: "Relatorio"};
    return `
      <button type="button" class="relatorios-mode-card ${tipo === current ? "active" : ""}" data-relatorio-mode="${escapeHtml(tipo)}">
        <span>${escapeHtml(meta.tag)}</span>
        <strong>${escapeHtml(meta.title)}</strong>
        <small>${escapeHtml(meta.desc)}</small>
      </button>
    `;
  }).join("");
  wrap.querySelectorAll("[data-relatorio-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      el("relatoriosTipo").value = button.dataset.relatorioMode;
      onRelatoriosTipoChange();
    });
  });
}

function normalizeRelatoriosButtonLabels() {
  const labels = {
    relatoriosResumoButton: "Gerar resumo",
    relatoriosExcelButton: "Exportar Excel",
    relatoriosPdfButton: "Gerar PDF",
    relatoriosPrintProgramacaoButton: "Reimprimir programacao",
    relatoriosPrintPrestacaoButton: "Reimprimir prestacao",
    relatoriosPrintRomaneiosButton: "Reimprimir romaneios",
    relatoriosFinalizarButton: "Finalizar rota",
    relatoriosReabrirButton: "Reabrir rota",
  };
  Object.entries(labels).forEach(([id, label]) => {
    const button = document.getElementById(id);
    if (button) button.textContent = label;
  });
}

function relatorioTipoDisplay(value) {
  const key = clean(value).toUpperCase();
  if (key === "NOTA FISCAL / TRANSBORDO") return "Nota Fiscal / Transbordo";
  if (key === "PROGRAMACOES") return "Planejamentos";
  if (key === "PRESTACAO DE CONTAS") return "Fechamento Operacional";
  if (key === "DESPESAS") return "Custos e Despesas";
  if (key === "MORTALIDADES") return "Mortalidades";
  if (key === "ABASTECIMENTOS") return "Abastecimentos";
  if (key === "BANHOS") return "Banhos";
  return value;
}

async function loadRelatoriosProgramacoes() {
  const params = new URLSearchParams({
    tipo: clean(el("relatoriosTipo").value) || "Programacoes",
    codigo: clean(el("relatoriosCodigo").value),
    nf: clean(el("relatoriosNf").value),
    motorista: clean(el("relatoriosMotorista").value),
    data: clean(el("relatoriosData").value),
    limit: "400",
  });
  try {
    state.relatoriosProgramacoes = await apiRequest(`/relatorios/programacoes?${params.toString()}`);
    renderRelatoriosProgramacoes();
    updateRelatoriosActionState();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderRelatoriosProgramacoes() {
  const select = el("relatoriosProgramacaoSelect");
  const current = clean(select.value);
  select.innerHTML = '<option value="">Selecione</option>';
  state.relatoriosProgramacoes.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.codigo_programacao;
    option.textContent = `${item.codigo_programacao} | NF ${item.nf_numero || "-"} | ${item.motorista || "-"} | ${item.veiculo || "-"} | ${item.status || "-"}`;
    select.appendChild(option);
  });
  if (current && state.relatoriosProgramacoes.some((item) => item.codigo_programacao === current)) {
    select.value = current;
  }
  el("relatoriosInfo").textContent = `${state.relatoriosProgramacoes.length} planejamento(s) encontrado(s).`;
  renderRelatoriosContextPanel(state.relatoriosResumo, selectedRelatorioProgramacaoInfo());
}

function onRelatoriosTipoChange() {
  state.relatoriosResumo = null;
  renderRelatoriosModeCards((state.relatoriosOptions && state.relatoriosOptions.tipos) || []);
  updateRelatoriosFilterHints();
  renderRelatoriosResumo();
  loadRelatoriosProgramacoes();
  updateRelatoriosActionState();
}

function updateRelatoriosFilterHints() {
  const tipo = clean(el("relatoriosTipo").value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  const codigo = el("relatoriosCodigo");
  const nf = el("relatoriosNf");
  const motorista = el("relatoriosMotorista");
  const data = el("relatoriosData");
  codigo.placeholder = "Codigo da programacao";
  nf.placeholder = "Numero da NF";
  motorista.placeholder = "Motorista";
  data.placeholder = "AAAA-MM-DD ou DD/MM/AAAA";
  if (tipo.includes("NOTA FISCAL") || tipo.includes("TRANSBORDO")) {
    codigo.placeholder = "Opcional: programacao";
    nf.placeholder = "Informe a NF para consolidar";
    motorista.placeholder = "Opcional: motorista";
    el("relatoriosInfo").textContent = "Busque pela NF para ver carga raiz, transbordos, transferencias, entregas, fotos, despesas e recebimentos.";
  } else if (tipo.includes("MORTALIDADE") || tipo.includes("OCORRENCIA")) {
    codigo.placeholder = "Opcional: programacao ou NF";
    motorista.placeholder = "Filtrar motorista";
    el("relatoriosInfo").textContent = "Analise mortalidades por motorista, rota e programacao.";
  } else if (tipo.includes("ABASTEC") || tipo.includes("BANHO")) {
    codigo.placeholder = "Opcional: programacao ou NF";
    motorista.placeholder = "Opcional: motorista";
    el("relatoriosInfo").textContent = "Filtre os lancamentos do app por programacao, NF, motorista ou data.";
  } else if (tipo.includes("KM DE VEICULOS")) {
    codigo.placeholder = "Opcional";
    motorista.placeholder = "Opcional";
    el("relatoriosInfo").textContent = "Compare KM rodado e media por veiculo.";
  }
  renderRelatoriosContextPanel(state.relatoriosResumo, selectedRelatorioProgramacaoInfo());
}

function clearRelatoriosFilters() {
  el("relatoriosTipo").value = "Detalhe Completo da Rota";
  el("relatoriosCodigo").value = "";
  el("relatoriosNf").value = "";
  el("relatoriosMotorista").value = "";
  el("relatoriosData").value = "";
  el("relatoriosProgramacaoSelect").value = "";
  state.relatoriosResumo = null;
  renderRelatoriosResumo();
  loadRelatoriosProgramacoes();
}

function relatorioTipoExigeProgramacao() {
  const tipo = clean(el("relatoriosTipo").value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  if (relatorioTipoPorNotaFiscal()) return false;
  return tipo.includes("PROGRAMACOES") || tipo.includes("PRESTACAO") || tipo.includes("DETALHE COMPLETO");
}

function relatorioTipoPorNotaFiscal() {
  const tipo = clean(el("relatoriosTipo").value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  return tipo.includes("NOTA FISCAL") || tipo.includes("TRANSBORDO");
}

function relatoriosResumoParams() {
  return new URLSearchParams({
    tipo: clean(el("relatoriosTipo").value) || "Programacoes",
    programacao: clean(el("relatoriosProgramacaoSelect").value),
    codigo: clean(el("relatoriosCodigo").value),
    nf: clean(el("relatoriosNf").value),
    motorista: clean(el("relatoriosMotorista").value),
    data: clean(el("relatoriosData").value),
    show_recebimentos: el("relatoriosShowRecebimentos").checked ? "true" : "false",
    show_despesas: el("relatoriosShowDespesas").checked ? "true" : "false",
  });
}

async function loadRelatoriosResumo() {
  if (relatorioTipoExigeProgramacao() && !clean(el("relatoriosProgramacaoSelect").value)) {
    notify("Selecione um planejamento.", true);
    return;
  }
  if (relatorioTipoPorNotaFiscal() && !clean(el("relatoriosNf").value)) {
    notify("Informe o numero da nota fiscal.", true);
    return;
  }
  try {
    state.relatoriosResumo = await apiRequest(`/relatorios/resumo?${relatoriosResumoParams().toString()}`);
    renderRelatoriosResumo();
  } catch (error) {
    notify(error.message, true);
  }
}

function refreshRelatoriosResumoIfLoaded() {
  if (state.relatoriosResumo) loadRelatoriosResumo();
}

function renderRelatoriosResumo() {
  const data = state.relatoriosResumo;
  if (!data) {
    el("relatoriosStatus").textContent = "-";
    el("relatoriosMetrics").innerHTML = "";
    el("relatoriosChart").innerHTML = '<p class="empty-chart">Sem dados para grafico.</p>';
    el("relatoriosText").textContent = "Gere um resumo para visualizar.";
    el("relatoriosTableHead").innerHTML = "";
    el("relatoriosRows").innerHTML = "";
    el("relatoriosRows").appendChild(emptyRow(1, "Sem dados."));
    el("relatoriosSectionsPanel").innerHTML = "";
    el("relatoriosTableSearch").value = "";
    const tableInfo = document.getElementById("relatoriosTableInfo");
    if (tableInfo) tableInfo.textContent = "Nenhum relatorio carregado.";
    renderRelatoriosContextPanel(null, null);
    updateRelatoriosActionState();
    return;
  }
  el("relatoriosStatus").textContent = data.status || "-";
  el("relatoriosText").textContent = data.text || "";
  el("relatoriosChartTitle").textContent = data.tipo || "Distribuicao visual";
  el("relatoriosMetrics").innerHTML = (data.kpis || []).map((item) => `
    <article class="metric relatorios-metric ${relatoriosKpiClass(item)}">
      <span>${escapeHtml(item.label)}</span>
      <strong>${escapeHtml(item.value)}</strong>
    </article>
  `).join("");
  const filteredRows = filterRelatoriosRows(data.columns || [], data.rows || []);
  const tableInfo = document.getElementById("relatoriosTableInfo");
  const selectedInfo = selectedRelatorioProgramacaoInfo();
  if (tableInfo) {
    const mainCount = (data.rows || []).length;
    const filteredCount = filteredRows.length;
    const sectionCount = (data.sections || []).reduce((acc, section) => acc + ((section.rows || []).length), 0);
    tableInfo.textContent = `${filteredCount}/${mainCount} linha(s) principais | ${sectionCount} linha(s) em detalhes.`;
  }
  el("relatoriosInfo").textContent = relatoriosInfoText(data, selectedInfo, filteredRows.length);
  renderRelatoriosContextPanel(data, selectedInfo);
  renderRelatoriosChart(data.chart || [], data.tipo || "");
  renderRelatoriosRows(data.columns || [], filteredRows);
  renderRelatoriosSections(data.sections || []);
  updateRelatoriosActionState();
}

function renderRelatoriosContextPanel(data, selectedInfo) {
  const panel = document.getElementById("relatoriosContextPanel");
  if (!panel) return;
  const tipo = clean(el("relatoriosTipo").value) || "Relatorio";
  const meta = RELATORIOS_MODE_META[tipo] || {title: relatorioTipoDisplay(tipo), desc: "", tag: "Relatorio"};
  const nf = clean(el("relatoriosNf").value) || selectedInfo?.nf_numero || "-";
  const codigo = selectedInfo?.codigo_programacao || clean(el("relatoriosProgramacaoSelect").value) || clean(el("relatoriosCodigo").value) || "-";
  const motorista = selectedInfo?.motorista || clean(el("relatoriosMotorista").value) || "-";
  const veiculo = selectedInfo?.veiculo || "-";
  const status = selectedInfo?.status || "-";
  const prestacao = selectedInfo?.prestacao_status || "-";
  const kpis = (data?.kpis || []).slice(0, 4);
  panel.innerHTML = `
    <div class="relatorios-context-main">
      <span>${escapeHtml(meta.tag)}</span>
      <strong>${escapeHtml(meta.title)}</strong>
      <small>${escapeHtml(meta.desc || "Selecione filtros e gere o relatorio para ver o contexto operacional.")}</small>
    </div>
    <div class="relatorios-context-grid">
      <div><span>Programacao</span><strong>${escapeHtml(codigo)}</strong></div>
      <div><span>Nota fiscal</span><strong>${escapeHtml(nf)}</strong></div>
      <div><span>Motorista</span><strong>${escapeHtml(motorista)}</strong></div>
      <div><span>Veiculo</span><strong>${escapeHtml(veiculo)}</strong></div>
      <div><span>Status</span><strong>${escapeHtml(status)}</strong></div>
      <div><span>Prestacao</span><strong>${escapeHtml(prestacao)}</strong></div>
    </div>
    <div class="relatorios-context-kpis">
      ${kpis.length ? kpis.map((item) => `<div><span>${escapeHtml(item.label)}</span><strong>${escapeHtml(item.value)}</strong></div>`).join("") : "<div><span>Resumo</span><strong>Aguardando geracao</strong></div>"}
    </div>
  `;
}

function selectedRelatorioProgramacaoInfo() {
  const codigo = selectedRelatorioProgramacao();
  if (!codigo) return null;
  return (state.relatoriosProgramacoes || []).find((item) => item.codigo_programacao === codigo) || null;
}

function relatoriosInfoText(data, selectedInfo, visibleRows) {
  const parts = [data.tipo || "Relatorio"];
  if (data.programacao) parts.push(data.programacao);
  if (selectedInfo?.nf_numero) parts.push(`NF ${selectedInfo.nf_numero}`);
  if (selectedInfo?.motorista) parts.push(`Motorista ${selectedInfo.motorista}`);
  if (selectedInfo?.veiculo) parts.push(`Veiculo ${selectedInfo.veiculo}`);
  if (selectedInfo?.status) parts.push(`Status ${selectedInfo.status}`);
  if (selectedInfo?.prestacao_status) parts.push(`Prestacao ${selectedInfo.prestacao_status}`);
  if (clean(el("relatoriosNf").value) && !selectedInfo?.nf_numero) parts.push(`NF ${clean(el("relatoriosNf").value)}`);
  parts.push(`${visibleRows} linha(s)`);
  return parts.join(" | ");
}

function relatoriosKpiClass(item) {
  const label = String(item?.label || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  const rawValue = String(item?.value || "");
  const numberValue = Number(rawValue.replace(/[^\d,-]/g, "").replace(/\./g, "").replace(",", "."));
  const classes = [];
  if (["LUCRO", "RESULTADO", "MARGEM", "DIFERENCA", "SALDO"].some((needle) => label.includes(needle))) {
    classes.push("is-result");
  }
  if (["MORTALIDADE", "DESPESA", "CUSTO", "DIVERG"].some((needle) => label.includes(needle))) {
    classes.push("is-attention");
  }
  if (Number.isFinite(numberValue) && rawValue.includes("-")) classes.push("is-negative");
  return classes.join(" ");
}

function filterRelatoriosRows(columns, rowsData) {
  const needle = clean(el("relatoriosTableSearch").value).toUpperCase();
  if (!needle) return rowsData;
  return rowsData.filter((row) => columns.some((column) => String(row[column.key] ?? "").toUpperCase().includes(needle)));
}

function renderRelatoriosChart(chart, tipo) {
  const wrap = el("relatoriosChart");
  wrap.innerHTML = "";
  if (!chart.length) {
    wrap.innerHTML = '<p class="empty-chart">Sem dados para grafico.</p>';
    return;
  }
  const maxValue = Math.max(...chart.map((item) => Math.abs(Number(item.value || 0))), 1);
  const moneyChart = ["PRESTACAO", "DESPESAS", "PROGRAMACOES", "DETALHE"].some((needle) => tipo.toUpperCase().includes(needle));
  chart.forEach((item) => {
    const value = Number(item.value || 0);
    const pct = Math.max((Math.abs(value) / maxValue) * 100, 2);
    const row = document.createElement("div");
    row.className = `relatorios-chart-row ${value < 0 ? "is-negative" : ""}`;
    row.innerHTML = `
      <span>${escapeHtml(item.label || "-")}</span>
      <div class="relatorios-chart-track">
        <div class="relatorios-chart-bar" style="width: ${pct}%"></div>
      </div>
      <strong>${escapeHtml(moneyChart ? formatCurrencyBR(value) : formatPlainNumber(value))}</strong>
    `;
    wrap.appendChild(row);
  });
}

function renderRelatoriosSections(sections) {
  const wrap = el("relatoriosSectionsPanel");
  wrap.innerHTML = "";
  if (!sections.length) return;
  sections.forEach((section) => {
    const panel = document.createElement("section");
    panel.className = "panel relatorios-section-panel";
    const columns = section.columns || [];
    const rows = section.rows || [];
    const headHtml = columns.length
      ? `<tr>${columns.map((column) => `<th class="${relatoriosColumnClass(column)}">${escapeHtml(column.label)}</th>`).join("")}</tr>`
      : "";
    const bodyHtml = rows.length
      ? rows.map((row) => `
          <tr>
            ${columns.map((column) => `<td class="${relatoriosCellClass(row[column.key], column)}">${escapeHtml(formatRelatoriosCell(row[column.key], column.kind))}</td>`).join("")}
          </tr>
        `).join("")
      : `<tr><td colspan="${Math.max(columns.length, 1)}" class="empty-row">Sem dados.</td></tr>`;
    panel.innerHTML = `
      <div class="panel-header">
        <h2>${escapeHtml(section.title || "Detalhes")}</h2>
      </div>
      <div class="table-wrap">
        <table class="wide-table relatorios-table relatorios-section-table">
          <thead>${headHtml}</thead>
          <tbody>${bodyHtml}</tbody>
        </table>
      </div>
    `;
    wrap.appendChild(panel);
  });
}

function renderRelatoriosRows(columns, rowsData) {
  const head = el("relatoriosTableHead");
  const body = el("relatoriosRows");
  head.innerHTML = "";
  body.innerHTML = "";
  if (!columns.length) {
    body.appendChild(emptyRow(1, "Sem colunas."));
    return;
  }
  const trHead = document.createElement("tr");
  trHead.innerHTML = columns.map((column) => `<th class="${relatoriosColumnClass(column)}">${escapeHtml(column.label)}</th>`).join("");
  head.appendChild(trHead);
  if (!rowsData.length) {
    body.appendChild(emptyRow(columns.length, "Sem dados no relatorio."));
    return;
  }
  rowsData.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = columns.map((column) => {
      const value = row[column.key];
      return `<td class="${relatoriosCellClass(value, column)}">${escapeHtml(formatRelatoriosCell(value, column.kind))}</td>`;
    }).join("");
    body.appendChild(tr);
  });
}

function relatoriosColumnClass(column) {
  if (["money", "number"].includes(column.kind)) return "is-numeric";
  return "";
}

function relatoriosCellClass(value, column) {
  const classes = [];
  if (["money", "number"].includes(column.kind)) classes.push("is-numeric");
  const key = String(column.key || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  const text = String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  const numeric = Number(String(value ?? "").replace(/[^\d,-]/g, "").replace(/\./g, "").replace(",", "."));
  if (["STATUS", "PRESTACAO", "OPERACAO"].some((needle) => key.includes(needle))) {
    classes.push("is-tag");
    if (["CANCEL", "RECUS", "NEGAT"].some((needle) => text.includes(needle))) classes.push("is-danger-tag");
    else if (["FECHADA", "FINALIZADA", "ENTREG"].some((needle) => text.includes(needle))) classes.push("is-success-tag");
    else if (["TRANSBORDO", "EM_ROTA", "EM ROTA", "ATIVA"].some((needle) => text.includes(needle))) classes.push("is-info-tag");
    else if (["PENDENTE", "ABERTA", "AGUARD"].some((needle) => text.includes(needle))) classes.push("is-warning-tag");
  }
  if (Number.isFinite(numeric)) {
    if (numeric < 0) classes.push("is-negative");
    if (numeric > 0 && ["LUCRO", "RESULTADO", "MARGEM"].some((needle) => key.includes(needle))) classes.push("is-positive");
    if (numeric > 0 && ["MORT", "DIVERG", "DESPESA", "CUSTO"].some((needle) => key.includes(needle))) classes.push("is-warning");
  }
  return classes.join(" ");
}

function formatRelatoriosCell(value, kind) {
  if (kind === "money") return formatCurrencyBR(value);
  if (kind === "number") return formatPlainNumber(value);
  return value === null || value === undefined || value === "" ? "-" : String(value);
}

function selectedRelatorioProgramacao() {
  return clean(el("relatoriosProgramacaoSelect").value);
}

async function downloadRelatoriosExcel() {
  const codigo = selectedRelatorioProgramacao();
  try {
    const useProgramacaoWorkbook = codigo && relatorioTipoExigeProgramacao();
    const result = useProgramacaoWorkbook
      ? await apiBlobRequest(`/relatorios/${encodeURIComponent(codigo)}/exportar-excel`)
      : await apiBlobRequest(`/relatorios/exportar-excel?${relatoriosResumoParams().toString()}`);
    downloadBlob(result.blob, result.filename || `RELATORIO_${codigo || "OPERACIONAL"}.xlsx`);
    notify("Excel gerado.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadRelatoriosPdf() {
  if (relatorioTipoExigeProgramacao() && !selectedRelatorioProgramacao()) {
    notify("Selecione um planejamento.", true);
    return;
  }
  try {
    const tipo = clean(el("relatoriosTipo").value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
    const codigo = selectedRelatorioProgramacao();
    if (tipo.includes("PROGRAMACOES") && codigo) {
      const result = await apiBlobRequest(`/programacao/${encodeURIComponent(codigo)}/pdf?reimpressao=1`);
      downloadBlob(result.blob, result.filename || `PROGRAMACAO_${codigo}.pdf`);
      notify("Reimpressao do planejamento gerada.");
      return;
    }
    if (tipo.includes("PRESTACAO") && codigo) {
      const result = await apiBlobRequest(`/despesas/${encodeURIComponent(codigo)}/pdf?reimpressao=1`);
      downloadBlob(result.blob, result.filename || `PRESTACAO_${codigo}.pdf`);
      notify("Reimpressao da prestacao gerada.");
      return;
    }
    const result = await apiBlobRequest(`/relatorios/pdf?${relatoriosResumoParams().toString()}`);
    downloadBlob(result.blob, result.filename || "RELATORIO.pdf");
    notify("PDF gerado.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadRelatoriosDocumento(tipoDocumento) {
  const codigo = selectedRelatorioProgramacao();
  if (!codigo) {
    notify("Selecione um planejamento.", true);
    return;
  }
  const docs = {
    programacao: {
      path: `/programacao/${encodeURIComponent(codigo)}/pdf?reimpressao=1`,
      filename: `PROGRAMACAO_${codigo}.pdf`,
      message: "Reimpressao do planejamento gerada.",
    },
    prestacao: {
      path: `/despesas/${encodeURIComponent(codigo)}/pdf?reimpressao=1`,
      filename: `PRESTACAO_${codigo}.pdf`,
      message: "Reimpressao da prestacao gerada.",
    },
    romaneios: {
      path: `/programacao/${encodeURIComponent(codigo)}/romaneios-pdf?reimpressao=1`,
      filename: `ROMANEIOS_${codigo}.pdf`,
      message: "Reimpressao dos romaneios gerada.",
    },
  };
  const doc = docs[tipoDocumento];
  if (!doc) return;
  try {
    const result = await apiBlobRequest(doc.path);
    downloadBlob(result.blob, result.filename || doc.filename);
    notify(doc.message);
  } catch (error) {
    notify(error.message, true);
  }
}

async function finalizarRelatoriosRota() {
  const codigo = selectedRelatorioProgramacao();
  if (!codigo) {
    notify("Selecione um planejamento.", true);
    return;
  }
  if (!window.confirm(`Finalizar a rota ${codigo}?`)) return;
  try {
    await apiRequest(`/relatorios/${encodeURIComponent(codigo)}/finalizar-rota`, {method: "POST"});
    notify(`Rota finalizada: ${codigo}`);
    await loadRelatoriosProgramacoes();
    await loadRelatoriosResumo();
  } catch (error) {
    notify(error.message, true);
  }
}

async function reabrirRelatoriosRota() {
  const codigo = selectedRelatorioProgramacao();
  if (!codigo) {
    notify("Selecione um planejamento.", true);
    return;
  }
  if (!window.confirm(`Reabrir a rota ${codigo}?`)) return;
  try {
    await apiRequest(`/relatorios/${encodeURIComponent(codigo)}/reabrir-rota`, {method: "POST"});
    notify(`Rota reaberta: ${codigo}`);
    await loadRelatoriosProgramacoes();
    await loadRelatoriosResumo();
  } catch (error) {
    notify(error.message, true);
  }
}

function updateRelatoriosActionState() {
  const needsProgramacao = relatorioTipoExigeProgramacao();
  const needsNf = relatorioTipoPorNotaFiscal();
  const hasProgramacao = Boolean(selectedRelatorioProgramacao());
  const hasNf = Boolean(clean(el("relatoriosNf").value));
  el("relatoriosProgramacaoSelect").disabled = false;
  el("relatoriosExcelButton").disabled = (needsProgramacao && !hasProgramacao) || (needsNf && !hasNf);
  el("relatoriosPrintProgramacaoButton").disabled = !hasProgramacao;
  el("relatoriosPrintPrestacaoButton").disabled = !hasProgramacao;
  el("relatoriosPrintRomaneiosButton").disabled = !hasProgramacao;
  el("relatoriosFinalizarButton").disabled = !hasProgramacao;
  el("relatoriosReabrirButton").disabled = !hasProgramacao;
  el("relatoriosResumoButton").disabled = (needsProgramacao && !hasProgramacao) || (needsNf && !hasNf);
  el("relatoriosPdfButton").disabled = (needsProgramacao && !hasProgramacao) || (needsNf && !hasNf);
}

async function loadSystemTools() {
  try {
    const [overview, diariasConfig] = await Promise.all([
      apiRequest("/system-tools/overview"),
      apiRequest("/system-tools/diarias"),
    ]);
    state.systemTools = {...overview, diariasConfig};
    renderBackupView();
    renderFerramentasView();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderBackupView() {
  const backups = state.systemTools.backups || [];
  el("backupInfo").textContent = `${backups.length} backup(s) disponivel(is).`;
  const rows = el("backupRows");
  rows.innerHTML = "";
  if (!backups.length) {
    rows.appendChild(emptyRow(4, "Nenhum backup encontrado."));
    return;
  }
  backups.forEach((backup) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(backup.arquivo)}</td>
      <td>${escapeHtml(formatPlainNumber(backup.tamanho_kb || 0))}</td>
      <td>${escapeHtml(formatDate(backup.data_criacao))}</td>
      <td>
        <div class="actions">
          <button type="button" class="secondary" data-backup-download="${escapeHtml(backup.arquivo)}">BAIXAR</button>
          <button type="button" class="danger" data-backup-restore="${escapeHtml(backup.arquivo)}">RESTAURAR</button>
        </div>
      </td>
    `;
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-backup-download]").forEach((button) => {
    button.addEventListener("click", () => downloadSavedBackup(button.dataset.backupDownload));
  });
  rows.querySelectorAll("[data-backup-restore]").forEach((button) => {
    button.addEventListener("click", () => restoreSavedBackup(button.dataset.backupRestore));
  });
}

function renderFerramentasView() {
  const info = state.systemTools.info || {};
  const logs = state.systemTools.logs || [];
  const backups = state.systemTools.backups || [];
  renderDiariasConfigForm();
  el("ferramentasInfo").textContent = `${logs.length} log(s) carregado(s).`;
  const metrics = [
    ["Backups", backups.length],
    ["Logs", logs.length],
    ["Banco", `${formatPlainNumber(info.tamanho_banco_kb || 0)} KB`],
    ["Ultima acao", info.ultima_acao_em ? formatDate(info.ultima_acao_em) : "-"],
  ];
  el("ferramentasMetrics").innerHTML = metrics.map(([label, value]) => `
    <article class="metric ferramentas-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `).join("");

  const registros = info.registros_por_tabela || {};
  const infoItems = [
    ["Data/Hora Atual", formatDate(info.data_hora_atual)],
    ["Tamanho do Banco", `${formatPlainNumber(info.tamanho_banco_kb || 0)} KB`],
    ["Diretorio de Backup", info.backup_dir || "-"],
    ["Banco", info.database_path || "-"],
    ...Object.entries(registros).sort(([a], [b]) => a.localeCompare(b)).map(([table, count]) => [`Tabela ${table}`, `${count} registro(s)`]),
  ];
  el("ferramentasSystemInfo").innerHTML = infoItems.map(([label, value]) => `
    <div class="ferramentas-info-item">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join("");

  const logRows = el("ferramentasLogRows");
  logRows.innerHTML = "";
  if (!logs.length) {
    logRows.appendChild(emptyRow(6, "Sem logs do sistema."));
    return;
  }
  logs.forEach((log) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(log.descricao || "-")}</td>
      <td>${escapeHtml(log.tipo_acao || "-")}</td>
      <td>${escapeHtml(log.usuario || "-")}</td>
      <td>${escapeHtml(log.status || "-")}</td>
      <td>${escapeHtml(formatDate(log.executado_em))}</td>
      <td>${escapeHtml(log.resultado || "-")}</td>
    `;
    logRows.appendChild(tr);
  });
}

function renderDiariasConfigForm() {
  const itens = (state.systemTools.diariasConfig && state.systemTools.diariasConfig.itens) || [];
  const byLocal = Object.fromEntries(itens.map((item) => [clean(item.local_rota).toUpperCase(), item]));
  const serra = byLocal.SERRA || {};
  const sertao = byLocal.SERTAO || {};
  el("diariaSerraMotorista").value = formatMoneyInput(serra.motorista || 0);
  el("diariaSerraAjudante").value = formatMoneyInput(serra.ajudante || 0);
  el("diariaSertaoMotorista").value = formatMoneyInput(sertao.motorista || 0);
  el("diariaSertaoAjudante").value = formatMoneyInput(sertao.ajudante || 0);
}

async function saveDiariasConfig(event) {
  event.preventDefault();
  const payload = {
    serra_motorista: parseDecimal(el("diariaSerraMotorista").value),
    serra_ajudante: parseDecimal(el("diariaSerraAjudante").value),
    sertao_motorista: parseDecimal(el("diariaSertaoMotorista").value),
    sertao_ajudante: parseDecimal(el("diariaSertaoAjudante").value),
  };
  try {
    el("diariasConfigSaveButton").disabled = true;
    state.systemTools.diariasConfig = await apiRequest("/system-tools/diarias", {
      method: "PUT",
      body: JSON.stringify(payload),
    });
    notify("Valores de diarias salvos.");
    await loadSystemTools();
  } catch (error) {
    notify(error.message, true);
  } finally {
    el("diariasConfigSaveButton").disabled = false;
  }
}

async function createSystemBackup() {
  try {
    const result = await apiRequest("/system-tools/backups", {method: "POST"});
    notify(`Backup criado: ${result.arquivo}`);
    await loadSystemTools();
  } catch (error) {
    notify(error.message, true);
  }
}

async function exportBackupVendas() {
  try {
    const result = await apiBlobRequest("/system-tools/vendas-importadas/export");
    downloadBlob(result.blob, result.filename || "VENDAS_IMPORTADAS.xlsx");
    notify("Pedidos importados exportados.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadMigrationPackage() {
  try {
    notify("Preparando banco e fotos para migracao...");
    const result = await apiBlobRequest("/system-tools/migration/export/download");
    downloadBlob(result.blob, result.filename || "rotahub_migration.zip");
    notify("Pacote de migracao baixado.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadSavedBackup(filename) {
  try {
    const result = await apiBlobRequest(`/system-tools/backups/${encodeURIComponent(filename)}/download`);
    downloadBlob(result.blob, result.filename || filename);
    notify("Backup baixado.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadSystemInstaller() {
  try {
    notify("Preparando download do instalador...");
    const result = await apiBlobRequest("/system-tools/installer/download");
    downloadBlob(result.blob, result.filename || "RotaHubDesktop_Setup.exe");
    notify("Download do instalador iniciado.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function restoreSavedBackup(filename) {
  if (!window.confirm(`Restaurar o backup ${filename}? Esta acao substitui o banco atual.`)) return;
  try {
    const result = await apiRequest(`/system-tools/backups/${encodeURIComponent(filename)}/restore`, {
      method: "POST",
      body: JSON.stringify({confirmar: true}),
    });
    notify(result.mensagem || "Backup restaurado.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function restoreBackupUpload(event) {
  event.preventDefault();
  const fileInput = el("backupRestoreFile");
  const confirmed = el("backupRestoreConfirm").checked;
  if (!fileInput.files.length) {
    notify("Selecione um arquivo .db.", true);
    return;
  }
  if (!confirmed) {
    notify("Marque a confirmacao antes de restaurar.", true);
    return;
  }
  if (!window.confirm("Restaurar este arquivo .db e substituir o banco atual?")) return;
  try {
    const body = new FormData();
    body.set("file", fileInput.files[0]);
    const result = await apiRequest("/system-tools/backups/upload-restore?confirmar=true", {
      method: "POST",
      body,
    });
    event.currentTarget.reset();
    notify(result.mensagem || "Backup restaurado.");
  } catch (error) {
    notify(error.message, true);
  }
}

async function checkSystemIntegrity() {
  try {
    const result = await apiRequest("/system-tools/integridade");
    notify(result.ok ? "Banco de dados integro." : `Problemas encontrados: ${result.integridade}`, !result.ok);
    await loadSystemTools();
  } catch (error) {
    notify(error.message, true);
  }
}

async function clearSystemLogs() {
  const days = Math.max(parseInt(clean(el("ferramentasLogDays").value) || "30", 10) || 30, 1);
  if (!window.confirm(`Limpar logs com mais de ${days} dias?`)) return;
  try {
    const result = await apiRequest(`/system-tools/logs?dias=${encodeURIComponent(days)}`, {method: "DELETE"});
    notify(`Logs removidos: ${result.linhas_deletadas || 0}`);
    await loadSystemTools();
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadPermissoes() {
  try {
    const selectedBefore = state.permissoes.selectedUserId;
    const overview = await apiRequest("/permissoes/overview?include_inactive=true");
    const selectedExists = overview.usuarios.some((user) => user.id === selectedBefore);
    const selectedUserId = selectedExists ? selectedBefore : (overview.usuarios[0] ? overview.usuarios[0].id : null);
    state.permissoes = {
      usuarios: overview.usuarios || [],
      permissoes: overview.permissoes || [],
      modulos: overview.modulos || [],
      selectedUserId,
      granted: [],
    };
    if (selectedUserId) {
      state.permissoes.granted = await apiRequest(`/permissoes/usuarios/${selectedUserId}`);
    }
    renderPermissoesView();
  } catch (error) {
    notify(error.message, true);
  }
}

function currentPermissoesUser() {
  return state.permissoes.usuarios.find((user) => user.id === state.permissoes.selectedUserId) || null;
}

function renderPermissoesView() {
  const data = state.permissoes;
  const users = data.usuarios || [];
  const permissions = data.permissoes || [];
  const granted = data.granted || [];
  const selectedUser = currentPermissoesUser();
  el("permissoesInfo").textContent = `${users.length} usuario(s), ${permissions.length} permissao(oes) cadastrada(s).`;
  el("permissoesSelectedUser").textContent = selectedUser
    ? `${selectedUser.nome || selectedUser.username} - ${granted.length} permissao(oes) concedida(s).`
    : "Selecione um usuario.";

  const moduloFilter = el("permissoesModuloFilter");
  const selectedModule = moduloFilter.value;
  moduloFilter.innerHTML = `<option value="">TODOS</option>${(data.modulos || []).map((modulo) => (
    `<option value="${escapeHtml(modulo)}">${escapeHtml(modulo.toUpperCase())}</option>`
  )).join("")}`;
  moduloFilter.value = selectedModule;

  if (selectedUser && ["ADMIN", "GERENTE", "OPERADOR", "VISUALIZADOR"].includes(selectedUser.permissoes)) {
    el("permissoesPerfilSelect").value = selectedUser.permissoes;
  }

  const userRows = el("permissoesUserRows");
  userRows.innerHTML = "";
  if (!users.length) {
    userRows.appendChild(emptyRow(6, "Sem usuarios."));
  } else {
    users.forEach((user) => {
      const tr = document.createElement("tr");
      tr.className = user.id === data.selectedUserId ? "selected-row" : "";
      tr.innerHTML = `
        <td>${user.id}</td>
        <td>
          <strong>${escapeHtml(user.nome || user.username)}</strong>
          <span>${escapeHtml(user.username || "-")}</span>
        </td>
        <td>${escapeHtml(user.permissoes || "-")}</td>
        <td>${statusPill(Boolean(user.is_active))}</td>
        <td>${escapeHtml(user.granted_count || 0)}</td>
        <td><button type="button" class="secondary" data-permission-user="${user.id}">Selecionar</button></td>
      `;
      userRows.appendChild(tr);
    });
  }

  const grantedIds = new Set(granted.map((permission) => permission.id));
  const availableRows = el("permissoesAvailableRows");
  availableRows.innerHTML = "";
  const visiblePermissions = permissions.filter((permission) => (
    !selectedModule || permission.modulo === selectedModule
  ));
  if (!visiblePermissions.length) {
    availableRows.appendChild(emptyRow(5, "Nenhuma permissao encontrada."));
  } else {
    visiblePermissions.forEach((permission) => {
      const alreadyGranted = grantedIds.has(permission.id);
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(permission.modulo || "-")}</td>
        <td>${escapeHtml(permission.nome || "-")}</td>
        <td>${escapeHtml(permission.descricao || "-")}</td>
        <td>${permission.ativo ? "ATIVA" : "INATIVA"}</td>
        <td>
          <button type="button" class="primary" data-permission-grant="${permission.id}" ${!selectedUser || alreadyGranted ? "disabled" : ""}>
            ${alreadyGranted ? "CONCEDIDA" : "CONCEDER"}
          </button>
        </td>
      `;
      availableRows.appendChild(tr);
    });
  }

  const grantedRows = el("permissoesGrantedRows");
  grantedRows.innerHTML = "";
  if (!selectedUser) {
    grantedRows.appendChild(emptyRow(6, "Selecione um usuario."));
  } else if (!granted.length) {
    grantedRows.appendChild(emptyRow(6, "Nenhuma permissao concedida."));
  } else {
    granted.forEach((permission) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(permission.modulo || "-")}</td>
        <td>${escapeHtml(permission.nome || "-")}</td>
        <td>${escapeHtml(permission.descricao || "-")}</td>
        <td>${escapeHtml(formatDate(permission.concedida_em))}</td>
        <td>${escapeHtml(permission.concedida_por || "-")}</td>
        <td><button type="button" class="danger" data-permission-revoke="${permission.id}">REVOGAR</button></td>
      `;
      grantedRows.appendChild(tr);
    });
  }

  userRows.querySelectorAll("[data-permission-user]").forEach((button) => {
    button.addEventListener("click", () => selectPermissoesUser(Number(button.dataset.permissionUser)));
  });
  availableRows.querySelectorAll("[data-permission-grant]").forEach((button) => {
    button.addEventListener("click", () => grantPermission(Number(button.dataset.permissionGrant)));
  });
  grantedRows.querySelectorAll("[data-permission-revoke]").forEach((button) => {
    button.addEventListener("click", () => revokePermission(Number(button.dataset.permissionRevoke)));
  });
}

async function selectPermissoesUser(userId) {
  state.permissoes.selectedUserId = userId;
  try {
    state.permissoes.granted = await apiRequest(`/permissoes/usuarios/${userId}`);
    renderPermissoesView();
  } catch (error) {
    notify(error.message, true);
  }
}

async function grantPermission(permissionId) {
  const selectedUser = currentPermissoesUser();
  if (!selectedUser) {
    notify("Selecione um usuario.", true);
    return;
  }
  try {
    await apiRequest(`/permissoes/usuarios/${selectedUser.id}/conceder`, {
      method: "POST",
      body: JSON.stringify({permissao_id: permissionId}),
    });
    notify("Permissao concedida.");
    await loadPermissoes();
  } catch (error) {
    notify(error.message, true);
  }
}

async function revokePermission(permissionId) {
  const selectedUser = currentPermissoesUser();
  if (!selectedUser) {
    notify("Selecione um usuario.", true);
    return;
  }
  if (!window.confirm(`Revogar permissao do usuario ${selectedUser.username}?`)) return;
  try {
    await apiRequest(`/permissoes/usuarios/${selectedUser.id}/${permissionId}`, {method: "DELETE"});
    notify("Permissao revogada.");
    await loadPermissoes();
  } catch (error) {
    notify(error.message, true);
  }
}

async function applyPermissionProfile() {
  const selectedUser = currentPermissoesUser();
  const perfil = clean(el("permissoesPerfilSelect").value) || "OPERADOR";
  if (!selectedUser) {
    notify("Selecione um usuario.", true);
    return;
  }
  if (!window.confirm(`Aplicar perfil ${perfil} para ${selectedUser.username}? As permissoes atuais serao substituidas.`)) return;
  try {
    const result = await apiRequest(`/permissoes/usuarios/${selectedUser.id}/perfil`, {
      method: "POST",
      body: JSON.stringify({perfil}),
    });
    notify(`Perfil aplicado: ${result.permissoes_atribuidas || 0} permissao(oes).`);
    await loadPermissoes();
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadSaasAdmin() {
  try {
    const selected = Number(state.saasAdmin.selectedCompanyId || 0);
    const query = selected ? `?company_id=${encodeURIComponent(selected)}` : "";
    const data = await apiRequest(`/saas-admin/dashboard${query}`);
    const selectedCompanyId = Number((data.company || {}).id || selected || 0) || null;
    state.saasAdmin = {
      companies: data.companies || [],
      company: data.company || null,
      subscription: data.subscription || null,
      usage: data.usage || {},
      plans: data.plans || [],
      payments: data.payments || [],
      audit_logs: data.audit_logs || [],
      features: data.features || null,
      selectedCompanyId,
    };
    renderSaasAdminView();
  } catch (error) {
    notify(error.message, true);
  }
}

function currentSaasCompanyId() {
  return Number(state.saasAdmin.selectedCompanyId || (state.saasAdmin.company || {}).id || 0) || null;
}

function renderSaasAdminView() {
  const data = state.saasAdmin;
  const company = data.company || {};
  const subscription = data.subscription || {};
  const usage = data.usage || {};
  const vehicles = usage.vehicles || {};
  const companies = data.companies || [];
  const plans = data.plans || [];
  const payments = data.payments || [];
  const auditLogs = data.audit_logs || [];
  const selectedCompanyId = currentSaasCompanyId();

  el("saasAdminInfo").textContent = company.id
    ? `${company.name || "-"} - ${company.status || "-"}`
    : "Nenhuma empresa carregada.";

  el("saasAdminCompanySelect").innerHTML = companies.map((item) => (
    `<option value="${item.id}">${escapeHtml(item.name || item.code || `Empresa ${item.id}`)} (${escapeHtml(item.status || "-")})</option>`
  )).join("");
  if (selectedCompanyId) {
    el("saasAdminCompanySelect").value = String(selectedCompanyId);
  }

  const limit = vehicles.vehicle_limit;
  const limitText = limit === null || limit === undefined ? "sem limite" : String(limit);
  const metrics = [
    ["Empresa", company.name || "-", company.status || "-"],
    ["Plano", subscription.plan_name || "-", subscription.status || "-"],
    ["Veiculos", `${vehicles.vehicle_count || 0} / ${limitText}`, "uso atual"],
    ["Cobranca", subscription.next_due_date || "-", "vencimento"],
  ];
  el("saasAdminMetrics").innerHTML = metrics.map(([label, value, detail]) => `
    <article class="metric saas-admin-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <em>${escapeHtml(detail || "")}</em>
    </article>
  `).join("");

  el("saasAdminPlanSelect").innerHTML = plans.map((plan) => (
    `<option value="${escapeHtml(plan.code || "")}">${escapeHtml(plan.name || plan.code || "-")} (${escapeHtml(plan.code || "-")})</option>`
  )).join("");
  if (subscription.plan_code && plans.some((plan) => plan.code === subscription.plan_code)) {
    el("saasAdminPlanSelect").value = subscription.plan_code;
  }

  const usageRows = [
    ["Veiculos", `${vehicles.vehicle_count || 0} / ${limitText}`],
    ["Usuarios", usage.users || 0],
    ["Motoristas", usage.motoristas || 0],
    ["Vendedores", usage.vendedores || 0],
    ["Clientes", usage.clientes || 0],
    ["Programacoes", usage.programacoes || 0],
  ];
  const usageBody = el("saasAdminUsageRows");
  usageBody.innerHTML = usageRows.map(([label, value]) => `
    <tr>
      <td>${escapeHtml(label)}</td>
      <td>${escapeHtml(value)}</td>
    </tr>
  `).join("");

  renderSaasPlanMatrix(plans, subscription.plan_code);
  renderSaasFeatures(data.features);
  renderSaasPayments(payments);
  renderSaasAudit(auditLogs);
}

function renderSaasPlanMatrix(plans, currentPlanCode) {
  const wrap = el("saasAdminPlanMatrix");
  const ordered = [...(plans || [])].sort((a, b) => {
    const aLimit = a.vehicle_limit === null || a.vehicle_limit === undefined ? 999999 : Number(a.vehicle_limit || 0);
    const bLimit = b.vehicle_limit === null || b.vehicle_limit === undefined ? 999999 : Number(b.vehicle_limit || 0);
    return aLimit - bLimit || Number(a.monthly_price || 0) - Number(b.monthly_price || 0);
  });
  if (!ordered.length) {
    wrap.innerHTML = `<p class="empty-note">Nenhum plano configurado.</p>`;
    return;
  }
  wrap.innerHTML = ordered.map((plan) => {
    const features = Object.entries(plan.features || {})
      .filter(([, enabled]) => Boolean(enabled))
      .map(([feature]) => saasFeatureLabels[feature] || feature.replaceAll("_", " "));
    const vehicleLimit = plan.vehicle_limit === null || plan.vehicle_limit === undefined ? "Sem limite" : `${plan.vehicle_limit} veic.`;
    const userLimit = plan.user_limit === null || plan.user_limit === undefined ? "Sem limite" : `${plan.user_limit} usuarios`;
    const price = Number(plan.monthly_price || 0) > 0 ? formatCurrencyBR(plan.monthly_price) : "Sob contrato";
    const isCurrent = plan.code && plan.code === currentPlanCode;
    const isPrivate = Boolean((plan.features || {}).private_deployment);
    return `
      <article class="saas-plan-card ${isCurrent ? "current" : ""} ${isPrivate ? "private" : ""}">
        <div class="saas-plan-card-head">
          <span>${escapeHtml(plan.code || "-")}</span>
          ${isCurrent ? "<em>ATUAL</em>" : ""}
        </div>
        <h3>${escapeHtml(plan.name || plan.code || "-")}</h3>
        <strong class="saas-plan-price">${escapeHtml(price)}</strong>
        <p>${escapeHtml(plan.description || "Sem descricao.")}</p>
        <div class="saas-plan-limits">
          <span>${escapeHtml(vehicleLimit)}</span>
          <span>${escapeHtml(userLimit)}</span>
        </div>
        <div class="saas-plan-features">
          ${features.length ? features.map((feature) => `<span>${escapeHtml(feature)}</span>`).join("") : "<span>Basico</span>"}
        </div>
      </article>
    `;
  }).join("");
}

function renderSaasFeatures(featureData) {
  const wrap = el("saasAdminFeatures");
  const features = (featureData || {}).features || {};
  const entries = Object.entries(features).sort(([a], [b]) => a.localeCompare(b));
  if (!entries.length) {
    wrap.innerHTML = `<p class="empty-note">Nenhum recurso configurado para o plano atual.</p>`;
    return;
  }
  wrap.innerHTML = entries.map(([feature, enabled]) => `
    <div class="saas-feature ${enabled ? "enabled" : "disabled"}">
      <strong>${escapeHtml(saasFeatureLabels[feature] || feature.replaceAll("_", " "))}</strong>
      <span>${enabled ? "ATIVO" : "INATIVO"}</span>
    </div>
  `).join("");
}

function renderSaasPayments(payments) {
  const rows = el("saasAdminPaymentRows");
  rows.innerHTML = "";
  if (!payments.length) {
    rows.appendChild(emptyRow(8, "Sem pagamentos."));
    return;
  }
  payments.forEach((payment) => {
    const statusText = clean(payment.status).toLowerCase();
    const canRegister = statusText !== "paid";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(payment.id)}</td>
      <td>${escapeHtml(payment.status || "-")}</td>
      <td>${escapeHtml(formatCurrencyBR(payment.amount || 0))}</td>
      <td>${escapeHtml(payment.due_date || "-")}</td>
      <td>${escapeHtml(formatDate(payment.paid_at))}</td>
      <td>${escapeHtml(payment.reference || "-")}</td>
      <td>${escapeHtml(payment.notes || "-")}</td>
      <td>
        <button type="button" class="primary" data-saas-register-payment="${payment.id}" ${canRegister ? "" : "disabled"}>
          REGISTRAR
        </button>
      </td>
    `;
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-saas-register-payment]").forEach((button) => {
    button.addEventListener("click", () => registerSaasPayment(Number(button.dataset.saasRegisterPayment)));
  });
}

function renderSaasAudit(logs) {
  const rows = el("saasAdminAuditRows");
  rows.innerHTML = "";
  if (!logs.length) {
    rows.appendChild(emptyRow(6, "Sem auditoria SaaS."));
    return;
  }
  logs.forEach((log) => {
    const metadata = JSON.stringify(log.metadata || {}, null, 2);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(log.id)}</td>
      <td>${escapeHtml(log.action || "-")}</td>
      <td>${escapeHtml(log.severity || "-")}</td>
      <td>${escapeHtml(entityLabel(log))}</td>
      <td>${escapeHtml(formatDate(log.created_at))}</td>
      <td>
        <details class="metadata">
          <summary>Ver</summary>
          <pre>${escapeHtml(metadata)}</pre>
        </details>
      </td>
    `;
    rows.appendChild(tr);
  });
}

async function updateSaasCompanyStatus(status) {
  const companyId = currentSaasCompanyId();
  if (!companyId) {
    notify("Selecione uma empresa.", true);
    return;
  }
  if (!window.confirm(`${status === "active" ? "Ativar" : "Suspender"} a empresa selecionada?`)) return;
  try {
    await apiRequest(`/saas-admin/companies/${companyId}/status`, {
      method: "PUT",
      body: JSON.stringify({status}),
    });
    notify("Status da empresa atualizado.");
    await loadSaasAdmin();
  } catch (error) {
    notify(error.message, true);
  }
}

async function changeSaasCompanyPlan(event) {
  event.preventDefault();
  const companyId = currentSaasCompanyId();
  const planCode = clean(el("saasAdminPlanSelect").value);
  const reason = clean(el("saasAdminPlanReason").value);
  if (!companyId || !planCode) {
    notify("Selecione empresa e plano.", true);
    return;
  }
  if (!window.confirm(`Aplicar o plano ${planCode} para a empresa selecionada?`)) return;
  try {
    await apiRequest(`/saas-admin/companies/${companyId}/plan`, {
      method: "PUT",
      body: JSON.stringify({plan_code: planCode, reason}),
    });
    el("saasAdminPlanReason").value = "";
    notify("Plano alterado.");
    await loadSaasAdmin();
  } catch (error) {
    notify(error.message, true);
  }
}

async function createSaasPayment(event) {
  event.preventDefault();
  const companyId = currentSaasCompanyId();
  if (!companyId) {
    notify("Selecione uma empresa.", true);
    return;
  }
  const amount = parseNumberInput(el("saasAdminPaymentAmount").value);
  if (amount <= 0) {
    notify("Informe um valor maior que zero.", true);
    return;
  }
  try {
    await apiRequest("/saas-admin/payments", {
      method: "POST",
      body: JSON.stringify({
        company_id: companyId,
        amount,
        due_date: clean(el("saasAdminPaymentDue").value) || null,
        notes: clean(el("saasAdminPaymentNotes").value) || null,
      }),
    });
    event.currentTarget.reset();
    notify("Cobranca criada.");
    await loadSaasAdmin();
  } catch (error) {
    notify(error.message, true);
  }
}

async function registerSaasPayment(paymentId) {
  const reference = window.prompt("Referencia do pagamento", "");
  if (reference === null) return;
  try {
    await apiRequest(`/saas-admin/payments/${paymentId}/registrar-pagamento`, {
      method: "POST",
      body: JSON.stringify({method: "manual", reference: clean(reference) || null}),
    });
    notify("Pagamento registrado.");
    await loadSaasAdmin();
  } catch (error) {
    notify(error.message, true);
  }
}

async function runSaasOverdueCheck() {
  const graceDays = Math.max(parseInt(clean(el("saasAdminOverdueDays").value) || "0", 10) || 0, 0);
  if (!window.confirm(`Rodar suspensao por atraso com ${graceDays} dia(s) de tolerancia?`)) return;
  try {
    const result = await apiRequest("/saas-admin/billing/run-overdue-check", {
      method: "POST",
      body: JSON.stringify({grace_days: graceDays}),
    });
    notify(`Assinaturas suspensas: ${(result.summary || {}).suspended || 0}`);
    await loadSaasAdmin();
  } catch (error) {
    notify(error.message, true);
  }
}

async function onImportarVendasUpload(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const fileInput = el("importarVendasFile");
  if (!fileInput.files.length) {
    notify("Selecione um arquivo Excel.", true);
    return;
  }

  try {
    const body = new FormData(form);
    const result = await apiRequest("/importar-vendas/upload", {
      method: "POST",
      body,
    });
    form.reset();
    notify(importarVendasResultMessage(result));
    await Promise.all([refreshImportarVendasData({clearSearch: true}), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

function importarVendasResultMessage(result) {
  const parts = [`Pedidos importados: ${result.importadas || 0}`];
  if (result.ignoradas) parts.push(`Ignoradas: ${result.ignoradas}`);
  if (result.invalidas) parts.push(`Invalidas: ${result.invalidas}`);
  if (result.opcionais_ausentes && result.opcionais_ausentes.length) {
    parts.push(`Opcionais ausentes: ${result.opcionais_ausentes.join(", ")}`);
  }
  return parts.join(" | ");
}

async function loadVendasImportadas() {
  const params = new URLSearchParams();
  const busca = clean(el("importarVendasBusca").value);
  if (busca) params.set("busca", busca);
  try {
    const query = params.toString() ? `?${params.toString()}` : "";
    state.vendasImportadas = await apiRequest(`/importar-vendas/${query}`);
    renderVendasImportadas();
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadImportarVendasProgramacoes() {
  try {
    state.vendasProgramacoes = await apiRequest("/importar-vendas/programacoes-vinculo");
    renderImportarVendasProgramacoes();
  } catch (error) {
    notify(error.message, true);
  }
}

async function refreshImportarVendasData(options = {}) {
  if (options.clearSearch) {
    el("importarVendasBusca").value = "";
  }
  await Promise.all([loadVendasImportadas(), loadImportarVendasProgramacoes()]);
}

function renderImportarVendasProgramacoes() {
  const select = el("importarVendasProgramacao");
  const current = select.value;
  select.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = "Selecione";
  select.appendChild(empty);
  state.vendasProgramacoes.forEach((codigo) => {
    const option = document.createElement("option");
    option.value = codigo;
    option.textContent = codigo;
    select.appendChild(option);
  });
  if (current && state.vendasProgramacoes.includes(current)) {
    select.value = current;
  } else if (state.vendasProgramacoes.length) {
    select.value = state.vendasProgramacoes[0];
  }
}

function renderVendasImportadas() {
  const rows = el("importarVendasRows");
  rows.innerHTML = "";
  const selectedCount = state.vendasImportadas.filter((venda) => Number(venda.selecionada) === 1).length;
  el("importarVendasInfo").textContent = `${state.vendasImportadas.length} registros carregados | Selecionadas: ${selectedCount}`;
  el("importarVendasTotalPedidos").textContent = formatPlainNumber(state.vendasImportadas.length);
  el("importarVendasSelectedPedidos").textContent = formatPlainNumber(selectedCount);
  if (!state.vendasImportadas.length) {
    rows.appendChild(emptyRow(14, "Sem pedidos livres."));
    updateImportarVendasTotals();
    return;
  }

  state.vendasImportadas.forEach((venda) => {
    const tr = document.createElement("tr");
    const isSelected = Number(venda.selecionada) === 1;
    tr.dataset.vendaId = String(venda.id);
    tr.className = isSelected ? "importar-vendas-selected-row" : "";
    tr.innerHTML = `
      <td><input type="checkbox" data-venda-row-select aria-label="Selecionar pedido" ${isSelected ? "checked" : ""}></td>
      <td class="num-cell">${venda.id}</td>
      <td class="center-cell">${isSelected ? "SIM" : "-"}</td>
      <td class="num-cell">${escapeHtml(venda.pedido || "")}</td>
      <td class="date-cell">${escapeHtml(venda.data_venda || "")}</td>
      <td class="num-cell">${escapeHtml(venda.cliente || "")}</td>
      <td>${escapeHtml(venda.nome_cliente || "")}</td>
      <td>${escapeHtml(venda.produto || "")}</td>
      <td class="num-cell">${escapeHtml(formatNumber(venda.vr_total || 0))}</td>
      <td class="num-cell">${escapeHtml(formatNumber(venda.qnt || 0))}</td>
      <td><input class="venda-caixas-input" type="number" min="0" step="1" value="${escapeHtml(venda.qnt_caixas || 0)}" data-venda-caixas="${venda.id}" aria-label="Caixas do pedido"></td>
      <td>${escapeHtml(venda.cidade || "")}</td>
      <td>${escapeHtml(venda.vendedor || "")}</td>
      <td>
        <div class="actions">
          <button type="button" class="secondary compact-action" data-venda-action="toggle" data-venda-id="${venda.id}">${Number(venda.selecionada) === 1 ? "Desmarcar" : "Marcar"}</button>
        </div>
      </td>
    `;
    tr.addEventListener("dblclick", () => toggleVendaImportada(venda.id));
    rows.appendChild(tr);
  });

  rows.querySelectorAll("[data-venda-action='toggle']").forEach((button) => {
    button.addEventListener("click", () => toggleVendaImportada(Number(button.dataset.vendaId)));
  });
  rows.querySelectorAll("[data-venda-caixas]").forEach((input) => {
    input.addEventListener("change", () => updateVendaImportadaCaixas(Number(input.dataset.vendaCaixas), input.value));
  });
  rows.querySelectorAll("[data-venda-row-select]").forEach((input) => {
    input.addEventListener("change", () => {
      input.closest("tr")?.classList.toggle("importar-vendas-selected-row", input.checked);
      updateImportarVendasTotals();
    });
  });
  updateImportarVendasTotals();
}

function importarVendasCurrentRows() {
  return [...el("importarVendasRows").querySelectorAll("tr[data-venda-id]")];
}

function importarVendasCaixasFromRow(row) {
  const input = row.querySelector("[data-venda-caixas]");
  return Math.max(Math.trunc(Number(input?.value || 0)), 0);
}

function updateImportarVendasTotals() {
  const rows = importarVendasCurrentRows();
  const checkedIds = checkedVendaIds();
  const checkedSet = new Set(checkedIds.map((id) => Number(id)));
  const markedSet = new Set(
    state.vendasImportadas
      .filter((venda) => Number(venda.selecionada) === 1)
      .map((venda) => Number(venda.id)),
  );
  const selectedSet = checkedSet.size ? checkedSet : markedSet;
  const totalCaixas = rows.reduce((sum, row) => sum + importarVendasCaixasFromRow(row), 0);
  const selectedCaixas = rows
    .filter((row) => selectedSet.has(Number(row.dataset.vendaId)))
    .reduce((sum, row) => sum + importarVendasCaixasFromRow(row), 0);
  el("importarVendasTotalCaixas").textContent = formatPlainNumber(totalCaixas);
  el("importarVendasSelectedCaixas").textContent = formatPlainNumber(selectedCaixas);
}

function checkedVendaIds() {
  return [...el("importarVendasRows").querySelectorAll("tr[data-venda-id]")]
    .filter((row) => row.querySelector("[data-venda-row-select]")?.checked)
    .map((row) => Number(row.dataset.vendaId))
    .filter(Boolean);
}

function selectedOrMarkedVendaIds() {
  const checked = checkedVendaIds();
  if (checked.length) return checked;
  return state.vendasImportadas
    .filter((venda) => Number(venda.selecionada) === 1)
    .map((venda) => Number(venda.id))
    .filter(Boolean);
}

async function toggleVendaImportada(id) {
  try {
    await apiRequest(`/importar-vendas/${id}/toggle-selecao`, {method: "POST"});
    await loadVendasImportadas();
  } catch (error) {
    notify(error.message, true);
  }
}

async function updateVendaImportadaCaixas(id, value) {
  const caixas = Math.max(Math.trunc(Number(value || 0)), 0);
  try {
    const updated = await apiRequest(`/importar-vendas/${id}/caixas`, {
      method: "PUT",
      body: JSON.stringify({qnt_caixas: caixas}),
    });
    state.vendasImportadas = state.vendasImportadas.map((item) => Number(item.id) === Number(id) ? updated : item);
    renderVendasImportadas();
  } catch (error) {
    notify(error.message, true);
    await loadVendasImportadas();
  }
}

async function setAllVendasImportadas(selected) {
  try {
    const result = await apiRequest(`/importar-vendas/marcar-todas?selected=${selected ? 1 : 0}`, {method: "POST"});
    notify(`Pedidos atualizados: ${result.updated || 0}`);
    await loadVendasImportadas();
  } catch (error) {
    notify(error.message, true);
  }
}

async function markSelectedVendasImportadas() {
  const ids = checkedVendaIds();
  if (!ids.length) {
    notify("Selecione um ou mais pedidos na tabela.", true);
    return;
  }
  try {
    const result = await apiRequest("/importar-vendas/marcar-ids", {
      method: "POST",
      body: JSON.stringify({ids}),
    });
    notify(`Pedidos marcados: ${result.updated || 0}`);
    await loadVendasImportadas();
  } catch (error) {
    notify(error.message, true);
  }
}

async function deleteSelectedVendasImportadas() {
  const ids = checkedVendaIds();
  if (!ids.length) {
    notify("Selecione um ou mais pedidos para excluir.", true);
    return;
  }
  if (!confirm(`Excluir ${ids.length} pedido(s) importado(s)?`)) return;
  try {
    const result = await apiRequest("/importar-vendas/ids", {
      method: "DELETE",
      body: JSON.stringify({ids}),
    });
    notify(`Pedidos excluidos: ${result.deleted || 0}`);
    await loadVendasImportadas();
  } catch (error) {
    notify(error.message, true);
  }
}

async function clearAllVendasImportadas() {
  if (!confirm("Apagar TODOS os pedidos importados livres?")) return;
  try {
    const result = await apiRequest("/importar-vendas/", {method: "DELETE"});
    notify(`Pedidos excluidos: ${result.deleted || 0}`);
    await loadVendasImportadas();
  } catch (error) {
    notify(error.message, true);
  }
}

async function linkSelectedVendasToProgramacao() {
  const codigo = clean(el("importarVendasProgramacao").value);
  if (!codigo) {
    notify("Selecione o planejamento ativo para vincular.", true);
    return;
  }
  const ids = selectedOrMarkedVendaIds();
  if (!ids.length) {
    notify("Selecione ou marque um ou mais pedidos para vincular.", true);
    return;
  }

  const caixasPorVenda = {};
  for (const id of ids) {
    const venda = state.vendasImportadas.find((item) => Number(item.id) === Number(id));
    const caixas = Math.trunc(Number(venda?.qnt_caixas || 0));
    if (caixas <= 0) {
      notify(`Informe as caixas na grade para o pedido ${venda?.pedido || id}.`, true);
      return;
    }
    caixasPorVenda[String(id)] = caixas;
  }

  try {
    const result = await apiRequest("/importar-vendas/vincular", {
      method: "POST",
      body: JSON.stringify({codigo_programacao: codigo, ids, caixas_por_venda: caixasPorVenda}),
    });
    notify(`${result.vendas_vinculadas} pedido(s) vinculado(s) a ${result.codigo_programacao}.`);
    await Promise.all([refreshImportarVendasData({clearSearch: true}), loadProgramacoes(), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadProgramacaoOptions() {
  try {
    state.programacaoOptions = await apiRequest("/programacao/options");
    renderProgramacaoOptions();
    if (!el("progCodigo").value) {
      el("progCodigo").value = state.programacaoOptions.proximo_codigo || "";
    }
  } catch (error) {
    notify(error.message, true);
  }
}

function renderProgramacaoOptions() {
  const options = state.programacaoOptions || {};
  fillSelect(
    el("progMotorista"),
    sortedProgramacaoMotoristas(options.motoristas || []),
    (item) => item.display || item.nome || "",
    (item) => item.display || item.nome || "",
    (option, item) => {
      option.dataset.nome = item.nome || "";
      option.dataset.codigo = item.codigo || "";
    },
  );
  fillSelect(el("progVeiculo"), options.veiculos || [], (item) => item.placa || "", (item) => item.placa || "");

  const ajudantes = el("progAjudantes");
  const selected = new Set([...ajudantes.selectedOptions].map((option) => option.value));
  ajudantes.innerHTML = "";
  sortedProgramacaoAjudantes(options.ajudantes || []).forEach((item) => {
    const option = document.createElement("option");
    option.value = item.id || "";
    option.textContent = item.display || [item.nome, item.sobrenome].filter(Boolean).join(" ");
    option.selected = selected.has(option.value);
    ajudantes.appendChild(option);
  });
  updateProgramacaoAjudantesResumo();
}

function rankingIndex(collection, matcher) {
  const ranking = collection || [];
  const found = ranking.find(matcher);
  return found ? Number(found.posicao_ranking || 9999) : 9999;
}

function sortedProgramacaoMotoristas(items) {
  const ranking = state.programacaoRankings?.motoristas || [];
  return [...items].sort((a, b) => {
    const aCodigo = clean(a.codigo).toUpperCase();
    const bCodigo = clean(b.codigo).toUpperCase();
    const aNome = clean(a.nome).toUpperCase();
    const bNome = clean(b.nome).toUpperCase();
    const aPos = rankingIndex(ranking, (item) => (aCodigo && clean(item.codigo).toUpperCase() === aCodigo) || clean(item.nome).toUpperCase() === aNome);
    const bPos = rankingIndex(ranking, (item) => (bCodigo && clean(item.codigo).toUpperCase() === bCodigo) || clean(item.nome).toUpperCase() === bNome);
    return aPos - bPos || aNome.localeCompare(bNome);
  });
}

function sortedProgramacaoAjudantes(items) {
  const ranking = state.programacaoRankings?.ajudantes || [];
  return [...items].sort((a, b) => {
    const aId = clean(a.id);
    const bId = clean(b.id);
    const aNome = clean(a.display || `${a.nome || ""} ${a.sobrenome || ""}`).toUpperCase();
    const bNome = clean(b.display || `${b.nome || ""} ${b.sobrenome || ""}`).toUpperCase();
    const aPos = rankingIndex(ranking, (item) => clean(item.id) === aId || clean(item.nome).toUpperCase() === aNome);
    const bPos = rankingIndex(ranking, (item) => clean(item.id) === bId || clean(item.nome).toUpperCase() === bNome);
    return aPos - bPos || aNome.localeCompare(bNome);
  });
}

async function loadProgramacaoRankings() {
  const periodo = clean(el("progRankingPeriodo").value) || "30";
  try {
    state.programacaoRankings = await apiRequest(`/programacao/rankings?periodo=${encodeURIComponent(periodo)}`);
    renderProgramacaoRankings();
    renderProgramacaoOptions();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderProgramacaoRankings() {
  const data = state.programacaoRankings || {motoristas: [], ajudantes: []};
  el("progRankingInfo").textContent = `Periodo analisado: ${data.periodo_dias || clean(el("progRankingPeriodo").value) || 30} dias.`;
  el("progMotoristaRankingSummary").textContent = data.resumo_motoristas || "Top motoristas: sem candidatos elegiveis.";
  el("progAjudanteRankingSummary").textContent = data.resumo_ajudantes || "Top ajudantes: sem candidatos elegiveis.";
  renderProgramacaoRankingRows("progMotoristaRankingRows", data.motoristas || []);
  renderProgramacaoRankingRows("progAjudanteRankingRows", data.ajudantes || []);
}

function renderProgramacaoRankingRows(targetId, rowsData) {
  const rows = el(targetId);
  rows.innerHTML = "";
  if (!rowsData.length) {
    rows.appendChild(emptyRow(7, "Sem candidatos elegiveis."));
    return;
  }
  rowsData.slice(0, 8).forEach((item) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.posicao_ranking || "-")}</td>
      <td>${escapeHtml(item.display || item.nome || "-")}</td>
      <td>${escapeHtml(formatNumber(item.score_exibicao || 0))}</td>
      <td>${escapeHtml(item.total_viagens || 0)}</td>
      <td>${escapeHtml(formatNumber(item.total_horas_trabalhadas || 0))}</td>
      <td>${escapeHtml(item.dias_desde_ultima_programacao ?? "-")}</td>
      <td>${escapeHtml(item.motivo_resumido || "-")}</td>
    `;
    rows.appendChild(tr);
  });
}

function fillSelect(select, items, valueFn, labelFn, decorate) {
  const current = select.value;
  select.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = "Selecione";
  select.appendChild(empty);
  items.forEach((item) => {
    const option = document.createElement("option");
    option.value = valueFn(item);
    option.textContent = labelFn(item);
    if (decorate) decorate(option, item);
    select.appendChild(option);
  });
  if (current) select.value = current;
}

async function loadProgramacoes() {
  try {
    state.programacoes = await apiRequest("/programacao/");
    renderProgramacoesTable();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderProgramacoesTable() {
  const rows = el("progProgramacoesRows");
  rows.innerHTML = "";
  if (!state.programacoes.length) {
    rows.appendChild(emptyRow(8, "Sem programacoes."));
    return;
  }
  state.programacoes.forEach((programacao) => {
    const tr = document.createElement("tr");
    tr.className = state.programacaoSelectedCodigo === programacao.codigo_programacao ? "selected-row" : "";
    tr.innerHTML = `
      <td>${escapeHtml(programacao.codigo_programacao)}</td>
      <td>${escapeHtml(programacao.motorista)}</td>
      <td>${escapeHtml(programacao.veiculo)}</td>
      <td>${escapeHtml(programacao.local_rota || "-")}</td>
      <td>${escapeHtml(programacao.status_operacional || programacao.status || "-")}</td>
      <td>${escapeHtml(programacao.total_caixas ?? 0)}</td>
      <td>${escapeHtml(formatNumber(programacao.quilos || programacao.kg_estimado || 0))}</td>
      <td><button type="button" class="secondary" data-prog-load="${escapeHtml(programacao.codigo_programacao)}">Carregar</button></td>
    `;
    tr.addEventListener("click", (event) => {
      if (event.target.closest("button")) return;
      loadProgramacao(programacao.codigo_programacao);
    });
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-prog-load]").forEach((button) => {
    button.addEventListener("click", () => loadProgramacao(button.dataset.progLoad));
  });
}

async function loadProgramacao(codigo) {
  if (!codigo) return;
  try {
    const programacao = await apiRequest(`/programacao/${encodeURIComponent(codigo)}`);
    fillProgramacaoForm(programacao);
    notify(`Planejamento ${programacao.codigo_programacao} carregado.`);
  } catch (error) {
    notify(error.message, true);
  }
}

function updateProgramacaoEmissionButtons() {
  const hasSavedProgramacao = Boolean(state.programacaoSelectedCodigo);
  const pdfButton = document.getElementById("progPdfButton");
  const reciboButton = document.getElementById("progAdiantamentoPdfButton");
  if (pdfButton) pdfButton.disabled = !hasSavedProgramacao;
  if (reciboButton) reciboButton.disabled = !hasSavedProgramacao;
  el("progRomaneiosButton").disabled = !hasSavedProgramacao;
}

function selectedProgramacaoCodigoForEmission() {
  const codigo = clean(state.programacaoSelectedCodigo);
  if (!codigo) {
    throw new Error("Carregue ou salve um planejamento antes de emitir.");
  }
  return codigo;
}

async function downloadProgramacaoPdf() {
  try {
    const codigo = selectedProgramacaoCodigoForEmission();
    const result = await apiBlobRequest(`/programacao/${encodeURIComponent(codigo)}/pdf`);
    downloadBlob(result.blob, result.filename || `PROGRAMACAO_${codigo}.pdf`);
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadProgramacaoAdiantamentoRecibo() {
  try {
    const codigo = selectedProgramacaoCodigoForEmission();
    const result = await apiBlobRequest(`/programacao/${encodeURIComponent(codigo)}/recibo-adiantamento-pdf`);
    downloadBlob(result.blob, result.filename || `RECIBO_ADIANTAMENTO_${codigo}.pdf`);
  } catch (error) {
    notify(error.message, true);
  }
}

async function downloadProgramacaoRomaneios() {
  try {
    const codigo = selectedProgramacaoCodigoForEmission();
    const result = await apiBlobRequest(`/programacao/${encodeURIComponent(codigo)}/romaneios-pdf`);
    downloadBlob(result.blob, result.filename || `ROMANEIOS_${codigo}.pdf`);
  } catch (error) {
    notify(error.message, true);
  }
}

async function loadSelectedVendasIntoProgramacao() {
  try {
    const result = await apiRequest("/programacao/vendas-selecionadas");
    const itens = result.itens || [];
    if (!itens.length) {
      notify("Nenhum pedido marcado para carregar.", true);
      return;
    }
    state.programacaoItems = itens.map((item) => ({
      venda_id: Number(item.venda_id || 0),
      cod_cliente: item.cod_cliente || "",
      nome_cliente: item.nome_cliente || "",
      produto: item.produto || "",
      endereco: item.endereco || "",
      qnt_caixas: item.qnt_caixas || 0,
      kg: item.kg || 0,
      preco: item.preco || 0,
      vendedor: item.vendedor || "",
      pedido: item.pedido || "",
      obs: item.obs || "",
      recomendacao: item.recomendacao || "",
      ordem_sugerida: Number(item.ordem_sugerida || 0) || null,
      distancia: Number(item.distancia || item.distancia_anterior_km || 0) || 0,
      confianca_localizacao: Number(item.confianca_localizacao || 0) || 0,
    }));
    state.programacaoLoadedVendaIds = (result.ids || []).map((id) => Number(id || 0)).filter(Boolean);
    renderProgramacaoItems();
    const ignored = result.invalidas ? ` | Invalidas ignoradas: ${result.invalidas}` : "";
    notify(`${state.programacaoItems.length} pedido(s) carregado(s) para o planejamento.${ignored}`);
  } catch (error) {
    notify(error.message, true);
  }
}

async function suggestProgramacaoRoute() {
  try {
    const veiculo = clean(el("progVeiculo").value);
    if (!veiculo) {
      throw new Error("Selecione um veiculo antes de sugerir rota.");
    }
    const currentItems = collectProgramacaoItems({allowBlank: false});
    const payload = {
      veiculo,
      local_rota: clean(el("progLocalRota").value),
      itens: currentItems,
    };
    const result = await apiRequest("/programacao/sugestao", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    const vendaIdsByKey = new Map();
    for (const row of el("progItemsRows").querySelectorAll("tr[data-item-index]")) {
      const item = programacaoItemFromRow(row);
      const key = `${clean(item.cod_cliente).toUpperCase()}|${clean(item.pedido).toUpperCase()}|${clean(item.produto).toUpperCase()}`;
      vendaIdsByKey.set(key, item.venda_id || 0);
    }
    state.programacaoItems = (result.itens || []).map((item) => {
      const key = `${clean(item.cod_cliente).toUpperCase()}|${clean(item.pedido).toUpperCase()}|${clean(item.produto).toUpperCase()}`;
      return {
        venda_id: vendaIdsByKey.get(key) || Number(item.venda_id || 0) || 0,
        cod_cliente: item.cod_cliente || "",
        nome_cliente: item.nome_cliente || "",
        produto: item.produto || "",
        endereco: item.endereco || "",
        qnt_caixas: item.qnt_caixas || 0,
        kg: item.kg || 0,
        preco: item.preco || 0,
        vendedor: item.vendedor || "",
        pedido: item.pedido || "",
        obs: item.obs || "",
        recomendacao: item.recomendacao || "",
        ordem_sugerida: Number(item.ordem_sugerida || 0) || null,
        distancia: Number(item.distancia || item.distancia_anterior_km || 0) || 0,
        confianca_localizacao: Number(item.confianca_localizacao || 0) || 0,
      };
    });
    renderProgramacaoSuggestion(result);
    renderProgramacaoItems();
    notify("Sugestao aplicada aos itens do planejamento.");
  } catch (error) {
    notify(error.message, true);
  }
}

function renderProgramacaoSuggestion(result) {
  const box = el("progSugestaoResumo");
  const alertas = result.alertas || [];
  const revisar = (result.itens || [])
    .filter((item) => clean(item.recomendacao) && clean(item.recomendacao) !== "OK")
    .slice(0, 8);
  box.classList.remove("hidden");
  box.innerHTML = `
    <div>
      <strong>Sugestao inteligente</strong>
      <p>${escapeHtml(result.resumo || "")}</p>
    </div>
    <div class="programacao-suggestion-grid">
      <span>Veiculo: <strong>${escapeHtml(result.veiculo || "-")}</strong></span>
      <span>Capacidade: <strong>${escapeHtml(result.capacidade_cx || 0)} cx</strong></span>
      <span>Total: <strong>${escapeHtml(result.total_caixas || 0)} cx</strong></span>
      <span>Excedente: <strong>${escapeHtml(result.caixas_excedentes || 0)} cx</strong></span>
      <span>Com GPS: <strong>${escapeHtml(result.clientes_com_localizacao || 0)}</strong></span>
      <span>Sem GPS: <strong>${escapeHtml(result.clientes_sem_localizacao || 0)}</strong></span>
      <span>Distancia estimada: <strong>${escapeHtml(formatPlainNumber(result.distancia_estimativa_km || 0))} km</strong></span>
    </div>
    ${alertas.length ? `<ul>${alertas.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>` : ""}
    ${revisar.length ? `
      <div class="programacao-suggestion-review">
        <strong>Clientes para revisar</strong>
        ${revisar.map((item) => `
          <span>${escapeHtml(item.ordem_sugerida || "-")} - ${escapeHtml(item.cod_cliente || "")} ${escapeHtml(item.nome_cliente || "")}: ${escapeHtml(item.recomendacao || "")}</span>
        `).join("")}
      </div>
    ` : ""}
  `;
}

function fillProgramacaoForm(programacao) {
  state.programacaoSelectedCodigo = programacao.codigo_programacao || "";
  state.programacaoLoadedVendaIds = [];
  el("progCodigo").value = programacao.codigo_programacao || "";
  setProgramacaoMotorista(programacao.motorista, programacao.motorista_codigo);
  ensureSelectValue(el("progVeiculo"), programacao.veiculo || "", programacao.veiculo || "");
  el("progLocalRota").value = normalizeLocalRotaJs(programacao.local_rota || "");
  el("progTipoEstimativa").value = programacao.tipo_estimativa || "KG";
  el("progEstimativa").value = programacao.tipo_estimativa === "CX"
    ? String(programacao.caixas_estimado || "")
    : formatNumber(programacao.kg_estimado || 0);
  el("progCarregamento").value = programacao.local_carregamento || "";
  el("progTransbordoModalidade").value = programacao.tipo_estimativa === "CX" ? "EMPRESA_BUSCA" : "CIF";
  el("progTransbordoObs").value = programacao.transbordo_observacao || "";
  el("progAdiantamento").value = formatMoneyInput(programacao.adiantamento || 0);
  el("progAdiantamentoOrigem").value = programacao.adiantamento_origem || "";
  setProgramacaoAjudantes(programacao.ajudantes || []);
  el("progStatusBadge").textContent = programacao.status_operacional || programacao.status || "-";
  state.programacaoItems = (programacao.itens || []).map((item) => ({
    venda_id: 0,
    cod_cliente: item.cod_cliente || "",
    nome_cliente: item.nome_cliente || "",
    produto: item.produto || "",
    endereco: item.endereco || "",
    qnt_caixas: item.qnt_caixas || 0,
    kg: item.kg || 0,
    preco: item.preco || 0,
    vendedor: item.vendedor || "",
    pedido: item.pedido || "",
    obs: item.obs || "",
    recomendacao: item.recomendacao || "",
    ordem_sugerida: Number(item.ordem_sugerida || 0) || null,
    distancia: Number(item.distancia || 0) || 0,
    confianca_localizacao: Number(item.confianca_localizacao || 0) || 0,
    carga_raiz_programacao: item.carga_raiz_programacao || "",
    carga_origem_imediata: item.carga_origem_imediata || "",
    transferencia_origem_id: item.transferencia_origem_id || "",
  }));
  renderProgramacaoItems();
  renderProgramacoesTable();
  updateProgramacaoEmissionButtons();
  updateProgramacaoEstimateLabel();
  setProgramacaoEditingEnabled(canEditProgramacaoStatus(programacao));
}

function normalizeProgramacaoStatus(value) {
  return clean(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replaceAll("_", " ");
}

function canEditProgramacaoStatus(programacao) {
  const status = normalizeProgramacaoStatus(programacao?.status_operacional || programacao?.status || "");
  const prestacao = normalizeProgramacaoStatus(programacao?.prestacao_status || "");
  if (prestacao === "FECHADA") return false;
  return !status || status === "ATIVA" || status === "EM ROTA";
}

function currentProgramacaoEditable() {
  const status = normalizeProgramacaoStatus(el("progStatusBadge").textContent || "");
  if (!status || status === "-") return true;
  return status === "ATIVA" || status === "EM ROTA";
}

function setProgramacaoEditingEnabled(enabled) {
  const form = el("programacaoForm");
  form.querySelectorAll("input, select, textarea").forEach((field) => {
    if (field.id === "progCodigo") {
      field.readOnly = true;
      return;
    }
    field.disabled = !enabled;
  });
  el("progItemsRows").querySelectorAll("input, select, textarea").forEach((field) => {
    field.disabled = !enabled;
  });
  el("progAjudantesButton").disabled = !enabled;
  ["progLoadVendasButton", "progSuggestButton", "progAddItemButton", "progRemoveItemButton", "progClearItemsButton", "progSaveButton"].forEach((id) => {
    el(id).disabled = !enabled;
  });
  el("progRefreshOptions").disabled = false;
}

async function editarPlanejamentoAtual() {
  const codigo = clean(el("progCodigo").value || state.programacaoSelectedCodigo);
  if (!codigo) {
    notify("Selecione ou carregue um planejamento para editar.", true);
    return;
  }
  try {
    const programacao = await apiRequest(`/programacao/${encodeURIComponent(codigo)}`);
    fillProgramacaoForm(programacao);
    if (!canEditProgramacaoStatus(programacao)) {
      notify("Este planejamento nao pode ser editado. Edicao permitida somente nos status ATIVA e EM ROTA.", true);
      return;
    }
    setProgramacaoEditingEnabled(true);
    notify(`Planejamento ${codigo} liberado para edicao.`);
  } catch (error) {
    notify(error.message, true);
  }
}

function setProgramacaoMotorista(nome, codigo) {
  const select = el("progMotorista");
  const codigoNorm = clean(codigo).toUpperCase();
  const nomeNorm = clean(nome).toUpperCase();
  for (const option of select.options) {
    if ((codigoNorm && clean(option.dataset.codigo).toUpperCase() === codigoNorm) || clean(option.dataset.nome).toUpperCase() === nomeNorm) {
      select.value = option.value;
      return;
    }
  }
  const label = codigoNorm ? `${nomeNorm} (${codigoNorm})` : nomeNorm;
  ensureSelectValue(select, label, label, (option) => {
    option.dataset.nome = nomeNorm;
    option.dataset.codigo = codigoNorm;
  });
}

function ensureSelectValue(select, value, label, decorate) {
  if (!value) {
    select.value = "";
    return;
  }
  const existing = [...select.options].find((option) => option.value === value);
  if (!existing) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label || value;
    if (decorate) decorate(option);
    select.appendChild(option);
  }
  select.value = value;
}

function setProgramacaoAjudantes(values) {
  const select = el("progAjudantes");
  const selected = new Set((values || []).map((value) => clean(value).toUpperCase()));
  for (const option of select.options) {
    const value = clean(option.value).toUpperCase();
    const label = clean(option.textContent).toUpperCase();
    option.selected = selected.has(value) || selected.has(label);
  }
  updateProgramacaoAjudantesResumo();
}

function selectedProgramacaoAjudantes() {
  return [...el("progAjudantes").selectedOptions].map((option) => ({
    value: option.value,
    label: clean(option.textContent),
  }));
}

function updateProgramacaoAjudantesResumo() {
  const selected = selectedProgramacaoAjudantes();
  const resumo = el("progAjudantesResumo");
  if (!selected.length) {
    resumo.textContent = "Nenhum ajudante selecionado";
  } else {
    resumo.textContent = selected.map((item) => item.label).join(" | ");
  }
  el("progAjudantesButton").textContent = selected.length ? "Alterar ajudantes" : "Escolher ajudantes";
}

function openProgramacaoAjudantesDialog() {
  renderProgramacaoAjudantesDialog();
  el("progAjudantesDialog").showModal();
}

function closeProgramacaoAjudantesDialog() {
  el("progAjudantesDialog").close();
}

function renderProgramacaoAjudantesDialog() {
  const target = el("progAjudantesDialogList");
  const selectedValues = new Set(selectedProgramacaoAjudantes().map((item) => item.value));
  const options = [...el("progAjudantes").options];
  target.innerHTML = "";
  if (!options.length) {
    target.appendChild(emptyDiv("Sem ajudantes cadastrados. Atualize os cadastros ou cadastre ajudantes."));
    return;
  }
  options.forEach((option) => {
    const label = document.createElement("label");
    label.className = "check-row ajudante-picker-item";
    label.innerHTML = `
      <input type="checkbox" value="${escapeHtml(option.value)}"${selectedValues.has(option.value) ? " checked" : ""}>
      <span>${escapeHtml(option.textContent || option.value)}</span>
    `;
    target.appendChild(label);
  });
}

function applyProgramacaoAjudantesDialog() {
  const checked = [...el("progAjudantesDialogList").querySelectorAll("input[type='checkbox']:checked")];
  if (checked.length !== 2) {
    notify("Selecione exatamente 2 ajudantes.", true);
    return;
  }
  const selected = new Set(checked.map((input) => input.value));
  for (const option of el("progAjudantes").options) {
    option.selected = selected.has(option.value);
  }
  updateProgramacaoAjudantesResumo();
  closeProgramacaoAjudantesDialog();
}

function clearProgramacaoForm() {
  state.programacaoSelectedCodigo = "";
  state.programacaoLoadedVendaIds = [];
  state.programacaoItems = [];
  el("programacaoForm").reset();
  el("progTipoEstimativa").value = "KG";
  el("progTransbordoModalidade").value = "CIF";
  el("progTransbordoObs").value = "";
  el("progAdiantamento").value = "0,00";
  el("progStatusBadge").textContent = "-";
  el("progCodigo").value = state.programacaoOptions.proximo_codigo || "";
  el("progSugestaoResumo").classList.add("hidden");
  el("progSugestaoResumo").innerHTML = "";
  renderProgramacaoItems();
  renderProgramacoesTable();
  updateProgramacaoEmissionButtons();
  updateProgramacaoEstimateLabel();
  setProgramacaoEditingEnabled(true);
}

function updateProgramacaoEstimateLabel(options = {}) {
  const { syncFrete = false, syncEstimativa = false } = options;
  const freteSelect = el("progTransbordoModalidade");
  const estimativaSelect = el("progTipoEstimativa");
  if (syncEstimativa) {
    estimativaSelect.value = freteSelect.value === "EMPRESA_BUSCA" ? "CX" : "KG";
  } else if (syncFrete) {
    freteSelect.value = estimativaSelect.value === "CX" ? "EMPRESA_BUSCA" : "CIF";
  }
  const input = el("progEstimativa");
  const config = currentLogisticaConfig();
  const isTransbordo = estimativaSelect.value === "CX";
  input.placeholder = isTransbordo
    ? `${config.embalagem_label || "Caixas"} estimadas`
    : `${config.unidade_padrao || "KG"} estimado`;
  el("progTransbordoObsField").classList.toggle("hidden", !isTransbordo);
  el("progTransbordoObs").disabled = !isTransbordo;
  el("progCarregamento").placeholder = isTransbordo ? "Fornecedor ou ponto de transbordo" : "";
}

function addProgramacaoItemRow(item = {}) {
  state.programacaoItems.push({
    venda_id: item.venda_id || 0,
    cod_cliente: item.cod_cliente || "",
    nome_cliente: item.nome_cliente || "",
    produto: item.produto || "",
    endereco: item.endereco || "",
    qnt_caixas: item.qnt_caixas ?? 1,
    kg: item.kg ?? 0,
    preco: item.preco ?? 0,
    vendedor: item.vendedor || "",
    pedido: item.pedido || "",
    obs: item.obs || "",
    recomendacao: item.recomendacao || "",
  });
  renderProgramacaoItems();
}

function renderProgramacaoItems() {
  const rows = el("progItemsRows");
  rows.innerHTML = "";
  if (!state.programacaoItems.length) {
    rows.appendChild(emptyRow(11, "Sem itens. Use INSERIR LINHA para adicionar clientes manualmente."));
    updateProgramacaoItemTotals();
    return;
  }
  state.programacaoItems.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.dataset.itemIndex = String(index);
    tr.dataset.vendaId = String(item.venda_id || "");
    tr.dataset.ordemSugerida = String(item.ordem_sugerida || index + 1);
    tr.dataset.distancia = String(item.distancia || 0);
    tr.dataset.confiancaLocalizacao = String(item.confianca_localizacao || 0);
    tr.dataset.recomendacao = String(item.recomendacao || "");
    tr.dataset.cargaRaizProgramacao = String(item.carga_raiz_programacao || "");
    tr.dataset.cargaOrigemImediata = String(item.carga_origem_imediata || "");
    tr.dataset.transferenciaOrigemId = String(item.transferencia_origem_id || "");
    const recomendacao = clean(item.recomendacao);
    if (recomendacao && recomendacao !== "OK") {
      tr.classList.add(recomendacao.includes("EXCEDE") ? "programacao-row-danger" : "programacao-row-warning");
      tr.title = recomendacao;
    }
    tr.innerHTML = `
      <td><input type="checkbox" data-prog-item-select aria-label="Selecionar item"></td>
      ${programacaoInputCell("cod_cliente", item.cod_cliente, "text", true)}
      ${programacaoInputCell("nome_cliente", item.nome_cliente, "text", true)}
      ${programacaoInputCell("produto", item.produto)}
      ${programacaoInputCell("endereco", item.endereco)}
      ${programacaoInputCell("qnt_caixas", item.qnt_caixas, "number")}
      ${programacaoInputCell("kg", formatNumber(item.kg), "text")}
      ${programacaoInputCell("preco", formatNumber(item.preco), "text")}
      ${programacaoInputCell("vendedor", item.vendedor)}
      ${programacaoInputCell("pedido", item.pedido)}
      ${programacaoInputCell("obs", item.obs)}
    `;
    rows.appendChild(tr);
  });
  updateProgramacaoItemTotals();
}

function currentProgramacaoRows() {
  return [...el("progItemsRows").querySelectorAll("tr[data-item-index]")];
}

function updateProgramacaoItemTotals() {
  const rows = currentProgramacaoRows();
  let totalCaixas = 0;
  let totalClientes = 0;
  for (const row of rows) {
    const item = programacaoItemFromRow(row);
    if (clean(item.cod_cliente) || clean(item.nome_cliente)) totalClientes += 1;
    totalCaixas += Math.max(Math.trunc(Number(item.qnt_caixas || 0)), 0);
  }
  el("progTotalCaixas").textContent = formatPlainNumber(totalCaixas);
  el("progTotalClientes").textContent = formatPlainNumber(totalClientes);
}

function programacaoInputCell(name, value = "", type = "text", required = false) {
  const requiredAttr = required ? " required" : "";
  const minAttr = type === "number" ? ' min="0" step="1"' : "";
  return `<td><input name="${name}" type="${type}" value="${escapeHtml(value ?? "")}"${requiredAttr}${minAttr}></td>`;
}

function removeSelectedProgramacaoRows() {
  const rows = [...el("progItemsRows").querySelectorAll("tr[data-item-index]")];
  const selectedIndexes = rows
    .filter((row) => row.querySelector("[data-prog-item-select]")?.checked)
    .map((row) => Number(row.dataset.itemIndex));
  if (!selectedIndexes.length) {
    notify("Selecione uma linha para remover.", true);
    return;
  }
  state.programacaoItems = rows
    .filter((row) => !selectedIndexes.includes(Number(row.dataset.itemIndex)))
    .map(programacaoItemFromRow);
  renderProgramacaoItems();
}

function clearProgramacaoItems() {
  state.programacaoLoadedVendaIds = [];
  state.programacaoItems = [];
  renderProgramacaoItems();
}

function programacaoItemFromRow(row) {
  const item = {};
  for (const input of row.querySelectorAll("input[name]")) {
    item[input.name] = clean(input.value);
  }
  return {
    venda_id: Number(row.dataset.vendaId || 0) || 0,
    cod_cliente: item.cod_cliente || "",
    nome_cliente: item.nome_cliente || "",
    produto: item.produto || "",
    endereco: item.endereco || "",
    qnt_caixas: item.qnt_caixas === "" ? 0 : Number(item.qnt_caixas),
    kg: parseDecimal(item.kg),
    preco: parseDecimal(item.preco),
    vendedor: item.vendedor || "",
    pedido: item.pedido || "",
    obs: item.obs || "",
    recomendacao: row.dataset.recomendacao || "",
    ordem_sugerida: Number(row.dataset.ordemSugerida || 0) || null,
    distancia: Number(row.dataset.distancia || 0) || 0,
    confianca_localizacao: Number(row.dataset.confiancaLocalizacao || 0) || 0,
    carga_raiz_programacao: row.dataset.cargaRaizProgramacao || "",
    carga_origem_imediata: row.dataset.cargaOrigemImediata || "",
    transferencia_origem_id: row.dataset.transferenciaOrigemId || "",
  };
}

function collectProgramacaoItems({allowBlank = false} = {}) {
  const out = [];
  for (const row of el("progItemsRows").querySelectorAll("tr[data-item-index]")) {
    const item = programacaoItemFromRow(row);
    const hasAny = Object.entries(item).some(([key, value]) => key !== "venda_id" && clean(value));
    if (!hasAny && !allowBlank) continue;
    if (!item.cod_cliente || !item.nome_cliente) {
      throw new Error("Ha linhas sem COD CLIENTE ou NOME CLIENTE.");
    }
    out.push({
      cod_cliente: item.cod_cliente,
      nome_cliente: item.nome_cliente,
      produto: item.produto || null,
      endereco: item.endereco || null,
      qnt_caixas: item.qnt_caixas === "" ? 0 : Number(item.qnt_caixas),
      kg: parseDecimal(item.kg),
      preco: parseDecimal(item.preco),
      vendedor: item.vendedor || null,
      pedido: item.pedido || null,
      obs: item.obs || null,
      ordem_sugerida: item.ordem_sugerida || null,
      distancia: Number(item.distancia || 0) || 0,
      confianca_localizacao: Number(item.confianca_localizacao || 0) || 0,
      carga_raiz_programacao: item.carga_raiz_programacao || null,
      carga_origem_imediata: item.carga_origem_imediata || null,
      transferencia_origem_id: item.transferencia_origem_id || null,
    });
  }
  return out;
}

function programacaoPayloadFromForm() {
  const motoristaOption = el("progMotorista").selectedOptions[0];
  const tipoFrete = clean(el("progTransbordoModalidade").value) === "EMPRESA_BUSCA" ? "EMPRESA_BUSCA" : "CIF";
  const tipoEstimativa = tipoFrete === "EMPRESA_BUSCA" ? "CX" : "KG";
  el("progTipoEstimativa").value = tipoEstimativa;
  const estimativa = parseDecimal(el("progEstimativa").value);
  const ajudantes = [...el("progAjudantes").selectedOptions].map((option) => option.value).filter(Boolean);
  const payload = {
    codigo_programacao: clean(el("progCodigo").value) || null,
    motorista: motoristaOption?.dataset.nome || clean(el("progMotorista").value),
    motorista_codigo: motoristaOption?.dataset.codigo || null,
    veiculo: clean(el("progVeiculo").value),
    ajudantes,
    local_rota: clean(el("progLocalRota").value),
    tipo_estimativa: tipoEstimativa,
    kg_estimado: tipoEstimativa === "KG" ? estimativa : 0,
    caixas_estimado: tipoEstimativa === "CX" ? Math.trunc(estimativa) : 0,
    operacao_tipo: tipoEstimativa === "CX" ? "TRANSBORDO" : "VENDA",
    transbordo_modalidade: tipoFrete,
    transbordo_observacao: tipoEstimativa === "CX" ? clean(el("progTransbordoObs").value) : null,
    local_carregamento: clean(el("progCarregamento").value),
    adiantamento: parseDecimal(el("progAdiantamento").value),
    adiantamento_origem: clean(el("progAdiantamentoOrigem").value) || null,
    itens: collectProgramacaoItems(),
    venda_ids: programacaoVendaIdsFromRows(),
  };
  return payload;
}

function programacaoVendaIdsFromRows() {
  return [...el("progItemsRows").querySelectorAll("tr[data-item-index]")]
    .map((row) => Number(row.dataset.vendaId || 0))
    .filter(Boolean);
}

async function onSaveProgramacao(event) {
  event.preventDefault();
  try {
    if (!currentProgramacaoEditable()) {
      throw new Error("Este planejamento nao pode ser alterado. Edicao permitida somente nos status ATIVA e EM ROTA.");
    }
    const payload = programacaoPayloadFromForm();
    const usaVendasImportadas = Boolean((payload.venda_ids || []).length);
    if (payload.ajudantes.length !== 2) {
      throw new Error("Selecione exatamente 2 ajudantes.");
    }
    const saved = await apiRequest("/programacao/", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    fillProgramacaoForm(saved);
    notify(`Planejamento salvo: ${saved.codigo_programacao}`);
    const refreshTasks = [loadProgramacaoOptions(), loadProgramacoes(), loadDashboard()];
    if (usaVendasImportadas) {
      refreshTasks.push(refreshImportarVendasData({clearSearch: true}));
    }
    await Promise.all(refreshTasks);
    await openProgramacaoEmissionPreview(saved);
  } catch (error) {
    notify(error.message, true);
  }
}

async function openProgramacaoEmissionPreview(programacao) {
  const codigo = clean(programacao?.codigo_programacao || state.programacaoSelectedCodigo);
  if (!codigo) return;
  const docs = [];
  const itens = Array.isArray(programacao?.itens) ? programacao.itens : state.programacaoItems;
  const hasItens = Array.isArray(itens) && itens.length > 0;
  const adiantamento = Number(programacao?.adiantamento || 0);
  if (hasItens) {
    const pdf = await apiBlobRequest(`/programacao/${encodeURIComponent(codigo)}/pdf`);
    docs.push({label: "Planejamento", blob: pdf.blob, filename: pdf.filename || `PROGRAMACAO_${codigo}.pdf`});
  }
  if (adiantamento > 0) {
    const recibo = await apiBlobRequest(`/programacao/${encodeURIComponent(codigo)}/recibo-adiantamento-pdf`);
    docs.push({label: "Recibo", blob: recibo.blob, filename: recibo.filename || `RECIBO_ADIANTAMENTO_${codigo}.pdf`});
  }
  if (!docs.length) return;
  openPdfPreviewWindow(docs, `Previsualizacao do planejamento ${codigo}`);
}

async function deleteSelectedProgramacao() {
  const codigo = clean(el("progCodigo").value || state.programacaoSelectedCodigo);
  if (!codigo) {
    notify("Carregue um planejamento para excluir.", true);
    return;
  }
  if (!confirm(`Excluir planejamento ${codigo}?`)) return;
  try {
    await apiRequest(`/programacao/${encodeURIComponent(codigo)}?devolver_vendas=true`, {method: "DELETE"});
    notify(`Planejamento ${codigo} excluido.`);
    clearProgramacaoForm();
    await Promise.all([loadProgramacaoOptions(), loadProgramacoes(), loadDashboard(), refreshImportarVendasData({clearSearch: true})]);
  } catch (error) {
    notify(error.message, true);
  }
}

function parseDecimal(value) {
  const raw = clean(value);
  if (!raw) return 0;
  let text = raw.replace("R$", "").replaceAll(" ", "");
  const hasDot = text.includes(".");
  const hasComma = text.includes(",");
  if (hasDot && hasComma) {
    text = text.lastIndexOf(",") > text.lastIndexOf(".")
      ? text.replaceAll(".", "").replace(",", ".")
      : text.replaceAll(",", "");
  } else if (hasComma) {
    text = text.replaceAll(".", "").replace(",", ".");
  }
  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
}

function formatNumber(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) ? number.toFixed(2) : "0.00";
}

function formatPlainNumber(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return "0";
  return Number.isInteger(number) ? String(number) : number.toFixed(2);
}

function formatCurrencyBR(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return "R$ 0,00";
  return number.toLocaleString("pt-BR", {style: "currency", currency: "BRL"});
}

function formatMoneyInput(value) {
  return formatNumber(value).replace(".", ",");
}

function normalizeLocalRotaJs(value) {
  const text = clean(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  return text === "SERTAO" || text === "SERRA" ? text : "";
}

function renderModuleView(module) {
  el("moduleTitle").textContent = module.title;
  el("moduleStatus").textContent = "Rotina presente no desktop, em migracao para web";
  el("moduleDescription").textContent = module.description;
  el("moduleMigration").textContent = module.migration;

  const grid = el("moduleFeatureGrid");
  grid.innerHTML = "";
  module.features.forEach(([title, description]) => {
    const card = document.createElement("article");
    card.className = "feature-card";
    card.innerHTML = `
      <strong>${escapeHtml(title)}</strong>
      <p>${escapeHtml(description)}</p>
    `;
    grid.appendChild(card);
  });
}

async function loadDashboard() {
  try {
    const [home, audit] = await Promise.all([
      apiRequest("/home/overview?limit=40"),
      apiRequest("/audit-logs/?limit=8"),
    ]);
    state.home = home;
    state.audit = audit;
    renderHomeDashboard();
    renderDashboardAudit();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderHomeDashboard() {
  const dashboard = state.home || {};
  const metrics = dashboard.metrics || {};
  const pendencias = dashboard.pendencias || {};
  const sistema = dashboard.sistema || {};
  el("homeLastUpdate").textContent = `Atualizado em ${formatDate(dashboard.generated_at)}`;

  const metricGrid = el("homeMetricGrid");
  metricGrid.innerHTML = [
    homeMetricCard("Programacoes Ativas", metrics.programacoes_ativas || 0, "Rotas liberadas ou em andamento"),
    homeMetricCard("Pedidos Importados", metrics.vendas_importadas || 0, "Pedidos disponiveis no fluxo de importacao"),
    homeMetricCard("Clientes Ativos", metrics.clientes_ativos || 0, "Clientes vinculados a rotas abertas"),
    homeMetricCard("Rotas em Aberto", pendencias.rotas_abertas || 0, "Operacao ainda sem fechamento"),
    homeMetricCard("Prestacoes Pendentes", pendencias.prestacoes_pendentes || 0, "Prestacao ainda nao fechada"),
    homeMetricCard("Sem Despesa", pendencias.sem_despesa || 0, "Rotas ativas sem custo lancado"),
  ].join("");

  renderHomeRoutes(dashboard.rotas || []);
  renderHomePending(pendencias);
  renderHomeSystem(sistema);
}

function homeMetricCard(label, value, detail) {
  return `
    <article class="metric home-metric">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(formatPlainNumber(value))}</strong>
      <em>${escapeHtml(detail)}</em>
    </article>
  `;
}

function renderHomeRoutes(rotas) {
  const rows = el("homeRouteRows");
  rows.innerHTML = "";
  if (!rotas.length) {
    rows.appendChild(emptyRow(7, "Sem rotas ativas."));
    return;
  }
  rotas.forEach((rota) => {
    const tr = document.createElement("tr");
    const routeKey = rota.codigo_programacao || rota.programacao_id || rota.id || "";
    tr.className = "home-route-row";
    tr.tabIndex = 0;
    tr.innerHTML = `
      <td><strong>${escapeHtml(rota.codigo_exibicao || rota.codigo_programacao || (rota.programacao_id ? `ID ${rota.programacao_id}` : "-"))}</strong></td>
      <td>${escapeHtml(rota.motorista || "-")}</td>
      <td>${escapeHtml(rota.veiculo || "-")}</td>
      <td>${escapeHtml(rota.rota || "-")}</td>
      <td>${escapeHtml(rota.data || "-")}</td>
      <td>${escapeHtml(rota.status || "-")}</td>
      <td><button type="button" class="secondary" data-home-route="${escapeHtml(routeKey)}">Abrir</button></td>
    `;
    tr.addEventListener("click", (event) => {
      if (event.target.closest("button")) return;
      openHomeRoutePreview(routeKey);
    });
    tr.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openHomeRoutePreview(routeKey);
      }
    });
    rows.appendChild(tr);
  });
  rows.querySelectorAll("[data-home-route]").forEach((button) => {
    button.addEventListener("click", () => openHomeRoutePreview(button.dataset.homeRoute));
  });
}

function renderHomePending(pendencias) {
  el("homePendingGrid").innerHTML = [
    homeInfoItem("Rotas em aberto", formatPlainNumber(pendencias.rotas_abertas || 0)),
    homeInfoItem("Prestacoes pendentes", formatPlainNumber(pendencias.prestacoes_pendentes || 0)),
    homeInfoItem("Sem despesa lancada", formatPlainNumber(pendencias.sem_despesa || 0)),
  ].join("");
}

function renderHomeSystem(sistema) {
  const items = [
    homeInfoItem("Versao atual", sistema.versao_local || "-"),
  ];
  if (Object.prototype.hasOwnProperty.call(sistema, "versao_disponivel")) {
    items.push(homeInfoItem("Versao disponivel", sistema.versao_disponivel || "-"));
  }
  if (Object.prototype.hasOwnProperty.call(sistema, "api")) {
    items.push(homeInfoItem("API", sistema.api || "-"));
  }
  if (Object.prototype.hasOwnProperty.call(sistema, "ambiente")) {
    items.push(homeInfoItem("Ambiente", sistema.ambiente || "-"));
  }
  if (Object.prototype.hasOwnProperty.call(sistema, "banco")) {
    items.push(homeInfoItem("Banco", sistema.banco || "-"));
  }
  if (Object.prototype.hasOwnProperty.call(sistema, "data_hora_atual")) {
    items.push(homeInfoItem("Data/Hora", formatDate(sistema.data_hora_atual)));
  }
  el("homeSystemInfo").innerHTML = items.join("");
}

function homeInfoItem(label, value) {
  return `
    <article class="home-info-item">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `;
}

async function openHomeRoutePreview(codigoProgramacao) {
  const codigo = clean(codigoProgramacao);
  if (!codigo) return;
  try {
    const data = await apiRequest(`/home/rotas/${encodeURIComponent(codigo)}/preview`);
    renderHomeRoutePreview(data);
    el("homeRouteDialog").showModal();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderHomeRoutePreview(data) {
  state.homeRoutePreview = data || {};
  const programacao = data.programacao || {};
  const resumo = data.resumo || {};
  el("homeRouteDialogTitle").textContent = `Rota ${programacao.codigo_programacao || "-"}`;
  el("homeRouteDialogSubtitle").textContent = `${programacao.motorista || "-"} | ${programacao.veiculo || "-"} | ${programacao.rota || "-"}`;
  renderHomeRouteHeader(programacao, resumo);
  el("homeRouteSummary").innerHTML = [
    homeInfoItem("Clientes", formatPlainNumber(resumo.clientes || 0)),
    homeInfoItem("Caixas", formatPlainNumber(resumo.caixas || 0)),
    homeInfoItem("KG", formatPlainNumber(resumo.kg || 0)),
    homeInfoItem("Entregues", formatPlainNumber(resumo.entregues || 0)),
    homeInfoItem("Pendentes", formatPlainNumber(resumo.pendentes || 0)),
    homeInfoItem("Com GPS", formatPlainNumber(resumo.com_localizacao || 0)),
    homeInfoItem("Recebido", formatCurrencyBR(resumo.recebido || 0)),
    homeInfoItem("Despesas", formatCurrencyBR(resumo.despesas || 0)),
    homeInfoItem("Saldo", formatCurrencyBR(resumo.saldo || 0)),
  ].join("");

  const rows = el("homeRoutePreviewRows");
  rows.innerHTML = "";
  const itens = data.itens || [];
  if (!itens.length) {
    rows.appendChild(emptyRow(10, "Sem clientes nesta rota."));
  } else {
    itens.forEach((item, index) => {
      const tr = document.createElement("tr");
      tr.className = "home-route-preview-row";
      tr.innerHTML = `
        <td>${escapeHtml(item.cod_cliente || "-")}</td>
        <td>${escapeHtml(item.nome_cliente || "-")}</td>
        <td>
          <button type="button" class="link-button home-trace-link" data-home-trace-index="${index}">
            ${escapeHtml(item.pedido || "-")}
          </button>
        </td>
        <td>${escapeHtml(formatPlainNumber(item.caixas || 0))}</td>
        <td>${escapeHtml(formatPlainNumber(item.kg || 0))}</td>
        <td>${escapeHtml(formatCurrencyBR(item.preco || 0))}</td>
        <td>${homePedidoStatusPill(item.status_pedido || "-")}</td>
        <td>${homeTraceGpsBadge(item)}</td>
        <td>${escapeHtml(formatCurrencyBR(item.valor_recebido || 0))}</td>
        <td><button type="button" class="secondary" data-home-trace-index="${index}">Ver</button></td>
      `;
      tr.addEventListener("dblclick", () => openHomeTraceabilityByIndex(index));
      rows.appendChild(tr);
    });
    rows.querySelectorAll("[data-home-trace-index]").forEach((button) => {
      button.addEventListener("click", () => openHomeTraceabilityByIndex(Number(button.dataset.homeTraceIndex)));
    });
  }

  renderHomeMiniList(
    "homeRouteRecebimentos",
    data.recebimentos || [],
    (item) => `${item.cod_cliente || "-"} - ${item.nome_cliente || "-"}: ${formatCurrencyBR(item.valor || 0)}`
  );
  renderHomeMiniList(
    "homeRouteDespesas",
    data.despesas || [],
    (item) => `${item.categoria || "ROTA"} - ${item.descricao || "-"}: ${formatCurrencyBR(item.valor || 0)}`
  );
}

function renderHomeRouteHeader(programacao, resumo) {
  const kgCarregado = Number(programacao.kg_carregado || programacao.nf_kg_carregado || resumo.kg || 0);
  const caixasCarregadas = Number(programacao.caixas_carregadas || programacao.nf_caixas || resumo.caixas_programadas || resumo.caixas || 0);
  const nfKg = Number(programacao.nf_kg || 0);
  const saldoNf = Number(programacao.nf_saldo || 0) || (nfKg > 0 && kgCarregado > 0 ? Math.max(nfKg - kgCarregado, 0) : 0);
  const precoNf = Number(programacao.nf_preco || programacao.preco_nf || 0);
  const descontoFornecedor = Number(programacao.desconto_fornecedor || programacao.nf_saldo_valor || 0) || (saldoNf * precoNf);
  const kmInicial = Number(programacao.km_inicial || 0);
  const kmFinal = Number(programacao.km_final || 0);
  const kmTexto = kmInicial || kmFinal
    ? `${formatPlainNumber(kmInicial)} -> ${formatPlainNumber(kmFinal)}`
    : "-";
  const saidaTexto = formatRouteDateTime(
    programacao.saida_data || programacao.data_saida,
    programacao.saida_hora || programacao.hora_saida
  );
  const carregamentoPeriodo = formatLoadingPeriod(programacao.inicio_carregamento, programacao.fim_carregamento);
  const caixaFinal = Number(programacao.aves_caixa_final || programacao.qnt_aves_caixa_final || 0);
  const headerItems = [
    ["Status", programacao.status || "-"],
    ["Motorista", programacao.motorista || "-"],
    ["Equipe", programacao.equipe || "-"],
    ["NF", programacao.num_nf || "-"],
    ["Adiantamento", formatCurrencyBR(programacao.adiantamento || 0)],
    ["Local da rota", programacao.rota || "-"],
    ["Saida", saidaTexto],
    ["Carregou em", programacao.local_carregamento || "-"],
    ["KG carregado", formatPlainNumber(kgCarregado)],
    ["Caixas carregadas", formatPlainNumber(caixasCarregadas)],
    ["Caixas entregues", formatPlainNumber(resumo.caixas_entregues || 0)],
    ["Saldo NF desconto", formatPlainNumber(saldoNf), true],
    ["Desconto fornecedor", formatCurrencyBR(descontoFornecedor), true],
    ["KM", kmTexto],
    ["Media carregada", formatPlainNumber(programacao.media || 0)],
    ["Caixa final", caixaFinal > 0 ? formatPlainNumber(caixaFinal) : "-"],
    ["Carregamento", carregamentoPeriodo],
    ["Fonte", "API ONLINE", true],
  ];
  el("homeRouteHeader").innerHTML = headerItems.map(([label, value, destaque]) => homeRouteHeaderItem(label, value, destaque)).join("");
}

function homeRouteHeaderItem(label, value, destaque = false) {
  return `
    <div class="home-route-header-item${destaque ? " featured" : ""}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `;
}

function formatRouteDateTime(dateValue, timeValue) {
  const dateText = clean(dateValue);
  const timeText = clean(timeValue);
  const joined = [dateText, timeText].filter(Boolean).join(" ");
  return joined || "-";
}

function formatLoadingPeriod(startValue, endValue) {
  const start = clean(startValue);
  const end = clean(endValue);
  if (start && end) return `${start} -> ${end}`;
  if (start) return `Inicio ${start}`;
  if (end) return `Fim ${end}`;
  return "-";
}

function homePedidoStatusPill(status) {
  const label = clean(status) || "PENDENTE";
  const normalized = label.toUpperCase();
  let css = "status-pendente";
  if (["ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO"].includes(normalized)) css = "status-entregue";
  if (normalized === "ALTERADO") css = "status-alterado";
  if (["CANCELADO", "CANCELADA"].includes(normalized)) css = "status-cancelado";
  return `<span class="status-pill ${css}">${escapeHtml(label)}</span>`;
}

function homeTraceGpsBadge(item) {
  return item?.tem_localizacao
    ? '<span class="status-pill status-entregue">OK</span>'
    : '<span class="status-pill status-pendente">PENDENTE</span>';
}

function openHomeTraceabilityByIndex(index) {
  const item = state.homeRoutePreview?.itens?.[index];
  if (!item) return;
  renderHomeTraceability(item);
  const dialog = el("homeTraceDialog");
  if (!dialog.open) dialog.showModal();
}

function renderHomeTraceability(item) {
  const programacao = state.homeRoutePreview?.programacao || {};
  const clienteLabel = `${item.cod_cliente || "-"} - ${item.nome_cliente || "-"}`;
  const pedidoLabel = item.pedido || "-";
  el("homeTraceDialogTitle").textContent = `Pedido ${pedidoLabel}`;
  el("homeTraceDialogSubtitle").textContent = clienteLabel;
  el("homeTraceSummary").innerHTML = [
    homeInfoItem("Cliente", clienteLabel),
    homeInfoItem("Status", item.status_pedido || "PENDENTE"),
    homeInfoItem("Programacao", programacao.codigo_programacao || "-"),
    homeInfoItem("Motorista", programacao.motorista || "-"),
    homeInfoItem("Veiculo", programacao.veiculo || "-"),
    homeInfoItem("Ultima alteracao", item.alterado_em || "-"),
  ].join("");

  const caixasOriginal = Number(item.caixas_original || 0);
  const caixasAtual = Number(item.caixas_atual ?? item.caixas ?? 0);
  const precoOriginal = Number(item.preco_original || 0);
  const precoAtual = Number(item.preco_atual ?? item.preco ?? 0);
  const kg = Number(item.kg || 0);
  const mediaCliente = caixasOriginal > 0 ? kg / caixasOriginal : 0;

  renderHomeTraceList("homeTraceDelivery", [
    ["Caixas origem", formatPlainNumber(caixasOriginal)],
    ["Caixas atual", `${formatPlainNumber(caixasAtual)} (${formatSignedNumber(caixasAtual - caixasOriginal)})`],
    ["Peso KG", formatPlainNumber(kg)],
    ["Peso previsto", formatPlainNumber(item.peso_previsto || 0)],
    ["Media cliente KG/CX", formatPlainNumber(mediaCliente)],
    ["Media aplicada", item.media_aplicada === null || item.media_aplicada === undefined ? "-" : formatPlainNumber(item.media_aplicada)],
    ["Preco origem", formatCurrencyBR(precoOriginal)],
    ["Preco atual", `${formatCurrencyBR(precoAtual)} (${formatSignedCurrency(precoAtual - precoOriginal)})`],
    ["Valor origem", formatCurrencyBR(item.valor_original || 0)],
    ["Valor atual", formatCurrencyBR(item.valor_atual || 0)],
    ["Recebido", `${formatCurrencyBR(item.valor_recebido || 0)} | ${item.forma_recebimento || "-"}`],
    ["Obs recebimento", item.obs_recebimento || "-"],
    ["Ocorrencias unidades", formatPlainNumber(item.mortalidade_aves || 0)],
  ]);

  renderHomeTraceList("homeTraceTracking", [
    ["NF", programacao.num_nf || "-"],
    ["Local carregado", programacao.local_carregamento || "-"],
    ["Local rota", programacao.rota || "-"],
    ["Equipe", programacao.equipe || "-"],
    ["Origem do status", item.status_origem || "-"],
    ["Alterado por", item.alterado_por || "-"],
    ["Hora entrega/alteracao", item.alterado_em || "-"],
    ["Tipo alteracao", item.alteracao_tipo || "-"],
    ["Detalhe alteracao", item.alteracao_detalhe || "-"],
    ["Vendedor", item.vendedor || "-"],
    ["Produto", item.produto || "-"],
    ["Ordem sugerida", item.ordem_sugerida ? formatPlainNumber(item.ordem_sugerida) : "-"],
    ["ETA previsto", item.eta || "-"],
    ["Distancia estimada", item.distancia ? formatPlainNumber(item.distancia) : "-"],
    ["Confianca localizacao", formatConfidence(item.confianca_localizacao)],
  ]);

  renderHomeTraceList("homeTraceLocation", [
    ["Latitude", formatCoord(item.lat_evento) || "-"],
    ["Longitude", formatCoord(item.lon_evento) || "-"],
    ["Latitude entrega", formatCoord(item.lat_entrega) || "-"],
    ["Longitude entrega", formatCoord(item.lon_entrega) || "-"],
    ["Precisao entrega", item.accuracy_entrega ? `${formatPlainNumber(item.accuracy_entrega)} m` : "-"],
    ["Capturado em", item.timestamp_entrega || item.alterado_em || "-"],
    ["Endereco evento", item.endereco_evento || item.endereco || "-"],
    ["Cidade", item.cidade_evento || "-"],
    ["Bairro", item.bairro_evento || "-"],
    ["Localizacao", item.localizacao || "-"],
    ["Foto ocorrencia", item.foto_mortalidade_path || item.mortalidade_foto_path || (item.foto_mortalidade && (item.foto_mortalidade.storage_path || item.foto_mortalidade.path_local || item.foto_mortalidade.arquivo_nome)) || "-"],
  ]);

  renderHomeTraceChecklist(item);
  renderHomeTraceMap(item);
}

function renderHomeTraceList(id, rows) {
  el(id).innerHTML = rows.map(([label, value]) => `
    <div class="home-trace-line">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join("");
}

function renderHomeTraceChecklist(item) {
  const status = clean(item.status_pedido).toUpperCase();
  const checks = [
    ["Status recebido do APP", Boolean(status)],
    ["Hora da alteracao", Boolean(clean(item.alterado_em))],
    ["Pedido entregue", ["ENTREGUE", "FINALIZADO", "FINALIZADA", "CONCLUIDO"].includes(status)],
    ["Latitude / Longitude salvas", Boolean(item.tem_localizacao)],
    ["Endereco / cidade / bairro salvos", Boolean(clean(item.endereco_evento) || clean(item.cidade_evento) || clean(item.bairro_evento))],
    ["Recebimento registrado", Number(item.valor_recebido || 0) > 0],
    ["Alteracao do pedido capturada", Boolean(clean(item.alteracao_tipo) || clean(item.alteracao_detalhe) || Number(item.delta_caixas || 0) !== 0 || Number(item.delta_preco || 0) !== 0)],
    ["Sugestao de roteirizacao", Boolean(Number(item.ordem_sugerida || 0) > 0 || clean(item.eta) || Number(item.distancia || 0) > 0)],
  ];
  el("homeTraceChecklist").innerHTML = checks.map(([label, ok]) => `
    <div class="home-trace-check">
      <span>${escapeHtml(label)}</span>
      <strong class="${ok ? "ok" : "pending"}">${ok ? "OK" : "PENDENTE"}</strong>
    </div>
  `).join("");
}

function renderHomeTraceMap(item) {
  const lat = Number(item.lat_evento);
  const lon = Number(item.lon_evento);
  const hasGps = Number.isFinite(lat) && Number.isFinite(lon);
  const link = el("homeTraceMapLink");
  if (hasGps) {
    link.href = item.map_url || `https://www.google.com/maps?q=${encodeURIComponent(`${lat},${lon}`)}`;
    link.classList.remove("hidden");
  } else {
    link.href = "#";
    link.classList.add("hidden");
  }
  const card = el("homeTraceMapCard");
  if (!hasGps) {
    card.innerHTML = `
      <div class="home-trace-map-empty">
        <strong>Sem latitude/longitude capturada</strong>
        <span>${escapeHtml(item.endereco_evento || item.endereco || "Endereco nao informado")}</span>
      </div>
    `;
    return;
  }
  card.innerHTML = `
    <div class="home-trace-map-canvas">
      <span class="home-trace-map-pin" style="${homeTraceMapPinStyle(lat, lon)}"></span>
      <strong>${escapeHtml(item.nome_cliente || "-")}</strong>
      <em>${escapeHtml(`${lat.toFixed(6)}, ${lon.toFixed(6)}`)}</em>
    </div>
  `;
}

function homeTraceMapPinStyle(lat, lon) {
  const latMin = -34.0;
  const latMax = 5.5;
  const lonMin = -74.0;
  const lonMax = -32.0;
  const x = Math.max(4, Math.min(96, ((lon - lonMin) / (lonMax - lonMin)) * 100));
  const y = Math.max(8, Math.min(92, (1 - ((lat - latMin) / (latMax - latMin))) * 100));
  return `left: ${x}%; top: ${y}%;`;
}

function formatSignedNumber(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return "+0";
  return `${number >= 0 ? "+" : ""}${formatPlainNumber(number)}`;
}

function formatSignedCurrency(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return formatCurrencyBR(0);
  return `${number >= 0 ? "+" : "-"}${formatCurrencyBR(Math.abs(number))}`;
}

function formatConfidence(value) {
  if (value === null || value === undefined || value === "") return "-";
  const number = Number(value);
  if (!Number.isFinite(number)) return "-";
  const normalized = number <= 1 ? number * 100 : number;
  return `${formatPlainNumber(normalized)}%`;
}

function renderHomeMiniList(id, items, labelFactory) {
  const node = el(id);
  node.innerHTML = "";
  if (!items.length) {
    const empty = document.createElement("p");
    empty.className = "home-mini-empty";
    empty.textContent = "Sem registros.";
    node.appendChild(empty);
    return;
  }
  items.slice(0, 8).forEach((item) => {
    const div = document.createElement("div");
    div.className = "home-mini-line";
    div.textContent = labelFactory(item);
    node.appendChild(div);
  });
}

function closeHomeRouteDialog() {
  closeHomeTraceDialog();
  el("homeRouteDialog").close();
}

function closeHomeTraceDialog() {
  const dialog = el("homeTraceDialog");
  if (dialog.open) dialog.close();
}

function renderMetrics() {
  const activeUsersMetric = document.getElementById("metricActiveUsers");
  const inactiveUsersMetric = document.getElementById("metricInactiveUsers");
  const auditEventsMetric = document.getElementById("metricAuditEvents");
  if (!activeUsersMetric || !inactiveUsersMetric || !auditEventsMetric) return;
  const active = state.users.filter((user) => user.is_active).length;
  activeUsersMetric.textContent = active;
  inactiveUsersMetric.textContent = state.users.length - active;
  auditEventsMetric.textContent = state.audit.length;
}

function renderDashboardAudit() {
  const rows = el("dashboardAuditRows");
  rows.innerHTML = "";
  if (!state.audit.length) {
    rows.appendChild(emptyRow(4, "Sem eventos."));
    return;
  }
  state.audit.forEach((log) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(log.action)}</td>
      <td>${escapeHtml(entityLabel(log))}</td>
      <td>${escapeHtml(log.user_id || "-")}</td>
      <td>${escapeHtml(formatDate(log.created_at))}</td>
    `;
    rows.appendChild(tr);
  });
}

async function loadUsers() {
  try {
    const includeInactive = el("includeInactive").checked;
    state.users = await apiRequest(`/users/?include_inactive=${includeInactive}`);
    renderUsers();
    renderMetrics();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderUsers() {
  const rows = el("usersRows");
  rows.innerHTML = "";
  if (!state.users.length) {
    rows.appendChild(emptyRow(7, "Sem usuarios."));
    return;
  }

  state.users.forEach((user) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${user.id}</td>
      <td>${escapeHtml(user.username)}</td>
      <td>${escapeHtml(user.nome)}</td>
      <td>${escapeHtml(user.permissoes)}</td>
      <td>${statusPill(user.is_active)}</td>
      <td>${escapeHtml(user.telefone || user.cpf || "-")}</td>
      <td>
        <div class="actions">
          <button type="button" class="secondary" data-user-action="edit" data-user-id="${user.id}">Editar</button>
          ${user.is_active
            ? `<button type="button" class="danger" data-user-action="deactivate" data-user-id="${user.id}">Desativar</button>`
            : `<button type="button" class="warning" data-user-action="reactivate" data-user-id="${user.id}">Reativar</button>`}
        </div>
      </td>
    `;
    rows.appendChild(tr);
  });

  rows.querySelectorAll("[data-user-action]").forEach((button) => {
    button.addEventListener("click", () => handleUserAction(button.dataset.userAction, Number(button.dataset.userId)));
  });
}

async function onCreateUser(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const payload = userPayloadFromForm(new FormData(form), {creating: true});

  try {
    await apiRequest("/users/", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    form.reset();
    form.elements.is_active.checked = true;
    notify("Usuario criado.");
    await Promise.all([loadUsers(), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

function handleUserAction(action, userId) {
  const user = state.users.find((item) => item.id === userId);
  if (!user) return;

  if (action === "edit") {
    openUserDialog(user);
    return;
  }

  if (action === "deactivate") {
    deactivateUser(user);
    return;
  }

  if (action === "reactivate") {
    reactivateUser(user);
  }
}

function openUserDialog(user) {
  const form = el("editUserForm");
  form.elements.id.value = user.id;
  form.elements.username.value = user.username || "";
  form.elements.nome.value = user.nome || "";
  form.elements.password.value = "";
  form.elements.permissoes.value = user.permissoes || "OPERADOR";
  form.elements.cpf.value = user.cpf || "";
  form.elements.telefone.value = user.telefone || "";
  form.elements.idade.value = user.idade ?? "";
  form.elements.is_active.checked = Boolean(user.is_active);
  el("userDialog").showModal();
}

function closeUserDialog() {
  el("userDialog").close();
}

async function onSaveUser(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const userId = Number(form.elements.id.value);
  const payload = userPayloadFromForm(new FormData(form), {creating: false});

  try {
    await apiRequest(`/users/${userId}`, {
      method: "PATCH",
      body: JSON.stringify(payload),
    });
    closeUserDialog();
    notify("Usuario atualizado.");
    await Promise.all([loadUsers(), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

async function deactivateUser(user) {
  if (!confirm(`Desativar ${user.username}?`)) return;
  try {
    await apiRequest(`/users/${user.id}`, {method: "DELETE"});
    notify("Usuario desativado.");
    await Promise.all([loadUsers(), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

async function reactivateUser(user) {
  try {
    await apiRequest(`/users/${user.id}`, {
      method: "PATCH",
      body: JSON.stringify({is_active: true}),
    });
    notify("Usuario reativado.");
    await Promise.all([loadUsers(), loadDashboard()]);
  } catch (error) {
    notify(error.message, true);
  }
}

function userPayloadFromForm(formData, options) {
  const payload = {
    username: clean(formData.get("username")),
    nome: clean(formData.get("nome")),
    permissoes: clean(formData.get("permissoes")) || "OPERADOR",
    is_active: formData.get("is_active") === "on",
    cpf: clean(formData.get("cpf")) || null,
    telefone: clean(formData.get("telefone")) || null,
  };

  const idade = clean(formData.get("idade"));
  payload.idade = idade === "" ? null : Number(idade);

  const password = clean(formData.get("password"));
  if (options.creating || password) {
    payload.password = password;
  }
  return payload;
}

async function loadAudit() {
  try {
    const form = new FormData(el("auditFilterForm"));
    const params = new URLSearchParams({limit: "100"});
    for (const key of ["action", "entity_type", "entity_id"]) {
      const value = clean(form.get(key));
      if (value) params.set(key, value);
    }
    state.audit = await apiRequest(`/audit-logs/?${params.toString()}`);
    renderAudit();
    renderMetrics();
  } catch (error) {
    notify(error.message, true);
  }
}

function renderAudit() {
  const rows = el("auditRows");
  rows.innerHTML = "";
  if (!state.audit.length) {
    rows.appendChild(emptyRow(8, "Sem eventos."));
    return;
  }

  state.audit.forEach((log) => {
    const tr = document.createElement("tr");
    const metadata = JSON.stringify(log.metadata || {}, null, 2);
    tr.innerHTML = `
      <td>${log.id}</td>
      <td>${escapeHtml(log.action)}</td>
      <td>${escapeHtml(log.severity)}</td>
      <td>${escapeHtml(entityLabel(log))}</td>
      <td>${escapeHtml(log.user_id || "-")}</td>
      <td>${escapeHtml(log.ip_address || "-")}</td>
      <td>${escapeHtml(formatDate(log.created_at))}</td>
      <td>
        <details class="metadata">
          <summary>Ver</summary>
          <pre>${escapeHtml(metadata)}</pre>
        </details>
      </td>
    `;
    rows.appendChild(tr);
  });
}

function entityLabel(log) {
  if (!log.entity_type && !log.entity_id) return "-";
  return `${log.entity_type || ""}${log.entity_id ? `:${log.entity_id}` : ""}`;
}

function statusPill(isActive) {
  const css = isActive ? "status-active" : "status-inactive";
  const label = isActive ? "Ativo" : "Inativo";
  return `<span class="status-pill ${css}">${label}</span>`;
}

function emptyRow(colspan, message) {
  const tr = document.createElement("tr");
  const td = document.createElement("td");
  td.colSpan = colspan;
  td.className = "empty-row";
  td.textContent = message;
  tr.appendChild(td);
  return tr;
}

function emptyDiv(message) {
  const div = document.createElement("div");
  div.className = "empty-row";
  div.textContent = message;
  return div;
}

function installTableEnhancer() {
  const schedule = () => {
    window.clearTimeout(tableEnhanceTimer);
    tableEnhanceTimer = window.setTimeout(enhanceVisibleTables, 80);
  };
  const observer = new MutationObserver(schedule);
  observer.observe(document.body, {childList: true, subtree: true});
  schedule();
}

function enhanceVisibleTables() {
  document.querySelectorAll(".view:not(.hidden) table").forEach((table) => {
    if (table.classList.contains("no-column-filter")) return;
    const thead = table.tHead;
    const tbody = table.tBodies && table.tBodies[0];
    if (!thead || !tbody || !thead.rows.length) return;
    const headerRow = [...thead.rows].find((row) => !row.classList.contains("column-filter-row"));
    if (!headerRow) return;
    const colCount = headerRow.cells.length;
    if (colCount < 2) return;
    thead.querySelectorAll(".column-filter-row").forEach((row) => row.remove());
    ensureColumnVisibilityControl(table, headerRow);
    [...headerRow.cells].forEach((cell, index) => enhanceTableHeaderCell(table, cell, index));
    applyColumnVisibility(table);
    applyColumnFilters(table);
  });
}

function enhanceTableHeaderCell(table, cell, index) {
  const title = getColumnTitle(cell);
  const normalizedTitle = title.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  if (!title || normalizedTitle === "acoes" || index === 0 && title === "") return;
  cell.classList.add("column-filter-header");
  cell.dataset.columnIndex = String(index);
  cell.tabIndex = 0;
  cell.title = "Clique para ordenar. Passe o mouse para filtrar.";
  if (!cell.querySelector(":scope > .column-filter-button")) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "column-filter-button";
    button.dataset.columnIndex = String(index);
    button.setAttribute("aria-label", `Filtrar ${title}`);
    button.textContent = "v";
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openColumnFilterMenu(table, cell, index);
    });
    cell.appendChild(button);
  }
  if (cell.dataset.sortBound) return;
  cell.dataset.sortBound = "1";
  cell.addEventListener("click", (event) => {
    if (event.target.closest(".column-filter-button")) return;
    sortTableByColumn(table, index, cell);
  });
  cell.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    sortTableByColumn(table, index, cell);
  });
}

function getColumnTitle(cell) {
  return clean([...cell.childNodes]
    .filter((node) => !(node.classList && (
      node.classList.contains("column-filter-button") ||
      node.classList.contains("column-visibility-button")
    )))
    .map((node) => node.textContent || "")
    .join(" "));
}

function ensureColumnVisibilityControl(table, headerRow) {
  const columns = getTableColumns(table, headerRow);
  const configurableColumns = columns.filter((column) => column.canToggle);
  if (configurableColumns.length < 2) return;
  const wrap = table.closest(".table-wrap, .compact-table-wrap, .compras-table-wrap, .importar-vendas-table-wrap") || table.parentElement;
  if (!wrap) return;
  wrap.querySelectorAll(":scope > .column-visibility-button").forEach((item) => {
    if (item.__columnVisibilityTable && item.__columnVisibilityTable !== table && !item.__columnVisibilityTable.isConnected) item.remove();
  });
  let button = tableColumnVisibilityButtons.get(table);
  if (!button || !button.isConnected) {
    button = document.createElement("button");
    button.type = "button";
    button.className = "column-visibility-button";
    button.textContent = "*";
    button.title = "Escolher colunas";
    button.setAttribute("aria-label", "Escolher colunas visiveis");
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openColumnVisibilityMenu(table, button);
    });
    button.__columnVisibilityTable = table;
    wrap.appendChild(button);
    tableColumnVisibilityButtons.set(table, button);
  }
  wrap.classList.add("has-column-visibility");
  restoreColumnVisibility(table);
}

function getTableColumns(table, headerRow = null) {
  const row = headerRow || [...(table.tHead?.rows || [])].find((item) => !item.classList.contains("column-filter-row"));
  if (!row) return [];
  return [...row.cells].map((cell, index) => {
    const title = getColumnTitle(cell) || `Coluna ${index + 1}`;
    const normalized = title.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    return {
      index,
      key: `${index}:${title}`,
      title,
      canToggle: title.trim() !== "" && normalized !== "acoes",
    };
  });
}

function getTableVisibilityKey(table) {
  const view = table.closest(".view");
  const viewId = view?.id || "global";
  const tables = view ? [...view.querySelectorAll("table")] : [...document.querySelectorAll("table")];
  const tableIndex = Math.max(0, tables.indexOf(table));
  const headers = getTableColumns(table).map((column) => column.title).join("|");
  return `${viewId}:${table.id || tableIndex}:${headers}`;
}

function readColumnVisibilityStore() {
  try {
    return JSON.parse(localStorage.getItem(COLUMN_VISIBILITY_STORAGE_KEY) || "{}") || {};
  } catch (_error) {
    return {};
  }
}

function writeColumnVisibilityStore(store) {
  try {
    localStorage.setItem(COLUMN_VISIBILITY_STORAGE_KEY, JSON.stringify(store));
  } catch (_error) {
    // Ignore storage errors; the current screen still keeps the choice in memory.
  }
}

function restoreColumnVisibility(table) {
  if (tableColumnVisibility.has(table)) return;
  const store = readColumnVisibilityStore();
  tableColumnVisibility.set(table, new Set(store[getTableVisibilityKey(table)] || []));
}

function saveColumnVisibility(table) {
  const store = readColumnVisibilityStore();
  const key = getTableVisibilityKey(table);
  const hidden = [...(tableColumnVisibility.get(table) || new Set())];
  if (hidden.length) store[key] = hidden;
  else delete store[key];
  writeColumnVisibilityStore(store);
}

function openColumnVisibilityMenu(table, button) {
  closeColumnFilterMenu();
  restoreColumnVisibility(table);
  const columns = getTableColumns(table);
  const toggleColumns = columns.filter((column) => column.canToggle);
  const hidden = new Set(tableColumnVisibility.get(table) || []);
  const menu = document.createElement("div");
  menu.className = "column-filter-menu column-visibility-menu";
  const rect = button.getBoundingClientRect();
  menu.style.left = `${Math.max(8, rect.left)}px`;
  menu.style.top = `${rect.bottom + 4}px`;

  const title = document.createElement("div");
  title.className = "column-visibility-title";
  title.textContent = "Colunas visiveis";
  menu.appendChild(title);

  const list = document.createElement("div");
  list.className = "column-filter-list";
  toggleColumns.forEach((column) => {
    const label = document.createElement("label");
    label.className = "column-filter-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = column.key;
    checkbox.checked = !hidden.has(column.key);
    label.append(checkbox, document.createTextNode(column.title));
    list.appendChild(label);
  });
  menu.appendChild(list);

  const actions = document.createElement("div");
  actions.className = "column-filter-actions";
  const showAllButton = document.createElement("button");
  showAllButton.type = "button";
  showAllButton.className = "secondary";
  showAllButton.textContent = "Todas";
  const applyButton = document.createElement("button");
  applyButton.type = "button";
  applyButton.className = "primary";
  applyButton.textContent = "Aplicar";
  actions.append(showAllButton, applyButton);
  menu.appendChild(actions);

  showAllButton.addEventListener("click", () => {
    tableColumnVisibility.set(table, new Set());
    saveColumnVisibility(table);
    applyColumnVisibility(table);
    closeColumnFilterMenu();
  });
  applyButton.addEventListener("click", () => {
    const checked = new Set([...list.querySelectorAll("input[type='checkbox']:checked")].map((checkbox) => checkbox.value));
    if (!checked.size) return;
    const nextHidden = new Set(toggleColumns.filter((column) => !checked.has(column.key)).map((column) => column.key));
    tableColumnVisibility.set(table, nextHidden);
    saveColumnVisibility(table);
    applyColumnVisibility(table);
    closeColumnFilterMenu();
  });

  document.body.appendChild(menu);
  activeColumnFilterMenu = menu;
  window.setTimeout(() => {
    document.addEventListener("click", closeColumnFilterMenuOnOutside);
    document.addEventListener("keydown", closeColumnFilterMenuOnEscape);
  }, 0);
}

function applyColumnVisibility(table) {
  restoreColumnVisibility(table);
  const hidden = tableColumnVisibility.get(table) || new Set();
  const columns = getTableColumns(table);
  columns.forEach((column) => {
    const isHidden = column.canToggle && hidden.has(column.key);
    table.querySelectorAll("tr").forEach((row) => {
      const cell = row.cells[column.index];
      if (!cell || cell.colSpan > 1) return;
      cell.classList.toggle("column-visibility-hidden", isHidden);
    });
  });
  const button = tableColumnVisibilityButtons.get(table);
  if (button) button.classList.toggle("column-visibility-active", hidden.size > 0);
}

function openColumnFilterMenu(table, headerCell, index) {
  closeColumnFilterMenu();
  const values = getColumnValues(table, index);
  const current = tableColumnFilters.get(table)?.get(index) || null;
  const selected = current ? new Set(current) : new Set(values);
  const menu = document.createElement("div");
  menu.className = "column-filter-menu";
  const rect = headerCell.getBoundingClientRect();
  menu.style.left = `${Math.max(8, rect.right - 220)}px`;
  menu.style.top = `${rect.bottom + 4}px`;

  const allLabel = document.createElement("label");
  allLabel.className = "column-filter-option column-filter-all";
  const allCheck = document.createElement("input");
  allCheck.type = "checkbox";
  allCheck.checked = !current || selected.size === values.length;
  allLabel.append(allCheck, document.createTextNode("(Todos)"));
  menu.appendChild(allLabel);

  const list = document.createElement("div");
  list.className = "column-filter-list";
  values.forEach((value) => {
    const label = document.createElement("label");
    label.className = "column-filter-option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = value;
    checkbox.checked = selected.has(value);
    label.append(checkbox, document.createTextNode(value || "(Vazio)"));
    list.appendChild(label);
  });
  menu.appendChild(list);

  const actions = document.createElement("div");
  actions.className = "column-filter-actions";
  const clearButton = document.createElement("button");
  clearButton.type = "button";
  clearButton.className = "secondary";
  clearButton.textContent = "Limpar";
  const applyButton = document.createElement("button");
  applyButton.type = "button";
  applyButton.className = "primary";
  applyButton.textContent = "Aplicar";
  actions.append(clearButton, applyButton);
  menu.appendChild(actions);

  allCheck.addEventListener("change", () => {
    list.querySelectorAll("input[type='checkbox']").forEach((checkbox) => {
      checkbox.checked = allCheck.checked;
    });
  });
  list.addEventListener("change", () => {
    const checks = [...list.querySelectorAll("input[type='checkbox']")];
    allCheck.checked = checks.every((checkbox) => checkbox.checked);
  });
  clearButton.addEventListener("click", () => {
    setColumnFilter(table, index, null);
    closeColumnFilterMenu();
  });
  applyButton.addEventListener("click", () => {
    const checked = [...list.querySelectorAll("input[type='checkbox']:checked")].map((checkbox) => checkbox.value);
    setColumnFilter(table, index, checked.length === values.length ? null : checked);
    closeColumnFilterMenu();
  });

  document.body.appendChild(menu);
  activeColumnFilterMenu = menu;
  window.setTimeout(() => {
    document.addEventListener("click", closeColumnFilterMenuOnOutside);
    document.addEventListener("keydown", closeColumnFilterMenuOnEscape);
  }, 0);
}

function closeColumnFilterMenuOnOutside(event) {
  if (activeColumnFilterMenu && !activeColumnFilterMenu.contains(event.target)) closeColumnFilterMenu();
}

function closeColumnFilterMenuOnEscape(event) {
  if (event.key === "Escape") closeColumnFilterMenu();
}

function closeColumnFilterMenu() {
  if (activeColumnFilterMenu) activeColumnFilterMenu.remove();
  activeColumnFilterMenu = null;
  document.removeEventListener("click", closeColumnFilterMenuOnOutside);
  document.removeEventListener("keydown", closeColumnFilterMenuOnEscape);
}

function getColumnValues(table, index) {
  const values = new Set();
  [...table.tBodies[0].rows].forEach((row) => {
    if (row.querySelector(".empty-row")) return;
    values.add(clean(row.cells[index]?.textContent || ""));
  });
  return [...values].sort((a, b) => a.localeCompare(b, "pt-BR", {numeric: true, sensitivity: "base"}));
}

function setColumnFilter(table, index, values) {
  let filters = tableColumnFilters.get(table);
  if (!filters) {
    filters = new Map();
    tableColumnFilters.set(table, filters);
  }
  if (!values) filters.delete(index);
  else filters.set(index, new Set(values));
  updateColumnFilterState(table);
  applyColumnFilters(table);
}

function applyColumnFilters(table) {
  if (!table.tBodies.length) return;
  const activeFilters = tableColumnFilters.get(table) || new Map();
  [...table.tBodies[0].rows].forEach((row) => {
    if (row.querySelector(".empty-row")) return;
    const visible = [...activeFilters.entries()].every(([index, values]) => values.has(clean(row.cells[index]?.textContent || "")));
    row.classList.toggle("column-filter-hidden", !visible);
  });
  updateColumnFilterState(table);
}

function updateColumnFilterState(table) {
  const filters = tableColumnFilters.get(table) || new Map();
  table.querySelectorAll(".column-filter-header").forEach((cell) => {
    const index = Number(cell.dataset.columnIndex);
    cell.classList.toggle("column-filter-active", filters.has(index));
  });
}

function sortTableByColumn(table, index, headerCell) {
  const tbody = table.tBodies && table.tBodies[0];
  if (!tbody) return;
  const currentIndex = Number(table.dataset.sortIndex);
  const currentOrder = table.dataset.sortOrder || "desc";
  const order = currentIndex === index && currentOrder === "asc" ? "desc" : "asc";
  const rows = [...tbody.rows].filter((row) => !row.querySelector(".empty-row"));
  rows.sort((rowA, rowB) => compareTableCellValues(rowA.cells[index]?.textContent, rowB.cells[index]?.textContent, order));
  rows.forEach((row) => tbody.appendChild(row));
  table.dataset.sortIndex = String(index);
  table.dataset.sortOrder = order;
  table.querySelectorAll(".column-filter-header").forEach((cell) => {
    cell.classList.remove("column-sort-asc", "column-sort-desc");
  });
  headerCell.classList.add(order === "asc" ? "column-sort-asc" : "column-sort-desc");
}

function compareTableCellValues(valueA, valueB, order) {
  const textA = clean(valueA);
  const textB = clean(valueB);
  const numA = parseTableNumber(textA);
  const numB = parseTableNumber(textB);
  const result = Number.isFinite(numA) && Number.isFinite(numB)
    ? numA - numB
    : textA.localeCompare(textB, "pt-BR", {numeric: true, sensitivity: "base"});
  return order === "asc" ? result : -result;
}

function parseTableNumber(value) {
  const normalized = clean(value)
    .replace(/[^\d,.-]/g, "")
    .replace(/\.(?=\d{3}(?:\D|$))/g, "")
    .replace(",", ".");
  if (!normalized || /^[-.]?$/.test(normalized)) return NaN;
  return Number(normalized);
}

function clean(value) {
  return String(value ?? "").trim();
}

function formatDate(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString("pt-BR");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function notify(message, isError = false) {
  const toast = el("toast");
  toast.textContent = message;
  toast.style.background = isError ? "#7a271a" : "#111827";
  toast.classList.remove("hidden");
  window.clearTimeout(notify.timer);
  notify.timer = window.setTimeout(() => toast.classList.add("hidden"), 3200);
}
