# Módulos do Sistema TMS - Mapeamento Detalhado

## 1. Dashboard/Home (HomePage)
**Localização**: main.py (linhas ~1-2000)
**Funcionalidades**:
- KPIs operacionais (rotas ativas, motoristas, etc.)
- Gráfico de distribuição de carga
- Status do sistema
- Acesso rápido a módulos

**Regras de Negócio**:
- Controle de permissões por plano
- Cálculo de métricas em tempo real

## 2. Cadastros (CadastrosPage)
**Localização**: main.py (linhas ~3000-5000)
**Entidades**:
- Usuários (nome, senha, permissões, CPF, idade, telefone)
- Motoristas (código único, nome, CNH, CPF, idade, telefone, senha)
- Veículos (placa única, modelo, ano, cor)
- Equipes (código único, ajudante1, ajudante2)
- Vendedores (nome, telefone, rota)

**Regras de Negócio**:
- Códigos únicos para motoristas e equipes
- Validação de CPF/CNH
- Controle de acesso por usuário

## 3. Rotas (RotasPage)
**Localização**: main.py (linhas ~2400-3000)
**Funcionalidades**:
- Cadastro de rotas
- Associação com motoristas/equipes
- Controle de status

**Regras de Negócio**:
- Rota obrigatória para entregas
- Validação de disponibilidade de equipe

## 4. Programação de Entrega (ProgramacaoPage)
**Localização**: main.py (linhas ~5700-10000)
**Funcionalidades**:
- Agendamento de entregas
- Itens de programação (datas, horários, equipes)
- Controle de status (ATIVA, EM_ROTA, FINALIZADA, CANCELADA)
- Cálculo de horas trabalhadas

**Regras de Negócio**:
- Data/hora de saída e chegada obrigatórias
- Cálculo automático de horas (máximo 72h)
- Validação de equipe completa
- Status transitions controladas

## 5. Recebimentos (RecebimentosPage)
**Localização**: main.py (linhas ~11800-13000)
**Funcionalidades**:
- Lançamentos de recebimentos por rota
- Categorização (tipo, valor, data)
- Fechamentos mensais

**Regras de Negócio**:
- Recebimento vinculado a rota
- Validação de valores positivos
- Controle de fechamento (impede alterações)

## 6. Despesas (DespesasPage)
**Localização**: main.py (linhas ~14600-18000)
**Funcionalidades**:
- Lançamentos de despesas
- Categorias (combustível, manutenção, etc.)
- Fechamentos por rota

**Regras de Negócio**:
- Despesa vinculada a rota ou geral
- Cálculo de custos por KM
- Validação de categorias

## 7. Prestação de Contas
**Localização**: Integrado em Recebimentos/Despesas
**Funcionalidades**:
- Relatórios de custos por rota
- Cálculo de lucro/prejuízo
- Mortalidade de aves
- Performance por motorista

**Regras de Negócio**:
- Cálculo de mortalidade (aves perdidas/total)
- KM/Litro automático
- Indicadores de eficiência

## 8. Escala (EscalaPage)
**Localização**: main.py (linhas ~20500-22200)
**Funcionalidades**:
- Distribuição de carga por motorista/ajudante
- KPIs: rotas, horas, KM, mortalidade
- Gráfico de barras (top 8 motoristas)
- Recomendações automáticas
- Filtros por período/status

**Regras de Negócio**:
- Cálculo de média por motorista
- Indicadores de sobrecarga (>1.25x média)
- Recomendações baseadas em localização
- Tags visuais (equilibrada, alerta, sobrecarga)

## 9. Centro de Custos (CentroCustosPage)
**Localização**: app/ui/pages/centro_custos_page.py
**Funcionalidades**:
- Controle orçamentário
- Categorização de custos
- Relatórios por centro

**Regras de Negócio**:
- Vinculação de despesas a centros
- Controle de orçamento vs. realizado

## 10. Relatórios (RelatoriosPage)
**Localização**: app/ui/pages/relatorios_page.py
**Funcionalidades**:
- Geração de PDF com ReportLab
- Exportação Excel com Pandas
- Relatórios operacionais e financeiros

**Regras de Negócio**:
- Filtros por período e status
- Formatação específica para PDFs

## 11. Permissões (PermissionsPage)
**Localização**: app/ui/pages/permissions_page.py
**Funcionalidades**:
- Controle de acesso por rotina
- Perfis de usuário

**Regras de Negócio**:
- RBAC (Role-Based Access Control)
- Validação de permissões no backend

## 12. Backup/Exportação (BackupExportarPage)
**Localização**: app/ui/pages/backup_exportar_page.py
**Funcionalidades**:
- Exportação de dados
- Backup do banco
- Restauração

**Regras de Negócio**:
- Formatos de exportação padronizados

## 13. Configurações
**Localização**: Disperso em main.py
**Funcionalidades**:
- Configurações do sistema
- Parâmetros operacionais

## Regras de Negócio Transversais

### Autenticação e Autorização
- Usuários com hash PBKDF2
- Controle por empresa (tenant)
- Permissões por rotina

### Validações Comuns
- Datas no formato brasileiro
- Valores monetários com vírgula
- Códigos únicos
- Campos obrigatórios

### Cálculos Automáticos
- Horas trabalhadas = chegada - saída
- Mortalidade = perdidas/total * 100
- KM médio = total_km / motoristas
- Custo por KM = despesas / km_rodado

### Status de Rotas
- ATIVA: Programada
- EM_ROTA: Saiu para entrega
- CARREGADA: Carregamento concluído
- FINALIZADA: Entrega completa
- CANCELADA: Cancelada

### Sincronização API
- Desktop pode ler da API quando configurado
- Prioridade: API > Local
- Tratamento de conflitos

## Dependências entre Módulos
- Cadastros → Todos os outros
- Programação → Rotas, Recebimentos, Despesas, Escala
- Recebimentos/Despesas → Prestação de Contas
- Programação + Recebimentos/Despesas → Relatórios

## Pontos de Atenção na Migração
- Preservar todos os cálculos automáticos
- Manter validações de negócio
- Garantir compatibilidade de dados
- Não quebrar fluxos de status
- Manter interface familiar para usuários