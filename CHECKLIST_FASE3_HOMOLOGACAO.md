# Checklist Fase 3 - Homologacao Ponta a Ponta

Data base: 2026-03-05

## Pre-condicoes
- API online em `https://rotahub-api.onrender.com`
- Desktop com `ROTA_SECRET` configurado
- APK motorista apontando para a mesma API
- Banco local atualizado (`C:\rotahub\rota_granja.db`)

## 1. Cadastros (Desktop)
- [ ] Cadastrar 1 motorista novo
- [ ] Cadastrar 1 veiculo novo
- [ ] Cadastrar 1 ajudante novo
- [ ] Cadastrar 1 cliente novo
- [ ] Validar se todos aparecem na grade com colunas corretas
- [ ] Editar status de um cadastro e confirmar persistencia

## 2. Importar Vendas
- [ ] Importar planilha de teste (minimo 3 linhas validas)
- [ ] Marcar/desmarcar vendas
- [ ] Confirmar que nao duplica linha ja importada
- [ ] Confirmar filtro por busca funcionando

## 3. Programacao
- [ ] Criar programacao nova com motorista, veiculo e ajudante
- [ ] Carregar vendas selecionadas
- [ ] Salvar programacao
- [ ] Editar programacao salva e confirmar alteracoes
- [ ] Imprimir/pre-visualizar romaneio sem erro

## 4. Vinculo com APK Motorista
- [ ] APK recebe programacao criada (mesmo codigo)
- [ ] Alteracao de status da rota no APK reflete no Desktop
- [ ] GPS/monitoramento da rota aparece na tela Rotas
- [ ] Finalizacao da rota no APK volta para Desktop

## 5. Recebimentos e Despesas
- [ ] Programacao finalizada aparece para prestacao
- [ ] Lancar ao menos 1 recebimento
- [ ] Lancar ao menos 1 despesa
- [ ] Salvar prestacao sem erro
- [ ] Validar totais (entrada, saida, resultado)

## 6. Fechamento e Relatorios
- [ ] Finalizar prestacao
- [ ] Gerar relatorio resumido
- [ ] Gerar PDF
- [ ] Conferir tela Escala
- [ ] Conferir tela Centro de Custos
- [ ] Conferir Backup/Exportar

## 7. Validacoes de Integracao (tecnico)
- [ ] `GET /admin/motoristas/acesso` retorna 200
- [ ] `GET /desktop/overview` retorna 200
- [ ] `GET /desktop/monitoramento/rotas` retorna 200
- [ ] `GET /desktop/rotas/{codigo_programacao}` retorna 200 para codigo real

## Critero de Aprovacao
- Aprovado somente se todas as etapas acima estiverem marcadas sem erro bloqueante.
- Se houver falha, registrar: tela, acao, mensagem exibida e codigo da programacao.

