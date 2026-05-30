# Matriz inicial de planos SaaS

Esta matriz organiza a oferta comercial sem alterar o fluxo operacional atual.
Os codigos dos planos existentes foram preservados para reduzir risco em testes,
assinaturas e trocas de plano.

| Plano | Codigo | Veiculos | Usuarios | Objetivo |
| --- | --- | ---: | ---: | --- |
| Inicial 5 Veiculos | `starter` | 5 | 6 | Para comecar organizado: programacao, recebimentos, despesas e app motorista. |
| Crescimento 10 Veiculos | `growth` | 10 | 15 | Para controlar perdas, rotas, financeiro basico e centro de custos. |
| Profissional 15 Veiculos | `professional` | 15 | 30 | Para operacoes mais exigentes: escala, relatorios, rotas e controles avancados. |
| Empresarial Mais Veiculos | `enterprise` | sob contrato | sob contrato | Para frotas acima de 15 veiculos, API e suporte prioritario. |
| Corporativo Privado 50 | `corporate_private` | 50 | sem limite definido | Plano privado da implantacao corporativa com acesso total. |
| Internal | `internal` | sem limite | sem limite | Desenvolvimento e testes. |

## Etapas recomendadas

1. Seed dos planos e limites de veiculos.
2. Exibir a matriz no Admin SaaS com leitura clara de recursos por plano.
3. Mapear endpoints/telas para `features` sem bloquear usuarios atuais.
4. Ativar bloqueios por plano gradualmente, com mensagens de upgrade.
5. Revisar precos comerciais depois de medir uso real por modulo.
