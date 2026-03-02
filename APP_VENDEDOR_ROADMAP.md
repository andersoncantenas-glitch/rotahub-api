# App Vendedor (Flutter separado) - Inicio Rapido

Objetivo: criar programacao avulsa no domingo sem afetar o app do motorista.

## 1) Projeto separado (nao compromete o app atual)

Crie outro app Flutter:

```powershell
cd C:\
flutter create flutter_vendedor_app
```

Use backend existente (`api_server.py`) com endpoints `desktop/*` protegidos por `X-Desktop-Secret`.

## 2) Endpoints ja prontos no backend

- `GET /desktop/cadastros/motoristas`
- `GET /desktop/cadastros/veiculos`
- `GET /desktop/cadastros/ajudantes`
- `GET /desktop/clientes/base?q=...`
- `POST /desktop/avulsas`
- `GET /desktop/avulsas`
- `GET /desktop/avulsas/{codigo_avulsa}`
- `POST /desktop/avulsas/{codigo_avulsa}/conciliar`

Header obrigatorio:

- `X-Desktop-Secret: <ROTA_SECRET>`

## 3) Telas MVP (fase 1)

1. Login/Config (salvar API URL + secret)
2. Nova Programacao Avulsa
   - motorista, veiculo, equipe, local, observacao
   - adicionar clientes (busca em `/desktop/clientes/base`)
3. Lista de Avulsas
4. Detalhe da Avulsa

## 4) Payload de criacao (exemplo)

```json
{
  "data_programada": "2026-03-02",
  "motorista_codigo": "MT001",
  "motorista_nome": "JUNIOR",
  "veiculo": "NNC0B42",
  "equipe": "MARCIO / PAULO",
  "local_rota": "SERTAO",
  "observacao": "DOMINGO AVULSA",
  "criado_por": "VENDEDOR_JOAO",
  "itens": [
    {
      "cod_cliente": "123",
      "nome_cliente": "CLIENTE A",
      "cidade": "IPU",
      "bairro": "CENTRO",
      "ordem": 1
    }
  ]
}
```

## 5) Conciliacao na segunda (fase 2)

No Desktop, apos importar vendas, conciliar avulsa com programacao oficial:

`POST /desktop/avulsas/{codigo_avulsa}/conciliar`

Payload:

```json
{
  "codigo_programacao_oficial": "PG202615",
  "usuario": "ADMIN"
}
```

> Nesta fase inicial, a conciliacao marca vinculo de cabecalho.
> A amarracao automatica pedido/NF por cliente entra na fase 2.

