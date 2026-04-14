# Testes de Contrato de Services

Este diretório concentra os testes de contrato dos services de domínio.

## Estrutura

- `test_auth_service_contract.py`
- `test_cliente_service_contract.py`
- `test_motorista_service_contract.py`
- `test_vendedor_service_contract.py`
- `test_programacao_service_contract.py`
- `test_runtime_flags.py`
- `_contract_test_helpers.py` (helpers compartilhados)

## Executar todos (discover)

```powershell
python -m unittest discover -s tests -p "test_*_service_contract.py" -v
```

## Executar toda a suíte deste diretório

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

## Executar um domínio específico

```powershell
python -m unittest tests.test_auth_service_contract -v
python -m unittest tests.test_cliente_service_contract -v
python -m unittest tests.test_motorista_service_contract -v
python -m unittest tests.test_vendedor_service_contract -v
python -m unittest tests.test_programacao_service_contract -v
```

## Validação de sintaxe dos testes

```powershell
python -m py_compile tests\_contract_test_helpers.py tests\test_auth_service_contract.py tests\test_cliente_service_contract.py tests\test_motorista_service_contract.py tests\test_vendedor_service_contract.py tests\test_programacao_service_contract.py
```

```powershell
python -m py_compile tests\_contract_test_helpers.py tests\test_auth_service_contract.py tests\test_cliente_service_contract.py tests\test_motorista_service_contract.py tests\test_vendedor_service_contract.py tests\test_programacao_service_contract.py tests\test_runtime_flags.py
```

## Testes de runtime flags

```powershell
python -m unittest tests.test_runtime_flags -v
```

## Escopo desses testes

- Validam o contrato de retorno: `ok`, `data`, `error`, `source`.
- Cobrem cenários de sucesso e erro controlado.
- Usam mocks para evitar dependência de banco/API reais.
