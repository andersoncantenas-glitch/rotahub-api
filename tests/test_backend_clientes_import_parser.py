import unittest

import pandas as pd

from backend.api.v1.endpoints.cadastros import clientes_rows_from_dataframe


class BackendClientesImportParserTests(unittest.TestCase):
    def test_accepts_accented_wibi_like_headers(self):
        df = pd.DataFrame(
            [["001", "Cliente Um", "Rua A", "88999990000", "Vendedor Um"]],
            columns=["CÓDIGO CLIENTE", "RAZÃO SOCIAL", "ENDEREÇO", "TEL", "REPRESENTANTE"],
        )

        rows = clientes_rows_from_dataframe(df)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].cod_cliente, "001")
        self.assertEqual(rows[0].nome_cliente, "Cliente Um")
        self.assertEqual(rows[0].endereco, "Rua A")
        self.assertEqual(rows[0].telefone, "88999990000")
        self.assertEqual(rows[0].vendedor, "Vendedor Um")

    def test_falls_back_to_first_two_columns_when_headers_are_unknown(self):
        df = pd.DataFrame([["002", "Cliente Dois"]], columns=["COLUNA A", "COLUNA B"])

        rows = clientes_rows_from_dataframe(df)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].cod_cliente, "002")
        self.assertEqual(rows[0].nome_cliente, "Cliente Dois")

    def test_accepts_multiple_phone_numbers_in_one_cell(self):
        phones = "(88) 98896-8871 / (88) 9...5399 / (88) 99988-0332"
        df = pd.DataFrame(
            [["003", "Cliente Tres", phones]],
            columns=["COD CLIENTE", "NOME CLIENTE", "TELEFONE"],
        )

        rows = clientes_rows_from_dataframe(df)

        self.assertGreater(len(phones), 40)
        self.assertEqual(rows[0].telefone, phones)


if __name__ == "__main__":
    unittest.main()
