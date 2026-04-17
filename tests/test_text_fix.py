import unittest

from app.utils.text_fix import normalize_ui_collection, normalize_ui_text


class TextFixTests(unittest.TestCase):
    def test_normalize_ui_text_repairs_mojibake(self):
        self.assertEqual(normalize_ui_text("VersÃ£o local"), "Versão local")

    def test_normalize_ui_text_applies_ptbr_accents(self):
        self.assertEqual(normalize_ui_text("Rotina Programacao"), "Rotina Programação")

    def test_normalize_ui_text_repairs_message_text(self):
        self.assertEqual(
            normalize_ui_text("ATENCAO: Nao e possivel salvar a programacao."),
            "ATENÇÃO: Não é possível salvar a programação.",
        )

    def test_normalize_ui_collection_preserves_sequence_type(self):
        self.assertEqual(
            normalize_ui_collection(("Prestacao de Contas", "SERTAO")),
            ("Prestação de Contas", "SERTÃO"),
        )

    def test_normalize_ui_text_repairs_operational_terms(self):
        self.assertEqual(
            normalize_ui_text("Preco de saida e lancamento"),
            "Preço de saída e lançamento",
        )


if __name__ == "__main__":
    unittest.main()
