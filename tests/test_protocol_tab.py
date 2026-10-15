import unittest

from docgen import SAMPLE, docx, p
import protocol_app

INS = 'w:id="{}" w:author="Сторона-2" w:date="2026-01-01T00:00:00Z"'


class AutoNumberedProtocolTest(unittest.TestCase):
    def test_labels_from_autonumbering(self):
        new_para = ('<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'
                    f'<w:rPr><w:ins {INS.format(1)}/></w:rPr></w:pPr>'
                    f'<w:ins {INS.format(2)}><w:r><w:t>Новый пункт о штрафе.</w:t></w:r></w:ins></w:p>')
        changed = ('<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>'
                   '<w:r><w:t xml:space="preserve">Срок оплаты </w:t></w:r>'
                   f'<w:del {INS.format(3)}><w:r><w:delText>10</w:delText></w:r></w:del>'
                   f'<w:ins {INS.format(4)}><w:r><w:t>30</w:t></w:r></w:ins>'
                   '<w:r><w:t xml:space="preserve"> дней.</w:t></w:r></w:p>')
        d = docx([p("Предмет", 1, 0), p("Текст пункта один.", 1, 1), new_para,
                  p("Цена", 1, 0), p("Цена составляет 100 рублей.", 1, 1), changed])
        items = protocol_app.parse_contract(d)
        labels = {it["kind"]: it["label"] for it in items}
        self.assertEqual(labels["new"], "1.2")
        self.assertEqual(labels["changed"], "2.2")
        # сдвиг автонумерации сам по себе правкой не считается
        self.assertEqual(len(items), 2)


@unittest.skipUnless(SAMPLE.exists(), "нет образца договора")
class SampleTest(unittest.TestCase):
    def test_sample_parses(self):
        items = protocol_app.parse_contract(SAMPLE.read_bytes())
        self.assertGreater(len(items), 20)


if __name__ == "__main__":
    unittest.main()
