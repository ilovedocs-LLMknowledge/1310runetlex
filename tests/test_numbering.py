import unittest

from docgen import docx, lvl, numbering, p, styles
from docxcore import Package, iter_paragraphs
from numbering import fmt_number, number_paragraphs


def keys(data, view="new"):
    pkg = Package(data)
    paras = iter_paragraphs(pkg.body)
    return [n["text"] if n else None for n in number_paragraphs(pkg, paras, view)]


class FormatTest(unittest.TestCase):
    def test_formats(self):
        self.assertEqual(fmt_number(4, "russianLower"), "г")
        self.assertEqual(fmt_number(10, "russianLower"), "к")  # без «й»
        self.assertEqual(fmt_number(14, "upperRoman"), "XIV")
        self.assertEqual(fmt_number(28, "lowerLetter"), "bb")
        self.assertEqual(fmt_number(3, "decimalZero"), "03")


class AutoNumberingTest(unittest.TestCase):
    def test_multilevel_and_restart(self):
        d = docx([p("Предмет", 1, 0), p("Текст", 1, 1), p("Ещё", 1, 1), p("Подпункт", 1, 2),
                  p("без номера"), p("Цена", 1, 0), p("Первый", 1, 1), p("буква", 1, 3), p("буква", 1, 3)])
        self.assertEqual(keys(d), ["1.", "1.1.", "1.2.", "1.2.1.", None, "2.", "2.1.", "а)", "б)"])

    def test_style_numbering_and_basedon(self):
        st = styles('<w:style w:type="paragraph" w:styleId="H1"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
                    '<w:style w:type="paragraph" w:styleId="H2"><w:basedOn w:val="H1"/>'
                    '<w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:style>')
        d = docx([p("Раздел", style="H1"), p("Пункт", style="H2"), p("Пункт", style="H2"),
                  p("Раздел", style="H1"), p("Пункт", style="H2")], styles_xml=st)
        self.assertEqual(keys(d), ["1.", "1.1.", "1.2.", "2.", "2.1."])

    def test_lvl_pstyle(self):
        nb = numbering({0: [lvl(0, "%1.", extra='<w:pStyle w:val="H1"/>'),
                            lvl(1, "%1.%2.", extra='<w:pStyle w:val="H2"/>')]}, {5: (0, "")})
        st = styles('<w:style w:type="paragraph" w:styleId="H1"><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style>'
                    '<w:style w:type="paragraph" w:styleId="H2"><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style>')
        d = docx([p("Раздел", style="H1"), p("Пункт", style="H2"), p("Пункт", style="H2")], nb, st)
        self.assertEqual(keys(d), ["1.", "1.1.", "1.2."])

    def test_start_override_and_shared_abstract(self):
        nb = numbering({0: [lvl(0, "%1."), lvl(1, "%1.%2.")]},
                       {1: (0, ""), 2: (0, ""),
                        3: (0, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>')})
        d = docx([p("а", 1, 0), p("б", 1, 0), p("в", 2, 0), p("г", 3, 0), p("д", 3, 1)], nb)
        # numId 2 продолжает общий abstractNum, numId 3 перезапускает
        self.assertEqual(keys(d), ["1.", "2.", "3.", "1.", "1.1."])

    def test_legal_and_numid_zero(self):
        nb = numbering({0: [lvl(0, "%1.", "upperRoman"), lvl(1, "%1.%2.", extra="<w:isLgl/>")]}, {1: (0, "")})
        d = docx([p("Раздел", 1, 0), p("Пункт", 1, 1), p("выкл", 0, 0)], nb)
        self.assertEqual(keys(d), ["I.", "1.1.", None])

    def test_text_numbering_fallback(self):
        d = docx([p("1. Предмет"), p("1.1. Текст"), p("Без номера"), p("Сам", 1, 0)])
        self.assertEqual(keys(d), ["1.", "1.1.", None, "1."])

    def test_views_skip_merged_paragraphs(self):
        ins_mark = ('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
                    '<w:rPr><w:ins w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>'
                    '<w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Новый</w:t></w:r></w:ins></w:p>')
        d = docx([p("а", 1, 0), ins_mark, p("б", 1, 0)])
        self.assertEqual(keys(d, "new"), ["1.", "2.", "3."])
        self.assertEqual(keys(d, "old"), ["1.", None, "2."])


if __name__ == "__main__":
    unittest.main()
