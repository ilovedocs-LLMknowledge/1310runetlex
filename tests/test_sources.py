import os
import unittest
from unittest import mock

from docgen import docx, p
import edits
import gemini
import redline
import sources

CONTRACT = docx([p("1. Предмет"), p("1.1. Поставщик передаёт Товар Покупателю."),
                 p("2. Цена"), p("2.1. Цена Товара составляет 100 000 (сто тысяч) рублей, включая НДС."),
                 p("3. Поставка"), p("3.1. Товар поставляется партиями."),
                 p("5. Споры"), p("5.3. Стороны решают споры путём переговоров."),
                 p("5.4. Споры рассматриваются в Арбитражном суде г. Москвы.")], numbering_xml=None)

# ответ Gemini на скан протокола (записан с реального ответа на синтетический протокол)
SCAN_RESPONSE = {
    "source": {"kind": "protocol", "number": "", "date": "", "title": "Протокол разногласий"},
    "ocr_text": "ПРОТОКОЛ РАЗНОГЛАСИЙ …",
    "edits": [
        {"action": "replace_clause", "target": "2.1", "anchor": "", "note": "",
         "old_text": "Цена Товара составляет 100 000 (сто тысяч) рублей, включая НДС.",
         "new_text": "Цена Товара составляет 120 000 (сто двадцать тысяч) рублей, включая НДС."},
        {"action": "insert", "target": "3.2", "anchor": "", "old_text": "", "note": "",
         "new_text": "Поставщик обязан передать Товар не позднее 10 (десяти) рабочих дней с даты оплаты."},
        {"action": "delete_clause", "target": "5.4", "anchor": "", "new_text": "", "note": "",
         "old_text": "Споры рассматриваются в Арбитражном суде г. Москвы."},
    ],
}

DS_TEXT = """ДОПОЛНИТЕЛЬНОЕ СОГЛАШЕНИЕ № 2
к Договору поставки № 15-П от 01.09.2026
г. Москва «15» марта 2026 г.
1. Пункт 5.3 Договора изложить в следующей редакции:
«5.3. Стороны решают споры путём переговоров в течение 30 дней.
Претензионный порядок обязателен.»
2. Дополнить Договор пунктом 3.2 следующего содержания: «3.2. Товар маркируется.»
3. Пункт 5.4 Договора исключить.
4. В пункте 1.1 слова «передаёт Товар» заменить словами «передаёт в собственность Товар».
5. Пункты 9.1–9.3 считать пунктами 9.2–9.4 соответственно.
6. Остальные условия Договора остаются неизменными.
7. Настоящее Соглашение вступает в силу с момента подписания."""


class ScanTest(unittest.TestCase):
    def test_scan_edits_apply(self):
        with mock.patch.object(gemini, "call", return_value=SCAN_RESPONSE):
            res = sources.parse_source(b"%PDF-fake", "скан.pdf")
        self.assertEqual(res["via"], "gemini")
        self.assertEqual([e["action"] for e in res["edits"]], ["replace_clause", "insert", "delete_clause"])
        out, rep = redline.build_final(CONTRACT, [{"label": res["meta"]["label"], "edits": res["edits"]}])
        self.assertEqual([r["status"] for r in rep["sources"][0]["results"]], ["exact", "exact", "exact"])
        texts = redline.view_paragraphs(out)
        self.assertIn("2.1. Цена Товара составляет 120 000 (сто двадцать тысяч) рублей, включая НДС.", texts)
        self.assertIn("3.2. Поставщик обязан передать Товар не позднее 10 (десяти) рабочих дней с даты оплаты.", texts)
        self.assertNotIn("5.4. Споры рассматриваются в Арбитражном суде г. Москвы.", texts)

    def test_no_key(self):
        with mock.patch("config.gemini_key", return_value=""):
            with self.assertRaises(gemini.GeminiError):
                gemini.call([{"text": "x"}])


class AmendmentRegexTest(unittest.TestCase):
    def test_parse_and_apply(self):
        found, unparsed = edits.parse_amendment_text(DS_TEXT)
        self.assertEqual(unparsed, [])
        self.assertEqual([(e["action"], e["target"]) for e in found], [
            ("replace_clause", "5.3"), ("insert", "3.2"), ("delete_clause", "5.4"), ("replace_words", "1.1"),
            ("renumber", "9.1"), ("renumber", "9.2"), ("renumber", "9.3")])
        meta = edits.amendment_meta(DS_TEXT)
        self.assertEqual(meta["label"], "ДС № 2 от 15.03.2026")
        out, rep = redline.build_final(CONTRACT, [{"label": meta["label"], "edits": found[:4]}])
        texts = redline.view_paragraphs(out)
        self.assertIn("1.1. Поставщик передаёт в собственность Товар Покупателю.", texts)
        self.assertIn("5.3. Стороны решают споры путём переговоров в течение 30 дней.", texts)
        self.assertIn("Претензионный порядок обязателен.", texts)
        self.assertIn("3.2. Товар маркируется.", texts)

    def test_docx_without_key_uses_regex(self):
        ds = docx([p(line) for line in DS_TEXT.split("\n")], numbering_xml=None)
        with mock.patch("config.gemini_key", return_value=""):
            res = sources.parse_source(ds, "ДС 2.docx")
        self.assertEqual(res["via"], "regex")
        self.assertEqual(len(res["edits"]), 7)


@unittest.skipUnless(os.environ.get("RUN_GEMINI") == "1", "живой запрос к Gemini: RUN_GEMINI=1")
class LiveGeminiTest(unittest.TestCase):
    def test_amendment_text(self):
        res = sources._norm_gemini(gemini.extract_from_text(DS_TEXT), "s1")
        actions = [e["action"] for e in res[1]]
        self.assertIn("replace_clause", actions)
        self.assertIn("delete_clause", actions)
        self.assertEqual(res[0]["number"], "2")


if __name__ == "__main__":
    unittest.main()
