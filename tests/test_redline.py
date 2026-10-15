import io
import unittest
import zipfile
import xml.etree.ElementTree as ET

from docgen import SAMPLE, TEMPLATE, docx, p
from docxcore import W, Package, iter_paragraphs
from numbering import number_paragraphs
import edits
import protocol_app
import redline


def check_structure(test, data):
    """Базовая проверка корректности правок в document.xml."""
    xml = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")
    root = ET.fromstring(xml)
    parents = {c: p for p in root.iter() for c in p}
    ids = []
    for el in root.iter():
        if el.tag in (W + "ins", W + "del"):
            ids.append(el.get(W + "id"))
            par = parents[el]
            test.assertNotEqual(par.tag, el.tag, "вложенные w:ins/w:ins или w:del/w:del")
            if el.tag == W + "del":
                test.assertIsNone(el.find(".//" + W + "t"), "w:t внутри w:del")
            test.assertTrue(el.get(W + "author"))
        if el.tag == W + "delText":
            chain, cur = [], el
            while cur in parents:
                cur = parents[cur]
                chain.append(cur.tag)
            test.assertTrue(W + "del" in chain or W + "moveFrom" in chain, "w:delText вне w:del")
    test.assertEqual(len(ids), len(set(ids)), "w:id правок не уникальны")
    test.assertIn(b"mc:Ignorable", xml[:3000])


def accepted_numbers(data):
    pkg = Package(data)
    paras = iter_paragraphs(pkg.body)
    nums = number_paragraphs(pkg, paras, "new")
    from docxcore import exists_in, para_text
    return [(n["text"] if n else "", para_text(q).strip()) for q, n in zip(paras, nums)
            if exists_in(q, "new") and para_text(q).strip()]


AUTO = docx([
    p("Предмет договора", 1, 0), p("Поставщик поставляет товар.", 1, 1),
    p("Цена", 1, 0), p("Цена указана в спецификации.", 1, 1),
    p("Порядок", 1, 0), p("Первое условие.", 1, 1), p("Второе условие.", 1, 1),
    p("Третье условие.", 1, 1),
])


class AutoNumberedTest(unittest.TestCase):
    def test_insert_and_delete(self):
        out, rep = redline.build_final(AUTO, [{"label": "ДС № 1", "edits": [
            {"id": "a", "action": "delete_clause", "target": "2.1", "old_text": ""},
            {"id": "b", "action": "insert", "target": "3.3", "new_text": "3.3. Новое условие о сроках."},
            {"id": "c", "action": "replace_clause", "target": "3.1", "new_text": "Первое условие с уточнением."},
        ]}])
        check_structure(self, out)
        statuses = [r["status"] for r in rep["sources"][0]["results"]]
        self.assertEqual(statuses, ["exact", "exact", "exact"])
        self.assertEqual(accepted_numbers(out), [
            ("1.", "Предмет договора"), ("1.1.", "Поставщик поставляет товар."),
            ("2.", "Цена"),
            ("3.", "Порядок"), ("3.1.", "Первое условие с уточнением."), ("3.2.", "Второе условие."),
            ("3.3.", "Новое условие о сроках."), ("3.4.", "Третье условие."),
        ])
        # отклонение правок возвращает исходник
        self.assertEqual(redline.view_paragraphs(out, "old"), redline.view_paragraphs(AUTO))

    def test_insert_first_subclause(self):
        out, _ = redline.build_final(AUTO, [{"label": "ДС", "edits": [
            {"action": "insert", "target": "2.1", "new_text": "Новый первый подпункт."}]}])
        nums = accepted_numbers(out)
        self.assertIn(("2.1.", "Новый первый подпункт."), nums)
        self.assertIn(("2.2.", "Цена указана в спецификации."), nums)


class SequentialSourcesTest(unittest.TestCase):
    def test_two_sources_two_authors(self):
        out, rep = redline.build_final(AUTO, [
            {"label": "ДС № 1 от 01.02.2026", "edits": [
                {"action": "insert", "target": "1.2", "new_text": "Товар поставляется партиями в течение 10 дней."}]},
            {"label": "ДС № 2 от 01.03.2026", "edits": [
                # ДС № 2 ссылается на пункт, появившийся в ДС № 1, и меняет его текст
                {"action": "replace_words", "target": "1.2", "old_text": "10 дней", "new_text": "15 рабочих дней"},
                {"action": "replace_words", "target": "", "old_text": "спецификации", "new_text": "Приложении № 1"}]},
        ])
        check_structure(self, out)
        self.assertTrue(all(r["status"] == "exact" for s in rep["sources"] for r in s["results"]))
        texts = redline.view_paragraphs(out)
        self.assertIn("Товар поставляется партиями в течение 15 рабочих дней.", texts)
        self.assertIn("Цена указана в Приложении № 1.", texts)
        xml = zipfile.ZipFile(io.BytesIO(out)).read("word/document.xml").decode()
        self.assertIn('w:author="ДС № 1 от 01.02.2026"', xml)
        self.assertIn('w:author="ДС № 2 от 01.03.2026"', xml)
        # удаление текста, вставленного ДС № 1, — w:del внутри w:ins
        root = ET.fromstring(xml)
        nested = [d for i in root.iter(W + "ins") for d in i.findall(W + "del")]
        self.assertTrue(nested)

    def test_not_found(self):
        _, rep = redline.build_final(AUTO, [{"label": "ДС", "edits": [
            {"action": "replace_clause", "target": "9.9", "old_text": "Такого текста нет в договоре.",
             "new_text": "x"}]}])
        self.assertEqual(rep["sources"][0]["results"][0]["status"], "not_found")


class TextNumberingTest(unittest.TestCase):
    DOC = docx([p("1. Предмет"), p("1.1. Поставщик поставляет товар."), p("1.2. Покупатель оплачивает товар."),
                p("2. Срок"), p("2.1. Договор действует год.")], numbering_xml=None)

    def test_insert_keeps_text_number_and_warns_duplicates(self):
        out, rep = redline.build_final(self.DOC, [{"label": "ДС", "edits": [
            {"action": "insert", "target": "1.2", "new_text": "Товар поставляется партиями."}]}])
        texts = redline.view_paragraphs(out)
        self.assertEqual(texts[2], "1.2. Товар поставляется партиями.")
        self.assertTrue(any("1.2" in w for w in rep["warnings"]))

    def test_insert_without_previous_sibling(self):
        # п. 1.4 вставляется, хотя п. 1.3 нет: после ближайшего 1.2, не «куда попало»
        out, rep = redline.build_final(self.DOC, [{"label": "ДС", "edits": [
            {"action": "replace_clause", "target": "2.1", "new_text": "Договор действует два года."},
            {"action": "insert", "target": "1.4", "new_text": "Риск переходит при передаче."}]}])
        texts = redline.view_paragraphs(out)
        self.assertEqual(texts.index("1.4. Риск переходит при передаче."), 3)
        self.assertEqual(rep["sources"][0]["results"][1]["status"], "by_text")

    def test_number_is_preserved_on_replace(self):
        out, _ = redline.build_final(self.DOC, [{"label": "ДС", "edits": [
            {"action": "replace_clause", "target": "2.1", "new_text": "Договор действует два года."}]}])
        self.assertIn("2.1. Договор действует два года.", redline.view_paragraphs(out))


@unittest.skipUnless(SAMPLE.exists() and TEMPLATE.exists(), "нет образца договора")
class RoundTripTest(unittest.TestCase):
    """Договор с правками -> протокол (вкладка 1) -> применение к исходнику (вкладка 2)."""

    def test_round_trip(self):
        sample = SAMPLE.read_bytes()
        base = redline.reject_all(sample)
        items = protocol_app.parse_contract(sample)
        prot = protocol_app.build_protocol(TEMPLATE.read_bytes(), {}, items)
        parsed = edits.parse_protocol(prot)
        self.assertEqual(len(parsed["edits"]), len(items))
        out, rep = redline.build_final(base, [{"label": "Протокол разногласий", "edits": parsed["edits"]}])
        check_structure(self, out)
        bad = [r for r in rep["sources"][0]["results"] if r["status"] not in ("exact", "by_text")]
        self.assertEqual(bad, [])
        self.assertEqual(redline.view_paragraphs(out, "new"), redline.view_paragraphs(sample, "new"))
        self.assertEqual(redline.view_paragraphs(out, "old"), redline.view_paragraphs(base, "new"))


if __name__ == "__main__":
    unittest.main()
