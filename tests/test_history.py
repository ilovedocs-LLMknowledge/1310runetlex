import json
import tempfile
import unittest
from pathlib import Path

import history


class HistoryTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name) / "Архив"

    def tearDown(self):
        self.tmp.cleanup()

    def test_save_list_resolve(self):
        m1 = history.save_run("protocol", "Договор поставки", [("Договор поставки.docx", b"1")],
                              [("Протокол.docx", b"2")], root=self.root, extra={"items": 3})
        m2 = history.save_run("final", "Договор/аренды", [("Договор.docx", b"1"), ("ДС 1.pdf", b"3")],
                              [("Договор — итоговая редакция.docx", b"4")], report_md="# отчёт", root=self.root,
                              extra={"sources": [{"label": "ДС № 1", "total": 2, "applied": 2}]})
        self.assertNotIn("/", m2["id"])
        runs = history.list_runs(root=self.root)
        self.assertEqual({r["id"] for r in runs}, {m1["id"], m2["id"]})
        self.assertEqual([r["id"] for r in history.list_runs("дс № 1", root=self.root)], [m2["id"]])
        self.assertEqual([r["id"] for r in history.list_runs(kind="protocol", root=self.root)], [m1["id"]])
        out = history.resolve(m2["id"], m2["outputs"][0], root=self.root)
        self.assertEqual(out.read_bytes(), b"4")
        meta = json.loads((self.root / m2["id"] / "meta.json").read_text(encoding="utf8"))
        self.assertEqual(meta["report"], "отчёт.md")

    def test_resolve_rejects_escape(self):
        m = history.save_run("protocol", "Д", [], [("a.docx", b"1")], root=self.root)
        for bad in ("..", "../..", "/etc", ""):
            with self.assertRaises(ValueError):
                history.resolve(bad, root=self.root)
        with self.assertRaises(ValueError):
            history.resolve(m["id"], "../../meta.json", root=self.root)


if __name__ == "__main__":
    unittest.main()


class SettingsTest(unittest.TestCase):
    def test_key_never_exposed(self):
        import config
        from unittest import mock
        with mock.patch("config.gemini_key", return_value="AIzaSyTEST-secret-key-1234567890"):
            pub = config.public_settings()
        self.assertTrue(pub["gemini_key_set"])
        self.assertNotIn("7890", json.dumps(pub))

    def test_write_env_keeps_other_lines(self):
        import config
        path = Path(tempfile.mkdtemp()) / ".env"
        path.write_text("# комментарий\nOTHER=1\nGEMINI_MODEL=old\n", encoding="utf8")
        config.write_env({"GEMINI_MODEL": "new", "GEMINI_API_KEY": "k"}, path)
        self.assertEqual(config.read_env(path), {"OTHER": "1", "GEMINI_MODEL": "new", "GEMINI_API_KEY": "k"})
        self.assertEqual(path.stat().st_mode & 0o777, 0o600)
