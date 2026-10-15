#!/usr/bin/env python3
"""
Протокол разногласий из правок Word.

Запуск: python3 protocol_app.py — откроется страница в браузере.
Работает на этом компьютере: сервер слушает 127.0.0.1, только стандартная
библиотека Python. В сеть (Gemini) уходят только сканы и допсоглашения во
вкладке «Итоговая редакция» и только если задан ключ (см. gemini.py).
Страница интерфейса — page.html.
"""
import html
import io
import json
import re
import socket
import sys
import threading
import unicodedata
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote

from numbering import Numbering

APP_DIR = Path(__file__).resolve().parent
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ---------------------------------------------------------------------------
# Чтение правок из договора
# ---------------------------------------------------------------------------

SKIP_TAGS = {W + "pPr", W + "rPr", W + "instrText", W + "delInstrText",
             W + "txbxContent", W + "drawing", W + "pict",
             "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"}

# Номер пункта в начале абзаца: «1.», «2.1», «6.16.1.» (обязательно с точкой)
NUM_RE = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,3})*\.?)(?=\s)")


def num_match(text):
    m = NUM_RE.match(text)
    return m if m and "." in m.group(1) else None


def para_segments(p):
    """Абзац -> список (вид, текст), вид: same / ins / del."""
    out = []

    def walk(el, ins, dl):
        tag = el.tag
        if tag in SKIP_TAGS:
            return
        if tag in (W + "ins", W + "moveTo"):
            ins = True
        if tag in (W + "del", W + "moveFrom"):
            dl = True
        text = None
        if tag in (W + "t", W + "delText"):
            text = el.text or ""
        elif tag in (W + "tab", W + "br", W + "cr"):
            text = " "
        elif tag == W + "noBreakHyphen":
            text = "-"
        if text and not (ins and dl):  # вставлено и тут же удалено — нигде нет
            kind = "ins" if ins else "del" if dl else "same"
            if out and out[-1][0] == kind:
                out[-1][1] += text
            else:
                out.append([kind, text])
        for child in el:
            walk(child, ins, dl)

    walk(p, False, False)
    return out


def para_mark(p):
    """Знак абзаца вставлен ('ins') / удалён ('del') / без изменений (None)."""
    rpr = p.find(W + "pPr/" + W + "rPr")
    if rpr is not None:
        if rpr.find(W + "ins") is not None:
            return "ins"
        if rpr.find(W + "del") is not None:
            return "del"
    return None


def version_text(segs, version):
    drop = "ins" if version == "old" else "del"
    return "".join(t for k, t in segs if k != drop)


def split_points(paras, version):
    """
    Делит документ на пункты в одной из редакций (old / new).
    Возвращает: список пунктов {num, logical: [[индексы абзацев], ...]}
    и словарь «индекс абзаца -> номер пункта».
    """
    # 1) логические абзацы: если знак абзаца вставлен, в старой редакции
    #    этот абзац был слит со следующим (и наоборот для удалённого знака)
    merge_mark = "ins" if version == "old" else "del"
    logical, cur = [], []
    for i, pa in enumerate(paras):
        cur.append(i)
        if pa["mark"] != merge_mark:
            logical.append(cur)
            cur = []
    if cur:
        logical.append(cur)

    # 2) пункты: начинаются с абзаца с номером, дальше — абзацы без номера
    texts = ["".join(version_text(paras[i]["segs"], version) for i in lp).strip() for lp in logical]
    nonempty = [k for k, t in enumerate(texts) if t]

    def lp_num(k):
        """Номер логического абзаца: автонумерация (у знака абзаца — последнего
        в группе) или номер, набранный текстом."""
        auto = paras[logical[k][-1]].get("auto", {}).get(version)
        if auto:
            return auto
        m = num_match(texts[k])
        return m.group(1).rstrip(".") if m else ""

    def is_heading(k):
        """Короткий абзац без номера перед пронумерованным пунктом — заголовок."""
        t = texts[k]
        if not t or lp_num(k) or len(t) > 80 or t[-1] in ".;:,":
            return False
        nxt = next((j for j in nonempty if j > k), None)
        return nxt is not None and bool(lp_num(nxt))

    points, owner = [], {}
    for k, lp in enumerate(logical):
        num = lp_num(k)
        heading = is_heading(k)
        if num or heading or not points:
            points.append({"num": num,
                           "heading": texts[k] if heading else "", "logical": []})
        points[-1]["logical"].append(lp)
        for i in lp:
            owner[i] = len(points) - 1
    return points, owner


def norm(s):
    return re.sub(r"\s+", " ", s).strip()


def strip_num(s):
    m = num_match(s)
    return s[m.end():].strip() if m else s.strip()


def auto_numbers(zf, elements):
    """Автонумерация Word для каждого абзаца в обеих редакциях: [{old, new}]."""
    names = zf.namelist()
    nb = Numbering(zf.read("word/numbering.xml") if "word/numbering.xml" in names else None,
                   zf.read("word/styles.xml") if "word/styles.xml" in names else None)
    views = {v: nb.compute(elements, v) for v in ("old", "new")}
    return [{v: (views[v][i] or {}).get("key", "") for v in views} for i in range(len(elements))]


def parse_contract(data):
    zf = zipfile.ZipFile(io.BytesIO(data))
    root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(W + "body")
    elements = list(body.iter(W + "p"))
    paras = [{"segs": para_segments(p), "mark": para_mark(p), "auto": a}
             for p, a in zip(elements, auto_numbers(zf, elements))]

    old_pts, old_owner = split_points(paras, "old")
    new_pts, new_owner = split_points(paras, "new")

    # Связываем старые и новые пункты, у которых есть общий неизменённый текст
    parent = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        parent[find(a)] = find(b)

    for key in [("o", i) for i in range(len(old_pts))] + [("n", j) for j in range(len(new_pts))]:
        find(key)
    for i, pa in enumerate(paras):
        if any(k == "same" and t.strip() for k, t in pa["segs"]):
            union(("o", old_owner[i]), ("n", new_owner[i]))

    groups = {}
    for key in list(parent):
        groups.setdefault(find(key), []).append(key)

    def render(pts, version):
        """Абзацы редакции: список списков сегментов."""
        drop = "ins" if version == "old" else "del"
        result = []
        for pt in pts:
            for lp in pt["logical"]:
                segs = []
                for i in lp:
                    for k, t in paras[i]["segs"]:
                        if k == drop:
                            continue
                        if segs and segs[-1][0] == k:
                            segs[-1][1] += t
                        else:
                            segs.append([k, t])
                if "".join(t for _, t in segs).strip():
                    result.append(segs)
        return result

    def first_para(opts, npts):
        """Первый абзац, где у пункта есть текст (в любой редакции)."""
        idx = [i for pts, v in ((opts, "old"), (npts, "new")) for pt in pts
               for lp in pt["logical"] for i in lp if version_text(paras[i]["segs"], v).strip()]
        return min(idx) if idx else 0

    items = []
    for keys in groups.values():
        opts = [old_pts[i] for s, i in sorted(keys) if s == "o"]
        npts = [new_pts[i] for s, i in sorted(keys) if s == "n"]
        old = render(opts, "old")
        new = render(npts, "new")
        old_txt = "\n".join("".join(t for _, t in p) for p in old)
        new_txt = "\n".join("".join(t for _, t in p) for p in new)
        if norm(old_txt) == norm(new_txt):
            continue  # правок нет (или поменялись только пробелы)

        old_num = ", ".join(pt["num"] for pt in opts if pt["num"] and render([pt], "old"))
        new_num = ", ".join(pt["num"] for pt in npts if pt["num"] and render([pt], "new"))
        if not old:
            kind = "new"
        elif not new:
            kind = "deleted"
        elif norm(strip_num(old_txt)) == norm(strip_num(new_txt)):
            kind = "renum"
        else:
            kind = "changed"

        if kind == "new" or not old_num:
            label = new_num
        elif kind == "deleted" or not new_num or old_num == new_num:
            label = old_num
        else:
            label = f"{old_num} ({new_num} в редакции Стороны-2)"
        heading = next((pt["heading"] for pt in opts + npts if pt["heading"]), "")
        if not label and heading:
            label = f"Заголовок «{heading}»"
        if not label:
            label = "Вводная часть" if first_para(opts, npts) < 30 else "Без номера"

        items.append({
            "order": first_para(opts, npts),
            "label": label,
            "kind": kind,
            "old": old,
            "new": new,
        })

    items.sort(key=lambda x: x["order"])
    return items


# ---------------------------------------------------------------------------
# Сборка протокола по шаблону
# ---------------------------------------------------------------------------

PH_RE = re.compile(r"\{\{\s*(.+?)\s*\}\}")
T_RE = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)


def default_template():
    for folder in (APP_DIR, APP_DIR / "docs"):
        preferred = folder / "шаблон_протокола_разногласий.docx"
        if preferred.exists():
            return preferred
    found = sorted(p for folder in (APP_DIR, APP_DIR / "docs") for p in folder.glob("*.docx")
                   if "шаблон" in unicodedata.normalize("NFC", p.name).lower() and not p.name.startswith("~$"))
    return found[0] if found else None


def template_placeholders(tpl_bytes):
    xml = zipfile.ZipFile(io.BytesIO(tpl_bytes)).read("word/document.xml").decode("utf8")
    names = []
    for p in re.findall(r"<w:p[ >].*?</w:p>|<w:p/>", xml, re.S):
        text = html.unescape("".join(m.group(2) for m in T_RE.finditer(p)))
        for n in PH_RE.findall(text):
            if n not in names:
                names.append(n)
    return names


def esc(s):
    return html.escape(s, quote=False)


def fill_placeholders(xml, fields):
    fields = {k: v for k, v in fields.items() if v.strip()}
    if not fields:
        return xml

    def repl_text(text):
        return PH_RE.sub(lambda m: fields.get(m.group(1).strip(), m.group(0)), text)

    def fix_para(m):
        p = m.group(0)
        texts = [html.unescape(t.group(2)) for t in T_RE.finditer(p)]
        if "{{" not in "".join(texts):
            return p
        # сначала заменяем внутри каждого фрагмента текста
        p = T_RE.sub(lambda t: t.group(1) + esc(repl_text(html.unescape(t.group(2)))) + t.group(3), p)
        joined = "".join(html.unescape(t.group(2)) for t in T_RE.finditer(p))
        if not any("{{" + k in joined or "{{ " + k in joined for k in fields):
            return p
        # поле разбито Word'ом на несколько кусков — собираем текст в первом
        full = repl_text(joined)
        first = [True]

        def put(t):
            if first[0]:
                first[0] = False
                return '<w:t xml:space="preserve">' + esc(full) + "</w:t>"
            return "<w:t></w:t>"
        return T_RE.sub(put, p)

    return re.sub(r"<w:p[ >].*?</w:p>", fix_para, xml, flags=re.S)


def run_xml(text, bold=False):
    b = "<w:b/>" if bold else '<w:b w:val="0"/>'
    return f'<w:r><w:rPr>{b}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def cell_paras(paragraphs, ppr, bold_ins):
    out = []
    for segs in paragraphs:
        runs = "".join(run_xml(t, bold=(bold_ins and k == "ins")) for k, t in segs)
        out.append(f"<w:p>{ppr}{runs}</w:p>")
    return "".join(out) or f"<w:p>{ppr}</w:p>"


def fill_table(xml, items):
    tables = list(re.finditer(r"<w:tbl>.*?</w:tbl>", xml, re.S))

    def tbl_text(t):
        return html.unescape(re.sub(r"<[^>]+>", "", t.group(0)))

    target = next((t for t in tables if "Редакция" in tbl_text(t)), None)
    if target is None:
        target = next((t for t in tables if t.group(0).count("<w:gridCol ") == 3), None)
    if target is None:
        raise ValueError("В шаблоне не найдена таблица для пунктов "
                         "(ожидается таблица с заголовком «Редакция …» из 3 колонок).")

    tbl = target.group(0)
    rows = list(re.finditer(r"<w:tr[ >].*?</w:tr>", tbl, re.S))
    sample = rows[-1].group(0)
    cells = re.findall(r"<w:tc[ >].*?</w:tc>|<w:tc>.*?</w:tc>", sample, re.S)
    if len(cells) < 3:
        raise ValueError("В таблице шаблона меньше 3 колонок.")
    tr_open = re.match(r"<w:tr[^>]*>", sample).group(0)
    tr_pr = (re.search(r"<w:trPr>.*?</w:trPr>", sample, re.S) or [""])[0]
    tc_prs = [(re.search(r"<w:tcPr>.*?</w:tcPr>", c, re.S) or [""])[0] for c in cells[:3]]
    p_prs = [(re.search(r"<w:p[ >]\s*(<w:pPr>.*?</w:pPr>)", c, re.S) or ["", ""])[1] for c in cells[:3]]
    if len(rows) == 1:  # в таблице только шапка — не копируем выравнивание заголовка
        p_prs = ["", "", ""]

    new_rows = []
    for it in items:
        old = it["old"] or [[["same", "Отсутствует"]]]
        new = it["new"] or [[["same", "Исключить"]]]
        contents = [
            cell_paras([[["same", it["label"]]]], p_prs[0], False),
            cell_paras(old, p_prs[1], False),
            cell_paras(new, p_prs[2], True),
        ]
        tcs = "".join(f"<w:tc>{pr}{c}</w:tc>" for pr, c in zip(tc_prs, contents))
        new_rows.append(f"{tr_open}{tr_pr}{tcs}</w:tr>")

    if len(rows) == 1:
        new_tbl = tbl.replace(sample, sample + "".join(new_rows), 1)
    else:
        new_tbl = tbl[:rows[-1].start()] + "".join(new_rows) + tbl[rows[-1].end():]
    return xml[:target.start()] + new_tbl + xml[target.end():]


def build_protocol(tpl_bytes, fields, items):
    src = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    xml = src.read("word/document.xml").decode("utf8")
    xml = fill_placeholders(xml, fields)
    xml = fill_table(xml, items)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = xml.encode("utf8") if info.filename == "word/document.xml" else src.read(info)
            dst.writestr(info, data)
    return out.getvalue()


# ---------------------------------------------------------------------------
# Локальный веб-сервер
# ---------------------------------------------------------------------------

import config  # noqa: E402
import edits as edits_mod  # noqa: E402
import gemini  # noqa: E402
import history  # noqa: E402
import redline  # noqa: E402
import sources as sources_mod  # noqa: E402

PAGE_PATH = APP_DIR / "page.html"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

STATE = {
    "tpl_bytes": None, "tpl_name": None,
    "contract": None, "contract_name": None,          # вкладка 1
    "final": {"contract": None, "name": None, "sources": {}, "seq": 0},  # вкладка 2
}
LOCK = threading.Lock()


def load_default_template():
    path = default_template()
    if path:
        STATE["tpl_bytes"] = path.read_bytes()
        STATE["tpl_name"] = path.name


def template_info():
    if not STATE["tpl_bytes"]:
        return {"name": None, "placeholders": []}
    return {"name": STATE["tpl_name"], "placeholders": template_placeholders(STATE["tpl_bytes"])}


def stem(name):
    return re.sub(r"\.(docx|pdf|jpe?g|png|heic|heif|tiff?|webp)$", "", name or "", flags=re.I)


def final_sources(req):
    """Источники из запроса интерфейса + их файлы с сервера, в порядке применения."""
    stored = STATE["final"]["sources"]
    out = []
    for s in req.get("sources", []):
        if s.get("id") not in stored:
            raise ValueError("Документ-источник не найден — загрузите его заново.")
        out.append({"id": s["id"], "label": (s.get("label") or "").strip() or stored[s["id"]]["meta"]["label"],
                    "edits": s.get("edits", []), "filename": stored[s["id"]]["filename"]})
    return out


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def send(self, code, body, ctype, headers=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, obj, code=200):
        self.send(code, json.dumps(obj, ensure_ascii=False).encode("utf8"), "application/json; charset=utf-8")

    def body(self):
        return self.rfile.read(int(self.headers.get("Content-Length", 0)))

    def filename(self, default="файл.docx"):
        return unicodedata.normalize("NFC", unquote(self.headers.get("X-Filename", default)))

    def do_GET(self):
        path, _, query = self.path.partition("?")
        params = {k: unquote(v) for k, _, v in (p.partition("=") for p in query.split("&") if p)}
        if path == "/":
            self.send(200, PAGE_PATH.read_bytes(), "text/html; charset=utf-8")
        elif path == "/api/template":
            self.send_json(template_info())
        elif path == "/api/settings":
            self.send_json(config.public_settings())
        elif path == "/api/history":
            self.send_json(history.list_runs(params.get("q", ""), params.get("kind", "")))
        else:
            self.send(404, b"not found", "text/plain")

    def do_POST(self):
        try:
            self.route()
        except (zipfile.BadZipFile, KeyError):
            self.send_json({"error": "Это не похоже на файл Word (.docx)."}, 400)
        except gemini.GeminiError as e:
            self.send_json({"error": str(e)}, 400)
        except Exception as e:  # показываем ошибку в интерфейсе
            self.send_json({"error": str(e)}, 400)

    def route(self):
        p = self.path
        # ---- вкладка 1 ----
        if p == "/api/parse":
            data = self.body()
            items = parse_contract(data)
            STATE["contract"], STATE["contract_name"] = data, self.filename("договор.docx")
            self.send_json({"items": items})
        elif p == "/api/template":
            data = self.body()
            template_placeholders(data)  # проверка, что это docx
            STATE["tpl_bytes"] = data
            STATE["tpl_name"] = self.filename("шаблон.docx")
            self.send_json(template_info())
        elif p == "/api/generate":
            req = json.loads(self.body())
            if not STATE["tpl_bytes"]:
                raise ValueError("Не выбран шаблон протокола.")
            doc = build_protocol(STATE["tpl_bytes"], req.get("fields", {}), req["items"])
            name = stem(STATE["contract_name"] or "договор")
            inputs = [(STATE["contract_name"], STATE["contract"])] if STATE["contract"] else []
            meta = history.save_run("protocol", name, inputs,
                                    [(f"Протокол разногласий — {name}.docx", doc)],
                                    extra={"items": len(req["items"]), "template": STATE["tpl_name"]})
            self.send(200, doc, DOCX_MIME, {"X-Run-Id": quote(meta["id"])})
        # ---- вкладка 2 ----
        elif p == "/api/final/contract":
            data = self.body()
            info = redline.contract_info(data)
            with LOCK:
                STATE["final"].update(contract=data, name=self.filename("договор.docx"))
            info["name"] = STATE["final"]["name"]
            self.send_json(info)
        elif p == "/api/final/source":
            data = self.body()
            name = self.filename()
            with LOCK:
                STATE["final"]["seq"] += 1
                sid = f"s{STATE['final']['seq']}"
            res = sources_mod.parse_source(data, name, sid, use_gemini=self.headers.get("X-Use-Gemini") == "1")
            STATE["final"]["sources"][sid] = {"bytes": data, "filename": name, "meta": res["meta"],
                                              "ocr_text": res.get("ocr_text", "")}
            res.update(id=sid, filename=name)
            res.pop("ocr_text", None)
            self.send_json(res)
        elif p == "/api/final/reparse":
            req = json.loads(self.body())
            src = STATE["final"]["sources"][req["id"]]
            res = edits_mod.parse_protocol(src["bytes"], req["id"], new_col=req.get("new_col"))
            self.send_json(res)
        elif p == "/api/final/drop-source":
            STATE["final"]["sources"].pop(json.loads(self.body()).get("id"), None)
            self.send_json({"ok": True})
        elif p in ("/api/final/preview", "/api/final/build"):
            req = json.loads(self.body())
            fin = STATE["final"]
            if not fin["contract"]:
                raise ValueError("Сначала загрузите исходный договор.")
            srcs = final_sources(req)
            doc, report = redline.build_final(fin["contract"], srcs)
            if p.endswith("preview"):
                self.send_json(report)
                return
            name = stem(fin["name"])
            out_name = f"{name} — итоговая редакция.docx"
            inputs = [(fin["name"], fin["contract"])]
            outputs = [(out_name, doc)]
            for s in srcs:
                stored = fin["sources"][s["id"]]
                inputs.append((stored["filename"], stored["bytes"]))
                if stored.get("ocr_text"):
                    outputs.append((stem(stored["filename"]) + ".распознано.txt", stored["ocr_text"]))
            src_meta = []
            for s, rep in zip(srcs, report["sources"]):
                on = [r for r in rep["results"] if r.get("status") != "off"]
                src_meta.append({"label": s["label"], "file": s["filename"], "total": len(on),
                                 "applied": sum(1 for r in on if r.get("status") in ("exact", "by_text"))})
            md = history.report_md(fin["name"], srcs, report)
            meta = history.save_run("final", name, inputs, outputs, report_md=md,
                                    extra={"sources": src_meta, "warnings": report["warnings"]})
            self.send(200, doc, DOCX_MIME, {"X-Run-Id": quote(meta["id"]), "X-Filename": quote(out_name)})
        # ---- история ----
        elif p == "/api/reveal":
            req = json.loads(self.body())
            what = req.get("what", "folder")
            if what == "archive":
                history.ARCHIVE.mkdir(exist_ok=True)
                history.open_file(history.ARCHIVE)
            else:
                run_id = unquote(req.get("id", ""))
                run_dir = history.resolve(run_id)
                meta = json.loads((run_dir / "meta.json").read_text(encoding="utf8"))
                first = (meta.get("outputs") or [None])[0]
                if what == "open" and first:
                    history.open_file(history.resolve(run_id, first))
                elif what == "result" and first:
                    history.reveal(history.resolve(run_id, first))
                else:
                    history.reveal(run_dir)
            self.send_json({"ok": True})
        # ---- настройки ----
        elif p == "/api/settings":
            req = json.loads(self.body() or b"{}")
            self.send_json(config.save_settings(req.get("key"), req.get("model")))
        elif p == "/api/settings/test":
            gemini.ping()
            self.send_json({"ok": True})
        else:
            self.send(404, b"not found", "text/plain")


def free_port(start=8765):
    for port in range(start, start + 50):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("Нет свободного порта")


def main():
    load_default_template()
    port = free_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"Программа запущена: {url}")
    print("Чтобы закрыть программу, закройте это окно или нажмите Ctrl+C.")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
