#!/usr/bin/env python3
"""
Протокол разногласий из правок Word.

Запуск: python3 protocol_app.py — откроется страница в браузере.
Работает только на этом компьютере: сервер слушает 127.0.0.1,
никаких внешних запросов, только стандартная библиотека Python.
"""
import html
import io
import json
import re
import socket
import sys
import threading
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

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

    def is_heading(k):
        """Короткий абзац без номера перед пронумерованным пунктом — заголовок."""
        t = texts[k]
        if not t or num_match(t) or len(t) > 80 or t[-1] in ".;:,":
            return False
        nxt = next((j for j in nonempty if j > k), None)
        return nxt is not None and bool(num_match(texts[nxt]))

    points, owner = [], {}
    for k, lp in enumerate(logical):
        m = num_match(texts[k])
        heading = is_heading(k)
        if m or heading or not points:
            points.append({"num": m.group(1).rstrip(".") if m else "",
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


def parse_contract(data):
    root = ET.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml"))
    body = root.find(W + "body")
    paras = [{"segs": para_segments(p), "mark": para_mark(p)} for p in body.iter(W + "p")]

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
    preferred = APP_DIR / "шаблон_протокола_разногласий.docx"
    if preferred.exists():
        return preferred
    found = sorted(p for p in APP_DIR.glob("*.docx")
                   if "шаблон" in p.name.lower() and not p.name.startswith("~$"))
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

STATE = {"tpl_bytes": None, "tpl_name": None}


def load_default_template():
    path = default_template()
    if path:
        STATE["tpl_bytes"] = path.read_bytes()
        STATE["tpl_name"] = path.name


def template_info():
    if not STATE["tpl_bytes"]:
        return {"name": None, "placeholders": []}
    return {"name": STATE["tpl_name"], "placeholders": template_placeholders(STATE["tpl_bytes"])}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def send(self, code, body, ctype):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, obj, code=200):
        self.send(code, json.dumps(obj, ensure_ascii=False).encode("utf8"), "application/json; charset=utf-8")

    def body(self):
        return self.rfile.read(int(self.headers.get("Content-Length", 0)))

    def do_GET(self):
        if self.path == "/":
            self.send(200, PAGE.encode("utf8"), "text/html; charset=utf-8")
        elif self.path == "/api/template":
            self.send_json(template_info())
        else:
            self.send(404, b"not found", "text/plain")

    def do_POST(self):
        try:
            if self.path == "/api/parse":
                self.send_json({"items": parse_contract(self.body())})
            elif self.path == "/api/template":
                data = self.body()
                template_placeholders(data)  # проверка, что это docx
                STATE["tpl_bytes"] = data
                STATE["tpl_name"] = unquote(self.headers.get("X-Filename", "шаблон.docx"))
                self.send_json(template_info())
            elif self.path == "/api/generate":
                req = json.loads(self.body())
                if not STATE["tpl_bytes"]:
                    raise ValueError("Не выбран шаблон протокола.")
                doc = build_protocol(STATE["tpl_bytes"], req.get("fields", {}), req["items"])
                self.send(200, doc, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                self.send(404, b"not found", "text/plain")
        except (zipfile.BadZipFile, KeyError):
            self.send_json({"error": "Это не похоже на файл Word (.docx)."}, 400)
        except Exception as e:  # показываем ошибку в интерфейсе
            self.send_json({"error": str(e)}, 400)


def free_port(start=8765):
    for port in range(start, start + 50):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("Нет свободного порта")


PAGE = r"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Протокол разногласий</title>
<style>
  :root {
    --bg: #f6f5f2; --card: #ffffff; --ink: #1d1d1f; --muted: #6b6b70; --line: #e3e1dc;
    --accent: #2f5d8a; --accent-ink: #ffffff;
    --ins-bg: #dcf3df; --ins-ink: #14532d; --del-bg: #fbe0de; --del-ink: #8a1c13;
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--ink);
         font: 15px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
  main { max-width: 1200px; margin: 0 auto; padding: 28px 20px 120px; }
  h1 { font-size: 24px; margin: 0 0 4px; }
  .sub { color: var(--muted); margin: 0 0 24px; }
  .drop { border: 2px dashed #c9c6bf; border-radius: 14px; background: var(--card);
          padding: 36px 20px; text-align: center; cursor: pointer; transition: .15s; }
  .drop.over { border-color: var(--accent); background: #eef3f8; }
  .drop b { font-size: 17px; }
  .drop small { display: block; color: var(--muted); margin-top: 6px; }
  .loaded { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .panel { background: var(--card); border: 1px solid var(--line); border-radius: 14px;
           padding: 18px 20px; margin-top: 18px; }
  .panel h2 { font-size: 15px; margin: 0 0 12px; }
  .fields { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 12px; }
  .fields label { display: block; font-size: 13px; color: var(--muted); }
  .fields input { width: 100%; margin-top: 4px; padding: 8px 10px; border: 1px solid var(--line);
                  border-radius: 8px; font: inherit; }
  .muted { color: var(--muted); font-size: 13px; }
  button, .btn { font: inherit; border: 1px solid var(--line); background: var(--card); color: var(--ink);
           padding: 7px 14px; border-radius: 8px; cursor: pointer; }
  button.primary { background: var(--accent); color: var(--accent-ink); border-color: var(--accent);
                   padding: 11px 22px; font-weight: 600; }
  button:disabled { opacity: .5; cursor: default; }
  .toolbar { display: flex; align-items: center; gap: 10px; margin: 26px 0 12px; flex-wrap: wrap; }
  .toolbar .count { margin-left: auto; color: var(--muted); }
  .legend span { padding: 1px 6px; border-radius: 4px; margin-left: 6px; font-size: 13px; }
  .item { background: var(--card); border: 1px solid var(--line); border-radius: 14px;
          margin-bottom: 12px; overflow: hidden; }
  .item.off { opacity: .55; }
  .item-head { display: flex; align-items: center; gap: 10px; padding: 12px 18px;
               border-bottom: 1px solid var(--line); cursor: pointer; user-select: none; }
  .item-head input { width: 18px; height: 18px; accent-color: var(--accent); }
  .num { font-weight: 700; }
  .tag { font-size: 12px; padding: 2px 8px; border-radius: 99px; background: #eceae5; color: var(--muted); }
  .cols { display: grid; grid-template-columns: 1fr 1fr; }
  .col { padding: 12px 18px 14px; font-size: 14px; }
  .col + .col { border-left: 1px solid var(--line); }
  .col h3 { margin: 0 0 6px; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }
  .col p { margin: 0 0 6px; }
  .col .empty { color: var(--muted); font-style: italic; }
  ins { background: var(--ins-bg); color: var(--ins-ink); text-decoration: none; border-radius: 3px; }
  del { background: var(--del-bg); color: var(--del-ink); border-radius: 3px; }
  .bar { position: fixed; left: 0; right: 0; bottom: 0; background: rgba(255,255,255,.95);
         border-top: 1px solid var(--line); padding: 14px 20px; }
  .bar-in { max-width: 1200px; margin: 0 auto; display: flex; align-items: center; gap: 16px; }
  .err { color: var(--del-ink); }
  [hidden] { display: none !important; }
  @media (max-width: 720px) { .cols { grid-template-columns: 1fr; } .col + .col { border-left: 0; border-top: 1px solid var(--line); } }
</style>
</head>
<body>
<main>
  <h1>Протокол разногласий из правок Word</h1>
  <p class="sub">Файлы обрабатываются только на этом компьютере и никуда не отправляются.</p>

  <div id="drop" class="drop">
    <b>Перетащите сюда договор с правками (.docx)</b>
    <small>или нажмите, чтобы выбрать файл</small>
  </div>
  <input id="file" type="file" accept=".docx" hidden>

  <div id="loadedBox" class="panel loaded" hidden>
    <span>Договор: <b id="fname"></b></span>
    <button id="another">Загрузить другой</button>
  </div>

  <div class="panel">
    <h2>Шаблон и данные для протокола</h2>
    <div class="loaded" style="margin-bottom:14px">
      <span>Шаблон: <b id="tplName">—</b></span>
      <button id="tplBtn">Выбрать другой шаблон</button>
      <input id="tplFile" type="file" accept=".docx" hidden>
    </div>
    <div id="fields" class="fields"></div>
    <p class="muted" id="fieldsHint" style="margin:10px 0 0">Пустые поля останутся в протоколе как {{…}} — их можно заполнить в Word.</p>
  </div>

  <p id="error" class="err"></p>

  <div id="listBox" hidden>
    <div class="toolbar">
      <button id="all">Выбрать все</button>
      <button id="none">Снять все</button>
      <span class="legend muted">Подсветка: <span style="background:var(--del-bg);color:var(--del-ink)"><s>удалено</s></span><span style="background:var(--ins-bg);color:var(--ins-ink)">добавлено</span></span>
      <span class="count" id="count"></span>
    </div>
    <div id="list"></div>
  </div>
</main>

<div class="bar" id="bar" hidden>
  <div class="bar-in">
    <button class="primary" id="make">Создать протокол разногласий</button>
    <span class="muted" id="barInfo"></span>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
let items = [], contractName = "";
const KIND = { new: "новый пункт", deleted: "пункт исключён", renum: "изменён только номер", changed: "" };

function esc(s) { return s.replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }

function renderParas(paras, emptyText) {
  if (!paras.length) return `<p class="empty">${emptyText}</p>`;
  return paras.map(segs => "<p>" + segs.map(([k, t]) =>
    k === "ins" ? `<ins>${esc(t)}</ins>` : k === "del" ? `<del>${esc(t)}</del>` : esc(t)).join("") + "</p>").join("");
}

function render() {
  $("list").innerHTML = items.map((it, i) => `
    <div class="item ${it.on ? "" : "off"}" data-i="${i}">
      <label class="item-head">
        <input type="checkbox" ${it.on ? "checked" : ""} data-i="${i}">
        <span class="num">П. ${esc(it.label)}</span>
        ${KIND[it.kind] ? `<span class="tag">${KIND[it.kind]}</span>` : ""}
        <span class="muted" style="margin-left:auto">включить в протокол разногласий</span>
      </label>
      <div class="cols">
        <div class="col"><h3>Было</h3>${renderParas(it.old, "Отсутствует")}</div>
        <div class="col"><h3>Стало</h3>${renderParas(it.new, "Исключить")}</div>
      </div>
    </div>`).join("");
  updateCount();
}

function updateCount() {
  const n = items.filter(x => x.on).length;
  $("count").textContent = `Пунктов с правками: ${items.length}, выбрано: ${n}`;
  $("barInfo").textContent = `В протокол попадёт пунктов: ${n}`;
  $("make").disabled = n === 0;
  document.querySelectorAll(".item").forEach(el => el.classList.toggle("off", !items[el.dataset.i].on));
}

$("list").addEventListener("change", e => {
  if (e.target.dataset.i !== undefined) { items[e.target.dataset.i].on = e.target.checked; updateCount(); }
});
$("all").onclick = () => { items.forEach(x => x.on = true); render(); };
$("none").onclick = () => { items.forEach(x => x.on = false); render(); };

async function post(url, body, headers = {}) {
  const r = await fetch(url, { method: "POST", body, headers });
  if (!r.ok) { let m = "Ошибка"; try { m = (await r.json()).error; } catch (_) {} throw new Error(m); }
  return r;
}

async function loadContract(file) {
  $("error").textContent = "";
  if (!file.name.toLowerCase().endsWith(".docx")) { $("error").textContent = "Нужен файл Word в формате .docx"; return; }
  try {
    const data = await (await post("/api/parse", file)).json();
    contractName = file.name.replace(/\.docx$/i, "");
    items = data.items.map(x => ({ ...x, on: true }));
    $("fname").textContent = file.name;
    $("drop").hidden = true; $("loadedBox").hidden = false;
    $("listBox").hidden = false; $("bar").hidden = false;
    if (!items.length) $("error").textContent = "В этом файле не найдено правок в режиме рецензирования.";
    render();
  } catch (e) { $("error").textContent = e.message; }
}

const drop = $("drop");
drop.onclick = () => $("file").click();
$("another").onclick = () => $("file").click();
$("file").onchange = e => { if (e.target.files[0]) loadContract(e.target.files[0]); e.target.value = ""; };
["dragenter", "dragover"].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
document.addEventListener("drop", e => { const f = e.dataTransfer.files[0]; if (f) loadContract(f); });

function showTemplate(info) {
  $("tplName").textContent = info.name || "не найден — выберите файл шаблона";
  const old = {};
  document.querySelectorAll("#fields input").forEach(i => old[i.name] = i.value);
  $("fields").innerHTML = info.placeholders.map(p =>
    `<label>${esc(p)}<input name="${esc(p)}" value="${esc(old[p] || "")}"></label>`).join("");
  $("fieldsHint").hidden = !info.placeholders.length;
}
$("tplBtn").onclick = () => $("tplFile").click();
$("tplFile").onchange = async e => {
  const f = e.target.files[0]; e.target.value = "";
  if (!f) return;
  try { showTemplate(await (await post("/api/template", f, { "X-Filename": encodeURIComponent(f.name) })).json()); $("error").textContent = ""; }
  catch (err) { $("error").textContent = err.message; }
};
fetch("/api/template").then(r => r.json()).then(showTemplate);

$("make").onclick = async () => {
  $("error").textContent = "";
  const fields = {};
  document.querySelectorAll("#fields input").forEach(i => fields[i.name] = i.value);
  const chosen = items.filter(x => x.on).map(({ label, old, new: nw }) => ({ label, old, new: nw }));
  try {
    const blob = await (await post("/api/generate", JSON.stringify({ fields, items: chosen }),
                                   { "Content-Type": "application/json" })).blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `Протокол разногласий — ${contractName}.docx`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    $("barInfo").textContent = "Готово: файл сохранён в папку «Загрузки».";
  } catch (e) { $("error").textContent = e.message; }
};
</script>
</body>
</html>
"""


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
