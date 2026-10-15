"""
Применение правок к договору в режиме рецензирования Word (w:ins / w:del).

Модель: документ меняется «на месте» — неизменённый текст остаётся в своих
run'ах с исходным форматированием, удалённый оборачивается в w:del, новый
вставляется в w:ins. Автор каждой правки — документ-источник.
"""
import copy
import datetime as dt
import difflib
import re
import xml.etree.ElementTree as ET

from docxcore import (MC_FALLBACK, W, XML_SPACE, Package, exists_in, iter_paragraphs, norm,
                      norm_cmp, para_mark)
from numbering import Numbering, NUM_RE, text_num

REV_TAGS = {W + "ins", W + "del", W + "moveTo", W + "moveFrom"}
_ATOM_SKIP = {W + "pPr", W + "rPr", W + "instrText", W + "delInstrText", W + "txbxContent",
              W + "drawing", W + "pict", MC_FALLBACK}
TOKEN_RE = re.compile(r"\w+|\s+|[^\w\s]")


def now_iso():
    return dt.datetime.now().replace(microsecond=0).isoformat() + "Z"


def tokens(text):
    return [(m.start(), m.end(), m.group(0)) for m in TOKEN_RE.finditer(text)]


def token_key(t):
    return " " if t.isspace() else norm(t) or t


def strip_num(text):
    m = text_num(text)
    return text[m.end():].lstrip() if m else text.lstrip()


def similarity(a, b):
    a, b = norm_cmp(strip_num(a)), norm_cmp(strip_num(b))
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    if sm.real_quick_ratio() < 0.3:
        return 0.0
    return sm.ratio()


def split_lines(text):
    return [ln for ln in (norm(x) for x in (text or "").split("\n")) if ln]


# ---------------------------------------------------------------------------
# Документ с правками
# ---------------------------------------------------------------------------

class Doc:
    def __init__(self, data):
        self.pkg = Package(data)
        self.body = self.pkg.body
        self.numbering = Numbering(self.pkg.read("word/numbering.xml"), self.pkg.read("word/styles.xml"))
        ids = [int(v) for el in self.pkg.root.iter() for k, v in el.attrib.items()
               if k == W + "id" and v.lstrip("-").isdigit()]
        self.next_id = max(ids + [0]) + 1
        self.date = now_iso()
        self.author = "Правки"
        self.fresh_dels = []  # w:del, созданные текущей операцией (к ним можно дописывать)

    def save(self):
        return self.pkg.save()

    # --- служебное -----------------------------------------------------------

    def new_id(self):
        self.next_id += 1
        return str(self.next_id)

    def rev(self, tag):
        return ET.Element(W + tag, {W + "id": self.new_id(), W + "author": self.author, W + "date": self.date})

    def parents(self, root=None):
        root = root if root is not None else self.body
        return {c: p for p in root.iter() for c in p}

    # --- текст абзаца: «атомы» — видимые символы с привязкой к XML -------------

    def atoms(self, p):
        """[(символ, run, узел, смещение)] видимого (принятого) текста абзаца."""
        out = []

        def walk(el, run):
            tag = el.tag
            if tag in _ATOM_SKIP or tag in (W + "del", W + "moveFrom"):
                return
            if tag == W + "r":
                run = el
            if run is not None:
                if tag == W + "t":
                    for i, ch in enumerate(el.text or ""):
                        out.append((ch, run, el, i))
                elif tag == W + "tab":
                    out.append(("\t", run, el, 0))
                elif tag in (W + "br", W + "cr"):
                    out.append(("\n", run, el, 0))
                elif tag == W + "noBreakHyphen":
                    out.append(("-", run, el, 0))
            for ch in el:
                walk(ch, run)
        walk(p, None)
        return out

    def text(self, p):
        return "".join(a[0] for a in self.atoms(p))

    # --- разрезание run'а ------------------------------------------------------

    def split_run_at(self, p, pos):
        """Гарантирует границу run'ов перед символом pos."""
        at = self.atoms(p)
        if pos <= 0 or pos >= len(at):
            return
        (_, r1, n1, _), (_, r2, n2, off2) = at[pos - 1], at[pos]
        if r1 is not r2:
            return
        run = r2
        children = list(run)
        if n2.tag == W + "t" and off2 > 0:
            # делим w:t на два
            tail = ET.Element(W + "t", {XML_SPACE: "preserve"})
            tail.text = n2.text[off2:]
            n2.text = n2.text[:off2]
            n2.set(XML_SPACE, "preserve")
            idx = children.index(n2)
            run.insert(idx + 1, tail)
            children = list(run)
            first_moved = idx + 1
        else:
            first_moved = children.index(n2)
        new_run = ET.Element(W + "r", dict(run.attrib))
        rpr = run.find(W + "rPr")
        if rpr is not None:
            new_run.append(copy.deepcopy(rpr))
        for ch in children[first_moved:]:
            run.remove(ch)
            new_run.append(ch)
        parent = self.parents(p)[run]
        parent.insert(list(parent).index(run) + 1, new_run)

    # --- размещение элемента рядом с run'ом (не внутри чужой правки) ---------

    def place(self, p, ref, new, after=True):
        parents = self.parents(p)
        parent = parents[ref]
        while parent.tag in REV_TAGS:
            kids = list(parent)
            idx = kids.index(ref)
            rest = kids[idx + 1:] if after else kids[idx:]
            if rest and (after or idx > 0):
                clone = ET.Element(parent.tag, dict(parent.attrib))
                clone.set(W + "id", self.new_id())
                for ch in rest:
                    parent.remove(ch)
                    clone.append(ch)
                gp = parents[parent]
                gp.insert(list(gp).index(parent) + 1, clone)
                ref = parent if after else clone
            else:
                ref = parent
            parents = self.parents(p)
            parent = parents[ref]
        parent.insert(list(parent).index(ref) + (1 if after else 0), new)

    def make_runs(self, text, rpr):
        runs = []
        for part in re.split(r"([\t\n])", text):
            if not part:
                continue
            r = ET.Element(W + "r")
            if rpr is not None:
                rp = copy.deepcopy(rpr)
                for bad in rp.findall(W + "rPrChange"):
                    rp.remove(bad)
                r.append(rp)
            if part == "\t":
                ET.SubElement(r, W + "tab")
            elif part == "\n":
                ET.SubElement(r, W + "br")
            else:
                t = ET.SubElement(r, W + "t", {XML_SPACE: "preserve"})
                t.text = part
            runs.append(r)
        return runs

    # --- удаление run'а ----------------------------------------------------

    def delete_run(self, p, run):
        for el in run.iter():
            if el.tag == W + "t":
                el.tag = W + "delText"
            elif el.tag == W + "instrText":
                el.tag = W + "delInstrText"
        parent = self.parents(p)[run]
        idx = list(parent).index(run)
        prev = parent[idx - 1] if idx > 0 else None
        if prev is not None and any(prev is d for d in self.fresh_dels):
            parent.remove(run)
            prev.append(run)
            return prev
        wrap = self.rev("del")
        self.fresh_dels.append(wrap)
        parent.remove(run)
        wrap.append(run)
        parent.insert(idx, wrap)
        return wrap

    def visible_runs(self, p):
        out = []

        def walk(el):
            if el.tag in _ATOM_SKIP or el.tag in (W + "del", W + "moveFrom"):
                return
            if el.tag == W + "r":
                out.append(el)
                return
            for ch in el:
                walk(ch)
        walk(p)
        return out

    # --- главная операция: заменить символы [a, b) строкой s ----------------

    def replace_range(self, p, a, b, s):
        self.split_run_at(p, b)
        self.split_run_at(p, a)
        at = self.atoms(p)
        anchor, after = None, True
        if b > a:
            first, last = at[a][1], at[b - 1][1]
            runs = self.visible_runs(p)
            i1, i2 = runs.index(first), runs.index(last)
            for run in runs[i1:i2 + 1]:
                anchor = self.delete_run(p, run)
        if s:
            rpr_src = at[a - 1][1] if a > 0 else (at[a][1] if a < len(at) else None)
            rpr = rpr_src.find(W + "rPr") if rpr_src is not None else None
            if rpr is None and rpr_src is None:
                rpr = p.find(W + "pPr/" + W + "rPr")
                if rpr is not None:
                    rpr = copy.deepcopy(rpr)
                    for bad in list(rpr):
                        if bad.tag in REV_TAGS or bad.tag == W + "rPrChange":
                            rpr.remove(bad)
            ins = self.rev("ins")
            for r in self.make_runs(s, rpr):
                ins.append(r)
            if anchor is not None:
                self.place(p, anchor, ins, after=True)
            elif a > 0:
                self.place(p, at[a - 1][1], ins, after=True)
            elif at:
                self.place(p, at[0][1], ins, after=False)
            else:
                p.append(ins)
        self.fresh_dels = []

    def diff_ops(self, old, new):
        """Пословный diff: [(a, b, вставка)] в координатах old."""
        ta, tb = tokens(old), tokens(new)
        sm = difflib.SequenceMatcher(None, [token_key(t[2]) for t in ta], [token_key(t[2]) for t in tb],
                                     autojunk=False)
        ops = [list(o) for o in sm.get_opcodes()]
        # пробел между двумя изменениями приклеиваем к ним — правка читается целиком
        merged = []
        for op in ops:
            if (merged and op[0] != "equal" and len(merged) >= 2 and merged[-1][0] == "equal"
                    and merged[-2][0] != "equal"
                    and all(ta[k][2].isspace() for k in range(merged[-1][1], merged[-1][2]))):
                eq = merged.pop()
                prev = merged.pop()
                op = ["replace", prev[1], op[2], prev[3], op[4]]
            elif merged and op[0] != "equal" and merged[-1][0] != "equal":
                prev = merged.pop()
                op = ["replace", prev[1], op[2], prev[3], op[4]]
            merged.append(op)
        out = []
        for tag, i1, i2, j1, j2 in merged:
            if tag == "equal":
                continue
            a = ta[i1][0] if i1 < len(ta) else len(old)
            b = ta[i2 - 1][1] if i2 > i1 else a
            s = new[tb[j1][0]:tb[j2 - 1][1]] if j2 > j1 else ""
            out.append((a, b, s))
        return out

    def set_text(self, p, new_text):
        old = self.text(p)
        # пробелы по краям абзаца в протоколах теряются — исходные не трогаем
        if new_text.strip():
            lead = old[:len(old) - len(old.lstrip())]
            trail = old[len(old.rstrip()):]
            new_text = lead + new_text.strip() + trail
        ops = self.diff_ops(old, new_text)
        for a, b, s in reversed(ops):
            self.replace_range(p, a, b, s)
        return bool(ops)

    # --- абзацы целиком ---------------------------------------------------

    def ensure_mark_rpr(self, p):
        ppr = p.find(W + "pPr")
        if ppr is None:
            ppr = ET.Element(W + "pPr")
            p.insert(0, ppr)
        rpr = ppr.find(W + "rPr")
        if rpr is None:
            rpr = ET.Element(W + "rPr")
            tail = [i for i, ch in enumerate(ppr) if ch.tag in (W + "sectPr", W + "pPrChange")]
            ppr.insert(tail[0] if tail else len(ppr), rpr)
        return rpr

    def delete_paragraph(self, p):
        n = len(self.atoms(p))
        if n:
            self.replace_range(p, 0, n, "")
        parent = self.parents()[p]
        kids = list(parent)
        nxt = kids[kids.index(p) + 1] if kids.index(p) + 1 < len(kids) else None
        if nxt is not None and nxt.tag == W + "p":
            rpr = self.ensure_mark_rpr(p)
            if rpr.find(W + "del") is None:
                pos = 1 if rpr.find(W + "ins") is not None else 0
                rpr.insert(pos, self.rev("del"))

    def new_paragraph(self, template, text, numbering="keep", ilvl=None):
        p = ET.Element(W + "p")
        tppr = template.find(W + "pPr") if template is not None else None
        ppr = copy.deepcopy(tppr) if tppr is not None else ET.Element(W + "pPr")
        for bad in ppr.findall(W + "sectPr") + ppr.findall(W + "pPrChange"):
            ppr.remove(bad)
        npr = ppr.find(W + "numPr")
        if numbering == "none":
            if npr is not None:
                ppr.remove(npr)
            if template is not None and self.numbering.para_numpr(template) is not None:
                npr = ET.Element(W + "numPr")
                ET.SubElement(npr, W + "ilvl", {W + "val": "0"})
                ET.SubElement(npr, W + "numId", {W + "val": "0"})
                ppr.insert(1 if ppr.find(W + "pStyle") is not None else 0, npr)
        elif ilvl is not None:
            if npr is None:
                cur = self.numbering.para_numpr(template)
                if cur is not None:
                    npr = ET.Element(W + "numPr")
                    ET.SubElement(npr, W + "ilvl", {W + "val": str(ilvl)})
                    ET.SubElement(npr, W + "numId", {W + "val": cur[0]})
                    ppr.insert(1 if ppr.find(W + "pStyle") is not None else 0, npr)
            else:
                lv = npr.find(W + "ilvl")
                if lv is None:
                    lv = ET.Element(W + "ilvl")
                    npr.insert(0, lv)
                lv.set(W + "val", str(ilvl))
        p.append(ppr)
        rpr = self.ensure_mark_rpr(p)
        for bad in list(rpr):
            if bad.tag in REV_TAGS or bad.tag == W + "rPrChange":
                rpr.remove(bad)
        rpr.insert(0, self.rev("ins"))
        run_rpr = None
        if template is not None:
            runs = self.visible_runs(template)
            if runs:
                run_rpr = runs[0].find(W + "rPr")
        if text:
            ins = self.rev("ins")
            for r in self.make_runs(text, run_rpr):
                ins.append(r)
            p.append(ins)
        return p

    def insert_paragraph(self, ref, new_p, after=True):
        parent = self.parents()[ref]
        parent.insert(list(parent).index(ref) + (1 if after else 0), new_p)
        return new_p


# ---------------------------------------------------------------------------
# Карта пунктов текущей (принятой) редакции
# ---------------------------------------------------------------------------

def is_heading_text(t):
    return bool(t) and len(t) <= 80 and t[-1] not in ".;:," and not text_num(t)


class Snapshot:
    """Абзацы принятой редакции с номерами пунктов."""

    def __init__(self, doc):
        self.doc = doc
        allp = iter_paragraphs(doc.body)
        nums = doc.numbering.compute(allp, "new")
        self.paras, self.index = [], {}
        for p, n in zip(allp, nums):
            if not exists_in(p, "new"):
                continue
            text = doc.text(p)
            key, auto = "", False
            if n and n.get("key"):
                key, auto = n["key"], True
            elif n is None or not n.get("text"):
                m = text_num(text)
                if m:
                    key = m.group(1).rstrip(".")
            elif n and n.get("text"):
                auto = True  # «а)», «•» — подпункт, не пункт
            self.index[p] = len(self.paras)
            self.paras.append({"p": p, "key": key, "auto": auto, "text": text,
                               "numbered": bool(n and n.get("text")) or bool(key)})
        self.by_key = {}
        for i, pa in enumerate(self.paras):
            if pa["key"]:
                self.by_key.setdefault(pa["key"], []).append(i)

    def _heading_at(self, j):
        t = norm(self.paras[j]["text"])
        if not is_heading_text(t) or self.paras[j]["key"]:
            return False
        nxt = next((k for k in range(j + 1, len(self.paras)) if norm(self.paras[k]["text"])), None)
        return nxt is not None and bool(self.paras[nxt]["key"])

    def own(self, i):
        """Абзац пункта и следующие за ним абзацы без номера."""
        out = [i]
        for j in range(i + 1, len(self.paras)):
            if self.paras[j]["key"] or self._heading_at(j):
                break
            out.append(j)
        while len(out) > 1 and not norm(self.paras[out[-1]]["text"]):
            out.pop()
        return out

    def full(self, i):
        """Пункт со всеми подпунктами."""
        key = self.paras[i]["key"]
        out = [i]
        for j in range(i + 1, len(self.paras)):
            k = self.paras[j]["key"]
            if (k and not k.startswith(key + ".")) or self._heading_at(j):
                break
            out.append(j)
        while len(out) > 1 and not norm(self.paras[out[-1]]["text"]):
            out.pop()
        return out

    def window(self, i, n):
        """n непустых абзацев подряд начиная с i (пустые между ними — тоже)."""
        out, cnt = [], 0
        for j in range(i, len(self.paras)):
            out.append(j)
            if norm(self.paras[j]["text"]):
                cnt += 1
                if cnt == n:
                    return out
        return out if cnt == n else None

    def nonempty(self, idxs):
        return [j for j in idxs if norm(self.paras[j]["text"])]

    def text_of(self, idxs):
        return " ".join(norm(self.paras[j]["text"]) for j in idxs if norm(self.paras[j]["text"]))


# ---------------------------------------------------------------------------
# Применение правок источника
# ---------------------------------------------------------------------------

EXACT, LOOSE = 0.85, 0.6


def align(old, new):
    """Выравнивание абзацев: [(i|None, j|None)] с максимальным сходством."""
    n, m = len(old), len(new)
    sim = [[similarity(old[i], new[j]) for j in range(m)] for i in range(n)]
    best = [[0.0] * (m + 1) for _ in range(n + 1)]
    move = [[None] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        for j in range(m + 1):
            if i == 0 and j == 0:
                continue
            cands = []
            if i > 0:
                cands.append((best[i - 1][j], "d"))
            if j > 0:
                cands.append((best[i][j - 1], "i"))
            if i > 0 and j > 0:
                bonus = 0.3 if i == 1 and j == 1 else 0
                s = sim[i - 1][j - 1]
                if s >= 0.35 or bonus:
                    cands.append((best[i - 1][j - 1] + s - 0.35 + bonus, "m"))
            best[i][j], move[i][j] = max(cands, key=lambda c: c[0])
    out, i, j = [], n, m
    while i > 0 or j > 0:
        mv = move[i][j]
        if mv == "m":
            out.append((i - 1, j - 1))
            i, j = i - 1, j - 1
        elif mv == "d":
            out.append((i - 1, None))
            i -= 1
        else:
            out.append((None, j - 1))
            j -= 1
    return out[::-1]


def predecessor(key):
    parts = key.split(".")
    try:
        last = int(parts[-1])
    except ValueError:
        return None, None
    if last > 1:
        return "sibling", ".".join(parts[:-1] + [str(last - 1)])
    if len(parts) > 1:
        return "parent", ".".join(parts[:-1])
    return None, None


class Applier:
    def __init__(self, doc):
        self.doc = doc
        self.warnings = []

    def apply_source(self, edits, author):
        doc = self.doc
        doc.author = author
        base = Snapshot(doc)
        base_keys = {pa["p"]: pa["key"] for pa in base.paras if pa["key"]}
        base_dups = {k for k, v in base.by_key.items() if len(v) > 1}
        touched = set()
        cursor = None
        results = []
        for e in edits:
            if not e.get("enabled", True):
                results.append({"id": e.get("id"), "status": "off"})
                continue
            try:
                res = self.apply_edit(e, base, cursor)
            except Exception as ex:  # одна неудачная правка не должна ронять сборку
                res = {"status": "error", "note": f"Ошибка применения: {ex}"}
            res["id"] = e.get("id")
            if res.get("cursor") is not None:
                cursor = res.pop("cursor")
            touched.update(res.pop("touched", []))
            results.append(res)
        self.check_numbering(base_keys, base_dups, touched, author)
        return results

    # --- поиск ---------------------------------------------------------------

    def locate(self, snap, e, cursor, full=False):
        """Абзацы пункта в snap: (индексы, статус, score)."""
        old_lines = split_lines(e.get("old_text"))
        targets = [t for t in (e.get("targets") or [e.get("target")]) if t]
        cur_idx = snap.index.get(cursor, -1) if cursor is not None else -1

        def extent(i):
            if old_lines:
                w = snap.window(i, len(old_lines))
                if w:
                    return w
            return snap.full(i) if full else snap.own(i)

        cands = []
        for t in targets[:1]:
            for i in snap.by_key.get(t, []):
                ext = extent(i)
                if len(targets) > 1 and not old_lines:
                    for t2 in targets[1:]:
                        for k in snap.by_key.get(t2, []):
                            if k > ext[-1]:
                                ext = ext + list(range(ext[-1] + 1, (snap.own(k))[-1] + 1))
                                break
                score = similarity(" ".join(old_lines), snap.text_of(ext)) if old_lines else 1.0
                cands.append((score, -(abs(i - cur_idx) if i >= cur_idx else 10000 + cur_idx - i), ext))
        if cands:
            cands.sort(key=lambda c: (c[0], c[1]), reverse=True)
            score, _, ext = cands[0]
            if not old_lines or score >= EXACT:
                return ext, "exact", score
        if not old_lines:
            return None, "not_found", 0.0
        # поиск по тексту
        want = " ".join(old_lines)
        first = old_lines[0]
        best = (0.0, None)
        for i, pa in enumerate(snap.paras):
            t = norm(pa["text"])
            if not t or similarity(first, t) < 0.5:
                continue
            w = snap.window(i, len(old_lines))
            if not w:
                continue
            s = similarity(want, snap.text_of(w))
            if s > best[0]:
                best = (s, w)
        if cands and cands[0][0] >= best[0] and cands[0][0] >= LOOSE:
            return cands[0][2], "by_text", cands[0][0]
        if best[1] is not None and best[0] >= LOOSE:
            return best[1], "by_text", best[0]
        return None, "not_found", max(best[0], cands[0][0] if cands else 0.0)

    # --- правки -------------------------------------------------------------

    def apply_edit(self, e, base, cursor):
        action = e.get("action")
        if action == "insert":
            return self.do_insert(e, cursor)
        if action == "replace_words":
            return self.do_replace_words(e, base, cursor)
        if action == "renumber":
            return self.do_renumber(e, base, cursor)
        ext, status, score = self.locate(base, e, cursor, full=(action == "delete_clause"))
        if ext is None:
            return {"status": "not_found", "score": round(score, 2),
                    "note": "Пункт не найден в договоре — исправьте номер или текст"}
        paras = [base.paras[j]["p"] for j in ext]
        if any(not exists_in(p, "new") for p in paras):
            return {"status": "skipped", "note": "Пункт уже исключён предыдущей правкой"}
        if action == "delete_clause":
            for p in paras:
                self.doc.delete_paragraph(p)
            return {"status": status, "score": round(score, 2), "cursor": paras[-1], "touched": paras,
                    "found": base.paras[ext[0]]["key"], "found_text": base.text_of(ext)[:600]}
        if action == "replace_clause":
            last = self.replace_clause(base, ext, e.get("new_text", ""))
            return {"status": status, "score": round(score, 2), "cursor": last, "touched": paras,
                    "found": base.paras[ext[0]]["key"], "found_text": base.text_of(ext)[:600]}
        return {"status": "error", "note": f"Неизвестное действие: {action}"}

    def fix_number(self, text, para, auto):
        """Номер в начале нового текста: для автонумерации убираем, для текстовой — сохраняем."""
        m = text_num(text)
        if auto:
            return text[m.end():].lstrip() if m else text
        if not m and para is not None:
            pm = text_num(para["text"])
            if pm:
                sep = para["text"][pm.end():pm.end() + 1] if pm.end() < len(para["text"]) else " "
                return pm.group(1) + (sep if sep.isspace() else " ") + text.lstrip()
        return text

    def replace_clause(self, snap, ext, new_text):
        doc = self.doc
        idxs = snap.nonempty(ext) or ext[:1]
        paras = [snap.paras[j] for j in idxs]
        lines = split_lines(new_text)
        if paras and lines:
            lines[0] = self.fix_number(lines[0], paras[0], paras[0]["auto"])
        pairs = align([pa["text"] for pa in paras], lines)
        last_p = paras[-1]["p"] if paras else None
        prev_p = None
        for i, j in pairs:
            if i is not None and j is not None:
                doc.set_text(paras[i]["p"], lines[j])
                prev_p = paras[i]["p"]
            elif i is not None:
                doc.delete_paragraph(paras[i]["p"])
                prev_p = paras[i]["p"]
            else:
                if prev_p is None:
                    tpl = paras[0]["p"]
                    newp = doc.new_paragraph(tpl, lines[j], "keep")
                    doc.insert_paragraph(tpl, newp, after=False)
                else:
                    # продолжение пункта: формат соседнего абзаца без номера, если есть
                    tpl_pa = next((pa for pa in paras[1:] if not pa["numbered"]), None)
                    has_num = bool(text_num(lines[j]))
                    if tpl_pa is not None and not has_num:
                        newp = doc.new_paragraph(tpl_pa["p"], lines[j], "keep")
                    elif has_num and not paras[0]["auto"]:
                        newp = doc.new_paragraph(paras[0]["p"], lines[j], "keep")
                    else:
                        newp = doc.new_paragraph(paras[0]["p"], lines[j], "none")
                    doc.insert_paragraph(prev_p, newp, after=True)
                    prev_p = newp
                last_p = prev_p if prev_p is not None else last_p
        return prev_p if prev_p is not None else last_p

    def do_insert(self, e, cursor):
        doc = self.doc
        snap = Snapshot(doc)
        lines = split_lines(e.get("new_text"))
        if not lines:
            return {"status": "not_found", "note": "Нет текста для вставки"}
        key = (e.get("target") or "").strip()
        anchor_key = (e.get("anchor") or "").strip()
        cur_idx = snap.index.get(cursor, -1) if cursor is not None else -1

        def pick(cands):
            above = [i for i in cands if i <= cur_idx]
            return above[-1] if above else cands[0]

        anchor_i, mode, tpl_i, status = None, "full", None, "exact"
        if anchor_key and snap.by_key.get(anchor_key):
            anchor_i = pick(snap.by_key[anchor_key])
            tpl_i = anchor_i
        elif key:
            kind, pk = predecessor(key)
            if kind and snap.by_key.get(pk):
                anchor_i = pick(snap.by_key[pk])
                if kind == "parent":
                    mode = "own"
                else:
                    tpl_i = anchor_i
            elif kind:
                # нет пункта X−1: ближайший предыдущий пункт того же уровня, иначе конец раздела
                parts = key.split(".")
                for k in range(int(parts[-1]) - 2, 0, -1):
                    sib = ".".join(parts[:-1] + [str(k)])
                    if snap.by_key.get(sib):
                        anchor_i, tpl_i, status = pick(snap.by_key[sib]), None, "by_text"
                        tpl_i = anchor_i
                        break
                if anchor_i is None and len(parts) > 1:
                    parent = ".".join(parts[:-1])
                    if snap.by_key.get(parent):
                        anchor_i, status = pick(snap.by_key[parent]), "by_text"
        if anchor_i is None and cursor is not None and cursor in snap.index and not key:
            anchor_i, mode, status = snap.index[cursor], "self", "by_text"
        if anchor_i is None:
            return {"status": "not_found", "note": "Не удалось определить, после какого пункта вставить"}
        if mode == "full":
            after_i = snap.full(anchor_i)[-1]
        elif mode == "own":
            after_i = snap.own(anchor_i)[-1]
        else:
            after_i = anchor_i
        # если курсор (предыдущая правка) стоит внутри этого же места — вставляем после него
        if cursor is not None and cursor in snap.index and anchor_i <= snap.index[cursor] <= after_i + 0:
            pass
        ilvl = None
        if tpl_i is None:
            depth = key.count(".") if key else None
            same = [i for i, pa in enumerate(snap.paras)
                    if depth is not None and pa["key"] and pa["key"].count(".") == depth]
            after_same = [i for i in same if i > after_i]
            if after_same:
                tpl_i = after_same[0]
            elif same:
                tpl_i = same[-1]
            else:
                tpl_i = anchor_i
                if key and snap.paras[anchor_i]["auto"]:
                    npr = doc.numbering.para_numpr(snap.paras[anchor_i]["p"])
                    ilvl = (npr[1] + 1) if npr else None
        tpl = snap.paras[tpl_i]
        numbered_new = bool(key)
        auto = tpl["auto"] and numbered_new
        first = self.fix_number(lines[0], None, auto)
        if not auto and key and not text_num(first) and tpl["key"] and not tpl["auto"]:
            pm = text_num(tpl["text"])
            dot = "." if pm and pm.group(1).endswith(".") else ""
            first = f"{key}{dot} {first}"
        if numbered_new:
            num_mode = "keep"
        else:
            num_mode = "none" if tpl["numbered"] else "keep"
        ref = snap.paras[after_i]["p"]
        newp = doc.new_paragraph(tpl["p"], first, num_mode, ilvl=ilvl)
        doc.insert_paragraph(ref, newp)
        last = newp
        cont_tpl = next((snap.paras[j] for j in snap.own(tpl_i)[1:] if not snap.paras[j]["numbered"]), None)
        for ln in lines[1:]:
            if cont_tpl is not None and not text_num(ln):
                q = doc.new_paragraph(cont_tpl["p"], ln, "keep")
            elif text_num(ln) and not tpl["auto"]:
                q = doc.new_paragraph(tpl["p"], ln, "keep")
            else:
                q = doc.new_paragraph(tpl["p"], self.fix_number(ln, None, tpl["auto"]), "none")
            doc.insert_paragraph(last, q)
            last = q
        return {"status": status, "score": 1.0, "cursor": last, "touched": [newp],
                "found": snap.paras[anchor_i]["key"] or "место предыдущей правки",
                "found_text": "после: " + norm(snap.paras[after_i]["text"])[:300]}

    def do_replace_words(self, e, base, cursor):
        old, new = e.get("old_text") or "", e.get("new_text") or ""
        if not norm(old):
            return {"status": "not_found", "note": "Не указаны заменяемые слова"}
        targets = [t for t in (e.get("targets") or [e.get("target")]) if t]
        if targets:
            idxs = []
            for t in targets:
                for i in base.by_key.get(t, [])[:1]:
                    idxs += base.full(i)
            if not idxs:
                return {"status": "not_found", "note": f"Пункт {targets[0]} не найден"}
        else:
            idxs = list(range(len(base.paras)))
        words = [re.escape(w.translate(_PLAIN)) for w in norm(old).lower().split(" ")]
        rx = re.compile(r"\s+".join(words))
        count, last = 0, None
        for j in idxs:
            p = base.paras[j]["p"]
            if not exists_in(p, "new"):
                continue
            text = self.doc.text(p)
            hits = list(rx.finditer(text.translate(_PLAIN).lower()))
            for m in reversed(hits):
                self.doc.replace_range(p, m.start(), m.end(), new)
                count += 1
                last = p
        if not count:
            return {"status": "not_found", "note": "Заменяемые слова не найдены в пункте"}
        return {"status": "exact", "score": 1.0, "cursor": last, "touched": [last],
                "note": f"Замен: {count}", "found": targets[0] if targets else "весь договор"}

    def do_renumber(self, e, base, cursor):
        t, new = (e.get("target") or "").strip(), (e.get("new_text") or "").strip().rstrip(".")
        idx = base.by_key.get(t)
        if not idx:
            return {"status": "not_found", "note": f"Пункт {t} не найден"}
        pa = base.paras[idx[0]]
        if pa["auto"]:
            return {"status": "exact", "note": "Автонумерация — Word перенумерует сам"}
        m = text_num(pa["text"])
        repl = new + ("." if m.group(1).endswith(".") else "")
        lead = len(pa["text"]) - len(pa["text"].lstrip())
        self.doc.replace_range(pa["p"], lead, m.end(), repl)
        return {"status": "exact", "score": 1.0, "cursor": pa["p"], "touched": [pa["p"]], "found": t}

    # --- проверки после источника -------------------------------------------

    REF_RE = re.compile(r"(?:\bп\.|\bпп\.|\bпункт\w*|\bподпункт\w*)\s*(\d+(?:\.\d+)+)", re.I)

    def check_numbering(self, base_keys, base_dups, touched, author):
        snap = Snapshot(self.doc)
        now_keys = {pa["p"]: pa["key"] for pa in snap.paras}
        moved = {}
        for p, k in base_keys.items():
            nk = now_keys.get(p)
            if nk != k:
                moved.setdefault(k, nk)
        existing = set(snap.by_key)
        # дубли и пропуски при текстовой нумерации
        dups = [k for k, v in snap.by_key.items()
                if len(v) > 1 and k not in base_dups and not snap.paras[v[0]]["auto"]]
        for k in dups:
            self.warnings.append(f"{author}: номер п. {k} встречается {len(snap.by_key[k])} раза — "
                                 "номера набраны текстом, проверьте нумерацию.")
        touched = set(touched)
        for pa in snap.paras:
            for m in self.REF_RE.finditer(pa["text"]):
                ref = m.group(1)
                where = pa["key"] or "без номера"
                if ref not in existing and ref not in moved:
                    continue
                if ref not in existing:
                    self.warnings.append(f"{author}: в п. {where} ссылка на п. {ref}, которого больше нет.")
                elif ref in moved and pa["p"] not in touched:
                    nk = moved[ref]
                    self.warnings.append(
                        f"{author}: в п. {where} ссылка на п. {ref} — "
                        + (f"этот пункт теперь имеет номер {nk}." if nk else "этот пункт исключён."))


_PLAIN = str.maketrans({"«": '"', "»": '"', "“": '"', "”": '"', "„": '"', "–": "-", "—": "-",
                        " ": " ", "\t": " ", "\n": " ", "ё": "е", "Ё": "Е"})


def build_final(contract_bytes, sources):
    """
    sources: [{label, edits: [...]}] в порядке применения.
    Возвращает (docx bytes, отчёт {sources: [{label, results}], warnings}).
    """
    doc = Doc(contract_bytes)
    from docxcore import has_revisions
    warnings = []
    if has_revisions(doc.body):
        warnings.append("В исходном договоре уже есть непринятые правки: они сохранены, "
                        "базой считается принятая редакция.")
    applier = Applier(doc)
    report = []
    for src in sources:
        res = applier.apply_source(src.get("edits", []), src.get("label") or "Изменения")
        report.append({"label": src.get("label"), "results": res})
    warnings += list(dict.fromkeys(applier.warnings))
    return doc.save(), {"sources": report, "warnings": warnings}


# ---------------------------------------------------------------------------
# Для тестов: принять / отклонить все правки
# ---------------------------------------------------------------------------

def view_paragraphs(data, view="new"):
    """Тексты абзацев редакции (с учётом слияния абзацев по знакам)."""
    from docxcore import para_text
    doc = Package(data)
    out, buf = [], ""
    for p in iter_paragraphs(doc.body):
        buf += para_text(p, view)
        if exists_in(p, view):
            if norm(buf):
                out.append(norm(buf))
            buf = ""
    if norm(buf):
        out.append(norm(buf))
    return out


def reject_all(data):
    """Отклонить все правки (для получения исходной редакции в тестах)."""
    pkg = Package(data)
    body = pkg.body
    # абзац со вставленным знаком сливается со следующим (берёт его свойства)
    parents = {c: p for p in pkg.root.iter() for c in p}
    for p in reversed(iter_paragraphs(body)):
        if para_mark(p) not in ("ins", "insdel"):
            continue
        par = parents[p]
        kids = list(par)
        nxt = kids[kids.index(p) + 1] if kids.index(p) + 1 < len(kids) else None
        if nxt is None or nxt.tag != W + "p":
            continue
        content = [c for c in p if c.tag != W + "pPr"]
        pos = 1 if nxt.find(W + "pPr") is not None else 0
        for k, c in enumerate(content):
            nxt.insert(pos + k, c)
        par.remove(p)
    parents = {c: p for p in pkg.root.iter() for c in p}
    for el in list(body.iter()):
        if el.tag in (W + "ins", W + "moveTo") and el in parents:
            par = parents[el]
            if par.tag == W + "rPr":
                par.remove(el)
            elif par.tag == W + "numPr":
                continue
            else:
                par.remove(el)
    parents = {c: p for p in pkg.root.iter() for c in p}
    for el in list(body.iter()):
        if el.tag in (W + "del", W + "moveFrom") and el in parents:
            par = parents[el]
            if par.tag == W + "rPr":
                par.remove(el)
                continue
            idx = list(par).index(el)
            par.remove(el)
            for k, ch in enumerate(list(el)):
                par.insert(idx + k, ch)
    for el in body.iter():
        if el.tag == W + "delText":
            el.tag = W + "t"
        elif el.tag == W + "delInstrText":
            el.tag = W + "instrText"
    for ppr in body.iter(W + "pPr"):
        ch = ppr.find(W + "pPrChange")
        if ch is not None:
            old = ch.find(W + "pPr")
            keep_rpr = ppr.find(W + "rPr")
            for c in list(ppr):
                ppr.remove(c)
            if old is not None:
                for c in old:
                    ppr.append(c)
            if keep_rpr is not None:
                ppr.append(keep_rpr)
    for rpr in body.iter(W + "rPr"):
        ch = rpr.find(W + "rPrChange")
        if ch is not None:
            old = ch.find(W + "rPr")
            for c in list(rpr):
                rpr.remove(c)
            if old is not None:
                for c in old:
                    rpr.append(c)
    # абзацы, у которых знак был вставлен, сливаются со следующим — для текста
    # достаточно перенести их содержимое в начало следующего абзаца
    return pkg.save()


def contract_info(data):
    """Сводка по договору для интерфейса."""
    from docxcore import has_revisions
    doc = Doc(data)
    snap = Snapshot(doc)
    keyed = [pa for pa in snap.paras if pa["key"]]
    auto = sum(1 for pa in keyed if pa["auto"])
    kind = "none" if not keyed else "auto" if auto == len(keyed) else "text" if auto == 0 else "mixed"
    return {"paragraphs": len(snap.paras), "clauses": len(keyed), "numbering": kind,
            "has_revisions": has_revisions(doc.body)}
