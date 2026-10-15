"""
Номера абзацев так, как их показывает Word: автонумерация (numbering.xml +
стили) и, если её нет, номер, набранный текстом в начале абзаца.
"""
import re
import xml.etree.ElementTree as ET

from docxcore import W, exists_in, para_text, wattr

# Номер пункта, набранный текстом: «1.», «2.1», «6.16.1.» (обязательно с точкой)
NUM_RE = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,3})*\.?)(?=\s)")
KEY_RE = re.compile(r"\d+(?:\.\d+)*")


def text_num(text):
    m = NUM_RE.match(text)
    return m if m and "." in m.group(1) else None


def clause_key(rendered):
    """«5.3.» -> «5.3»; «Статья 5.» -> «5»; «а)» -> «» (не пункт)."""
    found = KEY_RE.findall(rendered or "")
    return found[-1] if found else ""


# ---------------------------------------------------------------------------
# Форматы номеров
# ---------------------------------------------------------------------------

RU_LETTERS = "абвгдежзиклмнопрстуфхцчшщэюя"  # без ё, й, ъ, ы, ь — как в Word
ROMAN = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
         (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]


def _letters(n, alphabet):
    if n <= 0:
        return ""
    k = len(alphabet)
    return alphabet[(n - 1) % k] * ((n - 1) // k + 1)


def _roman(n):
    out = ""
    for v, s in ROMAN:
        while n >= v:
            out += s
            n -= v
    return out


def fmt_number(n, fmt):
    if fmt in ("decimal", None, ""):
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "lowerLetter":
        return _letters(n, "abcdefghijklmnopqrstuvwxyz")
    if fmt == "upperLetter":
        return _letters(n, "abcdefghijklmnopqrstuvwxyz").upper()
    if fmt == "lowerRoman":
        return _roman(n)
    if fmt == "upperRoman":
        return _roman(n).upper()
    if fmt == "russianLower":
        return _letters(n, RU_LETTERS)
    if fmt == "russianUpper":
        return _letters(n, RU_LETTERS).upper()
    if fmt in ("none", "bullet"):
        return ""
    return str(n)


# ---------------------------------------------------------------------------
# Описание списков
# ---------------------------------------------------------------------------

class Level:
    def __init__(self, el=None):
        self.start = 1
        self.fmt = "decimal"
        self.text = None
        self.restart = None   # None — сброс после любого более высокого уровня
        self.legal = False
        self.pstyle = None
        if el is not None:
            self.update(el)

    def update(self, el):
        def val(tag):
            return wattr(el.find(W + tag), "val")
        if val("start") is not None:
            self.start = int(val("start"))
        if val("numFmt") is not None:
            self.fmt = val("numFmt")
        if el.find(W + "lvlText") is not None:
            self.text = val("lvlText") or ""
        if val("lvlRestart") is not None:
            self.restart = int(val("lvlRestart"))
        if el.find(W + "isLgl") is not None:
            self.legal = wattr(el.find(W + "isLgl"), "val", "1") not in ("0", "false")
        if val("pStyle") is not None:
            self.pstyle = val("pStyle")


class Numbering:
    def __init__(self, numbering_xml=None, styles_xml=None):
        self.abstract = {}      # abstractNumId -> {ilvl: Level}
        self.abs_link = {}      # abstractNumId -> numStyleLink
        self.abs_stylelink = {}  # styleLink name -> abstractNumId
        self.nums = {}          # numId -> (abstractNumId, {ilvl: startOverride}, {ilvl: Level})
        self.styles = {}        # styleId -> {basedOn, numId, ilvl}
        self.default_style = None
        if numbering_xml:
            self._load_numbering(ET.fromstring(numbering_xml))
        if styles_xml:
            self._load_styles(ET.fromstring(styles_xml))

    def _load_numbering(self, root):
        for an in root.findall(W + "abstractNum"):
            aid = wattr(an, "abstractNumId")
            self.abstract[aid] = {int(wattr(l, "ilvl", "0")): Level(l) for l in an.findall(W + "lvl")}
            link = wattr(an.find(W + "numStyleLink"), "val")
            if link:
                self.abs_link[aid] = link
            slink = wattr(an.find(W + "styleLink"), "val")
            if slink:
                self.abs_stylelink[slink] = aid
        for num in root.findall(W + "num"):
            nid = wattr(num, "numId")
            aid = wattr(num.find(W + "abstractNumId"), "val")
            overrides, lvls = {}, {}
            for ov in num.findall(W + "lvlOverride"):
                ilvl = int(wattr(ov, "ilvl", "0"))
                so = ov.find(W + "startOverride")
                if so is not None:
                    overrides[ilvl] = int(wattr(so, "val", "1"))
                lv = ov.find(W + "lvl")
                if lv is not None:
                    lvls[ilvl] = lv
            self.nums[nid] = (aid, overrides, lvls)

    def _load_styles(self, root):
        for st in root.findall(W + "style"):
            if wattr(st, "type") != "paragraph":
                continue
            sid = wattr(st, "styleId")
            npr = st.find(W + "pPr/" + W + "numPr")
            self.styles[sid] = {
                "basedOn": wattr(st.find(W + "basedOn"), "val"),
                "numId": wattr(npr.find(W + "numId"), "val") if npr is not None else None,
                "ilvl": wattr(npr.find(W + "ilvl"), "val") if npr is not None else None,
            }
            if wattr(st, "default") in ("1", "true"):
                self.default_style = sid

    def style_numpr(self, sid):
        """numId/ilvl из стиля с учётом цепочки basedOn."""
        num_id = ilvl = None
        seen = set()
        while sid and sid not in seen and sid in self.styles:
            seen.add(sid)
            st = self.styles[sid]
            if num_id is None and st["numId"] is not None:
                num_id = st["numId"]
            if ilvl is None and st["ilvl"] is not None:
                ilvl = st["ilvl"]
            sid = st["basedOn"]
        return num_id, ilvl

    def resolve_abstract(self, aid, depth=0):
        """abstractNum с numStyleLink указывает на список через стиль."""
        link = self.abs_link.get(aid)
        if link and depth < 5:
            if link in self.abs_stylelink and self.abs_stylelink[link] != aid:
                return self.abs_stylelink[link]
            nid, _ = self.style_numpr(link)
            if nid in self.nums:
                return self.resolve_abstract(self.nums[nid][0], depth + 1)
        return aid

    def levels(self, num_id):
        aid, _, lvl_over = self.nums[num_id]
        aid = self.resolve_abstract(aid)
        base = self.abstract.get(aid, {})
        if not lvl_over:
            return aid, base
        merged = dict(base)
        for ilvl, el in lvl_over.items():
            lv = Level()
            if ilvl in base:
                lv.__dict__.update(base[ilvl].__dict__)
            lv.update(el)
            merged[ilvl] = lv
        return aid, merged

    def para_numpr(self, p, view="new"):
        """(numId, ilvl) абзаца или None."""
        ppr = p.find(W + "pPr")
        if view == "old" and ppr is not None and ppr.find(W + "pPrChange/" + W + "pPr") is not None:
            ppr = ppr.find(W + "pPrChange/" + W + "pPr")
        num_id = ilvl = None
        sid = None
        if ppr is not None:
            npr = ppr.find(W + "numPr")
            if npr is not None:
                if view == "old" and npr.find(W + "ins") is not None:
                    npr = None
            if npr is not None:
                num_id = wattr(npr.find(W + "numId"), "val")
                ilvl = wattr(npr.find(W + "ilvl"), "val")
            sid = wattr(ppr.find(W + "pStyle"), "val")
        sid = sid or self.default_style
        s_num, s_ilvl = self.style_numpr(sid)
        if num_id is None:
            num_id = s_num
        if ilvl is None:
            ilvl = s_ilvl
        if num_id is None or num_id == "0" or num_id not in self.nums:
            return None
        if ilvl is None:
            # уровень может быть привязан к стилю внутри abstractNum
            _, lvls = self.levels(num_id)
            ilvl = next((str(k) for k, lv in lvls.items() if sid and lv.pstyle == sid), "0")
        return num_id, int(ilvl)

    def compute(self, paras, view="new"):
        """
        Номера для списка абзацев. Возвращает список той же длины:
        {"text": "5.3.", "key": "5.3", "level": 1, "auto": True} или None.
        Абзацы, которых нет в редакции view, получают None и не считаются.
        """
        counters = {}       # abstractNumId -> {ilvl: значение}
        started = set()     # numId, для которых startOverride уже применён
        out = []
        for p in paras:
            if not exists_in(p, view):
                out.append(None)
                continue
            npr = self.para_numpr(p, view)
            if npr is None:
                out.append(None)
                continue
            num_id, ilvl = npr
            aid, lvls = self.levels(num_id)
            lv = lvls.get(ilvl)
            if lv is None:
                out.append(None)
                continue
            cnt = counters.setdefault(aid, {})
            overrides = self.nums[num_id][1]
            if num_id not in started and overrides:
                started.add(num_id)
                for ol, sv in overrides.items():
                    cnt[ol] = sv - 1
                    for deeper in [k for k in cnt if k > ol]:
                        del cnt[deeper]
            cnt[ilvl] = cnt.get(ilvl, lv.start - 1) + 1
            # сброс более глубоких уровней
            for deeper in [k for k in cnt if k > ilvl]:
                dl = lvls.get(deeper)
                restart = dl.restart if dl else None
                if restart is None or (restart != 0 and ilvl < restart):
                    del cnt[deeper]

            def sub(m):
                k = int(m.group(1)) - 1
                klv = lvls.get(k, Level())
                value = cnt.get(k, klv.start)
                fmt = "decimal" if (lv.legal and klv.fmt not in ("none", "bullet")) else klv.fmt
                return fmt_number(value, fmt)

            if lv.fmt == "bullet":
                out.append({"text": "", "key": "", "level": ilvl, "auto": True, "bullet": True})
                continue
            text = re.sub(r"%(\d)", sub, lv.text if lv.text is not None else f"%{ilvl + 1}.")
            out.append({"text": text, "key": clause_key(text) if re.search(r"\d", text) else "",
                        "level": ilvl, "auto": True})
        return out


def number_paragraphs(pkg_or_parts, paras, view="new"):
    """
    Номер каждого абзаца: автонумерация, иначе номер текстом.
    pkg_or_parts — docxcore.Package или (numbering_xml, styles_xml).
    """
    if hasattr(pkg_or_parts, "read"):
        nb = Numbering(pkg_or_parts.read("word/numbering.xml"), pkg_or_parts.read("word/styles.xml"))
    else:
        nb = Numbering(*pkg_or_parts)
    auto = nb.compute(paras, view)
    out = []
    for p, a in zip(paras, auto):
        if a and a["key"]:
            out.append(a)
            continue
        if not exists_in(p, view):
            out.append(None)
            continue
        m = text_num(para_text(p, view))
        if m:
            out.append({"text": m.group(1), "key": m.group(1).rstrip("."),
                        "level": m.group(1).rstrip(".").count("."), "auto": False})
        else:
            out.append(a if a and a.get("text") else None)  # «а)», «•» — подпункт без ключа
    return out
