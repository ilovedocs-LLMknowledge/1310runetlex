"""
Общие функции для работы с XML документа Word (.docx).

ElementTree при сохранении теряет префиксы и «лишние» объявления пространств
имён, без которых Word считает файл повреждённым (mc:Ignorable). Поэтому
load_part()/dump_part() регистрируют все префиксы и возвращают корневому тегу
исходные объявления.
"""
import io
import re
import xml.etree.ElementTree as ET
import zipfile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"

XMLNS_RE = re.compile(r'xmlns:([\w.-]+)="([^"]+)"')
ROOT_TAG_RE = re.compile(r"<(?!\?)(?!!)[^>]+>", re.S)


def register_namespaces(xml_text):
    for prefix, uri in XMLNS_RE.findall(xml_text):
        if re.match(r"ns\d+$", prefix):
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass


def load_part(data):
    """bytes части пакета -> (корневой элемент, исходный корневой тег)."""
    text = data.decode("utf8")
    register_namespaces(text)
    m = ROOT_TAG_RE.search(text)
    return ET.fromstring(data), (m.group(0) if m else None)


def dump_part(root, orig_root_tag):
    body = ET.tostring(root, encoding="unicode")
    if orig_root_tag:
        m = ROOT_TAG_RE.match(body)
        new_tag = m.group(0)
        have = dict(XMLNS_RE.findall(orig_root_tag))
        extra = [f'xmlns:{p}="{u}"' for p, u in XMLNS_RE.findall(new_tag) if p not in have]
        tag = orig_root_tag
        if extra:
            close = "/>" if tag.endswith("/>") else ">"
            tag = tag[: -len(close)] + " " + " ".join(extra) + close
        body = tag + body[m.end():]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body).encode("utf8")


class Package:
    """Пакет .docx: части в памяти, document.xml — разобранным деревом."""

    def __init__(self, data):
        self.zip = zipfile.ZipFile(io.BytesIO(data))
        self.names = self.zip.namelist()
        self.root, self.root_tag = load_part(self.zip.read("word/document.xml"))
        self.body = self.root.find(W + "body")

    def read(self, name):
        return self.zip.read(name) if name in self.names else None

    def save(self):
        out = io.BytesIO()
        doc = dump_part(self.root, self.root_tag)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in self.zip.infolist():
                data = doc if info.filename == "word/document.xml" else self.zip.read(info)
                dst.writestr(info, data)
        return out.getvalue()


# ---------------------------------------------------------------------------
# Абзацы
# ---------------------------------------------------------------------------

# Контейнеры, внутри которых лежат абзацы основного текста
_DESCEND = {W + t for t in ("tbl", "tr", "tc", "sdt", "sdtContent", "customXml",
                            "ins", "del", "moveTo", "moveFrom", "smartTag")}


def iter_paragraphs(body):
    """Абзацы основного текста по порядку (с таблицами, без надписей и Fallback)."""
    out = []

    def walk(el):
        for ch in el:
            if ch.tag == W + "p":
                out.append(ch)
            elif ch.tag in _DESCEND:
                walk(ch)
    walk(body)
    return out


def para_mark(p):
    """Состояние знака абзаца: 'ins', 'del', 'insdel' или None."""
    rpr = p.find(W + "pPr/" + W + "rPr")
    if rpr is None:
        return None
    ins = rpr.find(W + "ins") is not None
    dl = rpr.find(W + "del") is not None
    return "insdel" if ins and dl else "ins" if ins else "del" if dl else None


def exists_in(p, view):
    """
    Есть ли абзац как самостоятельный в редакции view:
    'old' — до правок (все отклонены), 'new' — после (все приняты).
    Абзац со вставленным (old) / удалённым (new) знаком сливается со следующим.
    """
    mark = para_mark(p)
    if view == "old":
        return mark not in ("ins", "insdel")
    return mark not in ("del", "insdel")


_SKIP_TEXT = {W + "pPr", W + "rPr", W + "instrText", W + "delInstrText",
              W + "txbxContent", W + "drawing", W + "pict", MC_FALLBACK}


def para_text(p, view="new"):
    """Текст абзаца в редакции view ('old' / 'new'). Табуляция и разрыв -> пробел."""
    out = []

    def walk(el, ins, dl):
        tag = el.tag
        if tag in _SKIP_TEXT:
            return
        if tag in (W + "ins", W + "moveTo"):
            ins = True
        elif tag in (W + "del", W + "moveFrom"):
            dl = True
        visible = (not dl) if view == "new" else (not ins)
        if visible:
            if tag in (W + "t", W + "delText"):
                out.append(el.text or "")
            elif tag in (W + "tab", W + "br", W + "cr"):
                out.append(" ")
            elif tag == W + "noBreakHyphen":
                out.append("-")
        for ch in el:
            walk(ch, ins, dl)
    walk(p, False, False)
    return "".join(out)


def has_revisions(body):
    for tag in ("ins", "del", "moveFrom", "moveTo"):
        if body.find(".//" + W + tag) is not None:
            return True
    return False


def wattr(el, name, default=None):
    return el.get(W + name, default) if el is not None else default


# ---------------------------------------------------------------------------
# Нормализация текста для сравнения
# ---------------------------------------------------------------------------

_QUOTES = str.maketrans({"«": '"', "»": '"', "“": '"', "”": '"', "„": '"', "‟": '"',
                         "’": "'", "‘": "'", "–": "-", "—": "-", "−": "-", "‑": "-",
                         " ": " ", " ": " ", "\t": " "})


def norm(s):
    return re.sub(r"\s+", " ", (s or "").translate(_QUOTES)).strip()


def norm_cmp(s):
    """Строже: без регистра и без пробелов перед знаками препинания."""
    s = norm(s).lower().replace("ё", "е")
    return re.sub(r"\s+([,.;:!?)])", r"\1", s)
