"""
Модель правки и разбор протокола разногласий (.docx с таблицей).

Правка: {id, action, target, targets, anchor, old_text, new_text, label, note, enabled}
action: replace_clause | insert | delete_clause | replace_words | renumber
"""
import re

from docxcore import W, Package, exists_in, iter_paragraphs, norm, para_text

ACTIONS = {
    "replace_clause": "изложить в новой редакции",
    "insert": "дополнить пунктом",
    "delete_clause": "исключить пункт",
    "replace_words": "заменить слова",
    "renumber": "изменить номер",
}

ABSENT = {"отсутствует", "-", "—", "–", "нет", "отсутствуют", "пункт отсутствует"}
DELETE = {"исключить", "исключен", "исключён", "исключено", "исключена", "-", "—", "–",
          "пункт исключить", "исключить пункт", "удалить", "исключить полностью"}
KEY_RE = re.compile(r"\d+(?:\.\d+)*")
MONTHS = {"января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6, "июля": 7,
          "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12}


def find_date(text):
    """Первая дата в тексте -> 'ДД.ММ.ГГГГ' или ''."""
    m = re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})\b", text)
    if m:
        y = m.group(3) if len(m.group(3)) == 4 else "20" + m.group(3)
        return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}.{y}"
    m = re.search(r"[«\"]?\s*(\d{1,2})\s*[»\"]?\s+(" + "|".join(MONTHS) + r")\s+(\d{4})", text, re.I)
    if m:
        return f"{int(m.group(1)):02d}.{MONTHS[m.group(2).lower()]:02d}.{m.group(3)}"
    return ""


def doc_date(text):
    """Дата самого документа: даты в ссылке на договор («к Договору … от …») пропускаем."""
    cleaned = re.sub(r"(?:к|по)\s+договор\w*[^\n]{0,120}?\bот\s+(?:[«\"]?\s*\d{1,2}\s*[»\"]?\s+\w+\s+\d{4}|\d{1,2}\.\d{1,2}\.\d{2,4})",
                     " ", text, flags=re.I)
    return find_date(cleaned)


def date_key(d):
    m = re.match(r"(\d\d)\.(\d\d)\.(\d{4})", d or "")
    return (m.group(3), m.group(2), m.group(1)) if m else ("9999",)


def parse_label(label):
    """
    «7.2 (7.2, 7.3 в редакции Стороны-2)» -> (['7.2'], ['7.2', '7.3'], '')
    «Заголовок «Форс-мажор»» -> ([], [], 'Форс-мажор')
    """
    s = norm(label)
    m = re.match(r"заголовок\s*[«\"](.+)[»\"]$", s, re.I)
    if m:
        return [], [], m.group(1)
    head, new_nums = s, []
    m = re.search(r"\((.*?)в\s+редакции[^)]*\)", s, re.I)
    if m:
        head = s[:m.start()]
        new_nums = KEY_RE.findall(m.group(1))
    head = re.sub(r"^(п\.|пп\.|пункт|подпункт|статья|ст\.)\s*", "", head.strip(), flags=re.I)
    return KEY_RE.findall(head), new_nums, ""


def cell_lines(tc):
    lines = []
    for p in tc.iter(W + "p"):
        if not exists_in(p, "new"):
            continue
        t = norm(para_text(p, "new"))
        if t:
            lines.append(t)
    return lines


def table_rows(tbl):
    rows = []
    for tr in tbl.findall(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            span = tc.find(W + "tcPr/" + W + "gridSpan")
            lines = cell_lines(tc)
            n = int(span.get(W + "val", "1")) if span is not None else 1
            cells.extend([lines] + [[] for _ in range(n - 1)])
        rows.append(cells)
    return rows


def detect_columns(header):
    hdr = [norm(" ".join(c)).lower() for c in header]
    red = [i for i, h in enumerate(hdr) if "редакц" in h or "текст" in h or "предлага" in h]
    agreed = [i for i in red if any(w in hdr[i] for w in ("согласован", "итогов", "принят", "окончательн"))]
    num = next((i for i, h in enumerate(hdr) if i not in red and ("пункт" in h or "статья" in h or "п." in h)), None)
    if num is None and red:
        num = red[0] - 1 if red[0] > 0 else None
    old = next((i for i in red if i not in agreed), None)
    return {"headers": [" ".join(c) for c in header], "num": num, "old": old,
            "new": agreed[0] if agreed else (red[-1] if len(red) > 1 else None), "red": red}


def pick_table(pkg):
    best = None
    for tbl in pkg.body.iter(W + "tbl"):
        rows = table_rows(tbl)
        if not rows:
            continue
        for k, row in enumerate(rows[:3]):
            cols = detect_columns(row)
            if cols["old"] is not None and cols["new"] is not None:
                return tbl, rows, k, cols
        width = max(len(r) for r in rows)
        if best is None and 3 <= width <= 5 and len(rows) > 1:
            cols = {"headers": [" ".join(c) for c in rows[0]], "num": 0, "old": 1, "new": width - 1,
                    "red": list(range(1, width))}
            best = (tbl, rows, 0, cols)
    return best


def is_absent(lines):
    return not lines or norm(" ".join(lines)).lower().rstrip(".") in ABSENT


def is_delete(lines):
    return bool(lines) and norm(" ".join(lines)).lower().rstrip(".") in DELETE


def parse_protocol(data, source_id="s1", new_col=None):
    """Протокол .docx -> {meta, columns, edits}."""
    pkg = Package(data)
    found = pick_table(pkg)
    if not found:
        raise ValueError("В протоколе не найдена таблица с редакциями пунктов.")
    tbl, rows, header_idx, cols = found
    if new_col is not None:
        cols["new"] = int(new_col)
    intro = " ".join(norm(para_text(p)) for p in iter_paragraphs(pkg.body)[:40])
    date = doc_date(intro)
    title = "Протокол согласования разногласий" if "согласования разногласий" in intro.lower() else "Протокол разногласий"
    edits, n = [], 0
    for row in rows[header_idx + 1:]:
        cell = lambda i: row[i] if i is not None and i < len(row) else []
        flat = [norm(" ".join(c)) for c in row]
        if not any(flat) or all(re.fullmatch(r"\d*", f) for f in flat):
            continue
        label = norm(" ".join(cell(cols["num"])))
        old, new = cell(cols["old"]), cell(cols["new"])
        if not label and not old and not new:
            continue
        old_nums, new_nums, heading = parse_label(label)
        note = ""
        # «Принять в редакции Стороны-1» в колонке согласованной редакции
        joined_new = norm(" ".join(new)).lower()
        if new and len(joined_new) < 80 and "редакци" in joined_new:
            other = [i for i in cols["red"] if i != cols["new"]]
            hit = None
            for i in other:
                words = [w for w in re.findall(r"\w{4,}", cols["headers"][i].lower()) if w != "редакция"]
                if any(w[:-1] in joined_new for w in words):
                    hit = i
            if hit == cols["old"]:
                continue  # пункт остаётся в прежней редакции
            if hit is not None:
                new = cell(hit)
                note = f"Принята редакция из колонки «{cols['headers'][hit]}»"
        if is_absent(old) and not is_absent(new):
            action = "insert"
            target = (new_nums or old_nums or [""])[0]
            if not target and new:
                m = re.match(r"\s*(\d+(?:\.\d+)+)\.?\s", new[0])
                target = m.group(1) if m else ""
            old_text = ""
        elif is_delete(new) or (is_absent(new) and not is_absent(old)):
            action, target, old_text = "delete_clause", (old_nums or [""])[0], "\n".join(old)
            new = []
        else:
            if norm(" ".join(old)) == norm(" ".join(new)):
                continue
            action, target, old_text = "replace_clause", (old_nums or [""])[0], "\n".join(old)
        n += 1
        edits.append({
            "id": f"{source_id}-{n}", "action": action, "target": target,
            "targets": old_nums if action != "insert" else [], "anchor": "",
            "old_text": old_text, "new_text": "\n".join(new), "label": label or (f"«{heading}»" if heading else ""),
            "note": note or ("Заголовок без номера" if heading else ""), "enabled": True,
        })
    return {
        "meta": {"kind": "protocol", "title": title, "date": date,
                 "label": f"{title} от {date}" if date else title},
        "columns": {"headers": cols["headers"], "old": cols["old"], "new": cols["new"], "red": cols["red"]},
        "edits": edits,
    }


def docx_text(data):
    """Плоский текст .docx (для ДС): абзацы и таблицы по порядку."""
    pkg = Package(data)
    return "\n".join(t for t in (norm(para_text(p)) for p in iter_paragraphs(pkg.body)
                                 if exists_in(p, "new")) if t)


def looks_like_protocol(data):
    try:
        return pick_table(Package(data)) is not None and any(
            "редакц" in h.lower() for h in pick_table(Package(data))[3]["headers"])
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Дополнительное соглашение: разбор типовых формулировок без нейросети
# ---------------------------------------------------------------------------

REF = r"(?:пункт\w*|подпункт\w*|п\.|пп\.|абзац\w*|раздел\w*|стать\w+|ст\.)\s*(\d+(?:\.\d+)*)"
INSTR_RE = re.compile(r"изложить|дополнить|исключить|заменить|считать|признать утратившим", re.I)
FINAL_RE = re.compile(r"остальн\w+ услови|вступает в силу|составлен\w* в \w+ экземпляр|"
                      r"неотъемлем\w+ част\w+|реквизит|подпис", re.I)


def _unquote(text):
    t = text.strip().rstrip(";").strip()
    if t.endswith("."):
        core = t[:-1].rstrip()
        if core.endswith(("»", '"', "”")):
            t = core
    if t[:1] in "«\"“" and t[-1:] in "»\"”":
        t = t[1:-1]
    return t.strip()


def _strip_leading_num(text, key):
    m = re.match(r"\s*(\d+(?:\.\d+)*)\.?\s+", text)
    return text[m.end():] if m and m.group(1) == key else text


def amendment_meta(text):
    head = text[:600]
    m = re.search(r"(дополнительн\w+\s+соглашени\w+)\s*(?:№\s*([\w/.-]+))?", head, re.I)
    number = m.group(2) if m and m.group(2) else ""
    date = doc_date(head)
    label = "ДС" + (f" № {number}" if number else "") + (f" от {date}" if date else "")
    return {"kind": "amendment", "title": "Дополнительное соглашение", "number": number, "date": date,
            "label": label}


def _expand(a, b):
    """«9.1»–«9.3» -> ['9.1', '9.2', '9.3']"""
    pa, pb = a.split("."), b.split(".")
    if len(pa) != len(pb) or pa[:-1] != pb[:-1] or int(pb[-1]) < int(pa[-1]):
        return [a, b]
    return [".".join(pa[:-1] + [str(k)]) for k in range(int(pa[-1]), int(pb[-1]) + 1)]


def parse_amendment_text(text, source_id="s1"):
    """Типовые формулировки ДС -> правки. Нераспознанное — в unparsed."""
    # кавычки «» не нормализуем — по ним видно, где кончается новая редакция
    lines = [ln for ln in (re.sub(r"\s+", " ", x.replace("\u00a0", " ")).strip() for x in text.split("\n")) if ln]
    items, cur, depth, expect = [], None, 0, False
    for ln in lines:
        opens, closes = ln.count("«"), ln.count("»")
        if depth <= 0 and not expect and INSTR_RE.search(ln) and re.search(REF, ln, re.I):
            cur = [ln]
            items.append(cur)
            depth = opens - closes
            expect = depth <= 0 and ln.endswith(":")  # новая редакция — со следующей строки
        elif cur is not None and (depth > 0 or expect):
            cur.append(ln)
            depth += opens - closes
            expect = False
        else:
            cur, depth = None, 0
            if INSTR_RE.search(ln) and not FINAL_RE.search(ln):
                items.append([ln])
    edits, unparsed, n = [], [], 0

    def add(action, target, old="", new="", anchor="", note=""):
        nonlocal n
        n += 1
        edits.append({"id": f"{source_id}-{n}", "action": action, "target": target, "targets": [target] if target else [],
                      "anchor": anchor, "old_text": old, "new_text": new, "label": target,
                      "note": note, "enabled": True})

    for it in items:
        head = re.sub(r"^\d{1,2}\.\s*", "", it[0])
        body_first = head.split(":", 1)[1] if ":" in head else ""
        quoted = "\n".join(([body_first] if body_first.strip() else []) + it[1:])
        refs = re.findall(REF, head.split(":", 1)[0], re.I)
        low = head.lower()
        if "слова" in low and "заменить" in low:
            m = re.search(r"слов\w*\s*[«\"](.+?)[»\"]\s*замен\w+\s*слов\w*\s*[«\"](.+?)[»\"]", head, re.I)
            if m:
                whole = re.search(r"по\s+(?:всему\s+)?тексту", low)
                add("replace_words", "" if whole else (refs[0] if refs else ""), m.group(1), m.group(2))
                continue
        if "изложить" in low and refs:
            add("replace_clause", refs[0], new=_strip_leading_num(_unquote(quoted), refs[0]))
            continue
        if "дополнить" in low and refs:
            m = re.search(r"(?:пункт\w*|подпункт\w*)\s+(\d+(?:\.\d+)*)\s+следующ", head, re.I)
            target = m.group(1) if m else refs[-1]
            m2 = re.search(r"после\s+" + REF, head, re.I)
            add("insert", target, new=_strip_leading_num(_unquote(quoted), target),
                anchor=m2.group(1) if m2 else "")
            continue
        if ("исключить" in low or "утратившим" in low) and refs:
            for r in refs:
                add("delete_clause", r)
            continue
        if "считать" in low and len(refs) >= 2:
            rng = re.findall(r"(\d+(?:\.\d+)*)\s*[-–—]\s*(\d+(?:\.\d+)*)", head)
            olds, news = (_expand(*rng[0]), _expand(*rng[1])) if len(rng) >= 2 else ([refs[0]], [refs[1]])
            if len(olds) == len(news):
                for o, nw in zip(olds, news):
                    add("renumber", o, new=nw)
                continue
        unparsed.append(" ".join(it)[:400])
    return edits, unparsed
