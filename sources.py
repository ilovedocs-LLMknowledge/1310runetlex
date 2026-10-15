"""
Источник изменений (протокол / ДС, docx / скан) -> единый список правок.
"""
import re

import config
import edits as E
import gemini
from docxcore import norm

DOCX = (".docx",)


def _norm_gemini(res, source_id):
    src = res.get("source") or {}
    kind = src.get("kind") if src.get("kind") in ("protocol", "amendment") else "amendment"
    number, date = norm(src.get("number", "")), norm(src.get("date", ""))
    if kind == "protocol":
        title = "Протокол согласования разногласий" if "согласован" in src.get("title", "").lower() \
            else "Протокол разногласий"
        label = title + (f" от {date}" if date else "")
    else:
        title = "Дополнительное соглашение"
        label = "ДС" + (f" № {number}" if number else "") + (f" от {date}" if date else "")
    out = []
    for n, e in enumerate(res.get("edits") or [], 1):
        action = e.get("action")
        if action not in E.ACTIONS:
            continue
        target = re.sub(r"[^\d.]", "", e.get("target") or "").strip(".")
        old, new = e.get("old_text") or "", e.get("new_text") or ""
        if action in ("replace_clause", "insert"):
            new = E._strip_leading_num(E._unquote(new), target) if new else new
        out.append({"id": f"{source_id}-{n}", "action": action, "target": target,
                    "targets": [target] if target else [], "anchor": re.sub(r"[^\d.]", "", e.get("anchor") or ""),
                    "old_text": old, "new_text": new, "label": target, "note": e.get("note") or "",
                    "enabled": True})
    meta = {"kind": kind, "title": title, "number": number, "date": date, "label": label}
    return meta, out


def parse_source(data, filename, source_id="s1", use_gemini=True, new_col=None):
    """
    Возвращает {meta, edits, via: local|gemini|regex, columns?, ocr_text?, unparsed?}.
    """
    name = filename.lower()
    if name.endswith(DOCX):
        if E.looks_like_protocol(data):
            res = E.parse_protocol(data, source_id, new_col=new_col)
            res["via"] = "local"
            return res
        text = E.docx_text(data)
        if use_gemini and config.gemini_key():
            meta, edits = _norm_gemini(gemini.extract_from_text(text), source_id)
            local_meta = E.amendment_meta(text)
            meta["number"] = meta["number"] or local_meta["number"]
            meta["date"] = meta["date"] or local_meta["date"]
            return {"meta": meta, "edits": edits, "via": "gemini"}
        edits, unparsed = E.parse_amendment_text(text, source_id)
        return {"meta": E.amendment_meta(text), "edits": edits, "via": "regex", "unparsed": unparsed}
    if gemini.mime_for(name):
        if not use_gemini:
            raise ValueError("Скан можно разобрать только через Gemini.")
        res = gemini.extract_from_file(data, filename)
        meta, edits = _norm_gemini(res, source_id)
        return {"meta": meta, "edits": edits, "via": "gemini", "ocr_text": res.get("ocr_text", "")}
    raise ValueError("Поддерживаются .docx, .pdf, .jpg, .png, .heic, .tiff.")
