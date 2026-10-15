"""
История прогонов: папка Архив/<дата>_<договор>/ с входными файлами,
результатом, отчётом и meta.json. Список истории — сканирование meta.json.
"""
import datetime as dt
import json
import os
import re
import subprocess
import unicodedata
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
# ARCHIVE_DIR — другой путь к архиву (для тестов)
ARCHIVE = Path(os.environ["ARCHIVE_DIR"]) if os.environ.get("ARCHIVE_DIR") else APP_DIR / "Архив"
KINDS = {"protocol": "Протокол разногласий", "final": "Итоговая редакция"}


def safe_name(name, limit=60):
    name = unicodedata.normalize("NFC", name or "")
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name).strip(" .")
    return (name[:limit].rstrip(" .") or "документ")


def unique(folder, name):
    path = folder / name
    stem, suffix = path.stem, path.suffix
    k = 2
    while path.exists():
        path = folder / f"{stem} ({k}){suffix}"
        k += 1
    return path


def save_run(kind, contract_name, inputs, outputs, report_md=None, extra=None, root=None):
    """
    inputs / outputs: [(имя файла, bytes)]. Возвращает meta (с id — именем папки).
    """
    root = root or ARCHIVE
    now = dt.datetime.now()
    run_dir = unique(root, f"{now:%Y-%m-%d_%H%M%S}_{safe_name(contract_name, 50)}")
    (run_dir / "вход").mkdir(parents=True)
    (run_dir / "результат").mkdir()
    saved_in, saved_out = [], []
    for sub, files, acc in (("вход", inputs, saved_in), ("результат", outputs, saved_out)):
        for name, data in files:
            path = unique(run_dir / sub, safe_name(name, 120))
            path.write_bytes(data if isinstance(data, bytes) else data.encode("utf8"))
            acc.append(f"{sub}/{path.name}")
    if report_md:
        (run_dir / "отчёт.md").write_text(report_md, encoding="utf8")
    meta = {"id": run_dir.name, "created": now.isoformat(timespec="seconds"), "kind": kind,
            "contract": contract_name, "inputs": saved_in, "outputs": saved_out,
            "report": "отчёт.md" if report_md else None}
    meta.update(extra or {})
    (run_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf8")
    return meta


def list_runs(q="", kind="", root=None):
    root = root or ARCHIVE
    if not root.exists():
        return []
    q = unicodedata.normalize("NFC", q or "").lower().strip()
    runs = []
    for meta_path in root.glob("*/meta.json"):
        try:
            meta = json.loads(meta_path.read_text(encoding="utf8"))
        except (OSError, ValueError):
            continue
        meta["id"] = meta_path.parent.name
        if kind and meta.get("kind") != kind:
            continue
        hay = " ".join([meta.get("contract", ""), meta["id"]] +
                       [s.get("label", "") for s in meta.get("sources", [])]).lower()
        if q and q not in unicodedata.normalize("NFC", hay):
            continue
        meta["kind_title"] = KINDS.get(meta.get("kind"), meta.get("kind", ""))
        runs.append(meta)
    runs.sort(key=lambda m: m.get("created", ""), reverse=True)
    return runs


def resolve(run_id, rel=None, root=None):
    """Путь внутри архива по id прогона; всё, что ведёт наружу, отклоняется."""
    root = (root or ARCHIVE).resolve()
    if not run_id or "/" in run_id or "\\" in run_id or run_id in (".", ".."):
        raise ValueError("Неверный идентификатор прогона.")
    run_dir = (root / run_id).resolve()
    if run_dir.parent != root or not run_dir.is_dir():
        raise ValueError("Прогон не найден в архиве.")
    if rel is None:
        return run_dir
    path = (run_dir / rel).resolve()
    if run_dir not in path.parents or not path.exists():
        raise ValueError("Файл не найден в папке прогона.")
    return path


def reveal(path):
    """Показать в Finder (папку — выделить, файл — выделить в его папке)."""
    subprocess.run(["open", "-R", str(path)], check=False)


def open_file(path):
    subprocess.run(["open", str(path)], check=False)


STATUS_RU = {"exact": "применена", "by_text": "применена (найдено по тексту)", "not_found": "НЕ НАЙДЕНА",
             "skipped": "пропущена", "off": "отключена", "error": "ОШИБКА"}


ACTIONS_RU = {"replace_clause": "изложить в новой редакции", "insert": "дополнить пунктом",
              "delete_clause": "исключить пункт", "replace_words": "заменить слова", "renumber": "изменить номер"}


def report_md(contract_name, sources, report):
    lines = [f"# Итоговая редакция: {contract_name}", "",
             f"Собрано: {dt.datetime.now():%d.%m.%Y %H:%M}", ""]
    for src, rep in zip(sources, report["sources"]):
        lines += [f"## {src.get('label')}", f"Файл: {src.get('filename', '')}", ""]
        by_id = {r.get("id"): r for r in rep["results"]}
        for e in src.get("edits", []):
            r = by_id.get(e.get("id"), {})
            st = STATUS_RU.get(r.get("status"), r.get("status", ""))
            where = f"п. {e.get('target')}" if e.get("target") else (e.get("label") or "без номера")
            action = ACTIONS_RU.get(e.get("action"), e.get("action"))
            lines.append(f"- {where} — {action}: **{st}**"
                         + (f" ({r['note']})" if r.get("note") else ""))
        lines.append("")
    if report.get("warnings"):
        lines += ["## Предупреждения", ""] + [f"- {w}" for w in report["warnings"]] + [""]
    return "\n".join(lines)
