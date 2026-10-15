"""
Настройки программы в файле .env рядом с программой.

GEMINI_API_KEY — ключ Gemini (наружу не отдаётся: ни в интерфейс, ни в логи).
GEMINI_MODEL   — модель Gemini.
"""
import os
import re
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
ENV_PATH = APP_DIR / ".env"
DEFAULT_MODEL = "gemini-3.8-flash"

LINE_RE = re.compile(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$")


def read_env(path=ENV_PATH):
    values = {}
    if not path.exists():
        return values
    for line in path.read_text(encoding="utf8").splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        m = LINE_RE.match(line)
        if m:
            v = m.group(2)
            if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
                v = v[1:-1]
            values[m.group(1)] = v
    return values


def write_env(updates, path=ENV_PATH):
    """Обновляет переменные в .env, сохраняя остальные строки. Права — только владельцу."""
    lines = path.read_text(encoding="utf8").splitlines() if path.exists() else []
    done = set()
    for i, line in enumerate(lines):
        m = LINE_RE.match(line)
        if m and m.group(1) in updates:
            lines[i] = f"{m.group(1)}={updates[m.group(1)]}"
            done.add(m.group(1))
    lines += [f"{k}={v}" for k, v in updates.items() if k not in done]
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf8") as f:
        f.write("\n".join(lines) + "\n")
    os.chmod(path, 0o600)


def gemini_key():
    return os.environ.get("GEMINI_API_KEY") or read_env().get("GEMINI_API_KEY", "")


def gemini_model():
    return os.environ.get("GEMINI_MODEL") or read_env().get("GEMINI_MODEL") or DEFAULT_MODEL


def public_settings():
    """То, что можно показать в интерфейсе: без значения ключа."""
    key = gemini_key()
    return {"gemini_key_set": bool(key), "gemini_model": gemini_model()}


def save_settings(key=None, model=None):
    updates = {}
    if key is not None:
        key = key.strip()
        if key and not re.fullmatch(r"[A-Za-z0-9_\-]{20,200}", key):
            raise ValueError("Ключ выглядит некорректно: ожидаются латинские буквы, цифры, «-» и «_».")
        updates["GEMINI_API_KEY"] = key
    if model is not None and model.strip():
        if not re.fullmatch(r"[A-Za-z0-9.\-_]{3,80}", model.strip()):
            raise ValueError("Некорректное имя модели.")
        updates["GEMINI_MODEL"] = model.strip()
    if updates:
        write_env(updates)
    return public_settings()
