"""
Gemini: распознавание сканов и разбор свободного текста допсоглашений
в список правок (модель правки — edits.py).

Только стандартная библиотека: urllib. Ключ — из .env (config.py),
передаётся в заголовке x-goog-api-key и никуда не пишется.
"""
import base64
import json
import socket
import urllib.error
import urllib.request

import config

ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
TIMEOUT = 180
MAX_INLINE = 19 * 1024 * 1024

MIME = {".pdf": "application/pdf", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
        ".webp": "image/webp", ".heic": "image/heic", ".heif": "image/heif", ".tif": "image/tiff",
        ".tiff": "image/tiff"}


class GeminiError(Exception):
    pass


EDIT_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "source": {
            "type": "OBJECT",
            "properties": {
                "kind": {"type": "STRING", "enum": ["protocol", "amendment", "other"],
                         "description": "protocol — протокол разногласий / согласования разногласий; "
                                        "amendment — дополнительное соглашение"},
                "number": {"type": "STRING", "description": "Номер документа, если есть, иначе пусто"},
                "date": {"type": "STRING", "description": "Дата документа ДД.ММ.ГГГГ, если есть, иначе пусто"},
                "title": {"type": "STRING", "description": "Заголовок документа"},
            },
            "required": ["kind", "number", "date", "title"],
        },
        "ocr_text": {"type": "STRING",
                     "description": "Полный распознанный текст документа (для сканов); для текстового входа — пусто"},
        "edits": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "action": {"type": "STRING",
                               "enum": ["replace_clause", "insert", "delete_clause", "replace_words", "renumber"]},
                    "target": {"type": "STRING", "description": "Номер пункта: '5.3', '7'. Пусто — если номера нет"},
                    "anchor": {"type": "STRING",
                               "description": "Для insert: номер пункта, после которого вставить, если указан"},
                    "old_text": {"type": "STRING"},
                    "new_text": {"type": "STRING"},
                    "note": {"type": "STRING", "description": "Сомнения, неразборчивые места"},
                },
                "required": ["action", "target", "anchor", "old_text", "new_text", "note"],
            },
        },
    },
    "required": ["source", "ocr_text", "edits"],
}

PROMPT = """Ты — юрист, который вносит в договор изменения из документа-источника. Источник — протокол разногласий (протокол согласования разногласий) или дополнительное соглашение к договору. Твоя задача — точно перечислить изменения договора, которые из него следуют. Ничего не придумывай и не улучшай.

Правила:
1. Новый текст пунктов переписывай ДОСЛОВНО, символ в символ, без пересказа и без исправления опечаток. Абзацы внутри одного пункта разделяй символом перевода строки.
2. Номер пункта (target) пиши так, как в источнике: «5.3», «7», «6.18.1» — без слов «пункт», «п.».
3. В new_text не включай номер пункта, если он стоит отдельно от текста (например, «5.3. Текст» → new_text = «Текст», target = «5.3»). Внешние кавычки «…», в которые заключена новая редакция, не включай.
4. Действия:
   - replace_clause — «пункт X изложить в следующей редакции»; строка таблицы протокола с отличающимися редакциями. old_text — прежняя редакция, если она приведена (в протоколе — колонка редакции стороны, предложившей договор), иначе пусто.
   - insert — «дополнить пунктом X следующего содержания»; в протоколе «Отсутствует» в колонке прежней редакции. target — номер нового пункта; anchor — после какого пункта вставить, если это прямо сказано.
   - delete_clause — «пункт X исключить»; в протоколе «Исключить» в колонке новой редакции. old_text — текст пункта, если приведён.
   - replace_words — «в пункте X слова «A» заменить словами «B»»: old_text = A, new_text = B. Если замена «по всему тексту договора» — target пустой.
   - renumber — «пункты X считать пунктами Y»: по одной записи на пункт, target = старый номер, new_text = новый номер.
5. Протокол разногласий: какая редакция итоговая — смотри по заголовкам колонок. Если есть колонка «согласованная / итоговая редакция» — бери её. Иначе итоговой считается редакция стороны, которая направила протокол (обычно последняя колонка). Если в ячейке написано «принять в редакции Стороны N» — возьми текст из колонки этой стороны; если это редакция договора — правки нет.
6. Пункты «настоящее соглашение вступает в силу…», «остальные условия договора остаются неизменными», реквизиты и подписи — это НЕ изменения договора, пропускай их.
7. Неразборчивые на скане места помечай в note и в тексте ставь […].
8. Если документ — скан, в ocr_text верни весь распознанный текст документа; иначе ocr_text пустой.
"""


def mime_for(filename):
    name = filename.lower()
    for ext, m in MIME.items():
        if name.endswith(ext):
            return m
    return None


def call(parts, key=None, model=None, timeout=TIMEOUT):
    key = key or config.gemini_key()
    if not key:
        raise GeminiError("Не задан ключ Gemini. Добавьте его в .env (GEMINI_API_KEY) или в ⚙ Настройках.")
    model = model or config.gemini_model()
    body = {
        "system_instruction": {"parts": [{"text": PROMPT}]},
        "contents": [{"role": "user", "parts": parts}],
        "generation_config": {"response_mime_type": "application/json", "response_schema": EDIT_SCHEMA,
                              "temperature": 0},
    }
    req = urllib.request.Request(ENDPOINT.format(model=model), data=json.dumps(body).encode("utf8"),
                                 headers={"Content-Type": "application/json", "x-goog-api-key": key},
                                 method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf8"))
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read().decode("utf8")).get("error", {}).get("message", "")
        except Exception:
            msg = ""
        msg = msg.replace(key, "***")
        if e.code in (401, 403):
            raise GeminiError(f"Gemini отклонил ключ ({e.code}). Проверьте ключ в .env. {msg}")
        if e.code == 429:
            raise GeminiError("Превышена квота Gemini, попробуйте позже. " + msg)
        if e.code == 404:
            raise GeminiError(f"Модель «{model}» не найдена. Укажите другую в ⚙ Настройках. {msg}")
        raise GeminiError(f"Ошибка Gemini {e.code}: {msg}")
    except (urllib.error.URLError, socket.timeout, TimeoutError) as e:
        raise GeminiError(f"Нет связи с Gemini: {getattr(e, 'reason', e)}")
    cands = data.get("candidates") or []
    if not cands:
        reason = (data.get("promptFeedback") or {}).get("blockReason", "пустой ответ")
        raise GeminiError(f"Gemini не вернул ответ: {reason}")
    text = "".join(p.get("text", "") for p in cands[0].get("content", {}).get("parts", []))
    try:
        return json.loads(text)
    except ValueError:
        raise GeminiError("Gemini вернул ответ не в формате JSON — попробуйте ещё раз.")


def extract_from_file(data, filename):
    mime = mime_for(filename)
    if not mime:
        raise GeminiError("Этот формат файла не поддерживается: нужен PDF, JPG, PNG, HEIC или TIFF.")
    if len(data) > MAX_INLINE:
        raise GeminiError("Файл больше 19 МБ — уменьшите разрешение скана или разделите его.")
    parts = [{"inline_data": {"mime_type": mime, "data": base64.b64encode(data).decode("ascii")}},
             {"text": "Это скан документа. Распознай его и перечисли изменения договора."}]
    return call(parts)


def extract_from_text(text):
    return call([{"text": "Текст документа:\n\n" + text}])


def ping():
    """Пробный запрос без данных договора: проверка ключа и модели."""
    res = call([{"text": "Дополнительное соглашение № 0. 1. Пункт 1.1 Договора изложить в следующей "
                         "редакции: «1.1. Проверка связи.»"}], timeout=60)
    return bool(res.get("edits"))
