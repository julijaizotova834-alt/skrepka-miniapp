"""
webapp_server.py (RU)
Backend для Telegram Mini App "Brief AI" (skrepka-bot).
"""

import os
import hmac
import hashlib
import json
import tempfile
import logging
import time
from urllib.parse import parse_qsl

from fastapi import FastAPI, UploadFile, File, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import telegram

import core

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("webapp_server")

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "*")

bot = telegram.Bot(token=TELEGRAM_TOKEN)

app = FastAPI(title="Brief AI Mini App backend (RU)")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[ALLOWED_ORIGIN] if ALLOWED_ORIGIN != "*" else ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

core.init_db()

UPLOADS = {}
UPLOAD_TTL_SECONDS = 60 * 60
MAX_UPLOAD_BYTES = 1024 * 1024 * 1024


def validate_init_data(init_data: str) -> dict:
    if not init_data:
        raise HTTPException(status_code=401, detail="Нет initData")
    parsed = dict(parse_qsl(init_data, strict_parsing=True))
    received_hash = parsed.pop("hash", None)
    if not received_hash:
        raise HTTPException(status_code=401, detail="Нет hash в initData")
    data_check_string = "\n".join(f"{k}={v}" for k, v in sorted(parsed.items()))
    secret_key = hmac.new(b"WebAppData", TELEGRAM_TOKEN.encode(), hashlib.sha256).digest()
    computed_hash = hmac.new(secret_key, data_check_string.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(computed_hash, received_hash):
        raise HTTPException(status_code=401, detail="Неверная подпись initData")
    auth_date = int(parsed.get("auth_date", "0"))
    if time.time() - auth_date > 86400:
        raise HTTPException(status_code=401, detail="initData устарела, переоткройте мини-приложение")
    user = json.loads(parsed.get("user", "{}"))
    if not user.get("id"):
        raise HTTPException(status_code=401, detail="Нет данных пользователя")
    return user


def cleanup_old_uploads():
    now = time.time()
    expired = [uid for uid, v in UPLOADS.items() if now - v["created_at"] > UPLOAD_TTL_SECONDS]
    for uid in expired:
        try:
            os.remove(UPLOADS[uid]["path"])
        except OSError:
            pass
        UPLOADS.pop(uid, None)


@app.get("/api/status")
def status(x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    user_id = user["id"]
    username = user.get("username", "")
    full_name = (user.get("first_name", "") + " " + user.get("last_name", "")).strip()
    core.register_user_if_new(user_id, username, full_name)
    info = core.get_user_info(user_id)
    if not info:
        return {"is_subscribed": False, "remaining": core.FREE_LIMIT, "free_limit": core.FREE_LIMIT, "audit_credits": 0}
    if info['is_subscribed'] and info['until'] and info['until'].timestamp() > time.time():
        return {
            "is_subscribed": True,
            "tier": info['tier'] or 'basic',
            "subscription_until": info['until'].strftime("%d.%m.%Y"),
            "audit_credits": info['audit_credits'],
        }
    return {
        "is_subscribed": False,
        "remaining": max(0, core.FREE_LIMIT - info['usage_count']),
        "free_limit": core.FREE_LIMIT,
        "audit_credits": info['audit_credits'],
    }


@app.post("/api/upload")
async def upload(file: UploadFile = File(...), x_telegram_initdata: str = Header(None)):
    cleanup_old_uploads()
    user = validate_init_data(x_telegram_initdata)
    user_id = user["id"]
    username = user.get("username", "")
    if not core.check_access(user_id, username):
        raise HTTPException(status_code=403, detail="Лимит бесплатных документов исчерпан. Оформите подписку командой /subscribe в боте.")
    suffix = os.path.splitext(file.filename or "audio")[1] or ".audio"
    fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    size = 0
    with open(tmp_path, "wb") as out:
        while True:
            chunk = await file.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                out.close()
                os.remove(tmp_path)
                raise HTTPException(status_code=413, detail="Файл слишком большой.")
            out.write(chunk)
    upload_id = hashlib.sha256(f"{user_id}-{tmp_path}-{time.time()}".encode()).hexdigest()[:24]
    UPLOADS[upload_id] = {"path": tmp_path, "user_id": user_id, "username": username, "created_at": time.time()}
    logger.info(f"Файл загружен: user={user_id}, size={size/1024/1024:.1f} МБ, upload_id={upload_id}")
    return {"upload_id": upload_id, "size_mb": round(size / 1024 / 1024, 1)}


class GenerateRequest(BaseModel):
    upload_id: str
    doc_type: str


@app.post("/api/generate")
async def generate(req: GenerateRequest, x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    user_id = user["id"]
    username = user.get("username", "")

    record = UPLOADS.get(req.upload_id)
    if not record or record["user_id"] != user_id:
        raise HTTPException(status_code=404, detail="Загрузка не найдена. Прикрепите файл ещё раз.")

    if req.doc_type not in core.DOC_NAMES:
        raise HTTPException(status_code=400, detail="Неизвестный тип документа.")

    if not core.check_access(user_id, username):
        raise HTTPException(status_code=403, detail="Лимит бесплатных документов исчерпан.")

    # проверка доступа к аудиту
    if req.doc_type == 'audit':
        audit_status = core.check_audit_access(user_id, username)
        if audit_status == 'no_sub':
            raise HTTPException(status_code=403, detail="Лимит бесплатных документов исчерпан. Оформите подписку: /subscribe")
        if audit_status == 'need_pro':
            raise HTTPException(status_code=403, detail="Аудит звонков доступен на тарифе Про или при покупке пакета аудитов. Оформить: /subscribe")
        if audit_status == 'no_credits':
            raise HTTPException(status_code=403, detail="Аудиты закончились. Докупите пакет: /subscribe")

    path = record["path"]
    try:
        transcription = core.transcribe_audio_file(path)
    except Exception as e:
        logger.error(f"Ошибка транскрибации: {e}")
        raise HTTPException(status_code=500, detail="Не удалось распознать запись.")
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
        UPLOADS.pop(req.upload_id, None)

    if not transcription or len(transcription.strip()) < 10:
        raise HTTPException(status_code=422, detail="Речь в записи не распознана.")

    try:
        doc_name, clean_text = core.generate_document(req.doc_type, transcription)
    except Exception as e:
        logger.error(f"Ошибка генерации документа: {e}")
        raise HTTPException(status_code=500, detail="Не удалось составить документ.")

    fname = core.safe_filename(doc_name)
    docx_buf = core.build_docx(doc_name, clean_text)
    txt_buf = core.build_txt(doc_name, clean_text)

    await bot.send_document(chat_id=user_id, document=docx_buf, filename=f"{fname}.docx", caption="📎 .docx")
    await bot.send_document(chat_id=user_id, document=txt_buf, filename=f"{fname}.txt", caption="📎 .txt")

    # списание аудита
    if req.doc_type == 'audit':
        a_status = core.check_audit_access(user_id, username)
        if a_status == 'ok':
            core.use_audit_credit(user_id)

    core.increment_usage(user_id)

    finish_text = "✨ Документ сформирован и отправлен.\n\n⚠️ Проверьте цифры, имена и ссылки перед использованием."
    if req.doc_type == 'audit':
        info = core.get_user_info(user_id)
        if info:
            finish_text += f"\n\n🔍 Осталось аудитов: {info['audit_credits']}"

    await bot.send_message(chat_id=user_id, text=finish_text)

    result = {"status": "ok", "doc_name": doc_name}
    if req.doc_type == 'audit':
        info = core.get_user_info(user_id)
        result["audit_credits"] = info['audit_credits'] if info else 0
    return result
