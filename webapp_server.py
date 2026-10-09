"""Skrepka AI Mini App backend."""
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
from telegram import Bot, ReplyKeyboardMarkup, KeyboardButton

import core

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("webapp_server")

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
if not TELEGRAM_TOKEN:
    logger.warning("TELEGRAM_TOKEN is not configured")

ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "*")
bot = Bot(token=TELEGRAM_TOKEN) if TELEGRAM_TOKEN else None

app = FastAPI(title="Skrepka AI Mini App backend")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[ALLOWED_ORIGIN] if ALLOWED_ORIGIN != "*" else ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
core.init_db()

UPLOADS = {}
UPLOAD_TTL = 3600
MAX_UPLOAD = 1024 * 1024 * 1024
MENU_KEYBOARD = ReplyKeyboardMarkup(
    [[KeyboardButton("🏠 Главное меню"), KeyboardButton("💫 Мой тариф")],
    resize_keyboard=True,
    is_persistent=True,
)


def validate_init_data(init_data: str | None):
    if not init_data:
        raise HTTPException(status_code=401, detail="Нет initData")
    try:
        parsed = dict(parse_qsl(init_data, strict_parsing=True))
    except ValueError:
        raise HTTPException(status_code=401, detail="Некорректная initData")
    received_hash = parsed.pop("hash", None)
    if not received_hash:
        raise HTTPException(status_code=401, detail="Нет hash")
    if not TELEGRAM_TOKEN:
        raise HTTPException(status_code=503, detail="Сервис временно недоступен")
    data_check_string = "\n".join(f"{k}={v}" for k, v in sorted(parsed.items()))
    secret_key = hmac.new(b"WebAppData", TELEGRAM_TOKEN.encode(), hashlib.sha256).digest()
    expected_hash = hmac.new(secret_key, data_check_string.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected_hash, received_hash):
        raise HTTPException(status_code=401, detail="Неверная подпись")
    try:
        auth_date = int(parsed.get("auth_date", "0"))
    except ValueError:
        raise HTTPException(status_code=401, detail="Некорректная дата авторизации")
    if auth_date <= 0 or time.time() - auth_date > 86400 or auth_date > time.time() + 60:
        raise HTTPException(status_code=401, detail="initData устарела")
    try:
        user = json.loads(parsed.get("user", "{}"))
    except json.JSONDecodeError:
        raise HTTPException(status_code=401, detail="Некорректные данные пользователя")
    if not user.get("id"):
        raise HTTPException(status_code=401, detail="Нет данных пользователя")
    return user


def cleanup():
    now = time.time()
    expired = [upload_id for upload_id, record in UPLOADS.items()
               if now - record["created_at"] > UPLOAD_TTL]
    for upload_id in expired:
        try:
            os.remove(UPLOADS[upload_id]["path"])
        except OSError:
            pass
        UPLOADS.pop(upload_id, None)


def user_details(user):
    uid = int(user["id"])
    username = user.get("username", "")
    full_name = (user.get("first_name", "") + " " + user.get("last_name", "")).strip()
    core.register_user_if_new(uid, username, full_name)
    info = core.get_user_info(uid)
    if not info:
        info = {"usage": 0, "sub": False, "until": None, "credits": 0}
    active = core.has_active_sub(info)
    return uid, username, info, active


@app.get("/api/status")
def status(x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    _, _, info, active = user_details(user)
    result = {
        "is_subscribed": bool(active),
        "remaining": max(0, core.FREE_LIMIT - info["usage"]),
        "free_limit": core.FREE_LIMIT,
        "audit_credits": info["credits"],
        "subscription_until": info["until"].strftime("%d.%m.%Y") if active else None,
        "price_sub": core.PRICE_SUB,
        "audit_packs": core.AUDIT_PACKS,
    }
    return result


@app.post("/api/upload")
async def upload(file: UploadFile = File(...), x_telegram_initdata: str = Header(None)):
    """Upload is allowed before choosing a document type. Access is checked in /generate.

    This is necessary so a user who has exhausted the free document limit can still
    choose the independently purchasable audit feature.
    """
    cleanup()
    user = validate_init_data(x_telegram_initdata)
    uid = int(user["id"])
    uname = user.get("username", "")
    user_details(user)
    suffix = os.path.splitext(file.filename or "audio")[1] or ".audio"
    fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    size = 0
    try:
        with open(tmp_path, "wb") as out:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                size += len(chunk)
                if size > MAX_UPLOAD:
                    raise HTTPException(status_code=413, detail="Файл слишком большой (максимум 1 ГБ).")
                out.write(chunk)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise
    finally:
        await file.close()
    upload_id = hashlib.sha256(f"{uid}-{tmp_path}-{time.time()}".encode()).hexdigest()[:24]
    UPLOADS[upload_id] = {
        "path": tmp_path,
        "user_id": uid,
        "username": uname,
        "created_at": time.time(),
    }
    return {"upload_id": upload_id, "size_mb": round(size / 1024 / 1024, 1)}


class GenerateRequest(BaseModel):
    upload_id: str
    doc_type: str


@app.post("/api/generate")
async def generate(req: GenerateRequest, x_telegram_initdata: str = Header(None)):
    user = validate_init_data(x_telegram_initdata)
    uid, uname, _, _ = user_details(user)
    record = UPLOADS.get(req.upload_id)
    if not record or record["user_id"] != uid:
        raise HTTPException(status_code=404, detail="Загрузка не найдена. Прикрепите файл ещё раз.")
    if req.doc_type not in core.DOC_NAMES:
        raise HTTPException(status_code=400, detail="Неизвестный тип документа.")

    is_audit = req.doc_type == "audit"
    audit_status = None
    if is_audit:
        audit_status = core.check_audit_access(uid, uname)
        if audit_status == "need_credits":
            raise HTTPException(
                status_code=403,
                detail="Бесплатные обработки закончились, а кредитов на аудит нет. Купите пакет аудитов в Telegram-боте.",
            )
    elif not core.check_access(uid, uname):
        raise HTTPException(
            status_code=403,
            detail="Бесплатный лимит основных функций исчерпан. Оформите подписку Skrepka в Telegram-боте.",
        )

    path = record["path"]
    try:
        try:
            transcription = core.transcribe_audio_file(path)
        except Exception as exc:
            logger.exception("Транскрибация завершилась ошибкой: %s", exc)
            raise HTTPException(status_code=500, detail="Не удалось распознать запись. Попробуйте ещё раз.")
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
        UPLOADS.pop(req.upload_id, None)

    if not transcription or len(transcription.strip()) < 10:
        raise HTTPException(status_code=422, detail="Речь не распознана. Попробуйте запись с более разборчивой речью.")

    try:
        doc_name, clean_text = core.generate_document(req.doc_type, transcription)
    except Exception as exc:
        logger.exception("Генерация документа завершилась ошибкой: %s", exc)
        raise HTTPException(status_code=500, detail="Не удалось составить документ. Попробуйте ещё раз.")

    # If this audit uses a purchased credit, reserve it atomically after successful generation.
    paid_audit = is_audit and audit_status == "ok"
    if paid_audit and not core.try_use_audit_credit(uid):
        raise HTTPException(status_code=409, detail="Кредиты на аудит закончились. Обновите баланс и попробуйте снова.")

    try:
        if not bot:
            raise RuntimeError("Telegram bot is not configured")
        filename = core.safe_filename(doc_name)
        await bot.send_document(chat_id=uid, document=core.build_docx(doc_name, clean_text),
                                filename=f"{filename}.docx", caption="📎 .docx")
        await bot.send_document(chat_id=uid, document=core.build_txt(doc_name, clean_text),
                                filename=f"{filename}.txt", caption="📎 .txt")
    except Exception as exc:
        logger.exception("Не удалось отправить результат в Telegram: %s", exc)
        if paid_audit:
            try:
                core.refund_audit_credit(uid)
            except Exception:
                logger.exception("Не удалось вернуть кредит после ошибки отправки")
        raise HTTPException(status_code=502, detail="Документ подготовлен, но отправить его в Telegram не удалось. Попробуйте позже.")

    # Paid audits do not consume the free quota for core features.
    if not paid_audit:
        core.increment_usage(uid)

    info = core.get_user_info(uid) or {"credits": 0, "usage": 0, "sub": False, "until": None}
    if is_audit:
        finish = (
            "✨ Аудит готов!\n\nДокументы отправлены в чат.\n"
            f"\n🔍 Осталось кредитов на аудит: {info['credits']}"
            "\n\n⚠️ Проверяйте выводы по исходной записи."
        )
    else:
        finish = "✨ Документ готов! Файлы .docx и .txt отправлены в чат.\n\n⚠️ Проверьте цифры, имена и ссылки перед использованием."

    try:
        await bot.send_message(chat_id=uid, text=finish, reply_markup=MENU_KEYBOARD)
    except Exception:
        logger.exception("Не удалось отправить завершающее сообщение с меню")

    return {
        "status": "ok",
        "doc_name": doc_name,
        "audit_credits": info["credits"],
        "is_subscribed": core.has_active_sub(info),
        "remaining": max(0, core.FREE_LIMIT - info["usage"]),
        "message": finish,
    }
