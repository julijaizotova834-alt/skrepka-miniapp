"""
core.py — Skrepka AI
Общая логика для бота и мини-приложения.
"""

import os, re, logging, tempfile
from io import BytesIO
from datetime import datetime

import psycopg2
from groq import Groq
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

logger = logging.getLogger(__name__)

groq_client = Groq(api_key=os.environ.get('GROQ_API_KEY'))
genai.configure(api_key=os.environ.get('GEMINI_API_KEY'))
gemini_model = genai.GenerativeModel('gemini-2.5-flash')

BOT_NAME = 'skrepka_ru'
FREE_LIMIT = 7
GROQ_SIZE_LIMIT_MB = 24
PRICE_SUB = 750
SUB_DAYS = 30
AUDIT_PACKS = [
    {'size': 10, 'price': 150}, {'size': 50, 'price': 550}, {'size': 100, 'price': 900},
]
FREE_USERNAMES = {'black_physicist', 'juli_aleksandrova', 'Aliia85'}

PROMPTS = {
    'narada': """Проанализируй транскрипцию встречи и подготовь краткий отчет на русском языке. Структура: 1. Резюме встречи (3-5 предложений - суть, цель, результат) 2. Основные темы разговора (краткий список) 3. Ключевые решения (краткий список) 4. Договоренности сторон (краткий список) 5. Задачи - каждая в одной строке по шаблону: • [Ответственный] - [Задача] | Срок: [срок] | [Комментарий если есть] 6. Важные цифры, суммы и сроки 7. Открытые вопросы 8. Следующие шаги. Правила: не выдумывай информацию, которой нет в записи; при ошибках распознавания - логически восстанови смысл; НЕ используй таблицы ни в каком виде; только заголовки, списки и абзацы; деловой стиль без лишней «воды».""",

    'audit': """Действуй как РОП (руководитель отдела продаж) и бизнес-тренер по переговорам. Проанализируй транскрипцию разговора менеджера с клиентом (или заказчиком) и сформируй полный разбор звонка на русском языке. ПРАВИЛА ФОРМАТИРОВАНИЯ - выполняй строго: - НЕ используй таблицы ни в каком виде - Каждый пункт - это один короткий абзац или одна строка списка - Только заголовки, маркированные списки и короткие абзацы - Деловой, но прямой стиль — без «воды» и повторов ПРАВИЛА СОДЕРЖАНИЯ: - Не выдумывай информацию, которой нет в транскрипции - Если данные отсутствуют - «Не выявлено» или «Не прозвучало в разговоре» - При ошибках распознавания - логически восстанови смысл без изменения сути - Будь честным в оценке менеджера — не приукрашивай и не смягчай СТРУКТУРА: 1. Краткий итог Один абзац 3-4 предложения: кто с кем общался, цель звонка, чем закончился разговор, общее впечатление. 2. Профиль клиента На основе того, что клиент сказал и как себя вёл: • Тип клиента (горячий / тёплый / холодный) • Уровень осведомлённости о продукте/услуге • Ключевой запрос — что именно клиент хочет решить • Скрытая потребность — что стоит за запросом, если можно определить • Боли клиента — какие проблемы, страхи или неудобства озвучил • Критерии принятия решения — по каким параметрам выбирает (цена, сроки, доверие, опыт) • Бюджет и ожидания по цене — если прозвучало 3. Возражения и сомнения клиента Каждое возражение в формате: • [Возражение] — [как менеджер отработал или не отработал] Если возражений не было — «Возражений не прозвучало.» 4. Коммерческие условия Если обсуждались цены, сроки, объёмы — кратким списком: • [Что]: [сумма / срок / условие] Если не обсуждались — «Коммерческие условия не затрагивались.» 5. Оценка менеджера Оцени работу менеджера по каждому критерию — коротко, конкретно, с примерами из разговора: • Установление контакта — поздоровался, представился, задал тон разговору или нет • Выявление потребности — задавал открытые вопросы или сразу начал продавать • Активное слушание — слышал клиента, уточнял, перефразировал или перебивал и гнул своё • Презентация решения — подал продукт через выгоды клиента или через характеристики • Работа с возражениями — отработал, проигнорировал, спорил или согласился и сдался • Закрытие / следующий шаг — предложил конкретное действие или отпустил клиента без договорённости • Тон и манера — уверенность, дружелюбие, давление, неуверенность, суета 6. Ошибки менеджера Конкретный список с примерами — что менеджер сделал не так: • [Ошибка] — [что было сказано или сделано] — [как надо было] Если критических ошибок нет — отмечай точки роста: что можно было сделать лучше. 7. Сильные стороны менеджера Что менеджер сделал хорошо — конкретно, с примерами из разговора. 8. Этап сделки и прогноз Одним абзацем: • На каком этапе воронки клиент (первичный контакт / выявление потребности / презентация / обсуждение условий / принятие решения / отказ) • Вероятность сделки (высокая / средняя / низкая) — с аргументом почему • Что может помешать закрытию 9. Задачи и следующие шаги Каждая задача — ОДНА строка: • [Ответственный] — [Задача] | Срок: [срок] Если ответственный неизвестен — «Уточнить». Если срок неизвестен — «Не определено». 10. Рекомендации менеджеру 3-5 конкретных рекомендаций что делать с этим клиентом дальше и что улучшить в технике продаж для подобных звонков.""",

    'client_summary': """Действуй как профессиональный ассистент руководителя. Проанализируй транскрипцию встречи или звонка с клиентом (заказчиком, партнёром) и сформируй резюме встречи на русском языке, которое можно отправить клиенту сразу после разговора. ПРАВИЛА ФОРМАТИРОВАНИЯ - выполняй строго: - НЕ используй таблицы ни в каком виде - Только заголовки, короткие списки и абзацы - Вежливый, профессиональный деловой стиль - Документ должен выглядеть как готовое письмо или сообщение для отправки клиенту - НЕ добавляй внутренние пометки, оценки или комментарии — клиент это увидит ПРАВИЛА СОДЕРЖАНИЯ: - Не выдумывай информацию, которой нет в транскрипции - Если что-то не обсуждалось — не упоминай этот пункт вообще - При ошибках распознавания - логически восстанови смысл - НЕ включай внутренние задачи команды — только то, что касается клиента - НЕ включай цены, бюджеты и финансовые условия, если они не были окончательно согласованы обеими сторонами — черновые обсуждения цен не выносить в резюме - Тон: уважительный, конкретный, без подобострастия и канцелярита СТРУКТУРА: Обращение «Добрый день, [имя клиента если прозвучало / коллеги]!» и одно предложение: «Направляю резюме нашей встречи от [дата если известна].» 1. Что обсудили Компактный список ключевых тем разговора — 3-7 пунктов, каждый в одну строку: • [Тема]: [суть в одном предложении] Без внутренних деталей — только то, что важно для клиента. 2. Договорённости Чёткий список того, о чём договорились обе стороны: • [Что согласовано] — [кто делает / что происходит дальше] Только подтверждённые обеими сторонами решения. 3. Следующие шаги Список конкретных действий с нашей стороны и со стороны клиента: • С нашей стороны: [действие] — [срок если обсуждался] • С вашей стороны: [действие] — [срок если обсуждался] Если сроки не обсуждались — не придумывай их. 4. Открытые вопросы Если остались вопросы, требующие уточнения или дополнительного обсуждения — короткий список. Если всё решено — этот раздел не включай. Завершение Одно предложение: «Если я что-то упустил или вы хотите дополнить — буду рад обратной связи.» и подпись: «С уважением, [место для имени]» НЕ добавляй ничего после завершения. Документ должен быть готов к копированию и отправке клиенту.""",

    'ideas': """Действуй как помощник руководителя и стратегический консультант. Проанализируй голосовые заметки или мозговой штурм и преврати их в структурированный план действий на русском языке. Правила: не транскрибируй текст дословно - структурируй и отбирай главное; не выдумывай цифры, бюджеты или сроки которых не было; если данных нет - «Не определено»; НЕ используй таблицы ни в каком виде - только заголовки, маркированные списки и абзацы; управленческий стиль без «воды»; отвечай компактно и без повторов. Структура: 1. Главная суть и цель (3-4 предложения) 2. Ключевые идеи и их ценность (список) 3. Стратегические направления и перспективы 4. Операционные задачи - каждая в одной строке: • [Задача] | Ответственный: [кто] | Приоритет: [высокий/средний/низкий] | Срок: [когда] 5. Срочные вопросы и критические блокеры 6. Что делегировать и кому (роли или специалисты) 7. Риски - финансовые, юридические, технические или операционные (коротко) 8. Ресурсы - люди, инструменты, бюджет 9. Следующие шаги - план на 24 ч / 7 дней / 30 дней 10. Рекомендация - что самое перспективное и что проверить в первую очередь. Если в записи есть идея продукта или бизнеса - дополнительно: целевая аудитория, основная ценность, минимальная версия, монетизация, ключевые функции, сложность запуска.""",
}

DOC_NAMES = {
    'narada': 'Совещание', 'audit': 'Аудит звонка менеджера',
    'client_summary': 'Резюме для клиента', 'ideas': 'Структурированные идеи',
    'transcript': 'Транскрипция',
}

def get_db(): return psycopg2.connect(os.environ.get('DATABASE_URL'))

def init_db():
    conn = get_db(); cur = conn.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS users (
        user_id BIGINT, bot_name TEXT, username TEXT, full_name TEXT,
        registered_at TIMESTAMP DEFAULT NOW(), usage_count INTEGER DEFAULT 0,
        is_subscribed BOOLEAN DEFAULT FALSE, subscription_until TIMESTAMP,
        audit_credits INTEGER DEFAULT 0, last_used_at TIMESTAMP,
        PRIMARY KEY (user_id, bot_name))''')
    cur.execute('''CREATE TABLE IF NOT EXISTS payments (
        id SERIAL PRIMARY KEY, user_id BIGINT, bot_name TEXT,
        telegram_payment_charge_id TEXT UNIQUE, stars_amount INTEGER,
        payload TEXT, paid_at TIMESTAMP DEFAULT NOW())''')
    for col, ct, df in [('audit_credits','INTEGER','0')]:
        try: cur.execute(f"ALTER TABLE users ADD COLUMN {col} {ct} DEFAULT {df}")
        except: conn.rollback()
    try: cur.execute("ALTER TABLE payments ADD COLUMN payload TEXT")
    except: conn.rollback()
    try: cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_pay_charge ON payments(telegram_payment_charge_id)")
    except: conn.rollback()
    conn.commit(); cur.close(); conn.close()

def register_user_if_new(uid, uname, fname):
    conn = get_db(); cur = conn.cursor()
    cur.execute('''INSERT INTO users (user_id,bot_name,username,full_name,registered_at,usage_count,audit_credits,last_used_at)
        VALUES (%s,%s,%s,%s,NOW(),0,0,NOW()) ON CONFLICT (user_id,bot_name) DO UPDATE
        SET username=EXCLUDED.username, full_name=EXCLUDED.full_name, last_used_at=NOW()''', (uid, BOT_NAME, uname, fname))
    conn.commit(); cur.close(); conn.close()

def increment_usage(uid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('UPDATE users SET usage_count=usage_count+1, last_used_at=NOW() WHERE user_id=%s AND bot_name=%s', (uid, BOT_NAME))
    conn.commit(); cur.close(); conn.close()

def get_user_info(uid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('SELECT usage_count,is_subscribed,subscription_until,audit_credits FROM users WHERE user_id=%s AND bot_name=%s', (uid, BOT_NAME))
    r = cur.fetchone(); cur.close(); conn.close()
    if not r: return None
    return {'usage': r[0], 'sub': r[1], 'until': r[2], 'credits': r[3] or 0}

def has_active_sub(info): return info and info['sub'] and info['until'] and info['until'] > datetime.now()

def check_access(uid, uname):
    if uname and uname.lower() in FREE_USERNAMES: return True
    info = get_user_info(uid)
    if not info: return True
    if has_active_sub(info): return True
    return info['usage'] < FREE_LIMIT

def check_audit_access(uid, uname):
    if uname and uname.lower() in FREE_USERNAMES: return 'ok'
    info = get_user_info(uid)
    if not info: return 'free_ok'
    if not has_active_sub(info) and info['usage'] < FREE_LIMIT: return 'free_ok'
    if not has_active_sub(info) and info['credits'] > 0: return 'need_sub_for_credits'
    if not has_active_sub(info): return 'no_sub'
    if info['credits'] > 0: return 'ok'
    return 'need_credits'

def try_use_audit_credit(uid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('UPDATE users SET audit_credits=audit_credits-1 WHERE user_id=%s AND bot_name=%s AND audit_credits>0', (uid, BOT_NAME))
    ok = cur.rowcount > 0; conn.commit(); cur.close(); conn.close(); return ok

def clean_markdown(text):
    if not text: return text
    text = re.sub(r'^#{1,6}\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text); text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text); text = re.sub(r'_(.+?)_', r'\1', text)
    text = re.sub(r'```[a-zA-Z]*\n?', '', text); text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'^\s*[\*\-]\s+', '• ', text, flags=re.MULTILINE)
    return text.strip()

def build_docx(title, body):
    doc = Document(); s = doc.styles['Normal']; s.font.name = 'Calibri'; s.font.size = Pt(11)
    h = doc.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = h.add_run(title.upper()); run.bold = True; run.font.size = Pt(16); run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    doc.add_paragraph()
    for raw in body.split('\n'):
        line = raw.rstrip()
        if not line.strip(): doc.add_paragraph(); continue
        is_h = bool(re.match(r'^\d+\.\s+\S', line))
        ltrs = [c for c in line if c.isalpha()]
        is_caps = len(ltrs) >= 3 and all(c.isupper() for c in ltrs) and len(line) < 80
        if is_h or is_caps:
            p = doc.add_paragraph(); r = p.add_run(line); r.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
        elif line.lstrip().startswith('•'): doc.add_paragraph(line.lstrip()[1:].strip(), style='List Bullet')
        else: doc.add_paragraph(line)
    buf = BytesIO(); doc.save(buf); buf.seek(0); return buf

def build_txt(title, body):
    buf = BytesIO(f"{title.upper()}\n{'='*len(title)}\n\n{body}".encode('utf-8')); buf.seek(0); return buf

def safe_filename(n): return re.sub(r'[^\w\-]', '', n.replace(' ', '_'))

def compress_audio(p):
    import subprocess; o = p + '.compressed.mp3'
    try:
        subprocess.run(['ffmpeg','-y','-i',p,'-ac','1','-ar','16000','-b:a','32k',o], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
        if os.path.exists(o) and os.path.getsize(o) > 0: return o
    except: pass
    return p

def split_audio_chunks(input_path, chunk_minutes=25):
    import subprocess; out_dir = tempfile.mkdtemp()
    pattern = os.path.join(out_dir, "chunk_%03d.mp3")
    subprocess.run(['ffmpeg','-y','-i',input_path,'-f','segment','-segment_time',str(chunk_minutes*60),'-ac','1','-ar','16000','-b:a','32k',pattern], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=1800)
    return sorted(os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.startswith("chunk_"))

def transcribe_audio_file(path):
    size_mb = os.path.getsize(path) / (1024*1024)
    chunk_paths = [path]; cleanup = []
    if size_mb > GROQ_SIZE_LIMIT_MB:
        try: chunk_paths = split_audio_chunks(path); cleanup = chunk_paths
        except:
            c = compress_audio(path); chunk_paths = [c]
            if c != path: cleanup = [c]
    parts = []
    try:
        for p in chunk_paths:
            with open(p, 'rb') as f:
                t = groq_client.audio.transcriptions.create(file=(os.path.basename(p), f.read()), model='whisper-large-v3-turbo', language='ru')
            parts.append(t.text)
    finally:
        for p in cleanup:
            try: os.remove(p)
            except: pass
    return "\n\n".join(parts)

def generate_document(doc_type, transcription):
    doc_name = DOC_NAMES[doc_type]
    if doc_type == 'transcript': return doc_name, transcription.strip()
    resp = gemini_model.generate_content(f"{PROMPTS[doc_type]}\n\nВот транскрипция записи:\n\n{transcription}")
    return doc_name, clean_markdown(resp.text)
