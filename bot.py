import os
import io
import json
import sqlite3
from datetime import datetime
from contextlib import asynccontextmanager
from PIL import Image
from fastapi import FastAPI
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, BufferedInputFile, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery
from aiogram.filters import CommandStart, Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application
from google import genai
from docx import Document
from docx.shared import Inches as DocxInches, Pt as DocxPt, RGBColor as DocxRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "PUT_YOUR_BOT_TOKEN_HERE")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "PUT_YOUR_GEMINI_KEY_HERE")
ADMIN_TELEGRAM_ID = int(os.getenv("ADMIN_TELEGRAM_ID", "123456789"))

WEBHOOK_HOST = os.getenv("RENDER_EXTERNAL_URL", "https://your-app-name.onrender.com")
WEBHOOK_PATH = f"/webhook/{TELEGRAM_BOT_TOKEN}"
WEBHOOK_URL = f"{WEBHOOK_HOST}{WEBHOOK_PATH}"

bot = Bot(token=TELEGRAM_BOT_TOKEN)
dp = Dispatcher()
client = genai.Client(api_key=GEMINI_API_KEY)

# ================= دورة حياة FastAPI (Lifespan) المتوافقة مع أحدث الإصدارات =================
@asynccontextmanager
async def lifespan(app: FastAPI):
    # ما يتم تنفيذه عند تشغيل السيرفر وتفعيل الويب هوك
    await bot.set_webhook(url=WEBHOOK_URL, drop_pending_updates=True)
    yield
    # ما يتم تنفيذه عند إيقاف السيرفر وحذف الويب هوك
    await bot.delete_webhook()

app = FastAPI(lifespan=lifespan)

# ================= حالات المحادثة (FSM) =================
class TeacherStates(StatesGroup):
    choosing_service = State()
    choosing_subject = State()
    choosing_grade = State()
    waiting_for_input = State()
    waiting_for_pdf = State()
    choosing_saved_book = State()
    admin_granting = State()

# ================= قاعدة البيانات (SQLite) =================
def init_db():
    conn = sqlite3.connect("teachers_bot.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            requests_count INTEGER DEFAULT 0,
            last_date TEXT,
            is_premium INTEGER DEFAULT 0
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS saved_books (
            book_id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject TEXT,
            grade TEXT,
            book_title TEXT,
            gemini_file_uri TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()

def check_and_update_quota(user_id: int) -> bool:
    if user_id == ADMIN_TELEGRAM_ID:
        return True
        
    conn = sqlite3.connect("teachers_bot.db")
    cursor = conn.cursor()
    cursor.execute("SELECT requests_count, last_date, is_premium FROM users WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    today = datetime.now().strftime("%Y-%m-%d")
    
    if row is None:
        cursor.execute("INSERT INTO users (user_id, requests_count, last_date, is_premium) VALUES (?, 1, ?, 0)", (user_id, today))
        conn.commit()
        conn.close()
        return True
        
    requests_count, last_date, is_premium = row
    if is_premium == 1:
        conn.close()
        return True
    if last_date != today:
        cursor.execute("UPDATE users SET requests_count = 1, last_date = ? WHERE user_id = ?", (today, user_id))
        conn.commit()
        conn.close()
        return True
    if requests_count < 3:
        cursor.execute("UPDATE users SET requests_count = requests_count + 1 WHERE user_id = ?", (user_id,))
        conn.commit()
        conn.close()
        return True
        
    conn.close()
    return False

def save_book_to_db(subject, grade, title, uri):
    conn = sqlite3.connect("teachers_bot.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO saved_books (subject, grade, book_title, gemini_file_uri) VALUES (?, ?, ?, ?)", 
                   (subject, grade, title, uri))
    conn.commit()
    conn.close()

def get_saved_books(subject, grade):
    conn = sqlite3.connect("teachers_bot.db")
    cursor = conn.cursor()
    cursor.execute("SELECT book_id, book_title, gemini_file_uri FROM saved_books WHERE subject = ? AND grade = ?", (subject, grade))
    books = cursor.fetchall()
    conn.close()
    return books

# ================= دالة تصدير ملف Word (.docx) =================
def create_arabic_word_doc(title: str, content: str) -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = DocxInches(0.8)
        section.bottom_margin = DocxInches(0.8)
        section.left_margin = DocxInches(0.8)
        section.right_margin = DocxInches(0.8)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(title)
    title_run.font.size = DocxPt(18)
    title_run.font.bold = True
    title_run.font.color.rgb = DocxRGB(0, 51, 102)

    doc.add_paragraph()

    for line in content.split("\n"):
        if line.strip():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(line)
            run.font.size = DocxPt(12)
            run.font.name = "Arial"

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ================= دالة تصدير عرض تقديمي PowerPoint (.pptx) =================
def create_arabic_pptx(main_title: str, subtitle: str, slides_data: list) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    cover_slide = prs.slides.add_slide(blank_layout)
    bg = cover_slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(15, 44, 89)

    title_box = cover_slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.333), Inches(1.5))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = main_title
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    sub_box = cover_slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(11.333), Inches(1))
    sub_tf = sub_box.text_frame
    sub_p = sub_tf.paragraphs[0]
    sub_p.alignment = PP_ALIGN.CENTER
    sub_run = sub_p.add_run()
    sub_run.text = subtitle
    sub_run.font.size = Pt(22)
    sub_run.font.color.rgb = RGBColor(218, 192, 163)

    for item in slides_data:
        slide = prs.slides.add_slide(blank_layout)
        
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1))
        h_tf = header_box.text_frame
        h_tf.word_wrap = True
        h_p = h_tf.paragraphs[0]
        h_p.alignment = PP_ALIGN.RIGHT
        h_run = h_p.add_run()
        h_run.text = item.get("title", "")
        h_run.font.size = Pt(28)
        h_run.font.bold = True
        h_run.font.color.rgb = RGBColor(15, 44, 89)

        body_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5))
        b_tf = body_box.text_frame
        b_tf.word_wrap = True

        points = item.get("points", [])
        for idx, point in enumerate(points):
            bp = b_tf.paragraphs[0] if idx == 0 else b_tf.add_paragraph()
            bp.alignment = PP_ALIGN.RIGHT
            bp.space_after = Pt(14)
            b_run = bp.add_run()
            b_run.text = f"• {point}"
            b_run.font.size = Pt(20)
            b_run.font.color.rgb = RGBColor(40, 40, 40)

    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ================= لوحة تحكم الأدمن =================
@dp.message(Command("admin"))
async def admin_panel(message: Message):
    if message.from_user.id != ADMIN_TELEGRAM_ID:
        return
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📚 عرض وحذف الكتب المخزنة", callback_data="admin_list_books")],
        [InlineKeyboardButton(text="⭐ ترقية معلم لباقة بلس", callback_data="admin_add_premium")]
    ])
    await message.answer("🛠️ لوحة تحكم الأدمن:", reply_markup=keyboard)

@dp.callback_query(F.data == "admin_list_books")
async def admin_list_books(callback: CallbackQuery):
    if callback.from_user.id != ADMIN_TELEGRAM_ID:
        return
    conn = sqlite3.connect("teachers_bot.db")
    cursor = conn.cursor()
    cursor.execute("SELECT book_id, book_title, subject, grade FROM saved_books")
    books = cursor.fetchall()
    conn.close()

    if not books:
        await callback.message.edit_text("لا توجد كتب مخزنة.")
        await callback.answer()
        return

    text = "📚 الكتب المخزنة:\n"
    buttons = []
    for b_id, title, subj, grade in books:
        text += f"ID: {b_id} | {title} ({subj})\n"
        buttons.append([InlineKeyboardButton(text=f"🗑️ حذف ID: {b_id}", callback_data=f"admin_del_book_{b_id}")])
    await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await callback.answer()

@dp.callback_query(F.data.startswith("admin_del_book_"))
async def admin_delete_book(callback: CallbackQuery):
    if callback.from_user.id != ADMIN_TELEGRAM_ID:
        return
    book_id = int(callback.data.split("_")[3])
    conn = sqlite3.connect("teachers_bot.db")
    cursor = conn.cursor()
    cursor.execute("DELETE FROM saved_books WHERE book_id = ?", (book_id,))
    conn.commit()
    conn.close()
    await callback.message.answer(f"✅ تم حذف الكتاب {book_id}.")
    await callback.answer()

@dp.callback_query(F.data == "admin_add_premium")
async def admin_prompt_premium(callback: CallbackQuery, state: FSMContext):
    if callback.from_user.id != ADMIN_TELEGRAM_ID:
        return
    await callback.message.answer("أرسل معرف المستخدم (User ID) لترقيته:")
    await state.set_state(TeacherStates.admin_granting)
    await callback.answer()

@dp.message(TeacherStates.admin_granting)
async def process_grant_premium(message: Message, state: FSMContext):
    if message.from_user.id != ADMIN_TELEGRAM_ID:
        return
    try:
        target_id = int(message.text.strip())
        conn = sqlite3.connect("teachers_bot.db")
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO users (user_id, requests_count, last_date, is_premium) VALUES (?, 0, '', 1)", (target_id,))
        conn.commit()
        conn.close()
        await message.answer(f"⭐ تمت ترقية المستخدم {target_id} بنجاح.")
    except Exception as e:
        await message.answer(f"خطأ: {e}")
    await state.clear()

# ================= تفاعل المعلمين =================
@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📝 خطة درس", callback_data="serv_lesson"), InlineKeyboardButton(text="💡 شرح درس", callback_data="serv_explain")],
        [InlineKeyboardButton(text="📊 عرض تقديمي (PPTX)", callback_data="serv_pptx"), InlineKeyboardButton(text="📄 ورقة عمل", callback_data="serv_workSheet")],
        [InlineKeyboardButton(text="📌 ملخص سريع", callback_data="serv_summary"), InlineKeyboardButton(text="📸 امتحان من صورة", callback_data="serv_exam")],
        [InlineKeyboardButton(text="📚 تحليل كتاب PDF", callback_data="serv_pdf")]
    ])
    await message.answer("أهلاً بك يا أستاذ/ة في منصة المعلم الذكي 👨‍🏫\nاختر الخدمة المطلوبة:", reply_markup=keyboard)
    await state.set_state(TeacherStates.choosing_service)

@dp.callback_query(TeacherStates.choosing_service, F.data.startswith("serv_"))
async def process_service(callback: CallbackQuery, state: FSMContext):
    service_map = {
        "serv_lesson": "خطة درس",
        "serv_explain": "شرح درس",
        "serv_pptx": "عرض تقديمي",
        "serv_workSheet": "ورقة عمل",
        "serv_summary": "ملخص سريع",
        "serv_exam": "امتحان",
        "serv_pdf": "pdf_book"
    }
    await state.update_data(service_type=service_map.get(callback.data))
    
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📚 اللغة العربية", callback_data="subj_arabic"), InlineKeyboardButton(text="🔢 الرياضيات", callback_data="subj_math")],
        [InlineKeyboardButton(text="🧪 العلوم", callback_data="subj_science")]
    ])
    await callback.message.edit_text("اختر المادة الدراسية:", reply_markup=keyboard)
    await state.set_state(TeacherStates.choosing_subject)
    await callback.answer()

@dp.callback_query(TeacherStates.choosing_subject, F.data.startswith("subj_"))
async def process_subject(callback: CallbackQuery, state: FSMContext):
    subjects_map = {"subj_arabic": "اللغة العربية", "subj_math": "الرياضيات", "subj_science": "العلوم"}
    await state.update_data(subject=subjects_map.get(callback.data, "عام"))

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="الأول الإعدادي", callback_data="grade_1prep"), InlineKeyboardButton(text="الثاني الإعدادي", callback_data="grade_2prep")]
    ])
    await callback.message.edit_text("اختر الصف الدراسي:", reply_markup=keyboard)
    await state.set_state(TeacherStates.choosing_grade)
    await callback.answer()

@dp.callback_query(TeacherStates.choosing_grade, F.data.startswith("grade_"))
async def process_grade(callback: CallbackQuery, state: FSMContext):
    grades_map = {"grade_1prep": "الأول الإعدادي", "grade_2prep": "الثاني الإعدادي"}
    selected_grade = grades_map.get(callback.data, "عام")
    await state.update_data(grade=selected_grade)

    data = await state.get_data()
    service_type = data.get("service_type")

    if service_type == "pdf_book":
        saved = get_saved_books(data.get("subject"), selected_grade)
        buttons = [[InlineKeyboardButton(text=f"📖 {t}", callback_data=f"use_saved_{i}")] for i, t, _ in saved]
        buttons.append([InlineKeyboardButton(text="📥 رفع كتاب PDF جديد", callback_data="upload_new_pdf")])
        await callback.message.edit_text("اختر الكتاب أو ارفع كتاباً جديداً:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
        await state.set_state(TeacherStates.choosing_saved_book)
        await callback.answer()
        return

    prompts_map = {
        "خطة درس": "اكتب اسم أو موضوع الدرس لتحضيره:",
        "شرح درس": "اكتب عنوان أو موضوع الدرس الذي تريد شرحه:",
        "عرض تقديمي": "اكتب عنوان الدرس لتوليد عرض تقديمي (PowerPoint) جاهز للشرح:",
        "ورقة عمل": "اكتب موضوع الدرس لإنشاء ورقة عمل وتدريبات للطلاب:",
        "ملخص سريع": "اكتب موضوع الدرس لعمل ملخص مركز ومبسط:",
        "امتحان": "أرسل صورة صفحة الكتاب لتوليد الامتحان:"
    }
    
    await callback.message.edit_text(prompts_map.get(service_type, "أرسل طلبك:"))
    await state.set_state(TeacherStates.waiting_for_input)
    await callback.answer()

@dp.callback_query(TeacherStates.choosing_saved_book, F.data.startswith("use_saved_"))
async def use_saved(callback: CallbackQuery, state: FSMContext):
    book_id = int(callback.data.split("_")[2])
    conn = sqlite3.connect("teachers_bot.db")
    cursor = conn.cursor()
    cursor.execute("SELECT book_title, gemini_file_uri FROM saved_books WHERE book_id = ?", (book_id,))
    book = cursor.fetchone()
    conn.close()
    
    await state.update_data(book_title=book[0], file_uri=book[1])
    await callback.message.edit_text(f"✅ تم اختيار الكتاب: {book[0]}\nأرسل طلبك الآن:")
    await callback.answer()

@dp.callback_query(TeacherStates.choosing_saved_book, F.data == "upload_new_pdf")
async def req_pdf(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text("أرسل ملف الـ PDF كملف مستند:")
    await state.set_state(TeacherStates.waiting_for_pdf)
    await callback.answer()

@dp.message(TeacherStates.waiting_for_pdf, F.document)
async def handle_pdf(message: Message, state: FSMContext):
    if not check_and_update_quota(message.from_user.id):
        await message.answer("⚠️ نفذت محاولاتك اليومية.")
        await state.clear()
        return
    
    path = f"temp_{message.from_user.id}.pdf"
    file_info = await bot.get_file(message.document.file_id)
    await bot.download(file_info.file_path, destination=path)
    
    g_file = client.files.upload(file=path)
    data = await state.get_data()
    save_book_to_db(data.get("subject"), data.get("grade"), message.document.file_name, g_file.name)
    await state.update_data(file_uri=g_file.name)
    await message.answer("✅ تم رفع وحفظ الكتاب! أرسل طلبك الآن:")
    if os.path.exists(path):
        os.remove(path)

@dp.message(F.text, ~F.state.is_(None))
async def process_query(message: Message, state: FSMContext):
    data = await state.get_data()
    service_type = data.get("service_type")
    
    if not data.get("file_uri") and service_type not in ["خطة درس", "شرح درس", "عرض تقديمي", "ورقة عمل", "ملخص سريع"]:
        return
        
    if not check_and_update_quota(message.from_user.id):
        await message.answer("⚠️ نفذت محاولاتك اليومية.")
        await state.clear()
        return

    await message.answer("⏳ جاري المعالجة وتوليد الملف بالذكاء الاصطناعي...")
    try:
        subject = data.get('subject')
        grade = data.get('grade')

        if service_type == "عرض تقديمي":
            prompt = (
                f"أنشئ محتوى عرض تقديمي (7 شرائح) لمادة {subject} للصف {grade} حول درس: {message.text}. "
                "أرجع النتيجة بصيغة JSON فقط كقائمة (Array) من الكائنات، كل كائن يحتوي على: "
                "'title' (عنوان الشريحة القصير) و 'points' (قائمة من 3 إلى 4 نقاط مختصرة وواضحة للطلاب)."
            )
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config={
                    "response_mime_type": "application/json",
                    "temperature": 0.3
                }
            )
            slides_data = json.loads(resp.text)
            pptx_stream = create_arabic_pptx(
                main_title=message.text,
                subtitle=f"مادة {subject} — {grade}",
                slides_data=slides_data
            )
            await message.answer_document(
                BufferedInputFile(pptx_stream.getvalue(), filename=f"عرض_{message.text[:15]}.pptx"),
                caption="✅ تفضل العرض التقديمي (PowerPoint) جاهزاً للعرض والتعديل."
            )
            await state.clear()
            return

        if data.get("file_uri"):
            g_file = client.files.get(name=data.get("file_uri"))
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[g_file, f"طلب المعلم: {message.text}"],
                config={"temperature": 0.2}
            )
            doc_title = f"محتوى منهج {subject}"
            
        elif service_type == "شرح درس":
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"اكتب شرحاً تفصيلياً ومبسطاً للدرس في مادة {subject} للصف {grade}: {message.text}. مقسماً إلى مقدمة، أفكار رئيسية، وأمثلة توضيحية.",
                config={"temperature": 0.3}
            )
            doc_title = f"شرح درس - {message.text}"
            
        elif service_type == "ورقة عمل":
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"أنشئ ورقة عمل تعليمية احترافية لمادة {subject} للصف {grade} عن درس: {message.text}. تتضمن أسئلة تقويمية، تدريبات تطبيقية، وفراغات للإجابة.",
                config={"temperature": 0.3}
            )
            doc_title = f"ورقة عمل - {message.text}"
            
        elif service_type == "ملخص سريع":
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"اكتب ملخصاً سريعاً ومركزاً وشاملاً لأهم نقاط درس: {message.text} في مادة {subject} للصف {grade}.",
                config={"temperature": 0.3}
            )
            doc_title = f"ملخص الدرس - {message.text}"
            
        else:
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"موضوع الدرس: {message.text}",
                config={"temperature": 0.3}
            )
            doc_title = f"خطة درس - {subject}"
            
        stream = create_arabic_word_doc(doc_title, resp.text)
        await message.answer_document(BufferedInputFile(stream.getvalue(), filename="output.docx"), caption="✅ تفضل الملف جاهزاً للطباعة أو العرض.")
    except Exception as e:
        await message.answer(f"خطأ: {e}")
    await state.clear()

# ================= ربط وتفعيل الـ Webhook مع إطار العمل =================
webhook_requests_handler = SimpleRequestHandler(
    dispatcher=dp,
    bot=bot,
)
webhook_requests_handler.register(app, path=WEBHOOK_PATH)
setup_application(app, dp, bot=bot)