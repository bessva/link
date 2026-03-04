# -*- coding: utf-8 -*-
"""
линк_v4_app.py — Streamlit интерфейс для агента
Запуск: streamlit run линк_v4_app.py
"""

import os
import re
import uuid
import time
import glob
import requests
import streamlit as st
from datetime import datetime

# ==============================
# ⚙️ НАСТРОЙКИ
# ==============================

GIGACHAT_AUTH_KEY = "MDE5Y2E1MTQtMzEzYS03NzRkLTk0NWEtYTk0YjU0ZGFlMmM3OmE2ODk0ZjViLTMyMmMtNDUwYy1hZTdiLTRlY2ZiOGMxMTdlNw=="

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
DATA_DIR      = os.path.join(BASE_DIR, "data")
KNOWLEDGE_DIR = os.path.join(BASE_DIR, "knowledge")
VECTOR_DB_DIR = os.path.join(BASE_DIR, "vector_db")

for d in [DATA_DIR, KNOWLEDGE_DIR, VECTOR_DB_DIR]:
    os.makedirs(d, exist_ok=True)

# ==============================
# 🎨 Стили
# ==============================

st.set_page_config(
    page_title="ЛИНК — Энергетический ассистент",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.google.com/specimen/Montserrat?lang=ru_Cyrl');

/* Основной фон */
.stApp {
    background-color: #0d1117;
    font-family: 'Montserrat', sans-serif;
}

/* Сайдбар */
[data-testid="stSidebar"] {
    background-color: #161b22;
    border-right: 1px solid #21262d;
}
[data-testid="stSidebar"] * {
    color: #c9d1d9 !important;
}

/* Заголовок в сайдбаре */
.sidebar-title {
    font-family: 'Montserrat', monospace;
    font-size: 18px;
    font-weight: 600;
    color: #58a6ff !important;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 4px;
}
.sidebar-subtitle {
    font-size: 11px;
    color: #8b949e !important;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 24px;
}

/* Секции в сайдбаре */
.sidebar-section {
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #58a6ff !important;
    margin: 20px 0 8px 0;
    padding-bottom: 4px;
    border-bottom: 1px solid #21262d;
}

/* Элементы списка в сайдбаре */
.sidebar-item {
    font-size: 12px;
    color: #8b949e !important;
    padding: 3px 0;
    display: flex;
    align-items: center;
    gap: 6px;
}
.sidebar-item-active {
    color: #c9d1d9 !important;
}

/* Главная область */
.main-header {
    font-family: 'Montserrat', monospace;
    font-size: 13px;
    color: #58a6ff;
    letter-spacing: 3px;
    text-transform: uppercase;
    padding: 16px 0 4px 0;
    border-bottom: 1px solid #21262d;
    margin-bottom: 24px;
}

/* Сообщения чата */
.chat-container {
    display: flex;
    flex-direction: column;
    gap: 16px;
    padding-bottom: 20px;
}

.msg-user {
    background: #1c2128;
    border: 1px solid #30363d;
    border-radius: 8px 8px 2px 8px;
    padding: 12px 16px;
    margin-left: 60px;
    color: #c9d1d9;
    font-size: 14px;
    line-height: 1.6;
}

.msg-assistant {
    background: #161b22;
    border: 1px solid #21262d;
    border-left: 3px solid #58a6ff;
    border-radius: 2px 8px 8px 8px;
    padding: 12px 16px;
    margin-right: 60px;
    color: #c9d1d9;
    font-size: 14px;
    line-height: 1.6;
}

.msg-label-user {
    font-size: 10px;
    color: #8b949e;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 6px;
    text-align: right;
}
.msg-label-assistant {
    font-size: 10px;
    color: #58a6ff;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 6px;
    font-family: 'Montserrat', monospace;
}

/* Источник ответа */
.source-badge {
    display: inline-block;
    font-size: 9px;
    letter-spacing: 1px;
    text-transform: uppercase;
    padding: 2px 6px;
    border-radius: 3px;
    margin-right: 4px;
    font-family: 'Montserrat', monospace;
}
.source-excel { background: #1a3a1a; color: #3fb950; border: 1px solid #238636; }
.source-kb    { background: #1a1f3a; color: #79c0ff; border: 1px solid #1f6feb; }
.source-both  { background: #2d1f3a; color: #d2a8ff; border: 1px solid #6e40c9; }

/* Поле ввода */
.stTextInput input {
    background-color: #161b22 !important;
    border: 1px solid #30363d !important;
    border-radius: 6px !important;
    color: #c9d1d9 !important;
    font-family: 'Montserrat', sans-serif !important;
    font-size: 14px !important;
    padding: 10px 14px !important;
}
.stTextInput input:focus {
    border-color: #58a6ff !important;
    box-shadow: 0 0 0 3px rgba(88, 166, 255, 0.1) !important;
}

/* Кнопки */
.stButton button {
    background-color: #21262d !important;
    color: #c9d1d9 !important;
    border: 1px solid #30363d !important;
    border-radius: 6px !important;
    font-family: 'Montserrat', sans-serif !important;
    font-size: 12px !important;
    letter-spacing: 0.5px !important;
    transition: all 0.2s !important;
}
.stButton button:hover {
    background-color: #30363d !important;
    border-color: #58a6ff !important;
    color: #58a6ff !important;
}

/* Статус */
.status-ok  { color: #3fb950; font-size: 11px; }
.status-err { color: #f85149; font-size: 11px; }

/* Скрываем стандартные элементы streamlit */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 20px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)


# ==============================
# 🔐 GigaChat
# ==============================

def get_access_token():
    if ("_token" in st.session_state and
            time.time() < st.session_state.get("_token_exp", 0) - 60):
        return st.session_state["_token"]

    url = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
        "RqUID": str(uuid.uuid4()),
        "Authorization": f"Basic {GIGACHAT_AUTH_KEY}",
    }
    r = requests.post(url, headers=headers, data={"scope": "GIGACHAT_API_PERS"},
                      verify=False, timeout=15)
    r.raise_for_status()
    d = r.json()
    st.session_state["_token"]     = d["access_token"]
    st.session_state["_token_exp"] = d.get("expires_at", 0) / 1000
    return st.session_state["_token"]


def call_gigachat(system_prompt, user_message, history):
    token = get_access_token()
    messages = [{"role": "system", "content": system_prompt}]
    for q, a, _ in history[-5:]:
        messages.append({"role": "user",      "content": q})
        messages.append({"role": "assistant", "content": a})
    messages.append({"role": "user", "content": user_message})

    r = requests.post(
        "https://gigachat.devices.sberbank.ru/api/v1/chat/completions",
        headers={"Accept": "application/json", "Content-Type": "application/json",
                 "Authorization": f"Bearer {token}"},
        json={"model": "GigaChat-Pro", "messages": messages,
              "temperature": 0.3, "max_tokens": 1000},
        verify=False, timeout=30
    )
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"].strip()


# ==============================
# 📊 Excel
# ==============================

def _find_col(headers, keywords):
    for i, h in enumerate(headers):
        if h and any(kw in str(h).lower() for kw in keywords):
            return i
    return None

@st.cache_data(show_spinner=False)
def load_excel_data(folder):
    try:
        import openpyxl
    except ImportError:
        return {}
    all_data = {}
    for filepath in glob.glob(os.path.join(folder, "*.xlsx")):
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                rows = list(ws.iter_rows(values_only=True))
                if len(rows) < 3:
                    continue
                headers    = rows[0]
                col_tes    = _find_col(headers, ["наименование тэс", "наименование"])
                col_month  = _find_col(headers, ["месяц"])
                col_fuel   = _find_col(headers, ["вид топлива"])
                sheet_upper = sheet_name.upper()
                result_col = None
                if "ННЗТ" in sheet_upper: result_col = _find_col(headers, ["ннзт"])
                elif "НЭЗТ" in sheet_upper: result_col = _find_col(headers, ["нэзт, т.н.т", "нэзт"])
                elif "НАЗТ" in sheet_upper: result_col = _find_col(headers, ["назт, т.н.т", "назт"])
                elif "ОНЗТ" in sheet_upper: result_col = _find_col(headers, ["онзт, т.н.т", "онзт"])
                if result_col is None or col_month is None:
                    continue
                current_tes = current_fuel = None
                for row in rows[2:]:
                    if col_tes is not None and row[col_tes]:
                        current_tes = str(row[col_tes]).strip()
                    if col_fuel is not None and row[col_fuel]:
                        current_fuel = str(row[col_fuel]).strip()
                    month_val = row[col_month] if col_month is not None else None
                    result    = row[result_col]
                    if not current_tes or not month_val or result is None:
                        continue
                    all_data.setdefault(current_tes, {}).setdefault(sheet_upper, {})[str(month_val).strip()] = {
                        "value": round(float(result), 3), "fuel": current_fuel or "—"
                    }
        except Exception:
            pass
    return all_data


def find_relevant_excel(excel_data, question):
    q = question.lower()
    norm_types = [n for n in ["ННЗТ","НЭЗТ","НАЗТ","ОНЗТ"] if n.lower() in q] or ["ННЗТ","НЭЗТ","НАЗТ","ОНЗТ"]
    months_ru  = ["январь","февраль","март","апрель","май","июнь",
                  "июль","август","сентябрь","октябрь","ноябрь","декабрь"]
    target_months   = [m.capitalize() for m in months_ru if m in q]
    target_stations = [t for t in excel_data if any(w in q for w in t.lower().split())] or list(excel_data.keys())
    lines = []
    for tes in target_stations:
        block = [f"📍 {tes}:"]
        for norm in norm_types:
            if norm in excel_data.get(tes, {}):
                block.append(f"  {norm}:")
                for month, info in excel_data[tes][norm].items():
                    if not target_months or month in target_months:
                        block.append(f"    {month}: {info['value']} т.н.т.")
        if len(block) > 1:
            lines.extend(block)
    return "\n".join(lines)


# ==============================
# 📚 База знаний
# ==============================

def extract_pdf(path):
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: text += t + "\n"
        return text
    except Exception:
        return ""

def extract_docx(path):
    try:
        from docx import Document
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception:
        return ""

def chunk_text(text, size=800, overlap=100):
    words = text.split()
    chunks, i = [], 0
    while i < len(words):
        c = " ".join(words[i:i+size])
        if c.strip(): chunks.append(c)
        i += size - overlap
    return chunks

@st.cache_resource(show_spinner=False)
def build_knowledge_base(knowledge_dir, vector_db_dir):
    try:
        import chromadb
        from sentence_transformers import SentenceTransformer
    except ImportError:
        return None, None

    files = (glob.glob(os.path.join(knowledge_dir, "*.pdf")) +
             glob.glob(os.path.join(knowledge_dir, "*.docx")))
    if not files:
        return None, None

    client = chromadb.PersistentClient(path=vector_db_dir)
    existing = [c.name for c in client.list_collections()]
    model = SentenceTransformer("paraphrase-multilingual-mpnet-base-v2")

    if "knowledge" in existing:
        col = client.get_collection("knowledge")
        if col.count() > 0:
            return col, model

    if "knowledge" in existing:
        client.delete_collection("knowledge")
    col = client.create_collection("knowledge")

    all_chunks, all_ids, all_metas = [], [], []
    for filepath in files:
        fname = os.path.basename(filepath)
        ext   = os.path.splitext(filepath)[1].lower()
        text  = extract_pdf(filepath) if ext == ".pdf" else extract_docx(filepath)
        if not text.strip(): continue
        for i, chunk in enumerate(chunk_text(text)):
            all_chunks.append(chunk)
            all_ids.append(f"{fname}_{i}")
            all_metas.append({"source": fname, "chunk": i})

    if all_chunks:
        embeddings = model.encode(all_chunks, show_progress_bar=False).tolist()
        col.add(documents=all_chunks, embeddings=embeddings,
                ids=all_ids, metadatas=all_metas)
    return col, model


def search_kb(question, collection, model, top_k=3):
    if collection is None or model is None: return ""
    try:
        q_emb   = model.encode([question]).tolist()
        results = collection.query(query_embeddings=q_emb, n_results=top_k)
        docs    = results.get("documents", [[]])[0]
        metas   = results.get("metadatas", [[]])[0]
        return "\n\n---\n\n".join(f"[{m.get('source','?')}]\n{d}" for d, m in zip(docs, metas))
    except Exception:
        return ""


# ==============================
# 🧠 Классификация вопроса
# ==============================

def classify_question(q):
    q = q.lower()
    excel_kw = ["ннзт","нэзт","назт","онзт","нвзт","январь","февраль","март","апрель",
                "май","июнь","июль","август","сентябрь","октябрь","ноябрь","декабрь",
                "тэц","грэс","станци","топлив","сколько","значение","показател","данные"]
    know_kw  = ["как","почему","что делать","рекоменд","метод","подход","принцип",
                "системн","ситуацион","чс","аварий","управлен","решени","анализ",
                "оценк","риск","действи","план","алгоритм","порядок"]
    has_e = any(k in q for k in excel_kw)
    has_k = any(k in q for k in know_kw)
    if has_e and has_k: return "both"
    if has_k: return "both"
    if has_e: return "both"
    return "both"


SYSTEM_PROMPT = (
    "Ты — эксперт по управлению энергетическими объектами и нормативам "
    "топливных запасов (ННЗТ, НЭЗТ, НАЗТ, ОНЗТ, НВЗТ). "
    "Ты умеешь анализировать числовые данные по станциям и давать "
    "управленческие рекомендации на основе системно-ситуационного подхода. "
    "Отвечай КРАТКО и по существу — максимум 3-5 предложений или короткий список. "
    "Без вступлений, без повторений вопроса, без лишних слов. Только факты и выводы. "
    "ВАЖНО: отвечай ТОЛЬКО на вопросы про энергетику, ТЭЦ, ГРЭС, топливные запасы "
    "и управление энергообъектами. На остальное отвечай: "
    "'Я специализируюсь на энергетике и не отвечаю на вопросы вне этой темы.'"
)


# ==============================
# 🖥️ Интерфейс
# ==============================

# Инициализация состояния
if "history" not in st.session_state:
    st.session_state.history = []

# Загрузка данных
excel_data = load_excel_data(DATA_DIR)

with st.spinner("Загружаю базу знаний..."):
    collection, st_model = build_knowledge_base(KNOWLEDGE_DIR, VECTOR_DB_DIR)

# ── Сайдбар ──────────────────────────────────────────

with st.sidebar:
    st.markdown('<div class="sidebar-title">⚡ ЛИНК</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-subtitle">Энергетический ассистент</div>', unsafe_allow_html=True)

    # Станции
    st.markdown('<div class="sidebar-section">Станции</div>', unsafe_allow_html=True)
    if excel_data:
        for tes in excel_data.keys():
            st.markdown(f'<div class="sidebar-item sidebar-item-active">🏭 {tes}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="sidebar-item">Нет данных в папке data/</div>', unsafe_allow_html=True)

    # База знаний
    st.markdown('<div class="sidebar-section">База знаний</div>', unsafe_allow_html=True)
    kb_files = (glob.glob(os.path.join(KNOWLEDGE_DIR, "*.pdf")) +
                glob.glob(os.path.join(KNOWLEDGE_DIR, "*.docx")))
    if kb_files:
        for f in kb_files:
            st.markdown(f'<div class="sidebar-item sidebar-item-active">📄 {os.path.basename(f)}</div>',
                        unsafe_allow_html=True)
    else:
        st.markdown('<div class="sidebar-item">Нет документов в папке knowledge/</div>',
                    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Кнопка очистить
    if st.button("🗑  Очистить чат"):
        st.session_state.history = []
        st.rerun()

    # Статус подключения
    st.markdown('<div class="sidebar-section">Статус</div>', unsafe_allow_html=True)
    try:
        get_access_token()
        st.markdown('<div class="status-ok">● GigaChat подключён</div>', unsafe_allow_html=True)
    except Exception:
        st.markdown('<div class="status-err">● Ошибка подключения</div>', unsafe_allow_html=True)

    if collection is not None:
        st.markdown(f'<div class="status-ok">● База знаний активна ({collection.count()} фрагм.)</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-err">● База знаний не загружена</div>', unsafe_allow_html=True)


# ── Основная область ─────────────────────────────────

st.markdown('<div class="main-header">// Диалог с ассистентом</div>', unsafe_allow_html=True)

# История чата
for user_msg, bot_msg, source in st.session_state.history:
    source_badge = {
        "excel":     '<span class="source-badge source-excel">📊 Данные</span>',
        "knowledge": '<span class="source-badge source-kb">📚 База знаний</span>',
        "both":      '<span class="source-badge source-both">📊📚 Оба источника</span>',
    }.get(source, "")

    st.markdown(f"""
        <div class="msg-label-user">Вы</div>
        <div class="msg-user">{user_msg}</div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
        <div class="msg-label-assistant">⚡ ЛИНК {source_badge}</div>
        <div class="msg-assistant">{bot_msg}</div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

# Поле ввода
st.markdown("<br>", unsafe_allow_html=True)
with st.form("chat_form", clear_on_submit=True):
    col1, col2 = st.columns([6, 1])
    with col1:
        user_input = st.text_input(
            "",
            placeholder="Введите вопрос... (например: ННЗТ Челябинская ТЭЦ-1 за январь)",
            label_visibility="collapsed"
        )
    with col2:
        submitted = st.form_submit_button("Отправить", use_container_width=True)

if submitted and user_input.strip():
    q_type = classify_question(user_input)
    context_parts = []

    if q_type in ("excel", "both") and excel_data:
        ec = find_relevant_excel(excel_data, user_input)
        if ec: context_parts.append(f"📊 ДАННЫЕ ПО СТАНЦИЯМ:\n{ec}")

    if q_type in ("knowledge", "both") and collection is not None:
        kc = search_kb(user_input, collection, st_model)
        if kc: context_parts.append(f"📚 ИЗ БАЗЫ ЗНАНИЙ:\n{kc}")

    user_message = user_input + ("\n\n" + "\n\n".join(context_parts) if context_parts else "")

    with st.spinner("Думаю..."):
        try:
            answer = call_gigachat(SYSTEM_PROMPT, user_message, st.session_state.history)
        except Exception as e:
            answer = f"❌ Ошибка: {e}"

    st.session_state.history.append((user_input, answer, q_type))
    st.rerun()
