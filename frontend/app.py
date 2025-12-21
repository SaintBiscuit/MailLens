import sys
import os
import json
import torch
import streamlit as st
import pandas as pd
from datetime import datetime


os.environ["STREAMLIT_FRAGMENTS"] = "0"
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.mail_parser import parse_email, prepare_for_classification
from backend.classifier import MailClassifier
from backend.injection_guard import detect_injection


@st.cache_resource(show_spinner=True)
def load_classifier_once():
    print("=" * 50)
    print("ЗАГРУЗКА МОДЕЛИ")
    
    # Очистка GPU памяти
    if torch.cuda.is_available():
        torch.cuda.empty_cache()
        print(f"До загрузки: {torch.cuda.memory_allocated()/1e9:.2f} GB")
    
    classifier = MailClassifier()
    
    if torch.cuda.is_available():
        print(f"После загрузки: {torch.cuda.memory_allocated()/1e9:.2f} GB")
    print("=" * 50)
    
    return classifier

st.set_page_config(page_title="MailLens", layout="wide")
st.title("MailLens — интеллектуальная категоризация писем")

def auto_load_categories_on_startup():
    """Автоматически загружает категории при первом запуске"""
    
    default_cats = {
        'Техническая поддержка': {
            'folder_with_examples': 'Technical support',
            'description': 'Письма от клиентов или сотрудников с запросами о работе систем, программ, оборудования: ошибки, сбои, вопросы по функционалу, просьбы о помощи, инциденты.',
            },
        'Финансовые операции, чеки и счета': {
            'folder_with_examples': 'Financial transactions, checks and invoices',
            'description': 'Счета на оплату, выставленные/полученные счета-фактуры, чеки, платёжные уведомления, запросы на возврат средств, подтверждения транзакций.',
            },
        'Вакансии и карьера': {
            'folder_with_examples': 'Vacancies and careers',
            'description': 'Резюме соискателей, письма от рекрутинговых агентств, запросы на стажировки, внутренние анонсы вакансий, приглашения на собеседования, запросы на оценку кандидатов.',
            }, 
        'Рекламная рассылка': {
            'folder_with_examples': 'Promotional mailing',
            'description': 'Письма с коммерческими предложениями, акциями, скидками, презентациями продуктов и услуг от внешних компаний или собственного маркетинга.',
            }, 
        'Новостные рассылки': {
            'folder_with_examples': 'Newsletters',
            'description': 'Информационные бюллетени: отраслевые новости, обновления законодательства, корпоративные анонсы, аналитика, обзоры рынка — без прямого призыва к действию.',
            }, 
        'Регистрация и подтверждение': {
            'folder_with_examples': 'Registration and confirmation',
            'description': 'Письма, связанные с созданием или верификацией аккаунтов: подтверждение email, сброс пароля, двухфакторная аутентификация, привязка устройств.',
            }, 
        'Транспорт и путешествия': {
            'folder_with_examples': 'Transport and travel',
            'description': 'Бронирование билетов и отелей, уведомления о перелётах/поездах на такси, запросы на командировки.',
            }, 
        'Неприемлемый контент': {
            'folder_with_examples': 'Harm content',
            'description': 'Спам с порнографией, насилием, лотереями, экстремизмом',
            }, 
        'Бизнес-корреспонденция': {
            'folder_with_examples': 'Business and correspondence',
            'description': 'Официальные письма от партнёров, поставщиков, клиентов и госорганов: предложения сотрудничества, переговоры, юридические запросы, договоры, претензии.',
            }, 
        'Системные и сервисные уведомления': {
            'folder_with_examples': 'System and service notifications',
            'description': 'Автоматические сообщения от IT-систем: отчёты, алерты, уведомления о бэкапах, обновлениях, ошибках в интеграциях, статусы задач из CRM/ERP и т.п.',
            },
    }
    
    if st.session_state.auto_categories_loaded:
        base_path = "emails_by_catrgories"
        if os.path.exists(base_path):
            categories_loaded = []
            for category, info in default_cats.items():
                if info['folder_with_examples'] in os.listdir(base_path):
                    category_path = os.path.join(base_path, info['folder_with_examples'])
                    
                    if os.path.isdir(category_path):
                        files = [f for f in os.listdir(category_path) 
                                if f.endswith(('.eml', '.msg'))]
                        
                        if files:
                            example_texts = []
                            
                            for filename in files:
                                filepath = os.path.join(category_path, filename)
                                with open(filepath, 'rb') as f:
                                    parsed = parse_email(f.read(), filename)
                                    text = prepare_for_classification(parsed)
                                    example_texts.append(text)
                            st.session_state.classifier.add_category(
                                category, 
                                description=info['description'],
                                example_texts=example_texts,
                            )
                            categories_loaded.append(category)
            
            st.toast(f"Автоматически загружено {len(categories_loaded)} категорий", icon="✅")
    
    st.session_state.auto_categories_loaded = False


# Инициализация хранилища результатов
if "classifier" not in st.session_state:
    st.session_state.classifier = load_classifier_once()
if "disabled_uploder" not in st.session_state:
    st.session_state.disabled_uploder = False
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 1
if "auto_categories_loaded" not in st.session_state:
    st.session_state.auto_categories_loaded = True
if st.session_state.auto_categories_loaded:
    auto_load_categories_on_startup()


@st.dialog("Добавление новой категории")
def add_catigory():
    category_name = st.text_input("Введите название категории")
    description = st.text_area(label="Введите описание категории")
    example_files = st.file_uploader(
        f"Примеры писем для данной категории",
        type=["eml", "msg"],
        accept_multiple_files=True,
        key=st.session_state.uploader_key + 1
    )

    if st.button("Добавить категорию"):
        example_texts = []
        if not category_name:
            st.toast(f"Введите название категории", icon="⚠️")
        elif not description:
            st.toast(f"Добавьте небольшое описание категории", icon="⚠️")
        elif not example_files:
            st.toast(f"Добавьте примеры для категории", icon="⚠️")
        else:
            for file in example_files:
                parsed = parse_email(file.read(), file.name)
                data_for_classifier = prepare_for_classification(parsed)
                example_texts.append(data_for_classifier)
            st.session_state.classifier.add_category(
                category=category_name,
                description=description,
                example_texts=example_texts
            )
            st.rerun()

with st.sidebar:
    st.header("Доступные для распознования категории")
    if st.session_state.classifier.categories:
        for category in st.session_state.classifier.categories:
            st.write(f"• {category}")
    else:
        st.write('Добавьте категории для распознования или добавьте стандартные категории')
        if st.sidebar.button(f"Добавить стандартные категории"):
            st.session_state.auto_categories_loaded = True
            st.rerun()
    if st.sidebar.button(f"Добавить новую категорию"):
        add_catigory()
    if st.session_state.classifier.categories:
        if st.sidebar.button(f"Сбросить все категории"):
            st.session_state.classifier.categories = {}
            st.rerun()

    threshold = st.slider("Порог «Не определена»", 0.05, 0.90, 0.85, 0.005, 
                          help="Чем ниже — тем больше писем будет классифицировано")
    st.session_state.classifier.threshold = threshold

def process_new_email(file):
    """Обрабатывает новое письмо и кэширует результат"""
    try:
        parsed = parse_email(file.read(), file.name)
        data_for_classifier = prepare_for_classification(parsed)
        data_for_classifier = detect_injection(data_for_classifier)
        prediction = st.session_state.classifier.predict(data_for_classifier)
        result = {
            "file_name": file.name,
            "file_size": file.size,
            "predicted_category": prediction["predicted_category"],
            "best_similarity": prediction['best_similarity'],
            "all_scores": prediction["all_scores"],
            "data_for_classifier": data_for_classifier,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'error': None
        }
    except Exception as e:
        st.error(f"Ошибка: {e}")
        result = {
            "file_name": file.name,
            "file_size": file.size,
            "predicted_category": None,
            "best_similarity": None,
            "all_scores": None,
            "data_for_classifier": None,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'error': str(e)
        }
    st.session_state.results.append(result)

# Загрузка писем
uploaded_files = st.file_uploader(
    "Загрузите письма для оценки (Доступные форматы: .eml, .msg)", 
    type=["eml", "msg"], 
    accept_multiple_files=True,
    key=st.session_state.uploader_key,
    disabled=st.session_state.disabled_uploder,
)
if uploaded_files:
    for file in uploaded_files:
        process_new_email(file)
    else:
        st.session_state.uploader_key += 1
        st.rerun()

if "results" not in st.session_state:
    st.session_state.results = []
else:
    for result in st.session_state.results:
        with st.expander(f"Письмо: {result['file_name']} ({result['file_size']:,} байт)"):
            if result['error'] is not None:
                st.error(f"Ошибка: {result['error']}")
            else:
                st.success(f"**{result['predicted_category']}** (Векторная близость: {result['best_similarity']:.3f})")
                st.json(result['all_scores'], expanded=False)
                st.text(f'Оцениваемые данные\n{result["data_for_classifier"]}', width='stretch')

# === ЭКСПОРТ ===
if st.session_state.results:
    df = pd.DataFrame(st.session_state.results)

    col1, col2 = st.columns(2)
    with col1:
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Скачать результаты CSV",
            csv,
            "maillens_results.csv",
            "text/csv"
        )
    with col2:
        jsonl = "\n".join(json.dumps(r, ensure_ascii=False) for r in st.session_state.results)
        st.download_button(
            "Скачать результаты JSONL",
            jsonl,
            "maillens_results.jsonl",
            "application/jsonlines"
        )
