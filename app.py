import streamlit as st
from processor import process_document

st.set_page_config(page_title="Plagiat Highlighter", page_icon="🔍", layout="centered")
st.title("🔍 Plagiat Highlighter")
st.markdown("Автоматичне підсвічування збігів у порівняльній таблиці плагіату")

uploaded_file = st.file_uploader("Завантажте .docx файл з таблицею", type=["docx"])

with st.expander("⚙️ Налаштування"):
    threshold = st.slider(
        "Поріг схожості слів (%)", 60, 100, 75,
        help="75% — оптимально для морфології укр. мови. 'суд'/'суду' = збіг."
    )
    font_name = st.selectbox("Шрифт", ["Times New Roman", "Calibri", "Arial"], index=0)
    font_size = st.number_input("Розмір шрифту (pt)", 8, 16, 12)

if uploaded_file and st.button("🚀 Обробити файл", type="primary"):
    progress_bar = st.progress(0, text="Починаємо обробку...")

    def on_progress(current, total, row_num):
        pct = int(current / total * 100)
        progress_bar.progress(pct, text=f"Обробка рядка {row_num} з {total}...")

    result_bytes, warnings, stats = process_document(
        uploaded_file.read(),
        threshold=threshold,
        font_name=font_name,
        font_size=font_size,
        progress_callback=on_progress
    )
    progress_bar.progress(100, text="Готово!")

    col1, col2, col3 = st.columns(3)
    col1.metric("Оброблено рядків", stats['processed'])
    col2.metric("⚠️ Проблем",        stats['warnings'])
    col3.metric("Пропущено",         stats['skipped'])

    if warnings:
        with st.expander(f"⚠️ Детальні попередження ({len(warnings)})", expanded=True):
            for w in warnings:
                st.warning(w)

    st.success("✅ Файл готовий!")
    st.download_button(
        "📥 Завантажити результат",
        data=result_bytes,
        file_name="highlighted_" + uploaded_file.name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
