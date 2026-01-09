import re
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Расчет среднего балла за семестр (преобразование квалификатива в отметку с сотыми без округления)",
    page_icon="📘",
    layout="centered"
)

st.title("📘 Расчет среднего балла за семестр (преобразование квалификатива в отметку с сотыми без округления)")
st.write(
    "Загрузите журнал по предмету, скачанный с платформы Studii.md.\n\n"
    "Программа добавит в отчет дополнительную колонку «Средний балл отметка», "
    "в которой будет выполнен расчет среднего арифметического показателя "
    "за семестр путем преобразования каждого квалификатива в баллы.\n\n"
    "Скачайте новый дополненный отчет."
)

mapping = {
    'ОТ': 10, 'OT': 10, 'EX': 10,
    'ОХ': 9,  'OX': 9,  'FB': 9,
    'Х': 8,   'X': 8,   'B': 8,
    'У': 6,   'U': 6,   'S': 6
}

def is_date_col(col):
    s = str(col).strip()
    return bool(re.fullmatch(r'\d{1,2}[.,]\d{1,2}([.,]\d{2,4})?', s))

def process_file(file_bytes):
    sheets = pd.read_excel(BytesIO(file_bytes), sheet_name=None, header=1)

    for sheet_name, df in sheets.items():
        date_cols = [c for c in df.columns if is_date_col(c)]
        if not date_cols:
            sheets[sheet_name] = df
            continue

        grades = df[date_cols].astype(str).applymap(
            lambda x: str(x).strip().upper()
        )
        numeric = grades.replace(mapping)
        numeric = numeric.apply(pd.to_numeric, errors='coerce')

        avg = numeric.mean(axis=1)
        avg_trunc = np.floor(avg * 100) / 100

        target_col = None
        for c in df.columns:
            name = str(c).strip().lower()
            if name.startswith('media') or name.startswith('semestr'):
                target_col = c
                break

        if target_col is not None:
            insert_pos = df.columns.get_loc(target_col)
            df.insert(insert_pos, 'Средний балл отметка', avg_trunc)
        else:
            df['Средний балл отметка'] = avg_trunc

        sheets[sheet_name] = df

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output.getvalue()

uploaded_file = st.file_uploader(
    "📂 Загрузите журнал по предмету в формате Excel (.xlsx)",
    type=["xlsx"]
)

if uploaded_file:
    try:
        result_bytes = process_file(uploaded_file.read())

        st.success("Фаил был успешно обработан")
        st.download_button(
            label="⬇️ Скачайте готовый результат",
            data=result_bytes,
            file_name=f"result_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error("Ошибка при обработке файла. Проверьте структуру журнала.")
        st.exception(e)
