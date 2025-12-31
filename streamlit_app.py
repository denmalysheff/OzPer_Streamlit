import streamlit as st
import pandas as pd
import io
import os
import plotly.express as px

# 1. Функция исправления заголовков
def fix_headers(df):
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

# Функция поиска нужного листа
def find_sheet(xl, target_name):
    target_cleaned = target_name.replace(" ", "").upper()
    for sheet in xl.sheet_names:
        if sheet.replace(" ", "").upper() == target_cleaned:
            return sheet
    return None

# --- НАСТРОЙКА СТРАНИЦЫ ---
st.set_page_config(page_title="Детальный мониторинг Nуч", layout="wide")

# --- ОФОРМЛЕНИЕ (ЗАСТАВКА) ---
if os.path.exists("header.png"):
    st.image("header.png", use_container_width=True)

st.title("🚂 Сравнительный анализ Nуч и изменений")
st.markdown("---")

# --- ЗАГРУЗКА БАЗЫ СТАНЦИЙ ---
base_file_name = "stations_base.xlsx"
if os.path.exists(base_file_name):
    try:
        df_base_raw = pd.read_excel(base_file_name)
        df_base = fix_headers(df_base_raw)
        df_base = df_base.dropna(subset=['КООРДИНАТА', 'НАПРАВЛЕНИЕ'])
        df_base['КООРДИНАТА'] = pd.to_numeric(df_base['КООРДИНАТА'], errors='coerce')
        df_base = df_base.dropna(subset=['КООРДИНАТА'])
    except Exception as e:
        st.error(f"Ошибка в базе станций: {e}")
        st.stop()
else:
    st.error(f"❌ Файл '{base_file_name}' не найден!")
    st.stop()

# --- ЗАГРУЗКА ФАЙЛОВ ---
col_up1, col_up2 = st.columns(2)
with col_up1:
    file_prev = st.file_uploader("📂 ПРОШЛЫЙ месяц", type="xlsx")
with col_up2:
    file_curr = st.file_uploader("📂 ТЕКУЩИЙ месяц", type="xlsx")

def process_excel_data(file):
    if file is None: return None
    try:
        xl = pd.ExcelFile(file)
        sheet = find_sheet(xl, "Оценка КМ")
        if not sheet:
            st.warning(f"Лист 'Оценка КМ' не найден в {file.name}")
            return None
        df = pd.read_excel(file, sheet_name=sheet)
        df = fix_headers(df)
        cols = ['КМ', 'ОЦЕНКА', 'КОДНАПР', 'ПУТЬ']
        df = df.dropna(subset=cols)
        for c in cols:
            df[c] = pd.to_numeric(df[c], errors='coerce')
        return df.dropna(subset=cols)
    except: return None

def get_detailed_results(df_eval, df_base):
    if df_eval is None: return {}
    results = {}
    valid_dirs = {24602, 24603, 24701}
    for direction in df_base['НАПРАВЛЕНИЕ'].unique():
        if direction not in valid_dirs: continue
        stations = df_base[df_base['НАПРАВЛЕНИЕ'] == direction].sort_values('КООРДИНАТА')
        paths = df_eval[df_eval['КОДНАПР'] == direction]['ПУТЬ'].unique()
        for path in paths:
            for i in range(len(stations) - 1):
                st_a, st_b = stations.iloc[i], stations.iloc[i+1]
                km_s, km_e = int(st_a['КООРДИНАТА']) + 1, int(st_b['КООРДИНАТА'])
                seg = df_eval[(df_eval['КОДНАПР'] == direction) & (df_eval['ПУТЬ'] == path) & 
                              (df_eval['КМ'] >= km_s) & (df_eval['КМ'] <= km_e)]
                if not seg.empty:
                    s5, s4, s3, s2 = (seg['ОЦЕНКА']==5).sum(), (seg['ОЦЕНКА']==4).sum(), \
                                     (seg['ОЦЕНКА']==3).sum(), (seg['ОЦЕНКА']==2).sum()
                    n_uch = round((s5*5 + s4*4 + s3*3 - s2*5) / len(seg), 2)
                    km_map = dict(zip(seg['КМ'].astype(int), seg['ОЦЕНКА'].astype(int)))
                    key = f"{direction}_{path}_{st_a['СТАНЦИЯ']}_{st_b['СТАНЦИЯ']}"
                    results[key] = {
                        'Направление': int(direction), 'Путь': int(path),
                        'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                        'Км нач': int(km_s), 'Км кон': int(km_e), 'Nуч': n_uch,
                        'Отл': int(s5), 'Хор': int(s4), 'Удов': int(s3), 'Неуд': int(s2),
                        'km_map': km_map
                    }
    return results

if file_curr:
    df_c_data = process_excel_data(file_curr)
    res_curr = get_detailed_results(df_c_data, df_base)
    res_prev = get_detailed_results(process_excel_data(file_prev), df_base) if file_prev else {}

    comparison = []
    for key, data in res_curr.items():
        prev = res_prev.get(key, {})
        data['Прошлый Nуч'] = prev.get('Nуч', data['Nуч'])
        data['Динамика'] = round(data['Nуч'] - data['Прошлый Nуч'], 2)
        
        changes = []
        curr_map = data.pop('km_map', {})
        prev_map = prev.get('km_map', {})
        for km, score in curr_map.items():
            if km in prev_map and score != prev_map[km]:
                changes.append(f"{km}км({prev_map[km]}→{score})")
        
        data['Изменившиеся км'] = ", ".join(changes) if changes else "Без изменений"
        comparison.append(data)

    df_final = pd.DataFrame(comparison).sort_values('Nуч')

    # График
    st.plotly_chart(px.bar(df_final, x='Перегон', y='Динамика', color='Динамика', 
                           color_continuous_scale='RdYlGn', title="Динамика Nуч"), use_container_width=True)

    # Таблица (исправленный стайлинг)
    def color_dyn(val):
        if isinstance(val, (int, float)):
            return 'color: green' if val > 0 else ('color: red' if val < 0 else '')
        return ''

    st.dataframe(
        df_final.style.format({"Nуч": "{:.2f}", "Прошлый Nуч": "{:.2f}", "Динамика": "{:+.2f}"})
        .background_gradient(subset=['Nуч'], cmap='RdYlGn')
        .map(color_dyn, subset=['Динамика']), # В новых версиях .map вместо .applymap
        use_container_width=True
    )

    # Скачивание
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Анализ')
    
    st.download_button(label="📥 Скачать отчет", data=output.getvalue(), 
                       file_name="Nuch_Report.xlsx", mime="application/vnd.ms-excel")
