import streamlit as st
import pandas as pd
import io
import os
import plotly.express as px

# 1. Функция исправления заголовков (нормализация символов)
def fix_headers(df):
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

# Функция поиска нужного листа (игнорирует пробелы и регистр)
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
else:
    # Если картинки нет, просто выводим красивый заголовок
    st.title("🚂 Мониторинг состояния пути")

st.title("Сравнительный анализ Nуч и изменений по километрам")
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
    st.error(f"❌ Файл '{base_file_name}' не найден в папке приложения!")
    st.stop()

# --- ИНТЕРФЕЙС ЗАГРУЗКИ ФАЙЛОВ ---
col_up1, col_up2 = st.columns(2)
with col_up1:
    file_prev = st.file_uploader("📂 ПРОШЛЫЙ месяц (База для сравнения)", type="xlsx")
with col_up2:
    file_curr = st.file_uploader("📂 ТЕКУЩИЙ месяц (Результат)", type="xlsx")

def process_excel_data(file):
    if file is None: return None
    try:
        xl = pd.ExcelFile(file)
        sheet = find_sheet(xl, "Оценка КМ")
        if not sheet:
            st.warning(f"В файле {file.name} не найден лист 'Оценка КМ' или 'ОценкаКМ'")
            return None
        
        df = pd.read_excel(file, sheet_name=sheet)
        df = fix_headers(df)
        cols = ['КМ', 'ОЦЕНКА', 'КОДНАПР', 'ПУТЬ']
        
        if not all(c in df.columns for c in cols):
            st.error(f"В файле {file.name} отсутствуют колонки: {cols}")
            return None
            
        df = df.dropna(subset=cols)
        for c in cols:
            df[c] = pd.to_numeric(df[c], errors='coerce')
        return df.dropna(subset=cols)
    except Exception as e:
        st.error(f"Ошибка при обработке {file.name}: {e}")
        return None

def get_detailed_results(df_eval, df_base):
    if df_eval is None: return {}
    results = {}
    valid_dirs = {24602, 24603, 24701} # Список анализируемых направлений
    
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
                    
                    # Формула Nуч
                    n_uch = round((s5*5 + s4*4 + s3*3 - s2*5) / len(seg), 2)
                    km_map = dict(zip(seg['КМ'].astype(int), seg['ОЦЕНКА'].astype(int)))
                    
                    key = f"{direction}_{path}_{st_a['СТАНЦИЯ']}_{st_b['СТАНЦИЯ']}"
                    results[key] = {
                        'Направление': int(direction), 'Путь': int(path),
                        'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                        'Км нач': int(km_s), 'Км кон': int(km_e), 'Всего Км': int(len(seg)),
                        'Nуч': n_uch, 'Отл': int(s5), 'Хор': int(s4), 'Удов': int(s3), 'Неуд': int(s2),
                        'km_map': km_map
                    }
    return results

# --- ОСНОВНАЯ ЛОГИКА ---
if file_curr:
    df_c_data = process_excel_data(file_curr)
    res_curr = get_detailed_results(df_c_data, df_base)
    
    df_p_data = process_excel_data(file_prev) if file_prev else None
    res_prev = get_detailed_results(df_p_data, df_base) if df_p_data is not None else {}

    comparison_results = []
    for key, data in res_curr.items():
        prev_data = res_prev.get(key, {})
        prev_nuch = prev_data.get('Nуч', None)
        prev_km_map = prev_data.get('km_map', {})
        curr_km_map = data.get('km_map', {})
        
        # Сравнение оценок по каждому километру
        changes = []
        for km, score in curr_km_map.items():
            if km in prev_km_map:
                old_score = prev_km_map[km]
                if score != old_score:
                    changes.append(f"{km}км({old_score}→{score})")
        
        change_str = ", ".join(changes) if changes else "Без изменений"
        data['Прошлый Nуч'] = prev_nuch if prev_nuch is not None else data['Nуч']
        data['Динамика'] = round(data['Nуч'] - data['Прошлый Nуч'], 2)
        data['Изменившиеся км'] = change_str
        
        # Убираем словарь для вывода в таблицу
        output_row = {k: v for k, v in data.items() if k != 'km_map'}
        comparison_results.append(output_row)

    if comparison_results:
        df_final = pd.DataFrame(comparison_results).sort_values('Nуч')

        # --- KPI ---
        st.subheader("📊 Анализ изменений")
        k1, k2, k3 = st.columns(3)
        k1.metric("Средний Nуч (текущий)", f"{df_final['Nуч'].mean():.2f}")
        k2.metric("Улучшилось перегонов", len(df_final[df_final['Динамика'] > 0]))
        k3.metric("Ухудшилось перегонов", len(df_final[df_final['Динамика'] < 0]))

        # --- ГРАФИК ДИНАМИКИ ---
        fig = px.bar(df_final, x='Перегон', y='Динамика', 
                     color='Динамика', color_continuous_scale='RdYlGn',
                     title="Динамика изменения Nуч (выше 0 - стало лучше)")
        st.plotly_chart(fig, use_container_width=True)

        # --- ТАБЛИЦА ---
        st.subheader("📋 Детальный отчет")
        
        def style_custom(row):
            styles = [''] * len(row)
            # Подсвечиваем ячейку с изменениями, если они есть
            if row['Изменившиеся км'] != "Без изменений":
                idx = row.index.get_loc('Изменившиеся км')
                styles[idx] = 'background-color: #e8f4f8; font-weight: bold; color: #005a8d'
            return styles

        st.dataframe(
            df_final.style.format({
                "Nуч": "{:.2f}", 
                "Прошлый Nуч": "{:.2f}", 
                "Динамика": "{:+.2f}"
            })
            .background_gradient(subset=['Nуч'], cmap='RdYlGn')
            .apply(style_custom, axis=1)
            .applymap(lambda x: 'color: green; font-weight: bold' if x > 0 else ('color: red; font-weight: bold' if x < 0 else ''), subset=['Динамика']),
            use_container_width=True
        )

        # --- EXCEL EXPORT ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Анализ')
            workbook = writer.book
            worksheet = writer.sheets['Анализ']
            
            # Автоматическая настройка ширины колонок
            for i, col in enumerate(df_final.columns):
                width = max(len(str(col)), 15)
                if col == 'Изменившиеся км': width = 50
                worksheet.set_column(i, i, width)

        st.download_
