import streamlit as st
import pandas as pd
import io
import os
import plotly.express as px

# --- 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def fix_headers(df):
    """Исправляет заголовки таблицы (кириллица/латиница, регистр)."""
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

def find_sheet(xl, target_name):
    """Ищет лист в Excel-файле."""
    target_cleaned = target_name.replace(" ", "").upper()
    for sheet in xl.sheet_names:
        if sheet.replace(" ", "").upper() == target_cleaned:
            return sheet
    return None

# --- 2. НАСТРОЙКА ИНТЕРФЕЙСА ---
st.set_page_config(page_title="Мониторинг Nуч + Целостность", layout="wide")

if os.path.exists("header.png"):
    st.image("header.png", use_container_width=True)

st.title("🚂 Анализ Nуч и проверка целостности данных")

# --- 3. ЗАГРУЗКА БАЗОВЫХ ДАННЫХ (СТАНЦИИ И СТРУКТУРА ПД) ---

@st.cache_data
def load_base_files():
    # 1. Загрузка базы станций
    base_file = "stations_base.xlsx"
    if not os.path.exists(base_file):
        st.error(f"❌ Файл '{base_file}' не найден!")
        st.stop()
    df_base = fix_headers(pd.read_excel(base_file))
    
    # 2. Загрузка структуры ПД (административная структура)
    struct_file = "adm_struktur.xlsx"
    if not os.path.exists(struct_file):
        st.error(f"❌ Файл '{struct_file}' не найден! (Нужен для проверки целостности)")
        st.stop()
    df_struct = fix_headers(pd.read_excel(struct_file))
    
    # Приведение типов для структуры
    struct_cols = ['НАПРАВЛЕНИЕ', 'ПУТЬ', 'КМ НАЧАЛА', 'КМ КОНЦА']
    for col in struct_cols:
        if col in df_struct.columns:
            df_struct[col] = pd.to_numeric(df_struct[col], errors='coerce')
    
    return df_base.dropna(subset=['КООРДИНАТА']), df_struct.dropna(subset=struct_cols)

df_base, df_struct = load_base_files()

# --- 4. ЗАГРУЗКА ПОЛЬЗОВАТЕЛЬСКИХ ФАЙЛОВ ---

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_prev = st.file_uploader("📂 Шаг 1: ПРОШЛЫЙ месяц", type="xlsx")
with col_up2:
    file_curr = st.file_uploader("📂 Шаг 2: ТЕКУЩИЙ месяц", type="xlsx")

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
        for c in cols: df[c] = pd.to_numeric(df[c], errors='coerce')
        return df.dropna(subset=cols)
    except Exception as e:
        st.error(f"Ошибка чтения {file.name}: {e}")
        return None

# --- 5. БЛОК ПРОВЕРКИ ЦЕЛОСТНОСТИ ---

def check_integrity(df_eval, df_struct):
    """Сравнивает наличие КМ в загруженном файле со справочником ПД."""
    missing_report = []
    
    # Итерируемся по участкам ПД (Линейным участкам)
    for _, row in df_struct.iterrows():
        dir_id = row['НАПРАВЛЕНИЕ']
        path_id = row['ПУТЬ']
        km_start = int(row['КМ НАЧАЛА'])
        km_end = int(row['КМ КОНЦА'])
        pd_name = row.get('ЛИНЕЙНЫЙ УЧАСТОК (ПД)', f"ПД-{_}")
        
        # Создаем множество эталонных километров для этого участка
        required_kms = set(range(km_start, km_end + 1))
        
        # Находим фактически присутствующие км в данных
        actual_kms = set(df_eval[
            (df_eval['КОДНАПР'] == dir_id) & 
            (df_eval['ПУТЬ'] == path_id) & 
            (df_eval['КМ'] >= km_start) & 
            (df_eval['КМ'] <= km_end)
        ]['КМ'].astype(int).unique())
        
        missing = required_kms - actual_kms
        
        if missing:
            missing_report.append({
                "Линейный участок": pd_name,
                "Направление": dir_id,
                "Путь": path_id,
                "Всего км": len(required_kms),
                "Пропущено": len(missing),
                "Список пропусков": ", ".join(map(str, sorted(list(missing))))
            })
            
    return pd.DataFrame(missing_report)

# --- 6. ОСНОВНОЙ РАСЧЕТ Nуч ---

def get_detailed_results(df_eval, df_base):
    if df_eval is None: return {}
    results = {}
    valid_dirs = set(df_base['НАПРАВЛЕНИЕ'].unique())
    
    for direction in valid_dirs:
        stations = df_base[df_base['НАПРАВЛЕНИЕ'] == direction].sort_values('КООРДИНАТА')
        paths = df_eval[df_eval['КОДНАПР'] == direction]['ПУТЬ'].unique()
        
        for path in paths:
            for i in range(len(stations) - 1):
                st_a, st_b = stations.iloc[i], stations.iloc[i+1]
                km_s, km_e = int(st_a['КООРДИНАТА']) + 1, int(st_b['КООРДИНАТА'])
                
                seg = df_eval[(df_eval['КОДНАПР'] == direction) & 
                              (df_eval['ПУТЬ'] == path) & 
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

# --- 7. ВЫВОД РЕЗУЛЬТАТОВ ---

if file_curr:
    df_curr_raw = process_excel_data(file_curr)
    
    # 7.1 ПРОВЕРКА ЦЕЛОСТНОСТИ (Выводим первой)
    st.subheader("⚠️ Проверка полноты данных (на основе adm_struktur)")
    if df_curr_raw is not None:
        df_integrity = check_integrity(df_curr_raw, df_struct)
        if not df_integrity.empty:
            st.error(f"Обнаружены пропуски километров на {len(df_integrity)} участках ПД!")
            st.dataframe(df_integrity, use_container_width=True)
        else:
            st.success("✅ Все километры согласно структуре ПД присутствуют в файле.")

    # 7.2 РАСЧЕТ ДИНАМИКИ
    res_curr = get_detailed_results(df_curr_raw, df_base)
    res_prev = get_detailed_results(process_excel_data(file_prev), df_base) if file_prev else {}

    comparison = []
    for key, data in res_curr.items():
        prev = res_prev.get(key, {})
        data['Прошлый Nуч'] = prev.get('Nуч', data['Nуч'])
        data['Динамика'] = round(data['Nуч'] - data['Прошлый Nуч'], 2)
        
        curr_map = data.pop('km_map', {})
        prev_map = prev.get('km_map', {})
        changes = [f"{k}км({prev_map[k]}→{v})" for k, v in curr_map.items() if k in prev_map and v != prev_map[k]]
        
        data['Изменившиеся км'] = ", ".join(changes) if changes else "Без изменений"
        comparison.append(data)

    if comparison:
        df_final = pd.DataFrame(comparison).sort_values('Nуч')
        
        # График
        st.subheader("📈 Динамика изменения Nуч")
        fig = px.bar(df_final, x='Перегон', y='Динамика', color='Динамика', 
                     color_continuous_scale='RdYlGn', hover_data=['Направление', 'Путь'])
        st.plotly_chart(fig, use_container_width=True)

        # Основная таблица
        st.subheader("📋 Детальный отчет")
        def color_dyn(val):
            if isinstance(val, (int, float)):
                if val > 0: return 'color: #008000; font-weight: bold'
                if val < 0: return 'color: #FF0000; font-weight: bold'
            return ''

        st.dataframe(
            df_final.style.format({"Nуч": "{:.2f}", "Прошлый Nуч": "{:.2f}", "Динамика": "{:+.2f}"})
            .background_gradient(subset=['Nуч'], cmap='RdYlGn')
            .map(color_dyn, subset=['Динамика']),
            use_container_width=True
        )

        # Экспорт
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Анализ')
        st.download_button("📥 Скачать Excel отчет", output.getvalue(), "Analiz_Nuch.xlsx")
else:
    st.info("Загрузите файлы для начала анализа.")
