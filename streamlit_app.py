import streamlit as st
import pandas as pd
import io
import os

# 1. Функция исправления заголовков
def fix_headers(df):
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

# Настройка страницы
st.set_page_config(page_title="Мониторинг Nуч: Динамика", layout="wide")

# --- ОФОРМЛЕНИЕ ---
st.title("🚂 Сравнительный анализ Nуч: Текущий vs Прошлый проезд")

if os.path.exists("header.png"):
    st.image("header.png", use_container_width=True)

st.markdown("---")

# 2. Загрузка базы станций
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

# 3. Интерфейс загрузки двух файлов
col_up1, col_up2 = st.columns(2)
with col_up1:
    file_prev = st.file_uploader("📂 ПРОШЛЫЙ месяц (База для сравнения)", type="xlsx")
with col_up2:
    file_curr = st.file_uploader("📂 ТЕКУЩИЙ месяц (Результат)", type="xlsx")

def process_excel_data(file):
    if file is None: return None
    try:
        df = pd.read_excel(file, sheet_name='Оценка КМ')
        df = fix_headers(df)
        cols = ['КМ', 'ОЦЕНКА', 'КОДНАПР']
        df = df.dropna(subset=cols)
        for c in cols:
            df[c] = pd.to_numeric(df[c], errors='coerce')
        return df.dropna(subset=cols)
    except:
        return None

def get_nuch_results(df_eval, df_base):
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
                    
                    # Уникальный ключ для сопоставления перегона
                    key = f"{direction}_{path}_{st_a['СТАНЦИЯ']}_{st_b['СТАНЦИЯ']}"
                    
                    results[key] = {
                        'Направление': int(direction), 'Путь': int(path),
                        'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                        'Км нач': int(km_s), 'Км кон': int(km_e), 'Всего Км': int(len(seg)),
                        'Nуч': n_uch, 'Отл': int(s5), 'Хор': int(s4), 'Удов': int(s3), 'Неуд': int(s2),
                        'Список Отл км': ", ".join(seg[seg['ОЦЕНКА']==5]['КМ'].astype(int).astype(str)),
                        'Список Хор км': ", ".join(seg[seg['ОЦЕНКА']==4]['КМ'].astype(int).astype(str)),
                        'Список Удов км': ", ".join(seg[seg['ОЦЕНКА']==3]['КМ'].astype(int).astype(str)),
                        'Список Неуд км': ", ".join(seg[seg['ОЦЕНКА']==2]['КМ'].astype(int).astype(str))
                    }
    return results

if file_curr:
    df_curr_data = process_excel_data(file_curr)
    res_curr = get_nuch_results(df_curr_data, df_base)
    
    # Если загружен прошлый месяц, делаем сопоставление
    df_prev_data = process_excel_data(file_prev)
    res_prev = get_nuch_results(df_prev_data, df_base) if file_prev else {}

    comparison_results = []
    for key, data in res_curr.items():
        prev_nuch = res_prev.get(key, {}).get('Nуч', None)
        
        if prev_nuch is not None:
            delta = round(data['Nуч'] - prev_nuch, 2)
        else:
            delta = 0.0
            
        data['Прошлый Nуч'] = prev_nuch if prev_nuch is not None else data['Nуч']
        data['Динамика'] = delta
        comparison_results.append(data)

    df_final = pd.DataFrame(comparison_results).sort_values('Nуч')

    # --- KPI КАРТОЧКИ ---
    st.subheader("📈 Итоги проезда")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    
    avg_curr = df_final['Nуч'].mean()
    avg_prev = df_final['Прошлый Nуч'].mean()
    delta_total = avg_curr - avg_prev

    kpi1.metric("Средний Nуч (Тек)", f"{avg_curr:.2f}", delta=f"{delta_total:+.2f}")
    kpi2.metric("Кол-во Неуд км", df_final['Неуд'].sum())
    kpi3.metric("Перегонов в работе", len(df_final))
    kpi4.metric("Всего Км", df_final['Всего Км'].sum())

    # --- ТАБЛИЦА В БРАУЗЕРЕ ---
    st.subheader("📊 Детальная таблица сравнения")
    
    def style_delta(val):
        color = 'green' if val > 0 else 'red' if val < 0 else 'black'
        return f'color: {color}; font-weight: bold'

    styled_res = df_final.style.format({
        "Nуч": "{:.2f}", "Прошлый Nуч": "{:.2f}", "Динамика": "{:+.2f}"
    }).applymap(style_delta, subset=['Динамика']).background_gradient(subset=['Nуч'], cmap='RdYlGn')

    st.dataframe(styled_res, use_container_width=True)

    # --- ГЕНЕРАЦИЯ EXCEL ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Анализ Динамики', startrow=1)
        workbook  = writer.book
        worksheet = writer.sheets['Анализ Динамики']
        
        # Форматы
        f_int = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0'})
        f_float = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.00'})
        f_bold_float = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.00', 'bold': True})
        f_hdr = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'align': 'center'})
        
        # Заголовки
        for c_idx, col in enumerate(df_final.columns):
            worksheet.write(1, c_idx, col, f_hdr)
        
        # Данные
        for r_idx in range(len(df_final)):
            row = r_idx + 2
            for c_idx, col in enumerate(df_final.columns):
                val = df_final.iloc[r_idx][col]
                
                if col in ['Nуч', 'Прошлый Nуч', 'Динамика']:
                    worksheet.write(row, c_idx, val, f_bold_float)
                elif any(x in col for x in ['Направление', 'Путь', 'Км', 'Отл', 'Хор', 'Удов', 'Неуд']):
                    try:
                        worksheet.write(row, c_idx, int(val), f_int)
                    except:
                        worksheet.write(row, c_idx, val, f_int)
                else:
                    worksheet.write(row, c_idx, val, f_int)

        for i, col in enumerate(df_final.columns):
            worksheet.set_column(i, i, 40 if "Список" in col else 12)

    st.download_button("📥 Скачать сравнительный отчет (Excel)", output.getvalue(), "Nuch_Dynamics_Report.xlsx")

else:
    st.info("💡 Пожалуйста, загрузите текущий файл оценки для начала расчета.")
