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
st.set_page_config(page_title="Мониторинг Nуч", layout="wide")

# --- ОФОРМЛЕНИЕ ---
st.title("🚂 Мониторинг и динамика оценки состояния пути (Nуч)")

if os.path.exists("header.png"):
    st.image("header.png", use_container_width=True)

st.markdown("---")

# 2. База станций
base_file_name = "stations_base.xlsx"
if os.path.exists(base_file_name):
    df_base_raw = pd.read_excel(base_file_name)
    df_base = fix_headers(df_base_raw)
    df_base = df_base.dropna(subset=['КООРДИНАТА', 'НАПРАВЛЕНИЕ'])
    df_base['КООРДИНАТА'] = pd.to_numeric(df_base['КООРДИНАТА'], errors='coerce')
    df_base = df_base.dropna(subset=['КООРДИНАТА'])
else:
    st.error(f"❌ Файл '{base_file_name}' не найден!")
    st.stop()

# 3. Загрузка файлов
col_up1, col_up2 = st.columns(2)
with col_up1:
    file_curr = st.file_uploader("📂 ТЕКУЩИЙ месяц (Excel)", type="xlsx")
with col_up2:
    file_prev = st.file_uploader("📂 ПРОШЛЫЙ месяц (для сравнения)", type="xlsx")

def process_file(file):
    if file is None: return None
    df = pd.read_excel(file, sheet_name='Оценка КМ')
    df = fix_headers(df)
    cols = ['КМ', 'ОЦЕНКА', 'КОДНАПР']
    df = df.dropna(subset=cols)
    for c in cols: df[c] = pd.to_numeric(df[c], errors='coerce')
    return df.dropna(subset=cols)

def calculate_nuch(df_eval, df_base):
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
                    key = f"{direction}_{path}_{st_a['СТАНЦИЯ']}_{st_b['СТАНЦИЯ']}"
                    results[key] = {
                        'Направление': int(direction), 'Путь': int(path),
                        'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                        'Км нач': km_s, 'Км кон': km_e, 'Всего Км': len(seg),
                        'Nуч': n_uch, 'Отл': s5, 'Хор': s4, 'Удов': s3, 'Неуд': s2,
                        'Список Отл км': ", ".join(seg[seg['ОЦЕНКА']==5]['КМ'].astype(int).astype(str)),
                        'Список Хор км': ", ".join(seg[seg['ОЦЕНКА']==4]['КМ'].astype(int).astype(str)),
                        'Список Удов км': ", ".join(seg[seg['ОЦЕНКА']==3]['КМ'].astype(int).astype(str)),
                        'Список Неуд км': ", ".join(seg[seg['ОЦЕНКА']==2]['КМ'].astype(int).astype(str))
                    }
    return results

if file_curr:
    df_c = process_file(file_curr)
    res_c = calculate_nuch(df_c, df_base)
    
    # Сравнение
    final_data = []
    res_p = calculate_nuch(process_file(file_prev), df_base) if file_prev else {}

    for key, data in res_c.items():
        prev_val = res_p.get(key, {}).get('Nуч', None)
        delta = round(data['Nуч'] - prev_val, 2) if prev_val is not None else 0
        data['Динамика'] = delta
        final_data.append(data)

    df_res = pd.DataFrame(final_data).sort_values('Nуч')

    # --- KPI КАРТОЧКИ ---
    avg_nuch = round(df_res['Nуч'].mean(), 2)
    bad_segs = len(df_res[df_res['Nуч'] < 2.5])
    total_km = df_res['Всего Км'].sum()
    
    st.subheader("📈 Общие показатели")
    c1, c2, c3 = st.columns(3)
    c1.metric("Средний Nуч", avg_nuch, delta=round(avg_nuch - pd.DataFrame(res_p.values())['Nуч'].mean(), 2) if res_p else None)
    c2.metric("Неуд. перегоны (Nуч < 2.5)", bad_segs)
    c3.metric("Км в анализе", int(total_km))

    # Таблица
    st.subheader("📊 Детальный расчет по перегонам")
    
    def color_delta(val):
        color = 'green' if val > 0 else 'red' if val < 0 else 'gray'
        return f'color: {color}; font-weight: bold'

    styled_df = df_res.style.format({"Nуч": "{:.2f}", "Динамика": "{:+.2f}"})\
        .applymap(color_delta, subset=['Динамика'])\
        .background_gradient(subset=['Nуч'], cmap='RdYlGn')
    
    st.dataframe(styled_df, use_container_width=True)

    # --- EXCEL ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_res.to_excel(writer, index=False, sheet_name='Анализ', startrow=1)
        workbook, worksheet = writer.book, writer.sheets['Анализ']
        
        # Стили
        f_int = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0'})
        f_float = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.00', 'bold': True})
        f_hdr = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'align': 'center'})
        
        for c_idx, col in enumerate(df_res.columns):
            worksheet.write(1, c_idx, col, f_hdr)
            worksheet.set_column(c_idx, c_idx, 15 if "Список" not in col else 30)

        for r_idx in range(len(df_res)):
            row = r_idx + 2
            for c_idx, col in enumerate(df_res.columns):
                val = df_res.iloc[r_idx][col]
                fmt = f_float if col in ['Nуч', 'Динамика'] else f_int
                worksheet.write(row, c_idx, val, fmt)

    st.download_button("📥 Скачать полный отчет", output.getvalue(), "Nuch_Full_Report.xlsx")
