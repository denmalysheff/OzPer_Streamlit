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
st.set_page_config(page_title="Учет Nуч по перегонам", layout="wide")

# --- ОФОРМЛЕНИЕ ---
st.title("🚂 Расчет балловой оценки состояния пути (Nуч)")

if os.path.exists("header.png"):
    st.image("header.png", use_container_width=True)
else:
    st.info("💡 Загрузите файл 'header.png' для отображения баннера.")

st.markdown("---")

# 2. Поиск базы станций
base_file_name = "stations_base.xlsx"
if os.path.exists(base_file_name):
    try:
        df_base_raw = pd.read_excel(base_file_name)
        df_base = fix_headers(df_base_raw)
        df_base = df_base.dropna(subset=['КООРДИНАТА', 'НАПРАВЛЕНИЕ'])
        df_base['КООРДИНАТА'] = pd.to_numeric(df_base['КООРДИНАТА'], errors='coerce')
        df_base = df_base.dropna(subset=['КООРДИНАТА'])
        st.sidebar.success("✅ База станций подключена")
    except Exception as e:
        st.error(f"Ошибка в файле базы: {e}")
        st.stop()
else:
    st.error(f"❌ Файл '{base_file_name}' не найден!")
    st.stop()

# 3. Загрузка файла оценки
file_eval = st.file_uploader("Загрузите файл ОЦЕНКИ (лист 'Оценка КМ')", type="xlsx")

if file_eval:
    try:
        df_eval_raw = pd.read_excel(file_eval, sheet_name='Оценка КМ')
        df_eval = fix_headers(df_eval_raw)

        # Очистка данных
        cols_to_check = ['КМ', 'ОЦЕНКА', 'КОДНАПР']
        df_eval = df_eval.dropna(subset=cols_to_check)
        for col in cols_to_check:
            df_eval[col] = pd.to_numeric(df_eval[col], errors='coerce')
        df_eval = df_eval.dropna(subset=cols_to_check)

        results = []
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
                        all_km = len(seg)
                        
                        n_uch_val = (s5*5 + s4*4 + s3*3 - s2*5) / all_km
                        n_uch = round(float(n_uch_val), 2)
                        
                        neud_list = seg[seg['ОЦЕНКА'] == 2]['КМ'].astype(int).astype(str).tolist()
                        neud_str = ", ".join(neud_list)
                        
                        results.append({
                            'Направление': int(direction),
                            'Путь': int(path),
                            'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                            'Км нач': int(km_s),
                            'Км кон': int(km_e),
                            'Всего Км': int(all_km),
                            'Nуч': n_uch,
                            'Отл': int(s5), 'Хор': int(s4), 'Удов': int(s3), 'Неуд': int(s2),
                            'Список Неуд км': neud_str
                        })

        if results:
            df_res = pd.DataFrame(results).sort_values(by='Nуч', ascending=True)
            
            # Приведение всех "числовых" колонок к INT для корректного отображения
            int_cols = ['Направление', 'Путь', 'Км нач', 'Км кон', 'Всего Км', 'Отл', 'Хор', 'Удов', 'Неуд']
            for c in int_cols:
                if c in df_res.columns:
                    df_res[c] = df_res[c].astype(int)

            st.subheader("📊 Результаты расчета")
            
            # Форматирование в браузере (Nуч - 2 знака, остальное - целое)
            styled_df = df_res.style.format({
                "Nуч": "{:.2f}",
                "Направление": "{:d}", "Путь": "{:d}", "Км нач": "{:d}", 
                "Км кон": "{:d}", "Всего Км": "{:d}", "Отл": "{:d}", 
                "Хор": "{:d}", "Удов": "{:d}", "Неуд": "{:d}"
            })

            # Пытаемся применить градиент, если matplotlib установлен
            try:
                st.dataframe(styled_df.background_gradient(subset=['Nуч'], cmap='RdYlGn'), use_container_width=True)
            except ImportError:
                st.warning("⚠️ Для цветной подсветки таблицы выполните: pip install matplotlib")
                st.dataframe(styled_df, use_container_width=True)

            # --- ГЕНЕРАЦИЯ EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False, sheet_name='Результат', startrow=1)
                workbook  = writer.book
                worksheet = writer.sheets['Результат']
                
                fmt_int = '0'
                fmt_float = '0.00'
                base = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
                
                # Цвета для Excel
                styles = {
                    'green':  [workbook.add_format({**base, 'bg_color': '#C6EFCE', 'num_format': fmt_int}),
                               workbook.add_format({**base, 'bg_color': '#C6EFCE', 'num_format': fmt_float, 'bold': True})],
                    'blue':   [workbook.add_format({**base, 'bg_color': '#DDEBF7', 'num_format': fmt_int}),
                               workbook.add_format({**base, 'bg_color': '#DDEBF7', 'num_format': fmt_float, 'bold': True})],
                    'orange': [workbook.add_format({**base, 'bg_color': '#FFEB9C', 'num_format': fmt_int}),
                               workbook.add_format({**base, 'bg_color': '#FFEB9C', 'num_format': fmt_float, 'bold': True})],
                    'red':    [workbook.add_format({**base, 'bg_color': '#FFC7CE', 'num_format': fmt_int}),
                               workbook.add_format({**base, 'bg_color': '#FFC7CE', 'num_format': fmt_float, 'bold': True})]
                }
                
                fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#F2F2F2'})
                worksheet.merge_range(0, 0, 0, len(df_res.columns)-1, "Отчет по балловой оценке состояния пути", fmt_header)

                n_uch_idx = df_res.columns.get_loc('Nуч')

                for r_idx in range(len(df_res)):
                    val = df_res.iloc[r_idx]['Nуч']
                    row_num = r_idx + 2
                    
                    if val > 4: key = 'green'
                    elif 3 < val <= 4: key = 'blue'
                    elif 2.5 < val <= 3: key = 'orange'
                    else: key = 'red'
                    
                    st_i, st_f = styles[key]
                    
                    # Записываем ячейки: Nуч дробно, остальное целым
                    for c_idx, col_name in enumerate(df_res.columns):
                        cell_val = df_res.iloc[r_idx][col_name]
                        if col_name == 'Nуч':
                            worksheet.write(row_num, c_idx, cell_val, st_f)
                        elif col_name == 'Список Неуд км' or col_name == 'Перегон':
                            worksheet.write(row_num, c_idx, cell_val, st_i) # Для текста формат игнорируется
                        else:
                            worksheet.write(row_num, c_idx, int(cell_val), st_i)

                for i, col in enumerate(df_res.columns):
                    worksheet.set_column(i, i, 40 if col == 'Список Неуд км' else 12)

            st.download_button(label="📥 Скачать Excel", data=output.getvalue(), 
                               file_name="Nuch_Report.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ Совпадений не найдено.")
    except Exception as e:
        st.error(f"❌ Ошибка: {e}")
