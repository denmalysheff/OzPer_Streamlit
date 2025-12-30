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

st.set_page_config(page_title="Учет Nуч по перегонам", layout="wide")

st.title("🚂 Расчет оценки состояния пути (Nуч)")
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
        st.sidebar.success(f"✅ База станций подключена")
    except Exception as e:
        st.error(f"Ошибка в файле базы: {e}")
        st.stop()
else:
    st.error(f"❌ Файл '{base_file_name}' не найден на GitHub!")
    st.stop()

# 3. Загрузка файла оценки
file_eval = st.file_uploader("Загрузите файл ОЦЕНКИ (км)", type="xlsx")

if file_eval:
    try:
        df_eval_raw = pd.read_excel(file_eval, sheet_name='Оценка КМ')
        df_eval = fix_headers(df_eval_raw)

        # Очистка данных от NaN
        df_eval = df_eval.dropna(subset=['КМ', 'ОЦЕНКА', 'КОДНАПР'])
        df_eval['КМ'] = pd.to_numeric(df_eval['КМ'], errors='coerce')
        df_eval['ОЦЕНКА'] = pd.to_numeric(df_eval['ОЦЕНКА'], errors='coerce')
        df_eval['КОДНАПР'] = pd.to_numeric(df_eval['КОДНАПР'], errors='coerce')
        df_eval = df_eval.dropna(subset=['КМ', 'ОЦЕНКА', 'КОДНАПР'])

        results = []
        valid_dirs = {24602, 24603, 24701}
        
        for direction in df_base['НАПРАВЛЕНИЕ'].unique():
            if direction not in valid_dirs:
                continue
            
            stations = df_base[df_base['НАПРАВЛЕНИЕ'] == direction].sort_values('КООРДИНАТА')
            paths = df_eval[df_eval['КОДНАПР'] == direction]['ПУТЬ'].unique()

            for path in paths:
                for i in range(len(stations) - 1):
                    st_a = stations.iloc[i]
                    st_b = stations.iloc[i+1]
                    
                    km_start = int(st_a['КООРДИНАТА']) + 1
                    km_end = int(st_b['КООРДИНАТА'])
                    
                    seg = df_eval[
                        (df_eval['КОДНАПР'] == direction) & 
                        (df_eval['ПУТЬ'] == path) & 
                        (df_eval['КМ'] >= km_start) & 
                        (df_eval['КМ'] <= km_end)
                    ]
                    
                    if not seg.empty:
                        s5 = int((seg['ОЦЕНКА'] == 5).sum())
                        s4 = int((seg['ОЦЕНКА'] == 4).sum())
                        s3 = int((seg['ОЦЕНКА'] == 3).sum())
                        s2 = int((seg['ОЦЕНКА'] == 2).sum())
                        all_km = len(seg)
                        
                        # Расчет Nуч с принудительным округлением
                        n_uch_val = (s5*5 + s4*4 + s3*3 - s2*5) / all_km
                        n_uch = round(float(n_uch_val), 2)
                        
                        # Список КМ с оценкой 2
                        neud_list = seg[seg['ОЦЕНКА'] == 2]['КМ'].astype(int).astype(str).tolist()
                        neud_str = ", ".join(neud_list)
                        
                        # Собираем данные в строго заданном порядке столбцов
                        results.append({
                            'Направление': direction,
                            'Путь': path,
                            'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                            'Км нач': km_start,
                            'Км кон': km_end,
                            'Всего Км': all_km,
                            'Nуч': n_uch,
                            'Отл': s5,
                            'Хор': s4,
                            'Удов': s3,
                            'Неуд': s2,
                            'Список Неуд км': neud_str
                        })

        if results:
            df_res = pd.DataFrame(results).sort_values(by='Nуч', ascending=True)
            
            st.subheader("📊 Результаты расчета")
            
            # Принудительное форматирование отображения в браузере (3.66)
            try:
                st.dataframe(
                    df_res.style.format({"Nуч": "{:.2f}"})
                    .background_gradient(subset=['Nуч'], cmap='RdYlGn'), 
                    use_container_width=True
                )
            except:
                st.dataframe(df_res, use_container_width=True)

            # --- ГЕНЕРАЦИЯ EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False, sheet_name='Результат', startrow=1)
                workbook  = writer.book
                worksheet = writer.sheets['Результат']
                
                # Стили Excel
                fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
                fmt_num    = workbook.add_format({'num_format': '0.00', 'border': 1}) # Формат для Nуч
                fmt_red    = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1, 'num_format': '0.00'})
                fmt_orange = workbook.add_format({'bg_color': '#FFEB9C', 'border': 1, 'num_format': '0.00'})
                fmt_blue   = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1, 'num_format': '0.00'})
                fmt_green  = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1, 'num_format': '0.00'})

                worksheet.merge_range(0, 0, 0, len(df_res.columns)-1, "Отчет по Nуч по перегонам", fmt_header)

                for row_num in range(2, len(df_res) + 2):
                    val = df_res.iloc[row_num-2]['Nуч']
                    if val > 4: curr_fmt = fmt_green
                    elif 3 < val <= 4: curr_fmt = fmt_blue
                    elif 2.5 < val <= 3: curr_fmt = fmt_orange
                    else: curr_fmt = fmt_red
                    
                    # Применяем формат ко всей строке
                    worksheet.set_row(row_num, None, curr_fmt)

                # Ширина колонок
                for i, col in enumerate(df_res.columns):
                    w = 30 if col == 'Список Неуд км' else 15
                    worksheet.set_column(i, i, w)

            st.download_button(
                label="📥 Скачать отчет в Excel",
                data=output.getvalue(),
                file_name="Nuch_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Данные не найдены.")

    except Exception as e:
        st.error(f"❌ Ошибка: {e}")
