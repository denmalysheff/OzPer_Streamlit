import streamlit as st
import pandas as pd
import io
import os

# 1. Функция исправления заголовков (убирает пробелы и фиксит латиницу)
def fix_headers(df):
    def clean_text(text):
        if not isinstance(text, str): return text
        # Заменяем похожие латинские буквы на русские
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    df.columns = [clean_text(col) for col in df.columns]
    return df

st.set_page_config(page_title="Учет Nуч по перегонам", layout="wide")

st.title("🚂 Расчет оценки состояния пути (Nуч)")
st.markdown("---")

# 2. Поиск базы станций в корне проекта
base_file_name = "stations_base.xlsx"

if os.path.exists(base_file_name):
    try:
        df_base_raw = pd.read_excel(base_file_name)
        df_base = fix_headers(df_base_raw)
        
        # --- ОЧИСТКА БАЗЫ ---
        df_base = df_base.dropna(subset=['КООРДИНАТА'])
        df_base['КООРДИНАТА'] = pd.to_numeric(df_base['КООРДИНАТА'], errors='coerce')
        df_base = df_base.dropna(subset=['КООРДИНАТА', 'НАПРАВЛЕНИЕ'])
        
        st.sidebar.success(f"✅ База станций загружена")
    except Exception as e:
        st.error(f"Ошибка в файле базы: {e}")
        st.stop()
else:
    st.error(f"❌ Файл '{base_file_name}' не найден на GitHub!")
    st.stop()

# 3. Загрузка файла оценки пользователем
file_eval = st.file_uploader("Загрузите файл ОЦЕНКИ (км)", type="xlsx")

if file_eval:
    try:
        # Читаем лист 'Оценка КМ'
        df_eval_raw = pd.read_excel(file_eval, sheet_name='Оценка КМ')
        df_eval = fix_headers(df_eval_raw)

        # --- КРИТИЧЕСКАЯ ОЧИСТКА (Исправление ошибки NaN) ---
        # Удаляем строки, где нет ключевых данных
        df_eval = df_eval.dropna(subset=['КМ', 'ОЦЕНКА', 'КОДНАПР'])
        
        # Принудительно в числа. Все, что не число (пробелы, текст), станет NaN
        df_eval['КМ'] = pd.to_numeric(df_eval['КМ'], errors='coerce')
        df_eval['ОЦЕНКА'] = pd.to_numeric(df_eval['ОЦЕНКА'], errors='coerce')
        df_eval['КОДНАПР'] = pd.to_numeric(df_eval['КОДНАПР'], errors='coerce')
        
        # Удаляем строки, которые превратились в NaN после конвертации
        df_eval = df_eval.dropna(subset=['КМ', 'ОЦЕНКА', 'КОДНАПР'])
        # ---------------------------------------------------

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
                    
                    # Теперь int() сработает гарантированно
                    km_start = int(st_a['КООРДИНАТА']) + 1
                    km_end = int(st_b['КООРДИНАТА'])
                    
                    seg = df_eval[
                        (df_eval['КОДНАПР'] == direction) & 
                        (df_eval['ПУТЬ'] == path) & 
                        (df_eval['КМ'] >= km_start) & 
                        (df_eval['КМ'] <= km_end)
                    ]
                    
                    if not seg.empty:
                        s5 = (seg['ОЦЕНКА'] == 5).sum()
                        s4 = (seg['ОЦЕНКА'] == 4).sum()
                        s3 = (seg['ОЦЕНКА'] == 3).sum()
                        s2 = (seg['ОЦЕНКА'] == 2).sum()
                        all_km = len(seg)
                        
                        n_uch = round((s5*5 + s4*4 + s3*3 - s2*5) / all_km, 2)
                        
                        results.append({
                            'Направление': direction, 'Путь': path,
                            'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                            'КМ нач': km_start, 'КМ кон': km_end,
                            '5 (Отл)': s5, '4 (Хор)': s4, '3 (Удов)': s3, '2 (Неуд)': s2,
                            'Всего КМ': all_km, 'Nуч': n_uch
                        })

        if results:
            df_res = pd.DataFrame(results).sort_values(by='Nуч')
            st.write("### Результаты расчета")
            st.dataframe(df_res.style.background_gradient(subset=['Nуч'], cmap='RdYlGn'), use_container_width=True)

            # Генерация Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False, sheet_name='Результат', startrow=1)
                workbook  = writer.book
                worksheet = writer.sheets['Результат']
                
                # Форматы
                fmt_green  = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1})
                fmt_blue   = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1})
                fmt_orange = workbook.add_format({'bg_color': '#FFEB9C', 'border': 1})
                fmt_red    = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1})

                for row_num in range(2, len(df_res) + 2):
                    val = df_res.iloc[row_num-2]['Nуч']
                    if val > 4: curr_fmt = fmt_green
                    elif 3 < val <= 4: curr_fmt = fmt_blue
                    elif 2.5 < val <= 3: curr_fmt = fmt_orange
                    else: curr_fmt = fmt_red
                    worksheet.set_row(row_num, None, curr_fmt)
                
                for i, col in enumerate(df_res.columns):
                    worksheet.set_column(i, i, 15)

            st.download_button(label="📥 Скачать в Excel", data=output.getvalue(), file_name="N_uch_Report.xlsx")
        else:
            st.warning("Данные по указанным направлениям не найдены.")

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
