import streamlit as st
import pandas as pd
import io
import os


# Функция исправления заголовков
def fix_headers(df):
    def clean_text(text):
        if not isinstance(text, str): return text
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)

    df.columns = [clean_text(col) for col in df.columns]
    return df


st.set_page_config(page_title="Учет Nуч", layout="wide")
st.title("🚂 Расчет Nуч по перегонам")

# 1. Автоматическая загрузка базы станций
base_file_path = "stations_base.xlsx"  # Файл должен лежать в той же папке на GitHub

if os.path.exists(base_file_path):
    df_base = fix_headers(pd.read_excel(base_file_path))
    st.info(f"✅ База станций подключена (Найдено станций: {len(df_base)})")
else:
    st.error("❌ Файл 'stations_base.xlsx' не найден в корне проекта!")
    st.stop()

# 2. Поле для загрузки файла оценки пользователем
file_eval = st.file_uploader("Загрузите файл ОЦЕНКИ (км)", type="xlsx")

if file_eval:
    try:
        # Читаем данные
        df_eval = fix_headers(pd.read_excel(file_eval, sheet_name='Оценка КМ'))

        # Очистка данных
        df_eval = df_eval.dropna(subset=['КМ', 'ОЦЕНКА'])
        df_eval['КМ'] = pd.to_numeric(df_eval['КМ'], errors='coerce')
        df_eval['ОЦЕНКА'] = pd.to_numeric(df_eval['ОЦЕНКА'], errors='coerce')
        df_base['КООРДИНАТА'] = pd.to_numeric(df_base['КООРДИНАТА'], errors='coerce')

        results = []
        valid_dirs = {24602, 24603, 24701}

        # Логика расчета
        for direction in df_base['НАПРАВЛЕНИЕ'].unique():
            if direction not in valid_dirs: continue
            stations = df_base[df_base['НАПРАВЛЕНИЕ'] == direction].sort_values('КООРДИНАТА')
            paths = df_eval[df_eval['КОДНАПР'] == direction]['ПУТЬ'].unique()

            for path in paths:
                for i in range(len(stations) - 1):
                    st_a, st_b = stations.iloc[i], stations.iloc[i + 1]
                    km_s, km_e = int(st_a['КООРДИНАТА']) + 1, int(st_b['КООРДИНАТА'])

                    seg = df_eval[(df_eval['КОДНАПР'] == direction) & (df_eval['ПУТЬ'] == path) &
                                  (df_eval['КМ'] >= km_s) & (df_eval['КМ'] <= km_e)]

                    if not seg.empty:
                        s5, s4, s3, s2 = (seg['ОЦЕНКА'] == 5).sum(), (seg['ОЦЕНКА'] == 4).sum(), \
                            (seg['ОЦЕНКА'] == 3).sum(), (seg['ОЦЕНКА'] == 2).sum()
                        all_km = len(seg)
                        n_uch = round((s5 * 5 + s4 * 4 + s3 * 3 - s2 * 5) / all_km, 2)

                        results.append({
                            'Направление': direction, 'Путь': path,
                            'Перегон': f"{st_a['СТАНЦИЯ']} - {st_b['СТАНЦИЯ']}",
                            'КМ нач': km_s, 'КМ кон': km_e,
                            '5 (Отл)': s5, '4 (Хор)': s4, '3 (Удов)': s3, '2 (Неуд)': s2,
                            'Всего КМ': all_km, 'Nуч': n_uch
                        })

        if results:
            df_res = pd.DataFrame(results).sort_values(by='Nуч')
            st.write("### Предварительный просмотр результата:")
            st.dataframe(df_res.style.background_gradient(subset=['Nуч'], cmap='RdYlGn'), use_container_width=True)

            # Генерация Excel для скачивания
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False, sheet_name='Результат', startrow=1)
                # (Здесь остается ваш код форматирования и раскраски Excel из прошлых сообщений)
                writer.close()

            st.download_button(
                label="📥 Скачать отчет Nуч в Excel",
                data=output.getvalue(),
                file_name="Nуч_по_перегонам_.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")