import streamlit as st
import pandas as pd
import io
import os
import plotly.express as px

# --- 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def fix_headers(df):
    """
    Исправляет заголовки таблицы.
    Иногда в Excel 'КМ' написано английскими буквами, иногда русскими.
    Функция переводит всё в верхний регистр и меняет латиницу на кириллицу.
    """
    def clean_text(text):
        if not isinstance(text, str): return text
        # Словарь замен: ключ - латиница, значение - кириллица
        trans = str.maketrans("KMABOCPETX", "КМАВОСРЕТХ")
        return text.strip().upper().translate(trans)
    
    df.columns = [clean_text(col) for col in df.columns]
    return df

def find_sheet(xl, target_name):
    """
    Ищет лист в Excel-файле, игнорируя пробелы и регистр.
    Например, найдет и 'Оценка КМ', и 'оценкакм'.
    """
    target_cleaned = target_name.replace(" ", "").upper()
    for sheet in xl.sheet_names:
        if sheet.replace(" ", "").upper() == target_cleaned:
            return sheet
    return None

# --- 2. НАСТРОЙКА ИНТЕРФЕЙСА ---

# Устанавливаем широкое отображение страницы
st.set_page_config(page_title="Мониторинг Nуч", layout="wide")

# Проверяем наличие логотипа и выводим его
if os.path.exists("header.png"):
    st.image("header.png", use_container_width=True)

st.title("🚂 Сравнительный анализ Nуч по километрам")
st.info("Загрузите файлы, чтобы увидеть динамику изменений состояния пути.")

# --- 3. ЗАГРУЗКА БАЗОВЫХ ДАННЫХ ---

base_file_name = "stations_base.xlsx"
if os.path.exists(base_file_name):
    try:
        # Читаем базу станций (координаты начала и конца перегонов)
        df_base_raw = pd.read_excel(base_file_name)
        df_base = fix_headers(df_base_raw)
        # Очищаем данные от пустых строк в важных колонках
        df_base = df_base.dropna(subset=['КООРДИНАТА', 'НАПРАВЛЕНИЕ'])
        df_base['КООРДИНАТА'] = pd.to_numeric(df_base['КООРДИНАТА'], errors='coerce')
        df_base = df_base.dropna(subset=['КООРДИНАТА'])
    except Exception as e:
        st.error(f"Ошибка в файле базы станций: {e}")
        st.stop()
else:
    st.error(f"❌ Файл '{base_file_name}' не найден в папке с программой!")
    st.stop()

# --- 4. ЗАГРУЗКА ПОЛЬЗОВАТЕЛЬСКИХ ФАЙЛОВ ---

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_prev = st.file_uploader("📂 Шаг 1: Загрузите ПРОШЛЫЙ месяц", type="xlsx")
with col_up2:
    file_curr = st.file_uploader("📂 Шаг 2: Загрузите ТЕКУЩИЙ месяц", type="xlsx")

def process_excel_data(file):
    """Функция для извлечения данных из загруженного Excel"""
    if file is None: return None
    try:
        xl = pd.ExcelFile(file)
        sheet = find_sheet(xl, "Оценка КМ")
        if not sheet:
            st.warning(f"Лист 'Оценка КМ' не найден в {file.name}")
            return None
        
        df = pd.read_excel(file, sheet_name=sheet)
        df = fix_headers(df)
        
        # Оставляем только нужные колонки для расчета
        cols = ['КМ', 'ОЦЕНКА', 'КОДНАПР', 'ПУТЬ']
        df = df.dropna(subset=cols)
        for c in cols:
            df[c] = pd.to_numeric(df[c], errors='coerce')
        return df.dropna(subset=cols)
    except Exception as e:
        st.error(f"Не удалось прочитать данные: {e}")
        return None

# --- 5. ГЛАВНЫЙ АЛГОРИТМ РАСЧЕТА ---

def get_detailed_results(df_eval, df_base):
    """Распределяет километры по перегонам и считает Nуч"""
    if df_eval is None: return {}
    results = {}
    # Коды направлений, которые мы анализируем
    valid_dirs = {24602, 24603, 24701}
    
    for direction in df_base['НАПРАВЛЕНИЕ'].unique():
        if direction not in valid_dirs: continue
        
        # Получаем список станций для конкретного направления
        stations = df_base[df_base['НАПРАВЛЕНИЕ'] == direction].sort_values('КООРДИНАТА')
        # Определяем, какие пути есть в данных (1-й, 2-й и т.д.)
        paths = df_eval[df_eval['КОДНАПР'] == direction]['ПУТЬ'].unique()
        
        for path in paths:
            # Идем по парам станций (станция А и станция Б = перегон)
            for i in range(len(stations) - 1):
                st_a, st_b = stations.iloc[i], stations.iloc[i+1]
                km_s, km_e = int(st_a['КООРДИНАТА']) + 1, int(st_b['КООРДИНАТА'])
                
                # Фильтруем километры, попавшие в границы этого перегона
                seg = df_eval[(df_eval['КОДНАПР'] == direction) & 
                              (df_eval['ПУТЬ'] == path) & 
                              (df_eval['КМ'] >= km_s) & (df_eval['КМ'] <= km_e)]
                
                if not seg.empty:
                    # Считаем количество каждой оценки
                    s5, s4, s3, s2 = (seg['ОЦЕНКА']==5).sum(), (seg['ОЦЕНКА']==4).sum(), \
                                     (seg['ОЦЕНКА']==3).sum(), (seg['ОЦЕНКА']==2).sum()
                    
                    # Формула расчета Nуч (балловая оценка участка)
                    n_uch = round((s5*5 + s4*4 + s3*3 - s2*5) / len(seg), 2)
                    
                    # Сохраняем "карту" оценок по километрам для этого перегона
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

# --- 6. ОБРАБОТКА И ВЫВОД РЕЗУЛЬТАТОВ ---

if file_curr:
    # Обрабатываем текущий месяц
    df_c_data = process_excel_data(file_curr)
    res_curr = get_detailed_results(df_c_data, df_base)
    
    # Обрабатываем прошлый месяц (если загружен)
    res_prev = get_detailed_results(process_excel_data(file_prev), df_base) if file_prev else {}

    comparison = []
    for key, data in res_curr.items():
        prev = res_prev.get(key, {})
        # Если данных за прошлый месяц нет, считаем динамику от текущего (0)
        data['Прошлый Nуч'] = prev.get('Nуч', data['Nуч'])
        data['Динамика'] = round(data['Nуч'] - data['Прошлый Nуч'], 2)
        
        # Сравниваем оценки по каждому километру (аудит изменений)
        changes = []
        curr_map = data.pop('km_map', {}) # Извлекаем карту и удаляем из словаря для таблицы
        prev_map = prev.get('km_map', {})
        
        for km, score in curr_map.items():
            if km in prev_map and score != prev_map[km]:
                changes.append(f"{km}км({prev_map[km]}→{score})")
        
        data['Изменившиеся км'] = ", ".join(changes) if changes else "Без изменений"
        comparison.append(data)

    if comparison:
        df_final = pd.DataFrame(comparison).sort_values('Nуч')

        # Отображаем график динамики
        st.subheader("📈 График изменений (Динамика)")
        fig = px.bar(df_final, x='Перегон', y='Динамика', color='Динамика', 
                     color_continuous_scale='RdYlGn', 
                     hover_data=['Направление', 'Путь'])
        st.plotly_chart(fig, use_container_width=True)

        # Функция для раскраски текста динамики (зеленый/красный)
        def color_dyn(val):
            if isinstance(val, (int, float)):
                if val > 0: return 'color: #008000; font-weight: bold'
                if val < 0: return 'color: #FF0000; font-weight: bold'
            return ''

        # Вывод основной таблицы
        st.subheader("📋 Детальные данные по перегонам")
        st.dataframe(
            df_final.style.format({
                "Nуч": "{:.2f}", 
                "Прошлый Nуч": "{:.2f}", 
                "Динамика": "{:+.2f}"
            })
            .background_gradient(subset=['Nуч'], cmap='RdYlGn') # Цвет фона для Nуч
            .map(color_dyn, subset=['Динамика']),               # Цвет текста для Динамики
            use_container_width=True
        )

        # --- 7. ЭКСПОРТ В EXCEL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Анализ')
            # Здесь можно добавить доп. форматирование Excel при необходимости
        
        st.download_button(
            label="📥 Скачать полный отчет в Excel",
            data=output.getvalue(),
            file_name="Analiz_Nuch.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Не удалось сопоставить данные. Проверьте коды направлений в файлах.")
else:
    st.info("👋 Добро пожаловать! Начните с загрузки файлов в верхней части страницы.")
