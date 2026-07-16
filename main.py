import os, re, datetime, sqlite3, streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Расчёт Заказов", page_icon="⚙️")
DAYS = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

# --- ИНИЦИАЛИЗАЦИЯ СОСТОЯНИЯ СЕССИИ ДЛЯ СБРОСА ПОЛЕЙ ---
if 'storage' not in st.session_state: st.session_state.storage = []
if 'ops_val' not in st.session_state: st.session_state.ops_val = ""
if 'serials_val' not in st.session_state: st.session_state.serials_val = ""
if 'search_query' not in st.session_state: st.session_state.search_query = ""

def expand_serial_input(text):
    text = text.strip()
    if text.lower() == 'today': return True, DAYS[datetime.datetime.now().weekday()], 1
    parts = [p.strip() for p in re.split(r'[\s,]+', text) if p.strip()]
    if not parts: return False, "Строка пуста", 0
    count, res = 0, []
    for p in parts:
        if '-' in p:
            sub = p.split('-')
            s, e = sub[0].strip(), sub[1].strip()
            if len(e) < len(s): e = s[:len(s) - len(e)] + e
            count += int(e) - int(s) + 1
            res.append(f"{s}-{e}")
        elif p.isdigit(): count += 1; res.append(str(int(p)))
        else: return False, f"Ошибка в: '{p}'", 0
    return True, ", ".join(res), count

def generate_excel_bytes(data):
    wb, ws = Workbook(), Workbook().active
    ws.title = "Отчет за день"
    ws.append(["наименование", "номер чертежа", "номер операции", "стоимость за единицу", "номера изделий", "количество", "общая стоимость (операция)", "общая ... смену"])
    for c in range(1, 9): ws.cell(row=1, column=c).font = Font(bold=True)
    
    l_name, l_draw = "", ""
    for i in data:
        f_op = f"{i['op_num']} {i['desc']}"
        same = i['name'].lower() == l_name.lower() and i['drawing'] == l_draw
        ws.append(["" if same else i['name'], "" if same else i['drawing'], f_op, f"{i['price']:.2f} руб.", i['serials'], i['count'], f"{i['total']:.2f} руб.", ""])
        l_name, l_draw = i['name'], i['drawing']

    ws.cell(row=2, column=8).value = f"{sum(i['total'] for i in data):.2f} руб."
    for c in range(1, 9):
        ws.column_dimensions[get_column_letter(c)].width = max(max([len(str(ws.cell(row=r, column=c).value or '')) for r in range(1, ws.max_row + 1)]) + 4, 12)
    f = BytesIO()
    wb.save(f)
    return f.getvalue()

# --- МЕТРИКА СУММЫ В ПРАВОМ ВЕРХНЕМ УГЛУ ---
grand_total_now = sum(i['total'] for i in st.session_state.storage)
header_col, metric_col = st.columns(2)
with header_col: st.title("⚙️ Расчёт заказов")
with metric_col: st.metric(label="Сумма за смену", value=f"{grand_total_now:,.2f} руб.")

if not os.path.exists('production.db'):
    st.error("Файл 'production.db' не найден!")
else:
    with sqlite3.connect('production.db') as conn:
        db_names = [r[0] for r in conn.execute("SELECT DISTINCT name FROM items").fetchall()]

    # --- НАСТОЯЩИЙ ПОИСКОВЫЙ САДЖЕСТ НА ПЕРВОМ ПЛАНЕ ---
    # st.text_input с мгновенным фокусом и клавиатурой
    typed_text = st.text_input(
        "Изделие:", 
        value=st.session_state.search_query, 
        placeholder="Введите название детали..."
    ).strip()

    selected_name = typed_text

    # Выдвижное меню подсказок прямо на первом плане
    if typed_text:
        matches = [name for name in db_names if typed_text.lower() in name.lower()]
        
        # Если нашли совпадения и пользователь еще не выбрал точное слово
        if matches and (len(matches) > 1 or matches[0].lower() != typed_text.lower()):
            # Создаем красивый выпадающий блок-контейнер
            with st.container(border=True):
                st.write("📋 *Варианты из базы данных:*")
                for match in matches[:5]: # Показываем топ-5 совпадений
                    # Кнопки оформлены как строки выпадающего меню
                    if st.button(f"🔍 {match}", key=f"btn_{match}", use_container_width=True):
                        st.session_state.search_query = match
                        st.rerun()

    # Поля ввода операций и номеров изделий
    ops_raw = st.text_input("Операции:", value=st.session_state.ops_val)
    serials_raw = st.text_input("Номера изделий:", value=st.session_state.serials_val)

    if st.button("➕ Рассчитать и добавить", use_container_width=True):
        if not selected_name or not ops_raw or not serials_raw: 
            st.warning("Заполните все поля ввода!")
        else:
            ok, serials, count = expand_serial_input(serials_raw)
            if not ok: st.error(serials)
            else:
                ops = [o.strip() for o in ops_raw.split(',') if o.strip()]
                found = []
                with sqlite3.connect('production.db') as conn:
                    for op in ops:
                        res = conn.execute("SELECT drawing_number, work_description, price_per_unit FROM items WHERE LOWER(name)=LOWER(?) AND (work_description LIKE ? OR work_description=?)", (selected_name, f'{op},%', op)).fetchone()
                        if res: found.append({'op_num': op, 'desc': re.sub(r'^\d+\s*,\s*', '', str(res[1])).strip(), 'price': float(res[2]), 'drawing': res[0]})
                
                if not found: st.error("Операции не найдены.")
                else:
                    for o in found: 
                        st.session_state.storage.append({'name': selected_name, 'drawing': o['drawing'], 'op_num': o['op_num'], 'desc': o['desc'], 'price': o['price'], 'serials': serials, 'count': count, 'total': o['price'] * count})
                    
                    # ПОЛНОЕ АВТОМАТИЧЕСКОЕ ОБНУЛЕНИЕ ВСЕХ ПОЛЕЙ
                    st.session_state.ops_val = ""
                    st.session_state.serials_val = ""
                    st.session_state.search_query = ""
                    st.success("Успешно добавлено!")
                    st.rerun()

    # --- КНОПКА ПОДРОБНЕЕ ДЛЯ СКРЫТИЯ ДАННЫХ ---
    if st.session_state.storage:
        st.write("---")
        with st.expander("🔍 Подробнее"):
            for i in st.session_state.storage: st.write(f"**{i['name']}** | Оп. {i['op_num']} ({i['desc']}) | {i['count']} шт. (№ {i['serials']}) — *{i['total']:.2f} руб.*")
        st.download_button("💾 Скачать отчет Excel на iPhone", generate_excel_bytes(st.session_state.storage), f"{datetime.datetime.now().strftime('%d.%m.%Y')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        if st.button("🗑️ Сбросить смену", use_container_width=True):
            st.session_state.storage, st.session_state.ops_val, st.session_state.serials_val, st.session_state.search_query = [], "", "", ""
            st.rerun()

    st.write("---")
    with st.expander("🔐 Редактор базы данных"):
        if st.text_input("Пароль администратора:", type="password", key="adm_p") == "1234":
            add_name, add_draw, add_desc, add_price = st.text_input("Наименование:").strip(), st.text_input("Чертеж:").strip(), st.text_input("Описание:").strip(), st.number_input("Цена:", min_value=0.0, step=0.5)
            if st.button("💾 Сохранить в базу данных", use_container_width=True):
                if not add_name or not add_draw or not add_desc or add_price <= 0: st.error("Заполните поля!")
                else:
                    with sqlite3.connect('production.db') as conn: conn.execute("INSERT INTO items (name, drawing_number, work_description, price_per_unit) VALUES (?, ?, ?, ?)", (add_name, add_draw, add_desc, add_price)); conn.commit()
                    st.success("Добавлено!"); st.rerun()
