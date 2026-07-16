import os, re, datetime, sqlite3, streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --- НАСТРОЙКИ ---
st.set_page_config(page_title="Расчёт Заказов", page_icon="⚙️")
DAYS = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']
DB_PATH = 'production.db'

if 'storage' not in st.session_state: 
    st.session_state.storage = []

# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

@st.cache_data(ttl=300)
def get_items_from_db():
    if not os.path.exists(DB_PATH):
        return []
    try:
        with sqlite3.connect(DB_PATH) as conn:
            return [r[0] for r in conn.execute("SELECT DISTINCT name FROM items").fetchall()]
    except Exception:
        return []

@st.cache_data
def generate_excel_bytes(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет за день"
    
    headers = ["наименование", "номер чертежа", "номер операции", "стоимость за единицу", 
               "номера изделий", "количество", "общая стоимость (операция)", "общая сумма за смену"]
    ws.append(headers)
    
    for c in range(1, 9): 
        ws.cell(row=1, column=c).font = Font(bold=True)
    
    l_name, l_draw = "", ""
    
    for i in data:
        f_op = f"{i['op_num']} {i['desc']}"
        same = i['name'].lower() == l_name.lower() and i['drawing'] == l_draw
        
        row_data = [
            "" if same else i['name'], 
            "" if same else i['drawing'], 
            f_op, 
            f"{i['price']:.2f} руб.", 
            i['serials'], 
            i['count'], 
            f"{i['total']:.2f} руб.", 
            ""
        ]
        ws.append(row_data)
        l_name, l_draw = i['name'], i['drawing']
    
    if data:
        total_sum = sum(i['total'] for i in data)
        ws.cell(row=2, column=8).value = f"{total_sum:.2f} руб."

    for c in range(1, 9):
        max_len = 0
        for r in range(1, ws.max_row + 1):
            cell_val = str(ws.cell(row=r, column=c).value or '')
            max_len = max(max_len, len(cell_val))
        ws.column_dimensions[get_column_letter(c)].width = max(max_len + 4, 12)

    f = BytesIO()
    wb.save(f)
    return f.getvalue()

def expand_serial_input(text):
    text = text.strip()
    if text.lower() == 'today': 
        return True, DAYS[datetime.datetime.now().weekday()], 1
    
    parts = [p.strip() for p in re.split(r'[\s,]+', text) if p.strip()]
    if not parts: 
        return False, "Строка пуста", 0
    
    count, res = 0, []
    for p in parts:
        if '-' in p:
            sub = p.split('-')
            s, e = sub[0].strip(), sub[1].strip()
            if len(e) < len(s): 
                e = s[:len(s) - len(e)] + e
            
            try:
                start_int = int(s)
                end_int = int(e)
                if end_int < start_int:
                    return False, f"Некорректный диапазон: {p}", 0
                count += (end_int - start_int + 1)
                res.append(f"{s}-{e}")
            except ValueError:
                return False, f"Ошибка в диапазоне: '{p}'", 0
        elif p.isdigit(): 
            count += 1
            res.append(str(int(p)))
        else: 
            return False, f"Ошибка в значении: '{p}'", 0
            
    return True, ", ".join(res), count

def init_db_indexes():
    if not os.path.exists(DB_PATH):
        return
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_items_name_lower ON items(LOWER(name))")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_items_desc ON items(work_description)")
        conn.commit()

# --- ОСНОВНАЯ ЛОГИКА ---

if not os.path.exists(DB_PATH):
    st.error(f"Файл базы данных '{DB_PATH}' не найден!")
else:
    init_db_indexes()

db_names = get_items_from_db()

with st.form("order_form", clear_on_submit=True):
    col1, col2 = st.columns([3, 1])
    
    with col1:
        item = st.text_input("Изделие", placeholder="Начните вводить название...", autocomplete="off")
        ops_input = st.text_input("Операции (через запятую)", placeholder="10, 20, 30")
        serials_input = st.text_input("Номера изделий", placeholder="101-110 или 101,102,103 или 'today'")
    
    with col2:
        st.write("")
        submitted = st.form_submit_button("➕ Рассчитать и добавить", type="primary", use_container_width=True)

if submitted:
    # Важно: не делаем st.rerun() здесь. Streamlit сам перерисует страницу после обработки формы.
    if not item or not ops_input or not serials_input:
        st.warning("Пожалуйста, заполните все поля формы.")
    else:
        ok, serials_str, count = expand_serial_input(serials_input)
        if not ok:
            st.error(serials_str)
        else:
            ops_list = [o.strip() for o in ops_input.split(',') if o.strip()]
            found_ops = []
            
            with sqlite3.connect(DB_PATH) as conn:
                cursor = conn.cursor()
                for op in ops_list:
                    cursor.execute(
                        "SELECT drawing_number, work_description, price_per_unit FROM items WHERE LOWER(name)=LOWER(?) AND work_description LIKE ?", 
                        (item, f"{op}%")
                    )
                    res = cursor.fetchone()
                    
                    if res:
                        desc_raw = str(res[1])
                        clean_desc = re.sub(r'^\d+\s*,\s*', '', desc_raw).strip()
                        
                        found_ops.append({
                            'op_num': op, 
                            'desc': clean_desc, 
                            'price': float(res[2]), 
                            'drawing': str(res[0])
                        })
            
            if not found_ops:
                st.error(f"Операции '{ops_input}' для изделия '{item}' не найдены в базе.")
            else:
                for o in found_ops:
                    st.session_state.storage.append({
                        'name': item, 
                        'drawing': o['drawing'], 
                        'op_num': o['op_num'], 
                        'desc': o['desc'], 
                        'price': o['price'], 
                        'serials': serials_str, 
                        'count': count, 
                        'total': o['price'] * count
                    })
                # Просто показываем сообщение. Страница обновится сама.
                st.success("Заказ успешно добавлен!")

st.write("---")
grand_total_now = sum(i['total'] for i in st.session_state.storage)
header_col, metric_col = st.columns(2)
with header_col: st.title("⚙️ Расчёт заказов")
with metric_col: st.metric(label="Сумма за смену", value=f"{grand_total_now:,.2f} руб.")

if st.session_state.storage:
    MAX_SHOW = 100
    to_show = st.session_state.storage[-MAX_SHOW:]
    
    with st.expander("🔍 Подробнее", expanded=True):
        for i in to_show: 
            st.write(f"**{i['name']}** | Оп. {i['op_num']} ({i['desc']}) | {i['count']} шт. (№ {i['serials']}) — *{i['total']:.2f} руб.*")
        
        if len(st.session_state.storage) > MAX_SHOW:
            st.caption(f"Показано последние {MAX_SHOW} из {len(st.session_state.storage)} записей.")

    excel_file = generate_excel_bytes(st.session_state.storage)
    st.download_button(
        "💾 Скачать отчет Excel", 
        excel_file, 
        f"report_{datetime.datetime.now().strftime('%d.%m.%Y_%H-%M')}.xlsx", 
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        use_container_width=True
    )
    
    # Сброс смены: только тут нужен rerun, потому что мы очищаем session_state
    if st.button("🗑️ Сбросить смену", use_container_width=True, type="secondary"):
        st.session_state.storage = []
        st.rerun()

st.write("---")
with st.expander("🔐 Редактор базы данных"):
    pwd = st.text_input("Пароль администратора", type="password", key="adm_p")
    
    if pwd == "1234":
        add_name = st.text_input("Наименование изделия")
        add_draw = st.text_input("Номер чертежа")
        add_desc = st.text_input("Описание операции (начните с номера, например: '10, Токарная')")
        add_price = st.number_input("Цена за единицу", min_value=0.0, step=0.5)
        
        if st.button("💾 Сохранить в базу данных", use_container_width=True):
            if not add_name or not add_draw or not add_desc or add_price <= 0: 
                st.error("Заполните корректно все поля!")
            else:
                try:
                    with sqlite3.connect(DB_PATH) as conn: 
                        conn.execute(
                            "INSERT INTO items (name, drawing_number, work_description, price_per_unit) VALUES (?, ?, ?, ?)", 
                            (add_name, add_draw, add_desc, add_price)
                        )
                        conn.commit()
                    st.success("Успешно добавлено в базу данных!")
                    # Здесь rerun нужен, чтобы обновился список автодополнения
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка при записи в БД: {e}")
    elif pwd:
        st.error("Неверный пароль администратора")

if not os.path.exists(DB_PATH):
    st.info("""
    Для работы приложения требуется файл `production.db` с таблицей `items`.
    """)
