import os, re, datetime, sqlite3, streamlit as st, streamlit.components.v1 as components
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Расчёт Заказов", page_icon="⚙️")
DAYS = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

# --- ИНИЦИАЛИЗАЦИЯ СОСТОЯНИЯ СЕССИИ ДЛЯ СТИРАНИЯ ---
if 'storage' not in st.session_state: st.session_state.storage = []
if 'ops_val' not in st.session_state: st.session_state.ops_val = ""
if 'serials_val' not in st.session_state: st.session_state.serials_val = ""
if 'item_name_val' not in st.session_state: st.session_state.item_name_val = ""

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
        elif p.isdigit():
            count += 1
            res.append(str(int(p)))
        else: return False, f"Ошибка в: '{p}'", 0
    return True, ", ".join(res), count

def generate_excel_bytes(session_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет за день"
    headers = ["наименование", "номер чертежа", "номер операции", "стоимость за единицу", "номера изделий", "количество", "общая стоимость (операция)", "общая сумма за смену"]
    ws.append(headers)
    for c in range(1, 9): ws.cell(row=1, column=c).font = Font(bold=True)
    
    l_name, l_draw = "", ""
    for item in session_data:
        f_op = f"{item['op_num']} {item['desc']}"
        same = item['name'].lower() == l_name.lower() and item['drawing'] == l_draw
        ws.append(["" if same else item['name'], "" if same else item['drawing'], f_op, f"{item['price']:.2f} руб.", item['serials'], item['count'], f"{item['total']:.2f} руб.", ""])
        l_name, l_draw = item['name'], item['drawing']

    ws.cell(row=2, column=8).value = f"{sum(i['total'] for i in session_data):.2f} руб."
    for col_idx in range(1, 9):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(max([len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(1, ws.max_row + 1)]) + 4, 12)
    
    f = BytesIO()
    wb.save(f)
    return f.getvalue()

# --- ИСПРАВЛЕНО: Явное указание пропорций колонок для st.columns ---
grand_total_now = sum(i['total'] for i in st.session_state.storage)
header_col, metric_col = st.columns([3, 1])
with header_col: st.title("⚙️ Расчёт заказов")
with metric_col: st.metric(label="Сумма за смену", value=f"{grand_total_now:,.2f} руб.")

if not os.path.exists('production.db'):
    st.error("Файл 'production.db' не найден!")
else:
    with sqlite3.connect('production.db') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT name FROM items")
        db_names = [r[0] for r in cursor.fetchall()]

    # --- РЕАЛИЗАЦИЯ НАТИВНОГО САДЖЕСТА ДЛЯ IPHONE (DATALIST) ---
    st.write("**Изделие:**")
    
    options_html = "".join([f'<option value="{name}">' for name in db_names])
    html_code = f"""
    <input type="text" id="item_input" list="items_list" value="{st.session_state.item_name_val}" placeholder="Начните писать название..." style="width:100%; padding:8px; border:1px solid #ccc; border-radius:4px; font-size:16px; font-family:sans-serif; box-sizing:border-box;">
    <datalist id="items_list">{options_html}</datalist>
    <script>
        const input = document.getElementById('item_input');
        input.addEventListener('input', (e) => {{
            window.parent.postMessage({{type: 'streamlit:setComponentValue', value: e.target.value}}, '*');
        }});
    </script>
    """
    
    raw_selected_name = components.html(html_code, height=45)
    
    if raw_selected_name is None:
        selected_name = str(st.session_state.item_name_val)
    else:
        selected_name = str(raw_selected_name)
        st.session_state.item_name_val = selected_name

    # Поля ввода операций и номеров изделий
    ops_raw = st.text_input("Операции (через запятую):", value=st.session_state.ops_val)
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
                    cursor = conn.cursor()
                    for op in ops:
                        cursor.execute("SELECT drawing_number, work_description, price_per_unit FROM items WHERE LOWER(name)=LOWER(?) AND (work_description LIKE ? OR work_description=?)", (selected_name, f'{op},%', op))
                        res = cursor.fetchone()
                        if res: found.append({'op_num': op, 'desc': re.sub(r'^\d+\s*,\s*', '', str(res[1])).strip(), 'price': float(res[2]), 'drawing': res[0]})

                if not found: st.error("Операции не найдены в базе.")
                else:
                    for o in found:
                        st.session_state.storage.append({'name': selected_name, 'drawing': o['drawing'], 'op_num': o['op_num'], 'desc': o['desc'], 'price': o['price'], 'serials': serials, 'count': count, 'total': o['price'] * count})
                    
                    # Автоматическое обнуление всех трех полей ввода при нажатии
                    st.session_state.ops_val = ""
                    st.session_state.serials_val = ""
                    st.session_state.item_name_val = ""
                    st.success("Успешно добавлено!")
                    st.rerun()

    # --- КНОПКА ПОДРОБНЕЕ ДЛЯ СКРЫТИЯ ДАННЫХ ---
    if st.session_state.storage:
        st.write("---")
        with st.expander("🔍 Подробнее"):
            for item in st.session_state.storage:
                st.write(f"**{item['name']}** | Оп. {item['op_num']} ({item['desc']}) | {item['count']} шт. (№ {item['serials']}) — *{item['total']:.2f} руб.*")

        st.download_button(
            label="💾 Скачать отчет Excel на iPhone",
            data=generate_excel_bytes(st.session_state.storage),
            file_name=f"{datetime.datetime.now().strftime('%d.%m.%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        if st.button("🗑️ Сбросить смену", use_container_width=True):
            st.session_state.storage = []; st.session_state.ops_val = ""; st.session_state.serials_val = ""; st.session_state.item_name_val = ""
            st.rerun()

    # --- РЕДАКТОР БАЗЫ ДАННЫХ ---
    st.write("---")
    with st.expander("🔐 Редактор базы данных (Добавить новую деталь)"):
        pwd = st.text_input("Пароль администратора:", type="password", key="adm_p")
        if pwd == "1234":
            add_name = st.text_input("Наименование нового изделия:").strip()
            add_draw = st.text_input("Номер чертежа:").strip()
            add_desc = st.text_input("Описание (формат: '10, описание_работ'):").strip()
            add_price = st.number_input("Стоимость за единицу (руб):", min_value=0.0, step=0.5, key="adm_pr")

            if st.button("💾 Сохранить в базу данных", use_container_width=True):
                if not add_name or not add_draw or not add_desc or add_price <= 0: st.error("Заполните все поля!")
                else:
                    with sqlite3.connect('production.db') as conn:
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO items (name, drawing_number, work_description, price_per_unit) VALUES (?, ?, ?, ?)", (add_name, add_draw, add_desc, add_price))
                        conn.commit()
                    st.success("Успешно добавлено!")
                    st.rerun()
        elif pwd != "": st.error("Неверный пароль!")
