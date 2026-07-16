import os, re, datetime, json, streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Расчёт Заказов", page_icon="⚙️")
DAYS = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']

if 'storage' not in st.session_state: 
    st.session_state.storage = []

def expand_serial_input(text):
    text = text.strip()
    if text.lower() == 'today': return True, DAYS[datetime.datetime.now().weekday()], 1
    parts = [p.strip() for p in re.split(r'[\s,]+', text) if p.strip()]
    if not parts: return False, "Строка пуста", 0
    count, res = 0, []
    for p in parts:
        if '-' in p:
            sub = p.split('-')
            if len(sub) != 2: return False, f"Ошибка в диапазоне: '{p}'", 0
            s, e = sub.strip(), sub.strip()
            if len(e) < len(s): e = s[:len(s) - len(e)] + e
            count += int(e) - int(s) + 1
            res.append(f"{s}-{e}")
        elif p.isdigit(): 
            count += 1
            res.append(str(int(p)))
        else: 
            return False, f"Ошибка в: '{p}'", 0
    return True, ", ".join(res), count

def generate_excel_bytes(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет за день"
    ws.append(["наименование", "номер чертежа", "номер операции", "стоимость за единицу", "номера изделий", "количество", "общая стоимость (операция)", "общая сумма за смену"])
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

grand_total_now = sum(i['total'] for i in st.session_state.storage)
header_col, metric_col = st.columns(2)
with header_col: st.title("⚙️ Расчёт заказов")
with metric_col: st.metric(label="Сумма за смену", value=f"{grand_total_now:,.2f} руб.")

# Чтение Google Таблицы
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(spreadsheet=st.secrets["public_gsheets_url"], ttl="5m")
    db_names = df["name"].dropna().unique().tolist()
except Exception as e:
    st.error(f"Не удалось подключиться к Google Таблице: {e}")
    db_names = []

with st.form(key="main_order_form", clear_on_submit=True):
    selected_name = st.text_input("Изделие")
    ops_raw = st.text_input("Операции")
    serials_raw = st.text_input("Номера изделий")
    submit_button = st.form_submit_button(label="➕ Рассчитать и добавить", use_container_width=True)
# ОБРАБОТКА ДАННЫХ (Вне контекста формы)
if submit_button:
    if not db_names:
        st.error("База данных временно недоступна.")
    elif not selected_name.strip():
        st.error("Пожалуйста, введите наименование изделия!")
    elif not ops_raw.strip():
        st.error("Пожалуйста, укажите номера операций!")
    elif not serials_raw.strip():
        st.error("Пожалуйста, укажите номера изделий!")
    else:
        selected_name = selected_name.strip()
        ops_raw = ops_raw.strip()
        serials_raw = serials_raw.strip()
        
        name_exists = any(selected_name.lower() == str(name).lower() for name in db_names)
        
        if not name_exists:
            st.error(f"Изделие '{selected_name}' не найдено в Google Таблице!")
        else:
            for name in db_names:
                if selected_name.lower() == str(name).lower():
                    selected_name = str(name)
                    break
            
            ok, serials, count = expand_serial_input(serials_raw)
            if not ok: 
                st.error(serials)
            else:
                ops = [o.strip() for o in ops_raw.split(',') if o.strip()]
                found = []
                
                # Фильтруем строки по названию изделия
                sub_df = df[df["name"].astype(str).str.lower() == selected_name.lower()]
                
                for op in ops:
                    # ИСПРАВЛЕНИЕ: Теперь ищем строгое текстовое или числовое совпадение в НОВОМ столбце op_num
                    match = sub_df[sub_df["op_num"].astype(str).str.strip() == op]
                    
                    if not match.empty:
                        row = match.iloc[0]
                        # Текстовое описание теперь берется напрямую без вырезания цифр
                        desc_clean = str(row["work_description"]).strip()
                        found.append({
                            'op_num': op,
                            'desc': desc_clean,
                            'price': float(row["price_per_unit"]),
                            'drawing': str(row["drawing_number"])
                        })

                if not found: 
                    st.error(f"Операции {ops_raw} для изделия '{selected_name}' не найдены в Google Таблице.")
                else:
                    for o in found:
                        st.session_state.storage.append({
                            'name': selected_name, 
                            'drawing': o['drawing'], 
                            'op_num': o['op_num'], 
                            'desc': o['desc'], 
                            'price': o['price'], 
                            'serials': serials, 
                            'count': count, 
                            'total': o['price'] * count
                        })
                    st.success("Успешно добавлено!")
                    st.rerun()

# Вывод результатов смены
if st.session_state.storage:
    st.write("---")
    with st.expander("🔍 Подробнее", expanded=True):
        for i in st.session_state.storage: 
            st.write(f"**{i['name']}** | Оп. {i['op_num']} ({i['desc']}) | {i['count']} шт. (№ {i['serials']}) — *{i['total']:.2f} руб.*")
    
    excel_file = generate_excel_bytes(st.session_state.storage)
    st.download_button(
        "💾 Скачать отчет Excel на iPhone", 
        excel_file, 
        f"{datetime.datetime.now().strftime('%d.%m.%Y')}.xlsx", 
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        use_container_width=True
    )
    
    if st.button("🗑️ Сбросить смену", use_container_width=True):
        st.session_state.storage = []
        st.rerun()

st.write("---")
st.caption("Данные калькулятора успешно синхронизированы с Google Таблицей. Новые изделия и цены добавляются напрямую в файл на вашем Google Диске.")
