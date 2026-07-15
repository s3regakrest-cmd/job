import os, re, datetime, sqlite3, streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Расчёт Заказов", page_icon="⚙️")
DAYS = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']


def expand_serial_input(text):
    text = text.strip()
    if text.lower() == 'today': return True, DAYS[datetime.datetime.now().weekday()], 1
    parts = [p.strip() for p in re.split(r'[\s,]+', text) if p.strip()]
    if not parts: return False, "Строка пуста", 0
    count, res = 0, []
    for p in parts:
        if '-' in p:
            sub = p.split('-')
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


def generate_excel_bytes(session_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет за день"
    headers = ["наименование", "номер чертежа", "номер операции", "стоимость за единицу", "номера изделий",
               "количество", "общая стоимость (операция)", "общая сумма за смену"]
    ws.append(headers)
    for c in range(1, 9): ws.cell(row=1, column=c).font = Font(bold=True)

    l_name, l_draw = "", ""
    for item in session_data:
        f_op = f"{item['op_num']} {item['desc']}"
        same = item['name'].lower() == l_name.lower() and item['drawing'] == l_draw
        ws.append(["" if same else item['name'], "" if same else item['drawing'], f_op, f"{item['price']:.2f} руб.",
                   item['serials'], item['count'], f"{item['total']:.2f} руб.", ""])
        l_name, l_draw = item['name'], item['drawing']

    ws.cell(row=2, column=8).value = f"{sum(i['total'] for i in session_data):.2f} руб."
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = max(
            max([len(str(ws.cell(row=r, column=col).value or '')) for r in range(1, ws.max_row + 1)]) + 4, 12)

    f = BytesIO()
    wb.save(f)
    return f.getvalue()


if 'storage' not in st.session_state: st.session_state.storage = []
st.title("⚙️ Расчёт заказов")

if not os.path.exists('production.db'):
    st.error("Файл 'production.db' не найден!")
else:
    with sqlite3.connect('production.db') as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT name FROM items")
        db_names = [r for r in cursor.fetchall()]

    name = st.selectbox("Изделие:", db_names)
    ops_raw = st.text_input("Операции (через запятую):", placeholder="25, 45")
    serials_raw = st.text_input("Номера изделий:", placeholder="140-42, 147")

    if st.button("➕ Рассчитать и добавить", use_container_width=True):
        if not ops_raw or not serials_raw:
            st.warning("Заполните все поля!")
        else:
            ok, serials, count = expand_serial_input(serials_raw)
            if not ok:
                st.error(serials)
            else:
                ops = [o.strip() for o in ops_raw.split(',') if o.strip()]
                found = []
                with sqlite3.connect('production.db') as conn:
                    cursor = conn.cursor()
                    for op in ops:
                        cursor.execute(
                            "SELECT drawing_number, work_description, price_per_unit FROM items WHERE LOWER(name)=LOWER(?) AND (work_description LIKE ? OR work_description=?)",
                            (name, f'{op},%', op))
                        res = cursor.fetchone()
                        if res: found.append(
                            {'op_num': op, 'desc': re.sub(r'^\d+\s*,\s*', '', str(res)).strip(), 'price': float(res),
                             'drawing': res})

                if not found:
                    st.error("Операции не найдены.")
                else:
                    for o in found:
                        st.session_state.storage.append(
                            {'name': name, 'drawing': o['drawing'], 'op_num': o['op_num'], 'desc': o['desc'],
                             'price': o['price'], 'serials': serials, 'count': count, 'total': o['price'] * count})
                    st.success("Успешно добавлено!")

    if st.session_state.storage:
        st.write("---")
        for item in st.session_state.storage:
            st.write(
                f"**{item['name']}** | Оп. {item['op_num']} ({item['desc']}) | {item['count']} шт. (№ {item['serials']}) — *{item['total']:.2f} руб.*")

        st.metric(label="Общая сумма за смену", value=f"{sum(i['total'] for i in st.session_state.storage):.2f} руб.")

        st.download_button(
            label="💾 Скачать отчет Excel на iPhone",
            data=generate_excel_bytes(st.session_state.storage),
            file_name=f"{datetime.datetime.now().strftime('%d.%m.%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        if st.button("🗑️ Сбросить смену", use_container_width=True):
            st.session_state.storage = []
            st.rerun()
