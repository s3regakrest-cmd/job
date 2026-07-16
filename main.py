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
            s, e = sub[0].strip(), sub[1].strip()  # ИСПРАВЛЕНО: берем элементы списка по индексам
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
        db_names = [r[0] for r in cursor.fetchall()]

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
                        if res: found.append({'op_num': op, 'desc': re.sub(r'^\d+\s*,\s*', '', str(res[1])).strip(),
                                              'price': float(res[2]), 'drawing': res[0]})

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

    # --- СЕКРЕТНЫЙ БЛОК ДОБАВЛЕНИЯ В БАЗУ ПРЯМО С ТЕЛЕФОНА ---
    st.write("---")
    with st.expander("🔐 Редактор базы данных (Добавить новую деталь)"):
        pwd = st.text_input("Пароль администратора:", type="password", key="adm_p")
        if pwd == "1234":  # Сюда можно вписать любой ваш пароль
            add_name = st.text_input("Наименование нового изделия (например: бэшка):").strip()
            add_draw = st.text_input("Номер чертежа:").strip()
            # Напоминалка: номер операции должен быть в начале строки
            add_desc = st.text_input("Описание (формат: '10, описание_работ'):").strip()
            add_price = st.number_input("Стоимость за единицу (руб):", min_value=0.0, step=0.5, key="adm_pr")

            if st.button("💾 Сохранить в базу данных", use_container_width=True):
                if not add_name or not add_draw or not add_desc or add_price <= 0:
                    st.error("Заполните все поля корректно!")
                else:
                    with sqlite3.connect('production.db') as conn:
                        cursor = conn.cursor()
                        cursor.execute("""
                            INSERT INTO items (name, drawing_number, work_description, price_per_unit) 
                            VALUES (?, ?, ?, ?)
                        """, (add_name, add_draw, add_desc, add_price))
                        conn.commit()
                    st.success(f"Запись для '{add_name}' успешно сохранена в базу!")
                    st.rerun()
        elif pwd != "":
            st.error("Неверный пароль!")
