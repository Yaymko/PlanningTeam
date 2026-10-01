"""
Еженедельное планирование загрузки команды.
Запуск: streamlit run app.py
"""
from datetime import datetime, timezone

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core import DEFAULT_NORM, build_export_workbook, clean_week_label, parse_workbook

st.set_page_config(page_title="Загрузка команды", layout="wide")

parse_workbook_cached = st.cache_data(show_spinner=False)(parse_workbook)

st.title("Еженедельное планирование загрузки команды")

uploaded = st.file_uploader("Загрузите файл с Agile-доски (.xlsx)", type="xlsx")

if not uploaded:
    st.info("Загрузите еженедельный экспорт с доски — приложение разберёт тикеты по исполнителям, "
            "вы впишете часы прямо здесь и увидите загрузку команды на графиках.")
    st.stop()

if "file_name" not in st.session_state or st.session_state.file_name != uploaded.name:
    people_data, week_labels = parse_workbook_cached(uploaded.getvalue())
    st.session_state.file_name = uploaded.name
    # base_people_data is the baseline handed to each data_editor widget. It must stay
    # byte-for-byte identical across reruns (same values/dtypes) — if we fed the editor
    # its own previously-edited-and-cleaned output back in as "data", Streamlit treats
    # it as a new dataset and can reset in-progress edits, which is why typed hours
    # sometimes didn't stick on the first try. people_data (below) is the live, edited
    # copy used for the dashboard/export and is never passed back into data_editor.
    st.session_state.base_people_data = {k: v.copy() for k, v in people_data.items()}
    st.session_state.people_data = {k: v.copy() for k, v in people_data.items()}
    st.session_state.week_labels = week_labels
    st.session_state.norms = {name: DEFAULT_NORM for name in people_data}

people_data = st.session_state.people_data
base_people_data = st.session_state.base_people_data
week_labels = st.session_state.week_labels

if not week_labels:
    st.error("Не удалось найти недельные колонки часов в файле. Проверьте структуру листов исполнителей.")
    st.stop()

tab_input, tab_dashboard = st.tabs(["Ввод часов", "Дашборд"])

with tab_input:
    st.caption("Часы по тикетам с доски впишите в нужную неделю. Строки «Доп. работа» — для задач, "
               "которых нет на доске: впишите название и часы. Новые строки можно добавлять кнопкой "
               "«+» внизу таблицы, лишние — удалять через иконку корзины слева от строки.")
    names = list(people_data.keys())
    selected = st.selectbox("Исполнитель", names)
    norm_val = st.number_input(f"Норма часов в неделю — {selected}",
                                min_value=0.0, value=float(st.session_state.norms[selected]), step=1.0)
    st.session_state.norms[selected] = norm_val

    column_config = {
        "Тикет": st.column_config.TextColumn(),
        "Название": st.column_config.TextColumn(width="large"),
        "Колонка": st.column_config.TextColumn(),
        "Доска": st.column_config.TextColumn(),
        "Срок": st.column_config.DateColumn(format="DD.MM.YYYY"),
        "В работе": st.column_config.CheckboxColumn(),
    }
    for wl in week_labels:
        column_config[wl] = st.column_config.NumberColumn(clean_week_label(wl), min_value=0.0, step=0.5)

    # Always hand the editor the same untouched baseline for this person — Streamlit
    # owns the live edits internally via `key` from that point on. Passing our cleaned-up
    # `edited` copy back in here on the next rerun is what caused dropped/stuck edits.
    edited = st.data_editor(base_people_data[selected], column_config=column_config, hide_index=True,
                             use_container_width=True, num_rows="dynamic", key=f"editor_{selected}")

    # Newly added rows arrive with NaN/None in unfilled cells — normalise before aggregation/export.
    edited["Тикет"] = edited["Тикет"].fillna("")
    edited["Название"] = edited["Название"].fillna("")
    edited["Колонка"] = edited["Колонка"].fillna("Доп. работа")
    edited["Доска"] = edited["Доска"].fillna("")
    edited["В работе"] = edited["В работе"].fillna(False)
    for wl in week_labels:
        edited[wl] = edited[wl].fillna(0.0)

    people_data[selected] = edited
    st.session_state.people_data[selected] = edited

with tab_dashboard:
    rows = []
    for name, df in people_data.items():
        row = {"Исполнитель": name, "Норма, ч/нед": st.session_state.norms[name]}
        for wl in week_labels:
            row[clean_week_label(wl)] = df[wl].sum()
        rows.append(row)
    summary = pd.DataFrame(rows)
    week_cols = [clean_week_label(w) for w in week_labels]

    st.subheader("Сводка по команде")

    def highlight_overload(row):
        styles = []
        for col in summary.columns:
            if col in week_cols and row[col] > row["Норма, ч/нед"] > 0:
                styles.append("background-color:#F4CCCC")
            elif col in week_cols and 0 < row[col] <= row["Норма, ч/нед"]:
                styles.append("background-color:#D9EAD3")
            else:
                styles.append("")
        return styles

    st.dataframe(summary.style.apply(highlight_overload, axis=1), use_container_width=True, hide_index=True)

    col1, col2 = st.columns(2)
    with col1:
        long_df = summary.melt(id_vars=["Исполнитель"], value_vars=week_cols,
                                var_name="Неделя", value_name="Часы")
        fig = px.bar(long_df, x="Исполнитель", y="Часы", color="Неделя", barmode="group",
                     title="Загрузка по неделям, по каждому исполнителю")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        team_total = summary[week_cols].sum()
        team_norm = summary["Норма, ч/нед"].sum()
        fig2 = go.Figure()
        fig2.add_bar(x=week_cols, y=team_total.values, name="Запланировано, ч")
        fig2.add_scatter(x=week_cols, y=[team_norm] * len(week_cols),
                          mode="lines+markers", name="Норма команды, ч")
        fig2.update_layout(title="Суммарная загрузка команды по неделям vs норма")
        st.plotly_chart(fig2, use_container_width=True)

    overloaded = summary[summary.apply(
        lambda r: any(r[c] > r["Норма, ч/нед"] > 0 for c in week_cols), axis=1)]
    if not overloaded.empty:
        st.warning("Перегружены на одной или нескольких неделях: " + ", ".join(overloaded["Исполнитель"]))
    else:
        st.success("Перегрузов не обнаружено.")

st.divider()
st.subheader("Экспорт")
st.caption("Скачайте заполненный файл в формате Excel — с формулами, сводкой и графиками, для архива.")
if st.button("Сформировать Excel"):
    data = build_export_workbook(people_data, week_labels, st.session_state.norms)
    st.download_button(
        "Скачать файл",
        data=data,
        file_name=f"Загрузка_команды_{datetime.now(timezone.utc):%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
