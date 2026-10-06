"""
Еженедельное планирование загрузки команды.
Запуск: streamlit run app.py

Сотрудники заранее заполняют свои часы во вкладке «Мои часы» и сохраняют их в общую базу;
на встрече во вкладке «Планирование» видны часы всех и выгружается отчёт на неделю.
"""
from datetime import date, datetime, timezone

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core import (
    DELETE_COL,
    FIXED_COLS,
    actual_label,
    build_export_workbook,
    clean_week_label,
    drop_marked_rows,
    hour_cols,
    nearest_weeks,
    parse_workbook,
)
from storage import Store, db_path_from_env

st.set_page_config(page_title="Загрузка команды", layout="wide")


@st.cache_resource(show_spinner=False)
def get_store(path: str) -> Store:
    return Store(path)


def visible_weeks(week_labels, key):
    """Ближайшие 2 недели (прошлая и текущая) или все, если включён переключатель."""
    show_all = st.toggle("Показать все недели плана", key=key,
                         help="По умолчанию видны только прошлая и текущая неделя: "
                              "в начале недели можно внести факт за прошедшую.")
    return list(week_labels) if show_all else nearest_weeks(week_labels, date.today())


def fmt_saved(ts):
    if not ts:
        return "—"
    return datetime.fromisoformat(ts).strftime("%d.%m %H:%M")


store = get_store(str(db_path_from_env()))
plan = store.get_plan()

st.title("Еженедельное планирование загрузки команды")
if plan:
    weeks = ", ".join(clean_week_label(w) for w in plan["week_labels"])
    st.caption(f"План недели: {plan['file_name']} (загружен {fmt_saved(plan['uploaded_at'])}). Недели: {weeks}")

tab_mine, tab_planning = st.tabs(["Мои часы", "Планирование"])

# --- Мои часы: каждый сотрудник заполняет и сохраняет свои часы ------------------------
with tab_mine:
    if not plan:
        st.info("План недели ещё не загружен. Загрузите еженедельный экспорт с доски во вкладке "
                "«Планирование» — после этого каждый сотрудник сможет заполнить здесь свои часы.")
    else:
        week_labels = plan["week_labels"]
        st.caption("Выберите себя, впишите плановые часы по тикетам на неделю и нажмите «Сохранить». "
                   "В конце недели впишите в колонку «факт» сколько часов реально ушло и снова сохраните. Строки "
                   "«Доп. работа» — для задач, которых нет на доске: впишите название и часы. Новые строки "
                   "добавляются кнопкой «+» внизу таблицы. Чтобы убрать тикет или работу, отметьте "
                   "галочку «Удалить» в начале строки: строка удалится при сохранении.")
        selected = st.selectbox("Сотрудник", store.people(), index=None, placeholder="Выберите себя из списка")
        if selected:
            # The editor's baseline must stay identical across reruns, otherwise Streamlit resets
            # in-progress edits. So it is loaded from the base once per (person, version), and the
            # version is bumped after each save to start a fresh editor from the saved data.
            versions = st.session_state.setdefault("editor_versions", {})
            ver = versions.get(selected, 0)
            editor_key = f"editor_{selected}_{plan['uploaded_at']}_{ver}"
            baselines = st.session_state.setdefault("baselines", {})
            if editor_key not in baselines:
                df, norm, saved_at = store.load_person(selected)
                df["Срок"] = pd.to_datetime(df["Срок"])  # proper date column for the editor (empty, not "None")
                df.insert(0, DELETE_COL, False)
                baselines[editor_key] = (df, norm, saved_at)
            base_df, saved_norm, saved_at = baselines[editor_key]

            if saved_at:
                st.success(f"Ваши часы сохранены {fmt_saved(saved_at)}. Можно поправить и сохранить ещё раз.")
            else:
                st.warning("Вы ещё не сохраняли часы на эту неделю.")

            norm_val = st.number_input("Норма часов в неделю", min_value=0.0, value=float(saved_norm), step=1.0,
                                       help="Уменьшите, если часть недели в отпуске или на больничном.",
                                       key=f"norm_{editor_key}")

            column_config = {
                DELETE_COL: st.column_config.CheckboxColumn(
                    help="Отметьте, чтобы удалить строку. Удаление применяется при нажатии «Сохранить».",
                    default=False),
                "Тикет": st.column_config.TextColumn(),
                "Название": st.column_config.TextColumn(width="large"),
                "Колонка": st.column_config.TextColumn(),
                "Доска": st.column_config.TextColumn(),
                "Срок": st.column_config.DateColumn(format="DD.MM.YYYY"),
                "В работе": st.column_config.CheckboxColumn(),
            }
            for wl in week_labels:
                column_config[wl] = st.column_config.NumberColumn(
                    f"{clean_week_label(wl)} план", min_value=0.0, step=0.01, format="%g")
                column_config[actual_label(wl)] = st.column_config.NumberColumn(
                    f"{clean_week_label(wl)} факт", min_value=0.0, step=0.01, format="%g",
                    help="Сколько часов реально ушло на задачу за неделю. Заполняется в конце недели.")

            shown_weeks = visible_weeks(week_labels, key="all_weeks_mine")
            # Hidden weeks stay in the data (column_order only hides them), so saving keeps their hours.
            column_order = [DELETE_COL, *FIXED_COLS, *hour_cols(shown_weeks)]
            edited = st.data_editor(base_df, column_config=column_config, column_order=column_order,
                                    hide_index=True, use_container_width=True, num_rows="dynamic",
                                    key=editor_key)
            kept = drop_marked_rows(edited)
            marked = len(edited) - len(kept)
            if marked:
                st.caption(f":red[Отмечено к удалению строк: {marked}. Они удалятся при сохранении "
                           "и уже не учитываются в итогах ниже.]")

            def col_sum(col):
                return pd.to_numeric(kept[col], errors="coerce").fillna(0).sum()

            cols = st.columns(len(shown_weeks))
            for col, wl in zip(cols, shown_weeks):
                hours, actual = col_sum(wl), col_sum(actual_label(wl))
                delta = hours - norm_val
                col.metric(f"План {clean_week_label(wl)}", f"{hours:g} ч", f"{delta:+g} ч к норме",
                           delta_color="inverse" if delta > 0 else "off")
                col.caption(f"Факт: {actual:g} ч")

            editor_state = st.session_state.get(editor_key, {})
            dirty = any(editor_state.get(k) for k in ("edited_rows", "added_rows", "deleted_rows")) \
                or norm_val != float(saved_norm)
            if dirty:
                st.caption(":orange[Есть несохранённые изменения.]")

            if st.button("Сохранить", type="primary"):
                store.save_person(selected, kept, norm_val)
                versions[selected] = ver + 1
                baselines.pop(editor_key, None)
                st.toast("Часы сохранены")
                st.rerun()

# --- Планирование: загрузка выгрузки с доски, часы всей команды, отчёт ------------------
with tab_planning:
    with st.expander("Загрузить выгрузку с доски на новую неделю", expanded=not plan):
        st.caption("Файл становится планом недели для всей команды. Уже введённые часы по совпадающим "
                   "тикетам и неделям сохраняются, отметки «заполнено» сбрасываются, если недели в файле новые.")
        uploaded = st.file_uploader("Еженедельный экспорт с Agile-доски (.xlsx)", type="xlsx")
        if uploaded and st.button("Сделать планом недели"):
            people_data, week_labels = parse_workbook(uploaded.getvalue())
            if not week_labels:
                st.error("Не удалось найти недельные колонки часов в файле. Проверьте структуру листов исполнителей.")
            else:
                store.save_plan(uploaded.name, people_data, week_labels)
                st.session_state.pop("baselines", None)
                st.rerun()

    if plan:
        week_labels = plan["week_labels"]
        everyone = store.load_all()
        st.button("Обновить", help="Подтянуть часы, сохранённые коллегами с момента открытия страницы")
        shown_weeks = visible_weeks(week_labels, key="all_weeks_planning")
        week_cols = [clean_week_label(w) for w in shown_weeks]
        fact_cols = [f"{wc} факт" for wc in week_cols]

        rows = []
        for name, (df, norm, saved_at) in everyone.items():
            row = {"Сотрудник": name, "Сохранено": fmt_saved(saved_at), "Норма, ч/нед": norm}
            for wl, wc, fc in zip(shown_weeks, week_cols, fact_cols):
                row[wc] = df[wl].sum()
                row[fc] = df[actual_label(wl)].sum()
            rows.append(row)
        paired_cols = [c for pair in zip(week_cols, fact_cols) for c in pair]
        summary = pd.DataFrame(rows, columns=["Сотрудник", "Сохранено", "Норма, ч/нед", *paired_cols])

        not_filled = [name for name, (_, _, saved_at) in everyone.items() if not saved_at]
        filled = len(everyone) - len(not_filled)
        if not_filled:
            st.warning(f"Заполнили часы: {filled} из {len(everyone)}. Ещё не заполнили: " + ", ".join(not_filled))
        else:
            st.success(f"Все {len(everyone)} сотрудников заполнили часы.")

        st.subheader("Сводка по команде")
        st.caption("По каждой неделе — план и рядом факт. Подсветка перегруза — по плану.")

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

        styled = (summary.style.apply(highlight_overload, axis=1)
                  .format(precision=2, subset=paired_cols)
                  .format("{:g}", subset=["Норма, ч/нед"]))
        st.dataframe(styled, use_container_width=True, hide_index=True)

        if not summary.empty:
            col1, col2 = st.columns(2)
            with col1:
                long_df = summary.melt(id_vars=["Сотрудник"], value_vars=week_cols,
                                       var_name="Неделя", value_name="Часы")
                fig = px.bar(long_df, x="Сотрудник", y="Часы", color="Неделя", barmode="group",
                             title="Загрузка по неделям, по каждому сотруднику")
                st.plotly_chart(fig, use_container_width=True)

            with col2:
                team_total = summary[week_cols].sum()
                team_norm = summary["Норма, ч/нед"].sum()
                fig2 = go.Figure()
                fig2.add_bar(x=week_cols, y=team_total.values, name="Запланировано, ч")
                fig2.add_bar(x=week_cols, y=summary[fact_cols].sum().values, name="Факт, ч")
                fig2.add_scatter(x=week_cols, y=[team_norm] * len(week_cols),
                                 mode="lines+markers", name="Норма команды, ч")
                fig2.update_layout(title="Суммарная загрузка команды по неделям vs норма")
                st.plotly_chart(fig2, use_container_width=True)

            overloaded = summary[summary.apply(
                lambda r: any(r[c] > r["Норма, ч/нед"] > 0 for c in week_cols), axis=1)]
            if not overloaded.empty:
                st.warning("Перегружены на одной или нескольких неделях: " + ", ".join(overloaded["Сотрудник"]))
            else:
                st.success("Перегрузов не обнаружено.")

        st.divider()
        st.subheader("Отчёт на неделю")
        st.caption("Excel с сохранёнными часами всех сотрудников (план и факт по всем неделям плана): "
                   "лист на каждого, сводка по команде с формулами и графиками.")
        if st.button("Сформировать Excel"):
            people_data = {name: df for name, (df, _, _) in everyone.items()}
            norms = {name: norm for name, (_, norm, _) in everyone.items()}
            st.download_button(
                "Скачать отчёт",
                data=build_export_workbook(people_data, week_labels, norms),
                file_name=f"Загрузка_команды_{datetime.now(timezone.utc):%Y-%m-%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
