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
    DEFAULT_STATUSES,
    DELETE_COL,
    FIXED_COLS,
    PRIORITIES,
    PRIORITY_COL,
    STATUS_COL,
    actual_label,
    build_export_workbook,
    clean_week_label,
    drop_marked_rows,
    hour_cols,
    nearest_weeks,
    parse_workbook,
    sort_by_priority,
)
from fact import (
    DATE_COL,
    GROUP_COLS,
    HOURS_COL,
    build_comparison_workbook,
    compare_plan_fact,
    deviation_level,
    fact_breakdown,
    parse_fact_workbook,
    week_title,
)
from storage import Store, db_path_from_env
from table_filter import apply_filter, filter_controls, filter_table, merge_view_edits

st.set_page_config(page_title="Загрузка команды", layout="wide")

ROW_ID = "_row_id"  # hidden row id that ties edits in a filtered editor back to the full table


@st.cache_resource(show_spinner=False)
def get_store(path: str) -> Store:
    return Store(path)


def visible_weeks(week_labels):
    """В колонках только 2 ближайшие недели плана: прошлая и текущая (в начале недели можно внести
    факт за прошедшую). Часы остальных недель плана остаются в базе и при сохранении не теряются."""
    return nearest_weeks(week_labels, date.today())


def status_reference(store):
    """Справочник статусов: стандартные и свои; свой можно добавить или убрать."""
    with st.expander("Справочник статусов"):
        statuses = store.statuses()
        st.caption("Статусы для колонки «Статус»: " + ", ".join(statuses) + ". Добавленный статус сразу "
                   "появляется в списке у всех сотрудников.")
        with st.form("add_status", clear_on_submit=True, border=False):
            c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
            new_status = c1.text_input("Новый статус", placeholder="например, на ревью")
            if c2.form_submit_button("Добавить статус") and new_status.strip():
                if store.add_status(new_status):
                    st.toast(f"Статус «{new_status.strip()}» добавлен")
                    st.rerun()
                st.warning(f"Статус «{new_status.strip()}» уже есть в справочнике.")
        custom = statuses[len(DEFAULT_STATUSES):]
        if custom:
            c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
            to_remove = c1.selectbox("Убрать свой статус", custom, index=None, placeholder="Выберите статус")
            if c2.button("Убрать", disabled=not to_remove):
                store.remove_status(to_remove)
                st.rerun()


def fmt_saved(ts):
    if not ts:
        return "—"
    return datetime.fromisoformat(ts).strftime("%d.%m %H:%M")


store = get_store(str(db_path_from_env()))
plan = store.get_plan()

st.title("Еженедельное планирование загрузки команды")
if plan:
    weeks = ", ".join(clean_week_label(w) for w in visible_weeks(plan["week_labels"]))
    st.caption(f"План недели: {plan['file_name']} (загружен {fmt_saved(plan['uploaded_at'])}). Недели: {weeks}")

tab_mine, tab_planning, tab_history, tab_fact = st.tabs(["Мои часы", "Планирование", "История", "План vs факт"])

# --- Мои часы: каждый сотрудник заполняет и сохраняет свои часы ------------------------
with tab_mine:
    if not plan:
        st.info("План недели ещё не загружен. Загрузите еженедельный экспорт с доски во вкладке "
                "«Планирование» — после этого каждый сотрудник сможет заполнить здесь свои часы.")
    else:
        week_labels = plan["week_labels"]
        st.caption("Выберите себя, впишите плановые часы по тикетам на неделю и нажмите «Сохранить». "
                   "В конце недели впишите в колонку «факт» сколько часов реально ушло и снова сохраните. Для задач, "
                   "которых нет на доске, добавьте строку кнопкой «+» внизу таблицы и впишите название и часы. "
                   "Чтобы убрать тикет или работу, отметьте "
                   "галочку «Удалить» в начале строки: строка удалится при сохранении. Статус и приоритет "
                   "выбираются из списка; после сохранения строки встают по приоритету — «срочно» наверху.")
        selected = st.selectbox("Сотрудник", store.people(), index=None, placeholder="Выберите себя из списка")
        status_reference(store)
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
                df = sort_by_priority(df)
                for col in (STATUS_COL, PRIORITY_COL):  # empty cell in the editor, not an empty-string option
                    df[col] = df[col].astype(object).where(df[col] != "", None)
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

            statuses = store.statuses()
            column_config = {
                DELETE_COL: st.column_config.CheckboxColumn(
                    help="Отметьте, чтобы удалить строку. Удаление применяется при нажатии «Сохранить».",
                    default=False),
                "Тикет": st.column_config.TextColumn(),
                "Название": st.column_config.TextColumn(width="large"),
                "Колонка": st.column_config.TextColumn(),
                "Доска": st.column_config.TextColumn(),
                "Срок": st.column_config.DateColumn(format="DD.MM.YYYY"),
                STATUS_COL: st.column_config.SelectboxColumn(
                    options=[*statuses, *sorted(set(base_df[STATUS_COL].dropna()) - set(statuses))],
                    help="Статус задачи. Свой статус можно добавить в «Справочнике статусов» над таблицей."),
                PRIORITY_COL: st.column_config.SelectboxColumn(
                    options=PRIORITIES,
                    help="Срочно > высокий > средний > низкий. После сохранения строки встают по приоритету."),
            }
            for wl in week_labels:
                column_config[wl] = st.column_config.NumberColumn(
                    f"{clean_week_label(wl)} план", min_value=0.0, step=0.01, format="%g")
                column_config[actual_label(wl)] = st.column_config.NumberColumn(
                    f"{clean_week_label(wl)} факт", min_value=0.0, step=0.01, format="%g",
                    help="Сколько часов реально ушло на задачу за неделю. Заполняется в конце недели.")

            shown_weeks = visible_weeks(week_labels)
            # Hidden weeks stay in the data (column_order only hides them), so saving keeps their hours.
            column_order = [DELETE_COL, *FIXED_COLS, *hour_cols(shown_weeks)]

            # Search and filters show part of the rows in the editor. The editor's baseline must stay
            # fixed while the filter is unchanged, so edits are merged into the full table every run and
            # that merged table becomes the new baseline only when the filter changes (new editor key).
            works = st.session_state.setdefault("filter_works", {})
            if editor_key not in works:
                works[editor_key] = {"full": base_df.assign(**{ROW_ID: range(len(base_df))}),
                                     "spec": None, "merged": None, "gen": 0, "carried": False,
                                     "view_dirty": False}
            work = works[editor_key]
            labels = {c: column_config[c]["label"] for c in hour_cols(shown_weeks)}
            spec = filter_controls(work["full"], key=f"filter_{editor_key}", columns=column_order[1:],
                                   labels=labels)
            if spec != work["spec"]:
                if work["merged"] is not None:
                    work["carried"] = work["carried"] or work["view_dirty"]
                    work["full"] = work["merged"]
                work["spec"], work["gen"] = spec, work["gen"] + 1
            view = apply_filter(work["full"], spec)
            if len(view) != len(work["full"]):
                st.caption(f"Показано строк: {len(view)} из {len(work['full'])}. Правки в скрытых фильтром "
                           "строках не теряются, итоги ниже — по всем строкам.")
            view_key = f"{editor_key}_view{work['gen']}"
            edited = st.data_editor(view, column_config=column_config, column_order=column_order,
                                    hide_index=True, use_container_width=True, num_rows="dynamic",
                                    key=view_key)
            work["merged"] = merge_view_edits(work["full"], view, edited, ROW_ID)
            kept = drop_marked_rows(work["merged"]).drop(columns=ROW_ID)
            marked = len(work["merged"]) - len(kept)
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

            editor_state = st.session_state.get(view_key, {})
            work["view_dirty"] = any(editor_state.get(k) for k in ("edited_rows", "added_rows", "deleted_rows"))
            dirty = work["view_dirty"] or work["carried"] or norm_val != float(saved_norm)
            if dirty:
                st.caption(":orange[Есть несохранённые изменения.]")

            if st.button("Сохранить", type="primary"):
                store.save_person(selected, kept, norm_val)
                versions[selected] = ver + 1
                baselines.pop(editor_key, None)
                works.pop(editor_key, None)
                st.toast("Часы сохранены")
                st.rerun()

# --- Планирование: загрузка выгрузки с доски, часы всей команды, отчёт ------------------
with tab_planning:
    with st.expander("Загрузить выгрузку с доски на новую неделю", expanded=not plan):
        st.caption("Файл становится планом недели для всей команды. Уже введённые часы по совпадающим "
                   "тикетам и неделям сохраняются, отметки «заполнено» сбрасываются, если недели в файле новые. "
                   "Недели, которых нет в новом файле, со всеми часами переходят во вкладку «История». "
                   "Пустые строки (без тикета, названия и часов) удаляются.")
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
        shown_weeks = visible_weeks(week_labels)
        week_cols = [clean_week_label(w) for w in shown_weeks]
        fact_cols = [f"{wc} факт" for wc in week_cols]

        rows = []
        for name, (df, norm, saved_at) in everyone.items():
            row = {"Сотрудник": name, "Сохранено": fmt_saved(saved_at), "Норма, ч/нед": norm}
            for wl, wc, fc in zip(shown_weeks, week_cols, fact_cols):
                row[wc] = round(df[wl].sum(), 2)
                row[fc] = round(df[actual_label(wl)].sum(), 2)
            rows.append(row)
        paired_cols = [c for pair in zip(week_cols, fact_cols) for c in pair]
        summary = pd.DataFrame(rows, columns=["Сотрудник", "Сохранено", "Норма, ч/нед", *paired_cols])

        # Не заполнил: ни разу не сохранял или на какой-то из показанных недель 0 ч плана
        # (например, лишние строки удалил, а часы ещё не проставил).
        not_filled = []
        for name, (_, _, saved_at) in everyone.items():
            row = summary.loc[summary["Сотрудник"] == name].iloc[0]
            empty_weeks = [wc for wc in week_cols if row[wc] == 0]
            if not saved_at:
                not_filled.append(name)
            elif empty_weeks:
                not_filled.append(f"{name} (0 ч на {', '.join(empty_weeks)})")
        filled = len(everyone) - len(not_filled)
        if not_filled:
            st.warning(f"Заполнили часы: {filled} из {len(everyone)}. Ещё не заполнили: " + ", ".join(not_filled))
        else:
            st.success(f"Все {len(everyone)} сотрудников заполнили часы.")

        st.subheader("Сводка по команде")
        st.caption("По каждой неделе — план и рядом факт. Подсветка — по плану: оранжевый — часы не заполнены "
                   "(0 ч), жёлтый — меньше нормы, зелёный — ровно норма, красный — больше нормы.")

        def highlight_load(row):
            norm = row["Норма, ч/нед"]
            missing = any(row[wc] == 0 for wc in week_cols)
            styles = []
            for col in summary.columns:
                if col == "Сотрудник" and missing:
                    styles.append("background-color:#FCE5CD")
                elif col not in week_cols:
                    styles.append("")
                elif row[col] == 0:
                    styles.append("background-color:#FCE5CD")
                elif row[col] > norm > 0:
                    styles.append("background-color:#F4CCCC")
                elif row[col] < norm:
                    styles.append("background-color:#FFF2CC")
                else:
                    styles.append("background-color:#D9EAD3")
            return styles

        shown_summary = filter_table(summary, key="filter_planning_summary")
        styled = (shown_summary.style.apply(highlight_load, axis=1)
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
        st.caption("Excel с сохранёнными часами всех сотрудников (план и факт по двум ближайшим неделям): "
                   "лист на каждого с тикетами по приоритету, сводка по команде с формулами и графиками.")
        if st.button("Сформировать Excel"):
            people_data = {name: df for name, (df, _, _) in everyone.items()}
            norms = {name: norm for name, (_, norm, _) in everyone.items()}
            st.download_button(
                "Скачать отчёт",
                data=build_export_workbook(people_data, shown_weeks, norms),
                file_name=f"Загрузка_команды_{datetime.now(timezone.utc):%Y-%m-%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# --- История: прошедшие недели, ушедшие из плана при загрузке новой выгрузки -------------
with tab_history:
    past_weeks = store.history_weeks()
    if not past_weeks:
        st.info("История пока пустая. Когда загружается выгрузка с доски, недели, которых в ней уже нет, "
                "попадают сюда со всеми часами (план и факт).")
    else:
        chosen = st.selectbox("Неделя", past_weeks, format_func=lambda w: clean_week_label(w["label"]),
                              key="history_week")
        label, people = store.load_history_week(chosen["key"])
        st.caption(f"Из плана {chosen['file_name']}, в истории с {fmt_saved(chosen['archived_at'])}.")

        hist_rows = [{"Сотрудник": name, "Сохранено": fmt_saved(saved_at), "Норма, ч/нед": norm,
                      "План, ч": df[label].sum(), "Факт, ч": df[actual_label(label)].sum()}
                     for name, (df, norm, saved_at) in people.items()]
        hist = pd.DataFrame(hist_rows)
        # Фактическая отметка из учёта времени за эту неделю, если её загружали во вкладке «План vs факт».
        try:
            hist_fact = store.load_fact(chosen["key"])
        except KeyError:
            hist_fact = None
        hist_hour_cols = ["План, ч", "Факт, ч"]
        if hist_fact is not None:
            hist_cmp, hist_fact = compare_plan_fact(people, label, hist_fact)
            fact_by_name = hist_cmp.set_index("Сотрудник")["Факт по отметке, ч"]
            hist["Факт по отметке, ч"] = hist["Сотрудник"].map(fact_by_name).fillna(0.0)
            hist["Отклонение, ч"] = (hist["Факт по отметке, ч"] - hist["План, ч"]).round(2)
            hist_hour_cols += ["Факт по отметке, ч", "Отклонение, ч"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("План команды", f"{hist['План, ч'].sum():g} ч")
        c2.metric("Факт команды", f"{hist['Факт, ч'].sum():g} ч")
        c3.metric("Факт по отметке", f"{hist_fact['Часы'].sum():g} ч" if hist_fact is not None else "—")
        c4.metric("Норма команды", f"{hist['Норма, ч/нед'].sum():g} ч")
        if hist_fact is None:
            st.caption("Фактическую отметку за эту неделю не загружали. Её можно загрузить во вкладке «План vs факт».")
        else:
            only_fact = sorted(set(hist_fact["Сотрудник"]) - set(people))
            if only_fact:
                st.caption("Есть в отметке, но не было в плане: " + ", ".join(only_fact))
        shown_hist = filter_table(hist, key=f"filter_history_{chosen['key']}")
        st.dataframe(shown_hist.style.format(precision=2, subset=hist_hour_cols)
                     .format("{:g}", subset=["Норма, ч/нед"]),
                     use_container_width=True, hide_index=True)

        for name, (df, _, _) in people.items():
            with st.expander(f"{name}: тикеты" + (" и факт по отметке" if hist_fact is not None else "")):
                if df.empty:
                    st.caption("Часов на эту неделю не было.")
                else:
                    df = df.assign(Срок=pd.to_datetime(df["Срок"]))
                    df = df.rename(columns={label: "План, ч", actual_label(label): "Факт, ч"})
                    st.dataframe(filter_table(df, key=f"filter_history_{chosen['key']}_{name}"),
                                 use_container_width=True, hide_index=True,
                                 column_config={"Срок": st.column_config.DateColumn(format="DD.MM.YYYY")})
                if hist_fact is not None:
                    mine = hist_fact[hist_fact["Сотрудник"] == name]
                    st.markdown(f"**Факт по отметке: {GROUP_COLS[0]}**")
                    if mine.empty:
                        st.caption("Отметок за эту неделю нет.")
                    else:
                        st.dataframe(fact_breakdown(mine, GROUP_COLS[0]), use_container_width=True, hide_index=True,
                                     column_config={"Доля": st.column_config.NumberColumn(format="percent"),
                                                    "Часы": st.column_config.NumberColumn(format="%.2f")})

        st.download_button(
            "Скачать Excel за эту неделю",
            data=build_export_workbook({n: df for n, (df, _, _) in people.items()}, [label],
                                       {n: norm for n, (_, norm, _) in people.items()}),
            file_name=f"Загрузка_команды_{chosen['key']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# --- План vs факт: плановые часы недели против фактической отметки из учёта времени ------
with tab_fact:
    with st.expander("Загрузить фактическую отметку", expanded=not store.fact_weeks()):
        st.caption("Выгрузка из учёта рабочего времени («Сводная таблица по выполненным работам», .xlsx): "
                   "строка на сотрудника, день и работу с колонками «Дата», «Фамилия И.О.», «Часы», «Проект» "
                   "и др. Часы раскладываются по неделям; недели из файла заменяют ранее загруженный факт за них.")
        fact_file = st.file_uploader("Выгрузка факта (.xlsx)", type="xlsx", key="fact_upload")
        if fact_file and st.button("Загрузить факт"):
            fact_rows = parse_fact_workbook(fact_file.getvalue())
            if fact_rows.empty:
                st.error("Не нашлось строк с часами. Нужны колонки «Дата», «Фамилия И.О.» и «Часы».")
            else:
                keys = store.save_fact(fact_file.name, fact_rows)
                st.session_state["fact_week"] = max(keys)
                st.rerun()

    fact_weeks = store.fact_weeks()
    if not fact_weeks:
        st.info("Загрузите выгрузку фактической отметки, чтобы сравнить её с планом команды за неделю.")
    else:
        week_keys = [w["key"] for w in fact_weeks]
        if st.session_state.get("fact_week") not in week_keys:
            st.session_state["fact_week"] = week_keys[0]
        key = st.selectbox("Неделя", week_keys, format_func=week_title, key="fact_week")
        info = next(w for w in fact_weeks if w["key"] == key)
        fact_rows = store.load_fact(key)
        label, plan_people = store.plan_for_week(key)
        days = sorted(fact_rows[DATE_COL].unique())
        st.caption(f"Факт из {info['file_name']} (загружен {fmt_saved(info['uploaded_at'])}), "
                   f"дни с отметками: {days[0]:%d.%m}–{days[-1]:%d.%m} ({len(days)} дн.).")
        if not label:
            st.warning(f"Плана на неделю {week_title(key)} нет ни в текущем плане, ни в истории: "
                       "видно только, куда ушли часы.")

        summary, fact_rows = compare_plan_fact(plan_people, label, fact_rows)
        team_plan, team_fact = summary["План, ч"].sum(), summary["Факт по отметке, ч"].sum()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("План команды", f"{team_plan:g} ч")
        c2.metric("Факт команды", f"{team_fact:g} ч", f"{team_fact - team_plan:+g} ч к плану", delta_color="off")
        c3.metric("Факт / план", f"{team_fact / team_plan:.0%}" if team_plan else "—")
        c4.metric("Норма команды", f"{summary['Норма, ч'].sum():g} ч")

        only_fact = summary.loc[summary["Норма, ч"].isna(), "Сотрудник"].tolist()
        no_fact = summary.loc[summary["Норма, ч"].notna() & (summary["Факт по отметке, ч"] == 0), "Сотрудник"]
        if only_fact:
            st.warning("Есть в отметке, но нет в планировщике: " + ", ".join(only_fact))
        if not no_fact.empty:
            st.warning("Нет фактической отметки за неделю: " + ", ".join(no_fact))

        st.subheader("План и факт по сотрудникам")
        st.caption("«Факт по отметке» — из загруженной выгрузки, «Факт в планировщике» — что сотрудник сам вписал "
                   "в колонку «факт» во вкладке «Мои часы». Подсветка отклонения отметки от плана: "
                   "зелёный — до 10%, жёлтый — до 25%, красный — больше или план не заполнен.")
        colors = {"ok": "#D9EAD3", "warn": "#FFF2CC", "bad": "#F4CCCC"}

        def highlight_deviation(row):
            color = colors[deviation_level(row["План, ч"], row["Факт по отметке, ч"])]
            return [f"background-color:{color}" if c in ("Отклонение, ч", "Факт / план") else ""
                    for c in summary.columns]

        hour_fmt = ["План, ч", "Факт в планировщике, ч", "Факт по отметке, ч"]
        shown_summary = filter_table(summary, key=f"filter_fact_summary_{key}")
        st.dataframe(shown_summary.style.apply(highlight_deviation, axis=1)
                     .format(precision=2, subset=hour_fmt)
                     .format("{:+.2f}", subset=["Отклонение, ч"])
                     .format("{:.0%}", subset=["Факт / план"], na_rep="—")
                     .format("{:g}", subset=["Норма, ч"], na_rep="—"),
                     use_container_width=True, hide_index=True)

        group_by = st.selectbox("Куда ушли часы — группировать по", GROUP_COLS, key="fact_group_by")
        long_df = summary.melt(id_vars=["Сотрудник"], value_vars=["План, ч", "Факт по отметке, ч"],
                               var_name="Часы по", value_name="Часы")
        st.plotly_chart(px.bar(long_df, x="Сотрудник", y="Часы", color="Часы по", barmode="group",
                               title="План и факт по сотрудникам"), use_container_width=True)

        # Самые затратные группы — сверху столбцов и первыми в легенде, с долей от всей отметки команды.
        by_person = fact_rows.groupby(["Сотрудник", group_by])[HOURS_COL].sum().reset_index()
        group_totals = by_person.groupby(group_by)[HOURS_COL].sum().sort_values(ascending=False)
        share = group_totals / group_totals.sum() if group_totals.sum() else group_totals
        legend_label = {g: f"{g} — {share[g]:.1%}" for g in group_totals.index}
        person_totals = by_person.groupby("Сотрудник")[HOURS_COL].transform("sum")
        by_person = by_person.assign(**{
            "Группа": by_person[group_by].map(legend_label),
            "Доля у сотрудника": by_person[HOURS_COL] / person_totals,
            "Доля в команде": by_person[HOURS_COL] / group_totals.sum(),
        })
        fig_groups = px.bar(
            by_person, x="Сотрудник", y=HOURS_COL, color="Группа",
            category_orders={"Сотрудник": summary["Сотрудник"].tolist(),
                             "Группа": [legend_label[g] for g in reversed(group_totals.index)]},
            custom_data=[group_by, "Доля у сотрудника", "Доля в команде"],
            title=f"Факт по отметке: {group_by} (доля от всей отметки команды)")
        fig_groups.update_traces(hovertemplate="%{x}<br>%{customdata[0]}<br>%{y:.2f} ч — "
                                               "%{customdata[1]:.0%} часов сотрудника, "
                                               "%{customdata[2]:.1%} всей отметки<extra></extra>")
        fig_groups.update_layout(legend={"traceorder": "reversed", "title": {"text": group_by}}, height=650)
        st.plotly_chart(fig_groups, use_container_width=True)

        st.markdown(f"**Команда: факт по «{group_by}»**")
        st.dataframe(filter_table(fact_breakdown(fact_rows, group_by), key=f"filter_fact_team_{key}"),
                     use_container_width=True, hide_index=True,
                     column_config={"Доля": st.column_config.NumberColumn(format="percent"),
                                    "Часы": st.column_config.NumberColumn(format="%.2f")})

        st.subheader("По каждому сотруднику")
        for _, row in summary.iterrows():
            name = row["Сотрудник"]
            mine = fact_rows[fact_rows["Сотрудник"] == name]
            with st.expander(f"{name}: план {row['План, ч']:g} ч, факт {row['Факт по отметке, ч']:g} ч "
                             f"({row['Отклонение, ч']:+g} ч)"):
                left, right = st.columns(2)
                with left:
                    st.markdown("**План в планировщике**")
                    tickets = plan_people[name][0] if name in plan_people and label else None
                    if tickets is None or tickets.empty:
                        st.caption("Плана на эту неделю нет.")
                    else:
                        t = tickets[(tickets[label] != 0) | (tickets[actual_label(label)] != 0)]
                        t = t[["Тикет", "Название", label, actual_label(label)]].rename(
                            columns={label: "План, ч", actual_label(label): "Факт в планировщике, ч"})
                        t = filter_table(t.sort_values("План, ч", ascending=False),
                                         key=f"filter_fact_plan_{key}_{name}")
                        st.dataframe(t, use_container_width=True,
                                     hide_index=True)
                with right:
                    st.markdown(f"**Факт по отметке: {group_by}**")
                    if mine.empty:
                        st.caption("Отметок за эту неделю нет.")
                    else:
                        st.dataframe(filter_table(fact_breakdown(mine, group_by), key=f"filter_fact_mine_{key}_{name}"),
                                     use_container_width=True, hide_index=True,
                                     column_config={"Доля": st.column_config.NumberColumn(format="percent"),
                                                    "Часы": st.column_config.NumberColumn(format="%.2f")})
                        per_day = mine.groupby(DATE_COL)[HOURS_COL].sum()
                        st.caption("По дням: " + ", ".join(f"{d:%d.%m} — {h:g} ч" for d, h in per_day.items()))

        st.download_button(
            "Скачать Excel план vs факт",
            data=build_comparison_workbook(summary, fact_rows, group_by),
            file_name=f"План_факт_{key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
