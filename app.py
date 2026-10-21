"""
Dimax HR v2 — веб-інтерфейс. Запуск:  streamlit run app.py
"""
import os
import time
import uuid
from datetime import datetime

import pandas as pd
import streamlit as st

import hr_core as core

WORK_DIR = os.environ.get("HR_WORK_DIR", "hr_data")
st.set_page_config(page_title="Dimax HR · Відбір кандидатів", page_icon="🧭", layout="wide")


def secret(name):
    try:
        v = st.secrets.get(name)
        if v:
            return str(v)
    except Exception:
        pass
    return os.environ.get(name, "")


# ---------------------------------------------------------------- пароль
pw = secret("APP_PASSWORD")
if pw and not st.session_state.get("auth"):
    st.title("🧭 Dimax HR")
    entered = st.text_input("Пароль", type="password")
    if entered == pw:
        st.session_state.auth = True
        st.rerun()
    elif entered:
        st.error("Невірний пароль")
    st.stop()


@st.cache_resource
def get_store():
    return core.Store(WORK_DIR)


store = get_store()

# ---------------------------------------------------------------- налаштування
with st.sidebar:
    st.header("Налаштування")
    mode_label = st.radio("Режим запитів до ШІ", ["💳 Платний: швидко", "🆓 Безкоштовний: повільно"],
                          help="Платний: 8 запитів паралельно.\n\nБезкоштовний: 1 запит раз на кілька секунд і денний ліміт Google. "
                               "Коли ліміт закінчиться, оцінка зупиниться; наступного дня натисніть «Оцінити» ще раз — продовжить з місця зупинки.")
    mode = "free" if mode_label.startswith("🆓") else "paid"
    provider = st.selectbox("Постачальник ШІ", ["gemini", "openai"], format_func={"gemini": "Google Gemini", "openai": "OpenAI"}.get)
    if provider == "openai" and mode == "free":
        st.warning("OpenAI не має безкоштовного тарифу — використовується платний режим.")
        mode = "paid"
    model = st.text_input("Модель", value=core.DEFAULT_MODEL[provider])
    server_key = secret("LLM_API_KEY")
    if server_key:
        st.success("API-ключ задано на сервері")
        api_key = server_key
    else:
        api_key = st.text_input("API-ключ", type="password", help="Ключ Google AI Studio (aistudio.google.com) або OpenAI. Ніде не зберігається.")
    pause = None
    if mode == "free":
        pause = st.number_input("Пауза між запитами, с", 1.0, 120.0, 5.0, 1.0)
        st.info("На безкоштовному тарифі Google може використовувати надіслані дані для покращення своїх продуктів. Резюме містять персональні дані.")
    smart = st.checkbox("Економний фільтр", value=True,
                        help="Для вакансій, де українська чи російська обов'язкова, не надсилати в ШІ резюме без жодних ознак "
                             "UA/СНД/Балтії (міста, вузи, телефонні коди, кирилиця, прізвища). Економить 60–80% запитів.")
    threshold = st.slider("Поріг «можливо»", 40, 70, 50, 5, help="Бали від порогу до 69 — «можливо», 70+ — «підходить».")

settings = core.Settings(api_key=api_key, mode=mode, provider=provider, model=model, pause=pause, smart_filter=smart)

st.title("🧭 Dimax HR: відбір кандидатів")
tab_v, tab_c, tab_r = st.tabs(["📋 Вакансії", "👥 Кандидати та оцінка", "📊 Результати"])

PILL = {"good": "🟢", "warn": "🟠", "neutral": "⚪"}

# =====================================================================
# ВАКАНСІЇ
# =====================================================================
def load_editor(v):
    ss = st.session_state
    ss.ed = v
    ss.ed_n = ss.get("ed_n", 0) + 1
    ss.v_title = v["title"]
    ss.v_place = v["location"].get("place", "")
    ss.v_mode = v["location"].get("mode", "pref")
    ss.v_exp = v["experience"].get("years", "")
    ss.v_exp_req = v["experience"].get("req", "pref")
    ss.v_exp_field = v["experience"].get("field", "")
    ss.v_edu = v["education"].get("type", "")
    ss.v_edu_req = v["education"].get("req", "pref")
    ss.v_notes = v.get("notes", "")
    ss.v_prompt = v.get("prompt", "")
    ss.v_prompt_for = v.get("promptFor", "")
    ss.v_prompt_edited = v.get("promptEdited", False)


def editor_vacancy():
    ss = st.session_state
    n = ss.ed_n
    langs = ss.get(f"langs_df_{n}")
    tools = ss.get(f"tools_df_{n}")
    v = dict(ss.ed)
    v.update(title=ss.v_title.strip(), notes=ss.v_notes, prompt=ss.v_prompt, promptFor=ss.v_prompt_for, promptEdited=ss.v_prompt_edited,
             location={"place": ss.v_place, "mode": ss.v_mode},
             experience={"years": ss.v_exp, "field": ss.v_exp_field, "req": ss.v_exp_req},
             education={"type": ss.v_edu, "req": ss.v_edu_req})
    rev = {t: k for k, t in core.REQ.items()}
    if langs is not None:
        v["languages"] = [{"lang": str(r["Мова"]).strip(), "level": r["Рівень"] or "вільний", "req": rev.get(r["Важливість"], "pref")}
                          for _, r in langs.iterrows() if str(r["Мова"] or "").strip()]
    if tools is not None:
        v["tools"] = [{"name": str(r["Інструмент / навичка"]).strip(), "req": rev.get(r["Важливість"], "pref")}
                      for _, r in tools.iterrows() if str(r["Інструмент / навичка"] or "").strip()]
    return v


with tab_v:
    if "pending_load" in st.session_state:  # завантаження у форму — до створення віджетів
        load_editor(st.session_state.pop("pending_load"))
    if st.session_state.pop("saved_toast", False):
        st.toast("Вакансію збережено")
    vacs = sorted(store.vacancies.values(), key=lambda v: v["title"].lower())
    left, right = st.columns([1, 2.6], gap="large")
    with left:
        st.subheader("Вакансії")
        for v in vacs:
            cls, txt = core.prompt_state(v)
            n = sum(1 for r in store.results.values() if v["id"] in r.get("results", {}))
            sel = st.session_state.get("ed", {}).get("id") == v["id"]
            if st.button(f"{'▶ ' if sel else ''}{v['title']}", key=f"pick_{v['id']}", width="stretch", type="primary" if sel else "secondary"):
                load_editor(v)
                st.rerun()
            st.caption(f"{PILL[cls]} {txt} · оцінено: {n}")
        st.divider()
        if st.button("➕ Нова вакансія", width="stretch"):
            load_editor(core.blank_vacancy())
            st.rerun()
        tpl = st.selectbox("З шаблону", ["—"] + list(core.TEMPLATES))
        if tpl != "—" and st.button("Створити з шаблону", width="stretch"):
            load_editor(core.from_template(tpl))
            st.rerun()
        with st.expander("Резервна копія вакансій"):
            st.download_button("⬇️ Скачати vacancies.json", data=core.json.dumps(store.vacancies, ensure_ascii=False, indent=1),
                               file_name="vacancies.json", mime="application/json")
            up = st.file_uploader("Завантажити vacancies.json", type=["json"], key="vac_up")
            if up is not None and st.button("Імпортувати (додати/замінити)"):
                try:
                    data = core.json.loads(up.getvalue())
                    store.vacancies.update({k: {**core.blank_vacancy(), **v, "id": k} for k, v in data.items()})
                    store.save_vacancies()
                    st.success(f"Імпортовано: {len(data)}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Не вдалося прочитати файл: {e}")

    with right:
        if "ed" not in st.session_state:
            st.info("Оберіть вакансію ліворуч або створіть нову. Критерії: назва посади, мови, інструменти, локація, досвід, освіта.")
        else:
            ss = st.session_state
            n = ss.ed_n
            st.subheader(ss.ed["title"] or "Нова вакансія")
            st.markdown("**Основні вимоги**")
            st.text_input("Назва посади", key="v_title", placeholder="Напр.: MEA Sales Manager")
            st.caption("Знання мов")
            ss[f"langs_df_{n}"] = st.data_editor(pd.DataFrame([{"Мова": l["lang"], "Рівень": l["level"], "Важливість": core.REQ.get(l["req"], "бажано")} for l in ss.ed["languages"]]
                                        or [{"Мова": "", "Рівень": "вільний", "Важливість": "обов'язково"}]),
                           key=f"langs_ed_{n}", num_rows="dynamic", width="stretch", hide_index=True,
                           column_config={"Мова": st.column_config.TextColumn(width="large"),
                                          "Рівень": st.column_config.SelectboxColumn(options=core.LEVELS, default="вільний"),
                                          "Важливість": st.column_config.SelectboxColumn(options=list(core.REQ.values()), default="обов'язково")}).fillna("")
            st.caption("Інструменти, навички, програми")
            ss[f"tools_df_{n}"] = st.data_editor(pd.DataFrame([{"Інструмент / навичка": t["name"], "Важливість": core.REQ.get(t["req"], "бажано")} for t in ss.ed["tools"]]
                                        or [{"Інструмент / навичка": "", "Важливість": "обов'язково"}]),
                           key=f"tools_ed_{n}", num_rows="dynamic", width="stretch", hide_index=True,
                           column_config={"Інструмент / навичка": st.column_config.TextColumn(width="large"),
                                          "Важливість": st.column_config.SelectboxColumn(options=list(core.REQ.values()), default="обов'язково")}).fillna("")
            c1, c2 = st.columns([1.2, 1.4])
            c1.text_input("Локація", key="v_place", placeholder="Місто, країна")
            c2.selectbox("Вимога до локації", list(core.LOC_MODES), key="v_mode", format_func=core.LOC_MODES.get)

            st.markdown("**Додаткові вимоги** · необов'язково")
            c1, c2, c3, c4 = st.columns([1.3, 1, 1.3, 1])
            c1.selectbox("Досвід", core.EXP_OPTIONS, key="v_exp", format_func=lambda x: x or "Не важливо")
            c2.selectbox("Важливість досвіду", list(core.REQ), key="v_exp_req", format_func=core.REQ.get)
            c3.selectbox("Освіта", core.EDU_OPTIONS, key="v_edu", format_func=lambda x: x or "Не важливо")
            c4.selectbox("Важливість освіти", list(core.REQ), key="v_edu_req", format_func=core.REQ.get)
            st.text_input("Досвід у сфері", key="v_exp_field", placeholder="Напр.: B2B-продажі в регіоні MEA")
            st.text_area("Інші вимоги та нюанси", key="v_notes", height=90,
                         placeholder="Напр.: не розглядаємо кандидатів, які зараз у Росії; кандидати з СНД часто не вказують російську")

            st.divider()
            cur = editor_vacancy()
            cls, txt = core.prompt_state(cur)
            h1, h2 = st.columns([2, 1.4])
            h1.markdown("**Промпт для оцінки**")
            h2.markdown(f"{PILL[cls]} {txt}")
            with st.expander("Критерії, які отримає ШІ"):
                st.text(core.criteria_text(cur))

            if ss.pop("do_gen", False):
                if not api_key:
                    st.error("Введіть API-ключ у бічній панелі.")
                else:
                    with st.spinner("ШІ складає промпт… (10–40 с)"):
                        try:
                            ss.v_prompt = core.generate_prompt(settings, cur)
                            ss.v_prompt_for = core.crit_hash(cur)
                            ss.v_prompt_edited = False
                            st.toast("Промпт готовий. Перевірте його та збережіть вакансію.")
                        except Exception as e:
                            st.error(f"Не вдалося згенерувати промпт: {str(e)[:300]}")

            def _gen():
                if not st.session_state.v_title.strip():
                    st.session_state.gen_err = True
                else:
                    st.session_state.do_gen = True
            st.button("✨ Згенерувати заново" if ss.v_prompt.strip() else "✨ Згенерувати промпт", on_click=_gen, type="primary")
            if ss.pop("gen_err", False):
                st.warning("Спочатку вкажіть назву посади.")
            st.text_area("Промпт (можна редагувати)", key="v_prompt", height=380,
                         on_change=lambda: st.session_state.update(v_prompt_edited=True),
                         placeholder="Натисніть «Згенерувати промпт»: ШІ складе обов'язкові й бажані вимоги, ваги та шкалу. Потім текст можна змінити тут.")

            b1, b2, _, b3 = st.columns([1, 1, 2, 1])
            if b1.button("💾 Зберегти", type="primary"):
                v = editor_vacancy()
                if not v["title"]:
                    st.error("Вкажіть назву посади.")
                elif not v["prompt"].strip():
                    st.error("Згенеруйте або напишіть промпт для оцінки.")
                else:
                    v["id"] = v.get("id") or uuid.uuid4().hex[:8]
                    v["updatedAt"] = int(time.time())
                    store.vacancies[v["id"]] = v
                    store.save_vacancies()
                    ss.pending_load = v
                    ss.saved_toast = True
                    st.rerun()
            if ss.ed.get("id"):
                if b3.button("🗑 Видалити"):
                    ss.confirm_del = ss.ed["id"]
                if ss.get("confirm_del") == ss.ed["id"]:
                    st.warning("Видалити вакансію? Оцінки кандидатів за нею залишаться в історії, але не показуватимуться.")
                    if st.button("Так, видалити"):
                        store.vacancies.pop(ss.ed["id"], None)
                        store.save_vacancies()
                        for k in ("ed", "confirm_del"):
                            ss.pop(k, None)
                        st.rerun()

# =====================================================================
# КАНДИДАТИ
# =====================================================================
with tab_c:
    st.subheader("1. Завантажте кандидатів")
    c1, c2 = st.columns(2)
    with c1:
        csvs = st.file_uploader("CSV з Betterteam (можна кілька)", type=["csv"], accept_multiple_files=True, key="csvs")
        if csvs and st.button("Додати з CSV", type="primary"):
            try:
                for f in csvs:
                    f.seek(0)
                rows, new, upd = core.import_csvs(store, [(f.name, f) for f in csvs])
                st.success(f"Рядків: {rows} · нових кандидатів: {new} · уже були: {upd}")
            except Exception as e:
                st.error(f"Не вдалося прочитати CSV: {e}")
    with c2:
        files = st.file_uploader("Файли резюме або ZIP (необов'язково)", type=["pdf", "docx", "doc", "txt", "zip"], accept_multiple_files=True, key="cvfiles",
                                 help="Якщо посилання з CSV не відкриваються або кандидатів немає в CSV. Файли прив'язуються до кандидатів "
                                      "за назвою файлу з посилання або за ім'ям; решта стають окремими кандидатами.")
        if files and st.button("Додати файли"):
            m, s_ = core.import_files(store, [(f.name, f.getvalue()) for f in files])
            st.success(f"Прив'язано до кандидатів з CSV: {m} · нових кандидатів із файлів: {s_}")

    total = len(store.candidates)
    stat = pd.Series([store.results.get(k, {}).get("status", "") for k in store.candidates], dtype=str)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Кандидатів", total)
    m2.metric("Оцінено", int((stat == "ok").sum()))
    m3.metric("Не прочитано", int((stat == "unreadable").sum()))
    m4.metric("Помилки ШІ", int((stat == "error").sum()))

    st.subheader("2. Оцінка")
    ready = [v for v in sorted(store.vacancies.values(), key=lambda v: v["title"].lower()) if v.get("prompt", "").strip()]
    sel_ids = st.multiselect("Вакансії", [v["id"] for v in ready], default=[v["id"] for v in ready],
                             format_func=lambda i: store.vacancies[i]["title"])
    sel = [v for v in ready if v["id"] in sel_ids]
    rescore = st.checkbox("Переоцінити вже оцінених", help="Інакше оцінюються лише нові кандидати, помилкові та ті, чий промпт змінився.")
    retry_unread = st.checkbox("Повторити неприйняті резюме", help="Ще раз спробувати завантажити резюме, які раніше не вдалося прочитати.")
    keys = [k for k in store.candidates if sel and core.needs_scoring(store, k, sel, rescore)]
    if retry_unread and sel:
        keys += [k for k in store.candidates if k not in keys and store.results.get(k, {}).get("status") == "unreadable"]
    if sel and keys:
        share = (0.35 if smart and all(core.requires_uk_ru(v) for v in sel) else 1.0)
        calls = int(len(keys) * share)
        in_tok = 2500 + 900 * len(sel)
        cost = calls * (in_tok * 0.75 + (300 + 350 * len(sel)) * 3.75) / 1e6
        eta = calls * settings.request_pause / 60 + len(keys) * 1.5 / settings.n_workers / 60
        st.caption(f"До оцінки: **{len(keys)}** кандидатів · запитів до ШІ ≈ {calls} · "
                   + (f"≈ ${cost:.2f} (Gemini Flash) · " if mode == "paid" else "") + f"≈ {max(1, round(eta))} хв")
    elif sel:
        st.caption("Усіх кандидатів уже оцінено за вибраними вакансіями.")
    elif not ready:
        st.warning("Спочатку збережіть хоча б одну вакансію з промптом.")

    if st.button("▶️ Оцінити", type="primary", disabled=not (keys and sel and api_key)):
        for k in keys:
            if store.results.get(k, {}).get("status") == "unreadable":
                store.append_result({"key": k, "status": "", "error": ""})
        bar = st.progress(0.0, text="Старт…")
        info = st.empty()
        st.caption("Щоб зупинити, натисніть **Stop** угорі праворуч. Оцінене зберігається.")
        t0 = time.time()

        def prog(s_, c):
            left = (time.time() - t0) / max(s_["done"], 1) * (s_["total"] - s_["done"])
            bar.progress(s_["done"] / s_["total"], text=f"{s_['done']} / {s_['total']} · залишилось ≈ {left / 60:.0f} хв")
            info.write(f"ШІ: **{s_['ai']}** · лише фільтр: **{s_['filtered']}** · не прочитано: **{s_['unreadable']}** · "
                       f"помилки: **{s_['errors']}** · останній: {c.get('name', '')}")

        res = core.run_scoring(store, settings, keys, sel, rescore, on_progress=prog)
        if res["quota_exhausted"]:
            st.warning(f"⏸ Ліміт запитів вичерпано. Оцінено {res['done']}, залишилось {res['remaining']}. "
                       "Безкоштовний ліміт оновлюється щодня о 10:00 за Києвом — тоді натисніть «Оцінити» ще раз.")
        elif res["errors"]:
            st.warning(f"{res['errors']} резюме не оцінено через помилки ШІ — натисніть «Оцінити» ще раз, щоб повторити лише їх.")
        else:
            st.success(f"Готово за {res['seconds'] / 60:.1f} хв. Перейдіть на вкладку «Результати».")
        if res.get("tokens_in"):
            st.caption(f"Токени: вхід {res['tokens_in']:,}, вихід {res['tokens_out']:,} · запитів {res['calls']}")

    with st.expander("Список кандидатів"):
        df = core.results_frame(store)
        if df.empty:
            st.write("Поки що порожньо.")
        else:
            st.dataframe(df[["Кандидат", "Статус", "Проблема", "Email", "Вакансії відгуку", "Дата відгуку"]], hide_index=True, width="stretch")
    with st.expander("💾 Резервна копія всіх даних"):
        st.write("Вакансії, кандидати, тексти резюме й оцінки в одному ZIP. На Streamlit Cloud файли стираються при перезапуску "
                 "застосунку — скачуйте копію після роботи й відновлюйте її тут.")
        st.download_button("⬇️ Скачати резервну копію", data=core.backup_zip(store), file_name=f"dimax_hr_backup_{datetime.now():%Y-%m-%d_%H%M}.zip",
                           mime="application/zip")
        bk = st.file_uploader("Відновити з резервної копії", type=["zip"], key="restore")
        if bk is not None and st.button("Відновити"):
            try:
                core.restore_zip(store, bk.getvalue())
                st.success("Дані відновлено")
                st.rerun()
            except Exception as e:
                st.error(f"Не вдалося відновити: {e}")
    with st.expander("Очистити всіх кандидатів"):
        st.write("Видаляє кандидатів, тексти резюме та оцінки. Вакансії залишаються.")
        if st.checkbox("Я розумію") and st.button("🗑 Очистити"):
            store.clear_candidates()
            st.rerun()

# =====================================================================
# РЕЗУЛЬТАТИ
# =====================================================================
with tab_r:
    vacs = sorted(store.vacancies.values(), key=lambda v: v["title"].lower())
    df = core.results_frame(store)
    if not vacs or df.empty:
        st.info("Тут з'являться результати після оцінки кандидатів.")
    else:
        top = st.columns([3, 1])
        labels, seen = {}, {}
        for v in vacs:  # однакові назви вакансій → різні підписи
            seen[v["title"]] = seen.get(v["title"], 0) + 1
            labels[v["id"]] = v["title"] if seen[v["title"]] == 1 else f"{v['title']} ({seen[v['title']]})"
        vid = top[0].radio("Вакансія", list(labels), horizontal=True, format_func=labels.get)
        vac = store.vacancies[vid]
        top[1].download_button("⬇️ Excel", data=core.build_excel(store, threshold),
                               file_name=f"Dimax_кандидати_{datetime.now():%Y-%m-%d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
        f1, f2, f3 = st.columns([1, 1, 2])
        mn = f1.slider("Мінімальний бал", 0, 100, 0, 5)
        vf = f2.selectbox("Висновок", ["Усі", "Підходить (70+)", f"Можливо ({threshold}–69)", f"Не підходить (<{threshold})", "Ще не оцінено"])
        q = f3.text_input("Пошук", placeholder="Ім'я, мова, місто, навичка…")

        col = f"bal::{vac['id']}"
        d = df.copy()
        d["Бал"] = d[col]
        d["Коротко"] = d[f"sum::{vac['id']}"]
        d["Промпт змінено"] = d[f"stale::{vac['id']}"]
        scored = d["Бал"].notna()
        good, maybe = (d["Бал"] >= 70).sum(), ((d["Бал"] >= threshold) & (d["Бал"] < 70)).sum()
        st.caption(f"Оцінено {int(scored.sum())} з {len(d)} · 🟢 підходить: {int(good)} · 🟠 можливо: {int(maybe)} · 🔴 не підходить: {int(scored.sum() - good - maybe)}")
        if mn:
            d = d[d["Бал"] >= mn]
        if vf.startswith("Підходить"):
            d = d[d["Бал"] >= 70]
        elif vf.startswith("Можливо"):
            d = d[(d["Бал"] >= threshold) & (d["Бал"] < 70)]
        elif vf.startswith("Не підходить"):
            d = d[d["Бал"] < threshold]
        elif vf == "Ще не оцінено":
            d = d[d["Бал"].isna()]
        if q.strip():
            hay = d[["Кандидат", "Мови", "Локація", "Навички", "Освіта", "Коротко", "Email"]].astype(str).agg(" ".join, axis=1).str.lower()
            d = d[hay.str.contains(q.strip().lower(), regex=False)]
        d = d.sort_values("Бал", ascending=False, na_position="last")
        others = [v for v in vacs if v["id"] != vac["id"]]
        for v in others:
            d[labels[v["id"]]] = d[f"bal::{v['id']}"]
        show = ["Кандидат", "Бал", *[labels[v["id"]] for v in others], "Коротко", "Мови", "Локація", "Досвід, р.", "Тип освіти", "Резюме", "Betterteam", "Статус", "Промпт змінено"]
        ev = st.dataframe(d[show + ["key"]], hide_index=True, width="stretch", height=520, on_select="rerun", selection_mode="single-row",
                          column_order=show,
                          column_config={"Бал": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d"),
                                         "Резюме": st.column_config.LinkColumn(display_text="відкрити"),
                                         "Betterteam": st.column_config.LinkColumn(display_text="картка"),
                                         "Коротко": st.column_config.TextColumn(width="large"),
                                         **{labels[v["id"]]: st.column_config.NumberColumn(format="%d") for v in others}})
        rows = ev.selection.rows if ev and ev.selection else []
        if rows:
            key = d.iloc[rows[0]]["key"]
            c = store.candidates[key]
            r = store.results.get(key, {})
            p = r.get("profile") or {}
            st.markdown(f"### {c.get('name', '')}")
            st.caption(" · ".join(x for x in [c.get("email"), c.get("phone"), p.get("location"), core.lang_str(p)] if x))
            cols = st.columns(max(1, len(r.get("results", {})) + 1))
            for i, v in enumerate([v for v in vacs if v["id"] in r.get("results", {})]):
                x = r["results"][v["id"]]
                with cols[i]:
                    st.markdown(f"**{v['title']}** · {x['score']} — {core.verdict(x['score'])}")
                    for m in x.get("matched", []):
                        st.markdown(f"✅ {m}")
                    for m in x.get("missing", []):
                        st.markdown(f"❌ {m}")
                    st.caption(x.get("summary", ""))
            with cols[-1]:
                st.markdown("**Профіль**")
                st.write(f"Освіта: {p.get('education') or '—'}")
                st.write(f"Досвід: {p.get('experience_years') if p.get('experience_years') is not None else '—'} р.")
                st.write(f"Навички: {', '.join(p.get('tools', [])) or '—'}")
                if r.get("error"):
                    st.error(r["error"])
        else:
            st.caption("Виберіть рядок, щоб побачити, що збіглося і чого бракує.")
