import io
from datetime import date

import streamlit as st
from docx import Document
from docx.shared import Pt

st.set_page_config(page_title="Кардио-осмотр: АГ+ХСН+ИБС", page_icon="🩺", layout="wide")
st.title("Первичный осмотр кардиолога: АГ + ХСН + ИБС")
st.caption("Учебный прототип. Логика упрощена и не заменяет клинические рекомендации.")

# ---------------- 1. Пациент ----------------
with st.expander("1. Пациент", expanded=True):
    c1, c2, c3 = st.columns(3)
    fio = c1.text_input("ФИО", value="Иванов И. И.")
    age = c2.number_input("Возраст, лет", min_value=18, max_value=120, value=58)
    sex = c3.selectbox("Пол", ["Мужской", "Женский"])
    pregnant = False
    if sex == "Женский":
        pregnant = st.checkbox("Беременность")
    first_visit = st.checkbox("Первичный осмотр (иначе — повторный визит на терапии)", value=True)

# ---------------- 2. Антропометрия ----------------
with st.expander("2. Антропометрия"):
    c1, c2, c3 = st.columns(3)
    height = c1.number_input("Рост, см", min_value=120, max_value=220, value=175)
    weight = c2.number_input("Вес, кг", min_value=30, max_value=250, value=96)
    waist = c3.number_input("Окружность талии, см", min_value=40, max_value=180, value=104)

bmi = weight / (height / 100) ** 2

# ---------------- 3. Жалобы и анамнез ----------------
with st.expander("3. Жалобы и анамнез"):
    complaints = st.multiselect("Жалобы", [
        "Головная боль", "Головокружение", "Шум в ушах",
        "Одышка при нагрузке", "Перебои в работе сердца",
        "Боли за грудиной", "Отёки ног",
        "Одышка лёжа / пароксизмальная ночная одышка (ортопноэ)",
        "Прибавка веса более 2 кг за неделю",
        "Уменьшение количества мочи",
        "Снижение памяти", "Мелькание мушек перед глазами",
    ], default=["Головная боль", "Одышка при нагрузке", "Отёки ног"])
    duration = st.text_input("Длительность жалоб / давность повышения АД", value="около 3 лет")
    anamnesis = st.text_area(
        "Анамнез (перенесённые заболевания, операции, аллергия)",
        value="Гипертоническая болезнь в анамнезе около 3 лет, постоянной терапии не получал. "
              "Туберкулёз, гепатиты, ВИЧ отрицает. Операций не было. Аллергологический анамнез не отягощён.",
    )
    heredity = st.checkbox("Семейный анамнез ССЗ в молодом возрасте / АГ у родителей в молодом возрасте")

# ---------------- 4. Факторы риска ----------------
with st.expander("4. Факторы риска"):
    c1, c2 = st.columns(2)
    smoking = c1.selectbox("Курение", ["Никогда не курил", "Курит сейчас", "Курил раньше"])
    alcohol = c1.selectbox("Алкоголь", ["Не употребляет", "Редко", "Регулярно"])
    dyslipid = c2.checkbox("Дислипидемия (ОХ > 4,9; ХС ЛНП > 3,0; ХС ЛВП < 1,0 (муж.) / < 1,2 (жен.); ТГ > 1,7 ммоль/л)")
    glucose = c2.checkbox("Нарушение гликемии натощак / нарушение толерантности к глюкозе")
    inactivity = c2.checkbox("Малоподвижный образ жизни")
    uric_acid = c2.checkbox("Мочевая кислота ≥ 360 мкмоль/л")
    c3a, c3b = st.columns(2)
    early_menopause = c3a.checkbox("Ранняя менопауза") if sex == "Женский" else False
    psych = c3b.checkbox("Психологические и социально-экономические факторы")

# ---------------- 5. Скрининг вторичной АГ ----------------
with st.expander("5. Скрининг вторичной АГ"):
    st.caption("Отметьте признаки, если они есть — по ним расширится план обследования.")
    flags = st.multiselect("Настораживающие признаки", [
        "Пароксизмы АД, потливость, тахикардия",
        "Храп, дневная сонливость",
        "Мышечная слабость, эпизоды гипокалиемии",
        "Регулярный приём НПВС/деконгестантов",
        "Систолический шум над почечными артериями",
        "АД, резистентное к 3 препаратам в максимальных дозах",
    ])

# ---------------- 6. Сопутствующие заболевания ----------------
with st.expander("6. Сопутствующие заболевания"):
    c1, c2, c3 = st.columns(3)
    dm = c1.checkbox("Сахарный диабет 2 типа", value=True)
    ibs = c1.checkbox("ИБС (ишемическая болезнь сердца)")
    stroke = c1.checkbox("Перенесённый инсульт/ТИА")
    chf = c1.checkbox("ХСН")
    chd = c1.checkbox("ХОБЛ")
    glj = c2.checkbox("ГЛЖ по ЭхоКГ")
    athero = c2.checkbox("Атеросклероз: атероматозные бляшки при визуализации (стеноз ≥ 50%)")
    anemia = c2.checkbox("Анемия")
    pad = c2.checkbox("Заболевание периферических артерий")
    fp_alone = c3.checkbox("Фибрилляция предсердий (вне ХСН)")
    ckd = c3.selectbox("ХБП", [
        "нет", "C3a (СКФ 45-59)", "C3b (СКФ 30-44)",
        "C4 (СКФ 15-29)", "C5 (СКФ <15)",
    ])
    others = st.text_input("Прочие сопутствующие (текстом)", value="")

# ---------------- 6а. Поражение органов-мишеней ----------------
with st.expander("6а. Поражение органов-мишеней (бессимптомное ПОМ)"):
    st.caption("Заполняется по данным ЭКГ/ЭхоКГ/анализов. Влияет на стадию АГ.")
    c1, c2, c3 = st.columns(3)
    ecg_glv = c1.checkbox("ЭКГ-признаки ГЛЖ (индекс Соколова-Лайона > 35 мм; R aVL ≥ 11 мм; "
                          "Корнельское произведение > 2440 мм×мс; SV3+RaVL > 28 мм (муж.) / > 20 мм (жен.))")
    ecg_q = c1.checkbox("Зубец Q или изменения ST-T на ЭКГ")
    lp_dil = c2.checkbox("Расширение левого предсердия (ИОЛП > 34 мл/м² при синусовом ритме, "
                         "> 40 мл/м² при ФП)")
    albuminuria = c2.checkbox("Альбуминурия 30-300 мг/24 ч (или А/К 30-300 мг/г, 3,4-34 мг/ммоль) "
                              "или протеинурия по тест-полоске")
    retinopathy = c3.checkbox("Выраженная ретинопатия")
    pwv = c1.selectbox("Каротидно-феморальная СЛПВ", ["нет данных", "≤ 10 м/с", "> 10 м/с"])
    lpi = c3.selectbox("Лодыжечно-плечевой индекс", ["нет данных", "≥ 0,9", "< 0,9"])
    st.caption("Пульсовое давление рассчитается после заполнения АД в разделе 7.")

# ---------------- 7. Объективный статус (включая АД) ----------------
with st.expander("7. Объективный статус", expanded=True):
    c1, c2 = st.columns(2)
    sys_r = c1.number_input("САД правая рука", min_value=80, max_value=250, value=165)
    dia_r = c1.number_input("ДАД правая рука", min_value=50, max_value=150, value=100)
    sys_l = c2.number_input("САД левая рука", min_value=80, max_value=250, value=160)
    dia_l = c2.number_input("ДАД левая рука", min_value=50, max_value=150, value=98)
    st.caption("АД измерено на обеих руках; в расчёт берётся максимальное значение.")

    consciousness = st.selectbox("Сознание", ["ясное", "спутанное", "оглушение", "сопор", "кома"])
    general = st.selectbox("Общее состояние", ["удовлетворительное", "средней тяжести", "тяжёлое"])
    c1, c2 = st.columns(2)
    skin = c1.selectbox("Кожные покровы", [
        "обычной окраски и влажности", "бледные", "акроцианоз", "холодные конечности"])
    edema = c1.selectbox("Отёки", [
        "нет", "пастозность голеней", "до нижней трети голеней",
        "до средней трети голеней", "до верхней трети голеней", "анасарка"])
    breath = c2.selectbox("Дыхание", ["везикулярное", "жёсткое", "ослабленное"])
    rales = c2.selectbox("Хрипы", ["нет", "сухие рассеянные", "влажные мелкопузырчатые в нижних отделах"])
    rr = c1.number_input("ЧДД в мин", min_value=10, max_value=40, value=16)
    spo2 = c2.number_input("SpO2, %", min_value=70, max_value=100, value=98)
    hr = st.number_input("ЧСС, уд/мин", min_value=35, max_value=200, value=76)
    tones = st.selectbox("Тоны сердца", [
        "ясные, ритмичные", "приглушены, ритмичные", "аритмичные (ФП)",
        "приглушены, ритм галопа"])
    murmur = st.text_input("Шумы", value="систолический шум на верхушке")
    c1, c2 = st.columns(2)
    jugular = c1.selectbox("Яремные вены", ["не набухшие", "набухшие"])
    hyr = c2.selectbox("Гепатоюгулярный рефлюкс", ["отрицательный", "положительный"])
    abdomen = st.selectbox("Живот", ["мягкий, безболезненный", "болезненный в эпигастрии", "увеличен, вздут"])
    liver = st.selectbox("Печень", ["у края рёберной дуги", "выступает на 1-2 см (гепатомегалия)", "значительно увеличена"])
    c1, c2 = st.columns(2)
    stool = c1.selectbox("Стул", ["оформленный, регулярный", "склонность к запорам", "диарея"])
    urination = c2.selectbox("Мочеиспускание", ["свободное, безболезненное", "олигурия", "никтурия", "дизурия"])
    punch = st.selectbox("Симптом Пастернацкого", [
        "отрицательный с обеих сторон", "положительный слева",
        "положительный справа", "положительный с обеих сторон"])
    pulse_deficit = st.number_input("Дефицит пульса при ФП (уд/мин; 0 если нет)", 0, 60, 0)

ad_s = int(max(sys_r, sys_l))
ad_d = int(max(dia_r, dia_l))
pulse_pressure = ad_s - ad_d
dp = round(ad_s * hr / 100)  # двойное произведение

# --- степень АГ (нужна ПТВ и разделам ниже) ---
if ad_s >= 180 or ad_d >= 110:
    degree = 3
elif ad_s >= 160 or ad_d >= 100:
    degree = 2
elif ad_s >= 140 or ad_d >= 90:
    degree = 1
else:
    degree = 0

# --- авторасчёт ФК и стадии ХСН ---
edema_list = ["нет", "пастозность голеней", "до нижней трети голеней",
              "до средней трети голеней", "до верхней трети голеней", "анасарка"]
edema_idx = edema_list.index(edema)
HF_COMPLAINTS = {
    "Одышка при нагрузке",
    "Одышка лёжа / пароксизмальная ночная одышка (ортопноэ)",
    "Прибавка веса более 2 кг за неделю",
    "Уменьшение количества мочи",
    "Отёки ног",
}
hf_present = HF_COMPLAINTS & set(complaints)
orthopnoe = "Одышка лёжа / пароксизмальная ночная одышка (ортопноэ)" in hf_present

if not hf_present and edema_idx == 0 and rales == "нет" and jugular == "не набухшие" \
   and liver == "у края рёберной дуги":
    chf_stage_auto = "предстадия"
elif edema_idx >= 2 or rales.startswith("влажные") or jugular == "набухшие" \
     or liver != "у края рёберной дуги":
    chf_stage_auto = "II"
else:
    chf_stage_auto = "I"

if skin == "акроцианоз" or spo2 <= 90:
    chf_fc_auto = "IV"
elif orthopnoe or edema_idx >= 3 or jugular == "набухшие" or liver != "у края рёберной дуги":
    chf_fc_auto = "III"
elif hf_present or edema_idx >= 1:
    chf_fc_auto = "II"
else:
    chf_fc_auto = "I"

# ---------------- 7а. ИБС — детали (типичность боли, ФК, ПТВ) ----------------
ibs_c1 = ibs_c2 = ibs_c3 = False
angina_form = ""
angina_fc = ""
ibs_typicality = ""
ptv_score = 0
ptv_text = ""
if ibs:
    with st.expander("7а. ИБС — детали (характер боли, ФК, ПТВ)", expanded=True):
        st.caption("Критерии типичной стенокардии (все три признака одновременно): "
                   "боль в области грудины с иррадиацией; возникает при нагрузке/стрессе/"
                   "холоде/обильной еде/кризе АД; проходит в покое или через 1-3 мин после нитроглицерина.")
        c1, c2, c3 = st.columns(3)
        ibs_c1 = c1.checkbox("Боль/дискомфорт за грудиной с иррадиацией (левая рука, челюсть, спина), 2-5 мин (не более 20)")
        ibs_c2 = c2.checkbox("Возникает при физической нагрузке / стрессе / холоде / обильной еде / кризе АД")
        ibs_c3 = c3.checkbox("Проходит в покое или через 1-3 мин после нитроглицерина")
        n_ang = sum([ibs_c1, ibs_c2, ibs_c3])
        if n_ang == 3:
            ibs_typicality = "типичная (несомненная) стенокардия"
        elif n_ang == 2:
            ibs_typicality = "атипичная стенокардия"
        else:
            ibs_typicality = "неангинозные боли"
        st.info(f"Характер боли: **{ibs_typicality}** ({n_ang} из 3 признаков).")

        c1, c2 = st.columns(2)
        angina_form = c1.selectbox("Форма ИБС", [
            "Стенокардия напряжения", "Стенокардия вазоспастическая",
            "Стенокардия микрососудистая", "Безболевая ишемия миокарда",
        ])
        if angina_form.startswith("Стенокардия"):
            angina_fc = c2.selectbox("ФК стенокардии (CCS)", ["I", "II", "III", "IV"],
                                     help="ФК I: приступы при необычной/очень большой нагрузке. "
                                          "ФК II: при обычной нагрузке (ходьба > 500 м, подъём > 1 этажа). "
                                          "ФК III: при лёгкой нагрузке (ходьба < 500 м, подъём на 1 этаж). "
                                          "ФК IV: в покое или при минимальной нагрузке.")

        st.markdown("**Предтестовая вероятность (ПТВ) ИБС**")
        n_fr_ptv = sum([smoking == "Курит сейчас", dyslipid, dm, degree >= 1, bmi >= 30])
        ptv_score = 0
        ptv_score += 2 if ibs_typicality == "типичная (несомненная) стенокардия" \
            else (1 if ibs_typicality == "атипичная стенокардия" else 0)
        if sex == "Мужской" and age >= 40:
            ptv_score += 1
        if age >= 50:
            ptv_score += 1
        if age >= 65:
            ptv_score += 1
        if n_fr_ptv >= 2:
            ptv_score += 1
        if ecg_q:
            ptv_score += 1
        if ptv_score >= 5:
            ptv_text = ">15% (умеренная)"
        elif ptv_score >= 2:
            ptv_text = "5-15% (низкая)"
        else:
            ptv_text = "<5% (очень низкая)"
        st.write(f"ПТВ ИБС: **{ptv_text}** ({ptv_score} б: типичность {n_ang}/3, возраст/пол, "
                 f"ФР ССЗ: {n_fr_ptv}, ЭКГ-изменения: {'да' if ecg_q else 'нет'}).")
        st.caption("Полосы ПТВ: < 5% — очень низкая, ИБС может быть исключена после первичного обследования; "
                   "5-15% — низкая, нужно первичное и специфическое обследование; "
                   ">15% — умеренная, при клинических факторах может потребоваться специфическое обследование. "
                   "Повышают ПТВ: зубец Q/ST-T, дисфункция ЛЖ, положительная нагрузочная ЭКГ, кальциноз КА. "
                   "Снижают: отрицательная нагрузочная ЭКГ, кальций КА = 0.")

# ---------------- 7б. Основной диагноз ----------------
main_dx = "Гипертоническая болезнь"
if ibs:
    with st.expander("7б. Основной диагноз"):
        auto = "ИБС" if (ibs and (ibs_c1 or ibs_c2)) else "ГБ"
        main_dx = st.radio(
            "Что выносить как основное заболевание",
            ["Автоматически (по жалобам)", "Гипертоническая болезнь", "ИБС"],
            help="Автоматически: ИБС — если она отмечена и есть ангинозный болевой синдром; "
                 "иначе ГБ. Основное заболевание — то, что определяет повод для визита.")
        if main_dx == "Автоматически (по жалобам)":
            main_dx = auto
        st.caption(f"Автопредложение: **{auto}**.")

# ---------------- 8. ХСН — детали + шкала H2FPEF ----------------
chf_type = chf_stage = chf_fc = ""
af = vent_arr = False
h2fpef_score = 0
h2fpef_items = []
h2fpef_concl = ""
ntprobnp = "нет данных"
CHF_TYPES = [
    "ФВ не известна",
    "ХСНнФВ (ФВ ≤ 40%)", "ХСНунФВ (ФВ 41-49%)", "ХСНсФВ (ФВ ≥ 50%)",
]
if chf:
    with st.expander("8. ХСН — детали (авторасчёт)", expanded=True):
        st.caption(f"Стадия и ФК подставлены по жалобам и объективному статусу: "
                   f"стадия «{chf_stage_auto}», ФК {chf_fc_auto}. При необходимости поправьте вручную.")
        c1, c2, c3 = st.columns(3)
        chf_type = c1.selectbox("Вид ХСН (по ФВ ЛЖ)", CHF_TYPES,
                                help="ХСНнФВ — ФВ ≤ 40%; ХСНунФВ — 41-49%; ХСНсФВ — ≥ 50%. "
                                     "Если ЭхоКГ не выполнена — «ФВ не известна».")
        stage_options = ["предстадия", "I", "II"]
        chf_stage = c2.selectbox("Стадия ХСН (РКО 2023)", stage_options,
                                 index=stage_options.index(chf_stage_auto))
        fc_options = ["I", "II", "III", "IV"]
        chf_fc = c3.selectbox("ФК по NYHA", fc_options,
                              index=fc_options.index(chf_fc_auto))
        c1, c2 = st.columns(2)
        af = c1.checkbox("Фибрилляция предсердий")
        vent_arr = c2.checkbox("Желудочковые нарушения ритма")

        st.markdown("**Шкала H2FPEF — вероятность ХСНсФВ**")
        st.caption("Показателей ЭхоКГ, которых нет, оставьте «нет данных» — балл за них не начисляется.")
        h1, h2, h3 = st.columns(3)
        antihtn_n = h1.number_input("Гипотензивных препаратов принимает", 0, 6, 0 if first_visit else 2,
                                    help="Балл H (Hypertensive): 1 б при приёме 2 и более препаратов")
        sdlg = h2.selectbox("СДЛА по ЭхоКГ", ["нет данных", "≤ 35 мм рт. ст.", "> 35 мм рт. ст."],
                            help="Балл P (Pulmonary): 1 б при СДЛА > 35 мм рт. ст.")
        ee = h3.selectbox("E/e' по ЭхоКГ", ["нет данных", "≤ 9", "> 9"],
                          help="Балл F (Filling): 1 б при E/e' > 9")
        n1, n2 = st.columns(2)
        ntprobnp = n1.selectbox("NT-proBNP/BNP",
                                ["нет данных", "в норме (< 125/35 пг/мл)",
                                 "повышен (> 125/35 пг/мл; > 365/105 при ФП)"],
                                help="Критерий 3 для ХСНсФВ: повышение НУП")

        h2fpef_score = 0
        h2fpef_items = []
        if bmi > 30:
            h2fpef_score += 2
            h2fpef_items.append("H (Heavy): ИМТ > 30 — 2 б")
        if antihtn_n >= 2:
            h2fpef_score += 1
            h2fpef_items.append("H (Hypertensive): ≥ 2 гипотензивных препаратов — 1 б")
        if af:
            h2fpef_score += 3
            h2fpef_items.append("F (Atrial fibrillation): ФП — 3 б")
        if sdlg == "> 35 мм рт. ст.":
            h2fpef_score += 1
            h2fpef_items.append("P (Pulmonary): СДЛА > 35 мм рт. ст. — 1 б")
        if age >= 60:
            h2fpef_score += 1
            h2fpef_items.append("E (Elder): возраст ≥ 60 лет — 1 б")
        if ee == "> 9":
            h2fpef_score += 1
            h2fpef_items.append("F (Filling): E/e' > 9 — 1 б")
        if h2fpef_score >= 6:
            h2fpef_concl = "ХСНсФВ высоковероятна"
        elif h2fpef_score >= 2:
            h2fpef_concl = "требуется дообследование"
        else:
            h2fpef_concl = "ХСНсФВ маловероятна"
        st.write(f"Сумма баллов H2FPEF: **{h2fpef_score} из 9** — {h2fpef_concl}.")

# ---------------- 9. Текущая терапия (повторный визит) ----------------
with st.expander("9. Текущая терапия (для повторного визита)"):
    current_therapy = st.text_input(
        "Что принимает сейчас",
        value="периндоприл 5 мг утром, амлодипин 5 мг утром" if not first_visit else "",
        disabled=first_visit,
    )

# ================= РАСЧЁТ =================

ckd_heavy = ckd in ("C4 (СКФ 15-29)", "C5 (СКФ <15)")
ckd_pom = ckd in ("C3a (СКФ 45-59)", "C3b (СКФ 30-44)")

chf_diagnosed = chf and (
    chf_type in ("ХСНнФВ (ФВ ≤ 40%)", "ХСНунФВ (ФВ 41-49%)")
    or (chf_type == "ХСНсФВ (ФВ ≥ 50%)" and h2fpef_score >= 6)
)

aks = ibs or stroke or athero or chf_diagnosed or ckd_heavy or pad \
      or (fp_alone or (chf and af)) or (dm and (glj or ecg_glv or albuminuria or ckd_pom
                                                or lpi == "< 0,9" or retinopathy
                                                or pwv == "> 10 м/с" or lp_dil))

pom_items = []
if glj: pom_items.append("ГЛЖ по ЭхоКГ")
if ecg_glv: pom_items.append("ЭКГ-признаки ГЛЖ")
if lp_dil: pom_items.append("расширение ЛП")
if albuminuria: pom_items.append("альбуминурия/протеинурия")
if ckd_pom: pom_items.append(f"ХБП {ckd}")
if lpi == "< 0,9": pom_items.append("ЛПИ < 0,9")
if retinopathy: pom_items.append("ретинопатия")
if pwv == "> 10 м/с": pom_items.append("СЛПВ > 10 м/с")
if pulse_pressure >= 60 and age >= 65: pom_items.append("пульсовое давление ≥ 60 мм рт. ст.")
pom = bool(pom_items)

stage = "III" if aks else ("II" if (pom or dm) else "I")

if degree == 0:
    risk = "—"
elif dm or ibs or stroke or athero or chf_diagnosed or ckd_heavy or pad \
     or fp_alone or (chf and af):
    risk = "4 (очень высокий)"
elif pom:
    risk = "3 (высокий)"
elif degree == 1:
    risk = "2 (средний)"
else:
    risk = "3 (высокий)"

if age < 65:
    target = "130/80 мм рт. ст."
    achieved = ad_s <= 130 and ad_d <= 80
else:
    target = "130-139/80 мм рт. ст."
    achieved = ad_s <= 139 and ad_d <= 80

comorb_list = []
if dm: comorb_list.append("Сахарный диабет 2 типа")
if stroke: comorb_list.append("Перенесённый инсульт/ТИА")
if chd: comorb_list.append("ХОБЛ")
if glj: comorb_list.append("ГЛЖ")
if athero: comorb_list.append("Атеросклероз (стеноз ≥ 50%)")
if anemia: comorb_list.append("Анемия")
if pad: comorb_list.append("Заболевание периферических артерий")
if fp_alone: comorb_list.append("Фибрилляция предсердий")
if ckd != "нет": comorb_list.append(f"ХБП {ckd}")
if others: comorb_list.append(others)

fr_list = []
if smoking == "Курит сейчас": fr_list.append("курение")
elif smoking == "Курил раньше": fr_list.append("курение в анамнезе")
if alcohol == "Регулярно": fr_list.append("регулярное употребление алкоголя")
if dyslipid: fr_list.append("дислипидемия")
if glucose: fr_list.append("нарушение гликемии натощак / толерантности к глюкозе")
if inactivity: fr_list.append("малоподвижный образ жизни")
if uric_acid: fr_list.append("мочевая кислота ≥ 360 мкмоль/л")
if heredity: fr_list.append("семейный анамнез ССЗ в молодом возрасте")
if psych: fr_list.append("психосоциальные факторы")
if early_menopause: fr_list.append("ранняя менопауза")
if bmi >= 40: fr_list.append(f"ожирение III степени (ИМТ {bmi:.1f})")
elif bmi >= 35: fr_list.append(f"ожирение II степени (ИМТ {bmi:.1f})")
elif bmi >= 30: fr_list.append(f"ожирение I степени (ИМТ {bmi:.1f})")
elif bmi >= 25: fr_list.append(f"избыточная масса тела (ИМТ {bmi:.1f})")
if (sex == "Мужской" and waist >= 102) or (sex == "Женский" and waist >= 88):
    fr_list.append("абдоминальное ожирение")
if (sex == "Мужской" and age >= 55) or (sex == "Женский" and age >= 65):
    fr_list.append("возраст ≥ 55 лет (муж.)" if sex == "Мужской" else "возраст ≥ 65 лет (жен.)")
if hr > 80: fr_list.append(f"ЧСС покоя > 80 уд/мин ({hr})")
if pregnant: fr_list.append("беременность")

# --- единый конструктор медикаментозной терапии ---
def build_med():
    meds = []
    def add(x):
        # не добавляем группу, если она уже есть (совпадает префикс до «:»)
        prefix = x.split(":")[0]
        if any(m.split(":")[0] == prefix for m in meds):
            return
        meds.append(x)

    if pregnant:
        add("Метилдопа 250-500 мг 2 раза в сутки или нифедипин ретард 20 мг 1-2 раза в сутки")
        return meds

    if degree and not chf:
        add("иРААС: периндоприл или лозартан")
        add("АК: амлодипин или фелодипин")
        if ckd_heavy:
            add("Петлевой диуретик: торасемид (тиазидоподобные при СКФ < 30 неэффективны)")
        else:
            add("Тиазидоподобный диуретик: индапамид")
        add("β-АБ: бисопролол или небиволол (при ИБС, тахикардии)")
        if not first_visit:
            add("При недостижении цели АД — усилить до тройной комбинации (иРААС + АК + диуретик)")

    if chf:
        if chf_type == "ХСНнФВ (ФВ ≤ 40%)":
            add("АРНИ: валсартан+сакубитрил (при недоступности/непереносимости — иАПФ)")
        elif chf_type == "ФВ не известна":
            add("иАПФ (при непереносимости — АРА); вид терапии уточнить после определения ФВ ЛЖ по ЭхоКГ")
        else:
            add("иАПФ (при непереносимости — АРА или валсартан+сакубитрил)")
        add("β-АБ: бисопролол, метопролол пролонгированный или небиволол (кардиоселективные)")
        if chd:
            add("Ивабрадин (альтернатива снижения ЧСС при ХОБЛ)")
        add("Антагонист альдостерона: спиронолактон или эплеренон")
        add("иНГЛТ2: дапаглифлозин или эмпаглифлозин")
        if edema_idx >= 1 or rales.startswith("влажные"):
            add("Петлевой диуретик: фуросемид или торасемид — для устранения застоя")
        elif ckd_heavy:
            add("Петлевой диуретик: торасемид (тиазидоподобные при СКФ < 30 неэффективны)")
        if af:
            add("Антикоагулянт: варфарин или НОАК (по CHA2DS2-VASc)")
            add("Амиодарон — контроль ритма и ЧСС")
        if vent_arr:
            add("Амиодарон; обсудить ИКД/СРТ и катетерную аблацию")
        if dm:
            add("Метформин; иПДПП-4 (ситаглиптин/линаглиптин) или аГПП-1")

    if dm and not chf:
        add("иНГЛТ2: дапаглифлозин или эмпаглифлозин")
        add("Метформин; иПДПП-4 (ситаглиптин/линаглиптин) или аГПП-1")

    if ibs:
        add("Антиагрегант: ацетилсалициловая кислота (при непереносимости — клопидогрел)")
        add("Статин высокой интенсивности (при недостижении цели ЛНП — добавить эзетемиб)")
        add("Нитраты короткого действия: нитроглицерин сублингвально при приступах")
        if angina_fc in ("III", "IV"):
            add("Нитраты длительного действия (изосорбида мононитрат/динитрат)")
            add("Ранолазин или триметазидин; альтернатива — никорандил")
        if not (degree or chf):
            add("β-АБ: бисопролол или небиволол")
            add("АК: амлодипин или фелодипин (если β-АБ недостаточно/не переносится)")
        if fp_alone or (chf and af):
            add("Антикоагулянт при ФП: НОАК или АВК (по CHA2DS2-VASc)")
    if pad:
        add("При заболевании ПА: статин, антиагрегант, цилостазол (при перемежающейся хромоте)")
    return meds

med_list = build_med()

base_plan = [
    "ОАК, ОАМ",
    "Глюкоза натощак, HbA1c",
    "Креатинин с расчётом СКФ (CKD-EPI), калий, натрий",
    "Липидограмма (ОХ, ЛПНП, ЛПВП, ТГ)",
    "АЛТ, АСТ, билирубин, общий белок, альбумин, мочевая кислота",
    "ТТГ",
    "ЭКГ в 12 отведениях" + (" (индекс Соколова-Лайона, Корнельское произведение)" if degree else ""),
    "ЭхоКГ" + (" (ФВ по Симпсону, диастолическая функция ЛЖ, ИММЛЖ, размер ЛП)" if chf else " (ГЛЖ, функция ЛЖ)"),
    "Осмотр глазного дна",
    "СМАД при неоднозначных значениях офисного АД",
]
flag_plan = {
    "Пароксизмы АД, потливость, тахикардия": "Метанефрины крови или суточной мочи (исключить феохромоцитому)",
    "Храп, дневная сонливость": "Скрининг СОАС, при подозрении — полисомнография",
    "Мышечная слабость, эпизоды гипокалиемии": "Ренин, альдостерон и их соотношение (исключить первичный гиперальдостеронизм)",
    "Регулярный приём НПВС/деконгестантов": "Отмена препарата с контролем АД через 2 недели",
    "Систолический шум над почечными артериями": "УЗДГ почечных артерий (исключить реноваскулярную АГ)",
    "АД, резистентное к 3 препаратам в максимальных дозах": "Резистентная АГ: суточная проба с 1/2 табл. спиронолактона, консультация эндокринолога/нефролога",
}
plan = list(base_plan) + [flag_plan[f] for f in flags if f in flag_plan]

ibs_plan = []
if ibs:
    ibs_plan = [
        "Нагрузочная ЭКГ (первичная оценка ишемии)",
        "Стресс-ЭхоКГ или сцинтиграфия миокарда с функциональными пробами (уточнить ишемию)",
        "КТ-скрининг коронарного кальция (индекс Агатстона; кальций = 0 снижает ПТВ)",
        "Дуплексное сканирование экстракраниальных отделов БЦА",
    ]
    if ptv_text == ">15% (умеренная)":
        ibs_plan.append("МСКТА коронарных артерий (ПТВ > 15%)")
    if angina_form == "Стенокардия вазоспастическая":
        ibs_plan.append("Холтер ЭКГ (подозрение на вазоспастическую стенокардию)")
    if angina_fc in ("III", "IV"):
        ibs_plan.append("КАГ — для стратификации риска ССО при тяжёлой стенокардии (ФК III-IV)")
    if chf:
        ibs_plan.append("МРТ сердца с контрастированием (оценка жизнеспособности и рубцовых изменений)")
plan += ibs_plan

if not albuminuria and degree:
    plan.append("Отношение альбумин-креатинин в моче (скрининг ПОМ)")
if pwv == "нет данных" and degree and age >= 65:
    plan.append("Каротидно-феморальная СЛПВ (скрининг артериальной жёсткости)")
if lpi == "нет данных" and (pad or degree >= 2):
    plan.append("Лодыжечно-плечевой индекс (скрининг ПА)")
if chf:
    chf_plan = [
        "NT-proBNP/BNP (>125/35 пг/мл при синусовом ритме, >365/105 пг/мл при ФП)",
        "ШОКС — шкала оценки клинического состояния",
        "Тест с 6-минутной ходьбой",
        "Прицельная рентгенография органов грудной клетки",
        "Ферритин, насыщение трансферрина железом (НТЖ)",
    ]
    if anemia:
        chf_plan.append("Развернутый ОАК с ретикулоцитами (анемия — коррекция дефицита железа)")
    if chf_type in ("ХСНсФВ (ФВ ≥ 50%)", "ФВ не известна"):
        if h2fpef_items:
            chf_plan.append(f"Шкала H2FPEF: {h2fpef_score} б ({'; '.join(h2fpef_items)}) — {h2fpef_concl}")
        else:
            chf_plan.append("Шкала H2FPEF: показатели не заполнены (заполнить в разделе 8)")
    if chf_type == "ФВ не известна":
        chf_plan.append("Определить ФВ ЛЖ по ЭхоКГ для установления вида ХСН")
    if af or vent_arr:
        chf_plan.append("Суточное мониторирование ЭКГ (холтер)")
    plan += chf_plan

if flags:
    st.warning("Отмечены признаки возможной вторичной АГ — в документ добавлен расширенный план обследования.")
if ibs and ibs_typicality == "неангинозные боли" and ptv_text == "<5% (очень низкая)":
    st.warning("ПТВ ИБС очень низкая, боли неангинозные: ИБС маловероятна — ищите другие причины боли "
               "(мышечно-скелетные, гастроэзофагеальные и др.).")
if chf and chf_type == "ХСНсФВ (ФВ ≥ 50%)" and h2fpef_score < 6:
    st.warning("Сумма H2FPEF < 6 б при выбранной ХСНсФВ: диагноз ХСН в документ не вносится, "
               "требуется дообследование.")
if chf and chf_type == "ФВ не известна":
    st.info("ФВ ЛЖ не известна: диагноз ХСН не выставляется до ЭхоКГ; в плане есть пункт определения ФВ ЛЖ.")
if chf and chf_type == "ХСНсФВ (ФВ ≥ 50%)" and h2fpef_score >= 6 \
   and not (ee == "> 9" or sdlg == "> 35 мм рт. ст." or glj or lp_dil
            or ntprobnp == "повышен (> 125/35 пг/мл; > 365/105 при ФП)"):
    st.warning("Критерий 3 для ХСНсФВ не выполнен: нет объективных признаков структурного/"
               "функционального поражения (E/e', СДЛА, ГЛЖ, расширение ЛП, NT-proBNP). "
               "Отметьте хотя бы один или вернитесь к ЭхоКГ.")

# --- сборка диагноза ---
def build_ibs_str():
    if not ibs:
        return ""
    s = f"ИБС: {angina_form}"
    if angina_fc:
        s += f", ФК {angina_fc}"
    s += f". ПТВ ИБС: {ptv_text}."
    return s

def build_dx():
    lines = []
    main_is_ibs = main_dx == "ИБС" and ibs
    ibs_str = build_ibs_str()
    if main_is_ibs:
        lines.append(ibs_str)
        if degree > 0:
            if first_visit:
                g = f"Гипертоническая болезнь {degree} степени"
                if fr_list:
                    g += ", ФР: " + ", ".join(fr_list)
                g += f", риск {risk}."
            else:
                g = f"Гипертоническая болезнь {stage} стадии"
                if fr_list:
                    g += ", ФР: " + ", ".join(fr_list)
                ach = "Целевое АД достигнуто" if achieved else "Целевое АД не достигнуто"
                g += f", риск {risk}. {ach} (цель: {target}; на визите {ad_s}/{ad_d} мм рт. ст.)."
            lines.append(g)
    else:
        if first_visit:
            if degree == 0:
                lines.append("Артериальная гипертония не выявлена (АД < 140/90 мм рт. ст.). Контроль АД рекомендован.")
            else:
                g = f"Гипертоническая болезнь {degree} степени"
                if fr_list:
                    g += ", ФР: " + ", ".join(fr_list)
                g += f", риск {risk}."
                lines.append(g)
        else:
            g = f"Гипертоническая болезнь {stage} стадии"
            if fr_list:
                g += ", ФР: " + ", ".join(fr_list)
            ach = "Целевое АД достигнуто" if achieved else "Целевое АД не достигнуто"
            g += f", риск {risk}. {ach} (цель: {target}; на визите {ad_s}/{ad_d} мм рт. ст.)."
            lines.append(g)
        if ibs_str:
            lines.append(ibs_str)
    return lines

def build_chf_dx():
    if not chf_diagnosed:
        return ""
    if chf_type == "ХСНсФВ (ФВ ≥ 50%)":
        dx2 = f"Осложнение: ХСНсФВ (ФВ ЛЖ ≥ 50%), {chf_stage} стадия, ФК {chf_fc}. "
        dx2 += f"Шкала H2FPEF: {h2fpef_score} б ({'; '.join(h2fpef_items)}) — ХСНсФВ высоковероятна."
    else:
        dx2 = f"Осложнение: {chf_type}, {chf_stage} стадия, ФК {chf_fc}."
    if af:
        dx2 += " ФП."
    if vent_arr:
        dx2 += " Желудочковые нарушения ритма."
    return dx2

# строка предварительного расчёта в интерфейсе
if first_visit:
    preview_ag = f"степень АГ {degree}" if degree else "АГ не выявлена"
else:
    preview_ag = "целевое АД достигнуто" if achieved else "целевое АД не достигнуто"
st.markdown("---")
st.markdown(
    f"**Предварительный расчёт:** основной — {main_dx if ibs else 'ГБ'} · {preview_ag} · стадия {stage} · риск {risk} · "
    f"цель АД {target} · фактическое {ad_s}/{ad_d}"
    + (f" · ИБС: {ibs_typicality}, ПТВ {ptv_text}" if ibs else "")
    + (f" · ХСН: стадия {chf_stage_auto}, ФК {chf_fc_auto}" if chf else "")
    + (f" · H2FPEF: {h2fpef_score} б — {h2fpef_concl}"
       if chf and chf_type in ("ХСНсФВ (ФВ ≥ 50%)", "ФВ не известна") else "")
)

# ---------------- Справка: классификация ----------------
with st.expander("Справка: классификация по КР156 / КР155"):
    st.markdown(
        "**По фракции выброса ЛЖ:** ХСНнФВ (≤ 40%), ХСНунФВ (41-49%), ХСНсФВ (≥ 50%). "
        "Для ХСНсФВ обязателен критерий 3 — объективные признаки структурного/функционального "
        "поражения (Э/e' > 9, СДЛА > 35 мм рт. ст., ГЛЖ, расширение ЛП, повышенный НУП)."
    )
    st.markdown(
        "**Стадии ХСН (РКО 2023):** предстадия — симптомов нет, есть структурное поражение/повышенный НУП; "
        "стадия I — симптомы есть или были; стадия II — тяжёлая ХСН с повторными госпитализациями "
        "или непереносимостью терапии."
    )
    st.markdown(
        "**Стадии АГ:** стадия I — нет ПОМ и АКС (возможны ФР); стадия II — бессимптомное ПОМ "
        "(артериальная жёсткость: пульсовое давление ≥ 60 у пожилых, СЛПВ > 10 м/с; ЭКГ/ЭхоКГ-ГЛЖ; "
        "альбуминурия 30-300 мг/24 ч; ХБП С3; ЛПИ < 0,9; ретинопатия) и/или СД без ПОМ; "
        "стадия III — АКС: ЦВЗ, ИБС, атеросклероз со стенозом ≥ 50%, ХСН, заболевание ПА, ФП, "
        "тяжёлая ХБП с СКФ < 30, СД с ПОМ."
    )
    st.markdown(
        "**Типичность стенокардии (КР155):** типичная — все 3 признака (боль за грудиной с иррадиацией, "
        "при нагрузке, купируется покоем/нитроглицерином); атипичная — 2 из 3; неангинозные боли — меньше 2. "
        "**ПТВ:** < 5% — очень низкая, 5-15% — низкая, > 15% — умеренная."
    )

# ================= ДОКУМЕНТ =================
def make_docx():
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)

    h = doc.add_paragraph()
    h.alignment = 1
    h.add_run("ОСМОТР ТЕРАПЕВТА-КАРДИОЛОГА").bold = True
    doc.add_paragraph(f"Дата: {date.today().strftime('%d.%m.%Y')}")

    doc.add_paragraph().add_run("Паспортная часть").bold = True
    sex_word = "мужской" if sex == "Мужской" else "женский"
    preg = ", беременность" if pregnant else ""
    doc.add_paragraph(f"{fio}, {age} лет, пол {sex_word}{preg}.")

    doc.add_paragraph().add_run("Жалобы").bold = True
    if complaints:
        doc.add_paragraph(", ".join(complaints).lower() + f". Длительность: {duration}.")
    else:
        doc.add_paragraph("Активных жалоб не предъявляет.")

    doc.add_paragraph().add_run("Анамнез заболевания").bold = True
    doc.add_paragraph(anamnesis)
    if not first_visit and current_therapy:
        doc.add_paragraph(f"Текущая гипотензивная терапия: {current_therapy}.")

    doc.add_paragraph().add_run("Объективный осмотр").bold = True
    obj_lines = [
        f"АД: правая рука {int(sys_r)}/{int(dia_r)}, левая рука {int(sys_l)}/{int(dia_l)} мм рт. ст. "
        f"(в расчёт взято максимальное {ad_s}/{ad_d}); пульсовое давление {pulse_pressure} мм рт. ст.",
        f"Общее состояние {general}, сознание {consciousness}.",
        f"Кожные покровы {skin}, отёки: {edema}.",
        f"Дыхание {breath}, хрипы: {rales}. ЧДД {rr} в мин, SpO2 {spo2}%.",
        f"ЧСС {hr} уд/мин"
        + (f", дефицит пульса {pulse_deficit} уд/мин" if pulse_deficit else "")
        + f", тоны сердца {tones}. Шумы: {murmur}. Двойное произведение: {dp}.",
        f"Яремные вены: {jugular}. Гепатоюгулярный рефлюкс {hyr}.",
        f"Живот {abdomen}. Печень {liver}. Стул: {stool}. Мочеиспускание: {urination}.",
        f"Симптом Пастернацкого {punch}.",
    ]
    for line in obj_lines:
        doc.add_paragraph(line)

    doc.add_paragraph().add_run("Диагноз").bold = True
    for line in build_dx():
        doc.add_paragraph(line)
    chf_dx = build_chf_dx()
    if chf_dx:
        doc.add_paragraph(chf_dx)
    if pom_items and not aks:
        doc.add_paragraph("Бессимптомное ПОМ: " + ", ".join(pom_items) + ".")
    if comorb_list:
        doc.add_paragraph("Сопутствующие/фоновые: " + ", ".join(comorb_list) + ".")

    doc.add_paragraph().add_run("Рекомендации").bold = True
    if med_list:
        doc.add_paragraph("Медикаментозная терапия:", style="List Bullet")
        for m in med_list:
            doc.add_paragraph(m, style="List Bullet 2")
    else:
        doc.add_paragraph("Медикаментозная терапия не показана: немедикаментозные меры "
                          "и контроль АД.", style="List Bullet")

    life = ["Ограничение соли до 5 г/сут."]
    if ibs:
        life.append("Регулярные аэробные упражнения (не менее 30 мин умеренной интенсивности 5-7 дней в неделю); "
                    "овощи, фрукты, рыба, орехи, ненасыщенные жиры (оливковое масло).")
    if chf:
        life.append("Ограничение жидкости при выраженном застое; ежедневное взвешивание "
                    "(прибавка > 2 кг за 3 дня — обратиться к врачу).")
    if bmi >= 25:
        life.append("Снижение массы тела (цель — ИМТ < 25).")
    if smoking == "Курит сейчас":
        life.append("Отказ от курения.")
    if alcohol == "Регулярно":
        life.append("Ограничение алкоголя.")
    if inactivity:
        life.append("Расширение физической активности до 150 мин/нед аэробной нагрузки.")
    life.append("Домашний контроль АД с дневником (утро и вечер, 2 измерения).")
    if not first_visit:
        life.append(f"Контроль эффективности терапии: повторный визит через 4 недели (цель АД {target}).")
    for item in life:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph().add_run("План обследования").bold = True
    for item in plan:
        doc.add_paragraph(item, style="List Bullet")

    if flags:
        doc.add_paragraph().add_run("Исключить вторичную АГ").bold = True
        for f in flags:
            doc.add_paragraph(f"{f} — {flag_plan.get(f, 'дообследование по показаниям')}", style="List Bullet")

    p = doc.add_paragraph()
    p.add_run("Прототип сформирован автоматически; перед клиническим использованием сверить "
              "с клиническими рекомендациями и подписать у врача.").italic = True
    return doc

# ---------------- 10. Документ осмотра ----------------
st.subheader("10. Документ осмотра")
if st.button("Сформировать осмотр", type="primary"):
    buf = io.BytesIO()
    make_docx().save(buf)
    st.download_button(
        "Скачать осмотр (docx)",
        data=buf.getvalue(),
        file_name=f"osmotr_{date.today().isoformat()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
