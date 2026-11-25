import io
from datetime import date

import streamlit as st
from docx import Document
from docx.shared import Pt

st.set_page_config(page_title="Кардио-осмотр: АГ+ХСН", page_icon="🩺", layout="wide")
st.title("Первичный осмотр кардиолога: АГ + ХСН")
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

# ---------------- 2. Артериальное давление ----------------
with st.expander("2. Артериальное давление", expanded=True):
    st.caption("АД измерено на обеих руках; в расчёт берётся максимальное значение.")
    c1, c2 = st.columns(2)
    sys_r = c1.number_input("САД правая рука", min_value=80, max_value=250, value=165)
    dia_r = c1.number_input("ДАД правая рука", min_value=50, max_value=150, value=100)
    sys_l = c2.number_input("САД левая рука", min_value=80, max_value=250, value=160)
    dia_l = c2.number_input("ДАД левая рука", min_value=50, max_value=150, value=98)
    hr = st.number_input("ЧСС, уд/мин", min_value=35, max_value=200, value=76)

ad_s = int(max(sys_r, sys_l))
ad_d = int(max(dia_r, dia_l))

# ---------------- 3. Антропометрия ----------------
with st.expander("3. Антропометрия"):
    c1, c2, c3 = st.columns(3)
    height = c1.number_input("Рост, см", min_value=120, max_value=220, value=175)
    weight = c2.number_input("Вес, кг", min_value=30, max_value=250, value=96)
    waist = c3.number_input("Окружность талии, см", min_value=40, max_value=180, value=104)

bmi = weight / (height / 100) ** 2

# ---------------- 4. Жалобы и анамнез ----------------
with st.expander("4. Жалобы и анамнез"):
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
    heredity = st.checkbox("Наследственность: ранние ССЗ у родственников (<55 лет у мужчин, <65 у женщин)")

# ---------------- 5. Факторы риска ----------------
with st.expander("5. Факторы риска"):
    c1, c2 = st.columns(2)
    smoking = c1.selectbox("Курение", ["Никогда не курил", "Курит сейчас", "Курил раньше"])
    alcohol = c1.selectbox("Алкоголь", ["Не употребляет", "Редко", "Регулярно"])
    dyslipid = c2.checkbox("Дислипидемия (ЛПНП ≥ 3,0 ммоль/л или терапия статинами)")
    glucose = c2.checkbox("Нарушение углеводного обмена (глюкоза натощак 5,6-6,9 или HbA1c 5,7-6,4%)")
    inactivity = c2.checkbox("Низкая физическая активность (<150 мин/нед)")

# ---------------- 6. Скрининг вторичной АГ ----------------
with st.expander("6. Скрининг вторичной АГ"):
    st.caption("Отметьте признаки, если они есть — по ним расширится план обследования.")
    flags = st.multiselect("Настораживающие признаки", [
        "Пароксизмы АД, потливость, тахикардия",
        "Храп, дневная сонливость",
        "Мышечная слабость, эпизоды гипокалиемии",
        "Регулярный приём НПВС/деконгестантов",
        "Систолический шум над почечными артериями",
        "АД, резистентное к 3 препаратам в максимальных дозах",
    ])

# ---------------- 7. Сопутствующие заболевания ----------------
with st.expander("7. Сопутствующие заболевания"):
    c1, c2, c3 = st.columns(3)
    dm = c1.checkbox("Сахарный диабет 2 типа", value=True)
    ibs = c1.checkbox("ИБС")
    stroke = c1.checkbox("Перенесённый инсульт/ТИА")
    chf = c1.checkbox("ХСН")
    chd = c1.checkbox("ХОБЛ")
    glj = c2.checkbox("ГЛЖ (ЭхоКГ/ЭКГ)")
    athero = c2.checkbox("Атеросклероз (СА/ПВА)")
    anemia = c2.checkbox("Анемия")
    ckd = c3.selectbox("ХБП", [
        "нет", "C3a (СКФ 45-59)", "C3b (СКФ 30-44)",
        "C4 (СКФ 15-29)", "C5 (СКФ <15)",
    ])
    others = st.text_input("Прочие сопутствующие (текстом)", value="")

# ---------------- 8. Объективный статус ----------------
with st.expander("8. Объективный статус"):
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

# --- авторасчёт ФК и стадии ХСН по жалобам и объективному статусу ---
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

# ---------------- 9. ХСН — детали + шкала H2FPEF ----------------
chf_type = chf_stage = chf_fc = ""
af = vent_arr = False
h2fpef_score = 0
h2fpef_items = []
h2fpef_concl = ""
CHF_TYPES = [
    "ФВ не известна",
    "ХСНнФВ (ФВ ≤ 40%)", "ХСНунФВ (ФВ 41-49%)", "ХСНсФВ (ФВ ≥ 50%)",
]
if chf:
    with st.expander("9. ХСН — детали (авторасчёт)", expanded=True):
        st.caption(f"Стадия и ФК подставлены по жалобам и объективному статусу: "
                   f"стадия «{chf_stage_auto}», ФК {chf_fc_auto}. При необходимости поправьте вручную.")
        c1, c2, c3 = st.columns(3)
        chf_type = c1.selectbox("Вид ХСН", CHF_TYPES,
                                help="Если ЭхоКГ ещё не выполнена — оставьте «ФВ не известна».")
        stage_options = ["предстадия", "I", "II"]
        chf_stage = c2.selectbox("Стадия ХСН", stage_options,
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

# ---------------- 10. Текущая терапия (повторный визит) ----------------
with st.expander("10. Текущая терапия (для повторного визита)"):
    current_therapy = st.text_input(
        "Что принимает сейчас",
        value="периндоприл 5 мг утром, амлодипин 5 мг утром" if not first_visit else "",
        disabled=first_visit,
    )

# ================= РАСЧЁТ =================
if ad_s >= 180 or ad_d >= 110:
    degree = 3
elif ad_s >= 160 or ad_d >= 100:
    degree = 2
elif ad_s >= 140 or ad_d >= 90:
    degree = 1
else:
    degree = 0

ckd_heavy = ckd in ("C4 (СКФ 15-29)", "C5 (СКФ <15)")
ckd_pom = ckd in ("C3a (СКФ 45-59)", "C3b (СКФ 30-44)")
ckz = dm or ibs or stroke or chf or ckd_heavy
pom = glj or athero or ckd_pom

stage = "III" if ckz else ("II" if pom else "I")

if degree == 0:
    risk = "—"
elif ckz:
    risk = "4 (очень высокий)"
elif pom:
    risk = "3 (высокий)"
elif degree == 1:
    risk = "2 (средний)"
else:
    risk = "3 (высокий)"

if age < 65:
    target = "130/80 мм рт. ст. (если переносится)"
    achieved = ad_s <= 130 and ad_d <= 80
else:
    target = "130-139/80 мм рт. ст."
    achieved = ad_s <= 139 and ad_d <= 80

comorb_list = []
if dm: comorb_list.append("сахарный диабет 2 типа")
if ibs: comorb_list.append("ИБС")
if stroke: comorb_list.append("перенесённый инсульт/ТИА")
if chd: comorb_list.append("ХОБЛ")
if glj: comorb_list.append("ГЛЖ")
if athero: comorb_list.append("атеросклероз аорты/ПВА")
if anemia: comorb_list.append("анемия")
if ckd != "нет": comorb_list.append(f"ХБП {ckd}")
if others: comorb_list.append(others)

fr_list = []
if smoking == "Курит сейчас": fr_list.append("курение")
elif smoking == "Курил раньше": fr_list.append("курение в анамнезе")
if alcohol == "Регулярно": fr_list.append("регулярное употребление алкоголя")
if dyslipid: fr_list.append("дислипидемия")
if glucose: fr_list.append("нарушение углеводного обмена")
if inactivity: fr_list.append("низкая физическая активность")
if heredity: fr_list.append("отягощённая наследственность по ранним ССЗ")
if bmi >= 30: fr_list.append(f"ожирение (ИМТ {bmi:.1f})")
elif bmi >= 25: fr_list.append(f"избыточная масса тела (ИМТ {bmi:.1f})")
if (sex == "Мужской" and waist >= 102) or (sex == "Женский" and waist >= 88):
    fr_list.append("абдоминальное ожирение")
if age >= 65: fr_list.append("возраст ≥ 65 лет")
if pregnant: fr_list.append("беременность")

# --- гипотензивная терапия ---
if pregnant:
    therapy = ("Метилдопа 250-500 мг 2 раза в сутки или нифедипин ретард 20 мг 1-2 раза в сутки. "
               "иРААС/АРА и тиазидные диуретики при беременности противопоказаны.")
elif degree == 0:
    therapy = ("Гипотензивная терапия не показана: немедикаментозные меры и домашний контроль АД, "
               "повторный визит через 3 месяца.")
elif chf:
    therapy = "Терапия ХСН (см. ниже) одновременно контролирует АД."
elif ckd_heavy:
    therapy = ("иРААС (периндоприл/лозартан) с осторожным титрованием + петлевой диуретик (торасемид). "
               "Тиазидоподобные диуретики при СКФ <30 неэффективны. "
               "Контроль креатинина и калия через 7-14 дней после старта/титрования. Целевое АД: " + target + ".")
else:
    therapy = ("Старт с фиксированной комбинации иРААС (периндоприл/лозартан) + АК (амлодипин) или "
               "тиазидоподобного диуретика (индапамид). При недостижении цели АД — тройная комбинация "
               "(иРААС + АК + диуретик). Целевое АД: " + target + ".")

# --- терапия ХСН (только если ХСН отмечена) ---
chf_therapy = ""
if chf:
    parts = []
    if chf_type == "ФВ не известна":
        parts.append("иАПФ (при непереносимости — АРА)")
        if h2fpef_items:
            parts.append(f"по шкале H2FPEF {h2fpef_score} б — {h2fpef_concl}; "
                         "вид терапии уточнить после определения ФВ ЛЖ по ЭхоКГ")
        else:
            parts.append("вид терапии уточнить после определения ФВ ЛЖ по ЭхоКГ")
    elif chf_type == "ХСНнФВ (ФВ ≤ 40%)":
        parts.append("валсартан+сакубитрил (или иАПФ при недоступности/непереносимости)")
    else:
        parts.append("иАПФ (при непереносимости — АРА или валсартан+сакубитрил)")
    if chd:
        parts.append("кардиоселективный β-АБ: бисопролол, метопролол пролонгированный или небиволол; ивабрадин")
    else:
        parts.append("β-АБ: бисопролол, метопролол пролонгированный или небиволол")
    parts.append("антагонист альдостерона (спиронолактон/эплеренон)")
    parts.append("ингибитор НГЛТ2 (дапаглифлозин/эмпаглифлозин)")
    parts.append("диуретик для устранения застоя (при отёках/застое)")
    if af:
        parts.append("при ФП: антикоагулянт (варфарин или НОАК по CHA2DS2-VASc), амиодарон; "
                     "антиаритмики I класса противопоказаны при ХСН")
    if vent_arr:
        parts.append("при желудочковых НРС: амиодарон; обсудить ИКД/СРТ, катетерную аблацию")
    if dm:
        parts.append("при СД 2 типа: дапаглифлозин/эмпаглифлозин, метформин, иПДПП-4 (ситаглиптин/линаглиптин), "
                     "аГПП-1")
    if ibs:
        parts.append("при ИБС: статины, антиагреганты, нитраты/никорандил")
    if ckd_heavy:
        parts.append("при СКФ <30: петлевой диуретик вместо тиазидоподобного, контроль калия и креатинина")
    chf_therapy = ("; ".join(parts) + ". "
                   "Противопоказаны при ХСН II-IV ФК: верапамил, дилтиазем, моксонидин, альфа-адреноблокаторы; "
                   "из АК только дигидропиридиновые (амлодипин, фелодипин).")

base_plan = [
    "ОАК, ОАМ",
    "Глюкоза натощак, HbA1c",
    "Креатинин с расчётом СКФ (CKD-EPI), калий, натрий",
    "Липидограмма (ОХ, ЛПНП, ЛПВП, ТГ)",
    "АЛТ, АСТ, билирубин, общий белок, альбумин, мочевая кислота",
    "ТТГ",
    "ЭКГ в 12 отведениях",
    "ЭхоКГ" + (" (ФВ по Симпсону, диастолическая функция ЛЖ)" if chf else " (ГЛЖ, функция ЛЖ)"),
    "Осмотр глазного дна",
    "АКМЕ (СМАД) при неоднозначных значениях офисного АД",
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

chf_plan = []
if chf:
    chf_plan += [
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
            detail = "; ".join(h2fpef_items)
            chf_plan.append(f"Шкала H2FPEF: {h2fpef_score} б ({detail}) — {h2fpef_concl}")
        else:
            chf_plan.append("Шкала H2FPEF: показатели не заполнены (заполнить в разделе 9)")
    if af or vent_arr:
        chf_plan.append("Суточное мониторирование ЭКГ (холтер)")

plan += chf_plan

if flags:
    st.warning("Отмечены признаки возможной вторичной АГ — в документ добавлен расширенный план обследования.")
if chf:
    st.info("ХСН отмечена: в документ добавлены терапия и обследование по КР156.")

# --- сборка диагноза ---
def build_dx():
    if first_visit:
        if degree == 0:
            return "Артериальная гипертония не выявлена (АД < 140/90 мм рт. ст.). Контроль АД рекомендован."
        dx = f"Гипертоническая болезнь {degree} степени"
        if fr_list:
            dx += ", ФР: " + ", ".join(fr_list)
        dx += f", риск {risk}."
        return dx
    dx = f"Гипертоническая болезнь {stage} стадии"
    if fr_list:
        dx += ", ФР: " + ", ".join(fr_list)
    ach = "целевое АД достигнуто" if achieved else "целевое АД не достигнуто"
    dx += f", риск {risk}. {ach.capitalize()} (цель: {target}; на визите {ad_s}/{ad_d} мм рт. ст.)."
    return dx

def build_chf_dx():
    if not chf:
        return ""
    tmap = {
        "ХСНнФВ (ФВ ≤ 40%)": "ХСНнФВ (ФВ ЛЖ ≤ 40%)",
        "ХСНунФВ (ФВ 41-49%)": "ХСНунФВ (ФВ ЛЖ 41-49%)",
        "ХСНсФВ (ФВ ≥ 50%)": "ХСНсФВ (ФВ ЛЖ ≥ 50%)",
        "ФВ не известна": "ХСН",
    }
    dx2 = "Осложнение: " + tmap.get(chf_type, "ХСН")
    dx2 += f", {chf_stage} стадия, ФК {chf_fc}."
    if chf_type == "ФВ не известна" and h2fpef_score >= 6:
        dx2 += " По шкале H2FPEF наиболее вероятна ХСНсФВ — подтвердить ЭхоКГ."
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
    f"**Предварительный расчёт:** {preview_ag} · риск {risk} · "
    f"цель АД {target} · фактическое {ad_s}/{ad_d}"
    + (f" · ХСН: стадия {chf_stage_auto}, ФК {chf_fc_auto}" if chf else "")
    + (f" · H2FPEF: {h2fpef_score} б — {h2fpef_concl}"
       if chf and chf_type in ("ХСНсФВ (ФВ ≥ 50%)", "ФВ не известна") else "")
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
        f"Общее состояние {general}, сознание {consciousness}.",
        f"Кожные покровы {skin}, отёки: {edema}.",
        f"Дыхание {breath}, хрипы: {rales}. ЧДД {rr} в мин, SpO2 {spo2}%.",
        f"ЧСС {hr} уд/мин"
        + (f", дефицит пульса {pulse_deficit} уд/мин" if pulse_deficit else "")
        + f", тоны сердца {tones}. Шумы: {murmur}.",
        f"АД: правая рука {int(sys_r)}/{int(dia_r)}, левая рука {int(sys_l)}/{int(dia_l)} мм рт. ст. "
        f"(в расчёт взято максимальное {ad_s}/{ad_d}).",
        f"Яремные вены: {jugular}. Гепатоюгулярный рефлюкс {hyr}.",
        f"Живот {abdomen}. Печень {liver}. Стул: {stool}. Мочеиспускание: {urination}.",
        f"Симптом Пастернацкого {punch}.",
    ]
    for line in obj_lines:
        doc.add_paragraph(line)

    doc.add_paragraph().add_run("Диагноз").bold = True
    doc.add_paragraph(build_dx())
    chf_dx = build_chf_dx()
    if chf_dx:
        doc.add_paragraph(chf_dx)
    if comorb_list:
        doc.add_paragraph("Сопутствующие/фоновые: " + ", ".join(comorb_list) + ".")

    doc.add_paragraph().add_run("Рекомендации").bold = True
    doc.add_paragraph("Медикаментозно: " + therapy)
    if chf_therapy:
        doc.add_paragraph("Терапия ХСН: " + chf_therapy)
    life = ["Ограничение соли до 5 г/сут."]
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
    for line in life:
        doc.add_paragraph(line)

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

# ---------------- 11. Документ осмотра ----------------
st.subheader("11. Документ осмотра")
if st.button("Сформировать осмотр", type="primary"):
    buf = io.BytesIO()
    make_docx().save(buf)
    st.download_button(
        "Скачать осмотр (docx)",
        data=buf.getvalue(),
        file_name=f"osmotr_{date.today().isoformat()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
