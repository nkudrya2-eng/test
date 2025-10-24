import re
from typing import Dict, List, Any, Tuple
from collections import defaultdict

from ferp_reference import FERP_TASKS, EQUIPMENT_FERP_MAP, PRIORITY_BY_TYPE

# ========== УСЛОВНЫЕ ФИЛЬТРЫ (детекция по тексту DOCX) ==========

# ========== PNR: генерация из фактических FERP ==========
from typing import Set

# Приоритет кодов для "main_tasks_5" (чтобы важное попадало первее)
FERP_PRIORITY = [
    "01-11-010-01",  # растекание заземляющего устройства
    "01-11-011-01",  # цепь заземления
    "01-11-012-01",  # напряжение/частота
    "01-11-026-01",  # фазировка/полярность
    "01-11-028-01",  # изоляция кабельных линий
    "01-11-029-01",  # изоляция цепей заземления
    "02-01-002-01",  # линии связи
    "02-01-004-01",  # оптическая мощность
    "01-06-022-01",  # СКС
    "02-01-003-01",  # сигнализация/ТМ
    "01-11-034-01",  # работа ИБП
    "01-11-011-02",  # экраны/заземление трасс
    "01-11-036-01",  # комплексное опробование (итог)
]

# Шаблоны текстов для чек-листа
CHECKLIST_TEXTS: Dict[str, List[str]] = {
    "01-11-010-01": [
        "Проверить целостность контура заземления и доступность точек измерения.",
        "Измерить сопротивление растеканию тока; зафиксировать результаты и сравнить с нормами.",
    ],
    "01-11-011-01": [
        "Проверить непрерывность защитной цепи между корпусами оборудования и ГЗШ.",
        "Проверить соединения PE/шины заземления; затяжку и маркировку.",
    ],
    "01-11-011-02": [
        "Проверить целостность экранов кабельных линий и их корректное заземление.",
        "Убедиться в отсутствии паразитных петель и многократных точек заземления, если это не предусмотрено.",
    ],
    "01-11-012-01": [
        "Проверить напряжение питающих цепей на соответствие допустимым отклонениям.",
        "Проверить частоту сети и фазировку вводов согласно схеме.",
    ],
    "01-11-026-01": [
        "Проверить правильность подключения L/N/PE и чередование фаз (при трёхфазных цепях).",
        "Проверить отсутствие переполюсовки в розеточных цепях.",
    ],
    "01-11-028-01": [
        "Измерить сопротивление изоляции кабелей напряжением до 1 кВ.",
        "Сравнить значения с требованиями; оформить протокол измерений.",
    ],
    "01-11-029-01": [
        "Проверить электрическую изоляцию цепей заземления от токоведущих частей.",
        "Подтвердить отсутствие утечек на корпус при нормальных условиях.",
    ],
    "02-01-002-01": [
        "Проверить целостность и параметры линий связи (длина, затухание/NEXT — по доступной методике).",
        "Выполнить прозвонку и контроль распиновки/полярности у медных линий.",
    ],
    "02-01-004-01": [
        "Измерить уровни оптической мощности на оконечном оборудовании (Tx/Rx).",
        "Сверить бюджет линии с паспортным; подтвердить отсутствие перегруза приёмника.",
    ],
    "01-06-022-01": [
        "Проверить функционирование портов СКС, соответствие категории канала.",
        "Выполнить тест-сертификацию (если предусмотрено) и сохранить отчёт.",
    ],
    "02-01-003-01": [
        "Проверить прохождение сигналов/телемеханики и корректность дискретных уровней.",
        "Подтвердить корректность адресации/канализации согласно проекту.",
    ],
    "01-11-034-01": [
        "Проверить работу ИБП от сети и от батарей, переход режимов без потерь нагрузки.",
        "Измерить время автономной работы под расчётной нагрузкой.",
    ],
    "01-11-036-01": [
        "Выполнить комплексное опробование системы в штатных режимах работы.",
        "Подтвердить готовность к приёмо-сдаточным испытаниям, оформить акт.",
    ],
}

def _collect_ferp_codes(equipment: List[Dict[str, Any]]) -> List[str]:
    codes: Set[str] = set()
    for it in equipment:
        for f in it.get("ferp", []) or []:
            code = f.get("code")
            if code:
                codes.add(code)
    # Сортируем по приоритету, затем по коду
    order = {c:i for i, c in enumerate(FERP_PRIORITY)}
    return sorted(codes, key=lambda c: (order.get(c, 999), c))

def _main_tasks_from_codes(sorted_codes: List[str], ferp_tasks: Dict[str, str]) -> List[str]:
    tasks: List[str] = []
    for c in sorted_codes:
        title = ferp_tasks.get(c)
        if not title:
            continue
        # Формируем краткую формулировку: "Код — Наименование"
        tasks.append(f"{c} — {title}")
        if len(tasks) >= 5:
            break
    return tasks

def _checklist_from_codes(sorted_codes: List[str]) -> List[str]:
    checklist: List[str] = []
    for c in sorted_codes:
        for line in CHECKLIST_TEXTS.get(c, []):
            checklist.append(line)
            if len(checklist) >= 15:  # ограничиваем верхнюю границу
                return checklist
    # Гарантируем минимум 10, при необходимости дублируем ключевые пункты итогового опробования
    while len(checklist) < 10:
        checklist.append("Подтвердить соответствие фактических параметров проектным и нормативным требованиям.")
        if len(checklist) >= 15:
            break
    return checklist

def _methodology_from_codes(sorted_codes: List[str], ferp_tasks: Dict[str, str]) -> str:
    # Краткая методика с перечислением кодов и общими принципами
    parts: List[str] = []
    names = [f"{c} ({ferp_tasks.get(c, '')})" for c in sorted_codes]
    if names:
        parts.append("Перечень выполняемых испытаний и проверок по ФЕРп: " + "; ".join(names) + ".")
    parts.append(
        "Измерения выполняются аттестованными средствами, поверенными в установленном порядке; "
        "условия проведения соответствуют паспортам приборов и нормативной документации. "
        "Результаты оформляются протоколами с указанием модели прибора, его номера и даты поверки."
    )
    if "02-01-004-01" in sorted_codes:
        parts.append("Оптические измерения: уровни Tx/Rx контролируются на оконечном оборудовании; "
                     "бюджет линии сверяется с расчётным, учитываются соединители/сплайсы.")
    if "01-11-028-01" in sorted_codes:
        parts.append("Сопротивление изоляции: измеряется мегомметром при нормируемом напряжении испытания; "
                     "значения не ниже требуемых по проекту/нормам.")
    if "01-11-012-01" in sorted_codes or "01-11-026-01" in sorted_codes:
        parts.append("Электрические проверки: напряжение, частота и фазировка оцениваются на вводах/розетках; "
                     "соответствие допускам подтверждается измерениями.")
    if "01-11-036-01" in sorted_codes:
        parts.append("Комплексное опробование: система проверяется в штатных режимах с имитацией рабочих сценариев; "
                     "фиксируется готовность к приёмо-сдаточным испытаниям.")
    return " ".join(parts)

def build_pnr_from_ferp(equipment: List[Dict[str, Any]], ferp_tasks: Dict[str, str]) -> Dict[str, Any]:
    sorted_codes = _collect_ferp_codes(equipment)
    main_tasks = _main_tasks_from_codes(sorted_codes, ferp_tasks)
    checklist = _checklist_from_codes(sorted_codes)
    methodology = _methodology_from_codes(sorted_codes, ferp_tasks)
    return {
        "main_tasks_5": main_tasks,
        "checklist_10_15": checklist,
        "methodology_section": methodology,
    }



def has_optics(text: str) -> bool:
    """Есть ли оптика/SFP/волокно: ключевые маркеры."""
    patterns = [
        r"\bSFP\+?\b", r"\bQSFP\+?\b", r"\bоптик[ао]\b", r"\bоптическ",
        r"\bSC\/APC\b", r"\bLC\/UPC\b", r"\bG\.65[24]\b", r"\bDWDM\b", r"\bCWDM\b"
    ]
    return any(re.search(pat, text, re.IGNORECASE) for pat in patterns)

def has_scs(text: str) -> bool:
    """Есть ли упоминание СКС/патч-панелей/категорий кабеля."""
    patterns = [
        r"\bСКС\b", r"\bструктурированн[аяо]\b", r"\bCat\s?5e\b", r"\bCat\s?6(a)?\b",
        r"\bпатч-?панел", r"\bпатчкорд", r"\bRJ-?45\b"
    ]
    return any(re.search(pat, text, re.IGNORECASE) for pat in patterns)

def has_ac_power(text: str) -> bool:
    """Наличие упоминаний AC-питания, фазировки, частоты."""
    patterns = [
        r"\b220 ?В\b", r"\b230 ?В\b", r"\b50 ?Гц\b", r"\bфазировк", r"\bL[1-3]\b", r"\bN\b", r"\bPE\b"
    ]
    return any(re.search(pat, text, re.IGNORECASE) for pat in patterns)

def mentions_ups(text: str) -> bool:
    return bool(re.search(r"\bИБП\b|\bUPS\b", text, re.IGNORECASE))

def mentions_voip(text: str) -> bool:
    return bool(re.search(r"\bVoIP\b|\bIP-?телефон", text, re.IGNORECASE))

# ========== ВСПОМОГАТЕЛЬНОЕ ==========

def extract_port_count(fragment: str) -> int:
    """
    Парсит количество портов из фрагмента: "24-портовый", "48xGE", "24 GE", "24x1G", "8-port"
    Возвращает 0, если не найдено.
    """
    patterns = [
        r"(\d+)\s*[- ]?порт\w*",       # 24-портовый
        r"(\d+)\s*x\s*(GE|1G|10G|100M)", # 48xGE
        r"(\d+)\s*(GE|1G|10G|100M)",     # 24 GE
        r"(\d+)\s*port"                  # 8-port
    ]
    for pat in patterns:
        m = re.search(pat, fragment, re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                continue
    return 0

def summarize_work_from_ferp(ferp_list: List[Dict[str, str]]) -> str:
    titles = [item["title"] for item in ferp_list]
    return ", ".join(titles)

def unique_equipment_key(item: Dict[str, Any]) -> Tuple[str, str, str]:
    return (
        item.get("equipment_type", "").lower(),
        (item.get("model_or_code") or "").strip().lower(),
        (item.get("name") or "").strip().lower(),
    )

# ========== ЯДРО СБОРКИ FERP ДЛЯ ЭЛЕМЕНТА ==========

def select_ferp_for_item(equipment_type: str, doc_text: str) -> List[Dict[str, str]]:
    """
    Берём максимум из EQUIPMENT_FERP_MAP по типу,
    фильтруем по уместности: оптика/СКС/AC и т.д.
    """
    codes = EQUIPMENT_FERP_MAP.get(equipment_type, [])
    filtered: List[str] = []

    optics = has_optics(doc_text)
    scs = has_scs(doc_text)
    ac = has_ac_power(doc_text)

    for code in codes:
        # Фильтры по смыслу:
        if code == "02-01-004-01" and not optics:
            continue  # уровни оптической мощности — только если есть оптика
        if code == "01-06-022-01" and not scs:
            continue  # проверка СКС — только если СКС упомянута
        if code in {"01-11-012-01", "01-11-026-01"} and not ac:
            # напряжение/частота и фазировка — только при признаках AC-питания
            continue
        filtered.append(code)

    # Компьютер: если вообще ничего не прошло, разрешим хотя бы комплексное опробование
    if equipment_type == "computer" and "01-11-036-01" not in filtered:
        filtered.append("01-11-036-01")

    return [{"code": c, "title": FERP_TASKS[c]} for c in filtered]

# ========== СВОДНЫЕ ПОРТЫ ==========

def compute_total_switch_ports(equipment: List[Dict[str, Any]]) -> int:
    total = 0
    for it in equipment:
        et = it.get("equipment_type")
        qty = int(it.get("quantity") or 0)
        if et in {"switch", "router"}:
            pc = int(it.get("port_count") or 0)
            total += (pc + 1) * qty  # +1 логический mgmt
        elif et == "voip_phone":
            total += 1 * qty
        elif et == "ups":
            total += 1 * qty
    return total

# ========== СБОРКА ОБОРУДОВАНИЯ ИЗ ТЕКСТА (ШАБЛОН) ==========

def extract_equipment_items(doc_text: str) -> List[Dict[str, Any]]:
    """
    Заглушка-парсер: вместо полноценного NER — несколько эвристик по ключевым словам.
    Реальный парсер можно подключить позднее.
    """
    items: List[Dict[str, Any]] = []

    # Примеры эвристик (добавь свои паттерны/таблицы соответствий моделей → тип):
    # Switch
    for m in re.finditer(r"(?:коммутатор|switch)\s+(?P<model>[A-Za-z0-9\-_/]+).*?(?P<frag>(?:\d+\s*[- ]?порт\w*|\d+\s*(?:x\s*)?(?:GE|1G|10G)))?", doc_text, re.IGNORECASE | re.DOTALL):
        model = m.group("model")
        frag = m.group("frag") or ""
        items.append({
            "equipment_type": "switch",
            "priority": PRIORITY_BY_TYPE["switch"],
            "vendor": "",
            "name": f"Коммутатор {model}",
            "model_or_code": model,
            "description": "",
            "quantity": 1,
            "port_count": extract_port_count(frag),
        })

    # Router
    for m in re.finditer(r"(?:маршрутизатор|router)\s+(?P<model>[A-Za-z0-9\-_/]+)", doc_text, re.IGNORECASE):
        model = m.group("model")
        items.append({
            "equipment_type": "router",
            "priority": PRIORITY_BY_TYPE["router"],
            "vendor": "",
            "name": f"Маршрутизатор {model}",
            "model_or_code": model,
            "description": "",
            "quantity": 1,
            "port_count": 0,
        })

    # SFP
    if has_optics(doc_text):
        # грубо: по упоминанию SFP считаем, что есть отдельные модули
        count = len(re.findall(r"\bSFP\+?\b", doc_text, re.IGNORECASE)) or 1
        items.append({
            "equipment_type": "sfp",
            "priority": PRIORITY_BY_TYPE["sfp"],
            "vendor": "",
            "name": "SFP(-+) модули",
            "model_or_code": "",
            "description": "",
            "quantity": count,
            "port_count": 0,
        })

    # UPS
    if mentions_ups(doc_text):
        qty = max(1, len(re.findall(r"\b(ИБП|UPS)\b", doc_text, re.IGNORECASE)))
        items.append({
            "equipment_type": "ups",
            "priority": PRIORITY_BY_TYPE["ups"],
            "vendor": "",
            "name": "Источник бесперебойного питания",
            "model_or_code": "",
            "description": "",
            "quantity": qty,
            "port_count": 0,
        })

    # VoIP phone
    if mentions_voip(doc_text):
        qty = max(1, len(re.findall(r"\b(IP-?телефон|VoIP)\b", doc_text, re.IGNORECASE)))
        items.append({
            "equipment_type": "voip_phone",
            "priority": PRIORITY_BY_TYPE["voip_phone"],
            "vendor": "",
            "name": "IP-телефон",
            "model_or_code": "",
            "description": "",
            "quantity": qty,
            "port_count": 0,
        })

    # Computer (рабочее место)
    for _ in re.finditer(r"\b(АРМ|рабочее место|компьютер)\b", doc_text, re.IGNORECASE):
        items.append({
            "equipment_type": "computer",
            "priority": PRIORITY_BY_TYPE["computer"],
            "vendor": "",
            "name": "Компьютер (АРМ)",
            "model_or_code": "",
            "description": "",
            "quantity": 1,
            "port_count": 0,
        })

    # TODO: добавить паттерны для pdu/breaker/socket_din/ground_busbar/din_panel/cctv/cabinet/access_control/ntp-clock/other

    # Мердж дубликатов:
    by_key: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
    for it in items:
        k = unique_equipment_key(it)
        if k in by_key:
            by_key[k]["quantity"] += it.get("quantity", 1)
            # максимум портов (если одна модель встретилась с уточнением)
            by_key[k]["port_count"] = max(by_key[k].get("port_count", 0), it.get("port_count", 0))
        else:
            by_key[k] = it

    return list(by_key.values())

# ========== СБОРКА FERP/WORK/СУММАРИЕС ==========

def attach_ferp_and_work(equipment: List[Dict[str, Any]], doc_text: str) -> None:
    for it in equipment:
        et = it["equipment_type"]
        ferp = select_ferp_for_item(et, doc_text)
        it["ferp"] = ferp
        it["work"] = summarize_work_from_ferp(ferp) if ferp else ""

def build_summaries(equipment: List[Dict[str, Any]]) -> Dict[str, Any]:
    return {
        "total_switch_ports": compute_total_switch_ports(equipment),
        "notes": "В расчёт включены +1 mgmt-порт на каждый switch/router, +1 порт на каждый VoIP-телефон и ИБП.",
    }

# ========== ВХОДНАЯ ТОЧКА ==========

def extract_all(docx_text: str, project_defaults: Dict[str, str]) -> Dict[str, Any]:
    equipment = extract_equipment_items(docx_text)
    attach_ferp_and_work(equipment, docx_text)
    summaries = build_summaries(equipment)

    project = {
        "project_code": project_defaults.get("project_code", ""),
        "project_title": project_defaults.get("project_title", ""),
        "site_title": project_defaults.get("site_title", ""),
        "customer": project_defaults.get("customer", ""),
        "designer": project_defaults.get("designer", ""),
        "shcaf": project_defaults.get("shcaf", ""),
    }

    pnr = build_pnr_from_ferp(equipment, FERP_TASKS)

    return {
        "project": project,
        "equipment": equipment,
        "summaries": summaries,
        "pnr": pnr,
    }
