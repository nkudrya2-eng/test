# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Iterable, Tuple, Optional
import re

FERP_CODE_RE = re.compile(r"^\d{2}-\d{2}-\d{3}-\d{2}$")

# === 1) ЕДИНЫЙ ИСТОЧНИК ПРАВДЫ ПО ФЕРп =======================================

@dataclass(frozen=True)
class FerpInfo:
    code: str
    title: str
    unit: str
    note: str

# Единый справочник: каждый код описывается полностью (title, unit, note).
# Все производные словари (FERP_TASKS/UNIT_BY_CODE/NOTE_BY_CODE) строятся из него.
FERP_INFO: Dict[str, FerpInfo] = {
    "01-11-010-01": FerpInfo(
        code="01-11-010-01",
        title="Измерение сопротивления растеканию тока заземляющего устройства",
        unit="измерение",
        note="Проверка сопротивления заземлителя и контура",
    ),
    "01-11-011-01": FerpInfo(
        code="01-11-011-01",
        title="Проверка наличия электрической цепи между заземлителями и заземлёнными элементами оборудования",
        unit="измерение",
        note="Контроль электрической связи контура, шкафа и металлических частей оборудования",
    ),
    "01-11-011-02": FerpInfo(
        code="01-11-011-02",
        title="Проверка целостности экранов и заземления кабельных трасс",
        unit="измерение",
        note="Контроль экранирования и заземления в шкафах и кабелях СКС",
    ),
    "01-11-012-01": FerpInfo(
        code="01-11-012-01",
        title="Проверка напряжения и частоты в цепях электропитания",
        unit="измерение",
        note="Контроль параметров электросети перед включением оборудования",
    ),
    "01-11-026-01": FerpInfo(
        code="01-11-026-01",
        title="Проверка фазировки и полярности цепей электропитания",
        unit="точка",
        note="Контроль правильности чередования фаз и подключения нулевого проводника",
    ),
    "01-11-028-01": FerpInfo(
        code="01-11-028-01",
        title="Измерение сопротивления изоляции кабельных линий напряжением до 1 кВ",
        unit="шт",
        note="Проверка изоляции силовых цепей электропитания",
    ),
    "01-11-029-01": FerpInfo(
        code="01-11-029-01",
        title="Проверка электрической изоляции цепей заземления от токоведущих частей",
        unit="измерение",
        note="Отсутствие замыканий PE на рабочие цепи",
    ),
    "01-11-034-01": FerpInfo(
        code="01-11-034-01",
        title="Проверка работы источника бесперебойного питания (ИБП)",
        unit="шт",
        note="Тест перехода на батареи, измерение времени автономной работы",
    ),
    "01-11-036-01": FerpInfo(
        code="01-11-036-01",
        title="Проверка готовности системы к приёмо-сдаточным испытаниям (общее/комплексное опробование)",
        unit="комплекс",
        note="Итоговое комплексное испытание всех подсистем",
    ),
    "01-06-022-01": FerpInfo(
        code="01-06-022-01",
        title="Проверка функционирования структурированной кабельной системы (СКС)",
        unit="точка",
        note="Тестирование линий витой пары, оформление протоколов",
    ),
    "02-01-002-01": FerpInfo(
        code="02-01-002-01",
        title="Проверка целостности и параметров линий связи",
        unit="линия",
        note="Контроль целостности и параметров каналов связи между узлами",
    ),
    "02-01-003-01": FerpInfo(
        code="02-01-003-01",
        title="Проверка функционирования систем сигнализации или телемеханики",
        unit="канал",
        note="Проверка передачи сигналов и индикации по каналам связи",
    ),
    "02-01-004-01": FerpInfo(
        code="02-01-004-01",
        title="Проверка уровней оптической мощности на оконечном оборудовании",
        unit="точка",
        note="Контроль потерь в волоконно-оптических линиях",
    ),
}

# Производные словари (для обратной совместимости с остальным кодом)
FERP_TASKS: Dict[str, str] = {k: v.title for k, v in FERP_INFO.items()}
UNIT_BY_CODE: Dict[str, str] = {k: v.unit for k, v in FERP_INFO.items()}
NOTE_BY_CODE: Dict[str, str] = {k: v.note for k, v in FERP_INFO.items()}

# Самопроверка корректности кодов:
for code in FERP_INFO:
    assert FERP_CODE_RE.match(code), f"Bad FERP code format: {code}"

# === 2) ПОРЯДОК ОТОБРАЖЕНИЯ (БЕЗ ДУБЛЕЙ В КЛЮЧЕВОМ КОДЕ) ====================

# Разрешаем псевдометки, чтобы один и тот же код можно было вывести в разных местах,
# но при нормализации они схлопываются в базовый код.
FERP_ORDER_LABELLED: List[str] = [
    "01-11-010-01",
    "01-11-011-01#bonding",     # непрерывность защитной цепи
    "01-11-028-01",
    "01-06-022-01",
    "02-01-002-01",
    "01-11-026-01",
    "01-11-012-01",
    "01-11-034-01",
    "01-11-011-02",
    "02-01-003-01",
    "02-01-004-01",
    "01-11-011-01#racks",       # заземление корпусов стоек/РУ
    "01-11-029-01",
    "01-11-036-01",
]

def _base_code(labelled_code: str) -> str:
    return labelled_code.split("#", 1)[0]

def normalize_ferp_order() -> List[str]:
    """Возвращает порядок без дублей по базовому коду, сохраняя первое вхождение."""
    seen = set()
    result: List[str] = []
    for item in FERP_ORDER_LABELLED:
        base = _base_code(item)
        if base not in seen:
            seen.add(base)
            result.append(base)
    return result

FERP_ORDER: List[str] = normalize_ferp_order()

# === 3) МАППИНГ ОБОРУДОВАНИЯ К ТИПОВЫМ ФЕРп ==================================

EQUIPMENT_FERP_MAP: Dict[str, List[str]] = {
    "switch": ["01-11-011-01", "01-11-012-01", "01-11-026-01", "02-01-002-01", "02-01-004-01", "01-11-036-01"],
    "router": ["01-11-011-01", "01-11-012-01", "02-01-002-01", "02-01-004-01", "01-11-036-01"],
    "sfp": ["02-01-004-01", "02-01-002-01"],
    "ups": ["01-11-010-01", "01-11-011-01", "01-11-012-01", "01-11-026-01", "01-11-034-01", "01-11-036-01"],
    "voip_phone": ["01-06-022-01", "02-01-002-01", "01-11-036-01"],
    "computer": ["01-06-022-01", "02-01-002-01", "01-11-036-01"],
    "pdu": ["01-11-011-01", "01-11-012-01", "01-11-026-01", "01-11-036-01"],
    "breaker": ["01-11-012-01", "01-11-026-01", "01-11-036-01"],
    "socket_din": ["01-11-012-01", "01-11-026-01", "01-11-036-01"],
    "ground_busbar": ["01-11-010-01", "01-11-011-01", "01-11-029-01", "01-11-036-01"],
    "din_panel": ["01-11-011-01", "01-11-036-01"],
    "cctv": ["02-01-002-01", "01-11-036-01"],
    "cabinet": ["01-11-011-01", "01-11-036-01"],
    "cable_management": ["01-11-011-02", "01-11-036-01"],
    "access_control": ["02-01-003-01", "01-11-036-01"],
    "ntp-clock": ["01-11-011-01", "02-01-002-01", "01-11-036-01"],
    "other": ["01-11-036-01"],
}

# Самопроверка: все коды из EQUIPMENT_FERP_MAP существуют в FERP_INFO
for eq_type, codes in EQUIPMENT_FERP_MAP.items():
    for code in codes:
        base = _base_code(code)
        assert base in FERP_INFO, f"Unknown FERP code in map: {eq_type} -> {code}"

PRIORITY_BY_TYPE: Dict[str, int] = {
    "switch": 1, "router": 1, "voip_phone": 2, "ups": 3, "pdu": 4, "sfp": 5,
    "breaker": 6, "socket_din": 7, "ground_busbar": 8, "din_panel": 9, "other": 10,
    "cctv": 10, "cabinet": 10, "cable_management": 10, "access_control": 10, "computer": 10, "ntp-clock": 10,
}

# === 4) ПУБЛИЧНЫЕ ХЕЛПЕРЫ =====================================================

def get_ferp_by_type(equipment_type: str) -> List[FerpInfo]:
    """Возвращает типовые FerpInfo для данного типа оборудования (без дублей и с валидацией)."""
    codes = EQUIPMENT_FERP_MAP.get(equipment_type, ["01-11-036-01"])  # fallback: комплексное опробование
    unique: List[str] = []
    seen = set()
    for c in codes:
        base = _base_code(c)
        if base not in seen:
            seen.add(base)
            unique.append(base)
    return [FERP_INFO[c] for c in unique]

def get_priority(equipment_type: str) -> int:
    """Приоритет типа оборудования: меньше — выше в выдаче."""
    return PRIORITY_BY_TYPE.get(equipment_type, 10)

def format_work(ferp_codes: Iterable[str]) -> str:
    """
    Компактная сводка 'work' на основе выбранных ФЕРп.
    Пример: 'Проверка заземления, проверка питания, проверка линий связи, комплексное опробование.'
    """
    bases = []
    seen = set()
    for c in ferp_codes:
        base = _base_code(c)
        if base in FERP_INFO and base not in seen:
            seen.add(base)
            bases.append(FERP_INFO[base].title)
    # деликатно укорачиваем заголовки к коротким смысловым фразам
    def short(title: str) -> str:
        for lead in ("Проверка ", "Измерение ", "Контроль ", "Проверка функционирования "):
            if title.startswith(lead):
                return title[len(lead):]
        return title
    parts = [short(t) for t in bases]
    return (", ".join(parts) + ".") if parts else ""

# === 5) ДОК-ТЕСТЫ И БЫСТРЫЕ ПРОВЕРКИ ==========================================

if __name__ == "__main__":
    # Проверка порядка без дублей
    assert "01-11-011-01" in FERP_ORDER
    assert FERP_ORDER.count("01-11-011-01") == 1

    # Проверка маппинга
    sw = get_ferp_by_type("switch")
    assert any(x.code == "02-01-002-01" for x in sw)
    assert all(FERP_CODE_RE.match(x.code) for x in sw)

    # Форматирование work
    w = format_work(["01-11-011-01#bonding", "02-01-002-01", "01-11-036-01"])
    assert "заземлителями" in w or "заземления" in w
