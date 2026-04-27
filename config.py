"""
config.py — Настройки АСУД-скрипта.

Читает config.json рядом с exe (если есть), иначе дефолты.
Менять настройки можно без пересборки — просто правишь JSON.
"""

import os
import sys
import json
import logging

log = logging.getLogger("asud.config")

# Дефолтные значения
DEFAULTS = {
    "asud_url": "https://asud.interrao.ru/asudik/",
    "timeout": 20,
    "outlook_dir": r"D:\OutlookSubjects\ТЭС",
    "addressees": [
        "Басманов Александр Владимирович",
        "Халецкая Юлия Владимировна",
    ],
    "unknown_correspondent": "Неизвестный Неизвестный Неизвестный",
    "delivery_method": "Электронная почта",
    "sheet_name": "Лист2",
}

# Маппинг индекса из Excel → название вида в АСУД
DOC_TYPE_MAP = {
    1: "Указы, распоряжения Президента Российской Федерации",
    2: "Документы Администрации Президента",
    3: "Документы Правительства Российской Федерации",
    4: "Документы Федеральных органов исполнительной и законодательной власти",
    5: "Письма юридических лиц",
    6: "Письма компаний Топливно-энергетического комплекса",
    7: "Документы органов законодательной и исполнительной власти субъектов",
    8: "Письма, заявления и жалобы граждан, акционеров",
}


def get_base_dir():
    """Папка где лежит exe/скрипт."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load():
    """Загружает config.json, мержит с дефолтами."""
    cfg = dict(DEFAULTS)
    path = os.path.join(get_base_dir(), "config.json")
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                user_cfg = json.load(f)
            cfg.update(user_cfg)
            log.info(f"Конфиг загружен: {path}")
        except Exception as e:
            log.warning(f"Ошибка чтения config.json: {e}, используем дефолты")
    return cfg
