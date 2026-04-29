"""
okrug_parser.py — определение округа Омска по адресу из текста.

Используется и в clean-mix (для записи в output-таблицу), и в clean-resolutions
(для сопоставления с DEFAULT_OKRUG_MAP). Источник — addresses.csv с колонками
street;house;okrug, который запекается в exe через PyInstaller --add-data.

Стратегия поиска адреса в TextBody:
  1. Блок "Суть обращения: ..." (там часто реальный адрес проблемы)
  2. Строка "Почтовый адрес: ..." (адрес регистрации, fallback)
  3. Весь TextBody (последний fallback)

Матчинг — поиск по списку известных улиц (~465). Сначала длинные,
чтобы не зацепить "Ленина" в "100 лет Ленина". Дом ищется в окне
50 символов после улицы.

На Почта_ТЭС.xlsx даёт ~57% точного автоматического определения,
~1% ошибочно, остальное — улица отсутствует в БД или нестандартный
формат записи.
"""

import os
import re
import sys
import csv
import logging
from collections import defaultdict

log = logging.getLogger("asud.okrug")

_PREFIX_RE = re.compile(
    r'^(?:г\s*омск\s*,?\s*)?'
    r'(?:улица|ул\.?|проспект|пр[-]?т\.?|пр[-]?кт\.?|пр\.?|переулок|пер\.?|'
    r'бульвар|б[-]?р\.?|площадь|пл\.?|шоссе|ш\.?|набережная|наб\.?|'
    r'линия|тупик|проезд|пр[-]?д\.?|микрорайон|мкр\.?)\s*',
    re.IGNORECASE)

_HOUSE_RE = re.compile(r'(\d+[а-я]?)', re.IGNORECASE)


def _norm_text(s):
    if not s:
        return ''
    s = str(s).lower().replace('ё', 'е')
    s = re.sub(r'[«»"\'`]', '', s)
    s = re.sub(r'[.,;:()\\/]+', ' ', s)
    # "3-я молодежная" / "3я молодежная" → "3 молодежная"
    s = re.sub(r'(\d+)[\s-]*я(?=\s)', r'\1', s)
    return re.sub(r'\s+', ' ', s).strip()


def _norm_street_name(s):
    return _PREFIX_RE.sub('', _norm_text(s)).strip()


def _norm_house(s):
    if not s:
        return ''
    m = _HOUSE_RE.search(str(s).lower().replace('ё', 'е'))
    return m.group(1) if m else ''


_index = None
_streets_sorted = None


def _addresses_csv_path(base_dir_fn=None):
    """Возвращает путь к addresses.csv: внутри exe (через _MEIPASS),
    либо рядом с exe / py-скриптом."""
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        path = os.path.join(meipass, 'addresses.csv')
        if os.path.exists(path):
            return path
    if base_dir_fn:
        path = os.path.join(base_dir_fn(), 'addresses.csv')
        if os.path.exists(path):
            return path
    # last resort — current dir
    if os.path.exists('addresses.csv'):
        return 'addresses.csv'
    return None


def _build_index(base_dir_fn=None):
    global _index, _streets_sorted
    if _index is not None:
        return _index, _streets_sorted
    path = _addresses_csv_path(base_dir_fn)
    if not path:
        log.warning("addresses.csv не найден — авто-определение округа отключено")
        _index = {}
        _streets_sorted = []
        return _index, _streets_sorted
    try:
        idx = defaultdict(set)
        with open(path, encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')
            header = next(reader, None)
            if not header:
                _index = {}; _streets_sorted = []
                return _index, _streets_sorted
            cols = {h.lower(): i for i, h in enumerate(header)}
            if 'street' not in cols or 'house' not in cols or 'okrug' not in cols:
                log.warning(f"Неподдерживаемый формат addresses.csv: {header}")
                _index = {}; _streets_sorted = []
                return _index, _streets_sorted
            for row in reader:
                if not row or len(row) <= max(cols['street'], cols['house'], cols['okrug']):
                    continue
                s = row[cols['street']].strip().lower()
                h = row[cols['house']].strip().lower()
                o = row[cols['okrug']].strip()
                if s and h and o:
                    # Базовый ключ
                    idx[s].add((h, o))
                    # Алиасы для номерных улиц:
                    # 'молодежная 3-я' → также 'молодежная 3', '3-я молодежная',
                    # '3 молодежная'
                    m = re.match(r'^(.+?)\s+(\d+)[\s-]*я$', s)
                    if m:
                        base, num = m.group(1).strip(), m.group(2)
                        idx[f"{base} {num}"].add((h, o))
                        idx[f"{num}-я {base}"].add((h, o))
                        idx[f"{num} {base}"].add((h, o))
        _index = idx
        _streets_sorted = sorted(idx.keys(), key=lambda x: -len(x))
        log.info(f"Street index: {len(idx)} улиц")
    except Exception as e:
        log.warning(f"Ошибка построения street index: {e}")
        _index = {}
        _streets_sorted = []
    return _index, _streets_sorted


def _find_street_house(text, idx, sorted_streets):
    norm = _norm_text(text)
    for street in sorted_streets:
        if len(street) < 3:
            continue
        pos = 0
        while True:
            i = norm.find(street, pos)
            if i < 0:
                break
            left_ok = i == 0 or not norm[i - 1].isalnum()
            end = i + len(street)
            right_ok = end == len(norm) or not norm[end].isalnum()
            if left_ok and right_ok:
                tail = norm[end:end + 50]
                m = re.search(r'\b(\d+[а-я]?)', tail)
                if m:
                    return (street, m.group(1))
            pos = i + 1
    return (None, None)


def okrug_from_textbody(textbody, base_dir_fn=None):
    """Извлекает округ ('КАО'/'САО'/'ЦАО'/'ОАО'/'ЛАО') из TextBody. Возвращает None если не нашли."""
    if not textbody:
        return None
    idx, sorted_streets = _build_index(base_dir_fn)
    if not idx:
        return None
    text = str(textbody)
    fragments = []
    m = re.search(r'суть\s+обращени[яе]\s*:?\s*([\s\S]+?)(?:\n\s*\n|$)',
                  text, re.IGNORECASE)
    if m:
        fragments.append(('суть', m.group(1)))
    m = re.search(r'почтов[а-я]+\s+адрес[а-я]*\s*:\s*([^\n]+)',
                  text, re.IGNORECASE)
    if m:
        fragments.append(('почт', m.group(1)))
    fragments.append(('весь', text))

    for name, frag in fragments:
        street, house = _find_street_house(frag, idx, sorted_streets)
        if street and house:
            # Нормализация дома для сравнения: '15г' → '15', '6в' → '6'.
            # Литерные суффиксы (а/б/в/г) обозначают разные строения одного
            # дома — обычно в одном округе, можно игнорировать.
            house_digits = re.match(r'\d+', house)
            house_num = house_digits.group(0) if house_digits else house
            for h, o in idx[street]:
                # В БД тоже могут быть дома с буквами — нормализуем обе стороны
                h_digits = re.match(r'\d+', h)
                h_num = h_digits.group(0) if h_digits else h
                if h == house or h_num == house_num:
                    return o
    return None
