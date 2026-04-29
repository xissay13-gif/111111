"""
resolutions.py — Выдача резолюций по обращениям из реестра.

Запускается под учёткой Есиной М.В., автоматически переключается
на учётку Халецкой Ю.В. (через выпадашку в шапке профиля).
Заходит в раздел "На резолюцию", для каждой строки реестра находит
соответствующий документ в АСУД и выдаёт резолюцию начальнику
абонентского отдела (по округу из колонки ao/fio в реестре).

Реестр: Лист2 с колонками Link, Subject, TextBody, Тема, To, LS, ao, fio.
Матч документа в АСУД: по номеру обращения из TextBody (regex
"обращение № NNNNNN"), fallback на подстроку первых 60 символов TextBody.
"""

import os
import re
import sys
import time
import logging
from datetime import date, datetime, timedelta

import openpyxl
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config as cfg
from ui import (click, wait_and_click, find_input_near_label,
                wait_asud_loaded, wait_modal_closed, close_open_modals, js_set_value)
from correspondent import match_correspondent

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger("asud.res")
start_time = time.monotonic()

settings = {}


# ============================================================
# EXCEL
# ============================================================

APPEAL_RE = re.compile(r'обращени[ея]\s*№?\s*(\d{4,10})', re.IGNORECASE)


# Lookup street_name → set of (house, okrug_short) — индекс для матчинга
# адресов из TextBody с адресной БД. Загружается лениво.
_STREET_INDEX = None
_ALL_STREETS_SORTED = None  # отсортированы по длине (длинные первыми)


def _addresses_csv_path():
    """Возвращает путь к addresses.csv: внутри exe (через _MEIPASS)
    или рядом с .py-скриптом в dev-режиме."""
    # PyInstaller --onefile распаковывает данные в sys._MEIPASS
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        path = os.path.join(meipass, 'addresses.csv')
        if os.path.exists(path):
            return path
    # Fallback: рядом с exe / скриптом
    base = cfg.get_base_dir()
    path = os.path.join(base, 'addresses.csv')
    if os.path.exists(path):
        return path
    return None


# --- Нормализация для парсинга адресов ---

_PREFIX_RE = re.compile(
    r'^(?:г\s*омск\s*,?\s*)?'
    r'(?:улица|ул\.?|проспект|пр[-]?т\.?|пр[-]?кт\.?|пр\.?|переулок|пер\.?|'
    r'бульвар|б[-]?р\.?|площадь|пл\.?|шоссе|ш\.?|набережная|наб\.?|'
    r'линия|тупик|проезд|пр[-]?д\.?|микрорайон|мкр\.?)\s*',
    re.IGNORECASE)

_HOUSE_RE = re.compile(r'(\d+[а-я]?)(?:[/\\-](\d+[а-я]?))?', re.IGNORECASE)


def _norm_text_for_match(s):
    """Нормализация для сопоставления: lower, ё→е, без пунктуации, single space."""
    if not s:
        return ''
    s = str(s).lower().replace('ё', 'е')
    s = re.sub(r'[«»"\'`]', '', s)
    s = re.sub(r'[.,;:()\\/]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _norm_street_name(s):
    """Только название улицы без префиксов."""
    return _PREFIX_RE.sub('', _norm_text_for_match(s)).strip()


def _norm_house_main(s):
    """Возвращает основной номер дома (первое число)."""
    if not s: return ''
    m = _HOUSE_RE.search(str(s).lower().replace('ё', 'е'))
    return m.group(1) if m else ''


def _build_street_index():
    """Парсит addresses.csv в индекс {street_norm: set((house, okrug_short))}.
    Используется один раз (кешируется в _STREET_INDEX)."""
    global _STREET_INDEX, _ALL_STREETS_SORTED
    if _STREET_INDEX is not None:
        return _STREET_INDEX, _ALL_STREETS_SORTED
    path = _addresses_csv_path()
    if not path:
        log.warning("addresses.csv не найден — авто-определение округа отключено")
        _STREET_INDEX = {}
        _ALL_STREETS_SORTED = []
        return _STREET_INDEX, _ALL_STREETS_SORTED

    # CSV: LS;okrug
    # Старая версия CSV содержит только LS+okrug, но для адрес-парсинга
    # нужны улицы из исходного xlsx. Если у нас только CSV без улиц —
    # парсер не сработает; идём через колонку ao реестра.
    # В будущем addresses.csv можно расширить колонкой улицы.
    try:
        import csv
        from collections import defaultdict
        idx = defaultdict(set)
        with open(path, encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')
            header = next(reader, None)
            # Пытаемся определить колонки
            if not header or 'okrug' not in [h.lower() for h in header]:
                _STREET_INDEX = {}
                _ALL_STREETS_SORTED = []
                return _STREET_INDEX, _ALL_STREETS_SORTED
            # Если в CSV есть street/house — используем
            cols = {h.lower(): i for i, h in enumerate(header)}
            if 'street' in cols and 'house' in cols:
                for row in reader:
                    if not row or len(row) <= max(cols['street'], cols['house'], cols['okrug']):
                        continue
                    street = _norm_street_name(row[cols['street']])
                    house = _norm_house_main(row[cols['house']])
                    okrug = row[cols['okrug']].strip()
                    if street and house and okrug:
                        idx[street].add((house, okrug))
        log.info(f"Street index: {len(idx)} улиц")
        _STREET_INDEX = idx
        _ALL_STREETS_SORTED = sorted(idx.keys(), key=lambda s: -len(s))
    except Exception as e:
        log.warning(f"Ошибка построения street index: {e}")
        _STREET_INDEX = {}
        _ALL_STREETS_SORTED = []
    return _STREET_INDEX, _ALL_STREETS_SORTED


def _find_street_house(text, idx, sorted_streets):
    """Ищет известную улицу в нормализованном тексте, потом дом рядом."""
    norm = _norm_text_for_match(text)
    for street in sorted_streets:
        if len(street) < 3:
            continue
        # Поиск как целое слово
        pos = 0
        while True:
            i = norm.find(street, pos)
            if i < 0: break
            left_ok = i == 0 or not norm[i-1].isalnum()
            end = i + len(street)
            right_ok = end == len(norm) or not norm[end].isalnum()
            if left_ok and right_ok:
                # Дом в окне 50 символов после улицы
                tail = norm[end:end+50]
                m = re.search(r'\b(\d+[а-я]?)', tail)
                if m:
                    return (street, m.group(1))
            pos = i + 1
    return (None, None)


def _okrug_from_textbody(textbody):
    """Извлекает округ из TextBody.

    Стратегия: суть обращения → почтовый адрес → весь текст.
    Возвращает короткий код округа ('КАО', 'ЦАО', ...) или None.
    """
    if not textbody:
        return None
    idx, sorted_streets = _build_street_index()
    if not idx:
        return None
    text = str(textbody)
    fragments = []
    # 1) Суть обращения
    m = re.search(r'суть\s+обращени[яе]\s*:?\s*([\s\S]+?)(?:\n\s*\n|$)',
                  text, re.IGNORECASE)
    if m:
        fragments.append(('суть', m.group(1)))
    # 2) Почтовый адрес
    m = re.search(r'почтов[а-я]+\s+адрес[а-я]*\s*:\s*([^\n]+)',
                  text, re.IGNORECASE)
    if m:
        fragments.append(('почт', m.group(1)))
    # 3) Весь текст как fallback
    fragments.append(('весь', text))

    for name, frag in fragments:
        street, house = _find_street_house(frag, idx, sorted_streets)
        if street and house:
            for h, o in idx[street]:
                if h == house:
                    log.info(f"  адрес [{name}]: {street} {house} → {o}")
                    return o
    return None


def _resolve_executor(ao, fio, ls=None, textbody=None):
    """Возвращает ФИО начальника по приоритету:
    1. fio из реестра (если заполнено)
    2. ao из реестра → DEFAULT_OKRUG_MAP
    3. адрес из TextBody → addresses.csv → DEFAULT_OKRUG_MAP

    Параметр ls пока не используется — формат LS в реестре не совпадает
    с адресной БД (11 цифр vs 6), нужно правило конвертации.
    """
    # 1. fio напрямую (вручную проставленный)
    if fio and str(fio).strip():
        return str(fio).strip()
    # 2. ao из реестра (вручную проставленный)
    if ao:
        key = str(ao).strip()
        v = cfg.DEFAULT_OKRUG_MAP.get(key)
        if v:
            return v
    # 3. Парсинг адреса из TextBody
    if textbody:
        ao_short = _okrug_from_textbody(textbody)
        if ao_short:
            v = cfg.DEFAULT_OKRUG_MAP.get(ao_short)
            if v:
                return v
    return None


def load_excel(file_path):
    """Читает Лист2. Колонки: Link, Subject, TextBody, Тема, To, LS, ao, fio."""
    sheet_name = settings.get("sheet_name", cfg.DEFAULTS["sheet_name"])
    wb = openpyxl.load_workbook(file_path, data_only=True)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        log.warning(f"Лист '{sheet_name}' не найден, использую активный: {wb.active.title}")
        ws = wb.active

    rows = []
    skipped = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True), 2):
        if not row or len(row) < 8:
            skipped += 1
            continue
        link = row[0]
        subject = row[1]
        textbody = row[2] or ''
        # row[3] = Тема (тип, не нужно для резолюций)
        # row[4] = To
        # row[5] = LS
        ls = row[5] if len(row) > 5 else None
        ao = row[6] if len(row) > 6 else None
        fio = row[7] if len(row) > 7 else None

        if not link and not subject:
            skipped += 1
            continue

        executor = _resolve_executor(ao, fio, ls, textbody)
        appeal_no = _extract_appeal_no(textbody)

        rows.append({
            "row_idx": row_idx,
            "link": link,
            "subject": str(subject or '').strip(),
            "textbody": str(textbody),
            "ls": _norm_ls(ls),
            "ao": str(ao or '').strip(),
            "executor": executor,
            "appeal_no": appeal_no,
        })

    wb.close()

    log.info(f"Загружено: {len(rows)}, пропущено: {skipped}")
    no_executor = sum(1 for r in rows if not r['executor'])
    no_match = sum(1 for r in rows if not r['appeal_no'])
    if no_executor:
        log.warning(f"  без исполнителя: {no_executor} (будут пропущены)")
    if no_match:
        log.warning(f"  без номера обращения: {no_match} (fallback на подстроку TextBody)")

    return rows


# ============================================================
# UI: переключение учётки
# ============================================================

def switch_account(driver, target_substring):
    """Переключается на учётку, ФИО которой содержит target_substring.
    1. Клик по dropdown-стрелке в шапке профиля
    2. Клик по пункту с нужным ФИО в выпадашке
    """
    log.info(f"Переключение на учётку: {target_substring}")
    time.sleep(2)

    # Шаг 1: клик ▼ рядом с именем профиля
    try:
        # Стрелочка вниз обычно лежит рядом с блоком пользователя в самом верху страницы
        # На скрине она 21x55px, расположена сразу справа от блока с именем
        # Ищем по клик-целям рядом с блоком пользователя
        triggers = driver.find_elements(By.CSS_SELECTOR,
            "img[class*='trigger'], img[class*='Trigger'], div[class*='trigger']")
        # Выбираем те что в верхней части страницы (top < 100px)
        candidate = None
        for t in triggers:
            try:
                if not t.is_displayed():
                    continue
                rect = t.rect
                if rect.get('y', 0) < 120 and rect.get('x', 0) < 400:
                    candidate = t
                    break
            except Exception:
                continue
        if not candidate:
            # Fallback — любая стрелка/треугольник в верхнем левом углу
            log.warning("Стрелка профиля не найдена по trigger-классам, пробую по позиции")
            all_imgs = driver.find_elements(By.TAG_NAME, "img")
            for im in all_imgs:
                try:
                    if not im.is_displayed(): continue
                    rect = im.rect
                    if rect.get('y', 0) < 120 and rect.get('x', 0) < 400 \
                            and 10 <= rect.get('width', 0) <= 30:
                        candidate = im
                        break
                except Exception:
                    continue
        if not candidate:
            log.error("Не нашёл dropdown-стрелку профиля")
            return False
        click(driver, candidate, "▼ профиль")
        time.sleep(1.5)
    except Exception as e:
        log.error(f"Ошибка клика по dropdown профиля: {e}")
        return False

    # Шаг 2: клик по пункту со строкой target_substring
    try:
        WebDriverWait(driver, 10).until(
            lambda d: any(target_substring in (e.text or '')
                          for e in d.find_elements(By.CSS_SELECTOR, "div, span, label")
                          if e.is_displayed()))
    except Exception:
        log.warning(f"Пункт '{target_substring}' в выпадашке не появился")

    items = driver.find_elements(By.XPATH,
        f"//*[contains(normalize-space(text()), '{target_substring}')]")
    target = None
    for it in items:
        try:
            if not it.is_displayed():
                continue
            rect = it.rect
            # должен быть всплывающий пункт под профилем (y > стрелки)
            if rect.get('y', 0) > 100:
                target = it
                break
        except Exception:
            continue
    if not target and items:
        target = next((i for i in items if i.is_displayed()), None)

    if not target:
        log.error(f"Пункт с '{target_substring}' не найден")
        return False

    click(driver, target, f"учётка {target_substring}")
    time.sleep(3)
    log.info("Переключение запущено, жду перезагрузку АСУД")
    wait_asud_loaded(driver)
    return True


# ============================================================
# UI: сайдбар → "На резолюцию"
# ============================================================

def click_sidebar_section(driver, section_text):
    """Клик по пункту в левом сайдбаре."""
    log.info(f"Сайдбар → '{section_text}'")
    items = driver.find_elements(By.XPATH,
        f"//*[normalize-space(text())='{section_text}']")
    target = None
    for it in items:
        try:
            if it.is_displayed():
                target = it
                break
        except Exception:
            continue
    if not target:
        log.error(f"Пункт сайдбара '{section_text}' не найден")
        return False
    click(driver, target, f"сайдбар: {section_text}")
    time.sleep(2)
    return True


# ============================================================
# UI: поиск и открытие документа в списке "На резолюцию"
# ============================================================

LIST_TABLE_ID = "CABINET_MENU__RECEIVED__ALL_ACTIVE__TO_RESOLUTION"


def find_doc_row(driver, doc, timeout=10):
    """Находит <tr> в таблице списка, который соответствует doc (строке реестра).

    Стратегии (по убыванию точности):
      1. По номеру обращения "5417313" из TextBody
      2. По первым 60 символам TextBody (substring)
      3. По Subject (как очень слабый fallback)
    """
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        # 1) номер обращения
        if doc.get('appeal_no'):
            try:
                row = driver.find_element(By.XPATH,
                    f"//table[@id='{LIST_TABLE_ID}']"
                    f"//tr[contains(., '{doc['appeal_no']}')]")
                if row.is_displayed():
                    log.info(f"  match: appeal № {doc['appeal_no']}")
                    return row
            except Exception:
                pass

        # 2) подстрока TextBody (первые 60 символов)
        body = (doc.get('textbody') or '').replace('\xa0', ' ').strip()
        snippet = re.sub(r'\s+', ' ', body)[:60].strip()
        # экранируем кавычки для XPath
        snippet_safe = snippet.replace("'", "")[:60].strip()
        if len(snippet_safe) >= 20:
            try:
                row = driver.find_element(By.XPATH,
                    f"//table[@id='{LIST_TABLE_ID}']"
                    f"//tr[contains(., \"{snippet_safe}\")]")
                if row.is_displayed():
                    log.info(f"  match: подстрока TextBody")
                    return row
            except Exception:
                pass

        # 3) по Subject
        subj = (doc.get('subject') or '').strip()
        if subj and len(subj) >= 10:
            try:
                row = driver.find_element(By.XPATH,
                    f"//table[@id='{LIST_TABLE_ID}']"
                    f"//tr[contains(., '{subj[:40]}')]")
                if row.is_displayed():
                    log.info(f"  match: subject")
                    return row
            except Exception:
                pass

        time.sleep(0.5)

    return None


def open_doc_card(driver, row):
    """Открывает карточку документа двойным кликом по строке."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
        time.sleep(0.3)
        ActionChains(driver).move_to_element(row).pause(0.2).double_click().perform()
        time.sleep(2)
        # Ждём появления кнопки "Создать резолюцию" — индикатор открытой карточки
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "header-action-btn-add_resolution")))
            return True
        except Exception:
            log.warning("Карточка не открылась за 15 сек")
            return False
    except Exception as e:
        log.error(f"Не удалось открыть карточку: {e}")
        return False


# ============================================================
# UI: диалог "Корневая резолюция"
# ============================================================

def _wait_data_value(driver, container, target_value="true", timeout=5):
    """Ждёт пока у контейнера-тоггла data-value станет target_value."""
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            if container.get_attribute('data-value') == target_value:
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def toggle_switch(driver, label_text, target_value="true"):
    """Переключает тоггл рядом с label_text в нужное состояние."""
    try:
        label = driver.find_element(By.XPATH,
            f"//*[normalize-space(text())='{label_text}']")
        container = label.find_element(By.XPATH,
            "./following::*[contains(@class,'switcherContainer')][1]")
    except Exception as e:
        log.warning(f"Тоггл '{label_text}' не найден: {e}")
        return False

    cur = container.get_attribute('data-value')
    if cur == target_value:
        log.info(f"Тоггл '{label_text}' уже = {target_value}")
        return True

    click(driver, container, f"тоггл {label_text} → {target_value}")
    if _wait_data_value(driver, container, target_value, timeout=3):
        log.info(f"Тоггл '{label_text}' = {target_value}")
        return True
    log.warning(f"Тоггл '{label_text}' не переключился")
    return False


def select_content_template(driver, template_text):
    """Выбирает в поле "Содержание" пункт из выпадашки."""
    try:
        # input с placeholder "Общие формулировки"
        inp = None
        candidates = driver.find_elements(By.CSS_SELECTOR,
            "input[placeholder='Общие формулировки']")
        for c in candidates:
            if c.is_displayed():
                inp = c
                break
        if not inp:
            # Fallback — поиск по label "Содержание"
            inp = find_input_near_label(driver, "Содержание")
        if not inp:
            log.error("Поле 'Содержание' не найдено")
            return False

        click(driver, inp, "Содержание")
        time.sleep(1)

        # Дропдаун с пунктами
        items = driver.find_elements(By.XPATH,
            f"//*[normalize-space(text())='{template_text}']")
        target = None
        for it in items:
            try:
                if it.is_displayed():
                    target = it
                    break
            except Exception:
                continue
        if not target:
            log.error(f"Пункт '{template_text}' в выпадашке не найден")
            return False
        click(driver, target, f"Содержание: {template_text}")
        time.sleep(0.5)
        return True
    except Exception as e:
        log.error(f"Ошибка выбора шаблона содержания: {e}")
        return False


def add_business_days(start, days):
    cur = start
    added = 0
    while added < days:
        cur += timedelta(days=1)
        if cur.weekday() < 5:  # 0-4 = пн-пт
            added += 1
    return cur


def set_stage_date(driver, n_workdays):
    """Заполняет дату в поле 'Контрольный этап' = today + n рабочих дней."""
    deadline = add_business_days(date.today(), n_workdays).strftime("%d.%m.%Y")
    try:
        inp = driver.find_element(By.CSS_SELECTOR,
            "input[id*='stage_control_date']")
        js_set_value(driver, inp, deadline)
        log.info(f"Контрольный этап: {deadline}")
        return True
    except Exception as e:
        log.warning(f"Поле даты этапа не найдено: {e}")
        return False


def fill_executor(driver, fio):
    """Вбивает ФИО в поле 'Исполнитель' (combobox), выбирает из выпадашки."""
    try:
        # Ищем поле по label "Исполнитель"
        inp = find_input_near_label(driver, "Исполнитель")
        if not inp:
            # Fallback — input id select_combobox-input
            inp = driver.find_element(By.ID, "select_combobox-input")
        if not inp:
            log.error("Поле 'Исполнитель' не найдено")
            return False

        surname = fio.split()[0]
        inp.click()
        time.sleep(0.3)
        inp.clear()
        time.sleep(0.3)
        for ch in surname:
            inp.send_keys(ch)
            time.sleep(0.08)
        log.info(f"Введена фамилия: {surname}")
        time.sleep(2)

        # Кандидаты в выпадашке
        results = driver.find_elements(By.XPATH,
            f"//*[contains(text(),'{surname}')]")
        candidates = [r for r in results
                      if r.is_displayed() and r != inp
                      and r.tag_name.lower() != 'input']
        if not candidates:
            inp.send_keys(Keys.ENTER)
            time.sleep(2)
            results = driver.find_elements(By.XPATH,
                f"//*[contains(text(),'{surname}')]")
            candidates = [r for r in results
                          if r.is_displayed() and r != inp
                          and r.tag_name.lower() != 'input']

        log.info(f"Кандидатов: {len(candidates)}")
        target = None
        for idx, r in enumerate(candidates, 1):
            try:
                txt = (r.text or '').strip()
                if len(txt) > 200:
                    continue
                ok = match_correspondent(txt, fio)
                preview = txt.replace('\n', ' ')[:80]
                log.info(f"  [{idx}] {'OK' if ok else '--'} | {preview!r}")
                if ok and target is None:
                    target = r
            except Exception:
                continue

        if not target:
            log.error(f"Исполнитель '{fio}' не найден в выпадашке")
            return False

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
        time.sleep(0.3)
        ActionChains(driver).move_to_element(target).pause(0.3).click().perform()
        time.sleep(1)
        log.info(f"Исполнитель выбран: {fio}")
        return True
    except Exception as e:
        log.error(f"Ошибка ввода исполнителя: {e}")
        return False


def _wait_button_enabled(driver, btn_id, timeout=15):
    """Ждёт пока кнопка с data-disabled='1' не станет активной."""
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            btn = driver.find_element(By.ID, btn_id)
            if btn.get_attribute('data-disabled') != '1' and btn.is_displayed():
                return btn
        except Exception:
            pass
        time.sleep(0.3)
    return None


def submit_resolution(driver):
    """Финальный шаг — клик 'Сохранить и отправить'. Проверяет что диалог закрылся."""
    btn = _wait_button_enabled(driver, "save_and_send_btn", timeout=15)
    if not btn:
        log.error("Кнопка 'Сохранить и отправить' не активировалась")
        return False
    click(driver, btn, "Сохранить и отправить")
    time.sleep(3)

    # Проверка что диалог закрылся
    try:
        title = driver.find_element(By.XPATH,
            "//*[contains(text(),'Корневая резолюция')]")
        if title.is_displayed():
            log.warning("Диалог не закрылся — пробую крестик")
            close_open_modals(driver)
            time.sleep(1)
    except Exception:
        pass
    return True


def close_card_after_resolution(driver):
    """После выдачи резолюции карточка может остаться открытой — закрываем."""
    time.sleep(2)
    try:
        close_btn = driver.find_element(By.ID, "header-close-btn")
        if close_btn.is_displayed():
            ActionChains(driver).move_to_element(close_btn).pause(0.3).click().perform()
            time.sleep(2)
            log.info("Карточка закрыта")
            return
    except Exception:
        pass
    log.info("Карточка уже закрыта или header-close-btn не найден")


# ============================================================
# DOCUMENT FLOW (один документ)
# ============================================================

def process_one(driver, doc, index, total):
    """Обработка одной строки реестра: найти, открыть, выдать резолюцию, закрыть."""
    log.info("=" * 50)
    log.info(f"ДОКУМЕНТ {index}/{total}: link={doc.get('link')!r}, "
             f"appeal_no={doc.get('appeal_no')}, "
             f"исполнитель={doc.get('executor')}")

    if not doc.get('executor'):
        log.warning(f"Row {doc['row_idx']}: исполнитель не определён → пропускаю")
        return False

    # 1. Найти строку в списке
    row = find_doc_row(driver, doc, timeout=10)
    if not row:
        log.warning(f"Row {doc['row_idx']}: документ в списке не найден → пропускаю")
        return False

    # 2. Открыть карточку
    if not open_doc_card(driver, row):
        return False

    # 3. Создать резолюцию
    try:
        btn = driver.find_element(By.ID, "header-action-btn-add_resolution")
        click(driver, btn, "Создать резолюцию")
        time.sleep(2)
    except Exception as e:
        log.error(f"Кнопка 'Создать резолюцию' не найдена: {e}")
        return False

    # 4. Содержание
    select_content_template(driver,
        settings.get("resolution_content", cfg.DEFAULTS["resolution_content"]))

    # 5. Тоггл "Требуется отчёт"
    if settings.get("require_report", True):
        toggle_switch(driver, "Требуется отчёт", "true")

    # 6. Тоггл "Контрольная резолюция"
    if settings.get("control_resolution", True):
        toggle_switch(driver, "Контрольная резолюция", "true")
        time.sleep(0.5)
        # 7. Дата контрольного этапа
        set_stage_date(driver,
            settings.get("workdays", cfg.DEFAULTS["workdays"]))

    # 8. Исполнитель
    if not fill_executor(driver, doc['executor']):
        log.error(f"Row {doc['row_idx']}: не получилось ввести исполнителя")
        # Закроем диалог чтобы продолжить — но НЕ сохраним
        close_open_modals(driver)
        return False

    # 9. Клик "Добавить"
    add_btn = _wait_button_enabled(driver, "add_btn", timeout=10)
    if not add_btn:
        log.error("Кнопка 'Добавить' не активировалась")
        close_open_modals(driver)
        return False
    click(driver, add_btn, "Добавить")
    time.sleep(1.5)

    # 10. Клик "Сохранить и отправить"
    if not submit_resolution(driver):
        log.error("Не удалось сохранить и отправить")
        return False

    # 11. Закрыть карточку (если осталась открыта)
    close_card_after_resolution(driver)
    log.info(f"Документ {index}/{total} ОБРАБОТАН")
    return True


# ============================================================
# MAIN
# ============================================================

def _choose_xlsx(base_dir):
    files = [f for f in os.listdir(base_dir) if f.lower().endswith('.xlsx')]
    if not files:
        log.error(f"Нет .xlsx в {base_dir}")
        input("Enter...")
        sys.exit(1)
    if len(files) == 1:
        log.info(f"Файл: {files[0]}")
        return os.path.join(base_dir, files[0])
    print(f"\nНайдено {len(files)} xlsx-файлов:")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    choice = input("Выбери номер: ").strip()
    try:
        return os.path.join(base_dir, files[int(choice) - 1])
    except (ValueError, IndexError):
        log.error("Неверный выбор")
        sys.exit(1)


def _is_port_open(host, port, timeout=0.5):
    import socket
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except (socket.error, OSError):
        return False


def _start_browser(base_dir):
    driver_path = os.path.join(base_dir, "msedgedriver.exe")
    if not os.path.exists(driver_path):
        log.error(f"msedgedriver.exe не найден в {base_dir}")
        input("Enter...")
        sys.exit(1)

    service = EdgeService(executable_path=driver_path)
    debugger_port = settings.get("debugger_port")
    if debugger_port and _is_port_open("127.0.0.1", int(debugger_port)):
        try:
            options = EdgeOptions()
            options.add_experimental_option("debuggerAddress",
                f"127.0.0.1:{debugger_port}")
            driver = webdriver.Edge(service=service, options=options)
            log.info(f"Подключился к открытому Edge на :{debugger_port} "
                     f"(URL: {driver.current_url or '?'})")
            return driver, True
        except Exception as e:
            log.warning(f"Не удалось подключиться: {e}")
    elif debugger_port:
        log.info(f"Edge не запущен с debug-портом {debugger_port}. Стартую новый.")

    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--auth-server-whitelist=*.interrao.ru")
    options.add_argument("--auth-negotiate-delegate-whitelist=*.interrao.ru")
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return webdriver.Edge(service=service, options=options), False


def _go_to_asud(driver, url, attached):
    try:
        cur = (driver.current_url or "").lower()
    except Exception:
        cur = ""
    if attached and "asudik" in cur:
        log.info(f"Использую вкладку АСУД: {driver.current_url}")
    else:
        log.info(f"Открываю {url}")
        driver.get(url)
    wait_asud_loaded(driver)


def _maybe_quit(driver, attached):
    if attached:
        log.info("Оставляю Edge открытым")
    else:
        try:
            driver.quit()
            log.info("Браузер закрыт")
        except Exception as e:
            log.warning(f"Ошибка закрытия: {e}")


def main():
    global settings
    settings = cfg.load()

    log.info("=" * 50)
    log.info("АСУД ИК — выдача резолюций (под Халецкой)")
    log.info("=" * 50)

    base_dir = cfg.get_base_dir()
    excel_path = _choose_xlsx(base_dir)

    docs = load_excel(excel_path)
    if not docs:
        log.error("Реестр пуст")
        input("Enter...")
        sys.exit(1)

    # Превью
    print(f"\nПервые 5:")
    for i, d in enumerate(docs[:5], 1):
        flag = '✓' if d['executor'] else '!'
        print(f"  {i}. [{flag}] {d.get('appeal_no') or '???':>10} | "
              f"{(d.get('executor') or 'ИСПОЛНИТЕЛЬ?')[:30]:30} | "
              f"{d.get('subject', '')[:40]}")
    no_executor = sum(1 for d in docs if not d['executor'])
    print(f"\nВсего: {len(docs)} (без исполнителя: {no_executor} — будут пропущены)")

    if input("Начать? (да/нет): ").strip().lower() not in ("да", "д", "y", "yes", ""):
        print("Отменено.")
        sys.exit(0)

    driver, attached = _start_browser(base_dir)
    try:
        url = settings.get("asud_url", cfg.DEFAULTS["asud_url"])
        _go_to_asud(driver, url, attached)

        # Переключение учётки
        if not switch_account(driver,
                settings.get("target_account", cfg.DEFAULTS["target_account"])):
            log.error("Не удалось переключиться на учётку. Прерываю.")
            input("Enter...")
            return

        # Сайдбар → "На резолюцию"
        click_sidebar_section(driver,
            settings.get("sidebar_section", cfg.DEFAULTS["sidebar_section"]))
        time.sleep(2)

        done, err, skip = 0, 0, 0
        for i, doc in enumerate(docs, 1):
            try:
                ok = process_one(driver, doc, i, len(docs))
                if ok:
                    done += 1
                else:
                    skip += 1
            except Exception as e:
                log.error(f"ОШИБКА документ {i}: {e}")
                err += 1
                # Возврат в список — обновим страницу
                try:
                    driver.get(url)
                    wait_asud_loaded(driver)
                    switch_account(driver,
                        settings.get("target_account", cfg.DEFAULTS["target_account"]))
                    click_sidebar_section(driver,
                        settings.get("sidebar_section", cfg.DEFAULTS["sidebar_section"]))
                    time.sleep(2)
                except Exception:
                    pass

        elapsed_seconds = time.monotonic() - start_time
        elapsed = timedelta(seconds=int(elapsed_seconds))
        avg = (timedelta(seconds=int(elapsed_seconds / done))
               if done else None)
        summary = [
            "",
            "=" * 60,
            "ГОТОВО!",
            f"  Обработано:  {done} / {len(docs)}",
            f"  Пропущено:   {skip}",
            f"  Ошибок:      {err}",
            f"  Затрачено:   {elapsed}" + (f"  (в среднем {avg}/док)" if avg else ""),
            "=" * 60,
        ]
        for line in summary:
            log.info(line)
            print(line)
        input("\nEnter для закрытия...")
    except Exception as e:
        log.error(f"Ошибка: {e}")
        input("Enter...")
    finally:
        _maybe_quit(driver, attached)


if __name__ == "__main__":
    main()
