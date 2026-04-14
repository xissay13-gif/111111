"""
Скрипт для пакетного создания Входящих документов в АСУД ИК
из Excel-файла (колонка B — краткое содержание, колонка C — корреспондент).

Установка:
    pip install selenium openpyxl pyinstaller

Сборка exe:
    pyinstaller --onefile --name asud_create_doc asud_create_doc.py

Положи msedgedriver.exe рядом с exe/скриптом.
"""

import time
import sys
import os
from datetime import date
import openpyxl
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


# ===== НАСТРОЙКИ =====
ASUD_URL = "https://asud.interrao.ru/asudik/"
TIMEOUT = 20
AUTO_REGISTER = False  # True — регистрирует автоматически, False — останавливается перед регистрацией


def get_driver_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    driver_path = os.path.join(base_dir, "msedgedriver.exe")

    if not os.path.exists(driver_path):
        print(f"!! msedgedriver.exe не найден в папке: {base_dir}")
        print("   Положи msedgedriver.exe рядом с этим файлом.")
        input("Enter для выхода...")
        sys.exit(1)

    return driver_path


def load_excel(file_path):
    """Читает Excel и возвращает список словарей {содержание, корреспондент}."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        content = row[1]  # колонка B
        correspondent = row[2]  # колонка C
        if content and correspondent:
            rows.append({
                "содержание": str(content).strip(),
                "корреспондент": str(correspondent).strip(),
            })
    wb.close()
    return rows


def js_click(driver, element, description=""):
    """Надёжный клик с fallback'ами для GWT/GXT-элементов."""
    # Прокручиваем к элементу
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});", element)
        time.sleep(0.2)
    except Exception:
        pass

    # Способ 1: ActionChains — имитирует реальную мышь с hover (лучше всего для GXT dropdown)
    try:
        from selenium.webdriver.common.action_chains import ActionChains as AC
        AC(driver).move_to_element(element).pause(0.2).click().perform()
        print(f"  ОК клик (mouse): {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass

    # Способ 2: обычный Selenium click
    try:
        element.click()
        print(f"  ОК клик: {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass

    # Способ 3: JS .click()
    try:
        driver.execute_script("arguments[0].click();", element)
        print(f"  ОК JS-клик: {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass

    # Способ 4: полный набор mouse-событий (для самых упрямых GWT)
    try:
        driver.execute_script("""
            var el = arguments[0];
            var events = ['mouseover', 'mousedown', 'mouseup', 'click'];
            events.forEach(function(type) {
                var evt = new MouseEvent(type, {bubbles: true, cancelable: true, view: window});
                el.dispatchEvent(evt);
            });
        """, element)
        print(f"  ОК mouse-events: {description}")
        time.sleep(0.5)
    except Exception as e:
        print(f"  !! Клик не удался: {description}: {e}")


def wait_asud_loaded(driver, max_wait=120):
    """Адаптивное ожидание полной загрузки АСУД.
    1) document.readyState === 'complete'
    2) Кнопка создания документа кликабельна
    3) В таблице документов появились строки (данные прогрузились)
    """
    print("  Жду готовности страницы...")
    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: d.execute_script("return document.readyState === 'complete'")
        )
    except Exception:
        print("  ! readyState не complete за отведённое время")

    print("  Жду кнопку создания документа...")
    try:
        WebDriverWait(driver, max_wait).until(
            EC.element_to_be_clickable((By.ID, "mainscreen-create-button"))
        )
    except Exception:
        print("  ! Кнопка создания не появилась")

    print("  Жду загрузку данных в таблице...")
    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR,
                "tr[class*='GridView-row'], tr[class*='grid-row'], "
                "tr[class*='OSHSGridStyle-row'], tr[class*='obj-list-rec']")) > 0
        )
    except Exception:
        print("  ! Данные в таблице не появились — продолжаем")

    # Дополнительная пауза чтобы GWT дорисовал UI
    time.sleep(5)
    print("  ОК АСУД загружен")


def wait_and_click(driver, by, selector, description="", timeout=TIMEOUT):
    print(f"  -> Ожидаю: {description or selector}")
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )
    time.sleep(0.5)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    print(f"  ОК Клик: {description or selector}")
    time.sleep(0.5)
    return el


def fio_to_initials(full_name):
    """Преобразует 'Калганова Тамара Алексеевна' -> 'Калганова Т А'
    для сопоставления с тем, что показывает АСУД."""
    parts = full_name.strip().split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]} {parts[2][0]}"
    elif len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}"
    return parts[0] if parts else full_name


def match_correspondent(text, full_name):
    """Проверяет совпадение: текст из АСУД (инициалы) vs полное ФИО из Excel."""
    text_clean = text.strip()
    # Прямое совпадение по полному ФИО
    if full_name in text_clean:
        return True
    # Совпадение по инициалам: "Калганова Т А" или "Калганова Т.А." и т.п.
    initials = fio_to_initials(full_name)
    # Убираем точки для сравнения
    text_norm = text_clean.replace('.', '').replace(',', '')
    initials_norm = initials.replace('.', '').replace(',', '')
    if initials_norm.lower() in text_norm.lower():
        return True
    # Совпадение только по фамилии (фоллбэк)
    surname = full_name.split()[0]
    if text_clean.lower().startswith(surname.lower()):
        return True
    return False


def find_input_near_label(driver, label_text):
    """Находит input combobox в той же tr/контейнере что и лейбл."""
    labels = driver.find_elements(By.XPATH,
        f"//*[normalize-space(text())='{label_text}']")
    for label in labels:
        try:
            if not label.is_displayed():
                continue
            # Поднимаемся вверх ища контейнер с input
            for level in range(1, 6):
                parent = label
                for _ in range(level):
                    parent = parent.find_element(By.XPATH, "..")
                inputs = parent.find_elements(By.CSS_SELECTOR,
                    "input[id*='select_combobox-input'], input[type='text']")
                visible = [i for i in inputs
                           if i.is_displayed() and i.get_attribute("readonly") is None]
                if visible:
                    return visible[0]
        except Exception:
            continue
    return None


def fill_correspondent(driver, person_name):
    """Заполняет поле Корреспондент через combobox."""
    print(f"  Корреспондент: {person_name}")
    time.sleep(1)

    # Ищем поле строго рядом с лейблом "Корреспондент" (не "Номер у корреспондента"!)
    corr_input = find_input_near_label(driver, "Корреспондент")

    if not corr_input:
        print("  !! Поле корреспондента не найдено!")
        return

    # Вводим фамилию для поиска
    surname = person_name.split()[0]
    initials = fio_to_initials(person_name)

    corr_input.click()
    time.sleep(0.3)
    corr_input.clear()
    time.sleep(0.3)
    for char in surname:
        corr_input.send_keys(char)
        time.sleep(0.1)
    print(f"  ОК Введена фамилия: {surname}")
    time.sleep(2)

    # Кликаем через ActionChains — так работало в Служебной записке
    def find_all_with_surname():
        results = driver.find_elements(By.XPATH,
            f"//*[contains(text(),'{surname}')]")
        found = []
        for r in results:
            try:
                if not r.is_displayed() or r == corr_input:
                    continue
                if r.tag_name.lower() == 'input':
                    continue
                found.append(r)
            except Exception:
                continue
        return found

    all_results = find_all_with_surname()
    if not all_results:
        corr_input.send_keys(Keys.ENTER)
        time.sleep(2)
        all_results = find_all_with_surname()

    print(f"  Найдено кандидатов с фамилией '{surname}': {len(all_results)}")

    # Ищем первое совпадение по инициалам
    target = None
    for r in all_results:
        try:
            if match_correspondent(r.text, person_name):
                target = r
                print(f"  Совпадение по инициалам: {r.text.strip()[:80]}")
                break
        except Exception:
            continue

    if not target and all_results:
        target = all_results[0]
        print(f"  ! По инициалам '{initials}' не нашли, беру первого: {target.text.strip()[:80]}")

    if not target:
        print(f"  !! Корреспондент не найден: {person_name}")
        return

    # Клик через ActionChains (как в Служебной записке — это работало)
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", target)
        time.sleep(0.3)
        ActionChains(driver).move_to_element(target).pause(0.3).click().perform()
        time.sleep(1)
        print(f"  ОК Корреспондент выбран: {person_name}")
    except Exception as e:
        print(f"  !! Ошибка выбора корреспондента: {e}")


def fill_corr_number(driver):
    """Заполняет поле 'Номер у корреспондента' значением 'б/н'."""
    print("  Номер у корреспондента: б/н")

    inp = find_input_near_label(driver, "Номер у корреспондента")

    if not inp:
        print("  !! Поле 'Номер у корреспондента' не найдено")
        return

    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", inp)
        time.sleep(0.3)
        inp.click()
        time.sleep(0.3)
        inp.clear()
        time.sleep(0.2)
        inp.send_keys("б/н")
        time.sleep(0.3)
        inp.send_keys(Keys.TAB)
        print("  ОК Номер заполнен")
    except Exception as e:
        # Fallback через JS
        try:
            driver.execute_script("""
                arguments[0].value = 'б/н';
                arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
                arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
                arguments[0].dispatchEvent(new Event('blur', {bubbles: true}));
            """, inp)
            print("  ОК Номер заполнен через JS")
        except Exception as e2:
            print(f"  !! Ошибка заполнения номера: {e2}")


def fill_corr_date(driver):
    """Заполняет поле 'Дата у корреспондента' сегодняшней датой.
    Использует точный поиск по лейблу, чтобы не попасть в 'Дата помещения в архив'
    или 'Дата регистрации'."""
    today = date.today().strftime("%d.%m.%Y")
    print(f"  Дата у корреспондента: {today}")

    # Ищем лейбл с ТОЧНЫМ совпадением 'Дата у корреспондента'
    labels = driver.find_elements(By.XPATH,
        "//*[normalize-space(text())='Дата у корреспондента']")

    inp = None
    for label in labels:
        try:
            if not label.is_displayed():
                continue
            # Поднимаемся вверх, ищем input в том же контейнере
            for level in range(1, 6):
                parent = label
                for _ in range(level):
                    parent = parent.find_element(By.XPATH, "..")
                inputs = parent.find_elements(By.CSS_SELECTOR, "input[type='text']")
                visible = [i for i in inputs
                           if i.is_displayed() and i.get_attribute("readonly") is None]
                if visible:
                    inp = visible[0]
                    break
            if inp:
                break
        except Exception:
            continue

    if not inp:
        print("  !! Поле 'Дата у корреспондента' не найдено")
        return

    try:
        # Прокручиваем к элементу чтобы он был в центре (и триггер не перекрывал)
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});", inp)
        time.sleep(0.5)

        # Клик через JS — обходит перекрытие другими элементами
        driver.execute_script("arguments[0].focus(); arguments[0].click();", inp)
        time.sleep(0.3)

        # Очистка: выделяем всё и удаляем
        inp.send_keys(Keys.CONTROL + "a")
        time.sleep(0.2)
        inp.send_keys(Keys.DELETE)
        time.sleep(0.2)

        # Вводим дату
        inp.send_keys(today)
        time.sleep(0.3)
        inp.send_keys(Keys.TAB)
        print(f"  ОК Дата заполнена: {today}")
    except Exception as e:
        # Финальный fallback: напрямую через JS set value + trigger change
        try:
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
                arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
                arguments[0].dispatchEvent(new Event('blur', {bubbles: true}));
            """, inp, today)
            print(f"  ОК Дата заполнена через JS: {today}")
        except Exception as e2:
            print(f"  !! Ошибка заполнения даты: {e2}")


def fill_delivery_method(driver):
    """Выбирает 'Электронная почта' в поле 'Способ получения'.
    Это triggerfield: клик открывает контекстное меню со списком опций."""
    print("  Способ получения: Электронная почта")

    # 1. Находим поле по ТОЧНОМУ лейблу
    trigger = find_input_near_label(driver, "Способ получения")

    # Запасной вариант: если нет input, ищем по лейблу и берём кликабельный div/trigger
    if not trigger:
        labels = driver.find_elements(By.XPATH,
            "//*[normalize-space(text())='Способ получения']")
        for label in labels:
            try:
                if not label.is_displayed():
                    continue
                for level in range(1, 6):
                    parent = label
                    for _ in range(level):
                        parent = parent.find_element(By.XPATH, "..")
                    # Ищем триггер — div с классом trigger или img
                    for sel in ["div[class*='trigger']", "img[class*='trigger']",
                                "[class*='ComboBox']", "[class*='combobox']"]:
                        try:
                            el = parent.find_element(By.CSS_SELECTOR, sel)
                            if el.is_displayed():
                                trigger = el
                                break
                        except Exception:
                            continue
                    if trigger:
                        break
                if trigger:
                    break
            except Exception:
                continue

    if not trigger:
        print("  !! Поле 'Способ получения' не найдено")
        return

    # 2. Прокручиваем и кликаем чтобы открыть dropdown
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", trigger)
        time.sleep(0.3)
    except Exception:
        pass

    js_click(driver, trigger, "Открыть список 'Способ получения'")
    time.sleep(1.5)  # ждём появления dropdown

    # 3. Ищем "Электронная почта" в появившемся меню
    #    Варианты: div в dropdown, li, td, span
    target_text = "Электронная почта"
    option = None

    # Сначала пробуем точное совпадение текста
    candidates = driver.find_elements(By.XPATH,
        f"//*[normalize-space(text())='{target_text}']")
    for c in candidates:
        try:
            if c.is_displayed():
                option = c
                break
        except Exception:
            continue

    # Если точного нет — частичное
    if not option:
        candidates = driver.find_elements(By.XPATH,
            f"//*[contains(text(),'{target_text}')]")
        for c in candidates:
            try:
                if c.is_displayed() and c.tag_name.lower() != 'input':
                    option = c
                    break
            except Exception:
                continue

    if option:
        # Прокручиваем к опции (выпадающий список может быть длинным)
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", option)
            time.sleep(0.3)
        except Exception:
            pass
        js_click(driver, option, target_text)
        time.sleep(0.5)
        print("  ОК Способ получения выбран: Электронная почта")
        return

    # Фоллбэк: <select>
    try:
        selects = driver.find_elements(By.TAG_NAME, "select")
        for sel in selects:
            if not sel.is_displayed():
                continue
            for opt in sel.find_elements(By.TAG_NAME, "option"):
                if target_text in opt.text:
                    opt.click()
                    print("  ОК Способ получения выбран (select)")
                    return
    except Exception:
        pass

    print("  !! 'Электронная почта' не найдена в списке")


def get_attachment_path():
    """Ищет .msg файл рядом с exe/скриптом."""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    msg_files = [f for f in os.listdir(base_dir) if f.lower().endswith('.msg')]
    if not msg_files:
        return None
    if len(msg_files) == 1:
        return os.path.join(base_dir, msg_files[0])
    # Если несколько — берём первый, но сообщаем
    print(f"  ! Найдено {len(msg_files)} .msg файлов, берём: {msg_files[0]}")
    return os.path.join(base_dir, msg_files[0])


def attach_content(driver, file_path):
    """Нажимает 'Присоединить содержимое' и загружает файл."""
    print(f"  Присоединение файла: {os.path.basename(file_path)}")

    # Кликаем кнопку "Присоединить содержимое"
    try:
        btn = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(text(),'Присоединить содержимое')]"))
        )
        js_click(driver, btn, "Prisoedinit soderzhimoe")
        time.sleep(2)
    except Exception as e:
        print(f"  !! Кнопка 'Присоединить содержимое' не найдена: {e}")
        return

    # Ищем скрытый input[type='file'] и отправляем путь к файлу
    try:
        file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
        file_input.send_keys(file_path)
        print(f"  ОК Файл выбран: {os.path.basename(file_path)}")
        time.sleep(2)
    except Exception:
        # Фоллбэк: ищем все input[type='file'], включая скрытые
        try:
            file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
            if file_inputs:
                # Делаем input видимым через JS и отправляем файл
                driver.execute_script(
                    "arguments[0].style.display='block'; arguments[0].style.visibility='visible';",
                    file_inputs[0])
                time.sleep(0.5)
                file_inputs[0].send_keys(file_path)
                print(f"  ОК Файл выбран (fallback): {os.path.basename(file_path)}")
                time.sleep(2)
            else:
                print("  !! input[type='file'] не найден на странице!")
                return
        except Exception as e:
            print(f"  !! Ошибка загрузки файла: {e}")
            return

    # Нажимаем кнопку подтверждения в диалоге
    try:
        confirm_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "#SetContentDialogBtnSend, [id*='SetContentDialogBtnSend']"))
        )
        js_click(driver, confirm_btn, "Podtverdit prisoedinenie")
        time.sleep(3)
        print("  ОК Файл присоединён!")
    except Exception:
        # Фоллбэк по тексту кнопки
        try:
            btns = driver.find_elements(By.XPATH,
                "//div[contains(text(),'Присоединить содержимое')] | //button[contains(text(),'Присоединить')]")
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                js_click(driver, visible[-1], "Podtverdit (fallback)")
                time.sleep(3)
                print("  ОК Файл присоединён!")
        except Exception as e:
            print(f"  !! Ошибка подтверждения: {e}")


def add_addressee(driver, person_name):
    """Добавляет адресата через combobox рядом с разделом 'Адресаты'.
    Использует ActionChains как в рабочей версии Служебной записки."""
    print(f"  Адресат: {person_name}")
    try:
        addr_input = find_input_near_label(driver, "Адресаты")

        if not addr_input:
            print("  !! Поле адресата не найдено")
            return

        surname = person_name.split()[0]
        initials = fio_to_initials(person_name)

        addr_input.click()
        time.sleep(0.5)
        addr_input.clear()
        time.sleep(0.3)
        for char in surname:
            addr_input.send_keys(char)
            time.sleep(0.1)
        print(f"  ОК Введена фамилия: {surname}")
        time.sleep(2)

        def find_all_with_surname():
            results = driver.find_elements(By.XPATH,
                f"//*[contains(text(),'{surname}')]")
            found = []
            for r in results:
                try:
                    if not r.is_displayed() or r == addr_input:
                        continue
                    if r.tag_name.lower() == 'input':
                        continue
                    found.append(r)
                except Exception:
                    continue
            return found

        all_results = find_all_with_surname()
        if not all_results:
            addr_input.send_keys(Keys.ENTER)
            time.sleep(2)
            all_results = find_all_with_surname()

        print(f"  Найдено кандидатов с фамилией '{surname}': {len(all_results)}")

        target = None
        for r in all_results:
            try:
                if match_correspondent(r.text, person_name):
                    target = r
                    print(f"  Совпадение по инициалам: {r.text.strip()[:80]}")
                    break
            except Exception:
                continue

        if not target and all_results:
            target = all_results[0]
            print(f"  ! По инициалам '{initials}' не нашли, беру первого: {target.text.strip()[:80]}")

        if not target:
            print(f"  !! Адресат не найден в списке: {person_name}")
            return

        # ActionChains click (как в Служебной записке)
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", target)
        time.sleep(0.3)
        ActionChains(driver).move_to_element(target).pause(0.3).click().perform()
        time.sleep(1)
        print(f"  ОК Адресат добавлен: {person_name}")
    except Exception as e:
        print(f"  !! Ошибка добавления адресата: {e}")


def go_to_distribution_tab(driver):
    """Переходит на вкладку 'Рассылка'."""
    print("  Переход на вкладку Рассылка...")
    try:
        tab = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(text(),'Рассылка')] | //span[contains(text(),'Рассылка')]"))
        )
        js_click(driver, tab, "Vkladka Rassylka")
        time.sleep(2)
        print("  ОК Вкладка Рассылка открыта")
    except Exception as e:
        print(f"  !! Вкладка Рассылка не найдена: {e}")


def add_distribution_addressee(driver, person_name):
    """Добавляет адресата на вкладке 'Рассылка' через поле 'Добавить адресатов'."""
    print(f"  Рассылка адресат: {person_name}")
    try:
        # Ищем поле "Добавить адресатов" — combobox внизу
        corr_input = None

        # По id combobox
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[id*='select_combobox-input']")
        visible = [i for i in inputs if i.is_displayed()]
        if visible:
            corr_input = visible[-1]  # берём последний видимый (внизу страницы)

        if not corr_input:
            # По лейблу
            try:
                label = driver.find_element(By.XPATH,
                    "//*[contains(text(),'Добавить адресатов')]")
                parent = label.find_element(By.XPATH, "./ancestor::div[contains(@class,'field')] | ./ancestor::tr")
                inp = parent.find_element(By.CSS_SELECTOR, "input[type='text']")
                if inp.is_displayed():
                    corr_input = inp
            except Exception:
                pass

        if not corr_input:
            print("  !! Поле 'Добавить адресатов' не найдено!")
            return

        # Вводим фамилию
        surname = person_name.split()[0]
        corr_input.click()
        time.sleep(0.3)
        corr_input.clear()
        corr_input.send_keys(surname)
        print(f"  ОК Введена фамилия: {surname}")
        time.sleep(2)

        # Ищем в выпадающем списке
        results = driver.find_elements(By.XPATH,
            f"//*[contains(text(),'{surname}')]")
        visible_results = [r for r in results if r.is_displayed() and r != corr_input]
        if visible_results:
            for r in visible_results:
                if match_correspondent(r.text, person_name):
                    js_click(driver, r, f"Vybor: {r.text.strip()}")
                    time.sleep(1)
                    print(f"  ОК Адресат добавлен: {person_name}")
                    return
            js_click(driver, visible_results[0], f"Vybor (pervyj): {visible_results[0].text.strip()}")
            time.sleep(1)
            print(f"  ОК Адресат добавлен: {person_name}")
        else:
            print(f"  !! Адресат не найден в списке: {person_name}")
    except Exception as e:
        print(f"  !! Ошибка: {e}")


def create_one_document(driver, doc_data, index, total):
    """Создаёт один входящий документ."""
    print(f"\n{'='*60}")
    print(f"ДОКУМЕНТ {index}/{total}")
    print(f"  Содержание: {doc_data['содержание'][:80]}...")
    print(f"  Корреспондент: {doc_data['корреспондент']}")
    print(f"{'='*60}")

    # ШАГ 1: Кнопка создания документа
    print("\n[1/5] Кнопка создания документа...")
    el = WebDriverWait(driver, TIMEOUT).until(
        EC.presence_of_element_located((By.ID, "mainscreen-create-button"))
    )
    time.sleep(1)
    js_click(driver, el, "Кнопка создания документа")
    time.sleep(3)

    # ШАГ 2: Тип — Входящий документ
    print("\n[2/5] Тип: Входящий документ...")
    wait_and_click(driver, By.XPATH,
        "//div[contains(text(),'Входящий документ')]",
        "Входящий документ")
    time.sleep(1)

    # ШАГ 3: Вид — Письма, заявления и жалобы граждан, акционеров
    print("\n[3/5] Вид: Письма, заявления...")
    wait_and_click(driver, By.XPATH,
        "//div[contains(text(),'Письма, заявления и жалобы граждан')] | //td[contains(text(),'Письма, заявления и жалобы граждан')]",
        "Письма, заявления и жалобы")
    time.sleep(0.5)

    # Кнопка "Создать документ"
    print("  Создать документ...")
    wait_and_click(driver, By.XPATH,
        "//button[contains(text(),'Создать документ')] | //div[contains(text(),'Создать документ')]",
        "Создать документ")
    print("  Жду загрузку формы (5 сек)...")
    time.sleep(5)

    # ШАГ 4: Заполнение формы
    print("\n[4/5] Заполняю форму...")

    # --- Краткое содержание ---
    print("\n  Краткое содержание:")
    try:
        textareas = driver.find_elements(By.TAG_NAME, "textarea")
        visible_ta = [ta for ta in textareas if ta.is_displayed()]
        if visible_ta:
            visible_ta[0].click()
            time.sleep(0.3)
            visible_ta[0].clear()
            visible_ta[0].send_keys(doc_data["содержание"])
            print(f"  ОК Заполнено")
        else:
            print("  !! Textarea не найдена")
    except Exception as e:
        print(f"  !! Ошибка: {e}")
    time.sleep(0.5)

    # --- Корреспондент ---
    print("\n  Корреспондент:")
    fill_correspondent(driver, doc_data["корреспондент"])
    time.sleep(1)

    # --- Номер у корреспондента ---
    print("\n  Номер:")
    fill_corr_number(driver)
    time.sleep(0.5)

    # --- Дата у корреспондента ---
    print("\n  Дата:")
    fill_corr_date(driver)
    time.sleep(0.5)

    # --- Адресат (Басманов) ---
    print("\n  Адресат:")
    add_addressee(driver, "Басманов Александр Владимирович")
    time.sleep(0.5)

    # --- Способ получения ---
    print("\n  Способ получения:")
    fill_delivery_method(driver)
    time.sleep(0.5)

    # ШАГ 5: Сохранение
    print("\n[5/8] Сохранение...")
    try:
        # Сначала пытаемся по id="header-save-btn"
        save_btn = None
        try:
            save_btn = WebDriverWait(driver, TIMEOUT).until(
                EC.element_to_be_clickable((By.ID, "header-save-btn"))
            )
        except Exception:
            # Фоллбэк: кнопка Сохранить в шапке по тексту
            btns = driver.find_elements(By.XPATH,
                "//*[normalize-space(text())='Сохранить']")
            for b in btns:
                if b.is_displayed():
                    save_btn = b
                    break

        if save_btn:
            js_click(driver, save_btn, "Сохранить")
            time.sleep(3)
            print(f"  ОК Документ {index}/{total} сохранён!")
        else:
            print("  !! Кнопка 'Сохранить' не найдена")
    except Exception as e:
        print(f"  !! Ошибка сохранения: {e}")

    # ШАГ 6: Присоединить содержимое
    if doc_data.get("файл"):
        print("\n[6/8] Присоединение содержимого...")
        attach_content(driver, doc_data["файл"])

    # ШАГ 7: Вкладка "Рассылка" — добавить Халецкую
    print("\n[7/9] Рассылка — добавить Халецкую...")
    go_to_distribution_tab(driver)
    add_distribution_addressee(driver, "Халецкая Юлия Владимировна")
    time.sleep(1)

    # ШАГ 8: Сохранить после рассылки
    print("\n[8/9] Сохранение после рассылки...")
    try:
        save_btn = None
        try:
            save_btn = WebDriverWait(driver, TIMEOUT).until(
                EC.element_to_be_clickable((By.ID, "header-save-btn"))
            )
        except Exception:
            btns = driver.find_elements(By.XPATH,
                "//*[normalize-space(text())='Сохранить']")
            for b in btns:
                if b.is_displayed():
                    save_btn = b
                    break
        if save_btn:
            js_click(driver, save_btn, "Сохранить")
            time.sleep(3)
    except Exception:
        pass

    # ШАГ 9: Зарегистрировать
    if AUTO_REGISTER:
        print("\n[9/9] Регистрация...")
        try:
            reg_btn = WebDriverWait(driver, TIMEOUT).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "#header-action-btn-register, [id*='header-action-btn-register']"))
            )
            js_click(driver, reg_btn, "Зарегистрировать")
            time.sleep(3)
            print(f"  ОК Документ {index}/{total} ЗАРЕГИСТРИРОВАН!")
        except Exception:
            try:
                btn = driver.find_element(By.XPATH,
                    "//div[contains(text(),'Зарегистрировать')]")
                js_click(driver, btn, "Зарегистрировать (fallback)")
                time.sleep(3)
                print(f"  ОК Документ {index}/{total} ЗАРЕГИСТРИРОВАН!")
            except Exception as e:
                print(f"  !! Кнопка 'Зарегистрировать' не найдена: {e}")
    else:
        print(f"\n[9/9] Документ {index}/{total} готов (без регистрации)")

    # Возвращаемся на главную для следующего документа
    time.sleep(2)
    driver.get(ASUD_URL)
    wait_asud_loaded(driver)


def main():
    print("=" * 60)
    print("АСУД ИК - Пакетное создание Входящих документов")
    print("=" * 60)

    # --- Поиск Excel-файла рядом с exe/скриптом ---
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    xlsx_files = [f for f in os.listdir(base_dir) if f.lower().endswith('.xlsx')]

    if not xlsx_files:
        print(f"!! Не найден .xlsx файл в папке: {base_dir}")
        print("   Положи Excel-файл рядом с exe.")
        input("Enter...")
        sys.exit(1)
    elif len(xlsx_files) == 1:
        excel_path = os.path.join(base_dir, xlsx_files[0])
        print(f"\nНайден файл: {xlsx_files[0]}")
    else:
        print(f"\nНайдено {len(xlsx_files)} xlsx-файлов:")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  {i}. {f}")
        choice = input("Выбери номер файла: ").strip()
        try:
            excel_path = os.path.join(base_dir, xlsx_files[int(choice) - 1])
        except (ValueError, IndexError):
            print("!! Неверный выбор")
            input("Enter...")
            sys.exit(1)

    # --- Поиск .msg файла ---
    msg_path = get_attachment_path()
    if msg_path:
        print(f"Файл для присоединения: {os.path.basename(msg_path)}")
    else:
        print("! .msg файл не найден — документы будут без вложения")

    # --- Загрузка данных ---
    print(f"\nЧтение файла: {excel_path}")
    docs = load_excel(excel_path)
    # Добавляем путь к файлу в каждый документ
    for doc in docs:
        doc["файл"] = msg_path
    print(f"Найдено документов: {len(docs)}")

    if not docs:
        print("!! Нет данных для создания!")
        input("Enter...")
        sys.exit(1)

    # Показать превью
    print("\nПервые 5 записей:")
    for i, d in enumerate(docs[:5], 1):
        print(f"  {i}. {d['корреспондент']} | {d['содержание'][:60]}...")

    print(f"\nВсего: {len(docs)} документов")
    confirm = input("Начать создание? (да/нет): ").strip().lower()
    if confirm not in ("da", "yes", "y", "д", "да"):
        print("Отменено.")
        sys.exit(0)

    # --- Запуск браузера ---
    driver_path = get_driver_path()
    print(f"\nEdgeDriver: {driver_path}")

    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--auth-server-whitelist=*.interrao.ru")
    options.add_argument("--auth-negotiate-delegate-whitelist=*.interrao.ru")
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    service = EdgeService(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)

    try:
        # Открываем АСУД
        print("\nОткрываю АСУД...")
        driver.get(ASUD_URL)
        wait_asud_loaded(driver)

        # Создаём документы в цикле
        for i, doc in enumerate(docs, 1):
            try:
                create_one_document(driver, doc, i, len(docs))
            except Exception as e:
                print(f"\n!! ОШИБКА при создании документа {i}: {e}")
                print("  Пробую следующий...")
                driver.get(ASUD_URL)
                wait_asud_loaded(driver)
                continue

        print(f"\n{'='*60}")
        print(f"ГОТОВО! Создано документов: {len(docs)}")
        print(f"{'='*60}")

        input("\nEnter для закрытия браузера...")

    except Exception as e:
        print(f"\n!! Ошибка: {e}")
        input("Enter для закрытия...")

    finally:
        driver.quit()
        print("\nОК Браузер закрыт.")


if __name__ == "__main__":
    main()
