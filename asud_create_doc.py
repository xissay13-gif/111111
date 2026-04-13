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
        print(f"!! msedgedriver.exe ne najden v papke: {base_dir}")
        print("   Polozhi msedgedriver.exe ryadom s etim fajlom.")
        input("Enter dlya vyhoda...")
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
    """Кликает через JavaScript — надёжнее для GWT-элементов."""
    driver.execute_script("arguments[0].click();", element)
    print(f"  OK JS-klik: {description}")
    time.sleep(0.5)


def wait_and_click(driver, by, selector, description="", timeout=TIMEOUT):
    print(f"  -> Ozhidayu: {description or selector}")
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )
    time.sleep(0.5)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    print(f"  OK Klik: {description or selector}")
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


def fill_correspondent(driver, person_name):
    """Заполняет поле Корреспондент через combobox."""
    print(f"  Korrespondent: {person_name}")
    time.sleep(1)

    # Ищем поле корреспондента по CSS-селектору combobox
    corr_input = None
    try:
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[id*='select_combobox-input']")
        visible = [i for i in inputs if i.is_displayed()]
        if visible:
            corr_input = visible[0]
    except Exception:
        pass

    if not corr_input:
        try:
            labels = driver.find_elements(By.XPATH,
                "//*[contains(text(),'Корреспондент')]")
            for label in labels:
                if label.is_displayed():
                    parent = label.find_element(By.XPATH, "./ancestor::tr")
                    inp = parent.find_element(By.CSS_SELECTOR, "input[type='text']")
                    if inp.is_displayed():
                        corr_input = inp
                        break
        except Exception:
            pass

    if not corr_input:
        print("  !! Pole korrespondenta ne najdeno!")
        return

    # Вводим фамилию для поиска
    surname = person_name.split()[0]
    corr_input.click()
    time.sleep(0.3)
    corr_input.clear()
    corr_input.send_keys(surname)
    print(f"  OK Vvedena familiya: {surname}")
    time.sleep(2)

    # Ждём выпадающий список и ищем совпадение по инициалам
    try:
        results = driver.find_elements(By.XPATH,
            f"//*[contains(text(),'{surname}')]")
        visible_results = [r for r in results if r.is_displayed() and r != corr_input]
        if visible_results:
            # Сначала ищем точное совпадение по инициалам
            for r in visible_results:
                if match_correspondent(r.text, person_name):
                    js_click(driver, r, f"Vybor: {r.text.strip()}")
                    time.sleep(1)
                    return
            # Если точного нет — берём первый с фамилией
            js_click(driver, visible_results[0], f"Vybor (pervyj): {visible_results[0].text.strip()}")
            time.sleep(1)
        else:
            corr_input.send_keys(Keys.ENTER)
            time.sleep(2)
            results = driver.find_elements(By.XPATH,
                f"//*[contains(text(),'{surname}')]")
            visible_results = [r for r in results if r.is_displayed() and r != corr_input]
            if visible_results:
                for r in visible_results:
                    if match_correspondent(r.text, person_name):
                        js_click(driver, r, f"Vybor: {r.text.strip()}")
                        time.sleep(1)
                        return
                js_click(driver, visible_results[0], f"Vybor (pervyj): {visible_results[0].text.strip()}")
                time.sleep(1)
            else:
                print(f"  !! Korrespondent ne najden: {person_name}")
    except Exception as e:
        print(f"  !! Oshibka vybora korrespondenta: {e}")


def fill_corr_number(driver):
    """Заполняет поле 'Номер у корреспондента' значением 'б/н'."""
    print("  Nomer u korrespondenta: b/n")
    try:
        # Ищем поле по лейблу
        label = driver.find_element(By.XPATH,
            "//*[contains(text(),'Номер у корреспондента')]")
        parent = label.find_element(By.XPATH, "./ancestor::tr | ./ancestor::div[contains(@class,'field')]")
        inp = parent.find_element(By.CSS_SELECTOR, "input[type='text']")
        if inp.is_displayed():
            inp.click()
            time.sleep(0.3)
            inp.clear()
            inp.send_keys("б/н")
            print("  OK Nomer zapolnen")
            return
    except Exception:
        pass

    # Фоллбэк: ищем все видимые input'ы и берём тот что рядом с "Номер у корреспондента"
    try:
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
        visible = [i for i in inputs if i.is_displayed()]
        for inp in visible:
            try:
                # Проверяем, есть ли рядом текст "Номер у корреспондента"
                parent_html = inp.find_element(By.XPATH, "./ancestor::tr").get_attribute("innerHTML")
                if "Номер у корреспондента" in parent_html:
                    inp.click()
                    time.sleep(0.3)
                    inp.clear()
                    inp.send_keys("б/н")
                    print("  OK Nomer zapolnen (fallback)")
                    return
            except Exception:
                continue
    except Exception:
        pass
    print("  !! Pole 'Nomer u korrespondenta' ne najdeno")


def fill_corr_date(driver):
    """Заполняет поле 'Дата у корреспондента' сегодняшней датой."""
    today = date.today().strftime("%d.%m.%Y")
    print(f"  Data u korrespondenta: {today}")
    try:
        # Ищем поле даты по лейблу
        label = driver.find_element(By.XPATH,
            "//*[contains(text(),'Дата у корреспондента')]")
        parent = label.find_element(By.XPATH, "./ancestor::tr | ./ancestor::div[contains(@class,'field')]")
        inp = parent.find_element(By.CSS_SELECTOR, "input[type='text']")
        if inp.is_displayed():
            inp.click()
            time.sleep(0.3)
            inp.clear()
            inp.send_keys(today)
            # Убираем фокус чтобы дата применилась
            inp.send_keys(Keys.TAB)
            print(f"  OK Data zapolnena: {today}")
            return
    except Exception:
        pass

    # Фоллбэк: ищем input с data-marker="date" или Css3DateCell в классе
    try:
        date_inputs = driver.find_elements(By.CSS_SELECTOR,
            "input[id*='x-auto'][class*='DateCell']")
        visible = [i for i in date_inputs if i.is_displayed()]
        if visible:
            # Берём последний видимый (дата корреспондента обычно ниже даты регистрации)
            inp = visible[-1]
            inp.click()
            time.sleep(0.3)
            inp.clear()
            inp.send_keys(today)
            inp.send_keys(Keys.TAB)
            print(f"  OK Data zapolnena (fallback): {today}")
            return
    except Exception:
        pass
    print("  !! Pole 'Data u korrespondenta' ne najdeno")


def fill_delivery_method(driver):
    """Выбирает 'Электронная почта' в поле 'Способ получения'."""
    print("  Sposob polucheniya: Elektronnaya pochta")
    try:
        # Ищем поле "Способ получения" — это выпадающий список (select или triggerfield)
        label = driver.find_element(By.XPATH,
            "//*[contains(text(),'Способ получения')]")
        # Кликаем по полю рядом с лейблом чтобы открыть выпадающий список
        parent = label.find_element(By.XPATH, "./ancestor::tr | ./ancestor::div[contains(@class,'field')]")
        # Ищем кликабельную область — input или div триггера
        clickable = None
        for sel in ["input[type='text']", "div[class*='trigger']", "img[class*='trigger']"]:
            try:
                el = parent.find_element(By.CSS_SELECTOR, sel)
                if el.is_displayed():
                    clickable = el
                    break
            except Exception:
                continue
        if not clickable:
            # Пробуем кликнуть сам родительский элемент
            clickable = parent

        js_click(driver, clickable, "Otkryt spisok")
        time.sleep(1)

        # Ищем "Электронная почта" в выпадающем списке
        option = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH,
                "//*[contains(text(),'Электронная почта')]"))
        )
        js_click(driver, option, "Elektronnaya pochta")
        time.sleep(0.5)
        print("  OK Sposob polucheniya vybran")
        return
    except Exception:
        pass

    # Фоллбэк: ищем select элемент
    try:
        selects = driver.find_elements(By.TAG_NAME, "select")
        visible = [s for s in selects if s.is_displayed()]
        for sel in visible:
            options = sel.find_elements(By.TAG_NAME, "option")
            for opt in options:
                if "Электронная почта" in opt.text:
                    opt.click()
                    print("  OK Sposob polucheniya vybran (select)")
                    return
    except Exception:
        pass
    print("  !! Pole 'Sposob polucheniya' ne najdeno")


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
    print(f"  ! Najdeno {len(msg_files)} .msg fajlov, berem: {msg_files[0]}")
    return os.path.join(base_dir, msg_files[0])


def attach_content(driver, file_path):
    """Нажимает 'Присоединить содержимое' и загружает файл."""
    print(f"  Prisoedinenie fajla: {os.path.basename(file_path)}")

    # Кликаем кнопку "Присоединить содержимое"
    try:
        btn = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(text(),'Присоединить содержимое')]"))
        )
        js_click(driver, btn, "Prisoedinit soderzhimoe")
        time.sleep(2)
    except Exception as e:
        print(f"  !! Knopka 'Prisoedinit soderzhimoe' ne najdena: {e}")
        return

    # Ищем скрытый input[type='file'] и отправляем путь к файлу
    try:
        file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
        file_input.send_keys(file_path)
        print(f"  OK Fajl vybran: {os.path.basename(file_path)}")
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
                print(f"  OK Fajl vybran (fallback): {os.path.basename(file_path)}")
                time.sleep(2)
            else:
                print("  !! input[type='file'] ne najden na stranice!")
                return
        except Exception as e:
            print(f"  !! Oshibka zagruzki fajla: {e}")
            return

    # Нажимаем кнопку подтверждения в диалоге
    try:
        confirm_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "#SetContentDialogBtnSend, [id*='SetContentDialogBtnSend']"))
        )
        js_click(driver, confirm_btn, "Podtverdit prisoedinenie")
        time.sleep(3)
        print("  OK Fajl prisoedinyon!")
    except Exception:
        # Фоллбэк по тексту кнопки
        try:
            btns = driver.find_elements(By.XPATH,
                "//div[contains(text(),'Присоединить содержимое')] | //button[contains(text(),'Присоединить')]")
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                js_click(driver, visible[-1], "Podtverdit (fallback)")
                time.sleep(3)
                print("  OK Fajl prisoedinyon!")
        except Exception as e:
            print(f"  !! Oshibka podtverzhdeniya: {e}")


def add_addressee(driver, person_name):
    """Добавляет адресата на вкладке Реквизиты через кнопку '+'."""
    print(f"  Adresat: {person_name}")
    try:
        # Ищем кнопку "+" рядом с разделом "Адресаты"
        plus_buttons = driver.find_elements(By.CSS_SELECTOR, "img[src*='add']")
        visible_plus = [b for b in plus_buttons if b.is_displayed()]
        if not visible_plus:
            # Фоллбэк: ищем кнопку "+" по другим селекторам
            plus_buttons = driver.find_elements(By.XPATH,
                "//div[contains(@class,'add')] | //img[contains(@class,'add')]")
            visible_plus = [b for b in plus_buttons if b.is_displayed()]

        if visible_plus:
            js_click(driver, visible_plus[0], "+ Adresat")
            time.sleep(2)
            add_person_from_directory(driver, person_name, "Adresat")
        else:
            print("  !! Knopka + ne najdena dlya adresatov")
    except Exception as e:
        print(f"  !! Oshibka dobavleniya adresata: {e}")


def go_to_distribution_tab(driver):
    """Переходит на вкладку 'Рассылка'."""
    print("  Perekhod na vkladku Rassylka...")
    try:
        tab = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(text(),'Рассылка')] | //span[contains(text(),'Рассылка')]"))
        )
        js_click(driver, tab, "Vkladka Rassylka")
        time.sleep(2)
        print("  OK Vkladka Rassylka otkryta")
    except Exception as e:
        print(f"  !! Vkladka Rassylka ne najdena: {e}")


def add_distribution_addressee(driver, person_name):
    """Добавляет адресата на вкладке 'Рассылка' через поле 'Добавить адресатов'."""
    print(f"  Rassylka adresat: {person_name}")
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
            print("  !! Pole 'Dobavit adresatov' ne najdeno!")
            return

        # Вводим фамилию
        surname = person_name.split()[0]
        corr_input.click()
        time.sleep(0.3)
        corr_input.clear()
        corr_input.send_keys(surname)
        print(f"  OK Vvedena familiya: {surname}")
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
                    print(f"  OK Adresat dobavlen: {person_name}")
                    return
            js_click(driver, visible_results[0], f"Vybor (pervyj): {visible_results[0].text.strip()}")
            time.sleep(1)
            print(f"  OK Adresat dobavlen: {person_name}")
        else:
            print(f"  !! Adresat ne najden v spiske: {person_name}")
    except Exception as e:
        print(f"  !! Oshibka: {e}")


def create_one_document(driver, doc_data, index, total):
    """Создаёт один входящий документ."""
    print(f"\n{'='*60}")
    print(f"DOKUMENT {index}/{total}")
    print(f"  Soderzhanie: {doc_data['содержание'][:80]}...")
    print(f"  Korrespondent: {doc_data['корреспондент']}")
    print(f"{'='*60}")

    # ШАГ 1: Кнопка создания документа
    print("\n[1/5] Knopka sozdaniya dokumenta...")
    el = WebDriverWait(driver, TIMEOUT).until(
        EC.presence_of_element_located((By.ID, "mainscreen-create-button"))
    )
    time.sleep(1)
    js_click(driver, el, "Knopka sozdaniya dokumenta")
    time.sleep(3)

    # ШАГ 2: Тип — Входящий документ
    print("\n[2/5] Tip: Vkhodyashchij dokument...")
    wait_and_click(driver, By.XPATH,
        "//div[contains(text(),'Входящий документ')]",
        "Vkhodyashchij dokument")
    time.sleep(1)

    # ШАГ 3: Вид — Письма, заявления и жалобы граждан, акционеров
    print("\n[3/5] Vid: Pisma, zayavleniya...")
    wait_and_click(driver, By.XPATH,
        "//div[contains(text(),'Письма, заявления и жалобы граждан')] | //td[contains(text(),'Письма, заявления и жалобы граждан')]",
        "Pisma, zayavleniya i zhaloby")
    time.sleep(0.5)

    # Кнопка "Создать документ"
    print("  Sozdat dokument...")
    wait_and_click(driver, By.XPATH,
        "//button[contains(text(),'Создать документ')] | //div[contains(text(),'Создать документ')]",
        "Sozdat dokument")
    print("  Zhdu zagruzku formy (5 sek)...")
    time.sleep(5)

    # ШАГ 4: Заполнение формы
    print("\n[4/5] Zapolnyayu formu...")

    # --- Краткое содержание ---
    print("\n  Kratkoe soderzhanie:")
    try:
        textareas = driver.find_elements(By.TAG_NAME, "textarea")
        visible_ta = [ta for ta in textareas if ta.is_displayed()]
        if visible_ta:
            visible_ta[0].click()
            time.sleep(0.3)
            visible_ta[0].clear()
            visible_ta[0].send_keys(doc_data["содержание"])
            print(f"  OK Zapolneno")
        else:
            print("  !! Textarea ne najdena")
    except Exception as e:
        print(f"  !! Oshibka: {e}")
    time.sleep(0.5)

    # --- Корреспондент ---
    print("\n  Korrespondent:")
    fill_correspondent(driver, doc_data["корреспондент"])
    time.sleep(1)

    # --- Номер у корреспондента ---
    print("\n  Nomer:")
    fill_corr_number(driver)
    time.sleep(0.5)

    # --- Дата у корреспондента ---
    print("\n  Data:")
    fill_corr_date(driver)
    time.sleep(0.5)

    # --- Адресат (Басманов) ---
    print("\n  Adresat:")
    add_addressee(driver, "Басманов Александр Владимирович")
    time.sleep(0.5)

    # --- Способ получения ---
    print("\n  Sposob polucheniya:")
    fill_delivery_method(driver)
    time.sleep(0.5)

    # ШАГ 5: Сохранение
    print("\n[5/8] Sokhranenie...")
    try:
        save_btn = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//button[contains(text(),'Сохранить')] | //div[contains(text(),'Сохранить')]"))
        )
        js_click(driver, save_btn, "Sokhranit")
        time.sleep(3)
        print(f"  OK Dokument {index}/{total} sokhranyon!")
    except Exception as e:
        print(f"  !! Oshibka sokhraneniya: {e}")

    # ШАГ 6: Присоединить содержимое
    if doc_data.get("файл"):
        print("\n[6/8] Prisoedinenie soderzhimogo...")
        attach_content(driver, doc_data["файл"])

    # ШАГ 7: Вкладка "Рассылка" — добавить Халецкую
    print("\n[7/9] Rassylka — dobavit Khaletskuyu...")
    go_to_distribution_tab(driver)
    add_distribution_addressee(driver, "Халецкая Юлия Владимировна")
    time.sleep(1)

    # ШАГ 8: Сохранить после рассылки
    print("\n[8/9] Sokhranenie posle rassylki...")
    try:
        save_btn = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(text(),'Сохранить')]"))
        )
        js_click(driver, save_btn, "Sokhranit")
        time.sleep(3)
    except Exception:
        pass

    # ШАГ 9: Зарегистрировать
    if AUTO_REGISTER:
        print("\n[9/9] Registratsiya...")
        try:
            reg_btn = WebDriverWait(driver, TIMEOUT).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "#header-action-btn-register, [id*='header-action-btn-register']"))
            )
            js_click(driver, reg_btn, "Zaregistrirovat")
            time.sleep(3)
            print(f"  OK Dokument {index}/{total} ZAREGISTRIROVAN!")
        except Exception:
            try:
                btn = driver.find_element(By.XPATH,
                    "//div[contains(text(),'Зарегистрировать')]")
                js_click(driver, btn, "Zaregistrirovat (fallback)")
                time.sleep(3)
                print(f"  OK Dokument {index}/{total} ZAREGISTRIROVAN!")
            except Exception as e:
                print(f"  !! Knopka 'Zaregistrirovat' ne najdena: {e}")
    else:
        print(f"\n[9/9] Dokument {index}/{total} gotov (bez registratsii)")

    # Возвращаемся на главную для следующего документа
    time.sleep(2)
    driver.get(ASUD_URL)
    print("  Zhdu zagruzku glavnoj...")
    time.sleep(5)


def main():
    print("=" * 60)
    print("ASUD IK - Paketnoe sozdanie Vkhodyashchikh dokumentov")
    print("=" * 60)

    # --- Поиск Excel-файла рядом с exe/скриптом ---
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    xlsx_files = [f for f in os.listdir(base_dir) if f.lower().endswith('.xlsx')]

    if not xlsx_files:
        print(f"!! Ne najden .xlsx fajl v papke: {base_dir}")
        print("   Polozhi Excel-fajl ryadom s exe.")
        input("Enter...")
        sys.exit(1)
    elif len(xlsx_files) == 1:
        excel_path = os.path.join(base_dir, xlsx_files[0])
        print(f"\nNajden fajl: {xlsx_files[0]}")
    else:
        print(f"\nNajdeno {len(xlsx_files)} xlsx-fajlov:")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  {i}. {f}")
        choice = input("Vyberi nomer fajla: ").strip()
        try:
            excel_path = os.path.join(base_dir, xlsx_files[int(choice) - 1])
        except (ValueError, IndexError):
            print("!! Nevvernyj vybor")
            input("Enter...")
            sys.exit(1)

    # --- Поиск .msg файла ---
    msg_path = get_attachment_path()
    if msg_path:
        print(f"Fajl dlya prisoedineniya: {os.path.basename(msg_path)}")
    else:
        print("! .msg fajl ne najden — dokumenty budut bez vlozheniya")

    # --- Загрузка данных ---
    print(f"\nChtenie fajla: {excel_path}")
    docs = load_excel(excel_path)
    # Добавляем путь к файлу в каждый документ
    for doc in docs:
        doc["файл"] = msg_path
    print(f"Najdeno dokumentov: {len(docs)}")

    if not docs:
        print("!! Net dannyh dlya sozdaniya!")
        input("Enter...")
        sys.exit(1)

    # Показать превью
    print("\nPervye 5 zapisej:")
    for i, d in enumerate(docs[:5], 1):
        print(f"  {i}. {d['корреспондент']} | {d['содержание'][:60]}...")

    print(f"\nVsego: {len(docs)} dokumentov")
    confirm = input("Nachat sozdanie? (da/net): ").strip().lower()
    if confirm not in ("da", "yes", "y", "д", "да"):
        print("Otmeneno.")
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
        print("\nOtkryvayu ASUD...")
        driver.get(ASUD_URL)
        print("  Zhdu zagruzku GWT (10 sek)...")
        time.sleep(10)
        print("  OK Stranica zagruzhena")

        # Создаём документы в цикле
        for i, doc in enumerate(docs, 1):
            try:
                create_one_document(driver, doc, i, len(docs))
            except Exception as e:
                print(f"\n!! OSHIBKA pri sozdanii dokumenta {i}: {e}")
                print("  Probuju sleduyushchij...")
                driver.get(ASUD_URL)
                time.sleep(5)
                continue

        print(f"\n{'='*60}")
        print(f"GOTOVO! Sozdano dokumentov: {len(docs)}")
        print(f"{'='*60}")

        input("\nEnter dlya zakrytiya brauzera...")

    except Exception as e:
        print(f"\n!! Oshibka: {e}")
        input("Enter dlya zakrytiya...")

    finally:
        driver.quit()
        print("\nOK Brauzer zakryt.")


if __name__ == "__main__":
    main()
