"""
Скрипт для создания Исходящего документа / Служебная записка в АСУД ИК

Положи msedgedriver.exe рядом с exe/скриптом.
"""

import time
import sys
import os
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

DOC_DATA = {
    "краткое_содержание": "О возврате денежных средств в размере",
    "адресаты": [
        "Басманов Александр Владимирович",
        "Халецкая Юлия Владимировна",
    ],
    "подписанты": [
        "Матус Елена Анатольевна",
    ],
    "проект": "00-",
}

TIMEOUT = 20
PAUSE = 3  # пауза между действиями (сек)


def get_driver_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    driver_path = os.path.join(base_dir, "msedgedriver.exe")
    if not os.path.exists(driver_path):
        print(f"!! msedgedriver.exe не найден в: {base_dir}")
        input("Enter...")
        sys.exit(1)
    return driver_path


def safe_click(driver, element, description=""):
    """Несколько способов клика — от надёжного к запасному."""
    print(f"  -> Клик: {description}")

    # Прокручиваем к элементу
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.5)
    except Exception:
        pass

    # Способ 1: ActionChains (лучше всего для GWT/GXT)
    try:
        ActionChains(driver).move_to_element(element).pause(0.5).click().perform()
        print(f"  ОК (ActionChains): {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass

    # Способ 2: Обычный клик
    try:
        element.click()
        print(f"  ОК (click): {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass

    # Способ 3: JavaScript .click()
    try:
        driver.execute_script("arguments[0].click();", element)
        print(f"  ОК (JS click): {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass

    # Способ 4: JavaScript dispatchEvent (запасной)
    try:
        driver.execute_script("""
            var evt = new MouseEvent('click', {
                bubbles: true, cancelable: true, view: window
            });
            arguments[0].dispatchEvent(evt);
        """, element)
        print(f"  ОК (JS dispatchEvent): {description}")
        time.sleep(0.5)
    except Exception as e:
        print(f"  !! Все способы клика не сработали: {e}")


def wait_and_click(driver, by, selector, description="", timeout=TIMEOUT):
    print(f"  -> Ожидаю: {description or selector}")
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, selector))
    )
    time.sleep(PAUSE)
    safe_click(driver, el, description or selector)
    time.sleep(PAUSE)
    return el


def find_section_input(driver, section_label):
    """Находит input-поле combobox рядом с указанной секцией."""
    # Находим лейбл секции
    try:
        section = driver.find_element(By.XPATH,
            f"//*[contains(text(),'{section_label}')]")
    except Exception:
        print(f"  !! Секция '{section_label}' не найдена")
        return None

    # Ищем input внутри родительских контейнеров
    for level in range(1, 8):
        try:
            parent = section
            for _ in range(level):
                parent = parent.find_element(By.XPATH, "..")
            inputs = parent.find_elements(By.CSS_SELECTOR,
                "input[type='text']")
            visible = [i for i in inputs if i.is_displayed() and i.get_attribute("readonly") is None]
            if visible:
                return visible[0]
        except Exception:
            continue
    return None


def add_person_to_combobox(driver, section_label, person_name):
    """Добавляет человека через combobox в указанной секции."""
    print(f"\n  Добавляю в '{section_label}': {person_name}")
    surname = person_name.split()[0]

    input_field = find_section_input(driver, section_label)
    if not input_field:
        print(f"  !! Поле ввода не найдено для '{section_label}'")
        return

    # Кликаем на поле и вводим фамилию
    safe_click(driver, input_field, f"Поле '{section_label}'")
    time.sleep(1)
    input_field.clear()
    input_field.send_keys(surname)
    print(f"  Ввожу фамилию: {surname}")
    time.sleep(PAUSE)

    # Ждём выпадающий список и выбираем
    try:
        options = driver.find_elements(By.XPATH,
            f"//*[contains(text(),'{surname}')]")
        for opt in options:
            try:
                if opt == input_field:
                    continue
                if not opt.is_displayed():
                    continue
                # Пропускаем сам лейбл секции
                tag = opt.tag_name.lower()
                if tag in ('label', 'span', 'td') and opt.text.strip() == section_label:
                    continue
                safe_click(driver, opt, f"Выбор: {person_name}")
                print(f"  ОК Выбран: {person_name}")
                time.sleep(PAUSE)
                return
            except Exception:
                continue
    except Exception:
        pass

    # Если выпадающий список не появился — Enter
    print(f"  Выпадающий список не найден, пробую Enter...")
    input_field.send_keys(Keys.ENTER)
    time.sleep(PAUSE)


def main():
    print("=" * 60)
    print("АСУД ИК - Создание Служебной записки")
    print("=" * 60)

    driver_path = get_driver_path()
    print(f"\nEdgeDriver: {driver_path}")

    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--auth-server-whitelist=*.interrao.ru")
    options.add_argument("--auth-negotiate-delegate-whitelist=*.interrao.ru")
    options.add_argument("--log-level=3")
    options.add_argument("--disable-gpu")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    service = EdgeService(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)

    try:
        # SHAG 1
        print("\n[1/7] Открываю АСУД...")
        driver.get(ASUD_URL)
        print("  Жду загрузку (60 сек)...")
        time.sleep(60)
        print("  ОК Загружено")

        # SHAG 2
        print("\n[2/7] Кнопка создания документа...")
        el = WebDriverWait(driver, TIMEOUT).until(
            EC.element_to_be_clickable((By.ID, "mainscreen-create-button"))
        )
        time.sleep(PAUSE)
        safe_click(driver, el, "Кнопка создания")
        print("  Жду открытие диалога...")
        time.sleep(PAUSE)

        # SHAG 3
        print("\n[3/7] Исходящий документ...")
        # У элемента есть id="Исходящий документ" (из DevTools)
        el = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//*[@id='Исходящий документ'] | //div[contains(text(),'Исходящий документ')]"))
        )
        time.sleep(PAUSE)
        safe_click(driver, el, "Исходящий документ")
        print("  Жду загрузку подтипов...")
        time.sleep(PAUSE)

        # SHAG 4
        print("\n[4/7] Служебная записка...")
        # Ищем в правой таблице Вид — может быть div, td или span
        el = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH,
                "//*[contains(text(),'Служебная записка')]"))
        )
        time.sleep(PAUSE)
        safe_click(driver, el, "Служебная записка")
        time.sleep(PAUSE)

        # SHAG 5
        print("\n[5/7] Создать документ...")
        wait_and_click(driver, By.XPATH,
            "//button[contains(text(),'Создать документ')] | //div[contains(text(),'Создать документ')]",
            "Создать документ")
        print("  Жду форму...")
        time.sleep(PAUSE)

        # SHAG 6
        print("\n[6/7] Заполняю форму...")

        print("\n  Краткое содержание:")
        try:
            textareas = driver.find_elements(By.TAG_NAME, "textarea")
            visible_ta = [ta for ta in textareas if ta.is_displayed()]
            if visible_ta:
                driver.execute_script("arguments[0].value = arguments[1];",
                    visible_ta[0], DOC_DATA["краткое_содержание"])
                driver.execute_script("""
                    var evt = new Event('input', {bubbles: true});
                    arguments[0].dispatchEvent(evt);
                    var evt2 = new Event('change', {bubbles: true});
                    arguments[0].dispatchEvent(evt2);
                """, visible_ta[0])
                print(f"  ОК Заполнено")
            else:
                print("  !! Textarea не найдена")
        except Exception as e:
            print(f"  !! Ошибка: {e}")
        time.sleep(PAUSE)

        print("\n  Адресаты:")
        for person in DOC_DATA["адресаты"]:
            try:
                add_person_to_combobox(driver, "Адресаты", person)
            except Exception as e:
                print(f"  !! Ошибка: {e}")

        print("\n  Подписанты:")
        for person in DOC_DATA["подписанты"]:
            try:
                add_person_to_combobox(driver, "Подписанты", person)
            except Exception as e:
                print(f"  !! Ошибка: {e}")

        print("\n  Проект:")
        try:
            # Кликаем "+" у "Добавление проекта"
            plus_btn = None
            # Ищем img с data-marker="select-btn" рядом с "Добавление проекта"
            try:
                section = driver.find_element(By.XPATH,
                    "//*[contains(text(),'Добавление проекта')]")
                parent = section
                for _ in range(5):
                    parent = parent.find_element(By.XPATH, "..")
                    btns = parent.find_elements(By.CSS_SELECTOR,
                        "img[data-marker='select-btn'], img.gwt-Image")
                    visible = [b for b in btns if b.is_displayed()]
                    if visible:
                        plus_btn = visible[-1]
                        break
            except Exception:
                pass

            # Запасной вариант: ищем все img.gwt-Image с select-btn
            if not plus_btn:
                btns = driver.find_elements(By.CSS_SELECTOR,
                    "img[data-marker='select-btn']")
                visible = [b for b in btns if b.is_displayed()]
                if visible:
                    plus_btn = visible[-1]

            if plus_btn:
                safe_click(driver, plus_btn, "+ Добавление проекта")
                time.sleep(PAUSE)
            else:
                print("  !! Кнопка + проекта не найдена")

            # В диалоге "Множественный выбор": ждём загрузку диалога
            print("  Жду загрузку диалога проектов...")
            time.sleep(PAUSE)

            # Находим поле поиска в диалоге
            search_input = None
            dialog_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
            for inp in dialog_inputs:
                try:
                    if inp.is_displayed() and inp.is_enabled():
                        search_input = inp
                        break
                except Exception:
                    continue

            if search_input:
                # Кликаем на поле чтобы оно стало активным
                safe_click(driver, search_input, "Поле поиска проекта")
                time.sleep(1)
                search_input.clear()
                time.sleep(0.5)
                # Вводим по символу для надёжности
                for char in DOC_DATA["проект"]:
                    search_input.send_keys(char)
                    time.sleep(0.2)
                print(f"  Ввожу код проекта: {DOC_DATA['проект']}")
                time.sleep(PAUSE)
                # Enter для поиска
                search_input.send_keys(Keys.ENTER)
                time.sleep(PAUSE)
            else:
                print("  !! Поле поиска проекта не найдено")

            # Выбираем первый результат (00-000-000 Нет проекта)
            try:
                result = WebDriverWait(driver, TIMEOUT).until(
                    EC.presence_of_element_located((By.XPATH,
                        "//*[contains(text(),'Нет проекта')] | //*[contains(text(),'00-000')]"))
                )
                safe_click(driver, result, "Выбор проекта")
                time.sleep(PAUSE)
            except Exception:
                print("  !! Проект не найден в списке")

            # Нажимаем "Готово"
            try:
                done_btn = driver.find_element(By.XPATH,
                    "//button[contains(text(),'Готово')]")
                if done_btn.is_displayed():
                    safe_click(driver, done_btn, "Готово")
                    time.sleep(PAUSE)
            except Exception:
                print("  !! Кнопка 'Готово' не найдена")

        except Exception as e:
            print(f"  !! Ошибка: {e}")

        # SHAG 7
        print("\n[7/7] Сохраняю документ...")
        try:
            save_btn = WebDriverWait(driver, TIMEOUT).until(
                EC.element_to_be_clickable((By.ID, "header-save-btn"))
            )
            time.sleep(PAUSE)
            safe_click(driver, save_btn, "Сохранить")
            time.sleep(PAUSE)
            print("  ОК Документ сохранён!")
        except Exception as e:
            print(f"  !! Ошибка сохранения: {e}")

        input("\n  Enter для закрытия браузера...")

    except Exception as e:
        print(f"\n!! Ошибка: {e}")
        input("Enter для закрытия...")

    finally:
        driver.quit()
        print("\nОК Браузер закрыт.")


if __name__ == "__main__":
    main()
