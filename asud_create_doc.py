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
    "проект": "000-",
}

TIMEOUT = 20


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
    
    # Способ 1: JavaScript dispatchEvent (обход бага EdgeDriver)
    try:
        driver.execute_script("""
            var evt = new MouseEvent('click', {
                bubbles: true, cancelable: true, view: window
            });
            arguments[0].dispatchEvent(evt);
        """, element)
        print(f"  ОК (JS dispatchEvent): {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass
    
    # Способ 2: JavaScript .click()
    try:
        driver.execute_script("arguments[0].click();", element)
        print(f"  ОК (JS click): {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass
    
    # Способ 3: ActionChains
    try:
        ActionChains(driver).move_to_element(element).pause(0.3).click().perform()
        print(f"  ОК (ActionChains): {description}")
        time.sleep(0.5)
        return
    except Exception:
        pass
    
    # Способ 4: Обычный клик
    try:
        element.click()
        print(f"  ОК (click): {description}")
        time.sleep(0.5)
    except Exception as e:
        print(f"  !! Все способы клика не сработали: {e}")


def wait_and_click(driver, by, selector, description="", timeout=TIMEOUT):
    print(f"  -> Ожидаю: {description or selector}")
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )
    time.sleep(0.5)
    safe_click(driver, el, description or selector)
    return el


def add_person_from_directory(driver, person_name, field_name):
    print(f"\n  Добавляю {field_name}: {person_name}")
    time.sleep(2)

    search_field = None
    for sel in [".gwt-TextBox", "input[type='text']"]:
        try:
            fields = driver.find_elements(By.CSS_SELECTOR, sel)
            visible = [f for f in fields if f.is_displayed()]
            if visible:
                search_field = visible[-1]
                break
        except Exception:
            continue

    if not search_field:
        print(f"  !! Не найдено поле поиска!")
        return

    surname = person_name.split()[0]
    search_field.clear()
    search_field.send_keys(surname)
    print(f"  ОК Фамилия: {surname}")
    time.sleep(0.5)

    clicked = False
    for sel in ["//button[contains(text(),'Найти')]", "//button[contains(text(),'Поиск')]"]:
        try:
            btns = driver.find_elements(By.XPATH, sel)
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                safe_click(driver, visible[0], "Поиск")
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        search_field.send_keys(Keys.ENTER)

    print("  Жду результаты...")
    time.sleep(3)

    try:
        result = driver.find_element(By.XPATH, f"//*[contains(text(),'{surname}')]")
        if result.is_displayed():
            try:
                ActionChains(driver).double_click(result).perform()
            except Exception:
                safe_click(driver, result, surname)
            print(f"  ОК Выбран: {person_name}")
    except Exception:
        print(f"  !! Не найден: {person_name}")

    time.sleep(1)

    for sel in ["//button[contains(text(),'Выбрать')]", "//button[contains(text(),'OK')]",
                "//button[contains(text(),'Добавить')]"]:
        try:
            btns = driver.find_elements(By.XPATH, sel)
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                safe_click(driver, visible[0], "Подтверждение")
                break
        except Exception:
            continue
    time.sleep(1)


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
        print("  Жду загрузку (10 сек)...")
        time.sleep(10)
        print("  ОК Загружено")

        # SHAG 2
        print("\n[2/7] Кнопка создания документа...")
        el = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.ID, "mainscreen-create-button"))
        )
        time.sleep(2)
        safe_click(driver, el, "Кнопка создания")
        time.sleep(3)

        # SHAG 3
        print("\n[3/7] Исходящий документ...")
        wait_and_click(driver, By.XPATH,
            "//div[contains(text(),'Исходящий документ')]",
            "Исходящий документ")
        time.sleep(1)

        # SHAG 4
        print("\n[4/7] Служебная записка...")
        wait_and_click(driver, By.XPATH,
            "//div[contains(text(),'Служебная записка')] | //td[contains(text(),'Служебная записка')]",
            "Служебная записка")
        time.sleep(0.5)

        # SHAG 5
        print("\n[5/7] Создать документ...")
        wait_and_click(driver, By.XPATH,
            "//button[contains(text(),'Создать документ')] | //div[contains(text(),'Создать документ')]",
            "Создать документ")
        print("  Жду форму (5 сек)...")
        time.sleep(5)

        # SHAG 6
        print("\n[6/7] Заполняю форму...")

        # Kratkoe soderzhanie
        print("\n  Краткое содержание:")
        try:
            textareas = driver.find_elements(By.TAG_NAME, "textarea")
            visible_ta = [ta for ta in textareas if ta.is_displayed()]
            if visible_ta:
                driver.execute_script("arguments[0].value = arguments[1];", 
                    visible_ta[0], DOC_DATA["краткое_содержание"])
                # Тригерим событие change для GWT
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
        time.sleep(0.5)

        # Adresaty
        print("\n  Адресаты:")
        for person in DOC_DATA["адресаты"]:
            try:
                plus_buttons = driver.find_elements(By.CSS_SELECTOR, "img[src*='add']")
                visible_plus = [b for b in plus_buttons if b.is_displayed()]
                if visible_plus:
                    safe_click(driver, visible_plus[0], f"+ {person}")
                    time.sleep(2)
                    add_person_from_directory(driver, person, "Адресат")
                else:
                    print("  !! Кнопка + не найдена")
            except Exception as e:
                print(f"  !! Ошибка: {e}")

        # Podpisanty
        print("\n  Подписанты:")
        for person in DOC_DATA["подписанты"]:
            try:
                plus_buttons = driver.find_elements(By.CSS_SELECTOR, "img[src*='add']")
                visible_plus = [b for b in plus_buttons if b.is_displayed()]
                if len(visible_plus) >= 2:
                    safe_click(driver, visible_plus[1], f"+ {person}")
                elif visible_plus:
                    safe_click(driver, visible_plus[-1], f"+ {person}")
                time.sleep(2)
                add_person_from_directory(driver, person, "Подписант")
            except Exception as e:
                print(f"  !! Ошибка: {e}")

        # Proekt
        print("\n  Проект:")
        try:
            inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
            visible_inputs = [i for i in inputs if i.is_displayed()]
            for inp in reversed(visible_inputs):
                try:
                    driver.execute_script("arguments[0].value = arguments[1];", 
                        inp, DOC_DATA["проект"])
                    driver.execute_script("""
                        var evt = new Event('input', {bubbles: true});
                        arguments[0].dispatchEvent(evt);
                    """, inp)
                    time.sleep(2)
                    try:
                        no_proj = driver.find_element(By.XPATH,
                            "//*[contains(translate(text(),'НЕТПРОКА','нетпрока'),'нет проекта')]")
                        if no_proj.is_displayed():
                            safe_click(driver, no_proj, "Нет проекта")
                    except Exception:
                        pass
                    break
                except Exception:
                    continue
        except Exception as e:
            print(f"  !! Ошибка: {e}")

        # SHAG 7
        print("\n[7/7] Готово!")
        print("  Документ НЕ сохранён - проверь данные.")

        input("\n  Enter для закрытия браузера...")

    except Exception as e:
        print(f"\n!! Ошибка: {e}")
        input("Enter для закрытия...")

    finally:
        driver.quit()
        print("\nОК Браузер закрыт.")


if __name__ == "__main__":
    main()
