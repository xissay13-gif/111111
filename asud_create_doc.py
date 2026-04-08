"""
Скрипт для создания Исходящего документа / Служебная записка в АСУД ИК

Установка:
    pip install selenium pyinstaller

Сборка exe:
    pyinstaller --onefile --name asud_create_doc --hidden-import=selenium.webdriver.edge.webdriver --hidden-import=selenium.webdriver.edge.service --hidden-import=selenium.webdriver.edge.options --hidden-import=selenium.webdriver.common.action_chains --hidden-import=selenium.webdriver.support.ui --hidden-import=selenium.webdriver.support.expected_conditions asud_create_doc.py

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
        print(f"!! msedgedriver.exe ne najden v papke: {base_dir}")
        print("   Polozhi msedgedriver.exe ryadom s etim fajlom.")
        input("Enter dlya vyhoda...")
        sys.exit(1)

    return driver_path


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
        # Если обычный клик не сработал — используем JS
        driver.execute_script("arguments[0].click();", el)
    print(f"  OK Klik: {description or selector}")
    time.sleep(0.5)
    return el


def add_person_from_directory(driver, person_name, field_name):
    print(f"\n  Dobavlyayu {field_name}: {person_name}")
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
        print(f"  !! Ne najdeno pole poiska v spravochnike!")
        return

    surname = person_name.split()[0]
    search_field.clear()
    search_field.send_keys(surname)
    print(f"  OK Vvedena familiya: {surname}")
    time.sleep(0.5)

    clicked = False
    for sel in ["//button[contains(text(),'Найти')]", "//button[contains(text(),'Поиск')]"]:
        try:
            btns = driver.find_elements(By.XPATH, sel)
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                js_click(driver, visible[0], "Knopka poiska")
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        search_field.send_keys(Keys.ENTER)

    print("  Zhdu rezultaty...")
    time.sleep(3)

    try:
        result = driver.find_element(By.XPATH, f"//*[contains(text(),'{surname}')]")
        if result.is_displayed():
            ActionChains(driver).double_click(result).perform()
            print(f"  OK Vybran: {person_name}")
    except Exception:
        print(f"  !! Ne udalos najti: {person_name}")

    time.sleep(1)

    for sel in ["//button[contains(text(),'Выбрать')]", "//button[contains(text(),'OK')]",
                "//button[contains(text(),'Добавить')]"]:
        try:
            btns = driver.find_elements(By.XPATH, sel)
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                js_click(driver, visible[0], "Podtverzhdenie")
                break
        except Exception:
            continue
    time.sleep(1)


def main():
    print("=" * 60)
    print("ASUD IK - Sozdanie Sluzhebnoj zapiski")
    print("=" * 60)

    driver_path = get_driver_path()
    print(f"\nEdgeDriver: {driver_path}")

    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--auth-server-whitelist=*.interrao.ru")
    options.add_argument("--auth-negotiate-delegate-whitelist=*.interrao.ru")
    # Отключаем лишние логи
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    service = EdgeService(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)

    try:
        # SHAG 1
        print("\n[1/7] Otkryvayu ASUD...")
        driver.get(ASUD_URL)
        print("  Zhdu zagruzku GWT (10 sek)...")
        time.sleep(10)
        print("  OK Stranica zagruzhena")

        # SHAG 2
        print("\n[2/7] Knopka sozdaniya dokumenta...")
        # Ждём пока кнопка появится в DOM
        el = WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.ID, "mainscreen-create-button"))
        )
        time.sleep(1)
        js_click(driver, el, "Knopka sozdaniya dokumenta")
        time.sleep(3)

        # SHAG 3
        print("\n[3/7] Tip: Iskhodyashchij dokument...")
        wait_and_click(driver, By.XPATH,
            "//div[contains(text(),'Исходящий документ')]",
            "Iskhodyashchij dokument")
        time.sleep(1)

        # SHAG 4
        print("\n[4/7] Vid: Sluzhebnaya zapiska...")
        wait_and_click(driver, By.XPATH,
            "//div[contains(text(),'Служебная записка')] | //td[contains(text(),'Служебная записка')]",
            "Sluzhebnaya zapiska")
        time.sleep(0.5)

        # SHAG 5
        print("\n[5/7] Sozdat dokument...")
        wait_and_click(driver, By.XPATH,
            "//button[contains(text(),'Создать документ')] | //div[contains(text(),'Создать документ')]",
            "Sozdat dokument")
        print("  Zhdu zagruzku formy (5 sek)...")
        time.sleep(5)

        # SHAG 6
        print("\n[6/7] Zapolnyayu formu...")

        # Kratkoe soderzhanie
        print("\n  Kratkoe soderzhanie:")
        try:
            textareas = driver.find_elements(By.TAG_NAME, "textarea")
            visible_ta = [ta for ta in textareas if ta.is_displayed()]
            if visible_ta:
                visible_ta[0].click()
                time.sleep(0.3)
                visible_ta[0].clear()
                visible_ta[0].send_keys(DOC_DATA["краткое_содержание"])
                print(f"  OK Zapolneno")
            else:
                print("  !! Textarea ne najdena")
        except Exception as e:
            print(f"  !! Oshibka: {e}")
        time.sleep(0.5)

        # Adresaty
        print("\n  Adresaty:")
        for person in DOC_DATA["адресаты"]:
            try:
                plus_buttons = driver.find_elements(By.CSS_SELECTOR, "img[src*='add']")
                visible_plus = [b for b in plus_buttons if b.is_displayed()]
                if visible_plus:
                    js_click(driver, visible_plus[0], f"+ dlya {person}")
                    time.sleep(2)
                    add_person_from_directory(driver, person, "Adresat")
                else:
                    print("  !! Knopka + ne najdena")
            except Exception as e:
                print(f"  !! Oshibka: {e}")

        # Podpisanty
        print("\n  Podpisanty:")
        for person in DOC_DATA["подписанты"]:
            try:
                plus_buttons = driver.find_elements(By.CSS_SELECTOR, "img[src*='add']")
                visible_plus = [b for b in plus_buttons if b.is_displayed()]
                if len(visible_plus) >= 2:
                    js_click(driver, visible_plus[1], f"+ dlya {person}")
                elif visible_plus:
                    js_click(driver, visible_plus[-1], f"+ dlya {person}")
                time.sleep(2)
                add_person_from_directory(driver, person, "Podpisant")
            except Exception as e:
                print(f"  !! Oshibka: {e}")

        # Proekt
        print("\n  Proekt:")
        try:
            inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
            visible_inputs = [i for i in inputs if i.is_displayed()]
            for inp in reversed(visible_inputs):
                try:
                    inp.click()
                    time.sleep(0.3)
                    inp.clear()
                    inp.send_keys(DOC_DATA["проект"])
                    inp.send_keys(Keys.ENTER)
                    time.sleep(2)
                    try:
                        no_proj = driver.find_element(By.XPATH,
                            "//*[contains(translate(text(),'НЕТПРОКА','нетпрока'),'нет проекта')]")
                        if no_proj.is_displayed():
                            js_click(driver, no_proj, "Net proekta")
                    except Exception:
                        pass
                    break
                except Exception:
                    continue
        except Exception as e:
            print(f"  !! Oshibka: {e}")

        # SHAG 7
        print("\n[7/7] Gotovo!")
        print("  Dokument NE sohranyeon - prover dannye.")

        # Raskомментируй для автосохранения:
        # wait_and_click(driver, By.XPATH, "//button[contains(text(),'Сохранить')]", "Sohranit")

        input("\n  Enter dlya zakrytiya brauzera...")

    except Exception as e:
        print(f"\n!! Oshibka: {e}")
        input("Enter dlya zakrytiya...")

    finally:
        driver.quit()
        print("\nOK Brauzer zakryt.")


if __name__ == "__main__":
    main()
