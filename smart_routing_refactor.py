import time
import os
import logging
from datetime import date, timedelta

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ================= LOGGING =================
# Время начало работы
start_time = time.monotonic()
# Настройка логирования
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S,%ms')


# ================= CONFIG =================
ASUD_URL = "https://asud.interrao.ru/asudik/"
TIMEOUT = 20


DOC_TYPE_MAP = {
    1: "Указы, распоряжения Президента Российской Федерации",
    2: "Документы Администрации Президента",
    3: "Документы Правительства Российской Федерации",
    4: "Документы Федеральных органов исполнительной и законодательной власти",
    5: "Письма юридических лиц",
    6: "Письма компаний ТЭК",
    7: "Документы субъектов",
    8: "Письма и жалобы граждан",
}


# ================= SELENIUM CORE =================

def wait(driver, condition, timeout=TIMEOUT):
    return WebDriverWait(driver, timeout).until(condition)


def safe_click(driver, element, name=""):
    """Унифицированный клик (устойчивый к GWT/GXT)"""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(element).pause(0.1).click().perform()
        return True
    except Exception:
        pass

    try:
        element.click()
        return True
    except Exception:
        pass

    try:
        driver.execute_script("arguments[0].click();", element)
        return True
    except Exception:
        logging.exception(f"CLICK FAILED: {name}")
        return False


def find_visible(driver, by, selector):
    els = driver.find_elements(by, selector)
    return [e for e in els if e.is_displayed()]


def js_set_value(driver, element, value):
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
    """, element, value)


# ================= EXCEL =================

COL = {
    "link": 0,
    "subject": 1,
    "body": 2,
    "type": 3,
}


def load_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    rows = []
    skipped = 0

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row or len(row) < 4:
            skipped += 1
            continue

        try:
            type_idx = int(row[COL["type"]]) if row[COL["type"]] else 0
        except:
            type_idx = 0

        if type_idx not in DOC_TYPE_MAP:
            skipped += 1
            continue

        subject = str(row[COL["subject"]] or "").strip()
        body = str(row[COL["body"]] or "").strip()

        if not subject:
            skipped += 1
            continue

        rows.append({
            "link": row[COL["link"]],
            "subject": subject,
            "body": body,
            "type": type_idx,
            "type_name": DOC_TYPE_MAP[type_idx],
        })

    logging.info(f"Loaded: {len(rows)} skipped: {skipped}")
    return rows


# ================= UI HELPERS =================

def find_input_near(driver, label):
    labels = driver.find_elements(By.XPATH, f"//*[normalize-space(text())='{label}']")
    for l in labels:
        try:
            parent = l
            for _ in range(5):
                parent = parent.find_element(By.XPATH, "..")
                inputs = parent.find_elements(By.CSS_SELECTOR, "input[type='text']")
                vis = [i for i in inputs if i.is_displayed()]
                if vis:
                    return vis[0]
        except:
            continue
    return None


def wait_click(driver, by, selector, text=""):
    el = wait(driver, EC.presence_of_element_located((by, selector)))
    safe_click(driver, el, text)
    return el


# ================= BUSINESS LOGIC =================

def fill_correspondent(driver, name):
    inp = find_input_near(driver, "Корреспондент")
    if not inp:
        logging.warning("No correspondent input")
        return

    surname = name.split()[0]

    inp.click()
    inp.clear()
    inp.send_keys(surname)

    wait(driver, lambda d: len(d.find_elements(By.XPATH, f"//*[contains(text(),'{surname}')]")) > 0)

    options = find_visible(driver, By.XPATH, f"//*[contains(text(),'{surname}')]")

    if options:
        safe_click(driver, options[0], "correspondent")
    else:
        logging.info("No match → creating new")


def fill_date(driver):
    today = date.today().strftime("%d.%m.%Y")

    inp = find_input_near(driver, "Дата у корреспондента")
    if not inp:
        return

    inp.click()
    inp.send_keys(Keys.CONTROL + "a")
    inp.send_keys(Keys.DELETE)
    inp.send_keys(today)


def fill_number(driver, link):
    inp = find_input_near(driver, "Номер у корреспондента")
    if not inp:
        return

    value = f"б/н {link}" if link else "б/н"
    js_set_value(driver, inp, value)


def fill_text(driver, text):
    areas = driver.find_elements(By.TAG_NAME, "textarea")
    vis = [a for a in areas if a.is_displayed()]
    if vis:
        vis[0].click()
        vis[0].clear()
        vis[0].send_keys(text)


# ================= DOCUMENT FLOW =================

def create_document(driver, doc, i, total):

    logging.info(f"[{i}/{total}] {doc['subject'][:60]}")

    wait_click(driver, By.ID, "mainscreen-create-button", "create")
    wait_click(driver, By.XPATH, "//div[contains(text(),'Входящий документ')]")
    wait_click(driver, By.XPATH, "//button[contains(text(),'Создать документ')]")

    fill_text(driver, doc["body"])
    fill_correspondent(driver, "Басманов Александр Владимирович")
    fill_number(driver, doc["link"])
    fill_date(driver)

    save = wait(driver, EC.element_to_be_clickable((By.ID, "header-save-btn")))
    safe_click(driver, save, "save")

    logging.info("Saved")


# ================= MAIN =================

def main():

    base = os.path.dirname(os.path.abspath(__file__))
    excel = [f for f in os.listdir(base) if f.endswith(".xlsx")][0]

    data = load_excel(os.path.join(base, excel))

    driver = webdriver.Edge(service=EdgeService("msedgedriver.exe"))
    driver.get(ASUD_URL)

    wait(driver, lambda d: d.execute_script("return document.readyState") == "complete")

    for i, doc in enumerate(data, 1):
        create_document(driver, doc, i, len(data))

    driver.quit()


if __name__ == "__main__":
    main()

    end_time = time.monotonic()
    elapsed_time = timedelta(seconds=end_time - start_time)
    logging.info(f"Время выполнения: {elapsed_time}")