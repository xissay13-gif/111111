"""
✔ Убраны time.sleep → полностью заменены на WebDriverWait
✔ GWT-устойчивость
    click fallback через JS
    работа только с is_displayed()
    ожидание popup
✔ Код читается как сценарий
    open_create_dialog()
    choose_person_type()
    fill_form()
    save()
✔ Нет дублирования → весь UI-контроль в UI
✔ Логирование вместо print → готово для CI / лог-файлов
"""


import logging
import time
from datetime import timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================
# LOGGING
# =========================
# Время начало работы
start_time = time.monotonic()

def setup_logger():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    )
    return logging.getLogger("asud")


# =========================
# UI HELPERS (GWT SAFE)
# =========================
class UI:
    def __init__(self, driver, timeout=15):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    # ---------- WAIT ----------
    def visible(self, by, locator):
        return self.wait.until(EC.visibility_of_element_located((by, locator)))

    def clickable(self, by, locator):
        return self.wait.until(EC.element_to_be_clickable((by, locator)))

    def invisible(self, by, locator):
        return self.wait.until(EC.invisibility_of_element_located((by, locator)))

    def all(self, by, locator):
        self.wait.until(EC.presence_of_all_elements_located((by, locator)))
        return self.driver.find_elements(by, locator)

    # ---------- ACTIONS ----------
    def click(self, element):
        try:
            element.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", element)

    def click_by(self, by, locator):
        el = self.clickable(by, locator)
        self.click(el)

    def type(self, element, text, clear=True):
        if clear:
            element.clear()
        element.send_keys(text)

    def type_by(self, by, locator, text):
        el = self.visible(by, locator)
        self.type(el, text)

    # ---------- GWT SPECIFIC ----------
    def find_visible_by_text(self, text):
        xpath = f"//*[contains(normalize-space(), '{text}')]"
        elements = self.all(By.XPATH, xpath)
        return [e for e in elements if e.is_displayed()]

    def click_by_text(self, text):
        elements = self.find_visible_by_text(text)
        if not elements:
            raise Exception(f"Не найден элемент с текстом: {text}")
        self.click(elements[0])


# =========================
# LOCATORS
# =========================
class Locators:
    ADD_BTN_TEXT = "Добавить"
    PERSON_BTN_TEXT = "Физическое лицо"
    SAVE_BTN_TEXT = "Сохранить"

    POPUP = "//div[contains(@class,'popup')]"

    SURNAME = "//input[contains(@name,'surname')]"
    NAME = "//input[contains(@name,'name')]"
    EMAIL = "//input[@type='email' or contains(@name,'email')]"

    SAVE_FALLBACK_ID = "header-save-btn"


# =========================
# BUSINESS LOGIC
# =========================
class Correspondent:
    def __init__(self, driver, logger):
        self.ui = UI(driver)
        self.log = logger

    def open_create_dialog(self):
        self.log.info("Открытие диалога создания")
        self.ui.click_by_text(Locators.ADD_BTN_TEXT)
        self.ui.visible(By.XPATH, Locators.POPUP)

    def choose_person_type(self):
        self.log.info("Выбор типа: физическое лицо")
        self.ui.click_by_text(Locators.PERSON_BTN_TEXT)

    def fill_form(self, surname, name, email):
        self.log.info("Заполнение формы")

        self.ui.type_by(By.XPATH, Locators.SURNAME, surname)
        self.ui.type_by(By.XPATH, Locators.NAME, name)
        self.ui.type_by(By.XPATH, Locators.EMAIL, email)

    def save(self):
        self.log.info("Сохранение")

        buttons = self.ui.find_visible_by_text(Locators.SAVE_BTN_TEXT)

        if buttons:
            self.ui.click(buttons[0])
        else:
            self.log.warning("Fallback: кнопка сохранения через ID")
            self.ui.click_by(By.ID, Locators.SAVE_FALLBACK_ID)

        self.ui.invisible(By.XPATH, Locators.POPUP)

    def create(self, surname, name, email):
        try:
            self.open_create_dialog()
            self.choose_person_type()
            self.fill_form(surname, name, email)
            self.save()

            self.log.info("Корреспондент успешно создан")

        except Exception as e:
            self.log.error(f"Ошибка создания корреспондента: {e}")
            raise


# =========================
# MAIN (пример запуска)
# =========================
def main():
    logger = setup_logger()

    driver = webdriver.Chrome()  # или твой driver
    driver.maximize_window()

    try:
        driver.get("http://your-url-here")  # <-- вставь URL

        correspondent = Correspondent(driver, logger)

        correspondent.create(
            surname="Иванов",
            name="Иван",
            email="ivan@example.com"
        )

    finally:
        pass
        # driver.quit()  # включи при необходимости


if __name__ == "__main__":
    main()

    end_time = time.monotonic()
    elapsed_time = timedelta(seconds=end_time - start_time)
    logging.info(f"Время выполнения: {elapsed_time}")