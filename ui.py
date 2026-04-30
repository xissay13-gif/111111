"""
ui.py — Selenium UI-хелперы для АСУД (GWT/GXT).

Единый click(), find_input_near_label(), ожидания, работа с модалками.
Все паузы сохранены — GWT без них ломается.
"""

import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

log = logging.getLogger("asud.ui")


def click(driver, element, description=""):
    """Единый клик с fallback'ами для GWT/GXT.
    Порядок: ActionChains → native → JS → mouse events.

    Без post-sleep — caller сам ждёт следующий элемент через WebDriverWait.
    """
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center',inline:'center'});", element)
    except Exception:
        pass

    # ActionChains — лучше всего для GXT dropdown/autocomplete
    try:
        ActionChains(driver).move_to_element(element).pause(0.15).click().perform()
        log.info(f"Клик (mouse): {description}")
        return True
    except Exception:
        pass

    # Обычный Selenium click
    try:
        element.click()
        log.info(f"Клик (native): {description}")
        return True
    except Exception:
        pass

    # JS .click()
    try:
        driver.execute_script("arguments[0].click();", element)
        log.info(f"Клик (JS): {description}")
        return True
    except Exception:
        pass

    # Полный набор mouse-событий
    try:
        driver.execute_script("""
            var el = arguments[0];
            ['mouseover','mousedown','mouseup','click'].forEach(function(type) {
                el.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, view:window}));
            });
        """, element)
        log.info(f"Клик (events): {description}")
        return True
    except Exception as e:
        log.error(f"Клик не удался: {description}: {e}")
        return False


def wait_and_click(driver, by, selector, description="", timeout=20):
    """Ждёт элемент и кликает. Без post-sleep."""
    log.info(f"Ожидаю: {description or selector}")
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, selector))
    )
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    log.info(f"Клик: {description or selector}")
    return el


def find_input_near_label(driver, label_text):
    """Находит input combobox рядом с лейблом (точное совпадение)."""
    labels = driver.find_elements(By.XPATH,
        f"//*[normalize-space(text())='{label_text}']")
    for label in labels:
        try:
            if not label.is_displayed():
                continue
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


def wait_asud_loaded(driver, max_wait=120):
    """Адаптивное ожидание полной загрузки АСУД."""
    log.info("Жду загрузку АСУД...")
    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: d.execute_script("return document.readyState === 'complete'"))
    except Exception:
        log.warning("readyState не complete")

    try:
        WebDriverWait(driver, max_wait).until(
            EC.element_to_be_clickable((By.ID, "mainscreen-create-button")))
    except Exception:
        log.warning("Кнопка создания не появилась")

    try:
        WebDriverWait(driver, max_wait).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR,
                "tr[class*='GridView-row'], tr[class*='grid-row'], "
                "tr[class*='OSHSGridStyle-row'], tr[class*='obj-list-rec']")) > 0)
    except Exception:
        log.warning("Данные в таблице не появились")

    log.info("АСУД загружен")


def wait_modal_closed(driver, timeout=15):
    """Ждёт пока закроется модальное окно GXT ModalPanel."""
    log.info("Жду закрытия модалки...")
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: not any(
                m.is_displayed() for m in d.find_elements(
                    By.CSS_SELECTOR, "div[class*='ModalPanel'][class*='panel']")))
        log.info("Модалка закрыта")
    except Exception:
        log.warning("Модалка не закрылась — Escape")
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
        except Exception:
            pass


def close_open_modals(driver, max_escapes=5):
    """Закрывает все модалки через Escape."""
    log.info("Закрываю модалки...")
    for i in range(max_escapes):
        modals = driver.find_elements(By.CSS_SELECTOR,
            "div[class*='ModalPanel'][class*='panel']")
        visible = [m for m in modals if m.is_displayed()]
        if not visible:
            log.info(f"Модалки закрыты (попыток: {i})")
            return
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)
    log.warning(f"Не все модалки закрылись после {max_escapes} Escape")


def js_set_value(driver, element, value):
    """Устанавливает значение input через JS + dispatch events."""
    driver.execute_script("""
        arguments[0].focus();
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
    """, element, value)
