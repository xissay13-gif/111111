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


_FIND_INPUT_JS = """
const labelText = arguments[0];
const xp = `//*[normalize-space(text())='${labelText}']`;
const snap = document.evaluate(xp, document, null,
    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
for (let i = 0; i < snap.snapshotLength; i++) {
    const label = snap.snapshotItem(i);
    if (!label.offsetParent) continue;
    let parent = label;
    for (let level = 1; level <= 5; level++) {
        parent = parent.parentElement;
        if (!parent) break;
        const inputs = parent.querySelectorAll(
            "input[id*='select_combobox-input'], input[type='text']");
        for (const inp of inputs) {
            if (inp.offsetParent && !inp.readOnly) return inp;
        }
    }
}
return null;
"""


def find_input_near_label(driver, label_text):
    """Находит input combobox рядом с лейблом — один JS-вызов вместо ~25 Selenium round-trips."""
    return driver.execute_script(_FIND_INPUT_JS, label_text)


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
    """Устанавливает значение input через JS + dispatch events.
    Подходит для plain-полей (textarea, дата, номер) — без autocomplete."""
    driver.execute_script("""
        arguments[0].focus();
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', {bubbles:true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles:true}));
    """, element, value)


_FIND_OPTIONS_JS = """
const text = arguments[0], inp = arguments[1];
const xp = `//*[contains(text(),"${text}")]`;
const snap = document.evaluate(xp, document, null,
    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
const out = [];
for (let i = 0; i < snap.snapshotLength; i++) {
    const r = snap.snapshotItem(i);
    if (!r.offsetParent) continue;
    if (r === inp) continue;
    if (r.tagName === 'INPUT') continue;
    if ((r.textContent || '').length > 150) continue;
    out.push(r);
}
return out;
"""


def find_dropdown_options(driver, text, anchor_input):
    """Возвращает видимые варианты выпадашки, где text встречается в textContent.
    Один JS-вызов вместо N+1 Selenium round-trips."""
    return driver.execute_script(_FIND_OPTIONS_JS, text, anchor_input)


def js_type_combobox(driver, element, value):
    """Печатает в combobox-autocomplete атомарно через JS.

    Для GXT/GWT-комбобоксов (корреспондент, адресат, исполнитель) —
    выпадашка фильтруется по событиям input/keyup. Просто `value=...`
    не открывает её. Здесь:
      1) фокус
      2) value = пустая строка → input/keyup (на всякий случай — сброс)
      3) value = искомая строка → input/keyup → keypress → change
    Этого хватает чтобы GXT-фильтр перерисовал список вариантов.
    """
    driver.execute_script("""
        const el = arguments[0], v = arguments[1];
        el.focus();
        el.value = '';
        el.dispatchEvent(new Event('input', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
        el.value = v;
        el.dispatchEvent(new Event('input', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keydown', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keypress', {bubbles:true}));
        el.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
    """, element, value)
