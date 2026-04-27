"""
attachments.py — Прикрепление .msg файла (пустышки) к документу.

Ищет .msg рядом с exe, прикрепляет через pywinauto.
"""

import os
import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from ui import click, wait_modal_closed

log = logging.getLogger("asud.attach")

try:
    from pywinauto.application import Application
    PYWINAUTO = True
except ImportError:
    PYWINAUTO = False


# Спецсимволы которые pywinauto.type_keys интерпретирует как клавиши-модификаторы
# (а не как буквальный символ для печати).
# +=Shift, ^=Ctrl, %=Alt, ~=Enter, (){}=группировки/escape
_TYPE_KEYS_SPECIAL = set('+^%~(){}')


def _escape_for_type_keys(text):
    """Экранирует спецсимволы pywinauto.type_keys через '{X}'.
    Без этого путь типа 'D:\\Работа+проекты\\Имя(копия).msg' будет
    напечатан с зажатием Shift и т.п. — реальное имя поломается."""
    return ''.join('{' + c + '}' if c in _TYPE_KEYS_SPECIAL else c for c in text)


def get_dummy_msg(base_dir):
    """Ищет .msg-пустышку рядом с exe."""
    msg_files = [f for f in os.listdir(base_dir) if f.lower().endswith('.msg')]
    if msg_files:
        return os.path.join(base_dir, msg_files[0])
    return None


def attach_content(driver, file_path):
    """Прикрепляет файл. Сначала через input[type=file], затем pywinauto."""
    log.info(f"Прикрепление: {os.path.basename(file_path)}")

    # Стратегия 1: input[type=file] уже в DOM
    inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
    if inputs:
        try:
            driver.execute_script("""
                var el = arguments[0];
                el.style.display='block'; el.style.visibility='visible';
                el.style.opacity='1'; el.removeAttribute('hidden');
            """, inputs[0])
            time.sleep(0.3)
            inputs[0].send_keys(file_path)
            time.sleep(1)
            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                inputs[0])
            log.info("Файл отправлен через input[type=file]")
            _confirm_attach(driver)
            return
        except Exception as e:
            log.warning(f"input[type=file] не сработал: {e}")

    # Стратегия 2: кнопка + pywinauto
    if not PYWINAUTO:
        log.error("pywinauto не установлен — пропускаю")
        return

    try:
        btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(text(),'Присоединить содержимое')]")))
        click(driver, btn, "Присоединить содержимое")
    except Exception as e:
        log.error(f"Кнопка 'Присоединить содержимое' не найдена: {e}")
        return

    time.sleep(2)

    try:
        app = None
        for title_re in [".*Открыт.*", ".*Open.*", ".*Выбор.*", ".*Choose.*"]:
            try:
                app = Application(backend='win32').connect(title_re=title_re, timeout=10)
                break
            except Exception:
                continue

        if not app:
            log.error("Окно Explorer не найдено")
            return

        dlg = app.top_window()
        dlg.set_focus()
        time.sleep(0.5)
        dlg.type_keys(_escape_for_type_keys(file_path),
                      with_spaces=True, pause=0.02)
        time.sleep(0.5)
        dlg.type_keys("{ENTER}")
        time.sleep(2)
        log.info("Файл выбран через Explorer")
    except Exception as e:
        log.error(f"Ошибка pywinauto: {e}")
        return

    _confirm_attach(driver)


def _confirm_attach(driver):
    """Подтверждает загрузку в модалке АСУД."""
    time.sleep(2)
    try:
        confirm_btn = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "#SetContentDialogBtnSend, [id*='SetContentDialogBtnSend']")))
        click(driver, confirm_btn, "Подтвердить присоединение")
        time.sleep(3)
        log.info("Файл присоединён!")
    except Exception:
        try:
            btns = driver.find_elements(By.XPATH,
                "//button[contains(text(),'Присоединить')] | //div[contains(text(),'Присоединить')]")
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                click(driver, visible[-1], "Подтвердить (fallback)")
                time.sleep(3)
                log.info("Файл присоединён (fallback)")
        except Exception as e:
            log.error(f"Ошибка подтверждения: {e}")
