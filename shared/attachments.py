"""
attachments.py — Поиск и прикрепление .msg файлов.

Ищет файл по Link из Excel в outlook_dir (рекурсивно по подпапкам),
прикрепляет через pywinauto (нативный Windows Explorer).
"""

import os
import re
import time
import logging
from datetime import datetime, date
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from shared.ui import click, wait_modal_closed

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


def _link_to_variants(link):
    """Генерирует все возможные имена файлов для link.
    Форматы: DD.MM.YYYY / YYYY-MM-DD, с/без ведущего нуля, суффиксы _1-_9."""
    variants = set()

    if isinstance(link, (datetime, date)):
        try:
            variants.add(link.strftime("%d.%m.%Y %H-%M-%S"))
            variants.add(link.strftime("%Y-%m-%d %H-%M-%S"))
            with_lead = link.strftime("%d.%m.%Y %H-%M-%S")
            m = re.match(r'^(\d{2}\.\d{2}\.\d{4})\s+0(\d)-(\d{2})-(\d{2})$', with_lead)
            if m:
                variants.add(f"{m.group(1)} {m.group(2)}-{m.group(3)}-{m.group(4)}")
        except Exception:
            pass
    else:
        s = str(link).strip()
        if s:
            variants.add(s)
            # DD.MM.YYYY → ISO (+ с/без ведущего нуля в часе)
            m = re.match(r'^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2})-(\d{2})-(\d{2})$', s)
            if m:
                dd, mm, yyyy, h, mn, sec = m.groups()
                h_lead = f"{int(h):02d}"
                h_no = str(int(h))
                variants.add(f"{dd}.{mm}.{yyyy} {h_lead}-{mn}-{sec}")
                variants.add(f"{dd}.{mm}.{yyyy} {h_no}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_lead}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_no}-{mn}-{sec}")
            # ISO → DD.MM.YYYY
            m = re.match(r'^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2})-(\d{2})-(\d{2})$', s)
            if m:
                yyyy, mm, dd, h, mn, sec = m.groups()
                h_lead = f"{int(h):02d}"
                h_no = str(int(h))
                variants.add(f"{dd}.{mm}.{yyyy} {h_lead}-{mn}-{sec}")
                variants.add(f"{dd}.{mm}.{yyyy} {h_no}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_lead}-{mn}-{sec}")
                variants.add(f"{yyyy}-{mm}-{dd} {h_no}-{mn}-{sec}")

    variants.discard('')
    return list(variants)


def find_msg_by_link(link, outlook_dir, fallback_path=None):
    """Ищет .msg файл в outlook_dir (рекурсивно по подпапкам) по link.
    Fallback → пустышка."""
    log.info(f"link={link!r} (тип: {type(link).__name__})")

    if link is None or (isinstance(link, str) and not link.strip()):
        log.warning("link пустой — пустышка")
        return fallback_path

    if not os.path.isdir(outlook_dir):
        log.warning(f"Папка {outlook_dir} не найдена — пустышка")
        return fallback_path

    variants = _link_to_variants(link)
    log.info(f"Варианты: {variants}")

    # Один проход по дереву — собираем все .msg
    all_msg = []  # [(full_path, filename_no_ext)]
    try:
        for root, _dirs, files in os.walk(outlook_dir):
            for f in files:
                if f.lower().endswith('.msg'):
                    name_no_ext = os.path.splitext(f)[0]
                    all_msg.append((os.path.join(root, f), name_no_ext))
    except Exception as e:
        log.error(f"Ошибка обхода {outlook_dir}: {e}")
        return fallback_path

    if not all_msg:
        log.warning(f"В {outlook_dir} нет .msg файлов")
        return fallback_path

    def _rel(full):
        try:
            return os.path.relpath(full, outlook_dir)
        except Exception:
            return os.path.basename(full)

    # Фаза 1: точное совпадение с вариантом
    variants_set = set(variants)
    for full, name in all_msg:
        if name in variants_set:
            log.info(f"Нашёл: {_rel(full)}")
            return full

    # Фаза 2: variant + _1.._9
    suffix_set = {f"{v}_{i}" for v in variants for i in range(1, 10)}
    for full, name in all_msg:
        if name in suffix_set:
            log.info(f"Нашёл (суффикс): {_rel(full)}")
            return full

    # Фаза 3: подстрока
    for full, name in all_msg:
        for v in variants:
            if v in name:
                log.info(f"Нашёл (подстрока): {_rel(full)}")
                return full

    log.warning("Файл не найден — пустышка")
    return fallback_path


def get_dummy_msg(base_dir):
    """Ищет .msg-пустышку рядом с exe."""
    msg_files = [f for f in os.listdir(base_dir) if f.lower().endswith('.msg')]
    if msg_files:
        return os.path.join(base_dir, msg_files[0])
    return None


def move_to_done(file_path, outlook_dir, done_dirname="Завершено"):
    """Перемещает обработанный .msg в <outlook_dir>/Завершено/.
    Папка создаётся если нет. Конфликт имён → суффикс _HHMMSS.
    Ничего не делает если file_path пустой/не существует.
    """
    if not file_path or not os.path.isfile(file_path):
        return
    if not outlook_dir or not os.path.isdir(outlook_dir):
        log.warning(f"outlook_dir '{outlook_dir}' не существует — "
                    f"не перемещаю {os.path.basename(file_path)}")
        return

    done_dir = os.path.join(outlook_dir, done_dirname)
    try:
        os.makedirs(done_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Не удалось создать {done_dir}: {e}")
        return

    name = os.path.basename(file_path)
    dest = os.path.join(done_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        ts = datetime.now().strftime("%H%M%S")
        dest = os.path.join(done_dir, f"{base}_{ts}{ext}")

    try:
        import shutil
        shutil.move(file_path, dest)
        log.info(f"→ Завершено/{os.path.basename(dest)}")
    except Exception as e:
        log.warning(f"Не удалось переместить {name}: {e}")


_DIAG_BUTTONS_JS = r"""
// Диагностика: вернуть видимые кандидаты на confirm-кнопку
const out = [];
for (const el of document.querySelectorAll(
        "[id*='SetContent'], [id*='Send'], [id*='Submit'], button, div, span")) {
    if (!el.offsetParent) continue;
    const t = (el.textContent || '').trim();
    if (t.length > 80) continue;
    if (!/(Присоединить|Подтвердить|Сохранить|Загрузить)/i.test(t)
        && !/(SetContent|Send|Submit|Save)/i.test(el.id || '')) continue;
    out.push((el.id || el.tagName) + '|' + t.slice(0, 60));
}
return out;
"""


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
                el.dispatchEvent(new Event('change',{bubbles:true}));
            """, inputs[0])
            inputs[0].send_keys(file_path)
            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                inputs[0])
            log.info("Файл отправлен через input[type=file]")
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
        # Пауза перед открытием модалки — даём АСУД полностью прогрузить состояние
        time.sleep(2)
        click(driver, btn, "Присоединить содержимое")
    except Exception as e:
        log.error(f"Кнопка 'Присоединить содержимое' не найдена: {e}")
        return

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
        dlg.type_keys(_escape_for_type_keys(file_path),
                      with_spaces=True, pause=0)
        dlg.type_keys("{ENTER}")
        log.info(f"Файл выбран через Explorer")
    except Exception as e:
        log.error(f"Ошибка pywinauto: {e}")
        return

    # Пауза после закрытия Explorer — даём АСУД серверный upload + отрисовку confirm
    time.sleep(2)

    # Поллинг: ищем видимую кнопку "Присоединить" (или по SetContent-id) до 20s.
    # Не observer — чтобы не кликать раньше чем АСУД закончит upload.
    end = time.time() + 20
    while time.time() < end:
        # 1) По id (приоритет)
        try:
            btns = driver.find_elements(By.CSS_SELECTOR,
                "[id*='SetContent'], [id*='Send'], [id*='Submit']")
            for b in btns:
                if b.is_displayed():
                    click(driver, b, f"confirm by-id {b.get_attribute('id')}")
                    log.info("Файл присоединён!")
                    return
        except Exception:
            pass
        # 2) По тексту "Присоединить" — последняя видимая (не та что открывала Explorer)
        try:
            btns = driver.find_elements(By.XPATH,
                "//button[contains(text(),'Присоединить')] | //div[contains(text(),'Присоединить')]")
            visible = [b for b in btns if b.is_displayed()]
            if len(visible) >= 1:
                # Если она одна и совпадает с первой кнопкой — это значит модалка
                # ещё не открылась (та же кнопка). Проверим что у нас НОВАЯ кнопка
                # появилась относительно момента до Enter — берём последнюю.
                target = visible[-1]
                click(driver, target, "confirm by-text 'Присоединить'")
                log.info("Файл присоединён!")
                return
        except Exception:
            pass
        time.sleep(0.5)

    # Не нашли — диагностика
    log.warning("Confirm-кнопка не найдена за 20s — диагностика:")
    try:
        candidates = driver.execute_script(_DIAG_BUTTONS_JS) or []
        log.warning(f"  Видимых кандидатов: {len(candidates)}")
        for c in candidates[:10]:
            log.warning(f"    - {c}")
    except Exception:
        pass
    log.error("Файл НЕ присоединён")
