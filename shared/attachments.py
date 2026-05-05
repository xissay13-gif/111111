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


_DIAG_FILE_INPUTS_JS = r"""
// Диагностика для фонового режима: переписать все <input type="file"> в DOM
// со всеми атрибутами и контекстом. Помогает понять есть ли скрытый input
// который мы можем использовать вместо pywinauto + Explorer.
const out = [];
for (const inp of document.querySelectorAll("input[type='file']")) {
    const r = inp.getBoundingClientRect();
    const cs = window.getComputedStyle(inp);
    const parentTag = inp.parentElement ? inp.parentElement.tagName : null;
    const parentCls = inp.parentElement ? (inp.parentElement.className || '').toString().slice(0, 80) : null;
    out.push({
        id: inp.id || null,
        name: inp.name || null,
        accept: inp.getAttribute('accept'),
        multiple: inp.multiple,
        disabled: inp.disabled,
        offsetParent: !!inp.offsetParent,  // true если отображается
        rect: {w: r.width, h: r.height, x: r.left, y: r.top},
        display: cs.display,
        visibility: cs.visibility,
        opacity: cs.opacity,
        position: cs.position,
        parent_tag: parentTag,
        parent_class: parentCls,
        outerHTML_preview: inp.outerHTML.slice(0, 200),
    });
}
return {count: out.length, inputs: out};
"""


def _diag_file_inputs(driver, label):
    """Дампит в лог состояние всех <input type='file'> на странице."""
    try:
        diag = driver.execute_script(_DIAG_FILE_INPUTS_JS) or {}
        log.info(f"[diag {label}] input[type=file] count={diag.get('count', 0)}")
        for i, inp in enumerate(diag.get('inputs', []) or []):
            log.info(f"  [diag {label}] input[{i}]: id={inp.get('id')!r} name={inp.get('name')!r} "
                     f"disabled={inp.get('disabled')} offsetParent={inp.get('offsetParent')} "
                     f"display={inp.get('display')} visibility={inp.get('visibility')} "
                     f"opacity={inp.get('opacity')}")
            log.info(f"  [diag {label}]    rect={inp.get('rect')} parent={inp.get('parent_tag')}.{inp.get('parent_class')}")
            log.info(f"  [diag {label}]    html: {inp.get('outerHTML_preview')!r}")
    except Exception as e:
        log.warning(f"[diag {label}] dump упал: {e}")


def _attach_via_input(driver, file_path):
    """Стратегия через CDP file chooser intercept + send_keys на скрытый input.

    Найдено диагностикой стадии 0:
      • АСУД использует GXT FileUploadButton — скрытый <input type='file'> с
        name='SetContentDialog' создаётся через ~0.5-0.8s после клика
        'Присоединить содержимое'.
      • input невидим (display:none, rect 0×0), но functional.

    Шаги:
      1. Включаем Page.setInterceptFileChooserDialog → native picker не открывается
      2. Кликаем 'Присоединить содержимое' → input создаётся в DOM, picker
         перехвачен CDP'ом и не показан
      3. Ждём появления input[name='SetContentDialog'] (до 3s)
      4. Делаем input visible через JS (чтобы send_keys прошёл visibility-check)
      5. input.send_keys(file_path) — Selenium через WebDriver-protocol устанавливает
         значение файлового input'а без участия native picker
      6. Выключаем intercept
      7. Ждём confirm-кнопку, кликаем

    Возвращает True если успех, False если что-то упало (caller fallback на pywinauto).
    """
    intercept_enabled = False
    try:
        # 1. Включаем перехват native file picker
        try:
            driver.execute_cdp_cmd("Page.setInterceptFileChooserDialog",
                                    {"enabled": True})
            intercept_enabled = True
            log.debug("CDP file chooser intercept включён")
        except Exception as e:
            log.warning(f"CDP intercept не доступен: {e}")
            return False

        # 2. Клик по 'Присоединить содержимое' — picker не откроется (intercept)
        try:
            btn = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH,
                    "//div[contains(text(),'Присоединить содержимое')]")))
            click(driver, btn, "Присоединить содержимое (CDP)")
        except Exception as e:
            log.error(f"Кнопка 'Присоединить содержимое' не найдена: {e}")
            return False

        # 3. Ждём появления input[name='SetContentDialog']
        end = time.monotonic() + 3
        inp = None
        while time.monotonic() < end:
            try:
                inp = driver.find_element(By.CSS_SELECTOR,
                    "input[type='file'][name='SetContentDialog']")
                break
            except Exception:
                pass
            time.sleep(0.1)
        if not inp:
            log.warning("input[name='SetContentDialog'] не появился за 3s")
            return False
        log.debug(f"input найден: id={inp.get_attribute('id')}")

        # 4. Делаем visible — иначе send_keys жалуется на visibility
        driver.execute_script("""
            var el = arguments[0];
            el.style.display='block';
            el.style.visibility='visible';
            el.style.opacity='1';
            el.style.position='absolute';
            el.style.left='0px';
            el.style.top='0px';
            el.style.width='1px';
            el.style.height='1px';
            el.removeAttribute('hidden');
            el.disabled = false;
        """, inp)

        # 5. send_keys устанавливает значение файлового input напрямую через
        # WebDriver-protocol (без участия native picker)
        inp.send_keys(file_path)
        # Триггерим change на случай если автоматический dispatch не сработал
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", inp)
        log.info(f"Файл отправлен через CDP-intercept + send_keys")
        return True

    except Exception as e:
        log.warning(f"CDP-стратегия упала: {e}")
        return False
    finally:
        # Выключаем intercept чтобы не мешать другим частям АСУД
        if intercept_enabled:
            try:
                driver.execute_cdp_cmd("Page.setInterceptFileChooserDialog",
                                        {"enabled": False})
            except Exception:
                pass


def _wait_confirm_and_click(driver, timeout=20):
    """Поллинг confirm-кнопки 'Присоединить' (после успешной загрузки файла)."""
    end = time.time() + timeout
    while time.time() < end:
        try:
            btns = driver.find_elements(By.CSS_SELECTOR,
                "[id*='SetContent'], [id*='Send'], [id*='Submit']")
            for b in btns:
                if b.is_displayed():
                    click(driver, b, f"confirm by-id {b.get_attribute('id')}")
                    log.info("Файл присоединён!")
                    return True
        except Exception:
            pass
        try:
            btns = driver.find_elements(By.XPATH,
                "//button[contains(text(),'Присоединить')] | //div[contains(text(),'Присоединить')]")
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                click(driver, visible[-1], "confirm by-text 'Присоединить'")
                log.info("Файл присоединён!")
                return True
        except Exception:
            pass
        time.sleep(0.5)

    log.warning("Confirm-кнопка не найдена за timeout — диагностика:")
    try:
        candidates = driver.execute_script(_DIAG_BUTTONS_JS) or []
        log.warning(f"  Видимых кандидатов: {len(candidates)}")
        for c in candidates[:10]:
            log.warning(f"    - {c}")
    except Exception:
        pass
    return False


def attach_content(driver, file_path):
    """Прикрепляет файл к карточке документа.

    Порядок стратегий:
      1) CDP intercept + send_keys на скрытый input — основной путь, работает
         без открытия Explorer'а, headless-совместим.
      2) pywinauto + native Explorer — fallback если CDP не сработал
         (не headless, требует interactive Windows session).
    """
    log.info(f"Прикрепление: {os.path.basename(file_path)}")

    # === Стратегия 1: CDP intercept + send_keys ===
    if _attach_via_input(driver, file_path):
        if _wait_confirm_and_click(driver, timeout=20):
            return
        log.warning("File загружен но confirm-кнопка не найдена — продолжаю")
        return

    # === Стратегия 2: pywinauto fallback ===
    log.info("CDP-стратегия не сработала, fallback в pywinauto")
    if not PYWINAUTO:
        log.error("pywinauto не установлен — невозможно прикрепить")
        return

    try:
        btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(text(),'Присоединить содержимое')]")))
        time.sleep(2)
        click(driver, btn, "Присоединить содержимое (pywinauto)")
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

    time.sleep(2)
    _wait_confirm_and_click(driver, timeout=20)
