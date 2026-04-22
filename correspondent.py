"""
correspondent.py — Создание нового корреспондента в АСУД.

7 шагов:
  1. Клик '+' у Корреспондент
  2. 'Добавить' в Поиске
  3. Поиск организации → 'Создать организацию'
  4. 'Добавить' в Физические лица
  5. Заполнить карточку (ФИО + Должность=ФЛ)
  6. 'Выбрать физ. лиц'
  7. 'Готово'
"""

import re
import time
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from ui import click, find_input_near_label, close_open_modals

log = logging.getLogger("asud.correspondent")


def fio_to_initials(full_name):
    """'Калганова Тамара Алексеевна' → 'Калганова Т А'"""
    parts = full_name.strip().split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]} {parts[2][0]}"
    elif len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}"
    return parts[0] if parts else full_name


def _norm_no_space(s):
    return s.replace('.', '').replace(',', '').replace(' ', '').replace('\xa0', '').lower()


def _norm_keep_space(s):
    import re as _re
    s = s.replace('\xa0', ' ').replace('.', '').replace(',', '')
    return _re.sub(r'\s+', ' ', s).strip().lower()


def match_correspondent(text, full_name):
    """Мягкий матч: полное ФИО / инициалы / фамилия. Для адресатов."""
    text_clean = text.strip()
    if full_name in text_clean:
        return True
    initials = fio_to_initials(full_name)
    if _norm_keep_space(initials) in _norm_keep_space(text_clean):
        return True
    if _norm_no_space(initials) in _norm_no_space(text_clean):
        return True
    if _norm_no_space(full_name) in _norm_no_space(text_clean):
        return True
    surname = full_name.split()[0]
    if text_clean.lower().startswith(surname.lower()):
        return True
    return False


def match_strict(text, full_name):
    """Строгий матч: только полное ФИО или инициалы. Для корреспондентов.
    Нормализует любые пробелы (включая NBSP) и точки/запятые."""
    text_clean = text.strip()
    if full_name in text_clean:
        return True
    if _norm_keep_space(full_name) in _norm_keep_space(text_clean):
        return True
    if _norm_no_space(full_name) in _norm_no_space(text_clean):
        return True
    initials = fio_to_initials(full_name)
    if _norm_keep_space(initials) in _norm_keep_space(text_clean):
        return True
    if _norm_no_space(initials) in _norm_no_space(text_clean):
        return True
    return False


def create_correspondent(driver, person_name):
    """Создаёт нового корреспондента через 7-шаговый flow."""
    parts = person_name.strip().split()
    if not parts:
        log.error("Пустое ФИО")
        return
    surname = parts[0]
    first_name = parts[1] if len(parts) >= 2 else "Н"
    middle_name = parts[2] if len(parts) >= 3 else "Н"
    if len(parts) < 3:
        log.warning(f"Неполное ФИО '{person_name}' → недостающие = 'Н'")

    log.info(f"Создаю корреспондента: {surname} {first_name} {middle_name}")

    # ШАГ 1: Клик "+" у Корреспондент
    log.info("[1/7] Клик '+' у Корреспондент")
    try:
        label = driver.find_element(By.XPATH,
            "//*[normalize-space(text())='Корреспондент']")
        parent = label
        plus_btn = None
        for _ in range(1, 7):
            parent = parent.find_element(By.XPATH, "..")
            btns = parent.find_elements(By.CSS_SELECTOR,
                "img[data-marker='select-btn'], img.gwt-Image")
            visible = [b for b in btns if b.is_displayed()]
            if visible:
                plus_btn = visible[-1]
                break
        if not plus_btn:
            log.error("Кнопка '+' не найдена")
            return
        click(driver, plus_btn, "+ Корреспондент")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'Поиск корреспондента')]")))
        except Exception:
            pass
    except Exception as e:
        log.error(f"Шаг 1: {e}")
        close_open_modals(driver)
        return

    # ШАГ 2: 'Добавить' в Поиске
    log.info("[2/7] 'Добавить' в Поиске корреспондента")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.XPATH, "//*[normalize-space(text())='Добавить']")))
        btns = driver.find_elements(By.XPATH, "//*[normalize-space(text())='Добавить']")
        add_btn = next((b for b in btns if b.is_displayed()), None)
        if not add_btn:
            log.error("Кнопка 'Добавить' не найдена")
            close_open_modals(driver)
            return
        click(driver, add_btn, "Добавить")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'Редактирование организации') or "
                 "contains(text(),'Поиск организации')]")))
        except Exception:
            pass
    except Exception as e:
        log.error(f"Шаг 2: {e}")
        close_open_modals(driver)
        return

    # ШАГ 3: Поиск организации → 'Создать организацию'
    log.info("[3/7] Поиск организации")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.XPATH, "//*[normalize-space(text())='Поиск организации']")))
        # Ввод через JS (атомарно, без stale)
        js_result = driver.execute_script("""
            var surname = arguments[0];
            var xpath = "//*[normalize-space(text())='Поиск организации']";
            var iter = document.evaluate(xpath, document, null,
                XPathResult.ORDERED_NODE_ITERATOR_TYPE, null);
            var labels = [], node;
            while ((node = iter.iterateNext()) !== null) {
                if (node.offsetParent !== null) labels.push(node);
            }
            if (!labels.length) return 'no-label';
            for (var li = 0; li < labels.length; li++) {
                var parent = labels[li];
                for (var lvl = 0; lvl < 6; lvl++) {
                    parent = parent.parentElement;
                    if (!parent) break;
                    var inputs = parent.querySelectorAll('input[type="text"]');
                    for (var i = 0; i < inputs.length; i++) {
                        var inp = inputs[i];
                        if (inp.offsetParent !== null && !inp.readOnly) {
                            inp.focus(); inp.value = surname;
                            inp.dispatchEvent(new Event('input', {bubbles:true}));
                            inp.dispatchEvent(new Event('keyup', {bubbles:true}));
                            inp.dispatchEvent(new Event('change', {bubbles:true}));
                            return 'ok';
                        }
                    }
                }
            }
            return 'no-input';
        """, surname)
        log.info(f"JS ввод: {js_result}")
        if js_result != 'ok':
            close_open_modals(driver)
            return

        # Ждём кнопку "Создать организацию"
        create_org_btn = None
        for _ in range(10):
            try:
                btn = driver.find_element(By.CSS_SELECTOR,
                    "[id*='create_custom_org'], [id*='custom_org_button']")
                if btn.is_displayed():
                    create_org_btn = btn
                    break
            except Exception:
                pass
            btns = driver.find_elements(By.XPATH,
                "//*[contains(text(),'Создать организацию')]")
            for b in btns:
                if b.is_displayed():
                    create_org_btn = b
                    break
            if create_org_btn:
                break
            time.sleep(1)

        if create_org_btn:
            click(driver, create_org_btn, "Создать организацию")
            time.sleep(1)
        else:
            log.warning("Кнопка 'Создать организацию' не найдена")
            close_open_modals(driver)
            return
    except Exception as e:
        log.error(f"Шаг 3: {e}")
        close_open_modals(driver)
        return

    # ШАГ 4: 'Добавить' в Физические лица
    log.info("[4/7] 'Добавить' в Физические лица")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(text(),'Физические лица')]")))
        time.sleep(3)
        add_user_btn = None
        for attempt in range(20):
            try:
                btn = None
                try:
                    btn = driver.find_element(By.CSS_SELECTOR,
                        "[id*='header-organization-dialog-add-a-user-button']")
                except Exception:
                    pass
                if not btn:
                    section = driver.find_element(By.XPATH,
                        "//*[contains(text(),'Физические лица')]")
                    parent = section
                    for _ in range(1, 6):
                        parent = parent.find_element(By.XPATH, "..")
                        bs = parent.find_elements(By.XPATH,
                            ".//*[normalize-space(text())='Добавить']")
                        vis = [b for b in bs if b.is_displayed()]
                        if vis:
                            btn = vis[0]
                            break
                if btn:
                    is_enabled = driver.execute_script("""
                        var el = arguments[0];
                        if (!el.offsetParent) return false;
                        if (el.getAttribute('aria-disabled')==='true') return false;
                        if (el.classList.contains('x-disabled')) return false;
                        var style = window.getComputedStyle(el);
                        if (style.pointerEvents==='none') return false;
                        if (parseFloat(style.opacity)<0.5) return false;
                        return true;
                    """, btn)
                    if is_enabled:
                        add_user_btn = btn
                        break
            except Exception:
                pass
            time.sleep(1)
        if not add_user_btn:
            log.error("Кнопка 'Добавить' не активировалась")
            close_open_modals(driver)
            return
        click(driver, add_user_btn, "Добавить физ. лицо")
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                (By.XPATH, "//*[normalize-space(text())='Фамилия']")))
        except Exception:
            pass
    except Exception as e:
        log.error(f"Шаг 4: {e}")
        close_open_modals(driver)
        return

    # ШАГ 5: Заполнить карточку
    log.info("[5/7] Заполнение карточки")
    try:
        fields = [
            ("Фамилия", surname, "outer_person_dialog-last_name-input"),
            ("Имя", first_name, "outer_person_dialog-first_name-input"),
            ("Отчество", middle_name, "outer_person_dialog-middle_name-input"),
            ("Должность", "ФЛ", "outer_person_dialog-position-input"),
        ]
        for label_text, value, input_id in fields:
            result = driver.execute_script("""
                var inputId = arguments[0]; var value = arguments[1];
                var el = document.getElementById(inputId);
                if (!el) {
                    var base = inputId.replace('-input','');
                    var container = document.getElementById(base);
                    if (container) {
                        var inputs = container.querySelectorAll('input[type="text"]');
                        for (var i=0;i<inputs.length;i++)
                            if (inputs[i].offsetParent!==null && !inputs[i].readOnly)
                                { el=inputs[i]; break; }
                    }
                }
                if (!el) return 'no-element';
                el.focus(); el.value = value;
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
                return 'ok:'+el.id;
            """, input_id, value)
            log.info(f"  {label_text}: {value} [{result}]")
            time.sleep(0.3)

        # Нажать "Добавить" в карточке
        save_btn = None
        try:
            save_btn = driver.find_element(By.CSS_SELECTOR,
                "[id*='Parton_person_dialog_save_button']")
        except Exception:
            btns = driver.find_elements(By.XPATH,
                "//*[normalize-space(text())='Добавить']")
            vis = [b for b in btns if b.is_displayed()]
            if vis:
                save_btn = vis[-1]
        if save_btn:
            click(driver, save_btn, "Сохранить карточку")
            try:
                WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(),'Выбрать физ')]")))
            except Exception:
                pass
    except Exception as e:
        log.error(f"Шаг 5: {e}")
        close_open_modals(driver)
        return

    # ШАГ 6: 'Выбрать физ. лиц'
    log.info("[6/7] Выбрать физ. лиц")
    try:
        select_btn = None
        try:
            select_btn = driver.find_element(By.CSS_SELECTOR,
                "[id*='Parton_organization_dialog_select_persons_button']")
        except Exception:
            btns = driver.find_elements(By.XPATH,
                "//*[contains(text(),'Выбрать физ')]")
            vis = [b for b in btns if b.is_displayed()]
            if vis:
                select_btn = vis[0]
        if select_btn:
            click(driver, select_btn, "Выбрать физ. лиц")
            try:
                WebDriverWait(driver, 15).until(EC.presence_of_element_located(
                    (By.ID, "oshs-select-button")))
            except Exception:
                pass
        else:
            log.error("Кнопка 'Выбрать физ. лиц' не найдена")
            close_open_modals(driver)
            return
    except Exception as e:
        log.error(f"Шаг 6: {e}")
        close_open_modals(driver)
        return

    # ШАГ 7: 'Готово'
    log.info("[7/7] Готово")
    try:
        done_btn = None
        try:
            done_btn = driver.find_element(By.ID, "oshs-select-button")
        except Exception:
            btns = driver.find_elements(By.XPATH,
                "//*[normalize-space(text())='Готово']")
            vis = [b for b in btns if b.is_displayed()]
            if vis:
                done_btn = vis[0]
        if done_btn:
            click(driver, done_btn, "Готово")
            from ui import wait_modal_closed
            wait_modal_closed(driver)
            log.info(f"Корреспондент создан: {person_name}")
        else:
            log.error("Кнопка 'Готово' не найдена")
            close_open_modals(driver)
    except Exception as e:
        log.error(f"Шаг 7: {e}")
        close_open_modals(driver)


def fill_correspondent_field(driver, person_name):
    """Заполняет поле Корреспондент через combobox.
    Если не найден по инициалам — создаёт нового."""
    log.info(f"Корреспондент: {person_name}")
    time.sleep(1)

    inp = find_input_near_label(driver, "Корреспондент")
    if not inp:
        log.warning("Поле корреспондента не найдено")
        return

    surname = person_name.split()[0]
    initials = fio_to_initials(person_name)

    inp.click()
    time.sleep(0.3)
    inp.clear()
    time.sleep(0.3)
    for char in surname:
        inp.send_keys(char)
        time.sleep(0.1)
    log.info(f"Введена фамилия: {surname}")
    time.sleep(2)

    def find_all():
        from selenium.webdriver.common.by import By as _By
        results = driver.find_elements(_By.XPATH, f"//*[contains(text(),'{surname}')]")
        return [r for r in results
                if r.is_displayed() and r != inp and r.tag_name.lower() != 'input']

    all_results = find_all()
    if not all_results:
        from selenium.webdriver.common.keys import Keys as _Keys
        inp.send_keys(_Keys.ENTER)
        time.sleep(3)
        all_results = find_all()

    log.info(f"Кандидатов: {len(all_results)} (ищем '{initials}')")

    # Строгий матч по инициалам
    target = None
    for idx, r in enumerate(all_results, 1):
        try:
            raw = r.text
            ok = match_strict(raw, person_name)
            preview = raw.strip().replace('\n', ' ')[:80]
            log.info(f"  [{idx}] {'OK' if ok else '--'} | {preview!r}")
            if ok and target is None:
                target = r
        except Exception as e:
            log.info(f"  [{idx}] ERR читаю text: {e}")
            continue

    if target:
        from selenium.webdriver.common.action_chains import ActionChains as _AC
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
        time.sleep(0.3)
        _AC(driver).move_to_element(target).pause(0.3).click().perform()
        time.sleep(1)
        log.info(f"Корреспондент выбран: {person_name}")
        return

    # Нет совпадения — создаём нового
    log.info(f"'{initials}' не найден — создаю нового")
    from selenium.webdriver.common.keys import Keys as _Keys
    try:
        inp.send_keys(_Keys.ESCAPE)
        time.sleep(0.5)
    except Exception:
        pass
    create_correspondent(driver, person_name)
