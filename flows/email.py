"""
flows/email.py — Email-direct: создание Входящих документов прямо из .msg файлов
в указанной папке (без xlsx-реестра).

Парсит каждый .msg через extract_msg, извлекает Subject/Body/Date,
ищет ФИО абонента в теле через extract_fio_from_text, прикрепляет сам же
.msg как content. Дальнейший flow (форма, регистрация, На резолюцию,
output xlsx) переиспользуется из flows/mix.py.

Запускается через app.py с --mode=email.
"""

import os
import re
import sys
import time
import logging
from datetime import date, datetime, timedelta

import openpyxl
import extract_msg
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService

from shared import config as cfg
from shared.ui import wait_asud_loaded
from shared.correspondent import extract_fio_from_text

# Переиспользуем создание/регистрацию/output из mix
from flows import mix as mix_flow

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger("asud")
start_time = time.monotonic()

settings = {}


# ================= EMAIL → DOC_DATA =================

def _msg_link(msg, file_path):
    """Генерирует Link для doc_data.

    Приоритет:
      1) Дата письма из .msg (msg.date)
      2) Имя файла без расширения
    Формат — как в mix-flow реестре: 'DD.MM.YYYY HH-MM-SS'.
    """
    try:
        if msg.date:
            return msg.date.strftime("%d.%m.%Y %H-%M-%S")
    except Exception:
        pass
    return os.path.splitext(os.path.basename(file_path))[0]


def load_emails(folder_path):
    """Берёт .msg ТОЛЬКО из корня folder_path (без рекурсии в подпапки).
    Подпапки типа 'Завершено' не должны попадать в работу.

    Возвращает list of dicts с теми же полями что mix.load_excel.
    """
    unknown = settings.get("unknown_correspondent",
                            cfg.DEFAULTS["unknown_correspondent"])
    type_idx = settings.get("default_type_idx", 8)
    type_name = cfg.DOC_TYPE_MAP.get(
        type_idx, "Письма, заявления и жалобы граждан, акционеров")

    msg_files = []
    try:
        for f in os.listdir(folder_path):
            full = os.path.join(folder_path, f)
            if os.path.isfile(full) and f.lower().endswith('.msg'):
                msg_files.append(full)
    except OSError as e:
        log.error(f"Не могу прочитать папку {folder_path}: {e}")
        return []

    log.info(f"Найдено .msg файлов: {len(msg_files)}")

    rows, skipped = [], 0
    for idx, msg_path in enumerate(sorted(msg_files), 1):
        try:
            msg = extract_msg.openMsg(msg_path)
            subject = msg.subject or ""
            body = msg.body or ""
            link = _msg_link(msg, msg_path)
            try:
                msg.close()
            except Exception:
                pass
        except Exception as e:
            log.warning(f"Не удалось прочитать {os.path.basename(msg_path)}: {e}")
            skipped += 1
            continue

        if not body and not subject:
            log.warning(f"Пустое письмо {os.path.basename(msg_path)} — пропускаю")
            skipped += 1
            continue

        # Очищаем subject от FW:/RE:/Fwd:
        clean_subject = re.sub(r'^(FW:|RE:|Fwd:)\s*', '',
                                str(subject).strip(), flags=re.IGNORECASE)
        # Body — очищаем тем же helper'ом что в mix
        body_clean = mix_flow._clean_body(body) if body else clean_subject

        # ФИО абонента из тела письма
        fio, fio_src = extract_fio_from_text(body_clean)
        if fio:
            corr_found = True
            correspondent = fio
        else:
            corr_found = False
            correspondent = unknown

        rows.append({
            "row_idx": idx,
            "содержание": body_clean,
            "корреспондент": correspondent,
            "корр_найден": corr_found,
            "корр_источник": fio_src,
            "тема": clean_subject,
            "тип_индекс": type_idx,
            "тип_название": type_name,
            "link": link,
            "файл": msg_path,  # сам .msg прикрепляется как content
        })

    log.info(f"Загружено: {len(rows)} писем, пропущено: {skipped}")
    return rows


# ================= MAIN =================

def main():
    global settings
    settings = cfg.load()
    cfg.setup_file_logger("email")

    log.info("=" * 50)
    log.info("АСУД ИК — Email-direct (создание из .msg-писем)")
    log.info("=" * 50)

    base_dir = cfg.get_base_dir()

    # Запрос пути к папке с .msg
    folder = os.environ.get('ASUD_EMAIL_FOLDER')
    if not folder:
        default = settings.get("email_folder", "")
        print(f"\nПапка с .msg-письмами (только из корня папки, подпапки игнорируются).")
        if default:
            print(f"Enter — использовать: {default}")
        user_dir = input("Путь: ").strip().strip('"').strip("'")
        folder = user_dir or default

    if not folder or not os.path.isdir(folder):
        log.error(f"Папка не найдена: {folder!r}")
        input("Enter...")
        sys.exit(1)

    log.info(f"Папка писем: {folder}")

    # Парсим письма
    docs = load_emails(folder)
    if not docs:
        log.error("Нет .msg файлов или все пропущены")
        input("Enter...")
        sys.exit(1)

    # Превью
    known = sum(1 for d in docs if d["корр_найден"])
    unknown_n = len(docs) - known
    print(f"\nПервые 5:")
    for i, d in enumerate(docs[:5], 1):
        flag = 'OK' if d["корр_найден"] else '!!'
        print(f"  {i}. [{d['тип_индекс']}] {flag} {d['корреспондент'][:30]} | {d['тема'][:50]}")
    print(f"\nВсего: {len(docs)}  (ФИО: {known}, заглушка: {unknown_n})")
    print("режим: EMAIL  —  создание + регистрация + На резолюцию + сам .msg как вложение")

    if input("Начать? (да/нет): ").strip().lower() not in ("да", "д", "y", "yes", ""):
        sys.exit(0)

    # === Запуск браузера и обработки (повторяем mix-loop, но с нашими docs)
    driver_path = os.path.join(base_dir, "msedgedriver.exe")
    if not os.path.exists(driver_path):
        log.error(f"msedgedriver.exe не найден в {base_dir}")
        input("Enter...")
        sys.exit(1)

    options = cfg.build_edge_options()
    service = EdgeService(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)

    # Настраиваем mix_flow.settings — он использует module-level global
    mix_flow.settings = settings

    try:
        url = settings.get("asud_url", cfg.DEFAULTS["asud_url"])
        log.info(f"Открываю {url}")
        driver.get(url)
        wait_asud_loaded(driver)

        # Output xlsx — имя из имени папки писем
        folder_name = os.path.basename(os.path.normpath(folder)) or "email"
        registered_dir = os.path.join(base_dir, "Registered")
        os.makedirs(registered_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(registered_dir,
            f"{folder_name}_{ts}_резолюции.xlsx")
        mix_flow._ensure_output_xlsx(output_path)
        log.info(f"Output xlsx: {output_path}")

        done_count, err_count = 0, 0
        for i, doc in enumerate(docs, 1):
            try:
                asud_id = mix_flow.create_one_document(driver, doc, i, len(docs))
                mix_flow._append_output_row(output_path, doc, asud_id)
                done_count += 1
            except Exception as e:
                log.error(f"ОШИБКА документ {i}: {e}")
                err_count += 1
                try:
                    driver.get(url)
                    wait_asud_loaded(driver)
                except Exception:
                    pass
                continue

        elapsed_seconds = time.monotonic() - start_time
        elapsed = timedelta(seconds=int(elapsed_seconds))
        avg = (timedelta(seconds=int(elapsed_seconds / done_count))
               if done_count else None)
        summary = [
            "",
            "=" * 60,
            "ГОТОВО!",
            f"  Обработано:   {done_count} / {len(docs)}",
            f"  Ошибок:       {err_count}",
            f"  В черновиках: {unknown_n}",
            f"  Затрачено:    {elapsed}" + (f"  (в среднем {avg}/док)" if avg else ""),
            "=" * 60,
        ]
        for line in summary:
            log.info(line)
            print(line)
        input("\nEnter для закрытия...")
    except Exception as e:
        log.error(f"Ошибка: {e}")
        input("Enter...")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
