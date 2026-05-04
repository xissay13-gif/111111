"""
app.py — Единая точка входа для АСУД-автоматизации.

Поддерживаемые режимы:
  • mix         — создание входящих документов с прикреплением .msg по Link
  • auto-create — создание входящих документов без поиска .msg (используется пустышка)
  • resolutions — выдача резолюций по уже зарегистрированным документам

Запуск:
  python app.py                  # auto-detect режима по xlsx
  python app.py --mode=mix
  python app.py --mode=auto-create
  python app.py --mode=resolutions
  python app.py --xlsx=path.xlsx --mode=...

Auto-detect:
  • Имя файла оканчивается на _резолюции.xlsx                 → resolutions
  • Лист содержит колонку 'Link' (старый mix-формат)          → mix
  • Лист 'результат' (новый формат) или Subject/корреспондент → auto-create
"""

import argparse
import logging
import os
import sys

import openpyxl

from shared import config as cfg

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger("asud.app")


def detect_mode(xlsx_path):
    """Определяет режим по структуре xlsx."""
    name = os.path.basename(xlsx_path).lower()
    if name.endswith('_резолюции.xlsx') or 'резолюци' in name:
        return 'resolutions'

    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
        # Если есть лист 'результат' — это новый auto-create формат
        if 'результат' in wb.sheetnames:
            wb.close()
            return 'auto-create'
        # Иначе смотрим заголовки активного листа
        ws = wb.active
        headers = [str(c.value or '').lower()
                   for c in next(ws.iter_rows(max_row=1))]
        wb.close()
        # Колонка 'link' в заголовках → mix-flow (есть .msg по ссылке)
        if any('link' in h for h in headers):
            return 'mix'
        # Колонки ОПТС/ОРТС → resolutions
        if any('опт' in h or 'орт' in h for h in headers):
            return 'resolutions'
    except Exception as e:
        log.warning(f"Не удалось прочитать {xlsx_path} для auto-detect: {e}")

    return 'auto-create'


def main():
    parser = argparse.ArgumentParser(
        description="АСУД ИК — автоматизация документооборота")
    parser.add_argument('--mode', choices=['mix', 'auto-create', 'resolutions'],
                        help="Режим работы (если не задан — auto-detect по xlsx)")
    parser.add_argument('--xlsx', help="Путь к реестру (если не задан — спрашиваем)")
    args = parser.parse_args()

    base_dir = cfg.get_base_dir()
    log.info(f"Базовая папка: {base_dir}")

    # Если xlsx не указан — выбираем интерактивно
    xlsx_path = args.xlsx
    if not xlsx_path:
        # Сначала пробуем выбрать из base_dir
        candidates = [f for f in os.listdir(base_dir)
                      if f.lower().endswith('.xlsx') and not f.startswith('~')]
        if len(candidates) == 1:
            xlsx_path = os.path.join(base_dir, candidates[0])
            log.info(f"Найден единственный xlsx: {candidates[0]}")
        elif len(candidates) > 1:
            print("\nДоступные реестры:")
            for i, name in enumerate(candidates, 1):
                print(f"  {i}. {name}")
            choice = input(f"Выбери (1-{len(candidates)}): ").strip()
            try:
                xlsx_path = os.path.join(base_dir, candidates[int(choice) - 1])
            except (ValueError, IndexError):
                log.error("Неверный выбор")
                sys.exit(1)
        else:
            log.error(f"Не нашёл .xlsx в {base_dir}")
            sys.exit(1)

    # Определяем режим
    mode = args.mode or detect_mode(xlsx_path)
    log.info(f"Режим: {mode}  (xlsx: {os.path.basename(xlsx_path)})")

    # Запуск соответствующего flow
    if mode == 'mix':
        from flows.mix import main as flow_main
    elif mode == 'auto-create':
        from flows.auto_create import main as flow_main
    elif mode == 'resolutions':
        from flows.resolutions import main as flow_main
    else:
        log.error(f"Неизвестный режим: {mode}")
        sys.exit(1)

    # Каждый flow.main() сам читает xlsx — передаём через env-переменную
    # (минимальный contract без переписывания main каждого flow)
    os.environ['ASUD_XLSX'] = xlsx_path
    flow_main()


if __name__ == "__main__":
    main()
