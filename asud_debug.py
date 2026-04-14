"""
АСУД ИК — Debug-инструмент для сбора HTML.

Открывает АСУД в Edge, ждёт пока вы дойдёте до нужной формы/диалога,
по Enter сохраняет HTML страницы в файл рядом с exe.

Запуск: положите msedgedriver.exe рядом с asud_debug.exe и запустите.
"""

import sys
import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions


ASUD_URL = "https://asud.interrao.ru/asudik/"


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_driver_path():
    base_dir = get_base_dir()
    driver_path = os.path.join(base_dir, "msedgedriver.exe")
    if not os.path.exists(driver_path):
        print(f"!! msedgedriver.exe не найден в: {base_dir}")
        input("Enter для выхода...")
        sys.exit(1)
    return driver_path


def save_snapshot(driver, label=""):
    """Сохраняет HTML страницы + скриншот."""
    base_dir = get_base_dir()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = f"_{label}" if label else ""

    html_path = os.path.join(base_dir, f"snapshot_{ts}{suffix}.html")
    png_path = os.path.join(base_dir, f"snapshot_{ts}{suffix}.png")

    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print(f"  ОК HTML сохранён: {os.path.basename(html_path)}")
    except Exception as e:
        print(f"  !! Ошибка сохранения HTML: {e}")

    try:
        driver.save_screenshot(png_path)
        print(f"  ОК Скриншот сохранён: {os.path.basename(png_path)}")
    except Exception as e:
        print(f"  !! Ошибка скриншота: {e}")


def main():
    print("=" * 60)
    print("АСУД ИК — Debug снимок страницы")
    print("=" * 60)
    print()
    print("Порядок работы:")
    print("  1. Браузер откроет АСУД")
    print("  2. Вы руками дойдёте до нужного места (диалог, форма)")
    print("  3. В окне консоли введите метку (или просто Enter)")
    print("     → сохранится HTML + скриншот рядом с exe")
    print("  4. Повторяйте сколько нужно")
    print("  5. Введите 'q' для выхода")
    print()

    driver_path = get_driver_path()
    print(f"EdgeDriver: {driver_path}")

    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--auth-server-whitelist=*.interrao.ru")
    options.add_argument("--auth-negotiate-delegate-whitelist=*.interrao.ru")
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    service = EdgeService(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)

    try:
        print("\nОткрываю АСУД...")
        driver.get(ASUD_URL)
        time.sleep(2)
        print("  Браузер готов. Перейдите на нужную страницу.")
        print()

        while True:
            label = input("Метка (или 'q' для выхода): ").strip()
            if label.lower() in ("q", "quit", "exit", "выход"):
                break
            # Чистим метку от недопустимых символов в имени файла
            safe_label = "".join(c if c.isalnum() or c in "-_" else "_"
                                 for c in label)[:50]
            save_snapshot(driver, safe_label)
            print()

    except Exception as e:
        print(f"\n!! Ошибка: {e}")
        input("Enter для закрытия...")
    finally:
        driver.quit()
        print("\nБраузер закрыт.")


if __name__ == "__main__":
    main()
