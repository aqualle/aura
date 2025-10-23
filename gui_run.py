# gui_run.py - ИСПРАВЛЕНО

import tkinter as tk
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def main():
    print("🚀 Запуск GUI парсера Microsoft Edge...")
    
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, scrolledtext
        print("✅ GUI модули загружены")
    except ImportError as e:
        print(f"❌ Ошибка импорта GUI: {e}")
        print("🔧 Установите tkinter: apt-get install python3-tk (Linux)")
        return
    
    # Расширенная проверка наличия msedgedriver (PATH/ENV/локально)
    def _check_driver():
        candidates = [
            os.path.join("browserdriver", "msedgedriver"),
            os.path.join("browserdriver", "msedgedriver.exe"),
        ]
        env_driver = os.environ.get("MSEDGEDRIVER")
        if env_driver and os.path.exists(env_driver):
            return env_driver
        try:
            import shutil as _shutil
            which_path = _shutil.which("msedgedriver")
            if which_path:
                return which_path
        except Exception:
            pass
        for c in candidates:
            if os.path.exists(c):
                return c
        return None

    driver_found = _check_driver()
    if driver_found:
        print(f"✅ Edge WebDriver найден: {driver_found}")
    else:
        print("❌ Edge WebDriver НЕ НАЙДЕН (PATH/$MSEDGEDRIVER/./browserdriver)")
    
    cookies_file = os.path.expanduser("~/.yandex_parser_auth/cookies.json")
    if os.path.exists(cookies_file):
        print("✅ Cookies для юрлиц найдены")
    else:
        print("❌ Cookies для юрлиц не найдены (опционально)")
    
    try:
        from concurrent.futures import ThreadPoolExecutor
        print("✅ Многопоточность поддерживается")
    except ImportError:
        print("❌ Многопоточность не поддерживается")
    
    try:
        # ИСПРАВЛЕНИЕ: импортируем ParserGUI вместо YandexMarketGUI
        from gui_parser import ParserGUI
        print("✅ GUI парсер загружен")
    except ImportError as e:
        print(f"❌ Ошибка импорта gui_parser: {e}")
        print("📁 Убедитесь что файл gui_parser.py в той же папке")
        return
    
    root = tk.Tk()
    app = ParserGUI(root)
    print("🌐 GUI запущен для Edge парсера")
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\n❌ Программа прервана пользователем")
    except Exception as e:
        print(f"❌ Ошибка GUI: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("✅ GUI завершен")

if __name__ == "__main__":
    main()
