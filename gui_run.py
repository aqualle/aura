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
    
    edge_driver_path = os.path.join("browserdriver", "msedgedriver.exe")
    if os.path.exists(edge_driver_path):
        print(f"✅ Edge WebDriver найден: {edge_driver_path}")
    else:
        print(f"❌ Edge WebDriver НЕ НАЙДЕН: {edge_driver_path}")
    
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
