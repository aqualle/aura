# gui_parser.py - упрощенный GUI с кнопкой cookies

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
import time
import os
import sys
from typing import Dict, List
import pandas as pd
from datetime import datetime

try:
    from tender_parser import get_prices
    from utils import extract_products_from_excel, save_results_into_tender_format
except ImportError as e:
    print(f"Ошибка импорта: {e}")
    sys.exit(1)


class ParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер Яндекс.Маркет (Microsoft Edge)")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        self.is_parsing = False
        self.current_thread = None
        self.products_data = []
        self.results_data = []
        self.queue = queue.Queue()
        self.auto_save_counter = 0
        
        self.input_file = tk.StringVar(value="tender_list.xlsx")
        
        app_dir = os.path.dirname(os.path.abspath(__file__))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        auto_output = os.path.join(app_dir, f"results_{timestamp}.xlsx")
        self.output_file = tk.StringVar(value=auto_output)
        
        self.headless_mode = tk.BooleanVar(value=False)
        self.driver_path = tk.StringVar(value="")
        self.auto_save_enabled = tk.BooleanVar(value=True)
        
        # Cookies
        self.cookies_file = os.path.join(app_dir, ".yandex_parser_auth", "cookies.json")
        self.has_cookies = os.path.exists(self.cookies_file)
        
        self.create_widgets()
        self.process_queue()
        
        self.log_message(f"GUI загружен", "INFO")
        self.log_message(f"Результаты: {auto_output}", "INFO")
        if self.has_cookies:
            self.log_message(f"Cookies найдены", "SUCCESS")
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Настройки
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки", padding=10)
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Файлы
        files_frame = ttk.Frame(settings_frame)
        files_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(files_frame, text="Входной файл:").pack(anchor=tk.W)
        input_frame = ttk.Frame(files_frame)
        input_frame.pack(fill=tk.X, pady=(2, 5))
        ttk.Entry(input_frame, textvariable=self.input_file, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(input_frame, text="Обзор", width=8, command=self.browse_input_file).pack(side=tk.RIGHT, padx=(5, 0))
        
        ttk.Label(files_frame, text="Выходной файл:").pack(anchor=tk.W)
        output_frame = ttk.Frame(files_frame)
        output_frame.pack(fill=tk.X, pady=(2, 5))
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_file, width=50, state="readonly")
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="Обзор", width=8, command=self.browse_output_file).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Опции
        options_frame = ttk.Frame(settings_frame)
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Checkbutton(options_frame, text="Скрытый режим браузера",
                       variable=self.headless_mode).pack(anchor=tk.W)
        ttk.Checkbutton(options_frame, text="Автосохранение каждые 3 товара", 
                       variable=self.auto_save_enabled).pack(anchor=tk.W, pady=(5, 0))
        
        # Cookies
        cookies_frame = ttk.Frame(settings_frame)
        cookies_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(cookies_frame, text="Cookies для цен юридических лиц:", 
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        
        self.cookies_status_label = ttk.Label(cookies_frame, 
                                              text="Не загружены" if not self.has_cookies else "Загружены",
                                              foreground="red" if not self.has_cookies else "green")
        self.cookies_status_label.pack(side=tk.LEFT, padx=(10, 20))
        
        ttk.Button(cookies_frame, text="Загрузить Cookies", 
                  command=self.load_cookies).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(cookies_frame, text="Очистить", 
                  command=self.clear_cookies).pack(side=tk.LEFT)
        
        # Кнопки управления
        controls_frame = ttk.Frame(settings_frame)
        controls_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.start_button = ttk.Button(controls_frame, text="Запустить парсинг", 
                                      command=self.start_parsing)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(controls_frame, text="Остановить", 
                                     command=self.stop_parsing, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_button = ttk.Button(controls_frame, text="Сохранить", 
                                     command=self.save_results_now)
        self.save_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(controls_frame, text="Очистить", 
                  command=self.clear_results).pack(side=tk.LEFT)
        
        # Статистика
        stats_frame = ttk.LabelFrame(main_frame, text="Статистика", padding=10)
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        stats_grid = ttk.Frame(stats_frame)
        stats_grid.pack(fill=tk.X)
        
        self.total_label = ttk.Label(stats_grid, text="Всего: 0", font=("Arial", 10, "bold"))
        self.total_label.grid(row=0, column=0, padx=(0, 15), sticky=tk.W)
        
        self.processed_label = ttk.Label(stats_grid, text="Обработано: 0", foreground="blue")
        self.processed_label.grid(row=0, column=1, padx=(0, 15), sticky=tk.W)
        
        self.success_label = ttk.Label(stats_grid, text="Успешно: 0", foreground="green")
        self.success_label.grid(row=0, column=2, padx=(0, 15), sticky=tk.W)
        
        self.error_label = ttk.Label(stats_grid, text="Ошибки: 0", foreground="red")
        self.error_label.grid(row=0, column=3, padx=(0, 15), sticky=tk.W)
        
        self.autosave_label = ttk.Label(stats_grid, text="Сохранений: 0", foreground="orange")
        self.autosave_label.grid(row=0, column=4, sticky=tk.W)
        
        self.progress = ttk.Progressbar(stats_grid, mode='determinate')
        self.progress.grid(row=1, column=0, columnspan=5, sticky=(tk.W, tk.E), pady=(10, 0))
        stats_grid.columnconfigure(0, weight=1)
        
        # Результаты
        results_frame = ttk.LabelFrame(main_frame, text="Результаты", padding=5)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        columns = ("№", "Название", "Цена", "Статус", "Ссылка")
        self.tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=10)
        
        self.tree.heading("№", text="№")
        self.tree.heading("Название", text="Название товара")
        self.tree.heading("Цена", text="Цена")
        self.tree.heading("Статус", text="Статус")
        self.tree.heading("Ссылка", text="Ссылка")
        
        self.tree.column("№", width=50, anchor=tk.CENTER)
        self.tree.column("Название", width=400, anchor=tk.W)
        self.tree.column("Цена", width=120, anchor=tk.CENTER)
        self.tree.column("Статус", width=100, anchor=tk.CENTER)
        self.tree.column("Ссылка", width=80, anchor=tk.CENTER)
        
        v_scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)
        
        self.tree.bind("<Double-1>", self.open_link)
        
        # Лог
        log_frame = ttk.LabelFrame(main_frame, text="Лог", padding=5)
        log_frame.pack(fill=tk.X)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("SUCCESS", foreground="green", font=("Arial", 9, "bold"))
        self.log_text.tag_config("ERROR", foreground="red", font=("Arial", 9, "bold"))
        self.log_text.tag_config("WARNING", foreground="orange", font=("Arial", 9, "bold"))
    
    def load_cookies(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл cookies",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                import shutil
                os.makedirs(os.path.dirname(self.cookies_file), exist_ok=True)
                shutil.copy2(filename, self.cookies_file)
                self.has_cookies = True
                self.cookies_status_label.config(text="Загружены", foreground="green")
                self.log_message(f"Cookies загружены: {filename}", "SUCCESS")
            except Exception as e:
                self.log_message(f"Ошибка загрузки cookies: {e}", "ERROR")
    
    def clear_cookies(self):
        try:
            if os.path.exists(self.cookies_file):
                os.remove(self.cookies_file)
            self.has_cookies = False
            self.cookies_status_label.config(text="Не загружены", foreground="red")
            self.log_message("Cookies удалены", "INFO")
        except Exception as e:
            self.log_message(f"Ошибка удаления cookies: {e}", "ERROR")
    
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите входной Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
    
    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="Выберите место сохранения",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
    
    def log_message(self, message: str, level: str = "INFO"):
        timestamp = time.strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, formatted_message, level)
        self.log_text.see(tk.END)
    
    def update_stats(self):
        total = len(self.results_data)
        processed = len([r for r in self.results_data if r.get("status") in ["success", "error", "not_found"]])
        success = len([r for r in self.results_data if r.get("status") == "success"])
        error = len([r for r in self.results_data if r.get("status") == "error"])
        
        self.total_label.config(text=f"Всего: {total}")
        self.processed_label.config(text=f"Обработано: {processed}")
        self.success_label.config(text=f"Успешно: {success}")
        self.error_label.config(text=f"Ошибки: {error}")
        self.autosave_label.config(text=f"Сохранений: {self.auto_save_counter}")
        
        if total > 0:
            progress_value = (processed / total) * 100
            self.progress['value'] = progress_value
    
    def add_result_row(self, index: int, product_name: str, price: str = "—",
                      status: str = "pending", url: str = ""):
        while len(self.results_data) <= index:
            self.results_data.append({})
        
        self.results_data[index].update({
            "name": product_name,
            "price": price,
            "status": status,
            "url": url
        })
        
        status_indicators = {
            "pending": "⏳",
            "processing": "🔄", 
            "success": "✅",
            "error": "❌",
            "not_found": "❓"
        }
        
        status_text = status_indicators.get(status, "❓")
        
        item_id = f"item_{index}"
        if self.tree.exists(item_id):
            self.tree.item(item_id, values=(
                index + 1,
                product_name[:60] + ("..." if len(product_name) > 60 else ""),
                price,
                status_text,
                "🔗" if url else ""
            ))
        else:
            self.tree.insert("", "end", iid=item_id, values=(
                index + 1,
                product_name[:60] + ("..." if len(product_name) > 60 else ""),
                price,
                status_text,
                "🔗" if url else ""
            ))
        
        self.tree.see(item_id)
        self.update_stats()
    
    def open_link(self, event):
        selection = self.tree.selection()
        if not selection:
            return
            
        item = selection[0]
        index = int(item.split("_")[1])
        
        if index < len(self.results_data) and self.results_data[index].get("url"):
            url = self.results_data[index]["url"]
            import webbrowser
            webbrowser.open(url)
            self.log_message(f"Открыта ссылка для товара #{index + 1}", "INFO")
    
    def clear_results(self):
        if self.is_parsing:
            messagebox.showwarning("Внимание", "Остановите парсинг перед очисткой")
            return
            
        self.tree.delete(*self.tree.get_children())
        self.results_data.clear()
        self.log_text.delete(1.0, tk.END)
        self.update_stats()
        self.progress['value'] = 0
        self.auto_save_counter = 0
        self.log_message("Результаты очищены", "INFO")
    
    def save_results_now(self):
        if not self.results_data:
            messagebox.showwarning("Внимание", "Нет данных для сохранения")
            return
        
        self.perform_save()
    
    def perform_save(self):
        try:
            df_data = []
            for result in self.results_data:
                df_data.append({
                    "наименование": result.get("name", ""),
                    "цена": result.get("price", "—"),
                    "цена для юрлиц": "",
                    "ссылка": result.get("url", "")
                })
            
            df = pd.DataFrame(df_data)
            
            input_path = self.input_file.get()
            output_path = self.output_file.get()
            
            success = save_results_into_tender_format(input_path, output_path, df)
            
            if success:
                self.auto_save_counter += 1
                self.log_message(f"Сохранено: {output_path}", "SUCCESS")
                self.update_stats()
            
        except Exception as e:
            error_msg = f"Ошибка сохранения: {e}"
            self.log_message(error_msg, "ERROR")
            messagebox.showerror("Ошибка сохранения", error_msg)
    
    def start_parsing(self):
        if self.is_parsing:
            return
        
        input_path = self.input_file.get()
        if not os.path.exists(input_path):
            messagebox.showerror("Ошибка", f"Файл не найден: {input_path}")
            return
        
        self.clear_results()
        
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.is_parsing = True
        
        self.current_thread = threading.Thread(target=self.parse_worker, daemon=True)
        self.current_thread.start()
    
    def stop_parsing(self):
        if not self.is_parsing:
            return
            
        self.is_parsing = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.log_message("Парсинг остановлен", "WARNING")
        
        if self.results_data:
            self.perform_save()
    
    def parse_worker(self):
        try:
            self.queue.put(("log", "Начинаем парсинг...", "INFO"))
            
            input_path = self.input_file.get()
            products_df = extract_products_from_excel(input_path)
            
            if products_df.empty:
                self.queue.put(("log", "В файле не найдено товаров", "ERROR"))
                return
            
            products_list = products_df["name"].tolist()
            self.queue.put(("log", f"Найдено товаров: {len(products_list)}", "INFO"))
            
            for i, product_name in enumerate(products_list):
                self.queue.put(("add_row", i, product_name, "—", "pending", ""))
            
            for i, product_name in enumerate(products_list):
                if not self.is_parsing:
                    break
                
                self.queue.put(("update_row", i, product_name, "—", "processing", ""))
                self.queue.put(("log", f"{i + 1}/{len(products_list)}: {product_name[:40]}...", "INFO"))
                
                try:
                    result = get_prices(
                        product_name=product_name,
                        headless=self.headless_mode.get(),
                        driver_path=self.driver_path.get() if self.driver_path.get() else None,
                        timeout=20,
                        use_business_auth=self.has_cookies
                    )
                    
                    price = result.get("цена", "—")
                    url = result.get("ссылка", "")
                    
                    status = "success" if price not in ["—", "ERR", "ОШИБКА"] else "not_found"
                    if price in ["ERR", "ОШИБКА"]:
                        status = "error"
                    
                    self.queue.put(("update_row", i, product_name, price, status, url))
                    
                    if status == "success":
                        self.queue.put(("log", f"{product_name[:30]}... -> {price}", "SUCCESS"))
                    else:
                        self.queue.put(("log", f"{product_name[:30]}... -> не найден", "WARNING"))
                    
                    if self.auto_save_enabled.get() and (i + 1) % 3 == 0:
                        self.queue.put(("auto_save",))
                        self.queue.put(("log", f"Автосохранение {i + 1}/{len(products_list)}", "INFO"))
                        
                except Exception as e:
                    error_msg = f"Ошибка {product_name[:30]}...: {str(e)[:100]}"
                    self.queue.put(("update_row", i, product_name, "ERR", "error", ""))
                    self.queue.put(("log", error_msg, "ERROR"))
                
                time.sleep(1)
            
            if self.auto_save_enabled.get():
                self.queue.put(("auto_save",))
            
            self.queue.put(("log", "Парсинг завершен!", "SUCCESS"))
            
        except Exception as e:
            error_msg = f"Критическая ошибка: {e}"
            self.queue.put(("log", error_msg, "ERROR"))
        finally:
            self.queue.put(("parsing_finished",))
    
    def process_queue(self):
        try:
            while True:
                try:
                    message = self.queue.get_nowait()
                    action = message[0]
                    
                    if action == "log":
                        _, text, level = message
                        self.log_message(text, level)
                        
                    elif action == "add_row":
                        _, index, name, price, status, url = message
                        self.add_result_row(index, name, price, status, url)
                        
                    elif action == "update_row":
                        _, index, name, price, status, url = message
                        self.add_result_row(index, name, price, status, url)
                        
                    elif action == "auto_save":
                        self.perform_save()
                    
                    elif action == "parsing_finished":
                        self.is_parsing = False
                        self.start_button.config(state=tk.NORMAL)
                        self.stop_button.config(state=tk.DISABLED)
                        
                except queue.Empty:
                    break
                    
        except Exception as e:
            print(f"Ошибка обработки очереди: {e}")
        
        self.root.after(100, self.process_queue)


def main():
    root = tk.Tk()
    app = ParserGUI(root)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\nПрограмма прервана")
        root.quit()


if __name__ == "__main__":
    main()
