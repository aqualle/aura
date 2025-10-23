# tender_parser.py - ФИНАЛЬНАЯ версия с тендерным форматом

import time
import logging
import json
import re
import tempfile
import shutil
import uuid
import atexit
import signal
import os
from typing import Dict, Optional, List, Any, Tuple
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, WebDriverException
from utils import extract_products_from_excel, save_results_into_tender_format

# Настройка логирования
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# Глобальные переменные для управления парсингом и автосохранения
STOP_PARSING = False
CREATED_PROFILES = set()
CURRENT_DATAFRAME = None
CURRENT_OUTPUT_FILE = None
CURRENT_INPUT_FILE = None

def setup_signal_handlers():
    """Настройка обработчиков сигналов для автосохранения при завершении"""
    def signal_handler(signum, frame):
        global STOP_PARSING
        logger.info(f"Получен сигнал завершения ({signum}), выполняю автосохранение...")
        STOP_PARSING = True
        force_save_results()
        cleanup_profiles()
        logger.info("Автосохранение завершено, выход из программы")
        os._exit(0)

    # Обработчики для Windows и Unix
    try:
        signal.signal(signal.SIGINT, signal_handler)   # Ctrl+C
        signal.signal(signal.SIGTERM, signal_handler)  # Terminate
        if hasattr(signal, 'SIGBREAK'):  # Windows
            signal.signal(signal.SIGBREAK, signal_handler)
    except Exception as e:
        logger.warning(f"Не удалось установить обработчики сигналов: {e}")

def force_save_results():
    """Принудительное сохранение результатов при завершении"""
    global CURRENT_DATAFRAME, CURRENT_OUTPUT_FILE, CURRENT_INPUT_FILE

    if CURRENT_DATAFRAME is not None and CURRENT_OUTPUT_FILE and CURRENT_INPUT_FILE:
        try:
            # Считаем сколько товаров обработано
            processed = len([r for r in CURRENT_DATAFRAME['цена'] if r and r not in ['', 'ОШИБКА']])
            total = len(CURRENT_DATAFRAME)

            # ИСПОЛЬЗУЕМ НОВУЮ ФУНКЦИЮ ТЕНДЕРНОГО ФОРМАТА
            save_results_into_tender_format(CURRENT_INPUT_FILE, CURRENT_OUTPUT_FILE, CURRENT_DATAFRAME)
            logger.info(f"🚨 ЭКСТРЕННОЕ СОХРАНЕНИЕ ТЕНДЕРА: обработано {processed}/{total} товаров в {CURRENT_OUTPUT_FILE}")
        except Exception as e:
            logger.error(f"Ошибка экстренного сохранения: {e}")
    else:
        logger.info("Нет данных для экстренного сохранения")

def stop_all_parsing():
    """Останавливает все процессы парсинга"""
    global STOP_PARSING
    STOP_PARSING = True
    logger.info("Получен сигнал остановки парсинга")

def cleanup_single_profile(profile_path: str) -> bool:
    """Аккуратно очищает один профиль Edge после закрытия драйвера"""
    if not profile_path or not os.path.exists(profile_path):
        return False

    try:
        time.sleep(0.3)

        try:
            import psutil
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    if proc.info['name'] and 'msedge' in proc.info['name'].lower():
                        if proc.info['cmdline']:
                            cmdline = ' '.join(proc.info['cmdline'])
                            if profile_path in cmdline:
                                return False
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
        except ImportError:
            time.sleep(0.5)

        shutil.rmtree(profile_path, ignore_errors=True)
        success = not os.path.exists(profile_path)

        return success

    except Exception as e:
        return False

def cleanup_profiles():
    """Глобальная очистка всех профилей"""
    global CREATED_PROFILES
    cleanup_count = 0
    for profile_path in CREATED_PROFILES.copy():
        try:
            if os.path.exists(profile_path):
                shutil.rmtree(profile_path, ignore_errors=True)
                cleanup_count += 1
        except:
            pass
    CREATED_PROFILES.clear()
    if cleanup_count > 0:
        logger.info(f"Очищено {cleanup_count} профилей Edge")

atexit.register(cleanup_profiles)

def kill_zombie_edges():
    """Убивает Edge процессы"""
    print("Закрываю Edge процессы...")
    try:
        import psutil
        killed_count = 0
        for p in psutil.process_iter(['pid', 'name']):
            if p.info['name'] and 'msedge' in p.info['name'].lower():
                try:
                    p.terminate()
                    killed_count += 1
                except:
                    pass
        if killed_count > 0:
            print(f"Закрыто {killed_count} процессов")
    except:
        pass

def create_driver(headless: bool = True, driver_path: Optional[str] = None, use_auth: bool = False) -> webdriver.Edge:
    """Создание оптимизированного Edge драйвера"""
    global CREATED_PROFILES
    options = webdriver.EdgeOptions()

    # Оптимизации
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage") 
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-plugins")
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-default-apps")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3")
    options.add_argument("--silent")
    options.add_argument("--page-load-strategy=eager")
    options.add_argument("--disable-images")

    profile_dir = None
    if use_auth:
        from pathlib import Path
        timestamp = int(time.time() * 1000)
        worker_id = uuid.uuid4().hex[:8]
        app_dir = Path.home() / ".yandex_parser_auth"
        app_dir.mkdir(exist_ok=True)
        profile_dir = app_dir / f"edge_profile_{worker_id}_{timestamp}"
        profile_dir.mkdir(parents=True, exist_ok=True)
        options.add_argument(f"--user-data-dir={profile_dir}")
        # ВАЖНО: отключаем автозаполнение и другие функции, которые могут мешать
        options.add_argument("--disable-features=AutofillServerCommunication")
        CREATED_PROFILES.add(str(profile_dir))
        logger.debug(f"Создан профиль для авторизации: {profile_dir}")
    else:
        temp_dir = tempfile.mkdtemp(prefix=f"edge_temp_{uuid.uuid4().hex[:8]}_")
        options.add_argument(f"--user-data-dir={temp_dir}")
        CREATED_PROFILES.add(temp_dir)

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1024,600")

    try:
        if driver_path:
            edge_driver_path = driver_path
        else:
            edge_driver_path = os.path.join("browserdriver", "msedgedriver.exe")

        if not os.path.exists(edge_driver_path):
            raise FileNotFoundError(f"Edge WebDriver не найден: {edge_driver_path}")

        service = Service(edge_driver_path)
        driver = webdriver.Edge(service=service, options=options)

        driver.set_page_load_timeout(15)
        driver.implicitly_wait(3)

        return driver

    except Exception as e:
        if profile_dir and str(profile_dir) in CREATED_PROFILES:
            try:
                shutil.rmtree(profile_dir, ignore_errors=True)
                CREATED_PROFILES.discard(str(profile_dir))
            except:
                pass
        logger.error(f"Ошибка создания Edge драйвера: {e}")
        raise

def load_cookies_for_auth(driver):
    """ПРАВИЛЬНАЯ загрузка cookies - СНАЧАЛА yandex.ru, ПОТОМ market.yandex.ru"""
    if STOP_PARSING:
        return False

    cookies_file = os.path.expanduser("~/.yandex_parser_auth/cookies.json")
    if not os.path.exists(cookies_file):
        logger.warning(f"Файл cookies не найден: {cookies_file}")
        return False

    try:
        with open(cookies_file, 'r', encoding='utf-8') as f:
            cookies_data = json.loads(f.read().strip())

        if isinstance(cookies_data, list):
            cookies = cookies_data
        elif isinstance(cookies_data, dict) and 'cookies' in cookies_data:
            cookies = cookies_data['cookies']
        else:
            logger.error("Неверный формат cookies файла")
            return False

        logger.info(f"Найдено {len(cookies)} cookies в файле")

        # КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: Сначала переходим на yandex.ru для загрузки базовых cookies
        logger.debug("Переход на yandex.ru для загрузки cookies...")
        driver.get("https://yandex.ru")
        time.sleep(2)  # Даём время загрузиться
        
        loaded_count = 0
        error_count = 0
        important_cookies_loaded = []
        
        # Загружаем ВСЕ cookies на yandex.ru
        for i, cookie in enumerate(cookies):
            if STOP_PARSING:
                break
            try:
                if not isinstance(cookie, dict) or 'name' not in cookie or 'value' not in cookie:
                    logger.debug(f"Cookie {i}: пропущен (нет name или value)")
                    continue

                clean_cookie = {
                    'name': str(cookie['name']),
                    'value': str(cookie['value']),
                    'path': str(cookie.get('path', '/'))
                }

                # ВАЖНО: корректная обработка domain
                if 'domain' in cookie:
                    domain = str(cookie['domain'])
                    # Убираем лидирующую точку если есть
                    if domain.startswith('.'):
                        domain = domain[1:]
                    clean_cookie['domain'] = domain
                
                if cookie.get('secure', False):
                    clean_cookie['secure'] = True
                
                # Добавляем sameSite если есть
                if 'sameSite' in cookie:
                    clean_cookie['sameSite'] = str(cookie['sameSite'])

                driver.add_cookie(clean_cookie)
                loaded_count += 1
                
                # Логируем важные cookies
                if cookie['name'] in ['Session_id', 'sessionid2', 'yandexuid', 'i', 'yandex_login']:
                    important_cookies_loaded.append(cookie['name'])
                    logger.debug(f"✓ Важный cookie загружен: {cookie['name']}")
                    
            except Exception as e:
                error_count += 1
                logger.debug(f"Cookie {i} ({cookie.get('name', '?')}): ошибка - {e}")
                continue

        logger.info(f"Загружено cookies: {loaded_count} успешно, {error_count} ошибок")
        if important_cookies_loaded:
            logger.info(f"Важные cookies: {', '.join(important_cookies_loaded)}")

        if loaded_count > 0:
            # Теперь переходим на market.yandex.ru С УСТАНОВЛЕННЫМИ cookies
            logger.debug("Переход на market.yandex.ru с установленными cookies...")
            driver.get("https://market.yandex.ru")
            time.sleep(2)  # Даём время загрузиться с авторизацией
            
            # Проверяем что авторизация сохранилась
            page_source = driver.page_source.lower()
            if 'для юрлиц' in page_source or 'для бизнеса' in page_source:
                logger.info("✓ АВТОРИЗАЦИЯ УСПЕШНА - обнаружены признаки бизнес-аккаунта!")
                return True
            else:
                logger.warning("⚠ Авторизация возможно не удалась - не найдены признаки бизнес-аккаунта")
                logger.warning("   Продолжаю работу, но цены для юрлиц могут не появиться")
                return True  # Всё равно продолжаем
        else:
            logger.error("Ни один cookie не загружен!")
            return False

    except Exception as e:
        logger.error(f"Ошибка загрузки cookies: {e}")
        import traceback
        traceback.print_exc()
        return False

def extract_prices_fast(driver):
    """Быстрое извлечение цен: массово считывает первые 4 ds.valueLine + подписи"""
    price_data = {
        'обычная цена': '',
        'цена для юрлиц': ''
    }

    if STOP_PARSING:
        return price_data

    try:
        logger.debug("Извлечение цен из карточки товара...")


        script = """
        var result = {
            prices: [],
            labels: []
        };

        var valuelines = document.querySelectorAll("span.ds-valueLine");
        var targetElements = Array.from(valuelines).slice(0, 4);

        for (var i = 0; i < targetElements.length; i++) {
            var element = targetElements[i];
            var priceText = element.textContent.trim();
            result.prices.push(priceText);

            // Поиск подписей в соседних элементах
            var labelText = "";
            var parent = element.parentElement;

            if (parent && parent.parentElement) {
                var textLines = parent.parentElement.querySelectorAll(".ds-textLine");
                for (var j = 0; j < Math.min(textLines.length, 3); j++) {
                    var text = textLines[j].textContent.trim().toLowerCase();
                    if (text && text.length < 25) {
                        labelText = text;
                        break;
                    }
                }
            }

            result.labels.push(labelText);
        }

        return result;
        """

        try:
            bulk_data = driver.execute_script(script)
        except Exception as e:
            logger.warning(f"JavaScript ошибка, используем fallback: {e}")
            # Fallback
            all_valuelines = driver.find_elements(By.CSS_SELECTOR, "span.ds-valueLine")
            target_valuelines = all_valuelines[:4] if all_valuelines else []

            if not target_valuelines:
                return price_data

            bulk_data = {'prices': [], 'labels': []}
            for valueline in target_valuelines:
                bulk_data['prices'].append(valueline.text.strip())
                bulk_data['labels'].append("")

        if not bulk_data or not bulk_data.get('prices'):
            return price_data

        prices = bulk_data['prices']
        labels = bulk_data['labels']

        # Формируем данные для классификации
        prices_with_labels = []
        for i, (price_text, label_text) in enumerate(zip(prices, labels)):
            prices_with_labels.append({
                'text': price_text,
                'label': label_text.lower(),
                'index': i + 1
            })

        # Классификация по подписям
        regular_found = False
        vat_found = False

        # 1. Ищем "пэй" для обычной цены
        for item in prices_with_labels:
            if 'пэй' in item['label'] or 'pay' in item['label']:
                price_data['обычная цена'] = item['text']
                regular_found = True
                break

        # 2. Ищем "с НДС" для юрлиц
        for item in prices_with_labels:
            if 'с ндс' in item['label'] or 'ндс' in item['label'] or 'для юрлиц' in item['label']:
                price_data['цена для юрлиц'] = item['text']
                vat_found = True
                break

        # 3. Если не нашли "пэй" → первая цена как обычная
        if not regular_found and prices_with_labels:
            price_data['обычная цена'] = prices_with_labels[0]['text']

        return price_data

    except Exception as e:
        logger.error(f"Ошибка извлечения цен: {e}")
        return price_data

def extract_products_smart(driver) -> List[Dict[str, Any]]:
    products = []

    try:
        script = """
        var products = [];

        // ИСПРАВЛЕННЫЙ селектор из примера пользователя
        var cards = document.querySelectorAll('span[role="link"][data-auto="snippet-title"]');

        for (var i = 0; i < Math.min(10, cards.length); i++) {
            var card = cards[i];
            var title = card.textContent.trim();
            if (!title) continue;

            // Поиск ссылки в родительских элементах
            var url = null;
            var parent = card;
            for (var j = 0; j < 5; j++) {
                if (parent.tagName === 'A' && parent.href) {
                    url = parent.href;
                    break;
                }
                parent = parent.parentElement;
                if (!parent) break;
            }

            // Если не нашли в родителях, ищем в соседних элементах
            if (!url) {
                var links = card.parentElement ? card.parentElement.querySelectorAll('a[href]') : [];
                if (links.length > 0) {
                    url = links[0].href;
                }
            }

            products.push({
                title: title,
                url: url,
                index: i
            });
        }

        return products;
        """

        products_data = driver.execute_script(script)

        if products_data:
            products = [
                {
                    'title': p['title'],
                    'url': p['url'],
                    'index': p['index']
                }
                for p in products_data[:5]  # Максимум 5 карточек
            ]

        if products:
            logger.debug(f"Найдено {len(products)} товаров")
            return products

    except Exception as e:
        logger.warning(f"Ошибка извлечения товаров: {e}")

    return products

def parse_price_to_number(price_str: str) -> float:
    """Конвертирует строку цены в число для сравнения"""
    if not price_str:
        return float('inf')  # Бесконечность для отсутствующих цен

    try:
        # Убираем все кроме цифр, запятых и точек
        clean_price = re.sub(r'[^\d,.]', '', price_str)

        # Заменяем запятые на точки для float
        clean_price = clean_price.replace(',', '.')

        # Убираем множественные точки (оставляем только последнюю)
        if clean_price.count('.') > 1:
            parts = clean_price.split('.')
            clean_price = ''.join(parts[:-1]) + '.' + parts[-1]

        return float(clean_price) if clean_price else float('inf')
    except:
        return float('inf')

def collect_prices_from_all_products(driver, products: List[Dict[str, Any]], search_term: str) -> Dict[str, str]:
    """Собирает цены со ВСЕХ 5 карточек и выбирает НАИМЕНЬШУЮ"""
    result = {"цена": "", "цена для юрлиц": "", "ссылка": ""}

    if not products:
        logger.warning("Нет товаров для обработки")
        return result

    # Контейнеры для всех найденных цен
    all_products_data = []

    logger.info(f"Собираю цены с {len(products)} карточек товаров:")

    # Проходим по ВСЕМ товарам и собираем цены
    for i, product in enumerate(products, 1):
        if STOP_PARSING:
            break

        if not product.get('url'):
            logger.debug(f"Товар {i}: нет ссылки, пропуск")
            continue

        try:
            short_title = product['title'][:45] + "..." if len(product['title']) > 45 else product['title']
            logger.info(f"  {i}. {short_title}")

            # БЕЗОПАСНЫЙ переход с защитой от stale elements
            for retry in range(2):
                try:
                    driver.get(product['url'])
                    time.sleep(1.2)
                    break
                except (WebDriverException, TimeoutException):
                    if retry == 1:
                        logger.warning(f"     Ошибка загрузки после повтора")
                        break
                    time.sleep(1)
                    continue

            if STOP_PARSING:
                break

            # Проверяем загрузку страницы
            try:
                WebDriverWait(driver, 5).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
            except:
                pass

            # Извлекаем цены
            prices = extract_prices_fast(driver)

            # Сохраняем данные товара с ценами
            product_data = {
                'title': product['title'],
                'url': product['url'],
                'index': i,
                'обычная цена': prices.get('обычная цена', ''),
                'цена для юрлиц': prices.get('цена для юрлиц', ''),
                'regular_price_num': parse_price_to_number(prices.get('обычная цена', '')),
                'vat_price_num': parse_price_to_number(prices.get('цена для юрлиц', ''))
            }

            all_products_data.append(product_data)

            # Логируем что нашли
            price_info = []
            if prices.get('обычная цена'):
                price_info.append(f"Обычная: {prices['обычная цена']}")
            if prices.get('цена для юрлиц'):
                price_info.append(f"Юрлица: {prices['цена для юрлиц']}")

            if price_info:
                logger.info(f"     {', '.join(price_info)}")
            else:
                logger.info(f"     цены не найдены")

        except StaleElementReferenceException as e:
            logger.warning(f"     StaleElement ошибка")
            continue
        except Exception as e:
            logger.warning(f"     Ошибка: {e}")
            continue

    if not all_products_data:
        logger.warning("Ни один товар не дал результата")
        return result

    # ВЫБИРАЕМ товар с НАИМЕНЬШЕЙ обычной ценой
    valid_products = [p for p in all_products_data if p['regular_price_num'] != float('inf')]

    if valid_products:
        # Сортируем по обычной цене (по возрастанию)
        best_product = min(valid_products, key=lambda x: x['regular_price_num'])

        result["цена"] = best_product['обычная цена']
        result["цена для юрлиц"] = best_product['цена для юрлиц']
        result["ссылка"] = best_product['url']

        logger.info(f"ЛУЧШИЙ ВЫБОР: товар {best_product['index']} - {best_product['обычная цена']}")

        # Показываем сравнение цен
        logger.info("Сравнение цен:")
        for p in sorted(all_products_data, key=lambda x: x['regular_price_num']):
            if p['regular_price_num'] != float('inf'):
                marker = "→ ВЫБРАН" if p == best_product else ""
                logger.info(f"  Товар {p['index']}: {p['обычная цена']} {marker}")
    else:
        # Если нет обычных цен, берем первый товар с любыми данными
        first_product = all_products_data[0]
        result["цена"] = first_product['обычная цена']
        result["цена для юрлиц"] = first_product['цена для юрлиц']
        result["ссылка"] = first_product['url']

        logger.warning("Обычные цены не найдены, взят первый товар")

    return result

def smart_search_input(driver, search_term: str, max_retries: int = 3) -> bool:
    """УЛУЧШЕННАЯ функция поиска с определением текущего состояния страницы"""
    current_url = driver.current_url

    # Проверяем, находимся ли мы уже на странице поиска
    if 'search' in current_url and 'text=' in current_url:
        logger.debug("Уже на странице поиска, обновляем запрос")
        # Пытаемся найти поле поиска на странице результатов
        return update_search_query(driver, search_term, max_retries)
    else:
        logger.debug("На главной странице, выполняем новый поиск")
        # Выполняем поиск с главной страницы
        return perform_new_search(driver, search_term, max_retries)

def update_search_query(driver, search_term: str, max_retries: int = 3) -> bool:
    """Обновляет поисковый запрос на странице результатов"""

    for retry in range(max_retries):
        if STOP_PARSING:
            return False

        try:
            # Ждем загрузки страницы
            WebDriverWait(driver, 3).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )

            # Расширенный список селекторов для поля поиска на странице результатов
            search_selectors = [
                'input[name="text"]',
                'input[data-auto="search-input"]',
                'input[placeholder*="искать"]',
                'input[placeholder*="поиск"]',
                '.search-input input',
                '.header-search input',
                '[data-zone="search"] input',
                'input.n-search__input',
                'input[type="search"]'
            ]

            searchbox = None
            for selector in search_selectors:
                try:
                    searchbox = WebDriverWait(driver, 2).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    logger.debug(f"Найдено поле поиска: {selector}")
                    break
                except TimeoutException:
                    continue

            if not searchbox:
                logger.warning(f"Попытка {retry + 1}: поле поиска не найдено на странице результатов")
                if retry < max_retries - 1:
                    # Пытаемся перейти на главную страницу
                    driver.get("https://market.yandex.ru")
                    time.sleep(1)
                    continue
                return False

            # Обновляем поисковый запрос
            try:
                # Очищаем текущий запрос
                searchbox.clear()
                time.sleep(0.3)

                # Вводим новый запрос
                searchbox.send_keys(search_term[:50])
                time.sleep(0.3)
                searchbox.send_keys(Keys.RETURN)
                time.sleep(1.5)
                return True

            except StaleElementReferenceException:
                logger.warning(f"Попытка {retry + 1}: StaleElement при обновлении запроса")
                if retry < max_retries - 1:
                    time.sleep(1)
                    continue
                return False

        except Exception as e:
            logger.warning(f"Попытка {retry + 1} обновления запроса: {e}")
            if retry < max_retries - 1:
                time.sleep(1)
                continue
            return False

    return False

def perform_new_search(driver, search_term: str, max_retries: int = 3) -> bool:
    """Выполняет новый поиск с главной страницы"""

    for retry in range(max_retries):
        if STOP_PARSING:
            return False

        try:
            # Ждем загрузки страницы
            WebDriverWait(driver, 5).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )

            # Находим поле поиска
            searchbox = None
            wait = WebDriverWait(driver, 5)

            for selector_type, selector in [
                (By.NAME, "text"),
                (By.CSS_SELECTOR, "input[name='text']"),
                (By.CSS_SELECTOR, "[data-auto='search-input']")
            ]:
                try:
                    searchbox = wait.until(EC.element_to_be_clickable((selector_type, selector)))
                    break
                except TimeoutException:
                    continue

            if not searchbox:
                logger.warning(f"Попытка {retry + 1}: поле поиска не найдено на главной")
                if retry < max_retries - 1:
                    time.sleep(1)
                    continue
                return False

            # Выполняем поиск
            try:
                searchbox.clear()
                searchbox.send_keys(search_term[:50])
                searchbox.send_keys(Keys.RETURN)
                time.sleep(1.5)
                return True

            except StaleElementReferenceException:
                logger.warning(f"Попытка {retry + 1}: StaleElement при новом поиске")
                if retry < max_retries - 1:
                    time.sleep(1)
                    continue
                return False

        except Exception as e:
            logger.warning(f"Попытка {retry + 1} нового поиска: {e}")
            if retry < max_retries - 1:
                time.sleep(1)
                continue
            return False

    return False

def get_prices(product_name: str, headless: bool = True, driver_path: Optional[str] = None,
              timeout: int = 15, use_business_auth: bool = False) -> Dict[str, str]:
    """Главная функция получения цен с выбором наименьшей из 5 карточек"""
    result = {"цена": "", "цена для юрлиц": "", "ссылка": ""}
    driver = None
    current_profile_path = None

    if STOP_PARSING:
        return result

    try:
        driver = create_driver(headless=headless, driver_path=driver_path, use_auth=use_business_auth)

        # Отслеживаем профиль для очистки
        if use_business_auth:
            for profile_path in CREATED_PROFILES:
                if 'edge_profile_' in profile_path:
                    current_profile_path = profile_path
                    break
        else:
            for profile_path in CREATED_PROFILES:
                if 'edge_temp_' in profile_path:
                    current_profile_path = profile_path
                    break

        # Загрузка cookies для авторизации
        if use_business_auth and not STOP_PARSING:
            auth_success = load_cookies_for_auth(driver)
            if auth_success:
                logger.info("✓ Авторизация успешна")
            else:
                logger.warning("⚠ Авторизация не удалась, продолжаю без неё")
            # load_cookies_for_auth уже перешёл на market.yandex.ru
        else:
            # Если НЕ используем авторизацию - переходим на маркет вручную
            if STOP_PARSING:
                return result
            try:
                driver.get("https://market.yandex.ru")
                time.sleep(1.0)
            except Exception as e:
                logger.error(f"Ошибка перехода на маркет: {e}")
                return result

        if STOP_PARSING:
            return result

        # УЛУЧШЕННЫЙ поиск с определением состояния страницы
        search_success = smart_search_input(driver, product_name)
        if not search_success:
            logger.error("Не удалось выполнить поиск")
            return result

        if STOP_PARSING:
            return result

        # Извлечение товаров
        products = extract_products_smart(driver)
        if not products:
            logger.warning("Товары не найдены")
            return result

        if STOP_PARSING:
            return result

        # Собираем цены со ВСЕХ товаров и выбираем НАИМЕНЬШУЮ
        result = collect_prices_from_all_products(driver, products, product_name)

        return result

    except Exception as e:
        logger.error(f"Ошибка обработки товара {product_name[:30]}...: {e}")
        return result

    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass

        # Очистка профиля
        if current_profile_path:
            success = cleanup_single_profile(current_profile_path)
            if success:
                CREATED_PROFILES.discard(current_profile_path)

def parse_tender_excel(input_file: str, output_file: str, headless: bool = True,
                      workers: int = 1, driver_path: Optional[str] = None,
                      auto_save: bool = True, use_business_auth: bool = False) -> pd.DataFrame:
    """ОСНОВНАЯ функция парсинга с автосохранением и ТЕНДЕРНЫМ ФОРМАТОМ"""
    global STOP_PARSING, CURRENT_DATAFRAME, CURRENT_OUTPUT_FILE, CURRENT_INPUT_FILE

    # Настройка автосохранения при завершении
    setup_signal_handlers()

    STOP_PARSING = False
    CURRENT_INPUT_FILE = input_file
    CURRENT_OUTPUT_FILE = output_file

    kill_zombie_edges()

    items = extract_products_from_excel(input_file)
    if items.empty:
        raise ValueError("Не найдены товары в файле")

    # DataFrame для хранения результатов парсинга
    df = pd.DataFrame({
        'наименование': items['name'],
        'цена': '',
        'цена для юрлиц': '',
        'ссылка': ''
    })

    CURRENT_DATAFRAME = df  # Для автосохранения

    auth_text = "с авторизацией" if use_business_auth else "без авторизации"
    logger.info(f"Начинаю обработку {len(df)} товаров {auth_text}")
    logger.info("🔄 Автосохранение при принудительном завершении АКТИВНО")
    logger.info("📋 РЕЗУЛЬТАТ: тендерная таблица с колонкой 'Яндекс Маркет'")
    logger.info("Режим: поиск наименьшей цены среди 5 карточек")

    try:
        for idx, row in df.iterrows():
            if STOP_PARSING:
                logger.info("Парсинг остановлен")
                break

            try:
                logger.info(f"Обработка: {idx + 1}/{len(df)} - {row['наименование'][:40]}...")

                prices = get_prices(row['наименование'], headless, driver_path, 20, use_business_auth)

                df.at[idx, 'цена'] = prices.get('цена', '')
                df.at[idx, 'цена для юрлиц'] = prices.get('цена для юрлиц', '')
                df.at[idx, 'ссылка'] = prices.get('ссылка', '')

                # Лог результата
                price_summary = []
                if prices.get('цена'):
                    price_summary.append(f"Лучшая цена: {prices['цена'][:15]}")
                if prices.get('цена для юрлиц'):
                    price_summary.append(f"Для юрлиц: {prices['цена для юрлиц'][:15]}")

                if price_summary:
                    logger.info(f"Результат {idx + 1}/{len(df)}: {', '.join(price_summary)}")
                else:
                    logger.info(f"Результат {idx + 1}/{len(df)}: цены не найдены")

                # Автосохранение каждые 3 товара В ТЕНДЕРНОМ ФОРМАТЕ
                if auto_save and (idx + 1) % 3 == 0:
                    try:
                        save_results_into_tender_format(input_file, output_file, df)
                        logger.info(f"Автосохранение тендера: {idx + 1}/{len(df)}")
                    except Exception as e:
                        logger.warning(f"Ошибка автосохранения: {e}")

            except Exception as e:
                logger.error(f"Ошибка товара {idx + 1}: {e}")
                df.at[idx, 'цена'] = "ОШИБКА"
                df.at[idx, 'цена для юрлиц'] = "ОШИБКА"

    finally:
        cleanup_profiles()
        CURRENT_DATAFRAME = None  # Очищаем глобальную переменную

    # Финальное сохранение В ТЕНДЕРНОМ ФОРМАТЕ
    if output_file != "auto":
        save_results_into_tender_format(input_file, output_file, df)
        logger.info(f"🎯 ТЕНДЕРНАЯ ТАБЛИЦА ГОТОВА: {output_file}")
        logger.info("📊 Создана точная копия оригинала + колонка 'Яндекс Маркет'")

    return df

if __name__ == "__main__":
    test_product = "Точка доступа Ubiquiti UniFi AC Pro AP"
    print("Тест финальной версии с тендерным форматом...")
    result = get_prices(test_product, headless=False, use_business_auth=True)

    print(f"Товар: {test_product}")
    print(f"Лучшая цена: {result['цена']}")
    print(f"Цена для юрлиц: {result['цена для юрлиц'] or 'НЕ НАЙДЕНА'}")
    print(f"Ссылка: {result['ссылка']}")
    print("-" * 50)
