#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для проверки авторизации Edge с cookies
"""

import os
import sys
import json
import time
from tender_parser import create_driver, load_cookies_for_auth
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

def test_cookies():
    """Тестирует cookies для авторизации"""
    
    print("\n" + "="*70)
    print("🔐 ТЕСТ АВТОРИЗАЦИИ EDGE ЧЕРЕЗ COOKIES")
    print("="*70 + "\n")
    
    # Проверка файла cookies
    cookies_file = os.path.expanduser("~/.yandex_parser_auth/cookies.json")
    
    if not os.path.exists(cookies_file):
        print(f"❌ Файл cookies не найден: {cookies_file}")
        print("\n📋 Инструкция:")
        print("1. Откройте Edge и войдите на market.yandex.ru")
        print("2. Экспортируйте cookies (например, расширением 'EditThisCookie')")
        print("3. Сохраните в: {cookies_file}")
        return False
    
    # Проверка содержимого
    try:
        with open(cookies_file, 'r', encoding='utf-8') as f:
            cookies_data = json.loads(f.read().strip())
        
        if isinstance(cookies_data, list):
            cookies = cookies_data
        elif isinstance(cookies_data, dict) and 'cookies' in cookies_data:
            cookies = cookies_data['cookies']
        else:
            print("❌ Неверный формат cookies файла")
            return False
        
        print(f"✅ Файл cookies найден: {cookies_file}")
        print(f"📊 Количество cookies: {len(cookies)}")
        
        # Проверяем важные cookies
        important_cookies = ['Session_id', 'sessionid2', 'yandexuid', 'i']
        found_important = []
        for cookie in cookies:
            if cookie.get('name') in important_cookies:
                found_important.append(cookie['name'])
        
        print(f"🔑 Важные cookies найдены: {', '.join(found_important) if found_important else 'НЕТ'}")
        
    except Exception as e:
        print(f"❌ Ошибка чтения cookies: {e}")
        return False
    
    print("\n" + "-"*70)
    print("🚀 ЗАПУСК БРАУЗЕРА EDGE...")
    print("-"*70 + "\n")
    
    driver = None
    try:
        # Создаём драйвер в видимом режиме
        driver = create_driver(headless=False, use_auth=True)
        print("✅ Драйвер Edge создан")
        
        # Загружаем cookies
        print("\n📥 Загрузка cookies...")
        auth_success = load_cookies_for_auth(driver)
        
        if auth_success:
            print("✅ Cookies загружены успешно!")
            print("\n⏳ Ожидаем 5 секунд для проверки авторизации...")
            print("   (проверьте в браузере, авторизованы ли вы)")
            time.sleep(5)
            
            # Проверяем текущий URL
            current_url = driver.current_url
            print(f"\n🔗 Текущая страница: {current_url}")
            
            # Проверяем авторизацию
            try:
                # Пытаемся найти элементы, которые указывают на авторизацию
                page_source = driver.page_source
                
                if 'для юрлиц' in page_source.lower() or 'для бизнеса' in page_source.lower():
                    print("✅ АВТОРИЗАЦИЯ УСПЕШНА! (найден признак бизнес-аккаунта)")
                elif 'войти' in page_source.lower() or 'sign in' in page_source.lower():
                    print("⚠️ АВТОРИЗАЦИЯ НЕ УДАЛАСЬ (найдена кнопка 'Войти')")
                else:
                    print("❓ Статус авторизации неизвестен")
                    
            except Exception as e:
                print(f"⚠️ Не удалось проверить статус авторизации: {e}")
            
            print("\n✅ ТЕСТ ЗАВЕРШЁН")
            print("📌 Если видите, что вы авторизованы - всё работает!")
            print("📌 Если нет - попробуйте экспортировать cookies заново")
            
            return True
        else:
            print("❌ Не удалось загрузить cookies!")
            return False
            
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        if driver:
            print("\n⏳ Закрытие браузера через 10 секунд...")
            time.sleep(10)
            try:
                driver.quit()
                print("✅ Браузер закрыт")
            except:
                pass

if __name__ == "__main__":
    try:
        success = test_cookies()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n❌ Прервано пользователем")
        sys.exit(1)
