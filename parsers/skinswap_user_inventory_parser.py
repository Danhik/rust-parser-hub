import os
import sys
import time
import json
import argparse
import random
import threading
import queue
import concurrent.futures
from typing import List, Dict, Any, Optional

import requests
from openpyxl import Workbook
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# -------------------------------------------------------------
# Configuration & Utils
# -------------------------------------------------------------

def format_price(cents: int) -> str:
    if cents is None:
        return ""
    dollars = cents / 100.0
    return f"{dollars:.2f}".replace(".", ",")

def get_random_user_agent() -> str:
    return "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0"

# -------------------------------------------------------------
# Session Management (Selenium)
# -------------------------------------------------------------

def _get_chrome_major_version() -> Optional[int]:
    """Определяет major-версию установленного Chrome."""
    import subprocess
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in paths:
        if os.path.exists(path):
            try:
                out = subprocess.check_output(
                    ['powershell', '-Command', f'(Get-Item "{path}").VersionInfo.ProductVersion'],
                    stderr=subprocess.DEVNULL, timeout=5
                ).decode().strip()
                major = int(out.split(".")[0])
                print(f"Обнаружена версия Chrome: {out} (major={major})")
                return major
            except Exception:
                pass
    return None


def get_session_cookies_via_selenium(session_file: str) -> dict:
    print("Запуск браузера для авторизации. Пожалуйста, войдите через Steam...")
    
    options = uc.ChromeOptions()
    # options.add_argument('--headless') # Нельзя headless, нужна ручная авторизация
    
    chrome_major = _get_chrome_major_version()
    try:
        if chrome_major:
            driver = uc.Chrome(options=options, version_main=chrome_major)
        else:
            driver = uc.Chrome(options=options)
    except Exception as e:
        print(f"Ошибка запуска UC с version_main={chrome_major}: {e}. Пробуем без указания версии...")
        driver = uc.Chrome(options=options)
    
    try:
        driver.get("https://skinswap.com/")
        
        print("Ожидаем авторизации (кука 'token')...")
        print("Вам дается 5 минут на вход через Steam. После успешного входа браузер закроется сам.")
        
        token_cookie = None
        cf_clearance_cookie = None
        user_agent = driver.execute_script("return navigator.userAgent;")
        
        # Ждем появления куки 'token' (обычно ставится после редиректа со Steam)
        timeout = time.time() + 300 # 5 минут
        while time.time() < timeout:
            cookies = driver.get_cookies()
            for c in cookies:
                if c['name'] == 'token':
                    token_cookie = c['value']
                elif c['name'] == 'cf_clearance':
                    cf_clearance_cookie = c['value']
            
            if token_cookie:
                break
            
            time.sleep(2)
            
        if not token_cookie:
            print("[Ошибка] Время ожидания авторизации истекло. Закрываем браузер.")
            return {}
            
        print("Успешная авторизация! Токен получен.")
        
        session_data = {
            "token": token_cookie,
            "cf_clearance": cf_clearance_cookie,
            "user_agent": user_agent,
            "timestamp": time.time()
        }
        
        with open(session_file, "w", encoding="utf-8") as f:
            json.dump(session_data, f, indent=4)
            
        print(f"Сессия сохранена в {session_file}")
        return session_data
        
    finally:
        driver.quit()

def load_session(session_file: str) -> dict:
    if os.path.exists(session_file):
        try:
            with open(session_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                
                # Базовая проверка "протухания" (допустим 24 часа)
                if time.time() - data.get("timestamp", 0) > 86400:
                    print("Сохраненная сессия устарела (прошло > 24 часов). Потребуется новая авторизация.")
                    return {}
                
                return data
        except Exception as e:
            print(f"Ошибка чтения сессии: {e}")
            return {}
    return {}

def ensure_session(session_file: str) -> dict:
    session_data = load_session(session_file)
    if not session_data or not session_data.get("token"):
        print("Нет сохраненной сессии или токена.")
        session_data = get_session_cookies_via_selenium(session_file)
    else:
        print("Найдена сохраненная сессия. Проверяем токен...")
        
    return session_data

# -------------------------------------------------------------
# API Worker Logic
# -------------------------------------------------------------

def create_api_session(session_data: dict) -> requests.Session:
    session = requests.Session()
    
    ua = session_data.get("user_agent", get_random_user_agent())
    token = session_data.get("token", "")
    cf_clearance = session_data.get("cf_clearance", "")
    
    session.headers.update({
        "User-Agent": ua,
        "Origin": "https://skinswap.com",
        "Referer": "https://skinswap.com/",
        "Accept": "*/*",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "x-requested-with": "SkinSwap v3.1.5"
    })
    
    # Добавляем куки
    if token:
        session.cookies.set("token", token, domain=".skinswap.com")
    if cf_clearance:
        session.cookies.set("cf_clearance", cf_clearance, domain=".skinswap.com")
        
    return session

def worker_task(
    offset: int, 
    limit: int, 
    appid: int, 
    session_data: dict,
    max_retries: int,
    timeout: int,
    verbose: bool,
    only_accepted: bool
) -> Optional[List[Dict[str, Any]]]:
    
    url = f"https://api.skinswap.com/api/user/inventory?offset={offset}&limit={limit}&appid={appid}&sort=price-desc"
    
    session = create_api_session(session_data)
    
    for attempt in range(1, max_retries + 1):
        if verbose:
            print(f"[API] Offset: {offset} | Попытка {attempt}/{max_retries}")
            
        try:
            # OPTIONS preflight для Cloudflare
            session.options(url, timeout=timeout)
            
            # Основной запрос GET
            resp = session.get(url, timeout=timeout)
            
            if resp.status_code == 200:
                data = resp.json()
                if data.get("success"):
                    # Возвращаем кортеж (данные, reached_end)
                    # Если endOfResults == true, значит больше дергать API не надо
                    items = data.get("data", [])
                    reached_end = data.get("endOfResults", False)
                    
                    if only_accepted:
                        items = [item for item in items if item.get("accepted") is True]
                        
                    return (items, reached_end)
                else:
                    if verbose: print(f"      [Ошибка] success=false на offset {offset}. Ответ: {data}")
                    return ([], True)
            elif resp.status_code == 401 or resp.status_code == 403:
                print(f"      [Auth Error] Код {resp.status_code}. Возможно, токен протух или Cloudflare заблокировал запрос.")
                return None # Фатальная ошибка, нужна переавторизация
            elif resp.status_code in [429, 503]:
                if verbose: print(f"      [RateLimit/503] Код {resp.status_code}. Ждем 5 сек...")
                time.sleep(5)
                continue
            else:
                if verbose: print(f"      [Ошибка] Код {resp.status_code} на offset {offset}")
                
        except requests.exceptions.RequestException as e:
            if verbose: print(f"      [Network Error] {type(e).__name__}")
            time.sleep(2)
            continue
            
        time.sleep(2)
        
    print(f"Не удалось получить данные для offset={offset} после {max_retries} попыток.")
    return None

# -------------------------------------------------------------
# Main Orchestrator
# -------------------------------------------------------------

def _save_workbook_safe(wb: Workbook, path: str, retries: int = 5):
    for i in range(retries):
        try:
            wb.save(path)
            return
        except PermissionError:
            print(f"[Writer] Ошибка доступа к файлу {path} (возможно открыт?). Попытка {i+1}/{retries} через 5 сек...")
            time.sleep(5)
        except Exception as e:
            print(f"[Writer] Ошибка сохранения {path}: {e}")
            return
    print(f"[Writer] Не удалось сохранить {path} после {retries} попыток. Данные могут быть потеряны.")


def save_to_excel(items: List[Dict[str, Any]], out_path: str, only_accepted: bool = False):
    wb = Workbook()
    ws = wb.active
    ws.title = "My Inventory"
    
    if only_accepted:
        headers = ["Name", "Price (Trade)", "Overstock Limit", "Overstock Count"]
    else:
        headers = [
            "Name", 
            "Tradable", 
            "Accepted on Skinswap", 
            "Reason (if rejected)", 
            "Trade Price (They pay)", 
            "Sell Price (Instant cash)", 
            "Overstock Limit", 
            "Overstock Count"
        ]
    ws.append(headers)
    
    seen_ids = set()
    saved = 0
    
    for item in items:
        item_id = item.get("id")
        if item_id and item_id in seen_ids:
            continue
        if item_id:
            seen_ids.add(item_id)
            
        name = item.get("name", "Unknown")
        prices = item.get("price", {})
        trade_price = format_price(prices.get("trade"))
        
        overstock = item.get("overstock", {})
        limit = overstock.get("limit", "")
        count = overstock.get("count", "")
        
        if only_accepted:
            ws.append([name, trade_price, limit, count])
        else:
            tradable = "Yes" if item.get("tradable") else "No"
            accepted = "Yes" if item.get("accepted") else "No"
            reason = item.get("reason", "")
            sell_price = format_price(prices.get("sell"))
            ws.append([name, tradable, accepted, reason, trade_price, sell_price, limit, count])
            
        saved += 1
        
    _save_workbook_safe(wb, out_path)
    print(f"Сохранено {saved} (уникальных) предметов в файл {out_path}.")


def check_auth_validity(session_data: dict, appid: int) -> bool:
    """Делает тестовый запрос 1 предмета, чтобы убедиться, что токен жив"""
    print("Проверка валидности сессии...")
    session = create_api_session(session_data)
    url = f"https://api.skinswap.com/api/user/inventory?offset=0&limit=1&appid={appid}&sort=price-desc"
    try:
        session.options(url, timeout=10)
        resp = session.get(url, timeout=10)
        if resp.status_code == 200 and resp.json().get("success") is True:
            print("Сессия активна!")
            return True
        else:
            print(f"Сессия невалидна (Code {resp.status_code}). Требуется переавторизация.")
            return False
    except Exception as e:
        print(f"Ошибка проверки сессии: {e}")
        return False


def main():
    ap = argparse.ArgumentParser(description="Парсер ЛИЧНОГО инвентаря Skinswap (Гибридный подход: Selenium -> Requests).")
    ap.add_argument("--appid", type=int, default=252490, help="ID игры (по умолчанию 252490 - Rust)")
    ap.add_argument("--workers", type=int, default=3, help="Количество потоков (не ставьте много для личного, чтобы избежать 429)")
    ap.add_argument("--task-retries", type=int, default=4, help="Количество попыток загрузить одну страницу")
    ap.add_argument("--limit", type=int, default=80, help="Количество предметов на страницу (по умолчанию 80 для пользователя)")
    ap.add_argument("--out", default="skinswap_user_items.xlsx", help="Имя выходного файла")
    ap.add_argument("--timeout", type=int, default=15, help="Таймаут запросов")
    ap.add_argument("--verbose", action="store_true", help="Включить подробные логи API")
    ap.add_argument("--accepted-only", action="store_true", help="ФИЛЬТР: сохранить только вещи, которые Skinswap принимает (accepted=true)")
    ap.add_argument("--user-profile", default="default", help="Имя профиля для сохранения сессии (например: user1).")

    args = ap.parse_args()
    
    session_file = f"skinswap_session_{args.user_profile}.json"

    # 1. Получение / проверка сессии
    session_data = load_session(session_file)
    
    if session_data and not check_auth_validity(session_data, args.appid):
        print("Запускаем браузер для обновления куки...")
        session_data = get_session_cookies_via_selenium(session_file)
        
    elif not session_data:
        session_data = get_session_cookies_via_selenium(session_file)
        
    if not session_data or not session_data.get("token"):
        print("Фатальная ошибка: Не удалось получить токен авторизации. Выход.")
        sys.exit(1)
        
    only_accepted = args.accepted_only
    
    msg_filter = "ТОЛЬКО ПРИНИМАЕМЫЕ" if only_accepted else "ВСЕ (в т.ч. непринимаемые с причиной)"
    print(f"Начинаем сбор личного инвентаря [AppID: {args.appid} | Потоков: {args.workers} | Режим: {msg_filter}]...")
    
    results_by_offset = {}
    
    # 2. Многопоточный парсинг
    offset_queue = queue.Queue()
    for i in range(args.workers):
        offset_queue.put(i * args.limit)
    
    highest_offset_queued = (args.workers - 1) * args.limit
    
    active_tasks = 0
    pages_processed = 0
    end_reached = False
    auth_failed = False
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {}
        
        while not offset_queue.empty():
            offset = offset_queue.get()
            fut = executor.submit(
                worker_task, 
                offset, args.limit, args.appid, session_data, args.task_retries, args.timeout, args.verbose, only_accepted
            )
            futures[fut] = offset
            active_tasks += 1
            
        while active_tasks > 0:
            done, _ = concurrent.futures.wait(futures.keys(), return_when=concurrent.futures.FIRST_COMPLETED)
            
            for fut in done:
                offset = futures.pop(fut)
                active_tasks -= 1
                pages_processed += 1
                
                try:
                    result = fut.result()
                    if result is not None:
                        items, batch_end_reached = result
                    else:
                        items, batch_end_reached = None, False
                except Exception as e:
                    print(f"Критическая ошибка в потоке offset={offset}: {e}")
                    items, batch_end_reached = None, False
                    
                if items is None:
                    auth_failed = True
                    end_reached = True
                else:
                    results_by_offset[offset] = items
                    current_total = sum(len(v) for v in results_by_offset.values())
                    print(f"[Прогресс] Страниц проверено: {pages_processed}. Собрано предметов (по фильтру): {current_total}")
                    
                    if batch_end_reached or len(items) == 0:
                        end_reached = True
                    
                    if not end_reached:
                        highest_offset_queued += args.limit
                        
                        new_fut = executor.submit(
                            worker_task, 
                            highest_offset_queued, args.limit, args.appid, session_data, args.task_retries, args.timeout, args.verbose, only_accepted
                        )
                        futures[new_fut] = highest_offset_queued
                        active_tasks += 1
                        
            time.sleep(0.1)

    if auth_failed:
        print("[Внимание] Возникли критические сетевые ошибки или токен устарел во время сбора.")
        print("В следующий раз файл сессии может быть обновлен автоматически.")
        
        if os.path.exists(session_file):
             os.remove(session_file)
             print("Локальный файл сессии удален для форсирования новой авторизации.")

    all_items = []
    for off in sorted(results_by_offset.keys()):
        all_items.extend(results_by_offset[off])

    print(f"Сбор завершен. Получено {len(all_items)} предметов.")
    if all_items:
        save_to_excel(all_items, args.out, only_accepted)

if __name__ == "__main__":
    main()
