import os
import sys
import time
import argparse
import threading
import concurrent.futures
from typing import List, Dict, Any, Optional

import requests
from openpyxl import Workbook

# -------------------------------------------------------------
# Configuration & Utils
# -------------------------------------------------------------

def load_proxies(filepath: str) -> List[str]:
    """Загрузка прокси из файла вида ip:port:user:pass"""
    proxies = []
    if not os.path.exists(filepath):
        print(f"Файл {filepath} не найден. Работа без прокси.")
        return proxies
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                proxies.append(line)
    return proxies

def format_proxy(proxy_str: str, scheme: str = "http") -> dict:
    if not proxy_str:
        return {}
    parts = proxy_str.split(":")
    if len(parts) == 4:
        ip, port, user, pwd = parts
        url = f"{scheme}://{user}:{pwd}@{ip}:{port}"
    elif len(parts) == 2:
        ip, port = parts
        url = f"{scheme}://{ip}:{port}"
    else:
        url = f"{scheme}://{proxy_str}"
    return {"http": url, "https": url}

def get_user_agent() -> str:
    return (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36 Edg/147.0.0.0"
    )

def format_price(cents) -> str:
    """Цена приходит в центах (целое), переводим в доллары с запятой"""
    if cents is None:
        return ""
    return f"{cents / 100:.2f}".replace(".", ",")

# -------------------------------------------------------------
# Proxy Pool
# -------------------------------------------------------------

class ProxyPool:
    def __init__(self, proxies: List[str], max_passes: int = 1):
        self._proxies = proxies.copy()
        self._max_passes = max_passes
        self._current_pass = 1
        self._index = 0
        self._lock = threading.Lock()

    def acquire_next(self) -> Optional[str]:
        if not self._proxies:
            return None
        with self._lock:
            if self._current_pass > self._max_passes:
                return None
            p = self._proxies[self._index]
            self._index += 1
            if self._index >= len(self._proxies):
                self._index = 0
                self._current_pass += 1
            return p

    def rotate(self) -> Optional[str]:
        """Вернуть следующую прокси при ошибке"""
        return self.acquire_next()

    def has_proxies(self) -> bool:
        return bool(self._proxies)

# -------------------------------------------------------------
# Worker
# -------------------------------------------------------------

def worker_task(
    offset: int,
    limit: int,
    game_id: int,
    sort_type: str,
    proxy_pool: ProxyPool,
    max_retries: int,
    timeout: int,
    verbose: bool,
) -> Optional[tuple]:
    """
    Возвращает (items: list, end_reached: bool) или None при фатальной ошибке.
    """
    url = (
        f"https://tradeit.gg/api/v2/inventory/data"
        f"?gameId={game_id}&offset={offset}&sortType={sort_type}"
        f"&searchValue=&context=trade&fresh=false&limit={limit}&isForStore=0"
    )

    ua = get_user_agent()
    headers = {
        "User-Agent": ua,
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "ru,en;q=0.9,en-GB;q=0.8,en-US;q=0.7",
        "Referer": "https://tradeit.gg/",
        "sec-ch-ua": '"Microsoft Edge";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
    }

    current_proxy_str = proxy_pool.acquire_next() if proxy_pool.has_proxies() else None

    for attempt in range(1, max_retries + 1):
        proxy_dict = format_proxy(current_proxy_str) if current_proxy_str else {}

        if verbose:
            proxy_label = current_proxy_str.split(":")[0] if current_proxy_str else "no-proxy"
            print(f"[Worker] offset={offset} | попытка {attempt}/{max_retries} | proxy={proxy_label}")

        try:
            resp = requests.get(
                url,
                headers=headers,
                proxies=proxy_dict,
                timeout=timeout,
            )

            if resp.status_code == 200:
                data = resp.json()
                items = data.get("items", [])
                # В динамичном инвентаре страница может прийти короче лимита, это не всегда конец.
                # Завершаем только когда страница реально пустая.
                end_reached = len(items) == 0
                return (items, end_reached)

            elif resp.status_code == 429:
                if verbose:
                    print(f"      [429] Rate limit на offset={offset}. Ждём 10 сек...")
                time.sleep(10)
                continue

            elif resp.status_code in (403, 401):
                if verbose:
                    print(f"      [{resp.status_code}] Ошибка доступа на offset={offset}. Меняем прокси...")
                current_proxy_str = proxy_pool.rotate()
                if current_proxy_str is None and not proxy_pool.has_proxies():
                    return None
                continue

            elif resp.status_code == 503:
                if verbose:
                    print(f"      [503] Сайт недоступен. Ждём 5 сек...")
                time.sleep(5)
                continue

            else:
                if verbose:
                    print(f"      [HTTP {resp.status_code}] offset={offset}")
                time.sleep(2)

        except requests.exceptions.ProxyError:
            if verbose:
                print(f"      [ProxyError] offset={offset}. Меняем прокси...")
            current_proxy_str = proxy_pool.rotate()
            if current_proxy_str is None and proxy_pool.has_proxies():
                return None
            time.sleep(1)
            continue

        except requests.exceptions.ConnectionError:
            if verbose:
                print(f"      [ConnectionError] offset={offset}. Меняем прокси...")
            current_proxy_str = proxy_pool.rotate()
            time.sleep(1)
            continue

        except requests.exceptions.Timeout:
            if verbose:
                print(f"      [Timeout] offset={offset}")
            time.sleep(2)
            continue

        except requests.exceptions.RequestException as e:
            if verbose:
                print(f"      [RequestException] {type(e).__name__} на offset={offset}")
            time.sleep(2)
            continue

    print(f"[FAIL] offset={offset} — не удалось получить данные после {max_retries} попыток.")
    return ([], False)

# -------------------------------------------------------------
# Excel
# -------------------------------------------------------------

def _save_workbook_safe(wb: Workbook, path: str, retries: int = 5):
    for i in range(retries):
        try:
            wb.save(path)
            return
        except PermissionError:
            print(f"[Writer] Файл {path} занят. Попытка {i+1}/{retries} через 5 сек...")
            time.sleep(5)
        except Exception as e:
            print(f"[Writer] Ошибка сохранения: {e}")
            return
    print(f"[Writer] Не удалось сохранить {path} после {retries} попыток.")


def save_to_excel(items: List[Dict[str, Any]], out_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tradeit Inventory"

    headers = ["Name", "Price (Trade)"]
    ws.append(headers)

    seen_ids = set()
    saved = 0

    for item in items:
        item_id = item.get("id") or item.get("groupId")
        if item_id in seen_ids:
            continue
        seen_ids.add(item_id)

        name = item.get("name", "Unknown")
        trade_price = format_price(item.get("priceForTrade") or item.get("price"))

        ws.append([name, trade_price])
        saved += 1

    _save_workbook_safe(wb, out_path)
    print(f"Сохранено {saved} предметов в {out_path}.")

# -------------------------------------------------------------
# Main
# -------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Парсер инвентаря сайта Tradeit.gg (через API)."
    )
    ap.add_argument("--game-id", type=int, default=252490,
                    help="ID игры: 252490=Rust, 730=CS:GO (по умолчанию: 252490)")
    ap.add_argument("--sort", default="Popularity",
                    choices=["Popularity", "Price", "PriceReversed", "Name"],
                    help="Тип сортировки (по умолчанию: Popularity)")
    ap.add_argument("--limit", type=int, default=160,
                    help="Предметов на страницу (по умолчанию: 160)")
    ap.add_argument("--workers", type=int, default=5,
                    help="Число параллельных потоков (по умолчанию: 5)")
    ap.add_argument("--task-retries", type=int, default=4,
                    help="Попыток на одну страницу (по умолчанию: 4)")
    ap.add_argument("--proxies", default="",
                    help="Файл с прокси (ip:port:user:pass)")
    ap.add_argument("--proxy-passes", type=int, default=2,
                    help="Сколько раз пройти по списку прокси (по умолчанию: 2)")
    ap.add_argument("--timeout", type=int, default=15,
                    help="Таймаут запроса в секундах")
    ap.add_argument("--out", default="tradeit_items.xlsx",
                    help="Выходной файл Excel")
    ap.add_argument("--verbose", action="store_true",
                    help="Подробные логи")

    args = ap.parse_args()

    # Прокси
    proxy_list = []
    if args.proxies:
        proxy_list = load_proxies(args.proxies)
        print(f"Загружено {len(proxy_list)} прокси.")
    proxy_pool = ProxyPool(proxy_list, max_passes=args.proxy_passes)

    print(
        f"Старт парсинга Tradeit.gg "
        f"[gameId={args.game_id} | workers={args.workers} | limit={args.limit} | sort={args.sort}]"
    )

    results_by_offset: Dict[int, list] = {}
    highest_offset_queued = (args.workers - 1) * args.limit
    pages_processed = 0
    end_reached = False

    with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as executor:
        # Запускаем первые N страниц
        futures = {}
        for i in range(args.workers):
            offset = i * args.limit
            fut = executor.submit(
                worker_task,
                offset, args.limit, args.game_id, args.sort,
                proxy_pool, args.task_retries, args.timeout, args.verbose,
            )
            futures[fut] = offset

        while futures:
            done, _ = concurrent.futures.wait(
                futures.keys(), return_when=concurrent.futures.FIRST_COMPLETED
            )

            for fut in done:
                offset = futures.pop(fut)
                pages_processed += 1

                try:
                    result = fut.result()
                except Exception as e:
                    print(f"[Error] offset={offset}: {e}")
                    result = None

                if result is None:
                    print("[Fatal] Прекращаем сбор из-за фатальной ошибки.")
                    end_reached = True
                    continue

                items, batch_end = result
                results_by_offset[offset] = items
                total = sum(len(v) for v in results_by_offset.values())
                print(f"[Прогресс] Страниц: {pages_processed} | Предметов: {total} | offset={offset}")

                if batch_end:
                    end_reached = True

                if not end_reached:
                    highest_offset_queued += args.limit
                    new_fut = executor.submit(
                        worker_task,
                        highest_offset_queued, args.limit, args.game_id, args.sort,
                        proxy_pool, args.task_retries, args.timeout, args.verbose,
                    )
                    futures[new_fut] = highest_offset_queued

            time.sleep(0.05)

    # Собираем в порядке offset
    all_items = []
    for off in sorted(results_by_offset.keys()):
        all_items.extend(results_by_offset[off])

    print(f"\nСбор завершён. Всего предметов: {len(all_items)}")

    if all_items:
        save_to_excel(all_items, args.out)
    else:
        print("Нет данных для сохранения.")


if __name__ == "__main__":
    main()
