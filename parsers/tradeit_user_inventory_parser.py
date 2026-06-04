import os
import sys
import time
import json
import argparse
from typing import Dict, Any, List, Optional

import requests
from openpyxl import Workbook
import undetected_chromedriver as uc


# -------------------------------------------------------------
# Utils
# -------------------------------------------------------------

def format_price(cents: Optional[int]) -> str:
    if cents is None:
        return ""
    return f"{cents / 100.0:.2f}".replace(".", ",")


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def get_default_user_agent() -> str:
    return (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36 Edg/147.0.0.0"
    )


def game_slug(game_id: int) -> str:
    # Используем slug только для страницы реферера/авторизации.
    mapping = {
        252490: "rust",
        730: "csgo",
        570: "dota2",
    }
    return mapping.get(game_id, "rust")


def inventory_url(game_id: int) -> str:
    return f"https://tradeit.gg/api/v2/inventory/my/data?gameId={game_id}&fresh=0&listing=0"


def _get_chrome_major_version() -> Optional[int]:
    """Пробует определить major версию установленного Chrome (Windows)."""
    import subprocess

    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in paths:
        if not os.path.exists(path):
            continue
        try:
            out = subprocess.check_output(
                ["powershell", "-Command", f'(Get-Item "{path}").VersionInfo.ProductVersion'],
                stderr=subprocess.DEVNULL,
                timeout=5,
            ).decode().strip()
            return int(out.split(".")[0])
        except Exception:
            continue
    return None


# -------------------------------------------------------------
# Session
# -------------------------------------------------------------

def create_api_session(session_data: Dict[str, Any], game_id: int) -> requests.Session:
    session = requests.Session()

    ua = session_data.get("user_agent") or get_default_user_agent()
    cookies = session_data.get("cookies") or {}

    session.headers.update(
        {
            "User-Agent": ua,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "ru,en;q=0.9,en-GB;q=0.8,en-US;q=0.7",
            "Referer": f"https://tradeit.gg/ru/{game_slug(game_id)}/trade",
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-origin",
        }
    )

    for name, value in cookies.items():
        session.cookies.set(name, value, domain=".tradeit.gg")

    return session


def load_session(session_file: str, max_age_hours: int) -> Dict[str, Any]:
    if not os.path.exists(session_file):
        return {}
    try:
        with open(session_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"Ошибка чтения сессии: {e}")
        return {}

    ts = float(data.get("timestamp", 0))
    max_age_sec = max_age_hours * 3600
    if ts and (time.time() - ts > max_age_sec):
        print(f"Сессия старше {max_age_hours} часов. Нужна переавторизация.")
        return {}

    if not data.get("cookies"):
        return {}

    return data


def save_session(session_file: str, session_data: Dict[str, Any]) -> None:
    with open(session_file, "w", encoding="utf-8") as f:
        json.dump(session_data, f, ensure_ascii=False, indent=2)


def check_auth_validity(session_data: Dict[str, Any], game_id: int, timeout: int, verbose: bool = False) -> bool:
    url = inventory_url(game_id)
    session = create_api_session(session_data, game_id)

    try:
        resp = session.get(url, timeout=timeout)
    except requests.exceptions.RequestException as e:
        if verbose:
            print(f"[AuthCheck] Сетевая ошибка: {type(e).__name__}")
        return False

    if resp.status_code != 200:
        if verbose:
            print(f"[AuthCheck] HTTP {resp.status_code}")
        return False

    try:
        payload = resp.json()
    except ValueError:
        if verbose:
            print("[AuthCheck] Не JSON ответ")
        return False

    ok = bool(payload.get("success") is True and isinstance(payload.get("items"), dict))
    if verbose:
        print(f"[AuthCheck] {'OK' if ok else 'FAIL'}")
    return ok


def get_session_via_selenium(session_file: str, game_id: int, timeout_sec: int = 300) -> Dict[str, Any]:
    print("Открываю браузер для авторизации Tradeit через Steam...")
    print("После входа сессия сохранится автоматически.")

    options = uc.ChromeOptions()
    chrome_major = _get_chrome_major_version()

    try:
        if chrome_major:
            driver = uc.Chrome(options=options, version_main=chrome_major)
        else:
            driver = uc.Chrome(options=options)
    except Exception as e:
        print(f"Ошибка запуска UC с version_main={chrome_major}: {e}. Пробую без версии...")
        driver = uc.Chrome(options=options)

    try:
        driver.get(f"https://tradeit.gg/ru/{game_slug(game_id)}/trade")
        ua = driver.execute_script("return navigator.userAgent;") or get_default_user_agent()

        deadline = time.time() + timeout_sec
        while time.time() < deadline:
            cookies = {c["name"]: c["value"] for c in driver.get_cookies() if c.get("name")}

            # sessionid не гарантирует авторизацию, поэтому валидируем реальным API вызовом.
            if "sessionid" in cookies:
                candidate = {
                    "cookies": cookies,
                    "user_agent": ua,
                    "timestamp": time.time(),
                }
                if check_auth_validity(candidate, game_id, timeout=12, verbose=False):
                    save_session(session_file, candidate)
                    print(f"Сессия сохранена в {session_file}")
                    return candidate

            time.sleep(3)

        print("Время ожидания авторизации истекло (5 минут).")
        return {}
    finally:
        driver.quit()


def ensure_session(session_file: str, game_id: int, timeout: int, max_age_hours: int, verbose: bool) -> Dict[str, Any]:
    session_data = load_session(session_file, max_age_hours=max_age_hours)
    if session_data:
        print("Найдена сохраненная сессия. Проверяю валидность...")
        if check_auth_validity(session_data, game_id, timeout=timeout, verbose=verbose):
            print("Сессия активна.")
            return session_data
        print("Сессия невалидна. Нужна повторная авторизация.")

    return get_session_via_selenium(session_file=session_file, game_id=game_id)


# -------------------------------------------------------------
# Fetch + Normalize
# -------------------------------------------------------------

def fetch_user_inventory(
    session_data: Dict[str, Any],
    game_id: int,
    timeout: int,
    max_retries: int,
    verbose: bool
) -> Optional[Dict[str, List[Dict[str, Any]]]]:
    url = inventory_url(game_id)

    for attempt in range(1, max_retries + 1):
        session = create_api_session(session_data, game_id)
        if verbose:
            print(f"[API] Попытка {attempt}/{max_retries}")

        try:
            resp = session.get(url, timeout=timeout)
        except requests.exceptions.Timeout:
            if verbose:
                print("      [Timeout] Жду 2 сек и повторяю...")
            time.sleep(2)
            continue
        except requests.exceptions.RequestException as e:
            if verbose:
                print(f"      [Network] {type(e).__name__}. Жду 2 сек...")
            time.sleep(2)
            continue

        if resp.status_code == 200:
            try:
                data = resp.json()
            except ValueError:
                if verbose:
                    print("      [Ошибка] Ответ не JSON")
                time.sleep(2)
                continue

            if data.get("success") is True and isinstance(data.get("items"), dict):
                return data.get("items", {})

            if verbose:
                print(f"      [Ошибка] success=false. Ответ: {str(data)[:400]}")
            return None

        if resp.status_code in (401, 403):
            print(f"      [Auth Error] HTTP {resp.status_code}. Сессия устарела или невалидна.")
            return None

        if resp.status_code == 429:
            if verbose:
                print("      [429] Rate limit. Жду 5 сек...")
            time.sleep(5)
            continue

        if resp.status_code >= 500:
            if verbose:
                print(f"      [HTTP {resp.status_code}] Серверная ошибка. Жду 3 сек...")
            time.sleep(3)
            continue

        if verbose:
            print(f"      [HTTP {resp.status_code}]")
        time.sleep(2)

    print("Не удалось получить личный инвентарь после всех попыток.")
    return None


def flatten_items(grouped_items: Dict[str, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for _, assets in grouped_items.items():
        if not isinstance(assets, list):
            continue
        out.extend(assets)
    return out


def classify_item(item: Dict[str, Any], min_trade_price_cents: int = 3) -> Dict[str, bool]:
    """
    Определяем, можно ли реально депнуть предмет в бот.
    Логика основана на наблюдениях из API:
    - tradable == 1
    - enableDeposit == 1
    - maxDeposit > 0
    - price >= min_trade_price_cents (по наблюдениям, вещи <= $0.02 часто недоступны)
    - отдельно считаем избыток (maxDeposit <= 0 или currentStock >= botMaxQuantity)
    """
    tradable_ok = item.get("tradable") == 1
    deposit_enabled = item.get("enableDeposit") == 1

    max_deposit = to_float(item.get("maxDeposit"))
    deficit = to_float(item.get("itemDeficit"))
    sell_price = to_float(item.get("sellPrice"))
    trade_price = to_float(item.get("price"))

    current_stock = to_float(item.get("currentStock"))
    bot_max = to_float(item.get("botMaxQuantity"))
    if bot_max is None:
        bot_max = to_float(item.get("wantedStock"))

    has_overstock = False
    if max_deposit is not None and max_deposit <= 0:
        has_overstock = True
    elif current_stock is not None and bot_max is not None and bot_max >= 0 and current_stock >= bot_max:
        has_overstock = True

    has_nonpositive_deficit = deficit is not None and deficit <= 0
    has_no_sell_price = sell_price is not None and sell_price <= 0
    has_too_low_trade_price = trade_price is not None and trade_price < float(min_trade_price_cents)

    # Ключевой флаг доступности депозита.
    can_deposit = (
        tradable_ok
        and deposit_enabled
        and (max_deposit is None or max_deposit > 0)
        and not has_too_low_trade_price
        and not has_overstock
    )

    unavailable = not can_deposit
    unavailable_and_overstock = unavailable and has_overstock

    return {
        "can_deposit": can_deposit,
        "tradable_ok": tradable_ok,
        "deposit_enabled": deposit_enabled,
        "has_overstock": has_overstock,
        "has_nonpositive_deficit": has_nonpositive_deficit,
        "has_no_sell_price": has_no_sell_price,
        "has_too_low_trade_price": has_too_low_trade_price,
        "unavailable": unavailable,
        "unavailable_and_overstock": unavailable_and_overstock,
    }


def save_to_excel(items: List[Dict[str, Any]], out_path: str, min_trade_price_cents: int = 3) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "My Tradeit Inventory"
    ws.append(["Name", "Price (Trade)", "In Bot Now", "Bot Max"])

    seen = set()
    saved = 0
    skipped = 0
    skipped_unavailable = 0
    skipped_overstock = 0
    skipped_unavailable_and_overstock = 0
    skipped_not_tradable = 0
    skipped_deposit_disabled = 0
    skipped_no_sell_price = 0
    skipped_too_low_trade_price = 0
    skipped_nonpositive_deficit = 0

    for item in items:
        unique_id = item.get("assetId") or item.get("id") or item.get("_id")
        if unique_id in seen:
            continue
        if unique_id is not None:
            seen.add(unique_id)

        status = classify_item(item, min_trade_price_cents=min_trade_price_cents)
        if not status["can_deposit"]:
            skipped += 1
            if status["unavailable"]:
                skipped_unavailable += 1
            if status["has_overstock"]:
                skipped_overstock += 1
            if status["unavailable_and_overstock"]:
                skipped_unavailable_and_overstock += 1
            if not status["tradable_ok"]:
                skipped_not_tradable += 1
            if not status["deposit_enabled"]:
                skipped_deposit_disabled += 1
            if status["has_no_sell_price"]:
                skipped_no_sell_price += 1
            if status["has_too_low_trade_price"]:
                skipped_too_low_trade_price += 1
            if status["has_nonpositive_deficit"]:
                skipped_nonpositive_deficit += 1
            continue

        name = item.get("name", "Unknown")
        trade_price = format_price(item.get("price"))
        current_in_bot = item.get("currentStock")
        if current_in_bot is None:
            current_in_bot = item.get("visibleItemStock", "")

        bot_max = item.get("botMaxQuantity")
        if bot_max is None:
            bot_max = item.get("wantedStock", "")

        ws.append([name, trade_price, current_in_bot, bot_max])
        saved += 1

    for i in range(5):
        try:
            wb.save(out_path)
            print(f"Сохранено {saved} предметов в {out_path}.")
            print(
                "Отфильтровано: "
                f"{skipped} | "
                f"недоступно={skipped_unavailable}, "
                f"избыток={skipped_overstock}, "
                f"недоступно+избыток={skipped_unavailable_and_overstock}, "
                f"tradable!=1={skipped_not_tradable}, "
                f"enableDeposit!=1={skipped_deposit_disabled}, "
                f"sellPrice<=0(справочно)={skipped_no_sell_price}, "
                f"price<{min_trade_price_cents}c={skipped_too_low_trade_price}, "
                f"itemDeficit<=0(справочно)={skipped_nonpositive_deficit}"
            )
            return
        except PermissionError:
            print(f"[Writer] Файл {out_path} занят. Попытка {i+1}/5 через 5 сек...")
            time.sleep(5)
        except Exception as e:
            print(f"[Writer] Ошибка сохранения: {e}")
            return

    print(f"[Writer] Не удалось сохранить {out_path} после 5 попыток.")


# -------------------------------------------------------------
# Main
# -------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Парсер личного инвентаря Tradeit.gg (Selenium авторизация -> requests API)."
    )
    ap.add_argument("--game-id", type=int, default=252490, help="ID игры: 252490=Rust, 730=CS:GO")
    ap.add_argument("--out", default="tradeit_user_items.xlsx", help="Выходной xlsx файл")
    ap.add_argument("--user-profile", default="default", help="Имя профиля сессии")
    ap.add_argument("--timeout", type=int, default=15, help="Таймаут запроса (сек)")
    ap.add_argument("--task-retries", type=int, default=5, help="Попыток загрузки API")
    ap.add_argument("--session-max-age-hours", type=int, default=168, help="Сколько часов хранить сессию")
    ap.add_argument(
        "--min-trade-price",
        type=int,
        default=3,
        help="Минимальная цена предмета для включения в Excel (в центах, по умолчанию: 3; отсекает <= $0.02)",
    )
    ap.add_argument("--verbose", action="store_true", help="Подробные логи")
    args = ap.parse_args()

    session_file = f"tradeit_session_{args.user_profile}.json"

    session_data = ensure_session(
        session_file=session_file,
        game_id=args.game_id,
        timeout=args.timeout,
        max_age_hours=args.session_max_age_hours,
        verbose=args.verbose,
    )

    if not session_data:
        print("Не удалось получить валидную сессию. Выход.")
        sys.exit(1)

    grouped_items = fetch_user_inventory(
        session_data=session_data,
        game_id=args.game_id,
        timeout=args.timeout,
        max_retries=args.task_retries,
        verbose=args.verbose,
    )

    if grouped_items is None:
        # Если API сказал auth error - удаляем сессию, чтобы следующий запуск форсировал login.
        if os.path.exists(session_file):
            try:
                os.remove(session_file)
                print("Сессия удалена. На следующем запуске будет новая авторизация.")
            except OSError:
                pass
        sys.exit(1)

    items = flatten_items(grouped_items)
    print(f"Сбор завершен. Получено предметов: {len(items)}")

    if not items:
        print("Инвентарь пуст или нет данных для сохранения.")
        return

    save_to_excel(items, args.out, min_trade_price_cents=args.min_trade_price)


if __name__ == "__main__":
    main()
