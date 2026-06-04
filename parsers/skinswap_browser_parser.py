#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
skinswap_browser_parser.py

Playwright открывает skinswap.com, скроллит инвентарь и перехватывает ответы API.
Cloudflare проходится через реальный браузер. Профиль Chrome сохраняется.

Запуск:
  python skinswap_browser_parser.py --out skinswap_items.xlsx --verbose

Ключи:
  --headless          Без видимого окна (не рекомендуется)
  --user-data-dir     Профиль Chrome (CF-куки сохраняются между запусками)
  --scroll-pause      Пауза между скроллами (сек, по умолчанию 1.0)
  --stable-rounds     Кол-во скроллов без новых предметов для остановки
"""

import argparse
import asyncio
from typing import Dict, List, Set

from openpyxl import Workbook
from playwright.async_api import async_playwright, Page, BrowserContext, Response

# ---------------------------------------------------------------------------
# Конфигурация
# ---------------------------------------------------------------------------

INVENTORY_URL_PATTERN = "api.skinswap.com/api/site/inventory"
SITE_URL = "https://skinswap.com/ru/trade-rust-skins"

# JS: ищет элемент с computed overflow:auto (vue-virtual-scroller инвентаря),
# кликает/фокусирует его и скроллит вниз на 85% высоты.
SCROLL_JS = """
(() => {
    const candidates = Array.from(document.querySelectorAll('*')).filter(el => {
        const st = window.getComputedStyle(el);
        return st.overflow === 'auto' && (el.scrollHeight - el.clientHeight) > 50;
    });
    if (candidates.length === 0) return {found: false};

    const el = candidates[0];
    el.scrollTop += Math.floor(el.clientHeight * 0.85) || 400;
    return {
        found: true,
        scrollTop: el.scrollTop,
        scrollHeight: el.scrollHeight,
        clientHeight: el.clientHeight,
        atBottom: el.scrollTop + el.clientHeight >= el.scrollHeight - 10,
        classes: el.className.slice(0, 80)
    };
})()
"""

# JS: кликает по центру скроллируемого элемента чтобы Vue "активировал" его.
# Это заменяет ручной первый скролл пользователя.
ACTIVATE_JS = """
(() => {
    const candidates = Array.from(document.querySelectorAll('*')).filter(el => {
        const st = window.getComputedStyle(el);
        return st.overflow === 'auto' && (el.scrollHeight - el.clientHeight) > 50;
    });
    if (candidates.length === 0) return false;
    const el = candidates[0];
    el.focus();
    // Небольшой начальный скролл для активации Vue scroll listener
    el.scrollTop = 1;
    setTimeout(() => { el.scrollTop = 0; }, 50);
    return true;
})()
"""

# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def format_price(cents) -> str:
    try:
        return f"{int(cents) / 100:.2f}".replace(".", ",")
    except Exception:
        return ""


def save_to_excel(items: List[Dict], out_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Skinswap Items"
    ws.append(["Name", "Price (Trade)", "Overstock Limit", "Overstock Count"])

    seen_ids: Set = set()
    saved = 0
    for item in items:
        item_id = item.get("id")
        if item_id:
            if item_id in seen_ids:
                continue
            seen_ids.add(item_id)

        name = item.get("name", "Unknown")
        price_cents = (item.get("price") or {}).get("trade")
        price_str = format_price(price_cents) if price_cents is not None else ""
        overstock = item.get("overstock") or {}
        limit = overstock.get("limit", "")
        count = overstock.get("count", "")

        ws.append([name, price_str, limit, count])
        saved += 1

    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    wb.save(out_path)
    print(f"Сохранено {saved} уникальных предметов в {out_path}")


# ---------------------------------------------------------------------------
# Playwright логика
# ---------------------------------------------------------------------------

async def run(args):
    all_items: List[Dict] = []
    seen_offsets: Set[int] = set()

    async def handle_response(response: Response):
        url = response.url
        if INVENTORY_URL_PATTERN not in url:
            return
        if response.status != 200:
            return
        try:
            data = await response.json()
            if not data.get("success"):
                return
            items = data.get("data", [])
            if not isinstance(items, list):
                return

            offset = 0
            for part in url.split("?", 1)[-1].split("&"):
                if part.startswith("offset="):
                    try:
                        offset = int(part.split("=", 1)[1])
                    except ValueError:
                        pass

            if offset in seen_offsets:
                return
            seen_offsets.add(offset)
            all_items.extend(items)
            if args.verbose:
                print(f"[net] offset={offset:>5} | +{len(items)} предметов | итого={len(all_items)}")
        except Exception as e:
            if args.verbose:
                print(f"[net] ошибка: {e}")

    async with async_playwright() as pw:
        launch_kwargs = {
            "headless": args.headless,
            "args": [
                "--disable-blink-features=AutomationControlled",
                "--window-size=1400,900",
                "--lang=ru-RU",
            ],
        }

        if args.user_data_dir:
            context: BrowserContext = await pw.chromium.launch_persistent_context(
                user_data_dir=args.user_data_dir,
                **launch_kwargs,
            )
            page: Page = context.pages[0] if context.pages else await context.new_page()
        else:
            browser = await pw.chromium.launch(**launch_kwargs)
            context = await browser.new_context(viewport={"width": 1400, "height": 900})
            page = await context.new_page()

        await page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            delete window.__playwright;
        """)

        page.on("response", lambda r: asyncio.ensure_future(handle_response(r)))

        print(f"Открываем {SITE_URL} ...")
        await page.goto(SITE_URL, wait_until="domcontentloaded", timeout=60_000)

        # Ждём CF challenge
        print("Ожидаем загрузки страницы...")
        for i in range(args.challenge_wait):
            title = await page.title()
            if "just a moment" in title.lower() or "checking" in title.lower():
                if i % 5 == 0:
                    print(f"  CF challenge... ({i}s)")
                await asyncio.sleep(1)
            else:
                break

        print(f"Страница загружена: {await page.title()}")

        # Ждём первых предметов
        print("Ждём первой загрузки инвентаря...")
        for _ in range(80):
            if all_items:
                break
            await asyncio.sleep(0.1)
        await asyncio.sleep(0.5)

        if not all_items:
            print("[warn] Первая загрузка не произошла за 8 сек.")
        else:
            print(f"Первая загрузка: {len(all_items)} предметов.")

        # Активируем скролл-контейнер (заменяет ручной первый скролл)
        activated = await page.evaluate(ACTIVATE_JS)
        if args.verbose:
            print(f"Активация скроллера: {'OK' if activated else 'не найден'}")
        await asyncio.sleep(0.3)

        # ---------------------------------------------------------------------------
        # Цикл скроллинга
        # ---------------------------------------------------------------------------
        print(f"\nСкроллинг инвентаря (пауза {args.scroll_pause}s)...")
        stable = 0
        step = 0
        prev_total = 0
        last_scroll_height = 0

        while True:
            if args.max_items > 0 and len(all_items) >= args.max_items:
                print(f"Достигнут лимит: {args.max_items}")
                break

            result = await page.evaluate(SCROLL_JS)

            if not result or not result.get("found"):
                await page.evaluate("window.scrollBy(0, 600)")
                if args.verbose:
                    print(f"[scroll] шаг={step} | скроллер не найден, window")
            else:
                scroll_height = result.get("scrollHeight", 0)
                at_bottom = result.get("atBottom", False)
                scroll_top = result.get("scrollTop", 0)

                if args.verbose:
                    print(f"[scroll] шаг={step} | scrollTop={scroll_top} | "
                          f"scrollH={scroll_height} | atBottom={at_bottom} | "
                          f"итого={len(all_items)}")

                last_scroll_height = scroll_height

            await asyncio.sleep(args.scroll_pause)
            step += 1

            total_now = len(all_items)
            gained = total_now - prev_total
            prev_total = total_now

            if gained > 0:
                stable = 0
                if not args.verbose:
                    print(f"[прогресс] шаг={step} | +{gained} | итого={total_now}")
            else:
                stable += 1
                if args.verbose:
                    print(f"  → без новых: {stable}/{args.stable_rounds}")

            if stable >= args.stable_rounds:
                print(f"Конец инвентаря ({args.stable_rounds} шагов без новых предметов).")
                break

        print(f"\nСбор завершён. Всего предметов: {len(all_items)}")
        await context.close()

    return all_items


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Парсер инвентаря Skinswap через Playwright (скролл + перехват сети).",
        formatter_class=argparse.RawTextHelpFormatter
    )
    ap.add_argument("--out", default="skinswap_items.xlsx", help="Выходной файл Excel")
    ap.add_argument("--max-items", type=int, default=0, help="Лимит предметов (0=без лимита)")
    ap.add_argument(
        "--scroll-pause", type=float, default=1.0,
        help="Пауза между скроллами (сек, по умолчанию 1.0)"
    )
    ap.add_argument(
        "--stable-rounds", type=int, default=5,
        help="Остановиться если N шагов подряд нет новых предметов (по умолчанию 5)"
    )
    ap.add_argument(
        "--challenge-wait", type=int, default=60,
        help="Сколько секунд ждать CF challenge"
    )
    ap.add_argument("--headless", action="store_true", help="Headless режим (не рекомендуется)")
    ap.add_argument(
        "--user-data-dir", default="./chrome_profile",
        help="Путь к профилю Chrome (куки CF сохраняются между запусками)"
    )
    ap.add_argument("--verbose", action="store_true", help="Подробные логи")

    args = ap.parse_args()
    items = asyncio.run(run(args))

    if items:
        save_to_excel(items, args.out)
    else:
        print("Предметы не найдены.")


if __name__ == "__main__":
    main()
