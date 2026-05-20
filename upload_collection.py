#!/usr/bin/env python3
"""
Загрузчик подборок рецептов на gastronom.ru.

Подключается к УЖЕ ЗАПУЩЕННОМУ Chrome через CDP (тот же chrome-debug.bat).

Перед запуском:
1. Запусти Chrome через chrome-debug.bat (Windows) или chrome-debug.sh (macOS)
2. Залогинься на admin.gastronom.ru
3. Открой форму создания/редактирования подборки
4. Запусти: python upload_collection.py

Входной файл: collections.json (создаётся через docx_to_collection_json.py)

Алгоритм заполнения "Основной текст":
  1. Кнопка W (Word import): вставляет HTML-текст (заголовки + описания рецептов)
  2. Сайт обрабатывает и копирует результат в буфер обмена
  3. Ctrl+V в поле "Основной текст"
  4. Затем для каждого рецепта добавляем блок изображения через "+"
"""

import asyncio
import html as html_lib
import json
import sys
from pathlib import Path

if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

try:
    from playwright.async_api import async_playwright, Page, TimeoutError as PWTimeout
except ImportError:
    print("Нужно установить Playwright: pip install playwright && playwright install chromium")
    sys.exit(1)

CDP_URL                = "http://localhost:9222"
IN_FILE                = Path("collections.json")
COLLECTION_URL_PATTERN = "admin.gastronom.ru"
CTRL                   = "Meta" if sys.platform == "darwin" else "Control"


# ── БУФЕР ОБМЕНА ─────────────────────────────────────────────────────────────

async def set_clipboard_html(page: Page, html: str) -> None:
    """
    Кладёт HTML в системный буфер обмена через contenteditable + execCommand.
    В отличие от textarea+execCommand (который даёт только text/plain),
    contenteditable копирует как text/html — это нужно для W-кнопки,
    которая читает именно HTML-формат и распознаёт <h2>, <p> и т.д.
    """
    await page.evaluate("""(raw_html) => {
        const div = document.createElement('div');
        div.contentEditable = 'true';
        div.style.cssText = 'position:fixed;top:-9999px;left:-9999px;opacity:0;';
        div.innerHTML = raw_html;
        document.body.appendChild(div);
        const range = document.createRange();
        range.selectNodeContents(div);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        document.execCommand('copy');
        sel.removeAllRanges();
        document.body.removeChild(div);
    }""", html)
    await page.wait_for_timeout(150)


async def set_clipboard_plain(page: Page, text: str) -> None:
    """Кладёт plain text в буфер обмена."""
    await page.evaluate("""(t) => {
        const ta = document.createElement('textarea');
        ta.value = t;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
    }""", text)


# ── W-КНОПКА (WORD IMPORT) ───────────────────────────────────────────────────

async def fill_section_via_w(page: Page, section_name: str, html_content: str) -> bool:
    """
    Заполняет секцию Editor.js через W-кнопку:
    1. Кладёт HTML в clipboard
    2. Находит W-кнопку нужной секции (по её заголовку) и кликает
    3. Ctrl+V в попапе → сайт обрабатывает и копирует результат в буфер
    4. Ctrl+V в редактор той же секции

    Возвращает True при успехе.
    """
    await set_clipboard_html(page, html_content)

    # Находим W-кнопку и редактор нужной секции через JS
    clicked = await page.evaluate("""(name) => {
        // Ищем элемент, чей прямой текст совпадает с именем секции
        for (const el of document.querySelectorAll('*')) {
            const direct = Array.from(el.childNodes)
                .filter(n => n.nodeType === 3)
                .map(n => n.textContent.trim())
                .join('');
            if (direct !== name) continue;
            // Нашли метку — ищем W-кнопку в ближайшем родителе-контейнере
            let node = el;
            for (let i = 0; i < 12; i++) {
                node = node.parentElement;
                if (!node) break;
                const use = node.querySelector('use[href="#icon-microsoft-word"]');
                if (use) {
                    const btn = use.closest('button') || use.parentElement;
                    btn.click();
                    return true;
                }
            }
        }
        return false;
    }""", section_name)

    if not clicked:
        print(f"  ✗ Не нашёл W-кнопку секции «{section_name}»")
        return False

    await page.wait_for_timeout(1000)

    # Фокусируем textarea попапа через JS (textarea есть только в попапе, не в редакторе)
    focused = await page.evaluate("""() => {
        for (const el of document.querySelectorAll('textarea')) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && r.height > 0) {
                el.focus();
                el.click();
                return true;
            }
        }
        return false;
    }""")

    if not focused:
        print(f"  ✗ Попап W открылся, но не нашёл поле для вставки")
        await page.keyboard.press("Escape")
        return False

    await page.wait_for_timeout(200)
    await page.keyboard.press(f"{CTRL}+v")

    # Ждём пока сайт обработает текст и скопирует результат в буфер (~3 сек)
    await page.wait_for_timeout(3000)

    # Находим редактор нужной секции и вставляем из буфера
    pasted = await page.evaluate("""(name) => {
        for (const el of document.querySelectorAll('*')) {
            const direct = Array.from(el.childNodes)
                .filter(n => n.nodeType === 3)
                .map(n => n.textContent.trim())
                .join('');
            if (direct !== name) continue;
            let node = el;
            for (let i = 0; i < 12; i++) {
                node = node.parentElement;
                if (!node) break;
                const editor = node.querySelector('[contenteditable="true"]');
                if (editor) {
                    editor.focus();
                    editor.click();
                    return true;
                }
            }
        }
        return false;
    }""", section_name)

    if not pasted:
        # Фолбек — любой первый contenteditable
        try:
            loc = page.locator('div[contenteditable="true"]').first
            if await loc.count() > 0 and await loc.is_visible(timeout=800):
                await loc.click()
                pasted = True
        except Exception:
            pass

    if not pasted:
        print(f"  ✗ Не нашёл редактор секции «{section_name}» для вставки")
        return False

    await page.wait_for_timeout(200)
    await page.keyboard.press(f"{CTRL}+v")
    await page.wait_for_timeout(1500)
    print(f"  ✓ «{section_name}» заполнена")
    return True


# ── EDITOR.JS БЛОКИ (+) ───────────────────────────────────────────────────────

async def editor_click_plus(page: Page) -> bool:
    """Имитирует hover над последним блоком Editor.js и нажимает '+'."""
    # Hover над последним contenteditable-блоком
    try:
        last_block = page.locator('.ce-block').last
        if await last_block.count() > 0:
            box = await last_block.bounding_box()
            if box:
                cx = box['x'] + box['width'] / 2
                cy = box['y'] + box['height'] / 2
                await page.mouse.move(50, 50)
                await page.wait_for_timeout(80)
                await page.mouse.click(cx, cy)
                await page.wait_for_timeout(350)
                for ox, oy in [(0,0),(4,0),(-4,0),(0,4)]:
                    await page.mouse.move(cx + ox, cy + oy, steps=2)
                    await page.wait_for_timeout(40)
                await page.wait_for_timeout(350)
    except Exception:
        pass

    # Нажимаем '+'
    return bool(await page.evaluate("""() => {
        for (const sel of [
            '.ce-toolbar--opened .ce-toolbar__plus',
            '.ce-toolbar__plus',
            '[class*="toolbar__plus"]',
        ]) {
            for (const el of document.querySelectorAll(sel)) {
                const r = el.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) { el.click(); return true; }
            }
        }
        // Кнопка с текстом "+"
        for (const el of document.querySelectorAll('button, div, span')) {
            if (el.children.length > 0) continue;
            if (el.textContent.trim() !== '+') continue;
            const r = el.getBoundingClientRect();
            if (r.width > 0 && r.width < 60 && r.height > 0) { el.click(); return true; }
        }
        return false;
    }"""))


async def toolbox_select(page: Page, block_name: str) -> bool:
    """Выбирает блок из тулбокса Editor.js по названию."""
    # Пробуем вбить в Filter-поиск тулбокса
    try:
        filter_inp = page.locator('input[placeholder="Filter"]').first
        if await filter_inp.is_visible(timeout=800):
            await filter_inp.fill("")
            await filter_inp.type(block_name, delay=30)
            await page.wait_for_timeout(700)
    except Exception:
        pass

    # Playwright CSS-селекторы
    for sel in [
        f'.ce-popover-item__title:text-is("{block_name}")',
        f'.ce-popover-item:has-text("{block_name}")',
        f'.ce-popover__item:has-text("{block_name}")',
        f'[data-item-name="{block_name.lower()}"]',
        f'[data-item-name*="{block_name.lower()}"]',
        f'button:has-text("{block_name}")',
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.count() > 0 and await loc.is_visible(timeout=500):
                await loc.click(timeout=1500)
                await page.wait_for_timeout(500)
                return True
        except Exception:
            pass

    # JS: ищем сначала внутри попапа, потом глобально по точному тексту
    clicked = await page.evaluate("""(term) => {
        const termL = term.toLowerCase();
        // Ищем контейнер тулбокса/попапа
        for (const popSel of [
            '.ce-popover', '.ce-toolbox', '[class*="ce-popover"]',
            '[class*="toolbox"]', '[role="listbox"]', '[role="menu"]'
        ]) {
            for (const pop of document.querySelectorAll(popSel)) {
                const r = pop.getBoundingClientRect();
                if (!r.width || !r.height) continue;
                for (const el of pop.querySelectorAll('*')) {
                    if (el.textContent.trim().toLowerCase() !== termL) continue;
                    const er = el.getBoundingClientRect();
                    if (er.width > 0 && er.height > 0) { el.click(); return true; }
                }
            }
        }
        // Глобальный фолбек: любой видимый элемент с точным текстом
        for (const el of document.querySelectorAll('*')) {
            if (el.textContent.trim().toLowerCase() !== termL) continue;
            const r = el.getBoundingClientRect();
            if (r.width > 0 && r.height > 0 && r.height < 80) { el.click(); return true; }
        }
        return false;
    }""", block_name)

    if clicked:
        await page.wait_for_timeout(500)
        return True

    await page.keyboard.press("Escape")
    return False


async def add_image_block(page: Page, image_uuid: str, recipe_name: str, recipe_url: str) -> str:
    """
    Добавляет блок изображения после описания рецепта:
    1. Находит .ce-block с <h2> = recipe_name, прокручивает в центр экрана
    2. Кликает на следующий блок (описание), нажимает «+» → выбирает «Изображение»
    3. Нажимает «Загрузить существующее изображение»
    4. Вводит UUID в фильтр, применяет, кликает картинку
    5. Заполняет «Введите ссылку»

    Возвращает "" при успехе или строку с предупреждением.
    """
    if not image_uuid:
        return f"«{recipe_name}»: UUID не задан — добавь изображение вручную"

    # Шаг 1: найти блок-описание (следующий после h2 с именем рецепта)
    #         и прокрутить его в центр экрана ДО получения координат
    scrolled = await page.evaluate("""(name) => {
        const norm = s => s.replace(/\\s+/g, ' ').trim();
        const blocks = Array.from(document.querySelectorAll('.ce-block'));
        for (let i = 0; i < blocks.length; i++) {
            const h2 = blocks[i].querySelector('h2');
            if (!h2 || norm(h2.textContent) !== norm(name)) continue;
            const target = blocks[i + 1] || blocks[i];
            target.scrollIntoView({block: 'center', behavior: 'instant'});
            return true;
        }
        return false;
    }""", recipe_name)

    if not scrolled:
        return f"«{recipe_name}»: не нашёл блок рецепта в редакторе — добавь изображение вручную"

    await page.wait_for_timeout(500)  # дать странице осесть после прокрутки

    # Шаг 2: получить координаты ПОСЛЕ прокрутки (теперь блок точно в viewport'е)
    block_xy = await page.evaluate("""(name) => {
        const norm = s => s.replace(/\\s+/g, ' ').trim();
        const blocks = Array.from(document.querySelectorAll('.ce-block'));
        for (let i = 0; i < blocks.length; i++) {
            const h2 = blocks[i].querySelector('h2');
            if (!h2 || norm(h2.textContent) !== norm(name)) continue;
            const target = blocks[i + 1] || blocks[i];
            const r = target.getBoundingClientRect();
            if (r.width > 0 && r.height > 0)
                return {x: r.left + r.width / 2, y: r.top + r.height / 2, vh: window.innerHeight};
        }
        return null;
    }""", recipe_name)

    if not block_xy:
        return f"«{recipe_name}»: блок не виден после прокрутки — добавь изображение вручную"

    cx, cy = block_xy['x'], block_xy['y']
    print(f"    → блок: x={cx:.0f}, y={cy:.0f}, vh={block_xy['vh']:.0f}")

    # Кликаем на блок → Editor.js показывает тулбар
    await page.mouse.move(50, 50)
    await page.wait_for_timeout(80)
    await page.mouse.click(cx, cy)
    await page.wait_for_timeout(400)
    for ox, oy in [(0, 0), (4, 0), (-4, 0), (0, 4)]:
        await page.mouse.move(cx + ox, cy + oy, steps=2)
        await page.wait_for_timeout(40)
    await page.wait_for_timeout(400)

    # Шаг 3: нажимаем «+» через Playwright locator (до JS-фолбека)
    plus_ok = False
    try:
        plus_loc = page.locator('.ce-toolbar__plus').first
        if await plus_loc.count() > 0 and await plus_loc.is_visible(timeout=1500):
            await plus_loc.click(timeout=2000)
            plus_ok = True
    except Exception:
        pass

    if not plus_ok:
        # JS-фолбек: только Editor.js-специфичные селекторы, без поиска по тексту «+»
        plus_ok = bool(await page.evaluate("""() => {
            for (const sel of [
                '.ce-toolbar--opened .ce-toolbar__plus',
                '.ce-toolbar__plus',
                '[class*="toolbar__plus"]',
            ]) {
                const el = document.querySelector(sel);
                if (!el) continue;
                const r = el.getBoundingClientRect();
                if (r.width > 0 && r.height > 0) { el.click(); return true; }
            }
            return false;
        }"""))

    if not plus_ok:
        return f"«{recipe_name}»: не нашёл кнопку «+» — добавь изображение вручную"
    await page.wait_for_timeout(800)

    # Шаг 3: выбрать «Изображение» из тулбокса
    img_ok = await toolbox_select(page, "Изображение")
    if not img_ok:
        return f"«{recipe_name}»: не нашёл «Изображение» в тулбоксе — добавь вручную"
    await page.wait_for_timeout(1000)

    # Шаг 4: «Загрузить существующее изображение» — data-testid="exist-button"
    clicked_existing = False

    for locator in [
        page.locator('button[data-testid="exist-button"]'),
        page.get_by_test_id("exist-button"),
        page.get_by_text("Загрузить существующее изображение", exact=True),
    ]:
        try:
            loc = locator.last
            if await loc.count() > 0:
                await loc.scroll_into_view_if_needed()
                await loc.click(force=True, timeout=2000)
                clicked_existing = True
                break
        except Exception:
            pass

    if not clicked_existing:
        clicked_existing = bool(await page.evaluate("""() => {
            const btn = document.querySelector('button[data-testid="exist-button"]');
            if (btn) { btn.click(); return true; }
            return false;
        }"""))

    if not clicked_existing:
        return f"«{recipe_name}»: не нашёл «Загрузить существующее изображение» — добавь вручную"
    await page.wait_for_timeout(2000)

    # Шаг 5: раскрываем панель фильтра (клик по «Фильтр ∨»)
    filter_expanded = False
    for sel in [
        'button:has-text("Фильтр")',
        '*:has-text("Фильтр"):not(h1):not(h2):not(h3):not(label)',
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.count() > 0 and await loc.is_visible(timeout=1000):
                await loc.click(timeout=1500)
                await page.wait_for_timeout(500)
                filter_expanded = True
                break
        except Exception:
            pass

    # Вводим UUID в поле «ID, имя или Alt»
    id_input = None
    for sel in [
        'input[placeholder*="ID"]',
        'input[placeholder*="имя или Alt"]',
        'input[placeholder*="ID, имя"]',
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.count() > 0 and await loc.is_visible(timeout=800):
                id_input = loc
                break
        except Exception:
            pass

    if id_input is None:
        # JS-фолбек
        focused = await page.evaluate("""() => {
            for (const el of document.querySelectorAll('input')) {
                const ph = (el.placeholder || '').toLowerCase();
                if (ph.includes('id') || ph.includes('имя') || ph.includes('alt')) {
                    el.focus(); return true;
                }
            }
            return false;
        }""")
        if not focused:
            return f"«{recipe_name}»: не нашёл поле фильтра — вставь UUID {image_uuid} вручную"
        await page.keyboard.press("Control+a")
        await page.keyboard.type(image_uuid, delay=40)
    else:
        await id_input.click(timeout=1000)
        await id_input.fill(image_uuid)

    await page.wait_for_timeout(300)

    # Применяем фильтр
    for sel in [
        'button:has-text("Применить фильтр")',
        'button:has-text("Применить")',
        'button:has-text("Найти")',
        'button[type="submit"]',
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.count() > 0 and await loc.is_visible(timeout=800):
                await loc.click(timeout=1500)
                await page.wait_for_timeout(2000)
                break
        except Exception:
            pass

    # Шаг 6: кликаем по найденной картинке
    img_clicked = False
    for sel in [
        '[class*="media-item"] img', '[class*="gallery"] img',
        '[class*="media-library"] img', '[class*="thumbnail"]',
        '.media-item',
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.count() > 0 and await loc.is_visible(timeout=2000):
                await loc.click(timeout=2000)
                img_clicked = True
                await page.wait_for_timeout(1000)
                break
        except Exception:
            pass

    if not img_clicked:
        return f"«{recipe_name}»: картинка UUID={image_uuid} не найдена — выбери вручную"

    # Шаг 7: заполняем «Введите ссылку»
    await page.wait_for_timeout(500)

    for sel in [
        'input[placeholder*="ссылку"]',  'input[placeholder*="Ссылку"]',
        'input[placeholder*="Ссылка"]',  'input[placeholder*="ссылка"]',
        'input[placeholder*="URL"]',     'input[placeholder*="url"]',
        'input[placeholder*="http"]',    'input[type="url"]',
    ]:
        try:
            loc = page.locator(sel).last
            if await loc.count() > 0 and await loc.is_visible(timeout=600):
                await loc.click(timeout=1000)
                await loc.fill(recipe_url)
                await page.wait_for_timeout(150)
                break
        except Exception:
            pass

    return ""


# ── ГЛАВНАЯ ЛОГИКА ────────────────────────────────────────────────────────────

async def fill_collection(page: Page, collection: dict) -> None:
    """Заполняет одну подборку на открытой странице формы."""
    title   = collection.get("title", "")
    author  = collection.get("author", "")
    intro   = collection.get("intro", "")
    recipes = collection.get("recipes", [])

    print(f"\n{'='*60}")
    print(f"  Подборка: {title}")
    print(f"  Рецептов: {len(recipes)}")
    print(f"{'='*60}")

    # ── Заголовок подборки
    if title:
        for sel in [
            'input[placeholder*="Заголовок"]', 'input[placeholder*="заголовок"]',
            'input[placeholder*="Название"]',  'input[name="title"]',
        ]:
            try:
                loc = page.locator(sel).first
                if await loc.count() > 0 and await loc.is_visible(timeout=600):
                    await loc.click()
                    await loc.fill(title)
                    print(f"  ✓ Заголовок: {title[:60]}")
                    break
            except Exception:
                pass

    # ── Автор
    if author:
        for sel in [
            'input[placeholder*="Автор"]', 'input[placeholder*="автор"]',
            'input[name="author"]',
        ]:
            try:
                loc = page.locator(sel).first
                if await loc.count() > 0 and await loc.is_visible(timeout=600):
                    await loc.click()
                    await loc.fill(author)
                    print(f"  ✓ Автор: {author}")
                    break
            except Exception:
                pass

    # ── Вводка (intro) через W-кнопку секции «Вводка»
    if intro:
        intro_html = "\n".join(
            f"<p>{html_lib.escape(p)}</p>"
            for p in intro.split("\n") if p.strip()
        )
        ok = await fill_section_via_w(page, "Вводка", intro_html)
        if not ok:
            print(f"  ⚠ Вводка: вставь вручную: {intro[:80]}...")

    print()

    # ── Основной текст через W-кнопку секции «Основной текст»
    body_html = collection.get("body_html", "")
    if body_html:
        ok = await fill_section_via_w(page, "Основной текст", body_html)
        if not ok:
            print("  ⚠ Основной текст: вставь вручную, затем нажми Enter")
            input("  → ")
        await page.wait_for_timeout(800)

    # ── Блоки изображений для каждого рецепта
    print()
    warnings = []
    for i, recipe in enumerate(recipes):
        name = recipe.get("name", "")
        url  = recipe.get("url", "")
        uuid = recipe.get("image_uuid", "")

        print(f"  [{i+1}/{len(recipes)}] Картинка: {name[:50]}")
        warn = await add_image_block(page, uuid, name, url)
        if warn:
            warnings.append(warn)
            print(f"    ⚠ {warn}")
        else:
            print(f"    ✓ UUID={uuid[:8]}... | {name[:30]}")

        await page.wait_for_timeout(500)

    print()
    if warnings:
        print("  Предупреждения:")
        for w in warnings:
            print(f"  • {w}")
    else:
        print("  Всё заполнено автоматически.")

    print()
    print("  Проверь форму в браузере и нажми 'Опубликовать'.")
    input("  Нажми Enter когда готов перейти к следующей подборке...")


async def main():
    if not IN_FILE.exists():
        print(f"Файл {IN_FILE} не найден.")
        print("Сначала запусти: python docx_to_collection_json.py подборка.docx")
        sys.exit(1)

    with open(IN_FILE, encoding="utf-8") as f:
        collections = json.load(f)

    if not isinstance(collections, list):
        collections = [collections]

    print(f"Загружено подборок: {len(collections)}")

    async with async_playwright() as pw:
        try:
            browser = await pw.chromium.connect_over_cdp(CDP_URL)
        except Exception as e:
            print(f"\n✗ НЕ УДАЛОСЬ ПОДКЛЮЧИТЬСЯ К CHROME")
            print(f"  Ошибка: {e}")
            print("\n  Запусти Chrome через chrome-debug.bat (Windows)")
            print("  или chrome-debug.sh (macOS), затем открой форму подборки.")
            sys.exit(1)

        contexts = browser.contexts
        if not contexts:
            print("Нет открытых контекстов браузера.")
            sys.exit(1)

        # Ищем вкладку с формой подборки
        page = None
        for ctx in contexts:
            for p in ctx.pages:
                u = p.url
                if COLLECTION_URL_PATTERN in u and (
                    "collection" in u or "подборк" in u
                    or "edit" in u or "create" in u
                ):
                    page = p
                    break
            if page:
                break

        if not page:
            for ctx in contexts:
                for p in ctx.pages:
                    if COLLECTION_URL_PATTERN in p.url:
                        page = p
                        break
                if page:
                    break

        if not page:
            print(f"Не нашёл открытую вкладку с {COLLECTION_URL_PATTERN}.")
            print("Открой форму создания подборки в Chrome и запусти скрипт снова.")
            sys.exit(1)

        print(f"Подключился к: {page.url}\n")

        for col in collections:
            await fill_collection(page, col)

    print("\nВсе подборки обработаны.")


if __name__ == "__main__":
    asyncio.run(main())
