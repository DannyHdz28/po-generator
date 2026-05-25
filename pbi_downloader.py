import asyncio
import os
import tempfile
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

PBI_URL = "https://reports.maximaapparel.com/reports/powerbi/RESOURCES/Style%20UPC%20Report"
USERNAME = "daniela.hernandez@maximaapparel.com"
PASSWORD = "Aez1234!1"
# Windows/NTLM username (without domain)
WIN_USERNAME = "daniela.hernandez"

SIZES = ["2T", "3T", "4", "4T", "5", "6", "7", "OS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]


async def _login(page, username, password):
    try:
        await page.wait_for_selector('input[type="email"]', timeout=15000)
        await page.fill('input[type="email"]', username)
        await page.click('input[type="submit"]')
        await page.wait_for_selector('input[type="password"]', timeout=15000)
        await page.fill('input[type="password"]', password)
        await page.click('input[type="submit"]')
        try:
            await page.wait_for_selector('input[type="submit"]', timeout=8000)
            await page.click('input[type="submit"]')
        except PWTimeout:
            pass
    except PWTimeout:
        pass


async def _set_size_filter(page, size, progress_fn=None):
    try:
        filter_btn = page.locator('[aria-label="Filtros"]').first
        if await filter_btn.is_visible():
            await filter_btn.click()
            await asyncio.sleep(1)
    except Exception:
        pass

    try:
        size_card = page.locator("text=SIZE").nth(1)
        await size_card.click()
        await asyncio.sleep(1)
    except Exception:
        pass

    try:
        select_all = page.get_by_text("Seleccionar todo").first
        is_checked = await select_all.get_attribute("aria-checked")
        if is_checked != "false":
            await select_all.click()
            await asyncio.sleep(0.5)
    except Exception:
        pass

    try:
        search = page.locator('input[placeholder="Buscar"]').last
        await search.fill("")
        await search.fill(size)
        await asyncio.sleep(1)
    except Exception:
        pass

    try:
        checkbox = page.get_by_text(size, exact=True).first
        await checkbox.click()
        await asyncio.sleep(2)
    except Exception as e:
        if progress_fn:
            progress_fn(f"⚠ No se pudo seleccionar talla {size}: {e}")


async def _export_file(page, download_dir, filename, progress_fn=None):
    try:
        async with page.expect_download(timeout=180000) as dl_info:
            export_btn = page.locator('[title="Exportar"]').first
            await export_btn.click()
            await asyncio.sleep(2)
            try:
                await page.get_by_role("button", name="Exportar").last.click()
            except Exception:
                pass
        download = await dl_info.value
        file_path = os.path.join(download_dir, filename)
        await download.save_as(file_path)
        return file_path
    except Exception as e:
        if progress_fn:
            progress_fn(f"⚠ Error exportando {filename}: {e}")
        return None


async def download_all_sizes(progress_fn=None):
    download_dir = tempfile.mkdtemp(prefix="upcs_")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, args=["--start-maximized"])
        context = await browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
            http_credentials={"username": WIN_USERNAME, "password": PASSWORD},
        )
        page = await context.new_page()

        if progress_fn:
            progress_fn("Abriendo Power BI...")
        await page.goto(PBI_URL, timeout=120000)

        if progress_fn:
            progress_fn("Iniciando sesión...")
        await _login(page, USERNAME, PASSWORD)
        await page.wait_for_load_state("networkidle", timeout=120000)
        await asyncio.sleep(8)

        if progress_fn:
            progress_fn("Configurando filtro ENTITY...")
        try:
            entity_dropdown = page.locator("text=Todas").first
            await entity_dropdown.click()
            await asyncio.sleep(2)
            pro_std = page.get_by_text("PRO STANDARD US", exact=True).first
            await pro_std.click()
            await asyncio.sleep(4)
        except Exception as e:
            if progress_fn:
                progress_fn(f"Filtro ENTITY: {e}")

        files = []
        for i, size in enumerate(SIZES):
            if progress_fn:
                progress_fn(f"Descargando talla {size} ({i+1}/{len(SIZES)})...")
            await _set_size_filter(page, size, progress_fn)
            file_path = await _export_file(page, download_dir, f"size_{size}.xlsx", progress_fn)
            if file_path:
                files.append(file_path)

        await browser.close()

    return files


def run_download(progress_fn=None):
    return asyncio.run(download_all_sizes(progress_fn))
