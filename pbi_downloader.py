import asyncio
import os
import tempfile
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

PBI_URL = "https://reports.maximaapparel.com/reports/powerbi/RESOURCES/Style%20UPC%20Report"
USERNAME = "daniela.hernandez@maximaapparel.com"
PASSWORD = "Aez1234!1"
WIN_USERNAME = "daniela.hernandez"


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
            progress_fn("Iniciando sesion...")
        await _login(page, USERNAME, PASSWORD)
        await page.wait_for_load_state("networkidle", timeout=120000)
        await asyncio.sleep(10)

        if progress_fn:
            progress_fn("Configurando filtro ENTITY = PRO STANDARD US...")
        try:
            entity_dropdown = page.locator("text=Todas").first
            await entity_dropdown.click()
            await asyncio.sleep(2)
            pro_std = page.get_by_text("PRO STANDARD US", exact=True).first
            await pro_std.click()
            await asyncio.sleep(5)
        except Exception as e:
            if progress_fn:
                progress_fn(f"Filtro ENTITY: {e}")

        if progress_fn:
            progress_fn("Exportando datos... (puede tardar 2-3 minutos)")

        files = []
        try:
            async with page.expect_download(timeout=300000) as dl_info:
                # Try toolbar export button first
                try:
                    export_btn = page.locator('[title="Exportar"]').first
                    await export_btn.click(timeout=5000)
                except Exception:
                    pass

                await asyncio.sleep(2)

                # Try Archivo > Exportar menu
                try:
                    await page.get_by_text("Archivo").first.click()
                    await asyncio.sleep(1)
                    await page.get_by_text("Exportar").first.click()
                    await asyncio.sleep(1)
                except Exception:
                    pass

                # Confirm dialog if appears
                try:
                    await page.get_by_role("button", name="Exportar").last.click()
                    await asyncio.sleep(1)
                except Exception:
                    pass

            download = await dl_info.value
            file_path = os.path.join(download_dir, "upcs_data.xlsx")
            await download.save_as(file_path)
            files.append(file_path)
            if progress_fn:
                progress_fn("Descarga completada.")
        except Exception as e:
            if progress_fn:
                progress_fn(f"Error exportando: {e}")

        await browser.close()

    return files


def run_download(progress_fn=None):
    return asyncio.run(download_all_sizes(progress_fn))
