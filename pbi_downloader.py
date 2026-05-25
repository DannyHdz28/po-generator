import asyncio
import os
import tempfile
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

PBI_URL = "https://reports.maximaapparel.com/reports/powerbi/RESOURCES/Style%20UPC%20Report"
USERNAME = "daniela.hernandez@maximaapparel.com"
PASSWORD = "Aez1234!1"

SIZES = ["2T", "3T", "4", "4T", "5", "6", "7", "OS", "XS", "S", "M", "L", "XL", "2XL", "3XL"]


async def _login(page, username, password):
    """Handle Microsoft / Azure AD login flow."""
    try:
        await page.wait_for_selector('input[type="email"]', timeout=10000)
        await page.fill('input[type="email"]', username)
        await page.click('input[type="submit"]')
        await page.wait_for_selector('input[type="password"]', timeout=10000)
        await page.fill('input[type="password"]', password)
        await page.click('input[type="submit"]')
        # "Stay signed in?" dialog
        try:
            await page.wait_for_selector('input[type="submit"]', timeout=5000)
            await page.click('input[type="submit"]')
        except PWTimeout:
            pass
    except PWTimeout:
        pass  # Already logged in or different auth flow


async def _wait_for_report(page):
    """Wait until the Power BI report finishes loading."""
    await page.wait_for_load_state("networkidle", timeout=60000)
    # Wait for the PBI visual table to appear
    await page.wait_for_selector("text=STYLE UPC REPORT", timeout=60000)
    await asyncio.sleep(3)


async def _set_size_filter(page, size):
    """Open the SIZE filter and select only the given size."""
    # Open Filtros panel if closed
    try:
        filtros_btn = page.locator("text=Filtros").first
        if await filtros_btn.is_visible():
            await filtros_btn.click()
            await asyncio.sleep(1)
    except Exception:
        pass

    # Find the SIZE filter section and expand it
    size_filter = page.locator("text=SIZE").nth(1)  # Second "SIZE" avoids "SIZE CATEGORY"
    await size_filter.click()
    await asyncio.sleep(1)

    # Clear all current selections via "Seleccionar todo" toggle
    try:
        select_all = page.get_by_text("Seleccionar todo").first
        # Uncheck all first
        await select_all.click()
        await asyncio.sleep(0.5)
        await select_all.click()
        await asyncio.sleep(0.5)
        # Now uncheck all
        await select_all.click()
        await asyncio.sleep(0.5)
    except Exception:
        pass

    # Type the size in the search box to filter options
    try:
        search_box = page.locator('input[placeholder="Buscar"]').last
        await search_box.fill(size)
        await asyncio.sleep(1)
    except Exception:
        pass

    # Check the matching size
    size_checkbox = page.get_by_text(size, exact=True).first
    await size_checkbox.click()
    await asyncio.sleep(2)


async def _export_data(page, download_dir, size):
    """Trigger the export from Power BI and return the downloaded file path."""
    async with page.expect_download(timeout=120000) as dl_info:
        # Try Archivo menu first
        try:
            archivo_menu = page.get_by_text("Archivo").first
            await archivo_menu.click()
            await asyncio.sleep(1)
            export_option = page.get_by_text("Exportar").first
            await export_option.click()
            await asyncio.sleep(1)
            # Confirm export dialog if it appears
            try:
                confirm = page.get_by_text("Exportar", exact=True).nth(1)
                await confirm.click()
            except Exception:
                pass
        except Exception:
            # Fallback: try the download icon at top right of browser chrome
            await page.keyboard.press("Control+Shift+Alt+e")

    download = await dl_info.value
    file_path = os.path.join(download_dir, f"size_{size}.xlsx")
    await download.save_as(file_path)
    return file_path


async def download_all_sizes(progress_fn=None):
    """
    Main entry point. Opens Chrome, logs into PBI, applies filters and downloads
    one Excel file per size. Returns list of downloaded file paths.
    """
    download_dir = tempfile.mkdtemp(prefix="upcs_")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized"],
        )
        context = await browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = await context.new_page()

        # Navigate to report
        if progress_fn:
            progress_fn("Abriendo Power BI...")
        await page.goto(PBI_URL, timeout=60000)

        # Login
        if progress_fn:
            progress_fn("Iniciando sesión...")
        await _login(page, USERNAME, PASSWORD)
        await _wait_for_report(page)

        # Apply permanent filters once: ENTITY = PRO STANDARD US
        # (The ENTITY slicer in the report sidebar handles this)
        # UPC filter (no blanks) should already be set as the default in PBI
        if progress_fn:
            progress_fn("Configurando filtros base...")

        # Set ENTITY = PRO STANDARD US via the on-canvas slicer
        try:
            entity_dropdown = page.locator("text=Todas").first
            await entity_dropdown.click()
            await asyncio.sleep(1)
            pro_std = page.get_by_text("PRO STANDARD US", exact=True).first
            await pro_std.click()
            await asyncio.sleep(2)
        except Exception:
            pass

        # Download each size
        files = []
        for i, size in enumerate(SIZES):
            if progress_fn:
                progress_fn(f"Descargando talla {size} ({i+1}/{len(SIZES)})...")
            try:
                await _set_size_filter(page, size)
                file_path = await _export_data(page, download_dir, size)
                files.append(file_path)
            except Exception as e:
                if progress_fn:
                    progress_fn(f"⚠ Error en talla {size}: {e}")

        await browser.close()

    return files


def run_download(progress_fn=None):
    """Synchronous wrapper for use in Streamlit."""
    return asyncio.run(download_all_sizes(progress_fn))
