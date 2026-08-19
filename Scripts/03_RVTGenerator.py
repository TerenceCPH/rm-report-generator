"""03 RVT Generator: download DOORS module views, format Excel, export PDF."""

import argparse
import importlib
import math
import os
import sys
import time

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import WorksheetProperties
from playwright.sync_api import sync_playwright

sys.modules.setdefault("env", importlib.import_module("00_EnvironmentVariables"))
from env import (
    INPUT_DIR,
    PASSWORD,
    ROOT_DIR,
    USERNAME,
    export_workbook_to_pdf,
    output_subdir,
)

DEFAULT_CSV_CANDIDATES = [
    os.path.join(INPUT_DIR, "AppD", "RM Module List.csv"),
    os.path.join(ROOT_DIR, "AppD", "RM Module List.csv"),
    os.path.join(ROOT_DIR, "RM Module List.csv"),
    r"C:\Users\terence.phchan\Downloads\RM\AppD\RM Module List.csv",
]

LOGIN_URL = (
    "https://1263doorsapp.shuion.com.hk:9443/rm/web#action="
    "com.ibm.rdm.web.pages.showProjectDashboard&componentURI="
    "https%3A%2F%2F1263doorsapp.shuion.com.hk%3A9443%2Frm%2Frm-projects"
    "%2F_wt4CcCcoEfCz_bTWSv7vxQ%2Fcomponents%2F_wvhBRScoEfCz_bTWSv7vxQ"
    "&vvc.configuration=https%3A%2F%2F1263doorsapp.shuion.com.hk%3A9443"
    "%2Frm%2Fcm%2Fstream%2F_wvqLJicoEfCz_bTWSv7vxQ&viewType=all"
)


def resolve_csv_path(explicit):
    if explicit:
        return explicit
    for path in DEFAULT_CSV_CANDIDATES:
        if os.path.exists(path):
            return path
    return DEFAULT_CSV_CANDIDATES[0]


def format_excel(file_path: str, doc_name: str):
    """
    Applies formatting to downloaded XLSX file:
    - Enables wrap_text on ALL columns.
    - Explicitly sets row height based on MAX estimated lines across Column D, Column H, and Column L.
    - Excludes Column I from driving height expansion.
    - Deletes METADATA rows before applying borders.
    """
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    if ws.max_column >= 13:
        ws.delete_cols(13, 8)

    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val and "METADATA" in str(val).upper():
            start_row = max(1, r - 1)
            ws.delete_rows(start_row, ws.max_row - start_row + 1)
            break

    max_row = ws.max_row
    max_col = ws.max_column

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    if ws.sheet_properties is None:
        ws.sheet_properties = WorksheetProperties()

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"

    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75
    ws.page_margins.left = 0.7
    ws.page_margins.right = 0.7
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3

    ws.oddHeader.left.text = '&"Arial,Regular"&10 Appendix D — RVT Attachment'
    ws.oddHeader.right.text = f'&"Arial,Regular"&10 {doc_name}'

    col_widths = {2: 16, 4: 60, 5: 12, 7: 12, 8: 12, 9: 32, 11: 12, 12: 32}
    for col_idx, width in col_widths.items():
        if max_col >= col_idx:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    font_regular = Font(name="Arial", size=10)
    font_bold = Font(name="Arial", size=10, bold=True)

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    target_last_col = min(max_col, 15)

    for row in range(1, max_row + 1):
        is_header = row == 1

        if is_header:
            ws.row_dimensions[row].height = 25
        else:
            col4_val = str(ws.cell(row=row, column=4).value or "")
            lines_d = 0
            for paragraph in col4_val.split("\n"):
                lines_d += math.ceil(len(paragraph) / 55) if paragraph else 1

            col8_val = str(ws.cell(row=row, column=8).value or "")
            lines_h = 0
            for paragraph in col8_val.split("\n"):
                lines_h += math.ceil(len(paragraph) / 10) if paragraph else 1

            col12_val = str(ws.cell(row=row, column=12).value or "")
            lines_l = 0
            for paragraph in col12_val.split("\n"):
                lines_l += math.ceil(len(paragraph) / 28) if paragraph else 1

            max_lines = max(lines_d, lines_h, lines_l)
            calculated_height = max(20, max_lines * 15)
            ws.row_dimensions[row].height = calculated_height

        for col in range(1, target_last_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.font = font_bold if is_header else font_regular

            if is_header:
                cell.alignment = center_align
            else:
                if col in (4, 9, 12):
                    cell.alignment = left_align
                else:
                    cell.alignment = center_align

    wb.save(file_path)


def maybe_pdf(xlsx_path, skip_pdf):
    if skip_pdf:
        return
    try:
        export_workbook_to_pdf(xlsx_path)
    except Exception as e:
        print(f"    [!] Error exporting PDF: {e}")


def automate_doors_export(csv_path, output_dir, skip_pdf=False):
    os.makedirs(output_dir, exist_ok=True)

    try:
        df = pd.read_csv(csv_path)
    except FileNotFoundError:
        print(f"Error: File not found at {csv_path}")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="msedge")
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()

        page.goto(LOGIN_URL)
        page.fill("#jazz_app_internal_LoginWidget_0_userId", USERNAME)
        page.fill("#jazz_app_internal_LoginWidget_0_password", PASSWORD)
        page.check('input[type="checkbox"]')
        page.click(".j-button-primary")
        page.wait_for_load_state("networkidle")

        artifacts_btn = page.locator('a[title="Artifacts"]').first
        if not artifacts_btn.evaluate('el => el.classList.contains("selected")'):
            artifacts_btn.click()
            page.wait_for_load_state("networkidle")

        artifacts_url = page.url

        for index, row in df.iterrows():
            module_id = str(row["Module ID"])
            doc_number = str(row.get("Document Number", "Doc"))
            doc_name = str(row.get("Document Name", doc_number))
            file_name = f"{doc_number}_{module_id}.xlsx"
            save_path = os.path.join(output_dir, file_name)

            if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                print(
                    f"[⇄] Module {index + 1}/{len(df)}: {module_id} already exists "
                    f"({file_name}). Formatting & skipping download."
                )
                try:
                    format_excel(save_path, doc_name)
                    maybe_pdf(save_path, skip_pdf)
                except Exception as fe:
                    print(f"    [!] Error applying Excel formatting: {fe}")
                continue

            print(f"\n[+] Processing Module {index + 1}/{len(df)}: {module_id}")

            try:
                search_input = page.locator(
                    'input.filterText[placeholder*="Type to filter"]'
                ).first
                search_input.wait_for(state="visible", timeout=15000)
                search_input.clear()
                search_input.fill(module_id)
                search_input.press("Enter")

                page.wait_for_load_state("networkidle")
                time.sleep(2)

                resource_link = page.locator(
                    f'a.jazz-ui-ResourceLink:has-text("{module_id}")'
                ).first
                if resource_link.count() == 0 or not resource_link.is_visible():
                    print(f"[-] Module {module_id} not found. Moving to next.")
                    continue

                resource_link.click()
                page.wait_for_load_state("networkidle")

                view_span = page.locator(
                    'span[title="App E View (T): Shared, All modules"]'
                ).first
                view_span.wait_for(state="visible", timeout=15000)
                view_span.click(button="right")

                export_option = page.get_by_text("Export View...").last
                export_option.wait_for(state="visible", timeout=10000)
                export_option.click(force=True)

                xlsx_label = page.get_by_text("XLSX", exact=True).last
                xlsx_label.wait_for(state="visible", timeout=15000)
                xlsx_label.click(force=True)

                ok_btn = page.locator('button.j-button-primary:has-text("OK")').last
                ok_btn.wait_for(state="visible", timeout=10000)

                with page.expect_download(timeout=60000) as download_info:
                    ok_btn.click()

                download = download_info.value
                download.save_as(save_path)

                while not os.path.exists(save_path) or os.path.getsize(save_path) == 0:
                    time.sleep(0.5)

                print(f"[✓] Downloaded: {file_name}")

                format_excel(save_path, doc_name)
                print(f"[✓] Formatted: {file_name}")
                maybe_pdf(save_path, skip_pdf)

            except Exception as e:
                print(f"[!] Error processing Module {module_id}: {e}")
                print("    Recovering state and continuing to next module...")

            try:
                page.goto(artifacts_url)
                page.wait_for_load_state("networkidle")
                page.locator(
                    'input.filterText[placeholder*="Type to filter"]'
                ).first.wait_for(state="visible", timeout=15000)
                time.sleep(1)
            except Exception as nav_e:
                print(f"    [!] Failed to reset page navigation: {nav_e}")

        browser.close()
        print("\n[✓] All module downloads and formattings completed!")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Download RVT module views, format Excel, and export PDF."
    )
    parser.add_argument(
        "--csv",
        default=None,
        help="Path to RM Module List.csv (Module ID, Document Number, Document Name).",
    )
    parser.add_argument(
        "--output-dir",
        default=output_subdir("03_RVTGenerator"),
        help="Folder for per-module xlsx/pdf files.",
    )
    parser.add_argument("--skip-pdf", action="store_true", help="Do not export PDF via Excel.")
    return parser.parse_args()


def main():
    args = parse_args()
    csv_path = resolve_csv_path(args.csv)
    automate_doors_export(csv_path, args.output_dir, skip_pdf=args.skip_pdf)


if __name__ == "__main__":
    main()
