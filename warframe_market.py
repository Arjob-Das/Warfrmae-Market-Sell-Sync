"""
Warframe Market Excel Automation Engine
=======================================
Modular entry point and CLI for Warframe.market synchronization, price updating,
session rollover, and Excel workbook automation.

Usage:
  python warframe_market.py                  # Interactive menu
  python warframe_market.py --sync           # Fetch sell orders and sync into Excel
  python warframe_market.py --push           # Read Excel prices/stock and push changes to warframe.market
  python warframe_market.py --dry-run        # Preview price/stock/visibility changes without pushing
  python warframe_market.py --commit         # Rollover session quantities & revenue to all-time
  python warframe_market.py --update-columns # Refresh formatting, formulas, and VBA macro buttons
  python warframe_market.py --export-xlsx    # Export clean .xlsx copy without macros
"""

import sys
import time
import argparse

# Re-export all core components for backwards compatibility
from modules.config import (
    CACHE_FILE, DEFAULT_EXCEL, SHEET_NAME, API_BASE_URL,
    API_CALL_DELAY_SEC, DEFAULT_EXPORT_DIR_PLACEHOLDER,
    FONT_NAME, FONT_HEADER, FONT_SUBHEADER, FONT_DATA,
    FONT_ITEM_NAME, FONT_PRICE, FONT_STOCK, FONT_CHECKED,
    FONT_UNCHECKED, FONT_QTY, FONT_REV, FONT_TOTAL,
    FONT_TOTAL_LABEL, FONT_MONO, FILL_CANVAS, FILL_ROW_ODD,
    FILL_ROW_EVEN, FILL_ZEBRA_EVEN, FILL_HEADER_DARK,
    FILL_HEADER_BLUE, FILL_HEADER_GREEN, FILL_SUBHEADER,
    FILL_TOTAL, FILL_TOKEN, FILL_CARD_DARK, ALIGN_LEFT,
    ALIGN_CENTER, ALIGN_RIGHT, BORDER_CELL, BORDER_TOTAL,
    BORDER_TOKEN_BOX, COLUMN_WIDTHS
)

from modules.sheet_layout import (
    get_open_excel_workbook, get_credentials_from_excel,
    save_credentials_to_excel, prompt_jwt_token,
    get_export_dir_from_excel, save_export_dir_to_excel,
    safe_merge, migrate_sheet_layout_if_needed,
    initialize_sheet_structure, apply_row_formulas_and_styling,
    populate_sidebar_column, adjust_column_widths
)

from modules.api import (
    get_auth_headers_and_cookies, is_jwt_expired,
    validate_jwt_token, load_items_catalog, fetch_orders
)

from modules.vba_manager import (
    VBA_MODULE_CODE, SHEET_EVENT_CODE,
    inject_vba_and_shapes, export_clean_xlsx
)

from modules.sync_engine import (
    sync_market_orders_to_excel, update_columns_and_formulas,
    commit_session_to_all_time
)

from modules.market_updater import (
    read_excel_prices_and_stock, read_excel_prices,
    diff_prices_and_stock, diff_prices,
    display_price_diff, push_price_updates
)

from modules.workflows import (
    run_sync_workflow, run_push_workflow,
    run_commit_workflow, interactive_menu
)


def populate_column_h(ws, username: str = "", jwt_token: str = "", export_dir: str = "") -> None:
    """Backward compatibility alias for populate_sidebar_column."""
    populate_sidebar_column(ws, username=username, jwt_token=jwt_token, export_dir=export_dir)


def main():
    parser = argparse.ArgumentParser(description="Warframe Market Excel Automation Engine")
    parser.add_argument("--sync", action="store_true", help="Fetch orders from warframe.market and sync to Excel")
    parser.add_argument("--push", action="store_true", help="Read Excel prices and push changes to warframe.market")
    parser.add_argument("--dry-run", action="store_true", help="Preview price changes without pushing")
    parser.add_argument("-y", "--yes", action="store_true", help="Skip confirmation prompt on price updates")
    parser.add_argument("--commit", action="store_true", help="Rollover Current Session quantities & revenue to All Time")
    parser.add_argument("--update-columns", action="store_true", help="Refresh formatting, formulas, and VBA macro buttons")
    parser.add_argument("--export-xlsx", action="store_true", help="Export clean .xlsx copy to configured directory")
    parser.add_argument("--user", type=str, help="Override Warframe.market Username")
    parser.add_argument("--token", type=str, help="Override / update JWT Token")
    parser.add_argument("--file", type=str, default=DEFAULT_EXCEL, help="Specify Excel file path")

    args = parser.parse_args()
    excel_file = args.file

    # Load credentials directly from Excel cells J2 & J4 (fallback I2/I4, H2/H4)
    excel_user, excel_jwt = get_credentials_from_excel(excel_file)
    username = (args.user or excel_user or "darksoulhunter2001").strip()
    jwt_token = (args.token or excel_jwt or "").strip()

    if args.user or args.token:
        save_credentials_to_excel(excel_file, username=args.user, jwt_token=args.token)

    # Command Line Flags Execution
    if args.sync:
        run_sync_workflow(excel_file, username, jwt_token)
        return
    elif args.push:
        run_push_workflow(excel_file, username, jwt_token, dry_run=False, auto_confirm=args.yes)
        return
    elif args.dry_run:
        run_push_workflow(excel_file, username, jwt_token, dry_run=True, auto_confirm=True)
        return
    elif args.commit:
        run_commit_workflow(excel_file)
        return
    elif args.update_columns:
        cnt = update_columns_and_formulas(excel_file)
        print(f"[+] Refreshed {cnt} row(s) with formulas and formatting.")
        time.sleep(1.2)
        return
    elif args.export_xlsx:
        res = export_clean_xlsx(excel_file)
        if res:
            print(f"[+] Clean .xlsx exported: {res}")
        else:
            print(f"[!] Export failed or no valid export directory configured.")
        time.sleep(1.2)
        return

    # Interactive Menu Mode
    interactive_menu(excel_file, username, jwt_token)


if __name__ == "__main__":
    main()
