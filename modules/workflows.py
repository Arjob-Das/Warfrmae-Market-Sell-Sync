"""
Warframe Market Sell Sync - High-Level Workflows & Interactive CLI
==================================================================
Orchestrates sync, push, commit, and refresh workflows for both CLI
and interactive menu sessions.
"""

import time
from .config import DEFAULT_EXCEL
from .sheet_layout import (
    get_credentials_from_excel, save_credentials_to_excel,
    prompt_jwt_token
)
from .api import validate_jwt_token, fetch_orders
from .sync_engine import (
    sync_market_orders_to_excel, update_columns_and_formulas,
    commit_session_to_all_time
)
from .market_updater import (
    read_excel_prices_and_stock, diff_prices_and_stock,
    display_price_diff, push_price_updates
)
from .vba_manager import export_clean_xlsx


def run_sync_workflow(excel_file: str, username: str, jwt_token: str) -> None:
    """Sync orders from warframe.market into Excel."""
    print(f"\n[*] Connecting to Warframe.market...")
    orders = fetch_orders(username, jwt_token, excel_file)
    if not orders:
        print(f"[!] No active sell orders found for '{username}'.")
        return

    print(f"[+] Retrieved {len(orders)} active sell order(s).")
    print(f"[*] Updating spreadsheet '{excel_file}'...")
    updated_cnt, added_cnt = sync_market_orders_to_excel(excel_file, orders, username=username, jwt_token=jwt_token)
    print(f"[+] Success! Added {added_cnt} new item(s), refreshed {updated_cnt} price(s).")
    print(f"[+] Stock formulas (=base-E), Revenue formulas (=B*E and =base+G), and Totals row updated.")
    time.sleep(1.5)


def run_push_workflow(excel_file: str, username: str, jwt_token: str, dry_run: bool = False, auto_confirm: bool = False) -> None:
    """Read Excel prices, stock, and visibility; compare with market, and push updates."""
    if not jwt_token or not validate_jwt_token(jwt_token):
        print("\n[!] Current JWT token is missing, expired, or invalid.")
        jwt_token = prompt_jwt_token(excel_file)
        if not jwt_token or not validate_jwt_token(jwt_token):
            print("[!] Operation aborted. No valid JWT token provided.")
            return

    excel_items = read_excel_prices_and_stock(excel_file)
    if not excel_items:
        print(f"[!] No items found in '{excel_file}'. Run --sync first.")
        return

    print(f"[*] Reading current orders from warframe.market...")
    market_orders = fetch_orders(username, jwt_token, excel_file)
    if not market_orders:
        print("[!] Could not fetch market orders.")
        return

    changed, unchanged, unmatched = diff_prices_and_stock(excel_items, market_orders)
    display_price_diff(changed, unchanged, unmatched)

    if not changed:
        time.sleep(1.5)
        return

    if dry_run:
        print("\n[*] Dry-run mode enabled. No changes pushed.")
        return

    if not auto_confirm:
        ans = input(f"\nPush these {len(changed)} change(s) to warframe.market? (y/N): ").strip().lower()
        if ans not in ("y", "yes"):
            print("[-] Cancelled by user.")
            return

    print(f"\n[*] Pushing {len(changed)} price, stock & visibility update(s) to warframe.market...")
    success, failed = push_price_updates(jwt_token, changed, excel_file)
    print("\n" + "=" * 50)
    print(f"  MARKET UPDATE RESULT: {success} Succeeded, {failed} Failed")
    print("=" * 50)
    export_clean_xlsx(excel_file)
    time.sleep(1.5)


def run_commit_workflow(excel_file: str) -> None:
    """Update All Time Revenue: add Current Session revenue to All Time Revenue and reset Current Session to 0."""
    print(f"[*] Updating All Time Revenue in '{excel_file}'...")
    committed = commit_session_to_all_time(excel_file)
    if not committed:
        print("[-] Current Session had no quantities > 0. All Time Revenue unchanged.")
    else:
        print("\n" + "=" * 55)
        print("  ALL TIME REVENUE UPDATE SUMMARY")
        print("=" * 55)
        tot_plat = 0
        for name, qty, rev in committed:
            print(f"  - {name:<32} x{qty:<2} (+{rev} plat)")
            tot_plat += rev
        print("-" * 55)
        print(f"  Total Session Revenue: {tot_plat} Platinum")
        print(f"[+] Current Session quantities reset to 0.")
        print(f"[+] All Time Revenue updated with accumulated session revenue.")
        print("=" * 55)


def interactive_menu(excel_file: str, username: str, jwt_token: str) -> None:
    """Runs interactive console terminal menu for all operations."""
    while True:
        excel_user, excel_jwt = get_credentials_from_excel(excel_file)
        u_display = excel_user or username or "[Not Set]"
        tok_display = "[Set in J4]" if (excel_jwt or jwt_token) else "[Not Set]"

        print("\n" + "=" * 60)
        print("       WARFRAME MARKET EXCEL AUTOMATION")
        print("=" * 60)
        print(f" Target File: {excel_file}")
        print(f" Account    : {u_display} | JWT: {tok_display}")
        print("-" * 60)
        print(" 1. Sync Market Sell Orders to Excel (Add missing items)")
        print(" 2. Push Excel Price & Stock Changes to Warframe.market")
        print(" 3. Preview Price Changes (Dry Run)")
        print(" 4. Update All Time Revenue (Roll Session Rev into All Time)")
        print(" 5. Refresh Columns, Formulas & Macro Buttons")
        print(" 6. Set / Update Browser JWT Token")
        print(" 7. Exit")
        print("=" * 60)

        choice = input("Select an option (1-7): ").strip()
        if choice == "1":
            run_sync_workflow(excel_file, username, jwt_token)
        elif choice == "2":
            run_push_workflow(excel_file, username, jwt_token)
        elif choice == "3":
            run_push_workflow(excel_file, username, jwt_token, dry_run=True)
        elif choice == "4":
            run_commit_workflow(excel_file)
        elif choice == "5":
            cnt = update_columns_and_formulas(excel_file)
            print(f"[+] Refreshed {cnt} row(s) with formulas and formatting.")
        elif choice == "6":
            new_tok = prompt_jwt_token(excel_file)
            if new_tok:
                jwt_token = new_tok
        elif choice == "7":
            print("Exiting. Happy trading, Tenno!")
            break
        else:
            print("[!] Invalid choice. Please enter 1-7.")
