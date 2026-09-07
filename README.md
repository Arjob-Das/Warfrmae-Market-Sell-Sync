# Warframe Market Excel Automation

Automated two-way synchronization between your [Warframe.market](https://warframe.market) sell orders and an Excel spreadsheet.

## Files

- **`warframe_market.py`**: Main Python automation script.
- **`Warframe Sell Stats.xlsm`**: Macro-enabled Excel workbook with live formulas, account credentials in Column H, and 1-click action buttons.

---

## Fresh Start (From Scratch / No Files Present)

If starting completely fresh with only `warframe_market.py`:

### Step 1: Install Dependencies

```powershell
pip install openpyxl requests pywin32
```

### Step 2: Generate Workbook & Initial Sync

```powershell
python warframe_market.py --sync
```

*Tip: If your Warframe.market username is different from `darksoulhunter2001`, pass it once:*

```powershell
python warframe_market.py --sync --user <YourUsername>
```

This single command:

1. Downloads the items catalog and fetches all your active sell orders.
2. Creates `Warframe Sell Stats.xlsm` from scratch with 2-row headers, table styling, and dynamic revenue formulas (`=B*C` and `=B*D`).
3. Appends the Totals summary row (`=SUM(E3:E...)` and `=SUM(F3:F...)`).
4. Automatically injects the VBA macros and one-click action buttons into the workbook.

### Step 3: Add Your JWT Token

1. Open `Warframe Sell Stats.xlsm` in Excel (click **Enable Macros** if prompted).
2. Paste your Warframe.market JWT cookie into cell **`H4`** *(see below on how to grab it)*.
3. Done! All buttons and cell shortcuts inside Excel are now fully operational.

---

## How to Use Inside Excel

Open `Warframe Sell Stats.xlsm`. In **Column H**, click the cells or buttons:

- **▶ Sync from Market** (`H7`): Fetches newest orders & prices from warframe.market. Terminal auto-closes.
- **⬆ Push Prices to Market** (`H9`): Pushes your edited Excel prices to warframe.market. Terminal auto-closes.
- **🔄 Update All Time Revenue** (`H11` or `G2`): Immediately rolls over Current Session quantities into All Time and resets Current Session to 0 (no confirmation dialog).
- **⚡ Refresh Formulas** (`H13`): Re-applies formulas and table styling. Terminal auto-closes.

> **Credentials**: Username is in cell **`H2`** and JWT Token is in cell **`H4`**.

---

## Command Line Usage

| Command                                        | Action                                                |
| ---------------------------------------------- | ----------------------------------------------------- |
| `python warframe_market.py`                  | Open interactive CLI menu                             |
| `python warframe_market.py --sync`           | Fetch market orders into Excel                        |
| `python warframe_market.py --dry-run`        | Preview Excel vs Market price differences             |
| `python warframe_market.py --push`           | Push Excel price changes to warframe.market           |
| `python warframe_market.py --push -y`        | Push price changes without confirmation prompt        |
| `python warframe_market.py --commit`         | Update All Time Revenue (archive session to All Time) |
| `python warframe_market.py --update-columns` | Refresh formulas and styling                          |

---

## How to Get Your JWT Token

1. Log into [warframe.market](https://warframe.market) in Chrome or Edge.
2. Press `F12` -> `Application` (or `Storage`) -> `Cookies` -> `https://warframe.market`.
3. Copy the value of the `JWT` cookie (starts with `eyJ...`).
4. Paste it directly into cell **`H4`** of `Warframe Sell Stats.xlsm` (or enter when prompted by the script).
