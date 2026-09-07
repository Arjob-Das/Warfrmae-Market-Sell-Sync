# Warframe Market Excel Automation

Automated, two-way synchronization between your [Warframe.market](https://warframe.market) sell orders and an interactive Microsoft Excel workbook.

---

## Key Features

1. **8-Column Inventory System with Live Formulas**:
   - **Column A**: `Item/Mod Name` (Left-aligned, text wrapped)
   - **Column B**: `Price` (Platinum Gold `#FBBF24`, Right-aligned)
   - **Column C**: `Stock` (Crisp White `#FFFFFF`, dynamic formula `={base_stock}-E{row}`)
   - **Column D**: `Visible` (Interactive toggle checkbox: `☑` Emerald `#34D399` / `☐` Slate `#64748B`)
   - **Column E**: `Quantity Sold: Current Session` (Ice Blue `#93C5FD`, editable sales tracker)
   - **Column F**: `Quantity Sold: All Time` (Cumulative quantity sold)
   - **Column G**: `Revenue Generated: Current Session` (Formula `=B{row}*E{row}`)
   - **Column H**: `Revenue Generated: All Time` (Formula `={base_all_time_rev}+G{row}`)
   - **Column I**: `Actions Gutter` (Cell `I2` quick action link)
   - **Column J**: `Sidebar Controls` (Account info, credentials, macro buttons, export directory)

2. **Accurate All-Time Revenue Preservation**:
   - All-time revenue accumulates dynamically: `Current All-Time Revenue + Current Session Revenue` (`={base_all_time_rev}+G{row}`).
   - Preserves historical sales revenue even when current listing prices fluctuate.

3. **Restock-Optimized Sorting**:
   - Always sorted by **Stock ASCENDING** (lowest stock first: 0, 1, 2...) so depleted items needing restock appear at the top.
   - When stock count is equal, items are sorted by **Price DESCENDING** (highest value items first).

4. **Stock Depletion & Delisting (1 -> 0)**:
   - Warframe.market's API rejects `quantity: 0` (`PATCH /v2/order/{id}` requires `quantity >= 1`).
   - When stock falls from 1 to 0 upon sales rollover, the **visibility checkbox is unchecked** (`☐` / `visible: false`), maintaining base stock at `1`.
   - When pushed to the market, the listing is hidden automatically without deleting or throwing an API error.

5. **Interactive Checkbox & VBA Auto-Save**:
   - Clicking any cell in Column D toggles `☑` <-> `☐` and updates font color immediately.
   - Every macro button forces active cell commit (`Range("A1").Select`) and saves the workbook (`ThisWorkbook.Save`) before running any action.

6. **Automatic Clean `.xlsx` Export**:
   - Automatically exports a clean copy of the workbook to your configured folder (`J16`, e.g. `C:\Users\titof\OneDrive\Documents\Warframe\Warframe Sell Stats.xlsx`).
   - Strips all macros and shape buttons, and sanitizes sensitive JWT credentials (`J4`, `I4`, `H4`).

---

## Project Structure

```
Warfrmae Market Sell Sync/
├── modules/
│   ├── __init__.py           # Package initializer
│   ├── config.py             # Constants, themes, fonts, fills, alignments, borders
│   ├── api.py                # Warframe.market API client, JWT validation, items cache
│   ├── sheet_layout.py       # Openpyxl layout, headers, zebra styling, sidebar (Col J)
│   ├── vba_manager.py        # VBA macros (SortTableByStock, Save, ExportNormalXLSX) & COM injection
│   ├── sync_engine.py        # Order syncing, column refresh, session commit, stock-ascending sort
│   ├── market_updater.py     # Excel price/stock/visibility reader, diff engine, patch pusher
│   └── workflows.py          # Orchestration workflows for CLI & interactive menu
├── warframe_market.py        # Unified CLI & VBA entry point
├── Warframe_Market_Macros.bas # Standalone VBA module
├── Warframe Sell Stats.xlsm  # Primary macro-enabled workbook
└── README.md                 # Documentation
```

---

## Setup & Installation

### Step 1: Install Dependencies

```powershell
pip install openpyxl requests pywin32
```

### Step 2: Initial Sync & Workbook Generation

```powershell
python warframe_market.py --sync
```

*(Optional: specify your username if different from `darksoulhunter2001`:)*

```powershell
python warframe_market.py --sync --user <YourUsername>
```

### Step 3: Add Your JWT Token

1. Open `Warframe Sell Stats.xlsm` in Excel (click **Enable Macros**).
2. Paste your Warframe.market JWT cookie into cell **`J4`**.
3. *(Optional)* Verify your export directory in cell **`J16`** (defaults to `~\OneDrive\Documents\Warframe`).

---

## How to Use Inside Excel

Open `Warframe Sell Stats.xlsm`. In **Column J**, click the buttons or cells:

- **▶ 1. Sync from Market** (`J7`): Fetches your active market orders into Excel.
- **⬆ 2. Push Prices & Stock** (`J9`): Pushes edited prices, stock, and checkbox visibility (`☑`/`☐`) to Warframe.market.
- **🔄 3. Update All Time Revenue** (`J11` or `I2`): Immediately commits Current Session revenue into All Time Revenue, resets Current Session sold to 0, handles 1->0 stock unchecks, and sorts by stock ascending / price descending.
- **⚡ 4. Refresh Columns & Formulas** (`J13`): Re-applies formulas, table styling, and sorting.

---

## Command Line Usage

| Command | Action |
| --- | --- |
| `python warframe_market.py` | Open interactive CLI menu |
| `python warframe_market.py --sync` | Fetch market orders into Excel |
| `python warframe_market.py --dry-run` | Preview Excel vs Market price/stock/visibility differences |
| `python warframe_market.py --push` | Push Excel price/stock/visibility changes to warframe.market |
| `python warframe_market.py --push -y` | Push changes without confirmation prompt |
| `python warframe_market.py --commit` | Update All Time Revenue (rollover session to all-time) |
| `python warframe_market.py --update-columns` | Refresh formulas, styling, and sort order |
| `python warframe_market.py --export-xlsx` | Export clean sanitized `.xlsx` copy without macros |

---

## How to Get Your JWT Token

1. Log into [warframe.market](https://warframe.market) in Chrome or Edge.
2. Press `F12` -> `Application` (or `Storage`) -> `Cookies` -> `https://warframe.market`.
3. Copy the value of the `JWT` cookie (starts with `eyJ...`).
4. Paste it directly into cell **`J4`** of `Warframe Sell Stats.xlsm` (or enter when prompted by the script).
