# Warframe Market Excel Automation

Automated, two-way synchronization between your [Warframe.market](https://warframe.market) sell orders and an interactive Microsoft Excel workbook.

---

## 🌟 Key Features

1. **Live Warframe.market Status Dropdown (`J6`)**:
   - Change your market presence directly inside Excel: **Online**, **Online in Game**, or **Invisible**.
   - Selecting any option immediately auto-syncs your presence to Warframe.market via WebSocket in the background.

2. **8-Column Inventory System with Live Formulas**:
   - **Column A**: `Item/Mod Name` (Left-aligned, text wrapped)
   - **Column B**: `Price` (Platinum Gold `#FBBF24`, Right-aligned)
   - **Column C**: `Stock` (Crisp White `#FFFFFF`, dynamic formula `={base_stock}-E{row}`)
   - **Column D**: `Visible` (Interactive toggle checkbox: `☑` Emerald `#34D399` / `☐` Slate `#64748B`)
   - **Column E**: `Quantity Sold: Current Session` (Ice Blue `#93C5FD`, editable sales tracker)
   - **Column F**: `Quantity Sold: All Time` (Cumulative quantity sold)
   - **Column G**: `Revenue Generated: Current Session` (Formula `=B{row}*E{row}`)
   - **Column H**: `Revenue Generated: All Time` (Formula `={base_all_time_rev}+G{row}`)
   - **Column I**: `Actions Gutter` (Cell `I2` quick action link)
   - **Column J**: `Sidebar Controls` (Account username in `J2`, JWT token in `J4`, live presence in `J6`, and 4 macro action buttons)

3. **Authoritative In-Sheet Credential Storage**:
   - Your Warframe.market Username is stored in cell **`J2`**.
   - Your JWT Token is stored in cell **`J4`**.
   - Once saved in Excel, cells `J2` and `J4` serve as the permanent single source of truth. Any edits you make in Excel are automatically picked up by Python and VBA macros.

4. **Accurate All-Time Revenue Preservation**:
   - All-time revenue accumulates dynamically: `Current All-Time Revenue + Current Session Revenue` (`={base_all_time_rev}+G{row}`).
   - Preserves historical sales revenue even when current listing prices fluctuate.

5. **Restock-Optimized Sorting**:
   - Always sorted by **Stock ASCENDING** (lowest stock first: 0, 1, 2...) so depleted items needing restock appear at the top.
   - When stock count is equal, items are sorted by **Price DESCENDING** (highest value items first).

6. **Stock Depletion & Auto-Delisting (1 -> 0)**:
   - Warframe.market's API rejects `quantity: 0` (`PATCH /v2/order/{id}` requires `quantity >= 1`).
   - When stock falls from 1 to 0 upon sales rollover, the **visibility checkbox is unchecked** (`☐` / `visible: false`), maintaining base stock at `1`.
   - When pushed to the market, the listing is hidden automatically without deleting or throwing an API error.

7. **Interactive Checkbox & VBA Auto-Save**:
   - Clicking any cell in Column D toggles `☑` <-> `☐` and updates font color immediately.
   - Every macro button forces active cell commit (`Range("A1").Select`) and saves the workbook (`ThisWorkbook.Save`) before running any action.

---

## 📁 Project Structure

```
Warframe Market Sell Sync/
├── modules/
│   ├── __init__.py           # Package initializer
│   ├── config.py             # Constants, themes, fonts, fills, alignments, borders
│   ├── api.py                # Warframe.market REST & WebSocket client, JWT validation, items cache
│   ├── sheet_layout.py       # Openpyxl layout, headers, zebra styling, sidebar (Col J)
│   ├── vba_manager.py        # VBA macros (SortTableByStock, SetUserStatus) & COM injection
│   ├── sync_engine.py        # Order syncing, column refresh, session commit, stock-ascending sort
│   ├── market_updater.py     # Excel price/stock/visibility reader, diff engine, patch pusher
│   └── workflows.py          # Orchestration workflows for CLI & interactive menu
├── warframe_market.py        # Unified CLI & VBA entry point
├── Warframe_Market_Macros.bas # Standalone VBA module
├── Warframe Sell Stats.xlsm  # Primary macro-enabled workbook
├── requirements.txt          # Python dependencies
├── config.json               # Default configuration template
└── README.md                 # Documentation
```

---

## 🚀 Setup & Quickstart

### Step 1: Install Python Dependencies

Make sure you have Python 3.10+ installed on Windows. Open PowerShell or Command Prompt in the project folder and run:

```powershell
pip install -r requirements.txt
```

*(Dependencies: `openpyxl`, `requests`, `pywin32`, `websockets`)*

### Step 2: Run Initial Sync & First-Time Setup

Simply run:

```powershell
python warframe_market.py --sync
```

On your first run, the script will interactively ask for:
1. **Your Warframe.market username**
2. **Your JWT token** *(optional on first sync; you can also paste it into cell `J4` later)*

The script connects to Warframe.market, fetches your active sell orders, and generates `Warframe Sell Stats.xlsm` with all formulas, formatting, macro buttons, and the presence status selector.

*(Alternatively, you can pass your username directly on the command line: `python warframe_market.py --sync --user <YOUR_WARFRAME_MARKET_USERNAME>`)*

### Step 3: Open Excel & Enable Macros

1. Open `Warframe Sell Stats.xlsm` in Microsoft Excel.
2. Click **Enable Content / Enable Macros** when prompted.
3. If you haven't entered your JWT token yet, paste it directly into cell **`J4`**.
4. You're ready to trade!

---

## 🔑 How to Get Your JWT Token (Easy Step-by-Step)

Warframe.market uses a `JWT` cookie to authorize actions like updating prices, delisting sold items, or changing your online presence. Here is how to grab it in seconds.

### Method 1: The 10-Second Console Shortcut (Fastest & Easiest)

1. Open [https://warframe.market](https://warframe.market) in **Google Chrome**, **Microsoft Edge**, **Brave**, or **Opera** and make sure you are logged in.
2. Press **`F12`** on your keyboard (or right-click anywhere on the page and select **Inspect**).
3. In the developer window that opens, click the **Console** tab at the top.
4. Paste the following single line and press **`Enter`**:
   ```javascript
   copy(document.cookie.split('; ').find(row => row.startsWith('JWT='))?.split('=')[1] || '')
   ```
5. **Done!** Your JWT token is now automatically copied to your clipboard.
6. Switch to Excel and press **`Ctrl+V`** to paste it into cell **`J4`** (or enter it when prompted by the terminal).

---

### Method 2: Visual DevTools Guide (If Console Pasting is Disabled)

If your browser shows Chrome's "Don't paste code you don't understand" warning, follow these visual steps:

1. Log in to [https://warframe.market](https://warframe.market).
2. Press **`F12`** to open Developer Tools.
3. Look at the top tabs bar (`Elements`, `Console`, `Sources`, `Network`...).
   - *If you don't see `Application`, click the `>>` icon at the right end of the tabs bar to find it.*
4. Click on **Application** (in Google Chrome / Microsoft Edge) or **Storage** (in Firefox).
5. In the left panel, find the **Storage** section, expand **Cookies**, and click on `https://warframe.market`.
6. Look through the list for the cookie named **`JWT`**.
7. Double-click the long text in the **Value** column (starts with `eyJ...`).
8. Press **`Ctrl+C`** to copy it.
9. Open `Warframe Sell Stats.xlsm` and paste it into cell **`J4`**.

---

### Method 3: Using `config.local.json` (Optional Git-Safe File)

If you prefer to keep credentials in a config file, create `config.local.json` in the root folder (it is automatically ignored by `.gitignore`):

```json
{
  "username": "<YOUR_WARFRAME_MARKET_USERNAME>",
  "jwt_token": "<YOUR_JWT_TOKEN_HERE>"
}
```

---

## 🎮 How to Use Inside Excel

Open `Warframe Sell Stats.xlsm`. In **Column J**, you have full control over your inventory:

- **🟢 Status Selector (`J6`)**: Click the dropdown in cell `J6` and choose **Online**, **Online in Game**, or **Invisible**. It auto-syncs your live status to Warframe.market immediately in the background.
- **▶ 1. Sync from Market (`J7`)**: Fetches new sell orders from Warframe.market into Excel, styles rows, and sets up formulas.
- **⬆ 2. Push Prices & Stock (`J9`)**: Reads any modified prices, stock amounts, and visibility checkboxes (`☑`/`☐`) and pushes them live to Warframe.market.
- **🔄 3. Update All Time Revenue (`J11` or `I2`)**:
  - Rolls your Current Session sales into All-Time Revenue (`={base_all_time_rev}+G{row}`).
  - Resets Current Session quantities to `0`.
  - Automatically unchecks visibility (`☐`) for items whose stock reached 0.
  - Re-sorts inventory by Stock ASCENDING / Price DESCENDING.
- **⚡ 4. Refresh Columns & Formulas (`J13`)**: Re-applies all table styling, formulas, and button bindings.

---

## 💻 Command Line Interface (CLI)

You can also run all actions from PowerShell / Command Prompt:

| Command | Description |
| --- | --- |
| `python warframe_market.py` | Opens interactive terminal menu |
| `python warframe_market.py --sync` | Syncs active sell orders from Warframe.market to Excel |
| `python warframe_market.py --push` | Reads Excel and pushes price/stock/visibility updates to market |
| `python warframe_market.py --push -y` | Pushes updates without confirmation prompt |
| `python warframe_market.py --dry-run` | Previews price/stock differences without applying changes |
| `python warframe_market.py --commit` | Rolls Current Session revenue into All-Time Revenue and resets session |
| `python warframe_market.py --status "Online in Game"` | Sets market presence (`Online`, `Online in Game`, `Invisible`) |
| `python warframe_market.py --update-columns` | Re-applies formulas, dark-theme styling, and sorting |

---

## 🛡️ Privacy & Safety

- Your JWT token and username are stored locally on your machine in cells `J2` and `J4` of `Warframe Sell Stats.xlsm`.
- No OneDrive or external folders are accessed. All operations take place strictly within the current working directory.
- `config.local.json` is included in `.gitignore` by default to prevent accidental credential commits.
