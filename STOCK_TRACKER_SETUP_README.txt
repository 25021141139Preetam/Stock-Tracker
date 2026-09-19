================================================================================
          AUTOMATED NSE STOCK PRICE TRACKER - SETUP & USER GUIDE
================================================================================

1. OVERVIEW
-----------
This system provides fully automated tracking of your NSE stock portfolio,
fetching live market data directly from Yahoo Finance, updating your Excel
tracker, calculating holding values and portfolio allocations, and updating
dashboard metrics and performance charts.

Tracked NSE Holdings:
  - JSW Steel   (Ticker: JSWSTEEL.NS)  - Quantity: 152 shares
  - Coal India  (Ticker: COALINDIA.NS) - Quantity: 239 shares
  - Eternal     (Ticker: ETERNAL.NS)   - Quantity: 313 shares


2. FILES IN THIS SYSTEM
-----------------------
  - automated_stock_price_tracker.xlsx : Main Excel workbook (Tracker & Dashboard).
  - portfolio_dashboard.html           : Interactive modern web dashboard with live charts.
  - stock_tracker.py                   : Core Python automation engine.
  - stock_tracker_automation.py        : Wrapper script maintaining full compatibility.
  - run_stock_tracker.bat              : One-click Windows batch launcher.
  - requirements.txt                   : Python package dependencies.
  - STOCK_TRACKER_SETUP_README.txt     : This setup and scheduling guide.
  - backups/                           : Automatic timestamped safety backups.


3. PYTHON REQUIREMENTS & INSTALLATION
-------------------------------------
The automation requires Python 3.10+ and the following libraries:
  - openpyxl (for safe Excel read/write without breaking charts/formulas)
  - yfinance (for automated Yahoo Finance market data retrieval)
  - pandas
  - requests

To install requirements:
  Open Command Prompt or PowerShell in this directory and run:
      pip install -r requirements.txt


4. HOW TO RUN MANUALLY
----------------------
Option A: One-Click Launcher (Easiest)
  - Double-click "run_stock_tracker.bat".
  - The script will detect your Python installation, fetch the latest closing
    prices, update Sheet 1 and Sheet 4, and display the update summary.
  - The terminal window will stay open until you press any key.

Option B: Command Prompt / PowerShell
  - Open terminal in this folder and run:
      python stock_tracker.py

Option C: Dry Run (Test Without Modifying Excel)
  - To test market data connectivity without saving to the Excel file:
      python stock_tracker.py --dry-run


5. HOW TO SCHEDULE DAILY (WINDOWS TASK SCHEDULER)
-------------------------------------------------
NSE trading hours are 09:15 to 15:30 IST (Monday - Friday).
It is recommended to schedule the script to run Monday to Friday at 16:00 IST
(4:00 PM), which allows 30 minutes for post-market closing prices to settle.

Method 1: Fast Setup via PowerShell (Run as Administrator)
----------------------------------------------------------
Open PowerShell as Administrator and run this command:

$Action = New-ScheduledTaskAction -Execute "C:\Users\ASUS\Desktop\Stock tracker\run_stock_tracker.bat" -WorkingDirectory "C:\Users\ASUS\Desktop\Stock tracker"
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 4:00PM
Register-ScheduledTask -TaskName "NSE_Stock_Price_Tracker" -Action $Action -Trigger $Trigger -Description "Daily automated NSE stock price tracker update"

Method 2: Using the Windows Task Scheduler GUI
----------------------------------------------
1. Press Windows Key + R, type "taskschd.msc" and press Enter.
2. In the right panel, click "Create Task...".
3. In the "General" tab:
   - Name: "NSE Stock Price Tracker"
   - Check: "Run only when user is logged on" (or "Run whether user is logged on or not" if preferred).
4. In the "Triggers" tab:
   - Click "New...".
   - Begin the task: "On a schedule".
   - Select "Weekly".
   - Check: Monday, Tuesday, Wednesday, Thursday, Friday.
   - Start time: 16:00:00 (4:00 PM).
   - Click OK.
5. In the "Actions" tab:
   - Click "New...".
   - Action: "Start a program".
   - Program/script:
       C:\Users\ASUS\Desktop\Stock tracker\run_stock_tracker.bat
   - Start in (optional, but HIGHLY recommended):
       C:\Users\ASUS\Desktop\Stock tracker
   - Click OK.
6. In the "Conditions" tab:
   - Check: "Start only if the following network connection is available" -> Any connection.
7. Click OK to save the task.


6. BUILT-IN SAFETY & ERROR HANDLING
-----------------------------------
- Duplicate Date Prevention:
  If run multiple times on the same date or over the weekend, the script detects
  the existing row and refreshes the closing prices in place rather than creating
  duplicate rows.
- Weekend & Holiday Detection:
  If run on a Saturday, Sunday, or market holiday, the script automatically uses
  the latest available completed NSE session and informs you.
- Safe Atomic Saves & Backups:
  Before modifying the Excel file, an exact timestamped copy is saved in the
  "backups/" folder. The update is written atomically using a temporary file to
  prevent file corruption.
- Network / Market Outage Resilience:
  Includes automated socket checks, yfinance fetching, and a direct HTTP API fallback.
  If the network is unavailable, a clear error message is logged without corrupting
  the Excel file.
- Chart & Formula Preservation:
  The script expands the Tracker table ("TrackerTable") and updates the line chart
  series in "Sheet 4 - Dashboard" so all charts and summary tiles automatically reflect
  the newest data.
================================================================================
