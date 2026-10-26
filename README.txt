SHOP POS - V5
=============
BUILD THE .EXE (once)
  1. Install Python 3.12 from python.org  (tick "Add Python to PATH").
  2. Keep these files together: build_exe.bat, pos_server.py, pos_core.py, index.html
  3. Double-click build_exe.bat  ->  dist\ShopPOS_V5.exe   (make a desktop shortcut)
  Quick test without building:  py pos_server.py   (needs: pip install pywebview cryptography)

V5 IS A SEPARATE APP - your POS V4 is not touched
  Data   : %APPDATA%\MobilePOS_V5\pos.db      Backups: Documents\MobilePOS_V5_Backups
  To bring your old data: Settings > "Move from POS V4" > Import data from POS V4.

WHAT IS NEW IN V5
  USERS & ROLES   Settings > Users & PIN. Owner = everything. Cashier = no costs/profit, no Inventory/Suppliers/
                  Expenses/Offers/Reports/Settings. Risky actions (void, refund, big discount, selling below
                  price, stock adjustments, deletions) ask the OWNER PIN. 5 wrong PINs = 60 s lock.
  AUTO-LOCK       Settings > "Auto-lock after N minutes". PINs are stored salted (PBKDF2).
  ACTIVITY LOG    Settings > Activity log: who did what and when (voids, price/stock changes, rate, wipe, logins...).
  RATE HISTORY    Settings > Exchange-rate history.
  BACKUPS         Backup password (Settings) encrypts every backup (.enc). A backup is also made every time you
                  close the drawer. Restore asks for the same password.
  DRAWER          Close-out counts USD and L.L separately; Cash in / out log (supplier paid, personal, change...).
  DAY SUMMARY     New tab: the day's sales, cash by currency, debts, Alfa. Print / WhatsApp to you / Email.
  VALIDITY DAYS   One-tap presets ("+ Save current as preset").
  VALIDITY CARDS  (V5.1) cards are wide tiles: tap to add, tap again for another one, "- Remove one" to undo.
  SELL SCREEN     Quick buttons for your favourite Alfa services (Alfa > edit service > Quick button = Yes).
  RECEIPTS        80 mm thermal layout (or A4 in Settings), logo, WhatsApp button to send the receipt.
  STOCK COUNT     New tab: scan the shelf, see differences, apply. Phones are checked by IMEI.
  LABELS          Inventory > Print labels (Code 39 barcodes) - test one label with your scanner first.
  RECURRING       Expenses > Recurring expenses (rent, generator... added once a month).
  REPORTS         New: Profit by category, Dead stock (60 days), Reorder list by supplier, Cash in/out.
                  Reports can be sent by WhatsApp.
  CUSTOMERS       Statement of account: print or send by WhatsApp.
  HOME            7-day sales/profit chart.
  LANGUAGE        Settings > Language > Arabic (right-to-left). Main screens are translated; some small texts stay English.

UPDATES WITHOUT REBUILDING (V5.20)
  FIRST TIME ONLY: build the program with build_exe.bat (it now builds from launcher.py).
  AFTER THAT: when you get a new version zip, open the program > Settings > Updates > "Install an update from a file", choose the zip, type the owner password, and the program restarts
  on the new version. Your data is backed up first, the last 3 versions are kept ("Use this version" goes back), and if a new version does not start properly twice the program returns to the
  previous one by itself. Only install zips that come from your developer. A new Python library (rare) still needs one rebuild - you will be told.

SIGN-IN WITH USERNAME AND PASSWORD (V5.19)
  Settings > Users: add each person with full name, username, role (Owner or Cashier) and a password of 6+ characters. The first user you create is the Owner and the program then asks for sign-in.
  WRITE DOWN THE RECOVERY CODE shown once when the Owner is created: it is the only way back if the owner password is forgotten ("Forgot the owner password?" on the sign-in screen).
  Cashiers cannot see costs or profit; risky actions (cancel an invoice, refund, big discount, stock changes) ask for the OWNER password. Old PINs keep working as passwords until you change them.

SIZES PER SCREEN (V5.18)
  The zoom (- 100% +) and the Size button now apply to the screen you are on only: set Home big and the Dashboard stays as it was. "Back to normal (this screen)" resets just that screen;
  "Use these sizes on all screens" copies them everywhere. Each device remembers its own sizes.

ALFA / TOUCH SQUARES AND TOP-UPS (V5.17)
  Alfa and Touch: four squares at the top (Services, Dollars, Cards, Validity days). Validity days is no longer a separate tab. Pictures for the squares: Settings > Pictures > Sale-type squares.
  Sell screen: "+ Top-up / Card" - type the name, the selling price and your profit (iTunes 11$, profit 1$; PUBG 10$, profit 0.50$). The cost is worked out for you and the profit goes into the
  normal profit (Dashboard category "Game cards", report "Top-up sales"). Items you sold before come back as quick buttons.

SIZE BUTTON (V5.16)
  Header button "Size": make the TEXT, the PICTURES or the BOXES (Sell, Home, Validity days) bigger or smaller, each by itself, with "Back to normal".
  It is remembered on each device. The - 100% + buttons next to it zoom everything together.

ZOOM AND PICTURES (V5.14, fit fixed in V5.15)
  ZOOM      buttons - / 100% / + at the top (or Ctrl + / Ctrl - / Ctrl 0, or Ctrl + mouse wheel). It is remembered on each device and does not affect printing.
  PICTURES  Settings > Pictures: choose what to add pictures to (Products, Alfa / Touch services, Alfa / Touch cards, or Screens), then "Choose picture" next to an item.
            Pictures show whole (nothing is cut off); you can switch to "fill the box" in the same card. Best size for a tile picture: about 2:1 wide, e.g. 400 x 200.
            They show on Sell tiles, Inventory, Validity days tiles, the Alfa / Touch lists and the Home buttons. A Screens picture appears at the top of that tab.

DASHBOARD (V5.13)
  New Dashboard tab (owner only): sales, profit, margin, items and new debt for the last 7 / 30 / 90 days compared with the period before, revenue and profit per day,
  sales by category, where the profit comes from, how customers paid (USD / LBP / debt), sales by item, and a "Right now" row. It uses live POS data.
  Home is now a launcher: quick buttons, low stock, installments due and a backup warning.

TOUCH, RETURNS, TRADE-IN, WARRANTY, PRICES (V5.11)
  TOUCH tab           same as Alfa for the Touch provider: its own credit stock and cost, services, card types, validity days (use the two buttons at the top),
                      fee per message in Settings. Touch services/cards appear on the Sell screen under "Touch Services" / "Touch Cards".
  RETURNS             Sales > Return on an Alfa or Touch sale now puts the credit back in the stock at its original cost.
                      (A Validity days sale cannot be returned - use Void.)
  TRADE-IN tab        buy a used phone from a customer (model, IMEI, price paid, asking price, seller details). It enters Inventory as a used phone,
                      the payout can be taken from the drawer, and the list shows the profit once sold.
  WARRANTY tab        search a sold phone by IMEI / invoice / customer, open a claim, follow its status until it is repaired / replaced / closed.
  BULK PRICES         Inventory > Bulk price change (by category, % up or down, rounding) with "Undo last price change".
  EXCEL EXPORT        Inventory > Export to Excel (all items) and Export phone IMEIs (phones in stock). Files go to the exports folder.

SERVICES AND RECHARGE CARDS (V5.9)
  Sell screen / Inventory categories:  "Services" = your starred Alfa services (edit in the Alfa tab),
  "Recharge Cards" = the Alfa cards you sell, i.e. card types with a selling price (edit in Validity days > Manage card types).
  Both are sold from the cart like normal products and need no stock.

RECHARGE CARDS IN INVENTORY (V5.8, now called Services)
  Alfa tab > edit a service > "Show as a Recharge Card" = Yes. It then appears in Inventory (category/type "Recharge Cards") and on the Sell
  screen under the "Recharge Cards" chip, and is sold from the cart like any product (several in one invoice, discounts, debt, LBP...).
  The Alfa is taken from the Alfa stock at the current cost automatically. Edit price/Alfa amount in Alfa > Services.

SELLING ALFA CARDS (V5.7)
  Validity days > Manage card types: give each card a cost and a selling price. Then Sell screen shows "Sell Card 7.5 · $9.11" buttons
  (or Alfa tab > Sale type: Alfa card). Profit = selling price - cost of the card. No stock needed.

SAFE DATA (V5.6)
  Automatic backup every N hours (Settings, default 3). With a "Google Drive for Desktop" folder set, a fresh copy
  "pos_latest.zip" is also placed there every time = a copy OUTSIDE this computer. Home shows "Backups: last / off-site"
  and turns red if there is no off-site copy for 2 days. New computer: install the program, Settings > Backups >
  "Restore from a backup file" and pick the backup (if it is encrypted, set the same backup password first).

SECOND DEVICE / SECOND REGISTER (V5.6)  Settings > Shop network
  Tick "Allow other devices on the shop Wi-Fi...", Save, restart the program. Settings then shows an address like
  http://192.168.1.20:8765/ - open it in the browser of the other PC / phone / tablet (same Wi-Fi) and sign in with
  your own user and PIN. It is ONE database on the main computer, so sales from both devices are always in sync.
  Needs: the Owner user created first, the main computer ON, Windows Firewall "Allow" (private network) the first time.
  Printing / scanners work on the device you are using.
  FROM HOME (V5.10): install the free Tailscale app on the office computer AND on your phone, sign in to the SAME account on both,
  turn the shop network on in Settings, then on the phone open the address that starts with 100. (shown in Settings > Shop network),
  for example http://100.101.102.103:8765/ . It is encrypted and nothing is opened on the office router. The office computer
  must be on, online, and Tailscale running.

FROM V4 (still there)
  Alfa credit + validity days, Lebanese lira, logo, calculator, Wipe all data (needs DELETE + PIN), accent colour.

GOOGLE DRIVE / EMAIL BACKUP: same as before (Settings > Cloud & email backup).
SHORTCUTS: F2 = scan box, F9 = complete sale.

CHANGING THE DESIGN: everything is in index.html (<style> at the top). Edit with Notepad++ / VS Code, then run
build_exe.bat again - or ask Claude.
