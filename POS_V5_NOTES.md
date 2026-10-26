# POS V5 - project notes (continue from here)

Owner: Amer (Lebanon, Arabic speaker), mobile shop POS (Windows .exe via PyInstaller). V5 is a SEPARATE app from V4
(own data folder `%APPDATA%\MobilePOS_V5`, exe `ShopPOS_V5.exe`). Upload this zip to a new chat and say "continue POS V5".

## Files
`index.html` (all UI, CSS at top) · `pos_core.py` (SQLite + logic) · `pos_server.py` (local API, sessions, roles) ·
`build_exe.bat`, `ShopPOS_V5.spec` · `README.txt` (user guide).

## Architecture notes
- API = functions decorated with `@api` in pos_server.py, called by the UI as `api(name,args)` / `act(name,args,okMsg)`.
- Sessions live in server memory (`SESSION`). If no users exist the app is open (owner). `OWNER_ONLY` = blocked for cashiers,
  `SENSITIVE` = cashier needs the owner PIN (`act()` prompts and retries with `_opin`). `redact()` strips costs/profit for cashiers.
- Every sale (days, Alfa, normal) is a normal `sales` row, so profit/debts/receipts/reports work everywhere.
- Money is stored in USD; LBP payments keep `cur` + `amount_lbp`. Rate in Settings (`lbp_rate`) + `rate_history`.
- Alfa stock = `alfa_ledger` (weighted average cost). Validity days: cards from products whose category contains "Alfa".
- `migrate()` adds missing tables/columns automatically. Wipe keeps settings, users, audit, rate history.
- Encrypted backups use `cryptography` (Fernet, PBKDF2) -> files `.zip.enc`.
- Arabic = dictionary `AR` + MutationObserver in index.html (add strings there to translate more).

## V5.1
Validity days: cards are chosen as wide tiles (tap = add one, "− Remove one" to undo) instead of a dropdown.

## V5.2
Fix: leftover cost in an empty Alfa stock inflated the cost of 1$ of the next Alfa (e.g. 7.047 instead of 0.889). `_alfa_zero_out()` now clears it
(visible "reset" row in Stock adjustments) whenever the Alfa stock is empty.

## V5.3
Inventory has a scan box: unknown barcode -> fills the Add product form (barcode filled, cursor on Name); known barcode -> "Add stock" dialog
(phones: opens the IMEIs window). API `stock_add`.

## V5.4
Validity days: cards no longer need to be in Inventory. "Card types" (table `alfa_cards`: name + cost, no stock) are managed from the Validity days
screen, plus one-off cards (type the cost). Card keys: `t:<id>` type, `p:<id>` inventory product, `c:<cost>:<name>` one-off (old numeric ids = products).

## V5.5
Suppliers > New purchase > "Import invoice file (Excel / CSV)": reads headers (English or Arabic: Item, Qty, Cost, Barcode, IMEI, Selling price, Category),
matches products by barcode/name, creates new ones (markup % when no selling price), fills the purchase invoice. See invoice_import_sample.xlsx.

## V5.6
- Safe data: `backup_hours` (hourly backups, 24 hourly files kept separately), `pos_latest` copy into the Drive folder, `last_backup`/`last_offsite`
  shown on Home, `backup_upload` API + "Restore from a backup file" in Settings.
- Shop network: `lan_mode` + `lan_port` (restart needed). Server binds 0.0.0.0 only if enabled AND at least one user exists; accepts only private-IP
  clients/Hosts. Sessions are per device (`X-Session` header, `SESSIONS` dict; `SESSION` = the session of the request being handled under LOCK).
  WhatsApp links open on the client device (`api("open_url")` uses window.open when not on localhost).

## V5.7
Selling Alfa cards to customers: card types have an optional selling price (`alfa_cards.price`). Alfa tab > Sale type "Alfa card" and quick buttons
"Sell Card 7.5 · $9.11" on the Sell screen. Profit = price - card cost (stored as a normal sale + `alfa_sales` row kind `card`); no Alfa stock involved.

## V5.8
Starred Alfa services (`alfa_services.fav=1`) are mirrored as real products (category "Recharge Cards", `service=1`, `products.alfa_service_id`) by
`sync_alfa_products()` (runs at startup and after every service save/delete). They show in Inventory (tag "Recharge Card", managed in the Alfa tab) and as tiles
on the Sell screen (the blue quick buttons were removed). `create_sale` takes the Alfa from the Alfa stock at the average cost for those lines (ledger + alfa_sales),
checks there is enough Alfa, and Void reverses it. Editing/deleting them from Inventory is blocked (edit in Alfa > Services).

## V5.9
Category layout: starred Alfa services -> products in category "Services" (linked by `alfa_service_id`; take Alfa from stock); Alfa card types that have a selling price
-> products in category "Recharge Cards" (linked by `alfa_card_id`, cost = card cost, no stock; sale logged in `alfa_sales` kind `card`). Both are kept in sync by
`sync_alfa_products()` and edited only in the Alfa tab / Validity days > Manage card types. The old top buttons on the Sell screen were removed.

## V5.10
Shop-network mode also accepts Tailscale addresses (100.64.0.0/10) so the system can be opened from home through a private Tailscale network
(no port forwarding, nothing exposed to the internet). Still requires users + PIN sign-in.

## V5.11
- **Touch** = a second credit provider with the same features as Alfa (stock ledger, services, card types, validity days, message fee `touch_fee`, adjustments).
  Implemented with a `provider` column ('alfa'|'touch') on alfa_ledger / alfa_services / alfa_cards / alfa_sales / validity_sales / validity_presets and a `provider`
  argument on the core functions (`PROVIDERS` dict in pos_core.py). The UI reuses viewAlfa/viewVal with a global `PROV` (the API wrapper injects `provider`,
  and the words "Alfa" are shown as "Touch" inside #main/#modal while PROV is "touch"). Touch products are categories "Touch Services" / "Touch Cards".
- **Returns of Alfa/Touch sales** give the credit back to the ledger at the original cost (`_alfa_return`, ledger kind `return`, negative alfa_sales line). Validity-days
  sales cannot be returned (use Void). Void also removes the return lines.
- **Trade-in** (used phones): tradeins table; creates/reuses a "Used phones" IMEI product, adds the IMEI at the paid cost, optional cash-out in the drawer.
  **Warranty claims**: claims table, linked to the sold IMEI/invoice; statuses Open / Sent to supplier / Repaired / Replaced / Rejected / Closed.
- **Bulk price change** by % with rounding (skips Alfa/Touch linked items), one-level **undo** (price_undo table). **Excel export** = reports "Inventory list" and
  "Phones in stock (IMEIs)".

## V5.12
Report "Sales for dashboard" (Reports > export as CSV): one row per item per day with columns day, category, item, qty, revenue, cost, paid_usd, paid_lbp_usd, on_debt
= the format the shared dashboard artifact "A-TECH shop sales" reads (claude.ai artifact QJi43eNGoYXzV9jNF5GKNg, dataset `sales`). Also fixed: the original
"Inventory valuation" report keeps its own columns (the new full list is "Inventory list").

## V5.13
- **Dashboard tab** (owner only) rebuilt from the shared artifact "A-TECH shop sales": period switch 7/30/90 days, tiles (revenue, profit, margin, items, new debt, each vs the previous
  period), revenue + profit per day chart, sales by category, where the profit comes from, how customers paid (USD / LBP / debt), searchable sortable sales-by-item table, and a
  "Right now" row (today's sales, debts, stock, Alfa/Touch stock). It reads live data through `dash_rows` (= report "Sales for dashboard", last 190 days) - no export needed.
- **Home** no longer shows the KPI cards/chart: it is a launcher (quick buttons) with the low-stock and installments lists and a warning if there is no off-site backup for 2 days.

## V5.14
- **Zoom**: header buttons - / 100% / + (60%-200%, step 10%), also Ctrl + / Ctrl - / Ctrl 0 and Ctrl + mouse wheel. Remembered per device (localStorage). CSS `zoom` on body, reset when printing.
- **Pictures** (managed only in Settings > Pictures): products, Alfa/Touch services, Alfa/Touch card types, and a picture for each screen (banner at the top of the tab + on its Home button).
  Stored in table `images` (kind product|service|card|tab, ref, mime, data, ver); the browser resizes before upload; served by `GET /img/<id>?v=<ver>&t=<token>` (cached). Products API returns
  `img` ("id.ver"; linked service/card products use the service/card picture). Wipe all data keeps screen pictures.

## V5.15
Pictures now fit completely (nothing cut off): CSS variable `--pfit` (setting `pic_fit`: contain = whole picture, cover = fill the box). Tile picture area is 100px high;
banners keep the picture's own proportions (max 200px high). Settings > Pictures has the "How pictures fit their box" choice.

## V5.16
Header button **Size** (modal): separate steps for Text (70-170%), Pictures (40-250%) and Boxes (60-200%), plus "Back to normal". CSS variables --fs / --pic / --box on <html>
(every `font-size:Npx` in the file is now `calc(Npx*var(--fs,1))`; picture sizes use --pic; the Sell/Home/Validity tile grids use --box). Saved per device in localStorage ("disp"),
ignored when printing. Independent of the -/100%/+ zoom (which scales everything).

## V5.17
- **Alfa and Touch tabs**: the separate "Validity days" tab is gone. Each tab shows four squares at the top (Services, Dollars, Cards, Validity days) instead of the "Sale type" select; the picture of each square
  is set in Settings > Pictures > "Sale-type squares" (images kind `sq`, ref `alfa:svc` ... `touch:val`; boot returns `sqpics`). `AMODE` = the selected square; `ACT.asq` switches.
- **Sell screen: "+ Top-up / Card"** (owner): name, selling price and profit; cost = price - profit. The line is a normal cart line without a product (`topup:true`); tables `topups` (per sale, with the
  player ID) and `topup_items` (quick buttons: the items used before, most used first, last price/profit). Cashiers only get the quick buttons: the server takes the profit from the saved item and ignores the
  client's cost; a new item asks for the owner PIN. Reports: "Top-up sales"; the dashboard category is "Game cards".

## V5.18
Sizes are per screen: zoom (- 100% +), text, pictures and boxes are saved separately for every tab (localStorage "disp2" = {all, tabs:{tabId:{z,fs,pic,box}}}); a tab without its own values uses
`all` (the old global values from V5.16/5.17 are migrated into `all`). Header zoom buttons and Ctrl +/-/0/wheel act on the current screen. Size panel: "Back to normal (this screen)" and
"Use these sizes on all screens". Alfa and Touch are separate screens. `applyDisp()` runs on every tab change (render()).

## V5.19
Sign-in with **username + password** (roles Owner / Cashier). `users.username` (unique, case-insensitive; old users get their name as username, their old PIN keeps working as the password until changed).
New passwords need 6+ characters (PBKDF2, salted). The sign-in screen shows the logo (84px) and shop name (28px); last username remembered per device; 5 wrong tries lock 60 s; the public API no longer lists user names.
The owner **recovery code** (XXXX-XXXX-XXXX, hash in setting `owner_recovery`, never returned by boot) is created with the first owner and shown once; "Forgot the owner password?" on the sign-in screen resets an
owner's password with it and issues a new code; Settings > Users > "New recovery code" rotates it. Cashier approvals (void, refund, big discount...) and "Wipe all data" now ask for the **owner password**.

## V5.20
**Updates from a file (no rebuild).** The .exe now starts `launcher.py` (rarely changes, `BUILTIN_VERSION` = the version the exe was built with). The launcher runs the version named in
`%APPDATA%\MobilePOS_V5\app\current.json` (folder `app\<version>\` with pos_core.py, pos_server.py, index.html) loading the files explicitly (importlib) so they win over the bundled copies, or the
built-in version. A new version starts with status `new`; the first successful `boot` call marks it `ok`; if it is not confirmed after 2 starts, or fails to import, the launcher goes back to the previous
version (or built-in). Settings > Updates: choose the zip (the same zip I deliver), it is validated (files present, Python compiles, `VERSION` found), a backup `before_update` is made, the version is
installed (last 3 kept), owner password required, then `app_restart` relaunches the exe (env `POS_RESTART_DELAY`, `PYINSTALLER_RESET_ENVIRONMENT`). A new Python library still needs one rebuild.
Security fix: `password`, `code` and `b64` are never written to the activity log (V5.19 logged typed passwords of user_save); `scrub_audit()` blanks old entries at start.

## Tested here
Backend API flows (roles, owner-PIN rules, audit, drawer USD/LBP, cash moves, recurring, reports, stock count, presets,
encrypted backup + restore + wrong password, import from V4 incl. old PIN) and the UI in headless Chromium (20 checks, no JS errors).
NOT tested here: building the .exe on Windows, real WhatsApp/email sending, scanning the Code 39 labels with a real scanner.

## Ideas not done
Touch/other credit providers, used phones (trade-in), warranty claims, bulk price update, Excel export of lists,
read-only phone view, encrypting the database file itself.
