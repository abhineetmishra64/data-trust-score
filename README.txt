Data Trust Dashboard
Double-click TrustScore.exe. A console window opens (keep it open) and the dashboard opens in your browser.
The first start takes ~20 seconds. Close the console window to stop the app.

Two sources:
- Trust Score from Purview: reads live from Purview.
- Trust Score from raw JSON: pick a file from the raw_json folder (or browse to one).
"Save raw JSON" keeps a copy of the loaded data in raw_json; "Excel report" saves into output.
Optional: a config.json next to the exe (weights, bands) overrides the built-in scoring.

Rebuild: pyinstaller --onefile --name TrustScore --add-data "config.json;." --add-data "trust_app/web;trust_app/web" webapp.py
