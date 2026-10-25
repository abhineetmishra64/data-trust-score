"# Trust Score App

Data Trust Dashboard for reviewing trust scores from Microsoft Purview and from saved raw JSON data.

## Overview

The application opens a local dashboard in a browser and lets you:

- load trust score data directly from Purview
- open or upload a raw JSON file from the `raw_json` folder
- view score summaries and details in the browser
- save the loaded raw data back to `raw_json`
- export an Excel report into the `output` folder
- override the built-in scoring with a local `config.json`

## Run the app

### From source

```bash
python webapp.py
```

This starts a local Flask app and opens the dashboard in your default browser.

### Windows desktop shortcut / .exe

- Double-click `TrustScore.exe`
- A console window opens and the dashboard opens in the browser
- Keep the console window open while the app is running
- Close the console window to stop the app

## Data sources

The dashboard supports two sources:

1. Trust Score from Purview
   - Reads live data from Purview
2. Trust Score from raw JSON
   - Select a file from the `raw_json` folder, or browse to another JSON file

The app stores saved raw JSON copies in `raw_json`, and exported Excel reports in `output`.

## Project structure

`purview_trust_score.py` remains the command-line entry point and compatibility interface used by the dashboard and the domain-based script. Shared logic lives in `purview_engine/`:

- `settings.py` loads configuration and defines shared paths and scoring settings.
- `api.py` handles Purview authentication and paginated requests.
- `data.py` converts raw responses to Trust Score rows and calculates scores.
- `helpers.py` contains shared response-shaping helpers.
- `report.py` builds the Excel workbook and dashboard.

## Configuration

A `config.json` file next to the executable can override the built-in scoring settings such as weights and score bands.

## Rebuild the executable

```bash
pyinstaller --onefile --name TrustScore --add-data "config.json;." --add-data "trust_app/web;trust_app/web" webapp.py
```

## Notes

- The first startup may take around 20 seconds.
- `Save raw JSON` keeps a copy of the loaded data in `raw_json`.
- `Excel report` saves the generated workbook in the `output` folder.
" 
