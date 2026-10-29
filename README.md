# PBIX Inspector

Drag a .pbix (or .pbit) file onto a local web page and get an Excel workbook with
Columns, Measures, Relationships, Visual Design, M Queries and Tables.
Everything runs locally; nothing is uploaded.

## Setup (once)
    python -m venv venv
    venv\Scripts\activate          (Windows)   |   source venv/bin/activate  (Mac/Linux)
    pip install -r requirements.txt

## Run
    python pbix_inspector.py                  # opens the drag-and-drop page
    python pbix_inspector.py MyReport.pbix    # writes MyReport_inspection.xlsx

## Notes
- Requires Python 3.9+.
- For complete "Is Hidden" and column folder info, export from Power BI Desktop via
  File > Export > Power BI template (.pbit) and drop the .pbit instead.