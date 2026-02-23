# Afterwork Winter Olympics (local webpage)

A tiny local webpage for showing teams and tracking points across 5 events.

## Run locally

From this folder:

```bash
python3 -m http.server 8000
```

## Stop the webserver

In the terminal where `python3 -m http.server 8000` is running, press `Ctrl+C`.

Then open:

- http://localhost:8000

## Excel format

The importer looks for a **Team** / **Team Name** column and up to 4 member columns.

Common header names supported:

- Team: `Team`, `Team Name`
- Members: `Member 1..4`, `Player 1..4`, `Name 1..4`

If no member headers match, it assumes the 4 columns after the team column are members.

## Save / Load results

Use the buttons in the **Teams** tab:

- **Save (.xlsx)**: Exports teams + all event results to an Excel workbook (multiple sheets).
- **Save (.csv)**: Exports teams + all event results to a single CSV file (with a `RecordType` column).
- **Load saved file**: Imports a previously saved `.xlsx` or `.csv` created by the buttons above.

Note: Excel save/load requires the `xlsx` library to be available in the browser. If you are offline or the CDN is blocked, use **Save (.csv)**.

## Scoring (current defaults)

- Event 1 (Bobsled, 1v1): winner=3
- Event 2 (Ice hockey, 1v1): winner=3
- Event 3 (Curling, 1v1): winner=3
- Event 4 (Biathlon, timed): only top 3 score (1st=3, 2nd=2, 3rd=1)
- Event 5 (Skijump, 1v1): winner=3

Data persists in your browser via `localStorage`.
