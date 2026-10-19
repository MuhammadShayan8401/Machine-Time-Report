# Machine Time Report

A small Streamlit app that turns daily machine **Production Report PDFs** into an hourly Excel sheet showing each machine's on time, idle time, off time and missing time.

Upload the PDFs, preview the result on screen, then download it as Excel or CSV.

## Features

- Upload several report PDFs at once (one per machine)
- Hourly table per machine, from 09:00 to 09:00 the next day, with shift (A or B)
- Missing time calculated per hour, with a red/green flag against a threshold
- On-screen preview with a machine filter and an "above threshold only" option
- Summary tab with totals per machine
- Download as Excel (`.xlsx`) or CSV
- Optional list of machines that must appear, so a machine with no data is flagged

## How the numbers are calculated

Each report row (for example `09:23 - 10:00`) is added to the clock hour it starts in.

| Column | Meaning |
|---|---|
| On Time | Sum of the report's On Time for rows starting in that hour |
| Idle Time | Sum of the report's Idle Time for rows starting in that hour |
| Off Time | Sum of the report's Off Time. An hour with no report rows at all counts as 1:00:00 off |
| Missing Time | `1 hour - (On + Idle + Off)`, never below 0 |
| Status | Red if Missing Time is above the threshold (default 10 minutes), otherwise green |

Shift A is 09:00-21:00 and Shift B is 21:00-09:00.

## Run locally

Requires Python 3.9 or newer.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the address shown in the terminal (usually http://localhost:8501).

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Choose the repo, set the main file to `app.py`, and click **Deploy**.
4. If the reports are sensitive, set the app to private in its settings.

## Use without the web page

`make_report.py` also works from the command line:

```bash
python make_report.py path/to/pdf_folder -o report.xlsx --threshold 10 --expect "B650 T"
```

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit page (upload, preview, download) |
| `make_report.py` | PDF parser and Excel builder |
| `requirements.txt` | Python dependencies |

## Limitations

- The parser reads the report's "Hours" table by column position. If the report layout changes, `parse_pdf` in `make_report.py` may need updating.
- Values are taken from the reports as they are. If a report shows more On Time than a row's time span, that error carries into the output.
- Machine names come from the PDF header, for example `Z140 T(A)` is shown as `Z140 T (A)`.
