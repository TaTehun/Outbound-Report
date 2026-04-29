# Daily Outbound Report — Automation Pipeline

An end-to-end Python automation pipeline that downloads daily outbound report attachments from email, consolidates them into a single Excel workbook, and distributes the final report via email — all without manual intervention.

---

## Overview

Each business day, multiple carriers and logistics partners send their outbound report files to a shared mailbox. This pipeline:

1. **Downloads** today's report attachments from the POP3 mailbox
2. **Transforms** each source file into a standardized format
3. **Consolidates** all sources into a single workbook with a refreshed pivot table
4. **Distributes** the final `.xlsb` report via SMTP
5. **Alerts** stakeholders if any files were not received

---

## Project Structure

```
outbound_report/
├── main.py                        # Data processing & Excel consolidation
├── email_pipeline/
│   ├── config.py                  # All settings (paths, servers, targets)
│   ├── data_downloader.py         # POP3 attachment downloader
│   ├── run_downloader.py          # Pipeline entry point
│   └── smtp_notifier.py           # Email sender (report + failure alert)
└── utils/
    └── date_utils.py              # Date formatting utilities
```

---

## Data Sources

| Source | Description |
|--------|-------------|
| Provider A | Daily exception list (Excel) |
| Provider B | Daily exception list (Excel) |
| Provider C | Outbound tracking report (Excel) |
| Provider D | Daily outbound report (CSV) |
| Provider E | Daily manifest report (Excel) |

> If a file is not received on a given day, that source is skipped and the remaining files are still consolidated.

---

## Output

All files are organized into date-stamped folders:

```
Outbound/
├── Data/
│   ├── template.xlsx              # Master template (column headers, pivot, formatting)
│   └── 04-29-2026/                # Today's downloaded attachments
└── Output/
    └── 04-29-2026/
        └── Consolidated Daily Manifest - 04.29.2026.xlsb
```

---

## How It Works

### 1. Download
The downloader connects to the POP3 server and scans incoming emails **newest-first**, stopping as soon as it reaches emails older than today. For each email, it checks the subject and sender against the configured targets before downloading the attachment — minimizing unnecessary data transfer.

### 2. Transform
Each source file goes through source-specific transformations (column reordering, address splitting, date formatting, blank column insertion) to conform to a standard 19-column layout.

### 3. Consolidate
All transformed DataFrames are concatenated and written into the `Daily Manifest` sheet of the template workbook, starting at row 2 to preserve the pre-formatted header. The pivot table is then refreshed against the new data.

### 4. Distribute
The final `.xlsb` file is attached to a report email and sent to the configured recipients. If any files were missing, a separate failure alert is sent detailing which files were not received and which ones were.

---

## Configuration

All settings are managed in `email_pipeline/config.py`:

| Setting | Description |
|---------|-------------|
| `DATA_DIR` / `OUTPUT_DIR` | Date-stamped output folders |
| `TEMPLATE_PATH` | Path to `template.xlsx` |
| `EMAIL_TARGETS` | List of expected email sources with sender/subject criteria |
| `EMAIL_SEARCH_COUNT` | Max number of recent emails to scan |
| `REPORT_TO` / `REPORT_CC` | Report email recipients |
| `NOTIFICATION_TO` | Failure alert recipients |

Credentials (POP3/SMTP username and password) are stored in a `.env` file and loaded at runtime.

---

## Requirements

```
pandas
xlwings
openpyxl
pytz
python-dotenv
```

---

## Running

```bash
python main.py
```

This runs the full pipeline: download → transform → consolidate → send.

To run the downloader only:

```bash
python -m email_pipeline.run_downloader
```
