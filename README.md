# Outlook Alerts Email – R Template

A **generic** R script template to build and preview an HTML email with alert tables in **Microsoft Outlook** (Windows via COM).
No company references, no real data — uses mock data by default.

> ⚠️ This repository is for **demonstration purposes only**. It does not represent the policies or practices of any institution.

---

## Requirements

* **Windows** with **Microsoft Outlook** configured (with a profile)
* **R** 4.1+ recommended
* R packages: `RDCOMClient`, `dotenv`

```r
install.packages(c("RDCOMClient", "dotenv"))
```

---

## Configuration

1. Copy `.env.example` to `.env` and update values:

   ```env
   ALERT_TO=alerts@example.com
   ALERT_CC=
   THRESHOLD_GENERAL=5
   THRESHOLD_SPECIAL=3
   ```
2. Do **not** commit `.env` (already ignored in `.gitignore`).

---

## Usage

In R (RGui or RStudio):

```r
source("liquidity_alerts_email.R")
send_email()
```

* The script opens an email in Outlook with tables and alerts (mock data).
* To send automatically, uncomment `mail$Send()` in the script (test first!).

---

## Customization

* **Replace mock data** with:

  * CSV/Excel files (`read.csv()`, `readxl::read_excel()`),
  * Database queries (`DBI`, `odbc`),
  * API calls (`httr`, `jsonlite`).
* **Thresholds**: controlled via `.env`.
* **Branding**: adjust the `<style>` block and signature HTML.
* **Security**: keep credentials/recipients in `.env` or OS environment variables.

---

## Windows / Outlook Notes

* `RDCOMClient` depends on local Outlook; not compatible with macOS/Linux.
* Automatic sending may trigger security prompts depending on IT policy.
* For errors, test manually:

  ```r
  RDCOMClient::COMCreate("Outlook.Application")
  ```

---

## Repository Structure

```
/r-outlook-alerts-template/
├─ liquidity_alerts_email.R   # main R script (sanitized, mock data)
├─ .env.example               # sample environment variables (no secrets)
├─ .gitignore                 # ensures .env and data files aren’t committed
└─ README.md                  # this file
```

---

## File Overview

* **`liquidity_alerts_email.R`**
  Builds an HTML email with three example tables (Negative Outliers, Top Performers, Lowest Performers), highlights alerts above thresholds, and opens a draft in Outlook.

* **`.env.example`**
  Template for environment variables (`ALERT_TO`, `ALERT_CC`, `THRESHOLD_*`). Copy to `.env` locally.

* **`.gitignore`**
  Prevents committing sensitive files (e.g., `.env`, data files).

---

## Quick Start

1. Clone or download this repository.
2. Create your `.env` from `.env.example` and set recipients/thresholds.
3. Install required packages:

   ```r
   install.packages(c("RDCOMClient", "dotenv"))
   ```
4. Run:

   ```r
   source("liquidity_alerts_email.R")
   send_email()
   ```
5. Review the Outlook draft. When confident, you can enable automatic send by uncommenting `mail$Send()`.

---

## Command Line (Optional)

If you want to run via `Rscript`:

```bat
Rscript -e "source('liquidity_alerts_email.R'); send_email()"
```

> Ensure your working directory is the repo root (where the `.env` lives), or set `dotenv::load('path/to/.env')` explicitly in the script.

---

## Troubleshooting

* **`COMCreate` fails**
  Ensure Outlook is installed and a profile is configured. Test:

  ```r
  RDCOMClient::COMCreate("Outlook.Application")
  ```

* **Security prompts when sending**
  Keep `mail$Display()` for manual send during testing. Coordinate with IT for MAPI/COM send policies if you need unattended sending.

* **Email rendering differences**
  Email clients vary. Keep HTML/CSS simple and inline. Test on Outlook desktop, Outlook web, Gmail mobile, etc.

* **Encoding issues**
  Ensure UTF-8. The template includes `<meta charset='UTF-8'>` in the HTML head.

---

## Security & Compliance Checklist

* [x] No company names, emails, or logos in code/README
* [x] Mock data only
* [x] Real recipients and thresholds live in **`.env`** (not committed)
* [x] `.gitignore` blocks `.env` and data files
* [x] No internal paths, databases, or API keys

---

## Roadmap (Optional Enhancements)

* Configurable per-fund thresholds from JSON/CSV
* Conditional row coloring (positive/negative) in tables
* Logging (write HTML snapshots and send status to disk)
* Cross-platform sending via SMTP (`blastula`) instead of Outlook COM
* Optional Python helper via `reticulate` for shared templates/recipients

---

## License

**MIT** — feel free to use, modify, and distribute.
