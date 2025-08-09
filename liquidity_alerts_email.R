# Outlook Alerts Email – R Template (Generic, Sanitized)
# ------------------------------------------------------
# Purpose: Build and preview an HTML email with alert tables in Microsoft Outlook.
# Notes:
#  - Uses only mock data. Replace with your data sources.
#  - No company names, emails, or internal paths.
#  - Recipients and thresholds are loaded from environment variables (.env).

rm(list = ls())

# (Optional) function to detach non-base packages
detachAllPackages <- function() {
  basic <- c("package:stats","package:graphics","package:grDevices",
             "package:utils","package:datasets","package:methods","package:base")
  pkgs <- setdiff(search()[grepl("^package:", search())], basic)
  for (p in pkgs) detach(p, character.only = TRUE)
}
# detachAllPackages()

# Minimal dependencies
if (!requireNamespace("RDCOMClient", quietly = TRUE)) {
  stop("Please install.packages('RDCOMClient')")
}
if (!requireNamespace("dotenv", quietly = TRUE)) {
  stop("Please install.packages('dotenv')")
}

library(RDCOMClient)
library(dotenv)

# Load environment variables from .env (if present)
dotenv::load(".env")

get_env_num <- function(key, default) {
  val <- Sys.getenv(key, unset = NA)
  if (is.na(val) || !nzchar(val)) return(as.numeric(default))
  suppressWarnings(as.numeric(val))
}

send_email <- function(
  recipients = list(
    to = Sys.getenv("ALERT_TO", "alerts@example.com"),
    cc = Sys.getenv("ALERT_CC", "")
  ),
  alert_thresholds = list(
    general = get_env_num("THRESHOLD_GENERAL", 5),
    special = get_env_num("THRESHOLD_SPECIAL", 3)
  )
) {
  # ----------------------------
  # MOCK DATA (replace with real)
  # ----------------------------
  negative_outliers <- data.frame(
    Fund   = c("Alpha Fund", "Beta Fund", "Special Fund", "Delta Fund"),
    Change = c(-6.5, -5.2, -3.5, -4.9),
    stringsAsFactors = FALSE
  )
  positive_top <- data.frame(
    Fund   = c("Epsilon Fund", "Phi Fund", "Gamma Fund"),
    Change = c(4.5, 3.8, 5.1),
    stringsAsFactors = FALSE
  )
  negative_bottom <- data.frame(
    Fund   = c("Eta Fund", "Theta Fund", "Iota Fund"),
    Change = c(-2.5, -3.2, -3.0),
    stringsAsFactors = FALSE
  )

  report_date <- Sys.Date()

  # HTML table builder
  build_html_table <- function(df, title){
    if (!nrow(df)) return(paste0("<p>No data for ", title, ".</p>"))
    header <- paste0("<h3>", title, "</h3>")
    rows <- apply(df, 1, function(r) {
      val <- suppressWarnings(as.numeric(r[["Change"]]))
      sprintf("<tr><td>%s</td><td>%+.2f%%</td></tr>", r[["Fund"]], val)
    })
    paste0(
      header,
      "<table><thead><tr><th>Fund</th><th>Change (%)</th></tr></thead><tbody>",
      paste(rows, collapse = ""),
      "</tbody></table>"
    )
  }

  # Alert message builder
  build_alerts <- function(df){
    alerts <- apply(df, 1, function(r){
      fund <- r[["Fund"]]
      chg  <- suppressWarnings(as.numeric(r[["Change"]]))
      thr  <- if (fund == "Special Fund") alert_thresholds$special else alert_thresholds$general
      if (is.na(chg)) return("")
      if (abs(chg) > thr)
        sprintf("<p style='color:#B00020;font-weight:bold;margin:6px 0;'>Alert: %s at %+.2f%% exceeds %.2f%%.</p>",
                fund, chg, thr)
      else ""
    })
    paste(alerts, collapse = "")
  }

  # Email HTML body
  body <- paste0(
    "<html><head><meta charset='UTF-8'><style>",
    "body{font-family:Calibri,Arial,sans-serif;line-height:1.3}",
    "table{border-collapse:collapse;width:100%;margin:8px 0}",
    "th,td{border:1px solid #333;padding:6px;text-align:left}",
    "th{background:#f2f2f2}",
    "h2,h3{margin:10px 0 6px 0}",
    "</style></head><body>",
    "<h2>Alerts Report – ", format(report_date, "%Y-%m-%d"), "</h2>",
    "<p>This is a generic template with mock data.</p>",
    build_html_table(negative_outliers, "Negative Outliers"),
    build_html_table(positive_top, "Top Performers"),
    build_html_table(negative_bottom, "Lowest Performers"),
    build_alerts(negative_outliers),
    "<p>Regards,<br><strong>Alerts Bot</strong></p>",
    "</body></html>"
  )

  # Outlook COM automation
  outlook <- RDCOMClient::COMCreate("Outlook.Application")
  mail <- outlook$CreateItem(0)
  mail[["To"]]      <- recipients$to
  mail[["Cc"]]      <- recipients$cc
  mail[["Subject"]] <- paste0("Alerts | ", format(report_date, "%Y-%m-%d"))
  mail[["HTMLBody"]]<- body
  mail$Display()    # Preview (recommended)
  # mail$Send()     # Automatic send (use with caution)
}

# Example run:
# send_email()
