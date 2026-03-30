# 📈 Amazon Sales Report Automation

A standalone Python automation script that pulls weekly and year-to-date sales data from Amazon Vendor Central, merges it with NetSuite cost data, generates charts, and assembles a formatted HTML email report for distribution to sales and management teams.

Developed as part of a broader internal operations automation suite. Runs weekly and replaces a manual reporting process that previously took several hours.

---

## What It Does

1. **Logs into Amazon Vendor Central** via Selenium with email/password/TOTP 2FA
2. **Downloads weekly sales data** — pulls WTD and YTD revenue and units from the Retail Analytics dashboard for the most recent completed weeks
3. **Exports item-level data** — downloads the weekly sales CSV for detailed SKU-level analysis
4. **Pulls NetSuite data** — logs into NetSuite via Selenium, exports COGS and inventory data, merges with Amazon sales figures to calculate margin
5. **Generates charts** using matplotlib:
   - YoY weekly revenue line chart
   - WTD and YTD sales by category (revenue + units) pie charts
   - Category trend line charts (weekly revenue and units)
6. **Builds a formatted Excel report** with:
   - Weekly top 40 SKUs snapshot
   - YTD top 50 by revenue and by units
   - Executive summary with week-over-week and YoY comparisons
7. **Assembles an HTML email** in Outlook with embedded charts, formatted tables, and the Excel report attached — ready to review and send

---

## Output

- Weekly Excel report with multiple sheets (weekly snapshot, YTD revenue, YTD units)
- PNG chart files saved to a configurable data directory
- Formatted Outlook email draft with:
  - Executive summary metrics
  - Top 40 SKUs table
  - YTD top 50 tables (revenue + units side by side)
  - Embedded YoY line chart
  - Category pie charts and trend lines

---

## Tech Stack

| Tool | Purpose |
|---|---|
| Python | Core scripting |
| Selenium + webdriver-manager | Browser automation (Amazon VC + NetSuite login) |
| Pandas + NumPy | Data processing and merging |
| matplotlib | Chart generation |
| openpyxl | Excel report building |
| win32com (pywin32) | Outlook email assembly |
| pyotp | TOTP 2FA authentication |
| AI-Assisted Development | Claude (Anthropic) used for code generation and iteration |

---

## Configuration

All paths and credentials are loaded from environment variables. Create a `.env` file in the project root (see `.env.example`):

```
# Credentials
EMAIL=your-amazon-email@example.com
AMAZON_PASSWORD=your-password
AMAZON_KEY=your-totp-secret-key

LONG_EMAIL=your-netsuite-email@example.com
NETSUITE_PASSWORD=your-netsuite-password
NETSUITE_KEY=your-netsuite-totp-key

# Paths
DOWNLOAD_DIR=C:/Users/yourname/Downloads
LOG_DIR=C:/path/to/logs
SALES_DATA_DIR=C:/path/to/sales/data

# Email
REPORT_EMAIL_TO=recipient@your-company.com
```

---

## Project Structure

```
amazon-sales-report-automation/
├── amazon_sales_report.py    # Main script
├── Utilities/
│   ├── amazon_login.py       # Selenium login helper — Amazon Vendor Central
│   └── netsuite_login.py     # Selenium login helper — NetSuite
├── .env.example              # Environment variable template
├── .gitignore
└── README.md
```

---

## Running It

```bash
pip install selenium webdriver-manager pandas numpy matplotlib openpyxl pyotp pywin32 python-dotenv
python amazon_sales_report.py
```

> ⚠️ Requires Windows (uses win32com for Outlook). Credentials must be set in `.env` before running.

---

## Notes

- The script pauses at the Outlook draft stage — you review the email before sending
- Chart files are saved to `SALES_DATA_DIR/PNG Files/` and reused in subsequent runs
- Logging output is written to `LOG_DIR` with timestamped filenames
- NetSuite 2FA field IDs may need updating depending on your instance configuration (see `netsuite_login.py`)

---

*Part of a broader ops automation suite — see [ops-automation-dashboard](https://github.com/yourusername/ops-automation-dashboard) for the full Streamlit application.*

*Built by Rich Hardisty · [linkedin.com/in/richhardisty](https://linkedin.com/in/richhardisty)*
