# clean_csv.py — Excel/CSV Cleanup & Transformation Tool

A professional Python script that takes a messy CSV or Excel file and produces a clean, formatted Excel workbook — automatically fixing the most common data quality problems found in real-world spreadsheets.

Built for IT teams, operations staff, and small businesses who need reliable data cleanup without paying for an enterprise ETL tool.

---

## What it fixes automatically

| Problem | What happens |
|---|---|
| Duplicate rows | Detected and removed |
| Mixed date formats (`15/04/2020`, `2020-04-15`) | Normalised to `YYYY-MM-DD` |
| Inconsistent capitalisation (`JOHN`, `john`) | Title-cased correctly |
| ALL-CAPS or lowercase emails | Lowercased |
| Invalid email addresses | Flagged for manual review |
| Unformatted phone numbers (`5551234567`) | Formatted to `(555) 123-4567` |
| Currency/symbol in number fields (`$75,000`) | Stripped to clean numeric value |
| Inconsistent status fields (`ACTIVE`, `active`) | Normalised to `Active` |
| Missing required fields | Flagged with row number |
| Whitespace padding in any cell | Stripped |

---

## Output

Running the script produces a single `.xlsx` workbook with **three sheets**:

### Sheet 1 — Cleaned Data
- All fixes applied
- Rows that had any issue are highlighted amber so you can spot-check them
- Alternating row colours for readability

### Sheet 2 — Summary
- Row counts before and after
- Count of every fix type applied
- Items still needing manual review highlighted in red

### Sheet 3 — Issues Log
- Every change logged: row number, column, what was wrong, what it was changed to
- Items that need manual review (invalid emails, missing fields) highlighted separately

---

## Sample console output

```
════════════════════════════════════════════════════
  CSV CLEANUP COMPLETE
════════════════════════════════════════════════════

  Input / Output
  Original rows:                       20
  Cleaned rows:                        19
  Duplicates removed:                  1

  Fixes Applied
  ►  Dates normalised                  2
  ►  Names corrected                   5
  ►  Emails lowercased                 1
  ✔  Phones formatted                  0
  ✔  Numeric values cleaned            0
  ►  Status fields corrected           5

  Needs Manual Review
  ►  Missing required fields           2
  ►  Invalid emails flagged            1

  Total issues logged:  17

  Output: /Users/you/employees_cleaned.xlsx
════════════════════════════════════════════════════
```

---

## Requirements

```
Python 3.8+
pandas
openpyxl
```

Install dependencies:
```bash
pip install pandas openpyxl
```

---

## Usage

**Basic — clean a CSV with default settings:**
```bash
python clean_csv.py employees.csv
```

**Specify output file name:**
```bash
python clean_csv.py data.xlsx --output cleaned_data.xlsx
```

**Require multiple fields to be non-blank:**
```bash
python clean_csv.py report.csv --required Email,Phone,Salary
```

**Disable colour output (for logging or CI):**
```bash
python clean_csv.py data.csv --no-color
```

---

## Configuration

The top of the script has a configuration section where you can define which columns to treat as dates, names, emails, phones, and numerics. The script auto-detects columns by matching against these lists — no code changes needed for common column names.

```python
DEFAULT_DATE_COLUMNS    = ["Start Date", "Date", "DOB", "Hire Date"]
DEFAULT_NAME_COLUMNS    = ["First Name", "Last Name", "Full Name", "Name"]
DEFAULT_EMAIL_COLUMNS   = ["Email", "Email Address"]
DEFAULT_PHONE_COLUMNS   = ["Phone", "Phone Number", "Mobile", "Tel"]
DEFAULT_NUMERIC_COLUMNS = ["Salary", "Amount", "Revenue", "Cost", "Price"]
DEFAULT_STATUS_COLUMNS  = ["Status", "Active", "State"]
```

Add your own column names to any list and the script picks them up automatically.

---

## Project structure

```
clean_csv/
├── clean_csv.py          # Main script
├── sample_dirty.csv      # Example messy input file (for testing)
├── sample_cleaned.xlsx   # Example output (3-sheet workbook)
└── README.md
```

---

## Common use cases

- **HR teams** — clean new hire import files before uploading to HRIS
- **Finance** — normalise exported reports before analysis
- **Operations** — fix vendor-supplied data files with inconsistent formatting
- **IT admins** — clean AD export CSVs before bulk import scripts run against them

---

## License

MIT — free to use, modify, and distribute.

---

*Built as part of my Python automation portfolio. I build custom Python and PowerShell automation scripts for IT teams and small businesses. [View my Upwork profile →](https://www.upwork.com)*
