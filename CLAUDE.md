# CRM Google Sheet Project — Full Context

## What Was Built
A CRM (Customer Relationship Management) Google Sheet with:
- **Tab 1 — CRM Data**: Main data entry sheet
- **Tab 2 — Charts**: Dashboard with summary tables and chart placeholders
- **Tab 3 — Master Data**: Reference data for customers and collectors

---

## Final Deliverable

**File:** `CRM_Sheet.xlsx`
**Location:** `C:\Users\KGN\Desktop\Projects\Excel\CRM_Sheet.xlsx`
**Builder script:** `C:\Users\KGN\Desktop\Projects\Excel\build_crm.py` (Python/openpyxl)

### How to Deploy
1. Upload `CRM_Sheet.xlsx` to Google Drive
2. Right-click → Open with → Google Sheets
3. All formulas, dropdowns, and formatting are already embedded

---

## CRM Data Sheet — Column Structure

| Col | Field | Type |
|-----|-------|------|
| A | Customer Name | Dropdown (from Master Data col A) |
| B | Company Name | Auto-filled via VLOOKUP from Master Data |
| C | Phone Number | Auto-filled via VLOOKUP from Master Data |
| D | Status | Dropdown: Pending / Paid / Overdue / Partial / Cancelled |
| E | Bill Generate Date | Date (dd/mm/yyyy) |
| F | Amount | Currency |
| G | Collected Amount | Currency |
| H | Collection Date | Date (dd/mm/yyyy) |
| I | Collected By | Dropdown (from Master Data col E) |
| J | Pending Amount | Formula: =F-G (auto-calculated) |
| K | Delay Days | Formula: =TODAY()-E (blank when Paid) |
| L | Contacted Date | Date (dd/mm/yyyy) |
| M | Last Contacted Date | Date (dd/mm/yyyy) |
| N | Notes | Free text |

---

## Master Data Sheet — Structure

| Column | Content |
|--------|---------|
| A | Customer Name |
| B | Company Name |
| C | Phone Number |
| E | Collectors list |
| G | Status options (not used for dropdown — Status uses hardcoded list) |

**To add customers:** Type in Master Data col A/B/C — dropdown updates automatically.
**To add collectors:** Type in Master Data col E — dropdown updates automatically.

---

## Conditional Formatting

| Rule | Colour |
|------|--------|
| Status = Paid | Green background |
| Status = Pending | Yellow background |
| Status = Overdue | Red background |
| Status = Partial | Blue background |
| Status = Cancelled | Grey background |
| Delay Days > 30 | Red background |
| Delay Days 1–30 | Yellow background |

---

## Charts Tab
Contains live summary tables with SUMIF/COUNTIF formulas:
- **By Status table** (rows 4–10): Count, Billed, Collected, Pending per status
- **Monthly Trend table** (rows 13–19): Last 6 months of Billed/Collected/Pending

**To add charts in Google Sheets:**
- Select A4:E9 → Insert → Chart → Pie or Column chart
- Select A13:D19 → Insert → Chart → Line chart

---

## Auto-Status Feature (Optional)
The Excel file does NOT auto-update Status when Collected Amount is typed.
To enable this in Google Sheets, paste this into Extensions → Apps Script and Save:

```javascript
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'CRM Data') return;
    var row = e.range.getRow(), col = e.range.getColumn();
    if (row < 2 || col !== 7) return;
    var amt  = sheet.getRange(row, 6).getValue();
    var coll = Number(e.value) || 0;
    if (!amt) return;
    sheet.getRange(row, 4).setValue(coll >= amt ? 'Paid' : coll > 0 ? 'Partial' : 'Pending');
  } catch (_) {}
}
```

No need to run it — just save. It activates automatically on edit.

---

## Why Apps Script Was Abandoned

Multiple attempts were made to build the sheet via Google Apps Script (`CRM_Sheet.gs` still exists in this folder for reference). All failed due to:

1. **Google Apps Script 6-minute execution limit** — formatting 100+ rows, building charts, and setting range-based data validations collectively exceeded this limit
2. **Google account storage nearly full** — shown in the "Almost out of storage" banner in Google Sheets; this severely throttles Apps Script API calls
3. **`requireValueInRange` is slow** — applying range-referenced dropdowns to 100+ rows takes minutes on its own
4. **Stale installable triggers** — earlier script versions installed `onCRMEdit` triggers that ran in the background consuming quota

**Solution:** Switched to Python/openpyxl to generate the `.xlsx` file locally, then upload to Google Drive.

---

## Rebuild Instructions
If the Excel file needs to be regenerated:

```bash
cd "C:\Users\KGN\Desktop\Projects\Excel"
python build_crm.py
```

Then re-upload `CRM_Sheet.xlsx` to Google Drive.

---

## User Notes
- User's Google account had "Almost out of storage" — this was a major factor in failures
- User is on Windows 11, using Google Sheets via browser
- Python 3.14.2 is installed at `C:\Python314\`
- openpyxl 3.1.5 is installed
