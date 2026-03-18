# Financial Chatbot Documentation

## Overview
A Next.js-based financial chatbot (v3) that queries financial data from Google Drive CSV files. Deployed on Vercel with Clerk authentication.

**Logo:** `public/logo.png` (CW_logo.png from project root)

---

## 🚀 Quick Start

Type any of these commands in the chatbot:

| Command | What It Does | Example |
|---------|-------------|---------|
| **Analyze** | Run full 6-comparison financial analysis | `Analyze` |
| **Detail X** | Drill into analysis comparison #X | `Detail 3` |
| **Detail X.Y** | Drill into specific item in comparison #X | `Detail 3.1` |
| **Total [Item] [Type]** | Sum sub-items under a parent | `Total Preliminaries Cashflow` |
| **Compare X with Y** | Compare two Financial Types | `compare projected gp with budget gp` |
| **[Type] [Metric]** | Direct query | `projected gp` |
| **[Metric] [Month] [Year]** | Query with date | `gp jan 2025` |

### Shortcut Cheat Sheet
- **gp** = Gross Profit, **np** = Net Profit
- **bp** / **budget** = Business Plan
- **wip** = Audit Report (WIP)
- **cf** / **cashflow** = Cash Flow
- **projection** / **projected** = Projection as at
- **committed** = Committed Value / Cost
- **revision** / **rev** = Revision as at
- **prelim** = Preliminaries, **subcon** = Subcontractor

---

## Data Source
- **Location:** `G:/My Drive/Ai Chatbot Knowledge Base`
- **Structure:** Year | Month | Sheet_Name | Financial_Type | Item_Code | Trade | Value

### ⚠️ IMPORTANT: Data Values are in Thousands ('000)

**All monetary values in the CSV data are in thousands.** This means:
- A value of **70549.9** in the data represents **70,549,900** actual dollars (i.e., 70,549.9 thousand)
- In the UI summary panel, this displays as **"70.5 Mil"** (70.5 million)
- The `formatSummaryNumber()` function in `page.tsx` handles this conversion

**Conversion Logic:**
- Raw data value: `70549.9` (in '000)
- This represents: `70549.9 × 1,000 = 70,549,900` actual dollars
- To display as millions: `70549.9 ÷ 1,000 = 70.5 Mil`

**⚠️ BUG FIX (2026-03-01): Summary Dashboard Display**

**Problem:** The summary dashboard was showing `70549.9 Mil` instead of `70.5 Mil`.

**Root Cause:** The `formatSummaryNumber()` function was NOT dividing by 1000 to convert from thousands to millions. It was displaying the raw value directly.

**Correct Implementation:**
```javascript
function formatSummaryNumber(value: number | string): string {
  // ... string handling code ...
  
  const absValue = Math.abs(value)
  const sign = value < 0 ? '-' : ''
  
  // Data is in thousands ('000), so we need to convert to display format:
  // - Values >= 1000000 (thousands) = 1 billion+ actual → display as "B" (divide by 1e6)
  // - Values >= 1000 (thousands) = 1 million+ actual → display as "Mil" (divide by 1000)
  // - Values >= 1 (thousands) = 1000+ actual → display as "K" (no division)
  // - Values < 1 (thousands) = less than 1000 actual → display as "K" (no division)
  // 
  // Example: 70549.9 (thousands) = 70,549,900 actual = 70.5 million
  // → 70549.9 >= 1000, so divide by 1000 = 70.5 Mil
  if (absValue >= 1e6) {
    // 1,000,000 thousand = 1 billion actual
    return `${sign}${(absValue / 1e6).toFixed(1)} B`
  } else if (absValue >= 1000) {
    // 1,000 thousand = 1 million actual
    return `${sign}${(absValue / 1000).toFixed(1)} Mil`
  }
  // Less than 1000 thousand = less than 1 million actual
  return `${sign}${absValue.toFixed(1)} K`
}
```

**Test Cases:**
| Input (thousands) | Actual Value | Output |
|-------------------|--------------|--------|
| 70549.9 | 70,549,900 | 70.5 Mil |
| -152926.9 | -152,926,900 | -152.9 Mil |
| 1500000 | 1,500,000,000 | 1.5 B |
| 500 | 500,000 | 500.0 K |

**Key Points:**
- `formatSummaryNumber()` receives values in **THOUSANDS** (e.g., 70549.9)
- It must **DIVIDE by 1000** to convert to millions for display
- **NEVER multiply values by 1000** anywhere in the chain
- The raw CSV values are preserved as-is (no transformation during data loading)

**Do NOT revert this formatting.** If values show as "70549.9 Mil" instead of "70.5 Mil", the `formatSummaryNumber()` function has been incorrectly modified. See the code comments in `page.tsx` for the correct implementation.

### Financial Types
- Tender A
- 1st Working Budget B
- Audit Report (WIP) J
- Projection as at I
- Cash Flow
- And others...

### Trades
- Gross Profit (Item 1.0-2.0) (Financial A/C)
- Gross Profit (Item 3.0-4.3)
- Income
- Original Contract Works
- And others...

### Sheets
- Financial Status
- Projection
- Committed Cost
- Accrual
- Cash Flow

---

## Query Functions

### Available Functions
```python
from financial_chatbot import initialize, get_projected_gross_profit, get_wip_gross_profit, get_cash_flow

# Initialize connection (call once at startup)
initialize()

# Query functions
get_projected_gross_profit()  # Projection sheet, Trade contains "Gross Profit"
get_wip_gross_profit()        # Financial Status, Financial_Type="Audit Report (WIP) J", Trade contains "Gross Profit"
get_cash_flow()               # Cash Flow sheet, Trade contains "Gross Profit"
```

### ⚠️ Important: Use Correct Queries
**OLD (DO NOT USE - returns 0):**
```python
query(Trade='Gross Profit (bf adj)')
```

**NEW (works correctly):**
```python
get_projected_gross_profit()
get_wip_gross_profit()
get_cash_flow()
```

---

## Running Locally

### Prerequisites
- Python 3.12 (recommended)
- Google Sheets API credentials
- Required packages: `streamlit`, `gspread`, `oauth2client`

### Setup
```bash
# Clone the repo
git clone https://github.com/yatbond/Financial-chatbot.git
cd Financial-chatbot

# Install dependencies
pip install -r requirements.txt

# Set up Google credentials (JSON key file)
# Place credentials.json in project root

# Run locally
streamlit run app.py
```

### Access
Local URL: `http://localhost:8501`

---

## Deploying to Streamlit Cloud

### Repository
- **GitHub Repo:** `yatbond/Financial-chatbot`
- **Streamlit App:** https://share.streamlit.io

### Deployment Steps
1. Push code to `yatbond/Financial-chatbot` on GitHub
2. Go to https://share.streamlit.io/apps
3. Select the app → Settings
4. Ensure Repository is set to `yatbond/Financial-chatbot`
5. Save and redeploy

### ⚠️ Critical: Repository Setting
**Common Issue:** Streamlit Cloud may point to an OLD/incorrect repo.

**Fix:**
1. Go to https://share.streamlit.io/apps
2. Click your app → Settings
3. Check "Repository" setting
4. Change to: `yatbond/Financial-chatbot`
5. Save

**Error Symptom:** `AttributeError: st.session_state has no attribute "current_month"`

---

## Troubleshooting

### "returns 0" or No Data
- **Cause:** Using old query syntax
- **Fix:** Use `get_projected_gross_profit()`, `get_wip_gross_profit()`, or `get_cash_flow()`

### Connection Errors
- Check Google Sheets API credentials
- Verify `credentials.json` is valid
- Ensure sheet sharing is enabled

### Streamlit Cloud Errors
- Verify repository is `yatbond/Financial-chatbot`
- Check requirements.txt has all dependencies
- Review Streamlit logs in the dashboard

---

## File Structure
```
Financial-chatbot/
├── app.py                  # Main Streamlit app
├── financial_chatbot.py    # Query functions & logic
├── requirements.txt        # Python dependencies
├── credentials.json        # Google API credentials (not committed)
├── .gitignore
└── README.md
```

---

## Adding New Query Functions

To add a new query function in `financial_chatbot.py`:

```python
def get_new_query():
    """Description of what this queries"""
    return query(
        Sheet_Name='Sheet_Name',
        Financial_Type='Financial_Type',
        Trade='Trade contains "keyword"'
    )
```

Then import and use in `app.py`:
```python
from financial_chatbot import get_new_query
result = get_new_query()
```

---

## Compare Query Behavior

### ✅ Implementation (Fixed 2026-02-28)

Compare queries are handled by a dedicated `handleComparisonQuery()` function in `route.ts` (v3).

**How it works:**
1. `isComparisonQuery()` detects comparison keywords: "compare", "vs", "versus", "compared to", "with" (when combined with financial terms)
2. `extractComparisonParts()` splits the query into two sides using the comparison keyword as delimiter
3. `matchFinancialType()` resolves each side to an actual Financial_Type from the data
4. `extractComparisonMetric()` identifies the metric (Data_Type) being compared
5. Both Financial Types are queried from the **Financial Status** sheet
6. Results are formatted as a comparison table with difference calculation

**Supported query patterns:**
- `"compare projected gp with business plan gp"`
- `"projected gp vs budget gp"`
- `"wip gp versus projection gp"`
- `"compare cash flow with projection"` (works for any metric)
- `"projected np compared to business plan np"`

**This pattern applies to ALL "Compare" queries:**
- Always compares the same metric (e.g., Gross Profit, Revenue, Net Profit)
- From the same table (Financial Status)
- For the same Sheet_Name
- But across different Financial Types (e.g., Projection vs Business Plan)

**Output format:**
```
## Comparing: Projection as at I vs 1st Working Budget B
Table: Financial Status
Metric: Gross Profit (Item 1.0-2.0) (Financial A/C)

| Financial Type | Gross Profit (Item 1.0-2.0) (Financial A/C) |
|----------------|----------------------------------------------|
| Projection as at I     | $120,550     |
| 1st Working Budget B   | $70,550      |
| Difference             | ↑ $50,000 (+70.7%) |

*Values in ('000)*
```

**Fallback:** If comparison detection fails (can't resolve both Financial Types), the query falls through to the normal single-query handler.

---

## Total Query Behavior

### ✅ Implementation (Added 2026-02-28)

Total queries sum up all sub-items under a parent item for a specified Financial Type.

**How it works:**
1. `isTotalQuery()` detects "Total" keyword at the start of a query
2. `parseTotalQuery()` extracts:
   - Item name (e.g., "Preliminaries", "Materials")
   - Financial type (e.g., "Cashflow" → "Cash Flow", "Committed" → "Committed Cost")
   - Optional month and year
3. `handleTotalQuery()` orchestrates:
   - **No month specified:** Uses "Financial Status" sheet, sums all sub-items under parent item code
   - **Month specified:** Goes to worksheet matching the Financial Type name, filters by month, sums sub-items

**Item name mappings:**
| Keyword | Parent Code |
|---------|-------------|
| Preliminaries, prelim | 2.1 |
| Materials, material | 2.2 |
| Plant, machinery | 2.3 |
| Subcontractor, subcon, sub | 2.4 |
| Labour, labor | 2.5 |
| Reinforcement, rebar | 2.6 |
| Concrete | 2.7 |
| Steelworks | 2.8 |
| Electrical | 2.9 |
| Plumbing | 2.10 |
| Others | 2.11 |

**Financial type mappings:**
| Keyword | Financial Type |
|---------|----------------|
| cashflow, cash flow, cf | Cash Flow Actual received & paid as at |
| committed, committed cost, committed value | Committed Value / Cost as at |
| projection, projected | Projection as at |
| business plan, budget, bp | Business Plan |
| 1st working budget, first working budget | 1st Working Budget |
| revision, rev, budget revision | Revision as at |
| actual, actual cost | Actual Cost |
| accrual, accrued | Accrual |
| tender, budget tender | Budget Tender |
| wip, audit, audit report | Audit Report (WIP) |

**Supported query patterns:**
- `"Total Preliminaries Cashflow"` → Sums 2.1.x items for Cash Flow on Financial Status
- `"Total Committed Materials"` → Sums 2.2.x items for Committed Cost on Financial Status
- `"Total Cashflow Preliminaries"` → Order flexible (Financial Type can be before or after item)
- `"Total Cashflow Preliminaries Jan 2025"` → Sums 2.1.x on Cash Flow sheet for January 2025
- `"Total Plant Projection"` → Sums 2.3.x for Projection on Financial Status

**Output format:**
```
## Total: Preliminaries (Cash Flow)

**Parent Item:** 2.1 - Preliminaries
**Financial Type:** Cash Flow
**Sheet:** Financial Status

| Sub-Item | Value ('000) |
|----------|--------------|
| 2.1.1    | $X           |
| 2.1.2    | $Y           |
| ...      | ...          |

**Total:** $XXX
```

---

---

## Analyze Query Behavior

### ✅ Implementation (Added 2026-02-28)

The "Analyze" command runs a comprehensive financial analysis across 6 comparisons.

**Trigger:** User types "Analyze" or "Analyse"

**How it works:**
1. Always uses **Financial Status** sheet (cumulative values)
2. Auto-discovers Financial_Type names from the data (Projection, Business Plan, WIP, Committed, Budget Revision)
3. Gets all 2nd tier items under Income (1.x) and Costs (2.x)
4. Runs 6 comparisons:

| # | Comparison | Items | Condition |
|---|-----------|-------|-----------|
| 1 | Projection vs Business Plan | Income (1.x) | Projected < Business Plan |
| 2 | Projection vs Audit Report (WIP) | Income (1.x) | Projected < WIP |
| 3 | Projection vs Business Plan | Cost (2.x) | Projected > Business Plan |
| 4 | Projection vs Audit Report (WIP) | Cost (2.x) | Projected > WIP |
| 5 | Committed vs Projection | Cost (2.x) | Committed > Projected |
| 6 | Projection vs Budget Revision | Cost (2.x) | Projected > Budget Revision |

5. Results are cached in memory (30-minute TTL) for Detail drill-down

**Output format:**
```
## Financial Analysis

### Income Analysis (Item 1.x)
**1. Projection vs Business Plan - Income Shortfalls**
   1.1 Preliminaries: Projected $X < Business Plan $Y (↓$Z, -A%)
   ...

### Cost Analysis (Item 2.x)
**3. Projection vs Business Plan - Cost Overruns**
   3.1 Preliminaries: Projected $X > Business Plan $Y (↑$Z, +A%)
   ...
```

---

## Detail Query Behavior

### ✅ Implementation (Added 2026-02-28)

The "Detail" command drills down into specific analysis results from the Analyze command.

**Trigger:** User types "Detail X" or "Detail X.Y"

**Prerequisites:** Must run "Analyze" first (results cached for 30 minutes)

**"Detail X"** (e.g., "Detail 3"):
- Shows all 2nd tier items from comparison #3
- For each, lists 3rd tier items where the condition holds
- Output: hierarchical list with item codes and values

**"Detail X.Y"** (e.g., "Detail 3.1"):
- Shows 3rd tier items for a specific 2nd tier item
- Output: table format with values and differences
- Includes total overrun/shortfall

**Output format (Detail X):**
```
## Detail 3: Projection vs Business Plan - Cost Overruns

### 3.1 Preliminaries (2.1)
   - 2.1.1 Site Management: Projected $X > Business Plan $Y (↑$Z, +A%)
   - 2.1.2 Site Labour: Projected $X > Business Plan $Y (↑$Z, +A%)
```

**Output format (Detail X.Y):**
```
## Detail 3.1: Preliminaries - Projection vs Business Plan

**Parent:** 2.1 - Preliminaries

| 3rd Tier Item | Projected | Business Plan | Difference |
|---------------|-----------|---------------|------------|
| 2.1.1 Site Management | $X | $Y | ↑$Z (+A%) |

**Total Overrun:** $XXX
```

---

## Query Priority Order

1. **Analyze** → Full financial analysis (6 comparisons)
2. **Detail** → Drill-down into analysis results
3. **Total** → Sum sub-items under a parent
4. **Compare** → Compare two Financial Types
5. **Normal** → Standard query with fuzzy matching

---

## Shortcuts & Commands Reference

### Financial Type Shortcuts

These keywords are resolved to the actual `Financial_Type` names in the data.

| Shortcut / Keyword | Actual Financial_Type |
|--------------------|-----------------------|
| `bp`, `budget`, `business plan` | **Business Plan** |
| `revision`, `rev`, `budget revision`, `revision as at` | **Revision as at** |
| `1st working budget`, `first working budget` | **1st Working Budget** |
| `tender`, `budget tender` | **Budget Tender** |
| `projection`, `projected` | **Projection as at** |
| `wip`, `audit report`, `audit` | **Audit Report (WIP)** |
| `committed`, `committed value`, `committed cost` | **Committed Value / Cost as at** |
| `accrual`, `accrued` | **Accrual** |
| `cash flow`, `cashflow`, `cf` | **Cash Flow Actual received & paid as at** |
| `adjustment`, `variation` | **Adjustment Cost/ variation** |
| `balance` | **Balance** |
| `balance to` | **Balance to** |
| `general` | **General** |

> ⚠️ **Important:** "budget" and "bp" map to **Business Plan**, NOT "1st Working Budget". To query the 1st Working Budget specifically, use "1st working budget" or "first working budget".

### Data Type Shortcuts (ACRONYM_MAP)

| Shortcut | Expands To |
|----------|-----------|
| `gp` | gross profit |
| `np` | net profit |
| `subcon`, `sub` | subcontractor |
| `rebar` | reinforcement |
| `staff` | manpower (mgt. & supervision) |
| `labour`, `labor` | manpower (labour) |
| `prelim`, `preliminary` | preliminaries |
| `material` | materials |
| `plant`, `machinery` | plant and machinery |
| `profit`, `income`, `revenue` | gross profit |
| `loss` | net loss |

### Command Reference

#### Trend
Show a metric's trend over multiple months.

**Syntax:** `trend [metric] [months]`

**Examples:**
- `trend gp 6` — Gross Profit trend over 6 months
- `trend np 3` — Net Profit trend over 3 months

#### Compare
Compare the same metric across two different Financial Types.

**Syntax:**
- `compare [type1] [metric] with [type2] [metric]`
- `[type1] [metric] vs [type2] [metric]`

**Examples:**
- `compare projected gp with business plan gp` — Compare Projection GP vs Business Plan GP
- `projected gp vs wip gp` — Same but shorthand
- `budget np vs revision np` — Business Plan NP vs Revision as at NP
- `committed vs projection` — Compare committed cost vs projection

#### Total
Sum all sub-items under a parent item for a given Financial Type.

**Syntax:**
- `Total [Item] [Financial Type]`
- `Total [Financial Type] [Item]`
- `Total [Item] [Financial Type] [Month] [Year]` (with date)

**Examples:**
- `Total Preliminaries Cashflow` — Sum 2.1.x items for Cash Flow
- `Total Materials Committed` — Sum 2.2.x items for Committed Cost
- `Total Plant Projection Jan 2025` — Sum 2.3.x items for Projection in January 2025
- `Total Cashflow Preliminaries` — Order is flexible

#### Analyze
Run a comprehensive financial analysis across 6 comparisons.

**Syntax:** `Analyze` or `Analyse`

**What it does:**
1. Uses Financial Status sheet (cumulative values)
2. Compares Income items (1.x) and Cost items (2.x)
3. Runs 6 comparisons:
   - #1: Projection vs Business Plan — Income Shortfalls (Projected < BP)
   - #2: Projection vs WIP — Income Shortfalls (Projected < WIP)
   - #3: Projection vs Business Plan — Cost Overruns (Projected > BP)
   - #4: Projection vs WIP — Cost Overruns (Projected > WIP)
   - #5: Committed vs Projection — Committed Exceeds Projection
   - #6: Projection vs Revision — Exceeds Budget Revision
4. Caches results for 30 minutes for Detail drill-down

#### Detail
Drill down into analysis results from the Analyze command.

**Prerequisites:** Must run `Analyze` first.

**Syntax:**
- `Detail X` — Show all 3rd tier items for comparison #X
- `Detail X.Y` — Show 3rd tier items for specific 2nd tier item #Y within comparison #X

**Examples:**
- `Detail 3` — Drill into Comparison #3 (Cost Overruns: Projection vs Business Plan)
- `Detail 3.1` — Drill into the first flagged item in Comparison #3
- `Detail 5` — Drill into Comparison #5 (Committed vs Projection)

### Usage Examples

| Query | What It Does |
|-------|-------------|
| `projected gp` | Show Projected Gross Profit from Financial Status |
| `business plan np` | Show Business Plan Net Profit |
| `budget gp` | Same as "business plan gp" |
| `revision gp` | Show Revision as at Gross Profit |
| `wip gp` | Show Audit Report (WIP) Gross Profit |
| `compare projected gp with budget gp` | Compare Projection vs Business Plan GP |
| `Total Preliminaries Cashflow` | Sum all preliminary sub-items for Cash Flow |
| `Analyze` | Run full 6-comparison financial analysis |
| `Detail 3` | Drill into cost overrun details |
| `gp jan 2025` | Gross Profit for January 2025 |
| `cf` | Cash Flow data |

### Available Financial Types in Data

For reference, these are the actual `Financial_Type` values present in the data:

1. General
2. Budget Tender
3. 1st Working Budget
4. Adjustment Cost/ variation
5. Revision as at
6. Business Plan
7. Audit Report (WIP)
8. Adjustment Cost / Variation k=I-J
9. Projection as at
10. Committed Value / Cost as at
11. Balance
12. Accrual
13. Cash Flow Actual received & paid as at
14. G1=G/E
15. Balance to

---

## Financial Data Preprocessor

### Overview

The `financial_preprocessor.py` script parses Excel financial reports and converts them to flat CSV format for the chatbot to query.

**Location:** `G:/My Drive/Ai Chatbot Knowledge Base/financial_preprocessor.py`

**Run Command:**
```cmd
cd "G:\My Drive\Ai Chatbot Knowledge Base"
run_preprocessor.bat
```

### CSV Output Structure

| Column | Description |
|--------|-------------|
| Year | Calendar year (e.g., 2025, 2026) |
| Month | Month number (1-12) |
| Sheet_Name | Excel sheet name (Financial Status, Projection, Cash Flow, etc.) |
| Financial_Type | Type of financial data (Business Plan, Projection, etc.) |
| Item_Code | Line item code (1, 2, 2.1, 2.1.1, etc.) |
| Data_Type | Type of data (field name for General rows) |
| Value | The value (number or string) |

---

### ⚠️ BUG FIX: Year Assignment Logic (2026-03-01)

**Problem:** The year logic was inconsistent across sheets:
- Financial Status: Month 1 (Jan) → 2026 ✓ CORRECT
- Other sheets (Projection, Cash Flow, etc.): Month 1 (Jan) → 2025 ❌ WRONG

**Root Cause:** 
1. The comparison operator was `<` instead of `<=`
2. Year fix was only applied to monthly sheets, not all sheets

**Correct Logic:**
```python
if data_month <= report_month:
    year = report_year      # Same year (up to and including report month)
else:
    year = report_year - 1  # Previous year (months after report month)
```

**Examples:**

| Report Date | Data Month | Old Logic | New Logic | Correct? |
|-------------|------------|-----------|-----------|----------|
| Jan 2026 (1) | Jan (1) | 2025 ❌ | 2026 ✅ | ✓ |
| Jan 2026 (1) | Feb (2) | 2025 | 2025 | ✓ |
| Jan 2026 (1) | Dec (12) | 2025 | 2025 | ✓ |
| Feb 2026 (2) | Jan (1) | 2026 | 2026 | ✓ |
| Feb 2026 (2) | Feb (2) | 2025 ❌ | 2026 ✅ | ✓ |
| Feb 2026 (2) | Mar (3) | 2025 | 2025 | ✓ |

**Key Changes:**
1. Changed `<` to `<=` in `fix_year_assignment()` function
2. Applied year logic to ALL sheets (not just monthly sheets)
3. Updated all docstrings and comments

**File:** `/mnt/g/My Drive/Ai Chatbot Knowledge Base/financial_preprocessor.py`
**Function:** `fix_year_assignment(df)` (line ~491)

---

### ⚠️ BUG FIX: General Entries (Project Info) (2026-03-01)

**Problem:** Time Consumed (%) and Target Completed (%) were missing from the CSV output.

**Root Cause:** 
- These values were stored as Excel formulas in the spreadsheet
- When `openpyxl` opened the file with `data_only=True`, formula cells returned `None`
- The parser skipped rows with `None` values

**Fix:** Calculate percentages from the dates instead of reading formula cells.

**8 General Entries Required:**
1. Project Code
2. Project Name
3. Report Date
4. Start Date
5. Complete Date
6. Target Complete Date
7. **Time Consumed (%)** - CALCULATED
8. **Target Completed (%)** - CALCULATED

**Calculation Formulas:**

```python
# Time Consumed (%)
Time Consumed (%) = (Report Date - Start Date) / (Complete Date - Start Date) × 100

# Target Completed (%)
if Target Complete Date is missing or "Nil":
    Target Completed (%) = "Nil"
else:
    Target Completed (%) = (Report Date - Start Date) / (Target Complete Date - Start Date) × 100
```

**Edge Cases:**
- Missing Target Complete Date → "Nil"
- Division by zero (Complete Date = Start Date) → "N/A"

**CSV Row Format for General Entries:**
```
Year,Month,Sheet_Name,Financial_Type,Item_Code,Data_Type,Value
2026,1,Financial Status,General,Project Code,Project Code,1014
2026,1,Financial Status,General,Project Name,Project Name,PolyU
2026,1,Financial Status,General,Time Consumed (%),Time Consumed (%),214.24
2026,1,Financial Status,General,Target Completed (%),Target Completed (%),92.77
```

**Key Points:**
- All General entries have `Sheet_Name = "Financial Status"`
- All General entries have `Financial_Type = "General"`
- For General entries: `Data_Type = Item_Code` (both are the field name)

**File:** `/mnt/g/My Drive/Ai Chatbot Knowledge Base/financial_preprocessor.py`
**Function:** `_extract_metadata(ws, report_year, report_month)` (line ~38)

---

### Query Scoring Fix (2026-03-01)

**Problem:** Query "projected gp" returned "Revision as at" instead of "Projection as at"

**Root Cause:**
- Common words "as", "at" were boosting scores incorrectly
- "Revision as at" matched these common words (not in user query)
- "Projection as at" matched the keyword "projection" (from user's "projected")

**Fix:**
1. Added **STOPWORDS** set to filter common words
2. Added **FINANCIAL_TYPE_KEYWORDS** map for primary keyword matching
3. **10x weight** for Financial Type keyword matches (50 pts vs 5 pts)
4. **100 pts** for exact Financial Type match

**Stopwords List:**
```python
STOPWORDS = new Set([
  'as', 'at', 'the', 'and', 'or', 'for', 'in', 'on', 'to', 'of', 
  'a', 'an', 'is', 'it', 'by', 'from', 'with', 'that', 'this'
])
```

**Financial Type Keywords:**
```python
FINANCIAL_TYPE_KEYWORDS = {
  'projection': ['projection', 'projected'],
  'revision': ['revision', 'rev'],
  'business plan': ['business plan', 'budget', 'bp'],
  'audit report': ['wip', 'audit', 'audit report'],
  'committed': ['committed'],
  'cash flow': ['cash flow', 'cashflow', 'cf'],
  'tender': ['tender'],
  'accrual': ['accrual', 'accrued'],
  '1st working budget': ['1st working budget', 'first working', 'working budget']
}
```

**File:** `app/api/chat/route.ts`
**Lines:** 155-179 (STOPWORDS and FINANCIAL_TYPE_KEYWORDS definitions)

---

### Summary Dashboard Fix (2026-03-01)

**Problem:** Summary dashboard showing `70549.9 Mil` instead of `70.5 Mil`

**Root Cause:** `formatSummaryNumber()` was not dividing by 1000 to convert from thousands to millions.

**Fix:** See "Data Values are in Thousands ('000)" section above for details.

**File:** `app/page.tsx`
**Function:** `formatSummaryNumber(value)` (line ~8)

---

---

### Parent Value Calculation (Added 2026-03-01)

**Feature:** Automatically calculates missing parent item values by summing their direct children's values, using bottom-up recursive aggregation.

**Problem:** The Excel financial reports contain hierarchical item codes (e.g., `2.1.1`, `2.1.2`, etc.) but the parent items (e.g., `2.1`, `2`, `1`) often don't have values in the spreadsheet — they're either missing entirely or have formulas that `openpyxl` can't read. This meant users couldn't query "Total Preliminaries" or "Total Cost" directly.

**How Parent-Child Relationships Work:**

Item codes follow a dot-separated hierarchy:
```
1           → Top-level (Income)
├── 1.1     → 2nd tier (Original Contract Works)
├── 1.2     → 2nd tier (V.O. / Compensation Events)
│   ├── 1.2.1  → 3rd tier (V.O. / C.E.)
│   ├── 1.2.2  → 3rd tier (Fee of C.E.)
│   └── 1.2.3  → 3rd tier (Stretch Target)
├── 1.3     → 2nd tier (Provisional Sum)
...
2           → Top-level (Less : Cost)
├── 2.1     → 2nd tier (Preliminaries)
│   ├── 2.1.1  → 3rd tier (Manpower - Mgt. & Supervision)
│   ├── 2.1.2  → 3rd tier (Manpower - RE)
│   ...
├── 2.2     → 2nd tier (Materials)
│   ├── 2.2.1  → 3rd tier (Concrete)
│   ...
```

- A parent has fewer dot-separated parts than its children
- `X.0` items (e.g., `3.0`, `5.0`, `7.0`) are treated as top-level, NOT as children of `X`
- Parent code is derived by removing the last dot-segment: `2.1.3` → parent `2.1` → parent `2`

**Calculation Algorithm:**

1. **Discover all parent-child relationships** from existing Item_Codes
2. **Recursively discover higher-level parents** (e.g., `2.1` is parent → `2` is grandparent)
3. **Sort parents by depth** (deepest first for bottom-up processing)
4. **Per (Sheet_Name, Financial_Type, Year, Month) group:**
   - Build a value map accumulating all numeric values per Item_Code
   - For each parent (deepest first):
     - Skip if parent already has a non-zero value in this group
     - Sum all direct children's numeric values
     - Store result for cascading to higher parents
     - Create new row or update existing empty row
5. **Append all new parent rows** to the DataFrame

**Example:**
```
Given data:
  2.1.1 (Manpower Mgt.)    = 52,051
  2.1.2 (Manpower RE)      = 0
  2.1.3 (Manpower Labour)  = 7,808
  2.1.4 (Insurance)        = 12,550
  ...
  2.1.14 (Levies)          = 2,500
  
Calculated:
  2.1 (Preliminaries)      = sum(2.1.1 + 2.1.2 + ... + 2.1.14) = 124,141.95
  
Then:
  2 (Less : Cost)           = sum(2.1 + 2.2 + 2.3 + ... + 2.14) = 594,775.70
```

**Edge Cases Handled:**
- **Duplicate Item_Codes:** Some codes appear twice with different Data_Types (e.g., `2.1.6` as "Less : Cost" and "Reconciliation"). Values are accumulated (summed) rather than overwritten.
- **Non-numeric values:** Strings like "N/A", "Nil" are skipped during summation.
- **Missing parent rows:** New rows are created with Data_Type derived from children (common prefix). E.g., children "Less : Cost - Preliminaries - -Manpower" → parent "Less : Cost - Preliminaries".
- **Per-group calculation:** Each (Sheet_Name, Financial_Type, Year, Month) combination is processed independently, so a parent may be calculated in one Financial_Type but not another.
- **Existing values preserved:** Parents with existing non-zero values are NOT overwritten.

**Generated Parent Items (typical):**

| Item_Code | Data_Type | Description |
|-----------|-----------|-------------|
| 1 | Income | Total Income |
| 1.2 | Income - V.O. / Compensation Events | V.O. subtotal |
| 1.12 | Income - Other Revenue / Pain Gain Sharing | Other revenue subtotal |
| 2 | Less : Cost | Total Cost |
| 2.1 | Less : Cost - Preliminaries | Preliminaries subtotal |
| 2.2 | Less : Cost - Materials | Materials subtotal |
| 2.3 | Less : Cost - Plant & Machinery | Plant & Machinery subtotal |
| 2.4 | Less : Cost - Subcontractor | Subcontractor subtotal |
| 2.6 | Less : Cost - Nominated Package | Nominated Package subtotal |
| 2.11 | Less : Cost - Incentive / POR Bonus | Incentive subtotal |
| 4 | Reconciliation | Reconciliation total |
| 6 | Overhead | Overhead total |
| 6.1 | Overhead - HO Overhead Rate % | HO Overhead subtotal |

**Integration:**
- Function: `calculate_parent_values(df)` in `financial_preprocessor.py`
- Called after `fix_year_assignment(df)` in `preprocess_folder()`
- Adds ~670 new rows per workbook (varies by file)

**Files Modified:**
- `financial_preprocessor.py` — Added `calculate_parent_values()`, `_get_parent_code()`, `_derive_parent_data_type()`

---

### Multi-Format Excel Parser (Added 2026-03-01)

**Feature:** The parser now supports **two different Excel layout formats** automatically.

**Problem:** The 966 NE201703 financial report had a completely different structure from PolyU files, causing only 2 rows to parse instead of thousands.

**Format Differences:**

| Aspect | PolyU Format | NE Format |
|--------|--------------|-----------|
| "Item" header column | Column A | Column B |
| Metadata label column | Column A | Column B |
| Metadata value column | Column B | Column E |
| Report Date label | `"Report Date"` | `"Financial Status as at"` |
| Complete Date label | `"Complete Date"` | `"Project Completion Date"` |
| Start Date label | `"Start Date"` | `"Project Start Date"` |
| Month column headers | Text ("Jan", "Feb", etc.) | DateTime objects (2018-08-01) |
| Sheet: Projection | `"Projection"` | `"Projected Cost"` |
| Sheet: Cash Flow | `"Cash Flow"` | `"Cashflow"` |

**Changes Made to `financial_preprocessor.py`:**

| Function | Change |
|----------|--------|
| `_find_header_row()` | Returns `(row, col)` tuple; scans columns A **and** B for "Item" |
| `_extract_metadata()` | Tries both (A,B) and (B,E) label/value layouts; adds alternative label mappings |
| `_get_report_date()` | Tries both column layouts; handles DateTime objects directly |
| `_parse_financial_status()` | Accepts `item_col` parameter; uses dynamic column positions |
| `_parse_monthly_sheet()` | Accepts `item_col` parameter; handles DateTime month headers; adds "projected cost", "cashflow", "budget" sheet matching |
| Sheet name matching | Added keywords: "projected cost", "cashflow", "budget", "actual rec" |

**Alternative Label Mappings:**

```python
meta_fields = {
    'Report Date': 'Report Date',
    'Financial Status as at': 'Report Date',  # NE format
    'Start Date': 'Start Date',
    'Project Start Date': 'Start Date',       # NE format
    'Complete Date': 'Complete Date',
    'Project Completion Date': 'Complete Date',  # NE format
    ...
}
```

**DateTime Month Headers:**

Some files use actual DateTime objects as column headers instead of text:
```python
# Old code only handled text
if header in month_map:
    month_cols[month_map[header]] = col

# New code handles DateTime too
if hasattr(val, 'month'):
    month_cols[val.month] = col
```

**Results:**

| File | Before Fix | After Fix |
|------|-----------|-----------|
| 966 NE201703 | 2 rows ❌ | 3,365 rows ✅ |
| 1014 PolyU | 5,777 rows | 5,777 rows ✅ (regression passed) |

**Sheets Now Parsing (966 NE201703):**
- Financial Status: 601 rows ✅
- Projected Cost: 674 rows ✅
- Committed Cost: 674 rows ✅
- Actual Rec'd & Cost: 696 rows ✅
- Actual Interest Working: 720 rows ✅

**Known Limitations:**
- **Budget** sheet: Uses Tender A/B summary columns instead of monthly data; not parsed
- **Cashflow** sheet: Uses IP/Month/Planned/Actual format instead of Item/Trade hierarchy; not parsed

---

*Last updated: 2026-03-01 (v6 - Multi-format parser, Parent value calculation, Year logic, General entries, Query scoring)*
