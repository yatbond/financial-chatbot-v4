# Financial Chatbot User Manual
## For Staff Use

---

## 📖 What is the Financial Chatbot?

The Financial Chatbot is an AI assistant that helps you query and analyze project financial data. Simply type your questions in plain English, and it will fetch the data from our financial records.

**Access:** Available 24/7 in this Telegram group

---

## 🚀 Getting Started

### Basic Query Format
Just type your question naturally:
- `projected gross profit`
- `cash flow for January 2025`
- `compare budget gp with actual gp`

### Quick Reference - Most Common Commands

| What You Want | Command | Example |
|--------------|---------|---------|
| **List all headings** | `list` | `list` |
| **See all sub-items** | `list all` or `more` | `list all`, `more` (after `list`) |
| **Check gross profit** | `[type] gp` | `projected gp`, `budget gp` |
| **Check net profit** | `[type] np` | `projected np`, `wip np` |
| **Query by item code** | `[type] [code]` | `projected 2.1`, `cashflow 2.4` |
| **Trend by item code** | `trend [type] [code]` | `trend cashflow 2.1` |
| **Compare by item code** | `compare [type1] [code] vs [type2] [code]` | `compare committed 2.1 vs cashflow 2.1` |
| **Compare two scenarios** | `compare [type1] gp with [type2] gp` | `compare projected gp with budget gp` |
| **See cash flow** | `cf` or `cashflow` | `cf`, `cash flow Jan 2025` |
| **Run full analysis** | `Analyze` | `Analyze` |
| **Drill into analysis** | `Detail X` | `Detail 3`, `Detail 3.1` |
| **Drill into Compare/Trend** | `detail` | After compare/trend query, type `detail` |
| **Total sub-items** | `Total [item] [type]` | `Total Preliminaries Cashflow` |

---

## 💡 Key Shortcuts

### Financial Types (Use these keywords)

| Shortcut | Full Name | Example |
|----------|-----------|---------|
| `projected` / `projection` | Projection as at | `projected gp` |
| `budget` / `bp` | Business Plan | `budget gp` |
| `wip` / `audit` | Audit Report (WIP) | `wip gp` |
| `revision` / `rev` | Revision as at | `revision gp` |
| `committed` | Committed Value/Cost | `committed gp` |
| `cashflow` / `cf` | Cash Flow | `cf Jan 2025` |

### Data Types (Metrics)

| Shortcut | Full Name | Item Code | Example |
|----------|-----------|-----------|---------|
| `gp` | Gross Profit | 3 | `projected gp` |
| `np` / `net profit` | Acc. Net Profit/(Loss) | 7 | `budget np` |
| `cost` / `total cost` | Less : Cost | 2 | `Total cost cashflow` |
| `prelim` / `preliminary` / `preliminaries` / `total prelim` | Preliminaries | 2.1 | `Total prelim cashflow` |
| `supervision` / `staff` | Manpower (Mgt. & Supervision) | 2.1.1 | `trend cashflow supervision` |
| `material` / `materials` / `material cost` | Materials | 2.2 | `Total material projection` |
| `plant` / `all plant` / `machinery` | Plant & Machinery | 2.3 | `Total plant committed` |
| `subcon` / `sub` / `subbie` / `subcontractor` / `subcontractors` | Less : Cost - Subcontractor | 2.4 | `Total subcon committed` |
| `contract works` / `subbie contract works` / `subcon contract works` / `subcontractors contract works` | -Contract Works | 2.4.1 | `subbie contract works` |
| `vo` / `variation` / `variations` / `subbie vo` / `subcon vo` / `subcontractors vo` | -Variation | 2.4.2 | `subcon variation` |
| `claim` / `claims` / `subbie claim` / `subcon claim` / `subcontractors claim` | -Claim | 2.4.3 | `subcon claim` |
| `labour` / `labor` | Manpower (Labour) | 2.5 | `Total labour committed` |
| `rebar` | Reinforcement | 2.6 | `Total rebar projection` |

---

## 📊 Main Features

### 1. List Financial Headings
See all available financial headings and item codes.

**Format:** `list` or `list all`

**Examples:**
- `list` → Show top 2 tiers (summary view)
- `list all` → Show all tiers including sub-items
- `more` → After `list`, show all sub-items

**What you'll see:**
```
## Financial Headings

### Income (Item 1)
  1.1 - Income - Original Contract Works
  1.2 - Income - V.O. / Compensation Events
  1.3 - Income - Provisional Sum

### Less : Cost (Item 2)
  2.1 - Less : Cost - Preliminaries
  2.2 - Less : Cost - Materials
  2.3 - Less : Cost - Plant & Machinery
  2.4 - Less : Cost - Subcontractor

### Gross Profit (Item 3)
  ...

💡 Type 'list all' or 'more' to see all sub-items (3rd and 4th tier)
```

---

### 2. Query by Item Code
Use item codes directly in any command for precise queries.

**Format:** `[command] [item_code]`

**Item code patterns:**
- Single level: `1`, `2`, `3` (top-level items)
- Two levels: `2.1`, `2.2`, `2.4` (categories)
- Three levels: `2.1.1`, `2.4.1`, `2.4.2` (sub-categories)

**Examples:**

| Command | What it does |
|---------|--------------|
| `projected 2.1` | Projected value for Item 2.1 (Preliminaries) |
| `cashflow 2.4` | Cash flow for Item 2.4 (Subcontractor) |
| `trend cashflow 2.1` | Trend for Item 2.1 on Cash Flow sheet |
| `trend projection 2.4.1` | Trend for Item 2.4.1 (Contract Works) |
| `compare committed 2.1 vs cashflow 2.1` | Compare Item 2.1 across sheets |
| `compare committed 2.4 vs cashflow subbie` | Mix item code with acronym |

**Benefits:**
- More precise than keyword matching
- Works with any command (trend, compare, normal query)
- Can combine item codes with acronyms in comparisons

---

### 3. Simple Queries
Ask for any financial metric by name.

**Examples:**
- `projected gross profit` → Shows projected GP
- `business plan net profit` → Shows budget NP
- `wip gp` → Shows WIP Gross Profit
- `committed cost` → Shows committed values

**With dates:**
- `gp Jan 2025` → GP for January 2025
- `cash flow Feb 2025` → Cash flow for February

---

### 4. Compare Two Scenarios
Compare the same metric across different financial types.

**Format:** `compare [type1] [metric] with [type2] [metric]`

**Examples:**
- `compare projected gp with budget gp`
- `wip gp vs projected gp`
- `committed vs projection`
- `compare committed income oct 2025 vs jan 2026` (simplified - auto-infers "committed income")

**What you'll see:**
```
Comparing: Projection as at vs Business Plan

| Financial Type | Gross Profit |
|----------------|--------------|
| Projection     | $120,550     |
| Business Plan  | $70,550      |
| Difference     | ↑ $50,000 (+70.7%) |

💡 Type 'detail' to compare sub-items
```

**Simplified Syntax:**
- `compare committed income oct 2025 vs jan 2026`
- `compare cashflow subcontractor 10 2025 vs 1 2026` (numeric months also work)
- Chatbot infers that "jan 2026" or "1 2026" means "committed income jan 2026"
- Only need to specify the changed part (date)
- Both month names (jan, oct) and month numbers (1, 10) are supported

**Drill-down with Detail:**
- Type `detail` after a comparison to see sub-item comparisons
- Each sub-item shown with numbered table
- Type `more` for next 20 sub-items

---

### 5. Trend Analysis with Detail Drill-Down
Show a metric's trend over time, then drill down into sub-items.

**Format:** `Trend [Financial Type] [Metric] [N months]`

**Examples:**
- `trend cashflow cost` → Cost trend for last 6 months
- `trend accrual prelim 8` → Preliminaries trend for 8 months
- `trend projection gp 12` → GP trend for 12 months

**Drill-down:**
- Type `detail` → Shows children items with trend tables (numbered 1-20)
- Type `detail 3` → Drills down into item 3's children
- Type `more` → Shows next 20 items

**What you'll see:**
```
## Trend: Less : Cost (Item 2) (Last 6 Months)

| Month | Value ('000) |
|-------|-------------|
| Aug 2025 | $49,031 |
| Sep 2025 | $52,143 |
...

💡 Type 'detail' to see sub-items

## Detail: Less : Cost (Item 2) - Sub-Items

### [1] Item 2.1 - Preliminaries
| Month | Value ('000) |
|-------|-------------|
| Aug 2025 | $9,341 |
...

### [2] Item 2.2 - Materials
...

💡 Type 'detail N' to drill down into sub-item N
*Showing 1-20 of 35 sub-items*
```

---

### 6. Total Sub-Items
Sum up all sub-items under a parent category.

**Format:** `Total [Item] [Financial Type]`

**Common parent items:**
- **Preliminaries** (2.1) - Site management, site labour, etc.
- **Materials** (2.2)
- **Plant/Machinery** (2.3)
- **Subcontractor** (2.4)
- **Labour** (2.5)
- **Reinforcement** (2.6)
- **Concrete** (2.7)

**Examples:**
- `Total Preliminaries Cashflow` → Sum all prelim sub-items for cash flow
- `Total Materials Committed` → Sum material costs in committed
- `Total Subcon Projection Jan 2025` → Subcontractor totals for January

**What you'll see:**
```
Total: Preliminaries (Cash Flow)

| Sub-Item | Value ('000) |
|----------|--------------|
| 2.1.1    | $X           |
| 2.1.2    | $Y           |
| ...      | ...          |
```

### 7. Query Specific Item Code
Query a specific item by its code. The system now applies an **exact filter** to return only the requested item.

**Format:** `[Financial Type] item_code [X.X]` or `[Financial Type] item [X.X]` or `[Financial Type] [X.X]`

**Examples:**
- `Projected item_code 2.1` → Returns Item 2.1 (Preliminaries) value: $723,786
- `Budget item 1.1` → Returns Item 1.1 under Budget
- `item_code 2` → Returns Item 2 (Less: Cost)
- `Projected 2.1` → Returns Item 2.1 (Preliminaries)

**What you'll see:**
```
## Query Results

Filters:
• Sheet: Financial Status
• Financial Type: Projection as at
• Month: All
• Year: All
• Data Type: All
• Item Code: 2.1

Total: $723,786 [Financial Status, Projection as at, Less : Cost - Preliminaries, 2.1, 1, 2026] ('000)
```

**Note:** The system now filters by exact Item_Code match, so you get the exact item you requested, not a list of sub-items.

Total: $XXX
```

---

### 8. Full Analysis (Most Powerful)
Run a comprehensive financial health check.

**Command:** `Analyze` or `Analyse`

**What it does:**
- Compares 6 key scenarios automatically
- Flags income shortfalls and cost overruns
- Shows where projected differs from budget/WIP
- Identifies committed costs exceeding projections

**The 6 comparisons:**
1. **Income Shortfalls** - Projection vs Business Plan
2. **Income Shortfalls** - Projection vs WIP
3. **Cost Overruns** - Projection vs Business Plan
4. **Cost Overruns** - Projection vs WIP
5. **Cost Alerts** - Committed vs Projection
6. **Budget Alerts** - Projection vs Revision

**What you'll see:**
```
## Financial Analysis

### Income Analysis
**1. Projection vs Business Plan - Income Shortfalls**
   1.1 Preliminaries: Projected $X < Business Plan $Y (↓$Z, -A%)

### Cost Analysis
**3. Projection vs Business Plan - Cost Overruns**
   3.1 Preliminaries: Projected $X > Business Plan $Y (↑$Z, +A%)
```

---

### 9. Detail Drill-Down
Dive deeper into analysis results or comparison queries.

**Prerequisites:** Run `Analyze`, `Compare`, or `Trend` first

**Format:** `detail` or `Detail X` or `Detail X.Y`

**After Compare queries:**
- Type `detail` after a comparison to see sub-item comparisons
- Works for both financial type comparisons and date comparisons
- Example: After `compare cash flow oct 2025 vs jan 2026`, type `detail` to see all sub-items with Oct vs Jan values

**After Trend queries:**
- Type `detail` after a trend query to see children items with trend tables
- Type `detail N` to drill down into item N's children

**After Analyze:**
- `Detail 3` → Show all items in comparison #3 (Cost Overruns)
- `Detail 3.1` → Show specific sub-items for item 3.1 (Preliminaries)
- `Detail 5` → Show committed costs exceeding projection

**What you'll see (after Compare):**
```
## Detail: Comparing Sub-Items

Comparing: Cash Flow (Oct 2025) vs Cash Flow (Jan 2026)
Metric: Subcontractor

### [1] Item 2.4.1 - Contract Works
| Type | Value ('000) |
|------|-------------|
| Cash Flow (Oct 2025) | $X |
| Cash Flow (Jan 2026) | $Y |
| Diff | ↑$Z (+A%) |

### [2] Item 2.4.2 - Variation
...
```

**What you'll see (after Analyze):**
```
## Detail 3.1: Preliminaries - Projection vs Business Plan

| 3rd Tier Item | Projected | Business Plan | Difference |
|---------------|-----------|---------------|------------|
| 2.1.1 Site Mgmt | $X | $Y | ↑$Z (+A%) |

Total Overrun: $XXX
```

---

## 🎯 Common Use Cases

### Daily Checks
- `projected gp` - Check current projected gross profit
- `cf` - Quick cash flow status
- `wip gp` - Compare with audit report

### Monthly Reviews
- `Analyze` - Run full analysis
- `Detail 3` - Check cost overruns
- `Detail 5` - Review committed costs

### Before Client Meetings
- `compare projected gp with budget gp` - See how you're tracking vs budget
- `Total Preliminaries Cashflow` - Breakdown of preliminaries
- `gp [month] [year]` - Specific month performance

### When Costs Change
- `committed gp` - Check committed costs
- `Total [item] Committed` - Breakdown by category
- `compare committed vs projection` - Are you over-committed?

---

## 📝 Query Examples

### Basic Queries
```
projected gp                    → Show projected gross profit
budget np                       → Show business plan net profit
wip gp                          → Show WIP gross profit
cf                              → Show cash flow data
```

### With Dates
```
gp Jan 2025                     → GP for January 2025
cash flow Feb 2025              → Cash flow for February
projected gp Mar 2025           → Projected GP for March
```

### Comparisons
```
compare projected gp with budget gp      → Projection vs Business Plan
wip gp vs projected gp                   → WIP vs Projection
committed vs projection                  → Committed costs vs Projection
budget np vs revision np                 → Business Plan vs Revision
```

### Totals
```
Total Preliminaries Cashflow             → Sum prelim sub-items (cash flow)
Total Materials Committed                → Sum materials (committed)
Total Subcon Projection                  → Sum subcontractor (projection)
Total Plant Committed Jan 2025           → Plant/machinery for January
```

### Analysis
```
Analyze                                  → Run 6-comparison analysis
Detail 3                                 → Drill into cost overruns
Detail 3.1                               → Specific prelim overrun details
Detail 5                                 → Committed vs projection details
```

---

## ⚠️ Important Notes

1. **"budget" = Business Plan** - Not 1st Working Budget
   - Use `budget gp` or `bp gp` for Business Plan
   - Use `1st working budget gp` for that specific budget

2. **Run "Analyze" first** before using "Detail" commands

3. **Values are in thousands ('000)** unless specified otherwise

4. **Analysis results cache for 30 minutes** - If data updates, run "Analyze" again

5. **Month names** can be full (January) or abbreviated (Jan)

---

## 🔧 Troubleshooting

### "No data found" or returns 0
- Check spelling of financial type or metric
- Try using the full name instead of shortcut
- Verify the month/year format (e.g., "Jan 2025")

### Comparison not working
- Ensure both financial types exist in data
- Use format: `compare [type1] gp with [type2] gp`
- Try simpler version: `[type1] gp vs [type2] gp`

### Detail command shows "No previous query found"
- Run a query first: `Analyze`, `Compare`, or `Trend`
- Type `detail` (not `detail N`) for Compare/Trend drill-down
- Use `Detail N` format only after `Analyze`
- Context expires after some time - re-run your query

---

## 📞 Support

For questions or issues:
- Ask in this Telegram group
- Tag @MaryJane123bot for assistance

---

**Quick Start Checklist:**
✅ Try: `list` - See all financial headings
✅ Try: `projected gp` - Check current GP
✅ Try: `projected 2.1` - Query by item code
✅ Try: `Analyze` - Run full analysis
✅ Try: `Detail 3` - See cost overrun details
✅ Try: `compare projected gp with budget gp` - Compare scenarios
✅ Try: `detail` after Compare/Trend - See sub-item comparisons

---

*Last Updated: 2026-03-03*
*Version: 1.2*
