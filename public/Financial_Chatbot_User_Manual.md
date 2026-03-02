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
| **Check gross profit** | `[type] gp` | `projected gp`, `budget gp` |
| **Check net profit** | `[type] np` | `projected np`, `wip np` |
| **Check parent items** | `[parent item] [type]` | `Preliminaries Cashflow`, `Materials Committed` |
| **Compare two scenarios** | `compare [type1] gp with [type2] gp` | `compare projected gp with budget gp` |
| **See cash flow** | `cf` or `cashflow` | `cf`, `cash flow Jan 2025` |
| **Run full analysis** | `Analyze` | `Analyze` |
| **Drill into analysis** | `Detail X` | `Detail 3`, `Detail 3.1` |

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

### Parent Items (Query directly for totals)

| Shortcut | Full Name | Example |
|----------|-----------|---------|
| `prelim` / `preliminaries` | Preliminaries (2.1) | `prelim cashflow` |
| `materials` | Materials (2.2) | `materials committed` |
| `plant` / `machinery` | Plant/Machinery (2.3) | `plant projection` |
| `subcon` / `sub` | Subcontractor (2.4) | `subcon committed` |
| `labour` / `labor` | Labour (2.5) | `labour projection` |
| `rebar` / `reinforcement` | Reinforcement (2.6) | `rebar projection` |
| `concrete` | Concrete (2.7) | `concrete committed` |

**Note:** Parent values are automatically calculated as the sum of their children. Just query the parent item directly!

---

## 📊 Main Features

### 1. Simple Queries
Ask for any financial metric by name.

**Examples:**
- `projected gross profit` → Shows projected GP
- `business plan net profit` → Shows budget NP
- `wip gp` → Shows WIP Gross Profit
- `committed cost` → Shows committed values

**With dates:**
- `gp Jan 2025` → GP for January 2025
- `cash flow Feb 2025` → Cash flow for February

**Parent items (automatically sum children):**
- `Preliminaries Cashflow` → Total of all prelim sub-items
- `Materials Committed` → Total of all material sub-items
- `Subcon Projection Jan 2025` → Subcontractor total for January

---

### 2. Compare Two Scenarios
Compare the same metric across different financial types.

**Format:** `compare [type1] [metric] with [type2] [metric]`

**Examples:**
- `compare projected gp with budget gp`
- `wip gp vs projected gp`
- `committed vs projection`

**What you'll see:**
```
Comparing: Projection as at vs Business Plan

| Financial Type | Gross Profit |
|----------------|--------------|
| Projection     | $120,550     |
| Business Plan  | $70,550      |
| Difference     | ↑ $50,000 (+70.7%) |
```

---

### 3. Full Analysis (Most Powerful)
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

### 4. Detail Drill-Down
Dive deeper into analysis results.

**Prerequisites:** Run `Analyze` first

**Format:** `Detail X` or `Detail X.Y`

**Examples:**
- `Detail 3` → Show all items in comparison #3 (Cost Overruns)
- `Detail 3.1` → Show specific sub-items for item 3.1 (Preliminaries)
- `Detail 5` → Show committed costs exceeding projection

**What you'll see:**
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
- `Preliminaries Cashflow` - Breakdown of preliminaries (parent = sum of children)
- `gp [month] [year]` - Specific month performance

### When Costs Change
- `committed gp` - Check committed costs
- `[parent item] Committed` - View by category (e.g., `Materials Committed`)
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

### Parent Items (Sum of Children)
```
Preliminaries Cashflow          → Sum of prelim sub-items (cash flow)
Materials Committed             → Sum of materials (committed)
Subcon Projection               → Sum of subcontractor (projection)
Plant Committed Jan 2025        → Plant/machinery for January
```

### Comparisons
```
compare projected gp with budget gp      → Projection vs Business Plan
wip gp vs projected gp                   → WIP vs Projection
committed vs projection                  → Committed costs vs Projection
budget np vs revision np                 → Business Plan vs Revision
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
   - Data values are in thousands (e.g., 70.5 = 70,500)
   - Summary panel displays these as "70.5 Mil" (millions)
   - This is the correct display format - do not revert

4. **Analysis results cache for 30 minutes** - If data updates, run "Analyze" again

5. **Month names** can be full (January) or abbreviated (Jan)

6. **Parent items automatically sum their children** - No need for a separate "Total" command. Just query the parent directly (e.g., `Preliminaries Cashflow` instead of `Total Preliminaries Cashflow`)

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

### Detail command shows "Run Analyze first"
- Run `Analyze` command before using `Detail`
- Analysis results expire after 30 minutes

---

## 📞 Support

For questions or issues:
- Ask in this Telegram group
- Tag @MaryJane123bot for assistance

---

**Quick Start Checklist:**
✅ Try: `projected gp` - Check current GP
✅ Try: `Analyze` - Run full analysis
✅ Try: `Detail 3` - See cost overrun details
✅ Try: `compare projected gp with budget gp` - Compare scenarios
✅ Try: `Preliminaries Cashflow` - View parent item (sum of children)

---

*Last Updated: 2026-03-02*
*Version: 1.1*
