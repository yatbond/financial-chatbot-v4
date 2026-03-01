import { NextRequest, NextResponse } from 'next/server'
import { google } from 'googleapis'

/**
 * ============================================================
 * FINANCIAL CHATBOT - FUNCTION GUIDE
 * ============================================================
 *
 * AVAILABLE COMMANDS (processed in priority order):
 *
 * 1. ANALYZE
 *    - Trigger: "Analyze" or "Analyse"
 *    - Purpose: Run comprehensive financial analysis (6 comparisons)
 *    - Uses Financial Status sheet (cumulative values)
 *    - Compares Income (1.x) and Cost (2.x) items across Financial Types
 *    - Comparisons:
 *      #1: Projection vs Business Plan — Income Shortfalls
 *      #2: Projection vs WIP — Income Shortfalls
 *      #3: Projection vs Business Plan — Cost Overruns
 *      #4: Projection vs WIP — Cost Overruns
 *      #5: Committed vs Projection — Committed Exceeds Projection
 *      #6: Projection vs Revision — Exceeds Budget Revision
 *    - Results are cached (30 min TTL) for Detail drill-down
 *    - Example: "Analyze"
 *
 * 2. DETAIL
 *    - Trigger: "Detail X" or "Detail X.Y"
 *    - Purpose: Drill down into Analyze results
 *    - Prerequisite: Must run "Analyze" first (cached 30 min)
 *    - Examples:
 *      - "Detail 3"   → Show all 3rd tier items for comparison #3
 *      - "Detail 3.1" → Show 3rd tier items for the 1st flagged item in comparison #3
 *
 * 3. TOTAL
 *    - Trigger: "Total [Item] [Financial Type]" or "Total [Financial Type] [Item]"
 *    - Purpose: Sum all sub-items under a parent item for a Financial Type
 *    - Without month: uses Financial Status sheet
 *    - With month: uses the specific Financial Type worksheet
 *    - Examples:
 *      - "Total Preliminaries Cashflow"
 *      - "Total Committed Materials"
 *      - "Total Cashflow Preliminaries Jan 2025"
 *      - "Total Plant Projection"
 *
 * 4. COMPARE
 *    - Trigger: "compare X with Y", "X vs Y", "X versus Y", "X compared to Y"
 *    - Purpose: Compare same metric across two Financial Types
 *    - Always from Financial Status sheet
 *    - Examples:
 *      - "compare projected gp with business plan gp"
 *      - "projection vs budget"
 *      - "wip gp versus projection gp"
 *      - "committed vs projection"
 *
 * 5. TREND (planned)
 *    - Trigger: "trend [metric] [N months]"
 *    - Purpose: Show metric trend over time
 *    - Examples:
 *      - "trend gp 6 months"
 *      - "trend cash flow 12"
 *
 * 6. NORMAL QUERY (fallback)
 *    - Any query not matching above patterns
 *    - Uses fuzzy matching to find Financial Type, Data Type, and optional month
 *    - Examples:
 *      - "projected gp"
 *      - "business plan np"
 *      - "gp jan 2025"
 *      - "wip gp"
 *
 * ============================================================
 * FINANCIAL TYPE SHORTCUTS (resolved by expandAcronyms + resolveFinancialType):
 *   bp, budget, business plan           → "Business Plan"
 *   projection, projected               → "Projection as at"
 *   wip, audit, audit report            → "Audit Report (WIP)"
 *   committed, committed value/cost     → "Committed Value / Cost as at"
 *   revision, rev, budget revision      → "Revision as at"
 *   tender, budget tender               → "Budget Tender"
 *   cf, cashflow, cash flow             → "Cash Flow Actual received & paid as at"
 *   accrual, accrued                    → "Accrual"
 *   1st working budget, first working   → "1st Working Budget"
 *
 * DATA TYPE / ITEM SHORTCUTS (ACRONYM_MAP below):
 *   gp                → Gross Profit
 *   np                → Net Profit
 *   prelim            → Preliminaries
 *   subcon, sub       → Subcontractor
 *   rebar             → Reinforcement
 *   staff             → Manpower (Mgt. & Supervision)
 *   labour, labor     → Manpower (Labour)
 *   material          → Materials
 *   plant, machinery  → Plant and Machinery
 *   profit, income    → Gross Profit
 *   loss              → Net Loss
 *
 * QUERY PRIORITY ORDER:
 *   Analyze → Detail → Total → Compare → Normal
 * ============================================================
 */

// Acronym mapping - comprehensive shortcut expansion
const ACRONYM_MAP: Record<string, string> = {
  // === Data Type shortcuts ===
  'gp': 'gross profit',
  'np': 'net profit',

  // === Financial Type shortcuts ===
  'bp': 'business plan',
  'budget': 'business plan',
  'revision': 'revision as at',
  'rev': 'revision as at',
  'tender': 'budget tender',
  'committed': 'committed value',
  'accrual': 'accrual',
  'wip': 'audit report (wip)',
  'projection': 'projection as at',
  'projected': 'projection as at',
  'cf': 'cash flow',
  'cashflow': 'cash flow',
  'cash': 'cash flow',

  // === Item / category shortcuts ===
  'subcon': 'subcontractor',
  'sub': 'subcontractor',
  'subcontractor': 'subcontractor',
  'rebar': 'reinforcement',
  'staff': 'manpower (mgt. & supervision)',
  'labour': 'manpower (labour)',
  'labor': 'manpower (labour)',
  'prelim': 'preliminaries',
  'preliminary': 'preliminaries',
  'material': 'materials',
  'plant': 'plant and machinery',
  'machinery': 'plant and machinery',
  'lab': 'labour',

  // === Other synonyms ===
  'profit': 'gross profit',
  'income': 'gross profit',
  'revenue': 'gross profit',
  'loss': 'net loss',
  'actual': 'actual cost',
  'accrued': 'accrual',
  'ytd': 'year to date',
  'mtd': 'month to date',
}

function expandAcronyms(text: string): string {
  const words = text.toLowerCase().split(/\s+/)
  return words.map(word => ACRONYM_MAP[word] || word).join(' ')
}

// Helper to convert Value to number safely
function toNumber(val: number | string): number {
  if (typeof val === 'number') return val
  return parseFloat(val) || 0
}

interface FinancialRow {
  Year: string
  Month: string
  Sheet_Name: string
  Financial_Type: string
  Data_Type: string
  Item_Code: string
  Value: number | string
  _project?: string
}

interface ProjectInfo {
  code: string
  name: string
  year: string
  month: string
  filename: string
  fileId?: string
}

interface FolderStructure {
  [year: string]: string[]
}

// Google Drive helper functions
async function getDriveService() {
  // For Vercel: use environment variable
  const credentials = process.env.GOOGLE_CREDENTIALS
  if (!credentials) {
    throw new Error('GOOGLE_CREDENTIALS not set')
  }

  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(credentials),
    scopes: ['https://www.googleapis.com/auth/drive.readonly']
  })

  return google.drive({ version: 'v3', auth })
}

async function findRootFolder(drive: any) {
  const res = await drive.files.list({
    q: "name='Ai Chatbot Knowledge Base' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    fields: 'files(id, name)'
  })

  return res.data.files?.[0] || null
}

async function listYearFolders(drive: any, rootId: string) {
  const allFiles: any[] = []
  let pageToken: string | null = null

  do {
    const res = await drive.files.list({
      q: `'${rootId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields: 'files(id, name), nextPageToken',
      pageSize: 100,
      pageToken: pageToken || undefined
    })
    if (res.data.files) allFiles.push(...res.data.files)
    pageToken = res.data.nextPageToken || null
  } while (pageToken)

  return allFiles
}

async function listMonthFolders(drive: any, yearId: string) {
  const allFiles: any[] = []
  let pageToken: string | null = null

  do {
    const res = await drive.files.list({
      q: `'${yearId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields: 'files(id, name), nextPageToken',
      pageSize: 100,
      pageToken: pageToken || undefined
    })
    if (res.data.files) allFiles.push(...res.data.files)
    pageToken = res.data.nextPageToken || null
  } while (pageToken)

  return allFiles
}

async function listCsvFiles(drive: any, monthId: string) {
  const allFiles: any[] = []
  let pageToken: string | null = null

  do {
    const res = await drive.files.list({
      q: `'${monthId}' in parents and name contains '_flat.csv' and trashed=false`,
      fields: 'files(id, name), nextPageToken',
      pageSize: 100,
      pageToken: pageToken || undefined
    })
    if (res.data.files) allFiles.push(...res.data.files)
    pageToken = res.data.nextPageToken || null
  } while (pageToken)

  return allFiles
}

// Extract project code and name from filename
function extractProjectInfo(filename: string): { code: string | null; name: string } {
  const name = filename.replace('_flat.csv', '')
  const match = name.match(/^(\d+)/)
  if (match) {
    const code = match[1]
    const projectName = name.slice(code.length).trim()
    const cleanName = projectName.replace(/\s*Financial\s*Report.*/i, '').trim()
    return { code, name: cleanName }
  }
  return { code: null, name }
}

// Get folder structure from Google Drive
async function getFolderStructure(): Promise<{ folders: FolderStructure; projects: Record<string, ProjectInfo>; error?: string }> {
  const folders: FolderStructure = {}
  const projects: Record<string, ProjectInfo> = {}

  try {
    const drive = await getDriveService()
    const rootFolder = await findRootFolder(drive)

    if (!rootFolder) {
      return { folders, projects, error: 'Folder "Ai Chatbot Knowledge Base" not found' }
    }

    const yearFolders = await listYearFolders(drive, rootFolder.id!)

    for (const yearFolder of yearFolders) {
      const year = yearFolder.name!
      const monthFolders = await listMonthFolders(drive, yearFolder.id!)

      for (const monthFolder of monthFolders) {
        const csvFiles = await listCsvFiles(drive, monthFolder.id!)

        if (csvFiles.length > 0) {
          if (!folders[year]) folders[year] = []
          if (!folders[year].includes(monthFolder.name!)) folders[year].push(monthFolder.name!)

          for (const file of csvFiles) {
            const { code, name } = extractProjectInfo(file.name!)
            if (code) {
              projects[file.name!] = {
                code,
                name,
                year,
                month: monthFolder.name!,
                filename: file.name!,
                fileId: file.id!
              }
            }
          }
        }
      }
    }
  } catch (error: any) {
    console.error('Error getting folder structure:', error)
    return { folders, projects, error: error.message || 'Unknown error' }
  }

  return { folders, projects }
}

// Load a single project CSV from Google Drive
async function loadProjectData(filename: string, year: string, month: string): Promise<FinancialRow[]> {
  try {
    const drive = await getDriveService()
    const rootFolder = await findRootFolder(drive)
    if (!rootFolder) return []

    const yearFolders = await listYearFolders(drive, rootFolder.id!)
    const yearFolder = yearFolders.find(f => f.name === year)
    if (!yearFolder) return []

    const monthFolders = await listMonthFolders(drive, yearFolder.id!)
    const monthFolder = monthFolders.find(f => f.name === month)
    if (!monthFolder) return []

    const csvFiles = await listCsvFiles(drive, monthFolder.id!)
    const targetFile = csvFiles.find(f => f.name === filename)
    if (!targetFile) return []

    // Download file
    const res = await drive.files.get({
      fileId: targetFile.id!,
      alt: 'media'
    }, { responseType: 'text' })

    // Parse CSV - columns are at fixed positions:
    // 0: Year, 1: Month, 2: Sheet_Name, 3: Financial_Type, 4: Item_Code, 5: Data_Type, 6: Value
    const lines = (res.data as string).split('\n').filter(line => line.trim())

    const { name } = extractProjectInfo(filename)
    const code = filename.match(/^(\d+)/)?.[1] || ''
    const projectLabel = `${code} - ${name}`

    const data: FinancialRow[] = []
    for (let i = 0; i < lines.length; i++) {
      // Handle quoted CSV values
      const values: string[] = []
      let inQuote = false
      let current = ''
      for (let j = 0; j < lines[i].length; j++) {
        const char = lines[i][j]
        if (char === '"') {
          inQuote = !inQuote
        } else if (char === ',' && !inQuote) {
          values.push(current.trim().replace(/"/g, ''))
          current = ''
        } else {
          current += char
        }
      }
      values.push(current.trim().replace(/"/g, ''))

      // Skip header row if present
      const firstValue = values[0]?.toLowerCase()
      if (i === 0 && (firstValue === 'year' || firstValue === 'sheet_name')) continue

      // Parse each column - columns are at fixed positions:
      // 0: Year, 1: Month, 2: Sheet_Name, 3: Financial_Type, 4: Item_Code, 5: Data_Type, 6: Value

      // Extract Data_Type directly from column 5 (not pattern matching)
      const dataType = values[5] || ''

      // Extract Item_Code from column 4
      const itemCode = values[4] || ''

      // Financial_Type from column 3
      const financialType = values[3] || ''

      // For Project Info (Financial_Type = "General"), keep values as strings (dates, percentages)
      // For financial data, parse as numbers
      const rawValue = values[values.length - 1] || ''
      let value: number | string

      if (financialType === 'General') {
        // Keep Project Info values as strings
        value = rawValue
      } else {
        // Parse financial values as numbers
        value = parseFloat(rawValue) || 0
      }

      const row: FinancialRow = {
        Year: values[0] || '',
        Month: values[1] || '',
        Sheet_Name: values[2] || '',
        Financial_Type: financialType,
        Data_Type: dataType || '',
        Item_Code: itemCode,
        Value: value,
        _project: projectLabel
      }
      data.push(row)
    }
    return data
  } catch (error) {
    console.error('Error loading project data:', error)
    return []
  }
}

// Calculate project metrics
function getProjectMetrics(data: FinancialRow[], project: string) {
  const projectData = data.filter(d => d._project === project)
  if (projectData.length === 0) {
    return {
      'Business Plan GP': 0,
      'Projected GP': 0,
      'WIP GP': 0,
      'Cash Flow': 0,
      'Start Date': 'N/A',
      'Complete Date': 'N/A',
      'Target Complete Date': 'N/A',
      'Time Consumed (%)': 'N/A',
      'Target Completed (%)': 'N/A'
    }
  }

  const gpFilter = (d: FinancialRow) => d.Item_Code === '3' && d.Data_Type?.toLowerCase().includes('gross profit')

  // Helper to convert Value to number safely
  const toNumber = (val: number | string): number => {
    if (typeof val === 'number') return val
    return parseFloat(val) || 0
  }

  const bp = projectData.filter(d =>
    d.Sheet_Name === 'Financial Status' &&
    d.Financial_Type?.toLowerCase().includes('business plan') &&
    gpFilter(d)
  ).reduce((sum, d) => sum + toNumber(d.Value), 0)

  const proj = projectData.filter(d =>
    d.Sheet_Name === 'Financial Status' &&
    d.Financial_Type?.toLowerCase().includes('projection') &&
    gpFilter(d)
  ).reduce((sum, d) => sum + toNumber(d.Value), 0)

  const wip = projectData.filter(d =>
    d.Sheet_Name === 'Financial Status' &&
    d.Financial_Type?.toLowerCase().includes('audit report') &&
    gpFilter(d)
  ).reduce((sum, d) => sum + toNumber(d.Value), 0)

  const cf = projectData.filter(d =>
    d.Sheet_Name === 'Financial Status' &&
    d.Financial_Type?.toLowerCase().includes('cash flow') &&
    gpFilter(d)
  ).reduce((sum, d) => sum + toNumber(d.Value), 0)

  // Extract Project Info (Financial_Type = "General")
  const projectInfo = projectData.filter(d => d.Financial_Type === 'General')

  const getProjectInfoValue = (dataType: string) => {
    const row = projectInfo.find(d => d.Data_Type === dataType)
    const val = row ? String(row.Value) : ''
    return (val && val !== 'Nil') ? val : 'N/A'
  }

  return {
    'Business Plan GP': bp,
    'Projected GP': proj,
    'WIP GP': wip,
    'Cash Flow': cf,
    'Start Date': getProjectInfoValue('Start Date'),
    'Complete Date': getProjectInfoValue('Complete Date'),
    'Target Complete Date': getProjectInfoValue('Target Complete Date'),
    'Time Consumed (%)': getProjectInfoValue('Time Consumed (%)'),
    'Target Completed (%)': getProjectInfoValue('Target Completed (%)')
  }
}

// Handle monthly category queries
function handleMonthlyCategory(data: FinancialRow[], project: string, question: string, month: string): string {
  const expandedQuestion = expandAcronyms(question).toLowerCase()

  const categoryMap: Record<string, string> = {
    'plant and machinery': '2.3',
    'preliminaries': '2.1',
    'materials': '2.2',
    'plant': '2.3',
    'machinery': '2.3',
    'labour': '2.4',
    'labor': '2.4',
    'manpower (labour) for works': '2.5',
    'manpower (labour)': '2.5',
    'manpower': '2.5',
    'subcontractor': '2.5',
    'subcon': '2.5',
    'staff': '2.6',
    'admin': '2.7',
    'administration': '2.7',
    'insurance': '2.8',
    'bond': '2.9',
    'others': '2.10',
    'other': '2.10',
    'contingency': '2.11',
  }

  if (!expandedQuestion.includes('monthly')) return ''

  let categoryPrefix = ''
  let categoryName = ''

  const sortedCategories = Object.entries(categoryMap).sort((a, b) => b[0].length - a[0].length)
  for (const [kw, prefix] of sortedCategories) {
    const pattern = new RegExp(`\\b${kw}\\b`, 'i')
    if (pattern.test(expandedQuestion)) {
      categoryPrefix = prefix
      categoryName = kw
      break
    }
  }

  if (!categoryPrefix) return ''

  const projectData = data.filter(d => d._project === project)

  const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
  const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

  let targetMonth = parseInt(month)
  for (let i = 0; i < monthNames.length; i++) {
    if (expandedQuestion.includes(monthNames[i]) || expandedQuestion.includes(monthAbbr[i])) {
      targetMonth = i + 1
      break
    }
  }

  const financialTypes = ['Projection', 'Committed Cost', 'Accrual', 'Cash Flow']
  const results: Record<string, number> = {}

  for (const ft of financialTypes) {
    const filtered = projectData.filter(d =>
      d.Sheet_Name === ft &&
      d.Month === String(targetMonth) &&
      (d.Item_Code.startsWith(categoryPrefix + '.') || d.Item_Code === categoryPrefix)
    )
    results[ft] = filtered.reduce((sum, d) => sum + toNumber(d.Value), 0)
  }

  const displayName = categoryName.charAt(0).toUpperCase() + categoryName.slice(1)
  let response = `## Monthly ${displayName} (${targetMonth}/${month}) ('000)\n\n`

  for (const [ft, value] of Object.entries(results)) {
    response += `- **${ft}:** $${value.toLocaleString()}\n`
  }

  return response
}

// ============================================
// New Query Logic (Revised)
// ============================================

interface ParsedQuery {
  year?: string
  month?: string
  sheetName?: string
  financialType?: string
  dataType?: string
  itemCode?: string
  monthsAgo?: number  // PHASE 2: for "last 3 months" type queries
  comparison?: string  // PHASE 2: for "vs", "compare" type queries
}

interface FuzzyResult {
  text: string
  candidates: Array<{
    id: number
    value: number | string
    score: number
    sheet: string
    financialType: string
    dataType: string
    itemCode: string
    month: string
    year: string
    matchedKeywords: string[]
  }>
}

// Parse date from question - maps "january 2025" or "2025 jan" or "jan" or "feb 25" or "1st month" or "last month" to month/year
function parseDate(question: string, defaultMonth: string): ParsedQuery {
  const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
  const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
  const monthMap: Record<string, string> = {}
  monthNames.forEach((name, i) => monthMap[name] = String(i + 1))
  monthAbbr.forEach((abbr, i) => monthMap[abbr] = String(i + 1))

  // Ordinal number map: 1st, 2nd, 3rd, etc.
  const ordinalMap: Record<string, string> = {
    '1st': '1', '2nd': '2', '3rd': '3', '4th': '4', '5th': '5', '6th': '6',
    '7th': '7', '8th': '8', '9th': '9', '10th': '10', '11th': '11', '12th': '12'
  }

  const result: ParsedQuery = {}
  const lowerQ = question.toLowerCase()
  const currentMonth = new Date().getMonth() + 1
  const currentYear = new Date().getFullYear()

  // === PHASE 2: Time Keywords ===
  // Handle "last month", "this month", "previous month", "last 3 months"
  if (lowerQ.includes('last month') || lowerQ.includes('previous month') || lowerQ.includes('prev month')) {
    const lastMonth = currentMonth === 1 ? 12 : currentMonth - 1
    const lastYear = currentMonth === 1 ? currentYear - 1 : currentYear
    result.month = String(lastMonth)
    result.year = String(lastYear)
    return result
  }
  
  if (lowerQ.includes('this month')) {
    result.month = String(currentMonth)
    result.year = String(currentYear)
    return result
  }

  // Handle "last 3 months", "last 6 months" (for aggregation - mark for special handling)
  const lastMonthsMatch = lowerQ.match(/last (\d+) months?/)
  if (lastMonthsMatch) {
    result.monthsAgo = parseInt(lastMonthsMatch[1])
    return result
  }

  // === PHASE 2: Comparison Queries ===
  // Handle "vs", "versus", "compare", "difference" for comparisons
  if (lowerQ.includes(' vs ') || lowerQ.includes(' versus ') || 
      lowerQ.includes(' compare ') || lowerQ.includes(' compared ') ||
      lowerQ.includes(' difference ')) {
    result.comparison = 'difference'
  }

  // Check for ordinal numbers first (1st, 2nd, 3rd... 12th)
  for (const [ordinal, monthNum] of Object.entries(ordinalMap)) {
    if (lowerQ.includes(ordinal)) {
      result.month = monthNum
      break
    }
  }

  // Find 4-digit year (2024, 2025, etc.)
  const yearMatch = lowerQ.match(/\b(20[2-4]\d)\b/)
  if (yearMatch) {
    result.year = yearMatch[1]
  }

  // Handle "M/YY" or "MM/YY" format like "2/25" or "02/25" → Feb 2025
  // This must be checked BEFORE 2-digit year
  const mmyyMatch = lowerQ.match(/(\d{1,2})\/(\d{2})\b/)
  if (mmyyMatch && !yearMatch) {
    const monthNum = mmyyMatch[1]
    const yearNum = mmyyMatch[2]
    // Validate month is 1-12
    const month = parseInt(monthNum)
    if (month >= 1 && month <= 12) {
      result.month = String(month)
      result.year = '20' + yearNum
    }
  }

  // Find 2-digit year (24, 25, etc.) - only if preceded by space and not part of month name
  // "feb 25" should be Feb + 2025, not Feb + month 25
  const twoDigitYearMatch = lowerQ.match(/\b(\d{2})\b(?!.*\d{4})/)
  if (twoDigitYearMatch && !yearMatch && !result.year) {
    const year = parseInt(twoDigitYearMatch[1])
    if (year >= 20 && year <= 30) {
      result.year = '20' + twoDigitYearMatch[1]
    }
  }

  // Find month name or abbreviation - only if it doesn't conflict with year
  for (let i = 0; i < monthNames.length; i++) {
    if (lowerQ.includes(monthNames[i]) || lowerQ.includes(monthAbbr[i])) {
      const monthNum = String(i + 1)
      // Only set month if it doesn't look like a year (e.g., don't treat "2025" as month 20)
      if (monthNum.length === 1 || (monthNum.length === 2 && monthNum !== '20' && monthNum !== '21' && monthNum !== '22' && monthNum !== '23')) {
        result.month = monthNum
      }
      break
    }
  }

  // If no month found, use default month
  if (!result.month) {
    result.month = defaultMonth
  }

  return result
}

// Find closest match for a value against a list of candidates using Levenshtein distance
function findClosestMatch(input: string, candidates: string[]): string | null {
  if (!input || candidates.length === 0) return null

  const normalizedInput = input.toLowerCase().trim()
  let bestMatch: string | null = null
  let bestDistance = Infinity

  for (const candidate of candidates) {
    const normalizedCand = candidate.toLowerCase().trim()

    // Exact match - return immediately
    if (normalizedInput === normalizedCand) return candidate

    // Check for EXACT SUBSTRING match - input must appear as a WHOLE WORD in candidate
    // Split candidate into words and check if input matches any word
    const candWords = normalizedCand.split(/\s+/)
    let isWordMatch = false
    for (let i = 0; i < candWords.length; i++) {
      const word = candWords[i]
      // Input must match a candidate word EXACTLY, or be a substantial part (50%+) of that word
      if (word === normalizedInput) {
        isWordMatch = true
        break
      }
      // Check if word contains input OR input contains word (and input is substantial part)
      if (word.includes(normalizedInput)) {
        // Input is substring of word - only accept if input is at least 50% of the word
        if (normalizedInput.length >= word.length * 0.5) {
          isWordMatch = true
          break
        }
      }
    }

    if (isWordMatch) {
      // It's a word-level match
      const distance = Math.abs(normalizedCand.length - normalizedInput.length)
      if (distance < bestDistance) {
        bestDistance = distance
        bestMatch = candidate
      }
      continue
    }

    // For fuzzy matching (not substring), check character similarity
    const inputChars: Record<string, boolean> = {}
    for (let i = 0; i < normalizedInput.length; i++) {
      const c = normalizedInput[i]
      if (c !== ' ') inputChars[c] = true
    }
    const candChars: Record<string, boolean> = {}
    for (let i = 0; i < normalizedCand.length; i++) {
      const c = normalizedCand[i]
      if (c !== ' ') candChars[c] = true
    }
    let sharedChars = 0
    for (const c in inputChars) {
      if (candChars[c]) sharedChars++
    }
    const totalChars = Object.keys(inputChars).length + Object.keys(candChars).length - sharedChars
    const similarity = totalChars > 0 ? sharedChars / totalChars : 0

    // Only consider fuzzy match if at least 60% character similarity
    if (similarity < 0.6) continue

    // Levenshtein distance for fuzzy matching
    const distance = levenshteinDistance(normalizedInput, normalizedCand)
    if (distance < bestDistance) {
      bestDistance = distance
      bestMatch = candidate
    }
  }

  // Only return if reasonably close (word match OR very close fuzzy match)
  if (bestMatch && bestDistance <= normalizedInput.length / 2) {
    return bestMatch
  }
  return null
}

// Levenshtein distance calculation
function levenshteinDistance(a: string, b: string): number {
  const matrix: number[][] = []
  for (let i = 0; i <= b.length; i++) matrix[i] = [i]
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1]
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        )
      }
    }
  }
  return matrix[b.length][a.length]
}

// Format currency without decimals - e.g., $20.01 → $20
function formatCurrency(value: number): string {
  return `$${Math.round(value).toLocaleString()}`
}

// Get all unique values for a column
function getUniqueValues(data: FinancialRow[], project: string, field: keyof FinancialRow): string[] {
  const projectData = data.filter(d => d._project === project)
  const values = new Set<string>()
  projectData.forEach(row => {
    const val = row[field]
    if (val) values.add(String(val))
  })
  return Array.from(values)
}

// ============================================
// Comparison Query Handler
// ============================================

// Detect if a query is a comparison query and extract the two sides
function isComparisonQuery(question: string): boolean {
  const lowerQ = question.toLowerCase()
  return (
    lowerQ.includes(' vs ') || lowerQ.includes(' versus ') ||
    lowerQ.includes('compare ') || lowerQ.includes('compared ') ||
    lowerQ.includes('comparison ') ||
    // "X with Y" pattern (but only if combined with financial terms)
    (lowerQ.includes(' with ') && (
      lowerQ.includes('projected') || lowerQ.includes('projection') ||
      lowerQ.includes('business plan') || lowerQ.includes('budget') ||
      lowerQ.includes('wip') || lowerQ.includes('audit') ||
      lowerQ.includes('cash flow') || lowerQ.includes('actual') ||
      lowerQ.includes('tender') || lowerQ.includes('committed') ||
      lowerQ.includes('gp') || lowerQ.includes('np') ||
      lowerQ.includes('gross profit') || lowerQ.includes('net profit')
    ))
  )
}

// Extract two Financial Type keywords from a comparison query
// e.g. "compare projected gp with business plan gp" → ["projection", "business plan"]
// e.g. "projected gp vs budget gp" → ["projection", "business plan"]
function extractComparisonParts(expandedQuestion: string): { side1: string; side2: string; metric: string } | null {
  const lowerQ = expandedQuestion.toLowerCase()

  // Split by comparison keywords
  let parts: string[] = []
  if (lowerQ.includes(' vs ')) {
    parts = lowerQ.split(' vs ')
  } else if (lowerQ.includes(' versus ')) {
    parts = lowerQ.split(' versus ')
  } else if (lowerQ.includes(' compared to ')) {
    parts = lowerQ.split(' compared to ')
  } else if (lowerQ.includes(' compared with ')) {
    parts = lowerQ.split(' compared with ')
  } else if (lowerQ.match(/compare\s+(.+?)\s+with\s+(.+)/)) {
    const match = lowerQ.match(/compare\s+(.+?)\s+with\s+(.+)/)!
    parts = [match[1], match[2]]
  } else if (lowerQ.match(/compare\s+(.+?)\s+and\s+(.+)/)) {
    const match = lowerQ.match(/compare\s+(.+?)\s+and\s+(.+)/)!
    parts = [match[1], match[2]]
  } else if (lowerQ.includes(' with ')) {
    // Generic "X with Y" pattern
    parts = lowerQ.split(' with ')
  }

  if (parts.length < 2) return null

  return {
    side1: parts[0].trim(),
    side2: parts[1].trim(),
    metric: '' // Will be determined later
  }
}

// Match a text fragment to a Financial_Type from the available types
function matchFinancialType(text: string, financialTypes: string[]): string | null {
  const lowerText = text.toLowerCase()

  // Known keyword-to-Financial_Type mappings (expanded text after acronym expansion)
  // IMPORTANT: "business plan" / "budget" → "Business Plan" (NOT "1st Working Budget")
  // "revision" → "Revision as at"
  const knownMappings: Record<string, string[]> = {
    'projection as at': ['projection'],
    'projection': ['projection'],
    'projected': ['projection'],
    'business plan': ['business plan'],
    'budget': ['business plan'],
    'audit report (wip)': ['audit report'],
    'audit report': ['audit report'],
    'wip': ['audit report'],
    'cash flow': ['cash flow'],
    'tender': ['tender', 'budget tender'],
    'actual': ['actual'],
    'committed': ['committed'],
    'committed value': ['committed'],
    'accrual': ['accrual'],
    '1st working budget': ['1st working budget'],
    'first working budget': ['1st working budget'],
    'working budget': ['1st working budget'],
    'revision as at': ['revision as at', 'revision'],
    'revision': ['revision as at', 'revision'],
    'budget revision': ['revision as at', 'revision'],
    'balance': ['balance'],
    'adjustment': ['adjustment'],
    'variation': ['adjustment', 'variation'],
  }

  // First: try known mappings
  for (const [keyword, targetPhrases] of Object.entries(knownMappings)) {
    if (lowerText.includes(keyword)) {
      // Find matching Financial_Type from available types
      for (const phrase of targetPhrases) {
        const match = financialTypes.find(ft => ft.toLowerCase().includes(phrase))
        if (match) return match
      }
    }
  }

  // Second: try direct word matching against Financial_Type values
  const textWords = lowerText.split(/\s+/).filter(w => w.length > 1)
  for (const ft of financialTypes) {
    const ftLower = ft.toLowerCase()
    const ftWords = ftLower.split(/\s+/)
    for (const tw of textWords) {
      for (const fw of ftWords) {
        if (tw === fw && tw.length > 2) return ft
        if (fw.includes(tw) && tw.length >= fw.length * 0.5) return ft
      }
    }
  }

  // Third: fuzzy match
  for (const word of textWords) {
    const match = findClosestMatch(word, financialTypes)
    if (match) return match
  }

  return null
}

// Extract the metric (Data_Type) from comparison text
function extractComparisonMetric(expandedQuestion: string, dataTypes: string[]): string | null {
  // Known acronym-to-DataType mappings
  const acronymDataMap: Record<string, string[]> = {
    'gross profit': ['gross profit', 'acc. gross profit'],
    'net profit': ['net profit', 'acc. net profit'],
    'cash flow': ['cash flow'],
    'revenue': ['revenue', 'gross profit'],
    'income': ['income', 'gross profit'],
  }

  const lowerQ = expandedQuestion.toLowerCase()

  // Check known mappings first
  for (const [keyword, expansions] of Object.entries(acronymDataMap)) {
    if (lowerQ.includes(keyword)) {
      for (const expansion of expansions) {
        const match = dataTypes.find(dt => dt.toLowerCase().includes(expansion))
        if (match) return match
      }
    }
  }

  // Try word-by-word matching
  const questionWords = lowerQ.split(/\s+/).filter(w => w.length > 2)
  let bestMatch: string | null = null
  let bestScore = 0

  for (const dt of dataTypes) {
    const dtLower = dt.toLowerCase()
    const dtWords = dtLower.split(/\s+/)
    let score = 0
    for (const qw of questionWords) {
      for (const dw of dtWords) {
        if (qw === dw) score += 2
        else if (dw.includes(qw) && qw.length >= dw.length * 0.5) score += 1
      }
    }
    if (score > bestScore) {
      bestScore = score
      bestMatch = dt
    }
  }

  return bestMatch
}

// Handle comparison query - returns formatted comparison result or null if not a comparison
function handleComparisonQuery(data: FinancialRow[], project: string, question: string, defaultMonth: string): FuzzyResult | null {
  const expandedQuestion = expandAcronyms(question).toLowerCase()

  // Check if this is a comparison query
  if (!isComparisonQuery(expandedQuestion)) return null

  const projectData = data.filter(d => d._project === project)
  if (projectData.length === 0) return null

  // Extract the two sides of the comparison
  const compParts = extractComparisonParts(expandedQuestion)
  if (!compParts) return null

  // Get available Financial Types and Data Types
  const financialTypes = getUniqueValues(data, project, 'Financial_Type')
  const dataTypes = getUniqueValues(data, project, 'Data_Type')

  // Match each side to a Financial_Type
  const finType1 = matchFinancialType(compParts.side1, financialTypes)
  const finType2 = matchFinancialType(compParts.side2, financialTypes)

  if (!finType1 || !finType2) {
    // Can't determine both Financial Types - fall back to normal query
    return null
  }

  // If both sides resolve to the same Financial_Type, it's not a meaningful comparison
  if (finType1 === finType2) return null

  // Extract the metric (Data_Type) to compare
  const targetDataType = extractComparisonMetric(expandedQuestion, dataTypes)

  // Default to Financial Status sheet for comparison
  const targetSheet = 'Financial Status'

  // Helper to get the value for a specific Financial_Type
  const getValueForType = (finType: string): { total: number; rows: FinancialRow[] } => {
    let filtered = projectData.filter(d => d.Sheet_Name === targetSheet)
    filtered = filtered.filter(d => d.Financial_Type === finType)
    if (targetDataType) {
      filtered = filtered.filter(d => d.Data_Type === targetDataType)
    }
    const total = filtered.reduce((sum, d) => sum + toNumber(d.Value), 0)
    return { total, rows: filtered }
  }

  const result1 = getValueForType(finType1)
  const result2 = getValueForType(finType2)

  // Calculate difference
  const diff = result1.total - result2.total
  const absDiff = Math.abs(diff)
  const pctChange = result2.total !== 0 ? (diff / Math.abs(result2.total)) * 100 : 0
  const arrow = diff > 0 ? '↑' : diff < 0 ? '↓' : '→'
  const sign = diff > 0 ? '+' : ''

  // Format response as comparison table
  const metricLabel = targetDataType || 'Total'
  let response = `## Comparing: ${finType1} vs ${finType2}\n`
  response += `**Table:** ${targetSheet}\n`
  response += `**Metric:** ${metricLabel}\n\n`

  // Table header
  response += `| Financial Type | ${metricLabel} |\n`
  response += `|----------------|${'-'.repeat(Math.max(metricLabel.length, 14))}|\n`
  response += `| ${finType1} | ${formatCurrency(result1.total)} |\n`
  response += `| ${finType2} | ${formatCurrency(result2.total)} |\n`
  response += `| Difference | ${arrow} ${formatCurrency(absDiff)} (${sign}${pctChange.toFixed(1)}%) |\n`

  response += `\n*Values in ('000)*`

  // Still provide candidates for drill-down
  const allRows = [...result1.rows, ...result2.rows]
  const candidates = allRows.slice(0, 10).map((d, i) => ({
    id: i + 1,
    value: d.Value,
    score: 100,
    sheet: d.Sheet_Name,
    financialType: d.Financial_Type,
    dataType: d.Data_Type,
    itemCode: d.Item_Code,
    month: d.Month,
    year: d.Year,
    matchedKeywords: []
  }))

  return { text: response, candidates }
}

// ============================================
// Total Query Handler
// ============================================

// Parent item name → item code mapping
const PARENT_ITEM_MAP: Record<string, { code: string; name: string }> = {
  'preliminaries': { code: '2.1', name: 'Preliminaries' },
  'prelim': { code: '2.1', name: 'Preliminaries' },
  'preliminary': { code: '2.1', name: 'Preliminaries' },
  'materials': { code: '2.2', name: 'Materials' },
  'material': { code: '2.2', name: 'Materials' },
  'plant and machinery': { code: '2.3', name: 'Plant and Machinery' },
  'plant': { code: '2.3', name: 'Plant and Machinery' },
  'machinery': { code: '2.3', name: 'Plant and Machinery' },
  'labour': { code: '2.4', name: 'Labour' },
  'labor': { code: '2.4', name: 'Labour' },
  'manpower': { code: '2.5', name: 'Manpower (Labour) for Works' },
  'subcontractor': { code: '2.5', name: 'Subcontractor' },
  'subcon': { code: '2.5', name: 'Subcontractor' },
  'staff': { code: '2.6', name: 'Staff' },
  'admin': { code: '2.7', name: 'Administration' },
  'administration': { code: '2.7', name: 'Administration' },
  'insurance': { code: '2.8', name: 'Insurance' },
  'bond': { code: '2.9', name: 'Bond' },
  'others': { code: '2.10', name: 'Others' },
  'other': { code: '2.10', name: 'Others' },
  'contingency': { code: '2.11', name: 'Contingency' },
}

// Financial type keyword → canonical Financial_Type name mapping
// These map user-facing keywords to the sheet names / Financial_Type values in the data
// IMPORTANT: "budget" / "bp" → "business plan" (NOT "1st working budget")
const FINANCIAL_TYPE_KEYWORDS: Record<string, string[]> = {
  'cash flow': ['cash flow', 'cashflow', 'cf'],
  'committed cost': ['committed', 'committed cost', 'committed value'],
  'projection': ['projection', 'projected'],
  'accrual': ['accrual', 'accrued'],
  'business plan': ['business plan', 'budget', 'bp'],
  'tender': ['tender', 'budget tender'],
  'actual': ['actual', 'actual cost'],
  'audit report': ['wip', 'audit', 'audit report'],
  '1st working budget': ['1st working budget', 'first working budget'],
  'revision as at': ['revision', 'rev', 'budget revision', 'revision as at'],
}

// Detect if a query starts with "Total" keyword
function isTotalQuery(question: string): boolean {
  const lowerQ = question.toLowerCase().trim()
  return lowerQ.startsWith('total ')
}

// Parse a Total query to extract item name, financial type, and optional month
function parseTotalQuery(question: string): {
  itemName: string | null
  itemCode: string | null
  itemDisplayName: string | null
  financialType: string | null
  month: string | null
  year: string | null
} {
  const lowerQ = question.toLowerCase().trim()
  // Remove "total" prefix
  const afterTotal = lowerQ.replace(/^total\s+/, '').trim()

  let itemName: string | null = null
  let itemCode: string | null = null
  let itemDisplayName: string | null = null
  let financialType: string | null = null
  let month: string | null = null
  let year: string | null = null

  // Parse month/year from query
  const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
  const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

  // Check for month names
  for (let i = 0; i < monthNames.length; i++) {
    if (afterTotal.includes(monthNames[i]) || afterTotal.includes(monthAbbr[i])) {
      month = String(i + 1)
      break
    }
  }

  // Check for 4-digit year
  const yearMatch = afterTotal.match(/\b(20[2-4]\d)\b/)
  if (yearMatch) {
    year = yearMatch[1]
  }

  // Check for 2-digit year
  if (!year) {
    const twoDigitMatch = afterTotal.match(/\b(\d{2})\b/)
    if (twoDigitMatch) {
      const y = parseInt(twoDigitMatch[1])
      if (y >= 20 && y <= 30) {
        year = '20' + twoDigitMatch[1]
      }
    }
  }

  // Try to find item name - sort by longest key first for greedy matching
  const sortedItems = Object.entries(PARENT_ITEM_MAP).sort((a, b) => b[0].length - a[0].length)
  for (const [keyword, info] of sortedItems) {
    // Use word boundary matching
    const pattern = new RegExp(`\\b${keyword.replace(/\s+/g, '\\s+')}\\b`, 'i')
    if (pattern.test(afterTotal)) {
      itemName = keyword
      itemCode = info.code
      itemDisplayName = info.name
      break
    }
  }

  // Try to find financial type - check multi-word phrases first
  const sortedFinTypes = Object.entries(FINANCIAL_TYPE_KEYWORDS).sort((a, b) => {
    const aMaxLen = Math.max(...a[1].map(k => k.length))
    const bMaxLen = Math.max(...b[1].map(k => k.length))
    return bMaxLen - aMaxLen
  })

  for (const [canonicalType, keywords] of sortedFinTypes) {
    for (const keyword of keywords.sort((a, b) => b.length - a.length)) {
      const pattern = new RegExp(`\\b${keyword.replace(/\s+/g, '\\s+')}\\b`, 'i')
      if (pattern.test(afterTotal)) {
        financialType = canonicalType
        break
      }
    }
    if (financialType) break
  }

  return { itemName, itemCode, itemDisplayName, financialType, month, year }
}

// Handle a Total query - sum all sub-items under a parent item for a given financial type
function handleTotalQuery(data: FinancialRow[], project: string, question: string, defaultMonth: string): FuzzyResult | null {
  if (!isTotalQuery(question)) return null

  const parsed = parseTotalQuery(question)

  if (!parsed.itemCode || !parsed.financialType) {
    // Build helpful error message
    let errorMsg = '❌ Could not parse Total query.\n\n'
    if (!parsed.itemCode) {
      errorMsg += '**Missing item name.** Supported items:\n'
      const seen = new Set<string>()
      for (const [, info] of Object.entries(PARENT_ITEM_MAP)) {
        if (!seen.has(info.code)) {
          seen.add(info.code)
          errorMsg += `• ${info.code} - ${info.name}\n`
        }
      }
    }
    if (!parsed.financialType) {
      errorMsg += '\n**Missing financial type.** Supported types:\n'
      for (const [canonicalType, keywords] of Object.entries(FINANCIAL_TYPE_KEYWORDS)) {
        errorMsg += `• ${canonicalType} (keywords: ${keywords.join(', ')})\n`
      }
    }
    errorMsg += '\n**Example:** "Total Preliminaries Cashflow", "Total Committed Materials Jan 2025"'
    return { text: errorMsg, candidates: [] }
  }

  const projectData = data.filter(d => d._project === project)
  if (projectData.length === 0) {
    return { text: 'No data found for this project.', candidates: [] }
  }

  // Determine which sheet to use and how to filter
  const hasMonth = !!parsed.month
  let sheetName: string
  let targetFinancialType: string | null = null

  if (hasMonth) {
    // User specified a month → go to the worksheet matching the financial type name
    // e.g., "Cash Flow" sheet, "Projection" sheet, "Committed Cost" sheet
    // The sheet name typically matches the canonical financial type name
    const sheetCandidates = Array.from(new Set(projectData.map(d => d.Sheet_Name)))

    // Try to find a sheet matching the financial type
    const ftLower = parsed.financialType.toLowerCase()
    let matchedSheet = sheetCandidates.find(s => s.toLowerCase() === ftLower)
    if (!matchedSheet) {
      // Try partial match
      matchedSheet = sheetCandidates.find(s => s.toLowerCase().includes(ftLower) || ftLower.includes(s.toLowerCase()))
    }
    if (!matchedSheet) {
      // Fallback to Financial Status
      matchedSheet = 'Financial Status'
    }
    sheetName = matchedSheet
  } else {
    // No month specified → use Financial Status sheet
    sheetName = 'Financial Status'
    // On Financial Status, we need to match Financial_Type column
    // Resolve the canonical name to an actual Financial_Type in the data
    const ftLower = parsed.financialType.toLowerCase()
    const availableFinTypes = Array.from(new Set(projectData.filter(d => d.Sheet_Name === 'Financial Status').map(d => d.Financial_Type)))

    // Try exact match
    targetFinancialType = availableFinTypes.find(ft => ft.toLowerCase() === ftLower) || null
    if (!targetFinancialType) {
      // Try partial/contains match
      targetFinancialType = availableFinTypes.find(ft => ft.toLowerCase().includes(ftLower) || ftLower.includes(ft.toLowerCase())) || null
    }
    if (!targetFinancialType) {
      // Try keyword-based matching using FINANCIAL_TYPE_KEYWORDS
      const keywords = FINANCIAL_TYPE_KEYWORDS[parsed.financialType] || []
      for (const keyword of keywords) {
        targetFinancialType = availableFinTypes.find(ft => ft.toLowerCase().includes(keyword)) || null
        if (targetFinancialType) break
      }
    }

    if (!targetFinancialType) {
      const availableList = availableFinTypes.map(ft => `• ${ft}`).join('\n')
      return {
        text: `❌ Could not find Financial Type "${parsed.financialType}" in Financial Status sheet.\n\nAvailable Financial Types:\n${availableList}`,
        candidates: []
      }
    }
  }

  // Filter data: get sub-items under the parent item code
  const parentCode = parsed.itemCode // e.g., "2.1"
  const subItemPrefix = parentCode + '.' // e.g., "2.1."

  let filtered = projectData.filter(d => {
    // Must be a sub-item (e.g., 2.1.1, 2.1.2, etc.) - NOT the parent itself
    if (!d.Item_Code.startsWith(subItemPrefix)) return false
    // Must be on the right sheet
    if (d.Sheet_Name !== sheetName) return false
    return true
  })

  // Apply Financial_Type filter (for Financial Status sheet)
  if (targetFinancialType && sheetName === 'Financial Status') {
    filtered = filtered.filter(d => d.Financial_Type === targetFinancialType)
  }

  // Apply month filter if specified
  if (parsed.month) {
    filtered = filtered.filter(d => d.Month === parsed.month)
  }

  // Apply year filter if specified
  if (parsed.year) {
    filtered = filtered.filter(d => d.Year === parsed.year)
  }

  if (filtered.length === 0) {
    // Try to provide helpful info about what's available
    const availableSubItems = projectData.filter(d => d.Item_Code.startsWith(subItemPrefix))
    const availableSheets = Array.from(new Set(availableSubItems.map(d => d.Sheet_Name)))
    const availableFinTypes = Array.from(new Set(availableSubItems.map(d => d.Financial_Type)))

    let msg = `No sub-items found for ${parsed.itemDisplayName} (${parentCode}) `
    if (targetFinancialType) msg += `with Financial Type "${targetFinancialType}" `
    msg += `on sheet "${sheetName}"`
    if (parsed.month) msg += ` for month ${parsed.month}`
    msg += '.\n\n'

    if (availableSheets.length > 0) {
      msg += `**Available sheets for ${parentCode}.x items:** ${availableSheets.join(', ')}\n`
    }
    if (availableFinTypes.length > 0) {
      msg += `**Available Financial Types:** ${availableFinTypes.join(', ')}\n`
    }

    return { text: msg, candidates: [] }
  }

  // Group by sub-item code and sum values
  const subItemGroups = new Map<string, { dataType: string; total: number; rows: FinancialRow[] }>()

  for (const row of filtered) {
    const key = row.Item_Code
    if (!subItemGroups.has(key)) {
      subItemGroups.set(key, { dataType: row.Data_Type, total: 0, rows: [] })
    }
    const group = subItemGroups.get(key)!
    group.total += toNumber(row.Value)
    group.rows.push(row)
  }

  // Sort sub-items by item code numerically
  const sortedSubItems = Array.from(subItemGroups.entries()).sort((a, b) => {
    const partsA = a[0].split('.').map(Number)
    const partsB = b[0].split('.').map(Number)
    for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
      const diff = (partsA[i] || 0) - (partsB[i] || 0)
      if (diff !== 0) return diff
    }
    return 0
  })

  // Calculate grand total
  const grandTotal = sortedSubItems.reduce((sum, [, group]) => sum + group.total, 0)

  // Format the display financial type name
  const displayFinType = targetFinancialType || parsed.financialType
  const displayFinTypeCapitalized = displayFinType.charAt(0).toUpperCase() + displayFinType.slice(1)

  // Build response
  let response = `## Total: ${parsed.itemDisplayName} (${displayFinTypeCapitalized})\n\n`
  response += `**Parent Item:** ${parentCode} - ${parsed.itemDisplayName}\n`
  response += `**Financial Type:** ${displayFinTypeCapitalized}\n`
  response += `**Sheet:** ${sheetName}\n`
  if (parsed.month) {
    const monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    response += `**Month:** ${monthNames[parseInt(parsed.month)] || parsed.month}`
    if (parsed.year) response += ` ${parsed.year}`
    response += '\n'
  }
  response += '\n'

  // Build table
  response += '| Sub-Item | Description | Value |\n'
  response += '|----------|-------------|-------|\n'

  for (const [itemCode, group] of sortedSubItems) {
    const description = group.dataType || '-'
    response += `| ${itemCode} | ${description} | ${formatCurrency(group.total)} |\n`
  }

  response += `\n**Total: ${formatCurrency(grandTotal)}** ('000)\n`

  // Build candidates for drill-down
  const candidates = sortedSubItems.map(([itemCode, group], i) => ({
    id: i + 1,
    value: group.total,
    score: 100,
    sheet: sheetName,
    financialType: displayFinType,
    dataType: group.dataType,
    itemCode: itemCode,
    month: parsed.month || defaultMonth,
    year: parsed.year || '',
    matchedKeywords: [parsed.itemDisplayName || '', displayFinType]
  }))

  return { text: response, candidates }
}

// ============================================
// Analyze & Detail Query Handlers
// ============================================

// Session cache for analysis results (keyed by project name, with TTL)
interface AnalysisItem {
  subIndex: number        // e.g., 1, 2, 3... within the comparison
  itemCode: string        // e.g., "1.1", "2.3"
  itemName: string        // e.g., "Preliminaries"
  value1: number          // first value (e.g., Projected)
  value2: number          // second value (e.g., Business Plan)
  label1: string          // e.g., "Projected"
  label2: string          // e.g., "Business Plan"
  difference: number
  percentage: number
}

interface AnalysisComparison {
  comparisonIndex: number // 1-6
  title: string           // e.g., "Projection vs Business Plan - Income Shortfalls"
  category: 'income' | 'cost'
  operator: 'lt' | 'gt'   // lt = value1 < value2, gt = value1 > value2
  finType1: string        // actual Financial_Type name for value1
  finType2: string        // actual Financial_Type name for value2
  items: AnalysisItem[]
}

interface AnalysisResult {
  comparisons: AnalysisComparison[]
  timestamp: number
  project: string
}

// In-memory cache: project -> analysis result (with 30-minute TTL)
const analysisCache = new Map<string, AnalysisResult>()
const ANALYSIS_CACHE_TTL = 30 * 60 * 1000 // 30 minutes

function isAnalyzeQuery(question: string): boolean {
  const lowerQ = question.toLowerCase().trim()
  return lowerQ === 'analyze' || lowerQ === 'analyse' || 
         lowerQ.startsWith('analyze ') || lowerQ.startsWith('analyse ')
}

function isDetailQuery(question: string): boolean {
  const lowerQ = question.toLowerCase().trim()
  return /^detail\s+\d+(\.\d+)?$/i.test(lowerQ)
}

function parseDetailQuery(question: string): { x: number; y?: number } | null {
  const lowerQ = question.toLowerCase().trim()
  const match = lowerQ.match(/^detail\s+(\d+)(?:\.(\d+))?$/)
  if (!match) return null
  const x = parseInt(match[1])
  const y = match[2] ? parseInt(match[2]) : undefined
  return { x, y }
}

// Resolve a Financial_Type keyword to actual Financial_Type value in the data
// Uses explicit mappings first, then falls back to fuzzy matching
function resolveFinancialType(data: FinancialRow[], project: string, keyword: string): string | null {
  const projectData = data.filter(d => d._project === project && d.Sheet_Name === 'Financial Status')
  const availableTypes = Array.from(new Set(projectData.map(d => d.Financial_Type).filter(Boolean)))
  
  const keywordLower = keyword.toLowerCase()
  
  // === Explicit keyword-to-Financial_Type mappings ===
  // These take highest priority to avoid ambiguous matches
  const EXPLICIT_FINANCIAL_TYPE_MAP: Record<string, string[]> = {
    'business plan':       ['Business Plan'],
    'bp':                  ['Business Plan'],
    'budget':              ['Business Plan'],
    'revision':            ['Revision as at'],
    'revision as at':      ['Revision as at'],
    'budget revision':     ['Revision as at'],
    'rev':                 ['Revision as at'],
    '1st working budget':  ['1st Working Budget'],
    'first working budget':['1st Working Budget'],
    'working budget':      ['1st Working Budget'],
    'projection':          ['Projection as at'],
    'projected':           ['Projection as at'],
    'projection as at':    ['Projection as at'],
    'tender':              ['Budget Tender'],
    'budget tender':       ['Budget Tender'],
    'audit report':        ['Audit Report (WIP)'],
    'wip':                 ['Audit Report (WIP)'],
    'audit report (wip)':  ['Audit Report (WIP)'],
    'committed':           ['Committed Value / Cost as at'],
    'committed value':     ['Committed Value / Cost as at'],
    'committed cost':      ['Committed Value / Cost as at'],
    'accrual':             ['Accrual'],
    'accrued':             ['Accrual'],
    'cash flow':           ['Cash Flow Actual received & paid as at'],
    'cashflow':            ['Cash Flow Actual received & paid as at'],
    'cf':                  ['Cash Flow Actual received & paid as at'],
    'general':             ['General'],
    'adjustment':          ['Adjustment Cost/ variation', 'Adjustment Cost / Variation k=I-J'],
    'variation':           ['Adjustment Cost/ variation', 'Adjustment Cost / Variation k=I-J'],
    'balance':             ['Balance'],
    'balance to':          ['Balance to'],
  }
  
  // Check explicit map first
  const explicitCandidates = EXPLICIT_FINANCIAL_TYPE_MAP[keywordLower]
  if (explicitCandidates) {
    for (const candidate of explicitCandidates) {
      // Find an available type that matches (case-insensitive)
      const match = availableTypes.find(ft => ft.toLowerCase() === candidate.toLowerCase())
      if (match) return match
      // Also try contains match for slight naming variations in data
      const containsMatch = availableTypes.find(ft => ft.toLowerCase().includes(candidate.toLowerCase()))
      if (containsMatch) return containsMatch
    }
  }
  
  // Fallback: Try exact match against available types
  let match = availableTypes.find(ft => ft.toLowerCase() === keywordLower)
  if (match) return match
  
  // Fallback: Try contains match
  match = availableTypes.find(ft => ft.toLowerCase().includes(keywordLower))
  if (match) return match
  
  // Fallback: Try reverse contains
  match = availableTypes.find(ft => keywordLower.includes(ft.toLowerCase()))
  if (match) return match
  
  return null
}

// Get all 2nd tier items under a parent (e.g., parentCode="1" → 1.1, 1.2, etc.)
function getSecondTierItems(data: FinancialRow[], project: string, parentCode: string): Array<{ itemCode: string; itemName: string }> {
  const projectData = data.filter(d => 
    d._project === project && 
    d.Sheet_Name === 'Financial Status'
  )
  
  const prefix = parentCode + '.'
  const items = new Map<string, string>()
  
  for (const row of projectData) {
    if (row.Item_Code.startsWith(prefix)) {
      // Check it's exactly 2nd tier (e.g., "1.1" not "1.1.1")
      const afterPrefix = row.Item_Code.slice(prefix.length)
      if (!afterPrefix.includes('.') && afterPrefix.length > 0) {
        if (!items.has(row.Item_Code)) {
          items.set(row.Item_Code, row.Data_Type || row.Item_Code)
        }
      }
    }
  }
  
  // Sort numerically
  return Array.from(items.entries())
    .sort((a, b) => {
      const partsA = a[0].split('.').map(Number)
      const partsB = b[0].split('.').map(Number)
      for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
        const diff = (partsA[i] || 0) - (partsB[i] || 0)
        if (diff !== 0) return diff
      }
      return 0
    })
    .map(([code, name]) => ({ itemCode: code, itemName: name }))
}

// Get all 3rd tier items under a 2nd tier parent (e.g., parentCode="2.1" → 2.1.1, 2.1.2, etc.)
function getThirdTierItems(data: FinancialRow[], project: string, parentCode: string): Array<{ itemCode: string; itemName: string }> {
  const projectData = data.filter(d => 
    d._project === project && 
    d.Sheet_Name === 'Financial Status'
  )
  
  const prefix = parentCode + '.'
  const items = new Map<string, string>()
  
  for (const row of projectData) {
    if (row.Item_Code.startsWith(prefix)) {
      // Check it's exactly 3rd tier (e.g., "2.1.1" not "2.1.1.1")
      const afterPrefix = row.Item_Code.slice(prefix.length)
      if (!afterPrefix.includes('.') && afterPrefix.length > 0) {
        if (!items.has(row.Item_Code)) {
          items.set(row.Item_Code, row.Data_Type || row.Item_Code)
        }
      }
    }
  }
  
  return Array.from(items.entries())
    .sort((a, b) => {
      const partsA = a[0].split('.').map(Number)
      const partsB = b[0].split('.').map(Number)
      for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
        const diff = (partsA[i] || 0) - (partsB[i] || 0)
        if (diff !== 0) return diff
      }
      return 0
    })
    .map(([code, name]) => ({ itemCode: code, itemName: name }))
}

// Get a value for a specific item code and financial type from Financial Status sheet
function getFinancialStatusValue(data: FinancialRow[], project: string, itemCode: string, financialType: string): number {
  const rows = data.filter(d => 
    d._project === project && 
    d.Sheet_Name === 'Financial Status' && 
    d.Item_Code === itemCode && 
    d.Financial_Type === financialType
  )
  return rows.reduce((sum, d) => sum + toNumber(d.Value), 0)
}

// Run a comparison between two financial types for a set of items
function runComparison(
  data: FinancialRow[], 
  project: string, 
  items: Array<{ itemCode: string; itemName: string }>,
  finType1: string, 
  finType2: string, 
  operator: 'lt' | 'gt',
  label1: string,
  label2: string
): AnalysisItem[] {
  const results: AnalysisItem[] = []
  let subIndex = 0
  
  for (const item of items) {
    const val1 = getFinancialStatusValue(data, project, item.itemCode, finType1)
    const val2 = getFinancialStatusValue(data, project, item.itemCode, finType2)
    
    // Skip if both are zero or either is missing
    if (val1 === 0 && val2 === 0) continue
    
    const meetsCondition = operator === 'lt' ? val1 < val2 : val1 > val2
    
    if (meetsCondition) {
      subIndex++
      const diff = Math.abs(val1 - val2)
      const pct = val2 !== 0 ? (diff / Math.abs(val2)) * 100 : 0
      
      results.push({
        subIndex,
        itemCode: item.itemCode,
        itemName: item.itemName,
        value1: val1,
        value2: val2,
        label1,
        label2,
        difference: diff,
        percentage: pct
      })
    }
  }
  
  return results
}

// Handle "Analyze" query
function handleAnalyzeQuery(data: FinancialRow[], project: string): FuzzyResult {
  const projectData = data.filter(d => d._project === project && d.Sheet_Name === 'Financial Status')
  
  if (projectData.length === 0) {
    return { text: '❌ No Financial Status data found for this project.', candidates: [] }
  }
  
  // Discover actual Financial_Type names from data
  const availableTypes = Array.from(new Set(projectData.map(d => d.Financial_Type).filter(Boolean)))
  
  // Try to resolve each needed Financial_Type
  // IMPORTANT: "Business Plan" maps to the actual "Business Plan" type, NOT "1st Working Budget"
  const projectionType = resolveFinancialType(data, project, 'projection')
  const businessPlanType = resolveFinancialType(data, project, 'business plan')
  const wipType = resolveFinancialType(data, project, 'wip')
  const committedType = resolveFinancialType(data, project, 'committed')
  // "Budget Revision" maps to "Revision as at" in the actual data
  const budgetRevisionType = resolveFinancialType(data, project, 'revision')
  
  // Get 2nd tier items
  const incomeItems = getSecondTierItems(data, project, '1')  // 1.1, 1.2, etc.
  const costItems = getSecondTierItems(data, project, '2')    // 2.1, 2.2, etc.
  
  const comparisons: AnalysisComparison[] = []
  let compIndex = 0
  
  // === INCOME COMPARISONS (1.x) ===
  
  // Comparison 1: Projection vs Business Plan - Income Shortfalls (Projected < Business Plan)
  if (projectionType && businessPlanType) {
    compIndex++
    const items = runComparison(data, project, incomeItems, projectionType, businessPlanType, 'lt', 'Projected', 'Business Plan')
    comparisons.push({
      comparisonIndex: compIndex,
      title: 'Projection vs Business Plan - Income Shortfalls',
      category: 'income',
      operator: 'lt',
      finType1: projectionType,
      finType2: businessPlanType,
      items
    })
  }
  
  // Comparison 2: Projection vs Audit Report (WIP) - Income Shortfalls (Projected < WIP)
  if (projectionType && wipType) {
    compIndex++
    const items = runComparison(data, project, incomeItems, projectionType, wipType, 'lt', 'Projected', 'WIP')
    comparisons.push({
      comparisonIndex: compIndex,
      title: 'Projection vs Audit Report (WIP) - Income Shortfalls',
      category: 'income',
      operator: 'lt',
      finType1: projectionType,
      finType2: wipType,
      items
    })
  }
  
  // === COST COMPARISONS (2.x) ===
  
  // Comparison 3: Projection vs Business Plan - Cost Overruns (Projected > Business Plan)
  if (projectionType && businessPlanType) {
    compIndex++
    const items = runComparison(data, project, costItems, projectionType, businessPlanType, 'gt', 'Projected', 'Business Plan')
    comparisons.push({
      comparisonIndex: compIndex,
      title: 'Projection vs Business Plan - Cost Overruns',
      category: 'cost',
      operator: 'gt',
      finType1: projectionType,
      finType2: businessPlanType,
      items
    })
  }
  
  // Comparison 4: Projection vs Audit Report (WIP) - Cost Overruns (Projected > WIP)
  if (projectionType && wipType) {
    compIndex++
    const items = runComparison(data, project, costItems, projectionType, wipType, 'gt', 'Projected', 'WIP')
    comparisons.push({
      comparisonIndex: compIndex,
      title: 'Projection vs Audit Report (WIP) - Cost Overruns',
      category: 'cost',
      operator: 'gt',
      finType1: projectionType,
      finType2: wipType,
      items
    })
  }
  
  // Comparison 5: Committed vs Projection - Committed Exceeds Projection
  if (committedType && projectionType) {
    compIndex++
    const items = runComparison(data, project, costItems, committedType, projectionType, 'gt', 'Committed', 'Projected')
    comparisons.push({
      comparisonIndex: compIndex,
      title: 'Committed vs Projection - Committed Exceeds Projection',
      category: 'cost',
      operator: 'gt',
      finType1: committedType,
      finType2: projectionType,
      items
    })
  }
  
  // Comparison 6: Projection vs Budget Revision - Exceeds Budget Revision
  if (budgetRevisionType && projectionType) {
    compIndex++
    const items = runComparison(data, project, costItems, projectionType, budgetRevisionType, 'gt', 'Projected', 'Budget Revision')
    comparisons.push({
      comparisonIndex: compIndex,
      title: 'Projection vs Budget Revision - Exceeds Budget Revision',
      category: 'cost',
      operator: 'gt',
      finType1: projectionType,
      finType2: budgetRevisionType,
      items
    })
  }
  
  // Store in cache
  const analysisResult: AnalysisResult = {
    comparisons,
    timestamp: Date.now(),
    project
  }
  analysisCache.set(project, analysisResult)
  
  // Build formatted output
  let response = `## Financial Analysis\n\n`
  
  // Show which Financial Types were detected
  response += `**Data Source:** Financial Status (cumulative)\n`
  response += `**Financial Types Detected:**\n`
  if (projectionType) response += `• Projection: "${projectionType}"\n`
  if (businessPlanType) response += `• Business Plan: "${businessPlanType}"\n`
  if (wipType) response += `• Audit Report (WIP): "${wipType}"\n`
  if (committedType) response += `• Committed: "${committedType}"\n`
  if (budgetRevisionType) response += `• Budget Revision: "${budgetRevisionType}"\n`
  
  // Show warnings for missing types
  const missingTypes: string[] = []
  if (!projectionType) missingTypes.push('Projection')
  if (!businessPlanType) missingTypes.push('Business Plan')
  if (!wipType) missingTypes.push('Audit Report (WIP)')
  if (!committedType) missingTypes.push('Committed')
  if (!budgetRevisionType) missingTypes.push('Budget Revision')
  
  if (missingTypes.length > 0) {
    response += `\n⚠️ **Not found:** ${missingTypes.join(', ')}\n`
    response += `Available types: ${availableTypes.join(', ')}\n`
  }
  
  response += `\n`
  
  // Income section
  const incomeComparisons = comparisons.filter(c => c.category === 'income')
  if (incomeComparisons.length > 0) {
    response += `### Income Analysis (Item 1.x)\n\n`
    for (const comp of incomeComparisons) {
      response += `**${comp.comparisonIndex}. ${comp.title}**\n`
      if (comp.items.length === 0) {
        response += `   ✅ No shortfalls found.\n`
      } else {
        for (const item of comp.items) {
          const arrow = comp.operator === 'lt' ? '↓' : '↑'
          const sign = comp.operator === 'lt' ? '-' : '+'
          response += `   ${comp.comparisonIndex}.${item.subIndex} ${item.itemName}: ${item.label1} ${formatCurrency(item.value1)} < ${item.label2} ${formatCurrency(item.value2)} (${arrow}${formatCurrency(item.difference)}, ${sign}${item.percentage.toFixed(1)}%)\n`
        }
      }
      response += `\n`
    }
  }
  
  // Cost section
  const costComparisons = comparisons.filter(c => c.category === 'cost')
  if (costComparisons.length > 0) {
    response += `### Cost Analysis (Item 2.x)\n\n`
    for (const comp of costComparisons) {
      response += `**${comp.comparisonIndex}. ${comp.title}**\n`
      if (comp.items.length === 0) {
        response += `   ✅ No issues found.\n`
      } else {
        for (const item of comp.items) {
          const arrow = '↑'
          const sign = '+'
          response += `   ${comp.comparisonIndex}.${item.subIndex} ${item.itemName}: ${item.label1} ${formatCurrency(item.value1)} > ${item.label2} ${formatCurrency(item.value2)} (${arrow}${formatCurrency(item.difference)}, ${sign}${item.percentage.toFixed(1)}%)\n`
        }
      }
      response += `\n`
    }
  }
  
  // Summary
  const totalIssues = comparisons.reduce((sum, c) => sum + c.items.length, 0)
  response += `---\n`
  response += `**Summary:** ${totalIssues} issue(s) found across ${comparisons.length} comparisons.\n`
  response += `💡 Type **"Detail X"** to drill down (e.g., "Detail 3" for Cost Overruns).\n`
  response += `💡 Type **"Detail X.Y"** for specific item details (e.g., "Detail 3.1").\n`
  
  return { text: response, candidates: [] }
}

// Handle "Detail X" or "Detail X.Y" query
function handleDetailQuery(data: FinancialRow[], project: string, question: string): FuzzyResult | null {
  const parsed = parseDetailQuery(question)
  if (!parsed) return null
  
  // Check cache
  const cached = analysisCache.get(project)
  if (!cached || (Date.now() - cached.timestamp > ANALYSIS_CACHE_TTL)) {
    return { 
      text: '❌ No analysis results found. Please run **"Analyze"** first to generate the analysis.', 
      candidates: [] 
    }
  }
  
  // Find the comparison
  const comparison = cached.comparisons.find(c => c.comparisonIndex === parsed.x)
  if (!comparison) {
    const availableIndices = cached.comparisons.map(c => c.comparisonIndex).join(', ')
    return { 
      text: `❌ Comparison #${parsed.x} not found. Available comparisons: ${availableIndices}`, 
      candidates: [] 
    }
  }
  
  if (parsed.y !== undefined) {
    // Detail X.Y - Show 3rd tier items for a specific 2nd tier item
    const analysisItem = comparison.items.find(item => item.subIndex === parsed.y)
    if (!analysisItem) {
      const availableItems = comparison.items.map(i => `${parsed.x}.${i.subIndex}`).join(', ')
      return { 
        text: `❌ Item ${parsed.x}.${parsed.y} not found in comparison #${parsed.x}. Available: ${availableItems}`, 
        candidates: [] 
      }
    }
    
    // Get 3rd tier items under this 2nd tier item
    const thirdTierItems = getThirdTierItems(data, project, analysisItem.itemCode)
    
    let response = `## Detail ${parsed.x}.${parsed.y}: ${analysisItem.itemName} - ${comparison.title}\n\n`
    response += `**Parent:** ${analysisItem.itemCode} - ${analysisItem.itemName}\n\n`
    
    // Build table
    response += `| 3rd Tier Item | ${comparison.items[0]?.label1 || 'Value 1'} | ${comparison.items[0]?.label2 || 'Value 2'} | Difference |\n`
    response += `|---------------|-----------|---------------|------------|\n`
    
    let totalDiff = 0
    let foundItems = 0
    
    for (const tier3 of thirdTierItems) {
      const val1 = getFinancialStatusValue(data, project, tier3.itemCode, comparison.finType1)
      const val2 = getFinancialStatusValue(data, project, tier3.itemCode, comparison.finType2)
      
      if (val1 === 0 && val2 === 0) continue
      
      const meetsCondition = comparison.operator === 'lt' ? val1 < val2 : val1 > val2
      if (!meetsCondition) continue
      
      foundItems++
      const diff = Math.abs(val1 - val2)
      const pct = val2 !== 0 ? (diff / Math.abs(val2)) * 100 : 0
      const arrow = comparison.operator === 'lt' ? '↓' : '↑'
      const sign = comparison.operator === 'lt' ? '-' : '+'
      totalDiff += diff
      
      response += `| ${tier3.itemCode} ${tier3.itemName} | ${formatCurrency(val1)} | ${formatCurrency(val2)} | ${arrow}${formatCurrency(diff)} (${sign}${pct.toFixed(1)}%) |\n`
    }
    
    if (foundItems === 0) {
      response += `| (no matching 3rd tier items) | - | - | - |\n`
    }
    
    const totalArrow = comparison.operator === 'lt' ? '↓' : '↑'
    response += `\n**Total ${comparison.operator === 'lt' ? 'Shortfall' : 'Overrun'}:** ${totalArrow}${formatCurrency(totalDiff)}\n`
    
    return { text: response, candidates: [] }
    
  } else {
    // Detail X - Show all 2nd tier items with their 3rd tier breakdown
    let response = `## Detail ${parsed.x}: ${comparison.title}\n\n`
    
    if (comparison.items.length === 0) {
      response += `✅ No issues found in this comparison.\n`
      return { text: response, candidates: [] }
    }
    
    for (const analysisItem of comparison.items) {
      response += `### ${parsed.x}.${analysisItem.subIndex} ${analysisItem.itemName} (${analysisItem.itemCode})\n`
      
      // Get 3rd tier items
      const thirdTierItems = getThirdTierItems(data, project, analysisItem.itemCode)
      
      let foundThirdTier = false
      for (const tier3 of thirdTierItems) {
        const val1 = getFinancialStatusValue(data, project, tier3.itemCode, comparison.finType1)
        const val2 = getFinancialStatusValue(data, project, tier3.itemCode, comparison.finType2)
        
        if (val1 === 0 && val2 === 0) continue
        
        const meetsCondition = comparison.operator === 'lt' ? val1 < val2 : val1 > val2
        if (!meetsCondition) continue
        
        foundThirdTier = true
        const diff = Math.abs(val1 - val2)
        const pct = val2 !== 0 ? (diff / Math.abs(val2)) * 100 : 0
        const arrow = comparison.operator === 'lt' ? '↓' : '↑'
        const sign = comparison.operator === 'lt' ? '-' : '+'
        const symbol = comparison.operator === 'lt' ? '<' : '>'
        
        response += `   - ${tier3.itemCode} ${tier3.itemName}: ${analysisItem.label1} ${formatCurrency(val1)} ${symbol} ${analysisItem.label2} ${formatCurrency(val2)} (${arrow}${formatCurrency(diff)}, ${sign}${pct.toFixed(1)}%)\n`
      }
      
      if (!foundThirdTier) {
        response += `   - (no 3rd tier breakdown available)\n`
      }
      
      response += `\n`
    }
    
    response += `💡 Type **"Detail ${parsed.x}.Y"** for a table view of a specific item.\n`
    
    return { text: response, candidates: [] }
  }
}

// Main query handler with new logic
function answerQuestion(data: FinancialRow[], project: string, question: string, defaultMonth: string): FuzzyResult {
  const expandedQuestion = expandAcronyms(question).toLowerCase()
  // Get significant words from the question (after acronym expansion)
  // IMPORTANT: Keep ALL words including short ones like "np" which may be acronyms
  const questionWords = expandedQuestion.split(/\s+/).filter(w => w.length > 0)
  const projectData = data.filter(d => d._project === project)

  if (projectData.length === 0) {
    return { text: 'No data found for this project.', candidates: [] }
  }

  // Step 0: Check if this is an Analyze or Detail query — highest priority
  if (isAnalyzeQuery(question)) {
    return handleAnalyzeQuery(data, project)
  }
  
  if (isDetailQuery(question)) {
    const detailResult = handleDetailQuery(data, project, question)
    if (detailResult) return detailResult
  }

  // Step 0a: Check if this is a Total query — handle with dedicated logic
  const totalResult = handleTotalQuery(data, project, question, defaultMonth)
  if (totalResult) return totalResult

  // Step 0b: Check if this is a comparison query — handle it with dedicated logic
  const comparisonResult = handleComparisonQuery(data, project, question, defaultMonth)
  if (comparisonResult) return comparisonResult

  // Step 1: Parse date → month, year
  const parsedDate = parseDate(expandedQuestion, defaultMonth)

  // Step 2: Check Sheet_Name
  // IF user didn't specify any date (only using defaults): default to "Financial Status"
  // IF user specified a date: check if user mentioned a specific sheet
  let targetSheet: string | undefined
  const sheets = getUniqueValues(data, project, 'Sheet_Name')

  // Detect if user actually specified a date (not just using defaults)
  const userSpecifiedYear = parsedDate.year && parsedDate.year !== String(new Date().getFullYear())
  const userSpecifiedMonth = parsedDate.month && parsedDate.month !== defaultMonth
  const hasUserDate = userSpecifiedYear || userSpecifiedMonth

  if (!hasUserDate) {
    // No date specified by user → Default to Financial Status, skip sheet detection
    targetSheet = 'Financial Status'
  } else {
    // Date specified by user → Check if user explicitly mentioned a sheet
    // First try: check if user explicitly mentioned a sheet
    for (const sheet of sheets) {
      const sheetLower = sheet.toLowerCase()
      // Check if sheet name appears in question (with or without "sheet")
      if (expandedQuestion.includes(sheetLower.replace(/\s+/g, '')) ||
          expandedQuestion.includes(sheetLower.split(' ')[0])) {
        targetSheet = sheet
        break
      }
    }

    // Second: check for common sheet name keywords even without "sheet" prefix
    if (!targetSheet) {
      const sheetKeywords: Record<string, string> = {
        'cashflow': 'Cash Flow',
        'cash flow': 'Cash Flow',
        'projection': 'Projection',
        'committed': 'Committed Cost',
        'accrual': 'Accrual',
        'financial status': 'Financial Status',
        'financial': 'Financial Status'
      }
      for (const [keyword, sheetName] of Object.entries(sheetKeywords)) {
        // Check if keyword is in expanded question (as standalone word/phrase)
        const keywordRegex = new RegExp(`\\b${keyword.replace(/\s+/g, '\\s+')}\\b`, 'i')
        if (keywordRegex.test(expandedQuestion)) {
          // Verify this sheet actually exists in the data (handle whitespace/blank issues)
          const found = sheets.find(s => s && s.trim() && s.trim().toLowerCase() === sheetName.toLowerCase())
          if (found) {
            targetSheet = found
            break
          }
        }
      }
    }
  }

  // Step 3: Get unique Financial_Type and Data_Type from data
  const financialTypes = getUniqueValues(data, project, 'Financial_Type')
  const dataTypes = getUniqueValues(data, project, 'Data_Type')
  
  // Step 4: Extract Financial_Type from question (find closest match)
  // IMPORTANT: "projected" should map to Financial_Type like "Projection as at"
  // We should NOT skip Financial_Type just because it contains a Sheet_Name
  let targetFinType: string | undefined
  let targetDataType: string | undefined
  for (const ft of financialTypes) {
    const ftLower = ft.toLowerCase()
    // Check if any question word matches any Financial_Type word
    // IMPORTANT: Keep short words like "np" which may be acronyms
    const ftWords = ftLower.split(/\s+/).filter(w => w.length > 0)
    for (const qWord of questionWords) {
      for (const ftWord of ftWords) {
        // Exact word match OR word contains substring with 50%+ threshold
        if (qWord === ftWord) {
          targetFinType = ft
          break
        }
        if (ftWord.includes(qWord) && qWord.length >= ftWord.length * 0.5) {
          targetFinType = ft
          break
        }
      }
      if (targetFinType) break
    }
    if (targetFinType) break
  }
  // If no match found, use fuzzy matching
  if (!targetFinType) {
    for (const word of questionWords) {
      const match = findClosestMatch(word, financialTypes)
      if (match) {
        targetFinType = match
        break
      }
    }
  }

  // Step 5: Extract Data_Type from question (find closest match)
  // IMPORTANT: "gp" / "np" should map to Data_Type like "Gross Profit" / "Net Profit"

  // Special mapping for common acronyms (must come first!)
  const acronymMap: Record<string, string[]> = {
    'np': ['net profit', 'acc. net profit'],
    'gp': ['gross profit', 'acc. gross profit'],
    'wip': ['work in progress'],
    'cf': ['cash flow']
  }

  // Check if question contains any known acronyms
  for (const [acronym, expansions] of Object.entries(acronymMap)) {
    if (questionWords.includes(acronym)) {
      // Try to find a matching Data_Type in the data
      for (const expansion of expansions) {
        const match = dataTypes.find(dt => dt.toLowerCase().includes(expansion))
        if (match) {
          targetDataType = match
          break
        }
      }
      if (targetDataType) break
    }
  }

  // If no acronym match found, continue with regular matching
  let bestDataTypeMatchCount = 0

  if (!targetDataType) {
    for (const dt of dataTypes) {
      const dtLower = dt.toLowerCase()
      const dtWords = dtLower.split(/\s+/).filter(w => w.length > 0)

      // Count how many question words match this Data_Type
      let matchCount = 0
      const matchedWords: string[] = []
      for (const qWord of questionWords) {
        for (const dtWord of dtWords) {
          if (qWord === dtWord) {
            matchCount++
            matchedWords.push(qWord)
            break
          }
        }
      }

      // Partial matches for longer words (4+ chars)
      for (const qWord of questionWords) {
        if (matchedWords.includes(qWord)) continue
        if (qWord.length <= 3) continue

        for (const dtWord of dtWords) {
          const qLen = qWord.length
          const dLen = dtWord.length

          const longer = qLen >= dLen ? qWord : dtWord
          const shorter = qLen >= dLen ? dtWord : qWord

          if (longer.includes(shorter) && shorter.length >= longer.length * 0.5) {
            matchCount++
            matchedWords.push(qWord)
            break
          }
        }
      }

      if (matchCount > bestDataTypeMatchCount) {
        bestDataTypeMatchCount = matchCount
        targetDataType = dt
      }
    }
  }

  // If no match found, use fuzzy matching with all significant words
  if (!targetDataType) {
    for (const word of questionWords) {
      const match = findClosestMatch(word, dataTypes)
      if (match) {
        targetDataType = match
        break
      }
    }
  }

  // Track which sheet was actually applied (for display)
  let appliedSheet = targetSheet

  // Build filter conditions
  let filtered = projectData

  // Apply Sheet_Name filter
  if (targetSheet) {
    filtered = filtered.filter(d => d.Sheet_Name === targetSheet)
    appliedSheet = targetSheet
  }

  // Apply month filter (if specified)
  if (parsedDate.month) {
    filtered = filtered.filter(d => d.Month === parsedDate.month)
  }

  // Apply year filter (if specified)
  if (parsedDate.year) {
    filtered = filtered.filter(d => d.Year === parsedDate.year)
  }

  // Apply Financial_Type filter (if found)
  if (targetFinType) {
    filtered = filtered.filter(d => d.Financial_Type === targetFinType)
  }

  // Apply Data_Type filter (if found)
  if (targetDataType) {
    filtered = filtered.filter(d => d.Data_Type === targetDataType)
  }

  // If no exact matches, relax filters progressively
  if (filtered.length === 0) {
    // Try without Financial_Type
    filtered = projectData
    if (targetSheet) filtered = filtered.filter(d => d.Sheet_Name === targetSheet)
    if (parsedDate.month) filtered = filtered.filter(d => d.Month === parsedDate.month)
    if (parsedDate.year) filtered = filtered.filter(d => d.Year === parsedDate.year)
    if (targetDataType) filtered = filtered.filter(d => d.Data_Type === targetDataType)
    appliedSheet = targetSheet || 'Financial Status'
  }

  if (filtered.length === 0) {
    // Try without Data_Type
    filtered = projectData
    if (targetSheet) filtered = filtered.filter(d => d.Sheet_Name === targetSheet)
    if (parsedDate.month) filtered = filtered.filter(d => d.Month === parsedDate.month)
    if (parsedDate.year) filtered = filtered.filter(d => d.Year === parsedDate.year)
    appliedSheet = targetSheet || 'Financial Status'
  }

  if (filtered.length === 0) {
    // PHASE 2: Generate helpful suggestions based on available data
    const availableMonths = Array.from(new Set(projectData.map(d => d.Month))).sort()
    const availableDataTypes = Array.from(new Set(projectData.map(d => d.Data_Type).filter(Boolean)))
    const availableFinTypes = Array.from(new Set(projectData.map(d => d.Financial_Type).filter(Boolean)))
    
    let suggestions = `No data found matching your query.\n\nFilters attempted:`
    if (appliedSheet) suggestions += `\n- Sheet: ${appliedSheet}`
    if (parsedDate.month) suggestions += `\n- Month: ${parsedDate.month}`
    if (parsedDate.year) suggestions += `\n- Year: ${parsedDate.year}`
    if (targetFinType) suggestions += `\n- Financial Type: ${targetFinType}`
    if (targetDataType) suggestions += `\n- Data Type: ${targetDataType}`
    
    // Add helpful suggestions
    suggestions += `\n\n💡 Suggestions:`
    if (availableMonths.length > 0) {
      suggestions += `\n• Available months: ${availableMonths.slice(0, 6).join(', ')}${availableMonths.length > 6 ? '...' : ''}`
    }
    if (availableDataTypes.length > 0) {
      suggestions += `\n• Try: "${availableDataTypes[0]}" (most common)`
    }
    suggestions += `\n• Try: "gross profit", "projection", "budget", "cash flow"`
    suggestions += `\n• Or try: "last month", "this month"`
    
    return { text: suggestions, candidates: [] }
  }

  // Helper to convert Value to number safely
  const toNumber = (val: number | string): number => {
    if (typeof val === 'number') return val
    return parseFloat(val) || 0
  }

  // Format results
  const total = filtered.reduce((sum, d) => sum + toNumber(d.Value), 0)

  // Get unique Item_Codes for display
  const itemGroups = new Map<string, FinancialRow[]>()
  filtered.forEach(d => {
    const key = d.Item_Code || 'Unknown'
    if (!itemGroups.has(key)) itemGroups.set(key, [])
    itemGroups.get(key)!.push(d)
  })

  let response = `## Query Results\n\n`
  response += `**Filters:**\n`
  // Show which sheet was used
  if (appliedSheet) response += `• Sheet: ${appliedSheet}\n`
  // Show Financial Type filter
  if (targetFinType) response += `• Financial Type: ${targetFinType}\n`
  // Show month - only show actual month if user specified it
  response += `• Month: ${hasUserDate && parsedDate.month ? parsedDate.month : 'All'}\n`
  // Show year - only show actual year if user specified it
  response += `• Year: ${hasUserDate && parsedDate.year ? parsedDate.year : 'All'}\n`
  response += `• Data Type: ${targetDataType || 'All'}\n`
  response += `• Item Code: all\n\n`

  response += `**Total: ${formatCurrency(total)}** ('000)\n\n`

  response += `**By Item Code:**\n`
  itemGroups.forEach((rows, itemCode) => {
    const itemTotal = rows.reduce((sum, d) => sum + toNumber(d.Value), 0)
    response += `• Item ${itemCode}: ${formatCurrency(itemTotal)}\n`
  })

  // Create candidates for clickable selection
  // ALWAYS score ALL project data records and show top 10 best matches
  // Include keyword matching across ALL fields for comprehensive scoring
  const allCandidates = projectData.map((d) => {
    let matchScore = 0
    const matchedKeywords: string[] = []

    // Build combined text from all searchable fields
    const combinedText = `${d.Sheet_Name} ${d.Financial_Type} ${d.Data_Type} ${d.Item_Code} ${d.Month} ${d.Year}`.toLowerCase()

    // Check each question word against ALL fields
    for (const qWord of questionWords) {
      // Financial_Type match
      if (d.Financial_Type.toLowerCase().includes(qWord)) {
        matchScore += 5
        matchedKeywords.push(qWord)
      }

      // Data_Type match (important!)
      if (d.Data_Type.toLowerCase().includes(qWord)) {
        matchScore += 8
        matchedKeywords.push(qWord)
      }

      // Item_Code match
      if (d.Item_Code.toLowerCase().includes(qWord)) {
        matchScore += 3
        matchedKeywords.push(qWord)
      }

      // Sheet_Name match
      if (d.Sheet_Name.toLowerCase().includes(qWord)) {
        matchScore += 2
        matchedKeywords.push(qWord)
      }
    }

    // Explicit Financial_Type match (high priority)
    if (targetFinType && d.Financial_Type === targetFinType) matchScore += 40
    else if (targetFinType && d.Financial_Type.toLowerCase().includes(targetFinType.toLowerCase())) {
      matchScore += 30
      matchedKeywords.push(targetFinType)
    }

    // Explicit Data_Type match (high priority)
    if (targetDataType && d.Data_Type === targetDataType) matchScore += 35
    else if (targetDataType && d.Data_Type.toLowerCase().includes(targetDataType.toLowerCase())) {
      matchScore += 25
      matchedKeywords.push(targetDataType)
    }

    // Month match
    if (parsedDate.month && d.Month === parsedDate.month) matchScore += 20

    // Year match
    if (parsedDate.year && d.Year === parsedDate.year) matchScore += 15

    // Bonus for common item codes
    if (d.Item_Code === '3' || d.Item_Code === '1' || d.Item_Code === '2') matchScore += 5

    // Bonus for Financial Status (default sheet)
    if (d.Sheet_Name === 'Financial Status') matchScore += 2

    return {
      id: 0, // Will be reassigned
      value: d.Value,
      score: matchScore,
      sheet: d.Sheet_Name,
      financialType: d.Financial_Type,
      dataType: d.Data_Type,
      itemCode: d.Item_Code,
      month: d.Month,
      year: d.Year,
      matchedKeywords: Array.from(new Set(matchedKeywords)) // Remove duplicates
    }
  }).sort((a, b) => b.score - a.score).slice(0, 10)

  // Reassign IDs after sorting
  const candidates = allCandidates.map((c, i) => ({ ...c, id: i + 1 }))

  if (candidates.length > 0) {
    response += `\n**Available Records (click to select):**\n`
    candidates.forEach((c) => {
      const matches = c.matchedKeywords.length > 0 ? ` [Matched: ${c.matchedKeywords.join(', ')}]` : ''
      response += `[${c.id}] ${c.month}/${c.year}/${c.sheet}/${c.financialType}/${c.dataType}/${c.itemCode}: ${formatCurrency(toNumber(c.value))} [Score: ${c.score}]${matches}\n`
    })
  }

  return { text: response, candidates }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json()
    const { action, year, month, project, projectFile, question } = body

    switch (action) {
      case 'getStructure': {
        const result = await getFolderStructure()
        if (result.error) {
          return NextResponse.json({
            folders: result.folders,
            projects: result.projects,
            error: result.error
          })
        }
        return NextResponse.json({
          folders: result.folders,
          projects: result.projects
        })
      }

      case 'loadProject': {
        const data = await loadProjectData(projectFile, year, month)
        const metrics = getProjectMetrics(data, project)

        // Debug info
        const debug = {
          source: `Google Drive: Ai Chatbot Knowledge Base/${year}/${month}/${projectFile}`,
          totalRows: data.length,
          // Show raw values from CSV
          sampleRows: data.slice(0, 3).map(d => ({
            Sheet: d.Sheet_Name,
            FinType: d.Financial_Type,
            Item: d.Item_Code,
            Data: d.Data_Type,
            Value: d.Value
          })),
          uniqueSheets: Array.from(new Set(data.map(d => d.Sheet_Name))),
          uniqueFinancialTypes: Array.from(new Set(data.map(d => d.Financial_Type))),
          uniqueItemCodes: Array.from(new Set(data.map(d => d.Item_Code))),
          uniqueDataTypes: Array.from(new Set(data.map(d => d.Data_Type))),
          gpRowsCount: data.filter(d =>
            d.Item_Code === '3' &&
            d.Data_Type?.toLowerCase().includes('gross profit')
          ).length
        }

        return NextResponse.json({ data, metrics, debug })
      }

      case 'query': {
        const data = await loadProjectData(projectFile, year, month)
        const result = answerQuestion(data, project, question, month)
        return NextResponse.json({ response: result.text, candidates: result.candidates })
      }

      case 'metrics': {
        const data = await loadProjectData(projectFile, year, month)
        const metrics = getProjectMetrics(data, project)
        return NextResponse.json({ metrics })
      }

      default:
        return NextResponse.json({ error: 'Unknown action' }, { status: 400 })
    }
  } catch (error) {
    console.error('API error:', error)
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 })
  }
}
