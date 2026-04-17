import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import { fmtEUR } from './utils'

const TEAL = [1, 105, 111]
const TEAL_SOFT = [206, 220, 216]
const GOLD_SOFT = [233, 196, 106]
const LIGHT = [249, 248, 245]
const LIGHT_2 = [243, 240, 236]
const DARK = [40, 37, 29]
const MUTED = [122, 121, 116]
const BORDER = [212, 209, 202]

const PAGE = {
  w: 210,
  h: 297,
  mx: 14,
  my: 16,
}

function currentYearKey() {
  return String(new Date().getFullYear())
}

function getMonthKey(dateStr) {
  const d = String(dateStr || '')
  return d.length >= 7 ? d.slice(0, 7) : ''
}

function getMonthsRemainingInYear() {
  const now = new Date()
  const year = now.getFullYear()
  const month = now.getMonth() + 1
  const arr = []
  for (let m = month + 1; m <= 12; m++) {
    arr.push(`${year}-${String(m).padStart(2, '0')}`)
  }
  return arr
}

function safeCategory(v) {
  return v || 'Other'
}

function safeMerchant(t) {
  return t.merchant || t.description || 'Unknown'
}

function isCurrentYear(dateStr, year = currentYearKey()) {
  return String(dateStr || '').startsWith(year)
}

function isSubscriptionCandidate(t) {
  const d = String(t.description || '')
  return /netflix|spotify|prime|apple|hbo|disney|subscription|adobe|microsoft|google one|icloud|dropbox|urban sports|revolut|youtube/i.test(d)
}

function getSubscriptionKeyFromDescription(description) {
  const d = String(description || '').toLowerCase()
  const m = d.match(/netflix|spotify|prime video|amazon prime|apple tv|apple music|hbo|disney|paramount|youtube|icloud|dropbox|dazn|patreon|subscription/i)
  if (m) return m[0].toLowerCase()
  return d.trim().split(' ')[0] || 'subscription'
}

function sum(arr, accessor) {
  return arr.reduce((s, item) => s + Number(accessor(item) || 0), 0)
}

function buildOwnTransactions(transactions) {
  return transactions
    .filter(t => t.account !== 'joint')
    .map(t => ({
      ...t,
      amountValue: Math.abs(Number(t.amount) || 0),
    }))
}

function buildJointTransactions(transactions) {
  return transactions
    .filter(t => t.account === 'joint' || t.joint)
    .map(t => {
      const amount = Math.abs(Number(t.amount) || 0)
      let jointAmount = amount
      if (t.account === 'joint') jointAmount = amount
      else if (t.jointMode === 'half') jointAmount = amount / 2
      else jointAmount = amount
      return {
        ...t,
        amountValue: jointAmount,
      }
    })
}

function buildMonthCategoryTable(rows) {
  const months = [...new Set(rows.map(r => getMonthKey(r.date)).filter(Boolean))].sort()
  const categories = [...new Set(rows.map(r => safeCategory(r.category)))].sort((a, b) => a.localeCompare(b))

  const byCatMonth = new Map()
  rows.forEach(r => {
    const month = getMonthKey(r.date)
    const category = safeCategory(r.category)
    if (!month) return
    const key = `${category}__${month}`
    byCatMonth.set(key, (byCatMonth.get(key) || 0) + Number(r.amountValue || 0))
  })

  const body = categories.map(category => {
    const values = months.map(month => byCatMonth.get(`${category}__${month}`) || 0)
    const total = values.reduce((a, b) => a + b, 0)
    return [
      category,
      ...values.map(v => (v ? fmtEUR(v) : '—')),
      fmtEUR(total),
    ]
  })

  const monthTotals = months.map(month =>
    rows
      .filter(r => getMonthKey(r.date) === month)
      .reduce((s, r) => s + Number(r.amountValue || 0), 0)
  )

  body.push([
    'Total',
    ...monthTotals.map(v => (v ? fmtEUR(v) : '—')),
    fmtEUR(monthTotals.reduce((a, b) => a + b, 0)),
  ])

  return {
    head: [['Category', ...months, 'Total']],
    body,
    grandTotal: monthTotals.reduce((a, b) => a + b, 0),
    months,
  }
}

function buildForecastByCategory(rows) {
  const currentYear = currentYearKey()
  const currentMonth = new Date().getMonth() + 1
  const monthsElapsed = currentMonth
  const remainingMonths = 12 - currentMonth

  const categories = [...new Set(rows.map(r => safeCategory(r.category)))].sort((a, b) => a.localeCompare(b))
  const body = categories.map(category => {
    const spentYTD = rows
      .filter(r => safeCategory(r.category) === category && isCurrentYear(r.date, currentYear))
      .reduce((s, r) => s + Number(r.amountValue || 0), 0)

    const avgPerMonth = monthsElapsed > 0 ? spentYTD / monthsElapsed : 0
    const projectedRest = avgPerMonth * remainingMonths
    const projectedFullYear = spentYTD + projectedRest

    return [
      category,
      fmtEUR(spentYTD),
      fmtEUR(avgPerMonth),
      fmtEUR(projectedRest),
      fmtEUR(projectedFullYear),
    ]
  })

  const totalSpentYTD = rows
    .filter(r => isCurrentYear(r.date, currentYear))
    .reduce((s, r) => s + Number(r.amountValue || 0), 0)

  const totalAvg = monthsElapsed > 0 ? totalSpentYTD / monthsElapsed : 0
  const totalProjectedRest = totalAvg * remainingMonths
  const totalProjectedFullYear = totalSpentYTD + totalProjectedRest

  body.push([
    'Total',
    fmtEUR(totalSpentYTD),
    fmtEUR(totalAvg),
    fmtEUR(totalProjectedRest),
    fmtEUR(totalProjectedFullYear),
  ])

  return {
    head: [['Category', 'Spent YTD', 'Avg / Month', 'Forecast to Dec', 'Projected Full Year']],
    body,
    totals: {
      spentYTD: totalSpentYTD,
      forecastToDec: totalProjectedRest,
      projectedFullYear: totalProjectedFullYear,
    },
  }
}

function buildTopMerchants(rows, limit = 12) {
  const m = new Map()
  rows.forEach(r => {
    const key = safeMerchant(r)
    m.set(key, (m.get(key) || 0) + Number(r.amountValue || 0))
  })

  const entries = [...m.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)

  return {
    head: [['Merchant', 'Amount']],
    body: [
      ...entries.map(([merchant, amount]) => [merchant, fmtEUR(amount)]),
      ['Total', fmtEUR(entries.reduce((s, [, v]) => s + v, 0))],
    ],
    rows: entries,
  }
}

function buildActiveSubscriptions(ownRows) {
  const now = new Date()
  const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`
  const candidates = ownRows.filter(isSubscriptionCandidate)
  const groups = new Map()

  for (const t of candidates) {
    const key = getSubscriptionKeyFromDescription(t.description)
    if (!groups.has(key)) groups.set(key, [])
    groups.get(key).push(t)
  }

  const monthDiff = (current, past) => {
    const [cy, cm] = current.split('-').map(Number)
    const [py, pm] = past.split('-').map(Number)
    return (cy - py) * 12 + (cm - pm)
  }

  const active = []
  for (const [key, txs] of groups.entries()) {
    const sorted = txs
      .filter(t => t.date)
      .slice()
      .sort((a, b) => (a.date > b.date ? -1 : 1))

    if (!sorted.length) continue
    const latest = sorted[0]
    const lastMonth = getMonthKey(latest.date)
    if (!lastMonth) continue

    const diff = monthDiff(currentMonthKey, lastMonth)
    if (diff >= 2) continue

    const totalSpent = sum(txs, t => t.amountValue)
    active.push([
      key,
      fmtEUR(totalSpent),
      fmtEUR(latest.amountValue || 0),
      latest.date || '—',
    ])
  }

  active.sort((a, b) => a[0].localeCompare(b[0]))

  return {
    head: [['Subscription', 'Total Spent', 'Last Charge', 'Last Charge Date']],
    body: active.length ? active : [['No active subscriptions detected', '—', '—', '—']],
    count: active.length,
  }
}

function addHeader(doc, title, subtitle) {
  doc.setFillColor(...TEAL)
  doc.rect(0, 0, PAGE.w, 22, 'F')
  doc.setTextColor(255, 255, 255)
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(16)
  doc.text(title, PAGE.mx, 13)
  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.text(subtitle, PAGE.w - PAGE.mx, 13, { align: 'right' })
}

function addSectionTitle(doc, title, y) {
  doc.setTextColor(...DARK)
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.text(title, PAGE.mx, y)
}

function addInsightBox(doc, title, lines, y, fill = LIGHT_2) {
  const boxHeight = 8 + lines.length * 6
  doc.setFillColor(...fill)
  doc.setDrawColor(...BORDER)
  doc.roundedRect(PAGE.mx, y, PAGE.w - PAGE.mx * 2, boxHeight, 2, 2, 'FD')
  doc.setTextColor(...DARK)
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(10)
  doc.text(title, PAGE.mx + 4, y + 6)
  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  lines.forEach((line, i) => {
    doc.text(`• ${line}`, PAGE.mx + 4, y + 12 + i * 6)
  })
  return y + boxHeight + 6
}

function tableTheme() {
  return {
    theme: 'grid',
    styles: {
      font: 'helvetica',
      fontSize: 8,
      textColor: DARK,
      lineColor: BORDER,
      lineWidth: 0.15,
      cellPadding: 2.2,
    },
    headStyles: {
      fillColor: TEAL,
      textColor: 255,
      fontStyle: 'bold',
    },
    bodyStyles: {
      textColor: DARK,
    },
    alternateRowStyles: {
      fillColor: LIGHT,
    },
    margin: {
      left: PAGE.mx,
      right: PAGE.mx,
    },
  }
}

function addTable(doc, title, head, body, startY, opts = {}) {
  addSectionTitle(doc, title, startY)
  autoTable(doc, {
    startY: startY + 3,
    head,
    body,
    ...tableTheme(),
    ...opts,
  })
  return doc.lastAutoTable.finalY + 8
}

function addInsights(doc, label, rows, merchants, forecast, y) {
  const total = sum(rows, r => r.amountValue)
  const monthsSeen = [...new Set(rows.map(r => getMonthKey(r.date)).filter(Boolean))].length
  const topCategoryRow = buildMonthCategoryTable(rows).body
    .slice(0, -1)
    .map(r => ({ category: r[0], total: Number(String(r[r.length - 1]).replace(/[^\d,-]/g, '').replace(',', '.')) || 0 }))

  const topMerchant = merchants.rows[0]
  const lines = [
    `${label} spend in ${currentYearKey()}: ${fmtEUR(total)} across ${rows.length} transactions.`,
    `Months with activity so far: ${monthsSeen}.`,
    topMerchant ? `Top merchant: ${topMerchant[0]} at ${fmtEUR(topMerchant[1])}.` : 'No merchant concentration detected yet.',
    `Projected full-year spend: ${forecast.totals.projectedFullYear ? fmtEUR(forecast.totals.projectedFullYear) : fmtEUR(0)}.`,
  ]
  return addInsightBox(doc, `${label} insights`, lines, y, label === 'Own expenses' ? TEAL_SOFT : [231, 215, 199])
}

export function generatePDFReport(transactions = []) {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' })
  const year = currentYearKey()

  const ownRows = buildOwnTransactions(transactions).filter(t => isCurrentYear(t.date, year))
  const jointRows = buildJointTransactions(transactions).filter(t => isCurrentYear(t.date, year))

  const ownMonthCategory = buildMonthCategoryTable(ownRows)
  const ownForecast = buildForecastByCategory(ownRows)
  const ownMerchants = buildTopMerchants(ownRows, 12)
  const ownSubscriptions = buildActiveSubscriptions(ownRows)

  const jointMonthCategory = buildMonthCategoryTable(jointRows)
  const jointForecast = buildForecastByCategory(jointRows)
  const jointMerchants = buildTopMerchants(jointRows, 12)

  addHeader(doc, 'Expense Report', `Generated ${new Date().toLocaleDateString('pt-PT')} · ${year}`)

  let y = 30
  doc.setTextColor(...DARK)
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(13)
  doc.text('Own Expenses', PAGE.mx, y)
  y += 6

  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.setTextColor(...MUTED)
  doc.text('Personal account overview for the current year, with category, month, forecast, merchants, and active subscriptions.', PAGE.mx, y)
  y += 8

  y = addInsights(doc, 'Own expenses', ownRows, ownMerchants, ownForecast, y)
  y = addTable(doc, `Current Year by Category and Month (${year})`, ownMonthCategory.head, ownMonthCategory.body, y, {
    styles: { fontSize: 7.5 },
  })

  y = addTable(doc, 'Forecast by Category to Year End', ownForecast.head, ownForecast.body, y, {
    styles: { fontSize: 8 },
  })

  y = addTable(doc, 'Top Merchants This Year', ownMerchants.head, ownMerchants.body, y, {
    columnStyles: { 0: { cellWidth: 120 }, 1: { halign: 'right' } },
  })

  if (y > 220) {
    doc.addPage()
    addHeader(doc, 'Expense Report', `Generated ${new Date().toLocaleDateString('pt-PT')} · ${year}`)
    y = 28
  }

  y = addTable(doc, 'Active Subscriptions', ownSubscriptions.head, ownSubscriptions.body, y, {
    columnStyles: { 0: { cellWidth: 70 } },
  })

  doc.addPage()
  addHeader(doc, 'Expense Report', `Generated ${new Date().toLocaleDateString('pt-PT')} · ${year}`)

  y = 30
  doc.setTextColor(...DARK)
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(13)
  doc.text('Joint Expenses', PAGE.mx, y)
  y += 6

  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.setTextColor(...MUTED)
  doc.text('Shared account and Joint-tab overview for the current year, with category, month, forecast, and merchant concentration.', PAGE.mx, y)
  y += 8

  y = addInsights(doc, 'Joint expenses', jointRows, jointMerchants, jointForecast, y)
  y = addTable(doc, `Current Year by Category and Month (${year})`, jointMonthCategory.head, jointMonthCategory.body, y, {
    styles: { fontSize: 7.5 },
  })

  y = addTable(doc, 'Forecast by Category to Year End', jointForecast.head, jointForecast.body, y, {
    styles: { fontSize: 8 },
  })

  y = addTable(doc, 'Top Merchants This Year', jointMerchants.head, jointMerchants.body, y, {
    columnStyles: { 0: { cellWidth: 120 }, 1: { halign: 'right' } },
  })

  doc.save(`expense-report-${year}.pdf`)
}

export function generateSplitPDFReport({
  currentUserName = 'You',
  partnerUserName = 'Partner',
  mySplitTransactions = [],
  partnerSplitTransactions = [],
  mySplitTotal = 0,
  partnerSplitTotal = 0,
  netBalance = 0,
} = {}) {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' })
  const today = new Date().toLocaleDateString('pt-PT')

  const normalizeRows = (rows = []) =>
    rows
      .map((t) => {
        const totalAmount = Math.abs(Number(t.amount) || 0)
        const owedAmount = totalAmount / 2

        return {
          date: t.date || '—',
          description: t.description || t.merchant || '—',
          category: t.category || 'Other',
          totalAmount,
          owedAmount,
        }
      })
      .sort((a, b) => String(a.date).localeCompare(String(b.date)))

  const myRows = normalizeRows(mySplitTransactions)
  const partnerRows = normalizeRows(partnerSplitTransactions)

  const balanceLabel =
    netBalance > 0.004
      ? `${partnerUserName} owes ${currentUserName}`
      : netBalance < -0.004
        ? `${currentUserName} owes ${partnerUserName}`
        : 'Balance is settled'

  const balanceValue = Math.abs(Number(netBalance) || 0)

  addHeader(doc, 'Split Balance Report', `Generated ${today}`)

  let y = 30

  doc.setTextColor(...DARK)
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(13)
  doc.text('Split overview', PAGE.mx, y)
  y += 6

  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.setTextColor(...MUTED)
  doc.text(
    'This report includes your split expenses, your partner split expenses, and the net balance of who owes who.',
    PAGE.mx,
    y,
    { maxWidth: PAGE.w - PAGE.mx * 2 }
  )
  y += 10

  y = addInsightBox(
    doc,
    'Balance summary',
    [
      `${currentUserName} split total: ${fmtEUR(mySplitTotal)}`,
      `${partnerUserName} split total: ${fmtEUR(partnerSplitTotal)}`,
      `${balanceLabel}: ${fmtEUR(balanceValue)}`,
    ],
    y,
    TEAL_SOFT
  )

  y = addTable(
    doc,
    `${currentUserName} split expenses`,
    [['Date', 'Description', 'Category', 'Total Amount', 'Owed Amount']],
    myRows.length
      ? [
          ...myRows.map((row) => [
            row.date,
            row.description,
            row.category,
            fmtEUR(row.totalAmount),
            fmtEUR(row.owedAmount),
          ]),
          ['Total', '', '', '', fmtEUR(mySplitTotal)],
        ]
      : [['—', 'No split transactions found', '—', '—', '—']],
    y,
    {
      styles: { fontSize: 8 },
      columnStyles: {
        0: { cellWidth: 24 },
        1: { cellWidth: 72 },
        2: { cellWidth: 34 },
        3: { halign: 'right', cellWidth: 28 },
        4: { halign: 'right', cellWidth: 28 },
      },
    }
  )

  if (y > 210) {
    doc.addPage()
    addHeader(doc, 'Split Balance Report', `Generated ${today}`)
    y = 28
  }

  y = addTable(
    doc,
    `${partnerUserName} split expenses`,
    [['Date', 'Description', 'Category', 'Total Amount', 'Owed Amount']],
    partnerRows.length
      ? [
          ...partnerRows.map((row) => [
            row.date,
            row.description,
            row.category,
            fmtEUR(row.totalAmount),
            fmtEUR(row.owedAmount),
          ]),
          ['Total', '', '', '', fmtEUR(partnerSplitTotal)],
        ]
      : [['—', 'No partner split transactions found', '—', '—', '—']],
    y,
    {
      styles: { fontSize: 8 },
      columnStyles: {
        0: { cellWidth: 24 },
        1: { cellWidth: 72 },
        2: { cellWidth: 34 },
        3: { halign: 'right', cellWidth: 28 },
        4: { halign: 'right', cellWidth: 28 },
      },
    }
  )

  y = addTable(
    doc,
    'Net balance',
    [['Item', 'Amount']],
    [
      [`${currentUserName} total owed amount`, fmtEUR(mySplitTotal)],
      [`${partnerUserName} total owed amount`, fmtEUR(partnerSplitTotal)],
      [balanceLabel, fmtEUR(balanceValue)],
    ],
    y,
    {
      styles: { fontSize: 9 },
      columnStyles: {
        0: { cellWidth: 130 },
        1: { halign: 'right', cellWidth: 40 },
      },
    }
  )

  doc.save(`split-balance-report-${new Date().toISOString().slice(0, 10)}.pdf`)
}