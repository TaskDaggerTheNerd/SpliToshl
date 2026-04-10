import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import { fmtEUR } from './utils'

const TEAL = [1, 105, 111]
const LIGHT = [249, 248, 245]
const DARK = [40, 37, 29]

const TABLE_STYLES = {
  headStyles: { fillColor: TEAL, textColor: 255, fontStyle: 'bold', fontSize: 8 },
  bodyStyles: { textColor: DARK, fontSize: 8 },
  alternateRowStyles: { fillColor: LIGHT },
  margin: { left: 14, right: 14 },
}

export function generatePDFReport(rows = [], forecast = { series: [], avg: 0 }, splitRows = [], splitTotal = 0) {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' })

  doc.setFillColor(...TEAL)
  doc.rect(0, 0, 210, 20, 'F')
  doc.setFontSize(14)
  doc.setTextColor(255, 255, 255)
  doc.setFont('helvetica', 'bold')
  doc.text('Expense Report', 14, 13)
  doc.setFontSize(9)
  doc.setFont('helvetica', 'normal')
  doc.text(new Date().toLocaleDateString('pt-PT', { year: 'numeric', month: 'long', day: 'numeric' }), 196, 13, { align: 'right' })

  const total = rows.reduce((s, r) => s + Math.abs(Number(r.amount || 0)), 0)
  doc.setTextColor(...DARK)
  doc.setFontSize(9)
  doc.setFont('helvetica', 'bold')
  doc.text('Summary', 14, 30)
  doc.setFont('helvetica', 'normal')
  doc.text(`Total spend: ${fmtEUR(total)}`, 14, 37)
  doc.text(`Transactions: ${rows.length}`, 80, 37)
  doc.text(`Split balance owed to you: ${fmtEUR(splitTotal)}`, 140, 37)

  const byCategory = new Map()
  rows.forEach(r => byCategory.set(r.category, (byCategory.get(r.category) || 0) + Math.abs(Number(r.amount || 0))))
  const catRows = [...byCategory.entries()].sort((a, b) => b[1] - a[1])
    .map(([cat, amt]) => [cat, fmtEUR(amt), ((amt / (total || 1)) * 100).toFixed(1) + '%'])

  doc.setFont('helvetica', 'bold')
  doc.text('Spending by Category', 14, 48)
  autoTable(doc, { startY: 51, head: [['Category', 'Amount', '% of Total']], body: catRows, ...TABLE_STYLES })

  const byMonth = new Map()
  rows.forEach(r => {
    const m = String(r.date || '').slice(0, 7)
    if (m.length === 7) byMonth.set(m, (byMonth.get(m) || 0) + Math.abs(Number(r.amount || 0)))
  })
  const monthRows = [...byMonth.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([m, amt]) => [m, fmtEUR(amt)])

  let y = doc.lastAutoTable?.finalY || 60
  doc.setFont('helvetica', 'bold')
  doc.text('Monthly Spend', 14, y + 10)
  autoTable(doc, { startY: y + 13, head: [['Month', 'Amount']], body: monthRows, ...TABLE_STYLES })

  y = doc.lastAutoTable?.finalY || y + 20
  doc.setFont('helvetica', 'bold')
  doc.text('3-Month Forecast', 14, y + 10)
  autoTable(doc, {
    startY: y + 13,
    head: [['Month', 'Actual', 'Projected']],
    body: forecast.series.map(r => [r.month, r.actual != null ? fmtEUR(r.actual) : '—', r.projected != null ? fmtEUR(r.projected) : '—']),
    ...TABLE_STYLES,
  })

  if (splitRows.length > 0) {
    y = doc.lastAutoTable?.finalY || y + 20
    doc.setFont('helvetica', 'bold')
    doc.text('Split Balance', 14, y + 10)
    autoTable(doc, {
      startY: y + 13,
      head: [['Month', 'Owed to You']],
      body: splitRows.map(r => [r.month, fmtEUR(r.owed)]).concat([['Total', fmtEUR(splitTotal)]]),
      ...TABLE_STYLES,
    })
  }

  y = doc.lastAutoTable?.finalY || y + 20
  if (y > 240) doc.addPage()
  y = doc.lastAutoTable?.finalY || y
  doc.setFont('helvetica', 'bold')
  doc.text('Transactions (up to 100)', 14, y + 10)
  autoTable(doc, {
    startY: y + 13,
    head: [['Date', 'Description', 'Category', 'Amount', 'Split']],
    body: rows.slice(0, 100).map(r => [
      r.date,
      String(r.description || '').slice(0, 40),
      r.category,
      fmtEUR(r.amount || 0),
      r.split ? 'Yes' : 'No',
    ]),
    ...TABLE_STYLES,
    columnStyles: { 1: { cellWidth: 65 } },
  })

  doc.save('expense-report.pdf')
}
