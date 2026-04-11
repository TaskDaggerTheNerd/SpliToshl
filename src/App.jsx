import React, { useEffect, useMemo, useRef, useState } from 'react'
import Papa from 'papaparse'
import * as XLSX from 'xlsx'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Legend
} from 'recharts'
import {
  normalizeTransactions,
  dedup,
  fmtEUR,
  fmtInt,
  fmtNumber,
  buildForecast,
  buildSplitSummary,
  DEFAULTCATEGORIES,
} from './utils'
import { saveToIDB, loadFromIDB, exportJSON, importJSON } from './storage'
import { generatePDFReport } from './report'

const PALETTE = ['#01696f', '#e9c46a', '#f4a261', '#e76f51', '#264653', '#8ab17d', '#6d597a', '#577590', '#bc4749', '#a8dadc']
const TABS = ['Own','Joint','Trends','Forecast','Merchants','Subscriptions','Transactions','Splits']

const MoonIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
  </svg>
)

const SunIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <circle cx="12" cy="12" r="5" />
    <path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42" />
  </svg>
)

const EmptyState = ({ message = 'Upload a CSV or Excel (.xlsx) file to get started.' }) => (
  <div className="empty-state">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path d="M9 17H7A5 5 0 0 1 7 7h1m6 0h1a5 5 0 0 1 0 10h-1M8 12h8" />
    </svg>
    <h3>No data yet</h3>
    <p>{message}</p>
  </div>
)

function parseXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array', cellDates: true })
        const sheetName = workbook.SheetNames[0]
        const sheet = workbook.Sheets[sheetName]
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })
        resolve(rows)
      } catch (err) {
        reject(err)
      }
    }
    reader.onerror = () => reject(reader.error)
    reader.readAsArrayBuffer(file)
  })
}

function parseCSV(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      delimiter: '',
      encoding: 'UTF-8',
      complete: ({ data }) => resolve(data || []),
      error: err => reject(new Error(err.message)),
    })
  })
}

function getMonthKey(dateStr) {
  const d = String(dateStr || '')
  return d.length >= 7 ? d.slice(0, 7) : ''
}

export default function App() {
  const [transactions, setTransactions] = useState([])
  const [activeTab, setActiveTab] = useState('Own')
  const [query, setQuery] = useState('')
  const [status, setStatus] = useState('Import a CSV or Excel (.xlsx) file to start.')
  const [darkMode, setDarkMode] = useState(() => window.matchMedia('(prefers-color-scheme: dark)').matches)

  const [isMobile, setIsMobile] = useState(window.innerWidth <= 640)

  const [dateFilter, setDateFilter] = useState('')
  const [transactionCategoryFilter, setTransactionCategoryFilter] = useState('all')
  const [overviewCategoryFilter, setOverviewCategoryFilter] = useState('all')

  const [editingId, setEditingId] = useState(null)
  const [editDraft, setEditDraft] = useState({ date: '', description: '', category: '', amount: '', split: false, joint: false,})

  const [sortConfig, setSortConfig] = useState({ key: 'date', direction: 'desc' })
  const [jointSortConfig, setJointSortConfig] = useState({ key: 'date', direction: 'desc' })
  const [pendingImport, setPendingImport] = useState(null)

  const fileRef = useRef(null)
  const jsonRef = useRef(null)

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', darkMode ? 'dark' : 'light')
  }, [darkMode])
  useEffect(() => {
  const onResize = () => setIsMobile(window.innerWidth <= 640)
  window.addEventListener('resize', onResize)
  return () => window.removeEventListener('resize', onResize)
}, [])

  // Load from IndexedDB once on mount
  useEffect(() => {
    ;(async () => {
      const saved = await loadFromIDB()
      if (saved?.transactions?.length) {
        const normalized = dedup(normalizeTransactions(saved.transactions))
        setTransactions(normalized)
        setStatus(`Loaded ${normalized.length} saved transactions.`)
      }
    })()
  }, [])

  // Persist to IndexedDB whenever transactions change
  useEffect(() => {
    saveToIDB({ transactions })
  }, [transactions])

  // My personal share of every transaction
const myTransactions = useMemo(
  () =>
    transactions.map((t) => {
      const amount = Math.abs(Number(t.amount) || 0)
      const isJoint = t.account === 'joint' || t.joint
      const myAmount = isJoint ? amount / 2 : amount

      return {
        ...t,
        myAmount,
      }
    }),
  [transactions]
)

// Joint tab: only rows from the joint account, shown at FULL cost
const jointTransactions = useMemo(
  () =>
    transactions
      .filter((t) => t.account === 'joint' || t.joint)
      .map((t) => {
        const amount = Math.abs(Number(t.amount) || 0)

        return {
          ...t,
          jointAmount: amount,
        }
      }),
  [transactions]
)

const filteredJointTransactions = useMemo(() => {
  const q = query.trim().toLowerCase()

  return jointTransactions.filter((t) => {
    const matchesQuery =
      !q ||
      [
        t.description,
        t.merchant,
        t.category,
        t.date,
        String(t.amount),
      ]
        .join(' ')
        .toLowerCase()
        .includes(q)

    const matchesDate =
      !dateFilter || String(t.date || '').startsWith(dateFilter)

    const matchesCategory =
      transactionCategoryFilter === 'all' ||
      (t.category || 'Other') === transactionCategoryFilter

    return matchesQuery && matchesDate && matchesCategory
  })
}, [jointTransactions, query, dateFilter, transactionCategoryFilter])

const jointTotal = useMemo(
  () =>
    filteredJointTransactions.reduce(
      (s, t) => s + Number(t.jointAmount || 0),
      0
    ),
  [filteredJointTransactions]
)

const jointCategoryData = useMemo(() => {
  const m = new Map()

  filteredJointTransactions.forEach((t) => {
    m.set(t.category, (m.get(t.category) || 0) + Number(t.jointAmount || 0))
  })

  return [...m.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({ name, value }))
}, [filteredJointTransactions])

const jointMonthData = useMemo(() => {
  const m = new Map()

  filteredJointTransactions.forEach((t) => {
    const key = String(t.date).slice(0, 7)
    if (key.length === 7) {
      m.set(key, (m.get(key) || 0) + Number(t.jointAmount || 0))
    }
  })

  return [...m.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([name, value]) => ({ name, value }))
}, [filteredJointTransactions])

  // Filtered list (using my share)
  const filtered = useMemo(() => {
  const q = query.trim().toLowerCase()

  return myTransactions.filter((t) => {
    const matchesQuery =
      !q ||
      [
        t.description,
        t.merchant,
        t.category,
        t.date,
        String(t.amount),
      ]
        .join(' ')
        .toLowerCase()
        .includes(q)

    const matchesDate =
      !dateFilter || String(t.date || '').startsWith(dateFilter)

    const matchesTransactionCategory =
      transactionCategoryFilter === 'all' ||
      (t.category || 'Other') === transactionCategoryFilter

    return matchesQuery && matchesDate && matchesTransactionCategory
  })
}, [myTransactions, query, dateFilter, transactionCategoryFilter])

const overviewFiltered = useMemo(() => {
  return myTransactions.filter((t) => {
    return (
      overviewCategoryFilter === 'all' ||
      (t.category || 'Other') === overviewCategoryFilter
    )
  })
}, [myTransactions, overviewCategoryFilter])

  const sortedTransactions = useMemo(() => {
    const arr = [...filtered]
    const { key, direction } = sortConfig
    const dir = direction === 'asc' ? 1 : -1
    arr.sort((a, b) => {
      if (key === 'amount') {
        const av = Number(a.amount) || 0
        const bv = Number(b.amount) || 0
        return (av - bv) * dir
      }
      if (key === 'date') {
        const av = a.date || ''
        const bv = b.date || ''
        if (av < bv) return -1 * dir
        if (av > bv) return 1 * dir
        return 0
      }
      const av = (a[key] || '').toString().toLowerCase()
      const bv = (b[key] || '').toString().toLowerCase()
      return av.localeCompare(bv) * dir
    })
    return arr
  }, [filtered, sortConfig])

  function handleSort(key) {
    setSortConfig(prev => {
      if (prev.key === key) {
        return { key, direction: prev.direction === 'asc' ? 'desc' : 'asc' }
      }
      return { key, direction: key === 'date' ? 'desc' : 'asc' }
    })
  }

  const sortedJointTransactions = useMemo(() => {
  const arr = [...filteredJointTransactions]
  const { key, direction } = jointSortConfig
  const dir = direction === 'asc' ? 1 : -1

  arr.sort((a, b) => {
    if (key === 'amount') {
      const av = Number(a.jointAmount ?? a.amount ?? 0)
      const bv = Number(b.jointAmount ?? b.amount ?? 0)
      return (av - bv) * dir
    }

    if (key === 'date') {
      const av = a.date
      const bv = b.date
      if (av < bv) return -1 * dir
      if (av > bv) return 1 * dir
      return 0
    }

    const av = (a[key] ?? '').toString().toLowerCase()
    const bv = (b[key] ?? '').toString().toLowerCase()
    return av.localeCompare(bv) * dir
  })

  return arr
}, [filteredJointTransactions, jointSortConfig])

function handleJointSort(key) {
  setJointSortConfig(prev => {
    if (prev.key === key) {
      return {
        key,
        direction: prev.direction === 'asc' ? 'desc' : 'asc',
      }
    }
    return {
      key,
      direction: key === 'date' ? 'desc' : 'asc',
    }
  })
}

  const myTotalSpend = useMemo(
  () => filtered.reduce((s, t) => s + (Number(t.myAmount) || 0), 0),
  [filtered]
)

  const categoryData = useMemo(() => {
  const m = new Map()
  filtered.forEach(t => {
    m.set(t.category, (m.get(t.category) || 0) + (Number(t.myAmount) || 0))
  })
  return [...m.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({ name, value }))
}, [filtered])

  const merchantData = useMemo(() => {
  const m = new Map()

  overviewFiltered.forEach((t) => {
    const key = t.merchant || t.description || 'Unknown'
    m.set(key, (m.get(key) || 0) + Number(t.myAmount || 0))
  })

  return [...m.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15)
    .map(([name, value]) => ({ name, value }))
}, [overviewFiltered])

  const monthData = useMemo(() => {
  const m = new Map()
  filtered.forEach(t => {
    const key = String(t.date || '').slice(0, 7)
    if (key.length === 7) {
      m.set(key, (m.get(key) || 0) + (Number(t.myAmount) || 0))
    }
  })
  return [...m.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([name, value]) => ({ name, value }))
}, [filtered])

    const forecast = useMemo(
  () => buildForecast(myTransactions.map(t => ({ ...t, amount: t.myAmount }))),
  [myTransactions]
)
  const splitSummary  = useMemo(() => buildSplitSummary(transactions), [transactions])
  const splitRows     = useMemo(
    () => [...splitSummary.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([month, owed]) => ({ month, owed })),
    [splitSummary]
  )
  const splitTotal    = useMemo(() => splitRows.reduce((s, r) => s + r.owed, 0),
  [splitRows])

  const splitBreakdownByMonthCategory = useMemo(() => {
  const rows = transactions
    .filter(t => t.split)
    .map(t => ({
      month: getMonthKey(t.date),
      category: t.category || 'Other',
      owed: Math.abs(Number(t.amount) || 0) / 2,
    }))
    .filter(r => r.month)

  const grouped = new Map()

  rows.forEach(r => {
    const key = `${r.month}__${r.category}`
    const current = grouped.get(key) || 0
    grouped.set(key, current + r.owed)
  })

  return [...grouped.entries()]
    .map(([key, owed]) => {
      const [month, category] = key.split('__')
      return { month, category, owed }
    })
    .sort((a, b) => {
      if (a.month !== b.month) return a.month.localeCompare(b.month)
      return a.category.localeCompare(b.category)
    })
}, [transactions])

  // --- subscription helpers (must be BEFORE subscriptions useMemo) ---

  const getSubscriptionKeyFromDescription = (description = '') => {
    const d = String(description).toLowerCase()
    // try to match a known subscription brand word first
    const m = d.match(
      /netflix|spotify|prime video|amazon prime|apple tv|apple music|hbo|disney|paramount|youtube|icloud|dropbox|dazn|patreon|subscription/i
    )
    if (m) return m[0].toLowerCase()
    // fallback: first word
    const first = d.trim().split(/\s+/)[0]
    return first || d || 'subscription'
  }

  const isSubscriptionCandidate = (t) => {
    const d = String(t.description || '')
    return /netflix|spotify|prime|apple|hbo|disney|subscription|adobe|microsoft|google one|icloud|dropbox|urban sports|revolut|youtube/i.test(d)
  }

      const subscriptions = useMemo(() => {
    const now = new Date()
    const currentMonthKey = `${now.getFullYear()}-${String(
      now.getMonth() + 1
    ).padStart(2, '0')}`

    const candidates = filtered.filter(isSubscriptionCandidate)

    // group by subscription key (service name)
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

    const summaries = []

    for (const [subKey, txs] of groups.entries()) {
      if (!txs.length) continue

      // sort newest → oldest
      const sorted = txs
        .filter(t => t.date)
        .slice()
        .sort((a, b) => (a.date > b.date ? -1 : 1))

      if (!sorted.length) continue

      const latest = sorted[0]
      const lastDate = latest.date
      const lastMonth = getMonthKey(lastDate)
      if (!lastMonth) continue

      const diff = monthDiff(currentMonthKey, lastMonth)
      const status = diff >= 2 ? 'over' : 'active' // over only if 2 months with no charge

      // total spent historically for this service
      const totalAmount = txs.reduce(
        (sum, t) => sum + Math.abs(Number(t.amount) || 0),
        0
      )

      const lastAmount = Math.abs(Number(latest.amount) || 0)

      summaries.push({
        ...latest,                // keeps id, date, description, category, amount, split
        subscriptionStatus: status,
        totalAmount,
        lastAmount,
      })
    }

    // newest subscriptions first
    summaries.sort((a, b) => {
      const ad = a.date || ''
      const bd = b.date || ''
      if (ad < bd) return 1
      if (ad > bd) return -1
      return 0
    })

    return summaries
  }, [filtered])

    const categoryOptions = useMemo(() => {
    const existing = [...new Set(transactions.map(t => t.category).filter(Boolean))]
    return [...new Set([...DEFAULTCATEGORIES, ...existing])].sort((a, b) => a.localeCompare(b))
  }, [transactions])

  const getCategoryColor = (name) => {
    const idx = categoryOptions.indexOf(name)
    const safeIdx = idx === -1 ? 0 : idx
    return PALETTE[safeIdx % PALETTE.length]
  }

  const hasData = transactions.length > 0

function importRows(rows, label) {
  const hasAccountFlag = rows.some(r => r.account)

  if (hasAccountFlag) {
    finishImport(rows, label, null)
  } else {
    setPendingImport({ rows, label })
  }
}

function finishImport(rows, label, accountType) {
  const isJoint = accountType === 'joint'
  const normalized = dedup(normalizeTransactions(rows))

  if (normalized.length === 0) {
    setStatus('No transactions found. Check file columns.')
    setPendingImport(null)
    return
  }

  const enriched = normalized.map((t) => ({
    ...t,
    account: t.account || (isJoint ? 'joint' : 'personal'),
    split: t.account ? Boolean(t.split) : false,
    joint: t.account === 'joint' ? true : (isJoint ? true : Boolean(t.joint)),
    jointMode:
      t.account === 'joint'
        ? 'full'
        : isJoint
          ? 'full'
          : (t.joint ? (t.jointMode || 'full') : null),
  }))

  const merged = dedup([...transactions, ...enriched])

  if (merged.length === transactions.length) {
    setStatus(`No new transactions found in ${label}; all were duplicates.`)
    setPendingImport(null)
    return
  }

  setTransactions(merged)
  setEditingId(null)
  setStatus(`Merged ${merged.length - transactions.length} new transactions from ${label}.`)
  setActiveTab('Own')
  setPendingImport(null)
}

async function handleFile(file) {
  
  if (!file) return
  const ext = file.name.split('.').pop().toLowerCase()
  setStatus(`Reading ${file.name}…`)
  try {
    if (ext === 'xlsx' || ext === 'xls') {
      const rows = await parseXLSX(file)
      importRows(rows, file.name)
    } else if (ext === 'csv' || ext === 'txt') {
      const rows = await parseCSV(file)
      importRows(rows, file.name)
    } else {
      setStatus(`Unsupported file type ".${ext}". Please use .xlsx or .csv.`)
    }
  } catch (err) {
    setStatus(`Error reading file: ${err.message}`)
  }
}

async function handleJSON(file) {
  if (!file) return
  try {
    const data = await importJSON(file)
    if (data?.transactions?.length) importRows(data.transactions, 'JSON')
    else setStatus('JSON file has no transactions.')
  } catch {
    setStatus('Failed to parse JSON.')
  }
}

  function handleExport() {
    exportJSON({ transactions }, 'expense-data.json').then(() => setStatus('JSON exported.'))
  }

  function handlePDF() {
  generatePDFReport(transactions)
  setStatus('PDF downloaded.')
}

  function handleClear() {
  if (!transactions.length) {
    setStatus('There is no data to clear.')
    return
  }

  const ok = window.confirm(
    'This will remove all imported transactions from this browser. Are you sure you want to clear everything?'
  )

  if (!ok) {
    setStatus('Clear cancelled.')
    return
  }

  setTransactions([])
  setEditingId(null)
  setStatus('All data cleared.')
}

function markAllSplitsPaid() {
  let changed = 0

  setTransactions(prev =>
    prev.map(t => {
      if (t.split) {
        changed++
        return { ...t, split: false, splitPaid: true }
      }
      return t
    })
  )

  if (changed === 0) {
    setStatus('No split transactions to mark as paid.')
  } else {
    setStatus(`Marked ${changed} split transactions as paid.`)
  }
}

function toggleSplit(id) {
  const tx = transactions.find(t => t.id === id)
  if (!tx) return

  if (tx.splitPaid) {
    setStatus('This transaction was already split and paid in the past.')
    return
  }

  setTransactions(prev =>
    prev.map(t => (t.id === id ? { ...t, split: !t.split } : t))
  )
  setStatus('Split updated.')
}

  function toggleJoint(id) {
  const tx = transactions.find((t) => t.id === id)
  if (!tx) return

  if (tx.account === 'joint') {
    setStatus('This transaction already comes from a Joint statement.')
    return
  }

  if (tx.joint) {
    setTransactions((prev) =>
      prev.map((t) =>
        t.id === id ? { ...t, joint: false, jointMode: null } : t
      )
    )
    setStatus('Removed from Joint tab.')
    return
  }

  setTransactions((prev) =>
    prev.map((t) =>
      t.id === id ? { ...t, joint: true, jointMode: null } : t
    )
  )
  setStatus('Added to Joint tab.')
}

  function startEdit(t) {
  setEditingId(t.id)
  setEditDraft({
    date: t.date,
    description: t.description,
    category: t.category || 'Other',
    amount: String(t.amount ?? ''),
    split: Boolean(t.split),
    joint: Boolean(t.joint || t.account === 'joint'),
  })
}

  function cancelEdit() {
  setEditingId(null)
  setEditDraft({
    date: '',
    description: '',
    category: '',
    amount: '',
    split: false,
    joint: false,
  })
}

  function saveEdit(id) {
  const amount = Number(String(editDraft.amount).replace(',', '.'))

  if (!editDraft.date) {
    setStatus('Date is required.')
    return
  }

  if (!Number.isFinite(amount) || amount < 0) {
    setStatus('Amount must be a valid positive number.')
    return
  }

  const original = transactions.find(t => t.id === id)
  const categoryChanged = original && original.category !== (editDraft.category || 'Other')
  const originalDescription = original?.description

  setTransactions(prev =>
    prev.map(t => {
      if (t.id === id)
        return {
          ...t,
          date: editDraft.date,
          description: editDraft.description.trim() || t.description,
          merchant: editDraft.description.trim() || t.merchant,
          category: editDraft.category || 'Other',
          amount,
          split: t.splitPaid ? false : Boolean(editDraft.split),
          splitPaid: Boolean(t.splitPaid),
          joint: t.account === 'joint' ? true : Boolean(editDraft.joint),
          jointMode: null,}

      if (categoryChanged && originalDescription && t.description === originalDescription)
        return { ...t, category: editDraft.category || 'Other' }

      return t
    })
  )

  if (categoryChanged && originalDescription) {
    const matchCount = transactions.filter(t => t.id !== id && t.description === originalDescription).length
    if (matchCount > 0)
      setStatus(`Updated. Category "${editDraft.category}" applied to ${matchCount + 1} transactions with description "${originalDescription}".`)
    else setStatus('Transaction updated.')
  } else {
    setStatus('Transaction updated.')
  }

  cancelEdit()
}

  function deleteTransaction(id) {
    setTransactions(prev => prev.filter(t => t.id !== id))
    if (editingId === id) cancelEdit()
    setStatus('Transaction deleted.')
  }

  const forecastByCategory = useMemo(() => {
    const byMonth = new Map()
    transactions.forEach(t => {
      const month = getMonthKey(t.date)
      if (!month) return
      const cat = t.category || 'Other'
      const key = `${cat}|${month}`
      byMonth.set(key, (byMonth.get(key) || 0) + Math.abs(Number(t.amount) || 0))
    })

    const months = [...new Set([...byMonth.keys()].map(k => k.split('|')[1]))].sort()
    const last6 = months.slice(-6)
    if (last6.length === 0) return []

    const last6Set = new Set(last6)
    const catTotals = new Map()
    byMonth.forEach((val, key) => {
      const [cat, month] = key.split('|')
      if (last6Set.has(month)) {
        catTotals.set(cat, (catTotals.get(cat) || 0) + val)
      }
    })

    return [...catTotals.entries()].map(([category, total]) => {
      const avgPerMonth = total / last6.length
      const projected3 = avgPerMonth * 3
      return { category, avgPerMonth, projected3 }
    })
  }, [transactions])

  const mostExpensiveCategory = useMemo(() => {
    if (!forecastByCategory.length) return null
    return forecastByCategory.reduce((max, row) =>
      !max || row.avgPerMonth > max.avgPerMonth ? row : max,
      null
    )
  }, [forecastByCategory])

  const tt = {
  contentStyle: {
    background: 'var(--color-surface)',
    border: '1px solid var(--color-border)',
    borderRadius: 8,
  },
  formatter: (v) => fmtEUR(v),
}

const axisTick = { fontSize: isMobile ? 10 : 11, fill: 'var(--color-text-muted)' }
const pieHeight = isMobile ? 220 : 260
const chartHeight = isMobile ? 220 : 300
const merchantChartHeight = isMobile ? 320 : 400

return (
  
  <div className="app">
      <header className="topbar">
        <div className="topbar-title">
          <h1>SpliToshl</h1>
          <p>{status}</p>
        </div>
        <div className="topbar-actions">
          <button className="btn btn-primary" onClick={() => fileRef.current?.click()}>
            Import CSV / Excel
          </button>
          <button className="btn" onClick={() => jsonRef.current?.click()}>
            Import JSON
          </button>
          <button className="btn" onClick={handleExport}>
            Export JSON
          </button>
          <button className="btn" onClick={handlePDF}>
            Download PDF
          </button>
          <button className="btn" onClick={handleClear}>
            Clear
            </button>
           <button
            className="btn btn-theme"
            onClick={() => setDarkMode(v => !v)}
            aria-label="Toggle theme"
          >
            {darkMode ? <SunIcon /> : <MoonIcon />}
          </button>
          <input
            ref={fileRef}
            type="file"
            accept=".csv,.xlsx,.xls,text/csv"
            hidden
            onChange={e => handleFile(e.target.files?.[0])}
          />
          <input
            ref={jsonRef}
            type="file"
            accept=".json,application/json"
            hidden
            onChange={e => handleJSON(e.target.files?.[0])}
          />
        </div>
      </header>

      <section className="kpis">
        <div className="kpi-card">
  <div className="label">My Spend</div>
  <div className="value">{fmtEUR(myTotalSpend)}</div>
</div>
        <div className="kpi-card">
          <div className="label">Transactions</div>
          <div className="value">{fmtInt(filtered.length)}</div>
        </div>
        <div className="kpi-card">
          <div className="label">Categories</div>
          <div className="value">{fmtInt(categoryData.length)}</div>
        </div>
        <div className="kpi-card">
  <div className="label">Joint Spend (full)</div>
  <div className="value">{fmtEUR(jointTotal)}</div>
</div>
        <div className="kpi-card accent">
          <div className="label">Split Balance (owed to you)</div>
          <div className="value">{fmtEUR(splitTotal)}</div>
        </div>
      </section>

      <nav className="tabs">
        {TABS.map(tab => (
          <button
            key={tab}
            className={`tab-btn${activeTab === tab ? ' active' : ''}`}
            onClick={() => setActiveTab(tab)}
          >
            {tab}
            {tab === 'Splits' && splitTotal > 0 ? ` (${fmtEUR(splitTotal)})` : ''}
          </button>
        ))}
      </nav>

<div className="toolbar">
  <input
    className="search-input"
    value={query}
    onChange={(e) => setQuery(e.target.value)}
    placeholder="Search transactions..."
  />

  {(activeTab === 'Trends' || activeTab === 'Merchants') && (
    <select
      className="field-input"
      value={overviewCategoryFilter}
      onChange={(e) => setOverviewCategoryFilter(e.target.value)}
      title="Filter by category"
    >
      <option value="all">All categories</option>
      {categoryOptions.map((cat) => (
        <option key={cat} value={cat}>
          {cat}
        </option>
      ))}
    </select>
  )}
</div>

{/* ── OWN OVERVIEW ── */}
{activeTab === 'Own' &&
  (!hasData ? (
    <EmptyState />
  ) : (
    <>
      <div className="grid-2">
        <div className="panel">
          <h2>Spending by Category</h2>
          <ResponsiveContainer width="100%" height={pieHeight}>
            <PieChart>
              <Pie
                data={categoryData}
                dataKey="value"
                nameKey="name"
                outerRadius={90}
                label={({ name, percent }) =>
                  `${name} ${fmtNumber(percent * 100, 0)}%`
                }
              >
                {categoryData.map((d) => (
                  <Cell key={d.name} fill={getCategoryColor(d.name)} />
                ))}
              </Pie>
              <Tooltip {...tt} />
            </PieChart>
          </ResponsiveContainer>
        </div>

        <div className="panel">
          <h2>Monthly Spend</h2>
          <ResponsiveContainer width="100%" height={chartHeight}>
            <BarChart data={monthData}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-divider)" />
              <XAxis dataKey="name" tick={axisTick} />
              <YAxis tick={axisTick} />
              <Tooltip {...tt} />
              <Bar dataKey="value" fill="#01696f" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid-2">
        <div className="panel">
          <h2>Own Transactions {sortedTransactions.length}</h2>
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th onClick={() => handleSort('date')}>Date</th>
                  <th onClick={() => handleSort('description')}>Description</th>
                  <th onClick={() => handleSort('category')}>Category</th>
                  <th onClick={() => handleSort('amount')}>Amount</th>
                </tr>
              </thead>
              <tbody>
                {sortedTransactions.map((t) => (
                  <tr key={t.id}>
                    <td>{t.date}</td>
                    <td>{t.description}</td>
                    <td>{t.category}</td>
                    <td className="amount">{fmtEUR(t.myAmount || 0)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="panel">
          <h2>Own Summary</h2>
          <div className="table-wrap">
            <table className="data-table">
              <tbody>
                <tr>
                  <td>Total Own Spend</td>
                  <td className="amount">{fmtEUR(myTotalSpend)}</td>
                </tr>
                <tr>
                  <td>Transactions</td>
                  <td className="amount">{fmtEUR(filtered.length)}</td>
                </tr>
                <tr>
                  <td>Categories</td>
                  <td className="amount">{fmtInt(categoryData.length)}</td>
                </tr>
                {categoryData.map((row) => (
                  <tr key={row.name}>
                    <td>{row.name}</td>
                    <td className="amount">{fmtEUR(row.value)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </>
  ))}

      {/* ── JOINT OVERVIEW ── */}
      {activeTab === 'Joint' &&
        (!hasData ? (
          <EmptyState message="Import a joint account statement to see joint costs." />
        ) : jointTransactions.length === 0 ? (
          <EmptyState message="No joint transactions detected yet." />
        ) : (
          <div className="grid-2">
            <div className="panel">
              <h2>Joint Spending by Category</h2>
              <ResponsiveContainer width="100%" height={pieHeight}>
                <PieChart>
                  <Pie
                    data={jointCategoryData}
                    dataKey="value"
                    nameKey="name"
                    outerRadius={90}
                    label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                  >
                    {jointCategoryData.map(d => (
                      <Cell key={d.name} fill={getCategoryColor(d.name)} />
                    ))}
                  </Pie>
                  <Tooltip {...tt} />
                </PieChart>
              </ResponsiveContainer>
            </div>

            <div className="panel">
              <h2>Joint Monthly Spend</h2>
              <ResponsiveContainer width="100%" height={chartHeight}>
                <BarChart data={jointMonthData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--color-divider)" />
                  <XAxis dataKey="name" tick={axisTick} />
                  <YAxis tick={axisTick} />
                  <Tooltip {...tt} />
                  <Bar dataKey="value" fill="#264653" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="panel">
              <h2>Joint Transactions {sortedJointTransactions.length}</h2>
              <div className="table-wrap">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th onClick={() => handleJointSort('date')}>Date</th>
                      <th onClick={() => handleJointSort('description')}>Description</th>
                      <th onClick={() => handleJointSort('category')}>Category</th>
                      <th onClick={() => handleJointSort('amount')}>Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedJointTransactions.map(t => (
                      <tr key={t.id}>
                        <td>{t.date}</td>
                        <td>{t.description}</td>
                        <td>{t.category}</td>
                        <td className="amount">{fmtEUR(t.jointAmount || 0)}</td>
                      </tr>
                     ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="panel">
  <h2>Joint Summary</h2>
  <div className="table-wrap">
    <table className="data-table">
      <tbody>
        <tr>
          <td>Total Joint Spend</td>
          <td className="amount">{fmtEUR(jointTotal)}</td>
        </tr>
        <tr>
          <td>Transactions</td>
          <td className="amount">{fmtEUR(filteredJointTransactions.length)}</td>
        </tr>
        {jointCategoryData.map((row) => (
          <tr key={row.name}>
            <td>{row.name}</td>
            <td className="amount">{fmtEUR(row.value)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>
</div>
</div>
        ))}

      {/* ── TRENDS ── */}
      {activeTab === 'Trends' &&
        (!hasData ? (
          <EmptyState />
        ) : (
          <div className="panel">
            <h2>Top Merchants by Spend</h2>
            <ResponsiveContainer width="100%" height={merchantChartHeight}>
              <BarChart data={merchantData} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--color-divider)" />
                <XAxis type="number" tick={axisTick} tickFormatter={v => fmtEUR(v)} />
                <YAxis type="category" dataKey="name" width={180} tick={axisTick} />
                <Tooltip {...tt} />
                <Bar dataKey="value" fill="#264653" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        ))}

      {/* ── FORECAST ── */}
      {activeTab === 'Forecast' &&
        (!hasData ? (
          <EmptyState />
        ) : (
          <>
            <div className="grid-2">
              <div className="panel">
                <h2>Spend Forecast</h2>
                <p className="subtle-note">
                  Average of recent monthly spend: {fmtEUR(forecast.avg)}
                </p>
                <ResponsiveContainer width="100%" height={chartHeight}>
                  <LineChart data={forecast.series}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--color-divider)" />
                    <XAxis dataKey="month" tick={axisTick} />
                    <YAxis tick={axisTick} tickFormatter={v => fmtEUR(v)} width={88} />
                    <Tooltip {...tt} />
                    <Legend />
                    <Line
                      type="monotone"
                      dataKey="actual"
                      stroke="#01696f"
                      strokeWidth={2}
                      dot={{ r: 4 }}
                      name="Actual"
                      connectNulls={false}
                    />
                    <Line
                      type="monotone"
                      dataKey="projected"
                      stroke="#f4a261"
                      strokeWidth={2}
                      strokeDasharray="6 4"
                      dot={{ r: 4 }}
                      name="Projected"
                      connectNulls={false}
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
              <div className="panel">
                <h2>Forecast Details</h2>
                <div className="table-wrap">
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>Month</th>
                        <th>Actual</th>
                        <th>Projected</th>
                      </tr>
                    </thead>
                    <tbody>
                      {forecast.series.map(row => (
                        <tr key={row.month}>
                          <td>{row.month}</td>
                          <td className="amount">
                            {row.actual != null ? fmtEUR(row.actual) : '—'}
                          </td>
                          <td className="amount">
                            {row.projected != null ? fmtEUR(row.projected) : '—'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {forecastByCategory.length > 0 && (
              <div className="panel">
                <h2>Forecast by Category</h2>
                <div className="table-wrap">
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>Category</th>
                        <th>Avg monthly (last 6 months)</th>
                        <th>Projected next 3 months</th>
                      </tr>
                    </thead>
                    <tbody>
                      {forecastByCategory.map(row => (
                        <tr key={row.category}>
                          <td>{row.category}</td>
                          <td className="amount">{fmtEUR(row.avgPerMonth)}</td>
                          <td className="amount">{fmtEUR(row.projected3)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {mostExpensiveCategory && (
                  <p className="subtle-note" style={{ marginTop: '1rem' }}>
                    Most expensive category in the recent period is{' '}
                    <strong>{mostExpensiveCategory.category}</strong> with an average of{' '}
                    <strong>{fmtEUR(mostExpensiveCategory.avgPerMonth)}</strong> per month.
                  </p>
                )}
              </div>
            )}
          </>
        ))}

      {/* ── MERCHANTS ── */}
      {activeTab === 'Merchants' &&
        (!hasData ? (
          <EmptyState />
        ) : (
          <div className="panel">
            <h2>Top Merchants</h2>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Merchant</th>
                    <th>Spend</th>
                  </tr>
                </thead>
                <tbody>
                  {merchantData.map((row, i) => (
                    <tr key={row.name}>
                      <td className="muted">{i + 1}</td>
                      <td>{row.name}</td>
                      <td className="amount">{fmtEUR(row.value)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ))}

      {/* ── SUBSCRIPTIONS ── */}
{activeTab === 'Subscriptions' &&
  (!hasData ? (
    <EmptyState />
  ) : subscriptions.length === 0 ? (
    <EmptyState message="No subscriptions detected." />
  ) : (
    <div className="panel">
      <h2>Subscriptions</h2>
      <div className="table-wrap">
        <table className="data-table">
          <thead>
            <tr>
              <th>Subscription</th>
              <th>Total spent</th>
              <th>Last charge</th>
              <th>Last charge date</th>
              <th>Status</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
  {subscriptions.map(t => {
    const isEditing = editingId === t.id

    return (
      <tr key={t.id} className={isEditing ? 'editing-row' : ''}>
        {/* Subscription name */}
        <td>
          {isEditing ? (
            <input
              className="field-input"
              type="text"
              value={editDraft.description}
              placeholder="Subscription"
              onChange={e =>
                setEditDraft(p => ({ ...p, description: e.target.value }))
              }
            />
          ) : (
            t.description || <span className="muted">—</span>
          )}
        </td>

        {/* Total spent historically */}
        <td className="amount">
          {fmtEUR(t.totalAmount || 0)}
        </td>

        {/* Last charge amount */}
        <td className="amount">
          {isEditing ? (
            <input
              className="field-input amount-input"
              type="number"
              min="0"
              step="0.01"
              value={editDraft.amount}
              onChange={e =>
                setEditDraft(p => ({ ...p, amount: e.target.value }))
              }
            />
          ) : (
            fmtEUR(t.lastAmount ?? t.amount ?? 0)
          )}
        </td>

        {/* Last charge date */}
        <td>
          {isEditing ? (
            <input
              className="field-input"
              type="date"
              value={editDraft.date}
              onChange={e =>
                setEditDraft(p => ({ ...p, date: e.target.value }))
              }
            />
          ) : (
            t.date
          )}
        </td>

        {/* Status */}
        <td>
          {t.subscriptionStatus === 'active' ? (
            <span className="muted">Active</span>
          ) : (
            <span className="muted">Over currently</span>
          )}
        </td>

        {/* Actions */}
        <td>
          <div className="row-actions">
            {isEditing ? (
              <>
                <button
                  className="btn btn-small btn-primary"
                  onClick={() => saveEdit(t.id)}
                >
                  Save
                </button>
                <button className="btn btn-small" onClick={cancelEdit}>
                  Cancel
                </button>
              </>
            ) : (
              <>
                <button
                  className="btn btn-small"
                  onClick={() => startEdit(t)}
                >
                  Edit
                </button>
                <button
                  className="btn btn-small btn-danger"
                  onClick={() => deleteTransaction(t.id)}
                >
                  Delete
                </button>
              </>
            )}
          </div>
        </td>
      </tr>
    )
  })}
</tbody>
        </table>
      </div>
    </div>
  ))}

      {/* ── TRANSACTIONS ── */}
      {activeTab === 'Transactions' &&
        (!hasData ? (
          <EmptyState />
        ) : (
          <div className="panel">
            <h2>Transactions ({sortedTransactions.length})</h2>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                   <th onClick={() => handleSort('date')}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.35rem' }}>
                      <span>Date</span>
                      <input
                        className="field-input"
                        type="month"
                        value={dateFilter}
                        onChange={(e) => setDateFilter(e.target.value)}
                        onClick={(e) => e.stopPropagation()}
                      />
                    </div>
                  </th>

                  <th onClick={() => handleSort('description')}>Description</th>

                  <th onClick={() => handleSort('category')}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.35rem' }}>
                      <span>Category</span>
                      <select
                        className="field-input"
                        value={transactionCategoryFilter}
                        onChange={(e) => setTransactionCategoryFilter(e.target.value)}
                        onClick={(e) => e.stopPropagation()}
                      >
                        <option value="all">All categories</option>
                        {categoryOptions.map((cat) => (
                          <option key={cat} value={cat}>
                            {cat}
                          </option>
                        ))}
                      </select>
                    </div>
                  </th>

                  <th onClick={() => handleSort('amount')}>Amount</th>
                  <th>Split</th>
                  <th>Joint?</th>
                  <th>
                    <button
                      className="btn btn-small"
                      onClick={() => {
                        setDateFilter('')
                        setTransactionCategoryFilter('all')
                      }}
                    >
                      Clear
                    </button>
                  </th>
                </tr>
              </thead>

                <tbody>
                  {sortedTransactions.map(t => {
                    const isEditing = editingId === t.id
                    return (
                      <tr
                        key={t.id}
                          className={`${t.split ? 'split-row' : ''} ${t.splitPaid ? 'split-paid-row' : ''} ${isEditing ? 'editing-row' : ''}`.trim()}
                          >
                            <td>
                              {isEditing ? (
                              <input
                              className="field-input"
                              type="date"
                              value={editDraft.date}
                              onChange={e => setEditDraft(p => ({ ...p, date: e.target.value }))}
                              />
                                ) : (
                                t.date
                                )}
                            </td>

                          <td>
    {isEditing ? (
      <input
        className="field-input"
        type="text"
        value={editDraft.description}
        placeholder="Description"
        onChange={e => setEditDraft(p => ({ ...p, description: e.target.value }))}
      />
    ) : (
      <span className="muted">{t.description}</span>
    )}
  </td>

  <td>
  {isEditing ? (
    <select
      className="field-input"
      value={editDraft.category}
      onChange={(e) =>
        setEditDraft((p) => ({ ...p, category: e.target.value }))
      }
    >
      {categoryOptions.map((cat) => (
        <option key={cat} value={cat}>
          {cat}
        </option>
      ))}
    </select>
  ) : (
    t.category
  )}
</td>

  <td className="amount">
    {isEditing ? (
      <input
        className="field-input amount-input"
        type="number"
        min="0"
        step="0.01"
        value={editDraft.amount}
        onChange={e => setEditDraft(p => ({ ...p, amount: e.target.value }))}
      />
    ) : (
      fmtEUR(t.amount || 0)
    )}
  </td>

  <td>
    {isEditing ? (
      <select
        className="field-input"
        value={editDraft.split ? 'yes' : 'no'}
        onChange={e => setEditDraft(p => ({ ...p, split: e.target.value === 'yes' }))}
      >
        <option value="no">No</option>
        <option value="yes">Yes</option>
      </select>
    ) : (
      <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
  <button
    className={`btn btn-split ${(t.split || t.splitPaid) ? 'yes' : ''} ${t.splitPaid ? 'disabled' : ''}`}
    onClick={() => toggleSplit(t.id)}
    disabled={t.splitPaid}
    title={t.splitPaid ? 'Already split and paid in the past' : ''}
  >
    {(t.split || t.splitPaid) ? 'Yes' : 'No'}
  </button>

  {t.splitPaid && (
    <span
      style={{
        fontSize: '0.72rem',
        padding: '0.2rem 0.45rem',
        borderRadius: '999px',
        background: 'var(--color-warning-highlight)',
        color: 'var(--color-warning)',
        whiteSpace: 'nowrap',
        fontWeight: 600,
      }}
    >
      Already Paid
    </span>
  )}
</div>
    )}
  </td>

  <td>
  {isEditing ? (
    <select
      className="field-input"
      value={(t.account === 'joint' || editDraft.joint) ? 'yes' : 'no'}
      onChange={(e) =>
        setEditDraft((p) => ({
          ...p,
          joint: e.target.value === 'yes',
        }))
      }
      disabled={t.account === 'joint'}
      title={t.account === 'joint' ? 'Already from a Joint statement' : ''}
    >
      <option value="no">No</option>
      <option value="yes">Yes</option>
    </select>
  ) : (
    <button
      className={`btn btn-split ${(t.joint || t.account === 'joint') ? 'yes' : ''}`}
      onClick={() => toggleJoint(t.id)}
      disabled={t.account === 'joint'}
      title={t.account === 'joint' ? 'Already from a Joint statement' : ''}
    >
      {(t.joint || t.account === 'joint') ? 'Yes' : 'No'}
    </button>
  )}
</td>

  <td>
    <div className="row-actions">
      {isEditing ? (
        <>
          <button className="btn btn-small btn-primary" onClick={() => saveEdit(t.id)}>
            Save
          </button>
          <button className="btn btn-small" onClick={cancelEdit}>
            Cancel
          </button>
        </>
      ) : (
        <>
          <button className="btn btn-small" onClick={() => startEdit(t)}>
            Edit
          </button>
          <button className="btn btn-small btn-danger" onClick={() => deleteTransaction(t.id)}>
            Delete
          </button>
        </>
      )}
    </div>
  </td>
</tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        ))}

      {/* ── SPLITS ── */}
      {activeTab === 'Splits' &&
        (splitRows.length === 0 ? (
          <EmptyState message='Go to Transactions and set Split to "Yes" on any shared expense.' />
        ) : (
          <>
            <div className="panel">
              <div
  style={{
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: '1rem',
    flexWrap: 'wrap',
  }}
>
  <h2>Monthly Split Balance</h2>
  <button className="btn btn-small btn-primary" onClick={markAllSplitsPaid}>Paid</button>
</div>
              <p className="subtle-note">
                50% of each split transaction is counted as owed to you.
              </p>
              <div className="table-wrap">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Month</th>
                      <th>Owed to You</th>
                    </tr>
                  </thead>
                  <tbody>
                    {splitRows.map(r => (
                      <tr key={r.month}>
                        <td>{r.month}</td>
                        <td className="amount">{fmtEUR(r.owed)}</td>
                      </tr>
                    ))}
                    <tr className="total-row">
                      <td>Total</td>
                      <td className="amount">{fmtEUR(splitTotal)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            {splitBreakdownByMonthCategory.length > 0 && (
              <div className="panel">
                <h2>Split Debt by Month & Category</h2>
                <div className="table-wrap">
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>Month</th>
                        <th>Category</th>
                        <th>Owed to You</th>
                      </tr>
                    </thead>
                    <tbody>
                      {splitBreakdownByMonthCategory.map(row => (
                        <tr key={`${row.month}-${row.category}`}>
                          <td>{row.month}</td>
                          <td>{row.category}</td>
                          <td className="amount">{fmtEUR(row.owed)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
              ))}

      {/* Import type modal */}
      {pendingImport && (
        <div className="modal-backdrop">
          <div className="modal">
            <h3>Which account is this statement from?</h3>
            <p>
              Choose <strong>Own</strong> for your personal account, or{' '}
              <strong>Joint</strong> for your shared account with your partner.
            </p>
            <div className="modal-actions">
              <button
                className="btn btn-modal"
                onClick={() => finishImport(pendingImport.rows, pendingImport.label, 'personal')}
              >
                Own
              </button>
              <button
                className="btn btn-modal"
                onClick={() => finishImport(pendingImport.rows, pendingImport.label, 'joint')}
              >
                Joint
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}