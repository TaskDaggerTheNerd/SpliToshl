import React, { useEffect, useMemo, useRef, useState } from 'react'
import Papa from 'papaparse'
import * as XLSX from 'xlsx'
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  Legend,
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
import { saveToIDB, clearIDB, exportJSON, importJSON } from './storage'
import { supabase } from './supabase'
import { generatePDFReport, generateSplitPDFReport } from './report'

const PALETTE = [
  '#01696f',
  '#e9c46a',
  '#f4a261',
  '#e76f51',
  '#264653',
  '#8ab17d',
  '#6d597a',
  '#577590',
  '#bc4749',
  '#a8dadc',
]

const TABS = [
  'Own',
  'Joint',
  'Trends',
  'Forecast',
  'Merchants',
  'Subscriptions',
  'Transactions',
  'Splits',
]

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

function parseXLSXfile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
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

function parseCSVfile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      delimiter: ',',
      encoding: 'UTF-8',
      complete: ({ data }) => resolve(data),
      error: (err) => reject(new Error(err.message)),
    })
  })
}

function makeJointGroupId(tx) {
  if (tx?.jointGroupId) return tx.jointGroupId
  return `joint_${tx.date || 'nodate'}_${(tx.description || 'expense')
    .replace(/\s+/g, '-')
    .slice(0, 30)}_${tx.id || Date.now()}`
}

function makeSplitGroupId(tx) {
  if (tx?.splitGroupId) return tx.splitGroupId
  return `split_${tx.date || 'nodate'}_${(tx.description || 'expense')
    .replace(/\s+/g, '-')
    .slice(0, 30)}_${tx.id || Date.now()}`
}

function getPartnerUserId(currentUser) {
  if (!currentUser?.id) return null
  return currentUser.id === 1 ? 2 : 1
}

function getMonthKey(dateStr) {
  const d = String(dateStr || '')
  return d.length >= 7 ? d.slice(0, 7) : ''
}

export default function App() {
  const [transactions, setTransactions] = useState([])
  const [activeTab, setActiveTab] = useState('Own')
  const [query, setQuery] = useState('')
  const [status, setStatus] = useState('Import a CSV or Excel .xlsx file to start.')
  const [darkMode, setDarkMode] = useState(window.matchMedia('(prefers-color-scheme: dark)').matches)
  const [isMobile, setIsMobile] = useState(window.innerWidth < 640)
  const [dateFilter, setDateFilter] = useState('')
  const [transactionCategoryFilter, setTransactionCategoryFilter] = useState('all')
  const [overviewCategoryFilter, setOverviewCategoryFilter] = useState('all')
  const [editingId, setEditingId] = useState(null)
  const [editDraft, setEditDraft] = useState({
    date: '',
    description: '',
    category: '',
    amount: '',
    split: false,
    joint: false,
  })
  const [sortConfig, setSortConfig] = useState({ key: 'date', direction: 'desc' })
  const [jointSortConfig, setJointSortConfig] = useState({ key: 'date', direction: 'desc' })
  const [pendingImport, setPendingImport] = useState(null)

  const [partnerUser, setPartnerUser] = useState(null)
  const [partnerSplitTransactions, setPartnerSplitTransactions] = useState([])

  const [user, setUser] = useState(null)
  const [loginName, setLoginName] = useState('')
  const [loginPassword, setLoginPassword] = useState('')

  const [showAddModal, setShowAddModal] = useState(false)
  const [manualDraft, setManualDraft] = useState({
    date: new Date().toISOString().slice(0, 10),
    description: '',
    category: 'Other',
    amount: '',
    split: false,
    joint: false,
  })

  const fileRef = useRef(null)
  const jsonRef = useRef(null)

  const currentYear = String(new Date().getFullYear())
  const [kpiYear, setKpiYear] = useState(currentYear)

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', darkMode ? 'dark' : 'light')
  }, [darkMode])


  useEffect(() => {
  setTransactions([])
}, [])

  useEffect(() => {
  if (!user) return
  saveToIDB({ transactions })
}, [transactions, user])

useEffect(() => {
  if (user) loadPartnerData(user)
}, [user])

useEffect(() => {
  const saved = localStorage.getItem('splitoshl_user')
  if (!saved) return

  try {
    const parsed = JSON.parse(saved)
    if (parsed?.id && parsed?.username) {
      setUser(parsed)
      loadTransactionsFromCloud(parsed)
      loadPartnerData(parsed)
    }
  } catch {
    localStorage.removeItem('splitoshl_user')
  }
}, [])

useEffect(() => {
  if (!user?.id) return

  const channel = supabase
    .channel(`transactions-${user.id}`)
    .on(
      'postgres_changes',
      {
        event: '*',
        schema: 'public',
        table: 'transactions',
      },
      () => {
        loadTransactionsFromCloud(user)
        loadPartnerData(user)
      }
    )
    .subscribe()

  return () => {
    supabase.removeChannel(channel)
  }
}, [user])

  const myTransactions = useMemo(() => {
    return transactions.map((t) => {
      const amount = Math.abs(Number(t.amount || 0))
      const isJoint = t.account === 'joint' || t.joint
      const myAmount = isJoint ? amount / 2 : amount
      return { ...t, myAmount }
    })
  }, [transactions])

  const jointTransactions = useMemo(() => {
    return transactions
      .filter((t) => t.account === 'joint' || t.joint)
      .map((t) => {
        const amount = Math.abs(Number(t.amount || 0))
        return { ...t, jointAmount: amount }
      })
  }, [transactions])

  const filteredJointTransactions = useMemo(() => {
    const q = query.trim().toLowerCase()
    return jointTransactions.filter((t) => {
      const matchesQuery =
        !q ||
        [t.description, t.merchant, t.category, t.date, String(t.amount)]
          .join(' ')
          .toLowerCase()
          .includes(q)

      const matchesDate = !dateFilter || String(t.date).startsWith(dateFilter)
      const matchesCategory =
        transactionCategoryFilter === 'all' || (t.category || 'Other') === transactionCategoryFilter

      return matchesQuery && matchesDate && matchesCategory
    })
  }, [jointTransactions, query, dateFilter, transactionCategoryFilter])

  const jointTotal = useMemo(() => {
    return filteredJointTransactions.reduce((s, t) => s + Number(t.jointAmount || 0), 0)
  }, [filteredJointTransactions])

  const jointCategoryData = useMemo(() => {
    const m = new Map()
    filteredJointTransactions.forEach((t) => {
      m.set(t.category, (m.get(t.category) || 0) + Number(t.jointAmount || 0))
    })
    return [...m.entries()].sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }))
  }, [filteredJointTransactions])

  const jointMonthData = useMemo(() => {
    const m = new Map()
    filteredJointTransactions.forEach((t) => {
      const key = String(t.date).slice(0, 7)
      if (key.length === 7) m.set(key, (m.get(key) || 0) + Number(t.jointAmount || 0))
    })
    return [...m.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([name, value]) => ({ name, value }))
  }, [filteredJointTransactions])

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    return myTransactions.filter((t) => {
      const matchesQuery =
        !q ||
        [t.description, t.merchant, t.category, t.date, String(t.amount)]
          .join(' ')
          .toLowerCase()
          .includes(q)

      const matchesDate = !dateFilter || String(t.date).startsWith(dateFilter)
      const matchesTransactionCategory =
        transactionCategoryFilter === 'all' || (t.category || 'Other') === transactionCategoryFilter

      return matchesQuery && matchesDate && matchesTransactionCategory
    })
  }, [myTransactions, query, dateFilter, transactionCategoryFilter])

  const overviewFiltered = useMemo(() => {
    return myTransactions.filter((t) => {
      return overviewCategoryFilter === 'all' || (t.category || 'Other') === overviewCategoryFilter
    })
  }, [myTransactions, overviewCategoryFilter])

  const sortedTransactions = useMemo(() => {
    const arr = [...filtered]
    const { key, direction } = sortConfig
    const dir = direction === 'asc' ? 1 : -1

    arr.sort((a, b) => {
      if (key === 'amount') {
        const av = Number(a.amount || 0)
        const bv = Number(b.amount || 0)
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
  }, [filtered, sortConfig])

  function handleSort(key) {
    setSortConfig((prev) => {
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
    setJointSortConfig((prev) => {
      if (prev.key === key) {
        return { key, direction: prev.direction === 'asc' ? 'desc' : 'asc' }
      }
      return { key, direction: key === 'date' ? 'desc' : 'asc' }
    })
  }

  const myTotalSpend = useMemo(() => filtered.reduce((s, t) => s + Number(t.myAmount || 0), 0), [filtered])

  const categoryData = useMemo(() => {
    const m = new Map()
    filtered.forEach((t) => {
      m.set(t.category, (m.get(t.category) || 0) + Number(t.myAmount || 0))
    })
    return [...m.entries()].sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }))
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
    filtered.forEach((t) => {
      const key = String(t.date).slice(0, 7)
      if (key.length === 7) m.set(key, (m.get(key) || 0) + Number(t.myAmount || 0))
    })
    return [...m.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([name, value]) => ({ name, value }))
  }, [filtered])

  const forecast = useMemo(
    () => buildForecast(myTransactions.map((t) => ({ ...t, amount: t.myAmount }))),
    [myTransactions]
  )

  // My open split transactions (unpaid)
const myOpenSplitTransactions = useMemo(
  () => transactions.filter((t) => t.split && !t.splitPaid),
  [transactions]
)

const partnerOpenSplitTotal = useMemo(
  () =>
    partnerSplitTransactions.reduce((sum, t) => {
      const half = Math.abs(Number(t.amount) || 0) / 2
      if (t.splitDirection === 'owed_to_me') return sum + half
      if (t.splitDirection === 'i_owe') return sum - half
      return sum
    }, 0),
  [partnerSplitTransactions]
)

const myOpenSplitTotal = useMemo(
  () =>
    myOpenSplitTransactions.reduce((sum, t) => {
      const half = Math.abs(Number(t.amount) || 0) / 2
      if (t.splitDirection === 'owed_to_me') return sum + half
      if (t.splitDirection === 'i_owe') return sum - half
      return sum
    }, 0),
  [myOpenSplitTransactions]
)

// Net: positive = partner owes you, negative = you owe partner
const netSplitBalance = useMemo(
  () => myOpenSplitTotal,
  [myOpenSplitTotal]
)

const splitBalanceLabel = useMemo(() => {
  if (netSplitBalance > 0.004) return `${partnerUser?.displayName || 'Partner'} owes you`
  if (netSplitBalance < -0.004) return `You owe ${partnerUser?.displayName || 'Partner'}`
  return 'You are settled'
}, [netSplitBalance, partnerUser])

const splitBalanceValue = Math.abs(netSplitBalance)

// Monthly rows: my side per month
const splitRows = useMemo(() => {
  const m = new Map()
  myOpenSplitTransactions.forEach((t) => {
    const month = String(t.date).slice(0, 7)
    if (month.length !== 7) return
    const half = Math.abs(Number(t.amount) || 0) / 2
    const signedHalf =
      t.splitDirection === 'owed_to_me' ? half : t.splitDirection === 'i_owe' ? -half : 0
    m.set(month, (m.get(month) || 0) + signedHalf)
  })
  return [...m.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([month, owed]) => ({ month, owed }))
}, [myOpenSplitTransactions])

const splitTotal = useMemo(() => splitRows.reduce((s, r) => s + r.owed, 0), [splitRows])


  const getSubscriptionKeyFromDescription = (description) => {
    const d = String(description || '').toLowerCase()
    const m = d.match(
      /netflix|spotify|prime video|amazon prime|apple tv|apple music|hbo|disney|paramount|youtube|icloud|dropbox|dazn|patreon|subscription/i
    )
    if (m) return m[0].toLowerCase()
    const first = d.trim().split(' ')[0]
    return first || 'subscription'
  }

  const isSubscriptionCandidate = (t) => {
    const d = String(t.description || '')
    return /netflix|spotify|prime|apple|hbo|disney|subscription|adobe|microsoft|google one|icloud|dropbox|urban sports|revolut|youtube/i.test(
      d
    )
  }

  const subscriptions = useMemo(() => {
    const now = new Date()
    const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`
    const candidates = filtered.filter(isSubscriptionCandidate)
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

    for (const [, txs] of groups.entries()) {
      if (!txs.length) continue

      const sorted = txs
        .filter((t) => t.date)
        .slice()
        .sort((a, b) => (a.date > b.date ? -1 : 1))

      if (!sorted.length) continue

      const latest = sorted[0]
      const lastDate = latest.date
      const lastMonth = getMonthKey(lastDate)
      if (!lastMonth) continue

      const diff = monthDiff(currentMonthKey, lastMonth)
      const subscriptionStatus = diff >= 2 ? 'over' : 'active'
      const totalAmount = txs.reduce((sum, t) => sum + Math.abs(Number(t.amount || 0)), 0)
      const lastAmount = Math.abs(Number(latest.amount || 0))

      summaries.push({
        ...latest,
        subscriptionStatus,
        totalAmount,
        lastAmount,
      })
    }

    summaries.sort((a, b) => {
      if (a.date < b.date) return 1
      if (a.date > b.date) return -1
      return 0
    })

    return summaries
  }, [filtered])

  const categoryOptions = useMemo(() => {
    const existing = [...new Set(transactions.map((t) => t.category).filter(Boolean))]
    return [...new Set([...DEFAULTCATEGORIES, ...existing])].sort((a, b) => a.localeCompare(b))
  }, [transactions])

  const getCategoryColor = (name) => {
    const idx = categoryOptions.indexOf(name)
    const safeIdx = idx === -1 ? 0 : idx
    return PALETTE[safeIdx % PALETTE.length]
  }

  const availableYears = useMemo(() => {
  const years = [
    ...new Set(
      transactions
        .map((t) => String(t.date || '').slice(0, 4))
        .filter((y) => /^\d{4}$/.test(y))
    ),
  ].sort((a, b) => Number(b) - Number(a))

  return years.length ? years : [currentYear]
}, [transactions, currentYear])

useEffect(() => {
  if (!availableYears.includes(kpiYear)) {
    setKpiYear(currentYear)
  }
}, [availableYears, kpiYear, currentYear])

const kpiTransactions = useMemo(() => {
  return transactions.filter((t) => String(t.date || '').slice(0, 4) === kpiYear)
}, [transactions, kpiYear])

const kpiMyTransactions = useMemo(() => {
  return kpiTransactions.map((t) => {
    const amount = Math.abs(Number(t.amount || 0))
    const isJoint = t.account === 'joint' || t.joint
    const myAmount = isJoint ? amount / 2 : amount
    return { ...t, myAmount }
  })
}, [kpiTransactions])

const kpiJointTransactions = useMemo(() => {
  return kpiTransactions
    .filter((t) => t.account === 'joint' || t.joint)
    .map((t) => {
      const amount = Math.abs(Number(t.amount || 0))
      return { ...t, jointAmount: amount }
    })
}, [kpiTransactions])

const kpiMySpend = useMemo(
  () => kpiMyTransactions.reduce((s, t) => s + Number(t.myAmount || 0), 0),
  [kpiMyTransactions]
)

const kpiTransactionCount = useMemo(
  () => kpiTransactions.length,
  [kpiTransactions]
)

const kpiCategoryCount = useMemo(
  () => new Set(kpiTransactions.map((t) => t.category || 'Other')).size,
  [kpiTransactions]
)

const kpiJointSpend = useMemo(
  () => kpiJointTransactions.reduce((s, t) => s + Number(t.jointAmount || 0), 0),
  [kpiJointTransactions]
)
  
  const hasData = transactions.length > 0

 async function signInWithNamePassword(username, password) {
  const cleanUsername = String(username).trim().toLowerCase()
  const cleanPassword = String(password).trim()
  if (!cleanUsername || !cleanPassword) { setStatus('Enter username and password.'); return }
  const { data, error } = await supabase
    .from('app_users')
    .select('*')
    .ilike('username', cleanUsername)
    .eq('password', cleanPassword)
  if (error) { setStatus('Login failed: ' + error.message); return }
  if (!data || data.length === 0) { setStatus('No matching user found.'); return }
  const row = data[0]
  const appUser = {
    id: row.id,
    username: String(row.username || '').trim().toLowerCase(),
    displayName: row.display_name || row.username,
    partnerUsername: String(row.partner_username || '').trim().toLowerCase(),
  }
  setUser(appUser)
localStorage.setItem('splitoshl_user', JSON.stringify(appUser))
setLoginPassword('')
setStatus('Signed in as ' + appUser.displayName + '.')
await loadTransactionsFromCloud(appUser)
await loadPartnerData(appUser)
}

async function loadPartnerData(currentUser) {
  if (!currentUser?.partnerUsername) {
    setPartnerUser(null)
    setPartnerSplitTransactions([])
    return
  }
  const { data: userData, error: userErr } = await supabase
    .from('app_users')
    .select('*')
    .ilike('username', currentUser.partnerUsername)
    .limit(1)
  if (userErr || !userData || userData.length === 0) {
    setPartnerUser(null)
    setPartnerSplitTransactions([])
    return
  }
  const pRow = userData[0]
  const partner = {
    id: pRow.id,
    username: String(pRow.username || '').trim().toLowerCase(),
    displayName: pRow.display_name || pRow.username,
  }
  setPartnerUser(partner)

  const { data: txData, error: txErr } = await supabase
    .from('transactions')
    .select('*')
    .eq('userid', partner.id)
    .eq('split', true)
    .eq('splitpaid', false)
  if (txErr) { setPartnerSplitTransactions([]); return }
  setPartnerSplitTransactions(
  txData.map((t) => ({
    id: t.id,
    date: t.date,
    description: t.description,
    merchant: t.merchant || t.description,
    amount: Number(t.amount || 0),
    category: t.category || 'Other',
    split: Boolean(t.split),
    splitPaid: Boolean(t.splitpaid),
    splitDirection: t.split_direction || null,
    splitGroupId: t.split_group_id || null,
  }))
)
}

async function syncJointTransaction(tx, currentUser) {
  if (!currentUser?.id) return null

  const partnerId = getPartnerUserId(currentUser)
  if (!partnerId) return null

  const jointGroupId = makeJointGroupId(tx)

  const base = {
    date: tx.date,
    description: tx.description,
    merchant: tx.merchant || tx.description,
    amount: Number(tx.amount) || 0,
    category: tx.category || 'Other',
    split: Boolean(tx.split),
    splitpaid: Boolean(tx.splitPaid),
    joint: true,
    joint_mode: 'full',
    account: 'joint',
    joint_group_id: jointGroupId,
  }

  const myRow = {
    ...base,
    id: String(tx.id).replace(/__u[12]$/, ''),
    userid: currentUser.id,
    created_by_user_id: currentUser.id,
    shared_with_user_id: partnerId,
  }

  const partnerRow = {
  id: partnerRowId,
  userid: partnerId,
  date: tx.date,
  description: tx.description,
  merchant: tx.merchant || tx.description,
  amount: Number(tx.amount) || 0,
  category: tx.category || 'Other',
  split: Boolean(tx.split),
  splitpaid: Boolean(tx.splitPaid),
  split_direction: tx.split ? 'i_owe' : null,
  joint: true,
  joint_mode: 'full',
  account: 'joint',
  joint_group_id: jointGroupId,
  created_by_user_id: user.id,
  shared_with_user_id: user.id,
}

  const { error } = await supabase
    .from('transactions')
    .upsert([myRow, partnerRow], { onConflict: 'id' })

  return error
}

async function unsyncJointTransaction(tx, currentUser) {
  if (!currentUser?.id) return null

  const partnerId = getPartnerUserId(currentUser)
  const jointGroupId = tx.jointGroupId || makeJointGroupId(tx)

  const myBaseId = String(tx.id).replace(/__u[12]$/, '')
  const partnerRowId = `${myBaseId}__u${partnerId}`

  const { error: partnerDeleteError } = await supabase
    .from('transactions')
    .delete()
    .eq('id', partnerRowId)

  if (partnerDeleteError) return partnerDeleteError

  const { error: myUpdateError } = await supabase
    .from('transactions')
    .update({
      joint: false,
      joint_mode: null,
      account: 'personal',
      joint_group_id: null,
      created_by_user_id: null,
      shared_with_user_id: null,
    })
    .eq('id', tx.id)

  return myUpdateError
}

  async function signOutUser() {
  await clearIDB()
  localStorage.removeItem('splitoshl_user')
  setUser(null)
  setTransactions([])
  setLoginName('')
  setLoginPassword('')
  setEditingId(null)
  setPendingImport(null)
  setShowAddModal(false)
  setQuery('')
  setDateFilter('')
  setTransactionCategoryFilter('all')
  setOverviewCategoryFilter('all')
  setStatus('Signed out.')
  setPartnerUser(null)
  setPartnerSplitTransactions([])
}

  async function loadTransactionsFromCloud(currentUser) {
    if (!currentUser?.id) return

    const pageSize = 1000
    let from = 0
    let allRows = []
    let keepLoading = true

    while (keepLoading) {
   const { data, error } = await supabase
  .from('transactions')
  .select('*')
  .eq('userid', currentUser.id)
  .order('date', { ascending: false })
  .range(from, from + pageSize - 1)

      if (error) {
        setStatus(`Cloud load failed: ${error.message}`)
        return
      }

      const rows = data || []
      allRows = [...allRows, ...rows]

      if (rows.length < pageSize) {
        keepLoading = false
      } else {
        from += pageSize
      }
    }

 const mapped = allRows.map((t) => ({
  id: t.id,
  date: t.date,
  description: t.description,
  merchant: t.merchant || t.description,
  amount: Number(t.amount || 0),
  category: t.category || 'Other',
  split: Boolean(t.split),
  splitPaid: Boolean(t.splitpaid),
  joint: Boolean(t.joint),
  jointMode: null,
  account: t.account || 'personal',
  jointGroupId: t.joint_group_id || null,
  splitGroupId: t.split_group_id || null,
  createdByUserId: t.created_by_user_id || null,
  sharedWithUserId: t.shared_with_user_id || null,
  splitDirection: t.split_direction || null,
}))

    const normalized = dedup(normalizeTransactions(mapped))
    setTransactions(normalized)
    setStatus(`Loaded ${normalized.length} cloud transactions for ${currentUser.displayName || currentUser.username}.`)
  }

async function addTransactionToCloud(tx, currentUser) {
  const { data, error } = await supabase
    .from('transactions')
    .upsert(
      {
  id: tx.id,
  userid: currentUser.id,
  date: tx.date,
  description: tx.description,
  merchant: tx.merchant,
  amount: tx.amount,
  category: tx.category,
  split: tx.split,
  splitpaid: tx.splitPaid,
  joint: tx.joint,
  joint_mode: tx.jointMode || null,
  account: tx.account || 'personal',
  joint_group_id: tx.jointGroupId || null,
  split_group_id: tx.splitGroupId || null,
  created_by_user_id: tx.createdByUserId || null,
  shared_with_user_id: tx.sharedWithUserId || null,
  split_direction: tx.splitDirection || null,
  splitGroupId: t.split_group_id || null,
},
      { onConflict: 'userid,id' }
    )
    .select()

  return { data, error }
}

async function updateTransactionInCloud(tx, currentUser) {
  const { error } = await supabase
    .from('transactions')
    .update({
  date: tx.date,
  description: tx.description,
  merchant: tx.merchant,
  amount: tx.amount,
  category: tx.category,
  split: tx.split,
  splitpaid: tx.splitPaid,
  joint: tx.joint,
  joint_mode: tx.jointMode || null,
  account: tx.account || 'personal',
  joint_group_id: tx.jointGroupId || null,
  split_group_id: tx.splitGroupId || null,
  created_by_user_id: tx.createdByUserId || null,
  shared_with_user_id: tx.sharedWithUserId || null,
  split_direction: tx.splitDirection || null,
  splitGroupId: t.split_group_id || null,
})
    .eq('userid', currentUser.id)
    .eq('id', tx.id)

  return error
}

async function deleteTransactionFromCloud(id, currentUser) {
  const { error } = await supabase
    .from('transactions')
    .delete()
    .eq('userid', currentUser.id)
    .eq('id', id)

  return error
}

  function importRows(rows, label) {
    const hasAccountFlag = rows.some((r) => r.account)
    if (hasAccountFlag) finishImport(rows, label, null)
    else setPendingImport({ rows, label })
  }

  async function finishImport(rows, label, accountType) {
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
    joint: t.account === 'joint' ? true : isJoint ? true : Boolean(t.joint),
    jointMode: t.account === 'joint' ? 'full' : isJoint ? 'full' : t.joint ? t.jointMode || 'full' : null,
  }))

  const merged = dedup([...transactions, ...enriched])

  if (merged.length === transactions.length) {
    setStatus(`No new transactions found in ${label}; all were duplicates.`)
    setPendingImport(null)
    return
  }

  if (user) {
  for (const t of enriched) {
    const shouldBeJoint = Boolean(t.joint || t.account === 'joint')

    const baseTx = {
      ...t,
      joint: false,
      jointMode: null,
      account: 'personal',
      jointGroupId: null,
      createdByUserId: null,
      sharedWithUserId: null,
    }

    const { data, error } = await addTransactionToCloud(baseTx, user)

    if (error) {
      setStatus(`Imported locally, but cloud save failed: ${error.message}`)
      setTransactions(merged)
      setEditingId(null)
      setActiveTab('Own')
      setPendingImport(null)
      return
    }

    let savedTx = {
      ...baseTx,
      ...data?.[0],
      splitPaid: Boolean(data?.[0]?.splitpaid),
      jointGroupId: data?.[0]?.joint_group_id || null,
      createdByUserId: data?.[0]?.created_by_user_id || null,
      sharedWithUserId: data?.[0]?.shared_with_user_id || null,
      jointMode: data?.[0]?.joint_mode || null,
      account: data?.[0]?.account || 'personal',
    }

    if (shouldBeJoint) {
      await handleJointToggle(savedTx)
    }
  }
}

  setTransactions(merged)
  setEditingId(null)
  setStatus(
    user
      ? `Merged ${merged.length - transactions.length} new transactions from ${label} and saved online.`
      : `Merged ${merged.length - transactions.length} new transactions from ${label}.`
  )
  setActiveTab('Own')
  setPendingImport(null)
}

  async function handleFile(file) {
    if (!file) return
    const ext = file.name.split('.').pop()?.toLowerCase()
    setStatus(`Reading ${file.name}...`)

    try {
      if (ext === 'xlsx' || ext === 'xls') {
        const rows = await parseXLSXfile(file)
        importRows(rows, file.name)
      } else if (ext === 'csv' || ext === 'txt') {
        const rows = await parseCSVfile(file)
        importRows(rows, file.name)
      } else {
        setStatus(`Unsupported file type ".${ext}". Please use .xlsx or .csv.`)
      }
    } catch (err) {
      setStatus(`Error reading file: ${err.message}`)
    }
  }

async function handleJointToggle(tx) {
  try {
    if (!user?.id) {
      setStatus('Please log in first.')
      return
    }

    const partnerId = user.id === 1 ? 2 : 1
    const baseId = String(tx.id)
    const nextIsJoint = !(tx.joint || tx.account === 'joint')
    const jointGroupId = tx.jointGroupId || `joint_${baseId}`
    const splitGroupId = tx.split ? (tx.splitGroupId || makeSplitGroupId(tx)) : null

    if (nextIsJoint) {
      const myUpdatePayload = {
  joint: true,
  joint_mode: 'full',
  account: 'joint',
  joint_group_id: jointGroupId,
  split_group_id: splitGroupId,
  split_direction: tx.split ? 'owed_to_me' : null,
  created_by_user_id: user.id,
  shared_with_user_id: partnerId,
  split_group_id: tx.split ? (tx.splitGroupId || jointGroupId) : null,
}

      console.log('Updating my row:', tx.id, myUpdatePayload)

      const { error: updateMineError } = await supabase
        .from('transactions')
        .update(myUpdatePayload)
        .eq('id', tx.id)

      if (updateMineError) {
        console.error('updateMineError:', updateMineError)
        setStatus(`Failed to update your row: ${updateMineError.message}`)
        return
      }

      const partnerRowId = `${baseId}__u${partnerId}`

      const partnerRow = {
  id: partnerRowId,
  userid: partnerId,
  date: tx.date,
  description: tx.description,
  merchant: tx.merchant || tx.description,
  amount: Number(tx.amount) || 0,
  category: tx.category || 'Other',
  split: Boolean(tx.split),
  splitpaid: Boolean(tx.splitPaid),
  split_direction: tx.split ? 'i_owe' : null,
  split_group_id: splitGroupId,
  joint: true,
  joint_mode: 'full',
  account: 'joint',
  joint_group_id: jointGroupId,
  created_by_user_id: user.id,
  shared_with_user_id: user.id,
  split_group_id: tx.split ? (tx.splitGroupId || jointGroupId) : null,
}

      console.log('Attempting partner insert:', partnerRow)

      const { error: partnerInsertError, data: partnerInsertData } = await supabase
        .from('transactions')
        .insert([partnerRow])
        .select()

      console.log('partnerInsertData:', partnerInsertData)

      if (partnerInsertError) {
        console.error('partnerInsertError:', partnerInsertError)

        const { error: partnerUpdateError, data: partnerUpdateData } = await supabase
          .from('transactions')
          .update({
  date: tx.date,
  description: tx.description,
  merchant: tx.merchant || tx.description,
  amount: Number(tx.amount) || 0,
  category: tx.category || 'Other',
  split: Boolean(tx.split),
  splitpaid: Boolean(tx.splitPaid),
  split_direction: tx.split ? 'i_owe' : null,
  split_group_id: splitGroupId,
  joint: true,
  joint_mode: 'full',
  account: 'joint',
  joint_group_id: jointGroupId,
  created_by_user_id: user.id,
  shared_with_user_id: user.id,
})
          .eq('id', partnerRowId)
          .eq('userid', partnerId)
          .select()

        console.log('partnerUpdateData:', partnerUpdateData)

        if (partnerUpdateError) {
          console.error('partnerUpdateError:', partnerUpdateError)
          setStatus(`Failed to create/update partner row: ${partnerUpdateError.message}`)
          return
        }
      }

      setStatus('Joint transaction synced to both users.')
      await loadTransactionsFromCloud(user)
      return
    }

    const partnerRowId = `${baseId}__u${partnerId}`

    const { error: deletePartnerError } = await supabase
      .from('transactions')
      .delete()
      .eq('id', partnerRowId)

    if (deletePartnerError) {
      console.error('deletePartnerError:', deletePartnerError)
      setStatus(`Failed to remove partner row: ${deletePartnerError.message}`)
      return
    }

    const resetPayload = {
      joint: false,
      joint_mode: null,
      account: 'personal',
      joint_group_id: null,
      created_by_user_id: null,
      shared_with_user_id: null,
    }

    console.log('Resetting my row:', tx.id, resetPayload)

    const { error: resetMineError } = await supabase
      .from('transactions')
      .update(resetPayload)
      .eq('id', tx.id)

    if (resetMineError) {
      console.error('resetMineError:', resetMineError)
      setStatus(`Failed to reset your row: ${resetMineError.message}`)
      return
    }

    setStatus('Joint transaction removed from both users.')
    await loadTransactionsFromCloud(user)
  } catch (err) {
    console.error('handleJointToggle crash:', err)
    setStatus(`Joint toggle crashed: ${err.message}`)
  }
}

  async function handleJSONfile(file) {
    if (!file) return

    try {
      const data = await importJSON(file)
      if (data?.transactions?.length) {
        await importRows(data.transactions, 'JSON')
        setStatus(`Imported ${data.transactions.length} transactions from JSON.`)
      } else {
        setStatus('JSON file has no transactions.')
      }
    } catch (err) {
      setStatus(`Failed to parse JSON: ${err.message}`)
    }
  }

  function handleExport() {
    exportJSON({ transactions }, 'expense-data.json').then(() => setStatus('JSON exported.'))
  }

  function handlePDF() {
    generatePDFReport(transactions)
    setStatus('PDF downloaded.')
  }

  function handleSplitPDF() {
  generateSplitPDFReport({
    currentUserName: user?.displayName || user?.username || 'You',
    partnerUserName: partnerUser?.displayName || partnerUser?.username || 'Partner',
    mySplitTransactions: myOpenSplitTransactions,
    partnerSplitTransactions,
    myOpenSplitTotal,
    partnerOpenSplitTotal,
    netSplitBalance,
    splitBalanceLabel,
    splitBalanceValue,
  })
  setStatus('Split PDF downloaded.')
}

async function handleClearAll() {
  if (!user) {
    setTransactions([])
    await clearIDB()
    setStatus('Local data cleared.')
    return
  }

  const confirmed = window.confirm(
    'This will permanently delete all your transactions from this account, both locally and online. Continue?'
  )

  if (!confirmed) return

  const error = await clearTransactionsFromCloud(user)

  if (error) {
    setStatus(`Failed to clear online data: ${error.message}`)
    return
  }

  setTransactions([])
  await clearIDB()
  setStatus('All transactions cleared locally and online.')
}

  function openManualAdd() {
    setManualDraft({
      date: new Date().toISOString().slice(0, 10),
      description: '',
      category: 'Other',
      amount: '',
      split: false,
      joint: false,
    })
    setShowAddModal(true)
  }

  function closeManualAdd() {
    setShowAddModal(false)
  }

  async function saveManualExpense() {
    const rawAmount = String(manualDraft.amount ?? '').trim()
    const normalizedAmount = rawAmount.replace(/\s/g, '').replace('€', '').replace(',', '.')
    const amount = Number(normalizedAmount)

    if (!manualDraft.date) {
      setStatus('Date is required.')
      return
    }

    if (!manualDraft.description.trim()) {
      setStatus('Description is required.')
      return
    }

    if (!Number.isFinite(amount) || amount <= 0) {
      setStatus('Amount must be a valid positive number.')
      return
    }

    const baseId = `${manualDraft.date}-${manualDraft.description.trim()}-${amount}-${Date.now()}`

const newTransaction = {
  id: baseId,
  date: manualDraft.date,
  description: manualDraft.description.trim(),
  merchant: manualDraft.description.trim(),
  amount,
  category: manualDraft.category || 'Other',
  split: Boolean(manualDraft.split),
  splitDirection: manualDraft.split ? 'owed_to_me' : null,
  splitGroupId: manualDraft.split ? makeSplitGroupId({ id: baseId, date: manualDraft.date, description: manualDraft.description.trim() }) : null,
  splitPaid: false,
  joint: false,
  jointMode: null,
  account: 'personal',
}

let savedTransaction = newTransaction

if (user) {
  const { data, error } = await addTransactionToCloud(newTransaction, user)
  if (error) {
    setStatus(`Cloud save failed: ${error.message}`)
    return
  }

  if (data?.[0]) {
    savedTransaction = {
      ...newTransaction,
      ...data[0],
      splitPaid: Boolean(data[0].splitpaid),
      jointGroupId: data[0].joint_group_id || null,
      createdByUserId: data[0].created_by_user_id || null,
      sharedWithUserId: data[0].shared_with_user_id || null,
      jointMode: data[0].joint_mode || null,
      account: data[0].account || 'personal',
    }
  }

  if (manualDraft.joint) {
    savedTransaction = {
      ...savedTransaction,
      joint: false,
      jointMode: null,
      account: 'personal',
      jointGroupId: null,
      createdByUserId: null,
      sharedWithUserId: null,
    }

    await handleJointToggle(savedTransaction)
    setShowAddModal(false)
    setEditingId(null)
    setStatus('Manual expense added and synced to both users.')
    setActiveTab('Transactions')
    return
  }
}

setTransactions((prev) => dedup([...prev, savedTransaction]))
setShowAddModal(false)
setEditingId(null)
setStatus(user ? 'Manual expense added and saved online.' : 'Manual expense added locally.')
setActiveTab('Transactions')
}

async function clearTransactionsFromCloud(currentUser) {
  if (!currentUser?.id) return null

  const { error } = await supabase
    .from('transactions')
    .delete()
    .eq('userid', currentUser.id)

  return error
}

async function handleClear() {
  if (!user) {
    setTransactions([])
    await clearIDB()
    setStatus('Local data cleared.')
    return
  }

  const confirmed = window.confirm(
    'This will permanently delete all your transactions from this account, both locally and online. Continue?'
  )

  if (!confirmed) return

  const error = await clearTransactionsFromCloud(user)

  if (error) {
    setStatus(`Failed to clear online data: ${error.message}`)
    return
  }

  setTransactions([])
  await clearIDB()
  setStatus('All transactions cleared locally and online.')
}

  async function markAllSplitsPaid() {
  if (!user) {
    setStatus('You need to be signed in to mark splits as paid.')
    return
  }

  const openSplitRows = transactions.filter((t) => t.split && !t.splitPaid)

  if (openSplitRows.length === 0) {
    setStatus('No split transactions to mark as paid.')
    return
  }

  const splitGroupIds = [
    ...new Set(
      openSplitRows
        .map((t) => t.splitGroupId || t.jointGroupId)
        .filter(Boolean)
    ),
  ]

  if (splitGroupIds.length === 0) {
    setStatus('No split groups found to mark as paid.')
    return
  }

  const { error } = await supabase
    .from('transactions')
    .update({
      split: false,
      splitpaid: true,
    })
    .in('split_group_id', splitGroupIds)

  if (error) {
    setStatus(`Cloud save failed: ${error.message}`)
    return
  }

  await loadTransactionsFromCloud(user)
  await loadPartnerData(user)

  setStatus(`Marked ${splitGroupIds.length} split group(s) as paid for both users.`)
}

async function toggleSplit(id) {
  const tx = transactions.find((t) => t.id === id)
  if (!tx) return

  if (tx.splitPaid) {
    setStatus('This transaction was already split and paid in the past.')
    return
  }

  if (!user?.id) {
    const updatedTransaction = {
      ...tx,
      split: !tx.split,
      splitPaid: false,
      splitDirection: !tx.split ? 'owed_to_me' : null,
    }

    setTransactions((prev) =>
      prev.map((t) => (t.id === id ? updatedTransaction : t))
    )

    setStatus(
      updatedTransaction.split
        ? 'Split updated locally.'
        : 'Split removed locally.'
    )
    return
  }

  try {
    const partnerId = getPartnerUserId(user)

    if (!partnerId) {
      setStatus('Partner user not found.')
      return
    }

    const isJoint = tx.joint || tx.account === 'joint'

    if (isJoint) {
      const splitGroupId = tx.splitGroupId || tx.jointGroupId || `split_${tx.id}`

      if (!tx.split) {
        const { data: paidRows, error: paidCheckError } = await supabase
          .from('transactions')
          .select('id')
          .eq('joint_group_id', tx.jointGroupId)
          .eq('splitpaid', true)
          .limit(1)

        if (paidCheckError) {
          setStatus(`Split check failed: ${paidCheckError.message}`)
          return
        }

        if (paidRows && paidRows.length > 0) {
          setStatus('This joint transaction was already settled and cannot be split again.')
          return
        }

        const { error: myError } = await supabase
          .from('transactions')
          .update({
            split: true,
            splitpaid: false,
            split_group_id: splitGroupId,
            split_direction: 'owed_to_me',
          })
          .eq('joint_group_id', tx.jointGroupId)
          .eq('userid', user.id)

        if (myError) {
          setStatus(`Cloud save failed: ${myError.message}`)
          return
        }

        const { error: partnerError } = await supabase
          .from('transactions')
          .update({
            split: true,
            splitpaid: false,
            split_group_id: splitGroupId,
            split_direction: 'i_owe',
          })
          .eq('joint_group_id', tx.jointGroupId)
          .eq('userid', partnerId)

        if (partnerError) {
          setStatus(`Partner update failed: ${partnerError.message}`)
          return
        }

        await loadTransactionsFromCloud(user)
        await loadPartnerData(user)
        setStatus('Split enabled for both users.')
        return
      }

      const { error: unsplitError } = await supabase
        .from('transactions')
        .update({
          split: false,
          splitpaid: false,
          split_group_id: null,
          split_direction: null,
        })
        .eq('joint_group_id', tx.jointGroupId)
        .eq('splitpaid', false)

      if (unsplitError) {
        setStatus(`Cloud save failed: ${unsplitError.message}`)
        return
      }

      await loadTransactionsFromCloud(user)
      await loadPartnerData(user)
      setStatus('Split removed for both users.')
      return
    }

    const updatedTransaction = {
      ...tx,
      split: !tx.split,
      splitPaid: false,
      splitGroupId: !tx.split ? (tx.splitGroupId || `split_${tx.id}`) : null,
      splitDirection: !tx.split ? 'owed_to_me' : null,
    }

    const error = await updateTransactionInCloud(updatedTransaction, user)
    if (error) {
      setStatus(`Cloud save failed: ${error.message}`)
      return
    }

    setTransactions((prev) =>
      prev.map((t) => (t.id === id ? updatedTransaction : t))
    )

    await loadPartnerData(user)

    setStatus(
      updatedTransaction.split
        ? 'Split updated and saved online.'
        : 'Split removed and saved online.'
    )
  } catch (err) {
    setStatus(`Split toggle failed: ${err.message}`)
  }
}

async function toggleJoint(id) {
  const tx = transactions.find((t) => t.id === id)
  if (!tx) return

  if (tx.account === 'joint') {
    setStatus('This transaction already comes from a Joint statement.')
    return
  }

  const updatedTransaction = tx.joint
    ? {
        ...tx,
        joint: false,
        jointMode: null,
        account: 'personal',
      }
    : {
        ...tx,
        joint: true,
        jointMode: null,
        account: 'joint',
      }

  if (user) {
    const error = await updateTransactionInCloud(updatedTransaction, user)
    if (error) {
      setStatus(`Cloud save failed: ${error.message}`)
      return
    }
  }

  setTransactions((prev) =>
    prev.map((t) => (t.id === id ? updatedTransaction : t))
  )

  setStatus(
    updatedTransaction.joint
      ? 'Added to Joint tab and saved online.'
      : 'Removed from Joint tab and saved online.'
  )
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

async function saveEdit(id) {
  const amount = Number(String(editDraft.amount).replace(',', '.'))

  if (!editDraft.date) {
    setStatus('Date is required.')
    return
  }

  if (!Number.isFinite(amount) || amount <= 0) {
    setStatus('Amount must be a valid positive number.')
    return
  }

  const original = transactions.find((t) => t.id === id)
  if (!original) {
    setStatus('Transaction not found.')
    return
  }

  const nextSplit = original.splitPaid ? false : Boolean(editDraft.split)

const updatedTransaction = {
  ...original,
  date: editDraft.date,
  description: editDraft.description.trim() || original.description,
  merchant: editDraft.description.trim() || original.merchant,
  category: editDraft.category || 'Other',
  amount,
  split: nextSplit,
  splitPaid: Boolean(original.splitPaid),
  splitDirection: nextSplit ? (original.splitDirection || 'owed_to_me') : null,
  splitGroupId: nextSplit ? (original.splitGroupId || makeSplitGroupId(original)) : original.splitGroupId,
  joint: original.account === 'joint' ? true : Boolean(editDraft.joint),
  jointMode: null,
  account: original.account === 'joint'
    ? 'joint'
    : (Boolean(editDraft.joint) ? 'joint' : 'personal'),
}

  if (user) {
    const error = await updateTransactionInCloud(updatedTransaction, user)
    if (error) {
      setStatus(`Cloud save failed: ${error.message}`)
      return
    }
  }

  const categoryChanged = original.category !== (editDraft.category || 'Other')
  const originalDescription = original.description

  setTransactions((prev) =>
    prev.map((t) => {
      if (t.id === id) return updatedTransaction

      if (categoryChanged && originalDescription && t.description === originalDescription) {
        return { ...t, category: editDraft.category || 'Other' }
      }

      return t
    })
  )

  if (categoryChanged && originalDescription) {
    const matchCount = transactions.filter(
      (t) => t.id !== id && t.description === originalDescription
    ).length

    if (matchCount > 0) {
      setStatus(
        `Updated. Category ${editDraft.category || 'Other'} applied to ${matchCount + 1} transactions with description "${originalDescription}".`
      )
    } else {
      setStatus('Transaction updated.')
    }
  } else {
    setStatus('Transaction updated.')
  }

  cancelEdit()
}

async function deleteTransaction(id) {
  if (user) {
    const error = await deleteTransactionFromCloud(id, user)
    if (error) {
      setStatus(`Cloud delete failed: ${error.message}`)
      return
    }
  }

  setTransactions((prev) => prev.filter((t) => t.id !== id))

  if (editingId === id) cancelEdit()

  setStatus(user ? 'Transaction deleted and removed online.' : 'Transaction deleted.')
}

  const forecastByCategory = useMemo(() => {
    const byMonth = new Map()

    transactions.forEach((t) => {
      const month = getMonthKey(t.date)
      if (!month) return
      const cat = t.category || 'Other'
      const key = `${cat}|${month}`
      byMonth.set(key, (byMonth.get(key) || 0) + Math.abs(Number(t.amount || 0)))
    })

    const months = [...new Set([...byMonth.keys()].map((k) => k.split('|')[1]))].sort()
    const last6 = months.slice(-6)
    if (last6.length === 0) return []

    const last6Set = new Set(last6)
    const catTotals = new Map()

    byMonth.forEach((val, key) => {
      const [cat, month] = key.split('|')
      if (last6Set.has(month)) catTotals.set(cat, (catTotals.get(cat) || 0) + val)
    })

    return [...catTotals.entries()].map(([category, total]) => {
      const avgPerMonth = total / last6.length
      const projected3 = avgPerMonth * 3
      return { category, avgPerMonth, projected3 }
    })
  }, [transactions])

  const mostExpensiveCategory = useMemo(() => {
    if (!forecastByCategory.length) return null
    return forecastByCategory.reduce((max, row) => (!max || row.avgPerMonth > max.avgPerMonth ? row : max), null)
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
          <button type="button" className="btn btn-primary" onClick={() => fileRef.current?.click()}>
            Import CSV / Excel
          </button>

          <button
            type="button"
            className="btn btn-small-icon"
            onClick={openManualAdd}
            title="Add expense"
            aria-label="Add expense"
          >
            +
          </button>

          <button type="button" className="btn" onClick={() => jsonRef.current?.click()}>
            Import JSON
          </button>

          <button type="button" className="btn" onClick={handleExport}>
            Export JSON
          </button>

          <button type="button" className="btn" onClick={handlePDF}>
            Download PDF
          </button>

          <button type="button" className="btn btn-quiet" onClick={handleClearAll}>
            Clear
          </button>

          <button
            type="button"
            className="btn btn-theme btn-theme-quiet"
            onClick={() => setDarkMode((v) => !v)}
            aria-label="Toggle theme"
            title="Toggle theme"
          >
            {darkMode ? <SunIcon /> : <MoonIcon />}
          </button>

          {user ? (
            <>
              <span className="muted">Signed in as {user.displayName || user.username}</span>
              <button type="button" className="btn btn-quiet" onClick={signOutUser}>
                Sign out
              </button>
            </>
          ) : (
            <div className="login-inline">
              <input
                className="field-input login-input"
                type="text"
                placeholder="Username"
                value={loginName}
                onChange={(e) => setLoginName(e.target.value)}
              />
              <input
                className="field-input login-input"
                type="password"
                placeholder="Password"
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
              />
              <button
                type="button"
                className="btn btn-quiet"
                onClick={() => signInWithNamePassword(loginName, loginPassword)}
              >
                Log in
              </button>
            </div>
          )}

          <input
            ref={fileRef}
            type="file"
            accept=".csv,.xlsx,.xls,text/csv"
            hidden
            onChange={async (e) => {
              const file = e.target.files?.[0]
              await handleFile(file)
              e.target.value = ''
            }}
          />

          <input
            ref={jsonRef}
            type="file"
            accept=".json,application/json"
            hidden
            onChange={async (e) => {
              const file = e.target.files?.[0]
              await handleJSONfile(file)
              e.target.value = ''
            }}
          />
        </div>
      </header>

<section className="kpis">
  <div className="kpi-card">
    <div className="label-row">
      <div className="label">My Spend</div>
      <select
        className="kpi-year-select"
        value={kpiYear}
        onChange={(e) => setKpiYear(e.target.value)}
      >
        {availableYears.map((year) => (
          <option key={year} value={year}>
            {year}
          </option>
        ))}
      </select>
    </div>
    <div className="value">{fmtEUR(kpiMySpend)}</div>
  </div>

  <div className="kpi-card">
    <div className="label-row">
      <div className="label">Transactions</div>
      <select
        className="kpi-year-select"
        value={kpiYear}
        onChange={(e) => setKpiYear(e.target.value)}
      >
        {availableYears.map((year) => (
          <option key={year} value={year}>
            {year}
          </option>
        ))}
      </select>
    </div>
    <div className="value">{fmtInt(kpiTransactionCount)}</div>
  </div>

  <div className="kpi-card">
    <div className="label-row">
      <div className="label">Categories</div>
      <select
        className="kpi-year-select"
        value={kpiYear}
        onChange={(e) => setKpiYear(e.target.value)}
      >
        {availableYears.map((year) => (
          <option key={year} value={year}>
            {year}
          </option>
        ))}
      </select>
    </div>
    <div className="value">{fmtInt(kpiCategoryCount)}</div>
  </div>

  <div className="kpi-card">
    <div className="label-row">
      <div className="label">Joint Spend (full)</div>
      <select
        className="kpi-year-select"
        value={kpiYear}
        onChange={(e) => setKpiYear(e.target.value)}
      >
        {availableYears.map((year) => (
          <option key={year} value={year}>
            {year}
          </option>
        ))}
      </select>
    </div>
    <div className="value">{fmtEUR(kpiJointSpend)}</div>
  </div>

  <div className="kpi-card accent">
    <div className="label">{splitBalanceLabel}</div>
    <div className="value">{fmtEUR(splitBalanceValue)}</div>
  </div>
</section>

      <nav className="tabs">
        {TABS.map((tab) => (
          <button
            type="button"
            key={tab}
            className={`tab-btn ${activeTab === tab ? 'active' : ''}`}
            onClick={() => setActiveTab(tab)}
          >
            {tab}
            {tab === 'Splits' && splitBalanceValue > 0.004 ? fmtEUR(splitBalanceValue) : null}
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
                      label={({ name, percent }) => `${name} ${fmtNumber(percent * 100, 0)}%`}
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
                <h2>Own Transactions ({sortedTransactions.length})</h2>
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
                        <td className="amount">{fmtInt(filtered.length)}</td>
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

      {activeTab === 'Joint' &&
        (!hasData ? (
          <EmptyState message="Import a joint account statement to see joint costs." />
        ) : jointTransactions.length === 0 ? (
          <EmptyState message="No joint transactions detected yet." />
        ) : (
          <>
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
                      {jointCategoryData.map((d) => (
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
            </div>

            <div className="grid-2">
              <div className="panel">
                <h2>Joint Transactions ({sortedJointTransactions.length})</h2>
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
                      {sortedJointTransactions.map((t) => (
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
                        <td className="amount">{fmtInt(filteredJointTransactions.length)}</td>
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
          </>
        ))}

      {activeTab === 'Trends' &&
        (!hasData ? (
          <EmptyState />
        ) : (
          <div className="panel">
            <h2>Top Merchants by Spend</h2>
            <ResponsiveContainer width="100%" height={merchantChartHeight}>
              <BarChart data={merchantData} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--color-divider)" />
                <XAxis type="number" tick={axisTick} tickFormatter={(v) => fmtEUR(v)} />
                <YAxis type="category" dataKey="name" width={180} tick={axisTick} />
                <Tooltip {...tt} />
                <Bar dataKey="value" fill="#264653" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        ))}

      {activeTab === 'Forecast' &&
        (!hasData ? (
          <EmptyState />
        ) : (
          <>
            <div className="grid-2">
              <div className="panel">
                <h2>Spend Forecast</h2>
                <p className="subtle-note">Average of recent monthly spend: {fmtEUR(forecast.avg)}</p>
                <ResponsiveContainer width="100%" height={chartHeight}>
                  <LineChart data={forecast.series}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--color-divider)" />
                    <XAxis dataKey="month" tick={axisTick} />
                    <YAxis tick={axisTick} tickFormatter={(v) => fmtEUR(v)} width={88} />
                    <Tooltip {...tt} />
                    <Legend />
                    <Line type="monotone" dataKey="actual" stroke="#01696f" strokeWidth={2} dot={{ r: 4 }} name="Actual" connectNulls={false} />
                    <Line type="monotone" dataKey="projected" stroke="#f4a261" strokeWidth={2} strokeDasharray="6 4" dot={{ r: 4 }} name="Projected" connectNulls={false} />
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
                      {forecast.series.map((row) => (
                        <tr key={row.month}>
                          <td>{row.month}</td>
                          <td className="amount">{row.actual != null ? fmtEUR(row.actual) : '—'}</td>
                          <td className="amount">{row.projected != null ? fmtEUR(row.projected) : '—'}</td>
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
                      {forecastByCategory.map((row) => (
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
                    Most expensive category in the recent period is <strong>{mostExpensiveCategory.category}</strong> with an average of{' '}
                    <strong>{fmtEUR(mostExpensiveCategory.avgPerMonth)}</strong> per month.
                  </p>
                )}
              </div>
            )}
          </>
        ))}

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
                  {subscriptions.map((t) => {
                    const isEditing = editingId === t.id
                    return (
                      <tr key={t.id} className={isEditing ? 'editing-row' : ''}>
                        <td>
                          {isEditing ? (
                            <input
                              className="field-input"
                              type="text"
                              value={editDraft.description}
                              placeholder="Subscription"
                              onChange={(e) => setEditDraft((p) => ({ ...p, description: e.target.value }))}
                            />
                          ) : (
                            t.description || '—'
                          )}
                        </td>
                        <td className="amount">{fmtEUR(t.totalAmount || 0)}</td>
                        <td className="amount">
                          {isEditing ? (
                            <input
                              className="field-input amount-input"
                              type="number"
                              min="0"
                              step="0.01"
                              value={editDraft.amount}
                              onChange={(e) => setEditDraft((p) => ({ ...p, amount: e.target.value }))}
                            />
                          ) : (
                            fmtEUR(t.lastAmount ?? t.amount ?? 0)
                          )}
                        </td>
                        <td>
                          {isEditing ? (
                            <input
                              className="field-input"
                              type="date"
                              value={editDraft.date}
                              onChange={(e) => setEditDraft((p) => ({ ...p, date: e.target.value }))}
                            />
                          ) : (
                            t.date
                          )}
                        </td>
                        <td>{t.subscriptionStatus === 'active' ? <span className="muted">Active</span> : <span className="muted">Over currently</span>}</td>
                        <td>
                          <div className="row-actions">
                            {isEditing ? (
                              <>
                                <button type="button" className="btn btn-small btn-primary" onClick={() => saveEdit(t.id)}>
                                  Save
                                </button>
                                <button type="button" className="btn btn-small" onClick={cancelEdit}>
                                  Cancel
                                </button>
                              </>
                            ) : (
                              <>
                                <button type="button" className="btn btn-small" onClick={() => startEdit(t)}>
                                  Edit
                                </button>
                                <button type="button" className="btn btn-small btn-danger" onClick={() => deleteTransaction(t.id)}>
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
                        type="button"
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
                  {sortedTransactions.map((t) => {
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
                              onChange={(e) => setEditDraft((p) => ({ ...p, date: e.target.value }))}
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
                              onChange={(e) => setEditDraft((p) => ({ ...p, description: e.target.value }))}
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
                              onChange={(e) => setEditDraft((p) => ({ ...p, category: e.target.value }))}
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
                              onChange={(e) => setEditDraft((p) => ({ ...p, amount: e.target.value }))}
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
                              onChange={(e) => setEditDraft((p) => ({ ...p, split: e.target.value === 'yes' }))}
                            >
                              <option value="no">No</option>
                              <option value="yes">Yes</option>
                            </select>
                          ) : (
                            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                              <button
                                type="button"
                                className={`btn btn-split ${t.split || t.splitPaid ? 'yes' : ''}`}
                                onClick={() => toggleSplit(t.id)}
                                disabled={t.splitPaid}
                                title={t.splitPaid ? 'Already split and paid in the past' : ''}
                              >
                                {t.split || t.splitPaid ? 'Yes' : 'No'}
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
                              value={t.account === 'joint' || editDraft.joint ? 'yes' : 'no'}
                              onChange={(e) => setEditDraft((p) => ({ ...p, joint: e.target.value === 'yes' }))}
                              disabled={t.account === 'joint'}
                              title={t.account === 'joint' ? 'Already from a Joint statement' : ''}
                            >
                              <option value="no">No</option>
                              <option value="yes">Yes</option>
                            </select>
                          ) : (
                            <button
                              type="button"
                              className={`btn btn-split ${t.joint || t.account === 'joint' ? 'yes' : ''}`}
                              onClick={() => handleJointToggle(t)}
                            >
                              {t.joint || t.account === 'joint' ? 'Yes' : 'No'}
                            </button>
                          )}
                        </td>
                        <td>
                          <div className="row-actions">
                            {isEditing ? (
                              <>
                                <button type="button" className="btn btn-small btn-primary" onClick={() => saveEdit(t.id)}>
                                  Save
                                </button>
                                <button type="button" className="btn btn-small" onClick={cancelEdit}>
                                  Cancel
                                </button>
                              </>
                            ) : (
                              <>
                                <button type="button" className="btn btn-small" onClick={() => startEdit(t)}>
                                  Edit
                                </button>
                                <button type="button" className="btn btn-small btn-danger" onClick={() => deleteTransaction(t.id)}>
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

      {activeTab === 'Splits' &&
        (splitRows.length === 0 ? (
          <EmptyState message="Go to Transactions and set Split to Yes on any shared expense." />
        ) : (
          <>
            <div className="panel">
<div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '1rem', flexWrap: 'wrap' }}>
  <h2>Monthly Split Balance</h2>

  <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap' }}>
    <button
      type="button"
      className="btn btn-small"
      onClick={handleSplitPDF}
    >
      Split PDF
    </button>

    <button
      type="button"
      className="btn btn-small btn-primary"
      onClick={markAllSplitsPaid}
    >
      Paid
    </button>
  </div>
</div>
              <p className="subtle-note">50% of each split transaction is counted as owed to you.</p>
              {partnerUser && (
  <div style={{
    display: 'flex', gap: '1rem', flexWrap: 'wrap', margin: '0.75rem 0',
    padding: '0.75rem 1rem',
    background: 'var(--color-surface-offset)',
    borderRadius: 'var(--radius-md)',
    fontSize: '0.9rem'
  }}>
    <span>
      <span className="muted">Your open splits: </span>
      <strong>{fmtEUR(myOpenSplitTotal)}</strong>
    </span>
    <span>
      <span className="muted">{partnerUser.displayName}'s open splits: </span>
      <strong>{fmtEUR(partnerOpenSplitTotal)}</strong>
    </span>
    <span style={{ color: netSplitBalance > 0.004 ? 'var(--color-success)' : netSplitBalance < -0.004 ? 'var(--color-warning)' : 'var(--color-text-muted)' }}>
      <strong>{splitBalanceLabel}: {fmtEUR(splitBalanceValue)}</strong>
    </span>
  </div>
)}
              <div className="table-wrap">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Month</th>
                      <th>Owed to You</th>
                    </tr>
                  </thead>
                  <tbody>
                    {splitRows.map((r) => (
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
          </>
        ))}

      {showAddModal && (
        <div className="modal-backdrop" onClick={closeManualAdd}>
          <div className="modal modal-form" onClick={(e) => e.stopPropagation()}>
            <h3>Add expense</h3>
            <p>Manually add one expense to your transactions.</p>

            <div className="manual-form">
              <label className="field-group">
                <span>Date</span>
                <input
                  className="field-input"
                  type="date"
                  value={manualDraft.date}
                  onChange={(e) => setManualDraft((p) => ({ ...p, date: e.target.value }))}
                />
              </label>

              <label className="field-group">
                <span>Description</span>
                <input
                  className="field-input"
                  type="text"
                  value={manualDraft.description}
                  placeholder="Expense description"
                  onChange={(e) => setManualDraft((p) => ({ ...p, description: e.target.value }))}
                />
              </label>

              <label className="field-group">
                <span>Category</span>
                <select
                  className="field-input"
                  value={manualDraft.category}
                  onChange={(e) => setManualDraft((p) => ({ ...p, category: e.target.value }))}
                >
                  {categoryOptions.map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </label>

              <label className="field-group">
                <span>Amount</span>
                <input
                  className="field-input amount-input"
                  type="text"
                  inputMode="decimal"
                  value={manualDraft.amount}
                  placeholder="0,00"
                  onChange={(e) => setManualDraft((p) => ({ ...p, amount: e.target.value }))}
                />
              </label>

              <label className="field-check">
                <input
                  type="checkbox"
                  checked={manualDraft.split}
                  onChange={(e) => setManualDraft((p) => ({ ...p, split: e.target.checked }))}
                />
                <span>Split</span>
              </label>

              <label className="field-check">
                <input
                  type="checkbox"
                  checked={manualDraft.joint}
                  onChange={(e) => setManualDraft((p) => ({ ...p, joint: e.target.checked }))}
                />
                <span>Show in Joint tab too</span>
              </label>
            </div>

            <div className="modal-actions">
              <button type="button" className="btn btn-primary" onClick={saveManualExpense}>
                Add expense
              </button>
              <button type="button" className="btn" onClick={closeManualAdd}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {pendingImport && (
        <div className="modal-backdrop">
          <div className="modal">
            <h3>Which account is this statement from?</h3>
            <p>
              Choose <strong>Own</strong> for your personal account, or <strong>Joint</strong> for your shared account with your partner.
            </p>
            <div className="modal-actions">
              <button
                type="button"
                className="btn btn-modal"
                onClick={() => finishImport(pendingImport.rows, pendingImport.label, 'personal')}
              >
                Own
              </button>
              <button
                type="button"
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