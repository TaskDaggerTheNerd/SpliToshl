export const DEFAULTCATEGORIES = [
  'Dining',
  'Groceries',
  'Health',
  'Housing',
  'Investments',
  'Other',
  'Shopping',
  'Subscriptions',
  'Transport',
  'Utilities',
  'Travel',
  'Leisure',
  'Sports',
]

function formatGroupedNumber(value, decimals = 2) {
  const n = Number(value ?? 0)
  const safe = Number.isFinite(n) ? n : 0
  const sign = safe < 0 ? '-' : ''
  const abs = Math.abs(safe)
  const fixed = abs.toFixed(decimals)
  const [intPart, decPart] = fixed.split('.')
  const groupedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ' ')
  return decimals > 0 ? `${sign}${groupedInt},${decPart}` : `${sign}${groupedInt}`
}

export function fmtEUR(v) {
  return `${formatGroupedNumber(v, 2)} €`
}

export function fmtNumber(v, decimals = 2) {
  return formatGroupedNumber(v, decimals)
}

export function fmtInt(v) {
  return formatGroupedNumber(v, 0)
}

export function r2(v) {
  return Math.round(Number(v || 0) * 100) / 100
}

export function inferCategory(description = '') {
  const d = String(description).toLowerCase()
  if (/uber|bolt|taxi|metro|cp |comboios|train|bus|fuel|gas|petrol|galp|bp |repsol/.test(d)) return 'Transport'
  if (/pingo doce|continente|lidl|aldi|minipreco|mercearia|supermercado|grocery|supermarket/.test(d)) return 'Groceries'
  if (/netflix|spotify|prime|apple|hbo|disney|subscription|nts|dazn|icloud|dropbox|youtube|adobe|microsoft/.test(d)) return 'Subscriptions'
  if (/restaurant|café|coffee|starbucks|mcdonald|burger|pizza|takeaway|delivery|glovo|uber eats/.test(d)) return 'Dining'
  if (/renda|rent|mortgage|apartment|house|condominio/.test(d)) return 'Housing'
  if (/gym|sport|fitness|farmacia|pharmacy|doctor|clinica/.test(d)) return 'Health'
  if (/amazon|ikea|shopping|store|zara|h&m|primark|retail|fnac/.test(d)) return 'Shopping'
  if (/invest|etf|fund|degiro|xbpi|trading|stock|crypto/.test(d)) return 'Investments'
  if (/agua|electricity|eletricidade|internet|nos |meo |vodafone|edp|endesa|utility/.test(d)) return 'Utilities'
  return 'Other'
}

export function parseAmount(value) {
  if (value == null || value === '') return 0
  let str = String(value).trim()
  if (!str) return 0

  str = str.replace(/\s/g, '').replace(/[€$£]/g, '')

  const hasComma = str.includes(',')
  const hasDot = str.includes('.')

  if (hasComma && hasDot) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.')
    } else {
      str = str.replace(/,/g, '')
    }
  } else if (hasComma && !hasDot) {
    str = str.replace(/\./g, '').replace(',', '.')
  } else {
    str = str.replace(/,/g, '')
  }

  str = str.replace(/[^0-9.-]/g, '')
  const n = parseFloat(str)
  return Number.isFinite(n) ? Math.abs(n) : 0
}

export function parseDate(value) {
  if (!value) return ''
  const raw = String(value).trim()
  if (!raw) return ''

  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10)

  if (/^\d{4,5}$/.test(raw)) {
    const serial = parseInt(raw, 10)
    if (serial > 40000 && serial < 60000) {
      const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000)
      if (!isNaN(date)) return date.toISOString().slice(0, 10)
    }
  }

  const clean = raw.replace(/\./g, '/').replace(/-/g, '/')
  const parts = clean.split('/').map((p) => p.trim())

  if (parts.length === 3) {
    let [a, b, c] = parts

    if (a.length === 4) {
      return `${a}-${String(b).padStart(2, '0')}-${String(c).padStart(2, '0')}`
    }

    const day = String(a).padStart(2, '0')
    const month = String(b).padStart(2, '0')
    let year = c

    if (year.length === 2) year = Number(year) >= 70 ? `19${year}` : `20${year}`
    if (year.length === 4) return `${year}-${month}-${day}`
  }

  const parsed = new Date(raw)
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10)

  return ''
}

function normaliseKey(k) {
  return String(k).replace(/^\uFEFF/, '').trim().toLowerCase()
}

export function normalizeTransactions(rows = []) {
  return (rows || [])
    .map((row, idx) => {
      const raw = Object.fromEntries(
        Object.entries(row || {}).map(([k, v]) => [normaliseKey(k), v])
      )

      const date = parseDate(
        raw['completed date'] ??
          raw['started date'] ??
          raw.date ??
          raw['data'] ??
          raw['transaction_date'] ??
          raw['data mov.'] ??
          raw['data mov'] ??
          raw['booking date'] ??
          raw['value date'] ??
          ''
      )

      const description = String(
        raw.detail ??
          raw.details ??
          raw.description ??
          raw['descrição'] ??
          raw['descricao'] ??
          raw.memo ??
          raw['merchant name'] ??
          raw.merchant ??
          ''
      ).trim()

      const rawCategory = String(raw.category ?? raw['categoria'] ?? '').trim()
      const category = rawCategory || inferCategory(description) || 'Other'

      const amount = parseAmount(
        raw['expense amount'] ??
          raw.amount ??
          raw['valor'] ??
          raw['montante'] ??
          raw.value ??
          raw.debit ??
          raw.credit ??
          0
      )

      const merchant =
        String(raw.merchant ?? description.split('/')[0] ?? '').trim() ||
        description ||
        'Unknown'

      const split =
        raw.split === true ||
        String(raw.split).toLowerCase() === 'true' ||
        String(raw.split).toLowerCase() === 'yes'

      return {
        ...row,
        id: String(
          raw.id ??
            [
              date || 'no-date',
              description || 'no-desc',
              amount || 0,
              raw['completed date'] ?? raw['started date'] ?? raw['value date'] ?? '',
              idx,
            ].join('|')
        ),
        date,
        description,
        merchant,
        amount,
        category,
        split,
      }
    })
    .filter((t) => t.description || t.amount || t.date)
}

export function dedup(rows = []) {
  const seen = new Set()

  return rows.filter((r, idx) => {
    const key = String(r.id ?? `row-${idx}`)
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

export function buildForecast(rows = []) {
  const byMonth = new Map()

  rows.forEach((r) => {
    const month = String(r.date || '').slice(0, 7)
    if (month.length === 7) {
      byMonth.set(month, (byMonth.get(month) || 0) + Math.abs(Number(r.amount) || 0))
    }
  })

  const sorted = [...byMonth.entries()].sort((a, b) => a[0].localeCompare(b[0]))
  const history = sorted.slice(-6)
  const avg = history.length
    ? history.reduce((sum, [, value]) => sum + value, 0) / history.length
    : 0

  const series = history.map(([month, actual]) => ({
    month,
    actual: r2(actual),
    projected: null,
  }))

  const baseMonth = history.at(-1)?.[0] || new Date().toISOString().slice(0, 7)
  let [year, month] = baseMonth.split('-').map(Number)

  for (let i = 0; i < 3; i++) {
    month += 1
    if (month > 12) {
      year += 1
      month = 1
    }

    series.push({
      month: `${year}-${String(month).padStart(2, '0')}`,
      actual: null,
      projected: r2(avg),
    })
  }

  return { series, avg: r2(avg) }
}

export function buildSplitSummary(rows = []) {
  const byMonth = new Map()

  rows
    .filter((r) => r.split)
    .forEach((r) => {
      const month = String(r.date || '').slice(0, 7)
      if (month.length === 7) {
        byMonth.set(
          month,
          (byMonth.get(month) || 0) + Math.abs(Number(r.amount) || 0) / 2
        )
      }
    })

  return byMonth
}