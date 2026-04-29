import { useState, useEffect, useRef } from 'react'
import { useTheme } from '../context/ThemeContext'
import Chart from 'chart.js/auto'
import './CpiAnalytics.css'

/* ======================================================
   STORE — module-level pub/sub (preserved from original)
   ====================================================== */
const _listeners = new Map()
const _state = { reports: [], mergedData: {}, selectedDomain: null }

function subscribe(key, fn) {
  if (!_listeners.has(key)) _listeners.set(key, new Set())
  _listeners.get(key).add(fn)
  return () => _listeners.get(key).delete(fn)
}

function _notify(key) {
  ;(_listeners.get(key) || []).forEach(fn => fn(_state[key]))
  ;(_listeners.get('*') || []).forEach(fn => fn(_state))
}

function getState(key) {
  return key ? _state[key] : { ..._state }
}

const mutations = {
  addReport(report) {
    if (_state.reports.find(r => r.id === report.id)) return
    _state.reports = [..._state.reports, report]
    _rebuildMerged()
    _notify('reports')
    _notify('mergedData')
  },
  removeReport(id) {
    _state.reports = _state.reports.filter(r => r.id !== id)
    _rebuildMerged()
    _notify('reports')
    _notify('mergedData')
  },
  setSelectedDomain(domain) {
    _state.selectedDomain = domain
    _notify('selectedDomain')
  },
}

function _rebuildMerged() {
  const merged = {}
  _state.reports.forEach(r => {
    Object.entries(r.sheets).forEach(([cpi, rows]) => {
      if (!merged[cpi]) merged[cpi] = []
      rows.forEach(row => merged[cpi].push({ ...row, _reportId: r.id }))
    })
  })
  _state.mergedData = merged
}

function getDomains() {
  const domains = new Set()
  Object.values(_state.mergedData).forEach(rows =>
    rows.forEach(r => {
      const d = r['GTS Product Domain']
      if (d && !String(d).includes('No Mapping')) domains.add(d)
    })
  )
  return [...domains].sort()
}

/* ======================================================
   ANALYTICS — pure functions (preserved from original)
   ====================================================== */
const _safeDiv = (a, b) => (b ? Math.round((a / b) * 100) : 0)
const _safeAvg = (arr, key) => {
  const vals = arr.map(r => parseFloat(r[key])).filter(v => !isNaN(v))
  return vals.length ? vals.reduce((s, v) => s + v, 0) / vals.length : null
}
const _safeSum = (arr, key) => arr.reduce((s, r) => s + (parseFloat(r[key]) || 0), 0)

function filterByDomain(rows, domain) {
  return (rows || []).filter(r => {
    const d = String(r['GTS Product Domain'] || '')
    return d === domain || d.includes(domain)
  })
}

function computeDomainKPIs(mergedData, domain) {
  const f = cpi => filterByDomain(mergedData[cpi], domain)
  const cpi1 = f('CPI1'), cpi3 = f('CPI3'), cpi4 = f('CPI4')
  const cpi5 = f('CPI5'), cpi2b = f('CPI2B'), cpi6 = f('CPI6')
  const committed = _safeSum(cpi1, 'Story Points Committed')
  const delivered = _safeSum(cpi1, 'Story Points Delivered')
  return {
    saydoRatio: _safeDiv(delivered, committed), committed, delivered,
    inflow: cpi3.length,
    inflowOpen: cpi3.filter(r => r.Status !== 'Done').length,
    inflowCritical: cpi3.filter(r => r.Priority === 'Critical').length,
    totalResolved: cpi4.length,
    slaAdherence: _safeDiv(cpi4.filter(r => r['Within SLA'] === 'Yes').length, cpi4.length),
    avgMttr: _safeAvg(cpi4, 'MTTR (in days)'),
    agingTotal: cpi5.length,
    agingCritical: cpi5.filter(r => r.Priority === 'Critical').length,
    avgTicketAge: _safeAvg(cpi5, 'Ticket Age (in days)'),
    totalOutageMins: _safeSum(cpi2b, 'Total Outage Minutes'),
    outageCount: cpi2b.length,
    avgCoverage: _safeAvg(cpi6, 'Coverage'),
    coverageSamples: cpi6.length,
  }
}

function computeOverallKPIs(mergedData, domains) {
  const all4 = mergedData.CPI4 || [], all1 = mergedData.CPI1 || [], all3 = mergedData.CPI3 || []
  return {
    domainCount: domains.length,
    overallSla: _safeDiv(all4.filter(r => r['Within SLA'] === 'Yes').length, all4.length),
    overallSaydo: _safeDiv(_safeSum(all1, 'Story Points Delivered'), _safeSum(all1, 'Story Points Committed')),
    totalInflow: all3.length,
    totalResolved: all4.length,
  }
}

function fmtNum(v, decimals = 0) {
  if (v === null || v === undefined) return 'N/A'
  if (v === 0) return '0'
  return decimals > 0 ? v.toFixed(decimals) : Math.round(v).toLocaleString()
}

function rateStatus(val, goodThreshold = 80, warnThreshold = 60) {
  if (val === null || val === undefined) return 'gray'
  if (val >= goodThreshold) return 'good'
  if (val >= warnThreshold) return 'warn'
  return 'bad'
}

/* ======================================================
   INFERENCE — text generators (preserved from original)
   ====================================================== */
function generateInferences(kpis) {
  const blocks = []
  if (kpis.committed > 0) {
    const level = kpis.saydoRatio >= 80 ? 'on track' : kpis.saydoRatio >= 60 ? 'showing moderate risk' : 'below target'
    const detail = kpis.saydoRatio < 80
      ? 'Story point leakage indicates scope creep, mid-sprint scope additions, or capacity constraints within sprints.'
      : 'Team is consistently meeting sprint commitments.'
    blocks.push({ cpi: 'CPI1 — Saydo Ratio', text: `Sprint delivery is ${level} at ${kpis.saydoRatio}% Saydo ratio (${kpis.delivered} of ${kpis.committed} story points delivered). ${detail}` })
  }
  if (kpis.inflow > 0) {
    const critNote = kpis.inflowCritical > 0
      ? `${kpis.inflowCritical} critical priority ticket${kpis.inflowCritical > 1 ? 's' : ''} indicate high-severity defects requiring immediate resolution.`
      : 'No critical tickets — defect severity is manageable.'
    blocks.push({ cpi: 'CPI3 — Support Ticket Inflow', text: `${kpis.inflow} tickets were raised during the period, with ${kpis.inflowOpen} remaining open. ${critNote} Sustained open-ticket growth signals insufficient resolution capacity.` })
  }
  if (kpis.totalResolved > 0) {
    const slaNote = kpis.slaAdherence < 80
      ? 'Significant SLA breaches detected — tickets are not resolved within agreed timeframes, risking contractual penalties.'
      : 'SLA performance is within acceptable bounds.'
    const mttrStr = kpis.avgMttr !== null ? ` Average MTTR stands at ${fmtNum(kpis.avgMttr, 1)} days.` : ''
    blocks.push({ cpi: 'CPI4 — SLA Adherence', text: `SLA adherence stands at ${kpis.slaAdherence}%.${mttrStr} ${slaNote} High MTTR values may indicate resolution bottlenecks or resource under-allocation.` })
  }
  if (kpis.agingTotal > 0) {
    const age = kpis.avgTicketAge !== null ? Math.round(kpis.avgTicketAge) : null
    const ageNote = age !== null && age > 30
      ? 'Average ticket age exceeding 30 days points to a systemic triage and prioritisation issue.'
      : age !== null ? 'Ticket age is within acceptable limits.' : ''
    const critNote = kpis.agingCritical > 0
      ? `${kpis.agingCritical} critical ticket${kpis.agingCritical > 1 ? 's are' : ' is'} aging — high risk to production stability and customer satisfaction.`
      : 'No critical aging tickets.'
    blocks.push({ cpi: 'CPI5 — Ticket Aging', text: `${kpis.agingTotal} tickets are aging${age !== null ? ` (avg ${age} days)` : ''}. ${critNote} ${ageNote}` })
  }
  if (kpis.outageCount > 0) {
    const hrs = (kpis.totalOutageMins / 60).toFixed(1)
    const severity = kpis.totalOutageMins > 10000
      ? 'Prolonged outages indicate reliability concerns requiring root cause analysis and preventive controls.'
      : 'Outage impact is limited but warrants continued monitoring.'
    blocks.push({ cpi: 'CPI2B — Outages', text: `${kpis.outageCount} outage event${kpis.outageCount > 1 ? 's' : ''} totalling ${Math.round(kpis.totalOutageMins).toLocaleString()} minutes (${hrs} hours). ${severity}` })
  }
  if (kpis.coverageSamples > 0 && kpis.avgCoverage !== null) {
    const coverageNote = kpis.avgCoverage < 80
      ? 'Coverage below the 80% quality gate increases regression risk for future releases.'
      : 'Code coverage meets the quality gate threshold.'
    blocks.push({ cpi: 'CPI6 — Code Coverage', text: `Average code coverage is ${fmtNum(kpis.avgCoverage, 1)}% across ${kpis.coverageSamples} merge request${kpis.coverageSamples > 1 ? 's' : ''}. ${coverageNote}` })
  }
  return blocks
}

function generateActions(kpis, domain) {
  const actions = []
  if (kpis.committed > 0 && kpis.saydoRatio < 80) {
    actions.push({ cpi: 'CPI1 — Sprint Delivery', text: 'Conduct sprint retrospectives to identify story-point leakage root causes. Introduce spill-over tracking and enforce WIP limits. Align sprint planning to actual team velocity. Target: Saydo ratio ≥80% within two sprints.' })
  }
  if (kpis.inflowCritical > 0 || kpis.inflowOpen > 5) {
    const escalate = kpis.inflowCritical > 0 ? `Immediately escalate ${kpis.inflowCritical} critical ticket${kpis.inflowCritical > 1 ? 's' : ''} for same-day triage. ` : ''
    actions.push({ cpi: 'CPI3 — Ticket Inflow', text: `${escalate}Implement weekly ticket review cadence with the client. Establish a dedicated bug-fix sprint or buffer capacity to reduce backlog. Analyse ticket patterns for systemic defect sources (recurring modules, test gaps).` })
  }
  if (kpis.totalResolved > 0 && kpis.slaAdherence < 80) {
    actions.push({ cpi: 'CPI4 — SLA Adherence', text: 'Triage all SLA-breached tickets immediately and establish an escalation matrix for tickets approaching breach thresholds. Streamline assignment workflows and expand resolver pool where needed. Target: SLA adherence ≥85% by end of next sprint cycle.' })
  }
  if (kpis.agingTotal > 0 && (kpis.agingCritical > 0 || (kpis.avgTicketAge || 0) > 30)) {
    const crit = kpis.agingCritical > 0 ? `Prioritise resolution of ${kpis.agingCritical} aging critical ticket${kpis.agingCritical > 1 ? 's' : ''} in the current sprint. ` : ''
    actions.push({ cpi: 'CPI5 — Ticket Aging', text: `${crit}Implement dashboard alerts at 14-day and 30-day aging thresholds. Schedule fortnightly aging-ticket reviews with team leads. Reject or close stale low-priority tickets after client approval.` })
  }
  if (kpis.outageCount > 0) {
    actions.push({ cpi: 'CPI2B — Outages', text: 'Conduct Post-Incident Reviews (PIRs) for all outages exceeding 60 minutes. Track recurring patterns and implement targeted preventive fixes. Improve monitoring and alerting to reduce time-to-detect. Present outage trend to client at the next governance review.' })
  }
  if (kpis.coverageSamples > 0 && kpis.avgCoverage !== null && kpis.avgCoverage < 80) {
    actions.push({ cpi: 'CPI6 — Code Coverage', text: 'Mandate minimum 80% code coverage as a merge-request quality gate. Allocate test-automation tasks in sprint planning. Identify untested modules and prioritise unit/integration test creation. Track coverage trend monthly.' })
  }
  if (actions.length === 0) {
    actions.push({ cpi: 'General', text: `${domain} is performing well across tracked CPIs. Maintain current practices and monitoring cadence. Prepare a month-on-month trend comparison to confirm sustained performance in the next reporting period.` })
  }
  return actions
}

function computeTeamKPIs(md, domain) {
  const df = rows => (rows || []).filter(r => { const d = String(r['GTS Product Domain'] || ''); return d === domain || d.includes(domain) })
  const c1 = df(md.CPI1), c3 = df(md.CPI3), c4 = df(md.CPI4), c5 = df(md.CPI5)
  const teamSet = new Set()
  ;[c1, c3, c4, c5].forEach(rows => rows.forEach(r => { if (r['Agile Team']) teamSet.add(String(r['Agile Team'])) }))
  return [...teamSet].sort().map(team => {
    const t1 = c1.filter(r => String(r['Agile Team']) === team)
    const t3 = c3.filter(r => String(r['Agile Team']) === team)
    const t4 = c4.filter(r => String(r['Agile Team']) === team)
    const t5 = c5.filter(r => String(r['Agile Team']) === team)
    const com = t1.reduce((s, r) => s + (parseFloat(r['Story Points Committed']) || 0), 0)
    const del = t1.reduce((s, r) => s + (parseFloat(r['Story Points Delivered']) || 0), 0)
    const saydo = com > 0 ? Math.round(del / com * 100) : null
    const slaTotal = t4.length, slaYes = t4.filter(r => r['Within SLA'] === 'Yes').length
    const sla = slaTotal > 0 ? Math.round(slaYes / slaTotal * 100) : null
    const inflow = t3.length, inflowOpen = t3.filter(r => r['Status'] !== 'Done').length
    const critical = t3.filter(r => r['Priority'] === 'Critical').length
    const agingTotal = t5.length, agingCritical = t5.filter(r => r['Priority'] === 'Critical').length
    const ageVals = t5.map(r => parseFloat(r['Ticket Age (in days)'])).filter(v => !isNaN(v))
    const avgAge = ageVals.length ? Math.round(ageVals.reduce((a, b) => a + b, 0) / ageVals.length) : null
    const pts = []
    if (saydo !== null) pts.push(saydo >= 80 ? `Saydo ${saydo}% — on track` : saydo >= 60 ? `Saydo ${saydo}% — moderate risk` : `Saydo ${saydo}% — below target`)
    if (sla !== null) pts.push(sla >= 80 ? `SLA ${sla}% — compliant` : sla >= 60 ? `SLA ${sla}% — needs attention` : `SLA ${sla}% — action required`)
    if (critical > 0) pts.push(`${critical} critical ticket${critical > 1 ? 's' : ''} open`)
    if (agingCritical > 0) pts.push(`${agingCritical} critical aging ticket${agingCritical > 1 ? 's' : ''}`)
    if (avgAge !== null && avgAge > 30) pts.push(`avg ticket age ${avgAge}d (high)`)
    if (inflowOpen > 5) pts.push(`${inflowOpen} open tickets`)
    const note = pts.length ? pts.join(' · ') : (inflow === 0 && agingTotal === 0 ? 'No ticket activity recorded for this team.' : 'Performance within acceptable bounds.')
    return { team, saydo, com, del, sla, slaTotal, inflow, inflowOpen, critical, agingTotal, agingCritical, avgAge, note }
  })
}

/* ======================================================
   CHART UTILITIES (preserved from original)
   ====================================================== */
const C = { blue: '#0033A0', accent: '#00A3E0', green: '#00843D', amber: '#F5A623', red: '#C8102E', purple: '#7B2D8B', mid: '#3A5BB8', teal: '#0F6856' }
const PALETTE = Object.values(C)

// CPI_DEFS references chart builder functions — hoisted as function declarations below
const CPI_DEFS = {
  CPI1:  { label: 'CPI1 — Saydo Ratio',          buildChart: trdCpi1Chart  },
  CPI3:  { label: 'CPI3 — Ticket Inflow',         buildChart: trdCpi3Chart  },
  CPI4:  { label: 'CPI4 — SLA Adherence',         buildChart: trdCpi4Chart  },
  CPI5:  { label: 'CPI5 — Ticket Aging',          buildChart: trdCpi5Chart  },
  CPI2B: { label: 'CPI2B — Outages',              buildChart: trdCpi2bChart },
  CPI6:  { label: 'CPI6 — Code Coverage',         buildChart: trdCpi6Chart  },
}

const CPI_PIVOT_FIELDS = {
  CPI1:  { cat: ['GTS Product Domain','Month','Agile Team','Initially Committed to Sprint','Completed in Sprint'], num: ['Story Points Committed','Story Points Delivered'] },
  CPI3:  { cat: ['GTS Product Domain','Month','Issue Type','Priority','Status'], num: [] },
  CPI4:  { cat: ['GTS Product Domain','Month','Issue Type','Priority','Within SLA'], num: ['MTTR (in days)'] },
  CPI5:  { cat: ['GTS Product Domain','Month','Issue Type','Priority','Status'], num: ['Ticket Age (in days)'] },
  CPI2B: { cat: ['GTS Product Domain','Month','Environment'], num: ['Total Outage Minutes'] },
  CPI6:  { cat: ['GTS Product Domain','Month','Agile Team'], num: ['Coverage'] },
}

function trdDs(label, data, color, type) {
  const isLine = type === 'line'
  return {
    label, data,
    backgroundColor: color + (isLine ? '33' : 'BB'),
    borderColor: color, borderWidth: isLine ? 2 : 0,
    fill: false, tension: 0.3,
    pointRadius: isLine ? 4 : 0, pointHoverRadius: 6,
    ...(isLine ? {} : { borderRadius: 4 }),
  }
}

function trdCpi1Chart(rows, chartType) {
  const months = trdSortedMonths(rows)
  if (!months.length) return null
  const type = chartType === 'line' ? 'line' : 'bar'
  const committed = months.map(m => rows.filter(r => r.Month === m).reduce((s, r) => s + (parseFloat(r['Story Points Committed']) || 0), 0))
  const delivered  = months.map(m => rows.filter(r => r.Month === m).reduce((s, r) => s + (parseFloat(r['Story Points Delivered'])  || 0), 0))
  return { type, yLabel: 'Story Points', legendItems: [{ label: 'Story Points Committed', color: C.blue }, { label: 'Story Points Delivered', color: C.green }], data: { labels: months, datasets: [trdDs('Story Points Committed', committed, C.blue, type), trdDs('Story Points Delivered', delivered, C.green, type)] } }
}

function trdCpi3Chart(rows, chartType) {
  const months = trdSortedMonths(rows)
  if (!months.length) return null
  const priorities = [...new Set(rows.map(r => r.Priority).filter(Boolean))].sort()
  const pColor = { Critical: C.red, High: C.amber, Medium: C.blue, Low: C.teal }
  const type = chartType === 'line' ? 'line' : 'bar'
  return {
    type, yLabel: 'Ticket Count',
    legendItems: priorities.map((p, i) => ({ label: p, color: pColor[p] || PALETTE[i % PALETTE.length] })),
    data: { labels: months, datasets: priorities.map((p, i) => {
      const color = pColor[p] || PALETTE[i % PALETTE.length]
      return trdDs(p, months.map(m => rows.filter(r => r.Month === m && r.Priority === p).length), color, type)
    })},
    extraOptions: type === 'bar' ? { scales: { x: { stacked: true }, y: { stacked: true } } } : {},
  }
}

function trdCpi4Chart(rows, chartType) {
  const months = trdSortedMonths(rows)
  if (!months.length) return null
  const type = chartType === 'line' ? 'line' : 'bar'
  return {
    type, yLabel: 'Tickets',
    legendItems: [{ label: 'Within SLA', color: C.green }, { label: 'SLA Breached', color: C.red }],
    data: { labels: months, datasets: [
      trdDs('Within SLA',   months.map(m => rows.filter(r => r.Month === m && r['Within SLA'] === 'Yes').length), C.green, type),
      trdDs('SLA Breached', months.map(m => rows.filter(r => r.Month === m && r['Within SLA'] === 'No').length),  C.red,   type),
    ]},
  }
}

function trdCpi5Chart(rows, chartType) {
  const months = trdSortedMonths(rows)
  if (!months.length) return null
  const type = chartType === 'line' ? 'line' : 'bar'
  const avgAge = months.map(m => {
    const vals = rows.filter(r => r.Month === m).map(r => parseFloat(r['Ticket Age (in days)'])).filter(v => !isNaN(v))
    return vals.length ? parseFloat((vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1)) : 0
  })
  const critCount = months.map(m => rows.filter(r => r.Month === m && r.Priority === 'Critical').length)
  return {
    type, yLabel: 'Avg Age (days)',
    legendItems: [{ label: 'Avg Ticket Age (days)', color: C.amber }, { label: 'Critical Count', color: C.red }],
    data: { labels: months, datasets: [
      { ...trdDs('Avg Ticket Age (days)', avgAge,    C.amber, type), yAxisID: 'y' },
      { ...trdDs('Critical Count',        critCount, C.red, 'line'), yAxisID: 'y1', type: 'line' },
    ]},
    extraScales: { y1: { position: 'right', beginAtZero: true, ticks: { font: { size: 11 } }, grid: { drawOnChartArea: false }, title: { display: true, text: 'Critical Count', font: { size: 11 }, color: 'rgba(200,16,46,0.8)' } } },
  }
}

function trdCpi2bChart(rows, chartType) {
  const months = trdSortedMonths(rows)
  if (!months.length) return null
  const type = chartType === 'line' ? 'line' : 'bar'
  const totMins  = months.map(m => parseFloat(rows.filter(r => r.Month === m).reduce((s, r) => s + (parseFloat(r['Total Outage Minutes']) || 0), 0).toFixed(1)))
  const incCount = months.map(m => rows.filter(r => r.Month === m).length)
  return {
    type, yLabel: 'Outage Minutes',
    legendItems: [{ label: 'Total Outage Minutes', color: C.red }, { label: 'Incident Count', color: C.amber }],
    data: { labels: months, datasets: [
      { ...trdDs('Total Outage Minutes', totMins,  C.red,   type),   yAxisID: 'y'  },
      { ...trdDs('Incident Count',       incCount, C.amber, 'line'), yAxisID: 'y1', type: 'line' },
    ]},
    extraScales: { y1: { position: 'right', beginAtZero: true, ticks: { font: { size: 11 } }, grid: { drawOnChartArea: false }, title: { display: true, text: 'Incidents', font: { size: 11 }, color: 'rgba(245,166,35,0.9)' } } },
  }
}

function trdCpi6Chart(rows, chartType) {
  const months = trdSortedMonths(rows)
  if (!months.length) return null
  const type = chartType === 'line' ? 'line' : 'bar'
  const avgCov = months.map(m => {
    const vals = rows.filter(r => r.Month === m).map(r => parseFloat(r.Coverage)).filter(v => !isNaN(v))
    return vals.length ? parseFloat((vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1)) : null
  })
  return {
    type, yLabel: 'Coverage (%)',
    legendItems: [{ label: 'Avg Code Coverage (%)', color: C.purple }],
    data: { labels: months, datasets: [{ ...trdDs('Avg Code Coverage (%)', avgCov, C.purple, type), fill: type === 'line' }] },
    extraScales: { y: { min: 0, max: 100 } },
  }
}

const TRD_MONTH_ORDER = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']

function trdSortedMonths(rows) {
  const months = [...new Set(rows.map(r => r.Month).filter(Boolean))]
  return months.sort((a, b) => {
    const pa = trdParseMY(a), pb = trdParseMY(b)
    if (pa.yr !== pb.yr) return pa.yr - pb.yr
    return TRD_MONTH_ORDER.indexOf(pa.mo) - TRD_MONTH_ORDER.indexOf(pb.mo)
  })
}
function trdParseMY(s) { const p = String(s).toUpperCase().split('-'); return { mo: p[0], yr: parseInt(p[1]) || 0 } }
function trdSlugify(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') }

/* ======================================================
   REACT COMPONENT
   ====================================================== */
export default function CpiAnalytics() {
  const { theme } = useTheme()

  // Core state
  const [loading, setLoading] = useState(true)
  const [error, setError]     = useState(null)
  const [activeTab, setActiveTab] = useState('overview')
  const [lastUpdated, setLastUpdated] = useState('')
  const [renderTick, setRenderTick] = useState(0) // incremented on any store change

  // Inferences tab state
  const [generating, setGenerating]   = useState(false)
  const [inferOutput, setInferOutput] = useState(null) // { domain, kpis, inferences, actions, teamRows }

  // Trend tab controls
  const [activeCpi, setActiveCpi]         = useState('CPI1')
  const [chartType, setChartType]         = useState('bar')
  const [domainFilter, setDomainFilter]   = useState('ALL')
  const [pivotState, setPivotState]       = useState({ rows: ['Month'], cols: ['GTS Product Domain'], vals: [] })
  const [pivotCpi, setPivotCpi]           = useState('CPI1')
  const [pivotChartType, setPivotChartType] = useState('bar')
  const [pivotAggFn, setPivotAggFn]       = useState('count')
  const [pivotTableData, setPivotTableData] = useState(null) // { rowKeys, colKeys, matrix, rowFlds }
  const [pivotLegendItems, setPivotLegendItems] = useState([])

  // Chart instance refs
  const trendCharts   = useRef({})
  const pivotChartRef = useRef(null)

  // ── Subscribe to store ────────────────────────────────────────────
  useEffect(() => subscribe('*', () => setRenderTick(n => n + 1)), [])

  // Clear inference output when domain selection changes
  useEffect(() => subscribe('selectedDomain', () => setInferOutput(null)), [])

  // ── Load data on mount ────────────────────────────────────────────
  useEffect(() => {
    if (_state.reports.length > 0) { setLoading(false); return }
    async function load() {
      try {
        const base = import.meta.env.VITE_STORAGE_URL ?? ''
        const cfgRes = await fetch(`${base}/cpi-analytics/config.json`)
        if (!cfgRes.ok) throw new Error(`Config fetch failed (HTTP ${cfgRes.status})`)
        const config = await cfgRes.json()
        setLastUpdated(config.lastUpdated || '')
        for (const fn of (config.reports || [])) {
          const res = await fetch(`${base}/cpi-analytics/${fn}`)
          if (!res.ok) throw new Error(`Failed to load ${fn} (HTTP ${res.status})`)
          mutations.addReport(await res.json())
        }
      } catch (err) {
        setError(err.message)
      } finally {
        setLoading(false)
      }
    }
    load()
  }, [])

  // ── Destroy charts on unmount ─────────────────────────────────────
  useEffect(() => () => destroyAll(), [])

  function destroyAll() {
    Object.values(trendCharts.current).forEach(c => { try { c.destroy() } catch (_) {} })
    trendCharts.current = {}
    if (pivotChartRef.current) { try { pivotChartRef.current.destroy() } catch (_) {} ; pivotChartRef.current = null }
  }

  // ── Rebuild trend charts whenever relevant deps change ────────────
  useEffect(() => {
    if (activeTab !== 'trend' || loading) {
      if (activeTab !== 'trend') destroyAll()
      return
    }
    const t = setTimeout(buildTrendCharts, 60)
    return () => clearTimeout(t)
  }, [activeTab, activeCpi, chartType, domainFilter, theme, renderTick, loading]) // eslint-disable-line

  // ── Rebuild pivot chart whenever pivot config changes ─────────────
  useEffect(() => {
    if (activeTab !== 'trend' || loading) return
    const t = setTimeout(buildPivotChart, 60)
    return () => clearTimeout(t)
  }, [pivotState, pivotCpi, pivotChartType, pivotAggFn, theme, activeTab, loading, renderTick]) // eslint-disable-line

  // ── Chart theme helper ────────────────────────────────────────────
  function getChartTheme() {
    const dark = theme === 'dark'
    return {
      grid: dark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.06)',
      tick: dark ? '#6E7681' : '#8892A4',
      tooltip: {
        backgroundColor: dark ? '#161B22' : '#FFFFFF',
        borderColor:      dark ? '#30363D' : '#D0D7E2',
        borderWidth: 1,
        titleColor: dark ? '#E6EDF3' : '#1A1F2E',
        bodyColor:  dark ? '#8B949E' : '#5A6478',
      },
    }
  }

  // ── Build per-domain trend charts ─────────────────────────────────
  function buildTrendCharts() {
    destroyAll()
    const mergedData = getState('mergedData')
    const cpiDef = CPI_DEFS[activeCpi]
    if (!cpiDef) return
    let domains = getDomains()
    if (domainFilter !== 'ALL') domains = domains.filter(d => d === domainFilter)
    const ct = getChartTheme()

    domains.forEach(domain => {
      const cid    = `cpi-chart-${activeCpi}-${trdSlugify(domain)}`
      const canvas = document.getElementById(cid)
      if (!canvas) return
      const rows = filterByDomain(mergedData[activeCpi] || [], domain)
        .filter(r => !String(r['GTS Product Domain'] || '').includes('No Mapping'))
      const cfg = cpiDef.buildChart(rows, chartType)
      if (!cfg) return

      const scaleDefaults = {
        x: { ticks: { maxRotation: 30, autoSkip: false, font: { size: 11 }, color: ct.tick }, grid: { color: ct.grid }, title: { display: true, text: 'Month', font: { size: 11 }, color: ct.tick } },
        y: { beginAtZero: true, ticks: { font: { size: 11 }, color: ct.tick }, grid: { color: ct.grid }, title: { display: !!(cfg.yLabel), text: cfg.yLabel || '', font: { size: 11 }, color: ct.tick } },
        ...(cfg.extraScales || {}),
      }

      trendCharts.current[cid] = new Chart(canvas, {
        type: cfg.type,
        data: cfg.data,
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false, ...ct.tooltip } },
          scales: scaleDefaults,
          ...(cfg.extraOptions || {}),
        },
      })
    })
    buildPivotChart()
  }

  // ── Build pivot chart ─────────────────────────────────────────────
  function buildPivotChart() {
    if (pivotChartRef.current) { try { pivotChartRef.current.destroy() } catch (_) {} ; pivotChartRef.current = null }

    const isStacked = pivotChartType === 'bar-stacked'
    const ct_type   = pivotChartType === 'line' ? 'line' : 'bar'
    let rows = (getState('mergedData')[pivotCpi] || [])
      .filter(r => r['GTS Product Domain'] && !String(r['GTS Product Domain'] || '').includes('No Mapping'))
    if (domainFilter !== 'ALL') rows = rows.filter(r => r['GTS Product Domain'] === domainFilter)

    const rowFlds = pivotState.rows.length ? pivotState.rows : ['Month']
    const colFlds = pivotState.cols
    const valFld  = pivotState.vals[0] || null
    const getKey  = (r, flds) => flds.map(f => r[f] ?? '—').join(' | ')
    const rowKeys = [...new Set(rows.map(r => getKey(r, rowFlds)))].sort()
    const colKeys = colFlds.length ? [...new Set(rows.map(r => getKey(r, colFlds)))].sort() : ['Count']

    const buckets = {}
    rows.forEach(r => {
      const rk = getKey(r, rowFlds), ck = colFlds.length ? getKey(r, colFlds) : 'Count'
      const k  = `${rk}||${ck}`
      if (!buckets[k]) buckets[k] = []
      buckets[k].push(r)
    })
    const agg = items => {
      if (valFld && pivotAggFn !== 'count') {
        const vs = items.map(r => parseFloat(r[valFld])).filter(v => !isNaN(v))
        return vs.length ? (pivotAggFn === 'sum' ? vs.reduce((a, b) => a + b, 0) : vs.reduce((a, b) => a + b, 0) / vs.length) : 0
      }
      return items.length
    }
    const matrix = {}
    rowKeys.forEach(rk => { matrix[rk] = {}; colKeys.forEach(ck => { matrix[rk][ck] = buckets[`${rk}||${ck}`] ? agg(buckets[`${rk}||${ck}`]) : 0 }) })

    // Update legend and table via state
    setPivotLegendItems(colKeys.map((ck, i) => ({ label: ck, color: PALETTE[i % PALETTE.length] })))
    setPivotTableData(rowKeys.length ? { rowKeys, colKeys, matrix, rowFlds } : null)

    const canvas = document.getElementById('cpi-pivot-chart')
    if (!canvas || !rowKeys.length) return
    const ct = getChartTheme()
    pivotChartRef.current = new Chart(canvas, {
      type: ct_type,
      data: {
        labels: rowKeys,
        datasets: colKeys.map((ck, i) => ({
          label: ck,
          data: rowKeys.map(rk => parseFloat((matrix[rk][ck] || 0).toFixed(2))),
          backgroundColor: PALETTE[i % PALETTE.length] + (ct_type === 'bar' ? 'CC' : '33'),
          borderColor: PALETTE[i % PALETTE.length],
          borderWidth: ct_type === 'line' ? 2 : 0,
          fill: false, tension: 0.3,
          pointRadius: ct_type === 'line' ? 4 : 0, pointHoverRadius: 6,
        })),
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false, ...ct.tooltip } },
        scales: {
          x: { stacked: isStacked, ticks: { maxRotation: 45, autoSkip: false, font: { size: 11 }, color: ct.tick }, grid: { color: ct.grid } },
          y: { stacked: isStacked, beginAtZero: true, ticks: { font: { size: 11 }, color: ct.tick }, grid: { color: ct.grid } },
        },
      },
    })
  }

  // ── Pivot field management ────────────────────────────────────────
  function pivotAddField(field) {
    const cfg = CPI_PIVOT_FIELDS[pivotCpi] || { cat: [], num: [] }
    setPivotState(prev => {
      const next = { ...prev, rows: [...prev.rows], cols: [...prev.cols], vals: [...prev.vals] }
      if (cfg.num.includes(field)) { if (!next.vals.includes(field)) next.vals.push(field) }
      else if (!next.rows.includes(field) && !next.cols.includes(field)) {
        if (!next.rows.length) next.rows.push(field)
        else next.cols.push(field)
      }
      return next
    })
  }
  function pivotRemoveField(field, zone) {
    setPivotState(prev => ({ ...prev, [zone]: prev[zone].filter(f => f !== field) }))
  }

  // ── Generate inferences ────────────────────────────────────────────
  async function handleGenerate() {
    if (generating) return
    const domain = getState('selectedDomain')
    if (!domain) return
    setGenerating(true)
    await new Promise(r => setTimeout(r, 50))
    const md       = getState('mergedData')
    const kpis     = computeDomainKPIs(md, domain)
    const inferences = generateInferences(kpis)
    const actions    = generateActions(kpis, domain)
    const teamRows   = computeTeamKPIs(md, domain)
    setInferOutput({ domain, kpis, inferences, actions, teamRows })
    setGenerating(false)
  }

  /* ──────────────────────────────────────────────────────────────────
     RENDER HELPERS
  ────────────────────────────────────────────────────────────────── */
  function chipCls(status) {
    const map = { good: 'green', warn: 'amber', bad: 'red', green: 'green', amber: 'amber', red: 'red', blue: 'blue' }
    return `cpi-chip cpi-chip--${map[status] || 'gray'}`
  }

  function MetricCard({ label, value, sub, status }) {
    const subCls = status ? `cpi-metric-card__sub--${status === 'good' ? 'good' : status === 'warn' ? 'warn' : status === 'bad' ? 'bad' : ''}` : ''
    return (
      <div className="cpi-metric-card">
        <div className="cpi-metric-card__label">{label}</div>
        <div className="cpi-metric-card__value">{value}</div>
        {sub && <div className={`cpi-metric-card__sub ${subCls}`}>{sub}</div>}
      </div>
    )
  }

  /* ──────────────────────────────────────────────────────────────────
     OVERVIEW TAB
  ────────────────────────────────────────────────────────────────── */
  function renderOverview() {
    const reports    = getState('reports')
    const mergedData = getState('mergedData')
    const domains    = getDomains()

    if (!reports.length) {
      return (
        <div className="cpi-empty-state">
          <svg className="cpi-empty-state__icon" width="48" height="48" viewBox="0 0 48 48" fill="none">
            <rect x="4" y="10" width="40" height="28" rx="4" stroke="currentColor" strokeWidth="1.5"/>
            <path d="M4 20h40M16 10v8M32 10v8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
          </svg>
          <div className="cpi-empty-state__title">No data loaded</div>
          <div className="cpi-empty-state__sub">Add CPI report files via the Upload Portal to see the overview.</div>
        </div>
      )
    }

    const o = computeOverallKPIs(mergedData, domains)
    const slaNote   = v => v >= 80 ? 'On target' : v >= 60 ? 'Needs attention' : 'Action required'
    const saydoNote = v => v >= 80 ? 'Meeting commitments' : v >= 60 ? 'Moderate risk' : 'Below target'

    return (
      <>
        <div className="cpi-metric-grid">
          <MetricCard label="GTS Domains"    value={o.domainCount} />
          <MetricCard label="Months Loaded"  value={reports.length} />
          <MetricCard label="Overall SLA"    value={`${o.overallSla}%`}    sub={slaNote(o.overallSla)}   status={rateStatus(o.overallSla)} />
          <MetricCard label="Overall Saydo"  value={`${o.overallSaydo}%`}  sub={saydoNote(o.overallSaydo)} status={rateStatus(o.overallSaydo)} />
          <MetricCard label="Ticket Inflow"  value={o.totalInflow.toLocaleString()} sub="across all domains" />
          <MetricCard label="SLA Tickets"    value={o.totalResolved.toLocaleString()} sub="resolved" />
        </div>

        <div className="cpi-card">
          <div className="cpi-card__title">Domain Performance Summary</div>
          <div className="cpi-table-wrap">
            {domains.length ? (
              <table className="cpi-data-table">
                <thead><tr>
                  <th>GTS Product Domain</th>
                  <th className="col-num">Inflow</th>
                  <th>Saydo Ratio</th>
                  <th>SLA Adherence</th>
                  <th className="col-num">Aging</th>
                  <th>Avg MTTR</th>
                  <th>Status</th>
                </tr></thead>
                <tbody>
                  {domains.map(domain => {
                    const kpis   = computeDomainKPIs(mergedData, domain)
                    const slaSt  = rateStatus(kpis.slaAdherence)
                    const saySt  = rateStatus(kpis.saydoRatio)
                    const ov     = slaSt === 'good' && saySt === 'good' ? 'good' : slaSt === 'bad' || saySt === 'bad' ? 'bad' : 'warn'
                    const ovLbl  = ov === 'good' ? 'Good' : ov === 'warn' ? 'Watch' : 'Action'
                    return (
                      <tr key={domain}>
                        <td>{domain}</td>
                        <td className="col-num">{kpis.inflow}</td>
                        <td><span className={chipCls(saySt)}>{kpis.saydoRatio}%</span></td>
                        <td>{kpis.totalResolved ? <span className={chipCls(slaSt)}>{kpis.slaAdherence}%</span> : <span className="cpi-chip cpi-chip--gray">N/A</span>}</td>
                        <td className="col-num">{kpis.agingTotal}</td>
                        <td>{kpis.avgMttr !== null ? fmtNum(kpis.avgMttr, 1) + 'd' : <span className="cpi-chip cpi-chip--gray">N/A</span>}</td>
                        <td><span className={chipCls(ov)}>{ovLbl}</span></td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            ) : <div className="cpi-empty-state__sub" style={{ padding: '1rem 0' }}>No domain data found.</div>}
          </div>
        </div>

        <div className="cpi-card" style={{ marginTop: '1rem' }}>
          <div className="cpi-card__title">Report Coverage</div>
          {reports.map(r => (
            <div key={r.id} className="cpi-report-row">
              <div>
                <span className="cpi-report-row__name">{r.name}</span>
                <span className="cpi-report-row__month">{r.month}</span>
              </div>
              <div className="cpi-report-row__chips">
                {(r.sheetKeys || []).map(k => <span key={k} className="cpi-chip cpi-chip--blue">{k}</span>)}
              </div>
            </div>
          ))}
        </div>
      </>
    )
  }

  /* ──────────────────────────────────────────────────────────────────
     INFERENCES TAB
  ────────────────────────────────────────────────────────────────── */
  function renderInferences() {
    const domains       = getDomains()
    const selectedDomain = getState('selectedDomain')
    const reports        = getState('reports')

    // Ensure selected domain is valid
    if (domains.length && (!selectedDomain || !domains.includes(selectedDomain))) {
      mutations.setSelectedDomain(domains[0])
    }

    return (
      <>
        <div className="cpi-card" style={{ marginBottom: '1rem' }}>
          <div className="cpi-card__title">Select GTS Product Domain</div>
          {domains.length ? (
            <div className="cpi-domain-pills">
              {domains.map(d => (
                <button
                  key={d}
                  className={`cpi-domain-pill${d === selectedDomain ? ' cpi-domain-pill--active' : ''}`}
                  onClick={() => mutations.setSelectedDomain(d)}
                >
                  {d}
                </button>
              ))}
            </div>
          ) : (
            <div style={{ marginBottom: '0.75rem' }}>
              <span className="cpi-chip cpi-chip--gray">No data — upload reports first</span>
            </div>
          )}
          <button
            className="cpi-btn cpi-btn--primary"
            onClick={handleGenerate}
            disabled={generating || !domains.length}
          >
            {generating ? (
              <><span className="cpi-spinner" /> Generating...</>
            ) : (
              <>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/>
                </svg>
                Generate Inferences &amp; Action Plan
              </>
            )}
          </button>
        </div>

        {inferOutput && (
          <>
            {/* Domain KPI summary */}
            <div className="cpi-card" style={{ marginBottom: '1rem' }}>
              <div className="cpi-section-row">
                <div className="cpi-section-row__title">{inferOutput.domain}</div>
                <span className="cpi-chip cpi-chip--blue">{reports.length} month{reports.length !== 1 ? 's' : ''} of data</span>
              </div>
              <div className="cpi-metric-grid">
                {[
                  { label: 'Saydo Ratio',    value: `${inferOutput.kpis.saydoRatio}%`,   sub: `${inferOutput.kpis.delivered}/${inferOutput.kpis.committed} pts`,                            status: rateStatus(inferOutput.kpis.saydoRatio) },
                  { label: 'Ticket Inflow',  value: inferOutput.kpis.inflow,              sub: `${inferOutput.kpis.inflowOpen} open`,                                                        status: inferOutput.kpis.inflowOpen > 0 ? 'warn' : 'good' },
                  { label: 'SLA Adherence',  value: `${inferOutput.kpis.slaAdherence}%`, sub: `Avg MTTR ${inferOutput.kpis.avgMttr !== null ? fmtNum(inferOutput.kpis.avgMttr, 1) : 'N/A'}d`, status: rateStatus(inferOutput.kpis.slaAdherence) },
                  { label: 'Aging Tickets',  value: inferOutput.kpis.agingTotal,          sub: `${inferOutput.kpis.agingCritical} critical`,                                                  status: inferOutput.kpis.agingCritical > 0 ? 'bad' : 'good' },
                  { label: 'Outage Minutes', value: Math.round(inferOutput.kpis.totalOutageMins).toLocaleString(), sub: `${(inferOutput.kpis.totalOutageMins / 60).toFixed(1)} hrs`, status: inferOutput.kpis.totalOutageMins > 10000 ? 'bad' : inferOutput.kpis.totalOutageMins > 0 ? 'warn' : 'good' },
                  { label: 'Code Coverage',  value: inferOutput.kpis.avgCoverage !== null ? fmtNum(inferOutput.kpis.avgCoverage, 1) + '%' : 'N/A', sub: inferOutput.kpis.avgCoverage !== null ? (inferOutput.kpis.avgCoverage >= 80 ? 'On target' : 'Below target') : 'No data', status: inferOutput.kpis.avgCoverage !== null ? rateStatus(inferOutput.kpis.avgCoverage) : '' },
                ].map(m => <MetricCard key={m.label} {...m} />)}
              </div>
            </div>

            {/* Inferences + Actions */}
            <div className="cpi-two-col" style={{ marginBottom: '1rem' }}>
              <div className="cpi-card">
                <div className="cpi-card__title">CPI Inferences</div>
                {inferOutput.inferences.length ? inferOutput.inferences.map((b, i) => (
                  <div key={i} className="cpi-inference-block">
                    <div className="cpi-inference-block__label">{b.cpi}</div>
                    <div className="cpi-inference-block__text">{b.text}</div>
                  </div>
                )) : <div className="cpi-empty-state__sub" style={{ padding: '1rem 0' }}>No data available for this domain.</div>}
              </div>
              <div className="cpi-card">
                <div className="cpi-card__title">Action Plan</div>
                {inferOutput.actions.map((a, i) => (
                  <div key={i} className="cpi-action-block">
                    <div className="cpi-action-block__label">{a.cpi}</div>
                    <div className="cpi-action-block__text">{a.text}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Team-level inference */}
            <div className="cpi-card">
              <div className="cpi-section-row" style={{ marginBottom: '0.75rem' }}>
                <div className="cpi-section-row__title">ACT / PCT Team-Level Inference</div>
                <span className="cpi-chip cpi-chip--blue">{inferOutput.teamRows.length} team{inferOutput.teamRows.length !== 1 ? 's' : ''}</span>
              </div>
              <p style={{ fontSize: '0.75rem', color: 'var(--txt2)', marginBottom: '1rem', lineHeight: 1.6 }}>
                The table below shows CPI metrics and a concise inference note for every ACT/PCT team
                within the <strong>{inferOutput.domain}</strong> domain.
              </p>
              {inferOutput.teamRows.length ? (
                <div className="cpi-table-wrap">
                  <table className="cpi-data-table">
                    <thead><tr>
                      <th>ACT / PCT</th><th>Saydo</th><th>SLA Adh.</th>
                      <th className="col-num">Inflow</th><th className="col-num">Critical</th>
                      <th className="col-num">Aging</th><th className="col-num">Avg Age</th>
                      <th>Status</th><th>Inference Note</th>
                    </tr></thead>
                    <tbody>
                      {inferOutput.teamRows.map(t => {
                        const saydoSt = t.saydo !== null ? rateStatus(t.saydo) : 'gray'
                        const slaSt   = t.sla   !== null ? rateStatus(t.sla)   : 'gray'
                        const ovSt    = (saydoSt === 'bad' || slaSt === 'bad' || t.critical > 0 || t.agingCritical > 0) ? 'bad'
                          : (saydoSt === 'warn' || slaSt === 'warn' || (t.avgAge !== null && t.avgAge > 30)) ? 'warn'
                          : (t.saydo !== null || t.sla !== null) ? 'good' : ''
                        const ovLabel = ovSt === 'bad' ? 'Action' : ovSt === 'warn' ? 'Watch' : ovSt === 'good' ? 'Good' : 'N/A'
                        return (
                          <tr key={t.team}>
                            <td style={{ fontWeight: 500, whiteSpace: 'nowrap' }}>{t.team}</td>
                            <td>{t.saydo !== null ? <span className={chipCls(saydoSt)}>{t.saydo}%</span> : <span className="cpi-chip cpi-chip--gray">N/A</span>}</td>
                            <td>{t.sla   !== null ? <span className={chipCls(slaSt)}>{t.sla}%</span>     : <span className="cpi-chip cpi-chip--gray">N/A</span>}</td>
                            <td className="col-num">{t.inflow}</td>
                            <td className="col-num">{t.critical > 0 ? <span style={{ color: 'var(--red)', fontWeight: 500 }}>{t.critical}</span> : '0'}</td>
                            <td className="col-num">{t.agingTotal}</td>
                            <td className="col-num">{t.avgAge !== null ? (t.avgAge > 30 ? <span style={{ color: 'var(--yellow)', fontWeight: 500 }}>{t.avgAge}d</span> : `${t.avgAge}d`) : '—'}</td>
                            <td>{ovSt ? <span className={chipCls(ovSt)}>{ovLabel}</span> : <span className="cpi-chip cpi-chip--gray">N/A</span>}</td>
                            <td style={{ fontSize: '0.7rem', color: 'var(--txt2)', lineHeight: 1.5, minWidth: 200 }}>{t.note}</td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              ) : <div className="cpi-empty-state__sub" style={{ padding: '1rem 0' }}>No team data found for this domain.</div>}
            </div>
          </>
        )}
      </>
    )
  }

  /* ──────────────────────────────────────────────────────────────────
     TREND TAB
  ────────────────────────────────────────────────────────────────── */
  function renderTrend() {
    const mergedData = getState('mergedData')
    let domains = getDomains()
    if (domainFilter !== 'ALL') domains = domains.filter(d => d === domainFilter)
    const allDomains = getDomains()
    const cpiDef = CPI_DEFS[activeCpi]

    const pivotCfg  = CPI_PIVOT_FIELDS[pivotCpi] || { cat: [], num: [] }
    const pivotAll  = [...pivotCfg.cat, ...pivotCfg.num]
    const pivotUsed = new Set([...pivotState.rows, ...pivotState.cols, ...pivotState.vals])

    const fmt = v => (!v ? '0' : v < 10 ? v.toFixed(2) : Math.round(v).toLocaleString())

    return (
      <>
        {/* Controls bar */}
        <div className="cpi-card" style={{ marginBottom: '1rem' }}>
          <div className="cpi-row-flex">
            <div>
              <div className="cpi-form-label" style={{ marginBottom: 6 }}>CPI Metric</div>
              <div className="cpi-tab-chips">
                {Object.entries(CPI_DEFS).map(([key, def]) => (
                  <button
                    key={key}
                    className={`cpi-tab-chip${activeCpi === key ? ' cpi-tab-chip--active' : ''}`}
                    onClick={() => setActiveCpi(key)}
                  >
                    {def.label}
                  </button>
                ))}
              </div>
            </div>
            <div className="cpi-form-group">
              <label className="cpi-form-label" htmlFor="trdChartType">Chart Type</label>
              <select className="cpi-form-select" id="trdChartType" value={chartType} onChange={e => setChartType(e.target.value)}>
                <option value="bar">Grouped Bar</option>
                <option value="line">Line</option>
              </select>
            </div>
            <div className="cpi-form-group">
              <label className="cpi-form-label" htmlFor="trdDomainFilter">Filter Domain</label>
              <select className="cpi-form-select" id="trdDomainFilter" value={domainFilter} onChange={e => setDomainFilter(e.target.value)}>
                <option value="ALL">All Domains</option>
                {allDomains.map(d => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>
          </div>
        </div>

        {/* Per-domain chart cards */}
        {!domains.length ? (
          <div className="cpi-empty-state">
            <svg className="cpi-empty-state__icon" width="48" height="48" viewBox="0 0 48 48" fill="none">
              <polyline points="4 40 16 24 24 32 34 18 44 28" stroke="currentColor" strokeWidth="2" fill="none"/>
            </svg>
            <div className="cpi-empty-state__title">No data yet</div>
            <div className="cpi-empty-state__sub">Upload CPI reports to see domain trends.</div>
          </div>
        ) : (
          domains.map(domain => {
            const cid  = `cpi-chart-${activeCpi}-${trdSlugify(domain)}`
            const rows = filterByDomain(mergedData[activeCpi] || [], domain)
              .filter(r => !String(r['GTS Product Domain'] || '').includes('No Mapping'))
            const cfg  = cpiDef.buildChart(rows, chartType)
            return (
              <div key={domain} className="cpi-card" style={{ marginBottom: '1rem' }}>
                <div className="cpi-section-row" style={{ marginBottom: '0.75rem' }}>
                  <div>
                    <div className="cpi-section-row__title">{domain}</div>
                    <div style={{ fontSize: '0.7rem', color: 'var(--txt3)', marginTop: 2 }}>{cpiDef.label}</div>
                  </div>
                  {cfg?.legendItems && (
                    <div className="cpi-legend">
                      {cfg.legendItems.map(li => (
                        <div key={li.label} className="cpi-legend__item">
                          <span className="cpi-legend__dot" style={{ background: li.color }} />
                          {li.label}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                <div style={{ position: 'relative', width: '100%', height: 240 }}>
                  {cfg ? (
                    <canvas id={cid} role="img" aria-label={`${cpiDef.label} trend for ${domain}`} />
                  ) : (
                    <div className="cpi-no-data">No data available for this domain.</div>
                  )}
                </div>
              </div>
            )
          })
        )}

        {/* Advanced pivot builder */}
        <details style={{ marginTop: '1rem' }}>
          <summary className="cpi-pivot-summary">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/>
              <rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/>
            </svg>
            Advanced — Custom Pivot Builder
            <svg className="cpi-details-arrow" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
              <polyline points="6 9 12 15 18 9"/>
            </svg>
          </summary>

          <div className="cpi-card" style={{ marginTop: '0.75rem' }}>
            <div className="cpi-card__title">Pivot Configuration</div>
            <div className="cpi-row-flex">
              <div className="cpi-form-group">
                <label className="cpi-form-label" htmlFor="pivotCpiSel">CPI Sheet</label>
                <select className="cpi-form-select" id="pivotCpiSel" value={pivotCpi} onChange={e => { setPivotCpi(e.target.value); setPivotState({ rows: ['Month'], cols: ['GTS Product Domain'], vals: [] }) }}>
                  {Object.entries(CPI_DEFS).map(([k, d]) => <option key={k} value={k}>{d.label}</option>)}
                </select>
              </div>
              <div className="cpi-form-group">
                <label className="cpi-form-label" htmlFor="pivotChartTypeSel">Chart Type</label>
                <select className="cpi-form-select" id="pivotChartTypeSel" value={pivotChartType} onChange={e => setPivotChartType(e.target.value)}>
                  <option value="bar">Grouped Bar</option>
                  <option value="line">Line</option>
                  <option value="bar-stacked">Stacked Bar</option>
                </select>
              </div>
              <div className="cpi-form-group">
                <label className="cpi-form-label" htmlFor="pivotAggFnSel">Aggregation</label>
                <select className="cpi-form-select" id="pivotAggFnSel" value={pivotAggFn} onChange={e => setPivotAggFn(e.target.value)}>
                  <option value="count">Count</option>
                  <option value="sum">Sum</option>
                  <option value="avg">Average</option>
                </select>
              </div>
            </div>

            <div className="cpi-pivot-builder">
              <div className="cpi-pivot-zone">
                <div className="cpi-pivot-zone__label">Available fields</div>
                <div className="cpi-pivot-chips">
                  {pivotAll.filter(f => !pivotUsed.has(f)).map(f => (
                    <span key={f} className="cpi-pivot-chip cpi-pivot-chip--available" onClick={() => pivotAddField(f)}>{f} +</span>
                  ))}
                  {pivotAll.filter(f => !pivotUsed.has(f)).length === 0 && <span className="cpi-pivot-zone__hint">All fields placed</span>}
                </div>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                {[
                  { zone: 'rows', label: 'Rows (group by)',   cls: 'rows' },
                  { zone: 'cols', label: 'Columns (series)',  cls: 'cols' },
                  { zone: 'vals', label: 'Values (measure)',  cls: 'vals' },
                ].map(({ zone, label, cls }) => (
                  <div key={zone} className="cpi-pivot-zone">
                    <div className={`cpi-pivot-zone__label cpi-pivot-zone__label--${cls}`}>{label}</div>
                    <div className="cpi-pivot-chips">
                      {pivotState[zone].map(f => (
                        <span key={f} className={`cpi-pivot-chip cpi-pivot-chip--${cls === 'rows' ? 'row' : cls === 'cols' ? 'col' : 'val'}`} onClick={() => pivotRemoveField(f, zone)}>{f} ×</span>
                      ))}
                      {!pivotState[zone].length && <span className="cpi-pivot-zone__hint">{zone === 'rows' ? 'Click a field to add' : 'Optional'}</span>}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          <div className="cpi-card" style={{ marginTop: '0.75rem' }}>
            {pivotLegendItems.length > 0 && (
              <div className="cpi-legend" style={{ marginBottom: '0.75rem' }}>
                {pivotLegendItems.map(li => (
                  <div key={li.label} className="cpi-legend__item">
                    <span className="cpi-legend__dot" style={{ background: li.color }} />
                    {li.label}
                  </div>
                ))}
              </div>
            )}
            <div style={{ position: 'relative', width: '100%', height: 300 }}>
              <canvas id="cpi-pivot-chart" role="img" aria-label="Custom pivot chart" />
            </div>
          </div>

          {pivotTableData && (
            <div className="cpi-card" style={{ marginTop: '0.75rem' }}>
              <div className="cpi-card__title">Pivot Table</div>
              <div className="cpi-table-wrap">
                <table className="cpi-data-table">
                  <thead>
                    <tr>
                      <th>{pivotTableData.rowFlds.join(' / ')}</th>
                      {pivotTableData.colKeys.map(ck => <th key={ck} className="col-num">{ck}</th>)}
                      <th className="col-num">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pivotTableData.rowKeys.map(rk => {
                      const vs  = pivotTableData.colKeys.map(ck => pivotTableData.matrix[rk][ck] || 0)
                      const tot = vs.reduce((a, b) => a + b, 0)
                      return (
                        <tr key={rk}>
                          <td>{rk}</td>
                          {vs.map((v, i) => <td key={i} className="col-num">{fmt(v)}</td>)}
                          <td className="col-num col-bold">{fmt(tot)}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </details>
      </>
    )
  }

  /* ──────────────────────────────────────────────────────────────────
     MAIN RENDER
  ────────────────────────────────────────────────────────────────── */
  if (loading) {
    return (
      <div className="cpi-root">
        <div className="cpi-loading">
          <span className="cpi-loading__spinner" />
          <span>Loading CPI data…</span>
        </div>
      </div>
    )
  }

  if (error) {
    return (
      <div className="cpi-root">
        <div className="cpi-error">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
          </svg>
          <div>
            <strong>Failed to load CPI data.</strong>
            <div style={{ marginTop: 4, fontSize: '0.8rem' }}>{error}</div>
          </div>
        </div>
      </div>
    )
  }

  const TABS = [
    { id: 'overview',    label: 'Overview'   },
    { id: 'inferences',  label: 'Inferences' },
    { id: 'trend',       label: 'Trend'      },
  ]

  return (
    <div className="cpi-root">
      {/* Internal tab bar */}
      <div className="cpi-topbar">
        <div className="cpi-tab-bar">
          {TABS.map(tab => (
            <button
              key={tab.id}
              className={`cpi-tab-btn${activeTab === tab.id ? ' cpi-tab-btn--active' : ''}`}
              onClick={() => setActiveTab(tab.id)}
            >
              {tab.label}
            </button>
          ))}
        </div>
        {lastUpdated && (
          <span className="cpi-last-updated">Data as of {lastUpdated}</span>
        )}
      </div>

      <div className="cpi-content">
        {activeTab === 'overview'   && renderOverview()}
        {activeTab === 'inferences' && renderInferences()}
        {activeTab === 'trend'      && renderTrend()}
      </div>
    </div>
  )
}
