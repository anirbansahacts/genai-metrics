import { useState, useEffect, useRef } from 'react'
import * as XLSX from 'xlsx'
import { Chart, registerables } from 'chart.js'
import './SprintProductivity.css'
import { useTheme } from '../context/ThemeContext'

Chart.register(...registerables)

// ── Data-labels plugin (registered once at module level) ──────────────────
Chart.register({
  id: 'spDataLabels',
  afterDatasetsDraw(chart) {
    const ctx = chart.ctx
    chart.data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di)
      if (meta.hidden) return
      meta.data.forEach((bar, i) => {
        const val = ds.data[i]
        if (!val) return
        const isHoriz = chart.options.indexAxis === 'y'
        const label = typeof val === 'number' && val % 1 !== 0 ? val.toFixed(1) : String(val)
        ctx.save()
        ctx.font = '500 9px Outfit,sans-serif'
        ctx.fillStyle = '#E6EDF3'
        ctx.textBaseline = 'middle'
        if (isHoriz) {
          ctx.textAlign = 'left'
          ctx.fillText(label + '%', Math.min(bar.x + 7, chart.chartArea.right - 28), bar.y)
        } else {
          ctx.textAlign = 'center'
          ctx.fillText(label, bar.x, bar.y - 8)
        }
        ctx.restore()
      })
    })
  }
})

// ── Constants ─────────────────────────────────────────────────────────────
const REQUIRED = ['Tower', 'ACT/PCT/Project', '#US/Defects', '#US/Defects SP', '#GenAI SubTask', '#GenAI Saved Hours', '#GenAI Saved SP']
const COLORS = ['#39D98A', '#58A6FF', '#F0883E', '#FF7EB6', '#BC8CFF', '#E3B341', '#F85149', '#3BC9DB', '#63E6BE', '#74C0FC']
const METRICS = [
  { key: '#US/Defects',        label: 'US/Def',  cls: 'cv-usd', thCls: 'th-metric-usd' },
  { key: '#US/Defects SP',     label: 'Def SP',  cls: 'cv-usp', thCls: 'th-metric-usp' },
  { key: '#GenAI SubTask',     label: 'SubTask', cls: 'cv-sub', thCls: 'th-metric-sub' },
  { key: '#GenAI Saved Hours', label: 'Hrs',     cls: 'cv-hrs', thCls: 'th-metric-hrs' },
  { key: '#GenAI Saved SP',    label: 'SavedSP', cls: 'cv-ssp', thCls: 'th-metric-ssp' },
  { key: '__prod__',           label: 'Prod%',   cls: 'cv-pct', thCls: 'th-metric-pct' },
]
// ══════════════════════════════════════════════════════════════════════════
export default function SprintProductivity() {
  const { theme } = useTheme()

  function getChartTheme() {
    const dark = theme === 'dark'
    return {
      tooltip: {
        backgroundColor: dark ? '#161B22' : '#FFFFFF',
        borderColor:     dark ? '#30363D' : '#D0D7E2',
        borderWidth: 1,
        titleColor:      dark ? '#E6EDF3' : '#1A1F2E',
        bodyColor:       dark ? '#8B949E' : '#5A6478',
      },
      grid:        dark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.06)',
      tickMuted:   dark ? '#6E7681' : '#8892A4',
      tickNormal:  dark ? '#8B949E' : '#5A6478',
      legendColor: dark ? '#8B949E' : '#5A6478',
    }
  }

  // ── React state ─────────────────────────────────────────────────────────
  const [activePage,        setActivePage]        = useState('dashboard')
  const [activeSprint,      setActiveSprint]      = useState('ALL')
  const [adoptActiveSprint, setAdoptActiveSprint] = useState('ALL')
  const [loading,           setLoading]           = useState(true)
  const [error,             setError]             = useState(null)
  const [uploadedAt,        setUploadedAt]        = useState('')
  const [dataVersion,       setDataVersion]       = useState(0)

  // ── Mutable refs (for imperative rendering functions) ────────────────────
  const sprintDataRef          = useRef({})
  const activeSprintRef        = useRef('ALL')
  const adoptActiveSprintRef   = useRef('ALL')
  const chartsRef              = useRef({ c1: null, c2: null, c3: null, cAdopt: null, cAiTrend: null })
  const lastParticipatedRef    = useRef([])
  const lastNotParticipatedRef = useRef([])
  const lastSprintTableRowsRef = useRef(null)
  const lastAdoptionDataRef    = useRef([])

  // Keep refs in sync with state
  useEffect(() => { activeSprintRef.current      = activeSprint      }, [activeSprint])
  useEffect(() => { adoptActiveSprintRef.current = adoptActiveSprint }, [adoptActiveSprint])

  // ── Auto-fetch on mount ─────────────────────────────────────────────────
  useEffect(() => {
    async function fetchData() {
      try {
        setLoading(true)
        setError(null)
        const base = import.meta.env.VITE_STORAGE_URL
        const cfgRes = await fetch(`${base}/sprint-productivity/config.json`)
        if (!cfgRes.ok) throw new Error(`Config fetch failed (HTTP ${cfgRes.status})`)
        const config = await cfgRes.json()
        setUploadedAt(config.uploadedAt || '')

        const xlsxRes = await fetch(`${base}/sprint-productivity/${config.latestFile}`)
        if (!xlsxRes.ok) throw new Error(`Data file fetch failed (HTTP ${xlsxRes.status})`)
        const buffer = await xlsxRes.arrayBuffer()
        const wb = XLSX.read(buffer, { type: 'array' })

        const data = {}
        for (const sn of wb.SheetNames) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: 0 })
          if (!rows.length) continue
          const missing = REQUIRED.filter(c => !Object.keys(rows[0]).includes(c))
          if (missing.length) { console.warn(`Sheet "${sn}" missing: ${missing.join(', ')}`); continue }
          data[sn] = rows
        }
        if (!Object.keys(data).length) throw new Error('No valid sprint sheets found in the Excel file.')

        sprintDataRef.current = data
        setDataVersion(v => v + 1)
        setLoading(false)
      } catch (err) {
        setError(err.message)
        setLoading(false)
      }
    }
    fetchData()
  }, [])

  // ── Re-render when data / active sprint / page changes ──────────────────
  useEffect(() => {
    if (loading || !Object.keys(sprintDataRef.current).length) return
    if (activePage === 'dashboard') {
      updateAllFilters()
      renderKPIs()
      renderSprintTable()
      renderParticipated()
      requestAnimationFrame(() => { renderChart1(); renderChart2(); renderChart3() })
    } else if (activePage === 'adoption') {
      updateAllFilters()
      requestAnimationFrame(() => { renderAdoptionTable(); renderAdoptionTrend() })
    } else if (activePage === 'ai') {
      updateAiScope()
      populateAiSprintScope()
    }
  }, [dataVersion, activeSprint, activePage, adoptActiveSprint, loading, theme]) // eslint-disable-line

  // ── Expose drill-toggle functions to window (called from injected HTML) ──
  useEffect(() => {
    window._sp_toggleMatrixDrill = toggleMatrixDrill
    window._sp_togglePartDrill   = togglePartDrill
    window._sp_toggleAdoptDrill  = toggleAdoptDrill
    return () => {
      delete window._sp_toggleMatrixDrill
      delete window._sp_togglePartDrill
      delete window._sp_toggleAdoptDrill
    }
  }, []) // eslint-disable-line

  // ── Destroy all charts on unmount ───────────────────────────────────────
  useEffect(() => {
    return () => { Object.values(chartsRef.current).forEach(c => { if (c) c.destroy() }) }
  }, [])

  // ══════════════════════════════════════════════════════════════════════════
  // HELPERS
  // ══════════════════════════════════════════════════════════════════════════
  function getSprints()    { return Object.keys(sprintDataRef.current).sort() }
  function getAllTowers()  { return [...new Set(Object.values(sprintDataRef.current).flat().map(r => r['Tower']).filter(Boolean))].sort() }
  function getAllProjects(tower) {
    let rows = Object.values(sprintDataRef.current).flat()
    if (tower && tower !== 'ALL') rows = rows.filter(r => r['Tower'] === tower)
    return [...new Set(rows.map(r => r['ACT/PCT/Project']).filter(Boolean))].sort()
  }
  function getRows(sprint, tower, project) {
    let rows = sprint === 'ALL' ? Object.values(sprintDataRef.current).flat() : (sprintDataRef.current[sprint] || [])
    if (tower   && tower   !== 'ALL') rows = rows.filter(r => r['Tower']            === tower)
    if (project && project !== 'ALL') rows = rows.filter(r => r['ACT/PCT/Project']  === project)
    return rows
  }
  function sumC(rows, col)  { return rows.reduce((a, r) => a + (parseFloat(r[col]) || 0), 0) }
  function prod(rows)        { const d = sumC(rows, '#US/Defects SP'); return d > 0 ? sumC(rows, '#GenAI Saved SP') / d : 0 }
  function fmtV(key, v) {
    if (key === '__prod__')            return v.pct.toFixed(2) + '%'
    if (key === '#US/Defects')         return Math.round(v.usd).toLocaleString()
    if (key === '#US/Defects SP')      return Math.round(v.usp).toLocaleString()
    if (key === '#GenAI SubTask')      return Math.round(v.sub).toLocaleString()
    if (key === '#GenAI Saved Hours')  return v.hrs.toFixed(1)
    if (key === '#GenAI Saved SP')     return v.ssp.toFixed(1)
    return ''
  }
  function calcV(rows) {
    const usd = sumC(rows, '#US/Defects'), usp = sumC(rows, '#US/Defects SP'),
          sub = sumC(rows, '#GenAI SubTask'), hrs = sumC(rows, '#GenAI Saved Hours'),
          ssp = sumC(rows, '#GenAI Saved SP')
    return { usd, usp, sub, hrs, ssp, pct: usp > 0 ? ssp / usp * 100 : 0 }
  }
  function today() { return new Date().toISOString().slice(0, 10) }
  function showToast(msg, type = '') {
    const t = document.getElementById('sp-toast')
    if (!t) return
    t.textContent = msg
    t.className = 'sp-toast show' + (type ? ' ' + type : '')
    setTimeout(() => { t.className = 'sp-toast' }, 3200)
  }
  function triggerDL(blob, filename) {
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = filename
    document.body.appendChild(a)
    a.click()
    setTimeout(() => { URL.revokeObjectURL(a.href); document.body.removeChild(a) }, 500)
  }

  // ══════════════════════════════════════════════════════════════════════════
  // FILTER UPDATES
  // ══════════════════════════════════════════════════════════════════════════
  function updateAllFilters() {
    ;['towerFilter', 'partTowerFilter', 'adoptTowerFilter'].forEach(id => {
      const s = document.getElementById(id); if (!s) return
      const p = s.value
      s.innerHTML = '<option value="ALL">All towers</option>'
      getAllTowers().forEach(t => { const o = document.createElement('option'); o.value = t; o.textContent = t; s.appendChild(o) })
      if (getAllTowers().includes(p)) s.value = p
    })
    const tf = document.getElementById('towerFilter');   if (tf) tf.value = 'ALL'
    const pf = document.getElementById('projectFilter'); if (pf) pf.value = 'ALL'
    updateProjectDropdown('projectFilter', 'ALL')
    updateProjectDropdown('partProjectFilter', document.getElementById('partTowerFilter')?.value || 'ALL')
  }
  function updateProjectDropdown(id, tower) {
    const sel = document.getElementById(id); if (!sel) return
    const prev = sel.value
    sel.innerHTML = '<option value="ALL">All projects</option>'
    getAllProjects(tower).forEach(p => { const o = document.createElement('option'); o.value = p; o.textContent = p.length > 32 ? p.slice(0, 30) + '…' : p; sel.appendChild(o) })
    if (getAllProjects(tower).includes(prev)) sel.value = prev
  }

  // ══════════════════════════════════════════════════════════════════════════
  // KPIs
  // ══════════════════════════════════════════════════════════════════════════
  function renderKPIs() {
    const rows    = getRows(activeSprintRef.current, 'ALL')
    const sprints = getSprints()
    const p = prod(rows)
    let delta = null
    if (activeSprintRef.current !== 'ALL') {
      const idx = sprints.indexOf(activeSprintRef.current)
      if (idx > 0) { const pp = prod(getRows(sprints[idx - 1], 'ALL')); delta = pp > 0 ? (p - pp) / pp * 100 : null }
    }
    const kpis = [
      { label: 'Productivity ratio', value: (p * 100).toFixed(1) + '%', sub: 'Saved SP / Total SP',   delta, color: 'var(--accent)' },
      { label: 'US / Defects',       value: Math.round(sumC(rows, '#US/Defects')).toLocaleString(),    sub: 'stories & defects', color: 'var(--blue)'   },
      { label: 'Story points',       value: Math.round(sumC(rows, '#US/Defects SP')).toLocaleString(), sub: 'total SP',          color: 'var(--purple)' },
      { label: 'GenAI subtasks',     value: Math.round(sumC(rows, '#GenAI SubTask')).toLocaleString(), sub: 'created',           color: 'var(--orange)' },
      { label: 'Hours saved',        value: Math.round(sumC(rows, '#GenAI Saved Hours')).toLocaleString(), sub: 'by GenAI',      color: 'var(--yellow)' },
      { label: 'SP saved',           value: Math.round(sumC(rows, '#GenAI Saved SP')).toLocaleString(),    sub: 'story points',  color: 'var(--pink)'   },
    ]
    const el = document.getElementById('kpiStrip'); if (!el) return
    el.innerHTML = kpis.map(k => `
      <div class="kpi-card" style="--kpi-color:${k.color}">
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value">${k.value}</div>
        <div class="kpi-sub">${k.sub}</div>
        ${k.delta !== null ? `<div class="kpi-delta ${k.delta >= 0 ? 'up' : 'down'}">${k.delta >= 0 ? '↑' : '↓'} ${Math.abs(k.delta).toFixed(1)}% vs prev</div>` : ''}
      </div>`).join('')
  }

  // ══════════════════════════════════════════════════════════════════════════
  // CHARTS
  // ══════════════════════════════════════════════════════════════════════════
  function renderChart1() {
    const sprints = getSprints()
    const data = sprints.map(s => parseFloat((prod(getRows(s, 'ALL')) * 100).toFixed(2)))
    const mx   = Math.max(...data, 1)
    if (chartsRef.current.c1) chartsRef.current.c1.destroy()
    const canvas = document.getElementById('chart1'); if (!canvas) return
    chartsRef.current.c1 = new Chart(canvas, {
      type: 'bar',
      data: { labels: sprints, datasets: [{ label: 'Productivity %', data, backgroundColor: sprints.map((_, i) => COLORS[i % COLORS.length]), borderRadius: 5, borderSkipped: false }] },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 20 } },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: v => v.raw + '%' }, ...getChartTheme().tooltip } },
        scales: {
          y: { beginAtZero: true, max: Math.ceil(mx * 1.3), ticks: { callback: v => v + '%', color: getChartTheme().tickMuted, font: { size: 10 } }, grid: { color: getChartTheme().grid }, border: { color: 'transparent' } },
          x: { ticks: { color: getChartTheme().tickNormal, font: { size: 10 } }, grid: { display: false }, border: { color: 'transparent' } }
        }
      }
    })
  }
  function renderChart2() {
    const tower   = document.getElementById('towerFilter')?.value   || 'ALL'
    const project = document.getElementById('projectFilter')?.value || 'ALL'
    const sprints = getSprints()
    const cols    = ['#US/Defects', '#US/Defects SP', '#GenAI SubTask', '#GenAI Saved Hours', '#GenAI Saved SP']
    const cc      = [COLORS[0], COLORS[1], COLORS[2], COLORS[3], COLORS[4]]
    const datasets = cols.map((col, i) => ({ label: col, data: sprints.map(s => parseFloat(sumC(getRows(s, tower, project), col).toFixed(1))), backgroundColor: cc[i], borderRadius: 3, borderSkipped: false }))
    const mx = Math.max(...datasets.flatMap(d => d.data), 1)
    const l2 = document.getElementById('legend2')
    if (l2) l2.innerHTML = cols.map((c, i) => `<div class="legend-item"><div class="legend-swatch" style="background:${cc[i]}"></div>${c}</div>`).join('')
    if (chartsRef.current.c2) chartsRef.current.c2.destroy()
    const canvas = document.getElementById('chart2'); if (!canvas) return
    chartsRef.current.c2 = new Chart(canvas, {
      type: 'bar', data: { labels: sprints, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 20 } },
        plugins: { legend: { display: false }, tooltip: { ...getChartTheme().tooltip } },
        scales: {
          x: { ticks: { color: getChartTheme().tickNormal, font: { size: 10 } }, grid: { display: false }, border: { color: 'transparent' } },
          y: { max: Math.ceil(mx * 1.2), ticks: { color: getChartTheme().tickMuted, font: { size: 10 } }, grid: { color: getChartTheme().grid }, border: { color: 'transparent' } }
        }
      }
    })
  }
  function renderChart3() {
    const sprints  = activeSprintRef.current === 'ALL' ? getSprints() : [activeSprintRef.current]
    const towers   = getAllTowers()
    const datasets = sprints.map((s, i) => ({
      label: s,
      data: towers.map(t => parseFloat((prod(getRows(s, t)) * 100).toFixed(2))),
      backgroundColor: COLORS[i % COLORS.length], borderRadius: 3, borderSkipped: false
    }))
    const mx = Math.max(...datasets.flatMap(d => d.data), 1)
    const h  = Math.max(220, towers.length * 46 + 80)
    const wrap = document.getElementById('chart3wrap'); if (wrap) wrap.style.height = h + 'px'
    if (chartsRef.current.c3) chartsRef.current.c3.destroy()
    const canvas = document.getElementById('chart3'); if (!canvas) return
    chartsRef.current.c3 = new Chart(canvas, {
      type: 'bar', data: { labels: towers, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, indexAxis: 'y', layout: { padding: { right: 50 } },
        plugins: {
          legend: { display: sprints.length > 1, labels: { color: getChartTheme().legendColor, boxWidth: 9, font: { size: 10 } } },
          tooltip: { callbacks: { label: v => v.raw + '%' }, ...getChartTheme().tooltip }
        },
        scales: {
          x: { beginAtZero: true, max: Math.ceil(mx * 1.35), ticks: { callback: v => v + '%', color: getChartTheme().tickMuted, font: { size: 10 } }, grid: { color: getChartTheme().grid }, border: { color: 'transparent' } },
          y: { ticks: { color: getChartTheme().tickNormal, font: { size: 10 } }, grid: { display: false }, border: { color: 'transparent' } }
        }
      }
    })
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SPRINT MATRIX — with drill-down
  // ══════════════════════════════════════════════════════════════════════════
  function renderSprintTable() {
    const sprints = getSprints()
    const towers  = getAllTowers()
    const nM = METRICS.length
    const pivot = {}
    towers.forEach(t => { pivot[t] = {}; sprints.forEach(s => { pivot[t][s] = calcV(getRows(s, t)) }) })
    const sprintTotals = {}
    sprints.forEach(s => {
      sprintTotals[s] = { usd: 0, usp: 0, sub: 0, hrs: 0, ssp: 0 }
      towers.forEach(t => { const v = pivot[t][s]; sprintTotals[s].usd += v.usd; sprintTotals[s].usp += v.usp; sprintTotals[s].sub += v.sub; sprintTotals[s].hrs += v.hrs; sprintTotals[s].ssp += v.ssp })
      sprintTotals[s].pct = sprintTotals[s].usp > 0 ? sprintTotals[s].ssp / sprintTotals[s].usp * 100 : 0
    })
    const projPivot = {}
    towers.forEach(t => {
      projPivot[t] = {}
      getAllProjects(t).forEach(pr => { projPivot[t][pr] = {}; sprints.forEach(s => { projPivot[t][pr][s] = calcV(getRows(s, t, pr)) }) })
    })
    lastSprintTableRowsRef.current = { pivot, sprints, towers, sprintTotals, METRICS, projPivot }

    let thead1 = '<tr><th class="th-tower" rowspan="2">Tower / Project</th>'
    let thead2 = '<tr>'
    sprints.forEach((s, si) => {
      const sep = si < sprints.length - 1 ? 'border-right:2px solid var(--border2);' : ''
      thead1 += `<th class="th-sprint" colspan="${nM}" style="${sep}">${s}</th>`
      METRICS.forEach((m, mi) => {
        const isLast = mi === nM - 1 && si < sprints.length - 1
        thead2 += `<th class="th-metric ${m.thCls}" style="${isLast ? 'border-right:2px solid var(--border2);' : ''}">${m.label}</th>`
      })
    })
    thead1 += '</tr>'; thead2 += '</tr>'

    let tbody = ''
    towers.forEach((t, ti) => {
      const rowId = 'mrow_' + ti
      let tRow = `<tr><td class="td-tower-name" onclick="window._sp_toggleMatrixDrill('${rowId}',this)"><span class="expand-icon">▶</span><strong>${t}</strong></td>`
      sprints.forEach((s, si) => {
        const v = pivot[t][s]
        METRICS.forEach((m, mi) => {
          const isLast = mi === nM - 1 && si < sprints.length - 1
          tRow += `<td class="td-val ${m.cls}" style="${isLast ? 'border-right:2px solid var(--border2);' : ''}">${fmtV(m.key, v)}</td>`
        })
      })
      tRow += '</tr>'; tbody += tRow
      getAllProjects(t).forEach(pr => {
        let dRow = `<tr class="drill-row" data-group="${rowId}">`
        dRow += `<td class="td-tower-name" style="padding-left:26px;font-size:10px;color:var(--txt2);font-weight:400;background:var(--bg3);cursor:default;">↳ ${pr}</td>`
        sprints.forEach((s, si) => {
          const v = projPivot[t][pr][s]
          METRICS.forEach((m, mi) => {
            const isLast = mi === nM - 1 && si < sprints.length - 1
            dRow += `<td class="td-val ${m.cls}" style="font-size:10px;background:var(--bg3);${isLast ? 'border-right:2px solid var(--border2);' : ''}">${fmtV(m.key, v)}</td>`
          })
        })
        dRow += '</tr>'; tbody += dRow
      })
    })
    let totRow = '<tr class="total-row"><td class="td-tower-name">Grand Total</td>'
    sprints.forEach((s, si) => {
      const v = sprintTotals[s]
      METRICS.forEach((m, mi) => {
        const isLast = mi === nM - 1 && si < sprints.length - 1
        totRow += `<td class="td-val ${m.cls}" style="font-weight:700;${isLast ? 'border-right:2px solid var(--border2);' : ''}">${fmtV(m.key, v)}</td>`
      })
    })
    totRow += '</tr>'
    const wrap = document.getElementById('sprintMatrixWrap'); if (!wrap) return
    wrap.innerHTML = `<table class="matrix-table"><thead>${thead1}${thead2}</thead><tbody>${tbody}${totRow}</tbody></table>`
  }
  function toggleMatrixDrill(groupId, tdEl) {
    const rows  = document.querySelectorAll(`tr[data-group="${groupId}"]`)
    const isOpen = rows.length && rows[0].classList.contains('visible')
    rows.forEach(r => r.classList.toggle('visible', !isOpen))
    const icon = tdEl.querySelector('.expand-icon')
    if (icon) icon.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(90deg)'
  }

  // ══════════════════════════════════════════════════════════════════════════
  // PARTICIPATED — grouped by tower with drill-down
  // ══════════════════════════════════════════════════════════════════════════
  function renderParticipated() {
    const tower   = document.getElementById('partTowerFilter')?.value   || 'ALL'
    const project = document.getElementById('partProjectFilter')?.value || 'ALL'
    const rows    = getRows(activeSprintRef.current, tower, project)
    const byTower = {}
    rows.forEach(r => {
      const t = r['Tower']; const pr = r['ACT/PCT/Project']
      if (!byTower[t]) byTower[t] = { tower: t, usd: 0, usp: 0, sub: 0, hours: 0, sp: 0, pct: 0, projects: {} }
      if (!byTower[t].projects[pr]) byTower[t].projects[pr] = { proj: pr, usd: 0, usp: 0, sub: 0, hours: 0, sp: 0, pct: 0 }
      const vals = [['usd', '#US/Defects'], ['usp', '#US/Defects SP'], ['sub', '#GenAI SubTask'], ['hours', '#GenAI Saved Hours'], ['sp', '#GenAI Saved SP']]
      vals.forEach(([k, c]) => { byTower[t][k] += (parseFloat(r[c]) || 0); byTower[t].projects[pr][k] += (parseFloat(r[c]) || 0) })
    })
    Object.values(byTower).forEach(t => {
      t.pct = t.usp > 0 ? t.sp / t.usp * 100 : 0
      Object.values(t.projects).forEach(p => { p.pct = p.usp > 0 ? p.sp / p.usp * 100 : 0 })
    })
    const partTowers    = Object.values(byTower).filter(t => t.hours > 0).sort((a, b) => b.hours - a.hours)
    const projMap = {}
    rows.forEach(r => {
      const k = r['Tower'] + '||' + r['ACT/PCT/Project']
      if (!projMap[k]) projMap[k] = { tower: r['Tower'], proj: r['ACT/PCT/Project'], hours: 0, usd: 0, usp: 0 }
      projMap[k].hours += (parseFloat(r['#GenAI Saved Hours']) || 0)
      projMap[k].usd   += (parseFloat(r['#US/Defects']) || 0)
      projMap[k].usp   += (parseFloat(r['#US/Defects SP']) || 0)
    })
    lastParticipatedRef.current    = partTowers
    lastNotParticipatedRef.current = Object.values(projMap).filter(p => p.hours === 0).sort((a, b) => a.proj.localeCompare(b.proj))

    const pc = document.getElementById('partCount');    if (pc) pc.textContent = partTowers.length + ' tower' + (partTowers.length !== 1 ? 's' : '')
    const nc = document.getElementById('notPartCount'); if (nc) nc.textContent = lastNotParticipatedRef.current.length + ' project' + (lastNotParticipatedRef.current.length !== 1 ? 's' : '')
    const nl = document.getElementById('notPartLabel'); if (nl) nl.textContent = ''

    let html = ''
    partTowers.forEach((t, ti) => {
      const gid = 'ptg_' + ti
      const rc  = ti === 0 ? 'top1' : ti === 1 ? 'top2' : ti === 2 ? 'top3' : ''
      html += `<tr class="tower-row">
        <td class="td-rank"><span class="rank-badge ${rc}">${ti + 1}</span></td>
        <td class="td-tower" onclick="window._sp_togglePartDrill('${gid}',this)" style="cursor:pointer;"><span class="expand-icon">▶</span><strong>${t.tower}</strong></td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--blue);">${Math.round(t.usd)}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--purple);">${Math.round(t.usp)}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--orange);">${Math.round(t.sub)}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent);font-weight:500;">${t.hours.toFixed(1)}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--blue);">${t.sp.toFixed(1)}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--pink);font-weight:600;">${t.pct.toFixed(2)}%</td>
      </tr>`
      Object.values(t.projects).filter(p => p.hours > 0).sort((a, b) => b.hours - a.hours).forEach(p => {
        html += `<tr class="drill-proj-row" data-pgroup="${gid}">
          <td></td>
          <td style="padding-left:22px;font-size:10px;color:var(--txt2);">↳ ${p.proj}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--blue);">${Math.round(p.usd)}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--purple);">${Math.round(p.usp)}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--orange);">${Math.round(p.sub)}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--accent);">${p.hours.toFixed(1)}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--blue);">${p.sp.toFixed(1)}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--pink);">${p.pct.toFixed(2)}%</td>
        </tr>`
      })
    })
    if (!html) html = '<tr><td colspan="8" class="empty">No participated projects</td></tr>'
    const pb = document.getElementById('participatedBody'); if (pb) pb.innerHTML = html

    const nb = document.getElementById('notParticipatedBody'); if (!nb) return
    nb.innerHTML = lastNotParticipatedRef.current.length
      ? lastNotParticipatedRef.current.map((it, i) => `<tr><td class="td-rank"><span class="rank-badge">${i + 1}</span></td><td class="td-tower">${it.tower}</td><td class="td-proj">${it.proj}</td><td style="text-align:right;font-family:'DM Mono',monospace;color:var(--orange);">${Math.round(it.usd)}</td><td style="text-align:right;font-family:'DM Mono',monospace;color:var(--purple);">${Math.round(it.usp)}</td></tr>`).join('')
      : '<tr><td colspan="5" class="empty">🎉 All projects have participated!</td></tr>'
  }
  function togglePartDrill(gid, tdEl) {
    const rows  = document.querySelectorAll(`tr[data-pgroup="${gid}"]`)
    const isOpen = rows.length && rows[0].classList.contains('visible')
    rows.forEach(r => r.classList.toggle('visible', !isOpen))
    const icon = tdEl.querySelector('.expand-icon'); if (icon) icon.style.transform = isOpen ? '' : 'rotate(90deg)'
  }

  // ══════════════════════════════════════════════════════════════════════════
  // ADOPTION TABLE & TREND
  // ══════════════════════════════════════════════════════════════════════════
  function renderAdoptionTable() {
    const towers = getAllTowers()
    let tbody = ''
    lastAdoptionDataRef.current = []
    towers.forEach((t, ti) => {
      const rows = getRows(adoptActiveSprintRef.current === 'ALL' ? 'ALL' : adoptActiveSprintRef.current, t)
      const usd  = sumC(rows, '#US/Defects')
      const sub  = sumC(rows, '#GenAI SubTask')
      const pct  = usd > 0 ? (sub / usd * 100) : 0
      lastAdoptionDataRef.current.push({ tower: t, usd, sub, pct })
      const barColor = pct >= 80 ? '#39D98A' : pct >= 50 ? '#E3B341' : '#F85149'
      const gid = 'adg_' + ti
      tbody += `<tr onclick="window._sp_toggleAdoptDrill('${gid}',this)" style="cursor:pointer;">
        <td style="font-weight:500;color:var(--txt);"><span class="expand-icon" style="display:inline-block;margin-right:5px;font-size:9px;color:var(--txt3);transition:transform .2s;">▶</span>${t}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--blue);">${Math.round(usd).toLocaleString()}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--orange);">${Math.round(sub).toLocaleString()}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:${barColor};font-weight:600;">${pct.toFixed(1)}%</td>
        <td><div class="adopt-pct-bar"><div class="adopt-pct-track"><div class="adopt-pct-fill" style="width:${Math.min(pct, 100)}%;background:${barColor};"></div></div></div></td>
      </tr>`
      getAllProjects(t).forEach(pr => {
        const pRows = getRows(adoptActiveSprintRef.current === 'ALL' ? 'ALL' : adoptActiveSprintRef.current, t, pr)
        const pUsd  = sumC(pRows, '#US/Defects')
        const pSub  = sumC(pRows, '#GenAI SubTask')
        const pPct  = pUsd > 0 ? (pSub / pUsd * 100) : 0
        const pBarColor = pPct >= 80 ? '#39D98A' : pPct >= 50 ? '#E3B341' : '#F85149'
        tbody += `<tr class="drill-row" data-adgroup="${gid}" style="display:none;">
          <td style="padding-left:22px;font-size:10px;color:var(--txt2);">↳ ${pr}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--blue);">${Math.round(pUsd).toLocaleString()}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:var(--orange);">${Math.round(pSub).toLocaleString()}</td>
          <td style="text-align:right;font-family:'DM Mono',monospace;font-size:10px;color:${pBarColor};font-weight:600;">${pPct.toFixed(1)}%</td>
          <td><div class="adopt-pct-bar"><div class="adopt-pct-track"><div class="adopt-pct-fill" style="width:${Math.min(pPct, 100)}%;background:${pBarColor};"></div></div></div></td>
        </tr>`
      })
    })
    const ab = document.getElementById('adoptionBody'); if (!ab) return
    ab.innerHTML = tbody || '<tr><td colspan="5" class="empty">No data</td></tr>'
  }
  function toggleAdoptDrill(gid, rowEl) {
    const drows = document.querySelectorAll(`tr[data-adgroup="${gid}"]`)
    const isOpen = drows.length && drows[0].style.display !== 'none'
    drows.forEach(r => r.style.display = isOpen ? 'none' : 'table-row')
    const icon = rowEl.querySelector('.expand-icon'); if (icon) icon.style.transform = isOpen ? '' : 'rotate(90deg)'
  }
  function renderAdoptionTrend() {
    const filterTower = document.getElementById('adoptTowerFilter')?.value || 'ALL'
    const sprints = getSprints()
    const towers  = filterTower === 'ALL' ? getAllTowers() : [filterTower]
    const datasets = towers.slice(0, 8).map((t, i) => ({
      label: t,
      data: sprints.map(s => { const rows = getRows(s, t); const usd = sumC(rows, '#US/Defects'); const sub = sumC(rows, '#GenAI SubTask'); return usd > 0 ? parseFloat((sub / usd * 100).toFixed(2)) : 0 }),
      borderColor: COLORS[i % COLORS.length], backgroundColor: 'transparent',
      borderWidth: 2, pointRadius: 4, pointBackgroundColor: COLORS[i % COLORS.length], tension: 0.35
    }))
    if (chartsRef.current.cAdopt) chartsRef.current.cAdopt.destroy()
    const canvas = document.getElementById('chartAdopt'); if (!canvas) return
    chartsRef.current.cAdopt = new Chart(canvas, {
      type: 'line', data: { labels: sprints, datasets },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 16 } },
        plugins: { legend: { display: towers.length > 1, labels: { color: getChartTheme().legendColor, boxWidth: 9, font: { size: 10 } } }, tooltip: { callbacks: { label: v => `${v.dataset.label}: ${v.raw}%` }, ...getChartTheme().tooltip } },
        scales: {
          y: { beginAtZero: true, ticks: { callback: v => v + '%', color: getChartTheme().tickMuted, font: { size: 10 } }, grid: { color: getChartTheme().grid }, border: { color: 'transparent' } },
          x: { ticks: { color: getChartTheme().tickNormal, font: { size: 10 } }, grid: { display: false }, border: { color: 'transparent' } }
        }
      }
    })
  }

  // ══════════════════════════════════════════════════════════════════════════
  // AI INSIGHTS PAGE
  // ══════════════════════════════════════════════════════════════════════════
  function updateAiScope() {
    const type  = document.getElementById('aiScopeType')?.value
    const wrap  = document.getElementById('aiScopeItemWrap')
    const label = document.getElementById('aiScopeItemLabel')
    const sel   = document.getElementById('aiScopeItem')
    if (!type || !wrap) return
    if (type === 'overall') { wrap.style.display = 'none'; return }
    wrap.style.display = 'block'
    if (type === 'tower') {
      if (label) label.textContent = 'Tower'
      if (sel) { sel.innerHTML = ''; getAllTowers().forEach(t => { const o = document.createElement('option'); o.value = t; o.textContent = t; sel.appendChild(o) }) }
    } else {
      if (label) label.textContent = 'Project'
      if (sel) { sel.innerHTML = ''; getAllProjects('ALL').forEach(p => { const o = document.createElement('option'); o.value = p; o.textContent = p; sel.appendChild(o) }) }
    }
  }
  function populateAiSprintScope() {
    const sel = document.getElementById('aiSprintScope'); if (!sel) return
    const prev = sel.value
    sel.innerHTML = '<option value="ALL">All sprints</option>'
    getSprints().forEach(s => { const o = document.createElement('option'); o.value = s; o.textContent = s; sel.appendChild(o) })
    if (prev) sel.value = prev
  }
  async function runAiInference() {
    const panel     = document.getElementById('aiResultPanel')
    const btn       = document.getElementById('aiRunBtn')
    const type      = document.getElementById('aiScopeType')?.value
    const sprint    = document.getElementById('aiSprintScope')?.value
    const scopeItem = type !== 'overall' ? document.getElementById('aiScopeItem')?.value : ''
    const sprints   = sprint === 'ALL' ? getSprints() : [sprint]

    let summary = ''
    if (type === 'overall') {
      summary = sprints.map(s => {
        const rows = getRows(s, 'ALL')
        const usd = sumC(rows, '#US/Defects'), usp = sumC(rows, '#US/Defects SP'), sub = sumC(rows, '#GenAI SubTask')
        const hrs = sumC(rows, '#GenAI Saved Hours'), ssp = sumC(rows, '#GenAI Saved SP')
        return `Sprint: ${s} | US/Defects: ${Math.round(usd)} | Defects SP: ${Math.round(usp)} | GenAI SubTasks: ${Math.round(sub)} | Hours Saved: ${hrs.toFixed(1)} | Saved SP: ${ssp.toFixed(1)} | Productivity: ${usp > 0 ? (ssp / usp * 100).toFixed(2) : 0}% | Adoption: ${usd > 0 ? (sub / usd * 100).toFixed(2) : 0}%`
      }).join('\n')
    } else {
      const filterT = type === 'tower'   ? scopeItem : 'ALL'
      const filterP = type === 'project' ? scopeItem : 'ALL'
      summary = sprints.map(s => {
        const rows = getRows(s, filterT, filterP)
        const usd = sumC(rows, '#US/Defects'), usp = sumC(rows, '#US/Defects SP'), sub = sumC(rows, '#GenAI SubTask')
        const hrs = sumC(rows, '#GenAI Saved Hours'), ssp = sumC(rows, '#GenAI Saved SP')
        return `Sprint: ${s} | US/Defects: ${Math.round(usd)} | Defects SP: ${Math.round(usp)} | GenAI SubTasks: ${Math.round(sub)} | Hours Saved: ${hrs.toFixed(1)} | Saved SP: ${ssp.toFixed(1)} | Productivity: ${usp > 0 ? (ssp / usp * 100).toFixed(2) : 0}% | Adoption: ${usd > 0 ? (sub / usd * 100).toFixed(2) : 0}%`
      }).join('\n')
    }

    const scopeLabel = type === 'overall' ? 'All Towers' : type === 'tower' ? `Tower: ${scopeItem}` : `Project: ${scopeItem}`
    const prompt = `You are a GenAI sprint analytics expert. Analyse the following sprint data and provide structured insights.\n\nScope: ${scopeLabel}\nSprint range: ${sprint === 'ALL' ? 'All sprints' : sprint}\n\nSprint Data:\n${summary}\n\nProvide a JSON response with this exact structure (no markdown, no backticks, just JSON):\n{\n  "summary": "2-3 sentence executive summary",\n  "goingWell": ["point 1", "point 2", "point 3"],\n  "goingWrong": ["point 1", "point 2"],\n  "trends": "1-2 sentence trend analysis",\n  "recommendations": ["action 1", "action 2", "action 3", "action 4"]\n}`

    if (panel) panel.innerHTML = '<div class="ai-loading"><svg viewBox="0 0 24 24"><path d="M12,4V2A10,10 0 0,0 2,12H4A8,8 0 0,1 12,4Z"/></svg>Analysing sprint data with AI…</div>'
    if (btn)   btn.disabled = true

    try {
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 1000, messages: [{ role: 'user', content: prompt }] })
      })
      const data = await resp.json()
      const text = data.content?.map(b => b.text || '').join('').trim()
      let parsed
      try { parsed = JSON.parse(text) } catch {
        const match = text.match(/\{[\s\S]*\}/)
        parsed = match ? JSON.parse(match[0]) : null
      }
      if (!parsed) throw new Error('Could not parse AI response')
      renderAiResult(parsed, scopeLabel)
      renderAiTrendChart(type, scopeItem, sprints)
    } catch (err) {
      if (panel) panel.innerHTML = `<div style="padding:1.5rem;color:var(--red);font-size:13px;">Error: ${err.message}.<br><br>Make sure you have API access enabled.</div>`
    }
    if (btn) btn.disabled = false
  }
  function renderAiResult(data, scopeLabel) {
    const panel = document.getElementById('aiResultPanel'); if (!panel) return
    const wellChips  = (data.goingWell     || []).map(w => `<span class="ai-action-chip">✓ ${w}</span>`).join('')
    const wrongChips = (data.goingWrong    || []).map(w => `<span class="ai-warn-chip">⚠ ${w}</span>`).join('')
    const recItems   = (data.recommendations || []).map((r, i) => `<div style="display:flex;gap:8px;margin-bottom:6px;"><span style="color:var(--accent);font-family:'DM Mono',monospace;font-size:11px;flex-shrink:0;">${i + 1}.</span><span style="font-size:12px;color:var(--txt2);line-height:1.6;">${r}</span></div>`).join('')
    panel.innerHTML = `
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:1rem;flex-wrap:wrap;gap:6px;">
        <div style="font-family:'Syne',sans-serif;font-size:14px;font-weight:600;color:var(--txt);">AI Analysis — ${scopeLabel}</div>
        <div style="font-size:10px;color:var(--txt3);background:var(--bg3);border:1px solid var(--border);border-radius:4px;padding:2px 8px;">Powered by Claude</div>
      </div>
      <div class="ai-section"><div class="ai-section-title">Summary</div><div class="ai-text">${data.summary || ''}</div></div>
      <div class="ai-section"><div class="ai-section-title">Going well</div><div>${wellChips || '<span style="color:var(--txt3);font-size:12px;">No highlights identified</span>'}</div></div>
      <div class="ai-section"><div class="ai-section-title">Areas of concern</div><div>${wrongChips || '<span style="color:var(--txt3);font-size:12px;">No concerns identified</span>'}</div></div>
      <div class="ai-section"><div class="ai-section-title">Trend analysis</div><div class="ai-text">${data.trends || ''}</div></div>
      <div class="ai-section"><div class="ai-section-title">Recommended actions</div>${recItems || '<span style="color:var(--txt3);font-size:12px;">No recommendations</span>'}</div>`
    const card = document.getElementById('aiTrendCard'); if (card) card.style.display = 'block'
  }
  function renderAiTrendChart(type, scopeItem, sprints) {
    const filterT  = type === 'tower'   ? scopeItem : 'ALL'
    const filterP  = type === 'project' ? scopeItem : 'ALL'
    const prodData = sprints.map(s => parseFloat((prod(getRows(s, filterT, filterP)) * 100).toFixed(2)))
    const adoptData = sprints.map(s => {
      const rows = getRows(s, filterT, filterP)
      const usd  = sumC(rows, '#US/Defects'); const sub = sumC(rows, '#GenAI SubTask')
      return usd > 0 ? parseFloat((sub / usd * 100).toFixed(2)) : 0
    })
    if (chartsRef.current.cAiTrend) chartsRef.current.cAiTrend.destroy()
    const canvas = document.getElementById('chartAiTrend'); if (!canvas) return
    chartsRef.current.cAiTrend = new Chart(canvas, {
      type: 'line',
      data: { labels: sprints, datasets: [
        { label: 'Productivity %', data: prodData,  borderColor: '#39D98A', backgroundColor: 'rgba(57,217,138,0.08)',  borderWidth: 2, pointRadius: 4, pointBackgroundColor: '#39D98A', fill: true, tension: 0.35 },
        { label: 'Adoption %',     data: adoptData, borderColor: '#58A6FF', backgroundColor: 'rgba(88,166,255,0.06)', borderWidth: 2, pointRadius: 4, pointBackgroundColor: '#58A6FF', fill: true, tension: 0.35 }
      ]},
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 16 } },
        plugins: { legend: { display: true, labels: { color: getChartTheme().legendColor, boxWidth: 9, font: { size: 10 } } }, tooltip: { callbacks: { label: v => `${v.dataset.label}: ${v.raw}%` }, ...getChartTheme().tooltip } },
        scales: {
          y: { beginAtZero: true, ticks: { callback: v => v + '%', color: getChartTheme().tickMuted, font: { size: 10 } }, grid: { color: getChartTheme().grid }, border: { color: 'transparent' } },
          x: { ticks: { color: getChartTheme().tickNormal, font: { size: 10 } }, grid: { display: false }, border: { color: 'transparent' } }
        }
      }
    })
  }

  // ══════════════════════════════════════════════════════════════════════════
  // DOWNLOADS / EXPORTS
  // ══════════════════════════════════════════════════════════════════════════
  function downloadChart(id, name) {
    const src = document.getElementById(id); if (!src) return
    const tmp = document.createElement('canvas'); tmp.width = src.width; tmp.height = src.height
    const ctx = tmp.getContext('2d'); ctx.fillStyle = '#161B22'; ctx.fillRect(0, 0, tmp.width, tmp.height); ctx.drawImage(src, 0, 0)
    const a = document.createElement('a'); a.href = tmp.toDataURL('image/png'); a.download = name + '_' + today() + '.png'
    document.body.appendChild(a); a.click(); document.body.removeChild(a)
    showToast('Chart saved', 'success')
  }
  function downloadSprintTablePNG() {
    if (!lastSprintTableRowsRef.current) return
    const { pivot, sprints, towers, sprintTotals } = lastSprintTableRowsRef.current
    const nM = METRICS.length
    const tW = 160, mW = 72; const sGW = nM * mW
    const totalW = tW + sprints.length * sGW
    const hH1 = 34, hH2 = 26, rH = 28
    const totalH = hH1 + hH2 + (towers.length + 1) * rH
    const canvas = document.createElement('canvas'); canvas.width = totalW; canvas.height = totalH
    const ctx = canvas.getContext('2d'); ctx.textBaseline = 'middle'
    ctx.fillStyle = '#161B22'; ctx.fillRect(0, 0, totalW, totalH)
    ctx.fillStyle = '#1C2128'; ctx.fillRect(0, 0, totalW, hH1 + hH2)
    ctx.font = '600 9px Outfit,sans-serif'; ctx.fillStyle = '#6E7681'; ctx.textAlign = 'left'; ctx.fillText('TOWER', 8, hH1 / 2)
    const sc = ['#39D98A', '#58A6FF', '#F0883E', '#FF7EB6', '#BC8CFF', '#E3B341']
    const mc = ['#58A6FF', '#BC8CFF', '#F0883E', '#E3B341', '#39D98A', '#FF7EB6']
    sprints.forEach((s, si) => {
      const gx = tW + si * sGW; ctx.fillStyle = sc[si % sc.length]; ctx.font = '600 10px Syne,sans-serif'; ctx.textAlign = 'center'
      ctx.fillText(s, gx + sGW / 2, hH1 / 2)
      if (si > 0) { ctx.strokeStyle = 'rgba(61,68,77,.9)'; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(gx, 0); ctx.lineTo(gx, hH1 + hH2); ctx.stroke() }
      METRICS.forEach((m, mi) => { ctx.fillStyle = mc[mi]; ctx.font = '500 8px DM Mono,monospace'; ctx.textAlign = 'center'; ctx.fillText(m.label.toUpperCase(), gx + mi * mW + mW / 2, hH1 + hH2 / 2) })
    })
    ctx.strokeStyle = 'rgba(61,68,77,.7)'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(0, hH1 + hH2); ctx.lineTo(totalW, hH1 + hH2); ctx.stroke()
    towers.forEach((t, ti) => {
      const y = hH1 + hH2 + ti * rH
      if (ti % 2 === 1) { ctx.fillStyle = 'rgba(255,255,255,.015)'; ctx.fillRect(0, y, totalW, rH) }
      ctx.font = '400 10px Outfit,sans-serif'; ctx.fillStyle = '#E6EDF3'; ctx.textAlign = 'left'; ctx.fillText(t, 8, y + rH / 2)
      sprints.forEach((s, si) => {
        const v = pivot[t][s]
        METRICS.forEach((m, mi) => { ctx.fillStyle = mc[mi]; ctx.font = '400 9px DM Mono,monospace'; ctx.textAlign = 'right'; ctx.fillText(fmtV(m.key, v), tW + si * sGW + mi * mW + mW - 3, y + rH / 2) })
        if (si < sprints.length - 1) { ctx.strokeStyle = 'rgba(61,68,77,.5)'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(tW + (si + 1) * sGW, y); ctx.lineTo(tW + (si + 1) * sGW, y + rH); ctx.stroke() }
      })
      ctx.strokeStyle = 'rgba(48,54,61,.5)'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(0, y + rH); ctx.lineTo(totalW, y + rH); ctx.stroke()
    })
    const ty = hH1 + hH2 + towers.length * rH
    ctx.fillStyle = 'rgba(57,217,138,.08)'; ctx.fillRect(0, ty, totalW, rH)
    ctx.strokeStyle = 'rgba(57,217,138,.35)'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(0, ty); ctx.lineTo(totalW, ty); ctx.stroke()
    ctx.font = '600 10px Outfit,sans-serif'; ctx.fillStyle = '#39D98A'; ctx.textAlign = 'left'; ctx.fillText('Grand Total', 8, ty + rH / 2)
    sprints.forEach((s, si) => { const v = sprintTotals[s]; METRICS.forEach((m, mi) => { ctx.fillStyle = mc[mi]; ctx.font = '600 9px DM Mono,monospace'; ctx.textAlign = 'right'; ctx.fillText(fmtV(m.key, v), tW + si * sGW + mi * mW + mW - 3, ty + rH / 2) }) })
    const a = document.createElement('a'); a.href = canvas.toDataURL('image/png'); a.download = 'sprint_consolidated_' + today() + '.png'
    document.body.appendChild(a); a.click(); document.body.removeChild(a); showToast('Matrix PNG saved', 'success')
  }
  function exportSprintTableCSV() {
    if (!lastSprintTableRowsRef.current) return
    const { pivot, sprints, towers, sprintTotals } = lastSprintTableRowsRef.current
    const lines = [['Tower', ...sprints.flatMap(s => METRICS.map((_, i) => i === 0 ? `"${s}"` : ''))].join(',')]
    lines.push([''  , ...sprints.flatMap(() => METRICS.map(m => `"${m.key === '__prod__' ? 'Productivity%' : m.key}"`))].join(','))
    towers.forEach(t => { lines.push([`"${t}"`, ...sprints.flatMap(s => METRICS.map(m => csvFmt(m.key, pivot[t][s])))].join(',')) })
    lines.push(['"Grand Total"', ...sprints.flatMap(s => METRICS.map(m => csvFmt(m.key, sprintTotals[s])))].join(','))
    triggerDL(new Blob([lines.join('\n')], { type: 'text/csv' }), 'sprint_consolidated_' + today() + '.csv')
    showToast('CSV downloaded', 'success')
  }
  function csvFmt(key, v) {
    if (key === '__prod__')            return v.pct.toFixed(2) + '%'
    if (key === '#US/Defects')         return Math.round(v.usd)
    if (key === '#US/Defects SP')      return Math.round(v.usp)
    if (key === '#GenAI SubTask')      return Math.round(v.sub)
    if (key === '#GenAI Saved Hours')  return v.hrs.toFixed(1)
    if (key === '#GenAI Saved SP')     return v.ssp.toFixed(1)
    return ''
  }
  function downloadParticipatedPNG() {
    if (!lastParticipatedRef.current.length) { showToast('No data', 'error'); return }
    const lp = lastParticipatedRef.current
    const cols = ['#', 'Tower', '#US/Defects', '#Def SP', '#SubTask', '#Saved Hrs', '#Saved SP', 'Prod%']
    const colW = [28, 140, 85, 72, 72, 82, 72, 72]
    const totalW = colW.reduce((a, b) => a + b, 0)
    const hH = 32; const rH = 28
    const totalH = hH + (lp.length + 1) * rH
    const canvas = document.createElement('canvas'); canvas.width = totalW; canvas.height = totalH
    const ctx = canvas.getContext('2d'); ctx.textBaseline = 'middle'
    ctx.fillStyle = '#161B22'; ctx.fillRect(0, 0, totalW, totalH)
    ctx.fillStyle = '#1C2128'; ctx.fillRect(0, 0, totalW, hH)
    ctx.font = '600 9px Outfit,sans-serif'; ctx.fillStyle = '#6E7681'
    let x = 0; cols.forEach((c, i) => { ctx.textAlign = i < 2 ? 'left' : 'right'; ctx.fillText(c.toUpperCase(), i < 2 ? x + 6 : x + colW[i] - 5, hH / 2); x += colW[i] })
    ctx.strokeStyle = 'rgba(61,68,77,.8)'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(0, hH); ctx.lineTo(totalW, hH); ctx.stroke()
    const cc = ['#6E7681', '#E6EDF3', '#58A6FF', '#BC8CFF', '#F0883E', '#39D98A', '#58A6FF', '#FF7EB6']
    lp.forEach((it, ri) => {
      const y = hH + ri * rH
      if (ri % 2 === 1) { ctx.fillStyle = 'rgba(255,255,255,.015)'; ctx.fillRect(0, y, totalW, rH) }
      const vals = [String(ri + 1), it.tower, Math.round(it.usd || 0).toString(), Math.round(it.usp || 0).toString(), Math.round(it.sub || 0).toString(), it.hours.toFixed(1), it.sp.toFixed(1), (it.pct || 0).toFixed(2) + '%']
      x = 0; vals.forEach((v, i) => { ctx.fillStyle = cc[i]; ctx.font = '400 9px DM Mono,monospace'; ctx.textAlign = i < 2 ? 'left' : 'right'; ctx.fillText(v, i < 2 ? x + 6 : x + colW[i] - 5, y + rH / 2); x += colW[i] })
      ctx.strokeStyle = 'rgba(48,54,61,.5)'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(0, y + rH); ctx.lineTo(totalW, y + rH); ctx.stroke()
    })
    const a = document.createElement('a'); a.href = canvas.toDataURL('image/png'); a.download = 'participated_' + today() + '.png'
    document.body.appendChild(a); a.click(); document.body.removeChild(a); showToast('PNG saved', 'success')
  }
  function exportParticipatedExcel() {
    if (!lastParticipatedRef.current.length) { showToast('No data', 'error'); return }
    const ws = XLSX.utils.aoa_to_sheet([['Tower', '#US/Defects', '#US/Defects SP', '#GenAI SubTask', '#GenAI Saved Hours', '#GenAI Saved SP', 'Productivity%'],
      ...lastParticipatedRef.current.map(it => [it.tower, Math.round(it.usd || 0), Math.round(it.usp || 0), Math.round(it.sub || 0), +(it.hours || 0).toFixed(1), +(it.sp || 0).toFixed(1), (it.pct || 0).toFixed(2) + '%'])])
    ws['!cols'] = [{ wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 14 }]
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Participated')
    XLSX.writeFile(wb, 'participated_' + today() + '.xlsx'); showToast('XLSX downloaded', 'success')
  }
  function exportNotParticipatedCSV() {
    if (!lastNotParticipatedRef.current.length) { showToast('No data', 'error'); return }
    const lines = [['#', 'Tower', 'ACT/PCT/Project', 'US/Defects', 'US/Defects SP'].join(',')]
    lastNotParticipatedRef.current.forEach((it, i) => lines.push([i + 1, `"${it.tower}"`, `"${it.proj}"`, Math.round(it.usd), Math.round(it.usp)].join(',')))
    triggerDL(new Blob([lines.join('\n')], { type: 'text/csv' }), 'not_participated_' + today() + '.csv'); showToast('CSV downloaded', 'success')
  }
  function exportNotParticipatedExcel() {
    if (!lastNotParticipatedRef.current.length) { showToast('No data', 'error'); return }
    const ws = XLSX.utils.aoa_to_sheet([['#', 'Tower', 'ACT/PCT/Project', 'US/Defects', 'US/Defects SP'],
      ...lastNotParticipatedRef.current.map((it, i) => [i + 1, it.tower, it.proj, Math.round(it.usd), Math.round(it.usp)])])
    ws['!cols'] = [{ wch: 5 }, { wch: 20 }, { wch: 42 }, { wch: 14 }, { wch: 14 }]
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Not Participated')
    XLSX.writeFile(wb, 'not_participated_' + today() + '.xlsx'); showToast('XLSX downloaded', 'success')
  }
  function exportMetricsCSV() {
    const tower   = document.getElementById('towerFilter')?.value   || 'ALL'
    const project = document.getElementById('projectFilter')?.value || 'ALL'
    const sprints = getSprints()
    const lines = [['Sprint', 'Tower', 'ACT/PCT/Project', '#US/Defects', '#US/Defects SP', '#GenAI SubTask', '#GenAI Saved Hours', '#GenAI Saved SP', 'Productivity%'].join(',')]
    sprints.forEach(s => getAllTowers().forEach(t => {
      if (tower !== 'ALL' && t !== tower) return
      getAllProjects(t).forEach(pr => {
        if (project !== 'ALL' && pr !== project) return
        const rows = getRows(s, t, pr); if (!rows.length) return
        lines.push([`"${s}"`, `"${t}"`, `"${pr}"`, sumC(rows, '#US/Defects').toFixed(0), sumC(rows, '#US/Defects SP').toFixed(0), sumC(rows, '#GenAI SubTask').toFixed(0), sumC(rows, '#GenAI Saved Hours').toFixed(1), sumC(rows, '#GenAI Saved SP').toFixed(1), (prod(rows) * 100).toFixed(2) + '%'].join(','))
      })
    }))
    triggerDL(new Blob([lines.join('\n')], { type: 'text/csv' }), 'overall_metrics_' + today() + '.csv'); showToast('CSV exported', 'success')
  }
  function exportAdoptionCSV() {
    if (!lastAdoptionDataRef.current.length) { showToast('No data', 'error'); return }
    const lines = [['Tower', '#US/Defects', '#GenAI SubTask', 'Adoption%'].join(',')]
    lastAdoptionDataRef.current.forEach(d => lines.push([`"${d.tower}"`, Math.round(d.usd), Math.round(d.sub), d.pct.toFixed(2) + '%'].join(',')))
    triggerDL(new Blob([lines.join('\n')], { type: 'text/csv' }), 'adoption_' + today() + '.csv'); showToast('CSV downloaded', 'success')
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SPRINT SIDEBAR JSX helper
  // ══════════════════════════════════════════════════════════════════════════
  function sprintSidebarItems(mode) {
    const sprints       = getSprints()
    const currentActive = mode === 'dashboard' ? activeSprint : adoptActiveSprint
    const avgPct        = sprints.length ? sprints.reduce((a, s) => a + prod(getRows(s, 'ALL')) * 100, 0) / sprints.length : 0
    const items = [
      { key: 'ALL', label: 'All Sprints', pct: avgPct },
      ...sprints.map(s => ({ key: s, label: s, pct: prod(getRows(s, 'ALL')) * 100 }))
    ]
    return items.map(({ key, label, pct }) => (
      <button
        key={key}
        className={`sprint-btn${key === currentActive ? ' active' : ''}`}
        onClick={() => {
          if (mode === 'dashboard') { activeSprintRef.current = key; setActiveSprint(key) }
          else                      { adoptActiveSprintRef.current = key; setAdoptActiveSprint(key) }
        }}
      >
        <span className="sprint-dot" />
        {label}
        {mode === 'dashboard' && <span className="sprint-prod">{pct.toFixed(1)}%</span>}
      </button>
    ))
  }

  // ══════════════════════════════════════════════════════════════════════════
  // RENDER
  // ══════════════════════════════════════════════════════════════════════════
  return (
    <div className="sp-root">

      {/* ── Loading ── */}
      {loading && (
        <div className="sp-loading">
          <div className="sp-spinner" />
          <p>Loading sprint data…</p>
        </div>
      )}

      {/* ── Error ── */}
      {!loading && error && (
        <div className="sp-error">
          <div className="sp-error-icon">⚠️</div>
          <h3>Failed to load data</h3>
          <p>{error}</p>
        </div>
      )}

      {/* ── Dashboard ── */}
      {!loading && !error && (
        <>
          {/* Internal tab bar */}
          <div className="sp-tab-bar">
            <button className={`nav-btn${activePage === 'dashboard' ? ' active' : ''}`} onClick={() => setActivePage('dashboard')}>Dashboard</button>
            <button className={`nav-btn${activePage === 'adoption'  ? ' active' : ''}`} onClick={() => setActivePage('adoption')}>GenAI Adoption</button>
            <button className={`nav-btn${activePage === 'ai'        ? ' active' : ''}`} onClick={() => setActivePage('ai')}>AI Insights</button>
            {uploadedAt && <span className="sp-updated">Last updated: {uploadedAt}</span>}
          </div>

          <div className="sp-pages">

            {/* ── DASHBOARD PAGE ── */}
            <div style={{ display: activePage === 'dashboard' ? 'flex' : 'none', flex: 1, overflow: 'hidden', minHeight: 0 }}>
              <div className="sp-dash-layout">
                <aside className="sp-dash-sidebar">
                  <div className="sidebar-label">Sprints</div>
                  {sprintSidebarItems('dashboard')}
                  <div className="sidebar-divider" />
                  <div className="sidebar-label" style={{ marginTop: '6px' }}>Source</div>
                  <div style={{ fontSize: '10px', color: 'var(--txt3)', padding: '0 7px', lineHeight: 1.9 }}>
                    📄 {getSprints().length} sprint{getSprints().length !== 1 ? 's' : ''} loaded
                  </div>
                </aside>
                <main className="sp-main">
                  <div className="sec-heading">Overall Metrics</div>
                  <div className="kpi-strip stagger" id="kpiStrip" />

                  <div className="charts-grid-2">
                    <div className="chart-card fade-in" style={{ marginBottom: 0 }}>
                      <div className="chart-top-row">
                        <div><div className="chart-title">Productivity by sprint</div><div className="chart-subtitle">Σ(GenAI Saved SP) / Σ(US/Defects SP)</div></div>
                        <button className="icon-btn" onClick={() => downloadChart('chart1', 'productivity_by_sprint')}>
                          <svg viewBox="0 0 24 24"><path d="M5,20H19V18H5M19,9H15V3H9V9H5L12,16L19,9Z"/></svg>PNG
                        </button>
                      </div>
                      <div className="chart-wrap" style={{ height: '220px' }}><canvas id="chart1" /></div>
                    </div>
                    <div className="chart-card fade-in" style={{ marginBottom: 0 }}>
                      <div className="chart-top-row">
                        <div><div className="chart-title">Tower-wise productivity</div><div className="chart-subtitle">All towers — selected sprint(s)</div></div>
                        <button className="icon-btn" onClick={() => downloadChart('chart3', 'tower_productivity')}>
                          <svg viewBox="0 0 24 24"><path d="M5,20H19V18H5M19,9H15V3H9V9H5L12,16L19,9Z"/></svg>PNG
                        </button>
                      </div>
                      <div className="chart-wrap" id="chart3wrap" style={{ height: '220px' }}><canvas id="chart3" /></div>
                    </div>
                  </div>

                  <div className="chart-card fade-in">
                    <div className="chart-top-row">
                      <div><div className="chart-title">Overall metrics comparison by sprint</div><div className="chart-subtitle">All columns across sprints — filter by tower &amp; project</div></div>
                      <div className="chart-actions">
                        <select id="towerFilter" onChange={e => { updateProjectDropdown('projectFilter', e.target.value); renderChart2(); renderChart3(); renderParticipated() }}><option value="ALL">All towers</option></select>
                        <select id="projectFilter" onChange={() => { renderChart2(); renderChart3(); renderParticipated() }}><option value="ALL">All projects</option></select>
                        <button className="icon-btn" onClick={() => downloadChart('chart2', 'overall_metrics')}><svg viewBox="0 0 24 24"><path d="M5,20H19V18H5M19,9H15V3H9V9H5L12,16L19,9Z"/></svg>PNG</button>
                        <button className="icon-btn" onClick={exportMetricsCSV}><svg viewBox="0 0 24 24"><path d="M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M18,20H6V4H13V9H18V20M10,19L6,15H9V11H11V15H14L10,19Z"/></svg>CSV</button>
                      </div>
                    </div>
                    <div id="legend2" className="legend-strip" />
                    <div className="chart-wrap" style={{ height: '260px' }}><canvas id="chart2" /></div>
                  </div>

                  <div className="chart-card fade-in">
                    <div className="chart-top-row">
                      <div><div className="chart-title">Sprint consolidated values</div><div className="chart-subtitle">Towers as rows · sprints as column groups · click a tower to drill down to projects</div></div>
                      <div className="chart-actions">
                        <button className="icon-btn" onClick={downloadSprintTablePNG}><svg viewBox="0 0 24 24"><path d="M5,20H19V18H5M19,9H15V3H9V9H5L12,16L19,9Z"/></svg>PNG</button>
                        <button className="icon-btn" onClick={exportSprintTableCSV}><svg viewBox="0 0 24 24"><path d="M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M18,20H6V4H13V9H18V20M10,19L6,15H9V11H11V15H14L10,19Z"/></svg>CSV</button>
                      </div>
                    </div>
                    <div className="data-table-wrap" id="sprintMatrixWrap"><p className="empty">No data</p></div>
                  </div>

                  <div className="filter-bar" style={{ marginTop: '4px' }}>
                    <label>Filter:</label>
                    <select id="partTowerFilter" onChange={e => { updateProjectDropdown('partProjectFilter', e.target.value); renderParticipated() }}><option value="ALL">All towers</option></select>
                    <select id="partProjectFilter" onChange={renderParticipated}><option value="ALL">All projects</option></select>
                  </div>

                  <div className="chart-card fade-in">
                    <div className="section-header">
                      <div><div className="chart-title">Participated ACT/PCT/Project</div><div className="chart-subtitle">GenAI Saved Hours &gt; 0 · click a tower row to drill to projects</div></div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '5px', flexWrap: 'wrap' }}>
                        <span className="section-count" id="partCount">0</span>
                        <button className="icon-btn" onClick={downloadParticipatedPNG}><svg viewBox="0 0 24 24"><path d="M5,20H19V18H5M19,9H15V3H9V9H5L12,16L19,9Z"/></svg>PNG</button>
                        <button className="icon-btn" onClick={exportParticipatedExcel}><svg viewBox="0 0 24 24"><path d="M21.17 3.25Q21.5 3.25 21.76 3.5 22 3.74 22 4.08V19.92Q22 20.26 21.76 20.5 21.5 20.75 21.17 20.75H7.83Q7.5 20.75 7.24 20.5 7 20.26 7 19.92V17H2.83Q2.5 17 2.24 16.76 2 16.5 2 16.17V7.83Q2 7.5 2.24 7.24 2.5 7 2.83 7H7V4.08Q7 3.74 7.24 3.5 7.5 3.25 7.83 3.25M7 13.06L8.18 15.28H9.97L8 12.06 9.97 8.84H8.18L7 11.06 6.2 9.56 5.86 8.84H4.07L6 12.06 4.07 15.28H5.86M13.88 19.25V17H8.25V19.25M13.88 15.5V12.63H12V15.5M13.88 11.13V8.25H12V11.13M13.88 6.75V4.75H8.25V6.75M20.75 19.25V17H15.13V19.25M20.75 15.5V12.63H15.13V15.5M20.75 11.13V8.25H15.13V11.13M20.75 6.75V4.75H15.13V6.75Z"/></svg>XLSX</button>
                      </div>
                    </div>
                    <div className="part-table-wrap">
                      <table className="part-table" id="participatedTable">
                        <thead><tr>
                          <th style={{ width: '28px' }}>#</th><th>Tower / Project</th>
                          <th style={{ textAlign: 'right' }}>#US/Defects</th><th style={{ textAlign: 'right' }}>#Def SP</th>
                          <th style={{ textAlign: 'right' }}>#SubTask</th><th style={{ textAlign: 'right' }}>#Saved Hrs</th>
                          <th style={{ textAlign: 'right' }}>#Saved SP</th><th style={{ textAlign: 'right' }}>Productivity%</th>
                        </tr></thead>
                        <tbody id="participatedBody"><tr><td colSpan={8} className="empty">No data</td></tr></tbody>
                      </table>
                    </div>
                  </div>

                  <div className="chart-card fade-in" style={{ marginBottom: 0 }}>
                    <div className="section-header">
                      <div>
                        <div className="chart-title">Not Participated ACT/PCT/Project <span className="not-part-label" id="notPartLabel" /></div>
                        <div className="chart-subtitle">Projects with no GenAI Saved Hours</div>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '5px', flexWrap: 'wrap' }}>
                        <span className="section-count" id="notPartCount">0</span>
                        <button className="icon-btn" onClick={exportNotParticipatedCSV}><svg viewBox="0 0 24 24"><path d="M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M18,20H6V4H13V9H18V20M10,19L6,15H9V11H11V15H14L10,19Z"/></svg>CSV</button>
                        <button className="icon-btn" onClick={exportNotParticipatedExcel}><svg viewBox="0 0 24 24"><path d="M21.17 3.25Q21.5 3.25 21.76 3.5 22 3.74 22 4.08V19.92Q22 20.26 21.76 20.5 21.5 20.75 21.17 20.75H7.83Q7.5 20.75 7.24 20.5 7 20.26 7 19.92V17H2.83Q2.5 17 2.24 16.76 2 16.5 2 16.17V7.83Q2 7.5 2.24 7.24 2.5 7 2.83 7H7V4.08Q7 3.74 7.24 3.5 7.5 3.25 7.83 3.25M7 13.06L8.18 15.28H9.97L8 12.06 9.97 8.84H8.18L7 11.06 6.2 9.56 5.86 8.84H4.07L6 12.06 4.07 15.28H5.86M13.88 19.25V17H8.25V19.25M13.88 15.5V12.63H12V15.5M13.88 11.13V8.25H12V11.13M13.88 6.75V4.75H8.25V6.75M20.75 19.25V17H15.13V19.25M20.75 15.5V12.63H15.13V15.5M20.75 11.13V8.25H15.13V11.13M20.75 6.75V4.75H15.13V6.75Z"/></svg>XLSX</button>
                      </div>
                    </div>
                    <div className="part-table-wrap">
                      <table className="part-table">
                        <thead><tr>
                          <th style={{ width: '28px' }}>#</th><th>Tower</th><th>ACT/PCT/Project</th>
                          <th style={{ textAlign: 'right' }}>US/Defects</th><th style={{ textAlign: 'right' }}>Defects SP</th>
                        </tr></thead>
                        <tbody id="notParticipatedBody"><tr><td colSpan={5} className="empty">No data</td></tr></tbody>
                      </table>
                    </div>
                  </div>
                </main>
              </div>
            </div>

            {/* ── ADOPTION PAGE ── */}
            <div style={{ display: activePage === 'adoption' ? 'flex' : 'none', flex: 1, overflow: 'hidden', minHeight: 0 }}>
              <div className="sp-dash-layout">
                <aside className="sp-dash-sidebar">
                  <div className="sidebar-label">Sprints</div>
                  {sprintSidebarItems('adoption')}
                </aside>
                <main className="sp-main">
                  <div className="chart-card fade-in">
                    <div className="chart-top-row">
                      <div><div className="chart-title">GenAI SubTask Adoption by Tower</div><div className="chart-subtitle">Adoption % = (#GenAI SubTask / #US/Defects) × 100 · click tower to drill into projects</div></div>
                      <div className="chart-actions">
                        <button className="icon-btn" onClick={exportAdoptionCSV}><svg viewBox="0 0 24 24"><path d="M14,2H6A2,2 0 0,0 4,4V20A2,2 0 0,0 6,22H18A2,2 0 0,0 20,20V8L14,2M18,20H6V4H13V9H18V20M10,19L6,15H9V11H11V15H14L10,19Z"/></svg>CSV</button>
                      </div>
                    </div>
                    <div className="data-table-wrap">
                      <table className="adoption-table" id="adoptionTable">
                        <thead><tr>
                          <th>Tower / Project</th><th style={{ textAlign: 'right' }}>#US/Defects</th>
                          <th style={{ textAlign: 'right' }}>#GenAI SubTask</th><th style={{ textAlign: 'right' }}>Adoption %</th>
                          <th style={{ minWidth: '120px' }}>Visual</th>
                        </tr></thead>
                        <tbody id="adoptionBody"><tr><td colSpan={5} className="empty">No data</td></tr></tbody>
                      </table>
                    </div>
                  </div>
                  <div className="chart-card fade-in" style={{ marginBottom: 0 }}>
                    <div className="chart-top-row">
                      <div><div className="chart-title">Adoption trend by sprint</div><div className="chart-subtitle">GenAI SubTask adoption % per tower per sprint</div></div>
                      <div className="chart-actions">
                        <select id="adoptTowerFilter" onChange={renderAdoptionTrend}><option value="ALL">All towers</option></select>
                        <button className="icon-btn" onClick={() => downloadChart('chartAdopt', 'adoption_trend')}><svg viewBox="0 0 24 24"><path d="M5,20H19V18H5M19,9H15V3H9V9H5L12,16L19,9Z"/></svg>PNG</button>
                      </div>
                    </div>
                    <div className="chart-wrap" style={{ height: '260px' }}><canvas id="chartAdopt" /></div>
                  </div>
                </main>
              </div>
            </div>

            {/* ── AI INSIGHTS PAGE ── */}
            <div style={{ display: activePage === 'ai' ? 'flex' : 'none', flex: 1, overflow: 'auto', minHeight: 0 }}>
              <div style={{ padding: '1.5rem', flex: 1 }}>
                <div className="ai-layout">
                  <div className="ai-panel">
                    <div className="sec-heading" style={{ marginBottom: '.8rem' }}>AI Inference Settings</div>
                    <div className="ai-controls">
                      <div>
                        <div className="ai-scope-label">Scope</div>
                        <select id="aiScopeType" style={{ width: '100%' }} onChange={updateAiScope}>
                          <option value="overall">Overall (all towers)</option>
                          <option value="tower">By Tower</option>
                          <option value="project">By ACT/PCT/Project</option>
                        </select>
                      </div>
                      <div id="aiScopeItemWrap" style={{ display: 'none' }}>
                        <div className="ai-scope-label" id="aiScopeItemLabel">Tower</div>
                        <select id="aiScopeItem" style={{ width: '100%' }} />
                      </div>
                      <div>
                        <div className="ai-scope-label">Sprint focus</div>
                        <select id="aiSprintScope" style={{ width: '100%' }}><option value="ALL">All sprints</option></select>
                      </div>
                      <div style={{ background: 'var(--bg3)', borderRadius: '8px', padding: '10px', fontSize: '11px', color: 'var(--txt3)', lineHeight: 1.6 }}>
                        AI will analyse productivity trends, GenAI adoption rates, hours saved, and recommend actions based on your data.
                      </div>
                      <button className="ai-run-btn" id="aiRunBtn" onClick={runAiInference}>
                        <svg viewBox="0 0 24 24"><path d="M12,2A10,10 0 0,1 22,12A10,10 0 0,1 12,22A10,10 0 0,1 2,12A10,10 0 0,1 12,2M12,4A8,8 0 0,0 4,12A8,8 0 0,0 12,20A8,8 0 0,0 20,12A8,8 0 0,0 12,4M11,8H13V12.41L15.29,14.71L13.88,16.12L11,13.24V8Z"/></svg>
                        Run AI Analysis
                      </button>
                    </div>
                  </div>
                  <div className="ai-result">
                    <div className="ai-result-inner" id="aiResultPanel">
                      <div className="ai-empty">Select a scope and click "Run AI Analysis" to get AI-powered insights about your sprint data.</div>
                    </div>
                    <div className="chart-card fade-in" style={{ marginTop: '14px', display: 'none' }} id="aiTrendCard">
                      <div className="chart-top-row">
                        <div><div className="chart-title">Productivity trend</div><div className="chart-subtitle">Sprint-over-sprint for selected scope</div></div>
                      </div>
                      <div className="chart-wrap" style={{ height: '220px' }}><canvas id="chartAiTrend" /></div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

          </div>{/* end sp-pages */}
        </>
      )}

      <div className="sp-toast" id="sp-toast" />
    </div>
  )
}
