import { useState, useEffect, useRef } from 'react'
import { Chart, registerables } from 'chart.js'
import * as XLSX from 'xlsx'
import './CopilotVsGit.css'
import { useTheme } from '../context/ThemeContext'

Chart.register(...registerables)

const STORAGE_URL = import.meta.env.VITE_STORAGE_URL || ''

const columnMap = {
  assocId:     'Associate ID',
  name:        'Associate Name',
  finalStatus: 'Final Status',
  tower:       'Tower',
  hasLicense:  'Has A Copilot License',
  linesAdded:  'Lines Added',
  linesDeleted:'Lines Deleted',
  addedLines:  'Added Lines',
  deletedLines:'Deleted Lines',
  interactions:'Copilot Interactions',
  usage:       'Copilot Usage',
  gitCommits:  'Gitlab Commits',
  gitlabActive:'Active Gitlab Account',
  actPct:      'ACT/PCT Mapping',
}

function col(row, key)     { return row[columnMap[key]] }
function g(row, key)       { return columnMap[key] ? row[columnMap[key]] : undefined }
function gn(row, key)      { return parseFloat(g(row, key)) || 0 }
function gs(row, key, def) { return columnMap[key] ? String(row[columnMap[key]] || def).trim() : def }
function avg(arr)          { return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0 }
function f2(v)             { return isNaN(v) || v === null || v === undefined ? '0.00' : Number(v).toFixed(2) }
function isYes(v)          { return ['yes','true','1','y'].includes(String(v || '').toLowerCase().trim()) }

function dlCSV(rows, fn) {
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n')
  const a = document.createElement('a')
  a.href = 'data:text/csv;charset=utf-8,\uFEFF' + encodeURIComponent(csv)
  a.download = fn; a.click()
}

const DISP_CATS = [
  'No Copilot License',
  'Low user (Interaction 0-5)',
  'Average user (Interaction 6-50 and Copilot Added Lines <100)',
  'Average user (Interaction 6-50 and Copilot Added Lines 100-500)',
  'Average user (Interaction 6-50 and Copilot Added Lines >500)',
  'Power user (Interaction >50 and Copilot Added Lines <100)',
  'Power user (Interaction >50 and Copilot Added Lines 100-500)',
  'Power user (Interaction >50 and Copilot Added Lines >500)',
  'No match',
]
const DISP_CSS = ['cn','cl','ca','ca','ca','cp','cp','cp','cm2']

function dispCat(r) {
  const l = r.license.toLowerCase()
  const hasLic = l==='yes'||l==='true'||l==='1'||l==='y'
  if (!hasLic) return 'No Copilot License'
  const ci = r.interactions, la = r.linesAdded
  if (ci >= 0 && ci <= 5)                    return 'Low user (Interaction 0-5)'
  if (ci >= 6 && ci <= 50 && la < 100)       return 'Average user (Interaction 6-50 and Copilot Added Lines <100)'
  if (ci >= 6 && ci <= 50 && la <= 500)      return 'Average user (Interaction 6-50 and Copilot Added Lines 100-500)'
  if (ci >= 6 && ci <= 50 && la > 500)       return 'Average user (Interaction 6-50 and Copilot Added Lines >500)'
  if (ci > 50 && la < 100)                   return 'Power user (Interaction >50 and Copilot Added Lines <100)'
  if (ci > 50 && la <= 500)                  return 'Power user (Interaction >50 and Copilot Added Lines 100-500)'
  if (ci > 50 && la > 500)                   return 'Power user (Interaction >50 and Copilot Added Lines >500)'
  return 'No match'
}

export default function CopilotVsGit() {
  const { theme } = useTheme()

  function getChartTheme() {
    const dark = theme === 'dark'
    return {
      grid:       dark ? '#30363D' : 'rgba(0,0,0,0.08)',
      tickColor:  dark ? '#8B949E' : '#5A6478',
      legendColor: dark ? '#8B949E' : '#5A6478',
      tooltip: {
        backgroundColor: dark ? '#161B22' : '#FFFFFF',
        borderColor:     dark ? '#30363D' : '#D0D7E2',
        borderWidth: 1,
        titleColor:      dark ? '#E6EDF3' : '#1A1F2E',
        bodyColor:       dark ? '#8B949E' : '#5A6478',
      },
    }
  }

  // ── FETCH STATE ────────────────────────────────────────────────
  const [fetchState, setFetchState]   = useState('loading')
  const [fetchError, setFetchError]   = useState('')
  const [allData, setAllData]         = useState([])

  // ── ANALYSIS STATE ─────────────────────────────────────────────
  const [analyzedData, setAnalyzedData] = useState([])
  const [metrics, setMetrics] = useState({
    total:'—', final:'—', finalSub:'at threshold',
    added:'—', addedSub:'at threshold', both:'—', bothSub:'immediate review'
  })
  const [currentPage, setCurrentPage] = useState('p-analysis')
  const [currentTab,  setCurrentTab]  = useState('all')
  const [pctThreshold,   setPctThreshold]   = useState(60)
  const [filterStatus,   setFilterStatus]   = useState('Active')
  const [filterLicense,  setFilterLicense]  = useState('Yes')
  const [searchInput,    setSearchInput]    = useState('')

  // ── TOWER STATE ────────────────────────────────────────────────
  const [towerRows,        setTowerRows]        = useState([])
  const [towerGrandTotal,  setTowerGrandTotal]  = useState(null)
  const [towerDisp,        setTowerDisp]        = useState(null)
  const [towerEmpty,       setTowerEmpty]       = useState('Load data to see tower summaries')
  const [towerInsights,    setTowerInsights]    = useState(null)
  const [twStatus,  setTwStatus]  = useState('All')
  const [twLicense, setTwLicense] = useState('All')
  const [twGitlab,  setTwGitlab]  = useState('All')
  const [twStatusOpts,  setTwStatusOpts]  = useState([])
  const [twGitlabOpts,  setTwGitlabOpts]  = useState([])

  // ── METRICS STATE ──────────────────────────────────────────────
  const [metricsRows,       setMetricsRows]       = useState([])
  const [metricsGrandTotal, setMetricsGrandTotal] = useState(null)
  const [metricsDisp,       setMetricsDisp]       = useState(null)
  const [metricsEmpty,      setMetricsEmpty]      = useState('Load data to see metrics')
  const [metricsInsights,   setMetricsInsights]   = useState(null)
  const [tmStatus, setTmStatus] = useState('All')
  const [tmGitlab, setTmGitlab] = useState('All')
  const [tmStatusOpts, setTmStatusOpts] = useState([])
  const [tmGitlabOpts, setTmGitlabOpts] = useState([])

  // ── UI STATE ───────────────────────────────────────────────────
  const [showConfirm, setShowConfirm] = useState(false)
  const [showEmail,   setShowEmail]   = useState(false)
  const [emailStep,   setEmailStep]   = useState(1)
  const [emailSubject, setEmailSubject] = useState('[CTS] Your Copilot vs Git Productivity Report')
  const [emailBodyTxt, setEmailBodyTxt] = useState(
    'Dear Associate,\n\nPlease find your Copilot utilisation metrics compared to your Git contribution data for the current review period.\n\nRegards,\nCTS Productivity Analytics Team'
  )
  const [sendProgress, setSendProgress] = useState({ label:'', pct:0 })
  const [sendLog,      setSendLog]      = useState('')
  const [emailStatus,  setEmailStatus]  = useState({ visible:false, msg:'', type:'' })
  const [msalSignedIn, setMsalSignedIn] = useState(false)
  const [msalUsername, setMsalUsername] = useState('')
  const [toasts,       setToasts]       = useState([])

  // ── DRILL STATE ────────────────────────────────────────────────
  const [drillOpen,     setDrillOpen]     = useState(false)
  const [drillTitle,    setDrillTitle]    = useState('')
  const [drillSubtitle, setDrillSubtitle] = useState('')
  const [drillRows,     setDrillRows]     = useState([])
  const [drillCols,     setDrillCols]     = useState([])
  const [drillIsDisp,   setDrillIsDisp]   = useState(false)
  const [drillSort,     setDrillSort]     = useState({ col:null, asc:true })

  // ── REFS ───────────────────────────────────────────────────────
  const towerRawRef   = useRef({})
  const metricsRawRef = useRef({})
  const drillBaseRef  = useRef([])   // unsorted rows for sort reset
  const chartFinalRef = useRef(null)
  const chartAddedRef = useRef(null)
  const chartBreachRef= useRef(null)
  const chartRatioRef = useRef(null)
  const canvasFinalRef = useRef(null)
  const canvasAddedRef = useRef(null)
  const canvasBreachRef= useRef(null)
  const canvasRatioRef = useRef(null)
  const msalInstRef    = useRef(null)
  const msalAccountRef = useRef(null)
  const toastIdRef     = useRef(0)

  // ── TOAST ──────────────────────────────────────────────────────
  function toast(msg, type = 'info', dur = 3500) {
    const id = ++toastIdRef.current
    setToasts(prev => [...prev, { id, msg, type }])
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), dur)
  }

  // ── FETCH ON MOUNT ─────────────────────────────────────────────
  useEffect(() => { fetchData() }, [])

  async function fetchData() {
    setFetchState('loading'); setFetchError('')
    try {
      const cfgRes = await fetch(`${STORAGE_URL}/copilot-git-analytics/config.json`)
      if (!cfgRes.ok) throw new Error(`Config fetch failed: ${cfgRes.status}`)
      const cfg = await cfgRes.json()
      const xlsxRes = await fetch(`${STORAGE_URL}/copilot-git-analytics/${cfg.latestFile}`)
      if (!xlsxRes.ok) throw new Error(`File fetch failed: ${xlsxRes.status}`)
      const buffer = await xlsxRes.arrayBuffer()
      const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' })
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' })
      if (!rows.length) throw new Error('No data rows found in Excel file')
      setAllData(rows)
      setFetchState('done')
      toast(`Loaded ${rows.length} rows`, 'success')
    } catch (e) {
      setFetchError(e.message); setFetchState('error')
    }
  }

  // ── POPULATE DROPDOWNS WHEN DATA LOADS ────────────────────────
  useEffect(() => {
    if (!allData.length) return
    function getUnique(key) {
      if (!columnMap[key]) return []
      const vals = new Set()
      allData.forEach(row => { const v = String(row[columnMap[key]] || '').trim(); if (v) vals.add(v) })
      return [...vals].sort()
    }
    const statuses = getUnique('finalStatus')
    const gitlabs  = getUnique('gitlabActive')
    setTwStatusOpts(statuses); setTwGitlabOpts(gitlabs)
    setTmStatusOpts(statuses); setTmGitlabOpts(gitlabs)
  }, [allData])

  // ── ANALYSIS EFFECT ────────────────────────────────────────────
  useEffect(() => {
    if (!allData.length) return
    doRunAnalysis(allData, pctThreshold, filterStatus, filterLicense)
  }, [allData, pctThreshold, filterStatus, filterLicense])

  // ── TOWER EFFECT ───────────────────────────────────────────────
  useEffect(() => {
    if (!allData.length) return
    doRenderTower(allData, twStatus, twLicense, twGitlab)
  }, [allData, twStatus, twLicense, twGitlab])

  // ── METRICS EFFECT ─────────────────────────────────────────────
  useEffect(() => {
    if (!allData.length) return
    doRenderMetrics(allData, tmStatus, tmGitlab)
  }, [allData, tmStatus, tmGitlab])

  // ── CHARTS EFFECT ──────────────────────────────────────────────
  useEffect(() => {
    if (!analyzedData.length) return
    updateCharts(analyzedData, pctThreshold)
    return () => destroyCharts()
  }, [analyzedData, pctThreshold, theme]) // eslint-disable-line

  // ── ANALYSIS ──────────────────────────────────────────────────
  function doRunAnalysis(data, pct, fs, fl) {
    let src = data.map(row => {
      const assocId = String(col(row,'assocId') || '').trim(); if (!assocId) return null
      const name    = String(col(row,'name')    || '').trim()
      const la = parseFloat(col(row,'linesAdded'))  || 0
      const ld = parseFloat(col(row,'linesDeleted'))|| 0
      const al = parseFloat(col(row,'addedLines'))  || 0
      const dl = parseFloat(col(row,'deletedLines'))|| 0
      const status  = String(col(row,'finalStatus') || 'Active').trim()
      const license = String(col(row,'hasLicense')  || 'Yes').trim()
      const fCop=la-ld, fGit=al-dl, aCop=la, aGit=al
      const fR = fGit!==0 ? (fCop/fGit)*100 : null
      const aR = aGit!==0 ? (aCop/aGit)*100 : null
      const ffinal = fR!==null && fR<pct
      const fadded = aR!==null && aR<pct
      return { assocId, name, fCop, fGit, aCop, aGit, fR, aR, ffinal, fadded, fboth:ffinal&&fadded, status, license }
    }).filter(Boolean)

    if (fs !== 'All') src = src.filter(r => r.status.toLowerCase() === fs.toLowerCase())
    if (fl !== 'All') {
      src = src.filter(r => {
        const l=r.license.toLowerCase(), y=l==='yes'||l==='true'||l==='1'||l==='y'
        return fl==='Yes' ? y : !y
      })
    }
    const t=src.length, ff=src.filter(r=>r.ffinal).length
    const fa=src.filter(r=>r.fadded).length, fb=src.filter(r=>r.fboth).length
    setMetrics({
      total:t, final:ff, finalSub:`${pct}% threshold — ${t?((ff/t)*100).toFixed(2):0}% flagged`,
      added:fa, addedSub:`${pct}% threshold — ${t?((fa/t)*100).toFixed(2):0}% flagged`,
      both:fb, bothSub:`${t?((fb/t)*100).toFixed(2):0}% of associates`,
    })
    setAnalyzedData(src)
  }

  // ── CHARTS ────────────────────────────────────────────────────
  function destroyCharts() {
    [chartFinalRef,chartAddedRef,chartBreachRef,chartRatioRef].forEach(ref => {
      try { if (ref.current) { ref.current.destroy(); ref.current=null } } catch {}
    })
  }
  function updateCharts(data, pct) {
    destroyCharts()
    if (!canvasFinalRef.current) return
    const ct = getChartTheme()
    const top=data.slice(0,20), lb=top.map(r=>r.assocId||r.name||'?')
    const grid = { color: ct.grid }
    const co = {
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ position:'top', labels:{ font:{size:11}, color:ct.legendColor } }, tooltip:{ ...ct.tooltip } },
      scales:{ x:{ ticks:{color:ct.tickColor}, grid }, y:{ ticks:{color:ct.tickColor}, grid } }
    }
    chartFinalRef.current = new Chart(canvasFinalRef.current, { type:'bar', data:{ labels:lb, datasets:[
      { label:'Final Copilot', data:top.map(r=>r.fCop), backgroundColor:'#58a6ff', borderRadius:3 },
      { label:'Final Git',     data:top.map(r=>r.fGit), backgroundColor:'#3fb950', borderRadius:3 },
    ]}, options:co })
    chartAddedRef.current = new Chart(canvasAddedRef.current, { type:'bar', data:{ labels:lb, datasets:[
      { label:'Added Copilot', data:top.map(r=>r.aCop), backgroundColor:'#0097a7', borderRadius:3 },
      { label:'Added Git',     data:top.map(r=>r.aGit), backgroundColor:'#e3b341', borderRadius:3 },
    ]}, options:co })
    const tot=data.length, ff=data.filter(r=>r.ffinal).length, fa=data.filter(r=>r.fadded).length
    const coStacked = { ...co, scales:{ x:{...co.scales.x,stacked:true}, y:{...co.scales.y,stacked:true} } }
    chartBreachRef.current = new Chart(canvasBreachRef.current, { type:'bar', data:{ labels:['Final','Added'], datasets:[
      { label:`Below ${pct}%`, data:[ff,fa],         backgroundColor:'#f85149', borderRadius:4 },
      { label:'Within',        data:[tot-ff,tot-fa], backgroundColor:'#3fb950', borderRadius:4 },
    ]}, options:coStacked })
    const sc=data.filter(r=>r.fR!==null).map((r,i)=>({x:i+1,y:parseFloat(r.fR.toFixed(2)),f:r.ffinal}))
    chartRatioRef.current = new Chart(canvasRatioRef.current, { type:'scatter', data:{ datasets:[
      { label:'Copilot%', data:sc, backgroundColor:sc.map(d=>d.f?'#f85149':'#3fb950'), pointRadius:5 },
      { label:`Threshold (${pct}%)`, type:'line', data:[{x:0,y:pct},{x:sc.length+1,y:pct}],
        borderColor:'#e3b341', borderDash:[5,5], pointRadius:0, borderWidth:2 },
    ]}, options:{ ...co, scales:{
      x:{ ...co.scales.x, title:{display:true,text:'Index',color:ct.tickColor} },
      y:{ ...co.scales.y, min:0, title:{display:true,text:'Copilot%',color:ct.tickColor} }
    }} })
  }

  // ── TOWER ─────────────────────────────────────────────────────
  function doRenderTower(data, fs, fl, fg) {
    const filtered = data.map(row => {
      const assocId = String(col(row,'assocId')||'').trim(); if (!assocId) return null
      const name=gs(row,'name',''), actPct=gs(row,'actPct','')
      const status=gs(row,'finalStatus','Active'), license=gs(row,'hasLicense','Yes')
      const gitlabActive=gs(row,'gitlabActive','Yes'), tower=gs(row,'tower','Unknown')||'Unknown'
      const interactions=gn(row,'interactions'), usage=gn(row,'usage')
      const linesAdded=gn(row,'linesAdded'), linesDeleted=gn(row,'linesDeleted')
      const addedLines=gn(row,'addedLines'), deletedLines=gn(row,'deletedLines')
      const gitCommits=gn(row,'gitCommits')
      return { assocId,name,actPct,status,license,gitlabActive,tower,
               interactions,usage,linesAdded,linesDeleted,addedLines,deletedLines,gitCommits,
               copilotFinal:linesAdded-linesDeleted, gitFinal:addedLines-deletedLines }
    }).filter(r => {
      if (!r) return false
      if (fs!=='All' && r.status!==fs) return false
      if (fl!=='All') {
        const y=['yes','true','1','y'].includes(r.license.toLowerCase())
        if (fl==='Yes'&&!y) return false; if (fl==='No'&&y) return false
      }
      if (fg!=='All' && r.gitlabActive!==fg) return false
      return true
    })
    if (!filtered.length) {
      setTowerEmpty('No records match the current filters — try setting filters to "All"')
      setTowerRows([]); setTowerGrandTotal(null); setTowerDisp(null); return
    }
    const byTower={}
    filtered.forEach(r=>{ if (!byTower[r.tower]) byTower[r.tower]=[]; byTower[r.tower].push(r) })
    const towers=Object.keys(byTower).sort()
    towerRawRef.current = byTower
    const rows = towers.map(tw=>{ const g=byTower[tw]; return {
      tower:tw, count:g.length,
      avgInt:avg(g.map(r=>r.interactions)), avgUse:avg(g.map(r=>r.usage)),
      avgCopAdd:avg(g.map(r=>r.linesAdded)), avgCopFin:avg(g.map(r=>r.copilotFinal)),
      avgCommit:avg(g.map(r=>r.gitCommits)), avgGitAdd:avg(g.map(r=>r.addedLines)),
      avgGitFin:avg(g.map(r=>r.gitFinal)),
    }})
    const gt={ tower:'Grand Total', count:filtered.length,
      avgInt:avg(filtered.map(r=>r.interactions)), avgUse:avg(filtered.map(r=>r.usage)),
      avgCopAdd:avg(filtered.map(r=>r.linesAdded)), avgCopFin:avg(filtered.map(r=>r.copilotFinal)),
      avgCommit:avg(filtered.map(r=>r.gitCommits)), avgGitAdd:avg(filtered.map(r=>r.addedLines)),
      avgGitFin:avg(filtered.map(r=>r.gitFinal)),
    }
    setTowerEmpty(''); setTowerRows(rows); setTowerGrandTotal(gt)
    const tDisp={}
    towers.forEach(tw=>{ tDisp[tw]={}; byTower[tw].forEach(r=>{ const c=dispCat(r); tDisp[tw][c]=(tDisp[tw][c]||0)+1 }) })
    setTowerDisp({ byTower:tDisp, towers, rawByTower:byTower })
    generateInsights('tower', rows, gt, setTowerInsights)
  }

  // ── METRICS ───────────────────────────────────────────────────
  function doRenderMetrics(data, fs, fg) {
    const parsed = data.map(row => {
      const assocId=String(col(row,'assocId')||'').trim(); if (!assocId) return null
      const name=gs(row,'name',''), actPct=gs(row,'actPct','')
      const status=gs(row,'finalStatus','Active'), license=gs(row,'hasLicense','Yes')
      const gitlabActive=gs(row,'gitlabActive','Yes'), tower=gs(row,'tower','Unknown')||'Unknown'
      const interactions=gn(row,'interactions'), usage=gn(row,'usage')
      const linesAdded=gn(row,'linesAdded'), linesDeleted=gn(row,'linesDeleted')
      const addedLines=gn(row,'addedLines'), deletedLines=gn(row,'deletedLines')
      const gitCommits=gn(row,'gitCommits'), hasLic=isYes(license)
      return { assocId,name,actPct,status,license,gitlabActive,tower,hasLic,
               interactions,usage,linesAdded,linesDeleted,addedLines,deletedLines,gitCommits,
               copilotFinal:linesAdded-linesDeleted, gitFinal:addedLines-deletedLines }
    }).filter(r => {
      if (!r) return false
      if (fs!=='All' && r.status!==fs) return false
      if (fg!=='All' && r.gitlabActive!==fg) return false
      return true
    })
    if (!parsed.length) {
      setMetricsEmpty('No records match the current filters — try setting filters to "All"')
      setMetricsRows([]); setMetricsGrandTotal(null); setMetricsDisp(null); return
    }
    const byTower={}
    parsed.forEach(r=>{ if (!byTower[r.tower]) byTower[r.tower]=[]; byTower[r.tower].push(r) })
    const towers=Object.keys(byTower).sort()
    metricsRawRef.current = byTower
    const rows = towers.map(tw=>{ const g=byTower[tw], licOnly=g.filter(r=>r.hasLic)
      let low=0,average=0,power=0
      g.forEach(r=>{ const c=dispCat(r); if(c.startsWith('Low'))low++; else if(c.startsWith('Average'))average++; else if(c.startsWith('Power'))power++ })
      return { tower:tw, count:g.length, licCount:licOnly.length,
        avgInt:avg(licOnly.map(r=>r.interactions)), avgUse:avg(licOnly.map(r=>r.usage)),
        avgCopAdd:avg(licOnly.map(r=>r.linesAdded)), avgCopFin:avg(licOnly.map(r=>r.copilotFinal)),
        avgCommit:avg(g.map(r=>r.gitCommits)), avgGitAdd:avg(g.map(r=>r.addedLines)),
        avgGitFin:avg(g.map(r=>r.gitFinal)), low, average, power,
      }
    })
    const allLicOnly=parsed.filter(r=>r.hasLic)
    let gtLow=0,gtAvg=0,gtPow=0
    parsed.forEach(r=>{ const c=dispCat(r); if(c.startsWith('Low'))gtLow++; else if(c.startsWith('Average'))gtAvg++; else if(c.startsWith('Power'))gtPow++ })
    const gtm={ count:parsed.length, licCount:allLicOnly.length,
      avgInt:avg(allLicOnly.map(r=>r.interactions)), avgUse:avg(allLicOnly.map(r=>r.usage)),
      avgCopAdd:avg(allLicOnly.map(r=>r.linesAdded)), avgCopFin:avg(allLicOnly.map(r=>r.copilotFinal)),
      avgCommit:avg(parsed.map(r=>r.gitCommits)), avgGitAdd:avg(parsed.map(r=>r.addedLines)),
      avgGitFin:avg(parsed.map(r=>r.gitFinal)), low:gtLow, average:gtAvg, power:gtPow,
    }
    setMetricsEmpty(''); setMetricsRows(rows); setMetricsGrandTotal(gtm)
    const mDisp={}
    towers.forEach(tw=>{ mDisp[tw]={}; byTower[tw].forEach(r=>{ const c=dispCat(r); mDisp[tw][c]=(mDisp[tw][c]||0)+1 }) })
    setMetricsDisp({ byTower:mDisp, towers, rawByTower:byTower })
    generateInsights('metrics', rows, gtm, setMetricsInsights)
  }

  // ── AI INSIGHTS ───────────────────────────────────────────────
  async function generateInsights(page, rows, gt, setInsights) {
    setInsights({ loading:true, html:'' })
    const summary = rows.map(r => page==='tower'
      ? `Tower: ${r.tower} | Associates: ${r.count} | Avg Copilot Interactions: ${f2(r.avgInt)} | Avg Copilot Usage: ${f2(r.avgUse)} | Avg Copilot Added Lines: ${f2(r.avgCopAdd)} | Avg Copilot Final Lines: ${f2(r.avgCopFin)} | Avg GitLab Commits: ${f2(r.avgCommit)} | Avg Git Added Lines: ${f2(r.avgGitAdd)} | Avg Git Final Lines: ${f2(r.avgGitFin)}`
      : `Tower: ${r.tower} | Associates: ${r.count} | With Copilot License: ${r.licCount} | Avg Copilot Interactions (licensed only): ${f2(r.avgInt)} | Avg Copilot Usage: ${f2(r.avgUse)} | Avg Copilot Added Lines: ${f2(r.avgCopAdd)} | Avg Git Added Lines: ${f2(r.avgGitAdd)} | Low Users: ${r.low} | Average Users: ${r.average} | Power Users: ${r.power}`
    ).join('\n')
    const prompt = page==='tower'
      ? `You are a productivity analytics expert at CTS (Cognizant Technology Solutions). Analyse this Tower Summary data showing Copilot vs Git contribution averages by tower. Provide 5-7 concise, actionable bullet-point insights. Focus on: which towers lead or lag in Copilot adoption, Copilot vs Git utilisation gaps, and specific recommendations. Use "•" for bullets. Reference specific numbers.\n\nData:\n${summary}`
      : `You are a productivity analytics expert at CTS (Cognizant Technology Solutions). Analyse this Tower Summary Metrics data. Copilot metrics are averaged over licensed associates only; Git metrics cover all. Provide 5-7 concise, actionable bullet-point insights on: licence utilisation efficiency, user tier distribution (Low/Average/Power) per tower, which towers need attention, and specific recommendations. Use "•" for bullets. Reference specific numbers.\n\nData:\n${summary}`
    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:1000, messages:[{role:'user',content:prompt}] })
      })
      if (!res.ok) { const e=await res.json().catch(()=>({})); throw new Error(e.error?.message||`API ${res.status}`) }
      const data = await res.json()
      const text = (data.content||[]).find(b=>b.type==='text')?.text || 'No insights generated.'
      const html = text.split('\n').map(line => {
        const l=line.trim(); if (!l) return ''
        if (/^[•\-\*]/.test(l)) return `<div class="insights-bullet"><span style="color:var(--ac);font-size:18px;line-height:1.2;flex-shrink:0">•</span><span>${l.replace(/^[•\-\*]\s*/,'')}</span></div>`
        if (l.length<80&&(l.endsWith(':')||(/^[A-Z]/.test(l)&&!l.includes('.')))) return `<div class="insights-heading">${l}</div>`
        return `<p style="margin:6px 0;color:var(--gt)">${l}</p>`
      }).filter(Boolean).join('')
      setInsights({ loading:false, html: html||`<p>${text}</p>` })
    } catch(e) {
      setInsights({ loading:false, html:`<div style="color:var(--re);padding:12px;background:rgba(248,81,73,.1);border-radius:var(--r)"><strong>Could not generate insights:</strong> ${e.message}</div>` })
    }
  }

  // ── DRILL DOWN ────────────────────────────────────────────────
  function openDrill(page, twEnc, colLabel, colKey) {
    const tower=decodeURIComponent(twEnc)
    const rawData = page==='tower' ? towerRawRef.current : metricsRawRef.current
    if (!rawData) return
    let rows = rawData[tower] || []
    if (colKey==='licOnly') rows=rows.filter(r=>r.hasLic)
    if (colKey==='low')     rows=rows.filter(r=>dispCat(r).startsWith('Low'))
    if (colKey==='average') rows=rows.filter(r=>dispCat(r).startsWith('Average'))
    if (colKey==='power')   rows=rows.filter(r=>dispCat(r).startsWith('Power'))
    drillBaseRef.current = rows
    setDrillIsDisp(false)
    setDrillCols([
      {key:'assocId',label:'Associate ID'},{key:'name',label:'Associate Name'},
      {key:'actPct',label:'ACT/PCT Mapping'},{key:'linesAdded',label:'Lines Added'},
      {key:'linesDeleted',label:'Lines Deleted'},{key:'usage',label:'Copilot Usage'},
      {key:'interactions',label:'Copilot Interactions'},{key:'gitCommits',label:'GitLab Commits'},
      {key:'addedLines',label:'Added Lines'},{key:'deletedLines',label:'Deleted Lines'},
    ])
    setDrillRows(rows); setDrillSort({col:null,asc:true})
    setDrillTitle(`${tower} — ${colLabel}`)
    setDrillSubtitle(`${rows.length} associate${rows.length!==1?'s':''} • Click any column header to sort`)
    setDrillOpen(true)
  }

  function openDispDrill(page, twEnc, catEnc) {
    const tower=decodeURIComponent(twEnc), cat=decodeURIComponent(catEnc)
    const rawData = page==='tower' ? towerRawRef.current : metricsRawRef.current
    if (!rawData) return
    const rows = (rawData[tower]||[]).filter(r=>dispCat(r)===cat)
    drillBaseRef.current = rows
    setDrillIsDisp(true)
    setDrillCols([
      {key:'assocId',label:'Associate ID'},{key:'name',label:'Associate Name'},
      {key:'actPct',label:'ACT/PCT Mapping'},{key:'linesAdded',label:'Lines Added'},
      {key:'linesDeleted',label:'Lines Deleted'},{key:'interactions',label:'Copilot Interactions'},
    ])
    setDrillRows(rows); setDrillSort({col:null,asc:true})
    setDrillTitle(`${tower} — Disposition Detail`)
    setDrillSubtitle(`Category: ${cat} • ${rows.length} associate${rows.length!==1?'s':''} • Click column header to sort`)
    setDrillOpen(true)
  }

  function handleDrillSort(key) {
    setDrillSort(prev => {
      const asc = prev.col===key ? !prev.asc : true
      const numeric=['linesAdded','linesDeleted','usage','interactions','gitCommits','addedLines','deletedLines'].includes(key)
      const sorted=[...drillBaseRef.current].sort((a,b)=>{
        const va=numeric?(parseFloat(a[key])||0):String(a[key]||'').toLowerCase()
        const vb=numeric?(parseFloat(b[key])||0):String(b[key]||'').toLowerCase()
        return asc?(va>vb?1:va<vb?-1:0):(va<vb?1:va>vb?-1:0)
      })
      setDrillRows(sorted)
      return { col:key, asc }
    })
  }

  function downloadDrillCSV() {
    const rows=drillBaseRef.current; if (!rows.length) return
    const hdr = drillIsDisp
      ? ['Associate ID','Associate Name','ACT/PCT Mapping','Lines Added','Lines Deleted','Copilot Interactions']
      : ['Associate ID','Associate Name','ACT/PCT Mapping','Lines Added','Lines Deleted','Copilot Usage','Copilot Interactions','GitLab Commits','Added Lines','Deleted Lines']
    const data = drillIsDisp
      ? rows.map(r=>[r.assocId,r.name||'',r.actPct||'',r.linesAdded,r.linesDeleted,r.interactions])
      : rows.map(r=>[r.assocId,r.name||'',r.actPct||'',r.linesAdded,r.linesDeleted,r.usage,r.interactions,r.gitCommits,r.addedLines,r.deletedLines])
    dlCSV([hdr,...data], `Drill_${drillTitle.replace(/[\s/—]+/g,'_').replace(/[^a-zA-Z0-9_]/g,'')}.csv`)
    toast('Drill-down CSV downloaded','success')
  }

  // ── DOWNLOADS ─────────────────────────────────────────────────
  function downloadCSV() {
    if (!analyzedData.length) return
    dlCSV([['Associate ID','Name','Status','License','Final Copilot','Final Git','Final Ratio%','Final','Added Copilot','Added Git','Added Ratio%','Added','Overall'],
      ...analyzedData.map(r=>[r.assocId,r.name,r.status,r.license,r.fCop,r.fGit,
        r.fR!==null?r.fR.toFixed(2):'N/A', r.ffinal?'Below':'OK',
        r.aCop,r.aGit, r.aR!==null?r.aR.toFixed(2):'N/A', r.fadded?'Below':'OK',
        r.fboth?'Both Below':r.ffinal||r.fadded?'Partial':'Pass'])],
      `CTS_Analysis_${pctThreshold}pct.csv`)
    toast('CSV downloaded','success')
  }
  function downloadXLSX() {
    if (!analyzedData.length) return
    const wb=XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['Associate ID','Name','Status','License','Final Copilot','Final Git','Final Ratio%','Final','Added Copilot','Added Git','Added Ratio%','Added','Overall'],
      ...analyzedData.map(r=>[r.assocId,r.name,r.status,r.license,r.fCop,r.fGit,
        r.fR!==null?parseFloat(r.fR.toFixed(2)):'N/A', r.ffinal?'Below':'OK',
        r.aCop,r.aGit, r.aR!==null?parseFloat(r.aR.toFixed(2)):'N/A', r.fadded?'Below':'OK',
        r.fboth?'Both Below':r.ffinal||r.fadded?'Partial':'Pass'])
    ]), 'Analysis')
    XLSX.writeFile(wb, `CTS_Analysis_${pctThreshold}pct.xlsx`)
    toast('XLSX downloaded','success')
  }
  function downloadTowerCSV() {
    if (!towerRows.length) return
    dlCSV([['Tower','Associates','Avg Copilot Interactions','Avg Copilot Usage','Copilot Added Lines','Copilot Final Lines','Avg GitLab Commits','Git Added Lines','Git Final Lines'],
      ...towerRows.map(r=>[r.tower,r.count,f2(r.avgInt),f2(r.avgUse),f2(r.avgCopAdd),f2(r.avgCopFin),f2(r.avgCommit),f2(r.avgGitAdd),f2(r.avgGitFin)])],
      'CTS_Tower_Summary.csv')
    toast('Tower CSV downloaded','success')
  }
  function downloadTowerXLSX() {
    if (!towerRows.length) return
    const wb=XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['Tower','Associates','Avg Copilot Interactions','Avg Copilot Usage','Copilot Added Lines','Copilot Final Lines','Avg GitLab Commits','Git Added Lines','Git Final Lines'],
      ...towerRows.map(r=>[r.tower,r.count,parseFloat(f2(r.avgInt)),parseFloat(f2(r.avgUse)),parseFloat(f2(r.avgCopAdd)),parseFloat(f2(r.avgCopFin)),parseFloat(f2(r.avgCommit)),parseFloat(f2(r.avgGitAdd)),parseFloat(f2(r.avgGitFin))])
    ]), 'Tower Summary')
    XLSX.writeFile(wb,'CTS_Tower_Summary.xlsx')
    toast('Tower XLSX downloaded','success')
  }
  function downloadMetricsCSV() {
    if (!metricsRows.length) return
    dlCSV([['Tower','Associates','Has Copilot License','Copilot Interactions','Copilot Usage','Copilot Added Lines','Copilot Final Lines','GitLab Commits','Git Added Lines','Git Final Lines','Low User','Average User','Power User'],
      ...metricsRows.map(r=>[r.tower,r.count,r.licCount,f2(r.avgInt),f2(r.avgUse),f2(r.avgCopAdd),f2(r.avgCopFin),f2(r.avgCommit),f2(r.avgGitAdd),f2(r.avgGitFin),r.low,r.average,r.power])],
      'CTS_Tower_Metrics.csv')
    toast('Metrics CSV downloaded','success')
  }
  function downloadMetricsXLSX() {
    if (!metricsRows.length) return
    const wb=XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['Tower','Associates','Has Copilot License','Copilot Interactions','Copilot Usage','Copilot Added Lines','Copilot Final Lines','GitLab Commits','Git Added Lines','Git Final Lines','Low User','Average User','Power User'],
      ...metricsRows.map(r=>[r.tower,r.count,r.licCount,parseFloat(f2(r.avgInt)),parseFloat(f2(r.avgUse)),parseFloat(f2(r.avgCopAdd)),parseFloat(f2(r.avgCopFin)),parseFloat(f2(r.avgCommit)),parseFloat(f2(r.avgGitAdd)),parseFloat(f2(r.avgGitFin)),r.low,r.average,r.power])
    ]), 'Tower Metrics')
    XLSX.writeFile(wb,'CTS_Tower_Metrics.xlsx')
    toast('Metrics XLSX downloaded','success')
  }

  // ── MSAL ──────────────────────────────────────────────────────
  function getMsal() {
    if (!msalInstRef.current) msalInstRef.current = new window.msal.PublicClientApplication({
      auth:{ clientId:'14d82eec-204b-4c2f-b7e8-296a70dab67e', authority:'https://login.microsoftonline.com/common', redirectUri:window.location.href.split('?')[0].split('#')[0] },
      cache:{ cacheLocation:'sessionStorage', storeAuthStateInCookie:false }
    })
    return msalInstRef.current
  }
  async function msalSignIn() {
    try {
      const app=getMsal(); await app.initialize()
      const r=await app.loginPopup({ scopes:['User.Read','Mail.Send'], prompt:'select_account' })
      msalAccountRef.current=r.account
      setMsalSignedIn(true); setMsalUsername(r.account.username)
      toast('Signed in as '+r.account.username,'success')
    } catch(e) { if (e.errorCode!=='user_cancelled') toast('Sign-in failed: '+(e.message||e.errorCode),'error',6000) }
  }
  async function getToken() {
    const app=getMsal()
    try { const s=await app.acquireTokenSilent({scopes:['User.Read','Mail.Send'],account:msalAccountRef.current}); return s.accessToken }
    catch(e) { const p=await app.acquireTokenPopup({scopes:['User.Read','Mail.Send'],account:msalAccountRef.current}); return p.accessToken }
  }
  async function sendOneEmail(token,to,subject,body) {
    const r=await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{
      method:'POST', headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},
      body:JSON.stringify({message:{subject,body:{contentType:'Text',content:body},toRecipients:[{emailAddress:{address:to}}]},saveToSentItems:true})
    })
    if (!r.ok) { const e=await r.json().catch(()=>({})); throw new Error(e.error?.message||`Graph ${r.status}`) }
  }

  function openEmailModal() {
    setEmailSubject(`[CTS] Your Copilot vs Git Productivity Report — ${pctThreshold}% Threshold`)
    setEmailStep(1); setEmailStatus({visible:false,msg:'',type:''}); setSendLog(''); setSendProgress({label:'',pct:0})
    setShowEmail(true)
  }
  async function sendViaGraph() {
    if (!msalAccountRef.current) { toast('Please sign in first','error'); return }
    const fails=analyzedData.filter(r=>r.ffinal||r.fadded)
    if (!fails.length) { toast('No flagged recipients','error'); return }
    setEmailStep(2); setEmailStatus({visible:false,msg:'',type:''}); setSendLog('')
    let sent=0, failed=0, token
    try {
      setSendProgress({label:'Acquiring access token...',pct:0})
      token=await getToken()
    } catch(e) {
      setEmailStatus({visible:true,msg:'<strong>Auth failed.</strong> '+e.message,type:'error'}); setEmailStep(1); return
    }
    for (let i=0; i<fails.length; i++) {
      const r=fails[i], to=`${r.assocId}@cognizant.com`
      const body=emailBodyTxt+'\n\n'+`--- Metrics (Threshold: ${pctThreshold}%) ---\n\n`+
        `Final Copilot Net : ${r.fCop}  |  Final Git Net : ${r.fGit}\n`+
        `Copilot % of Git  : ${r.fR!==null?r.fR.toFixed(2)+'%':'N/A'}\n`+
        `Final Status      : ${r.fR===null?'N/A':r.fR>pctThreshold?'ABOVE ✓':r.fR===pctThreshold?'EQUAL':'BELOW ✗'}\n\n`+
        `Added Copilot     : ${r.aCop}  |  Added Git : ${r.aGit}\n`+
        `Copilot % of Git  : ${r.aR!==null?r.aR.toFixed(2)+'%':'N/A'}\n`+
        `Added Status      : ${r.aR===null?'N/A':r.aR>pctThreshold?'ABOVE ✓':r.aR===pctThreshold?'EQUAL':'BELOW ✗'}\n\n`+
        `Overall: ${r.fboth?'Both Below':r.ffinal||r.fadded?'Partial':'Pass'} | Status: ${r.status} | License: ${r.license}\n---`
      setSendProgress({label:`Sending ${i+1}/${fails.length}: ${to}`,pct:Math.round(((i+.5)/fails.length)*100)})
      try { await sendOneEmail(token,to,emailSubject,body); sent++; setSendLog(p=>p+`✓ ${to}\n`) }
      catch(e) { failed++; setSendLog(p=>p+`✗ ${to} — ${e.message}\n`) }
      setSendProgress({label:`Sending ${i+1}/${fails.length}: ${to}`,pct:Math.round(((i+1)/fails.length)*100)})
      await new Promise(r=>setTimeout(r,250))
    }
    setSendProgress({label:`Complete — ${sent} sent, ${failed} failed.`,pct:100})
    setEmailStatus({visible:true, msg:failed===0?`<strong>✓ All ${sent} email(s) sent.</strong>`:`<strong>${sent} sent, ${failed} failed.</strong>`, type:failed===0?'success':'error'})
  }

  // ── RESET ─────────────────────────────────────────────────────
  function doReset() {
    setShowConfirm(false)
    setAllData([]); setAnalyzedData([])
    setTowerRows([]); setTowerGrandTotal(null); setTowerDisp(null); setTowerInsights(null)
    setMetricsRows([]); setMetricsGrandTotal(null); setMetricsDisp(null); setMetricsInsights(null)
    setMetrics({total:'—',final:'—',finalSub:'at threshold',added:'—',addedSub:'at threshold',both:'—',bothSub:'immediate review'})
    setCurrentTab('all'); setPctThreshold(60); setFilterStatus('Active'); setFilterLicense('Yes'); setSearchInput('')
    setTwStatus('All'); setTwLicense('All'); setTwGitlab('All')
    setTmStatus('All'); setTmGitlab('All')
    setTowerEmpty('Load data to see tower summaries'); setMetricsEmpty('Load data to see metrics')
    destroyCharts()
    fetchData()
  }

  // ── DERIVED ───────────────────────────────────────────────────
  const tableRows = analyzedData.filter(r => {
    if (searchInput && !(r.assocId+' '+r.name).toLowerCase().includes(searchInput.toLowerCase())) return false
    if (currentTab==='fail-final') return r.ffinal
    if (currentTab==='fail-added') return r.fadded
    if (currentTab==='fail-both')  return r.fboth
    return true
  })
  const emailFails = analyzedData.filter(r=>r.ffinal||r.fadded)

  const sboxStyle = (type) => ({
    success:{ bg:'rgba(63,185,80,.1)',   border:'var(--gr)', color:'#3fb950' },
    error:  { bg:'rgba(248,81,73,.1)',   border:'var(--re)', color:'#f85149' },
    info:   { bg:'rgba(88,166,255,.1)',  border:'var(--ac)', color:'#58a6ff' },
  }[type] || { bg:'#1C2128', border:'var(--bo)', color:'var(--dk)' })

  // ── DISPOSITION CARDS HELPER ──────────────────────────────────
  function DispCards({ dispData, page }) {
    if (!dispData) return (
      <div className="empty" style={{padding:'40px'}}>
        <div className="empty-ico">📋</div>
        <div className="empty-txt">Disposition categories will appear after data is loaded</div>
      </div>
    )
    return (
      <div className="disp-grid">
        {dispData.towers.map(tw => {
          const cats=dispData.byTower[tw], total=dispData.rawByTower[tw].length
          return (
            <div key={tw} className="disp-card">
              <div className="disp-card-hdr">🏗 {tw} <span style={{fontSize:'12px',color:'var(--gt)',fontWeight:400,marginLeft:8}}>{total} associates</span></div>
              {DISP_CATS.filter(c=>cats[c]>0).length===0
                ? <div className="disp-row"><span className="disp-label" style={{color:'var(--gt)'}}>No data</span></div>
                : DISP_CATS.filter(c=>cats[c]>0).map(c=>{
                  const ci=DISP_CATS.indexOf(c)
                  const label=c.includes('(')
                    ? <><span className={`tag ${DISP_CSS[ci]||'cm2'}`} style={{fontSize:'10px',marginRight:6}}>{c.split(' (')[0]}</span>{c.substring(c.indexOf('('))}</>
                    : c
                  return (
                    <div key={c} className="disp-row">
                      <span className="disp-label">{label}</span>
                      <span className="disp-val">
                        <span className="num-link" onClick={()=>openDispDrill(page,encodeURIComponent(tw),encodeURIComponent(c))}>{cats[c]}</span>
                        {' '}<span style={{fontSize:'11px',color:'var(--gt)',fontWeight:400}}>({((cats[c]/total)*100).toFixed(0)}%)</span>
                      </span>
                    </div>
                  )
                })
              }
            </div>
          )
        })}
      </div>
    )
  }

  // ── INSIGHTS CARD HELPER ──────────────────────────────────────
  function InsightsCard({ insights, page, rows, gt, setInsights, title }) {
    if (!insights) return null
    return (
      <div className="rs" style={{overflow:'visible',marginBottom:20}}>
        <div className="rh" style={{background:'linear-gradient(135deg,#21262D,#2C333B)'}}>
          <div style={{display:'flex',alignItems:'center',gap:10}}>
            <span style={{fontSize:'18px'}}>🧠</span>
            <div className="rt">{title}</div>
          </div>
          <button className="btn bo-btn" style={{color:'var(--on-header)',borderColor:'var(--on-header-muted)',fontSize:'12px'}}
            onClick={()=>generateInsights(page,rows,gt,setInsights)}>↺ Refresh</button>
        </div>
        <div className="insights-body">
          {insights.loading
            ? <div style={{color:'var(--gt)',fontStyle:'italic'}}>⏳ Analysing data with AI...</div>
            : <div dangerouslySetInnerHTML={{__html:insights.html}} />}
        </div>
      </div>
    )
  }

  // ── RENDER ────────────────────────────────────────────────────
  return (
    <div className="cvg-root">

      {/* TOASTS */}
      <div className="tcon">
        {toasts.map(t=>(
          <div key={t.id} className={`tst ${t.type}`}>
            <span style={{fontWeight:600}}>{t.type==='success'?'✓':t.type==='error'?'✕':'i'}</span> {t.msg}
          </div>
        ))}
      </div>

      {/* NAV */}
      <nav className="nav-bar">
        {[['p-analysis','📊 Analysis'],['p-tower','🏗 Tower Summary'],['p-metrics','📈 Tower Summary Metrics']].map(([id,label])=>(
          <button key={id} className={`nav-tab${currentPage===id?' active':''}`} onClick={()=>setCurrentPage(id)}>{label}</button>
        ))}
        <div style={{marginLeft:'auto',display:'flex',alignItems:'center',padding:'0 4px'}}>
          <button className="btn br-btn" style={{fontSize:'12px',padding:'7px 13px'}} onClick={()=>setShowConfirm(true)}>↺ Reset All</button>
        </div>
      </nav>

      <div className="main">
        {fetchState==='loading' && <div className="fetch-loading">⏳ Loading data...</div>}
        {fetchState==='error'   && <div className="fetch-error">⚠ Failed to load data: {fetchError}</div>}

        {/* ── PAGE 1: ANALYSIS ─────────────────────────────── */}
        <div className={`page${currentPage==='p-analysis'?' active':''}`}>
          <div className="ctrl-panel">
            <div className="ctrl-row">
              <div className="cg">
                <label>Threshold (%)</label>
                <div className="pw">
                  <input type="number" value={pctThreshold} min="0" max="100" step="1"
                    onChange={e=>setPctThreshold(parseFloat(e.target.value)||0)} />
                  <span className="ps">%</span>
                </div>
              </div>
              <div className="cg">
                <label>Final Status</label>
                <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
                  <option value="Active">Active</option><option value="All">All</option><option value="Inactive">Inactive</option>
                </select>
              </div>
              <div className="cg">
                <label>Has Copilot License</label>
                <select value={filterLicense} onChange={e=>setFilterLicense(e.target.value)}>
                  <option value="Yes">Yes</option><option value="All">All</option><option value="No">No</option>
                </select>
              </div>
              <div className="cg">
                <label>Search Associate</label>
                <input type="text" placeholder="ID or Name..." value={searchInput} onChange={e=>setSearchInput(e.target.value)} />
              </div>
            </div>
            <div className="cactions">
              <button className="btn bg-btn" onClick={downloadCSV}  disabled={!analyzedData.length}>⬇ CSV</button>
              <button className="btn bt-btn" onClick={downloadXLSX} disabled={!analyzedData.length}>⬇ XLSX</button>
              <button className="btn ba-btn" style={{fontSize:'12px',padding:'7px 14px'}} onClick={openEmailModal} disabled={!analyzedData.length}>✉ Email Reports</button>
            </div>
          </div>

          <div className="mrow">
            <div className="mc bl"><div className="mlabel">Total Associates</div><div className="mval">{metrics.total}</div><div className="msub">filtered dataset</div></div>
            <div className="mc gr"><div className="mlabel">Final: Copilot &lt; Git</div><div className="mval">{metrics.final}</div><div className="msub">{metrics.finalSub}</div></div>
            <div className="mc am"><div className="mlabel">Added: Copilot &lt; Git</div><div className="mval">{metrics.added}</div><div className="msub">{metrics.addedSub}</div></div>
            <div className="mc re"><div className="mlabel">Both Criteria Below</div><div className="mval">{metrics.both}</div><div className="msub">{metrics.bothSub}</div></div>
          </div>

          {analyzedData.length>0 && (
            <div className="cgrid">
              <div className="cc"><div className="cch"><div><div className="cct">Final Copilot vs Final Git</div><div className="ccs">Net lines — top 20</div></div></div><div className="cw"><canvas ref={canvasFinalRef} /></div></div>
              <div className="cc"><div className="cch"><div><div className="cct">Added Copilot vs Added Git</div><div className="ccs">Raw lines added — top 20</div></div></div><div className="cw"><canvas ref={canvasAddedRef} /></div></div>
              <div className="cc"><div className="cch"><div><div className="cct">Threshold Breach Overview</div><div className="ccs">Pass vs fail at threshold</div></div></div><div className="cw"><canvas ref={canvasBreachRef} /></div></div>
              <div className="cc"><div className="cch"><div><div className="cct">Copilot Utilization Ratio</div><div className="ccs">Copilot % of Git — red = below</div></div></div><div className="cw"><canvas ref={canvasRatioRef} /></div></div>
            </div>
          )}

          {analyzedData.length>0 ? (
            <div className="rs">
              <div className="rh">
                <div className="rt">Analysis Results — {pctThreshold}% Threshold | {tableRows.length} record(s)</div>
                <div className="ra">
                  <button className="btn bo-btn" style={{color:'var(--on-header)',borderColor:'var(--on-header-muted)',fontSize:'12px'}} onClick={downloadCSV}>⬇ CSV</button>
                  <button className="btn bo-btn" style={{color:'var(--on-header)',borderColor:'var(--on-header-muted)',fontSize:'12px'}} onClick={downloadXLSX}>⬇ XLSX</button>
                </div>
              </div>
              <div className="tbar">
                {[['all','All'],['fail-final','Final Below'],['fail-added','Added Below'],['fail-both','Both Below']].map(([k,l])=>(
                  <button key={k} className={`tbtn${currentTab===k?' active':''}`} onClick={()=>setCurrentTab(k)}>{l}</button>
                ))}
              </div>
              <div className="tw">
                <table>
                  <thead><tr>
                    <th>Associate ID</th><th>Name</th><th>Status</th><th>License</th>
                    <th>Final Copilot</th><th>Final Git</th><th>Copilot% of Git</th><th>Final</th>
                    <th>Added Copilot</th><th>Added Git</th><th>Copilot% of Git</th><th>Added</th><th>Overall</th>
                  </tr></thead>
                  <tbody>
                    {tableRows.length===0
                      ? <tr><td colSpan={13} style={{textAlign:'center',padding:'40px',color:'var(--gt)'}}>No records match the current filter</td></tr>
                      : tableRows.map((r,i)=>{
                        const fr=r.fR!==null?r.fR.toFixed(2)+'%':'N/A'
                        const ar=r.aR!==null?r.aR.toFixed(2)+'%':'N/A'
                        const l=r.license.toLowerCase(), isY=l==='yes'||l==='true'||l==='1'||l==='y'
                        const fTag=r.fR===null?<span className="tag" style={{background:'#333',color:'#888'}}>N/A</span>:Math.abs(r.fR-pctThreshold)<0.05?<span className="tag tb2">Equal</span>:r.fR>pctThreshold?<span className="tag tp">Above</span>:<span className="tag tf">Below</span>
                        const aTag=r.aR===null?<span className="tag" style={{background:'#333',color:'#888'}}>N/A</span>:Math.abs(r.aR-pctThreshold)<0.05?<span className="tag tb2">Equal</span>:r.aR>pctThreshold?<span className="tag tp">Above</span>:<span className="tag tf">Below</span>
                        const os=r.fboth?<span className="tag tf">Both Below</span>:r.ffinal||r.fadded?<span className="tag tw2">Partial</span>:<span className="tag tp">Pass</span>
                        return (
                          <tr key={i}>
                            <td style={{fontFamily:'var(--mo)',fontWeight:500,color:'var(--bm)'}}>{r.assocId}</td>
                            <td>{r.name}</td>
                            <td>{r.status.toLowerCase()==='active'?<span className="tag tp">Active</span>:<span className="tag" style={{background:'#333',color:'#aaa'}}>{r.status}</span>}</td>
                            <td>{isY?<span className="tag tp">Yes</span>:<span className="tag tf">No</span>}</td>
                            <td className={r.fCop<0?'dn':'dp'}>{r.fCop}</td><td>{r.fGit}</td>
                            <td className={r.ffinal?'dn':''}>{fr}</td><td>{fTag}</td>
                            <td className={r.aCop<r.aGit?'dn':'dp'}>{r.aCop}</td><td>{r.aGit}</td>
                            <td className={r.fadded?'dn':''}>{ar}</td><td>{aTag}</td><td>{os}</td>
                          </tr>
                        )
                      })
                    }
                  </tbody>
                </table>
              </div>
            </div>
          ) : fetchState==='done' ? (
            <div className="rs"><div className="empty" style={{padding:'70px 20px'}}>
              <div className="empty-ico">📊</div>
              <div className="empty-txt">No data to display</div>
              <div style={{fontSize:'13px',marginTop:'6px',color:'var(--gt)'}}>Adjust filters or reload data</div>
            </div></div>
          ) : null}
        </div>

        {/* ── PAGE 2: TOWER SUMMARY ────────────────────────── */}
        <div className={`page${currentPage==='p-tower'?' active':''}`}>
          <div className="ctrl-panel">
            <div className="ctrl-row">
              <div className="cg"><label>Final Status</label>
                <select value={twStatus} onChange={e=>setTwStatus(e.target.value)}>
                  <option value="All">All</option>{twStatusOpts.map(v=><option key={v}>{v}</option>)}
                </select>
              </div>
              <div className="cg"><label>Has Copilot License</label>
                <select value={twLicense} onChange={e=>setTwLicense(e.target.value)}>
                  <option value="All">All</option><option value="Yes">Yes</option><option value="No">No</option>
                </select>
              </div>
              <div className="cg"><label>Active GitLab Account</label>
                <select value={twGitlab} onChange={e=>setTwGitlab(e.target.value)}>
                  <option value="All">All</option>{twGitlabOpts.map(v=><option key={v}>{v}</option>)}
                </select>
              </div>
            </div>
          </div>

          <div className="rs">
            <div className="rh">
              <div className="rt">Tower Summary — Averages</div>
              <div className="ra">
                <button className="btn bo-btn" style={{color:'var(--on-header)',borderColor:'var(--on-header-muted)',fontSize:'12px'}} onClick={downloadTowerCSV}>⬇ CSV</button>
                <button className="btn bt-btn" style={{fontSize:'12px'}} onClick={downloadTowerXLSX}>⬇ XLSX</button>
              </div>
            </div>
            <div className="tw"><table>
              <thead><tr>
                <th>Tower</th>
                <th style={{textAlign:'right'}}>Associates</th><th style={{textAlign:'right'}}>Avg Copilot Interactions</th>
                <th style={{textAlign:'right'}}>Avg Copilot Usage</th><th style={{textAlign:'right'}}>Copilot Added Lines</th>
                <th style={{textAlign:'right'}}>Copilot Final Lines</th><th style={{textAlign:'right'}}>Avg GitLab Commits</th>
                <th style={{textAlign:'right'}}>Git Added Lines</th><th style={{textAlign:'right'}}>Git Final Lines</th>
              </tr></thead>
              <tbody>
                {towerRows.length===0
                  ? <tr><td colSpan={9}><div className="empty" style={{padding:'40px'}}><div className="empty-ico">🏗</div><div className="empty-txt">{towerEmpty}</div></div></td></tr>
                  : <>
                    {towerRows.map((r,i)=>(
                      <tr key={i}>
                        <td style={{fontWeight:600,color:'var(--bm)'}}>{r.tower}</td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'Associates','all')}>{r.count}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'Avg Copilot Interactions','interactions')}>{f2(r.avgInt)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'Avg Copilot Usage','usage')}>{f2(r.avgUse)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'Copilot Added Lines','linesAdded')}>{f2(r.avgCopAdd)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'Copilot Final Lines','copilotFinal')}>{f2(r.avgCopFin)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'GitLab Commits','gitCommits')}>{f2(r.avgCommit)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'Git Added Lines','addedLines')}>{f2(r.avgGitAdd)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('tower',encodeURIComponent(r.tower),'Git Final Lines','gitFinal')}>{f2(r.avgGitFin)}</span></td>
                      </tr>
                    ))}
                    {towerGrandTotal && (
                      <tr style={{background:'#1C2128',fontWeight:600,borderTop:'2px solid var(--bo)'}}>
                        <td style={{color:'var(--bm)'}}>Σ Grand Total / Avg</td>
                        <td className="tower-num">{towerGrandTotal.count}</td>
                        <td className="tower-num">{f2(towerGrandTotal.avgInt)}</td><td className="tower-num">{f2(towerGrandTotal.avgUse)}</td>
                        <td className="tower-num">{f2(towerGrandTotal.avgCopAdd)}</td><td className="tower-num">{f2(towerGrandTotal.avgCopFin)}</td>
                        <td className="tower-num">{f2(towerGrandTotal.avgCommit)}</td><td className="tower-num">{f2(towerGrandTotal.avgGitAdd)}</td>
                        <td className="tower-num">{f2(towerGrandTotal.avgGitFin)}</td>
                      </tr>
                    )}
                  </>
                }
              </tbody>
            </table></div>
          </div>

          <div className="rs">
            <div className="rh"><div className="rt">Disposition Analysis by Tower - Based on Copilot Added Lines</div></div>
            <DispCards dispData={towerDisp} page="tower" />
          </div>

          <InsightsCard insights={towerInsights} page="tower" rows={towerRows} gt={towerGrandTotal}
            setInsights={setTowerInsights} title="AI Insights — Tower Summary" />
        </div>

        {/* ── PAGE 3: TOWER SUMMARY METRICS ────────────────── */}
        <div className={`page${currentPage==='p-metrics'?' active':''}`}>
          <div className="ctrl-panel">
            <div className="ctrl-row">
              <div className="cg"><label>Final Status</label>
                <select value={tmStatus} onChange={e=>setTmStatus(e.target.value)}>
                  <option value="All">All</option>{tmStatusOpts.map(v=><option key={v}>{v}</option>)}
                </select>
              </div>
              <div className="cg"><label>Active GitLab Account</label>
                <select value={tmGitlab} onChange={e=>setTmGitlab(e.target.value)}>
                  <option value="All">All</option>{tmGitlabOpts.map(v=><option key={v}>{v}</option>)}
                </select>
              </div>
            </div>
          </div>

          <div className="rs">
            <div className="rh">
              <div className="rt">Tower Summary — Averages</div>
              <div className="ra">
                <button className="btn bo-btn" style={{color:'var(--on-header)',borderColor:'var(--on-header-muted)',fontSize:'12px'}} onClick={downloadMetricsCSV}>⬇ CSV</button>
                <button className="btn bt-btn" style={{fontSize:'12px'}} onClick={downloadMetricsXLSX}>⬇ XLSX</button>
              </div>
            </div>
            <div className="tw"><table>
              <thead><tr>
                <th>Tower</th>
                <th style={{textAlign:'right'}}>Associates</th><th style={{textAlign:'right'}}>Has Copilot License</th>
                <th style={{textAlign:'right'}}>Copilot Interactions</th><th style={{textAlign:'right'}}>Copilot Usage</th>
                <th style={{textAlign:'right'}}>Copilot Added Lines</th><th style={{textAlign:'right'}}>Copilot Final Lines</th>
                <th style={{textAlign:'right'}}>GitLab Commits</th><th style={{textAlign:'right'}}>Git Added Lines</th>
                <th style={{textAlign:'right'}}>Git Final Lines</th>
                <th style={{textAlign:'right'}}>Low User</th><th style={{textAlign:'right'}}>Average User</th><th style={{textAlign:'right'}}>Power User</th>
              </tr></thead>
              <tbody>
                {metricsRows.length===0
                  ? <tr><td colSpan={13}><div className="empty" style={{padding:'40px'}}><div className="empty-ico">🏗</div><div className="empty-txt">{metricsEmpty}</div></div></td></tr>
                  : <>
                    {metricsRows.map((r,i)=>(
                      <tr key={i}>
                        <td style={{fontWeight:600,color:'var(--bm)'}}>{r.tower}</td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Associates','all')}>{r.count}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Has Copilot License','licOnly')}>{r.licCount}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Copilot Interactions','interactions')}>{f2(r.avgInt)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Copilot Usage','usage')}>{f2(r.avgUse)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Copilot Added Lines','linesAdded')}>{f2(r.avgCopAdd)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Copilot Final Lines','copilotFinal')}>{f2(r.avgCopFin)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'GitLab Commits','gitCommits')}>{f2(r.avgCommit)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Git Added Lines','addedLines')}>{f2(r.avgGitAdd)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Git Final Lines','gitFinal')}>{f2(r.avgGitFin)}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Low Users','low')}>{r.low}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Average Users','average')}>{r.average}</span></td>
                        <td className="tower-num"><span className="num-link" onClick={()=>openDrill('metrics',encodeURIComponent(r.tower),'Power Users','power')}>{r.power}</span></td>
                      </tr>
                    ))}
                    {metricsGrandTotal && (
                      <tr style={{background:'#1C2128',fontWeight:600,borderTop:'2px solid var(--bo)'}}>
                        <td style={{color:'var(--bm)'}}>Σ Grand Total / Avg</td>
                        <td className="tower-num">{metricsGrandTotal.count}</td>
                        <td className="tower-num">{metricsGrandTotal.licCount}</td>
                        <td className="tower-num" title="Avg of licensed associates only">{f2(metricsGrandTotal.avgInt)}</td>
                        <td className="tower-num" title="Avg of licensed associates only">{f2(metricsGrandTotal.avgUse)}</td>
                        <td className="tower-num" title="Avg of licensed associates only">{f2(metricsGrandTotal.avgCopAdd)}</td>
                        <td className="tower-num" title="Avg of licensed associates only">{f2(metricsGrandTotal.avgCopFin)}</td>
                        <td className="tower-num" title="Avg across all">{f2(metricsGrandTotal.avgCommit)}</td>
                        <td className="tower-num" title="Avg across all">{f2(metricsGrandTotal.avgGitAdd)}</td>
                        <td className="tower-num" title="Avg across all">{f2(metricsGrandTotal.avgGitFin)}</td>
                        <td className="tower-num"><span className="tag cl">{metricsGrandTotal.low}</span></td>
                        <td className="tower-num"><span className="tag ca">{metricsGrandTotal.average}</span></td>
                        <td className="tower-num"><span className="tag cp">{metricsGrandTotal.power}</span></td>
                      </tr>
                    )}
                  </>
                }
              </tbody>
            </table></div>
          </div>

          <div className="rs">
            <div className="rh"><div className="rt">Disposition Analysis by Tower - Based on Copilot Added Lines</div></div>
            <DispCards dispData={metricsDisp} page="metrics" />
          </div>

          <InsightsCard insights={metricsInsights} page="metrics" rows={metricsRows} gt={metricsGrandTotal}
            setInsights={setMetricsInsights} title="AI Insights — Tower Summary Metrics" />
        </div>
      </div>

      {/* ── DRILL MODAL ──────────────────────────────────── */}
      <div className={`mov${drillOpen?' open':''}`}>
        <div className="mdl" style={{width:'92vw',maxWidth:'1100px',maxHeight:'88vh'}}>
          <div className="mhdr">
            <div>
              <div className="mttl">{drillTitle}</div>
              <div style={{fontSize:'12px',color:'var(--on-header-muted)',marginTop:2}}>{drillSubtitle}</div>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:10}}>
              <button className="btn bo-btn" style={{color:'var(--on-header)',borderColor:'var(--on-header-muted)',fontSize:'12px'}} onClick={downloadDrillCSV}>⬇ CSV</button>
              <button className="mclose" onClick={()=>setDrillOpen(false)}>✕</button>
            </div>
          </div>
          <div style={{overflow:'auto',maxHeight:'calc(88vh - 80px)'}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:'13px'}}>
              <thead><tr className="drill-hdr">
                {drillCols.map(c=>(
                  <th key={c.key} style={{cursor:'pointer',userSelect:'none'}} onClick={()=>handleDrillSort(c.key)}>
                    {c.label} <span style={{fontSize:'10px',opacity:drillSort.col===c.key?1:.5}}>
                      {drillSort.col===c.key?(drillSort.asc?'↑':'↓'):'⇅'}
                    </span>
                  </th>
                ))}
              </tr></thead>
              <tbody className="drill-body">
                {drillRows.length===0
                  ? <tr><td colSpan={drillCols.length} style={{textAlign:'center',padding:'24px',color:'var(--gt)'}}>No associates</td></tr>
                  : drillIsDisp
                  ? drillRows.map((r,i)=>(
                    <tr key={i}>
                      <td style={{fontFamily:'var(--mo)',fontWeight:500,color:'var(--bm)'}}>{r.assocId}</td>
                      <td>{r.name||'—'}</td><td>{r.actPct||'—'}</td>
                      <td className="tower-num">{r.linesAdded}</td><td className="tower-num">{r.linesDeleted}</td>
                      <td className="tower-num">{r.interactions}</td>
                    </tr>
                  ))
                  : drillRows.map((r,i)=>(
                    <tr key={i}>
                      <td style={{fontFamily:'var(--mo)',fontWeight:500,color:'var(--bm)'}}>{r.assocId}</td>
                      <td>{r.name||'—'}</td><td>{r.actPct||'—'}</td>
                      <td className="tower-num">{r.linesAdded}</td><td className="tower-num">{r.linesDeleted}</td>
                      <td className="tower-num">{r.usage}</td><td className="tower-num">{r.interactions}</td>
                      <td className="tower-num">{r.gitCommits}</td><td className="tower-num">{r.addedLines}</td>
                      <td className="tower-num">{r.deletedLines}</td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* ── EMAIL MODAL ──────────────────────────────────── */}
      <div className={`mov${showEmail?' open':''}`}>
        <div className="mdl">
          <div className="mhdr">
            <div className="mttl">✉ Send Individual Email Reports</div>
            <button className="mclose" onClick={()=>{ setShowEmail(false); setEmailStep(1) }}>✕</button>
          </div>
          <div className="mbody">
            {emailStep===1 ? (
              <>
                <div style={{background:'rgba(88,166,255,.08)',border:'1px solid rgba(88,166,255,.2)',borderRadius:'var(--r)',padding:14,marginBottom:14}}>
                  <div style={{fontSize:13,fontWeight:600,color:'var(--bm)',marginBottom:5}}>Step 1 — Sign in with your Cognizant Microsoft account</div>
                  <div style={{fontSize:12,color:'var(--gt)',marginBottom:10}}>Grants Mail.Send permission via Microsoft Graph. Token used only for this session.</div>
                  <button className="btn bp" onClick={msalSignIn} disabled={msalSignedIn} style={msalSignedIn?{background:'var(--gr)'}:{}}>
                    <svg width="16" height="16" viewBox="0 0 21 21" fill="none"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>
                    {msalSignedIn ? 'Signed in ✓' : 'Sign in with Microsoft'}
                  </button>
                  {msalSignedIn && <span style={{fontSize:12,color:'var(--gr)',marginLeft:12}}>✓ {msalUsername}</span>}
                </div>
                <div className="mst">Subject Line</div>
                <input type="text" value={emailSubject} onChange={e=>setEmailSubject(e.target.value)} />
                <div className="mst">Message Preamble</div>
                <textarea value={emailBodyTxt} onChange={e=>setEmailBodyTxt(e.target.value)} />
                <div className="mst">Recipients — <span>{emailFails.length}</span> flagged associates</div>
                <div className="rlist">
                  {emailFails.length===0
                    ? <div className="ritem" style={{color:'var(--gt)'}}>No flagged associates.</div>
                    : emailFails.map((r,i)=>(
                      <div key={i} className="ritem">
                        <div><div style={{fontWeight:500}}>{r.name||r.assocId}</div><div className="remail">{r.assocId}@cognizant.com</div></div>
                        <span className={`tag ${r.fboth?'tf':'tw2'}`}>{r.fboth?'Both below':r.ffinal?'Final below':'Added below'}</span>
                      </div>
                    ))
                  }
                </div>
              </>
            ) : (
              <>
                <div style={{fontSize:13,color:'var(--gt)',marginBottom:10}}>{sendProgress.label}</div>
                <div className="pb-wrap"><div className="pb-fill" style={{width:`${sendProgress.pct}%`}} /></div>
                <div className="slog" style={{whiteSpace:'pre'}}>{sendLog}</div>
              </>
            )}
            {emailStatus.visible && (() => {
              const s = sboxStyle(emailStatus.type)
              return <div style={{display:'block',marginTop:12,padding:'12px 14px',borderRadius:'var(--r)',fontSize:13,lineHeight:1.7,border:`1px solid ${s.border}`,background:s.bg,color:s.color}} dangerouslySetInnerHTML={{__html:emailStatus.msg}} />
            })()}
          </div>
          <div className="mfoot">
            <span className="mfnote">{msalSignedIn?`Signed in as ${msalUsername}`:'Sign in above to enable sending'}</span>
            <div className="mfbtns">
              <button className="btn bo-btn" onClick={()=>{ setShowEmail(false); setEmailStep(1) }}>Cancel</button>
              <button className="btn bp" onClick={sendViaGraph} disabled={!msalSignedIn||emailStep===2}>✉ Send Emails ({emailFails.length})</button>
            </div>
          </div>
        </div>
      </div>

      {/* ── CONFIRM RESET ────────────────────────────────── */}
      <div className={`mov${showConfirm?' open':''}`}>
        <div className="conf-mdl">
          <div className="conf-hdr"><span style={{fontSize:22,color:'#fff'}}>⚠</span><div style={{fontSize:15,fontWeight:600,color:'#fff'}}>Reset Dashboard?</div></div>
          <div className="conf-body">This will <strong>clear all analysis, charts and results</strong> and reload the data.</div>
          <div className="conf-foot">
            <button className="btn bo-btn" onClick={()=>setShowConfirm(false)}>Cancel</button>
            <button className="btn br-btn" onClick={doReset}>Yes, Reset Everything</button>
          </div>
        </div>
      </div>

    </div>
  )
}
