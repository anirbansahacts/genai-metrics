import { useState, useRef, useEffect } from 'react'
import XLSX from 'xlsx-js-style'
import JSZip from 'jszip'
import './BatchFlow.css'

// ── Module-level logger (set before each processing run) ─────────────────────
let _log = null

// ── Pure utilities ────────────────────────────────────────────────────────────
function colKey(name) { return String(name || '').toLowerCase().trim() }
function colLetter(idx) {
  let letter = '', i = idx + 1
  while (i > 0) { const m = (i - 1) % 26; letter = String.fromCharCode(65 + m) + letter; i = Math.floor((i - 1) / 26) }
  return letter
}
function hasCol(data, name) {
  if (!data?.length) return false
  const k = colKey(name)
  const sample = data.length > 5 ? data.slice(0, 5) : data
  return sample.some(row => Object.keys(row).some(h => colKey(h) === k))
}
function findColActual(row, name) {
  const k = colKey(name)
  return Object.keys(row).find(h => colKey(h) === k)
}
function getVal(row, name, def = null) {
  const actual = findColActual(row, name)
  return actual !== undefined ? row[actual] : def
}
function toNum(v, fallback = 0) {
  if (v === null || v === undefined || v === '') return fallback
  const n = parseFloat(String(v).replace(/,/g, ''))
  return isNaN(n) ? fallback : n
}
function normalizeId(v) { return String(v ?? '').trim().replace(/\.0+$/, '') }

function transformFinalStatus(billStatus) {
  if (String(billStatus ?? '').trim().toUpperCase() === 'Y') return 'Active'
  const now = new Date()
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
  return `Not Active as of ${now.getDate()} ${months[now.getMonth()]} Allocation`
}

function calcDisposition(hasLicense, interactions, linesAdded) {
  const lic = String(hasLicense ?? '').trim()
  const inter = toNum(interactions)
  const lines = toNum(linesAdded)
  if (lic !== 'Yes') return 'No Copilot License'
  if (inter >= 0 && inter <= 5) return 'Low user (Interaction 0-5)'
  if (inter >= 6 && inter <= 50) {
    if (lines < 100)  return 'Average user (Interaction 6-50 and Copilot Added Lines <100)'
    if (lines <= 500) return 'Average user (Interaction 6-50 and Copilot Added Lines 100\u2013500)'
    return 'Average user (Interaction 6-50 and Copilot Added Lines >500)'
  }
  if (inter > 50) {
    if (lines < 100)  return 'Power user (Interaction >50 and Copilot Added Lines <100)'
    if (lines <= 500) return 'Power user (Interaction >50 and Copilot Added Lines 100\u2013500)'
    return 'Power user (Interaction >50 and Copilot Added Lines >500)'
  }
  return 'No match'
}

function insertAfter(headers, newCol, afterCol) {
  if (headers.includes(newCol)) return headers
  const anchors = Array.isArray(afterCol) ? afterCol : [afterCol]
  for (const anchor of anchors) {
    const idx = headers.findIndex(h => colKey(h) === colKey(anchor))
    if (idx !== -1) { const result = [...headers]; result.splice(idx + 1, 0, newCol); return result }
  }
  return [...headers, newCol]
}
function applyInsRules(originalHeaders, rules, addedSet) {
  let headers = [...originalHeaders]
  for (const [newCol, afterCol] of rules) {
    if (addedSet.has(newCol)) headers = insertAfter(headers, newCol, afterCol)
  }
  return headers
}
function reorderRows(rows, headers) {
  return rows.map(row => { const out = {}; for (const col of headers) { out[col] = (col in row) ? row[col] : null }; return out })
}

// ── Constants ─────────────────────────────────────────────────────────────────
const JOB1_COLS = new Set([
  'Associate ID', 'Associate Name', 'Role Category', 'Service Line',
  'ACT/PCT Mapping', 'Final Copilot Lines', 'Copilot Usage > Interactions',
  'Final Git Lines', 'Final Copilot <50% of Final Git', 'Added Copilot <50% of Added Git',
  'Final Copilot <60% of Final Git', 'Added Copilot <60% of Added Git',
  'Copilot Training Complete', 'APEX Learning Assessment Completed?',
  'Either APEX or Copilot Training complete', 'CTS Copilot Activation',
  'Disposition parameter', 'Comments',
])
const INS_RULES_T1 = [
  ['Associate ID',                             'EID'],
  ['Associate Name',                           'Associate ID'],
  ['Role Category',                            'Associate Name'],
  ['Service Line',                             'Role Category'],
  ['ACT/PCT Mapping',                          'Service Line'],
  ['Final Copilot Lines',                      'Lines Deleted'],
  ['GHCP Pulse',                               'Lines Deleted'],
  ['Final Git Lines',                          'Deleted Lines'],
  ['GitHub Pulse',                             'Deleted Lines'],
  ['Copilot Usage > Interactions',             'Copilot Interactions'],
  ['Final Copilot <50% of Final Git',          'Final Git Lines'],
  ['Added Copilot <50% of Added Git',          'Final Copilot <50% of Final Git'],
  ['Final Copilot <60% of Final Git',          'Added Copilot <50% of Added Git'],
  ['Added Copilot <60% of Added Git',          'Final Copilot <60% of Final Git'],
  ['Copilot Training Complete',                ['Copilot Status', 'Copilot To Be Revoked', 'Current Copilot Status']],
  ['APEX Learning Assessment Completed?',      'Copilot Training Complete'],
  ['Either APEX or Copilot Training complete', 'APEX Learning Assessment Completed?'],
  ['CTS Copilot Activation',                   'Either APEX or Copilot Training complete'],
  ['Disposition parameter',                    'CTS Copilot Activation'],
  ['Comments',                                 'Disposition parameter'],
]
const INS_RULES_T2 = [
  ['Final Status',      'Associate Name'],
  ['Project ID',        'Service Line'],
  ['Project Name',      'Project ID'],
  ['Tower',             'Project Name'],
  ['Location (On/Off)', 'CTS Copilot Activation'],
]

// ── Task 1 ────────────────────────────────────────────────────────────────────
function processTask1(currentData, prevData, currentHeaders) {
  _log.plain('')
  _log.info(`Job 1 — EID Merge starting (${currentData.length} rows, ${prevData.length} prev rows)`)

  const origHeaders = (currentHeaders && currentHeaders.length) ? currentHeaders : Object.keys(currentData[0] || {})
  const hasEID = hasCol(currentData, 'EID')

  if (!hasEID) _log.warn('Job 1: "EID" column not found — EID matching skipped, formulas only.')
  else _log.info('Job 1: EID column found. Building lookup map…')

  const prevByEID = new Map()
  if (hasEID) {
    for (const row of prevData) { const eid = normalizeId(getVal(row, 'EID')); if (eid) prevByEID.set(eid, row) }
    _log.info(`Job 1: Lookup map built — ${prevByEID.size} unique EIDs in previous batch.`)
  }

  const COPY_MAP = [
    ['Associate ID',                        'Associate ID'],
    ['Associate Name',                      'Associate Name'],
    ['Role Category',                       'Role Category'],
    ['Service Line',                        'Service Line'],
    ['ACT/PCT Mapping',                     'ACT/PCT Mapping'],
    ['Copilot Training Complete',           'Copilot Training Complete'],
    ['APEX Learning Assessment Completed?', 'APEX Learning Assessment Completed?'],
    ['CTS Copilot Activation',              'CTS Copilot Activation'],
  ]
  const availCopyMap = []
  for (const [out, src] of COPY_MAP) {
    if (hasCol(prevData, src)) { availCopyMap.push({ out, src }); _log.info(`Job 1: "${out}" ← prev."${src}" — available`) }
    else _log.warn(`Job 1: prev column "${src}" not found — skipping "${out}"`)
  }

  const canCopilotLines = hasCol(currentData, 'Lines Added') && hasCol(currentData, 'Lines Deleted')
  const canGitLines     = hasCol(currentData, 'Added Lines') && hasCol(currentData, 'Deleted Lines')
  const canCopilotInter = hasCol(currentData, 'Copilot Usage') && hasCol(currentData, 'Copilot Interactions')
  const canDisposition  = hasCol(currentData, 'Has A Copilot License') && hasCol(currentData, 'Copilot Interactions') && hasCol(currentData, 'Lines Added')

  if (canCopilotLines) _log.info('Job 1: Will calculate "Final Copilot Lines" (Lines Added − Lines Deleted)')
  else _log.warn('Job 1: Skipping "Final Copilot Lines" — missing source columns')
  if (canGitLines) _log.info('Job 1: Will calculate "Final Git Lines" (Added Lines − Deleted Lines)')
  else _log.warn('Job 1: Skipping "Final Git Lines" — missing source columns')
  if (canCopilotInter) _log.info('Job 1: Will add "Copilot Usage > Interactions"')
  if (canCopilotLines && canGitLines) _log.info('Job 1: Will calculate 4 comparison columns')
  if (canDisposition) _log.info('Job 1: Will calculate "Disposition parameter"')
  else _log.warn('Job 1: Skipping "Disposition parameter" — missing required source columns')

  let matched = 0, unmatched = 0
  const processed = currentData.map(curr => {
    const row = { ...curr }
    if (hasEID) {
      const eid = normalizeId(getVal(curr, 'EID'))
      const prev = eid ? prevByEID.get(eid) : null
      if (prev) { matched++; for (const { out, src } of availCopyMap) { const val = getVal(prev, src); row[out] = (val === null || val === undefined || String(val).trim() === '') ? '#N/A' : val } }
      else { unmatched++; for (const { out } of availCopyMap) row[out] = '#N/A' }
    }
    row['Final Copilot Lines'] = canCopilotLines ? toNum(getVal(curr, 'Lines Added')) - toNum(getVal(curr, 'Lines Deleted')) : '#N/A'
    row['Final Git Lines']     = canGitLines     ? toNum(getVal(curr, 'Added Lines'))  - toNum(getVal(curr, 'Deleted Lines')) : '#N/A'
    if (canCopilotLines) row['GHCP Pulse']   = toNum(getVal(curr, 'Lines Added'))  - toNum(getVal(curr, 'Lines Deleted'))
    if (canGitLines)     row['GitHub Pulse'] = toNum(getVal(curr, 'Added Lines'))  - toNum(getVal(curr, 'Deleted Lines'))
    if (canCopilotInter) {
      const usage = getVal(curr, 'Copilot Usage'), inter = getVal(curr, 'Copilot Interactions')
      if (usage === null || inter === null || String(usage).trim() === '' || String(inter).trim() === '') row['Copilot Usage > Interactions'] = '#N/A'
      else row['Copilot Usage > Interactions'] = toNum(usage) > toNum(inter) ? 'TRUE' : 'FALSE'
    } else { row['Copilot Usage > Interactions'] = '#N/A' }

    const fcl = typeof row['Final Copilot Lines'] === 'number' ? row['Final Copilot Lines'] : null
    const fgl = typeof row['Final Git Lines'] === 'number' ? row['Final Git Lines'] : null
    const linesAdded = canCopilotLines ? toNum(getVal(curr, 'Lines Added')) : null
    const addedLines = canGitLines     ? toNum(getVal(curr, 'Added Lines')) : null
    row['Final Copilot <50% of Final Git'] = (fcl !== null && fgl !== null) ? (fcl < 0.5 * fgl ? 'TRUE' : 'FALSE') : '#N/A'
    row['Added Copilot <50% of Added Git'] = (linesAdded !== null && addedLines !== null) ? (linesAdded < 0.5 * addedLines ? 'TRUE' : 'FALSE') : '#N/A'
    row['Final Copilot <60% of Final Git'] = (fcl !== null && fgl !== null) ? (fcl < 0.6 * fgl ? 'TRUE' : 'FALSE') : '#N/A'
    row['Added Copilot <60% of Added Git'] = (linesAdded !== null && addedLines !== null) ? (linesAdded < 0.6 * addedLines ? 'TRUE' : 'FALSE') : '#N/A'

    const ct = row['Copilot Training Complete'], apx = row['APEX Learning Assessment Completed?']
    if (ct !== undefined || apx !== undefined) {
      const ctNA = ct === undefined || ct === '#N/A', apxNA = apx === undefined || apx === '#N/A'
      row['Either APEX or Copilot Training complete'] = (ctNA && apxNA) ? '#N/A' : (ct === 'Completed' || apx === 'Completed') ? 'Completed' : 'Not Completed'
    }
    row['Disposition parameter'] = canDisposition ? calcDisposition(getVal(curr, 'Has A Copilot License'), getVal(curr, 'Copilot Interactions', 0), getVal(curr, 'Lines Added', 0)) : '#N/A'
    row['Comments'] = ''
    return row
  })

  if (hasEID) _log.ok(`Job 1: EID matching — ${matched} matched, ${unmatched} unmatched`)

  const addedNew = new Set()
  if (hasEID) { for (const { out } of availCopyMap) { if (!origHeaders.find(h => colKey(h) === colKey(out))) addedNew.add(out) } }
  if (canCopilotLines && !origHeaders.find(h => colKey(h) === 'ghcp pulse'))   addedNew.add('GHCP Pulse')
  if (canGitLines     && !origHeaders.find(h => colKey(h) === 'github pulse')) addedNew.add('GitHub Pulse')
  const formulaCols = ['Final Copilot Lines','Final Git Lines','Copilot Usage > Interactions','Final Copilot <50% of Final Git','Added Copilot <50% of Added Git','Final Copilot <60% of Final Git','Added Copilot <60% of Added Git','Either APEX or Copilot Training complete','Disposition parameter','Comments']
  for (const col of formulaCols) { if (!origHeaders.find(h => colKey(h) === colKey(col))) addedNew.add(col) }

  let finalHeaders = applyInsRules(origHeaders, INS_RULES_T1, addedNew)
  const allKeys = new Set(processed.flatMap(r => Object.keys(r)))
  for (const k of allKeys) { if (!finalHeaders.find(h => colKey(h) === colKey(k))) finalHeaders.push(k) }
  const commIdx = finalHeaders.findIndex(h => colKey(h) === 'comments')
  if (commIdx !== -1 && commIdx !== finalHeaders.length - 1) finalHeaders.push(finalHeaders.splice(commIdx, 1)[0])

  const ordered = reorderRows(processed, finalHeaders)
  const allNewCols = new Set([...(hasEID ? availCopyMap.map(x => x.out) : []), ...formulaCols])
  _log.ok(`Job 1 complete — ${finalHeaders.length} columns × ${ordered.length} rows`)
  return { data: ordered, newCols: allNewCols, matched, unmatched }
}

// ── Task 2 ────────────────────────────────────────────────────────────────────
function processTask2(batchData, allocData, batchHeaders) {
  _log.plain('')
  _log.info(`Job 2 — Allocation Merge starting (${batchData.length} batch rows, ${allocData.length} alloc rows)`)

  if (!hasCol(batchData, 'Associate ID')) {
    const msg = 'Associate ID not found in Current Batch sheet. Please provide previous week batch sheet to generate the excel.'
    _log.err('Job 2: ' + msg)
    throw new Error(msg)
  }
  _log.info('Job 2: "Associate ID" column confirmed in batch.')

  const allocByID = new Map()
  for (const row of allocData) { const id = normalizeId(getVal(row, 'Associate ID')); if (id) allocByID.set(id, row) }
  _log.info(`Job 2: Allocation lookup map — ${allocByID.size} unique Associate IDs.`)

  const canProjID    = hasCol(allocData, 'Project ID')
  const canProjName  = hasCol(allocData, 'Project Description')
  const canFinalStat = hasCol(allocData, 'Billability Status')
  const canLocation  = hasCol(allocData, 'On/Off')
  const canTower     = hasCol(allocData, 'Tower')

  if (canProjID)    _log.info('Job 2: "Project ID" available from allocation')
  else              _log.warn('Job 2: "Project ID" not found in allocation — skipping')
  if (canProjName)  _log.info('Job 2: "Project Name" ← allocation."Project Description"')
  else              _log.warn('Job 2: "Project Description" not found — skipping "Project Name"')
  if (canFinalStat) _log.info('Job 2: "Final Status" ← transformed from "Billability Status"')
  else              _log.warn('Job 2: "Billability Status" not found — skipping "Final Status"')
  if (canLocation)  _log.info('Job 2: "Location" ← allocation."On/Off"')
  else              _log.warn('Job 2: "On/Off" not found — skipping "Location"')
  if (canTower)     _log.info('Job 2: "Tower" ← allocation."Tower"')
  else              _log.warn('Job 2: "Tower" not found — skipping')

  const origHeaders = (batchHeaders && batchHeaders.length) ? batchHeaders : Object.keys(batchData[0] || {})
  const canCopilotLinesT2 = hasCol(batchData, 'Lines Added') && hasCol(batchData, 'Lines Deleted')
  const canGitLinesT2     = hasCol(batchData, 'Added Lines') && hasCol(batchData, 'Deleted Lines')

  let matched = 0, unmatched = 0
  const processed = batchData.map(brow => {
    const row = { ...brow }
    const assocID = normalizeId(getVal(brow, 'Associate ID'))
    const alloc = assocID ? allocByID.get(assocID) : null
    if (alloc) {
      matched++
      if (canProjID)    row['Project ID']       = getVal(alloc, 'Project ID')          ?? '#N/A'
      if (canProjName)  row['Project Name']      = getVal(alloc, 'Project Description') ?? '#N/A'
      if (canFinalStat) row['Final Status']      = transformFinalStatus(getVal(alloc, 'Billability Status'))
      if (canLocation)  row['Location (On/Off)'] = getVal(alloc, 'On/Off')             ?? '#N/A'
      if (canTower)     row['Tower']             = getVal(alloc, 'Tower')               ?? '#N/A'
    } else {
      unmatched++
      if (canProjID)    row['Project ID']       = '#N/A'
      if (canProjName)  row['Project Name']      = '#N/A'
      if (canFinalStat) row['Final Status']      = transformFinalStatus('')
      if (canLocation)  row['Location (On/Off)'] = '#N/A'
      if (canTower)     row['Tower']             = '#N/A'
    }
    if (canCopilotLinesT2) row['GHCP Pulse']   = toNum(getVal(brow, 'Lines Added'))  - toNum(getVal(brow, 'Lines Deleted'))
    if (canGitLinesT2)     row['GitHub Pulse'] = toNum(getVal(brow, 'Added Lines'))  - toNum(getVal(brow, 'Deleted Lines'))
    return row
  })
  _log.ok(`Job 2: Allocation matching — ${matched} matched, ${unmatched} unmatched`)

  const addedNew = new Set()
  if (canProjID    && !origHeaders.find(h => colKey(h) === colKey('Project ID')))           addedNew.add('Project ID')
  if (canProjName  && !origHeaders.find(h => colKey(h) === colKey('Project Name')))          addedNew.add('Project Name')
  if (canFinalStat && !origHeaders.find(h => colKey(h) === colKey('Final Status')))          addedNew.add('Final Status')
  if (canLocation  && !origHeaders.find(h => colKey(h) === colKey('Location (On/Off)')))     addedNew.add('Location (On/Off)')
  if (canTower     && !origHeaders.find(h => colKey(h) === colKey('Tower')))                 addedNew.add('Tower')

  let finalHeaders = applyInsRules(origHeaders, INS_RULES_T2, addedNew)
  if (canCopilotLinesT2 && !finalHeaders.find(h => colKey(h) === 'ghcp pulse'))
    finalHeaders = insertAfter(finalHeaders, 'GHCP Pulse', 'Lines Deleted')
  if (canGitLinesT2 && !finalHeaders.find(h => colKey(h) === 'github pulse'))
    finalHeaders = insertAfter(finalHeaders, 'GitHub Pulse', 'Deleted Lines')
  const allKeys = new Set(processed.flatMap(r => Object.keys(r)))
  for (const k of allKeys) { if (!finalHeaders.find(h => colKey(h) === colKey(k))) finalHeaders.push(k) }

  const ordered = reorderRows(processed, finalHeaders)
  const allNewCols = new Set([
    ...(canProjID    ? ['Project ID']        : []),
    ...(canProjName  ? ['Project Name']      : []),
    ...(canFinalStat ? ['Final Status']      : []),
    ...(canLocation  ? ['Location (On/Off)'] : []),
    ...(canTower     ? ['Tower']             : []),
  ])
  _log.ok(`Job 2 complete — ${finalHeaders.length} columns × ${ordered.length} rows`)
  return { data: ordered, newCols: allNewCols, matched, unmatched }
}

// ── Excel writer ──────────────────────────────────────────────────────────────
function createWorksheetOnly(dataRows, newColsSet) {
  const ws = {}
  const outputHeaders = Object.keys(dataRows[0] || {})
  const numRows = dataRows.length
  const numCols = outputHeaders.length
  const clm = {}
  outputHeaders.forEach((h, i) => { clm[colKey(h)] = colLetter(i) })

  const FORMULA_DEFS = {
    'final copilot lines': r => (clm['lines added'] && clm['lines deleted']) ? `${clm['lines added']}${r}-${clm['lines deleted']}${r}` : null,
    'ghcp pulse':          r => (clm['lines added'] && clm['lines deleted']) ? `${clm['lines added']}${r}-${clm['lines deleted']}${r}` : null,
    'final git lines':     r => (clm['added lines'] && clm['deleted lines']) ? `${clm['added lines']}${r}-${clm['deleted lines']}${r}` : null,
    'github pulse':        r => (clm['added lines'] && clm['deleted lines']) ? `${clm['added lines']}${r}-${clm['deleted lines']}${r}` : null,
    'copilot usage > interactions': r => (clm['copilot usage'] && clm['copilot interactions']) ? `${clm['copilot usage']}${r}>${clm['copilot interactions']}${r}` : null,
    'final copilot <50% of final git': r => (clm['final copilot lines'] && clm['final git lines']) ? `${clm['final copilot lines']}${r}<0.5*${clm['final git lines']}${r}` : null,
    'added copilot <50% of added git': r => (clm['lines added'] && clm['added lines']) ? `${clm['lines added']}${r}<0.5*${clm['added lines']}${r}` : null,
    'final copilot <60% of final git': r => (clm['final copilot lines'] && clm['final git lines']) ? `${clm['final copilot lines']}${r}<0.6*${clm['final git lines']}${r}` : null,
    'added copilot <60% of added git': r => (clm['lines added'] && clm['added lines']) ? `${clm['lines added']}${r}<0.6*${clm['added lines']}${r}` : null,
    'either apex or copilot training complete': r => {
      const ct = clm['copilot training complete'], apx = clm['apex learning assessment completed?']
      if (!ct && !apx) return null
      const orPart = [ct && `${ct}${r}="Completed"`, apx && `${apx}${r}="Completed"`].filter(Boolean).join(',')
      return `IF(OR(${orPart}),"Completed","Not Completed")`
    },
    'disposition parameter': r => {
      const lic = clm['has a copilot license'], inter = clm['copilot interactions'], la = clm['lines added']
      if (!lic || !inter || !la) return null
      return `IF(${lic}${r}<>"Yes","No Copilot License",IF(AND(${inter}${r}>=0,${inter}${r}<=5),"Low user (Interaction 0-5)",IF(AND(${inter}${r}>=6,${inter}${r}<=50),IF(${la}${r}<100,"Average user (Interaction 6-50 and Copilot Added Lines <100)",IF(${la}${r}<=500,"Average user (Interaction 6-50 and Copilot Added Lines 100\u2013500)","Average user (Interaction 6-50 and Copilot Added Lines >500)")),IF(${la}${r}<100,"Power user (Interaction >50 and Copilot Added Lines <100)",IF(${la}${r}<=500,"Power user (Interaction >50 and Copilot Added Lines 100\u2013500)","Power user (Interaction >50 and Copilot Added Lines >500)")))))`
    },
  }

  const COL_COLORS = {
    'business name':'D0E0E3','name':'D0E0E3','eid':'D0E0E3','vzid':'D0E0E3','title':'D0E0E3','email':'D0E0E3','location':'D0E0E3','employee status':'D0E0E3','vendor name':'D0E0E3','worker id':'D0E0E3','work oder id':'D0E0E3','work order id':'D0E0E3','related sow':'D0E0E3',
    'tier 2 leader':'FFF2CC','tier 3 leader':'FFF2CC','tier 4 leader':'FFF2CC','tier 5 leader':'FFF2CC','tier 6 leader':'FFF2CC','tier 7 leader':'FFF2CC','tier 8 leader':'FFF2CC','tier 9 leader':'FFF2CC',
    'vendor paid subscription':'6FA8DC','has a copilot license':'6FA8DC','used agent mode':'6FA8DC','non agent mode acceptance rate':'6FA8DC','lines added':'6FA8DC','lines deleted':'6FA8DC','activation date':'6FA8DC','copilot usage':'6FA8DC','copilot interactions':'6FA8DC',
    'active gitlab account':'F9CB9C','gitlab commits':'F9CB9C','number of projects':'F9CB9C','added lines':'F9CB9C','deleted lines':'F9CB9C',
    'is a current employee':'F4CCCC','github id':'F4CCCC','github login':'F4CCCC','current copilot status':'F4CCCC',
  }
  const DISPLAY_NAMES = { 'Location (On/Off)': 'Location' }
  const hdrFont  = { name: 'Verizon NHG TX', sz: 12, bold: true, underline: true }
  const dataFont = { name: 'Verizon NHG TX', sz: 10 }
  const pinkFill = { patternType: 'solid', fgColor: { rgb: 'F2CEEF' } }
  const border   = { top:{style:'thin',color:{rgb:'D0D0D0'}}, bottom:{style:'thin',color:{rgb:'D0D0D0'}}, left:{style:'thin',color:{rgb:'D0D0D0'}}, right:{style:'thin',color:{rgb:'D0D0D0'}} }
  let lastRgb = 'FFFFFF'

  for (let c = 0; c < numCols; c++) {
    const colName = outputHeaders[c]
    const isNew   = newColsSet.has(colName)
    let fill
    if (isNew) { fill = pinkFill }
    else { const rgb = COL_COLORS[colKey(colName)]; if (rgb) lastRgb = rgb; fill = { patternType: 'solid', fgColor: { rgb: lastRgb } } }

    const displayName = DISPLAY_NAMES[colName] ?? colName
    const isComments  = colKey(colName) === 'comments'
    ws[XLSX.utils.encode_cell({ r: 0, c })] = {
      v: displayName, t: 's',
      s: isComments
        ? { font: { name:'Verizon NHG TX', sz:10 }, fill: { patternType:'solid', fgColor:{rgb:'FFFFFF'} }, alignment:{horizontal:'center',vertical:'center',wrapText:true}, border }
        : { font: hdrFont, fill, alignment:{horizontal:'center',vertical:'center',wrapText:true}, border }
    }
    const formulaDef = FORMULA_DEFS[colKey(colName)]
    for (let r = 1; r <= numRows; r++) {
      const val = dataRows[r - 1][colName]
      if (val === null || val === undefined || val === '') continue
      const formula = (formulaDef && val !== '#N/A') ? formulaDef(r + 1) : null
      if (formula) {
        const isBool = val === 'TRUE' || val === 'FALSE'
        const cellV  = isBool ? val === 'TRUE' : val
        const cellT  = isBool ? 'b' : typeof val === 'number' ? 'n' : 's'
        ws[XLSX.utils.encode_cell({ r, c })] = { v: cellV, t: cellT, f: formula, s: { font: dataFont } }
      } else {
        const t = typeof val === 'number' ? 'n' : val instanceof Date ? 'd' : 's'
        ws[XLSX.utils.encode_cell({ r, c })] = { v: val, t, s: { font: dataFont } }
      }
    }
  }
  ws['!ref']        = XLSX.utils.encode_range({ r:0, c:0 }, { r:numRows, c:numCols-1 })
  ws['!autofilter'] = { ref: ws['!ref'] }
  ws['!freeze']     = { xSplit: 0, ySplit: 1 }
  ws['!cols']       = outputHeaders.map(h => { const maxLen = Math.max(h.length, ...dataRows.slice(0,200).map(r => String(r[h]??'').length)); return { wch: Math.min(Math.max(maxLen+2,10),40) } })
  return ws
}

function createWorkbook(dataRows, newColsSet, _origWs, _sourceWb, targetSheetName = 'Copilot Gitlab Usage') {
  const ws = createWorksheetOnly(dataRows, newColsSet)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, targetSheetName)
  return wb
}

// ── Styles XML merger ─────────────────────────────────────────────────────────
function mergeStylesXml(origXml, newXml) {
  function extractSection(xml, tag) { const m = xml.match(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`)); return m ? m[1] : '' }
  function countElements(content, tag) { return (content.match(new RegExp(`<${tag}(?:[\\s>\/])`, 'g')) || []).length }
  function getCountAttr(xml, tag) { const m = xml.match(new RegExp(`<${tag}[^>]+count="(\\d+)"`)); return m ? parseInt(m[1]) : null }
  function appendToSection(xml, tag, content) { const idx = xml.indexOf(`</${tag}>`); return idx === -1 ? xml : xml.slice(0, idx) + content + xml.slice(idx) }
  function updateCountAttr(xml, tag, delta) { return xml.replace(new RegExp(`(<${tag}[^>]+count=")([0-9]+)(")`), (m, p1, n, p3) => `${p1}${parseInt(n)+delta}${p3}`) }

  const origFontsContent   = extractSection(origXml, 'fonts')
  const origFillsContent   = extractSection(origXml, 'fills')
  const origBordersContent = extractSection(origXml, 'borders')
  const origXfsContent     = extractSection(origXml, 'cellXfs')
  const newFontsContent    = extractSection(newXml, 'fonts')
  const newFillsContent    = extractSection(newXml, 'fills')
  const newBordersContent  = extractSection(newXml, 'borders')
  let   newXfsContent      = extractSection(newXml, 'cellXfs')

  const origFontCount   = getCountAttr(origXml,'fonts')   ?? countElements(origFontsContent,'font')
  const origFillCount   = getCountAttr(origXml,'fills')   ?? countElements(origFillsContent,'fill')
  const origBorderCount = getCountAttr(origXml,'borders') ?? countElements(origBordersContent,'border')
  const origXfsCount    = getCountAttr(origXml,'cellXfs') ?? countElements(origXfsContent,'xf')
  const newFontCount    = getCountAttr(newXml,'fonts')    ?? countElements(newFontsContent,'font')
  const newFillCount    = getCountAttr(newXml,'fills')    ?? countElements(newFillsContent,'fill')
  const newBorderCount  = getCountAttr(newXml,'borders')  ?? countElements(newBordersContent,'border')
  const newXfsCount     = getCountAttr(newXml,'cellXfs')  ?? countElements(newXfsContent,'xf')

  newXfsContent = newXfsContent
    .replace(/fontId="(\d+)"/g,   (_, id) => `fontId="${parseInt(id)+origFontCount}"`)
    .replace(/fillId="(\d+)"/g,   (_, id) => `fillId="${parseInt(id)+origFillCount}"`)
    .replace(/borderId="(\d+)"/g, (_, id) => `borderId="${parseInt(id)+origBorderCount}"`)

  let merged = origXml
  if (newFontCount   > 0) { merged = appendToSection(merged,'fonts',newFontsContent);     merged = updateCountAttr(merged,'fonts',newFontCount) }
  if (newFillCount   > 0) { merged = appendToSection(merged,'fills',newFillsContent);     merged = updateCountAttr(merged,'fills',newFillCount) }
  if (newBorderCount > 0) { merged = appendToSection(merged,'borders',newBordersContent); merged = updateCountAttr(merged,'borders',newBorderCount) }
  if (newXfsCount    > 0) { merged = appendToSection(merged,'cellXfs',newXfsContent);     merged = updateCountAttr(merged,'cellXfs',newXfsCount) }
  return { mergedStylesXml: merged, xfsOffset: origXfsCount }
}

// ── ZIP surgery ───────────────────────────────────────────────────────────────
async function buildOutputZip(processedWs, rawBuffer, targetSheetName) {
  const miniWb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(miniWb, processedWs, targetSheetName)
  const miniBytes = XLSX.write(miniWb, { bookType: 'xlsx', type: 'array', cellStyles: true, bookSST: false })
  const [origZip, miniZip] = await Promise.all([JSZip.loadAsync(rawBuffer), JSZip.loadAsync(miniBytes)])

  const origWbXml  = await origZip.file('xl/workbook.xml').async('string')
  const origWbRels = await origZip.file('xl/_rels/workbook.xml.rels').async('string')
  const escapedName = targetSheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const sheetNodeRe  = new RegExp(`<sheet[^>]+name="${escapedName}"[^>]+r:id="([^"]+)"`)
  const sheetNodeRe2 = new RegExp(`<sheet[^>]+r:id="([^"]+)"[^>]+name="${escapedName}"`)
  const sheetNodeMatch = origWbXml.match(sheetNodeRe) || origWbXml.match(sheetNodeRe2)

  let sheetTarget
  if (sheetNodeMatch) {
    const rId = sheetNodeMatch[1]
    const relRe  = new RegExp(`<Relationship[^>]+Id="${rId}"[^>]+Target="([^"]+)"`)
    const relRe2 = new RegExp(`<Relationship[^>]+Target="([^"]+)"[^>]+Id="${rId}"`)
    const relMatch = origWbRels.match(relRe) || origWbRels.match(relRe2)
    if (!relMatch) throw new Error(`Relationship ${rId} not found in workbook.xml.rels`)
    sheetTarget = relMatch[1]
  } else {
    const firstRel = origWbRels.match(/<Relationship[^>]+Target="(worksheets\/[^"]+)"/)
    if (!firstRel) throw new Error('No worksheet found in original file')
    sheetTarget = firstRel[1]
  }
  const sheetPath = sheetTarget.startsWith('/') ? sheetTarget.slice(1) : `xl/${sheetTarget}`

  const miniSheetFile = miniZip.file('xl/worksheets/sheet1.xml')
  if (!miniSheetFile) throw new Error('Mini workbook sheet XML not found')
  let newSheetXml = await miniSheetFile.async('string')

  const origStylesFile = origZip.file('xl/styles.xml')
  const miniStylesFile = miniZip.file('xl/styles.xml')
  if (origStylesFile && miniStylesFile) {
    const origStylesXml = await origStylesFile.async('string')
    const miniStylesXml = await miniStylesFile.async('string')
    const { mergedStylesXml, xfsOffset } = mergeStylesXml(origStylesXml, miniStylesXml)
    if (xfsOffset > 0) newSheetXml = newSheetXml.replace(/ s="(\d+)"/g, (m, idx) => ` s="${parseInt(idx)+xfsOffset}"`)
    origZip.file('xl/styles.xml', mergedStylesXml)
  }

  const wbXmlFile = origZip.file('xl/workbook.xml')
  if (wbXmlFile) {
    let wbXml = await wbXmlFile.async('string')
    if (/<calcPr/i.test(wbXml)) {
      wbXml = wbXml.replace(/(<calcPr[^>]*)\bcalcMode="[^"]*"/i,'$1calcMode="manual"').replace(/(<calcPr[^>]*)\bfullCalcOnLoad="[^"]*"/i,'$1')
      if (!/calcMode="manual"/i.test(wbXml)) wbXml = wbXml.replace(/(<calcPr)/i,'$1 calcMode="manual"')
    } else { wbXml = wbXml.replace(/<\/workbook>/i,'<calcPr calcMode="manual"/></workbook>') }
    origZip.file('xl/workbook.xml', wbXml)
  }

  origZip.file(sheetPath, newSheetXml)
  if (origZip.file('xl/calcChain.xml')) {
    origZip.remove('xl/calcChain.xml')
    const relsFile = origZip.file('xl/_rels/workbook.xml.rels')
    if (relsFile) { const relsXml = await relsFile.async('string'); origZip.file('xl/_rels/workbook.xml.rels', relsXml.replace(/<Relationship[^>]+calcChain[^>]*\/>/g,'')) }
    const ctFile = origZip.file('[Content_Types].xml')
    if (ctFile) { const ctXml = await ctFile.async('string'); origZip.file('[Content_Types].xml', ctXml.replace(/<Override[^>]+calcChain[^>]*\/>/gi,'')) }
  }
  return origZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

// ── Download helpers ──────────────────────────────────────────────────────────
function triggerDataUrlDownload(blob, filename) {
  const reader = new FileReader()
  reader.onload = () => { const a = document.createElement('a'); a.href = reader.result; a.download = filename; a.rel = 'noopener'; document.body.appendChild(a); a.click(); document.body.removeChild(a) }
  reader.onerror = () => console.error('Download encode failed: ' + filename)
  reader.readAsDataURL(blob)
}
function downloadWorkbook(wb, filename) {
  const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true })
  triggerDataUrlDownload(new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename)
}
function downloadBytes(bytes, filename) {
  triggerDataUrlDownload(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename)
}

// ── Scenario detection ────────────────────────────────────────────────────────
function detectScenario(files) {
  const { current, previous, allocation } = files
  if (!current) return { id:0, label:'Waiting for files…', desc:'Upload the Current Week Batch Sheet (required) + at least one additional file.', valid:false, pills:[] }
  if (!previous && !allocation) return { id:4, label:'Waiting for More Files', desc:'Please upload at least one additional file (Previous Week Batch or Allocation Sheet).', valid:false, pills:[{lbl:'⚠ Job 1: No previous week batch',cls:'pill-err'},{lbl:'⚠ Job 2: No allocation sheet',cls:'pill-err'}] }
  if (previous && allocation)   return { id:1, label:'Full Batch Processing', desc:'Will run Job 1 (EID merge), then Job 2 (Associate ID merge).', output:'New Batch - CTS - Copilot Gitlab Usage Weekly.xlsx', valid:true, pills:[{lbl:'▶ Job 1: EID Merge',cls:'pill-run'},{lbl:'▶ Job 2: Allocation Merge',cls:'pill-run'}] }
  if (previous && !allocation)  return { id:2, label:'Job 1 — Previous Batch Merge', hint:'Upload the Allocation file to run Full Batch Processing', desc:'Will run Job 1 only (EID merge).', output:'Job1_Only – CTS - Copilot Gitlab Usage Weekly.xlsx', valid:true, pills:[{lbl:'▶ Job 1: EID Merge',cls:'pill-run'},{lbl:'— Job 2: Skipped (no allocation)',cls:'pill-skip'}] }
  if (!previous && allocation)  return { id:3, label:'Job 2 — Allocation Merge', hint:'Upload the Previous Week Batch file to run Full Batch Processing', desc:'Will attempt Job 2 only. Requires "Associate ID" column in the current batch (added by Job 1 in a prior run).', output:'New Batch - CTS - Copilot Gitlab Usage Weekly.xlsx', valid:true, pills:[{lbl:'— Job 1: Skipped (no prev batch)',cls:'pill-skip'},{lbl:'▶ Job 2: Allocation Merge',cls:'pill-run'}] }
  return { id:0, label:'Waiting…', desc:'', valid:false, pills:[] }
}

function tick() { return new Promise(r => setTimeout(r, 0)) }

// ══════════════════════════════════════════════════════════════════════════════
export default function BatchFlow() {
  // ── Refs (data used in async runs) ────────────────────────────────────────
  const filesRef       = useRef({ current: null, previous: null, allocation: null })
  const dataRef        = useRef({ current: null, previous: null, allocation: null })
  const wsRef          = useRef({ current: null })
  const wbRef          = useRef({ current: null })
  const rawBufferRef   = useRef({ current: null })
  const headersRef     = useRef({ current: [] })
  const pendingDlRef   = useRef([])
  const toastTimerRef  = useRef(null)
  const logBodyRef     = useRef(null)

  // ── State ──────────────────────────────────────────────────────────────────
  const [filesState,    setFilesState]    = useState({ current: null, previous: null, allocation: null })
  const [cardStates,    setCardStates]    = useState({ current: 'idle', previous: 'idle', allocation: 'idle' })
  const [fileInfos,     setFileInfos]     = useState({ current: null, previous: null, allocation: null })
  const [processing,    setProcessing]    = useState(false)
  const [progressState, setProgressState] = useState({ show: false, pct: 0, label: '' })
  const [logs,          setLogs]          = useState([])
  const [results,       setResults]       = useState(null)
  const [showInfoModal, setShowInfoModal] = useState(false)
  const [errorToast,    setErrorToast]    = useState(null)

  // ── Log scroll ─────────────────────────────────────────────────────────────
  useEffect(() => {
    if (logBodyRef.current) logBodyRef.current.scrollTop = logBodyRef.current.scrollHeight
  }, [logs])

  // ── Init logs ──────────────────────────────────────────────────────────────
  useEffect(() => {
    const ts = () => new Date().toLocaleTimeString('en-GB', { hour12: false })
    setLogs([
      { level:'PLAIN', msg:'Application ready. Waiting for files…', ts: ts(), id: 1 },
      { level:'INFO',  msg:'  ⓘ  Zero data transmission — your files are read and processed entirely within this browser.', ts: ts(), id: 2 },
      { level:'INFO',  msg:'  ⓘ  Supported: .xlsx and .xls files. Max tested: 5,000+ rows.', ts: ts(), id: 3 },
    ])
    try {
      const blob = new Blob([new Uint8Array([0])])
      const r = new FileReader()
      r.onload = () => setLogs(p => [...p, { level:'PLAIN', msg:'  ✔  FileReader API: OK', ts: ts(), id: Date.now() }])
      r.onerror = () => setLogs(p => [...p, { level:'ERROR', msg:'  ✖  FileReader API blocked — file processing will fail. Check browser/OS security settings.', ts: ts(), id: Date.now() }])
      r.readAsArrayBuffer(blob)
    } catch(e) {
      setLogs(p => [...p, { level:'ERROR', msg:'  ✖  FileReader API unavailable: ' + e.message, ts: ts(), id: Date.now() }])
    }
  }, [])

  // ── Log helper ─────────────────────────────────────────────────────────────
  function addLog(level, msg) {
    const ts = new Date().toLocaleTimeString('en-GB', { hour12: false })
    setLogs(p => [...p, { level, msg, ts, id: Date.now() + Math.random() }])
  }

  // ── File handling ──────────────────────────────────────────────────────────
  function setFile(key, file) {
    const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase()
    if (!['.xlsx', '.xls'].includes(ext)) {
      setCardStates(p => ({ ...p, [key]: 'err' }))
      setFileInfos(p => ({ ...p, [key]: { err: `Invalid file type: ${ext}. Please upload .xlsx or .xls` } }))
      filesRef.current[key] = null
      dataRef.current[key] = null
      setFilesState(p => ({ ...p, [key]: null }))
      return
    }
    const sizeStr = file.size > 1024*1024 ? `${(file.size/1024/1024).toFixed(1)} MB` : `${Math.round(file.size/1024)} KB`
    setCardStates(p => ({ ...p, [key]: 'ok' }))
    setFileInfos(p => ({ ...p, [key]: { name: file.name, size: sizeStr } }))
    filesRef.current[key] = file
    if (key === 'current') { dataRef.current.current = null; wsRef.current.current = null; wbRef.current.current = null; rawBufferRef.current.current = null }
    dataRef.current[key] = null
    setFilesState(p => ({ ...p, [key]: file }))
    addLog('INFO', `  ⓘ  File set [${key}]: ${file.name} (${sizeStr})`)
  }

  // ── Error toast ────────────────────────────────────────────────────────────
  function showErrToast(msg) {
    setErrorToast(msg)
    clearTimeout(toastTimerRef.current)
    toastTimerRef.current = setTimeout(() => setErrorToast(null), 5000)
  }

  // ── Main process ───────────────────────────────────────────────────────────
  async function runProcess() {
    const sc = detectScenario(filesState)
    if (!sc.valid || processing) return
    setProcessing(true)
    pendingDlRef.current = []
    setResults(null)
    setProgressState({ show: true, pct: 0, label: 'Initializing…' })

    _log = {
      plain: m => addLog('PLAIN', m),
      info:  m => addLog('INFO', '  ⓘ  ' + m),
      ok:    m => addLog('SUCCESS', '  ✔  ' + m),
      warn:  m => addLog('WARN', '  ⚠  ' + m),
      err:   m => addLog('ERROR', '  ✖  ' + m),
    }

    try {
      const setP = (pct, label) => setProgressState({ show: true, pct, label })
      setP(5, 'Reading files…')
      addLog('PLAIN', '')
      addLog('INFO', `  ⓘ  Scenario ${sc.id}: ${sc.label}`)

      if (!dataRef.current.current) {
        addLog('INFO', '  ⓘ  Reading current batch file…')
        const r = await readExcelFile(filesRef.current.current, 'Copilot Gitlab Usage')
        dataRef.current.current = r.data; wsRef.current.current = r.ws
        headersRef.current.current = r.headers; wbRef.current.current = r.wb
        rawBufferRef.current.current = r.rawBuffer
        addLog('SUCCESS', `  ✔  Current batch: ${r.data.length} rows, ${r.headers.length} columns (sheet: "${r.sheetName}")`)
      }
      if (filesRef.current.previous && !dataRef.current.previous) {
        addLog('INFO', '  ⓘ  Reading previous batch file…')
        setP(15, 'Reading previous batch…')
        const rp = await readExcelFile(filesRef.current.previous, 'Copilot Gitlab Usage')
        dataRef.current.previous = rp.data
        addLog('SUCCESS', `  ✔  Previous batch: ${rp.data.length} rows (sheet: "${rp.sheetName}")`)
      }
      if (filesRef.current.allocation && !dataRef.current.allocation) {
        addLog('INFO', '  ⓘ  Reading allocation file…')
        setP(20, 'Reading allocation sheet…')
        const ra = await readExcelFile(filesRef.current.allocation, 'Verizon VDSI & Business')
        dataRef.current.allocation = ra.data
        addLog('SUCCESS', `  ✔  Allocation: ${ra.data.length} rows (sheet: "${ra.sheetName}")`)
      }

      let workingData = dataRef.current.current
      let allNewCols  = new Set()
      let t1Stats = null, t2Stats = null

      if (sc.id === 1 || sc.id === 2) {
        setP(35, 'Running Job 1 — EID merge…')
        await tick()
        const t1 = processTask1(dataRef.current.current, dataRef.current.previous, headersRef.current.current)
        workingData = t1.data; t1.newCols.forEach(c => allNewCols.add(c)); t1Stats = t1
        setP(60, 'Job 1 complete…')
      }
      if (sc.id === 3) { for (const col of JOB1_COLS) { if (hasCol(workingData, col)) allNewCols.add(col) } }
      if (sc.id === 1 || sc.id === 3) {
        setP(70, 'Running Job 2 — Allocation merge…')
        await tick()
        const t2 = processTask2(workingData, dataRef.current.allocation, sc.id === 3 ? headersRef.current.current : null)
        workingData = t2.data; t2.newCols.forEach(c => allNewCols.add(c)); t2Stats = t2
        setP(88, 'Job 2 complete…')
      }

      setP(92, 'Generating Excel file…')
      await tick()
      const filename = sc.id === 2 ? 'Job1_Only \u2013 CTS - Copilot Gitlab Usage Weekly.xlsx' : 'New Batch - CTS - Copilot Gitlab Usage Weekly.xlsx'
      const finalCols = Object.keys(workingData[0] || {})

      let outputBytes = null, outputWb = null
      if (rawBufferRef.current.current) {
        try {
          addLog('INFO', '  ⓘ  Applying ZIP surgery to preserve extra sheet formatting…')
          setP(94, 'Building ZIP…')
          const processedWs = createWorksheetOnly(workingData, allNewCols)
          outputBytes = await buildOutputZip(processedWs, rawBufferRef.current.current, 'Copilot Gitlab Usage')
          addLog('SUCCESS', '  ✔  ZIP surgery complete — extra sheet formatting & formulas preserved.')
        } catch (zipErr) {
          addLog('WARN', `  ⚠  ZIP surgery failed (${zipErr.message}) — falling back to standard output.`)
          outputBytes = null
        }
      }
      if (!outputBytes) {
        outputWb = createWorkbook(workingData, allNewCols, wsRef.current.current, wbRef.current.current, 'Copilot Gitlab Usage')
      }

      pendingDlRef.current = [{ wb: outputWb, bytes: outputBytes, filename, rows: workingData.length, cols: finalCols.length }]

      const statsArr = [
        { label: 'Total Rows',    value: workingData.length.toLocaleString() },
        { label: 'Total Columns', value: finalCols.length.toLocaleString() },
        { label: 'New Columns',   value: allNewCols.size },
      ]
      if (t1Stats) statsArr.push({ label: 'J1 EID Matched', value: t1Stats.matched.toLocaleString() })
      if (t2Stats) statsArr.push({ label: 'J2 ID Matched',  value: t2Stats.matched.toLocaleString() })

      setP(100, 'Done!')
      addLog('PLAIN', '')
      addLog('SUCCESS', `━━━━ Processing complete! Output: ${filename} ━━━━`)
      addLog('SUCCESS', `Rows: ${workingData.length} | Columns: ${finalCols.length} | New columns added: ${allNewCols.size}`)
      if (allNewCols.size > 0) addLog('SUCCESS', `New cols: ${[...allNewCols].join(', ')}`)

      setResults({ downloads: [{ filename, rows: workingData.length, cols: finalCols.length }], stats: statsArr })

    } catch (err) {
      addLog('ERROR', '  ✖  Error: ' + err.message)
      console.error(err)
      setProgressState(p => ({ ...p, pct: 0, label: 'Error occurred' }))
      showErrToast(err.message)
    } finally {
      _log = null
      setProcessing(false)
      setTimeout(() => setProgressState(p => ({ ...p, show: false })), 1500)
    }
  }

  function downloadEntry(idx) {
    const entry = pendingDlRef.current[idx]
    if (!entry) return
    if (entry.bytes) downloadBytes(entry.bytes, entry.filename)
    else downloadWorkbook(entry.wb, entry.filename)
    addLog('SUCCESS', '  ✔  Downloaded: ' + entry.filename)
  }

  // ── Computed ───────────────────────────────────────────────────────────────
  const sc = detectScenario(filesState)
  const SC_ICONS = { 0: '⏳', 1: '✅', 2: '✅', 3: '⚠️', 4: '❌' }

  const CARD_CONFIGS = [
    { key: 'current',    req: true,  ico: '📋', icoClass: 'ico-req', tag: 'Required', title: 'Current Week Batch Sheet',  desc: 'Weekly associate snapshot' },
    { key: 'previous',   req: false, ico: '🔃', icoClass: 'ico-opt', tag: 'Optional', title: 'Previous Week Batch Sheet', desc: 'Enables Job 1 (EID merge)' },
    { key: 'allocation', req: false, ico: '👥', icoClass: 'ico-opt', tag: 'Optional', title: 'Allocation Sheet',           desc: 'Enables Job 2 (Associate ID merge)' },
  ]

  // ── Render ─────────────────────────────────────────────────────────────────
  return (
    <div className="bf-root">
      {/* Info Modal */}
      <div className={`info-overlay${showInfoModal ? ' open' : ''}`} onClick={e => { if (e.target === e.currentTarget) setShowInfoModal(false) }}>
        <div className="info-modal">
          <div className="info-modal-header">
            <div className="info-modal-title">What each job does</div>
            <button className="info-modal-close" onClick={() => setShowInfoModal(false)}>✕</button>
          </div>
          <div className="info-job-grid">
            <div className="info-job-card job1">
              <div className="info-job-label">Job 1</div>
              <div className="info-job-name">Previous Batch Merge</div>
              <ul className="info-job-points">
                <li>Matches current week rows against the previous batch by EID</li>
                <li>Calculates Copilot and Git productivity metrics</li>
                <li>Adds usage comparisons and training status columns</li>
              </ul>
            </div>
            <div className="info-job-card job2">
              <div className="info-job-label">Job 2</div>
              <div className="info-job-name">Allocation Merge</div>
              <ul className="info-job-points">
                <li>Matches rows by Associate ID against the Allocation sheet</li>
                <li>Sets Final Status for each associate</li>
                <li>Adds project and location details</li>
              </ul>
              <div className="info-job2-warn">⚠ Requires the Associate ID column, which is only present after Job 1 has been run at least once on this batch.</div>
            </div>
          </div>
          <div className="info-combined">
            <div className="info-combined-ico">🔗</div>
            <div className="info-combined-body">
              <div className="info-combined-title">When both jobs run together</div>
              <div className="info-combined-desc">Job 1 runs first — it enriches the batch with EID-matched data and computed metrics. Job 2 then picks up that enriched output and merges it with the Allocation sheet. The final file contains all columns from both jobs in a single pass.</div>
            </div>
          </div>
          <div className="info-warning-box">
            <div className="info-combined-ico">🚫</div>
            <div className="info-combined-body">
              <div className="info-combined-title">Please Preserve the Output Column Names &amp; Structure</div>
              <ul className="info-job-points" style={{ marginTop: '2px' }}>
                <li>The output Excel is used as the input for the next week's batch — renaming any column may break matching in future runs and cause incorrect or incomplete results.</li>
                <li>If a column name needs to be changed, please reach out to the development team to ensure the tool is updated accordingly.</li>
              </ul>
            </div>
          </div>
        </div>
      </div>

      {/* Error Toast */}
      {errorToast && (
        <div className="error-toast show">
          <div className="error-toast-ico">✖</div>
          <div className="error-toast-msg">{errorToast}</div>
          <button className="error-toast-close" onClick={() => setErrorToast(null)}>✕</button>
        </div>
      )}

      <main className="bf-main">
        {/* Upload */}
        <div className="section-label">🗂️ Upload Files</div>
        <div className="upload-grid">
          {CARD_CONFIGS.map(({ key, req, ico, icoClass, tag, title, desc }) => {
            const state = cardStates[key]
            const info  = fileInfos[key]
            const cardClass = `upload-card${state === 'ok' ? ' file-ok' : state === 'err' ? ' file-err' : state === 'drag' ? ' drag-over' : ''}`
            return (
              <div key={key} className={cardClass}
                onDragOver={e => { e.preventDefault(); setCardStates(p => ({ ...p, [key]: 'drag' })) }}
                onDragLeave={() => setCardStates(p => ({ ...p, [key]: filesState[key] ? 'ok' : 'idle' }))}
                onDrop={e => { e.preventDefault(); setCardStates(p => ({ ...p, [key]: filesState[key] ? 'ok' : 'idle' })); const f = e.dataTransfer?.files?.[0]; if (f) setFile(key, f) }}
              >
                <input type="file" accept=".xlsx,.xls"
                  style={{ position:'absolute', inset:0, opacity:0, cursor:'pointer', width:'100%', height:'100%' }}
                  onChange={e => { const f = e.target.files?.[0]; if (f) setFile(key, f) }} />
                <div className="card-top">
                  <div className={`card-ico ${icoClass}`}>{ico}</div>
                  <div>
                    <span className={`req-tag ${req ? 'tag-req' : 'tag-opt'}`}>{tag}</span>
                    <div className="card-ttl">{title}</div>
                    <div className="card-desc">{desc}</div>
                  </div>
                </div>
                {state === 'ok' && info && !info.err && (
                  <div className="upload-status ok show">
                    <span>✔ {info.name}</span>
                    <span className="file-size-tag">{info.size}</span>
                  </div>
                )}
                {state === 'err' && info?.err && (
                  <div className="upload-status err show"><span>⚠ {info.err}</span></div>
                )}
                {(state === 'idle' || state === 'drag') && (
                  <div className="upload-hint">
                    <svg width="13" height="13" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
                    Click or drag to upload .xlsx
                  </div>
                )}
              </div>
            )
          })}
        </div>

        {/* Scenario + Process */}
        <div className="scenario-card">
          <div className="scenario-body">
            <div className="scenario-ico">{SC_ICONS[sc.id] ?? '⏳'}</div>
            <div className="scenario-info">
              <div className="scenario-title">{sc.label}</div>
              <div className="scenario-desc">
                {sc.id === 0 && <span style={{ color: '#e3d2c4', fontWeight: 500 }}>{sc.desc}</span>}
                {sc.id === 4 && <span style={{ color: '#b45309', fontWeight: 500 }}>{sc.desc}</span>}
                {sc.id !== 0 && sc.id !== 4 && sc.desc}
                {sc.output && <div style={{ marginTop: 5, fontWeight: 600, color: '#3fb950' }}>Output: {sc.output}</div>}
                {sc.hint   && <div style={{ fontSize: 12, fontWeight: 500, color: '#b45309', fontStyle: 'italic', marginTop: 4 }}>{sc.hint}</div>}
              </div>
              <div className="task-pills">
                {sc.pills.map((p, i) => <span key={i} className={`task-pill ${p.cls}`}>{p.lbl}</span>)}
              </div>
            </div>
            <button className="btn-info-toggle" style={{ width:36, height:36, borderRadius:'50%', background:'rgba(255,255,255,0.07)', border:'1px solid var(--border)', cursor:'pointer', fontSize:17, color:'var(--text-muted)', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0, transition:'background .2s' }}
              onClick={() => setShowInfoModal(true)} title="Jobs at a Glance">⚙️</button>
          </div>
          <div style={{ height: 16 }} />
          <button className={`btn-process${processing ? ' running' : ''}`} disabled={!sc.valid || processing} onClick={runProcess}>
            {processing
              ? <><span className="spinner" /><span>Processing…</span></>
              : <><svg width="17" height="17" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M13 10V3L4 14h7v7l9-11h-7z"/></svg><span>Run BatchFlow</span></>
            }
          </button>
          {progressState.show && (
            <div className="progress-wrap show">
              <div className="progress-meta">
                <span>{progressState.label}</span>
                <span className="pct">{progressState.pct}%</span>
              </div>
              <div className="progress-track">
                <div className="progress-fill" style={{ width: progressState.pct + '%' }} />
              </div>
            </div>
          )}
        </div>

        {/* Results */}
        {results && (
          <div className="results-card show">
            <div className="results-top">
              <div className="result-success-ico">✅</div>
              <div>
                <div className="result-title">Processing Complete!</div>
                <div className="result-sub">{results.downloads.length} file(s) ready to download.</div>
              </div>
            </div>
            <div className="stats-row">
              {results.stats.map((s, i) => (
                <div key={i} className="stat-box">
                  <div className="stat-num">{s.value}</div>
                  <div className="stat-lbl">{s.label}</div>
                </div>
              ))}
            </div>
            <div className="dl-grid">
              {results.downloads.map((d, i) => (
                <div key={i} className="dl-card">
                  <div className="dl-file-ico">📗</div>
                  <div className="dl-info">
                    <div className="dl-fname">{d.filename}</div>
                    <div className="dl-fmeta">{d.rows.toLocaleString()} rows × {d.cols.toLocaleString()} columns · XLSX</div>
                  </div>
                  <button className="btn-dl" onClick={() => downloadEntry(i)}>⬇ Download</button>
                </div>
              ))}
            </div>
            <div className="pink-note">
              <span className="pink-swatch" />
              <span>Columns added by this tool are highlighted with pink headers in the Excel output.</span>
            </div>
          </div>
        )}

        {/* Activity Log */}
        <div className="log-card">
          <div className="log-header">
            <div style={{ display:'flex', alignItems:'center', gap:10 }}>
              <div className="log-dots">
                <div className="log-dot dot-r" /><div className="log-dot dot-y" /><div className="log-dot dot-g" />
              </div>
              <span className="log-title">Activity Log</span>
            </div>
            <div className="log-actions">
              <button className="btn-log-action" onClick={() => {
                const ts = new Date().toLocaleTimeString('en-GB', { hour12: false })
                setLogs([{ level:'PLAIN', msg:'Log cleared.', ts, id: Date.now() }])
              }}>Clear</button>
              <button className="btn-log-action" onClick={() => {
                const lines = logs.map(l => l.msg).join('\n')
                navigator.clipboard?.writeText(lines).then(() => {
                  const ts = new Date().toLocaleTimeString('en-GB', { hour12: false })
                  setLogs(p => [...p, { level:'SUCCESS', msg:'  ✔  Log copied to clipboard.', ts, id: Date.now() }])
                })
              }}>Copy</button>
            </div>
          </div>
          <div className="log-body" ref={logBodyRef}>
            {logs.map(entry => (
              <div key={entry.id} className={`log-line ${entry.level}`}>
                <span className="log-ts">{entry.ts}</span>
                <span className="log-txt">{entry.msg}</span>
              </div>
            ))}
          </div>
        </div>

        <footer className="app-footer">
          <strong>BatchFlow Manager</strong> · Cognizant × Verizon · Zero data transmission — your files are read and processed entirely within this browser.
        </footer>
      </main>
    </div>
  )
}

// ── readExcelFile (defined after component to keep it out of the render scope) ─
function readExcelFile(file, targetSheetName = null) {
  return new Promise((resolve, reject) => {
    if (!file || !(file instanceof Blob)) { reject(new Error('Invalid file reference. Please re-select the file and try again.')); return }
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const arr = new Uint8Array(e.target.result)
        const wb = XLSX.read(arr, { type: 'array', raw: false, cellDates: true })
        const sheetName = (targetSheetName && wb.SheetNames.includes(targetSheetName)) ? targetSheetName : wb.SheetNames[0]
        const ws = wb.Sheets[sheetName]
        if (!ws) throw new Error('No worksheet found in file.')
        const HEADER_ANCHORS = ['vzid', 'name', 'email']
        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false })
        let headerRowIdx = 0
        for (let i = 0; i < allRows.length; i++) {
          const row = allRows[i]
          const found = row.some(cell => { const v = String(cell ?? '').trim().toLowerCase(); return HEADER_ANCHORS.includes(v) })
          if (found) { headerRowIdx = i; break }
        }
        const rawHdr  = allRows[headerRowIdx] || []
        const headers = rawHdr.map(h => String(h ?? '').trim()).filter(h => h)
        const dataRows = allRows.slice(headerRowIdx + 1).filter(row => row.some(c => c !== null && c !== ''))
        const data = dataRows.map(row => { const obj = {}; headers.forEach((h, i) => { obj[h] = row[i] ?? null }); return obj })
        if (!data.length) throw new Error('Worksheet appears to be empty.')
        resolve({ data, ws, headers, wb, sheetName, rawBuffer: e.target.result })
      } catch (err) { reject(new Error('Failed to parse Excel: ' + err.message)) }
    }
    reader.onerror = evt => {
      const domErr = evt?.target?.error
      const detail = domErr ? ` [${domErr.name}: ${domErr.message}]` : ''
      reject(new Error(`File read failed${detail}. Possible causes: file is open in Excel (close it first), browser security restriction, or the file is corrupted. Re-upload the file and try again.`))
    }
    try { reader.readAsArrayBuffer(file) }
    catch (syncErr) { reject(new Error('Could not start file read: ' + syncErr.message)) }
  })
}
