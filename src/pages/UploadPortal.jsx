import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import './UploadPortal.css'

const UPLOAD_PASSWORD = import.meta.env.VITE_UPLOAD_PASSWORD || 'genai2024'

const apps = [
  { id: 'weekly-status',     title: 'Weekly Status Report',   icon: '' },
  { id: 'sprint-productivity', title: 'Sprint Productivity',   icon: '' },
  { id: 'copilot-vs-git',    title: 'Copilot vs Git Analytics', icon: '' },
]

/* ── CPI sheet name → CPI key mapping ── */
const CPI_KEY_MAP = ['CPI1','CPI2A','CPI2B','CPI3','CPI4','CPI5','CPI6']
function matchCpiKey(sheetName) {
  return CPI_KEY_MAP.find(k => sheetName.startsWith(k)) || null
}

/* ── Month inference: filename → canonical "MMM-YYYY" ── */
function inferMonth(filename, sheets) {
  // Try filename patterns
  const name = filename.replace(/\.[^.]+$/, '') // strip extension
  const patterns = [
    /(?:^|[_\s-])(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[_\s-]?(\d{4})(?:[_\s-]|$)/i,
    /(\d{2})[_\s-](\d{4})/,
    /(\d{4})[_\s-](\d{2})/,
    /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{4}/i,
  ]
  for (const re of patterns) {
    const m = name.match(re)
    if (m) {
      // Check which capture groups we got
      if (/[a-z]/i.test(m[1])) {
        // Named month
        const mo = m[1].toUpperCase().slice(0, 3)
        const yr = m[2] || (m[0].match(/\d{4}/) || [])[0]
        if (yr) return `${mo}-${yr}`
      } else {
        // Numeric — try MM-YYYY or YYYY-MM
        const a = parseInt(m[1]), b = parseInt(m[2])
        const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
        if (a <= 12 && b > 2000) return `${months[a - 1]}-${b}`
        if (b <= 12 && a > 2000) return `${months[b - 1]}-${a}`
      }
    }
  }
  // Fall back to first Month cell in CPI1
  const cpi1 = sheets['CPI1']
  if (cpi1?.length && cpi1[0].Month) {
    return String(cpi1[0].Month).toUpperCase()
  }
  return null // signal caller to ask user
}

export default function UploadPortal() {
  const [authenticated, setAuthenticated] = useState(false)
  const [passwordInput, setPasswordInput] = useState('')
  const [authError, setAuthError]         = useState('')
  const [uploadedFiles, setUploadedFiles] = useState([])
  const fileInputRefs = useRef(apps.reduce((acc, app) => ({ ...acc, [app.id]: null }), {}))

  // ── CPI upload state ──────────────────────────────────────────────
  const cpiFileInputRef          = useRef(null)
  const [cpiParsing, setCpiParsing]       = useState(false)
  const [cpiError, setCpiError]           = useState('')
  const [cpiPreview, setCpiPreview]       = useState(null) // { name, month, sheetCount, totalRows, parsed }
  const [cpiMonthInput, setCpiMonthInput] = useState('')   // manual month entry
  const [cpiNeedsMonth, setCpiNeedsMonth] = useState(false)

  function handleLogin(e) {
    e.preventDefault()
    if (passwordInput === UPLOAD_PASSWORD) {
      setAuthenticated(true)
      setAuthError('')
    } else {
      setAuthError('Incorrect password. Please try again.')
      setPasswordInput('')
    }
  }

  function handleFiles(files, appId) {
    const newFiles = Array.from(files).map(file => ({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + ' KB',
      type: file.type || 'unknown',
      addedAt: new Date().toLocaleTimeString(),
      app: apps.find(app => app.id === appId)?.title || 'Unknown',
    }))
    setUploadedFiles(prev => [...prev, ...newFiles])
  }

  function handleFileInputChange(e, appId) {
    handleFiles(e.target.files, appId)
    e.target.value = ''
  }

  function removeFile(index) {
    setUploadedFiles(prev => prev.filter((_, i) => i !== index))
  }

  // ── CPI Analytics XLSX parser ─────────────────────────────────────
  async function handleCpiFileChange(e) {
    const file = e.target.files?.[0]
    e.target.value = ''
    if (!file) return

    setCpiError('')
    setCpiPreview(null)
    setCpiNeedsMonth(false)
    setCpiMonthInput('')
    setCpiParsing(true)

    try {
      const buf  = await file.arrayBuffer()
      const wb   = XLSX.read(buf, { type: 'array', cellDates: true })

      // Build sheets dict filtered to CPI sheets only
      const sheets = {}
      const sheetKeys = []
      for (const sn of wb.SheetNames) {
        const cpiKey = matchCpiKey(sn)
        if (!cpiKey) continue
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: null })
        // Strip PII
        const clean = rows.map(r => {
          const row = { ...r }
          ;['Author Mail','MR Link','email','Email'].forEach(f => delete row[f])
          return row
        })
        sheets[cpiKey] = clean
        sheetKeys.push(cpiKey)
      }

      if (!sheetKeys.length) {
        throw new Error('No CPI sheets found (expected sheets named CPI1, CPI2B, CPI3, CPI4, CPI5, or CPI6).')
      }

      const totalRows = Object.values(sheets).reduce((s, rows) => s + rows.length, 0)
      const name      = file.name.replace(/\.[^.]+$/, '')
      const month     = inferMonth(file.name, sheets)

      const parsed = {
        id:        crypto?.randomUUID?.() || String(Date.now()),
        name,
        month:     month || '',
        size:      file.size,
        sheetKeys,
        sheets,
      }

      setCpiPreview({ name, month, sheetCount: sheetKeys.length, totalRows, parsed })
      if (!month) setCpiNeedsMonth(true)

    } catch (err) {
      setCpiError(err.message)
    } finally {
      setCpiParsing(false)
    }
  }

  function handleCpiConfirm() {
    if (!cpiPreview) return
    const finalMonth = cpiNeedsMonth ? cpiMonthInput.trim().toUpperCase() : cpiPreview.month
    if (!finalMonth) { setCpiError('Please enter the report month before confirming.'); return }

    // TODO: replace with direct Azure Blob PUT when deployed to production
    const report   = { ...cpiPreview.parsed, month: finalMonth }
    const filename = finalMonth.toLowerCase().replace(/[^a-z0-9]+/g, '_') + '.json'
    const blob     = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' })
    const url      = URL.createObjectURL(blob)
    const a        = document.createElement('a')
    a.href         = url
    a.download     = filename
    a.click()
    URL.revokeObjectURL(url)

    setCpiPreview(prev => ({ ...prev, confirmed: true, finalMonth, filename }))
  }

  function handleCpiReset() {
    setCpiPreview(null)
    setCpiError('')
    setCpiNeedsMonth(false)
    setCpiMonthInput('')
  }

  if (!authenticated) {
    return (
      <div className="upload-portal">
        <div className="auth-card">
          <div className="auth-icon">🔒</div>
          <h1>Upload Portal</h1>
          <p>This area is password protected. Enter the password to continue.</p>
          <form onSubmit={handleLogin} className="auth-form">
            <input
              type="password"
              placeholder="Enter password"
              value={passwordInput}
              onChange={e => setPasswordInput(e.target.value)}
              autoFocus
              className={authError ? 'input-error' : ''}
            />
            {authError && <span className="error-msg">{authError}</span>}
            <button type="submit">Unlock</button>
          </form>
        </div>
      </div>
    )
  }

  return (
    <div className="upload-portal">
      <div className="upload-header">
        <h1>Upload Portal</h1>
        <p>Upload data files to populate the analytics dashboards.</p>
        <button className="logout-btn" onClick={() => setAuthenticated(false)}>
          Lock Portal
        </button>
      </div>

      <div className="upload-apps-list">
        {/* ── Standard app upload slots ── */}
        {apps.map(app => (
          <div key={app.id} className="upload-app-item">
            <div className="app-info-inline">
              <div className="app-title-section">
                <span className="upload-app-icon">{app.icon}</span>
                <h3>{app.title}</h3>
              </div>
              <button className="upload-btn" onClick={() => fileInputRefs.current[app.id]?.click()}>
                Upload Files
              </button>
              <input
                ref={el => fileInputRefs.current[app.id] = el}
                type="file"
                multiple
                accept=".xlsx,.csv,.json"
                onChange={e => handleFileInputChange(e, app.id)}
                style={{ display: 'none' }}
              />
            </div>
          </div>
        ))}

        {/* ── CPI Analytics upload slot ── */}
        <div className="upload-app-item cpi-upload-item">
          <div className="app-info-inline">
            <div className="app-title-section">
              <span className="upload-app-icon"></span>
              <h3>CPI Analytics — Add Monthly Report</h3>
            </div>
            {!cpiPreview && (
              <button
                className="upload-btn"
                onClick={() => cpiFileInputRef.current?.click()}
                disabled={cpiParsing}
              >
                {cpiParsing ? 'Parsing…' : 'Choose .xlsx'}
              </button>
            )}
            <input
              ref={cpiFileInputRef}
              type="file"
              accept=".xlsx"
              onChange={handleCpiFileChange}
              style={{ display: 'none' }}
            />
          </div>

          {/* Parse error */}
          {cpiError && (
            <div className="cpi-upload-error" style={{ marginTop: '0.75rem' }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
              </svg>
              {cpiError}
              <button className="cpi-upload-dismiss" onClick={() => setCpiError('')}>✕</button>
            </div>
          )}

          {/* Preview */}
          {cpiPreview && !cpiPreview.confirmed && (
            <div className="cpi-upload-preview">
              <div className="cpi-upload-preview__title">Preview</div>
              <table className="cpi-upload-preview__table">
                <tbody>
                  <tr><td>File name</td><td>{cpiPreview.name}.xlsx</td></tr>
                  <tr>
                    <td>Detected month</td>
                    <td>
                      {cpiNeedsMonth ? (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                          <span style={{ color: 'var(--yellow)' }}>Not detected —</span>
                          <input
                            className="cpi-month-input"
                            type="text"
                            placeholder="e.g. APR-2026"
                            value={cpiMonthInput}
                            onChange={e => setCpiMonthInput(e.target.value)}
                          />
                        </div>
                      ) : (
                        <span className="cpi-badge">{cpiPreview.month}</span>
                      )}
                    </td>
                  </tr>
                  <tr><td>CPI sheets found</td><td>{cpiPreview.sheetCount}</td></tr>
                  <tr><td>Total rows</td><td>{cpiPreview.totalRows.toLocaleString()}</td></tr>
                </tbody>
              </table>
              <div className="cpi-upload-preview__actions">
                <button className="upload-btn" onClick={handleCpiConfirm}>
                  Download JSON
                </button>
                <button className="cpi-upload-cancel" onClick={handleCpiReset}>
                  Cancel
                </button>
              </div>
            </div>
          )}

          {/* Post-confirm instruction */}
          {cpiPreview?.confirmed && (
            <div className="cpi-upload-success">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polyline points="20 6 9 17 4 12"/>
              </svg>
              <div>
                <strong>{cpiPreview.filename}</strong> downloaded.
                <ol className="cpi-upload-instructions">
                  <li>Move <code>{cpiPreview.filename}</code> into <code>local-storage/cpi-analytics/</code></li>
                  <li>Open <code>local-storage/cpi-analytics/config.json</code> and add <code>"{cpiPreview.filename}"</code> to the <code>reports</code> array (sorted chronologically), then update <code>lastUpdated</code> to today&apos;s date.</li>
                  <li>Refresh the app and navigate to <strong>CPI Analytics</strong> to see the new month.</li>
                </ol>
                <button className="cpi-upload-cancel" style={{ marginTop: '0.5rem' }} onClick={handleCpiReset}>
                  Upload another
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      {uploadedFiles.length > 0 && (
        <div className="file-list">
          <h2>Uploaded Files ({uploadedFiles.length})</h2>
          <table>
            <thead>
              <tr>
                <th>File Name</th>
                <th>Size</th>
                <th>Type</th>
                <th>Added</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {uploadedFiles.map((file, i) => (
                <tr key={i}>
                  <td>{file.name}</td>
                  <td>{file.size}</td>
                  <td>{file.type}</td>
                  <td>{file.addedAt}</td>
                  <td>
                    <button className="remove-btn" onClick={() => removeFile(i)}>
                      Remove
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
