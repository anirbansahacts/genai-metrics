import { useState, useRef } from 'react'
import './UploadPortal.css'

const UPLOAD_PASSWORD = 'genai2024'

const apps = [
  {
    id: 'ghcp-interaction',
    title: 'GHCP Interaction Usage',
    icon: '📊',
  },
  {
    id: 'copilot-usage',
    title: 'CoPilot Usage Report',
    icon: '🤖',
  },
  {
    id: 'weekly-status',
    title: 'Weekly Status Report',
    icon: '📅',
  },
  {
    id: 'sprint-productivity',
    title: 'Sprint Productivity',
    icon: '⚡',
  },
  {
    id: 'copilot-vs-git',
    title: 'Copilot vs Git Analytics',
    icon: '🔀',
  },
]

export default function UploadPortal() {
  const [authenticated, setAuthenticated] = useState(false)
  const [passwordInput, setPasswordInput] = useState('')
  const [authError, setAuthError] = useState('')
  const [uploadedFiles, setUploadedFiles] = useState([])
  const fileInputRefs = useRef(apps.reduce((acc, app) => ({ ...acc, [app.id]: null }), {}))

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
    const newFiles = Array.from(files).map((file) => ({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + ' KB',
      type: file.type || 'unknown',
      addedAt: new Date().toLocaleTimeString(),
      app: apps.find(app => app.id === appId)?.title || 'Unknown',
    }))
    setUploadedFiles((prev) => [...prev, ...newFiles])
  }

  function handleFileInputChange(e, appId) {
    handleFiles(e.target.files, appId)
    e.target.value = ''
  }

  function removeFile(index) {
    setUploadedFiles((prev) => prev.filter((_, i) => i !== index))
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
              onChange={(e) => setPasswordInput(e.target.value)}
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

      <div className="upload-apps-grid">
        {apps.map((app) => (
          <div key={app.id} className="upload-app-card">
            <div className="upload-app-icon">{app.icon}</div>
            <h3>{app.title}</h3>
            <button
              className="upload-btn"
              onClick={() => fileInputRefs.current[app.id]?.click()}
            >
              Upload Files
            </button>
            <input
              ref={(el) => fileInputRefs.current[app.id] = el}
              type="file"
              multiple
              accept=".xlsx,.csv,.json"
              onChange={(e) => handleFileInputChange(e, app.id)}
              style={{ display: 'none' }}
            />
          </div>
        ))}
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
