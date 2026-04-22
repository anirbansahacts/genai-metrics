import { useState, useRef } from 'react'
import './UploadPortal.css'

const UPLOAD_PASSWORD = 'genai2024'

export default function UploadPortal() {
  const [authenticated, setAuthenticated] = useState(false)
  const [passwordInput, setPasswordInput] = useState('')
  const [authError, setAuthError] = useState('')
  const [dragOver, setDragOver] = useState(false)
  const [uploadedFiles, setUploadedFiles] = useState([])
  const fileInputRef = useRef(null)

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

  function handleFiles(files) {
    const newFiles = Array.from(files).map((file) => ({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + ' KB',
      type: file.type || 'unknown',
      addedAt: new Date().toLocaleTimeString(),
    }))
    setUploadedFiles((prev) => [...prev, ...newFiles])
  }

  function handleDrop(e) {
    e.preventDefault()
    setDragOver(false)
    handleFiles(e.dataTransfer.files)
  }

  function handleFileInputChange(e) {
    handleFiles(e.target.files)
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

      <div
        className={`drop-zone ${dragOver ? 'drag-over' : ''}`}
        onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
        onDragLeave={() => setDragOver(false)}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current.click()}
      >
        <div className="drop-zone-content">
          <span className="drop-icon">📂</span>
          <p>Drag &amp; drop files here, or click to select</p>
          <span className="drop-hint">Supports .xlsx, .csv, .json</span>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept=".xlsx,.csv,.json"
          onChange={handleFileInputChange}
          style={{ display: 'none' }}
        />
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
