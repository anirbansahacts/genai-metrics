import { NavLink, useLocation } from 'react-router-dom'
import './Navbar.css'

export default function Navbar({ onToggleSidebar, theme, onToggleTheme }) {
  const location = useLocation()
  const showUploadPortal = location.pathname === '/genai-reports' ||
                          location.pathname.startsWith('/weekly-status') ||
                          location.pathname.startsWith('/sprint-productivity') ||
                          location.pathname.startsWith('/copilot-vs-git') ||
                          location.pathname.startsWith('/batch-flow') ||
                          location.pathname.startsWith('/upload')

  return (
    <nav className="navbar">
      <div className="navbar-left">
        <button className="hamburger" onClick={onToggleSidebar} aria-label="Toggle sidebar">
          ☰
        </button>
        <NavLink to="/" className="navbar-brand">
          Flex 2.0 Tools
        </NavLink>
      </div>
      <div className="navbar-right">
        {showUploadPortal && (
          <NavLink
            to="/upload"
            className={({ isActive }) =>
              `upload-btn ${isActive ? 'active' : ''}`
            }
          >
            Upload Portal
          </NavLink>
        )}
        <button
          className="theme-toggle"
          onClick={onToggleTheme}
          aria-label={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
          title={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
        >
          {theme === 'dark' ? '☀️' : '🌙'}
        </button>
      </div>
    </nav>
  )
}
