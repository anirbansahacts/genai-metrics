import { NavLink } from 'react-router-dom'
import './Navbar.css'

export default function Navbar({ onToggleSidebar }) {
  return (
    <nav className="navbar">
      <div className="navbar-left">
        <button className="hamburger" onClick={onToggleSidebar} aria-label="Toggle sidebar">
          ☰
        </button>
        <NavLink to="/" className="navbar-brand">
          GenAI Metrics
        </NavLink>
      </div>
      <div className="navbar-right">
        <NavLink
          to="/upload"
          className={({ isActive }) =>
            `upload-btn ${isActive ? 'active' : ''}`
          }
        >
          Upload Portal
        </NavLink>
      </div>
    </nav>
  )
}
