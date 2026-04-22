import { NavLink } from 'react-router-dom'
import './Navbar.css'

const apps = [
  { path: '/ghcp-interaction', label: 'GHCP Interaction Usage Compare' },
  { path: '/copilot-usage', label: 'CoPilot Usage Report' },
  { path: '/weekly-status', label: 'Weekly Status Report' },
  { path: '/sprint-productivity', label: 'Sprint Productivity Analytics' },
  { path: '/copilot-vs-git', label: 'Copilot vs Git Analytics' },
]

export default function Navbar() {
  return (
    <nav className="navbar">
      <NavLink to="/" className="navbar-brand">
        GenAI Metrics
      </NavLink>
      <ul className="navbar-links">
        {apps.map(({ path, label }) => (
          <li key={path}>
            <NavLink
              to={path}
              className={({ isActive }) => isActive ? 'nav-link active' : 'nav-link'}
            >
              {label}
            </NavLink>
          </li>
        ))}
        <li>
          <NavLink
            to="/upload"
            className={({ isActive }) => isActive ? 'nav-link upload-link active' : 'nav-link upload-link'}
          >
            Upload Portal
          </NavLink>
        </li>
      </ul>
    </nav>
  )
}
