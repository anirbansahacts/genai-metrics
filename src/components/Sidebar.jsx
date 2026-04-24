import { NavLink } from 'react-router-dom'
import './Sidebar.css'

const apps = [
  { path: '/ghcp-interaction', label: 'GHCP Interaction Usage Compare', icon: '📊' },
  { path: '/copilot-usage', label: 'CoPilot Usage Report', icon: '🤖' },
  { path: '/weekly-status', label: 'Weekly Status Report', icon: '📅' },
  { path: '/sprint-productivity', label: 'Sprint Productivity Analytics', icon: '⚡' },
  { path: '/copilot-vs-git', label: 'Copilot vs Git Analytics', icon: '🔀' },
]

export default function Sidebar({ isOpen, onClose }) {
  return (
    <>
      {/* Mobile backdrop */}
      {isOpen && <div className="sidebar-backdrop" onClick={onClose} />}

      <aside className={`sidebar ${isOpen ? 'open' : 'closed'}`}>
        <div className="sidebar-inner">
          <div className="sidebar-section-label">Dashboards</div>
          <nav>
            <ul className="sidebar-nav">
              {apps.map(({ path, label, icon }) => (
                <li key={path}>
                  <NavLink
                    to={path}
                    onClick={onClose}
                    className={({ isActive }) =>
                      `sidebar-link ${isActive ? 'active' : ''}`
                    }
                  >
                    <span className="sidebar-link-icon">{icon}</span>
                    <span className="sidebar-link-label">{label}</span>
                  </NavLink>
                </li>
              ))}
            </ul>
          </nav>
        </div>
      </aside>
    </>
  )
}
