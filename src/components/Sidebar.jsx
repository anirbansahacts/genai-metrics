import { NavLink, useLocation } from 'react-router-dom'
import { useState, useEffect } from 'react'
import './Sidebar.css'

const mainSections = [
  {
    id: 'genai-reports',
    path: '/genai-reports',
    label: 'GenAI Reports',
    icon: '',
    subItems: [
      { path: '/weekly-status', label: 'Weekly Status Report', icon: '' },
      { path: '/sprint-productivity', label: 'Sprint Productivity Analytics', icon: '' },
      { path: '/copilot-vs-git', label: 'Copilot vs Git', icon: '' },
      { path: '/batch-flow', label: 'BatchFlow', icon: '' },
    ]
  },
  {
    id: 'service-governance',
    path: '/cpi-analytics',
    label: 'Service Governance',
    icon: '',
    subItems: [
      { path: '/cpi-analytics', label: 'CPI Analytics', icon: '' },
    ]
  }
]

export default function Sidebar({ isOpen, onClose }) {
  const [expandedSection, setExpandedSection] = useState('genai-reports')
  const location = useLocation()

  // Auto-expand sections based on current path
  useEffect(() => {
    const genaiPaths = ['/weekly-status', '/sprint-productivity', '/copilot-vs-git', '/genai-reports']
    const servicePaths = ['/cpi-analytics']
    
    if (genaiPaths.some(path => location.pathname.startsWith(path))) {
      setExpandedSection('genai-reports')
    } else if (servicePaths.some(path => location.pathname.startsWith(path))) {
      setExpandedSection('service-governance')
    }
  }, [location.pathname])

  const toggleSection = (sectionId) => {
    setExpandedSection(expandedSection === sectionId ? null : sectionId)
  }

  return (
    <>
      {/* Mobile backdrop */}
      {isOpen && <div className="sidebar-backdrop" onClick={onClose} />}

      <aside className={`sidebar ${isOpen ? 'open' : 'closed'}`}>
        <div className="sidebar-inner">
          <div className="sidebar-section-label">FLEX 2.0 TOOLS</div>
          <nav>
            <ul className="sidebar-nav">
              {mainSections.map((section) => (
                <li key={section.id}>
                  {section.subItems ? (
                    // Expandable section
                    <>
                      <div
                        className={`sidebar-section-header ${expandedSection === section.id ? 'expanded' : ''}`}
                        onClick={() => toggleSection(section.id)}
                      >
                        <span className="sidebar-link-icon">{section.icon}</span>
                        <span className="sidebar-link-label">{section.label}</span>
                        <span className="expand-arrow">{expandedSection === section.id ? '▼' : '▶'}</span>
                      </div>
                      {expandedSection === section.id && (
                        <ul className="sidebar-subnav">
                          {section.subItems.map((item) => (
                            <li key={item.path}>
                              <NavLink
                                to={item.path}
                                onClick={onClose}
                                className={({ isActive }) =>
                                  `sidebar-link sub-item ${isActive ? 'active' : ''}`
                                }
                              >
                                <span className="sidebar-link-icon">{item.icon}</span>
                                <span className="sidebar-link-label">{item.label}</span>
                              </NavLink>
                            </li>
                          ))}
                        </ul>
                      )}
                    </>
                  ) : (
                    // Regular link
                    <NavLink
                      to={section.path}
                      onClick={onClose}
                      className={({ isActive }) =>
                        `sidebar-link ${isActive ? 'active' : ''}`
                      }
                    >
                      <span className="sidebar-link-icon">{section.icon}</span>
                      <span className="sidebar-link-label">{section.label}</span>
                    </NavLink>
                  )}
                </li>
              ))}
            </ul>
          </nav>
        </div>
      </aside>
    </>
  )
}
