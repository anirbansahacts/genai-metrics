import { Link } from 'react-router-dom'
import './Home.css'

const apps = [
  {
    path: '/ghcp-interaction',
    title: 'GHCP Interaction Usage Compare',
    description: 'Compare GitHub Copilot interaction usage metrics across teams and time periods.',
    icon: '📊',
  },
  {
    path: '/copilot-usage',
    title: 'CoPilot Usage Report',
    description: 'Detailed usage reports for GitHub Copilot adoption and engagement across the organization.',
    icon: '🤖',
  },
  {
    path: '/weekly-status',
    title: 'Weekly Status Report',
    description: 'Weekly aggregated status reports tracking GenAI tool performance and developer productivity.',
    icon: '📅',
  },
  {
    path: '/sprint-productivity',
    title: 'Sprint Productivity Analytics',
    description: 'Analyze sprint-level productivity metrics and the impact of AI-assisted development.',
    icon: '⚡',
  },
  {
    path: '/copilot-vs-git',
    title: 'Copilot vs Git Analytics Dashboard and Tower Metrics',
    description: 'Side-by-side comparison of Copilot-assisted commits vs standard Git activity with Tower metrics.',
    icon: '🔀',
  },
]

export default function Home() {
  return (
    <div className="home">
      <header className="home-header">
        <h1>GenAI Metrics</h1>
        <p>Consolidated analytics platform for GitHub Copilot and AI-assisted development insights</p>
      </header>
      <div className="apps-grid">
        {apps.map(({ path, title, description, icon }) => (
          <Link to={path} key={path} className="app-card">
            <div className="app-card-icon">{icon}</div>
            <h2>{title}</h2>
            <p>{description}</p>
            <span className="app-card-cta">Open Dashboard &rarr;</span>
          </Link>
        ))}
        <Link to="/upload" className="app-card upload-card">
          <div className="app-card-icon">🔒</div>
          <h2>Upload Portal</h2>
          <p>Securely upload data files to populate the analytics dashboards.</p>
          <span className="app-card-cta">Go to Upload &rarr;</span>
        </Link>
      </div>
    </div>
  )
}
