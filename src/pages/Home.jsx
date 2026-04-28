import { Link } from 'react-router-dom'
import './Home.css'

const mainSections = [
  {
    path: '/genai-reports',
    title: 'GenAI Reports',
    description: 'Comprehensive analytics and insights for GitHub Copilot and AI-assisted development metrics.',
    icon: '',
  },
  {
    path: '/cpi-analytics',
    title: 'Service Governance',
    description: 'CPI Analytics, service governance metrics, and process optimization insights.',
    icon: '',
  },
]

export default function Home() {
  return (
    <div className="home">
      <header className="home-header">
        <h1>Flex 2.0 Tools</h1>
        <p>Integrated platform for AI-assisted development analytics and service governance</p>
      </header>
      <div className="apps-grid">
        {mainSections.map(({ path, title, description, icon }) => (
          <Link to={path} key={path} className="app-card">
            <div className="app-card-icon">{icon}</div>
            <h2>{title}</h2>
            <p>{description}</p>
            <span className="app-card-cta">Explore &rarr;</span>
          </Link>
        ))}
      </div>
    </div>
  )
}
