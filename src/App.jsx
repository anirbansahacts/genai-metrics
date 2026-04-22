import { BrowserRouter, Routes, Route } from 'react-router-dom'
import Navbar from './components/Navbar'
import Home from './pages/Home'
import GhcpInteraction from './pages/GhcpInteraction'
import CopilotUsage from './pages/CopilotUsage'
import WeeklyStatus from './pages/WeeklyStatus'
import SprintProductivity from './pages/SprintProductivity'
import CopilotVsGit from './pages/CopilotVsGit'
import UploadPortal from './pages/UploadPortal'

export default function App() {
  return (
    <BrowserRouter>
      <Navbar />
      <main className="main-content">
        <Routes>
          <Route path="/" element={<Home />} />
          <Route path="/ghcp-interaction" element={<GhcpInteraction />} />
          <Route path="/copilot-usage" element={<CopilotUsage />} />
          <Route path="/weekly-status" element={<WeeklyStatus />} />
          <Route path="/sprint-productivity" element={<SprintProductivity />} />
          <Route path="/copilot-vs-git" element={<CopilotVsGit />} />
          <Route path="/upload" element={<UploadPortal />} />
        </Routes>
      </main>
    </BrowserRouter>
  )
}
