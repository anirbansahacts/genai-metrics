import { useState } from 'react'
import { BrowserRouter, Routes, Route } from 'react-router-dom'
import Navbar from './components/Navbar'
import Sidebar from './components/Sidebar'
import Home from './pages/Home'
import GenaiReports from './pages/GenaiReports'
import CpiAnalytics from './pages/CpiAnalytics'
import WeeklyStatus from './pages/WeeklyStatus'
import SprintProductivity from './pages/SprintProductivity'
import CopilotVsGit from './pages/CopilotVsGit'
import BatchFlow from './pages/BatchFlow'
import UploadPortal from './pages/UploadPortal'

export default function App() {
  const [sidebarOpen, setSidebarOpen] = useState(() => window.innerWidth >= 768)

  function toggleSidebar() {
    setSidebarOpen((prev) => !prev)
  }

  function closeSidebar() {
    // Only auto-close on mobile
    if (window.innerWidth < 768) setSidebarOpen(false)
  }

  return (
    <BrowserRouter>
      <Navbar onToggleSidebar={toggleSidebar} />
      <div className="app-body">
        <Sidebar isOpen={sidebarOpen} onClose={closeSidebar} />
        <main className="main-content">
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/genai-reports" element={<GenaiReports />} />
            <Route path="/cpi-analytics" element={<CpiAnalytics />} />
            <Route path="/weekly-status" element={<WeeklyStatus />} />
            <Route path="/sprint-productivity" element={<SprintProductivity />} />
            <Route path="/copilot-vs-git" element={<CopilotVsGit />} />
            <Route path="/batch-flow" element={<BatchFlow />} />
            <Route path="/upload" element={<UploadPortal />} />
          </Routes>
        </main>
      </div>
    </BrowserRouter>
  )
}
