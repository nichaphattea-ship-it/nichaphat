import React, { useState } from 'react';
import LineSimulator from './components/LineSimulator.tsx';
import StockDashboard from './components/StockDashboard.tsx';
import { Bot, BarChart3, HelpCircle, Sparkles } from 'lucide-react';

export default function App() {
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'simulator'>('dashboard');

  const handleDatabaseUpdate = () => {
    setRefreshTrigger(prev => prev + 1);
  };

  return (
    <div className="min-h-screen bg-[#0f0f13] text-gray-100 flex flex-col font-sans selection:bg-emerald-500/30">
      {/* Universal Workspace Header */}
      <header className="bg-[#15151b] border-b border-gray-800 px-6 py-4 flex items-center justify-between z-10 shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-emerald-500 to-teal-500 flex items-center justify-center text-white shadow-lg">
            <Bot className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-sm font-black tracking-wider uppercase bg-clip-text text-transparent bg-gradient-to-r from-emerald-400 to-teal-400">
              LINE STOCK BOT
            </h1>
            <p className="text-[10px] text-gray-500 leading-none mt-0.5">Restaurant Inventory & Reconciliation Engine</p>
          </div>
        </div>

        {/* Navigation Tabs for Mobile */}
        <div className="flex md:hidden bg-gray-900 p-1 rounded-xl border border-gray-800">
          <button
            onClick={() => setActiveTab('dashboard')}
            className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-colors flex items-center gap-1.5 ${
              activeTab === 'dashboard' ? 'bg-[#06c755] text-white' : 'text-gray-400 hover:text-white'
            }`}
          >
            <BarChart3 className="w-3.5 h-3.5" />
            Dashboard
          </button>
          <button
            onClick={() => setActiveTab('simulator')}
            className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-colors flex items-center gap-1.5 ${
              activeTab === 'simulator' ? 'bg-[#06c755] text-white' : 'text-gray-400 hover:text-white'
            }`}
          >
            <Bot className="w-3.5 h-3.5" />
            LINE Chat
          </button>
        </div>

        {/* Server & API Indicators */}
        <div className="hidden sm:flex items-center gap-4 text-xs">
          <div className="flex items-center gap-1.5 text-gray-400">
            <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-ping"></span>
            <span className="font-semibold text-gray-400">REST API Server:</span>
            <span className="text-gray-300 font-mono font-bold bg-gray-800/60 px-2 py-0.5 rounded border border-gray-700/40">PORT 3000</span>
          </div>
        </div>
      </header>

      {/* Main Container Layout */}
      <main className="flex-1 flex overflow-hidden relative">
        
        {/* Mobile View Routing */}
        <div className="md:hidden flex-1 flex flex-col h-full overflow-hidden">
          {activeTab === 'dashboard' ? (
            <div className="flex-1 overflow-y-auto">
              <StockDashboard 
                refreshTrigger={refreshTrigger} 
                onRefresh={() => setRefreshTrigger(prev => prev + 1)} 
              />
            </div>
          ) : (
            <div className="flex-1 bg-gray-950 flex items-center justify-center p-4 overflow-y-auto">
              <LineSimulator onDatabaseUpdate={handleDatabaseUpdate} />
            </div>
          )}
        </div>

        {/* Desktop View Routing (Side-by-side split screen for the ultimate demo!) */}
        <div className="hidden md:flex flex-1 w-full h-full overflow-hidden">
          
          {/* Main Stock Dashboard Panel (Left/Center) */}
          <div className="flex-1 h-full overflow-hidden flex flex-col">
            <StockDashboard 
              refreshTrigger={refreshTrigger} 
              onRefresh={() => setRefreshTrigger(prev => prev + 1)} 
            />
          </div>

          {/* LINE Chat Bot Simulator Panel (Right) */}
          <div className="w-[430px] bg-[#111115] border-l border-gray-800/80 p-5 shrink-0 flex flex-col items-center justify-center overflow-y-auto shadow-2xl">
            <div className="w-full max-w-[390px] mb-4 text-center px-2">
              <div className="inline-flex items-center gap-1.5 bg-green-950/40 border border-green-800/50 text-green-400 px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider mb-2">
                <Bot className="w-3.5 h-3.5" />
                LINE Bot Sandbox Simulator
              </div>
              <h2 className="text-sm font-bold text-white">ทดสอบพิมพ์คุยกับบอททางขวาได้ทันที</h2>
              <p className="text-[11px] text-gray-400 mt-1 font-sans">
                กดแถบ Rich Menu ด้านล่างจอเพื่อจำลองคำสั่ง เติมคงเหลือ หรือลากวางไฟล์ยอดขาย Excel เพื่อตัดสต๊อก
              </p>
            </div>
            
            <LineSimulator onDatabaseUpdate={handleDatabaseUpdate} />
          </div>

        </div>

      </main>
    </div>
  );
}
