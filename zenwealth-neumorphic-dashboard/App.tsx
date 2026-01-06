
import React, { useState, useEffect } from 'react';
import { 
  LayoutDashboard, 
  TrendingUp, 
  History, 
  Wallet, 
  Eye, 
  EyeOff, 
  Moon, 
  Sun,
  User,
  Plus
} from 'lucide-react';
import { AppRoute } from './types';
import Cockpit from './components/Cockpit';
import InvestmentPulse from './components/InvestmentPulse';
import ActivityLog from './components/ActivityLog';
import WalletOverview from './components/WalletOverview';
import InvestDetail from './components/InvestDetail';
import RecordsDetail from './components/RecordsDetail';
import AccountsDetail from './components/AccountsDetail';

const App: React.FC = () => {
  const [currentRoute, setCurrentRoute] = useState<AppRoute>('overview');
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [isPrivacyMode, setIsPrivacyMode] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    // Simulate initial loading for premium feel
    const timer = setTimeout(() => setIsLoading(false), 800);
    return () => clearTimeout(timer);
  }, [currentRoute]);

  const toggleTheme = () => {
    setIsDarkMode(!isDarkMode);
    document.documentElement.classList.toggle('dark');
  };

  const navItems = [
    { id: 'overview', icon: LayoutDashboard, label: 'Overview' },
    { id: 'invest', icon: TrendingUp, label: 'Invest' },
    { id: 'records', icon: History, label: 'History' },
    { id: 'accounts', icon: Wallet, label: 'Accounts' },
  ];

  const renderContent = () => {
    switch (currentRoute) {
      case 'overview':
        return (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
            <Cockpit isPrivacyMode={isPrivacyMode} isLoading={isLoading} />
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              <InvestmentPulse 
                isPrivacyMode={isPrivacyMode} 
                isLoading={isLoading} 
                onNavigate={() => setCurrentRoute('invest')}
              />
              <ActivityLog 
                isPrivacyMode={isPrivacyMode} 
                isLoading={isLoading} 
                onNavigate={() => setCurrentRoute('records')}
              />
            </div>
            <WalletOverview 
              isPrivacyMode={isPrivacyMode} 
              isLoading={isLoading} 
              onNavigate={() => setCurrentRoute('accounts')}
            />
          </div>
        );
      case 'invest':
        return <InvestDetail isPrivacyMode={isPrivacyMode} isLoading={isLoading} />;
      case 'records':
        return <RecordsDetail isPrivacyMode={isPrivacyMode} isLoading={isLoading} />;
      case 'accounts':
        return <AccountsDetail isPrivacyMode={isPrivacyMode} isLoading={isLoading} />;
      default:
        return null;
    }
  };

  return (
    <div className="flex flex-col min-h-screen pb-24 md:pb-0 md:pl-20 transition-colors duration-300">
      
      {/* Header (Mobile) */}
      <header className="fixed top-0 left-0 right-0 z-50 px-6 py-4 flex items-center justify-between bg-neu-light/80 dark:bg-neu-dark/80 glass-blur md:hidden">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-full neu-flat flex items-center justify-center">
            <User className="w-5 h-5 text-neu-accent" />
          </div>
          <div>
            <h1 className="text-sm font-bold">ZenWealth</h1>
            <p className="text-[10px] opacity-60">Wealth OS v2.5</p>
          </div>
        </div>
        <div className="flex gap-4">
          <button 
            onClick={() => setIsPrivacyMode(!isPrivacyMode)}
            className="w-10 h-10 rounded-full neu-flat flex items-center justify-center transition-all active:scale-90"
          >
            {isPrivacyMode ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
          </button>
          <button 
            onClick={toggleTheme}
            className="w-10 h-10 rounded-full neu-flat flex items-center justify-center transition-all active:scale-90"
          >
            {isDarkMode ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
          </button>
        </div>
      </header>

      {/* Sidebar Navigation (Desktop) */}
      <nav className="hidden md:flex fixed left-0 top-0 bottom-0 w-20 flex-col items-center py-8 gap-10 bg-neu-light dark:bg-neu-dark border-r border-slate-200/10 z-50">
        <div className="w-12 h-12 rounded-2xl neu-flat flex items-center justify-center bg-neu-accent mb-6">
          <TrendingUp className="text-white w-6 h-6" />
        </div>
        {navItems.map((item) => (
          <button
            key={item.id}
            onClick={() => setCurrentRoute(item.id as AppRoute)}
            className={`w-12 h-12 rounded-2xl flex items-center justify-center transition-all duration-300 group relative ${
              currentRoute === item.id 
                ? 'neu-inset text-neu-accent' 
                : 'hover:neu-flat text-slate-400'
            }`}
          >
            <item.icon className="w-6 h-6" />
            <span className="absolute left-16 px-2 py-1 rounded bg-slate-800 text-white text-[10px] font-bold opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none whitespace-nowrap">
              {item.label}
            </span>
          </button>
        ))}
        <div className="mt-auto flex flex-col gap-6">
          <button onClick={() => setIsPrivacyMode(!isPrivacyMode)} className="w-10 h-10 rounded-full neu-flat flex items-center justify-center hover:text-neu-accent transition-colors">
            {isPrivacyMode ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
          </button>
          <button onClick={toggleTheme} className="w-10 h-10 rounded-full neu-flat flex items-center justify-center hover:text-neu-accent transition-colors">
            {isDarkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
          </button>
        </div>
      </nav>

      {/* Bottom Navigation (Mobile) */}
      <nav className="md:hidden fixed bottom-6 left-6 right-6 h-16 px-6 bg-neu-light/90 dark:bg-neu-dark/90 glass-blur rounded-full neu-flat flex justify-between items-center z-50 border border-white/10">
        {navItems.map((item) => (
          <button
            key={item.id}
            onClick={() => setCurrentRoute(item.id as AppRoute)}
            className={`flex flex-col items-center gap-1 transition-all ${
              currentRoute === item.id ? 'text-neu-accent scale-110' : 'text-slate-400 opacity-60'
            }`}
          >
            <item.icon className="w-5 h-5" />
            <span className="text-[9px] font-bold tracking-tight">{item.label}</span>
          </button>
        ))}
      </nav>

      {/* Main Content Area */}
      <main className="flex-1 mt-24 md:mt-0 p-6 md:p-12 max-w-7xl mx-auto w-full relative z-10">
        {renderContent()}
      </main>
      
      {/* Decorative background elements */}
      <div className="fixed top-[-10%] right-[-10%] w-[40%] h-[40%] bg-neu-accent/5 rounded-full blur-[120px] pointer-events-none" />
      <div className="fixed bottom-[-10%] left-[-10%] w-[40%] h-[40%] bg-neu-accent/10 rounded-full blur-[120px] pointer-events-none" />
    </div>
  );
};

export default App;
