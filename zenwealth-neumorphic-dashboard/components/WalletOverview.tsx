
import React from 'react';
import Card from './Card';
import NumberFormatter from './NumberFormatter';
import { MOCK_ACCOUNTS, FX_RATE } from '../constants';
import { Building2, AlertCircle, ChevronRight } from 'lucide-react';

interface WalletOverviewProps {
  isPrivacyMode: boolean;
  isLoading: boolean;
  onNavigate: () => void;
}

const WalletOverview: React.FC<WalletOverviewProps> = ({ isPrivacyMode, isLoading, onNavigate }) => {
  // Sort: CTBC First, then by balance
  const sortedAccounts = [...MOCK_ACCOUNTS].sort((a, b) => {
    if (a.bank === 'CTBC') return -1;
    if (b.bank === 'CTBC') return 1;
    return b.balance - a.balance;
  });

  const totalLiquid = MOCK_ACCOUNTS.reduce((acc, curr) => {
    if (curr.isDebt) return acc;
    const value = curr.currency === 'USD' ? curr.balance * FX_RATE : curr.balance;
    return acc + value;
  }, 0);

  return (
    <Card 
      title="Wallet Distribution" 
      subtitle={`Total Liquid: $${new Intl.NumberFormat().format(totalLiquid)}`}
      onClick={onNavigate}
      isLoading={isLoading}
      className="lg:col-span-2"
    >
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {sortedAccounts.map(account => (
          <div 
            key={account.id} 
            className={`
              p-5 rounded-3xl neu-inset border-l-4 transition-all hover:scale-[1.02]
              ${account.isDebt ? 'border-rose-500/50 bg-rose-500/5' : 'border-neu-accent/50'}
            `}
          >
            <div className="flex justify-between items-start mb-4">
              <div className="w-10 h-10 rounded-2xl neu-flat flex items-center justify-center text-slate-400">
                {account.isDebt ? <AlertCircle size={20} className="text-rose-500" /> : <Building2 size={20} />}
              </div>
              <div className="px-2 py-1 rounded-md text-[9px] font-black uppercase tracking-tighter bg-slate-200/50 dark:bg-slate-800/50">
                {account.bank}
              </div>
            </div>
            
            <div className="space-y-1">
              <h5 className="text-sm font-bold truncate opacity-80">{account.name}</h5>
              <div className="flex items-baseline gap-1">
                <NumberFormatter 
                  value={account.balance} 
                  currency={account.currency} 
                  isPrivacyMode={isPrivacyMode} 
                  className={`text-2xl font-black ${account.isDebt ? 'text-rose-500' : 'text-slate-800 dark:text-white'}`}
                  precision={account.currency === 'USD' ? 2 : 0}
                />
              </div>
              {account.currency === 'USD' && !isPrivacyMode && (
                <p className="text-[10px] font-medium opacity-40">
                  ≈ <NumberFormatter value={account.balance * FX_RATE} currency="TWD" isPrivacyMode={false} />
                </p>
              )}
            </div>
          </div>
        ))}
        
        {/* Add Account Shortcut */}
        <div className="flex items-center justify-center p-5 rounded-3xl border-2 border-dashed border-slate-300 dark:border-slate-700 opacity-40 hover:opacity-100 transition-opacity cursor-pointer group">
          <div className="flex flex-col items-center gap-2">
            <div className="w-10 h-10 rounded-full neu-flat flex items-center justify-center group-hover:scale-110 transition-transform">
              <ChevronRight size={20} />
            </div>
            <span className="text-xs font-bold uppercase tracking-widest">Connect Bank</span>
          </div>
        </div>
      </div>
    </Card>
  );
};

export default WalletOverview;
