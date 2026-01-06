
import React from 'react';
import Card from './Card';
import NumberFormatter from './NumberFormatter';
import { MOCK_ACCOUNTS } from '../constants';
import { CreditCard, ExternalLink, ShieldCheck, Plus } from 'lucide-react';

interface AccountsDetailProps {
  isPrivacyMode: boolean;
  isLoading: boolean;
}

const AccountsDetail: React.FC<AccountsDetailProps> = ({ isPrivacyMode, isLoading }) => {
  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-3xl font-black">My Accounts</h2>
          <p className="opacity-50 text-sm">Managing 5 linked institutions</p>
        </div>
        <button className="px-6 py-3 neu-flat rounded-full text-neu-accent font-black text-xs flex items-center gap-2 hover:scale-105 active:scale-95 transition-all">
          <Plus size={16} />
          ADD ACCOUNT
        </button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <Card className="bg-gradient-to-br from-neu-accent/5 to-transparent">
          <div className="flex items-center gap-3 mb-6">
            <ShieldCheck className="text-emerald-500" />
            <h3 className="font-bold">Total Liquidity</h3>
          </div>
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <span className="text-sm opacity-50 font-medium">Available Now</span>
              <NumberFormatter value={1800000} prefix="$" isPrivacyMode={isPrivacyMode} className="text-2xl font-black text-neu-accent" />
            </div>
            <div className="w-full h-3 neu-inset rounded-full overflow-hidden">
              <div className="h-full bg-neu-accent rounded-full w-[70%] transition-all duration-1000" />
            </div>
            <p className="text-[10px] opacity-40 text-center font-bold uppercase tracking-widest">70% of targets achieved</p>
          </div>
        </Card>

        <Card>
          <div className="flex items-center gap-3 mb-6">
            <CreditCard className="text-amber-500" />
            <h3 className="font-bold">Debt Ratio</h3>
          </div>
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <span className="text-sm opacity-50 font-medium">Debt to Asset</span>
              <span className="text-2xl font-black text-rose-500">86.2%</span>
            </div>
            <div className="w-full h-3 neu-inset rounded-full overflow-hidden">
              <div className="h-full bg-rose-500 rounded-full w-[86%]" />
            </div>
            <p className="text-[10px] opacity-40 text-center font-bold uppercase tracking-widest">High utilization warning</p>
          </div>
        </Card>
      </div>

      <div className="space-y-6">
        <h3 className="text-xs font-black opacity-30 uppercase tracking-[0.2em]">Institutional Breakdown</h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {MOCK_ACCOUNTS.map((acc) => (
            <div key={acc.id} className="neu-flat p-8 rounded-[2.5rem] group hover:neu-inset transition-all cursor-pointer">
              <div className="flex justify-between items-start mb-6">
                <div className="w-12 h-12 rounded-2xl neu-flat flex items-center justify-center text-neu-accent group-hover:scale-110 transition-transform">
                  <CreditCard size={24} />
                </div>
                <button className="text-slate-300 hover:text-neu-accent transition-colors">
                  <ExternalLink size={18} />
                </button>
              </div>
              <div>
                <p className="text-[10px] font-black opacity-40 uppercase tracking-widest mb-1">{acc.bank}</p>
                <h4 className="text-lg font-black mb-4">{acc.name}</h4>
                <div className="flex items-baseline gap-2">
                  <NumberFormatter 
                    value={acc.balance} 
                    currency={acc.currency} 
                    isPrivacyMode={isPrivacyMode} 
                    className={`text-2xl font-black ${acc.isDebt ? 'text-rose-500' : 'text-slate-800 dark:text-white'}`}
                  />
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default AccountsDetail;
