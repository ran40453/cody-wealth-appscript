
import React from 'react';
import Card from './Card';
import NumberFormatter from './NumberFormatter';
import { Search, Filter, ArrowUp, ArrowDown } from 'lucide-react';

interface RecordsDetailProps {
  isPrivacyMode: boolean;
  isLoading: boolean;
}

const RecordsDetail: React.FC<RecordsDetailProps> = ({ isPrivacyMode, isLoading }) => {
  const transactions = [
    { date: '2023-10-27', title: 'Uber Eats', category: 'Food', amount: -450, bank: 'Richart' },
    { date: '2023-10-26', title: 'Salary Deposit', category: 'Income', amount: 82000, bank: 'CTBC' },
    { date: '2023-10-25', title: 'Netflix Subscription', category: 'Media', amount: -390, bank: 'Cathay' },
    { date: '2023-10-24', title: 'Starbucks Coffee', category: 'Food', amount: -155, bank: 'CTBC' },
    { date: '2023-10-23', title: 'Amazon Cloud', category: 'Services', amount: -1200, bank: 'Schwab' },
    { date: '2023-10-22', title: 'Dividend: VOO', category: 'Income', amount: 3200, bank: 'Schwab' },
  ];

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4">
        <div>
          <h2 className="text-3xl font-black">History</h2>
          <p className="opacity-50 text-sm">Reviewing last 30 days of activity</p>
        </div>
        <div className="flex gap-4">
          <div className="neu-inset px-4 py-2 rounded-full flex items-center gap-2">
            <Search size={16} className="opacity-40" />
            <input 
              type="text" 
              placeholder="Search..." 
              className="bg-transparent border-none outline-none text-xs font-bold w-32"
            />
          </div>
          <button className="w-10 h-10 rounded-full neu-flat flex items-center justify-center">
            <Filter size={16} />
          </button>
        </div>
      </div>

      <div className="space-y-4">
        {transactions.map((tx, i) => (
          <div key={i} className="neu-flat p-6 rounded-3xl flex items-center justify-between group hover:neu-inset transition-all">
            <div className="flex items-center gap-4">
              <div className={`w-12 h-12 rounded-2xl flex items-center justify-center ${tx.amount > 0 ? 'bg-emerald-500/10 text-emerald-500' : 'bg-rose-500/10 text-rose-500'}`}>
                {tx.amount > 0 ? <ArrowUp size={20} /> : <ArrowDown size={20} />}
              </div>
              <div>
                <h4 className="font-bold group-hover:text-neu-accent transition-colors">{tx.title}</h4>
                <div className="flex items-center gap-2 text-[10px] font-bold opacity-40 uppercase tracking-tighter">
                  <span>{tx.date}</span>
                  <span className="w-1 h-1 rounded-full bg-slate-400" />
                  <span>{tx.bank}</span>
                </div>
              </div>
            </div>
            <div className="text-right">
              <NumberFormatter 
                value={tx.amount} 
                prefix={tx.amount > 0 ? '+' : ''} 
                isPrivacyMode={isPrivacyMode} 
                className={`text-lg font-black ${tx.amount > 0 ? 'text-emerald-500' : 'text-slate-800 dark:text-white'}`}
              />
              <p className="text-[10px] font-bold opacity-30 uppercase">{tx.category}</p>
            </div>
          </div>
        ))}
      </div>

      <button className="w-full py-6 neu-flat rounded-3xl text-sm font-black opacity-40 hover:opacity-100 transition-opacity">
        LOAD MORE TRANSACTIONS
      </button>
    </div>
  );
};

export default RecordsDetail;
