
import React from 'react';
import Card from './Card';
import NumberFormatter from './NumberFormatter';
import { Receipt, Calendar, CreditCard } from 'lucide-react';

interface ActivityLogProps {
  isPrivacyMode: boolean;
  isLoading: boolean;
  onNavigate: () => void;
}

const ActivityLog: React.FC<ActivityLogProps> = ({ isPrivacyMode, isLoading, onNavigate }) => {
  const activities = [
    { id: 1, type: 'posted', title: 'Salary Credit', bank: 'CTBC', amount: 82000, date: 'Oct 25', isPositive: true },
    { id: 2, type: 'posted', title: 'Apple Store', bank: 'Cathay', amount: -2900, date: 'Oct 26', isPositive: false },
    { id: 3, type: 'projected', title: 'Mortgage Payment', bank: 'Fubon', amount: -35000, date: 'Nov 01', isPositive: false },
    { id: 4, type: 'projected', title: 'Utility Bill', bank: 'CTBC', amount: -1500, date: 'Nov 03', isPositive: false },
  ];

  return (
    <Card title="Activity Log" subtitle="Cashflow & Projections" onClick={onNavigate} isLoading={isLoading}>
      <div className="space-y-4">
        {/* Comparison Summary */}
        <div className="flex gap-4 mb-6">
          <div className="flex-1 text-center py-4 neu-inset rounded-3xl border-b-2 border-emerald-500/30">
            <p className="text-[10px] font-bold opacity-40 uppercase mb-1">Posted Sum</p>
            <NumberFormatter value={79100} isPrivacyMode={isPrivacyMode} className="text-lg font-bold text-emerald-500" />
          </div>
          <div className="flex-1 text-center py-4 neu-inset rounded-3xl border-b-2 border-amber-500/30">
            <p className="text-[10px] font-bold opacity-40 uppercase mb-1">Projected</p>
            <NumberFormatter value={-36500} isPrivacyMode={isPrivacyMode} className="text-lg font-bold text-amber-500" />
          </div>
        </div>

        {/* List */}
        <div className="space-y-2">
          {activities.map(act => (
            <div key={act.id} className={`flex items-center justify-between p-4 rounded-[1.5rem] transition-all hover:neu-inset ${act.type === 'projected' ? 'opacity-60 grayscale-[0.5]' : ''}`}>
              <div className="flex items-center gap-4">
                <div className={`w-10 h-10 rounded-full neu-flat flex items-center justify-center ${act.type === 'projected' ? 'text-amber-500' : 'text-neu-accent'}`}>
                  {act.type === 'projected' ? <Calendar size={18} /> : <Receipt size={18} />}
                </div>
                <div>
                  <h5 className="text-sm font-bold">{act.title}</h5>
                  <p className="text-[10px] opacity-40 font-medium uppercase">{act.bank} • {act.date}</p>
                </div>
              </div>
              <div className="text-right">
                <div className={`text-sm font-black ${act.isPositive ? 'text-emerald-500' : 'text-rose-500'}`}>
                  {act.isPositive ? '+' : ''}
                  <NumberFormatter value={act.amount} isPrivacyMode={isPrivacyMode} />
                </div>
                {act.type === 'projected' && (
                  <span className="text-[8px] font-bold px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-600 uppercase">Pending</span>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    </Card>
  );
};

export default ActivityLog;
