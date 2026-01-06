
import React from 'react';
import Card from './Card';
import NumberFormatter from './NumberFormatter';
import { MOCK_INVESTMENTS } from '../constants';
import { ArrowUpRight, ArrowDownRight, Target, Briefcase } from 'lucide-react';

interface InvestDetailProps {
  isPrivacyMode: boolean;
  isLoading: boolean;
}

const InvestDetail: React.FC<InvestDetailProps> = ({ isPrivacyMode, isLoading }) => {
  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-3xl font-black">Investment Portfolio</h2>
          <p className="opacity-50 text-sm">Diversification across 8 asset classes</p>
        </div>
        <button className="w-12 h-12 rounded-full neu-flat flex items-center justify-center text-neu-accent">
          <Target size={20} />
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
        <Card className="md:col-span-2">
          <div className="flex items-center gap-3 mb-8">
            <Briefcase className="text-neu-accent" />
            <h3 className="font-bold text-lg">Asset Allocation</h3>
          </div>
          <div className="h-48 flex items-end gap-2 px-2">
            {[65, 40, 85, 30, 55, 90, 45].map((h, i) => (
              <div key={i} className="flex-1 flex flex-col items-center gap-2">
                <div 
                  className="w-full neu-flat rounded-t-xl transition-all duration-1000 bg-neu-accent/20" 
                  style={{ height: `${h}%` }}
                />
                <span className="text-[10px] opacity-40 font-bold">M{i+1}</span>
              </div>
            ))}
          </div>
        </Card>

        <Card title="Performance" subtitle="All Time">
          <div className="space-y-6 text-center">
            <div className="neu-inset p-6 rounded-[2rem]">
              <p className="text-[10px] font-bold opacity-40 uppercase mb-2">Portfolio ROI</p>
              <h4 className="text-4xl font-black text-emerald-500">+14.2%</h4>
            </div>
            <div className="flex justify-between px-4">
              <div>
                <p className="text-[10px] opacity-40 font-bold">BEST</p>
                <p className="font-bold text-emerald-500">VOO</p>
              </div>
              <div className="w-px bg-slate-200 dark:bg-slate-800" />
              <div>
                <p className="text-[10px] opacity-40 font-bold">WORST</p>
                <p className="font-bold text-rose-500">TSLA</p>
              </div>
            </div>
          </div>
        </Card>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        {MOCK_INVESTMENTS.map((inv) => (
          <Card key={inv.symbol} className="flex items-center justify-between p-8">
            <div className="flex items-center gap-4">
              <div className="w-14 h-14 rounded-2xl neu-flat flex items-center justify-center font-black text-neu-accent">
                {inv.symbol.substring(0, 2)}
              </div>
              <div>
                <h4 className="font-bold">{inv.name}</h4>
                <p className="text-[10px] opacity-40 font-bold uppercase">{inv.symbol}</p>
              </div>
            </div>
            <div className="text-right">
              <NumberFormatter value={inv.value} prefix="$" isPrivacyMode={isPrivacyMode} className="text-xl font-bold" />
              <div className={`text-xs font-black flex items-center justify-end gap-1 ${inv.roi >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                {inv.roi >= 0 ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}
                {Math.abs(inv.roi)}%
              </div>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
};

export default InvestDetail;
