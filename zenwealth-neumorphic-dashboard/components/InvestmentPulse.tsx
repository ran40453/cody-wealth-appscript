
import React from 'react';
import Card from './Card';
import NumberFormatter from './NumberFormatter';
import { MOCK_INVESTMENTS } from '../constants';
import { PieChart, ArrowUpRight, ArrowDownRight } from 'lucide-react';

interface InvestmentPulseProps {
  isPrivacyMode: boolean;
  isLoading: boolean;
  onNavigate: () => void;
}

const InvestmentPulse: React.FC<InvestmentPulseProps> = ({ isPrivacyMode, isLoading, onNavigate }) => {
  const funds = MOCK_INVESTMENTS.filter(i => i.isFund);
  const stocks = MOCK_INVESTMENTS.filter(i => !i.isFund);

  return (
    <Card title="Investment Pulse" subtitle="Equities & Funds" onClick={onNavigate} isLoading={isLoading}>
      <div className="space-y-6">
        {/* Funds Section */}
        <div className="space-y-4">
          <div className="flex items-center gap-2 mb-2">
            <PieChart size={14} className="text-neu-accent" />
            <h4 className="text-xs font-bold opacity-60 uppercase">Managed Funds</h4>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div className="neu-inset p-4 rounded-2xl">
              <p className="text-[10px] font-bold opacity-40 mb-1">Total Market Value</p>
              <NumberFormatter value={450000} prefix="$" isPrivacyMode={isPrivacyMode} className="text-xl font-bold" />
            </div>
            <div className="neu-inset p-4 rounded-2xl">
              <p className="text-[10px] font-bold opacity-40 mb-1">Dividends Recv.</p>
              <NumberFormatter value={12400} prefix="$" isPrivacyMode={isPrivacyMode} className="text-xl font-bold text-emerald-500" />
            </div>
          </div>
        </div>

        {/* Stocks List */}
        <div className="space-y-3">
          <div className="flex items-center gap-2 mb-2">
            <ArrowUpRight size={14} className="text-neu-accent" />
            <h4 className="text-xs font-bold opacity-60 uppercase">Live Equity Holdings</h4>
          </div>
          <div className="space-y-2">
            {stocks.map(stock => (
              <div key={stock.symbol} className="flex items-center justify-between p-3 neu-inset rounded-2xl">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl neu-flat flex items-center justify-center text-[10px] font-black bg-slate-100 dark:bg-slate-800">
                    {stock.symbol}
                  </div>
                  <div>
                    <p className="text-sm font-bold">{stock.name}</p>
                    <p className="text-[10px] opacity-40 uppercase">{stock.symbol}</p>
                  </div>
                </div>
                <div className="text-right">
                  <div className="text-sm font-bold">
                    <NumberFormatter value={stock.value} isPrivacyMode={isPrivacyMode} prefix="$" />
                  </div>
                  <div className={`text-[10px] font-black flex items-center justify-end gap-1 ${stock.roi >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                    {stock.roi >= 0 ? <ArrowUpRight size={10} /> : <ArrowDownRight size={10} />}
                    {Math.abs(stock.roi)}%
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </Card>
  );
};

export default InvestmentPulse;
