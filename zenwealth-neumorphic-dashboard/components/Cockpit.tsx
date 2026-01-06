
import React from 'react';
import Card from './Card';
import NumberFormatter from './NumberFormatter';
import { FX_RATE, MOCK_STATS } from '../constants';
import { TrendingUp, TrendingDown, DollarSign } from 'lucide-react';

interface CockpitProps {
  isPrivacyMode: boolean;
  isLoading: boolean;
}

const Cockpit: React.FC<CockpitProps> = ({ isPrivacyMode, isLoading }) => {
  const stats = MOCK_STATS;
  const isPositive = stats.dailyPnL >= 0;

  return (
    <section className="space-y-6">
      <div className="flex flex-col md:flex-row gap-6">
        {/* Main Asset Summary */}
        <Card className="flex-grow min-w-[320px] bg-gradient-to-br from-neu-light to-[#f0f4f8] dark:from-neu-dark dark:to-[#1a1c1e]">
          <div className="flex justify-between items-start mb-4">
            <span className="text-xs font-bold opacity-40 uppercase tracking-widest">Global Net Worth</span>
            <div className={`
              px-3 py-1 rounded-full text-[10px] font-bold flex items-center gap-1
              ${isPositive ? 'neu-pill-green' : 'neu-pill-red'}
            `}>
              {isPositive ? <TrendingUp size={10} /> : <TrendingDown size={10} />}
              {isPositive ? '+' : ''}
              <NumberFormatter value={stats.dailyPnLPerc} suffix="%" isPrivacyMode={false} precision={1} />
            </div>
          </div>
          
          <div className="mb-8">
            <h2 className="text-5xl md:text-6xl font-black mb-2">
              <NumberFormatter 
                value={stats.totalAssets} 
                isPrivacyMode={isPrivacyMode} 
                prefix="$" 
                className="text-slate-800 dark:text-white"
              />
            </h2>
            <div className="flex items-center gap-2 opacity-50">
              <p className="text-xs font-medium">Updated {stats.lastUpdate}</p>
            </div>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-6 pt-6 border-t border-slate-200/20">
            <KPIItem 
              label="Net Assets" 
              value={stats.netAssets} 
              isPrivacyMode={isPrivacyMode} 
              color="text-neu-accent"
            />
            <KPIItem 
              label="Liquid Cash" 
              value={stats.cash} 
              isPrivacyMode={isPrivacyMode} 
            />
            <KPIItem 
              label="Disposable" 
              value={stats.disposable} 
              isPrivacyMode={isPrivacyMode} 
            />
            <KPIItem 
              label="USD Reserves" 
              value={stats.usdBalance} 
              currency="USD"
              isPrivacyMode={isPrivacyMode} 
              color="text-amber-500"
            />
          </div>
        </Card>
      </div>
    </section>
  );
};

const KPIItem = ({ label, value, currency = "TWD", isPrivacyMode, color = "" }: any) => (
  <div className="space-y-1">
    <p className="text-[10px] font-bold opacity-40 uppercase tracking-wider">{label}</p>
    <div className={`text-lg font-bold ${color}`}>
      <NumberFormatter 
        value={value} 
        currency={currency} 
        isPrivacyMode={isPrivacyMode} 
        precision={currency === 'USD' ? 2 : 0}
      />
    </div>
  </div>
);

export default Cockpit;
