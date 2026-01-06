
export type AppRoute = 'overview' | 'invest' | 'records' | 'accounts';

export interface Account {
  id: string;
  name: string;
  bank: string;
  balance: number;
  currency: 'TWD' | 'USD';
  isDebt?: boolean;
}

export interface Investment {
  symbol: string;
  name: string;
  value: number;
  roi: number;
  isFund: boolean;
}

export interface FinancialStats {
  totalAssets: number;
  netAssets: number;
  cash: number;
  disposable: number;
  usdBalance: number;
  lastUpdate: string;
  dailyPnL: number;
  dailyPnLPerc: number;
}
