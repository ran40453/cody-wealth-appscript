
import { Account, Investment, FinancialStats } from './types';

export const FX_RATE = 32.48; // USD to TWD

export const MOCK_ACCOUNTS: Account[] = [
  { id: '1', name: 'Primary Savings', bank: 'CTBC', balance: 1250400, currency: 'TWD' },
  { id: '2', name: 'Emergency Fund', bank: 'Cathay', balance: 450000, currency: 'TWD' },
  { id: '3', name: 'US Brokerage', bank: 'Schwab', balance: 15200, currency: 'USD' },
  { id: '4', name: 'Home Loan', bank: 'Fubon', balance: -4200000, currency: 'TWD', isDebt: true },
  { id: '5', name: 'Digital Wallet', bank: 'Richart', balance: 88500, currency: 'TWD' },
];

export const MOCK_INVESTMENTS: Investment[] = [
  { symbol: 'VOO', name: 'S&P 500 ETF', value: 450000, roi: 12.5, isFund: false },
  { symbol: 'TSLA', name: 'Tesla Inc', value: 120000, roi: -4.2, isFund: false },
  { symbol: 'FUND_A', name: 'Global Tech Growth Fund', value: 300000, roi: 8.7, isFund: true },
  { symbol: 'FUND_B', name: 'Fixed Income High Yield', value: 150000, roi: 3.2, isFund: true },
];

export const MOCK_STATS: FinancialStats = {
  totalAssets: 4850000,
  netAssets: 650000,
  cash: 1800000,
  disposable: 120000,
  usdBalance: 15200,
  lastUpdate: '2023-10-27 14:45',
  dailyPnL: 12450,
  dailyPnLPerc: 1.2
};
