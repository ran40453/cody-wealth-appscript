
import React from 'react';

interface NumberFormatterProps {
  value: number;
  currency?: string;
  isPrivacyMode?: boolean;
  className?: string;
  prefix?: string;
  suffix?: string;
  precision?: number;
}

const NumberFormatter: React.FC<NumberFormatterProps> = ({ 
  value, 
  currency, 
  isPrivacyMode, 
  className = "", 
  prefix = "",
  suffix = "",
  precision = 0
}) => {
  if (isPrivacyMode) {
    return <span className={`${className} font-mono blur-[4px] select-none pointer-events-none`}>••••••</span>;
  }

  const formatter = new Intl.NumberFormat('en-US', {
    minimumFractionDigits: precision,
    maximumFractionDigits: precision,
  });

  const formattedValue = formatter.format(Math.abs(value));
  const sign = value < 0 ? '-' : '';

  return (
    <span className={`${className} font-mono tracking-tighter`}>
      {sign}{prefix}{currency ? currency + ' ' : ''}{formattedValue}{suffix}
    </span>
  );
};

export default NumberFormatter;
