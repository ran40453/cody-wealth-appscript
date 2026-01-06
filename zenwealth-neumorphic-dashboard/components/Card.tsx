
import React from 'react';

interface CardProps {
  children: React.ReactNode;
  className?: string;
  onClick?: () => void;
  title?: string;
  subtitle?: string;
  isLoading?: boolean;
}

const Card: React.FC<CardProps> = ({ 
  children, 
  className = "", 
  onClick, 
  title, 
  subtitle,
  isLoading 
}) => {
  return (
    <div 
      onClick={onClick}
      className={`
        neu-flat rounded-[2.5rem] p-6 transition-all duration-300 
        ${onClick ? 'cursor-pointer hover:scale-[1.02] hover:-translate-y-1 active:scale-[0.98]' : ''} 
        ${className}
      `}
    >
      {(title || subtitle) && (
        <div className="mb-6 flex items-end justify-between">
          <div>
            {title && <h3 className="text-lg font-bold tracking-tight">{title}</h3>}
            {subtitle && <p className="text-xs opacity-50 font-medium uppercase tracking-wider">{subtitle}</p>}
          </div>
          {onClick && <div className="text-[10px] bg-neu-accent/10 text-neu-accent px-2 py-1 rounded-full font-bold">VIEW ALL</div>}
        </div>
      )}
      {isLoading ? (
        <div className="space-y-4 animate-pulse">
          <div className="h-20 bg-slate-300 dark:bg-slate-700 rounded-2xl opacity-20" />
          <div className="h-4 w-1/2 bg-slate-300 dark:bg-slate-700 rounded-2xl opacity-20" />
        </div>
      ) : children}
    </div>
  );
};

export default Card;
