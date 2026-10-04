import React, { useState, useRef, useEffect } from 'react';
import { Search, Check } from 'lucide-react';
import { DepoItem } from './DepoManagementModal';

interface PnAutocompleteProps {
  value: string;
  onChange: (pn: string, matchedItem?: DepoItem) => void;
  inventory: DepoItem[];
  placeholder?: string;
  className?: string;
  required?: boolean;
}

export const PnAutocomplete: React.FC<PnAutocompleteProps> = ({
  value,
  onChange,
  inventory,
  placeholder = "P/N arayınız veya yazınız...",
  className = "",
  required = false
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [query, setQuery] = useState(value);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setQuery(value);
  }, [value]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const cleanQuery = (query || '').toLowerCase().trim();

  const filteredItems = inventory.filter(item => {
    if (!cleanQuery) return true;
    const p = (item.partNumber || item.pn || '').toLowerCase();
    const d = (item.description || item.name || '').toLowerCase();
    return p.includes(cleanQuery) || d.includes(cleanQuery);
  }).slice(0, 15);

  const handleSelect = (item: DepoItem) => {
    const itemPn = item.partNumber || item.pn || '';
    setQuery(itemPn);
    onChange(itemPn, item);
    setIsOpen(false);
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newQuery = e.target.value;
    setQuery(newQuery);
    setIsOpen(true);
    
    // Check for exact PN match
    const exactMatch = inventory.find(i => 
      (i.partNumber || i.pn || '').toLowerCase() === newQuery.trim().toLowerCase()
    );
    
    onChange(newQuery, exactMatch);
  };

  return (
    <div ref={containerRef} className="relative w-full">
      <div className="relative flex items-center">
        <Search className="w-3.5 h-3.5 absolute left-2.5 text-slate-400 pointer-events-none" />
        <input
          type="text"
          value={query}
          onChange={handleInputChange}
          onFocus={() => setIsOpen(true)}
          placeholder={placeholder}
          required={required}
          className={`w-full pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold text-xs focus:outline-none focus:border-blue-500 shadow-inner ${className}`}
        />
      </div>

      {isOpen && (
        <div className="absolute top-full left-0 right-0 z-[11000] mt-1 bg-slate-950 border-2 border-blue-500/60 rounded-xl shadow-2xl max-h-52 overflow-y-auto divide-y divide-slate-800">
          {filteredItems.length === 0 ? (
            <div className="p-3 text-[11px] text-slate-400 font-sans italic text-center">
              Eşleşen Part Number (P/N) bulunamadı
            </div>
          ) : (
            filteredItems.map((item, idx) => {
              const itemPn = item.partNumber || item.pn || '-';
              const itemDesc = item.description || item.name || '-';
              const isSelected = itemPn.toLowerCase() === cleanQuery;

              return (
                <div
                  key={idx}
                  onClick={() => handleSelect(item)}
                  className={`p-2.5 hover:bg-blue-900/50 cursor-pointer text-xs flex items-center justify-between gap-2 transition ${
                    isSelected ? 'bg-blue-950/80 text-blue-200' : 'text-slate-200'
                  }`}
                >
                  <div className="min-w-0 flex-1">
                    <div className="font-mono font-bold text-blue-300 flex items-center gap-1.5">
                      <span>{itemPn}</span>
                      {isSelected && <Check className="w-3.5 h-3.5 text-emerald-400" />}
                    </div>
                    <div className="text-[10px] text-slate-400 font-sans truncate mt-0.5">
                      {itemDesc}
                    </div>
                  </div>
                  <span className="text-[10px] font-mono font-bold bg-slate-800 text-slate-300 border border-slate-700 px-2 py-0.5 rounded shrink-0">
                    Stok: {item.ankaraMevcut || item.toplamStok || 0}
                  </span>
                </div>
              );
            })
          )}
        </div>
      )}
    </div>
  );
};
