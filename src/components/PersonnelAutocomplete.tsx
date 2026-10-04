import React, { useState, useEffect, useRef } from 'react';
import { User, ChevronDown, Check } from 'lucide-react';
import { getStoredPersonnelList, syncPersonnelListFromScript } from '../utils/personnelData';

interface PersonnelAutocompleteProps {
  value: string;
  onChange: (val: string) => void;
  placeholder?: string;
  className?: string;
  id?: string;
}

export const PersonnelAutocomplete: React.FC<PersonnelAutocompleteProps> = ({
  value,
  onChange,
  placeholder = "Personel Adı Soyadı...",
  className = "",
  id
}) => {
  const [personnelList, setPersonnelList] = useState<string[]>([]);
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    // Initial from cache / defaults
    setPersonnelList(getStoredPersonnelList());
    // Background sync from Google script
    syncPersonnelListFromScript().then(list => {
      if (list && list.length > 0) setPersonnelList(list);
    });
  }, []);

  // Close when clicked outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Filtered recommendations based on user typed text - ONLY show if search query is entered
  const filtered = personnelList.filter(name => {
    if (!value || !value.trim()) return false;
    const q = value.toLowerCase().trim();
    return name.toLowerCase().includes(q);
  }).slice(0, 8);

  return (
    <div ref={containerRef} className="relative w-full">
      <div className="relative flex items-center">
        <input
          id={id}
          type="text"
          value={value}
          onChange={(e) => {
            onChange(e.target.value);
            setIsOpen(true);
          }}
          onFocus={() => {
            onChange(''); // Clear the input field on focus so text disappears instantly
            setIsOpen(false); // Do not open list on empty focus
          }}
          onClick={() => {
            onChange(''); // Clear the input field on click as well
            setIsOpen(false); // Close suggestions on click
          }}
          placeholder={placeholder}
          className={className || "w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold text-slate-800 focus:outline-none focus:border-emerald-600 font-sans"}
        />
      </div>

      {isOpen && filtered.length > 0 && (
        <div className="absolute top-full left-0 right-0 z-[10500] mt-1 bg-white border border-slate-300 rounded-xl shadow-2xl max-h-48 overflow-y-auto divide-y divide-slate-100 text-xs font-sans">
          <div className="p-1.5 bg-slate-100 text-[10px] font-black uppercase text-slate-600 font-mono flex items-center justify-between">
            <span>Teknisyen / Personel Seç:</span>
            <span className="text-emerald-700 font-bold">{filtered.length} Sonuç</span>
          </div>
          {filtered.map((person, idx) => (
            <div
              key={idx}
              onClick={() => {
                onChange(person);
                setIsOpen(false);
              }}
              className="px-3 py-1.5 hover:bg-emerald-50 cursor-pointer flex items-center justify-between transition-colors text-slate-800 font-bold"
            >
              <div className="flex items-center gap-2">
                <User className="w-3 h-3 text-emerald-600 shrink-0" />
                <span>{person}</span>
              </div>
              {value === person && <Check className="w-3.5 h-3.5 text-emerald-600" />}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
