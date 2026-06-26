import React, { useState, useRef, useEffect } from 'react';
import { Search, ChevronDown, Check, X } from 'lucide-react';

interface MultiSelectProps {
  label: string;
  options: string[];
  selected: string[];
  onChange: (selected: string[]) => void;
}

export function MultiSelect({ label, options, selected, onChange }: MultiSelectProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const filteredOptions = options.filter(opt => opt.toLowerCase().includes(searchTerm.toLowerCase()));

  const handleSelectAll = () => onChange(options);
  const handleClearAll = () => onChange([]);

  const toggleOption = (option: string) => {
    if (selected.includes(option)) {
      onChange(selected.filter(item => item !== option));
    } else {
      onChange([...selected, option]);
    }
  };

  const removeOption = (e: React.MouseEvent, option: string) => {
    e.stopPropagation();
    onChange(selected.filter(item => item !== option));
  };

  const isAllSelected = selected.length === 0 || selected.length === options.length;

  return (
    <div className="relative flex flex-col gap-1 w-full" ref={wrapperRef}>
      <label className="text-xs font-medium text-slate-400 ml-1">{label}</label>
      <div 
        className="min-h-[42px] bg-white/5 border border-white/10 rounded-lg p-2 cursor-pointer flex flex-wrap gap-1 items-center hover:bg-white/10 transition-colors"
        onClick={() => setIsOpen(!isOpen)}
      >
        {isAllSelected ? (
          <span className="text-sm text-slate-400 px-1 flex-1">All Selected</span>
        ) : (
          <div className="flex flex-wrap gap-1 flex-1">
            {selected.slice(0, 2).map(opt => (
              <span key={opt} className="bg-indigo-500/20 text-indigo-300 text-xs px-2 py-1 rounded-md flex items-center gap-1 border border-indigo-500/30">
                <span className="truncate max-w-[90px]">{opt}</span>
                <X className="w-3 h-3 hover:text-white" onClick={(e) => removeOption(e, opt)} />
              </span>
            ))}
            {selected.length > 2 && (
              <span className="bg-white/10 text-slate-300 text-xs px-2 py-1 rounded-md flex items-center border border-white/10">
                +{selected.length - 2} more
              </span>
            )}
          </div>
        )}
        <ChevronDown className={`w-4 h-4 text-slate-400 ml-auto flex-shrink-0 transition-transform ${isOpen ? 'rotate-180' : ''}`} />
      </div>

      {isOpen && (
        <div className="absolute top-full left-0 right-0 mt-1 bg-slate-800 border border-slate-700 rounded-lg shadow-xl z-50 overflow-hidden flex flex-col max-h-72">
          <div className="p-2 border-b border-slate-700">
            <div className="relative">
              <Search className="w-4 h-4 absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                className="w-full bg-slate-900 border border-slate-700 text-white text-sm rounded-md pl-8 pr-3 py-2 focus:outline-none focus:border-indigo-500"
                placeholder="Search..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                onClick={(e) => e.stopPropagation()}
              />
            </div>
            <div className="flex items-center justify-between mt-2 px-1">
              <button 
                className="text-xs text-indigo-400 hover:text-indigo-300 font-medium"
                onClick={(e) => { e.stopPropagation(); handleSelectAll(); }}
              >Select All</button>
              <button 
                className="text-xs text-slate-400 hover:text-slate-300 font-medium"
                onClick={(e) => { e.stopPropagation(); handleClearAll(); }}
              >Clear All</button>
            </div>
          </div>
          <div className="overflow-y-auto flex-1 p-1">
            {filteredOptions.length === 0 ? (
              <div className="p-3 text-center text-sm text-slate-500">No results found</div>
            ) : (
              filteredOptions.map(opt => {
                const isSelected = selected.includes(opt) || isAllSelected;
                return (
                  <div 
                    key={opt}
                    className="flex items-center gap-2 px-3 py-2 hover:bg-slate-700/50 rounded-md cursor-pointer text-sm text-slate-200"
                    onClick={(e) => { e.stopPropagation(); toggleOption(opt); }}
                  >
                    <div className={`w-4 h-4 rounded border flex items-center justify-center flex-shrink-0 transition-colors ${isSelected ? 'bg-indigo-500 border-indigo-500' : 'border-slate-600'}`}>
                      {isSelected && <Check className="w-3 h-3 text-white" />}
                    </div>
                    <span className="truncate">{opt}</span>
                  </div>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
}
