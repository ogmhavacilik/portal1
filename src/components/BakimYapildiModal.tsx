import React, { useState, useEffect } from 'react';
import { Calendar, CheckCircle2, X, AlertCircle, Wrench } from 'lucide-react';

interface BakimYapildiModalProps {
  isOpen: boolean;
  onClose: () => void;
  itemName?: string;
  prevSon: string;
  prevGelecek: string;
  intervalDays: number;
  initialSon: string;
  initialGelecek: string;
  onConfirm: (finalSon: string, finalGelecek: string) => Promise<void> | void;
}

export const BakimYapildiModal: React.FC<BakimYapildiModalProps> = ({
  isOpen,
  onClose,
  itemName,
  prevSon,
  prevGelecek,
  intervalDays,
  initialSon,
  initialGelecek,
  onConfirm
}) => {
  const [selectedSon, setSelectedSon] = useState<string>(initialSon);
  const [selectedGelecek, setSelectedGelecek] = useState<string>(initialGelecek);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);

  useEffect(() => {
    if (isOpen) {
      setSelectedSon(initialSon);
      setSelectedGelecek(initialGelecek);
    }
  }, [isOpen, initialSon, initialGelecek]);

  if (!isOpen) return null;

  // Helper to convert DD.MM.YYYY to YYYY-MM-DD for date inputs
  const dmyToYmd = (dmy: string): string => {
    if (!dmy) return '';
    const clean = dmy.split(/[\r\n;]+/)[0].trim();
    const parts = clean.split(/[\.\/-]/);
    if (parts.length === 3) {
      if (parts[2].length === 4) {
        return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
      }
      if (parts[0].length === 4) {
        return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
      }
    }
    return '';
  };

  // Helper to convert YYYY-MM-DD to DD.MM.YYYY
  const ymdToDmy = (ymd: string): string => {
    if (!ymd) return '';
    const parts = ymd.split('-');
    if (parts.length === 3) {
      return `${parts[2].padStart(2, '0')}.${parts[1].padStart(2, '0')}.${parts[0]}`;
    }
    return ymd;
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedGelecek) {
      alert('Lütfen bir sonraki bakım tarihini giriniz.');
      return;
    }
    try {
      setIsSubmitting(true);
      await onConfirm(selectedSon || initialSon, selectedGelecek);
      onClose();
    } catch (err) {
      console.error(err);
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[99999] flex items-center justify-center p-4 bg-black/75 backdrop-blur-sm animate-fade-in">
      <div className="bg-white rounded-3xl shadow-2xl border-2 border-emerald-600/30 max-w-lg w-full flex flex-col overflow-hidden text-slate-800">
        {/* Header */}
        <div className="bg-[#0b3d1d] text-white px-6 py-4 flex items-center justify-between border-b border-emerald-800">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-emerald-700/60 flex items-center justify-center">
              <Wrench className="w-5 h-5 text-emerald-300" />
            </div>
            <div>
              <h2 className="text-sm sm:text-base font-black uppercase tracking-wider text-white">
                BAKIM YAPILDI - TARİH BELİRLEME
              </h2>
              <p className="text-[11px] text-emerald-200/90 font-mono">
                BİR SONRAKİ BAKIM TARİHİ SEÇİMİ
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-emerald-900/60 hover:bg-emerald-800 text-slate-200 flex items-center justify-center transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <form onSubmit={handleSave} className="p-6 flex flex-col gap-5">
          {itemName && (
            <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3.5 flex flex-col gap-1">
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-wider">
                Teçhizat / Malzeme
              </span>
              <span className="text-xs sm:text-sm font-black text-slate-800 line-clamp-2">
                {itemName}
              </span>
            </div>
          )}

          {/* Previous maintenance info */}
          <div className="grid grid-cols-2 gap-3 bg-emerald-50/60 border border-emerald-200/70 rounded-2xl p-3.5">
            <div>
              <span className="text-[10px] font-bold text-slate-500 uppercase block">Önceki Son Bakım:</span>
              <span className="text-xs font-black text-slate-700 font-mono">{prevSon || '-'}</span>
            </div>
            <div>
              <span className="text-[10px] font-bold text-slate-500 uppercase block">Mevcut Gelecek Bakım:</span>
              <span className="text-xs font-black text-rose-700 font-mono">{prevGelecek || '-'}</span>
            </div>
            <div className="col-span-2 pt-2 border-t border-emerald-200/60 flex items-center justify-between">
              <span className="text-[11px] font-bold text-emerald-900">Hesaplanan Bakım Aralığı:</span>
              <span className="text-xs font-black text-emerald-800 font-mono bg-emerald-100 px-2 py-0.5 rounded-md">
                {intervalDays} Gün ({Math.round(intervalDays / 30)} Ay)
              </span>
            </div>
          </div>

          {/* Explanation Alert */}
          <div className="flex items-start gap-2.5 bg-sky-50 border border-sky-200 rounded-2xl p-3 text-sky-900 text-xs">
            <AlertCircle className="w-4 h-4 text-sky-600 shrink-0 mt-0.5" />
            <p className="leading-snug text-[11px]">
              Kural gereği takvimde bugünün tarihi değil, <strong>sistemde seçilen tarihin bir önceki tarih ile arasındaki fark ({intervalDays} gün)</strong> kadar ileri atılmış olan <strong>{initialGelecek}</strong> tarihi varsayılan olarak seçilmiştir.
            </p>
          </div>

          {/* Date inputs */}
          <div className="flex flex-col gap-4">
            <div className="flex flex-col gap-1.5">
              <label className="text-xs font-black text-slate-700 uppercase tracking-wide flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5 text-emerald-600" />
                <span>Yapılan Bakım / Kontrol Tarihi</span>
              </label>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <input
                  type="date"
                  value={dmyToYmd(selectedSon)}
                  onChange={(e) => {
                    const dmy = ymdToDmy(e.target.value);
                    setSelectedSon(dmy);
                  }}
                  className="px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                />
                <input
                  type="text"
                  value={selectedSon}
                  onChange={(e) => setSelectedSon(e.target.value)}
                  placeholder="GG.AA.YYYY"
                  className="px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none font-mono"
                />
              </div>
            </div>

            <div className="flex flex-col gap-1.5">
              <label className="text-xs font-black text-emerald-800 uppercase tracking-wide flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5 text-emerald-600" />
                <span>Bir Sonraki Bakım Tarihi</span>
              </label>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <input
                  type="date"
                  value={dmyToYmd(selectedGelecek)}
                  onChange={(e) => {
                    const dmy = ymdToDmy(e.target.value);
                    setSelectedGelecek(dmy);
                  }}
                  className="px-3 py-2 bg-emerald-50 border-2 border-emerald-500 rounded-xl text-xs font-black text-emerald-900 focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                />
                <input
                  type="text"
                  value={selectedGelecek}
                  onChange={(e) => setSelectedGelecek(e.target.value)}
                  placeholder="GG.AA.YYYY"
                  className="px-3 py-2 bg-emerald-50 border-2 border-emerald-500 rounded-xl text-xs font-black text-emerald-900 focus:ring-2 focus:ring-emerald-500 focus:outline-none font-mono"
                />
              </div>
            </div>
          </div>

          {/* Action buttons */}
          <div className="flex items-center justify-end gap-3 pt-3 border-t border-slate-100">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-black rounded-xl transition-all cursor-pointer"
            >
              Vazgeç
            </button>
            <button
              type="submit"
              disabled={isSubmitting}
              className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white text-xs font-black rounded-xl transition-all flex items-center gap-2 shadow-lg shadow-emerald-600/30 cursor-pointer disabled:opacity-50"
            >
              <CheckCircle2 className="w-4 h-4" />
              <span>{isSubmitting ? 'Kaydediliyor...' : 'Bakımı Tamamla ve Faal Yap'}</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
