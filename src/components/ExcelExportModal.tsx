import React from 'react';
import { FileSpreadsheet, Image as ImageIcon, X, Download } from 'lucide-react';

interface ExcelExportModalProps {
  isOpen: boolean;
  data?: {
    type: string;
    cols: string[];
    rows: string[][];
    title: string;
  } | null;
  title?: string;
  cols?: string[];
  rows?: string[][];
  type?: string;
  techizatImages?: Record<string, string>;
  onClose: () => void;
  onDownloadText?: (type: string, cols: string[], rows: string[][], title: string) => void;
  onDownloadWithImages?: (type: string, cols: string[], rows: string[][], title: string) => void;
  onExportTextOnly?: () => void;
  onExportWithImages?: () => void;
  isExportingWithImages?: boolean;
}

export const ExcelExportModal: React.FC<ExcelExportModalProps> = ({
  isOpen,
  data,
  title,
  cols,
  rows,
  type,
  techizatImages,
  onClose,
  onDownloadText,
  onDownloadWithImages,
  onExportTextOnly,
  onExportWithImages,
  isExportingWithImages = false,
}) => {
  if (!isOpen) return null;

  const exportTitle = title || data?.title || "Envanter Listesi";
  const exportCols = cols || data?.cols || [];
  const exportRows = rows || data?.rows || [];
  const exportType = type || data?.type || "techizat";

  const handleTextExport = () => {
    if (onExportTextOnly) {
      onExportTextOnly();
    } else if (onDownloadText) {
      onDownloadText(exportType, exportCols, exportRows, exportTitle);
      onClose();
    }
  };

  const handleImagesExport = () => {
    if (onExportWithImages) {
      onExportWithImages();
    } else if (onDownloadWithImages) {
      onDownloadWithImages(exportType, exportCols, exportRows, exportTitle);
    }
  };

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm animate-fade-in">
      <div className="bg-white rounded-3xl p-6 sm:p-8 max-w-lg w-full shadow-2xl border border-slate-100 flex flex-col gap-6 relative">
        <div className="flex items-center justify-between border-b border-slate-100 pb-4">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-800 flex items-center justify-center shadow-inner">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h3 className="text-base font-black text-slate-800 uppercase tracking-wider">
                EXCEL AKTARMA SEÇENEKLERİ
              </h3>
              <p className="text-xs text-slate-500 font-medium">
                {exportRows.length} kayıt dışa aktarılmaya hazır
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-400 hover:text-slate-600 flex items-center justify-center font-bold text-sm transition-all cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="bg-slate-50 rounded-2xl p-4 border border-slate-200 text-xs text-slate-700 leading-relaxed font-medium">
          <span className="font-bold text-slate-900 block mb-1">📋 {exportTitle}</span>
          Lütfen indirmek istediğiniz Excel formatını seçiniz. Hızlı indirme için metin tabanlı format önerilir.
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {/* Option 1: Fast Text Excel */}
          <button
            type="button"
            onClick={handleTextExport}
            className="p-5 rounded-2xl border-2 border-emerald-600/30 hover:border-emerald-600 bg-emerald-50/50 hover:bg-emerald-50 text-left transition-all group cursor-pointer flex flex-col justify-between gap-3"
          >
            <div className="flex items-center justify-between">
              <div className="p-2.5 rounded-xl bg-emerald-600 text-white group-hover:scale-110 transition-transform">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <span className="text-[10px] font-mono font-black uppercase text-emerald-800 bg-emerald-200/70 px-2 py-0.5 rounded-full">
                HIZLI
              </span>
            </div>
            <div>
              <h4 className="font-black text-sm text-slate-900 group-hover:text-emerald-900 transition-colors">
                Metin Formatı (.xls)
              </h4>
              <p className="text-[11px] text-slate-500 mt-1 leading-normal">
                Tüm sütunları, tarihler ve renkli durum göstergeleri ile anında indirir.
              </p>
            </div>
            <div className="pt-2 flex items-center gap-1.5 text-xs font-bold text-emerald-700">
              <Download className="w-3.5 h-3.5" /> İndir
            </div>
          </button>

          {/* Option 2: Image Embedded Excel */}
          <button
            type="button"
            disabled={isExportingWithImages}
            onClick={handleImagesExport}
            className="p-5 rounded-2xl border-2 border-indigo-600/30 hover:border-indigo-600 bg-indigo-50/50 hover:bg-indigo-50 text-left transition-all group cursor-pointer flex flex-col justify-between gap-3 disabled:opacity-60 disabled:cursor-not-allowed"
          >
            <div className="flex items-center justify-between">
              <div className="p-2.5 rounded-xl bg-indigo-600 text-white group-hover:scale-110 transition-transform">
                <ImageIcon className="w-5 h-5" />
              </div>
              <span className="text-[10px] font-mono font-black uppercase text-indigo-800 bg-indigo-200/70 px-2 py-0.5 rounded-full">
                FOTOĞRAFLI (.XLSX)
              </span>
            </div>
            <div>
              <h4 className="font-black text-sm text-slate-900 group-hover:text-indigo-900 transition-colors">
                Görselleri İçeren (.xlsx)
              </h4>
              <p className="text-[11px] text-slate-500 mt-1 leading-normal">
                Drive üzerindeki malzeme fotoğraflarını hücre içine gömerek oluşturur.
              </p>
            </div>
            <div className="pt-2 flex items-center gap-1.5 text-xs font-bold text-indigo-700">
              {isExportingWithImages ? "Hazırlanıyor..." : <><Download className="w-3.5 h-3.5" /> İndir</>}
            </div>
          </button>
        </div>

        <div className="flex justify-end pt-2">
          <button
            type="button"
            onClick={onClose}
            className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl cursor-pointer transition-all"
          >
            Vazgeç
          </button>
        </div>
      </div>
    </div>
  );
};
