import React, { useState } from 'react';
import { History, X, Search, FileSpreadsheet, ArrowRight } from 'lucide-react';
import { AuditLogEntry } from '../types';

interface AuditTrailModalProps {
  isOpen: boolean;
  onClose: () => void;
  auditLogs?: AuditLogEntry[];
  logs?: AuditLogEntry[];
  onClearLogs?: () => void;
}

export const AuditTrailModal: React.FC<AuditTrailModalProps> = ({
  isOpen,
  onClose,
  auditLogs,
  logs,
  onClearLogs,
}) => {
  const allLogs = auditLogs || logs || [];
  const [search, setSearch] = useState<string>('');
  const [filterUnit, setFilterUnit] = useState<string>('ALL');

  if (!isOpen) return null;

  const filteredLogs = allLogs.filter(log => {
    if (filterUnit !== 'ALL' && log.unit !== filterUnit) return false;
    if (search.trim()) {
      const q = search.toLowerCase();
      return (
        log.itemName.toLowerCase().includes(q) ||
        log.pn.toLowerCase().includes(q) ||
        log.fieldName.toLowerCase().includes(q) ||
        log.oldValue.toLowerCase().includes(q) ||
        log.newValue.toLowerCase().includes(q)
      );
    }
    return true;
  });

  const exportAuditToExcel = () => {
    let html = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head><meta charset="utf-8">
      <style>
        table { border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; }
        th { background-color: #0b3d1d; color: #ffffff; padding: 8px; font-size: 11px; }
        td { border: 1px solid #e2e8f0; padding: 6px 8px; font-size: 11px; }
      </style>
      </head>
      <body>
        <table>
          <thead>
            <tr>
              <th colspan="8" style="background-color: #0b3d1d; color: white; font-size: 14px; text-align: center; height: 35px;">
                DÜZENLEME VE DEĞİŞİKLİK GEÇMİŞİ (AUDIT LOGS)
              </th>
            </tr>
            <tr>
              <th>SIRA</th>
              <th>TARİH / SAAT</th>
              <th>BİRİM</th>
              <th>ÜRÜN / PARÇA ADI</th>
              <th>P/N NO</th>
              <th>DEĞİŞTİRİLEN ALAN</th>
              <th>ESKİ DEĞER</th>
              <th>YENİ DEĞER</th>
            </tr>
          </thead>
          <tbody>
    `;

    filteredLogs.forEach((log, idx) => {
      html += `
        <tr>
          <td style="text-align: center;">${idx + 1}</td>
          <td style="text-align: center;">${log.timestamp}</td>
          <td>${log.unit}</td>
          <td>${log.itemName}</td>
          <td style="text-align: center;">${log.pn || "-"}</td>
          <td style="font-weight: bold;">${log.fieldName}</td>
          <td style="color: #b91c1c; text-decoration: line-through;">${log.oldValue || "(Boş)"}</td>
          <td style="color: #15803d; font-weight: bold;">${log.newValue || "(Boş)"}</td>
        </tr>
      `;
    });

    html += `</tbody></table></body></html>`;

    const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Duzenleme_Gecmisi_${new Date().toISOString().slice(0, 10)}.xls`;
    link.click();
  };

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-3 sm:p-6 bg-slate-950/80 backdrop-blur-md animate-fade-in">
      <div className="bg-white rounded-3xl shadow-2xl max-w-5xl w-full max-h-[92vh] flex flex-col overflow-hidden border border-slate-200">
        {/* Header */}
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-amber-600/80 flex items-center justify-center text-white shadow-inner">
              <History className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-black uppercase tracking-wider flex items-center gap-2">
                DÜZENLEME GEÇMİŞİ & DEĞİŞİKLİK TAKİP TABLOSU
              </h3>
              <p className="text-[11px] text-slate-400 font-mono">
                {auditLogs.length} Kayıtlı Değişiklik İşlemi • Kim, Neyi, Ne Zaman Değiştirdi
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center font-bold text-sm transition-all cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Filters and Controls */}
        <div className="bg-slate-50 px-6 py-3 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-3 flex-1 min-w-[280px]">
            <div className="relative flex-1">
              <Search className="w-4 h-4 absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                placeholder="Parça adı, P/N, değiştirilen alan veya değer ara..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                className="w-full pl-9 pr-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none"
              />
            </div>

            <select
              value={filterUnit}
              onChange={(e) => setFilterUnit(e.target.value)}
              className="px-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
            >
              <option value="ALL">⚡ TÜM BİRİMLER</option>
              <option value="at802">AT-802F</option>
              <option value="bell429">Bell 429</option>
              <option value="t70">T-70</option>
              <option value="hangar">Hangar Yer Destek</option>
              <option value="kara_araclari">Kara Araçları</option>
              <option value="depo_sarf">Depo Sarf</option>
              <option value="depo_kimyasal">Depo Kimyasal</option>
            </select>
          </div>

          <button
            type="button"
            onClick={exportAuditToExcel}
            className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl flex items-center gap-2 shadow-sm transition-all cursor-pointer"
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>Excel Olarak İndir</span>
          </button>
        </div>

        {/* Content Table */}
        <div className="flex-1 overflow-y-auto p-6">
          {filteredLogs.length === 0 ? (
            <div className="p-16 text-center text-slate-400 text-xs font-medium border border-dashed border-slate-200 rounded-2xl">
              {search.trim() ? "Arama kriterlerine uygun değişiklik kaydı bulunamadı." : "Henüz sistemde kaydedilmiş bir düzenleme hareketi bulunmuyor. Teçhizat ve malzeme kartlarında yaptığınız her değişiklik burada kronolojik olarak saklanır."}
            </div>
          ) : (
            <div className="overflow-x-auto border border-slate-200 rounded-2xl shadow-sm">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-900 text-slate-200 uppercase tracking-wider text-[10px] font-black select-none">
                  <tr>
                    <th className="px-3 py-3 text-center">SIRA</th>
                    <th className="px-3 py-3">TARİH / SAAT</th>
                    <th className="px-3 py-3">BİRİM</th>
                    <th className="px-4 py-3">ÜRÜN / PARÇA ADI</th>
                    <th className="px-3 py-3">P/N</th>
                    <th className="px-3 py-3">DEĞİŞTİRİLEN ALAN</th>
                    <th className="px-4 py-3">ESKİ DEĞER</th>
                    <th className="px-4 py-3">YENİ DEĞER</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredLogs.map((log, idx) => (
                    <tr key={log.id || idx} className="hover:bg-slate-50 transition-colors">
                      <td className="px-3 py-2.5 text-center font-mono font-bold text-slate-500">{idx + 1}</td>
                      <td className="px-3 py-2.5 font-mono text-slate-700 whitespace-nowrap">{log.timestamp}</td>
                      <td className="px-3 py-2.5 font-bold text-slate-800 uppercase font-mono text-[10px]">{log.unit}</td>
                      <td className="px-4 py-2.5 font-bold text-slate-900">{log.itemName}</td>
                      <td className="px-3 py-2.5 font-mono text-slate-600">{log.pn || "-"}</td>
                      <td className="px-3 py-2.5 font-bold text-indigo-900 bg-indigo-50/50">
                        {log.fieldName}
                      </td>
                      <td className="px-4 py-2.5 font-mono text-rose-700 line-through bg-rose-50/30">
                        {log.oldValue || "(Boş)"}
                      </td>
                      <td className="px-4 py-2.5 font-mono font-bold text-emerald-800 bg-emerald-50/30 flex items-center gap-1.5">
                        <ArrowRight className="w-3 h-3 text-emerald-600 shrink-0" />
                        <span>{log.newValue || "(Boş)"}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
