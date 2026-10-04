import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Wrench,
  Search,
  SlidersHorizontal,
  RefreshCw,
  PlusCircle,
  Download,
  Boxes,
  Sparkles,
  Trash2,
  Plane,
  Layers,
  ArrowLeft,
  CheckCircle2,
  AlertTriangle,
  FileSpreadsheet,
  X,
  Lock,
  ExternalLink,
  Package
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { cleanAndFormatDateString, unmergeAndFillWorksheet } from '../utils/driveExcelSync';
import { publishCrossSystemEvent, subscribeCrossSystemEvents } from '../utils/crossSystemBridge';
import { downloadStandaloneSystemZip } from '../utils/zipExporter';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export interface FleetUnit {
  key: string;
  name: string;
  shortLabel: string;
  driveFileName: string;
  storageKey: string;
  iconType: 'plane' | 'heli' | 'hangar';
}

export const FLEET_UNITS: FleetUnit[] = [
  {
    key: 'at802',
    name: 'AT-802F Yangın Söndürme Uçağı',
    shortLabel: 'AT-802F',
    driveFileName: 'hava_araçları_yer_destek_at-802.xlsx',
    storageKey: 'excel_techizat_at802_data',
    iconType: 'plane'
  },
  {
    key: 'bell429',
    name: 'Bell-429 Çok Amaçlı Helikopter',
    shortLabel: 'BELL-429',
    driveFileName: 'hava_araçları_yer_destek_bell-429.xlsx',
    storageKey: 'excel_techizat_bell429_data',
    iconType: 'heli'
  },
  {
    key: 't70',
    name: 'T-70 Genel Maksat Helikopteri',
    shortLabel: 'T-70',
    driveFileName: 'hava_araçları_yer_destek_t-70.xlsx',
    storageKey: 'excel_techizat_t70_data',
    iconType: 'heli'
  },
  {
    key: 't70_bumbi_backet',
    name: 'T-70 Bumbi Bucket & Yangın Kiti',
    shortLabel: 'T-70 BUMBİ',
    driveFileName: 'hava_araçları_yer_destek_t-70_bumbi_backet.xlsx',
    storageKey: 'excel_techizat_t70_bumbi_backet_data',
    iconType: 'heli'
  },
  {
    key: 't70_helitak',
    name: 'T-70 Helitak Kurtarma & Aletleri',
    shortLabel: 'T-70 HELİTAK',
    driveFileName: 'hava_araçları_yer_destek_t-70_helitak.xlsx',
    storageKey: 'excel_techizat_t70_helitak_data',
    iconType: 'heli'
  },
  {
    key: 'b360',
    name: 'Beechcraft King Air B-360',
    shortLabel: 'B-360',
    driveFileName: 'hava_araçları_yer_destek_b-360.xlsx',
    storageKey: 'excel_techizat_b360_data',
    iconType: 'plane'
  },
  {
    key: 'c650',
    name: 'Cessna Citation C-650',
    shortLabel: 'C-650',
    driveFileName: 'hava_araçları_yer_destek_c-650.xlsx',
    storageKey: 'excel_techizat_c650_data',
    iconType: 'plane'
  },
  {
    key: 'hangar',
    name: 'Hangar Yer Destek ve Özel Takımlar',
    shortLabel: 'HANGAR',
    driveFileName: 'hava_araçları_yer_destek_hangar.xlsx',
    storageKey: 'excel_techizat_hangar_data',
    iconType: 'hangar'
  }
];

export const TABLE_COLUMNS = [
  "SIRA NO",
  "TEÇHİZAT ADI",
  "PARÇA NO (P/N) / MODEL",
  "SERİ NO (S/N)",
  "MİKTAR / KAPASİTE",
  "BULUNDUĞU YER",
  "DURUMU",
  "KALİBRASYONA TABİ",
  "SON KONTROL / KALİBRASYON / BAKIM",
  "GELECEK KONTROL / KALİBRASYON / BAKIM",
  "SON KONTROLÜ YAPAN FİRMA",
  "AÇIKLAMA"
];

export function YerDestekApp({ initialFleet, isEmbedded }: { initialFleet?: string; isEmbedded?: boolean }) {
  // Parse URL parameters
  const params = new URLSearchParams(window.location.search);
  const fleetFromUrl = params.get('fleet') || params.get('unit') || initialFleet || 'at802';
  const searchFromUrl = params.get('search') || '';
  const embeddedMode = isEmbedded ?? (params.get('embedded') === 'true');

  const [selectedFleetKey, setSelectedFleetKey] = useState<string>(fleetFromUrl);
  const [techizatSearchQuery, setTechizatSearchQuery] = useState<string>(searchFromUrl);
  const [techizatFirmaFilter, setTechizatFirmaFilter] = useState<string>('');
  const [techizatDurumFilter, setTechizatDurumFilter] = useState<string>('');
  const [techizatColorFilter, setTechizatColorFilter] = useState<string>('ALL');
  const [selectedColorFilter, setSelectedColorFilter] = useState<string>('all');
  const [sortByColor, setSortByColor] = useState<boolean>(false);
  const [isFilterOpen, setIsFilterOpen] = useState<boolean>(false);
  const [selectedItems, setSelectedItems] = useState<Record<string, boolean>>({});
  const [activeMatchIdx, setActiveMatchIdx] = useState<number>(0);
  const [isLoadingDrive, setIsLoadingDrive] = useState<boolean>(false);
  const [notification, setNotification] = useState<string | null>(null);

  // Row Edit Modal State
  const [activeEditRow, setActiveEditRow] = useState<{ idx: number; row: string[] } | null>(null);
  // New Product Modal State
  const [isNewProductOpen, setIsNewProductOpen] = useState<boolean>(false);
  const [newRowData, setNewRowData] = useState<Record<string, string>>({});

  // Password Lock Modal State
  const [isPasswordModalOpen, setIsPasswordModalOpen] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [passwordError, setPasswordError] = useState<boolean>(false);
  const [pendingPasswordAction, setPendingPasswordAction] = useState<(() => void) | null>(null);

  // Local storage for each unit data
  const [fleetData, setFleetData] = useState<Record<string, string[][]>>(() => {
    const initial: Record<string, string[][]> = {};
    for (const u of FLEET_UNITS) {
      try {
        const stored = localStorage.getItem(u.storageKey);
        if (stored) {
          initial[u.key] = JSON.parse(stored);
        } else {
          initial[u.key] = [];
        }
      } catch {
        initial[u.key] = [];
      }
    }
    return initial;
  });

  // Sayfalama (Pagination) State'leri - Her sayfada en fazla 50 ürün
  const [pageSize, setPageSize] = useState<number>(50);
  const [currentPage, setCurrentPage] = useState<number>(1);

  const showNotification = (msg: string) => {
    setNotification(msg);
    setTimeout(() => setNotification(null), 4000);
  };

  const currentUnit = FLEET_UNITS.find(u => u.key === selectedFleetKey) || FLEET_UNITS[0];
  const currentRows = fleetData[selectedFleetKey] || [];

  const saveRows = (unitKey: string, rows: string[][]) => {
    setFleetData(prev => ({ ...prev, [unitKey]: rows }));
    const unit = FLEET_UNITS.find(u => u.key === unitKey);
    if (unit) {
      try {
        localStorage.setItem(unit.storageKey, JSON.stringify(rows));
      } catch (e) {}
    }
    publishCrossSystemEvent('SYSTEM_DATA_UPDATED', 'yer-destek', { unitKey, count: rows.length });
  };

  // Cross-system event listener
  useEffect(() => {
    const unsub = subscribeCrossSystemEvents((ev) => {
      if (ev.type === 'REFRESH_ALL' || (ev.type === 'SYSTEM_DATA_UPDATED' && ev.sourceSystem !== 'yer-destek')) {
        // reload from storage
        const updated: Record<string, string[][]> = {};
        for (const u of FLEET_UNITS) {
          try {
            const val = localStorage.getItem(u.storageKey);
            if (val) updated[u.key] = JSON.parse(val);
          } catch {}
        }
        setFleetData(prev => ({ ...prev, ...updated }));
      }
    });
    return unsub;
  }, []);

  // Sayfa açılışında veya birim değiştiğinde Google Drive ile arka planda anlık senkronize et (Depo modülüyle birebir aynı açılış mimarisi)
  useEffect(() => {
    pullFromDrive();
  }, [selectedFleetKey]);

  // Calculate unique companies and statuses
  const uniqueFirmalar = useMemo(() => {
    const s = new Set<string>();
    for (const r of currentRows) {
      const f = (r[10] || '').trim();
      if (f && f !== '-' && f !== '--') s.add(f);
    }
    return Array.from(s).sort();
  }, [currentRows]);

  const uniqueDurumlar = useMemo(() => {
    const s = new Set<string>();
    for (const r of currentRows) {
      const d = (r[6] || '').trim().toUpperCase();
      if (d && d !== '-' && d !== '--') s.add(d);
    }
    return Array.from(s).sort();
  }, [currentRows]);

  // Date difference calculation in days
  const getDaysUntil = (dateStr: string): number | null => {
    if (!dateStr || dateStr === '-' || dateStr === '--') return null;
    const parts = dateStr.split(/[./-]/);
    if (parts.length < 3) return null;
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    const target = new Date(year, month, day);
    if (isNaN(target.getTime())) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    target.setHours(0, 0, 0, 0);
    return Math.round((target.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  };

  const getRowColorStatus = (row: string[]): 'red' | 'orange' | 'green' | 'neutral' => {
    const isKalibTabi = (row[7] || '').trim().toUpperCase();
    if (isKalibTabi === 'HAYIR' || isKalibTabi === 'HAYIR (MUAFIYET)' || isKalibTabi === 'MUAFIYET') {
      return 'neutral';
    }
    const days = getDaysUntil(row[9]);
    if (days === null) return 'neutral';
    if (days < 0) return 'red';
    if (days <= 90) return 'orange';
    return 'green';
  };

  // Filtered and sorted rows
  const processedRows = useMemo(() => {
    let result = currentRows.map((row, idx) => ({ row, idx }));

    // Search query
    if (techizatSearchQuery.trim()) {
      const q = techizatSearchQuery.toLowerCase().trim();
      result = result.filter(({ row }) =>
        row.some(cell => String(cell || '').toLowerCase().includes(q))
      );
    }

    // Company filter
    if (techizatFirmaFilter) {
      result = result.filter(({ row }) => (row[10] || '').trim() === techizatFirmaFilter);
    }

    // Durum filter
    if (techizatDurumFilter) {
      result = result.filter(({ row }) => (row[6] || '').trim().toUpperCase() === techizatDurumFilter);
    }

    // Color filter 1
    if (techizatColorFilter !== 'ALL') {
      result = result.filter(({ row }) => {
        const color = getRowColorStatus(row);
        if (techizatColorFilter === 'RED') return color === 'red';
        if (techizatColorFilter === 'ORANGE') return color === 'orange';
        if (techizatColorFilter === 'GREEN') return color === 'green';
        if (techizatColorFilter === 'GRAY') return color === 'neutral';
        return true;
      });
    }

    // Color filter 2
    if (selectedColorFilter !== 'all') {
      result = result.filter(({ row }) => {
        const color = getRowColorStatus(row);
        return color === selectedColorFilter;
      });
    }

    // Sort by color (Red -> Orange -> Green -> Neutral)
    if (sortByColor) {
      const rank = { red: 0, orange: 1, green: 2, neutral: 3 };
      result.sort((a, b) => {
        const cA = getRowColorStatus(a.row);
        const cB = getRowColorStatus(b.row);
        if (rank[cA] !== rank[cB]) return rank[cA] - rank[cB];
        const dA = getDaysUntil(a.row[9]) ?? 9999;
        const dB = getDaysUntil(b.row[9]) ?? 9999;
        return dA - dB;
      });
    }

    return result;
  }, [currentRows, techizatSearchQuery, techizatFirmaFilter, techizatDurumFilter, techizatColorFilter, selectedColorFilter, sortByColor]);

  // Reset page when filters or unit change
  useEffect(() => {
    setCurrentPage(1);
  }, [selectedFleetKey, techizatSearchQuery, techizatFirmaFilter, techizatDurumFilter, techizatColorFilter, selectedColorFilter, sortByColor, pageSize]);

  // Paginated Rows (en fazla 50 ürün listelenir)
  const paginatedRows = useMemo(() => {
    if (pageSize <= 0 || pageSize >= 99999) return processedRows;
    const startIndex = (currentPage - 1) * pageSize;
    return processedRows.slice(startIndex, startIndex + pageSize);
  }, [processedRows, currentPage, pageSize]);

  const totalPages = Math.max(1, Math.ceil(processedRows.length / (pageSize > 0 ? pageSize : 50)));

  // Pull from Google Drive (Canlı ve Otomatik Dosya Eşleştirmeli)
  const pullFromDrive = async () => {
    setIsLoadingDrive(true);
    showNotification(`Google Drive'dan [${currentUnit.driveFileName}] güncel verileri canlı alınıyor...`);
    try {
      // 1. Call server proxy /api/read-excel-from-drive with specific unit fileName
      const res = await fetch('/api/read-excel-from-drive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: currentUnit.driveFileName,
          folderId: DRIVE_FOLDER_ID
        })
      });

      const json = await res.json();
      if (json && json.status === 'success' && json.base64) {
        const binaryString = atob(json.base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }

        const wb = XLSX.read(bytes, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        if (sheet) {
          unmergeAndFillWorksheet(sheet);
          const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
          if (rawRows && rawRows.length > 0) {
            // Find header row
            let headerIdx = 0;
            for (let r = 0; r < Math.min(10, rawRows.length); r++) {
              const rowStr = (rawRows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
              if (rowStr.includes('TEÇHİZAT') || rowStr.includes('TECHIZAT') || rowStr.includes('MALZEME') || rowStr.includes('PARÇA') || rowStr.includes('P/N') || rowStr.includes('SIRA')) {
                headerIdx = r;
                break;
              }
            }

            const dataRows = rawRows.slice(headerIdx + 1).filter(r => r && r.some(c => String(c || '').trim() !== ''));
            if (dataRows.length > 0) {
              const cleanedRows = dataRows.map((r, i) => {
                const rowArr = [...r];
                while (rowArr.length < 12) rowArr.push('-');
                rowArr[0] = String(i + 1);
                if (rowArr[8]) rowArr[8] = cleanAndFormatDateString(String(rowArr[8]));
                if (rowArr[9]) rowArr[9] = cleanAndFormatDateString(String(rowArr[9]));
                return rowArr.map(c => String(c !== undefined && c !== null ? c : ''));
              });

              saveRows(selectedFleetKey, cleanedRows);
              showNotification(`✅ ${currentUnit.shortLabel} güncellendi: ${cleanedRows.length} kayıt Google Drive'dan aktarıldı.`);
              return;
            }
          }
        }
      }

      showNotification(`ℹ️ Drive üzerinde [${currentUnit.driveFileName}] için hazır tablonuz görüntülendi.`);
    } catch (e: any) {
      console.warn("Drive sync error:", e);
      showNotification(`❌ Drive bağlantı hatası: ${e?.message || 'Bilinmeyen hata'}`);
    } finally {
      setIsLoadingDrive(false);
    }
  };

  // Export to Excel
  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();
    const wsData = [
      TABLE_COLUMNS,
      ...processedRows.map(({ row }) => TABLE_COLUMNS.map((_, cIdx) => row[cIdx] || ''))
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, currentUnit.shortLabel);
    XLSX.writeFile(wb, `${currentUnit.driveFileName.replace('.xlsx', '')}_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showNotification('✅ Excel dosyası başarıyla indirildi.');
  };

  const requirePassword = (action: () => void) => {
    setPendingPasswordAction(() => action);
    setPasswordInput('');
    setPasswordError(false);
    setIsPasswordModalOpen(true);
  };

  const handlePasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (passwordInput === '1839' || passwordInput === '802') {
      setIsPasswordModalOpen(false);
      setPasswordError(false);
      if (pendingPasswordAction) {
        pendingPasswordAction();
        setPendingPasswordAction(null);
      }
    } else {
      setPasswordError(true);
    }
  };

  // Delete selected items
  const handleDeleteSelected = () => {
    const selectedIndices = new Set(Object.keys(selectedItems).filter(k => selectedItems[k]).map(Number));
    if (selectedIndices.size === 0) return;
    requirePassword(() => {
      const remaining = currentRows.filter((_, idx) => !selectedIndices.has(idx)).map((r, i) => {
        r[0] = String(i + 1);
        return r;
      });
      saveRows(selectedFleetKey, remaining);
      setSelectedItems({});
      showNotification(`✅ Seçilen ${selectedIndices.size} teçhizat silindi.`);
    });
  };

  const handleSaveNewProduct = () => {
    const newRow = [
      String(currentRows.length + 1),
      newRowData['1'] || '',
      newRowData['2'] || '-',
      newRowData['3'] || 'N/A',
      newRowData['4'] || '1',
      newRowData['5'] || 'Y/D HANGAR',
      newRowData['6'] || 'FAAL',
      newRowData['7'] || 'EVET',
      newRowData['8'] || cleanAndFormatDateString(new Date().toLocaleDateString('tr-TR')),
      newRowData['9'] || '',
      newRowData['10'] || 'YETKİLİ FİRMA',
      newRowData['11'] || ''
    ];
    const updated = [...currentRows, newRow];
    saveRows(selectedFleetKey, updated);
    setIsNewProductOpen(false);
    setNewRowData({});
    showNotification('✅ Yeni teçhizat başarıyla eklendi.');
  };

  const handleSaveEditRow = () => {
    if (!activeEditRow) return;
    const updated = [...currentRows];
    updated[activeEditRow.idx] = activeEditRow.row;
    saveRows(selectedFleetKey, updated);
    setActiveEditRow(null);
    showNotification('✅ Teçhizat kaydı güncellendi.');
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-emerald-500 selection:text-white">
      {/* NOTIFICATION TOAST */}
      {notification && (
        <div className="fixed top-4 right-4 z-[999] px-4 py-3 rounded-2xl bg-emerald-950 border-2 border-emerald-500 text-emerald-100 text-xs font-black shadow-2xl flex items-center gap-3 animate-fade-in">
          <CheckCircle2 className="w-4 h-4 text-emerald-400" />
          <span>{notification}</span>
        </div>
      )}

      {/* TOP HEADER (If standalone) */}
      {!embeddedMode && (
        <header className="bg-slate-900/90 border-b border-slate-800 backdrop-blur-md sticky top-0 z-40 px-4 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-4 select-none">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-800 flex items-center justify-center shadow-lg shadow-emerald-950/50 border border-emerald-400/30">
              <Wrench className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-black tracking-tight text-white uppercase">
                  Hava Araçları Yer Destek & Özel Aletler Takip Sistemi
                </h1>
                <span className="px-2 py-0.5 rounded-full text-[10px] font-mono font-bold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
                  MÜSTAKİL SİSTEM
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Orman Genel Müdürlüğü Havacılık Dairesi Başkanlığı
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => downloadStandaloneSystemZip('yer-destek')}
              className="px-3.5 py-1.5 rounded-xl bg-indigo-950/80 hover:bg-indigo-900 text-indigo-200 border border-indigo-500/50 text-xs font-bold flex items-center gap-1.5 transition-all shadow-md active:scale-95 cursor-pointer"
              title="Bu sistemi bağımsız zip web sitesi olarak indir"
            >
              <Package className="w-3.5 h-3.5 text-indigo-400" />
              <span>SİTEYİ ZİP İNDİR</span>
            </button>
            <a
              href="/"
              className="px-3.5 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white text-xs font-bold flex items-center gap-1.5 transition-colors border border-slate-700"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Ana Portala Dön</span>
            </a>
          </div>
        </header>
      )}

      {/* MAIN CONTAINER */}
      <div className="flex-1 p-4 lg:p-6 max-w-[1700px] mx-auto w-full flex flex-col">
        {/* FLEET SELECTION TABS (IDENTICAL TO MAIN APP FLEET PICKER) */}
        <div className="bg-slate-900 rounded-3xl p-2.5 mb-6 border border-slate-800 shadow-xl overflow-x-auto scrollbar-none flex items-center gap-2 select-none shrink-0">
          {FLEET_UNITS.map(u => {
            const count = (fleetData[u.key] || []).length;
            const isSelected = selectedFleetKey === u.key;
            return (
              <button
                key={u.key}
                type="button"
                onClick={() => setSelectedFleetKey(u.key)}
                className={`px-4 py-2.5 rounded-2xl text-xs font-black whitespace-nowrap transition-all flex items-center gap-2 cursor-pointer ${
                  isSelected
                    ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-950/40 border border-emerald-400/50 scale-102'
                    : 'bg-slate-800/80 text-slate-400 hover:text-slate-200 hover:bg-slate-800 border border-slate-700/50'
                }`}
              >
                {u.iconType === 'plane' ? (
                  <Plane className="w-4 h-4" />
                ) : u.iconType === 'heli' ? (
                  <Plane className="w-4 h-4 rotate-45" />
                ) : (
                  <Layers className="w-4 h-4" />
                )}
                <span>{u.shortLabel}</span>
                <span className={`px-2 py-0.5 rounded-full text-[10px] font-mono font-bold ${
                  isSelected ? 'bg-black/30 text-emerald-200' : 'bg-slate-900 text-slate-400'
                }`}>
                  {count}
                </span>
              </button>
            );
          })}
        </div>

        {/* COLLAPSIBLE FILTER AND SEARCH DRAWER */}
        <div className="mb-6 select-none print:hidden">
          <button
            type="button"
            onClick={() => setIsFilterOpen(prev => !prev)}
            className="w-full flex items-center justify-between px-5 py-3.5 bg-gradient-to-r from-slate-900 via-slate-850 to-slate-900 hover:from-slate-800 hover:to-slate-850 text-white rounded-2xl border border-slate-700/80 shadow-md transition-all cursor-pointer group"
          >
            <div className="flex items-center gap-3">
              <div className="p-2 bg-emerald-500/20 text-emerald-400 rounded-xl border border-emerald-500/30 group-hover:scale-105 transition-transform">
                <Search className="w-4 h-4" />
              </div>
              <div className="text-left">
                <div className="text-xs font-black tracking-wider uppercase text-slate-100 flex items-center gap-2">
                  <span>🔍 FİLTRELEME VE DETAYLI ARAMA PANELİ</span>
                  {(techizatSearchQuery || techizatFirmaFilter || techizatDurumFilter || techizatColorFilter !== 'ALL' || selectedColorFilter !== 'all' || sortByColor) && (
                    <span className="px-2 py-0.5 bg-amber-500 text-slate-950 font-black text-[10px] rounded-full animate-pulse">
                      FİLTRE AKTİF
                    </span>
                  )}
                </div>
                <div className="text-[11px] text-slate-400 font-medium">
                  {isFilterOpen 
                    ? "Filtre ve arama seçeneklerini gizlemek için tıklayın (▲ Kapat)" 
                    : "Malzeme arama, firma, durum, renk kodu filtrelerini açmak için tıklayın (▼ Alta Doğru Aç)"}
                </div>
              </div>
            </div>
            
            <div className="flex items-center gap-2">
              <span className={`text-xs font-black px-3.5 py-1.5 rounded-xl border transition-all ${
                isFilterOpen ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40' : 'bg-slate-800 text-slate-300 border-slate-700'
              }`}>
                {isFilterOpen ? '▲ KAPAT' : '▼ FİLTRELERİ GÖSTER'}
              </span>
            </div>
          </button>

          {isFilterOpen && (
            <div className="mt-3 bg-slate-900 text-slate-200 rounded-3xl p-5 flex flex-col xl:flex-row gap-4 items-center justify-between shadow-xl border border-slate-800 transition-all">
              <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
                {/* Magnifying glass circle icon container */}
                <div className="bg-slate-800 p-2.5 rounded-2xl border border-slate-700">
                  <Search className="w-5 h-5 text-emerald-400" />
                </div>

                {/* Metin Arama Input */}
                <div className="relative flex-1 sm:flex-initial">
                  <input
                    type="text"
                    placeholder="Teçhizat veya seri no..."
                    value={techizatSearchQuery}
                    onChange={(e) => {
                      setTechizatSearchQuery(e.target.value);
                      setActiveMatchIdx(0);
                    }}
                    className="bg-slate-800 text-white font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-48 border border-slate-700 placeholder-slate-400"
                  />
                </div>

                {/* Firma Filtreleme Dropdown */}
                <div className="relative flex-1 sm:flex-initial">
                  <select
                    value={techizatFirmaFilter}
                    onChange={(e) => {
                      setTechizatFirmaFilter(e.target.value);
                      setActiveMatchIdx(0);
                    }}
                    className="bg-slate-800 text-emerald-300 font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-52 border border-slate-700 cursor-pointer shadow-sm"
                  >
                    <option value="">🏢 TÜM FİRMALAR ({uniqueFirmalar.length})</option>
                    {uniqueFirmalar.map((f, i) => (
                      <option key={i} value={f}>{f}</option>
                    ))}
                  </select>
                </div>

                {/* Durum Filtreleme Dropdown */}
                <div className="relative flex-1 sm:flex-initial">
                  <select
                    value={techizatDurumFilter}
                    onChange={(e) => {
                      setTechizatDurumFilter(e.target.value);
                      setActiveMatchIdx(0);
                    }}
                    className="bg-slate-800 text-amber-300 font-black text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-48 border border-slate-700 cursor-pointer shadow-sm"
                  >
                    <option value="">⚡ TÜM DURUMLAR ({uniqueDurumlar.length})</option>
                    {uniqueDurumlar.length > 0 ? (
                      uniqueDurumlar.map((d, i) => (
                        <option key={i} value={d}>{d}</option>
                      ))
                    ) : (
                      <>
                        <option value="FAAL">FAAL</option>
                        <option value="GAYRİ FAAL">GAYRİ FAAL</option>
                      </>
                    )}
                  </select>
                </div>

                {/* Renk Koduna Seç Doğrudan Filtreleme Dropdown */}
                <div className="relative flex-1 sm:flex-initial">
                  <select
                    value={techizatColorFilter}
                    onChange={(e) => {
                      setTechizatColorFilter(e.target.value);
                      setActiveMatchIdx(0);
                    }}
                    className="bg-slate-800 text-cyan-300 font-black text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-56 border border-slate-700 cursor-pointer shadow-sm"
                    title="Renk Koduna Göre Seç / Filtrele"
                  >
                    <option value="ALL">🎨 RENK KODUNA SEÇ (TÜMÜ)</option>
                    <option value="RED">🔴 KIRMIZI (SÜRESİ DOLAN / ACİL)</option>
                    <option value="ORANGE">🟠 TURUNCU (&lt; 90 GÜN KALANLAR)</option>
                    <option value="GREEN">🟢 YEŞİL (SÜRESİ UYGUN / FAAL)</option>
                    <option value="GRAY">⚪ GRİ (BAKIMA TABİ DEĞİL / MUAF)</option>
                  </select>
                </div>

                {(techizatSearchQuery || techizatFirmaFilter || techizatDurumFilter || techizatColorFilter !== "ALL") && (
                  <button
                    onClick={() => {
                      setTechizatSearchQuery("");
                      setTechizatFirmaFilter("");
                      setTechizatDurumFilter("");
                      setTechizatColorFilter("ALL");
                      setActiveMatchIdx(0);
                    }}
                    className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition-all border border-slate-700 cursor-pointer"
                    title="Filtreleri Temizle"
                  >
                    ✕ Temizle
                  </button>
                )}

                {/* YUKARDAKİ FİLTER KISMINA RENK KODU SEÇ SEÇENEĞİ */}
                <div className="flex items-center gap-1.5 bg-slate-900 border border-slate-700/80 rounded-2xl px-3 py-2 shrink-0 shadow-md">
                  <span className="text-[11px] font-black font-mono text-slate-300 uppercase flex items-center gap-1">
                    <span>🎨</span>
                    <span className="hidden sm:inline">RENK KODU:</span>
                  </span>
                  <select
                    value={selectedColorFilter}
                    onChange={(e) => setSelectedColorFilter(e.target.value)}
                    className="bg-slate-800 text-white font-bold text-xs rounded-xl px-2.5 py-1.5 border border-slate-600 focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                  >
                    <option value="all">TÜMÜ (Renk Filtresi Yok)</option>
                    <option value="red">🔴 KIRMIZI (Günü Geçmiş / Süresi Dolan)</option>
                    <option value="orange">🟠 TURUNCU (90 Gün Altı / Yaklaşan)</option>
                    <option value="green">🟢 YEŞİL (Faal / Güvenli)</option>
                    <option value="neutral">⚪ NÖTR (Muaf / Süresiz)</option>
                  </select>
                  {selectedColorFilter !== 'all' && (
                    <button
                      onClick={() => setSelectedColorFilter('all')}
                      className="text-slate-400 hover:text-white text-xs px-1.5 py-0.5 rounded-lg bg-slate-800 hover:bg-slate-700 font-bold transition-all cursor-pointer"
                      title="Renk filtresini temizle"
                    >
                      ✕
                    </button>
                  )}
                </div>

                {/* Sort by color button */}
                <button
                  onClick={() => setSortByColor(!sortByColor)}
                  className={`px-4 py-3 active:scale-95 font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border shrink-0 ${
                    sortByColor
                      ? 'bg-red-600 hover:bg-red-700 text-white border-red-500 animate-pulse'
                      : 'bg-slate-800 hover:bg-slate-750 text-slate-200 border-slate-700'
                  }`}
                  title="Bakım gün sayısına göre (Kırmızı ➜ Turuncu ➜ Sarı ➜ Yeşil) sıralar"
                >
                  <SlidersHorizontal className="w-4 h-4 text-amber-400" />
                  <span>
                    {sortByColor ? "🔴 RENK SIRALAMASI AKTİF" : "⏳ RENK KODUNA GÖRE SIRALA"}
                  </span>
                </button>
              </div>

              {/* Action Buttons Right Side */}
              <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto justify-end">
                <button
                  onClick={pullFromDrive}
                  disabled={isLoadingDrive}
                  className="px-4 py-3 bg-[#0b3d1d] hover:bg-[#072612] disabled:opacity-50 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-emerald-600 shrink-0"
                  title="Google Drive üzerindeki Excel dosyasından anında canlı verileri tazeler"
                >
                  <RefreshCw className={`w-4 h-4 text-emerald-300 ${isLoadingDrive ? 'animate-spin' : ''}`} />
                  <span>{isLoadingDrive ? 'DRİVE BAĞLANIYOR...' : 'DRİVE İLE SENKRONİZE ET'}</span>
                </button>

                <button
                  onClick={() => requirePassword(pullFromDrive)}
                  className="px-4 py-3 bg-[#0b3d1d] hover:bg-[#072612] active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-emerald-600 shrink-0"
                >
                  <RefreshCw className="w-4 h-4 text-emerald-300" />
                  <span>VERİ GÜNCELLE</span>
                </button>

                <button
                  onClick={() => requirePassword(() => {
                    setNewRowData({});
                    setIsNewProductOpen(true);
                  })}
                  className="px-4 py-3 bg-emerald-600 hover:bg-emerald-500 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-950/20 border border-emerald-500 shrink-0"
                  title="Yetkili şifresi ile yeni ürün veya teçhizat kaydı ekle"
                >
                  <PlusCircle className="w-4 h-4 text-white" />
                  <span>YENİ ÜRÜN EKLE</span>
                </button>

                <button
                  onClick={exportToExcel}
                  className="px-4 py-3 bg-emerald-700 hover:bg-emerald-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-900/10 border border-emerald-600 shrink-0"
                >
                  <Download className="w-4 h-4" />
                  <span>EXCEL İNDİR</span>
                </button>

                <button
                  onClick={() => downloadStandaloneSystemZip('yer-destek')}
                  className="px-4 py-3 bg-indigo-700 hover:bg-indigo-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-indigo-500 shrink-0"
                  title="Bu sistemi bağımsız zip web sitesi olarak indir"
                >
                  <Package className="w-4 h-4 text-indigo-200" />
                  <span>SİTEYİ ZİP İNDİR</span>
                </button>
              </div>
            </div>
          )}
        </div>

        {/* DYNAMIC INTERACTIVE GRID CONTAINER (EXACT MATCH OF image.png) */}
        <div className="flex-1 bg-white border-2 border-slate-200/60 rounded-[2.5rem] shadow-xl overflow-hidden flex flex-col print:border-none print:shadow-none min-h-[500px]">
          {/* Table Title Bar */}
          <div className="bg-slate-900 px-6 py-4 border-b border-slate-800 flex items-center justify-between print:hidden shrink-0 flex-wrap gap-3">
            {Object.keys(selectedItems).filter(k => selectedItems[k]).length > 0 ? (
              <div className="flex items-center gap-3 animate-fade-in flex-wrap">
                <span className="text-xs font-black text-amber-400 bg-amber-950/40 border border-amber-900 px-3 py-1.5 rounded-xl">
                  ⚡ {Object.keys(selectedItems).filter(k => selectedItems[k]).length} TEÇHİZAT SEÇİLDİ
                </span>
                <button
                  onClick={handleDeleteSelected}
                  className="px-4 py-1.5 bg-rose-700 hover:bg-rose-600 active:scale-95 text-white font-black text-xs uppercase tracking-wider rounded-xl transition-all cursor-pointer flex items-center gap-2 shadow-lg border border-rose-600"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  <span>SEÇİLENLERİ SİL</span>
                </button>
              </div>
            ) : (
              <div className="flex items-center gap-2 text-[11px] font-mono font-black select-none">
                <span className="flex items-center gap-1.5 bg-emerald-950/60 text-emerald-300 border border-emerald-800/80 px-2.5 py-1 rounded-lg">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block shadow-sm"></span>
                  🟢 &ge;90 Gün
                </span>
                <span className="flex items-center gap-1.5 bg-orange-950/60 text-orange-300 border border-orange-800/80 px-2.5 py-1 rounded-lg">
                  <span className="w-2.5 h-2.5 rounded-full bg-orange-500 inline-block shadow-sm"></span>
                  🟠 &lt;90 Gün
                </span>
                <span className="flex items-center gap-1.5 bg-rose-950/60 text-rose-300 border border-rose-800/80 px-2.5 py-1 rounded-lg">
                  <span className="w-2.5 h-2.5 rounded-full bg-rose-600 inline-block shadow-sm animate-pulse"></span>
                  🔴 Bakım/Ömür Geçmiş
                </span>
              </div>
            )}

            <span className="text-[10px] font-mono font-black text-slate-400 bg-slate-800 px-3 py-1 rounded-full border border-slate-700">
              {processedRows.length} KALEM TEÇHİZAT LİSTELENDİ
            </span>
          </div>

          {/* Table Grid Scroll Wrapper */}
          <div className="flex-1 overflow-auto max-h-[75vh] relative">
            <table className="w-full border-collapse text-left min-w-[1200px]">
              <thead>
                <tr className="bg-slate-900 border-b border-slate-800 shrink-0 sticky top-0 z-10 select-none">
                  <th className="px-3 py-3.5 text-center text-[10px] font-black text-slate-300 uppercase font-mono border-r border-slate-800 w-[50px]">
                    <input
                      type="checkbox"
                      checked={processedRows.length > 0 && processedRows.every(({ idx }) => selectedItems[idx])}
                      onChange={(e) => {
                        const checked = e.target.checked;
                        const newSel: Record<string, boolean> = {};
                        if (checked) {
                          processedRows.forEach(({ idx }) => { newSel[idx] = true; });
                        }
                        setSelectedItems(newSel);
                      }}
                      className="rounded accent-emerald-500 w-4 h-4 cursor-pointer"
                    />
                  </th>
                  {TABLE_COLUMNS.map((col, idx) => (
                    <th
                      key={idx}
                      className="px-4 py-3.5 text-[11px] font-black text-slate-200 uppercase tracking-wider font-mono border-r border-slate-800 last:border-r-0"
                    >
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs text-slate-800 font-medium">
                {processedRows.length === 0 ? (
                  <tr>
                    <td colSpan={TABLE_COLUMNS.length + 1} className="py-16 text-center text-slate-400">
                      <div className="flex flex-col items-center justify-center gap-3">
                        <Wrench className="w-8 h-8 text-slate-300" />
                        <span className="font-bold">Eşleşen teçhizat veya kayıt bulunamadı.</span>
                      </div>
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map(({ row, idx }) => {
                    const colorStatus = getRowColorStatus(row);
                    const isSelected = Boolean(selectedItems[idx]);

                    let dateColorClass = 'text-slate-700';
                    if (colorStatus === 'red') dateColorClass = 'text-rose-600 font-black';
                    else if (colorStatus === 'orange') dateColorClass = 'text-orange-600 font-black';
                    else if (colorStatus === 'green') dateColorClass = 'text-emerald-700 font-bold';
                    else if (colorStatus === 'neutral') dateColorClass = 'text-slate-400 italic';

                    return (
                      <tr
                        key={idx}
                        onClick={() => setActiveEditRow({ idx, row: [...row] })}
                        className={`transition-colors cursor-pointer ${
                          isSelected ? 'bg-emerald-50/70 hover:bg-emerald-100/70' : 'bg-white hover:bg-slate-50'
                        }`}
                      >
                        <td
                          className="px-3 py-3 text-center border-r border-slate-100"
                          onClick={(e) => e.stopPropagation()}
                        >
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={(e) => {
                              setSelectedItems(prev => ({ ...prev, [idx]: e.target.checked }));
                            }}
                            className="rounded accent-emerald-500 w-4 h-4 cursor-pointer"
                          />
                        </td>
                        {/* SIRA NO */}
                        <td className="px-4 py-3 text-center font-bold text-slate-500 border-r border-slate-100 w-16">
                          {idx + 1}
                        </td>
                        {/* TEÇHİZAT ADI */}
                        <td className="px-4 py-3 font-bold text-slate-900 border-r border-slate-100 max-w-xs">
                          {row[1] || '-'}
                        </td>
                        {/* PARÇA NO */}
                        <td className="px-4 py-3 font-mono font-bold text-slate-700 border-r border-slate-100">
                          {row[2] || '-'}
                        </td>
                        {/* SERİ NO */}
                        <td className="px-4 py-3 font-mono text-slate-600 border-r border-slate-100">
                          {row[3] || '-'}
                        </td>
                        {/* MİKTAR */}
                        <td className="px-4 py-3 text-center font-bold text-slate-800 border-r border-slate-100">
                          {row[4] || '1'}
                        </td>
                        {/* BULUNDUĞU YER */}
                        <td className="px-4 py-3 text-slate-700 border-r border-slate-100">
                          {row[5] || '-'}
                        </td>
                        {/* DURUMU */}
                        <td className="px-4 py-3 text-center border-r border-slate-100">
                          <span className={`px-2.5 py-1 rounded-lg text-[10px] font-black ${
                            (row[6] || '').trim().toUpperCase() === 'FAAL'
                              ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                              : 'bg-rose-100 text-rose-800 border border-rose-300'
                          }`}>
                            {row[6] || 'FAAL'}
                          </span>
                        </td>
                        {/* KALİBRASYONA TABİ */}
                        <td className="px-4 py-3 text-center font-bold text-slate-700 border-r border-slate-100">
                          {row[7] || 'EVET'}
                        </td>
                        {/* SON KONTROL */}
                        <td className="px-4 py-3 text-center font-mono text-slate-600 border-r border-slate-100">
                          {row[8] || '-'}
                        </td>
                        {/* GELECEK KONTROL */}
                        <td className={`px-4 py-3 text-center font-mono border-r border-slate-100 ${dateColorClass}`}>
                          {row[9] || '-'}
                        </td>
                        {/* FİRMA */}
                        <td className="px-4 py-3 text-slate-700 border-r border-slate-100">
                          {row[10] || '-'}
                        </td>
                        {/* AÇIKLAMA */}
                        <td className="px-4 py-3 text-slate-500 max-w-sm truncate">
                          {row[11] || '-'}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* SAYFALAMA (PAGINATION) KONTROLLERİ */}
          <div className="bg-slate-900 px-5 py-3 border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-300 select-none">
            <div className="flex items-center gap-3">
              <span className="text-slate-400 font-mono text-[11px]">Sayfa Başına Teçhizat:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="bg-slate-800 border border-slate-700 text-white rounded-lg px-2.5 py-1 font-mono font-bold text-xs cursor-pointer focus:ring-1 focus:ring-emerald-500 outline-none"
              >
                <option value={50}>50 Ürün / Sayfa (Önerilen)</option>
                <option value={100}>100 Ürün</option>
                <option value={200}>200 Ürün</option>
                <option value={99999}>Tümü ({processedRows.length})</option>
              </select>

              <span className="text-slate-400 font-mono text-[11px] ml-2">
                {processedRows.length === 0 ? '0' : `${(currentPage - 1) * pageSize + 1} - ${Math.min(currentPage * pageSize, processedRows.length)}`} / Toplam {processedRows.length} ürün
              </span>
            </div>

            <div className="flex items-center gap-1.5 font-mono">
              <button
                disabled={currentPage <= 1}
                onClick={() => setCurrentPage(1)}
                className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed text-slate-200 border border-slate-700 rounded-lg font-bold text-xs transition"
                title="İlk Sayfa"
              >
                « İlk
              </button>
              <button
                disabled={currentPage <= 1}
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed text-slate-200 border border-slate-700 rounded-lg font-bold text-xs transition"
              >
                ‹ Önceki
              </button>

              {/* Numbered Page Buttons: 1, 2, 3... */}
              <div className="flex items-center gap-1">
                {Array.from({ length: totalPages }, (_, i) => i + 1)
                  .filter(p => p === 1 || p === totalPages || Math.abs(p - currentPage) <= 2)
                  .map((p, idx, arr) => {
                    const prevP = arr[idx - 1];
                    const showEllipsis = prevP && p - prevP > 1;
                    return (
                      <React.Fragment key={p}>
                        {showEllipsis && <span className="px-1 text-slate-500">...</span>}
                        <button
                          type="button"
                          onClick={() => setCurrentPage(p)}
                          className={`w-7 h-7 rounded-lg text-xs font-bold transition cursor-pointer flex items-center justify-center ${
                            currentPage === p
                              ? 'bg-emerald-600 text-white shadow-sm'
                              : 'bg-slate-800 text-slate-300 hover:bg-slate-700 border border-slate-700'
                          }`}
                        >
                          {p}
                        </button>
                      </React.Fragment>
                    );
                  })}
              </div>

              <button
                disabled={currentPage >= totalPages}
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed text-slate-200 border border-slate-700 rounded-lg font-bold text-xs transition"
              >
                Sonraki ›
              </button>
              <button
                disabled={currentPage >= totalPages}
                onClick={() => setCurrentPage(totalPages)}
                className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed text-slate-200 border border-slate-700 rounded-lg font-bold text-xs transition"
                title="Son Sayfa"
              >
                Son »
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* PASSWORD MODAL */}
      {isPasswordModalOpen && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border-2 border-emerald-500/60 p-6 rounded-3xl shadow-2xl max-w-sm w-full">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2 text-white font-black text-sm uppercase">
                <Lock className="w-4 h-4 text-emerald-400" />
                <span>YETKİLİ ŞİFRESİ GİRİNİZ</span>
              </div>
              <button
                onClick={() => setIsPasswordModalOpen(false)}
                className="text-slate-400 hover:text-white"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handlePasswordSubmit} className="space-y-4">
              <input
                type="password"
                autoFocus
                placeholder="Şifre"
                value={passwordInput}
                onChange={(e) => setPasswordInput(e.target.value)}
                className="w-full bg-slate-800 text-white font-mono text-center tracking-widest text-lg p-3 rounded-2xl border border-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500"
              />
              {passwordError && (
                <p className="text-xs text-rose-400 font-bold text-center">Hatalı yetkili şifresi!</p>
              )}
              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setIsPasswordModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-bold"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black"
                >
                  Onayla
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ROW EDIT MODAL */}
      {activeEditRow && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border-2 border-slate-700 p-6 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between mb-4 border-b border-slate-800 pb-3">
              <h3 className="text-sm font-black text-emerald-400 uppercase">
                Teçhizat Kaydı Düzenle (Sıra {activeEditRow.idx + 1})
              </h3>
              <button
                onClick={() => setActiveEditRow(null)}
                className="text-slate-400 hover:text-white"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {TABLE_COLUMNS.slice(1).map((colName, cIdx) => {
                const targetIdx = cIdx + 1;
                return (
                  <div key={targetIdx} className="space-y-1">
                    <label className="text-[10px] font-mono font-bold text-slate-400 uppercase">
                      {colName}
                    </label>
                    <input
                      type="text"
                      value={activeEditRow.row[targetIdx] || ''}
                      onChange={(e) => {
                        const val = e.target.value;
                        setActiveEditRow(prev => {
                          if (!prev) return null;
                          const newR = [...prev.row];
                          newR[targetIdx] = val;
                          return { ...prev, row: newR };
                        });
                      }}
                      className="w-full bg-slate-800 text-white text-xs p-2.5 rounded-xl border border-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                    />
                  </div>
                );
              })}
            </div>
            <div className="flex items-center justify-end gap-2 pt-6 mt-4 border-t border-slate-800">
              <button
                onClick={() => setActiveEditRow(null)}
                className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-bold"
              >
                Vazgeç
              </button>
              <button
                onClick={handleSaveEditRow}
                className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black"
              >
                Kaydet
              </button>
            </div>
          </div>
        </div>
      )}

      {/* NEW PRODUCT MODAL */}
      {isNewProductOpen && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border-2 border-emerald-500/60 p-6 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between mb-4 border-b border-slate-800 pb-3">
              <h3 className="text-sm font-black text-emerald-400 uppercase">
                Yeni Teçhizat / Alet Ekle ({currentUnit.shortLabel})
              </h3>
              <button
                onClick={() => setIsNewProductOpen(false)}
                className="text-slate-400 hover:text-white"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {TABLE_COLUMNS.slice(1).map((colName, cIdx) => {
                const targetIdx = String(cIdx + 1);
                return (
                  <div key={targetIdx} className="space-y-1">
                    <label className="text-[10px] font-mono font-bold text-slate-400 uppercase">
                      {colName}
                    </label>
                    <input
                      type="text"
                      value={newRowData[targetIdx] || ''}
                      onChange={(e) => {
                        const val = e.target.value;
                        setNewRowData(prev => ({ ...prev, [targetIdx]: val }));
                      }}
                      className="w-full bg-slate-800 text-white text-xs p-2.5 rounded-xl border border-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                    />
                  </div>
                );
              })}
            </div>
            <div className="flex items-center justify-end gap-2 pt-6 mt-4 border-t border-slate-800">
              <button
                onClick={() => setIsNewProductOpen(false)}
                className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-bold"
              >
                Vazgeç
              </button>
              <button
                onClick={handleSaveNewProduct}
                className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black"
              >
                Ekle ve Kaydet
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default YerDestekApp;
