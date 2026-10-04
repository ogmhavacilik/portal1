import React, { useState, useEffect, useMemo } from 'react';
import {
  Boxes,
  Search,
  SlidersHorizontal,
  RefreshCw,
  PlusCircle,
  Download,
  Trash2,
  ArrowLeft,
  CheckCircle2,
  FileSpreadsheet,
  X,
  Lock,
  Package,
  Layers,
  FlaskConical,
  ArrowUpRight,
  ArrowDownLeft,
  Upload,
  Calendar,
  MapPin
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { cleanAndFormatDateString } from '../utils/driveExcelSync';
import { publishCrossSystemEvent, subscribeCrossSystemEvents } from '../utils/crossSystemBridge';
import { downloadStandaloneSystemZip } from '../utils/zipExporter';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export const DEPO_CATEGORIES = [
  {
    key: 'sarf',
    name: 'Sarf Malzeme ve Parça Deposu',
    shortLabel: 'SARF DEPO',
    driveFileName: 'at-802_sarf_ve_parca_deposu.xlsx',
    fileId: '1fdSnAWA55Gj0DfTQDqSJSXhaIKpKNv5A',
    storageKey: 'excel_depo_sarf_data'
  },
  {
    key: 'kimyasal',
    name: 'Kimyasal ve Boya/Yağ Deposu',
    shortLabel: 'KİMYASAL DEPO',
    driveFileName: 'at-802_kimyasal_depo.xlsx',
    fileId: '',
    storageKey: 'excel_depo_kimyasal_data'
  }
];

export const DEPO_TABLE_COLUMNS = [
  "SIRA NO",
  "MALZEME ADI",
  "PARÇA NO (P/N) / MODEL",
  "SERİ NO (S/N) / LOT NO",
  "MİKTAR",
  "BİRİM / ÖLÇÜ",
  "DEPO YERİ / RAF",
  "DURUMU",
  "RAF ÖMRÜNE TABİ",
  "RAF ÖMRÜ GİRİŞ TARİHİ",
  "GİRİŞ / İMAL TARİHİ",
  "SKT / RAF ÖMRÜ BİTİŞ",
  "TEDARİKÇİ / FİRMA",
  "AÇIKLAMA",
  "RAF ÖMRÜ VAR MI?",
  "RAF ÖMRÜ BİTİŞ TARİHİ"
];

export function DepoApp({ initialCategory, isEmbedded }: { initialCategory?: string; isEmbedded?: boolean }) {
  const params = new URLSearchParams(window.location.search);
  const catFromUrl = params.get('cat') || params.get('unit') || initialCategory || 'sarf';
  const searchFromUrl = params.get('search') || '';
  const embeddedMode = isEmbedded ?? (params.get('embedded') === 'true');

  const [selectedCategoryKey, setSelectedCategoryKey] = useState<string>(catFromUrl);
  const [searchQuery, setSearchQuery] = useState<string>(searchFromUrl);
  const [firmaFilter, setFirmaFilter] = useState<string>('');
  const [durumFilter, setDurumFilter] = useState<string>('');
  const [colorFilter, setColorFilter] = useState<string>('ALL');
  const [selectedColorFilter, setSelectedColorFilter] = useState<string>('all');
  const [sortByColor, setSortByColor] = useState<boolean>(false);
  const [selectedItems, setSelectedItems] = useState<Record<string, boolean>>({});
  const [isLoadingDrive, setIsLoadingDrive] = useState<boolean>(false);
  const [notification, setNotification] = useState<string | null>(null);

  // Sayfalama (Virtualization / Chunking) - 1.779 ürünün aynı anda DOM'a yüklenmesinden kaynaklanan donmayı önler
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(100);

  // Modals
  const [activeEditRow, setActiveEditRow] = useState<{ idx: number; row: string[] } | null>(null);
  const [isNewProductOpen, setIsNewProductOpen] = useState<boolean>(false);
  const [newRowData, setNewRowData] = useState<Record<string, string>>({});
  const [isTxModalOpen, setIsTxModalOpen] = useState<boolean>(false);
  const [txType, setTxType] = useState<'in' | 'out'>('out');
  const [txData, setTxData] = useState({ qty: 1, target: 'HANGAR', person: '', notes: '' });

  // Password Lock
  const [isPasswordModalOpen, setIsPasswordModalOpen] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [passwordError, setPasswordError] = useState<boolean>(false);
  const [pendingPasswordAction, setPendingPasswordAction] = useState<(() => void) | null>(null);

  // Mobile View Specific States
  const [isMobile, setIsMobile] = useState<boolean>(false);
  const [mobileSearchQuery, setMobileSearchQuery] = useState<string>('');
  const [selectedMobileItem, setSelectedMobileItem] = useState<{ row: string[]; idx: number } | null>(null);
  const [showMobileSuggestions, setShowMobileSuggestions] = useState<boolean>(false);

  useEffect(() => {
    const checkMobile = () => {
      // Check both width and user agent to be more reliable
      const isTouch = 'ontouchstart' in window || navigator.maxTouchPoints > 0;
      const isSmallScreen = window.innerWidth < 1024; // Increased threshold for better mobile detection
      setIsMobile(isSmallScreen && isTouch);
    };
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);



  const [depoData, setDepoData] = useState<Record<string, string[][]>>(() => {
    const initial: Record<string, string[][]> = {};
    for (const c of DEPO_CATEGORIES) {
      try {
        const stored = localStorage.getItem(c.storageKey);
        if (stored) {
          initial[c.key] = JSON.parse(stored);
        } else {
          // Default mock/fallback data if empty
          initial[c.key] = [
            ['1', 'O-RING SIZDIRMAZLIK HALKASI', 'MS29513-010', 'LOT-9821', '50', 'ADET', 'RAF-A1/K2', 'FAAL', 'EVET', '01.01.2025', '01.01.2027', 'PARKER HANNIFIN', 'Standart hidrolik o-ring'],
            ['2', 'HİDROLİK SIVI MOBIL AERO HFA', 'MIL-PRF-5606', 'CAN-4412', '12', 'KUTU', 'KİMYA-DOLAP-2', 'FAAL', 'EVET', '15.03.2024', '15.03.2026', 'EXXON MOBIL', 'Kırmızı mineral hidrolik yağ'],
            ['3', 'SAFETY WIRE (KİLİTLEME TELİ)', 'MS20995C32', 'N/A', '8', 'MAKARA', 'RAF-B3/K1', 'FAAL', 'HAYIR', '10.02.2024', '-', 'TW METALS', 'Paslanmaz tel 0.032 inç'],
            ['4', 'SEALANT PR-1422 B-2', 'PR-1422-B2', 'LOT-2024-B', '6', 'TÜP', 'KİMYA-BUZDOLABI', 'FAAL', 'EVET', '10.05.2024', '10.11.2024', 'PPG AEROSPACE', 'Yakıt tankı dolgu sızdırmazlık']
          ];
        }
      } catch {
        initial[c.key] = [];
      }
    }
    return initial;
  });

  const showNotification = (msg: string) => {
    setNotification(msg);
    setTimeout(() => setNotification(null), 4000);
  };

  const currentCategory = DEPO_CATEGORIES.find(c => c.key === selectedCategoryKey) || DEPO_CATEGORIES[0];
  const currentRows = depoData[selectedCategoryKey] || [];

  const mobileSuggestions = useMemo(() => {
    if (mobileSearchQuery.trim().length < 2) return [];
    const q = mobileSearchQuery.toLowerCase().trim();
    return currentRows
      .map((row, idx) => ({ row, idx }))
      .filter(({ row }) => {
        const name = (row[1] || '').toLowerCase();
        const pn = (row[2] || '').toLowerCase();
        const sn = (row[3] || '').toLowerCase();
        const loc = (row[6] || '').toLowerCase();
        return name.includes(q) || pn.includes(q) || sn.includes(q) || loc.includes(q);
      })
      .slice(0, 15);
  }, [currentRows, mobileSearchQuery]);

  const saveRows = (catKey: string, rows: string[][]) => {
    setDepoData(prev => ({ ...prev, [catKey]: rows }));
    const cat = DEPO_CATEGORIES.find(c => c.key === catKey);
    if (cat) {
      try {
        localStorage.setItem(cat.storageKey, JSON.stringify(rows));
      } catch (e) {}

      // Kalıcı olarak Excel'e ve Google Drive'a kaydet (Canlı E-Tablo / Drive Entegrasyonu)
      try {
        const wb = XLSX.utils.book_new();
        // Ensure all rows match the column schema
        const processedRowsForUpload = rows.map(r => {
          const newR = [...r];
          // Pad rows to match DEPO_TABLE_COLUMNS length
          while (newR.length < DEPO_TABLE_COLUMNS.length) {
            newR.push('');
          }
          return newR;
        });
        const wsData = [
          DEPO_TABLE_COLUMNS,
          ...processedRowsForUpload
        ];
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, cat.shortLabel);
        const base64Data = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });

        fetch('/api/upload-techizat-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: cat.driveFileName,
            targetKey: `depo_${catKey}`,
            base64Data: base64Data,
            folderId: DRIVE_FOLDER_ID
          })
        }).catch(err => console.warn('Drive save warn:', err));
      } catch (excelErr) {
        console.warn('Excel export error:', excelErr);
      }
    }
    publishCrossSystemEvent('DEPO_ITEM_TRANSFERRED', 'depo', { catKey, count: rows.length });
  };

  useEffect(() => {
    const unsub = subscribeCrossSystemEvents((ev) => {
      if (ev.type === 'REFRESH_ALL') {
        const updated: Record<string, string[][]> = {};
        for (const c of DEPO_CATEGORIES) {
          try {
            const val = localStorage.getItem(c.storageKey);
            if (val) updated[c.key] = JSON.parse(val);
          } catch {}
        }
        setDepoData(prev => ({ ...prev, ...updated }));
      }
    });
    return unsub;
  }, []);

  // Kategori değiştiğinde veya sayfa yüklendiğinde Drive'dan verileri senkronize etme (Kullanıcı talebiyle kaldırıldı)
  // Sadece ilk açılışta verileri çek
  useEffect(() => {
    pullFromDrive();
  }, []);

  const uniqueFirmalar = useMemo(() => {
    const s = new Set<string>();
    for (const r of currentRows) {
      const f = (r[11] || '').trim();
      if (f && f !== '-' && f !== '--') s.add(f);
    }
    return Array.from(s).sort();
  }, [currentRows]);

  const uniqueDurumlar = useMemo(() => {
    const s = new Set<string>();
    for (const r of currentRows) {
      const d = (r[7] || '').trim().toUpperCase();
      if (d && d !== '-' && d !== '--') s.add(d);
    }
    return Array.from(s).sort();
  }, [currentRows]);

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
    const isRafTabi = (row[8] || '').trim().toUpperCase();
    if (isRafTabi === 'HAYIR' || isRafTabi === 'MUAFIYET') return 'neutral';
    const days = getDaysUntil(row[10]);
    if (days === null) return 'neutral';
    if (days < 0) return 'red';
    if (days <= 90) return 'orange';
    return 'green';
  };

  const processedRows = useMemo(() => {
    let result = currentRows.map((row, idx) => ({ row, idx }));

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter(({ row }) =>
        row.some(cell => String(cell || '').toLowerCase().includes(q))
      );
    }

    if (firmaFilter) {
      result = result.filter(({ row }) => (row[11] || '').trim() === firmaFilter);
    }

    if (durumFilter) {
      result = result.filter(({ row }) => (row[7] || '').trim().toUpperCase() === durumFilter);
    }

    if (colorFilter !== 'ALL') {
      result = result.filter(({ row }) => {
        const color = getRowColorStatus(row);
        if (colorFilter === 'RED') return color === 'red';
        if (colorFilter === 'ORANGE') return color === 'orange';
        if (colorFilter === 'GREEN') return color === 'green';
        if (colorFilter === 'GRAY') return color === 'neutral';
        return true;
      });
    }

    if (selectedColorFilter !== 'all') {
      result = result.filter(({ row }) => {
        const color = getRowColorStatus(row);
        return color === selectedColorFilter;
      });
    }

    if (sortByColor) {
      const rank = { red: 0, orange: 1, green: 2, neutral: 3 };
      result.sort((a, b) => {
        const cA = getRowColorStatus(a.row);
        const cB = getRowColorStatus(b.row);
        if (rank[cA] !== rank[cB]) return rank[cA] - rank[cB];
        const dA = getDaysUntil(a.row[10]) ?? 9999;
        const dB = getDaysUntil(b.row[10]) ?? 9999;
        return dA - dB;
      });
    }

    return result;
  }, [currentRows, searchQuery, firmaFilter, durumFilter, colorFilter, selectedColorFilter, sortByColor]);

  // Sayfalama hesaplamaları
  const totalPages = useMemo(() => {
    return Math.max(1, Math.ceil(processedRows.length / pageSize));
  }, [processedRows.length, pageSize]);

  // Filtreler veya arama değiştiğinde sayfa 1'e dönsün
  useEffect(() => {
    setCurrentPage(1);
  }, [searchQuery, firmaFilter, durumFilter, colorFilter, selectedColorFilter, selectedCategoryKey, pageSize]);

  const paginatedRows = useMemo(() => {
    if (pageSize >= 99999) return processedRows;
    const startIndex = (currentPage - 1) * pageSize;
    return processedRows.slice(startIndex, startIndex + pageSize);
  }, [processedRows, currentPage, pageSize]);

  const pullFromDrive = async () => {
    setIsLoadingDrive(true);
    showNotification(`Google Drive'dan ${currentCategory.name} verileri canlı olarak alınıyor...`);
    try {
      const res = await fetch('/api/read-excel-from-drive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: currentCategory.driveFileName,
          fileId: (currentCategory as any).fileId || '',
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
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        if (rawRows && rawRows.length > 1) {
          let headerIdx = 0;
          for (let r = 0; r < Math.min(15, rawRows.length); r++) {
            const rowStr = (rawRows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
            if (rowStr.includes('MALZEME') || rowStr.includes('DESCRIPTION') || rowStr.includes('PARÇA') || rowStr.includes('P/N')) {
              headerIdx = r;
              break;
            }
          }
          const headers = (rawRows[headerIdx] || []).map(c => String(c || '').toUpperCase().trim());
          const descCol = headers.findIndex(c => c.includes('DESCRIPTION') || c.includes('MALZEME'));
          const pnCol = headers.findIndex(c => c.includes('PART') || c.includes('P/N') || c.includes('PN'));
          const snCol = headers.findIndex(c => c.includes('SERİ') || c.includes('SERI') || c.includes('MÜKERRER'));
          const locCol = headers.findIndex(c => c.includes('LOKASYON') || c.includes('RAF') || c.includes('KONUM'));
          const qtyCol = headers.findIndex(c => c === 'GELEN' || c.includes('GELEN') || c.includes('TOPLAM STOK') || c.includes('MİKTAR'));

          const cleanedRows: string[][] = [];
          for (let r = headerIdx + 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row || row.length === 0) continue;
            const desc = String(row[descCol >= 0 ? descCol : 0] || '').trim();
            if (!desc || desc.toUpperCase() === 'DESCRIPTION' || desc.toUpperCase().includes('MALZEME ADI')) continue;

            const rowArr = new Array(DEPO_TABLE_COLUMNS.length).fill('');
            rowArr[0] = String(cleanedRows.length + 1);
            rowArr[1] = desc;
            rowArr[2] = String(row[pnCol >= 0 ? pnCol : 1] || '-').trim();
            rowArr[3] = String(row[snCol >= 0 ? snCol : 2] || '-').trim();
            rowArr[4] = String(row[qtyCol >= 0 ? qtyCol : 4] || '1').trim();
            rowArr[5] = 'ADET';
            rowArr[6] = String(row[locCol >= 0 ? locCol : 3] || 'DEPO').trim();
            rowArr[7] = 'FAAL';
            rowArr[8] = 'EVET';
            rowArr[9] = '-'; // RAF ÖMRÜ GİRİŞ TARİHİ
            rowArr[10] = '-'; // GİRİŞ / İMAL TARİHİ
            rowArr[11] = '-'; // SKT / RAF ÖMRÜ BİTİŞ
            rowArr[12] = '-'; // TEDARİKÇİ / FİRMA
            rowArr[13] = ''; // AÇIKLAMA

            cleanedRows.push(rowArr);
          }

          saveRows(selectedCategoryKey, cleanedRows);
          showNotification(`✅ ${currentCategory.shortLabel} güncellendi: ${cleanedRows.length} kayıt Google Drive'dan aktarıldı.`);
        }
      } else {
        showNotification(`⚠️ Drive üzerinde dosya okundu ancak satır bulunamadı.`);
      }
    } catch (e: any) {
      showNotification(`❌ Drive bağlantı hatası: ${e?.message || 'Bilinmeyen hata'}`);
    } finally {
      setIsLoadingDrive(false);
    }
  };

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();
    const wsData = [
      DEPO_TABLE_COLUMNS,
      ...processedRows.map(({ row }) => DEPO_TABLE_COLUMNS.map((_, cIdx) => row[cIdx] || ''))
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, currentCategory.shortLabel);
    XLSX.writeFile(wb, `${currentCategory.driveFileName.replace('.xlsx', '')}_${new Date().toISOString().slice(0, 10)}.xlsx`);
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

  const handleDeleteSelected = () => {
    const selectedIndices = new Set(Object.keys(selectedItems).filter(k => selectedItems[k]).map(Number));
    if (selectedIndices.size === 0) return;
    requirePassword(() => {
      const remaining = currentRows.filter((_, idx) => !selectedIndices.has(idx)).map((r, i) => {
        r[0] = String(i + 1);
        return r;
      });
      saveRows(selectedCategoryKey, remaining);
      setSelectedItems({});
      showNotification(`✅ Seçilen ${selectedIndices.size} depo kalemi silindi.`);
    });
  };

  const handleSaveNewProduct = () => {
    // Find the indices for the shelf life columns
    const rafOmruVarMiIdx = DEPO_TABLE_COLUMNS.indexOf("RAF ÖMRÜ VAR MI?");
    const rafOmruBitisIdx = DEPO_TABLE_COLUMNS.indexOf("RAF ÖMRÜ BİTİŞ TARİHİ");
    const rafOmruGirisTarihIdx = DEPO_TABLE_COLUMNS.indexOf("RAF ÖMRÜ GİRİŞ TARİHİ");
    const sktBitisIdx = DEPO_TABLE_COLUMNS.indexOf("SKT / RAF ÖMRÜ BİTİŞ");
    const imalTarihIdx = DEPO_TABLE_COLUMNS.indexOf("GİRİŞ / İMAL TARİHİ");

    // Get user selection for "RAF ÖMRÜ VAR MI?"
    const rafOmruVarMiVal = (newRowData[String(rafOmruVarMiIdx)] || '').trim().toUpperCase();

    const newRow = DEPO_TABLE_COLUMNS.map((colName, i) => {
        if (i === 0) return String(currentRows.length + 1);
        
        let val = (newRowData[String(i)] || '').trim();

        // If "HAYIR" is selected, force date fields to "-"
        if (rafOmruVarMiVal === 'HAYIR') {
          if (i === rafOmruBitisIdx ||
              i === rafOmruGirisTarihIdx ||
              i === sktBitisIdx ||
              i === imalTarihIdx) {
            val = '-';
          }
        }
        
        return val;
    });

    const updated = [...currentRows, newRow];
    saveRows(selectedCategoryKey, updated);
    setIsNewProductOpen(false);
    setNewRowData({});
    showNotification('✅ Yeni depo malzemesi ve raf ömrü bilgileri başarıyla eklendi.');
  };

  const handleSaveEditRow = () => {
    if (!activeEditRow) return;
    const updated = [...currentRows];
    updated[activeEditRow.idx] = activeEditRow.row;
    saveRows(selectedCategoryKey, updated);
    setActiveEditRow(null);
    showNotification('✅ Depo malzeme kaydı güncellendi.');
  };

  if (isMobile) {
    return (
      <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-emerald-500 selection:text-white">
        {notification && (
          <div className="fixed top-4 right-4 z-[999] px-4 py-3 rounded-2xl bg-emerald-950 border-2 border-emerald-500 text-emerald-100 text-xs font-black shadow-2xl flex items-center gap-3 animate-fade-in">
            <CheckCircle2 className="w-4 h-4 text-emerald-400" />
            <span>{notification}</span>
          </div>
        )}

        {/* MOBILE TOP HEADER */}
        <header className="bg-slate-900 border-b border-slate-800 p-4 sticky top-0 z-40 flex items-center justify-between shadow-lg select-none">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-amber-600 to-orange-800 flex items-center justify-center shadow-lg border border-amber-400/30">
              <Boxes className="w-4.5 h-4.5 text-white" />
            </div>
            <div>
              <h1 className="text-sm font-black tracking-wider text-white uppercase leading-none">
                OGM HAVACILIK
              </h1>
              <p className="text-[10px] text-slate-400 mt-0.5">Mobil Depo Sorgulama</p>
            </div>
          </div>
          <button
            onClick={() => pullFromDrive()}
            disabled={isLoadingDrive}
            className="p-2 rounded-lg bg-slate-800 border border-slate-700 hover:bg-slate-700 text-slate-300 flex items-center justify-center disabled:opacity-50 cursor-pointer"
            title="Drive'dan Yenile"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${isLoadingDrive ? 'animate-spin text-amber-500' : ''}`} />
          </button>
        </header>

        {/* MOBILE CONTAINER */}
        <div className="flex-1 p-4 flex flex-col gap-5 overflow-y-auto max-w-md mx-auto w-full">
          
          {/* STEP 1: CATEGORY SELECTION */}
          <div className="flex flex-col gap-2">
            <label className="text-xs font-black tracking-wider text-slate-400 uppercase">
              1. İŞLEM YAPILACAK DEPO SEÇİNİZ
            </label>
            <div className="grid grid-cols-2 gap-2.5">
              {DEPO_CATEGORIES.map(c => {
                const isSelected = selectedCategoryKey === c.key;
                const count = (depoData[c.key] || []).length;
                return (
                  <button
                    key={c.key}
                    type="button"
                    onClick={() => {
                      setSelectedCategoryKey(c.key);
                      setSelectedMobileItem(null);
                      setMobileSearchQuery('');
                      setShowMobileSuggestions(false);
                    }}
                    className={`p-3 rounded-2xl flex flex-col gap-1.5 text-left border transition-all cursor-pointer ${
                      isSelected
                        ? 'bg-gradient-to-b from-amber-600 to-amber-700 text-white border-amber-400 shadow-lg shadow-amber-950/40'
                        : 'bg-slate-900 hover:bg-slate-850 text-slate-400 border-slate-800 hover:border-slate-750'
                    }`}
                  >
                    <div className="flex items-center justify-between w-full">
                      {c.key === 'kimyasal' ? (
                        <FlaskConical className={`w-5 h-5 ${isSelected ? 'text-white' : 'text-emerald-500'}`} />
                      ) : (
                        <Layers className={`w-5 h-5 ${isSelected ? 'text-white' : 'text-amber-500'}`} />
                      )}
                      <span className={`px-1.5 py-0.5 rounded-full text-[9px] font-mono font-bold ${
                        isSelected ? 'bg-black/25 text-amber-200' : 'bg-slate-800 text-slate-400'
                      }`}>
                        {count}
                      </span>
                    </div>
                    <div>
                      <span className="text-[11px] font-black tracking-tight block uppercase leading-tight">
                        {c.shortLabel}
                      </span>
                      <span className={`text-[9px] block leading-none mt-0.5 ${
                        isSelected ? 'text-amber-100' : 'text-slate-500'
                      }`}>
                        {c.key === 'kimyasal' ? 'Kimyasal/Boya/Yağ' : 'Sarf/Yedek Parça'}
                      </span>
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          {/* STEP 2: SEARCH PRODUCT */}
          <div className="flex flex-col gap-2 relative">
            <label className="text-xs font-black tracking-wider text-slate-400 uppercase">
              2. MALZEME ARA / SEÇ
            </label>
            <div className="relative flex items-center">
              <div className="absolute left-3.5 text-slate-400">
                <Search className="w-4 h-4" />
              </div>
              <input
                type="text"
                value={mobileSearchQuery}
                placeholder="Malzeme Adı, P/N, S/N veya Raf Yeri..."
                onChange={(e) => {
                  setMobileSearchQuery(e.target.value);
                  setShowMobileSuggestions(true);
                  if (!e.target.value) {
                    setSelectedMobileItem(null);
                  }
                }}
                onFocus={() => setShowMobileSuggestions(true)}
                className="w-full pl-10 pr-10 py-3 rounded-xl bg-slate-900 border-2 border-slate-800 hover:border-slate-750 focus:border-amber-500 text-sm font-medium text-white placeholder-slate-500 transition-colors focus:outline-none"
              />
              {mobileSearchQuery && (
                <button
                  type="button"
                  onClick={() => {
                    setMobileSearchQuery('');
                    setSelectedMobileItem(null);
                    setShowMobileSuggestions(false);
                  }}
                  className="absolute right-3 p-1 rounded-full bg-slate-800 text-slate-400 hover:text-white cursor-pointer"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>

            {/* AUTOCOMPLETE SUGGESTIONS DROPDOWN */}
            {showMobileSuggestions && mobileSuggestions.length > 0 && (
              <div className="absolute top-[72px] left-0 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl max-h-64 overflow-y-auto z-50 divide-y divide-slate-850">
                {mobileSuggestions.map(({ row, idx }) => {
                  const name = row[1] || '';
                  const pn = row[2] || '-';
                  const sn = row[3] || '-';
                  const loc = row[6] || '-';
                  return (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => {
                        setSelectedMobileItem({ row, idx });
                        setMobileSearchQuery(name);
                        setShowMobileSuggestions(false);
                      }}
                      className="w-full px-4 py-3 text-left hover:bg-slate-850 transition-colors flex flex-col gap-1 active:bg-slate-800 cursor-pointer"
                    >
                      <span className="text-xs font-black text-slate-100 line-clamp-1">{name}</span>
                      <div className="flex flex-wrap items-center gap-x-2.5 gap-y-0.5 text-[10px] text-slate-400 font-medium">
                        <span className="font-mono bg-slate-800 px-1 py-0.2 rounded border border-slate-700/50">P/N: {pn}</span>
                        <span>S/N: {sn}</span>
                        <span className="text-amber-400 font-bold">Yer: {loc}</span>
                      </div>
                    </button>
                  );
                })}
              </div>
            )}

            {showMobileSuggestions && mobileSearchQuery.trim().length >= 2 && mobileSuggestions.length === 0 && (
              <div className="absolute top-[72px] left-0 w-full bg-slate-900 border border-slate-800 rounded-xl p-4 text-center z-50 shadow-xl">
                <p className="text-xs text-slate-400">Eşleşen malzeme bulunamadı.</p>
              </div>
            )}
          </div>

          {/* STEP 3: DETAILS CARD */}
          {selectedMobileItem ? (
            <div className="flex flex-col gap-3 animate-fade-in mt-2">
              <div className="flex items-center justify-between">
                <label className="text-xs font-black tracking-wider text-slate-400 uppercase">
                  3. MALZEME DETAYI & LOKASYONU
                </label>
                <button
                  type="button"
                  onClick={() => setSelectedMobileItem(null)}
                  className="text-[10px] font-black text-slate-400 hover:text-white uppercase tracking-wider cursor-pointer"
                >
                  Kartı Kapat
                </button>
              </div>

              {/* THE BEAUTIFUL CARD */}
              <div className="bg-slate-900 rounded-2xl border border-slate-800 shadow-2xl p-5 flex flex-col gap-4">
                <div className="flex items-start gap-3">
                  <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-amber-600 to-amber-700 flex items-center justify-center border border-amber-500/20 shrink-0">
                    <Package className="w-5 h-5 text-white" />
                  </div>
                  <div>
                    <h3 className="text-sm font-extrabold text-white leading-tight">
                      {selectedMobileItem.row[1] || 'İsimsiz Malzeme'}
                    </h3>
                    <span className="text-[10px] text-slate-500 uppercase font-black block mt-1 tracking-wider">
                      SIRA NO: {selectedMobileItem.row[0] || selectedMobileItem.idx + 1}
                    </span>
                  </div>
                </div>

                <hr className="border-slate-800" />

                {/* HIGHLIGHTED LOCATION BOX */}
                <div className="bg-amber-500/10 border-2 border-amber-500/40 rounded-xl p-3.5 flex items-center gap-3 shadow-inner">
                  <div className="p-2 bg-amber-500/20 text-amber-400 rounded-xl">
                    <MapPin className="w-5 h-5" />
                  </div>
                  <div className="flex-1">
                    <span className="text-[9px] font-black text-slate-400 uppercase tracking-wider leading-none block">
                      BULUNDUĞU DEPO / RAF YERİ
                    </span>
                    <span className="text-lg font-black text-amber-400 tracking-tight leading-normal uppercase">
                      {selectedMobileItem.row[6] || '-'}
                    </span>
                  </div>
                </div>

                {/* OTHER METRICS */}
                <div className="grid grid-cols-2 gap-3 mt-1">
                  <div className="bg-slate-950 rounded-xl p-3 border border-slate-800">
                    <span className="text-[9px] font-black text-slate-500 block leading-none uppercase">STOK MİKTARI</span>
                    <span className="text-sm font-extrabold text-emerald-400 mt-1 block">
                      {selectedMobileItem.row[4] || '0'} {selectedMobileItem.row[5] || 'ADET'}
                    </span>
                  </div>

                  <div className="bg-slate-950 rounded-xl p-3 border border-slate-800">
                    <span className="text-[9px] font-black text-slate-500 block leading-none uppercase">PARÇA NO (P/N)</span>
                    <span className="text-xs font-mono font-bold text-slate-200 mt-1 block truncate">
                      {selectedMobileItem.row[2] || '-'}
                    </span>
                  </div>

                  <div className="bg-slate-950 rounded-xl p-3 border border-slate-800">
                    <span className="text-[9px] font-black text-slate-500 block leading-none uppercase">SERİ NO (S/N)</span>
                    <span className="text-xs font-mono text-slate-300 mt-1 block truncate">
                      {selectedMobileItem.row[3] || '-'}
                    </span>
                  </div>

                  <div className="bg-slate-950 rounded-xl p-3 border border-slate-800">
                    <span className="text-[9px] font-black text-slate-500 block leading-none uppercase">DURUMU</span>
                    <span className="text-xs font-extrabold text-blue-400 mt-1 block">
                      {selectedMobileItem.row[7] || 'FAAL'}
                    </span>
                  </div>
                </div>

                {/* SHELF LIFE SECTION IF APPLICABLE */}
                {((selectedMobileItem.row[8] || '').trim().toUpperCase() === 'EVET' ||
                  (selectedMobileItem.row[14] || '').trim().toUpperCase() === 'EVET') && (
                  <div className="bg-slate-950 border border-slate-800/80 rounded-xl p-3 flex flex-col gap-1.5">
                    <div className="flex items-center justify-between">
                      <span className="text-[9px] font-black text-rose-400 uppercase tracking-wider">
                        RAF ÖMRÜNE TABİ ÜRÜN
                      </span>
                      <span className="px-1.5 py-0.5 rounded text-[8px] font-black bg-rose-500/20 text-rose-300 border border-rose-500/30">
                        SKT KONTROL
                      </span>
                    </div>
                    <div className="flex items-center gap-2 text-xs">
                      <Calendar className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                      <span className="font-semibold text-slate-200">
                        Bitiş: {selectedMobileItem.row[11] || selectedMobileItem.row[15] || '-'}
                      </span>
                    </div>
                  </div>
                )}
              </div>
            </div>
          ) : (
            <div className="flex flex-col items-center justify-center p-8 border-2 border-dashed border-slate-800 rounded-3xl mt-4 text-center gap-3">
              <div className="p-3 bg-slate-900 text-slate-400 rounded-2xl">
                <Search className="w-6 h-6" />
              </div>
              <div>
                <p className="text-xs font-black text-slate-300 uppercase tracking-wide">MALZEME SORGULAMA HAZIR</p>
                <p className="text-[10px] text-slate-500 max-w-[200px] mt-1 leading-normal">
                  Yukarıdaki arama çubuğunu kullanarak veya depo türünü seçerek envanter detaylarını sorgulayabilirsiniz.
                </p>
              </div>
            </div>
          )}

        </div>

        {/* MOBILE FOOTER INFO */}
        <footer className="bg-slate-900 py-3 px-4 border-t border-slate-800 text-center select-none shrink-0">
          <p className="text-[9px] text-slate-500 font-black tracking-wider">
            OGM HAVACILIK DEPO YÖNETİM PORTALI • CEP SÜRÜMÜ
          </p>
        </footer>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-emerald-500 selection:text-white">
      {notification && (
        <div className="fixed top-4 right-4 z-[999] px-4 py-3 rounded-2xl bg-emerald-950 border-2 border-emerald-500 text-emerald-100 text-xs font-black shadow-2xl flex items-center gap-3 animate-fade-in">
          <CheckCircle2 className="w-4 h-4 text-emerald-400" />
          <span>{notification}</span>
        </div>
      )}

      {/* TOP HEADER */}
      {!embeddedMode && (
        <header className="bg-slate-900/90 border-b border-slate-800 backdrop-blur-md sticky top-0 z-40 px-4 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-4 select-none">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-amber-600 to-orange-800 flex items-center justify-center shadow-lg shadow-amber-950/50 border border-amber-400/30">
              <Boxes className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-black tracking-tight text-white uppercase">
                  Hangar Yedek Parça ve Depo Yönetim Sistemi
                </h1>
                <span className="px-2 py-0.5 rounded-full text-[10px] font-mono font-bold bg-amber-500/20 text-amber-400 border border-amber-500/30">
                  MÜSTAKİL SİSTEM
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Orman Genel Müdürlüğü Havacılık Dairesi • Hangar Depo Envanteri
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => downloadStandaloneSystemZip('depo')}
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

      <div className="flex-1 p-4 lg:p-6 max-w-[1700px] mx-auto w-full flex flex-col">
        {/* CATEGORY TABS */}
        <div className="bg-slate-900 rounded-3xl p-2.5 mb-6 border border-slate-800 shadow-xl overflow-x-auto scrollbar-none flex items-center gap-2 select-none shrink-0">
          {DEPO_CATEGORIES.map(c => {
            const count = (depoData[c.key] || []).length;
            const isSelected = selectedCategoryKey === c.key;
            return (
              <button
                key={c.key}
                type="button"
                onClick={() => setSelectedCategoryKey(c.key)}
                className={`px-5 py-2.5 rounded-2xl text-xs font-black whitespace-nowrap transition-all flex items-center gap-2 cursor-pointer ${
                  isSelected
                    ? 'bg-amber-600 text-white shadow-lg shadow-amber-950/40 border border-amber-400/50 scale-102'
                    : 'bg-slate-800/80 text-slate-400 hover:text-slate-200 hover:bg-slate-800 border border-slate-700/50'
                }`}
              >
                {c.key === 'kimyasal' ? <FlaskConical className="w-4 h-4" /> : <Layers className="w-4 h-4" />}
                <span>{c.shortLabel}</span>
                <span className={`px-2 py-0.5 rounded-full text-[10px] font-mono font-bold ${
                  isSelected ? 'bg-black/30 text-amber-200' : 'bg-slate-900 text-slate-400'
                }`}>
                  {count}
                </span>
              </button>
            );
          })}
        </div>

        {/* 100% IDENTICAL SEARCH AND UTILITY CONTROLS TOOLBAR (EXACT MATCH OF image.png) */}
        <div className="bg-slate-900 text-slate-200 rounded-3xl p-5 mb-6 flex flex-col xl:flex-row gap-4 items-center justify-between shadow-xl border border-slate-800 print:hidden select-none">
          <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto">
            <div className="bg-slate-800 p-2.5 rounded-2xl border border-slate-700">
              <Search className="w-5 h-5 text-emerald-400" />
            </div>

            {/* Metin Arama Input */}
            <div className="relative flex-1 sm:flex-initial">
              <input
                type="text"
                placeholder="Malzeme, P/N veya seri no..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="bg-slate-800 text-white font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-48 border border-slate-700 placeholder-slate-400"
              />
            </div>

            {/* Firma Filtreleme Dropdown */}
            <div className="relative flex-1 sm:flex-initial">
              <select
                value={firmaFilter}
                onChange={(e) => setFirmaFilter(e.target.value)}
                className="bg-slate-800 text-emerald-300 font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-52 border border-slate-700 cursor-pointer shadow-sm"
              >
                <option value="">🏢 TÜM TEDARİKÇİLER ({uniqueFirmalar.length})</option>
                {uniqueFirmalar.map((f, i) => (
                  <option key={i} value={f}>{f}</option>
                ))}
              </select>
            </div>

            {/* Durum Filtreleme Dropdown */}
            <div className="relative flex-1 sm:flex-initial">
              <select
                value={durumFilter}
                onChange={(e) => setDurumFilter(e.target.value)}
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

            {/* Renk Koduna Seç Dropdown */}
            <div className="relative flex-1 sm:flex-initial">
              <select
                value={colorFilter}
                onChange={(e) => setColorFilter(e.target.value)}
                className="bg-slate-800 text-cyan-300 font-black text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-56 border border-slate-700 cursor-pointer shadow-sm"
              >
                <option value="ALL">🎨 RENK KODUNA SEÇ (TÜMÜ)</option>
                <option value="RED">🔴 KIRMIZI (SÜRESİ DOLAN / ACİL)</option>
                <option value="ORANGE">🟠 TURUNCU (&lt; 90 GÜN KALANLAR)</option>
                <option value="GREEN">🟢 YEŞİL (SÜRESİ UYGUN / FAAL)</option>
                <option value="GRAY">⚪ GRİ (RAF ÖMRÜNE TABİ DEĞİL)</option>
              </select>
            </div>

            {(searchQuery || firmaFilter || durumFilter || colorFilter !== "ALL") && (
              <button
                onClick={() => {
                  setSearchQuery("");
                  setFirmaFilter("");
                  setDurumFilter("");
                  setColorFilter("ALL");
                }}
                className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition-all border border-slate-700 cursor-pointer"
              >
                ✕ Temizle
              </button>
            )}

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
                <option value="red">🔴 KIRMIZI (Raf Ömrü Biten)</option>
                <option value="orange">🟠 TURUNCU (90 Gün Altı)</option>
                <option value="green">🟢 YEŞİL (Faal / Güvenli)</option>
                <option value="neutral">⚪ NÖTR (Süresiz)</option>
              </select>
            </div>

            <button
              onClick={() => setSortByColor(!sortByColor)}
              className={`px-4 py-3 active:scale-95 font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border shrink-0 ${
                sortByColor
                  ? 'bg-red-600 hover:bg-red-700 text-white border-red-500 animate-pulse'
                  : 'bg-slate-800 hover:bg-slate-750 text-slate-200 border-slate-700'
              }`}
            >
              <SlidersHorizontal className="w-4 h-4 text-amber-400" />
              <span>{sortByColor ? "🔴 RAF ÖMRÜ SIRALAMASI AKTİF" : "⏳ RAF ÖMRÜNE GÖRE SIRALA"}</span>
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-3 w-full xl:w-auto justify-end">
            <button
              onClick={() => requirePassword(() => {
                setNewRowData({});
                setIsNewProductOpen(true);
              })}
              className="px-4 py-3 bg-emerald-600 hover:bg-emerald-500 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-950/20 border border-emerald-500 shrink-0"
            >
              <PlusCircle className="w-4 h-4 text-white" />
              <span>YENİ ÜRÜN EKLE</span>
            </button>

            <button
              onClick={() => requirePassword(() => setIsTxModalOpen(true))}
              className="px-4 py-3 bg-amber-700 hover:bg-amber-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-amber-500 shrink-0"
            >
              <Boxes className="w-4 h-4 text-amber-200" />
              <span>DEPO YÖNETİM</span>
            </button>

            <button
              onClick={exportToExcel}
              className="px-4 py-3 bg-emerald-700 hover:bg-emerald-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-900/10 border border-emerald-600 shrink-0"
            >
              <Download className="w-4 h-4" />
              <span>EXCEL İNDİR</span>
            </button>

            <label className="px-4 py-3 bg-teal-700 hover:bg-teal-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-teal-500 shrink-0">
              <Upload className="w-4 h-4" />
              <span>EXCEL YÜKLE</span>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={async (e) => {
                  const file = e.target.files?.[0];
                  if (!file) return;
                  const reader = new FileReader();
                  reader.onload = async (evt) => {
                    try {
                      const data = new Uint8Array(evt.target?.result as ArrayBuffer);
                      const workbook = XLSX.read(data, { type: 'array' });
                      const sheet = workbook.Sheets[workbook.SheetNames[0]];
                      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
                      if (!raw || raw.length < 2) {
                        showNotification('⚠️ Excel dosyası boş veya satır bulunamadı.');
                        return;
                      }
                      const newRows: string[][] = [];
                      for (let r = 1; r < raw.length; r++) {
                        const row = raw[r];
                        if (!row || !row[1]) continue;
                        newRows.push([
                          String(newRows.length + 1),
                          String(row[1] || '').trim(),
                          String(row[2] || '-').trim(),
                          String(row[3] || 'N/A').trim(),
                          String(row[4] || '1').trim(),
                          String(row[5] || 'ADET').trim(),
                          String(row[6] || 'DEPO').trim(),
                          String(row[7] || 'FAAL').trim(),
                          String(row[8] || 'EVET').trim(),
                          String(row[9] || '').trim(),
                          String(row[10] || '').trim(),
                          String(row[11] || '-').trim(),
                          String(row[12] || '').trim()
                        ]);
                      }
                      saveRows(selectedCategoryKey, newRows);
                      showNotification(`✅ ${newRows.length} ürün yüklendi ve hem sisteme hem Google Drive Excel'e kalıcı aktarıldı!`);
                    } catch (err: any) {
                      showNotification(`❌ Excel hatası: ${err.message}`);
                    }
                  };
                  reader.readAsArrayBuffer(file);
                  e.target.value = '';
                }}
              />
            </label>

            <button
              onClick={() => downloadStandaloneSystemZip('depo')}
              className="px-4 py-3 bg-indigo-700 hover:bg-indigo-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg border border-indigo-500 shrink-0"
            >
              <Package className="w-4 h-4 text-indigo-200" />
              <span>SİTEYİ ZİP İNDİR</span>
            </button>
          </div>
        </div>

        {/* DYNAMIC TABLE CONTAINER */}
        <div className="flex-1 bg-white border-2 border-slate-200/60 rounded-[2.5rem] shadow-xl overflow-hidden flex flex-col print:border-none print:shadow-none min-h-[500px]">
          <div className="bg-slate-900 px-6 py-4 border-b border-slate-800 flex items-center justify-between print:hidden shrink-0 flex-wrap gap-3">
            {Object.keys(selectedItems).filter(k => selectedItems[k]).length > 0 ? (
              <div className="flex items-center gap-3 animate-fade-in flex-wrap">
                <span className="text-xs font-black text-amber-400 bg-amber-950/40 border border-amber-900 px-3 py-1.5 rounded-xl">
                  ⚡ {Object.keys(selectedItems).filter(k => selectedItems[k]).length} MALZEME SEÇİLDİ
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
                  🔴 Ömür Geçmiş
                </span>
              </div>
            )}

            <span className="text-[10px] font-mono font-black text-slate-400 bg-slate-800 px-3 py-1 rounded-full border border-slate-700">
              {processedRows.length} KALEM MALZEME LİSTELENDİ
            </span>
          </div>

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
                  {DEPO_TABLE_COLUMNS.map((col, idx) => (
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
                {paginatedRows.length === 0 ? (
                  <tr>
                    <td colSpan={DEPO_TABLE_COLUMNS.length + 1} className="py-16 text-center text-slate-400">
                      <div className="flex flex-col items-center justify-center gap-3">
                        <Boxes className="w-8 h-8 text-slate-300" />
                        <span className="font-bold">Eşleşen depo malzemesi bulunamadı.</span>
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
                          isSelected ? 'bg-amber-50/70 hover:bg-amber-100/70' : 'bg-white hover:bg-slate-50'
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
                        <td className="px-4 py-3 text-center font-bold text-slate-500 border-r border-slate-100 w-16">
                          {idx + 1}
                        </td>
                        <td className="px-4 py-3 font-bold text-slate-900 border-r border-slate-100 max-w-xs">
                          {row[1] || '-'}
                        </td>
                        <td className="px-4 py-3 font-mono font-bold text-slate-700 border-r border-slate-100">
                          {row[2] || '-'}
                        </td>
                        <td className="px-4 py-3 font-mono text-slate-600 border-r border-slate-100">
                          {row[3] || '-'}
                        </td>
                        <td className="px-4 py-3 text-center font-bold text-slate-800 border-r border-slate-100">
                          {row[4] || '1'}
                        </td>
                        <td className="px-4 py-3 text-center text-slate-700 border-r border-slate-100 font-semibold">
                          {row[5] || 'ADET'}
                        </td>
                        <td className="px-4 py-3 text-slate-700 border-r border-slate-100">
                          {row[6] || '-'}
                        </td>
                        <td className="px-4 py-3 text-center border-r border-slate-100">
                          <span className={`px-2.5 py-1 rounded-lg text-[10px] font-black ${
                            Number(row[4] || 0) > 0
                              ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                              : 'bg-rose-100 text-rose-800 border border-rose-300'
                          }`}>
                            {Number(row[4] || 0) > 0 ? 'VAR' : 'YOK'}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-center font-bold text-slate-700 border-r border-slate-100">
                          {row[8] || 'EVET'}
                        </td>
                        <td className="px-4 py-3 text-center font-mono text-slate-600 border-r border-slate-100">
                          {row[9] || '-'}
                        </td>
                        <td className={`px-4 py-3 text-center font-mono border-r border-slate-100 ${dateColorClass}`}>
                          {row[10] || '-'}
                        </td>
                        <td className="px-4 py-3 text-slate-700 border-r border-slate-100">
                          {row[11] || '-'}
                        </td>
                        <td className="px-4 py-3 text-slate-500 max-w-sm truncate border-r border-slate-100">
                          {row[12] || '-'}
                        </td>
                        <td className="px-4 py-3 text-center border-r border-slate-100 font-bold">
                          <span className={`px-2 py-0.5 rounded-full text-[10px] font-black ${
                            String(row[13] || row[8] || '').toUpperCase().includes('EVET')
                              ? 'bg-amber-100 text-amber-900 border border-amber-300'
                              : 'bg-slate-100 text-slate-600'
                          }`}>
                            {String(row[13] || row[8] || '').toUpperCase().includes('EVET') ? 'EVET' : 'HAYIR'}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-center font-mono font-bold text-slate-700">
                          {String(row[13] || row[8] || '').toUpperCase().includes('EVET') 
                            ? (row[14] || row[10] || '-') 
                            : '-'}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* SAYFALAMA KONTROLLERİ (1.779 ÜRÜNDE SIFIR KASILMA & VIRTUALIZATION HIZI) */}
          <div className="bg-slate-900 px-5 py-3 border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-300 select-none">
            <div className="flex items-center gap-3">
              <span className="text-slate-400 font-mono text-[11px]">Sayfa Başına Ürün:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="bg-slate-800 border border-slate-700 text-white rounded-lg px-2.5 py-1 font-mono font-bold text-xs cursor-pointer focus:ring-1 focus:ring-emerald-500 outline-none"
              >
                <option value={50}>50 Ürün</option>
                <option value={100}>100 Ürün / Sayfa (Önerilen)</option>
                <option value={200}>200 Ürün</option>
                <option value={500}>500 Ürün</option>
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

              <span className="px-3.5 py-1.5 bg-emerald-950/80 text-emerald-400 border border-emerald-700/60 font-black rounded-lg text-xs">
                Sayfa {currentPage} / {totalPages}
              </span>

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
              <button onClick={() => setIsPasswordModalOpen(false)} className="text-slate-400 hover:text-white">
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

      {/* NEW PRODUCT MODAL */}
      {isNewProductOpen && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border-2 border-emerald-500/60 p-6 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between mb-4 border-b border-slate-800 pb-3">
              <h3 className="text-sm font-black text-emerald-400 uppercase">
                Yeni Malzeme Ekle ({currentCategory.shortLabel})
              </h3>
              <button onClick={() => setIsNewProductOpen(false)} className="text-slate-400 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {DEPO_TABLE_COLUMNS.slice(1).map((colName, cIdx) => {
                const targetIdx = String(cIdx + 1);
                // Highlight shelf life fields
                const isShelfLife = colName.includes('RAF ÖMRÜ');
                return (
                  <div key={targetIdx} className="space-y-1">
                    <label className={`text-[10px] font-mono font-bold uppercase ${isShelfLife ? 'text-amber-400' : 'text-slate-400'}`}>
                      {colName}
                    </label>
                    <textarea
                      rows={colName.includes('TARİH') ? 1 : 2}
                      value={newRowData[targetIdx] || ''}
                      onChange={(e) => {
                        const val = e.target.value;
                        setNewRowData(prev => ({ ...prev, [targetIdx]: val }));
                      }}
                      placeholder={colName.includes('TARİH') ? 'Örn: 28.09.2026' : (colName.includes('RAF ÖMRÜ') ? 'Birden fazla tarih için Enter kullanın' : '')}
                      className={`w-full bg-slate-800 text-white text-xs p-2.5 rounded-xl border ${isShelfLife ? 'border-amber-500/50' : 'border-slate-700'} focus:outline-none focus:ring-2 focus:ring-emerald-500`}
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

      {/* EDIT MODAL */}
      {activeEditRow && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border-2 border-slate-700 p-6 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between mb-4 border-b border-slate-800 pb-3">
              <h3 className="text-sm font-black text-amber-400 uppercase">
                Malzeme Kaydı Düzenle (Sıra {activeEditRow.idx + 1})
              </h3>
              <button onClick={() => setActiveEditRow(null)} className="text-slate-400 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {DEPO_TABLE_COLUMNS.slice(1).map((colName, cIdx) => {
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

      {/* DEPO MANAGEMENT / TRANSACTION MODAL */}
      {isTxModalOpen && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border-2 border-amber-500/60 p-6 rounded-3xl shadow-2xl max-w-lg w-full">
            <div className="flex items-center justify-between mb-4 border-b border-slate-800 pb-3">
              <h3 className="text-sm font-black text-amber-400 uppercase flex items-center gap-2">
                <Boxes className="w-4 h-4 text-amber-400" />
                <span>Depo Giriş / Çıkış & Transfer İşlemi</span>
              </h3>
              <button onClick={() => setIsTxModalOpen(false)} className="text-slate-400 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="space-y-4">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setTxType('out')}
                  className={`flex-1 py-2.5 rounded-xl text-xs font-bold flex items-center justify-center gap-2 ${
                    txType === 'out' ? 'bg-rose-700 text-white' : 'bg-slate-800 text-slate-400'
                  }`}
                >
                  <ArrowUpRight className="w-4 h-4" />
                  <span>Depo Çıkış / Sarf</span>
                </button>
                <button
                  type="button"
                  onClick={() => setTxType('in')}
                  className={`flex-1 py-2.5 rounded-xl text-xs font-bold flex items-center justify-center gap-2 ${
                    txType === 'in' ? 'bg-emerald-700 text-white' : 'bg-slate-800 text-slate-400'
                  }`}
                >
                  <ArrowDownLeft className="w-4 h-4" />
                  <span>Depo Giriş / Kabul</span>
                </button>
              </div>

              <div>
                <label className="text-[10px] font-mono text-slate-400 uppercase font-bold">Hedef / Çıkış & Transfer Türü</label>
                <div className="flex flex-wrap gap-1 mt-1 mb-1.5">
                  {[
                    'ANKARA GELEN',
                    'ANKARA ÇIKAN',
                    'KARAİN TRANSFER',
                    'KARAİN ÇIKAN',
                    'ÇANAKKALE TRANSFER',
                    'ÇANAKKALE ÇIKAN',
                    'MİLAS TRANSFER',
                    'MİLAS ÇIKAN',
                    'BURSA TRANSFER',
                    'BURSA ÇIKAN'
                  ].map((tt) => (
                    <button
                      key={tt}
                      type="button"
                      onClick={() => setTxData(prev => ({ ...prev, target: tt }))}
                      className={`text-[9px] px-2 py-0.5 rounded-md font-bold transition ${
                        txData.target === tt
                          ? 'bg-amber-500 text-amber-950 ring-1 ring-amber-400'
                          : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                      }`}
                    >
                      {tt}
                    </button>
                  ))}
                </div>
                <input
                  type="text"
                  value={txData.target}
                  onChange={(e) => setTxData(prev => ({ ...prev, target: e.target.value }))}
                  placeholder="Örn: ANKARA ÇIKAN, KARAİN TRANSFER, MİLAS ÇIKAN..."
                  className="w-full bg-slate-800 text-white text-xs p-2.5 rounded-xl border border-slate-700 mt-1 font-bold"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] font-mono text-slate-400 uppercase font-bold">Miktar</label>
                  <input
                    type="number"
                    min="1"
                    value={txData.qty}
                    onChange={(e) => setTxData(prev => ({ ...prev, qty: parseInt(e.target.value) || 1 }))}
                    className="w-full bg-slate-800 text-white text-xs p-2.5 rounded-xl border border-slate-700 mt-1 font-mono font-bold text-center"
                  />
                </div>
                <div>
                  <label className="text-[10px] font-mono text-slate-400 uppercase font-bold">Teslim Alan / Eden Personel</label>
                  <input
                    type="text"
                    value={txData.person}
                    onChange={(e) => setTxData(prev => ({ ...prev, person: e.target.value }))}
                    placeholder="Ad Soyad"
                    className="w-full bg-slate-800 text-white text-xs p-2.5 rounded-xl border border-slate-700 mt-1"
                  />
                </div>
              </div>

              <div>
                <label className="text-[10px] font-mono text-slate-400 uppercase font-bold">Açıklama / İş Emri No</label>
                <input
                  type="text"
                  value={txData.notes}
                  onChange={(e) => setTxData(prev => ({ ...prev, notes: e.target.value }))}
                  placeholder="Örn: WO-2026-084 periyodik bakım sarfı"
                  className="w-full bg-slate-800 text-white text-xs p-2.5 rounded-xl border border-slate-700 mt-1"
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-4">
                <button
                  type="button"
                  onClick={() => setIsTxModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-bold"
                >
                  İptal
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setIsTxModalOpen(false);
                    showNotification(`✅ Depo ${txType === 'out' ? 'çıkış' : 'giriş'} işlemi kaydedildi.`);
                  }}
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black"
                >
                  İşlemi Tamamla
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default DepoApp;
