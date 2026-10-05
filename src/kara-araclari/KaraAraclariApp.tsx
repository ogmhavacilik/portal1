import React, { useState, useEffect, useMemo } from 'react';
import {
  Truck,
  Car,
  Search,
  SlidersHorizontal,
  RefreshCw,
  PlusCircle,
  Download,
  Trash2,
  ArrowLeft,
  CheckCircle2,
  X,
  Lock,
  Package,
  FileText,
  History,
  Calendar,
  ShieldCheck,
  Fuel
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { cleanAndFormatDateString } from '../utils/driveExcelSync';
import { publishCrossSystemEvent, subscribeCrossSystemEvents } from '../utils/crossSystemBridge';
import { downloadStandaloneSystemZip } from '../utils/zipExporter';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycby2TM-nlURR5iPR4s4I6BpE8hbor93Jin9g014k3XPaQ0rYtS2MWHwtlnlAoph8Y3mZ/exec";
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export const KARA_ARACLARI_COLUMNS = [
  "SIRA NO",
  "ARAÇ PLAKASI & MARKA MODEL",
  "ŞASİ NO / MOTOR NO",
  "ARAÇ TİPİ / CİNSİ",
  "MİKTAR",
  "BAĞLI BULUNDUĞU BİRİM",
  "DURUMU",
  "PERİYODİK MUAYENE TABİ",
  "SON MUAYENE / BAKIM TARİHİ",
  "GELECEK MUAYENE / BAKIM TARİHİ",
  "KASKO & SİGORTA BİTİŞ",
  "GÜNCEL KİLOMETRE / SAAT",
  "AÇIKLAMA"
];

export function KaraAraclariApp({ isEmbedded }: { isEmbedded?: boolean }) {
  const params = new URLSearchParams(window.location.search);
  const searchFromUrl = params.get('search') || '';
  const embeddedMode = isEmbedded ?? (params.get('embedded') === 'true');

  const [activeSubTab, setActiveSubTab] = useState<'list' | 'order_entry' | 'past_records'>('list');
  const [searchQuery, setSearchQuery] = useState<string>(searchFromUrl);
  const [firmaFilter, setFirmaFilter] = useState<string>('');
  const [durumFilter, setDurumFilter] = useState<string>('');
  const [colorFilter, setColorFilter] = useState<string>('ALL');
  const [selectedColorFilter, setSelectedColorFilter] = useState<string>('all');
  const [sortByColor, setSortByColor] = useState<boolean>(false);
  const [selectedItems, setSelectedItems] = useState<Record<string, boolean>>({});
  const [isLoadingDrive, setIsLoadingDrive] = useState<boolean>(false);
  const [notification, setNotification] = useState<string | null>(null);

  // Modals
  const [activeEditRow, setActiveEditRow] = useState<{ idx: number; row: string[] } | null>(null);
  const [isNewVehicleOpen, setIsNewVehicleOpen] = useState<boolean>(false);
  const [newRowData, setNewRowData] = useState<Record<string, string>>({});

  // Password Lock
  const [isPasswordModalOpen, setIsPasswordModalOpen] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [passwordError, setPasswordError] = useState<boolean>(false);
  const [pendingPasswordAction, setPendingPasswordAction] = useState<(() => void) | null>(null);

  // Vehicle data
  const [vehicleRows, setVehicleRows] = useState<string[][]>(() => {
    try {
      const stored = localStorage.getItem('excel_kara_araclari_data');
      if (stored) return JSON.parse(stored);
    } catch {}
    return [
      ['1', '06 OGM 101 - FORD TRANSIT ÇİFT KABİN', 'NM0XXXTT12345', 'HAFİF TİCARİ', '1', 'ANKARA HANGAR ŞUBE', 'FAAL', 'EVET', '15.01.2024', '15.01.2025', '01.06.2025', '84,500 KM', 'Hangar nöbetçi ve intikal aracı'],
      ['2', '06 OGM 102 - ISUZU D-MAX 4X4 YANGIN', 'MPATFS85J9876', 'ARAZİ ARACI', '1', 'ESENBOĞA UÇAK HANGARI', 'FAAL', 'EVET', '20.03.2024', '20.03.2025', '12.08.2025', '62,100 KM', 'Pist ve arazi kontrol aracı'],
      ['3', '06 OGM 103 - MERCEDES ATEGO YAKIT TANKERİ', 'WDB9702251K', 'JET-A1 TANKER', '1', 'YANGIN HAREKAT', 'FAAL', 'EVET', '10.02.2024', '10.02.2025', '15.09.2025', '41,200 KM', 'AT-802 mobil yakıt ikmal tankeri'],
      ['4', '06 OGM 104 - TOYOTA HILUX HANGAR', 'AHTFR22G456', 'PİKAP', '1', 'TEKNİK BAKIM ŞUBE', 'FAAL', 'EVET', '05.04.2024', '05.04.2025', '22.10.2025', '110,400 KM', 'Alet ve malzeme nakil aracı']
    ];
  });

  // Task Orders
  const [taskOrders, setTaskOrders] = useState<any[]>(() => {
    try {
      const stored = localStorage.getItem('kara_araclari_gorev_emirleri');
      if (stored) return JSON.parse(stored);
    } catch {}
    return [
      { id: 'GE-2026-001', plaka: '06 OGM 101', driver: 'Ahmet Yılmaz', destination: 'Esenboğa Hangar', date: '12.09.2026', status: 'Tamamlandı' },
      { id: 'GE-2026-002', plaka: '06 OGM 103', driver: 'Mehmet Demir', destination: 'Antalya Orman Yangın Üssü', date: '14.09.2026', status: 'Görevde' }
    ];
  });

  const [newOrder, setNewOrder] = useState({ plaka: '06 OGM 101', driver: '', dest: '', date: new Date().toISOString().slice(0, 10), notes: '' });

  const showNotification = (msg: string) => {
    setNotification(msg);
    setTimeout(() => setNotification(null), 4000);
  };

  const saveRows = (rows: string[][]) => {
    setVehicleRows(rows);
    try {
      localStorage.setItem('excel_kara_araclari_data', JSON.stringify(rows));
    } catch {}
    publishCrossSystemEvent('KARA_ARACI_UPDATED', 'kara-araclari', { count: rows.length });
  };

  useEffect(() => {
    const unsub = subscribeCrossSystemEvents((ev) => {
      if (ev.type === 'REFRESH_ALL') {
        try {
          const stored = localStorage.getItem('excel_kara_araclari_data');
          if (stored) setVehicleRows(JSON.parse(stored));
        } catch {}
      }
    });
    return unsub;
  }, []);

  // Sayfa açılışında Google Drive ile arka planda anlık senkronize et (Depo modülüyle aynı açılış mimarisi)
  useEffect(() => {
    pullFromDrive();
  }, []);

  const uniqueFirmalar = useMemo(() => {
    const s = new Set<string>();
    for (const r of vehicleRows) {
      const f = (r[5] || '').trim();
      if (f && f !== '-' && f !== '--') s.add(f);
    }
    return Array.from(s).sort();
  }, [vehicleRows]);

  const uniqueDurumlar = useMemo(() => {
    const s = new Set<string>();
    for (const r of vehicleRows) {
      const d = (r[6] || '').trim().toUpperCase();
      if (d && d !== '-' && d !== '--') s.add(d);
    }
    return Array.from(s).sort();
  }, [vehicleRows]);

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
    const isMuayeneTabi = (row[7] || '').trim().toUpperCase();
    if (isMuayeneTabi === 'HAYIR' || isMuayeneTabi === 'MUAFIYET') return 'neutral';
    const days = getDaysUntil(row[9]);
    if (days === null) return 'neutral';
    if (days < 0) return 'red';
    if (days <= 90) return 'orange';
    return 'green';
  };

  const processedRows = useMemo(() => {
    let result = vehicleRows.map((row, idx) => ({ row, idx }));

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter(({ row }) =>
        row.some(cell => String(cell || '').toLowerCase().includes(q))
      );
    }

    if (firmaFilter) {
      result = result.filter(({ row }) => (row[5] || '').trim() === firmaFilter);
    }

    if (durumFilter) {
      result = result.filter(({ row }) => (row[6] || '').trim().toUpperCase() === durumFilter);
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
        const dA = getDaysUntil(a.row[9]) ?? 9999;
        const dB = getDaysUntil(b.row[9]) ?? 9999;
        return dA - dB;
      });
    }

    return result;
  }, [vehicleRows, searchQuery, firmaFilter, durumFilter, colorFilter, selectedColorFilter, sortByColor]);

  const pullFromDrive = async () => {
    setIsLoadingDrive(true);
    showNotification("Google Drive'dan kara araçları güncel verileri alınıyor...");
    try {
      const url = `${GOOGLE_SCRIPT_URL}?action=read_excel&folderId=${DRIVE_FOLDER_ID}&fileName=${encodeURIComponent('hangar_yer_destek_kara_araclari.xlsx')}&t=${Date.now()}`;
      const res = await fetch(url);
      const json = await res.json();
      if (json && json.data && Array.isArray(json.data) && json.data.length > 0) {
        let rows: string[][] = json.data;
        if (rows[0] && (String(rows[0][0]).includes('SIRA') || String(rows[0][1]).includes('PLAKA'))) {
          rows = rows.slice(1);
        }
        const cleanedRows = rows.map((r, i) => {
          const rowArr = [...r];
          rowArr[0] = String(i + 1);
          if (rowArr[8]) rowArr[8] = cleanAndFormatDateString(rowArr[8]);
          if (rowArr[9]) rowArr[9] = cleanAndFormatDateString(rowArr[9]);
          return rowArr;
        });
        saveRows(cleanedRows);
        showNotification(`✅ Kara araçları güncellendi: ${cleanedRows.length} araç aktarıldı.`);
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
      KARA_ARACLARI_COLUMNS,
      ...processedRows.map(({ row }) => KARA_ARACLARI_COLUMNS.map((_, cIdx) => row[cIdx] || ''))
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Kara Araçları');
    XLSX.writeFile(wb, `hangar_kara_araclari_${new Date().toISOString().slice(0, 10)}.xlsx`);
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
      const remaining = vehicleRows.filter((_, idx) => !selectedIndices.has(idx)).map((r, i) => {
        r[0] = String(i + 1);
        return r;
      });
      saveRows(remaining);
      setSelectedItems({});
      showNotification(`✅ Seçilen ${selectedIndices.size} araç kaydı silindi.`);
    });
  };

  const handleSaveNewVehicle = () => {
    const newRow = [
      String(vehicleRows.length + 1),
      newRowData['1'] || '',
      newRowData['2'] || '-',
      newRowData['3'] || 'HAFİF TİCARİ',
      newRowData['4'] || '1',
      newRowData['5'] || 'HANGAR BAKIM',
      newRowData['6'] || 'FAAL',
      newRowData['7'] || 'EVET',
      newRowData['8'] || cleanAndFormatDateString(new Date().toLocaleDateString('tr-TR')),
      newRowData['9'] || '',
      newRowData['10'] || '',
      newRowData['11'] || '0 KM',
      newRowData['12'] || ''
    ];
    const updated = [...vehicleRows, newRow];
    saveRows(updated);
    setIsNewVehicleOpen(false);
    setNewRowData({});
    showNotification('✅ Yeni araç başarıyla eklendi.');
  };

  const handleSaveEditRow = () => {
    if (!activeEditRow) return;
    const updated = [...vehicleRows];
    updated[activeEditRow.idx] = activeEditRow.row;
    saveRows(updated);
    setActiveEditRow(null);
    showNotification('✅ Araç kaydı güncellendi.');
  };

  const handleCreateTaskOrder = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newOrder.driver || !newOrder.dest) {
      showNotification('Lütfen sürücü ve varış yerini doldurunuz.');
      return;
    }
    const orderObj = {
      id: `GE-2026-${String(taskOrders.length + 1).padStart(3, '0')}`,
      plaka: newOrder.plaka,
      driver: newOrder.driver,
      destination: newOrder.dest,
      date: newOrder.date,
      status: 'Görevde'
    };
    const updated = [orderObj, ...taskOrders];
    setTaskOrders(updated);
    try {
      localStorage.setItem('kara_araclari_gorev_emirleri', JSON.stringify(updated));
    } catch {}
    setNewOrder({ plaka: '06 OGM 101', driver: '', dest: '', date: new Date().toISOString().slice(0, 10), notes: '' });
    setActiveSubTab('past_records');
    showNotification('✅ Görev emri başarıyla oluşturuldu.');
  };

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
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-800 flex items-center justify-center shadow-lg shadow-blue-950/50 border border-blue-400/30">
              <Truck className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-black tracking-tight text-white uppercase">
                  Hangar Yer Destek Kara Araçları Takip Sistemi
                </h1>
                <span className="px-2 py-0.5 rounded-full text-[10px] font-mono font-bold bg-blue-500/20 text-blue-400 border border-blue-500/30">
                  MÜSTAKİL SİSTEM
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Orman Genel Müdürlüğü Havacılık Dairesi • Hangar Araç Filosu ve Görev Yönetimi
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => downloadStandaloneSystemZip('kara-araclari')}
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
        {/* SUB-TABS (EXACT MATCH OF ORIGINAL PORTAL KARA ARACLARI TABS) */}
        <div className="flex flex-wrap items-center gap-3 mb-6 select-none">
          <button
            onClick={() => setActiveSubTab('list')}
            className={`px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-wider transition-all cursor-pointer flex items-center gap-2 border shadow-sm ${
              activeSubTab === 'list'
                ? 'bg-[#0b3d1d] text-white border-[#0b3d1d] scale-105 shadow-md shadow-emerald-900/20'
                : 'bg-slate-900 hover:bg-slate-800 text-slate-300 border-slate-700'
            }`}
          >
            <Car className="w-4 h-4" />
            🚗 ARAÇ LİSTESİ & MUAYENE
          </button>
          <button
            onClick={() => setActiveSubTab('order_entry')}
            className={`px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-wider transition-all cursor-pointer flex items-center gap-2 border shadow-sm ${
              activeSubTab === 'order_entry'
                ? 'bg-[#0b3d1d] text-white border-[#0b3d1d] scale-105 shadow-md shadow-emerald-900/20'
                : 'bg-slate-900 hover:bg-slate-800 text-slate-300 border-slate-700'
            }`}
          >
            <FileText className="w-4 h-4" />
            📋 GÖREV EMRİ GİRİŞ
          </button>
          <button
            onClick={() => setActiveSubTab('past_records')}
            className={`px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-wider transition-all cursor-pointer flex items-center gap-2 border shadow-sm ${
              activeSubTab === 'past_records'
                ? 'bg-[#0b3d1d] text-white border-[#0b3d1d] scale-105 shadow-md shadow-emerald-900/20'
                : 'bg-slate-900 hover:bg-slate-800 text-slate-300 border-slate-700'
            }`}
          >
            <History className="w-4 h-4" />
            📜 GEÇMİŞ KAYITLAR
          </button>
        </div>

        {activeSubTab === 'list' && (
          <>
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
                    placeholder="Araç veya plaka no..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="bg-slate-800 text-white font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-48 border border-slate-700 placeholder-slate-400"
                  />
                </div>

                {/* Bağlı Birim Filtreleme Dropdown */}
                <div className="relative flex-1 sm:flex-initial">
                  <select
                    value={firmaFilter}
                    onChange={(e) => setFirmaFilter(e.target.value)}
                    className="bg-slate-800 text-emerald-300 font-extrabold text-xs px-4 py-3 rounded-2xl focus:outline-none focus:ring-4 focus:ring-emerald-500/20 w-full sm:w-52 border border-slate-700 cursor-pointer shadow-sm"
                  >
                    <option value="">🏢 TÜM BİRİMLER ({uniqueFirmalar.length})</option>
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
                    {uniqueDurumlar.map((d, i) => (
                      <option key={i} value={d}>{d}</option>
                    ))}
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
                    <option value="RED">🔴 KIRMIZI (MUAYENESİ GEÇMİŞ)</option>
                    <option value="ORANGE">🟠 TURUNCU (&lt; 90 GÜN KALANLAR)</option>
                    <option value="GREEN">🟢 YEŞİL (SÜRESİ UYGUN / FAAL)</option>
                    <option value="GRAY">⚪ GRİ (MUAF)</option>
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
                    <option value="red">🔴 KIRMIZI (Muayenesi Geçmiş)</option>
                    <option value="orange">🟠 TURUNCU (90 Gün Altı)</option>
                    <option value="green">🟢 YEŞİL (Faal / Güvenli)</option>
                    <option value="neutral">⚪ NÖTR (Muaf)</option>
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
                  <span>{sortByColor ? "🔴 MUAYENE SIRALAMASI AKTİF" : "⏳ RENK KODUNA GÖRE SIRALA"}</span>
                </button>
              </div>

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
                    setIsNewVehicleOpen(true);
                  })}
                  className="px-4 py-3 bg-emerald-600 hover:bg-emerald-500 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-950/20 border border-emerald-500 shrink-0"
                >
                  <PlusCircle className="w-4 h-4 text-white" />
                  <span>YENİ ARAÇ EKLE</span>
                </button>

                <button
                  onClick={exportToExcel}
                  className="px-4 py-3 bg-emerald-700 hover:bg-emerald-600 active:scale-95 text-white font-black font-mono text-xs rounded-2xl flex items-center gap-2 transition-all cursor-pointer shadow-lg shadow-emerald-900/10 border border-emerald-600 shrink-0"
                >
                  <Download className="w-4 h-4" />
                  <span>EXCEL İNDİR</span>
                </button>

                <button
                  onClick={() => downloadStandaloneSystemZip('kara-araclari')}
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
                      ⚡ {Object.keys(selectedItems).filter(k => selectedItems[k]).length} ARAÇ SEÇİLDİ
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
                      🔴 Muayene Geçmiş
                    </span>
                  </div>
                )}

                <span className="text-[10px] font-mono font-black text-slate-400 bg-slate-800 px-3 py-1 rounded-full border border-slate-700">
                  {processedRows.length} KALEM ARAÇ LİSTELENDİ
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
                      {KARA_ARACLARI_COLUMNS.map((col, idx) => (
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
                        <td colSpan={KARA_ARACLARI_COLUMNS.length + 1} className="py-16 text-center text-slate-400">
                          <div className="flex flex-col items-center justify-center gap-3">
                            <Car className="w-8 h-8 text-slate-300" />
                            <span className="font-bold">Eşleşen araç kaydı bulunamadı.</span>
                          </div>
                        </td>
                      </tr>
                    ) : (
                      processedRows.map(({ row, idx }) => {
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
                              isSelected ? 'bg-blue-50/70 hover:bg-blue-100/70' : 'bg-white hover:bg-slate-50'
                            }`}
                          >
                            <td className="px-3 py-3 text-center border-r border-slate-100" onClick={(e) => e.stopPropagation()}>
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
                            <td className="px-4 py-3 font-semibold text-slate-700 border-r border-slate-100">
                              {row[3] || '-'}
                            </td>
                            <td className="px-4 py-3 text-center font-bold text-slate-800 border-r border-slate-100">
                              {row[4] || '1'}
                            </td>
                            <td className="px-4 py-3 text-slate-700 border-r border-slate-100">
                              {row[5] || '-'}
                            </td>
                            <td className="px-4 py-3 text-center border-r border-slate-100">
                              <span className={`px-2.5 py-1 rounded-lg text-[10px] font-black ${
                                (row[6] || '').trim().toUpperCase() === 'FAAL'
                                  ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                  : 'bg-rose-100 text-rose-800 border border-rose-300'
                              }`}>
                                {row[6] || 'FAAL'}
                              </span>
                            </td>
                            <td className="px-4 py-3 text-center font-bold text-slate-700 border-r border-slate-100">
                              {row[7] || 'EVET'}
                            </td>
                            <td className="px-4 py-3 text-center font-mono text-slate-600 border-r border-slate-100">
                              {row[8] || '-'}
                            </td>
                            <td className={`px-4 py-3 text-center font-mono border-r border-slate-100 ${dateColorClass}`}>
                              {row[9] || '-'}
                            </td>
                            <td className="px-4 py-3 text-center font-mono text-slate-600 border-r border-slate-100">
                              {row[10] || '-'}
                            </td>
                            <td className="px-4 py-3 font-mono font-bold text-slate-800 border-r border-slate-100">
                              {row[11] || '-'}
                            </td>
                            <td className="px-4 py-3 text-slate-500 max-w-sm truncate">
                              {row[12] || '-'}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {/* SUBTAB 2: GÖREV EMRİ GİRİŞ */}
        {activeSubTab === 'order_entry' && (
          <div className="bg-slate-900 rounded-3xl p-6 border border-slate-800 shadow-2xl max-w-2xl mx-auto w-full">
            <h2 className="text-sm font-black text-emerald-400 uppercase mb-4 flex items-center gap-2">
              <FileText className="w-5 h-5" />
              <span>Yeni Araç Görev Emri Oluştur</span>
            </h2>
            <form onSubmit={handleCreateTaskOrder} className="space-y-4">
              <div>
                <label className="text-[11px] font-mono text-slate-400 uppercase font-bold">Görev Aracı Seçiniz</label>
                <select
                  value={newOrder.plaka}
                  onChange={(e) => setNewOrder(prev => ({ ...prev, plaka: e.target.value }))}
                  className="w-full bg-slate-800 text-white text-xs p-3 rounded-xl border border-slate-700 mt-1"
                >
                  {vehicleRows.map((v, i) => (
                    <option key={i} value={v[1]}>{v[1]}</option>
                  ))}
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[11px] font-mono text-slate-400 uppercase font-bold">Görevli Sürücü / Personel</label>
                  <input
                    type="text"
                    required
                    placeholder="Ad Soyad"
                    value={newOrder.driver}
                    onChange={(e) => setNewOrder(prev => ({ ...prev, driver: e.target.value }))}
                    className="w-full bg-slate-800 text-white text-xs p-3 rounded-xl border border-slate-700 mt-1"
                  />
                </div>
                <div>
                  <label className="text-[11px] font-mono text-slate-400 uppercase font-bold">Görev Tarihi</label>
                  <input
                    type="date"
                    value={newOrder.date}
                    onChange={(e) => setNewOrder(prev => ({ ...prev, date: e.target.value }))}
                    className="w-full bg-slate-800 text-white text-xs p-3 rounded-xl border border-slate-700 mt-1"
                  />
                </div>
              </div>
              <div>
                <label className="text-[11px] font-mono text-slate-400 uppercase font-bold">Gidilecek Yer / Güzergah</label>
                <input
                  type="text"
                  required
                  placeholder="Örn: Esenboğa Hangar - Etimesgut Yerleşkesi"
                  value={newOrder.dest}
                  onChange={(e) => setNewOrder(prev => ({ ...prev, dest: e.target.value }))}
                  className="w-full bg-slate-800 text-white text-xs p-3 rounded-xl border border-slate-700 mt-1"
                />
              </div>
              <div>
                <label className="text-[11px] font-mono text-slate-400 uppercase font-bold">Görev Amacı & Açıklama</label>
                <textarea
                  rows={3}
                  placeholder="Görev talimatı ve iş emri ayrıntıları..."
                  value={newOrder.notes}
                  onChange={(e) => setNewOrder(prev => ({ ...prev, notes: e.target.value }))}
                  className="w-full bg-slate-800 text-white text-xs p-3 rounded-xl border border-slate-700 mt-1"
                />
              </div>
              <div className="flex justify-end pt-4">
                <button
                  type="submit"
                  className="px-6 py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-2xl text-xs font-black uppercase tracking-wider cursor-pointer shadow-lg shadow-emerald-950/40"
                >
                  Görev Emrini Onayla ve Kaydet
                </button>
              </div>
            </form>
          </div>
        )}

        {/* SUBTAB 3: GEÇMİŞ KAYITLAR */}
        {activeSubTab === 'past_records' && (
          <div className="bg-slate-900 rounded-3xl p-6 border border-slate-800 shadow-2xl w-full">
            <h2 className="text-sm font-black text-emerald-400 uppercase mb-4 flex items-center gap-2">
              <History className="w-5 h-5" />
              <span>Kayıtlı Görev Emirleri ve Araç Hareketleri</span>
            </h2>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-slate-200">
                <thead>
                  <tr className="border-b border-slate-800 text-[11px] font-mono font-bold text-slate-400 uppercase">
                    <th className="py-3 px-4">GÖREV NO</th>
                    <th className="py-3 px-4">ARAÇ</th>
                    <th className="py-3 px-4">SÜRÜCÜ</th>
                    <th className="py-3 px-4">GÜZERGAH</th>
                    <th className="py-3 px-4">TARİH</th>
                    <th className="py-3 px-4">DURUM</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 font-medium">
                  {taskOrders.map((ord, i) => (
                    <tr key={i} className="hover:bg-slate-800/40 transition-colors">
                      <td className="py-3 px-4 font-mono font-bold text-emerald-400">{ord.id}</td>
                      <td className="py-3 px-4 font-bold text-white">{ord.plaka}</td>
                      <td className="py-3 px-4 text-slate-300">{ord.driver}</td>
                      <td className="py-3 px-4 text-slate-300">{ord.destination}</td>
                      <td className="py-3 px-4 font-mono text-slate-400">{ord.date}</td>
                      <td className="py-3 px-4">
                        <span className={`px-2.5 py-1 rounded-lg text-[10px] font-black ${
                          ord.status === 'Tamamlandı' ? 'bg-emerald-950 text-emerald-400 border border-emerald-800' : 'bg-amber-950 text-amber-400 border border-amber-800 animate-pulse'
                        }`}>
                          {ord.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
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

      {/* NEW VEHICLE MODAL */}
      {isNewVehicleOpen && (
        <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 border-2 border-emerald-500/60 p-6 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between mb-4 border-b border-slate-800 pb-3">
              <h3 className="text-sm font-black text-emerald-400 uppercase">
                Yeni Kara Aracı Ekle
              </h3>
              <button onClick={() => setIsNewVehicleOpen(false)} className="text-slate-400 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {KARA_ARACLARI_COLUMNS.slice(1).map((colName, cIdx) => {
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
                onClick={() => setIsNewVehicleOpen(false)}
                className="px-4 py-2 bg-slate-800 text-slate-300 rounded-xl text-xs font-bold"
              >
                Vazgeç
              </button>
              <button
                onClick={handleSaveNewVehicle}
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
              <h3 className="text-sm font-black text-blue-400 uppercase">
                Araç Kaydı Düzenle (Sıra {activeEditRow.idx + 1})
              </h3>
              <button onClick={() => setActiveEditRow(null)} className="text-slate-400 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {KARA_ARACLARI_COLUMNS.slice(1).map((colName, cIdx) => {
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
    </div>
  );
}

export default KaraAraclariApp;
