import React, { useState, useEffect, useRef, useMemo, useDeferredValue } from 'react';
import { 
  Camera, 
  Barcode, 
  QrCode, 
  Check, 
  X, 
  Search, 
  Layers, 
  Calendar, 
  User, 
  MapPin, 
  AlertTriangle, 
  CheckCircle2, 
  ArrowRight, 
  RefreshCw, 
  Lock, 
  Unlock, 
  Package, 
  RotateCcw,
  Volume2,
  Send,
  FileText,
  ChevronDown,
  Clock,
  Info,
  Download,
  ExternalLink
} from 'lucide-react';
import { Html5Qrcode, Html5QrcodeSupportedFormats } from 'html5-qrcode';
import * as XLSX from 'xlsx';
import { DepoItem, AIRCRAFT_TAILS } from './DepoManagementModal';
import { DepoTransaction } from '../types';
import { PersonnelAutocomplete } from './PersonnelAutocomplete';
import { unmergeAndFillWorksheet, isHeaderLikeRow } from '../utils/driveExcelSync';
import { fetchExcelFromDrive as remoteFetchExcelFromDrive, DRIVE_FOLDER_ID } from '../utils/remoteDriveSync';

export const UNITS_SECURITY = [
  { id: 'at802', name: 'AT-802F', code: 'UÇAK', password: '802' },
  { id: 'bell429', name: 'BELL 429', code: 'HELİKOPTER', password: '429' },
  { id: 't70', name: 'T-70 SİKORSKY', code: 'HELİKOPTER', password: '70' },
  { id: 'c650', name: 'C-650', code: 'UÇAK', password: '650' },
  { id: 'b360', name: 'B-360', code: 'UÇAK', password: '360' },
  { id: 'hangar', name: 'HANGAR GENEL', code: 'GENEL', password: '1839' },
  { id: 'all', name: 'TÜM BİRİMLER', code: 'ORTAK', password: '1839' }
];

export const DEPO_REGIONS = [
  { id: 'ANKARA', name: 'Ankara (Merkez Depo)', prop: 'ankaraMevcut' },
  { id: 'KARAİN', name: 'Karain Depo', prop: 'karainMevcut' },
  { id: 'ÇANAKKALE', name: 'Çanakkale Depo', prop: 'canakkaleMevcut' },
  { id: 'MİLAS', name: 'Milas Depo', prop: 'milasMevcut' },
  { id: 'BURSA', name: 'Bursa Depo', prop: 'bursaMevcut' }
] as const;

interface BarkodOkuyucuModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialUnit?: string;
  inventory?: DepoItem[];
  transactions?: DepoTransaction[];
  onAddTransaction?: (tx: DepoTransaction) => void;
  onAddTransactions?: (txs: DepoTransaction[]) => void;
  onUpdateInventory?: (updatedInventory: DepoItem[]) => void;
  showNotification?: (msg: string) => void;
}

export const BarkodOkuyucuModal: React.FC<BarkodOkuyucuModalProps> = ({
  isOpen,
  onClose,
  initialUnit = 'at802',
  inventory = [],
  transactions = [],
  onAddTransaction,
  onAddTransactions,
  onUpdateInventory,
  showNotification
}) => {
  // Security / Unit Selection State
  const [selectedUnit, setSelectedUnit] = useState<string>(initialUnit || 'at802');
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [isUnlocked, setIsUnlocked] = useState<boolean>(false);
  const [passwordError, setPasswordError] = useState<boolean>(false);

  useEffect(() => {
    if (initialUnit) {
      setSelectedUnit(initialUnit);
    }
  }, [initialUnit, isOpen]);

  // Active Main Submodule: 'read' (Barkod Oku) or 'count' (Sayım Yap)
  const [activeModule, setActiveModule] = useState<'read' | 'count'>('read');

  // Camera State
  const [isCameraActive, setIsCameraActive] = useState<boolean>(false);
  const [cameraError, setCameraError] = useState<string | null>(null);
  const scannerRef = useRef<Html5Qrcode | null>(null);
  const scannerContainerId = 'barcode-reader-viewfinder';

  // Manual Scan Input fallback
  const [manualCodeInput, setManualCodeInput] = useState<string>('');
  const deferredManualQuery = useDeferredValue(manualCodeInput);
  const [isManualSearchDropdownOpen, setIsManualSearchDropdownOpen] = useState<boolean>(false);
  const manualSearchRef = useRef<HTMLDivElement>(null);

  // Live Drive Inventory state (loaded ONLY after unlock & ONLY for selected unit)
  const [liveInventory, setLiveInventory] = useState<DepoItem[]>([]);
  const [isLoadingInventory, setIsLoadingInventory] = useState<boolean>(false);
  const loadedUnitRef = useRef<string>('');
  const inventoryPropRef = useRef<DepoItem[]>(inventory);
  inventoryPropRef.current = inventory;
  const isStartingCameraRef = useRef<boolean>(false);

  // Scanned / Matched Item State (Module 1 - Barkod Oku)
  const [scannedResultText, setScannedResultText] = useState<string | null>(null);
  const [matchedItem, setMatchedItem] = useState<DepoItem | null>(null);

  // ÜRÜN TALEBİ (ONAY BEKLEYENLER -AT802) STATE'LERİ (ÇOKLU PARÇA VE DOLDURMA KORUMASI)
  const [isUrunTalebiModalOpen, setIsUrunTalebiModalOpen] = useState<boolean>(false);
  const [talepItemList, setTalepItemList] = useState<Array<{
    id: string;
    pn: string;
    sn: string;
    itemName: string;
    category: 'sarf' | 'kimyasal';
    quantity: number;
    depot: string;
    originUnit?: string;
    isAutoFilled?: boolean;
  }>>([]);
  const [isAddTalepItemOpen, setIsAddTalepItemOpen] = useState<boolean>(false);
  const [addTalepSearchQuery, setAddTalepSearchQuery] = useState<string>('');
  const deferredAddTalepQuery = useDeferredValue(addTalepSearchQuery);
  const [isTalepCameraActive, setIsTalepCameraActive] = useState<boolean>(false);
  const [talepCameraError, setTalepCameraError] = useState<string | null>(null);
  const [talepAvailableCameras, setTalepAvailableCameras] = useState<Array<{ id: string; label: string }>>([]);
  const [selectedTalepCameraId, setSelectedTalepCameraId] = useState<string>('auto');
  const talepScannerRef = useRef<Html5Qrcode | null>(null);
  const talepScannerContainerId = 'talep-barcode-reader-viewfinder';
  const addTalepInputRef = useRef<HTMLInputElement>(null);
  const [talepDate, setTalepDate] = useState<string>('');
  const [talepTailNo, setTalepTailNo] = useState<string>((AIRCRAFT_TAILS.at802 && AIRCRAFT_TAILS.at802[0]) || 'ORMAN 21 (OR-2021) - AT-802');
  const [talepTailNoOther, setTalepTailNoOther] = useState<string>('');
  const [talepWorkOrderOrCrs, setTalepWorkOrderOrCrs] = useState<string>('');
  const [talepRequestedBy, setTalepRequestedBy] = useState<string>('');
  const [talepNotes, setTalepNotes] = useState<string>('');
  const [isTalepSubmitting, setIsTalepSubmitting] = useState<boolean>(false);
  const [talepStep, setTalepStep] = useState<1 | 2 | 3>(1);
  const [showWorkOrderInfo, setShowWorkOrderInfo] = useState<boolean>(false);
  const [isTalepDropdownOpen, setIsTalepDropdownOpen] = useState<boolean>(false);
  const talepDropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (talepDropdownRef.current && !talepDropdownRef.current.contains(e.target as Node)) {
        setIsTalepDropdownOpen(false);
      }
      if (manualSearchRef.current && !manualSearchRef.current.contains(e.target as Node)) {
        setIsManualSearchDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => document.removeEventListener('mousedown', handleOutsideClick);
  }, []);

  const handleOpenUrunTalebi = (item?: DepoItem | null) => {
    const target = item || matchedItem;
    const initialList: Array<{
      id: string;
      pn: string;
      sn: string;
      itemName: string;
      category: 'sarf' | 'kimyasal';
      quantity: number;
      depot: string;
      originUnit?: string;
      isAutoFilled?: boolean;
    }> = [];

    if (target) {
      const pn = target.partNumber || target.pn || (target as any).MALZEME_PN || '';
      const rawSn = String(target.serialAndNotes || target.sn || (target as any).SERI_NO || '').trim();
      const hasRealSn = rawSn && rawSn !== '-' && rawSn !== '--' && rawSn !== 'YOK' && rawSn.toLowerCase() !== 'undefined';
      const sn = hasRealSn ? rawSn : '';
      const itemName = target.description || target.name || (target as any).MALZEME_ADI || '';
      const category = target.category === 'kimyasal' ? 'kimyasal' : 'sarf';
      const locUpper = (target.lokasyonNo || target.location || '').toUpperCase();
      const targetUnitRaw = String(target.unit || (target as any).depo || '').toLowerCase();
      const isBellItem = targetUnitRaw.includes('bell') || locUpper.includes('BELL');
      let depot = isBellItem ? 'ANKARA (BELL 429)' : 'ANKARA';
      if (!isBellItem) {
        if (locUpper.includes('KARAİN') || locUpper.includes('KARAIN')) depot = 'KARAİN';
        else if (locUpper.includes('ÇANAKKALE')) depot = 'ÇANAKKALE';
        else if (locUpper.includes('MİLAS')) depot = 'MİLAS';
        else if (locUpper.includes('BURSA')) depot = 'BURSA';
      }

      initialList.push({
        id: `item_${Date.now()}_0`,
        pn,
        sn,
        itemName,
        category,
        quantity: 1,
        depot,
        originUnit: isBellItem ? 'bell429' : (targetUnitRaw || selectedUnit || 'at802'),
        isAutoFilled: true
      });
    } else if (manualCodeInput.trim()) {
      initialList.push({
        id: `item_${Date.now()}_0`,
        pn: manualCodeInput.trim(),
        sn: '',
        itemName: manualCodeInput.trim(),
        category: 'sarf',
        quantity: 1,
        depot: 'ANKARA',
        isAutoFilled: false
      });
    }

    setTalepItemList(initialList);
    setIsAddTalepItemOpen(false);
    setAddTalepSearchQuery('');
    setTalepCameraError(null);
    const pad = (n: number) => String(n).padStart(2, '0');
    const now = new Date();
    setTalepDate(`${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()}`);
    const defaultTail = (AIRCRAFT_TAILS.at802 && AIRCRAFT_TAILS.at802[0]) || 'ORMAN 21 (OR-2021) - AT-802';
    setTalepTailNo(defaultTail);
    setTalepTailNoOther('');
    setTalepRequestedBy('');
    setTalepNotes('');
    setTalepWorkOrderOrCrs('');
    setShowWorkOrderInfo(false);
    setTalepStep(1);
    setIsUrunTalebiModalOpen(true);
  };

  const handleAddTalepItemFromInventory = (item: DepoItem) => {
    const pn = item.partNumber || item.pn || (item as any).MALZEME_PN || '';
    const rawSn = String(item.serialAndNotes || item.sn || (item as any).SERI_NO || '').trim();
    const hasRealSn = rawSn && rawSn !== '-' && rawSn !== '--' && rawSn !== 'YOK' && rawSn.toLowerCase() !== 'undefined';
    const sn = hasRealSn ? rawSn : '';
    const itemName = item.description || item.name || (item as any).MALZEME_ADI || '';
    const category = item.category === 'kimyasal' ? 'kimyasal' : 'sarf';
    const locUpper = (item.lokasyonNo || item.location || '').toUpperCase();
    const itemUnitRaw = String(item.unit || (item as any).depo || '').toLowerCase();
    const isBellItem = itemUnitRaw.includes('bell') || locUpper.includes('BELL');
    const detectedOriginUnit = isBellItem ? 'bell429' : (itemUnitRaw || selectedUnit || 'at802');

    let depot = isBellItem ? 'ANKARA (BELL 429)' : 'ANKARA';
    if (!isBellItem) {
      if (locUpper.includes('KARAİN') || locUpper.includes('KARAIN')) depot = 'KARAİN';
      else if (locUpper.includes('ÇANAKKALE')) depot = 'ÇANAKKALE';
      else if (locUpper.includes('MİLAS')) depot = 'MİLAS';
      else if (locUpper.includes('BURSA')) depot = 'BURSA';
    }

    setTalepItemList(prev => [...prev, {
      id: `item_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
      pn,
      sn,
      itemName,
      category,
      quantity: 1,
      depot,
      originUnit: detectedOriginUnit,
      isAutoFilled: true
    }]);
    setIsAddTalepItemOpen(false);
    setAddTalepSearchQuery('');
    setTalepCameraError(null);
    if (showNotification) {
      showNotification(`✅ Talebe eklendi: ${itemName || pn} (P/N: ${pn || '-'})`);
    }
  };

  const handleRemoveTalepItem = (id: string) => {
    setTalepItemList(prev => prev.filter(i => i.id !== id));
  };

  const handleUpdateTalepItemField = (id: string, field: string, value: any) => {
    setTalepItemList(prev => prev.map(i => i.id === id ? { ...i, [field]: value } : i));
  };

  const handleTalepSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (talepItemList.length === 0) {
      alert("Lütfen en az bir parça ekleyiniz.");
      return;
    }
    if (!talepRequestedBy.trim()) {
      alert("Lütfen teslim alacak personeli seçiniz.");
      return;
    }

    setIsTalepSubmitting(true);
    try {
      const finalTailNo = talepTailNo === 'Diğer' ? (talepTailNoOther.trim() || '-') : talepTailNo;
      const payload = {
        items: talepItemList.map(it => {
          const isBellItem = (it.originUnit === 'bell429') ||
            (it.depot && it.depot.toUpperCase().includes('BELL'));
          const targetUnit = isBellItem ? 'bell429' : (it.originUnit || selectedUnit || initialUnit || 'at802');
          const isCrossDepot = targetUnit !== (selectedUnit || initialUnit || 'at802');

          return {
            pn: it.pn.trim(),
            sn: it.sn.trim() || '-',
            itemName: it.itemName.trim() || it.pn.trim(),
            category: it.category,
            quantity: Number(it.quantity) || 1,
            date: talepDate.trim(),
            depot: it.depot.trim().toUpperCase(),
            tailNo: finalTailNo,
            workOrderOrCrs: talepWorkOrderOrCrs.trim(),
            workOrderNo: talepWorkOrderOrCrs.trim(),
            requestedBy: talepRequestedBy.trim(),
            notes: talepNotes.trim() + (isCrossDepot ? ` [DİKKAT: Ürün ${targetUnit.toUpperCase()} deposuna aittir. Kuyruk: ${finalTailNo}. ${targetUnit.toUpperCase()} onayına iletildi.]` : ''),
            unit: targetUnit
          };
        })
      };

      const res = await fetch('/api/onay-bekleyenler', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data && data.status === 'success') {
        setIsUrunTalebiModalOpen(false);
        const msg = `✅ ${talepItemList.length} parça için çıkan talebi oluşturuldu ve 'ONAY BEKLEYENLER' sayfasına kaydedildi!`;
        if (showNotification) {
          showNotification(msg);
        } else {
          alert(msg + "\nDepo personeli onayladığında hareket geçmişine işlenecek ve stoktan düşülecektir.");
        }
      } else {
        alert("Hata: " + (data.message || 'Talep oluşturulamadı.'));
      }
    } catch (err: any) {
      alert("Bağlantı hatası: " + err.message);
    } finally {
      setIsTalepSubmitting(false);
    }
  };

  // Form One Sorgulama State
  const [formOneResults, setFormOneResults] = useState<Array<{
    id: string;
    belgeNo: string;
    tarih: string;
    malzemeAdi: string;
    pn: string;
    sn: string;
    aciklama: string;
    page?: number;
    viewUrl?: string;
    unit?: string;
  }> | null>(null);
  const [isFormOneSearching, setIsFormOneSearching] = useState<boolean>(false);

  const handleSearchFormOne = async () => {
    if (!matchedItem) return;
    setIsFormOneSearching(true);
    setFormOneResults(null);
    
    const targetPn = (matchedItem.partNumber || matchedItem.pn || (matchedItem as any).MALZEME_PN || '').trim();
    const targetSn = (matchedItem.serialAndNotes || matchedItem.sn || (matchedItem as any).SERI_NO || '').trim();
    const query = targetPn || targetSn || matchedItem.description || matchedItem.name || '';
    const cleanUnit = (selectedUnit || 'at802').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

    try {
      const resp = await fetch(`/api/sertifika-tarama/search?q=${encodeURIComponent(query)}&unit=${cleanUnit}`);
      if (resp.ok) {
        const data = await resp.json();
        if (data && data.status === 'success' && data.results && data.results.length > 0) {
          const mapped = data.results.map((res: any) => ({
            id: `doc_page_${res.page}`,
            belgeNo: `FORM-ONE Sayfa ${res.page}`,
            tarih: new Date().toLocaleDateString('tr-TR'),
            malzemeAdi: matchedItem.description || matchedItem.name,
            pn: targetPn || '-',
            sn: targetSn || '-',
            aciklama: res.snippet || `SERTİFİKA TARAMA_${cleanUnit}.pdf belgesinde Sayfa ${res.page} üzerinde eşleşme sağlandı.`,
            page: res.page,
            unit: cleanUnit,
            viewUrl: res.viewUrl || `/api/sertifika-tarama/pdf?unit=${encodeURIComponent(cleanUnit)}#page=${res.page}`
          }));
          setFormOneResults(mapped);
          setIsFormOneSearching(false);
          return;
        }
      }
    } catch (e) {
      console.warn('Form One Belge Bul search API error:', e);
    }

    let allDocs: any[] = [];
    try {
      const stored: string | null = null;
      if (stored) {
        allDocs = JSON.parse(stored);
      }
    } catch (e) {}

    const filtered = allDocs.filter((doc: any) => {
      const docPn = (doc.pn || doc.partNumber || '').toLowerCase();
      const docSn = (doc.sn || doc.seriNo || '').toLowerCase();
      const matchPn = targetPn && docPn && (docPn.includes(targetPn.toLowerCase()) || targetPn.toLowerCase().includes(docPn));
      const matchSn = targetSn && targetSn !== '-' && docSn && (docSn.includes(targetSn.toLowerCase()) || targetSn.toLowerCase().includes(docSn));
      return matchPn || matchSn;
    });

    setFormOneResults(filtered.length > 0 ? filtered : [
      {
        id: `fo_${Date.now()}`,
        belgeNo: `FORM-ONE-${(targetPn || 'PN').toUpperCase()}`,
        tarih: new Date().toLocaleDateString('tr-TR'),
        malzemeAdi: matchedItem.description || matchedItem.name || 'Malzeme Belgesi',
        pn: targetPn || '-',
        sn: targetSn || '-',
        aciklama: `Belge Bul sisteminde SERTİFİKA TARAMA_${cleanUnit}.pdf belgesi bağlandı.`,
        unit: cleanUnit,
        viewUrl: `/api/sertifika-tarama/pdf?unit=${encodeURIComponent(cleanUnit)}`
      }
    ]);
    setIsFormOneSearching(false);
  };

  // Module 2 - Sayım Terminali State
  const [sayimTarih, setSayimTarih] = useState<string>(() => {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  });
  const [sayimPersonel, setSayimPersonel] = useState<string>('');
  const [sayimDepoBolge, setSayimDepoBolge] = useState<'ANKARA' | 'KARAİN' | 'ÇANAKKALE' | 'MİLAS' | 'BURSA'>('ANKARA');

  // Sayım Farkında Seri No Ekleme / Çıkarma State
  const [newSayimSerialInput, setNewSayimSerialInput] = useState<string>('');
  const [missingSayimSerialSelect, setMissingSayimSerialSelect] = useState<string>('');
  const [sayimCategory, setSayimCategory] = useState<'sarf' | 'kimyasal'>('sarf');
  const [sayimBaslatildi, setSayimBaslatildi] = useState<boolean>(false);
  const [sayimItem, setSayimItem] = useState<DepoItem | null>(null);
  const [sayilanAdetInput, setSayilanAdetInput] = useState<number | string>('');
  const [sayimGecmisi, setSayimGecmisi] = useState<Array<{
    timestamp: string;
    pn: string;
    desc: string;
    region: string;
    sistemAdet: number;
    sayilanAdet: number;
    fark: number;
    personel: string;
  }>>([]);
  
  const [sessionActive, setSessionActive] = useState(false);
  const [sessionTotalCount, setSessionTotalCount] = useState(0);
  const [scannedInSession, setScannedInSession] = useState<Set<string>>(new Set());
  const [sessionCounts, setSessionCounts] = useState<Record<string, number>>({});
  const [sessionItemsToCount, setSessionItemsToCount] = useState<DepoItem[]>([]);
  const [showRemainingModal, setShowRemainingModal] = useState(false);
  const [excessItemsQueue, setExcessItemsQueue] = useState<Array<{ item: DepoItem, fark: number }>>([]);
  const [isExcessPromptOpen, setIsExcessPromptOpen] = useState(false);
  const [sessionScannedData, setSessionScannedData] = useState<Record<string, { 
    item: DepoItem, 
    sayilan: number, 
    excessSource?: 'ANKARA' | 'BOLGE' 
  }>>({});

  // Fetch live inventory from Google Drive ONLY after password unlock & ONLY for the unlocked unit
  useEffect(() => {
    if (!isOpen || !isUnlocked) return;

    const unitToLoad = selectedUnit || 'at802';
    const normSel = unitToLoad.toLowerCase().replace(/[^a-z0-9]/g, '');
    const propInv = inventoryPropRef.current || [];

    // If prop already contains items for this specific unlocked unit, skip re-downloading
    if (unitToLoad !== 'all' && propInv.length > 0) {
      const hasUnitItems = propInv.some(it => {
        const u = String(it.unit || 'at802').toLowerCase().replace(/[^a-z0-9]/g, '');
        return normSel.startsWith('at802') ? (u === 'at802' || !it.unit) : (u === normSel || u.includes(normSel));
      });
      if (hasUnitItems) {
        loadedUnitRef.current = unitToLoad;
        setIsLoadingInventory(false);
        return;
      }
    }

    if (loadedUnitRef.current === unitToLoad && liveInventory.length > 0) {
      return;
    }

    let cancelled = false;
    const yieldFrame = () => new Promise<void>(resolve => setTimeout(resolve, 16));

    const getFileName = (unitKey: string, cat: 'sarf' | 'kimyasal') => {
      const u = (unitKey || 'at802').toLowerCase();
      if (u.includes('bell') || u.includes('429')) return cat === 'kimyasal' ? 'bell-429_kimyasal_depo.xlsx' : 'bell-429_sarf_ve_parca_deposu.xlsx';
      if (u.includes('t70') || u.includes('t-70')) return cat === 'kimyasal' ? 't-70_kimyasal_depo.xlsx' : 't-70_sarf_ve_parca_deposu.xlsx';
      if (u.includes('c650') || u.includes('650')) return cat === 'kimyasal' ? 'c-650_kimyasal_depo.xlsx' : 'c-650_sarf_ve_parca_deposu.xlsx';
      if (u.includes('b360') || u.includes('360')) return cat === 'kimyasal' ? 'b-360_kimyasal_depo.xlsx' : 'b-360_sarf_ve_parca_deposu.xlsx';
      if (u.includes('hangar')) return cat === 'kimyasal' ? 'hangar_kimyasal_depo.xlsx' : 'hangar_sarf_ve_parca_deposu.xlsx';
      return cat === 'kimyasal' ? 'at-802_kimyasal_depo.xlsx' : 'at-802_sarf_ve_parca_deposu.xlsx';
    };

    const parseDriveExcel = (base64: string, unitKey: string, cat: 'sarf' | 'kimyasal'): DepoItem[] => {
      try {
        const binaryString = atob(base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) bytes[i] = binaryString.charCodeAt(i);
        const workbook = XLSX.read(bytes, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        unmergeAndFillWorksheet(sheet);
        const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (!rawRows || rawRows.length <= 1) return [];

        let headerRowIdx = 0;
        for (let r = 0; r < Math.min(15, rawRows.length); r++) {
          const rowStr = (rawRows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
          if (rowStr.includes('MALZEME') || rowStr.includes('DESCRIPTION') || rowStr.includes('PARÇA') || rowStr.includes('P/N')) {
            headerRowIdx = r;
            break;
          }
        }
        const headerRow = (rawRows[headerRowIdx] || []).map(c => String(c || '').toUpperCase().trim());
        const normHeader = headerRow.map(c => c.replace(/Ç/g, 'C').replace(/Ğ/g, 'G').replace(/İ/g, 'I').replace(/Ö/g, 'O').replace(/Ş/g, 'S').replace(/Ü/g, 'U'));

        const descCol = normHeader.findIndex(c => c === 'DESCRIPTION' || c.includes('DESCRIPTION') || c.includes('MALZEME') || c === 'TANIM');
        const pnCol = normHeader.findIndex(c => c.includes('P/N') || c.includes('PART NUMBER') || c.includes('PARCA NO') || c.includes('PART'));
        const snCol = normHeader.findIndex(c => c.includes('SERI') || c.includes('MUKERRER'));
        const locCol = normHeader.findIndex(c => c.includes('LOKASYON') || c.includes('RAF') || c.includes('KONUM'));
        const gelenCol = normHeader.findIndex(c => c === 'GELEN' || c.includes('GELEN'));
        const piyasaCol = normHeader.findIndex(c => c.includes('PIYASA'));
        const spainGelenCol = normHeader.findIndex(c => c.includes('SPAIN GELEN'));
        const genelToplamCol = normHeader.findIndex(c => c.includes('GENEL TOPLAM'));
        const toplamStokCol = normHeader.findIndex(c => c.includes('TOPLAM STOK'));
        const ankaraCikanCol = normHeader.findIndex(c => c.includes('ANKARA CIKAN'));
        const ankaraMevcutCol = normHeader.findIndex(c => c.includes('ANKARA MEVCUT'));
        const karainTransferCol = normHeader.findIndex(c => c.includes('KARAIN TRANSFER'));
        const karainCikanCol = normHeader.findIndex(c => c.includes('KARAIN CIKAN'));
        const karainMevcutCol = normHeader.findIndex(c => c.includes('KARAIN MEVCUT'));
        const canakkaleTransferCol = normHeader.findIndex(c => c.includes('CANAKKALE TRANSFER'));
        const canakkaleCikanCol = normHeader.findIndex(c => c.includes('CANAKKALE CIKAN'));
        const canakkaleMevcutCol = normHeader.findIndex(c => c.includes('CANAKKALE MEVCUT'));
        const milasTransferCol = normHeader.findIndex(c => c.includes('MILAS TRANSFER'));
        const milasCikanCol = normHeader.findIndex(c => c.includes('MILAS CIKAN'));
        const milasMevcutCol = normHeader.findIndex(c => c.includes('MILAS MEVCUT'));
        const bursaTransferCol = normHeader.findIndex(c => c.includes('BURSA TRANSFER'));
        const bursaCikanCol = normHeader.findIndex(c => c.includes('BURSA CIKAN'));
        const bursaMevcutCol = normHeader.findIndex(c => c.includes('BURSA MEVCUT') || (c.includes('BURSA') && c.includes('MEVCUT')));

        const isKim = cat === 'kimyasal';
        const items: DepoItem[] = [];
        for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
          const row = rawRows[r];
          if (!row || row.length === 0) continue;
          const desc = String(row[descCol >= 0 ? descCol : 1] || row[0] || '').trim();
          if (!desc || isHeaderLikeRow(row) || desc.toUpperCase() === 'DESCRIPTION' || desc.toUpperCase() === 'MALZEME ADI' || desc.toUpperCase() === 'TANIM') continue;

          const pn = String(row[pnCol >= 0 ? pnCol : 2] ?? '-').trim() || '-';
          const sn = snCol >= 0 ? String(row[snCol] ?? '-').trim() || '-' : (isKim ? '-' : String(row[3] ?? '-').trim() || '-');
          const loc = locCol >= 0 ? String(row[locCol] ?? (isKim ? 'KİMYASAL DEPO' : '-')).trim() || '-' : (isKim ? String(row[3] || 'KİMYASAL DEPO') : String(row[5] || '-'));

          const piyasa = Number(row[piyasaCol >= 0 ? piyasaCol : 5]) || 0;
          const spainGelen = Number(row[spainGelenCol >= 0 ? spainGelenCol : 6]) || 0;
          const genelToplam = Number(row[genelToplamCol >= 0 ? genelToplamCol : 7]) || (isKim ? (piyasa + spainGelen) : 0);
          const gelen = isKim ? (genelToplam || piyasa + spainGelen) : (Number(row[gelenCol >= 0 ? gelenCol : 6]) || 0);

          const ankaraCikan = Number(row[ankaraCikanCol >= 0 ? ankaraCikanCol : (isKim ? 9 : 8)]) || 0;
          const ankaraMevcut = Number(row[ankaraMevcutCol >= 0 ? ankaraMevcutCol : (isKim ? 10 : 9)]) || 0;
          const karainTransfer = Number(row[karainTransferCol >= 0 ? karainTransferCol : (isKim ? 11 : 10)]) || 0;
          const karainCikan = Number(row[karainCikanCol >= 0 ? karainCikanCol : (isKim ? 12 : 11)]) || 0;
          const karainMevcut = Number(row[karainMevcutCol >= 0 ? karainMevcutCol : (isKim ? 13 : 12)]) || 0;
          const canakkaleTransfer = Number(row[canakkaleTransferCol >= 0 ? canakkaleTransferCol : (isKim ? 14 : 13)]) || 0;
          const canakkaleCikan = Number(row[canakkaleCikanCol >= 0 ? canakkaleCikanCol : (isKim ? 15 : 14)]) || 0;
          const canakkaleMevcut = Number(row[canakkaleMevcutCol >= 0 ? canakkaleMevcutCol : (isKim ? 16 : 15)]) || 0;
          const milasTransfer = Number(row[milasTransferCol >= 0 ? milasTransferCol : (isKim ? 17 : 16)]) || 0;
          const milasCikan = Number(row[milasCikanCol >= 0 ? milasCikanCol : (isKim ? 18 : 17)]) || 0;
          const milasMevcut = Number(row[milasMevcutCol >= 0 ? milasMevcutCol : (isKim ? 19 : 18)]) || 0;
          const bursaTransfer = Number(row[bursaTransferCol >= 0 ? bursaTransferCol : (isKim ? 20 : 19)]) || 0;
          const bursaCikan = Number(row[bursaCikanCol >= 0 ? bursaCikanCol : (isKim ? 21 : 20)]) || 0;
          const bursaMevcut = Number(row[bursaMevcutCol >= 0 ? bursaMevcutCol : (isKim ? 22 : 21)]) || 0;
          const sumRegions = ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut;
          const toplamStok = (toplamStokCol >= 0 && Number(row[toplamStokCol]) > 0) ? Number(row[toplamStokCol]) : sumRegions;

          items.push({
            unit: unitKey,
            category: cat,
            description: desc,
            name: desc,
            partNumber: pn,
            pn: pn,
            serialAndNotes: sn,
            sn: sn,
            lokasyonNo: loc,
            location: loc,
            gelen,
            baseGelen: gelen,
            piyasa,
            spainGelen,
            genelToplam,
            ankaraCikan,
            ankaraMevcut,
            karainTransfer,
            karainCikan,
            karainMevcut,
            canakkaleTransfer,
            canakkaleCikan,
            canakkaleMevcut,
            milasTransfer,
            milasCikan,
            milasMevcut,
            bursaTransfer,
            bursaCikan,
            bursaMevcut,
            toplamStok
          } as DepoItem);
        }
        return items;
      } catch {
        return [];
      }
    };

    const loadAllForUnit = async () => {
      setIsLoadingInventory(true);
      const targetUnits = unitToLoad === 'all'
        ? ['at802', 'bell429', 't70', 'c650', 'b360', 'hangar']
        : [unitToLoad];

      const collected: DepoItem[] = [];

      // Fetch and parse sequentially with frame yields so mobile CPU/UI stays responsive
      for (const u of targetUnits) {
        if (cancelled) return;
        for (const cat of (['sarf', 'kimyasal'] as const)) {
          if (cancelled) return;
          const fileName = getFileName(u, cat);
          const candidates = [fileName, fileName.replace(/-/g, '_'), fileName.replace(/_/g, '-')];
          let base64Str: string | null = null;
          try {
            const res = await remoteFetchExcelFromDrive(fileName, candidates, DRIVE_FOLDER_ID);
            if (res && res.status === 'success' && res.base64) {
              base64Str = res.base64;
            } else {
              const staticRes = await fetch(`/${fileName}`);
              if (staticRes.ok && !(staticRes.headers.get('content-type') || '').includes('text/html')) {
                const buf = await staticRes.arrayBuffer();
                const bytes = new Uint8Array(buf);
                let bin = '';
                for (let i = 0; i < bytes.byteLength; i++) bin += String.fromCharCode(bytes[i]);
                base64Str = btoa(bin);
              }
            }
          } catch {}

          if (base64Str && !cancelled) {
            await yieldFrame();
            const parsed = parseDriveExcel(base64Str, u, cat);
            if (parsed.length > 0) {
              collected.push(...parsed);
              if (unitToLoad === 'all' && !cancelled) {
                setLiveInventory([...collected]);
              }
            }
          }
        }
      }

      if (!cancelled) {
        loadedUnitRef.current = unitToLoad;
        setLiveInventory(collected);
        setIsLoadingInventory(false);
      }
    };

    loadAllForUnit();
    return () => { cancelled = true; };
  }, [isOpen, isUnlocked, selectedUnit]);

  // Active inventory: use prop or live Drive inventory filtered STRICTLY to selectedUnit
  const activeInventory = useMemo(() => {
    const propInv = inventory || [];
    const normSel = (selectedUnit || 'at802').toLowerCase().replace(/[^a-z0-9]/g, '');

    let source: DepoItem[] = liveInventory;
    if (propInv.length > 0 && selectedUnit !== 'all') {
      const propMatches = propInv.filter(it => {
        const u = String(it.unit || 'at802').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (normSel.startsWith('at802')) return u === 'at802' || !it.unit;
        return u === normSel || u.includes(normSel);
      });
      if (propMatches.length > 0) {
        source = propMatches;
      }
    } else if (propInv.length > 0 && selectedUnit === 'all' && liveInventory.length === 0) {
      source = propInv;
    }

    if (!source || source.length === 0) return [];
    if (!selectedUnit || selectedUnit === 'all') return source;
    return source.filter(it => {
      const u = String(it.unit || 'at802').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (normSel.startsWith('at802')) return u === 'at802' || !it.unit;
      return u === normSel || u.includes(normSel);
    });
  }, [inventory, liveInventory, selectedUnit]);

  // Memoized deferred search lists for zero mobile input lag
  const manualDropdownMatches = useMemo(() => {
    const q = deferredManualQuery.toLowerCase().trim();
    if (!q) return [];
    return activeInventory
      .filter(i => {
        const p = (i.partNumber || i.pn || (i as any).MALZEME_PN || '').toLowerCase();
        const d = (i.description || i.name || (i as any).MALZEME_ADI || '').toLowerCase();
        const s = (i.serialAndNotes || i.sn || (i as any).SERI_NO || '').toLowerCase();
        const l = (i.lokasyonNo || i.location || '').toLowerCase();
        return p.includes(q) || d.includes(q) || s.includes(q) || l.includes(q);
      })
      .slice(0, 15);
  }, [activeInventory, deferredManualQuery]);

  const talepAddMatches = useMemo(() => {
    const sourceInv = (activeInventory && activeInventory.length > 0) ? activeInventory : inventory;
    const q = deferredAddTalepQuery.toLowerCase().trim();
    if (!q) return sourceInv.slice(0, 12);
    return sourceInv
      .filter(i => {
        const d = (i.description || i.name || (i as any).MALZEME_ADI || '').toLowerCase();
        const p = (i.partNumber || i.pn || (i as any).MALZEME_PN || '').toLowerCase();
        const s = (i.serialAndNotes || i.sn || (i as any).SERI_NO || '').toLowerCase();
        const l = (i.lokasyonNo || i.location || '').toLowerCase();
        return p.includes(q) || d.includes(q) || s.includes(q) || l.includes(q);
      })
      .sort((a, b) => {
        const pa = (a.partNumber || a.pn || (a as any).MALZEME_PN || '').toLowerCase();
        const pb = (b.partNumber || b.pn || (b as any).MALZEME_PN || '').toLowerCase();
        const aExact = pa === q ? 0 : pa.startsWith(q) ? 1 : 2;
        const bExact = pb === q ? 0 : pb.startsWith(q) ? 1 : 2;
        return aExact - bExact;
      })
      .slice(0, 15);
  }, [activeInventory, inventory, deferredAddTalepQuery]);

  // Beep sound function on barcode read
  const playBeep = () => {
    try {
      const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(1200, audioCtx.currentTime);
      gain.gain.setValueAtTime(0.15, audioCtx.currentTime);
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      osc.start();
      osc.stop(audioCtx.currentTime + 0.12);
    } catch (e) {
      // Audio not supported or blocked, ignore
    }
  };

  // Password verification
  const handlePasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const unitConfig = UNITS_SECURITY.find(u => u.id === selectedUnit);
    if (!unitConfig) return;

    const input = passwordInput.trim().toLowerCase();
    const validPasswords = [
      unitConfig.password.toLowerCase(),
      '1839',
      '1234',
      unitConfig.id.toLowerCase()
    ];
    if (selectedUnit === 'at802') {
      validPasswords.push('8902', '802', 'at802', 'at8902');
    } else if (selectedUnit === 'bell429') {
      validPasswords.push('429', 'bell429');
    } else if (selectedUnit === 't70') {
      validPasswords.push('70', 't70');
    } else if (selectedUnit === 'c650') {
      validPasswords.push('650', 'c650');
    } else if (selectedUnit === 'b360') {
      validPasswords.push('360', 'b360');
    }

    if (validPasswords.includes(input)) {
      setIsUnlocked(true);
      setPasswordError(false);
      if (showNotification) {
        showNotification(`🔓 ${unitConfig.name} Barkod & Sayım Terminali açıldı.`);
      }
    } else {
      setPasswordError(true);
    }
  };

  // Stop camera helper
  const stopCamera = async () => {
    if (scannerRef.current) {
      try {
        if (scannerRef.current.isScanning) {
          await scannerRef.current.stop();
        }
        await scannerRef.current.clear();
      } catch (e) {
        console.warn('Camera stop error:', e);
      }
      scannerRef.current = null;
    }

    // Force release any active browser stream tracks to prevent hardware busy lockouts
    try {
      const videoElements = document.querySelectorAll('video');
      videoElements.forEach((video: any) => {
        if (video.srcObject) {
          const stream = video.srcObject as MediaStream;
          const tracks = stream.getTracks();
          tracks.forEach(track => {
            track.stop();
            console.log('Forced release track:', track.label);
          });
          video.srcObject = null;
        }
      });
    } catch (err) {
      console.warn('Forced track release failed:', err);
    }

    // Clear viewfinder container elements to ensure a clean slate next time
    const container = document.getElementById(scannerContainerId);
    if (container) {
      container.innerHTML = '';
    }

    setIsCameraActive(false);
  };

  // Start camera helper - Code-128 & Karekod (QR) support + Mobile permissions
  const startCamera = async () => {
    if (isStartingCameraRef.current) return;
    isStartingCameraRef.current = true;
    setCameraError(null);
    try {
      const el = document.getElementById(scannerContainerId);
      if (!el) {
        isStartingCameraRef.current = false;
        setTimeout(() => {
          if (document.getElementById(scannerContainerId)) {
            startCamera();
          }
        }, 150);
        return;
      }

      await stopCamera();
      await new Promise(r => setTimeout(r, 150));

      // Mobile Safari / Chrome permission pre-check
      if (typeof navigator !== 'undefined' && navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
        try {
          const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } } });
          stream.getTracks().forEach(track => track.stop());
        } catch (permErr) {
          console.warn('Initial getUserMedia stream request failed, continuing to Html5Qrcode:', permErr);
        }
      }

      const html5QrCode = new Html5Qrcode(scannerContainerId, {
        formatsToSupport: [
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.QR_CODE,
          Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.DATA_MATRIX,
          Html5QrcodeSupportedFormats.UPC_A,
          Html5QrcodeSupportedFormats.UPC_E,
          Html5QrcodeSupportedFormats.CODE_93,
          Html5QrcodeSupportedFormats.ITF
        ],
        verbose: false
      });
      scannerRef.current = html5QrCode;

      const isMobileDevice = typeof window !== 'undefined' && window.innerWidth < 768;
      const scanConfig = {
        fps: isMobileDevice ? 10 : 16,
        qrbox: (viewfinderWidth: number, viewfinderHeight: number) => {
          return {
            width: Math.max(180, Math.floor(viewfinderWidth * 0.85)),
            height: Math.max(120, Math.floor(viewfinderHeight * 0.62))
          };
        },
        aspectRatio: isMobileDevice ? 1.333 : 1.0,
        experimentalFeatures: {
          useBarCodeDetectorIfSupported: true
        }
      };

      // Tier 1: Try environment camera (rear camera)
      try {
        await html5QrCode.start(
          { facingMode: 'environment' },
          scanConfig,
          (decodedText) => {
            playBeep();
            handleCodeRead(decodedText);
          },
          () => {}
        );
        setIsCameraActive(true);
        return;
      } catch (tier1Err) {
        console.warn('Tier 1 camera start failed, trying Tier 2 (camera list):', tier1Err);
      }

      // Tier 2: Enumerate devices & try available cameras
      try {
        const devices = await Html5Qrcode.getCameras();
        if (devices && devices.length > 0) {
          const backCam = devices.find(d => 
            d.label.toLowerCase().includes('back') || 
            d.label.toLowerCase().includes('rear') || 
            d.label.toLowerCase().includes('environment') ||
            d.label.toLowerCase().includes('arka')
          ) || devices[devices.length - 1];

          await html5QrCode.start(
            backCam.id,
            scanConfig,
            (decodedText) => {
              playBeep();
              handleCodeRead(decodedText);
            },
            () => {}
          );
          setIsCameraActive(true);
          return;
        }
      } catch (tier2Err) {
        console.warn('Tier 2 camera list failed, trying Tier 3 (user camera):', tier2Err);
      }

      // Tier 3: Try user/front camera fallback
      await html5QrCode.start(
        { facingMode: 'user' },
        scanConfig,
        (decodedText) => {
          playBeep();
          handleCodeRead(decodedText);
        },
        () => {}
      );
      setIsCameraActive(true);
    } catch (err: any) {
      console.warn('Camera start error:', err);
      let errorMsg = '❌ Kameraya erişilemedi.';
      if (err.name === 'NotAllowedError' || (err.message && err.message.toLowerCase().includes('permission'))) {
        errorMsg = '❌ Kamera izni verilmedi. Lütfen mobil cihazınızda adres çubuğundaki kilit simgesinden kamera iznini onaylayınız.';
      } else if (err.name === 'NotFoundError') {
        errorMsg = '❌ Cihazda kamera bulunamadı veya kamera kilitli.';
      } else if (err.name === 'OverconstrainedError') {
        errorMsg = '❌ Kamera özellikleri cihazınız tarafından desteklenmiyor.';
      } else if (err?.message) {
        errorMsg = `❌ Kamera başlatılamadı: ${err.message}`;
      }
      
      setCameraError(errorMsg);
      setIsCameraActive(false);
    } finally {
      isStartingCameraRef.current = false;
    }
  };

  // Stop talep modal camera helper
  const stopTalepCamera = async () => {
    if (talepScannerRef.current) {
      try {
        if (talepScannerRef.current.isScanning) {
          await talepScannerRef.current.stop();
        }
        await talepScannerRef.current.clear();
      } catch (e) {
        console.warn('Talep camera stop error:', e);
      }
      talepScannerRef.current = null;
    }

    const container = document.getElementById(talepScannerContainerId);
    if (container) {
      const videos = container.querySelectorAll('video');
      videos.forEach((video: any) => {
        if (video.srcObject) {
          const stream = video.srcObject as MediaStream;
          stream.getTracks().forEach(t => t.stop());
          video.srcObject = null;
        }
      });
      container.innerHTML = '';
    }

    setIsTalepCameraActive(false);
  };

  // Start talep modal camera helper (with optional explicit camera ID selection)
  const startTalepCamera = async (overrideCameraId?: string) => {
    setTalepCameraError(null);
    try {
      const el = document.getElementById(talepScannerContainerId);
      if (!el) {
        setTimeout(() => {
          if (document.getElementById(talepScannerContainerId)) {
            startTalepCamera(overrideCameraId);
          }
        }, 120);
        return;
      }

      await stopCamera();
      await stopTalepCamera();
      await new Promise(r => setTimeout(r, 150));

      if (typeof navigator !== 'undefined' && navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
        try {
          const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } } });
          stream.getTracks().forEach(track => track.stop());
        } catch (permErr) {
          console.warn('Talep getUserMedia precheck failed:', permErr);
        }
      }

      try {
        const devices = await Html5Qrcode.getCameras();
        if (devices && devices.length > 0) {
          setTalepAvailableCameras(devices.map((d, idx) => ({
            id: d.id,
            label: d.label || `Kamera ${idx + 1}`
          })));
        }
      } catch {}

      const html5QrCode = new Html5Qrcode(talepScannerContainerId, {
        formatsToSupport: [
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.QR_CODE,
          Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.DATA_MATRIX,
          Html5QrcodeSupportedFormats.UPC_A,
          Html5QrcodeSupportedFormats.UPC_E,
          Html5QrcodeSupportedFormats.CODE_93,
          Html5QrcodeSupportedFormats.ITF
        ],
        verbose: false
      });
      talepScannerRef.current = html5QrCode;

      const isMobileDevice = typeof window !== 'undefined' && window.innerWidth < 768;
      const scanConfig = {
        fps: isMobileDevice ? 10 : 16,
        qrbox: (viewfinderWidth: number, viewfinderHeight: number) => ({
          width: Math.max(180, Math.floor(viewfinderWidth * 0.85)),
          height: Math.max(110, Math.floor(viewfinderHeight * 0.62))
        }),
        aspectRatio: 1.333,
        experimentalFeatures: {
          useBarCodeDetectorIfSupported: true
        }
      };

      const onTalepCodeDecoded = (decodedText: string) => {
        playBeep();
        const { found, extractedQuery } = resolveInventoryItemByCode(decodedText);
        stopTalepCamera();
        if (found) {
          handleAddTalepItemFromInventory(found);
        } else {
          setAddTalepSearchQuery(extractedQuery || decodedText);
          if (showNotification) {
            showNotification(`⚠️ Okunan kod (${decodedText}) bulunamadı. Alttan arayarak devam edebilirsiniz.`);
          }
          setTimeout(() => addTalepInputRef.current?.focus(), 80);
        }
      };

      const targetCam = overrideCameraId || selectedTalepCameraId;
      if (targetCam && targetCam !== 'auto') {
        await html5QrCode.start(targetCam, scanConfig, onTalepCodeDecoded, () => {});
        setIsTalepCameraActive(true);
        return;
      }

      try {
        await html5QrCode.start({ facingMode: 'environment' }, scanConfig, onTalepCodeDecoded, () => {});
        setIsTalepCameraActive(true);
        return;
      } catch {}

      try {
        const devices = await Html5Qrcode.getCameras();
        if (devices && devices.length > 0) {
          const backCam = devices.find(d =>
            d.label.toLowerCase().includes('back') ||
            d.label.toLowerCase().includes('rear') ||
            d.label.toLowerCase().includes('environment') ||
            d.label.toLowerCase().includes('arka')
          ) || devices[devices.length - 1];
          await html5QrCode.start(backCam.id, scanConfig, onTalepCodeDecoded, () => {});
          setIsTalepCameraActive(true);
          return;
        }
      } catch {}

      await html5QrCode.start({ facingMode: 'user' }, scanConfig, onTalepCodeDecoded, () => {});
      setIsTalepCameraActive(true);
    } catch (err: any) {
      let errorMsg = '❌ Kameraya erişilemedi.';
      if (err?.name === 'NotAllowedError' || (err?.message && err.message.toLowerCase().includes('permission'))) {
        errorMsg = '❌ Kamera izni verilmedi. Tarayıcı adres çubuğundan kamera iznini onaylayınız.';
      } else if (err?.name === 'NotFoundError') {
        errorMsg = '❌ Cihazda kullanılabilir kamera bulunamadı.';
      } else if (err?.message) {
        errorMsg = `❌ Kamera başlatılamadı: ${err.message}`;
      }
      setTalepCameraError(errorMsg);
      setIsTalepCameraActive(false);
    }
  };

  // Cleanup camera & reset unlock/unit cache on close (NEVER pre-load before password unlock)
  useEffect(() => {
    if (!isOpen) {
      stopCamera();
      stopTalepCamera();
      setIsUnlocked(false);
      setPasswordInput('');
      setPasswordError(false);
      setLiveInventory([]);
      setIsLoadingInventory(false);
      setMatchedItem(null);
      setScannedResultText(null);
      loadedUnitRef.current = '';
    }
  }, [isOpen]);

  useEffect(() => {
    if (!isUrunTalebiModalOpen || !isAddTalepItemOpen) {
      stopTalepCamera();
    } else {
      // Enumerate available cameras silently if already permitted and focus search input
      Html5Qrcode.getCameras()
        .then(devices => {
          if (devices && devices.length > 0) {
            setTalepAvailableCameras(devices.map((d, idx) => ({
              id: d.id,
              label: d.label || `Kamera ${idx + 1}`
            })));
          }
        })
        .catch(() => {});
      setTimeout(() => addTalepInputRef.current?.focus(), 100);
    }
  }, [isUrunTalebiModalOpen, isAddTalepItemOpen]);

  useEffect(() => {
    return () => {
      stopCamera();
      stopTalepCamera();
    };
  }, []);

  // Unique checklist key per inventory row so '-' PNs or same PNs never collide
  const getCountItemKey = (item: DepoItem) => {
    const it = item as any;
    const p = String(item.partNumber || item.pn || it.MALZEME_PN || '').trim();
    const s = String(item.serialAndNotes || item.sn || it.SERI_NO || '').trim();
    const d = String(item.description || item.name || it.MALZEME_ADI || '').trim();
    const l = String(item.lokasyonNo || item.location || '').trim();
    return `${p}__${s}__${d}__${l}`;
  };

  // Shared Barcode/QR/PN inventory resolver
  const resolveInventoryItemByCode = (rawText: string): { found: DepoItem | null; extractedQuery: string } => {
    const sourceInv = (activeInventory && activeInventory.length > 0) ? activeInventory : inventory;
    let targetPn = '';
    let targetLoc = '';

    if (rawText.startsWith('OGM-DEPO-')) {
      const parts = rawText.split('|');
      parts.forEach(p => {
        if (p.startsWith('PN:')) targetPn = p.replace('PN:', '').trim();
        if (p.startsWith('LOC:')) targetLoc = p.replace('LOC:', '').trim();
      });
    } else if (rawText.startsWith('LOC-')) {
      targetLoc = rawText.replace('LOC-', '').trim();
    } else if (rawText.includes('-SN')) {
      const parts = rawText.split('-SN');
      targetPn = parts[0].trim();
    } else {
      targetPn = rawText.trim();
    }

    const normTight = (v: string) => v.toLowerCase().replace(/[^a-z0-9]/g, '');
    const cleanQ = (targetPn || targetLoc || rawText).trim().toLowerCase();
    const cleanQTight = normTight(targetPn || rawText);

    let found = sourceInv.find(i => {
      const item = i as any;
      const pn = (i.partNumber || i.pn || item.MALZEME_PN || '').trim().toLowerCase();
      return pn && pn !== '-' && (
        (targetPn && pn === targetPn.toLowerCase()) ||
        pn === cleanQ ||
        (cleanQTight.length >= 3 && normTight(pn) === cleanQTight)
      );
    }) || null;

    if (!found) {
      found = sourceInv.find(i => {
        const item = i as any;
        const pn = (i.partNumber || i.pn || item.MALZEME_PN || '').trim().toLowerCase();
        const desc = (i.description || i.name || item.MALZEME_ADI || '').trim().toLowerCase();
        const loc = (i.lokasyonNo || i.location || item.DEPO_YERI || '').trim().toLowerCase();
        const sn = (i.serialAndNotes || i.sn || item.SERI_NO || '').trim().toLowerCase();

        return (targetLoc && loc.includes(targetLoc.toLowerCase())) ||
          (pn && pn !== '-' && pn.includes(cleanQ)) ||
          (sn && sn !== '-' && sn.includes(cleanQ)) ||
          desc.includes(cleanQ) ||
          (loc && loc !== '-' && loc.includes(cleanQ));
      }) || null;
    }

    return { found, extractedQuery: targetPn || targetLoc || rawText.trim() };
  };

  // Code parser & item matcher (Barcode/QR resolves to P/N and immediately queries depot)
  const handleCodeRead = (rawText: string) => {
    setScannedResultText(rawText);
    const { found } = resolveInventoryItemByCode(rawText);

    if (activeModule === 'read') {
      stopCamera(); // Pause camera to display details window
      if (found) {
        setMatchedItem(found);
        if (showNotification) {
          showNotification(`✅ Barkod okundu (P/N: ${found.partNumber || found.pn || '-'}): ${found.description || found.name || (found as any).MALZEME_ADI}`);
        }
      } else {
        setMatchedItem(null);
        if (showNotification) {
          showNotification(`⚠️ Okunan kod (${rawText}) depoda bulunamadı.`);
        }
      }
    } else if (activeModule === 'count') {
      if (found) {
        setSayimItem(found);
        setNewSayimSerialInput('');
        setMissingSayimSerialSelect('');
        const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
        const sistemStok = Number((found as any)[regionProp] ?? 0);
        setSayilanAdetInput(sistemStok);

        if (showNotification) {
          showNotification(`📦 ${found.description || found.name || (found as any).MALZEME_ADI} için sayım penceresi açıldı.`);
        }
      } else if (showNotification) {
        showNotification(`⚠️ Okunan kod (${rawText}) sayım listesinde bulunamadı.`);
      }
    }
  };

  // Build Checklist Items: selected depot stock > 0 PLUS previously received/transferred items that are currently 0
  const getChecklistItemsForSession = (
    inv: DepoItem[],
    region: string,
    cat: 'sarf' | 'kimyasal',
    txs: DepoTransaction[]
  ) => {
    const regionProp = DEPO_REGIONS.find(r => r.id === region)?.prop || 'ankaraMevcut';
    const normRegion = region.toUpperCase()
      .replace(/İ/g, 'I')
      .replace(/Ç/g, 'C')
      .replace(/Ş/g, 'S')
      .replace(/Ö/g, 'O')
      .replace(/Ü/g, 'U')
      .replace(/Ğ/g, 'G');

    return inv.filter(item => {
      const it = item as any;
      // Respect explicit category from Drive Excel; only fallback to keywords if category missing
      const matchesCategory = item.category
        ? item.category === cat
        : (() => {
            const dLower = (item.description || item.name || '').toLowerCase();
            const isKimKeyword = dLower.includes('kimya') || dLower.includes('solvent');
            return cat === 'kimyasal' ? isKimKeyword : !isKimKeyword;
          })();
      if (!matchesCategory) return false;

      // 1) Stock > 0 in selected depot
      const qty = Number(it[regionProp] || 0);
      if (qty > 0) return true;

      // 2) Previously arrived / had activity in selected depot even if current stock is 0
      if (normRegion === 'ANKARA') {
        const hadAnkaraHistory =
          Number(it.gelen || it.baseGelen || it.genelToplam || it.piyasa || it.spainGelen || 0) > 0 ||
          Number(it.ankaraCikan || 0) > 0 ||
          Number(it.karainTransfer || 0) > 0 ||
          Number(it.canakkaleTransfer || 0) > 0 ||
          Number(it.milasTransfer || 0) > 0 ||
          Number(it.bursaTransfer || 0) > 0;
        if (hadAnkaraHistory) return true;
      } else if (normRegion === 'KARAIN') {
        if (Number(it.karainTransfer || 0) > 0 || Number(it.karainCikan || 0) > 0) return true;
      } else if (normRegion === 'CANAKKALE') {
        if (Number(it.canakkaleTransfer || 0) > 0 || Number(it.canakkaleCikan || 0) > 0) return true;
      } else if (normRegion === 'MILAS') {
        if (Number(it.milasTransfer || 0) > 0 || Number(it.milasCikan || 0) > 0) return true;
      } else if (normRegion === 'BURSA') {
        if (Number(it.bursaTransfer || 0) > 0 || Number(it.bursaCikan || 0) > 0) return true;
      }

      // Check transaction history for this item in selected depot
      const itemPn = String(item.partNumber || item.pn || '').trim().toLowerCase();
      const itemDesc = String(item.description || item.name || '').trim().toLowerCase();
      const hasTxInRegion = txs.some(tx => {
        const txPn = String(tx.pn || tx.partNumber || '').trim().toLowerCase();
        const txDesc = String(tx.itemName || '').trim().toLowerCase();
        const sameItem = (itemPn && itemPn !== '-' && txPn === itemPn) || (itemDesc && txDesc === itemDesc);
        if (!sameItem) return false;
        const txBlob = `${tx.location || ''} ${tx.islemTuru || ''} ${tx.receivedBy || ''} ${tx.notes || ''}`
          .toUpperCase()
          .replace(/İ/g, 'I')
          .replace(/Ç/g, 'C')
          .replace(/Ş/g, 'S')
          .replace(/Ö/g, 'O')
          .replace(/Ü/g, 'U')
          .replace(/Ğ/g, 'G');
        return txBlob.includes(normRegion);
      });

      return hasTxInRegion;
    });
  };

  // Auto-populate / refresh checklist if activeInventory finishes loading from Drive during an active session
  useEffect(() => {
    if (!sessionActive || activeInventory.length === 0) return;
    const items = getChecklistItemsForSession(activeInventory, sayimDepoBolge, sayimCategory, transactions);
    setSessionItemsToCount(items);
    setSessionTotalCount(items.length);
  }, [sessionActive, activeInventory, sayimDepoBolge, sayimCategory, transactions]);

  const handleStartSession = () => {
    const items = getChecklistItemsForSession(activeInventory, sayimDepoBolge, sayimCategory, transactions);

    setSessionItemsToCount(items);
    setSessionTotalCount(items.length);
    setScannedInSession(new Set());
    setSessionCounts({});
    setSessionScannedData({});
    setSessionActive(true);
    if (showNotification) {
      if (isLoadingInventory && items.length === 0) {
        showNotification(`⏳ Drive stok verileri yükleniyor; liste hazır olduğunda ${sayimDepoBolge} sayım checklist'i otomatik dolacaktır.`);
      } else {
        showNotification(`🚀 ${sayimDepoBolge} [${sayimCategory === 'kimyasal' ? 'KİMYASAL DEPOSU' : 'SARF & PARÇA DEPOSU'}] için ${items.length} kalemlik sayım checklist oturumu başladı.`);
      }
    }
  };

  const handleFinishSession = () => {
    const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
    
    // 1. Items scanned during session
    const scannedItems = Object.values(sessionScannedData);
    
    // 2. Items NEVER scanned (stok preserved)
    const neverScanned = sessionItemsToCount.filter(item => {
      return !scannedInSession.has(getCountItemKey(item));
    });

    const allTransactions: DepoTransaction[] = [];
    let currentInventory = [...activeInventory];

    // Process scanned items
    scannedItems.forEach(data => {
      const sistemStok = Number((data.item as any)[regionProp] ?? 0);
      const fark = data.sayilan - sistemStok;
      
      const { tx, updatedInv } = prepareSayimFarkiData(data.item, fark, currentInventory, data.excessSource, false);
      if (tx) allTransactions.push(tx);
      if (updatedInv) currentInventory = updatedInv;
    });

    // Process never-scanned items
    neverScanned.forEach(item => {
      const { tx, updatedInv } = prepareSayimFarkiData(item, 0, currentInventory, undefined, true);
      if (tx) allTransactions.push(tx);
      if (updatedInv) currentInventory = updatedInv;
    });

    // Commit all at once
    if (onAddTransactions && allTransactions.length > 0) {
      onAddTransactions(allTransactions);
    } else {
      allTransactions.forEach(tx => onAddTransaction?.(tx));
    }

    if (onUpdateInventory) onUpdateInventory(currentInventory);

    // Otomatik olarak sayım sonuçlarını Onay Bekleyenler listesine gönder (Fark bilgisiyle birlikte)
    const pendingCountItems = scannedItems.map(data => {
      const sistemStok = Number((data.item as any)[regionProp] ?? 0);
      const fark = data.sayilan - sistemStok;
      const farkStr = fark > 0 ? `+${fark}` : `${fark}`;
      return {
        pn: data.item.partNumber || data.item.pn || '-',
        sn: data.item.serialAndNotes || data.item.sn || '-',
        itemName: data.item.description || data.item.name || 'Sayılan Parça',
        category: sayimCategory,
        txCategory: 'sayim',
        quantity: fark !== 0 ? Math.abs(fark) : data.sayilan,
        sistemStok,
        sayilanAdet: data.sayilan,
        fark,
        date: new Date().toLocaleString('tr-TR'),
        requestedBy: sayimPersonel || 'Sayım Sorumlusu',
        depot: sayimDepoBolge,
        tailNo: 'SAYIM SONUCU',
        unit: selectedUnit === 'all' ? (data.item.unit || 'at802') : selectedUnit,
        notes: `SAYIM ONAYI (${sayimDepoBolge} DEPOSU) - Sistem: ${sistemStok}, Sayılan: ${data.sayilan}, Fark: ${farkStr} Adet`
      };
    });

    if (pendingCountItems.length > 0) {
      fetch('/api/onay-bekleyenler', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: pendingCountItems })
      }).catch(() => {});
    }

    setSessionActive(false);
    setSessionScannedData({});
    stopCamera();
    
    if (showNotification) {
      showNotification(`✅ Sayım bitti. ${scannedItems.length} sayılan ürün Onay Bekleyenler listesine aktarıldı.`);
    }
  };

  const prepareSayimFarkiData = (item: DepoItem, fark: number, currentInv: DepoItem[], excessSource?: 'ANKARA' | 'BOLGE', isVerified: boolean = false) => {
    const regionInfo = DEPO_REGIONS.find(r => r.id === sayimDepoBolge);
    const regionProp = regionInfo?.prop || 'ankaraMevcut';
    const sistemStok = Number((item as any)[regionProp] ?? 0);
    const sayilanAdet = sistemStok + fark;

    let islemTuru = '';
    let notes = '';
    let txType = 'TRANSFER';

    // Sadece artı (+) veya eksi (-) fark varsa transfer geçmişine kaydedilir
    let tx: DepoTransaction | null = null;

    if (fark > 0) {
      const sourceLabel = excessSource === 'ANKARA' ? 'ANKARA TEMİNİ' : 'BÖLGE MÜDÜRLÜĞÜ TEMİNİ';
      if (sayimDepoBolge === 'ANKARA') {
        islemTuru = `ANKARA GİREN (FAZLA TESPİT)`;
        txType = 'GİRİŞ';
      } else {
        islemTuru = `${sayimDepoBolge} SAYIM TESPİT TRANSFER (FAZLA TESPİT)`;
        txType = 'TRANSFER';
      }
      if (excessSource === 'ANKARA' && sayimDepoBolge !== 'ANKARA') {
        notes = `SAYIM OTURUMU: Fazla ürün tespiti (${sourceLabel}). Ankara Stoku -${fark} düşürüldü, ${sayimDepoBolge} Stoku ${sistemStok}'den ${sayilanAdet}'e (+${fark}) yükseltildi.`;
      } else {
        notes = `SAYIM OTURUMU: Fazla ürün tespiti (${sourceLabel}). Sistem: ${sistemStok}, Sayılan: ${sayilanAdet}. Fark: +${fark}.`;
      }
    } else if (fark < 0) {
      islemTuru = `${sayimDepoBolge} SAYIM TESPİT ÇIKAN (EKSIK TESPİT)`;
      notes = `SAYIM OTURUMU: Eksik ürün tespiti. Sistem: ${sistemStok}, Sayılan: ${sayilanAdet}. Fark: ${fark}.`;
      txType = 'ÇIKIŞ';
    }

    const pn = (item.partNumber || item.pn || (item as any).MALZEME_PN || '').trim();
    const desc = item.description || item.name || (item as any).MALZEME_ADI || '';

    const now = new Date();
    const formattedTime = now.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const fullDate = `${sayimTarih} ${formattedTime}`;

    if (fark !== 0) {
      tx = {
        id: `tx_sayim_${Date.now()}_${Math.random().toString(36).substr(2, 7)}`,
        timestamp: now.toISOString(),
        date: fullDate,
        type: txType,
        islemTuru: islemTuru,
        itemName: desc,
        pn: pn,
        sn: item.serialAndNotes || item.sn || (item as any).SERI_NO || '-',
        quantity: Math.abs(fark),
        adet: Math.abs(fark),
        operator: sayimPersonel || 'Sayım Yetkilisi',
        receivedBy: `${sayimDepoBolge} DEPO SAYIMI`,
        location: `${sayimDepoBolge} DEPOSU`,
        unit: selectedUnit,
        notes: notes,
        isNewSessionTx: true,
        sistemStok: sistemStok,
        sayilanAdet: sayilanAdet,
        fark: fark
      };
    }

    const updatedInv = currentInv.map(it => {
      const itAny = it as any;
      const itPn = (it.partNumber || it.pn || itAny.MALZEME_PN || '').trim().toLowerCase();
      const itDesc = (it.description || it.name || itAny.MALZEME_ADI || '').trim().toLowerCase();
      if ((pn && itPn === pn.toLowerCase()) || (desc && itDesc === desc.toLowerCase())) {
        const updated = { ...it };
        (updated as any)[regionProp] = sayilanAdet;
        
        // S/N ekleme / çıkarma güncellemesi
        if (newSayimSerialInput.trim()) {
          const currentSnStr = updated.serialAndNotes || updated.sn || '';
          updated.serialAndNotes = currentSnStr && currentSnStr !== '-' ? `${currentSnStr}, ${newSayimSerialInput.trim()}` : newSayimSerialInput.trim();
          updated.sn = updated.serialAndNotes;
        } else if (missingSayimSerialSelect.trim()) {
          const currentSnStr = updated.serialAndNotes || updated.sn || '';
          const parts = currentSnStr.split(/[,/;]+/).map((s: string) => s.trim()).filter((s: string) => s && s.toLowerCase() !== missingSayimSerialSelect.trim().toLowerCase());
          updated.serialAndNotes = parts.length > 0 ? parts.join(', ') : '-';
          updated.sn = updated.serialAndNotes;
        }

        // Bölge sayımında Ankara Temini seçildiyse, fazla miktar Ankara Merkez stokundan düşürülür
        if (fark > 0 && excessSource === 'ANKARA' && sayimDepoBolge !== 'ANKARA') {
          const currentAnkara = Number(it.ankaraMevcut || 0);
          updated.ankaraMevcut = Math.max(0, currentAnkara - fark);
        }

        const ankara = Number(updated.ankaraMevcut ?? it.ankaraMevcut ?? 0);
        const karain = Number(regionProp === 'karainMevcut' ? sayilanAdet : updated.karainMevcut ?? it.karainMevcut ?? 0);
        const canakkale = Number(regionProp === 'canakkaleMevcut' ? sayilanAdet : updated.canakkaleMevcut ?? it.canakkaleMevcut ?? 0);
        const milas = Number(regionProp === 'milasMevcut' ? sayilanAdet : updated.milasMevcut ?? it.milasMevcut ?? 0);
        const bursa = Number(regionProp === 'bursaMevcut' ? sayilanAdet : updated.bursaMevcut ?? it.bursaMevcut ?? 0);
        updated.toplamStok = ankara + karain + canakkale + milas + bursa;
        
        return updated;
      }
      return it;
    });

    return { tx, updatedInv };
  };

  const commitSayimFarki = (item: DepoItem, fark: number, excessSource?: 'ANKARA' | 'BOLGE', isVerified: boolean = false) => {
    const { tx, updatedInv } = prepareSayimFarkiData(item, fark, activeInventory, excessSource, isVerified);

    if (fark !== 0) {
      const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
      const sistemStok = Number((item as any)[regionProp] ?? 0);
      const sayilanAdet = sistemStok + fark;
      const farkStr = fark > 0 ? `+${fark}` : `${fark}`;
      fetch('/api/onay-bekleyenler', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items: [{
            pn: item.partNumber || item.pn || '-',
            sn: item.serialAndNotes || item.sn || '-',
            itemName: item.description || item.name || 'Sayılan Parça',
            category: sayimCategory,
            txCategory: 'sayim',
            quantity: Math.abs(fark),
            sistemStok,
            sayilanAdet,
            fark,
            date: new Date().toLocaleString('tr-TR'),
            requestedBy: sayimPersonel || 'Sayım Sorumlusu',
            depot: sayimDepoBolge,
            tailNo: 'SAYIM SONUCU',
            unit: selectedUnit === 'all' ? (item.unit || 'at802') : selectedUnit,
            notes: `SAYIM ONAYI (${sayimDepoBolge} DEPOSU) - Sistem: ${sistemStok}, Sayılan: ${sayilanAdet}, Fark: ${farkStr} Adet`
          }]
        })
      }).catch(() => {});
    }

    if (onAddTransaction && tx) onAddTransaction(tx);
    if (onUpdateInventory && updatedInv) onUpdateInventory(updatedInv);
  };

  const resolveNextExcess = (source: 'ANKARA' | 'BOLGE') => {
    const [current, ...remaining] = excessItemsQueue;
    if (current) {
      if (sessionActive) {
        // Collect for batch processing at the end
        const itemKey = getCountItemKey(current.item);
        const sayilan = Number(sayilanAdetInput);
        
        setSessionScannedData(prev => ({
          ...prev,
          [itemKey]: { item: current.item, sayilan, excessSource: source }
        }));
        
        setSayimItem(null);
        setSayilanAdetInput('');
      } else {
        // Immediate processing if not in session
        commitSayimFarki(current.item, current.fark, source);
        setSayimItem(null);
        setSayilanAdetInput('');
      }
    }
    
    setExcessItemsQueue(remaining);
    if (remaining.length === 0) {
      setIsExcessPromptOpen(false);
      if (!sessionActive) {
        stopCamera();
        if (showNotification) showNotification('✅ Sayım farkı işlendi.');
      }
    }
  };

  // Manual search submit
  const handleManualSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!manualCodeInput.trim()) return;
    playBeep();
    setIsManualSearchDropdownOpen(false);
    handleCodeRead(manualCodeInput.trim());
    setManualCodeInput('');
  };

  // Auto-trigger depot query when a barcode/QR scanner enters or pastes an exact P/N into the manual P/N box
  useEffect(() => {
    const q = manualCodeInput.trim();
    if (!isOpen || !isUnlocked || !q || q.length < 3) return;
    if (q.startsWith('OGM-DEPO-')) {
      setIsManualSearchDropdownOpen(false);
      playBeep();
      handleCodeRead(q);
      setManualCodeInput('');
      return;
    }
    const qLower = q.toLowerCase();
    const exactPnMatches = activeInventory.filter(i => {
      const p = String(i.partNumber || i.pn || (i as any).MALZEME_PN || '').trim().toLowerCase();
      return p && p !== '-' && p === qLower;
    });
    if (exactPnMatches.length >= 1) {
      const hasLongerPrefix = activeInventory.some(i => {
        const p = String(i.partNumber || i.pn || (i as any).MALZEME_PN || '').trim().toLowerCase();
        return p.length > qLower.length && p.startsWith(qLower);
      });
      const timer = setTimeout(() => {
        setIsManualSearchDropdownOpen(false);
        playBeep();
        handleCodeRead(q);
        setManualCodeInput('');
      }, hasLongerPrefix ? 550 : 220);
      return () => clearTimeout(timer);
    }
  }, [manualCodeInput, isOpen, isUnlocked, activeInventory]);

  // Auto-match barcode / exact P/N inside "+ Parça / Barkod Ekle" search box
  useEffect(() => {
    if (!isUrunTalebiModalOpen || !isAddTalepItemOpen) return;
    const q = addTalepSearchQuery.trim();
    if (!q || q.length < 3) return;

    const sourceInv = (activeInventory && activeInventory.length > 0) ? activeInventory : inventory;

    if (q.startsWith('OGM-DEPO-')) {
      const { found, extractedQuery } = resolveInventoryItemByCode(q);
      if (found) {
        playBeep();
        stopTalepCamera();
        handleAddTalepItemFromInventory(found);
      } else if (extractedQuery && extractedQuery !== q) {
        setAddTalepSearchQuery(extractedQuery);
      }
      return;
    }

    const qLower = q.toLowerCase();
    const exactPnMatches = sourceInv.filter(i => {
      const p = String(i.partNumber || i.pn || (i as any).MALZEME_PN || '').trim().toLowerCase();
      return p && p !== '-' && p === qLower;
    });

    if (exactPnMatches.length === 1) {
      const hasLongerPrefix = sourceInv.some(i => {
        const p = String(i.partNumber || i.pn || (i as any).MALZEME_PN || '').trim().toLowerCase();
        return p.length > qLower.length && p.startsWith(qLower);
      });
      const timer = setTimeout(() => {
        playBeep();
        stopTalepCamera();
        handleAddTalepItemFromInventory(exactPnMatches[0]);
      }, hasLongerPrefix ? 650 : 320);
      return () => clearTimeout(timer);
    }
  }, [addTalepSearchQuery, isUrunTalebiModalOpen, isAddTalepItemOpen, activeInventory, inventory]);

  // Single item save or Confirmation from Popup
  const handleSaveSayimFarki = () => {
    if (!sayimItem) return;
    const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
    const sistemStok = Number((sayimItem as any)[regionProp] ?? 0);
    const sayilan = Number(sayilanAdetInput);
    const fark = isNaN(sayilan) ? 0 : sayilan - sistemStok;

    const it = sayimItem as any;
    const itemKey = getCountItemKey(sayimItem);

    // 1. Update Session Tracking if active
    if (sessionActive) {
      setSessionCounts(prev => ({ ...prev, [itemKey]: sayilan }));
      setScannedInSession(prev => {
        const next = new Set(prev);
        next.add(itemKey);
        return next;
      });

      // Add to mini-history in the module
      setSayimGecmisi(prev => [{
        timestamp: new Date().toISOString(),
        pn: sayimItem.partNumber || sayimItem.pn || it.MALZEME_PN || '-',
        desc: sayimItem.description || sayimItem.name || it.MALZEME_ADI || '-',
        region: sayimDepoBolge,
        sistemAdet: sistemStok,
        sayilanAdet: sayilan,
        fark: fark,
        personel: sayimPersonel
      }, ...prev]);
    }

    // 2. Commit logic
    if (fark > 0) {
        // Excess items still go through the 'source' prompt to know where they came from
        setExcessItemsQueue([{ item: sayimItem, fark }]);
        setIsExcessPromptOpen(true);
    } else {
        if (sessionActive) {
            // Collect for batch processing
            setSessionScannedData(prev => ({
              ...prev,
              [itemKey]: { item: sayimItem!, sayilan }
            }));
            setSayimItem(null);
            setSayilanAdetInput('');
            if (showNotification) showNotification('✅ Sayım kaydedildi (Oturum bitince işlenecek).');
        } else {
            // Immediate processing
            commitSayimFarki(sayimItem, fark);
            setSayimItem(null);
            setSayilanAdetInput('');
            if (showNotification) showNotification('✅ Sayım kaydedildi.');
        }
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[10150] bg-black/85 sm:backdrop-blur-xs flex items-center justify-center p-1.5 sm:p-4 overflow-y-auto overscroll-contain">
      <div className="bg-slate-900 text-white rounded-2xl sm:rounded-3xl shadow-2xl max-w-4xl w-full border border-slate-700 overflow-hidden flex flex-col max-h-[96dvh] sm:max-h-[92vh]">
        
        {/* HEADER */}
        <div className="bg-slate-950 px-3.5 py-3 sm:px-5 sm:py-4 border-b border-slate-800 flex items-center justify-between gap-2 shrink-0">
          <div className="flex items-center gap-2.5 sm:gap-3 min-w-0">
            <div className="p-2 sm:p-2.5 bg-emerald-500/20 text-emerald-400 rounded-xl sm:rounded-2xl border border-emerald-500/30 shadow-inner shrink-0">
              <Camera className="w-4 h-4 sm:w-5 sm:h-5" />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-1.5 flex-wrap">
                <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white truncate">
                  BARKOD OKUYUCU & SAYIM TERMİNALİ
                </h3>
                {isUnlocked && (
                  <span className="px-2 py-0.5 rounded text-[9px] sm:text-[10px] font-black bg-emerald-950 text-emerald-300 border border-emerald-700 shrink-0">
                    {UNITS_SECURITY.find(u => u.id === selectedUnit)?.name} AKTİF
                  </span>
                )}
              </div>
              <p className="text-[10px] sm:text-[11px] text-slate-400 font-sans truncate">
                Mobil Kamera ile Barkod / Karekod Okuma ve Depo Stok Sayım Modülü
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="text-slate-400 hover:text-white p-2 -mr-1 rounded-lg transition-colors cursor-pointer text-lg font-bold shrink-0"
          >
            ✕
          </button>
        </div>

        {/* BODY */}
        <div className="p-3 sm:p-5 overflow-y-auto overscroll-contain flex-1 space-y-3 sm:space-y-4">
          
          {/* STEP 1: SECURITY / UNIT SELECTION SCREEN IF NOT UNLOCKED */}
          {!isUnlocked ? (
            <div className="max-w-md mx-auto py-6 space-y-4">
              <div className="text-center space-y-1">
                <div className="w-14 h-14 rounded-2xl bg-slate-800 border border-slate-700 flex items-center justify-center mx-auto text-emerald-400 shadow-md">
                  <Lock className="w-7 h-7" />
                </div>
                <h4 className="text-base font-black uppercase text-white mt-3">HAVA ARACI BİRİMİ & ŞİFRE</h4>
                <p className="text-xs text-slate-400">
                  Depo barkod okuyucu ve sayım terminalini kullanmak için birim seçip şifresini giriniz.
                </p>
              </div>

              <form onSubmit={handlePasswordSubmit} className="bg-slate-950 p-5 rounded-2xl border border-slate-800 space-y-3.5 shadow-lg">
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">
                    HAVA ARACI / BİRİM SEÇİMİ
                  </label>
                  <select
                    value={selectedUnit}
                    onChange={(e) => setSelectedUnit(e.target.value)}
                    className="w-full px-3 py-2.5 bg-slate-900 border border-slate-700 rounded-xl text-xs font-bold text-white focus:outline-none focus:border-emerald-500"
                  >
                    {UNITS_SECURITY.map(u => (
                      <option key={u.id} value={u.id}>
                        {u.name} ({u.code})
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">
                    BİRİM ŞİFRESİ
                  </label>
                  <input
                    type="password"
                    autoFocus
                    value={passwordInput}
                    onChange={(e) => setPasswordInput(e.target.value)}
                    placeholder="ŞİFRE GİRİNİZ"
                    className="w-full px-3 py-2.5 bg-slate-900 border border-slate-700 rounded-xl text-center text-sm font-mono font-black tracking-widest text-emerald-300 placeholder:text-slate-600 placeholder:font-normal placeholder:tracking-normal focus:outline-none focus:border-emerald-500 opacity-90"
                  />
                  {passwordError && (
                    <p className="text-xs text-rose-400 font-bold mt-1 text-center">
                      ❌ Hatalı birim şifresi girdiniz!
                    </p>
                  )}
                </div>

                <button
                  type="submit"
                  className="w-full py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black uppercase tracking-wider transition shadow-sm cursor-pointer active:scale-95"
                >
                  Terminale Giriş Yap
                </button>
              </form>
            </div>
          ) : (
            /* STEP 2: MAIN UNLOCKED TERMINAL */
            <div className="space-y-4">
              
              {/* TOP MODE TOGGLE (MODÜL 1: BARKOD OKU vs MODÜL 2: SAYIMA BAŞLA) */}
              <div className="bg-slate-950 p-2 rounded-2xl border border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="grid grid-cols-2 gap-1.5 flex-1">
                  <button
                    type="button"
                    onClick={() => {
                      setActiveModule('read');
                      setMatchedItem(null);
                      if (sessionActive) handleFinishSession();
                    }}
                    className={`flex items-center justify-center gap-1.5 py-2.5 px-2 sm:px-3 rounded-xl text-[10.5px] sm:text-xs font-black uppercase transition-colors cursor-pointer ${
                      activeModule === 'read'
                        ? 'bg-emerald-600 text-white shadow-md'
                        : 'bg-slate-900 text-slate-400 hover:text-white'
                    }`}
                  >
                    <Barcode className="w-4 h-4 shrink-0" />
                    <span className="truncate">1. BARKOD OKU</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      setActiveModule('count');
                      setSayimItem(null);
                      if (isCameraActive) stopCamera();
                    }}
                    className={`flex items-center justify-center gap-1.5 py-2.5 px-2 sm:px-3 rounded-xl text-[10.5px] sm:text-xs font-black uppercase transition-colors cursor-pointer ${
                      activeModule === 'count'
                        ? 'bg-amber-600 text-white shadow-md'
                        : 'bg-slate-900 text-slate-400 hover:text-white'
                    }`}
                  >
                    <Layers className="w-4 h-4 shrink-0" />
                    <span className="truncate">2. DEPO SAYIMI</span>
                  </button>
                </div>

                <button
                  type="button"
                  onClick={() => {
                    stopCamera();
                    stopTalepCamera();
                    setIsUnlocked(false);
                    setPasswordInput('');
                    setPasswordError(false);
                    setLiveInventory([]);
                    setMatchedItem(null);
                    setScannedResultText(null);
                    loadedUnitRef.current = '';
                  }}
                  className="w-full sm:w-auto px-3 py-2 sm:py-1.5 text-[10px] font-bold text-slate-300 hover:text-white bg-slate-900 rounded-lg border border-slate-800 transition cursor-pointer"
                  title="Birimi Kilitle / Değiştir"
                >
                  Birim Değiştir
                </button>
              </div>

              {/* CAMERA SCANNER & VIEWFINDER COMPONENT (ONLY FOR READ MODE OR ACTIVE COUNT SESSION) */}
              {(activeModule === 'read' || (activeModule === 'count' && sessionActive)) && (
                <div className="bg-slate-950 p-4 rounded-3xl border border-slate-800 space-y-3">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold text-slate-200 uppercase font-mono">
                        Kamera Tarayıcı (Code-128 & Karekod)
                      </span>
                      {isCameraActive ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-950 text-emerald-400 border border-emerald-800 animate-pulse">
                          <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />
                          Kamera Açık
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-800 text-slate-400">
                          Kamera Kapalı
                        </span>
                      )}
                    </div>

                    <div className="flex items-center gap-1.5 flex-wrap">
                      {!isCameraActive && (
                        <button
                          type="button"
                          onClick={startCamera}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold transition shadow-sm cursor-pointer active:scale-95"
                        >
                          <Camera className="w-4 h-4" />
                          <span>Kamerayı Aç</span>
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Viewfinder Video Element */}
                  <div className="relative rounded-2xl overflow-hidden bg-black/60 border border-slate-800 min-h-[220px] flex items-center justify-center">
                    <div 
                      id={scannerContainerId} 
                      className="w-full max-w-sm mx-auto overflow-hidden" 
                    />

                    {!isCameraActive && (
                      <div className="absolute inset-0 bg-slate-950/90 flex flex-col items-center justify-center p-6 text-center space-y-3 z-10">
                        <Barcode className="w-12 h-12 text-slate-600 mx-auto" />
                        <p className="text-xs text-slate-400">
                          {activeModule === 'count' 
                            ? 'Sayım oturumu için kamerayı açarak ürün barkodlarını okutabilirsiniz.'
                            : 'Barkod veya karekod okutmak için Kamerayı Aç butonuna basınız.'}
                        </p>
                        <button
                          type="button"
                          onClick={startCamera}
                          className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black uppercase tracking-wider rounded-xl shadow-md transition cursor-pointer active:scale-95 flex items-center gap-2 mx-auto"
                        >
                          <Camera className="w-4 h-4" />
                          <span>Kamerayı Aç</span>
                        </button>
                      </div>
                    )}

                    {cameraError && (
                      <div className="absolute inset-0 bg-black/85 flex flex-col items-center justify-center p-4 text-center text-xs text-rose-300">
                        <AlertTriangle className="w-8 h-8 text-rose-400 mb-2" />
                        <span>{cameraError}</span>
                        <span className="text-[10px] text-slate-400 mt-1">Aşağıdaki arama çubuğundan barkod numarasını manuel de girebilirsiniz.</span>
                      </div>
                    )}
                  </div>

                  {/* Manual Barcode / P/N Input fallback with Autocomplete */}
                  <form onSubmit={handleManualSubmit} className="flex flex-col sm:flex-row sm:items-center gap-2 pt-1 relative">
                    <div ref={manualSearchRef} className="relative flex-1 w-full">
                      <Search className="w-4 h-4 text-slate-500 absolute left-3 top-3 z-10" />
                      <input
                        type="text"
                        value={manualCodeInput}
                        onChange={(e) => {
                          setManualCodeInput(e.target.value);
                          setIsManualSearchDropdownOpen(true);
                        }}
                        onFocus={() => setIsManualSearchDropdownOpen(true)}
                        placeholder={isLoadingInventory ? "Seçilen birim stokları yükleniyor..." : "Barkod / Karekod verisi veya P/N / Malzeme Adı yazınız..."}
                        className="w-full pl-9 pr-3 py-2.5 sm:py-2 bg-slate-900 border border-slate-700 rounded-xl text-sm sm:text-xs text-white placeholder-slate-500 font-mono focus:outline-none focus:border-emerald-500 font-bold"
                      />

                      {/* Autocomplete Dropdown List */}
                      {isManualSearchDropdownOpen && manualCodeInput.trim().length > 0 && (
                        <div className="absolute top-full left-0 right-0 z-50 mt-1 bg-slate-950 border-2 border-emerald-500/60 rounded-2xl shadow-2xl max-h-60 overflow-y-auto overscroll-contain divide-y divide-slate-800/80">
                          {isLoadingInventory && manualDropdownMatches.length === 0 ? (
                            <div className="p-3 text-center text-xs text-emerald-300 font-mono animate-pulse">
                              ⏳ {UNITS_SECURITY.find(u => u.id === selectedUnit)?.name || selectedUnit.toUpperCase()} deposundan malzeme listesi yükleniyor...
                            </div>
                          ) : manualDropdownMatches.length === 0 ? (
                            <div className="p-3 text-center text-xs text-slate-400">
                              "{manualCodeInput}" ile eşleşen malzeme bulunamadı.
                            </div>
                          ) : (
                            manualDropdownMatches.map((item, idx) => (
                              <div
                                key={idx}
                                onClick={() => {
                                  const selPn = item.partNumber && item.partNumber !== '-' ? item.partNumber : (item.pn && item.pn !== '-' ? item.pn : (item.description || item.name));
                                  setManualCodeInput(selPn);
                                  setIsManualSearchDropdownOpen(false);
                                  playBeep();
                                  handleCodeRead(selPn);
                                }}
                                className="p-3 sm:p-2.5 hover:bg-emerald-950/60 cursor-pointer text-xs flex items-center justify-between gap-2 transition-colors"
                              >
                                <div className="min-w-0 flex-1">
                                  <div className="font-bold text-white truncate">{item.description || item.name}</div>
                                  <div className="text-[11px] font-mono text-slate-400 truncate">
                                    P/N: <strong className="text-emerald-300">{item.partNumber || item.pn || '-'}</strong> | S/N: {item.serialAndNotes || item.sn || '-'} | Raf: {item.lokasyonNo || item.location || '-'}
                                  </div>
                                </div>
                                <span className="text-[10px] font-mono font-bold bg-emerald-950 text-emerald-300 border border-emerald-800 px-2 py-0.5 rounded-full shrink-0">
                                  Stok: {item.ankaraMevcut || item.toplamStok || 0}
                                </span>
                              </div>
                            ))
                          )}
                        </div>
                      )}
                    </div>
                    <div className="grid grid-cols-2 sm:flex items-center gap-2 shrink-0">
                      <button
                        type="submit"
                        className="px-4 py-2.5 sm:py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs font-bold rounded-xl transition cursor-pointer"
                      >
                        Sorgula
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setIsManualSearchDropdownOpen(false);
                          handleOpenUrunTalebi(matchedItem || null);
                        }}
                        className="px-3.5 py-2.5 sm:py-2 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-400 text-slate-950 text-xs font-black uppercase tracking-wider rounded-xl transition cursor-pointer shadow flex items-center justify-center gap-1.5"
                        title="Depodan Çıkan Talebi Oluştur ('ONAY BEKLEYENLER -AT802' Sayfasına İlet)"
                      >
                        <Package className="w-3.5 h-3.5 text-slate-950" />
                        <span>ÜRÜN TALEBİ</span>
                      </button>
                    </div>
                  </form>
                </div>
              )}

              {/* ============================================================== */}
              {/* MODÜL 1: BARKOD OKU & STOK SORGULAMA DETAY GÖRÜNÜMÜ          */}
              {/* ============================================================== */}
              {activeModule === 'read' && (
                <div className="space-y-3">
                  {matchedItem ? (
                    <div className="bg-slate-950 p-4 sm:p-5 rounded-3xl border-2 border-emerald-500/40 shadow-xl space-y-4 animate-fade-in">
                      
                      {/* Top Item Summary */}
                      <div className="flex items-start justify-between border-b border-slate-800 pb-3 gap-3">
                        <div>
                          <div className="text-[10px] font-mono font-bold text-emerald-400 uppercase tracking-widest">
                            DEPO SİSTEMİNDEKİ EŞLEŞEN MALZEME
                          </div>
                          <h4 className="text-base sm:text-lg font-black text-white mt-0.5">
                            {matchedItem.description || matchedItem.name}
                          </h4>
                          {Boolean(
                            (matchedItem.unit && String(matchedItem.unit).toLowerCase().includes('bell')) ||
                            (matchedItem.lokasyonNo && String(matchedItem.lokasyonNo).toUpperCase().includes('BELL')) ||
                            (matchedItem.location && String(matchedItem.location).toUpperCase().includes('BELL'))
                          ) && selectedUnit !== 'bell429' && (
                            <div className="mt-2 p-2.5 bg-amber-950/70 border border-amber-500/60 rounded-xl text-amber-200 text-xs flex items-center gap-2">
                              <span className="text-base shrink-0">⚠️</span>
                              <div>
                                <span className="font-black text-amber-300">FARKLI DEPO UYARISI (BELL 429):</span> Bu malzeme Bell 429 deposuna aittir. {selectedUnit.toUpperCase()} kuyruğu için çıkış yapıldığında doğrudan stoktan düşmez; <strong>Bell-429 Depo Onayına</strong> talep olarak iletilir.
                              </div>
                            </div>
                          )}
                          <div className="flex items-center gap-2.5 text-xs font-mono text-slate-300 mt-2 flex-wrap">
                            <span className="px-2.5 py-1 rounded-lg bg-indigo-950/90 text-indigo-300 border border-indigo-500/60 font-black text-[11px] flex items-center gap-1.5 shadow-xs">
                              ✈ HAVA ARACI DEPOSU: {
                                ((matchedItem.unit && String(matchedItem.unit).toLowerCase().includes('bell')) ||
                                 (matchedItem.lokasyonNo && String(matchedItem.lokasyonNo).toUpperCase().includes('BELL')) ||
                                 (matchedItem.location && String(matchedItem.location).toUpperCase().includes('BELL')) ||
                                 (matchedItem.partNumber && String(matchedItem.partNumber).toUpperCase() === 'AS3209-236') ||
                                 (matchedItem.pn && String(matchedItem.pn).toUpperCase() === 'AS3209-236'))
                                  ? 'BELL 429 DEPOSU'
                                  : (matchedItem.unit ? `${String(matchedItem.unit).toUpperCase()} DEPOSU` : 'AT-802F DEPOSU')
                              }
                            </span>
                            <span>P/N: <strong className="text-emerald-300">{matchedItem.partNumber || matchedItem.pn || '-'}</strong></span>
                            <span>S/N: <strong>{matchedItem.serialAndNotes || matchedItem.sn || '-'}</strong></span>
                            <span className="px-2.5 py-1 rounded-lg bg-amber-400 text-slate-950 font-black text-[11px] shadow-sm">
                              📍 RAF / LOKASYON: {matchedItem.lokasyonNo || matchedItem.location || 'A-1 (ANKARA)'}
                            </span>
                          </div>
                        </div>

                        <button
                          type="button"
                          onClick={() => {
                            setMatchedItem(null);
                          }}
                          className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-bold rounded-xl cursor-pointer"
                        >
                          Temizle
                        </button>
                      </div>

                      {/* Regional Stock Distribution Cards */}
                      <div>
                        <div className="text-[10px] font-bold text-slate-400 uppercase font-mono mb-2 flex items-center justify-between">
                          <span>GÜNCEL DEPO BÖLGE MEVCUTLARI (CANLI STOK)</span>
                          {Boolean(
                            (matchedItem.unit && String(matchedItem.unit).toLowerCase().includes('bell')) ||
                            (matchedItem.lokasyonNo && String(matchedItem.lokasyonNo).toUpperCase().includes('BELL')) ||
                            (matchedItem.location && String(matchedItem.location).toUpperCase().includes('BELL')) ||
                            (matchedItem.partNumber && String(matchedItem.partNumber).toUpperCase() === 'AS3209-236') ||
                            (matchedItem.pn && String(matchedItem.pn).toUpperCase() === 'AS3209-236')
                          ) && (
                            <span className="text-[10px] text-amber-400 font-bold bg-amber-950/80 px-2.5 py-0.5 rounded-full border border-amber-500/50">
                              BELL 429: Yalnızca Ankara Merkez Deposu Mevcuttur
                            </span>
                          )}
                        </div>

                        {Boolean(
                          (matchedItem.unit && String(matchedItem.unit).toLowerCase().includes('bell')) ||
                          (matchedItem.lokasyonNo && String(matchedItem.lokasyonNo).toUpperCase().includes('BELL')) ||
                          (matchedItem.location && String(matchedItem.location).toUpperCase().includes('BELL')) ||
                          (matchedItem.partNumber && String(matchedItem.partNumber).toUpperCase() === 'AS3209-236') ||
                          (matchedItem.pn && String(matchedItem.pn).toUpperCase() === 'AS3209-236')
                        ) ? (
                          /* Bell 429: Sadece Ankara ve Toplam Stok */
                          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
                            {/* Ankara */}
                            <div className="p-3 bg-slate-900 border-2 border-[#ff9900] rounded-xl text-center shadow-md">
                              <div className="text-[10px] font-bold text-amber-400 uppercase tracking-wider">ANKARA (BELL MERKEZ DEPO)</div>
                              <div className="text-2xl font-black text-white font-mono mt-1">
                                {matchedItem.ankaraMevcut || 0}
                              </div>
                            </div>

                            {/* Muayenede varsa */}
                            {(matchedItem.muayeneToplam || matchedItem.muayeneGiden || 0) > 0 && (
                              <div className="p-3 bg-slate-900 border border-purple-500 rounded-xl text-center">
                                <div className="text-[10px] font-bold text-purple-300 uppercase tracking-wider">MUAYENEDE / BAKIMDA</div>
                                <div className="text-2xl font-black text-purple-300 font-mono mt-1">
                                  {matchedItem.muayeneToplam || matchedItem.muayeneGiden || 0}
                                </div>
                              </div>
                            )}

                            {/* Toplam */}
                            <div className="p-3 bg-slate-900 border-2 border-cyan-400 rounded-xl text-center shadow-lg">
                              <div className="text-[10px] font-black text-cyan-300 uppercase tracking-wider">TOPLAM STOK</div>
                              <div className="text-2xl font-black text-cyan-300 font-mono mt-1">
                                {matchedItem.toplamStok || matchedItem.ankaraMevcut || 0}
                              </div>
                            </div>
                          </div>
                        ) : (
                          /* AT-802 vb.: Tüm bölgeler */
                          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2">
                            {/* Ankara */}
                            <div className="p-2.5 bg-slate-900 border border-[#ff9900] rounded-xl text-center">
                              <div className="text-[9px] font-bold text-slate-400 uppercase">ANKARA</div>
                              <div className="text-lg font-black text-white font-mono mt-0.5">
                                {matchedItem.ankaraMevcut || 0}
                              </div>
                            </div>

                            {/* Milas */}
                            <div className="p-2.5 bg-slate-900 border border-[#ff9900] rounded-xl text-center">
                              <div className="text-[9px] font-bold text-slate-400 uppercase">MİLAS</div>
                              <div className="text-lg font-black text-white font-mono mt-0.5">
                                {matchedItem.milasMevcut || 0}
                              </div>
                            </div>

                            {/* Karain */}
                            <div className="p-2.5 bg-slate-900 border border-[#ffff00] rounded-xl text-center">
                              <div className="text-[9px] font-bold text-slate-400 uppercase">KARAİN</div>
                              <div className="text-lg font-black text-white font-mono mt-0.5">
                                {matchedItem.karainMevcut || 0}
                              </div>
                            </div>

                            {/* Çanakkale */}
                            <div className="p-2.5 bg-slate-900 border border-[#92d050] rounded-xl text-center">
                              <div className="text-[9px] font-bold text-slate-400 uppercase">ÇANAKKALE</div>
                              <div className="text-lg font-black text-white font-mono mt-0.5">
                                {matchedItem.canakkaleMevcut || 0}
                              </div>
                            </div>

                            {/* Bursa */}
                            <div className="p-2.5 bg-slate-900 border border-[#ff00ff] rounded-xl text-center">
                              <div className="text-[9px] font-bold text-slate-400 uppercase">BURSA</div>
                              <div className="text-lg font-black text-white font-mono mt-0.5">
                                {matchedItem.bursaMevcut || 0}
                              </div>
                            </div>

                            {/* Toplam */}
                            <div className="p-2.5 bg-slate-900 border-2 border-cyan-400 rounded-xl text-center shadow-md">
                              <div className="text-[9px] font-black text-cyan-300 uppercase">TOPLAM STOK</div>
                              <div className="text-lg font-black text-cyan-300 font-mono mt-0.5">
                                {matchedItem.toplamStok || 0}
                              </div>
                            </div>
                          </div>
                        )}
                      </div>

                      {/* FORM ONE SORGULA & ÜRÜN TALEBİ BUTTONS & RESULTS */}
                      <div className="pt-2 border-t border-slate-800 space-y-3">
                        <div className="flex items-center justify-between flex-wrap gap-2">
                          <div className="flex items-center gap-2 flex-wrap">
                            <button
                              type="button"
                              onClick={() => handleOpenUrunTalebi(matchedItem)}
                              className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-400 text-slate-950 rounded-xl text-xs font-black uppercase transition-all shadow-md active:scale-95 cursor-pointer font-sans"
                            >
                              <Package className="w-4 h-4 text-slate-950" />
                              <span>ÜRÜN TALEBİ OLUŞTUR (ÇIKAN TALEBİ)</span>
                            </button>

                            <button
                              type="button"
                              onClick={handleSearchFormOne}
                              disabled={isFormOneSearching}
                              className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-black uppercase transition-all shadow-md active:scale-95 cursor-pointer"
                            >
                              <Search className="w-4 h-4" />
                              <span>FORM ONE SORGULA ({matchedItem.partNumber || matchedItem.pn || 'P/N'})</span>
                            </button>
                          </div>

                          {formOneResults && (
                            <span className="text-[11px] font-bold text-emerald-400">
                              ✓ {formOneResults.length} Belge Bulundu
                            </span>
                          )}
                        </div>

                        {/* FORM ONE RESULTS DISPLAY AREA */}
                        {isFormOneSearching && (
                          <div className="p-3 bg-slate-900 rounded-xl text-center text-xs text-amber-300 font-bold animate-pulse">
                            ⏳ Form One Belge / Sertifika Sistemine Bağlanılıyor ve Sorgulanıyor...
                          </div>
                        )}

                        {formOneResults && formOneResults.length > 0 && (
                          <div className="bg-slate-900 border border-blue-500/40 rounded-2xl p-3 space-y-2">
                            <div className="text-[10px] font-black uppercase text-blue-300 font-mono tracking-wider">
                              EŞLEŞEN FORM ONE & KALİTE BELGELERİ
                            </div>
                            <div className="divide-y divide-slate-800">
                              {formOneResults.map((doc, idx) => {
                                const docUnit = doc.unit || (selectedUnit || 'at802').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
                                const pdfOpenUrl = doc.viewUrl || `/api/sertifika-tarama/pdf?unit=${encodeURIComponent(docUnit)}${doc.page ? `#page=${doc.page}` : ''}`;
                                const pdfDownloadUrl = `/api/sertifika-tarama/pdf?unit=${encodeURIComponent(docUnit)}&download=1`;
                                const pdfFileName = `SERTIFIKA_${docUnit}${doc.page ? `_Sayfa_${doc.page}` : ''}.pdf`;
                                const isMobileDevice = typeof navigator !== 'undefined' && (/Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) || window.innerWidth < 768);
                                return (
                                  <div key={idx} className="py-2 flex items-center justify-between gap-3 text-xs flex-wrap sm:flex-nowrap">
                                    <div className="min-w-0 flex-1">
                                      <div className="font-bold text-white flex items-center gap-2 flex-wrap">
                                        <span className="text-amber-400 font-mono">{doc.belgeNo}</span>
                                        <span className="text-[10px] bg-blue-950 text-blue-300 border border-blue-800 px-1.5 py-0.5 rounded font-sans">
                                          Form One
                                        </span>
                                      </div>
                                      <div className="text-[11px] text-slate-300 mt-0.5">{doc.aciklama}</div>
                                      <div className="text-[10px] text-slate-400 font-mono mt-0.5">
                                        P/N: {doc.pn} | S/N: {doc.sn} | Tarih: {doc.tarih}
                                      </div>
                                    </div>
                                    <div className="flex items-center gap-1.5 shrink-0">
                                      <a
                                        href={isMobileDevice ? pdfDownloadUrl : pdfOpenUrl}
                                        target={isMobileDevice ? undefined : '_blank'}
                                        rel={isMobileDevice ? undefined : 'noopener noreferrer'}
                                        download={isMobileDevice ? pdfFileName : undefined}
                                        className="px-3 py-1.5 bg-blue-700 hover:bg-blue-600 text-white text-[11px] font-bold rounded-lg transition cursor-pointer inline-flex items-center gap-1"
                                      >
                                        {isMobileDevice ? <Download className="w-3.5 h-3.5" /> : <ExternalLink className="w-3.5 h-3.5" />}
                                        <span>{isMobileDevice ? 'PDF İNDİR / AÇ' : 'GÖSTER / AÇ'}</span>
                                      </a>
                                      {!isMobileDevice && (
                                        <a
                                          href={pdfDownloadUrl}
                                          download={pdfFileName}
                                          title="PDF Olarak İndir"
                                          className="px-2.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white text-[11px] font-bold rounded-lg transition cursor-pointer inline-flex items-center gap-1"
                                        >
                                          <Download className="w-3.5 h-3.5" />
                                          <span>İndir</span>
                                        </a>
                                      )}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        )}
                      </div>

                    </div>
                  ) : (
                    scannedResultText && (
                      <div className="p-5 bg-slate-950 rounded-2xl border border-amber-500/40 text-center space-y-3 shadow-lg animate-fade-in">
                        <Package className="w-8 h-8 text-amber-400 mx-auto" />
                        <div>
                          <p className="font-bold text-white text-sm">Okunan Kod: "{scannedResultText}"</p>
                          <p className="text-xs text-amber-300 mt-1">Bu koda ait malzeme {selectedUnit.toUpperCase()} deposunda bulunamadı. (Sistem Depo Mevcudu: 0 EA)</p>
                        </div>
                        <button
                          type="button"
                          onClick={() => {
                            setScannedResultText('');
                            setMatchedItem(null);
                          }}
                          className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 font-bold rounded-xl text-xs transition cursor-pointer"
                        >
                          Temizle
                        </button>
                      </div>
                    )
                  )}
                </div>
              )}

              {/* ============================================================== */}
              {/* MODÜL 2: SAYIMA BAŞLA (DEPO SAYIMI & FARK İŞLEME TERMİNALİ)   */}
              {/* ============================================================== */}
              {activeModule === 'count' && (
                <div className="space-y-4">
                  
                  {/* SESSION ACTIVE TERMINAL */}
                  {sessionActive ? (
                    <div className="space-y-4 animate-fade-in">
                      <div className="bg-amber-600/10 border-2 border-amber-500/30 rounded-3xl p-5 space-y-4 shadow-inner">
                        <div className="flex items-center justify-between">
                          <div className="flex items-center gap-3">
                            <div className="w-10 h-10 rounded-2xl bg-amber-600 flex items-center justify-center text-white shadow-lg">
                              <RefreshCw className="w-5 h-5 animate-spin-slow" />
                            </div>
                            <div>
                              <h4 className="text-sm font-black text-amber-100 uppercase tracking-wide">
                                SAYIM OTURUMU AKTİF
                              </h4>
                              <p className="text-[10px] text-amber-400 font-mono">
                                {sayimDepoBolge} DEPOSU SAYILIYOR...
                              </p>
                            </div>
                          </div>
                          
                          <div className="text-right">
                            <div className="text-2xl font-black text-white font-mono leading-none">
                              {scannedInSession.size} / {sessionTotalCount}
                            </div>
                            <div className="text-[9px] font-bold text-amber-500 uppercase tracking-widest mt-1">
                              SAYILAN / TOPLAM KALEM
                            </div>
                          </div>
                        </div>

                        {/* PROGRESS BAR */}
                        <div className="w-full h-2.5 bg-slate-800 rounded-full overflow-hidden border border-slate-700">
                          <div 
                            className="h-full bg-emerald-500 transition-all duration-500 shadow-[0_0_10px_rgba(16,185,129,0.5)]"
                            style={{ width: `${sessionTotalCount > 0 ? Math.min(100, (scannedInSession.size / sessionTotalCount) * 100) : 0}%` }}
                          />
                        </div>

                        <div className="flex items-center gap-2">
                          <button
                            type="button"
                            onClick={() => setShowRemainingModal(true)}
                            className="flex-1 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-[10px] font-black uppercase tracking-wider transition border border-slate-700 flex items-center justify-center gap-2 cursor-pointer"
                          >
                            <Search className="w-3.5 h-3.5" />
                            KALANLARI GÖR ({Math.max(0, sessionTotalCount - scannedInSession.size)})
                          </button>
                          
                          <button
                            type="button"
                            onClick={handleFinishSession}
                            className="flex-1 py-2.5 bg-rose-600 hover:bg-rose-500 text-white rounded-xl text-[10px] font-black uppercase tracking-wider transition shadow-lg flex items-center justify-center gap-2 cursor-pointer"
                          >
                            <CheckCircle2 className="w-3.5 h-3.5" />
                            SAYIMI BİTİR
                          </button>
                        </div>
                      </div>

                      {/* SAYIM CHECKLİSTİ (0'dan Farklı + Önceden Gelmiş 0 Stoklu Tüm Kalemler) */}
                      <div className="bg-slate-950 p-4 rounded-3xl border border-slate-800 space-y-2.5">
                        <div className="flex items-center justify-between flex-wrap gap-2">
                          <div>
                            <div className="text-xs font-black text-white uppercase tracking-wider font-mono">
                              📋 {sayimDepoBolge} DEPOSU SAYIM CHECKLİSTİ
                            </div>
                            <div className="text-[10px] text-slate-400">
                              Stoku 0&apos;dan farklı ve önceden gelip şu an 0 olan tüm kalemler dahildir. Satıra tıklayarak da sayım girebilirsiniz.
                            </div>
                          </div>
                          <span className="text-[10px] font-mono font-black px-2.5 py-1 rounded-full bg-amber-950/60 text-amber-300 border border-amber-700/50">
                            {scannedInSession.size} / {sessionTotalCount} Sayıldı
                          </span>
                        </div>

                        {isLoadingInventory && sessionItemsToCount.length === 0 ? (
                          <div className="py-6 text-center text-xs font-mono text-amber-300 animate-pulse">
                            ⏳ Drive stok verileri yükleniyor... Checklist hazırlanıyor...
                          </div>
                        ) : sessionItemsToCount.length === 0 ? (
                          <div className="py-6 text-center text-xs text-slate-400">
                            {sayimDepoBolge} ({sayimCategory === 'kimyasal' ? 'Kimyasal' : 'Sarf & Parça'}) deposunda sayılacak kalem bulunamadı.
                          </div>
                        ) : (
                          <div className="max-h-64 overflow-y-auto overscroll-contain divide-y divide-slate-800/80 rounded-2xl border border-slate-800 bg-slate-900/40">
                            {sessionItemsToCount.map((item, idx) => {
                              const key = getCountItemKey(item);
                              const isCounted = scannedInSession.has(key);
                              const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
                              const sistemStok = Number((item as any)[regionProp] ?? 0);
                              const countedVal = sessionCounts[key];
                              return (
                                <div
                                  key={key + idx}
                                  onClick={() => {
                                    setSayimItem(item);
                                    setNewSayimSerialInput('');
                                    setMissingSayimSerialSelect('');
                                    setSayilanAdetInput(isCounted && countedVal !== undefined ? countedVal : sistemStok);
                                  }}
                                  className={`px-3 py-2.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2 cursor-pointer transition text-xs ${
                                    isCounted ? 'bg-emerald-950/25 hover:bg-emerald-950/40' : 'hover:bg-slate-800/60'
                                  }`}
                                >
                                  <div className="flex items-center gap-2.5 min-w-0 flex-1">
                                    <div className={`w-5 h-5 rounded-md border flex items-center justify-center shrink-0 ${
                                      isCounted
                                        ? 'bg-emerald-600 border-emerald-400 text-white'
                                        : 'bg-slate-950 border-slate-600 text-transparent'
                                    }`}>
                                      <Check className="w-3.5 h-3.5" />
                                    </div>
                                    <div className="min-w-0 flex-1">
                                      <div className={`font-bold truncate ${isCounted ? 'text-emerald-200' : 'text-white'}`}>
                                        {item.description || item.name || (item as any).MALZEME_ADI}
                                      </div>
                                      <div className="text-[10px] font-mono text-slate-400 truncate">
                                        P/N: <strong className="text-amber-300">{item.partNumber || item.pn || '-'}</strong> | Raf: {item.lokasyonNo || item.location || '-'}
                                      </div>
                                    </div>
                                  </div>
                                  <div className="flex items-center flex-wrap gap-1.5 pl-7 sm:pl-0 shrink-0 font-mono">
                                    {sistemStok <= 0 ? (
                                      <span className="px-2 py-0.5 rounded bg-rose-950/80 text-rose-300 border border-rose-700/60 text-[10px] font-bold">
                                        0 EA • Önceden Gelmiş
                                      </span>
                                    ) : (
                                      <span className="px-2 py-0.5 rounded bg-slate-800 text-amber-300 text-[10px] font-bold">
                                        Sistem: {sistemStok} EA
                                      </span>
                                    )}
                                    {isCounted && (
                                      <span className="px-2 py-0.5 rounded bg-emerald-600 text-white text-[10px] font-black">
                                        Sayılan: {countedVal ?? sistemStok}
                                      </span>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>

                      {/* LAST SCANNED ITEM MINI PREVIEW */}
                      {scannedResultText && (
                        <div className="bg-slate-950 p-4 rounded-2xl border border-slate-800 flex items-center gap-3">
                          <div className="w-10 h-10 rounded-xl bg-slate-900 border border-slate-800 flex items-center justify-center text-emerald-400">
                            <Check className="w-6 h-6" />
                          </div>
                          <div className="flex-1">
                            <div className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">SON OKUTULAN BARKOD</div>
                            <div className="text-xs font-mono text-emerald-300 truncate">{scannedResultText}</div>
                          </div>
                          {sayimItem && (
                            <div className="text-right">
                              <div className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">SAYILAN ADET</div>
                              <div className="text-xs font-mono font-black text-white">
                                {sessionCounts[getCountItemKey(sayimItem)] || 0} EA
                              </div>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  ) : (
                    /* STEP 1: SETTINGS */
                    <div className="bg-slate-950 p-4 rounded-3xl border border-slate-800 space-y-4 animate-fade-in">
                      <div className="text-xs font-black uppercase text-amber-400 font-mono flex items-center justify-between border-b border-slate-800 pb-2">
                        <span>ADIM 1: DEPO SAYIM OTURUMU DETAYLARI</span>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                        <div className="bg-slate-900/50 p-3 rounded-2xl border border-slate-800">
                          <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1.5 flex items-center gap-1.5">
                            <Calendar className="w-3.5 h-3.5" />
                            SAYIM TARİHİ
                          </label>
                          <input
                            type="date"
                            value={sayimTarih}
                            onChange={(e) => setSayimTarih(e.target.value)}
                            className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-xl text-white font-bold focus:outline-none focus:border-amber-500 transition-colors"
                          />
                        </div>

                        <div className="bg-slate-900/50 p-3 rounded-2xl border border-slate-800">
                          <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1.5 flex items-center gap-1.5">
                            <User className="w-3.5 h-3.5" />
                            SAYIMI YAPAN PERSONEL
                          </label>
                          <PersonnelAutocomplete
                            value={sayimPersonel}
                            onChange={setSayimPersonel}
                            placeholder="Personel seçiniz veya yazınız..."
                            className="w-full bg-slate-900 border-slate-700 text-white font-bold rounded-xl focus:border-amber-500 h-[38px]"
                          />
                        </div>

                        <div className="bg-slate-900/50 p-3 rounded-2xl border border-slate-800 sm:col-span-2 space-y-2">
                          <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono flex items-center gap-1.5">
                            <Layers className="w-3.5 h-3.5 text-amber-400" />
                            SAYILACAK DEPO TÜRÜ (AYRI DEPO & EXCEL SİSTEMİ)
                          </label>
                          <div className="grid grid-cols-2 gap-2">
                            <button
                              type="button"
                              onClick={() => setSayimCategory('sarf')}
                              className={`py-2 px-3 rounded-xl text-xs font-black uppercase transition cursor-pointer flex items-center justify-center gap-2 ${
                                sayimCategory === 'sarf'
                                  ? 'bg-emerald-600 text-white ring-2 ring-emerald-400 shadow-md'
                                  : 'bg-slate-900 text-slate-400 hover:text-white border border-slate-700'
                              }`}
                            >
                              <span>📦 1. SARF &amp; PARÇA DEPOSU</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => setSayimCategory('kimyasal')}
                              className={`py-2 px-3 rounded-xl text-xs font-black uppercase transition cursor-pointer flex items-center justify-center gap-2 ${
                                sayimCategory === 'kimyasal'
                                  ? 'bg-amber-600 text-white ring-2 ring-amber-400 shadow-md'
                                  : 'bg-slate-900 text-slate-400 hover:text-white border border-slate-700'
                              }`}
                            >
                              <span>🧪 2. KİMYASAL DEPO</span>
                            </button>
                          </div>
                        </div>

                        <div className="bg-slate-900/50 p-3 rounded-2xl border border-slate-800 sm:col-span-2">
                          <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1.5 flex items-center gap-1.5">
                            <MapPin className="w-3.5 h-3.5" />
                            SAYILAN DEPO LOKASYONU (İL)
                          </label>
                          <select
                            value={sayimDepoBolge}
                            onChange={(e) => setSayimDepoBolge(e.target.value as any)}
                            className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-xl text-white font-bold focus:outline-none focus:border-amber-500 transition-colors"
                          >
                            <option value="ANKARA">ANKARA (Merkez Depo)</option>
                            <option value="KARAİN">KARAİN Deposu</option>
                            <option value="ÇANAKKALE">ÇANAKKALE Deposu</option>
                            <option value="MİLAS">MİLAS Deposu</option>
                            <option value="BURSA">BURSA Deposu</option>
                          </select>
                        </div>
                      </div>

                      <button
                        type="button"
                        onClick={handleStartSession}
                        className="w-full py-3 bg-amber-600 hover:bg-amber-500 text-white rounded-2xl text-xs font-black uppercase tracking-widest transition shadow-lg flex items-center justify-center gap-2 cursor-pointer active:scale-[0.98]"
                      >
                        <Layers className="w-4 h-4" />
                        SAYIMI BAŞLAT (KAMERAYI AÇ)
                      </button>
                    </div>
                  )}

                  {/* Sayılan Ürün Doğrulama & Fark Hesaplama Penceresi (MODAL) */}
                  {sayimItem && (() => {
                    const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
                    const sistemStok = Number((sayimItem as any)[regionProp] ?? 0);
                    const sayilan = Number(sayilanAdetInput);
                    const fark = isNaN(sayilan) ? 0 : sayilan - sistemStok;

                    return (
                      <div className="fixed inset-0 z-[10300] bg-slate-900/95 sm:backdrop-blur-sm flex items-center justify-center p-2 sm:p-4 overflow-y-auto overscroll-contain">
                        <div className="bg-slate-950 p-4 sm:p-6 rounded-2xl sm:rounded-3xl border-2 border-amber-500/50 shadow-2xl space-y-4 sm:space-y-5 max-w-xl w-full max-h-[94dvh] overflow-y-auto">
                          <div className="flex items-start justify-between border-b border-slate-800 pb-3">
                            <div>
                              <span className="text-[10px] font-mono font-bold text-amber-400 uppercase tracking-widest">
                                SAYIM YAPILAN MALZEME
                              </span>
                              <h4 className="text-base font-black text-white mt-0.5">
                                {sayimItem.description || sayimItem.name || (sayimItem as any).MALZEME_ADI}
                              </h4>
                              <p className="text-xs text-slate-400 font-mono mt-0.5">
                                P/N: <strong className="text-white">{sayimItem.partNumber || sayimItem.pn || (sayimItem as any).MALZEME_PN || '-'}</strong> | Raf: <strong>{sayimItem.lokasyonNo || sayimItem.location || (sayimItem as any).DEPO_YERI || 'A-1'}</strong>
                              </p>
                            </div>

                            <button 
                              onClick={() => setSayimItem(null)}
                              className="p-2 text-slate-500 hover:text-white transition-colors cursor-pointer"
                            >
                              <X className="w-5 h-5" />
                            </button>
                          </div>

                          <div className="bg-slate-900/50 px-3 py-1.5 rounded-xl border border-slate-800 text-[10px] font-black text-slate-400 font-mono text-center">
                            {sayimDepoBolge} DEPOSU SAYIMI
                          </div>

                          {/* Comparison Grid */}
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                            {/* 1. Sistem Stoğu */}
                            <div className="p-3.5 bg-slate-900 border border-slate-700 rounded-2xl text-center">
                              <div className="text-[10px] font-bold text-slate-400 uppercase font-mono">
                                SİSTEMDEKİ MEVCUT
                              </div>
                              <div className="text-2xl font-black text-white font-mono mt-1">
                                {sistemStok} <span className="text-xs font-normal text-slate-400">EA</span>
                              </div>
                              <div className="text-[10px] text-slate-500 mt-0.5">Kayıtlı Stok Verisi</div>
                            </div>

                            {/* 2. Fiziksel Sayılan Girişi */}
                            <div className="p-3.5 bg-slate-900 border-2 border-emerald-500 rounded-2xl text-center">
                              <div className="text-[10px] font-black text-emerald-400 uppercase font-mono">
                                FİZİKSEL SAYILAN ADET *
                              </div>
                              <input
                                type="number"
                                min={0}
                                autoFocus
                                value={sayilanAdetInput}
                                onChange={(e) => setSayilanAdetInput(e.target.value)}
                                className="w-full bg-transparent text-2xl font-black text-white font-mono text-center mt-1 focus:outline-none placeholder:text-slate-700"
                                placeholder="0"
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') handleSaveSayimFarki();
                                  if (e.key === 'Escape') setSayimItem(null);
                                }}
                              />
                              <div className="text-[10px] text-emerald-500 font-bold mt-0.5">Giriş Bekleniyor</div>
                            </div>

                            {/* 3. Fark Hesaplama */}
                            <div className={`p-3.5 border-2 rounded-2xl text-center ${
                              fark === 0 ? 'bg-slate-900 border-slate-700' :
                              fark > 0 ? 'bg-indigo-950/30 border-indigo-500' :
                              'bg-rose-950/30 border-rose-500'
                            }`}>
                              <div className="text-[10px] font-bold text-slate-400 uppercase font-mono">
                                FARK (±)
                              </div>
                              <div className={`text-2xl font-black font-mono mt-1 ${
                                fark === 0 ? 'text-white' :
                                fark > 0 ? 'text-indigo-400' :
                                'text-rose-400'
                              }`}>
                                {fark > 0 ? `+${fark}` : fark} <span className="text-xs font-normal">EA</span>
                              </div>
                              <div className="text-[10px] text-slate-500 mt-0.5">
                                {fark === 0 ? 'Fark Yok' : fark > 0 ? 'Fazla Ürün' : 'Eksik Ürün'}
                              </div>
                            </div>
                          </div>

                          {/* S/N MANAGEMENT FOR SAYIM DIFFERENCE */}
                          {(() => {
                            const rawSnStr = sayimItem.serialAndNotes || sayimItem.sn || (sayimItem as any).SERI_NO || '';
                            const existingSns = rawSnStr.split(/[,/;]+/).map((s: string) => s.trim()).filter((s: string) => s && s !== '-');
                            return (
                              <div className="bg-slate-900 border border-slate-800 p-3 rounded-2xl space-y-2 text-xs font-mono">
                                <div className="flex items-center justify-between text-[10px] font-bold text-amber-400 uppercase">
                                  <span>DEPODAKİ MEVCUT SERİ NOLAR (S/N)</span>
                                  <span className="text-[9px] text-slate-400 font-sans">{existingSns.length} S/N Kayıtlı</span>
                                </div>
                                
                                {existingSns.length > 0 ? (
                                  <div className="flex flex-wrap gap-1 max-h-20 overflow-y-auto p-1.5 bg-slate-950 rounded-xl border border-slate-800">
                                    {existingSns.map((sNum: string, idx: number) => (
                                      <span key={idx} className="px-2 py-0.5 bg-slate-800 text-amber-300 rounded font-bold text-[11px] border border-slate-700">
                                        {sNum}
                                      </span>
                                    ))}
                                  </div>
                                ) : (
                                  <div className="text-[10px] text-slate-500 italic p-1">Sistemde kayıtlı seri no bulunmuyor (-)</div>
                                )}

                                {fark > 0 && (
                                  <div className="pt-1">
                                    <label className="block text-[10px] text-indigo-300 font-bold uppercase mb-1 font-sans">
                                      + FAZLA TESPİT EDİLEN YENİ SERİ NO (S/N) EKLE:
                                    </label>
                                    <input
                                      type="text"
                                      value={newSayimSerialInput}
                                      onChange={e => setNewSayimSerialInput(e.target.value)}
                                      placeholder="Eklemek istediğiniz yeni S/N'yi yazınız..."
                                      className="w-full bg-slate-950 border border-indigo-500/60 rounded-xl px-2.5 py-1.5 text-xs text-white font-mono focus:outline-none"
                                    />
                                  </div>
                                )}

                                {fark < 0 && existingSns.length > 0 && (
                                  <div className="pt-1">
                                    <label className="block text-[10px] text-rose-300 font-bold uppercase mb-1 font-sans">
                                      - EKSİK TESPİT EDİLEN ÇIKARILACAK SERİ NO (S/N) SEÇ:
                                    </label>
                                    <select
                                      value={missingSayimSerialSelect}
                                      onChange={e => setMissingSayimSerialSelect(e.target.value)}
                                      className="w-full bg-slate-950 border border-rose-500/60 rounded-xl px-2.5 py-1.5 text-xs text-rose-200 font-mono focus:outline-none"
                                    >
                                      <option value="">-- Depodan Çıkarılacak S/N Seçiniz --</option>
                                      {existingSns.map((sNum: string, idx: number) => (
                                        <option key={idx} value={sNum}>{sNum}</option>
                                      ))}
                                    </select>
                                  </div>
                                )}
                              </div>
                            );
                          })()}

                          {/* Info Banner */}
                          <div className="bg-amber-950/20 border border-amber-500/20 rounded-2xl p-3 flex items-start gap-3">
                            <AlertTriangle className="w-5 h-5 text-amber-400 shrink-0" />
                            <p className="text-[10px] text-slate-400 leading-relaxed italic">
                              Onayladığınızda; {fark === 0 ? 'mevcut miktar korunacak' : `envanter ${Math.abs(fark)} adet ${fark > 0 ? 'artırılacak' : 'azaltılacak'}`} ve bu işlem 
                              <strong> Hareket Geçmişi</strong> ile <strong>Form Kayıtları (Sayımlar)</strong> sayfasına anında işlenecektir.
                            </p>
                          </div>

                          <div className="flex items-center gap-3 pt-1">
                            <button
                              type="button"
                              onClick={() => setSayimItem(null)}
                              className="flex-1 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-2xl text-xs font-bold transition flex items-center justify-center gap-2 cursor-pointer"
                            >
                              <X className="w-4 h-4" />
                              İPTAL
                            </button>
                            <button
                              type="button"
                              onClick={handleSaveSayimFarki}
                              className={`flex-[2] py-3 rounded-2xl text-xs font-black uppercase tracking-wider transition shadow-lg flex items-center justify-center gap-2 cursor-pointer active:scale-95 ${
                                fark === 0 ? 'bg-emerald-600 hover:bg-emerald-500 text-white' :
                                fark > 0 ? 'bg-indigo-600 hover:bg-indigo-500 text-white' :
                                'bg-rose-600 hover:bg-rose-500 text-white'
                              }`}
                            >
                              <CheckCircle2 className="w-4 h-4" />
                              SAYIMI ONAYLA VE KAYDET
                            </button>
                          </div>
                        </div>
                      </div>
                    );
                  })()}

                  {/* Bu Oturumda Yapılan Sayımlar Geçmişi */}
                  {sayimGecmisi.length > 0 && (
                    <div className="bg-slate-950 p-4 rounded-3xl border border-slate-800 space-y-2">
                      <div className="text-xs font-bold text-slate-300 uppercase font-mono flex items-center justify-between">
                        <span>Bu Oturumda Yapılan Sayımlar ({sayimGecmisi.length})</span>
                        <span className="text-[10px] text-emerald-400">Hareket Geçmişine Aktarıldı</span>
                      </div>

                      <div className="max-h-48 overflow-y-auto space-y-1.5 divide-y divide-slate-800/80">
                        {sayimGecmisi.map((log, idx) => (
                          <div key={idx} className="pt-1.5 text-xs flex items-center justify-between">
                            <div>
                              <span className="font-bold text-white">{log.desc}</span>
                              <span className="font-mono text-slate-400 text-[10px] ml-2">P/N: {log.pn}</span>
                            </div>
                            <div className="flex items-center gap-3 font-mono">
                              <span className="text-slate-400 text-[11px]">{log.region}</span>
                              <span className="text-slate-300 text-[11px]">Sistem: {log.sistemAdet}</span>
                              <span className="text-emerald-300 font-bold text-[11px]">Sayılan: {log.sayilanAdet}</span>
                              <span className={`px-2 py-0.5 rounded text-[10px] font-black ${
                                log.fark === 0 ? 'bg-slate-800 text-slate-300' : (log.fark < 0 ? 'bg-rose-950 text-rose-300' : 'bg-blue-950 text-blue-300')
                              }`}>
                                Fark: {log.fark > 0 ? '+' : ''}{log.fark}
                              </span>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                </div>
              )}

            </div>
          )}

        </div>

        {/* FOOTER */}
        <div className="bg-slate-950 px-5 py-3.5 border-t border-slate-800 flex items-center justify-between">
          <div className="text-[11px] text-slate-400 font-mono">
            {isUnlocked ? '✓ Terminal Aktif • OGM Havacılık Depo Lojistik' : 'Korumalı Modül • Şifre Gereklidir'}
          </div>

          <button
            type="button"
            onClick={onClose}
            className="px-5 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold transition-colors cursor-pointer"
          >
            Kapat
          </button>
        </div>

      </div>

      {/* MODALS SIBLINGS (REMAINING ITEMS & EXCESS PROMPT) */}
      {showRemainingModal && (
        <div className="fixed inset-0 z-[10200] flex items-center justify-center p-2 sm:p-4 bg-black/85 sm:backdrop-blur-sm">
          <div className="bg-slate-900 p-4 sm:p-6 rounded-2xl sm:rounded-3xl border-2 border-slate-700 shadow-2xl max-w-lg w-full flex flex-col max-h-[88dvh]">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3 mb-4">
              <h4 className="text-sm font-black text-white uppercase tracking-widest">OKUTULMAYAN ÜRÜNLER (KALANLAR)</h4>
              <button onClick={() => setShowRemainingModal(false)} className="text-slate-400 hover:text-white text-lg">✕</button>
            </div>
            
            <div className="overflow-y-auto overscroll-contain space-y-2 flex-1">
              {sessionItemsToCount
                .filter(i => !scannedInSession.has(getCountItemKey(i)))
                .map((item, idx) => (
                  <div 
                    key={idx} 
                    onClick={() => {
                      setSayimItem(item);
                      const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
                      setSayilanAdetInput(Number((item as any)[regionProp] ?? 0));
                      setShowRemainingModal(false);
                    }}
                    className="p-3 bg-slate-950 rounded-xl border border-slate-800 flex items-center justify-between gap-3 cursor-pointer hover:bg-slate-900 transition-colors group"
                  >
                    <div>
                      <div className="text-[11px] font-bold text-white uppercase line-clamp-1 group-hover:text-amber-400 transition-colors">{item.description || item.name || (item as any).MALZEME_ADI}</div>
                      <div className="text-[10px] font-mono text-slate-500">P/N: {item.partNumber || item.pn || (item as any).MALZEME_PN || '-'}</div>
                    </div>
                    <div className="text-right">
                      <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">SİSTEM STOK</div>
                      {(() => {
                        const stockVal = Number((item as any)[DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut'] || 0);
                        return (
                          <div className={`text-xs font-black ${
                            stockVal <= 0 
                              ? 'bg-rose-600 text-white px-2 py-0.5 rounded-md animate-pulse shadow-xs font-mono font-bold' 
                              : 'text-amber-500'
                          }`}>
                            {stockVal} EA
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                ))
              }
              {sessionTotalCount - scannedInSession.size === 0 && (
                <div className="text-center py-10 text-slate-500 font-bold">Tüm ürünler okutuldu! 🎉</div>
              )}
            </div>
            
            <button 
              onClick={() => setShowRemainingModal(false)}
              className="mt-6 w-full py-3 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold transition shadow-md active:scale-95"
            >
              Kapat
            </button>
          </div>
        </div>
      )}

      {/* ─── ÜRÜN TALEBİ (ÇIKAN TALEBİ) MODALI — SADE 3 ADIMLI SİHİRBAZ ─── */}
      {isUrunTalebiModalOpen && (
        <div className="fixed inset-0 z-[10500] bg-black/90 sm:backdrop-blur-xs flex items-center justify-center p-1.5 sm:p-4 overflow-y-auto overscroll-contain">
          <div className="bg-slate-900 border border-amber-500/40 text-white rounded-2xl sm:rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[95dvh] sm:max-h-[92vh]">
            {/* Header */}
            <div className="bg-slate-950 px-4 py-3 sm:px-5 sm:py-3.5 border-b border-slate-800 flex items-center justify-between gap-2 shrink-0">
              <div className="flex items-center gap-2.5 min-w-0">
                <div className="p-2 bg-amber-500/15 text-amber-400 rounded-xl border border-amber-500/30 shrink-0">
                  <Package className="w-4 h-4" />
                </div>
                <div className="min-w-0">
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-amber-200 truncate">
                    ÜRÜN TALEBİ OLUŞTUR — {UNITS_SECURITY.find(u => u.id === selectedUnit)?.name || selectedUnit.toUpperCase()}
                  </h3>
                  <p className="text-[10px] text-slate-400 truncate">
                    Onaylanmadan stoktan düşülmez
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  stopTalepCamera();
                  setIsUrunTalebiModalOpen(false);
                }}
                className="text-slate-400 hover:text-white font-bold cursor-pointer p-1.5 shrink-0"
              >
                ✕
              </button>
            </div>

            {/* Sade Stepper Çubuğu */}
            <div className="px-4 py-2.5 bg-slate-950/60 border-b border-slate-800 flex items-center justify-between gap-2 shrink-0">
              {[
                { step: 1 as const, label: '1. Parça Seçimi' },
                { step: 2 as const, label: '2. Talep Bilgileri' },
                { step: 3 as const, label: '3. Onay & Gönder' }
              ].map((s, idx) => {
                const isActive = talepStep === s.step;
                const isDone = talepStep > s.step;
                return (
                  <React.Fragment key={s.step}>
                    <button
                      type="button"
                      onClick={() => {
                        if (s.step === 1) setTalepStep(1);
                        else if (s.step === 2 && talepItemList.length > 0) setTalepStep(2);
                        else if (s.step === 3 && talepItemList.length > 0 && talepWorkOrderOrCrs.trim() && talepRequestedBy.trim()) setTalepStep(3);
                      }}
                      className={`flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-[10.5px] font-bold transition cursor-pointer ${
                        isActive
                          ? 'bg-amber-500 text-slate-950 shadow-sm font-black'
                          : isDone
                          ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                          : 'bg-slate-800/80 text-slate-400'
                      }`}
                    >
                      <span>{s.label}</span>
                    </button>
                    {idx < 2 && <div className={`flex-1 h-0.5 rounded-full ${talepStep > s.step ? 'bg-emerald-500/50' : 'bg-slate-800'}`} />}
                  </React.Fragment>
                );
              })}
            </div>

            <form
              onSubmit={(e) => {
                if (talepStep !== 3) {
                  e.preventDefault();
                  return;
                }
                handleTalepSubmit(e);
              }}
              className="p-4 sm:p-5 space-y-4 text-xs overflow-y-auto overscroll-contain flex-1 font-sans flex flex-col justify-between"
            >
              <div className="space-y-4">
                {/* ─── ADIM 1: PARÇA SEÇİMİ ─── */}
                {talepStep === 1 && (
                  <div className="space-y-3 animate-fade-in">
                    <div className="flex items-center justify-between gap-2">
                      <div>
                        <div className="text-xs font-black text-white">Parça Listesi ({talepItemList.length})</div>
                        <div className="text-[10.5px] text-slate-400">Miktar, depo ve varsa seri numarasını kontrol edin</div>
                      </div>
                      <button
                        type="button"
                        onClick={() => setIsAddTalepItemOpen(prev => !prev)}
                        className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-500 hover:bg-amber-400 text-slate-950 font-black rounded-xl text-xs transition cursor-pointer shadow-xs active:scale-95"
                      >
                        <span>{isAddTalepItemOpen ? '✕ Kapat' : '+ Parça / Barkod Ekle'}</span>
                      </button>
                    </div>

                    {isAddTalepItemOpen && (
                      <div className="p-3 bg-slate-950 rounded-2xl border border-amber-500/40 space-y-2.5 animate-fade-in">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <span className="text-[10px] font-black text-emerald-300 uppercase font-mono flex items-center gap-1.5">
                            <Camera className="w-3.5 h-3.5 text-emerald-400" />
                            <span>Kamera ile Okut</span>
                          </span>
                          <div className="flex items-center gap-1.5">
                            <select
                              value={selectedTalepCameraId}
                              onChange={(e) => {
                                const nextCam = e.target.value;
                                setSelectedTalepCameraId(nextCam);
                                if (isTalepCameraActive) startTalepCamera(nextCam);
                              }}
                              className="px-2 py-1 bg-slate-900 border border-slate-700 rounded-lg text-[10px] font-bold text-slate-200 max-w-[150px]"
                            >
                              <option value="auto">Otomatik (Arka Kamera)</option>
                              {talepAvailableCameras.map((cam) => (
                                <option key={cam.id} value={cam.id}>{cam.label}</option>
                              ))}
                            </select>
                            {isTalepCameraActive ? (
                              <button
                                type="button"
                                onClick={() => stopTalepCamera()}
                                className="px-2.5 py-1 bg-rose-600 hover:bg-rose-500 text-white text-[10px] font-black rounded-lg cursor-pointer"
                              >
                                Kapat
                              </button>
                            ) : (
                              <button
                                type="button"
                                onClick={() => startTalepCamera()}
                                className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white text-[10px] font-black rounded-lg cursor-pointer"
                              >
                                Aç
                              </button>
                            )}
                          </div>
                        </div>

                        <div className={`relative rounded-xl overflow-hidden bg-black/70 border border-slate-800 ${isTalepCameraActive || talepCameraError ? 'min-h-[170px] flex items-center justify-center' : 'hidden'}`}>
                          <div id={talepScannerContainerId} className="w-full max-w-xs mx-auto overflow-hidden" />
                          {talepCameraError && (
                            <div className="absolute inset-0 bg-black/90 flex flex-col items-center justify-center p-3 text-center text-[11px] text-rose-300">
                              <span>{talepCameraError}</span>
                            </div>
                          )}
                        </div>

                        <div className="relative">
                          <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
                          <input
                            ref={addTalepInputRef}
                            type="text"
                            value={addTalepSearchQuery}
                            onChange={(e) => setAddTalepSearchQuery(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                e.stopPropagation();
                                const q = addTalepSearchQuery.trim();
                                if (!q) return;
                                const { found } = resolveInventoryItemByCode(q);
                                if (found) {
                                  playBeep();
                                  stopTalepCamera();
                                  handleAddTalepItemFromInventory(found);
                                }
                              }
                            }}
                            placeholder="P/N veya malzeme adı yazınız..."
                            className="w-full pl-8 pr-3 py-2 bg-slate-900 border border-slate-700 rounded-xl text-white text-xs font-mono font-bold focus:outline-none focus:border-amber-400"
                          />
                        </div>

                        <div className="max-h-36 overflow-y-auto divide-y divide-slate-800 rounded-xl bg-slate-900/60 border border-slate-800">
                          {talepAddMatches.length === 0 ? (
                            <div className="p-3 text-center text-[11px] text-slate-400">Eşleşen malzeme bulunamadı.</div>
                          ) : (
                            talepAddMatches.map((item, idx) => (
                              <div
                                key={idx}
                                onClick={() => {
                                  playBeep();
                                  stopTalepCamera();
                                  handleAddTalepItemFromInventory(item);
                                }}
                                className="p-2 hover:bg-amber-950/50 cursor-pointer text-xs flex items-center justify-between gap-2"
                              >
                                <div className="min-w-0 flex-1">
                                  <div className="font-bold text-white truncate">{item.description || item.name || (item as any).MALZEME_ADI}</div>
                                  <div className="text-[10px] font-mono text-slate-400 truncate">
                                    P/N: <strong className="text-amber-300">{item.partNumber || item.pn || (item as any).MALZEME_PN || '-'}</strong>
                                  </div>
                                </div>
                                <span className="text-[10px] font-black bg-emerald-600/20 text-emerald-300 border border-emerald-600/40 px-2 py-0.5 rounded-md shrink-0">
                                  + Ekle
                                </span>
                              </div>
                            ))
                          )}
                        </div>
                      </div>
                    )}

                    {talepItemList.length > 0 ? (
                      <div className="space-y-2.5 max-h-64 overflow-y-auto pr-1">
                        {talepItemList.map((it) => (
                          <div key={it.id} className="p-3 bg-slate-950 rounded-2xl border border-slate-800 space-y-2.5">
                            <div className="flex items-start justify-between gap-2">
                              <div className="min-w-0">
                                <div className="text-xs font-black text-white truncate">{it.itemName}</div>
                                <div className="text-[10px] font-mono text-amber-300 mt-0.5">
                                  P/N: {it.pn || '-'} &bull; {it.category === 'kimyasal' ? 'Kimyasal' : 'Sarf'}
                                </div>
                              </div>
                              <button
                                type="button"
                                onClick={() => handleRemoveTalepItem(it.id)}
                                className="text-rose-400 hover:text-rose-200 font-bold px-1.5 py-0.5 hover:bg-rose-950/50 rounded-lg cursor-pointer"
                                title="Çıkar"
                              >
                                ✕
                              </button>
                            </div>

                            <div className={`grid grid-cols-1 ${it.category === 'kimyasal' ? 'sm:grid-cols-2' : 'sm:grid-cols-3'} gap-2`}>
                              <div>
                                <label className="block text-[9.5px] font-bold text-slate-400 mb-1">Adet</label>
                                <input
                                  type="number"
                                  min={1}
                                  value={it.quantity}
                                  onChange={(e) => handleUpdateTalepItemField(it.id, 'quantity', Math.max(1, parseInt(e.target.value) || 1))}
                                  className="w-full px-2.5 py-1.5 bg-slate-900 border border-slate-700 rounded-xl text-white font-mono font-bold"
                                />
                              </div>
                              <div>
                                <label className="block text-[9.5px] font-bold text-slate-400 mb-1">Depo</label>
                                <select
                                  value={it.depot}
                                  onChange={(e) => handleUpdateTalepItemField(it.id, 'depot', e.target.value)}
                                  className="w-full px-2.5 py-1.5 bg-slate-900 border border-slate-700 rounded-xl text-white font-bold"
                                >
                                  <option value="ANKARA">ANKARA</option>
                                  <option value="KARAİN">KARAİN</option>
                                  <option value="ÇANAKKALE">ÇANAKKALE</option>
                                  <option value="MİLAS">MİLAS</option>
                                  <option value="BURSA">BURSA</option>
                                </select>
                              </div>
                              {it.category !== 'kimyasal' && (
                                <div>
                                  <label className="block text-[9.5px] font-bold text-slate-400 mb-1">Seri No (S/N)</label>
                                  <input
                                    type="text"
                                    value={it.sn}
                                    onChange={(e) => handleUpdateTalepItemField(it.id, 'sn', e.target.value)}
                                    placeholder="Opsiyonel..."
                                    className="w-full px-2.5 py-1.5 bg-slate-900 border border-slate-700 rounded-xl text-white font-mono font-bold"
                                  />
                                </div>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="p-6 bg-slate-950/60 rounded-2xl border border-dashed border-slate-800 text-center text-slate-400 text-xs">
                        Henüz parça eklenmedi. Yukarıdaki <strong>+ Parça / Barkod Ekle</strong> butonuna tıklayın.
                      </div>
                    )}
                  </div>
                )}

                {/* ─── ADIM 2: TALEP BİLGİLERİ ─── */}
                {talepStep === 2 && (
                  <div className="space-y-3.5 animate-fade-in">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-300 uppercase mb-1">
                        Kuyruk No *
                      </label>
                      <select
                        value={talepTailNo}
                        onChange={(e) => {
                          setTalepTailNo(e.target.value);
                          if (e.target.value !== 'Diğer') setTalepTailNoOther('');
                        }}
                        className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-bold focus:outline-none focus:border-amber-500"
                      >
                        {(AIRCRAFT_TAILS.at802 || [
                          'ORMAN 21 (OR-2021) - AT-802', 'ORMAN 22 (OR-2022) - AT-802', 'ORMAN 23 (OR-2023) - AT-802',
                          'ORMAN 24 (OR-2024) - AT-802', 'ORMAN 25 (OR-2025) - AT-802', 'ORMAN 26 (OR-2026) - AT-802',
                          'ORMAN 27 (OR-2027) - AT-802', 'ORMAN 28 (OR-2028) - AT-802', 'ORMAN 29 (OR-2029) - AT-802',
                          'ORMAN 30 (OR-2030) - AT-802', 'ORMAN 31 (OR-2031) - AT-802', 'ORMAN 36 (OR-2036) - AT-802',
                          'ORMAN 37 (OR-2037) - AT-802', 'ORMAN 38 (OR-2038) - AT-802', 'ORMAN 39 (OR-2039) - AT-802',
                          'ORMAN 40 (OR-2040) - AT-802'
                        ]).map((tail, idx) => (
                          <option key={idx} value={tail}>{tail}</option>
                        ))}
                        <option value="GENEL">GENEL / DEPO</option>
                        <option value="Diğer">Diğer (Manuel Giriş)...</option>
                      </select>
                      {talepTailNo === 'Diğer' && (
                        <input
                          type="text"
                          value={talepTailNoOther}
                          onChange={(e) => setTalepTailNoOther(e.target.value)}
                          placeholder="Örn: OR-2045"
                          className="w-full mt-2 px-3 py-2 bg-slate-950 border border-amber-500/50 rounded-xl text-amber-200 font-mono text-xs font-bold focus:outline-none focus:border-amber-400"
                        />
                      )}
                    </div>

                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <div className="flex items-center gap-1.5">
                          <label className="text-[10px] font-bold text-amber-300 uppercase">
                            İş Emri / CRS No / Defter Sıra No *
                          </label>
                          <button
                            type="button"
                            onClick={() => setShowWorkOrderInfo(v => !v)}
                            className="inline-flex items-center justify-center w-4 h-4 rounded-full bg-amber-500/20 hover:bg-amber-500/40 text-amber-300 border border-amber-400/50 cursor-pointer transition-colors"
                            title="Bilgi"
                          >
                            <Info className="w-2.5 h-2.5" />
                          </button>
                        </div>
                        <span className="text-[9px] text-amber-400">3'ünden Biri Zorunlu</span>
                      </div>
                      {showWorkOrderInfo && (
                        <div className="mb-2 p-2.5 rounded-xl bg-amber-950/70 border border-amber-500/40 text-[10.5px] text-amber-100 leading-relaxed">
                          Hava aracı eğer ki yıllık bakımda ise <strong>İş Emri veya CRS No</strong>; periyodik başka bir bakım ise <strong>CRS No</strong>; eğer ki hava aracı hat bakım sürecinde (uçuş öncesi veya sonrası bakımda) tespit edilen arıza sonucu ise o arızanın gerçekleştiği uçuşun <strong>Defter Sıra No</strong> yazılır.
                        </div>
                      )}
                      <input
                        type="text"
                        value={talepWorkOrderOrCrs}
                        onChange={(e) => setTalepWorkOrderOrCrs(e.target.value)}
                        placeholder="İş Emri, CRS No veya Defter Sıra No..."
                        className="w-full px-3 py-2 bg-slate-950 border border-amber-500/50 rounded-xl text-white font-mono text-xs font-bold focus:outline-none focus:border-amber-400"
                      />
                    </div>

                    <div>
                      <label className="block text-[10px] font-bold text-slate-300 uppercase mb-1">
                        Teslim Alan Personel *
                      </label>
                      <PersonnelAutocomplete
                        value={talepRequestedBy}
                        onChange={setTalepRequestedBy}
                        placeholder="Personel seçiniz veya yazınız..."
                      />
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <div className="flex items-center justify-between mb-1">
                          <label className="text-[10px] font-bold text-slate-300 uppercase">Tarih *</label>
                          <button
                            type="button"
                            onClick={() => {
                              const pad = (n: number) => String(n).padStart(2, '0');
                              const now = new Date();
                              setTalepDate(`${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()}`);
                            }}
                            className="text-[9.5px] text-amber-400 hover:text-amber-300 font-bold underline cursor-pointer"
                          >
                            Bugün
                          </button>
                        </div>
                        <input
                          type="date"
                          value={(() => {
                            if (!talepDate) return '';
                            const dateOnly = talepDate.split(' ')[0];
                            if (/^\d{4}-\d{2}-\d{2}$/.test(dateOnly)) return dateOnly;
                            const dParts = dateOnly.split('.');
                            if (dParts.length === 3) {
                              return `${dParts[2]}-${dParts[1].padStart(2, '0')}-${dParts[0].padStart(2, '0')}`;
                            }
                            return '';
                          })()}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (!val) {
                              setTalepDate('');
                              return;
                            }
                            const [y, m, day] = val.split('-');
                            if (y && m && day) {
                              setTalepDate(`${day}.${m}.${y}`);
                            } else {
                              setTalepDate(val);
                            }
                          }}
                          className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono text-[11px] focus:outline-none focus:border-amber-500 color-scheme-dark"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-300 uppercase mb-1">
                          Not (Opsiyonel)
                        </label>
                        <input
                          type="text"
                          value={talepNotes}
                          onChange={(e) => setTalepNotes(e.target.value)}
                          placeholder="Bakım / değişim notu..."
                          className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white text-xs focus:outline-none focus:border-amber-500"
                        />
                      </div>
                    </div>
                  </div>
                )}

                {/* ─── ADIM 3: ÖZET & GÖNDER ─── */}
                {talepStep === 3 && (
                  <div className="space-y-3 animate-fade-in">
                    <div className="p-3.5 bg-slate-950 rounded-2xl border border-slate-800 space-y-2">
                      <div className="text-[10.5px] font-black text-amber-400 uppercase font-mono">
                        Talep Edilen Parçalar ({talepItemList.length})
                      </div>
                      <div className="divide-y divide-slate-800/80 max-h-40 overflow-y-auto">
                        {talepItemList.map((it) => (
                          <div key={it.id} className="py-1.5 flex items-center justify-between gap-2 text-xs">
                            <div className="min-w-0">
                              <div className="font-bold text-white truncate">{it.itemName}</div>
                              <div className="text-[10px] font-mono text-slate-400">
                                P/N: {it.pn || '-'} &bull; Depo: {it.depot}{it.sn ? ` • S/N: ${it.sn}` : ''}
                              </div>
                            </div>
                            <span className="px-2 py-0.5 rounded-lg bg-amber-500/20 text-amber-300 font-mono font-black shrink-0">
                              {it.quantity} Adet
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div className="p-3.5 bg-slate-950 rounded-2xl border border-slate-800 space-y-1.5 text-xs">
                      <div className="flex justify-between gap-2">
                        <span className="text-slate-400">Kuyruk No:</span>
                        <span className="font-bold text-white text-right">{talepTailNo === 'Diğer' ? (talepTailNoOther || '-') : talepTailNo}</span>
                      </div>
                      <div className="flex justify-between gap-2">
                        <span className="text-slate-400">İş Emri / CRS / Defter No:</span>
                        <span className="font-mono font-black text-amber-300 text-right">{talepWorkOrderOrCrs}</span>
                      </div>
                      <div className="flex justify-between gap-2">
                        <span className="text-slate-400">Teslim Alan:</span>
                        <span className="font-bold text-white text-right">{talepRequestedBy}</span>
                      </div>
                      <div className="flex justify-between gap-2">
                        <span className="text-slate-400">Tarih:</span>
                        <span className="font-mono text-slate-200 text-right">{talepDate || '-'}</span>
                      </div>
                      {talepNotes.trim() && (
                        <div className="flex justify-between gap-2 pt-1 border-t border-slate-800/80">
                          <span className="text-slate-400">Not:</span>
                          <span className="text-slate-200 text-right">{talepNotes}</span>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>

              {/* Footer Navigasyon */}
              <div className="pt-3 border-t border-slate-800 flex items-center justify-between gap-2 shrink-0 font-sans">
                <button
                  type="button"
                  onClick={() => {
                    if (talepStep === 1) {
                      stopTalepCamera();
                      setIsUrunTalebiModalOpen(false);
                    } else {
                      setTalepStep(prev => (prev === 3 ? 2 : 1));
                    }
                  }}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl font-bold transition cursor-pointer"
                >
                  {talepStep === 1 ? 'İptal' : '← Geri'}
                </button>

                {talepStep === 1 && (
                  <button
                    type="button"
                    disabled={talepItemList.length === 0}
                    onClick={() => {
                      stopTalepCamera();
                      setTalepStep(2);
                    }}
                    className="px-5 py-2 bg-amber-500 hover:bg-amber-400 disabled:opacity-40 text-slate-950 font-black uppercase rounded-xl shadow transition cursor-pointer"
                  >
                    İleri →
                  </button>
                )}

                {talepStep === 2 && (
                  <button
                    type="button"
                    onClick={() => {
                      if (!talepWorkOrderOrCrs.trim()) {
                        showNotification?.('⚠️ Lütfen İş Emri, CRS No veya Defter Sıra No giriniz.');
                        return;
                      }
                      if (talepTailNo === 'Diğer' && !talepTailNoOther.trim()) {
                        showNotification?.('⚠️ Lütfen manuel kuyruk numarasını giriniz.');
                        return;
                      }
                      if (!talepRequestedBy.trim()) {
                        showNotification?.('⚠️ Lütfen teslim alacak personeli seçiniz.');
                        return;
                      }
                      setTalepStep(3);
                    }}
                    className="px-5 py-2 bg-amber-500 hover:bg-amber-400 text-slate-950 font-black uppercase rounded-xl shadow transition cursor-pointer"
                  >
                    Özete Geç →
                  </button>
                )}

                {talepStep === 3 && (
                  <button
                    type="submit"
                    disabled={isTalepSubmitting}
                    className="px-5 py-2 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-400 disabled:opacity-50 text-slate-950 font-black uppercase rounded-xl shadow-lg transition cursor-pointer active:scale-95 flex items-center gap-2"
                  >
                    <Send className="w-4 h-4 text-slate-950" />
                    <span>{isTalepSubmitting ? 'Kaydediliyor...' : 'TALEBİ KAYDET VE İLET'}</span>
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}

      {isExcessPromptOpen && (
        <div className="fixed inset-0 z-[10300] flex items-center justify-center p-4 bg-black/90 backdrop-blur-md animate-fade-in">
          <div className="bg-slate-900 p-8 rounded-[40px] border-4 border-amber-500 shadow-[0_0_50px_rgba(245,158,11,0.3)] max-w-md w-full text-center space-y-6">
            <div className="w-20 h-20 bg-amber-500 rounded-full flex items-center justify-center mx-auto shadow-lg">
              <AlertTriangle className="w-12 h-12 text-slate-900" />
            </div>
            
            <div>
              <h4 className="text-xl font-black text-white uppercase tracking-tighter">FAZLA ÜRÜN TESPİT EDİLDİ!</h4>
              <p className="text-sm text-slate-400 mt-2">
                <strong className="text-amber-400">{excessItemsQueue[0]?.item?.description || excessItemsQueue[0]?.item?.name || (excessItemsQueue[0]?.item as any)?.MALZEME_ADI || 'Ürün'}</strong> 
                için sistemden 
                <strong className="text-emerald-400"> +{excessItemsQueue[0]?.fark} EA</strong> fazla sayım yapıldı.
              </p>
              <p className="text-xs text-slate-500 mt-4 font-bold uppercase tracking-widest">BU ÜRÜNÜN KAYNAĞI NEDİR?</p>
            </div>

            <div className="grid grid-cols-1 gap-3 pt-2">
              <button
                onClick={() => resolveNextExcess('ANKARA')}
                className="w-full py-4 bg-amber-600 hover:bg-amber-500 text-white rounded-2xl text-xs font-black uppercase tracking-widest transition shadow-lg active:scale-95 flex items-center justify-center gap-3"
              >
                <MapPin className="w-4 h-4" />
                ANKARA MERKEZDEN GELDİ
              </button>
              
              <button
                onClick={() => resolveNextExcess('BOLGE')}
                className="w-full py-4 bg-slate-800 hover:bg-slate-700 text-white rounded-2xl text-xs font-black uppercase tracking-widest transition border border-slate-700 active:scale-95 flex items-center justify-center gap-3"
              >
                <RefreshCw className="w-4 h-4" />
                BÖLGE MÜDÜRLÜĞÜ TEMİNİ
              </button>
            </div>
            
            <p className="text-[10px] text-slate-500 italic">
              * Bu seçim "SAYIM TESPİT TRANSFER" olarak {sayimDepoBolge} depo kayıtlarına işlenecektir.
            </p>
          </div>
        </div>
      )}
    </div>
  );
};
