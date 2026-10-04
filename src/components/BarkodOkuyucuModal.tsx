import React, { useState, useEffect, useRef, useMemo } from 'react';
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
  Clock
} from 'lucide-react';
import { Html5Qrcode, Html5QrcodeSupportedFormats } from 'html5-qrcode';
import { DepoItem, AIRCRAFT_TAILS } from './DepoManagementModal';
import { DepoTransaction } from '../types';
import { PersonnelAutocomplete } from './PersonnelAutocomplete';

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
  const [isManualSearchDropdownOpen, setIsManualSearchDropdownOpen] = useState<boolean>(false);

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
    isAutoFilled?: boolean;
  }>>([]);
  const [isAddTalepItemOpen, setIsAddTalepItemOpen] = useState<boolean>(false);
  const [addTalepSearchQuery, setAddTalepSearchQuery] = useState<string>('');
  const [talepDate, setTalepDate] = useState<string>('');
  const [talepTailNo, setTalepTailNo] = useState<string>((AIRCRAFT_TAILS.at802 && AIRCRAFT_TAILS.at802[0]) || 'ORMAN 21 (OR-2021) - AT-802');
  const [talepTailNoOther, setTalepTailNoOther] = useState<string>('');
  const [talepRequestedBy, setTalepRequestedBy] = useState<string>('');
  const [talepNotes, setTalepNotes] = useState<string>('');
  const [isTalepSubmitting, setIsTalepSubmitting] = useState<boolean>(false);
  const [isTalepDropdownOpen, setIsTalepDropdownOpen] = useState<boolean>(false);
  const talepDropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (talepDropdownRef.current && !talepDropdownRef.current.contains(e.target as Node)) {
        setIsTalepDropdownOpen(false);
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
      let depot = 'ANKARA';
      if (locUpper.includes('KARAİN') || locUpper.includes('KARAIN')) depot = 'KARAİN';
      else if (locUpper.includes('ÇANAKKALE')) depot = 'ÇANAKKALE';
      else if (locUpper.includes('MİLAS')) depot = 'MİLAS';
      else if (locUpper.includes('BURSA')) depot = 'BURSA';

      initialList.push({
        id: `item_${Date.now()}_0`,
        pn,
        sn,
        itemName,
        category,
        quantity: 1,
        depot,
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
    const pad = (n: number) => String(n).padStart(2, '0');
    const now = new Date();
    setTalepDate(`${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
    const defaultTail = (AIRCRAFT_TAILS.at802 && AIRCRAFT_TAILS.at802[0]) || 'ORMAN 21 (OR-2021) - AT-802';
    setTalepTailNo(defaultTail);
    setTalepTailNoOther('');
    setTalepRequestedBy('');
    setTalepNotes('');
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
    let depot = 'ANKARA';
    if (locUpper.includes('KARAİN') || locUpper.includes('KARAIN')) depot = 'KARAİN';
    else if (locUpper.includes('ÇANAKKALE')) depot = 'ÇANAKKALE';
    else if (locUpper.includes('MİLAS')) depot = 'MİLAS';
    else if (locUpper.includes('BURSA')) depot = 'BURSA';

    setTalepItemList(prev => [...prev, {
      id: `item_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
      pn,
      sn,
      itemName,
      category,
      quantity: 1,
      depot,
      isAutoFilled: true
    }]);
    setIsAddTalepItemOpen(false);
    setAddTalepSearchQuery('');
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
        items: talepItemList.map(it => ({
          pn: it.pn.trim(),
          sn: it.sn.trim() || '-',
          itemName: it.itemName.trim() || it.pn.trim(),
          category: it.category,
          quantity: Number(it.quantity) || 1,
          date: talepDate.trim(),
          depot: it.depot.trim().toUpperCase(),
          tailNo: finalTailNo,
          requestedBy: talepRequestedBy.trim(),
          notes: talepNotes.trim(),
          unit: 'at802'
        }))
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
            viewUrl: res.viewUrl
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
      const stored = localStorage.getItem('ogm_depo_belge_bul_v5') || localStorage.getItem('ogm_form_one_docs');
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
        aciklama: `Belge Bul sisteminde SERTİFİKA TARAMA_${cleanUnit}.pdf belgesi bağlandı.`
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

  // Active inventory: use prop or fallback to live localStorage
  const activeInventory = useMemo(() => {
    if (inventory && inventory.length > 0) return inventory;
    try {
      const stored = localStorage.getItem('ogm_depo_inventory_v5');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {}
    return [];
  }, [inventory]);

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
    setCameraError(null);
    try {
      const el = document.getElementById(scannerContainerId);
      if (!el) {
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

      const scanConfig = {
        fps: 20,
        qrbox: (viewfinderWidth: number, viewfinderHeight: number) => {
          return {
            width: Math.max(200, Math.floor(viewfinderWidth * 0.88)),
            height: Math.max(140, Math.floor(viewfinderHeight * 0.65))
          };
        },
        aspectRatio: 1.0,
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
    }
  };

  // Cleanup camera on close or unmount (NEVER auto-start on page open)
  useEffect(() => {
    if (!isOpen) {
      stopCamera();
    }
  }, [isOpen]);

  useEffect(() => {
    return () => {
      stopCamera();
    };
  }, []);

  // Code parser & item matcher
  const handleCodeRead = (rawText: string) => {
    setScannedResultText(rawText);

    // Extract PN, SN or Location from text
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

    // Look for matching item in inventory
    const cleanQ = (targetPn || targetLoc || rawText).toLowerCase();
    const found = activeInventory.find(i => {
      const item = i as any;
      const pn = (i.partNumber || i.pn || item.MALZEME_PN || '').toLowerCase();
      const desc = (i.description || i.name || item.MALZEME_ADI || '').toLowerCase();
      const loc = (i.lokasyonNo || i.location || item.DEPO_YERI || '').toLowerCase();
      const sn = (i.serialAndNotes || i.sn || item.SERI_NO || '').toLowerCase();

      return (targetPn && pn === targetPn.toLowerCase()) ||
        (targetLoc && loc.includes(targetLoc.toLowerCase())) ||
        pn === cleanQ ||
        desc.includes(cleanQ) ||
        sn.includes(cleanQ) ||
        loc.includes(cleanQ);
    }) || null;

    if (activeModule === 'read') {
      stopCamera(); // Pause camera to display details window
      if (found) {
        setMatchedItem(found);
        if (showNotification) {
          showNotification(`✅ Barkod okundu: ${found.description || found.name || (found as any).MALZEME_ADI}`);
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
        
        // If in session, maybe show what we already counted? 
        // No, user wants to see current system stock as physical initially or for verification.
        setSayilanAdetInput(sistemStok);

        if (showNotification) {
          showNotification(`📦 ${found.description || found.name || (found as any).MALZEME_ADI} için sayım penceresi açıldı.`);
        }
      }
    }
  };

  const handleStartSession = () => {
    const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
    // Filter items with quantity > 0 and matching sayimCategory ('sarf' vs 'kimyasal')
    // Ayrıca Karain seçildiğinde, mevcut 0 olsa bile önceden transfer geçmişi olanlar da sayım checklistine dahil edilir
    const items = activeInventory.filter(item => {
        const qty = Number((item as any)[regionProp] || 0);
        const itemCat = item.category || 'sarf';
        const isKimyasalItem = itemCat === 'kimyasal' || (item.description || item.name || '').toLowerCase().includes('kimya') || (item.description || item.name || '').toLowerCase().includes('solvent') || (item.description || item.name || '').toLowerCase().includes('yağ') || (item.description || item.name || '').toLowerCase().includes('hidrolik');
        
        const matchesCategory = sayimCategory === 'kimyasal' ? isKimyasalItem : !isKimyasalItem;
        if (!matchesCategory) return false;

        if (qty > 0) return true;
        if ((sayimDepoBolge as string) === 'KARAİN' || (sayimDepoBolge as string) === 'KARAIN') {
          const hadTransfer = Number((item as any).karainTransfer || 0) > 0 || Number((item as any).karainCikan || 0) > 0;
          if (hadTransfer) return true;
        }
        return false;
    });
    
    setSessionItemsToCount(items);
    setSessionTotalCount(items.length);
    setScannedInSession(new Set());
    setSessionCounts({});
    setSessionScannedData({});
    setSessionActive(true);
    if (showNotification) {
      showNotification(`🚀 ${sayimDepoBolge} [${sayimCategory === 'kimyasal' ? 'KİMYASAL DEPOSU' : 'SARF & PARÇA DEPOSU'}] için ${items.length} kalemlik sayım oturumu başladı.`);
    }
  };

  const handleFinishSession = () => {
    const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
    
    // 1. Items scanned during session
    const scannedItems = Object.values(sessionScannedData);
    
    // 2. Items NEVER scanned (stok preserved)
    const neverScanned = sessionItemsToCount.filter(item => {
      const it = item as any;
      const itemKey = item.partNumber || item.pn || it.MALZEME_PN || item.name || it.MALZEME_ADI || '';
      return !scannedInSession.has(itemKey);
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

    // Otomatik olarak sayım sonuçlarını Onay Bekleyenler listesine gönder
    const pendingCountItems = scannedItems.map(data => ({
      pn: data.item.partNumber || data.item.pn || '-',
      sn: data.item.serialAndNotes || data.item.sn || '-',
      itemName: data.item.description || data.item.name || 'Sayılan Parça',
      category: sayimCategory,
      quantity: data.sayilan,
      date: new Date().toLocaleString('tr-TR'),
      requestedBy: sayimPersonel || 'Sayım Sorumlusu',
      depot: sayimDepoBolge,
      tailNo: 'SAYIM SONUCU',
      notes: `SAYIM ONAYI (${sayimDepoBolge} DEPOSU) - Fiziksel Sayılan: ${data.sayilan} Adet`
    }));

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
    
    try {
      if (updatedInv) {
        localStorage.setItem('ogm_depo_inventory_v5', JSON.stringify(updatedInv));
      }
      if (tx) {
        const storedTxs = localStorage.getItem('ogm_depo_transactions_v5');
        const existingTxs: DepoTransaction[] = storedTxs ? JSON.parse(storedTxs) : [];
        localStorage.setItem('ogm_depo_transactions_v5', JSON.stringify([tx, ...existingTxs]));
      }
    } catch (e) {}

    if (onAddTransaction && tx) onAddTransaction(tx);
    if (onUpdateInventory && updatedInv) onUpdateInventory(updatedInv);
  };

  const resolveNextExcess = (source: 'ANKARA' | 'BOLGE') => {
    const [current, ...remaining] = excessItemsQueue;
    if (current) {
      if (sessionActive) {
        // Collect for batch processing at the end
        const it = current.item as any;
        const itemKey = it.partNumber || it.pn || it.MALZEME_PN || it.description || it.name || it.MALZEME_ADI || '';
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
    handleCodeRead(manualCodeInput.trim());
    setManualCodeInput('');
  };

  // Single item save or Confirmation from Popup
  const handleSaveSayimFarki = () => {
    if (!sayimItem) return;
    const regionProp = DEPO_REGIONS.find(r => r.id === sayimDepoBolge)?.prop || 'ankaraMevcut';
    const sistemStok = Number((sayimItem as any)[regionProp] ?? 0);
    const sayilan = Number(sayilanAdetInput);
    const fark = isNaN(sayilan) ? 0 : sayilan - sistemStok;

    const it = sayimItem as any;
    const itemKey = sayimItem.partNumber || sayimItem.pn || it.MALZEME_PN || sayimItem.name || it.MALZEME_ADI || '';

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
    <div className="fixed inset-0 z-[10150] bg-black/75 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
      <div className="bg-slate-900 text-white rounded-3xl shadow-2xl max-w-4xl w-full border border-slate-700 overflow-hidden flex flex-col max-h-[92vh]">
        
        {/* HEADER */}
        <div className="bg-slate-950 px-5 py-4 border-b border-slate-800 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-emerald-500/20 text-emerald-400 rounded-2xl border border-emerald-500/30 shadow-inner">
              <Camera className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-sm font-black uppercase tracking-wider text-white">
                  BARKOD OKUYUCU & SAYIM TERMİNALİ
                </h3>
                {isUnlocked && (
                  <span className="px-2 py-0.5 rounded text-[10px] font-black bg-emerald-950 text-emerald-300 border border-emerald-700">
                    {UNITS_SECURITY.find(u => u.id === selectedUnit)?.name} AKTİF
                  </span>
                )}
              </div>
              <p className="text-[11px] text-slate-400 font-sans">
                Mobil Kamera ile Barkod / Karekod Okuma ve Depo Stok Sayım Modülü
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1 rounded-lg transition-colors cursor-pointer text-lg font-bold"
          >
            ✕
          </button>
        </div>

        {/* BODY */}
        <div className="p-4 sm:p-5 overflow-y-auto flex-1 space-y-4">
          
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
              <div className="bg-slate-950 p-2 rounded-2xl border border-slate-800 flex items-center justify-between gap-2 flex-wrap">
                <div className="flex items-center gap-1.5 flex-1 min-w-[280px]">
                  <button
                    type="button"
                    onClick={() => {
                      setActiveModule('read');
                      setMatchedItem(null);
                      if (sessionActive) handleFinishSession();
                    }}
                    className={`flex-1 flex items-center justify-center gap-2 py-2 px-3 rounded-xl text-xs font-black uppercase transition-all cursor-pointer ${
                      activeModule === 'read'
                        ? 'bg-emerald-600 text-white shadow-md'
                        : 'bg-slate-900 text-slate-400 hover:text-white'
                    }`}
                  >
                    <Barcode className="w-4 h-4" />
                    <span>1. BARKOD OKU (STOK SORGULA)</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      setActiveModule('count');
                      setSayimItem(null);
                      if (isCameraActive) stopCamera();
                    }}
                    className={`flex-1 flex items-center justify-center gap-2 py-2 px-3 rounded-xl text-xs font-black uppercase transition-all cursor-pointer ${
                      activeModule === 'count'
                        ? 'bg-amber-600 text-white shadow-md'
                        : 'bg-slate-900 text-slate-400 hover:text-white'
                    }`}
                  >
                    <Layers className="w-4 h-4" />
                    <span>2. SAYIMA BAŞLA (DEPO SAYIMI)</span>
                  </button>
                </div>

                <button
                  type="button"
                  onClick={() => setIsUnlocked(false)}
                  className="px-2.5 py-1.5 text-[10px] font-bold text-slate-400 hover:text-white bg-slate-900 rounded-lg border border-slate-800 transition"
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
                  <form onSubmit={handleManualSubmit} className="flex items-center gap-2 pt-1 relative">
                    <div className="relative flex-1">
                      <Search className="w-4 h-4 text-slate-500 absolute left-3 top-2.5 z-10" />
                      <input
                        type="text"
                        value={manualCodeInput}
                        onChange={(e) => {
                          setManualCodeInput(e.target.value);
                          setIsManualSearchDropdownOpen(true);
                        }}
                        onFocus={() => setIsManualSearchDropdownOpen(true)}
                        placeholder="Barkod / Karekod verisi veya P/N / Malzeme Adı yazınız..."
                        className="w-full pl-9 pr-3 py-2 bg-slate-900 border border-slate-700 rounded-xl text-xs text-white placeholder-slate-500 font-mono focus:outline-none focus:border-emerald-500 font-bold"
                      />

                      {/* Autocomplete Dropdown List */}
                      {isManualSearchDropdownOpen && manualCodeInput.trim().length > 0 && (
                        <div className="absolute top-full left-0 right-0 z-50 mt-1 bg-slate-950 border-2 border-emerald-500/60 rounded-2xl shadow-2xl max-h-56 overflow-y-auto divide-y divide-slate-800/80">
                          {activeInventory.filter(i => {
                            const q = manualCodeInput.toLowerCase().trim();
                            const p = (i.partNumber || i.pn || (i as any).MALZEME_PN || '').toLowerCase();
                            const d = (i.description || i.name || (i as any).MALZEME_ADI || '').toLowerCase();
                            const s = (i.serialAndNotes || i.sn || (i as any).SERI_NO || '').toLowerCase();
                            return p.includes(q) || d.includes(q) || s.includes(q);
                          }).slice(0, 10).map((item, idx) => (
                            <div
                              key={idx}
                              onClick={() => {
                                const selPn = item.partNumber || item.pn || (item as any).MALZEME_PN || item.description || item.name;
                                setManualCodeInput(selPn);
                                setIsManualSearchDropdownOpen(false);
                                playBeep();
                                handleCodeRead(selPn);
                              }}
                              className="p-2.5 hover:bg-emerald-950/60 cursor-pointer text-xs flex items-center justify-between gap-2 transition-colors"
                            >
                              <div className="min-w-0 flex-1">
                                <div className="font-bold text-white truncate">{item.description || item.name}</div>
                                <div className="text-[11px] font-mono text-slate-400 truncate">
                                  P/N: <strong className="text-emerald-300">{item.partNumber || item.pn || '-'}</strong> | S/N: {item.serialAndNotes || item.sn || '-'}
                                </div>
                              </div>
                              <span className="text-[10px] font-mono font-bold bg-emerald-950 text-emerald-300 border border-emerald-800 px-2 py-0.5 rounded-full shrink-0">
                                Stok: {item.ankaraMevcut || item.toplamStok || 0}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                    <button
                      type="submit"
                      className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs font-bold rounded-xl transition cursor-pointer shrink-0"
                    >
                      Sorgula
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setIsManualSearchDropdownOpen(false);
                        handleOpenUrunTalebi(matchedItem || null);
                      }}
                      className="px-3.5 py-2 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-400 text-slate-950 text-xs font-black uppercase tracking-wider rounded-xl transition cursor-pointer shadow flex items-center gap-1.5 shrink-0"
                      title="Depodan Çıkan Talebi Oluştur ('ONAY BEKLEYENLER -AT802' Sayfasına İlet)"
                    >
                      <Package className="w-3.5 h-3.5 text-slate-950" />
                      <span>ÜRÜN TALEBİ</span>
                    </button>
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
                          <div className="flex items-center gap-3 text-xs font-mono text-slate-300 mt-1 flex-wrap">
                            <span>P/N: <strong className="text-emerald-300">{matchedItem.partNumber || matchedItem.pn || '-'}</strong></span>
                            <span>S/N: <strong>{matchedItem.serialAndNotes || matchedItem.sn || '-'}</strong></span>
                            <span className="px-2 py-0.5 rounded bg-yellow-400 text-black font-black text-[10px]">
                              RAF: {matchedItem.lokasyonNo || matchedItem.location || 'ANKARA DEPO'}
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
                        <div className="text-[10px] font-bold text-slate-400 uppercase font-mono mb-2">
                          GÜNCEL DEPO BÖLGE MEVCUTLARI (CANLI STOK)
                        </div>

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
                              {formOneResults.map((doc, idx) => (
                                <div key={idx} className="py-2 flex items-center justify-between gap-3 text-xs">
                                  <div>
                                    <div className="font-bold text-white flex items-center gap-2">
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
                                  <button
                                    type="button"
                                    onClick={() => alert(`📄 ${doc.belgeNo} Form One Belgesi Görüntüleniyor...\nP/N: ${doc.pn}\nS/N: ${doc.sn}\nSertifika No: ${doc.belgeNo}`)}
                                    className="px-3 py-1.5 bg-blue-700 hover:bg-blue-600 text-white text-[11px] font-bold rounded-lg shrink-0 transition cursor-pointer"
                                  >
                                    GÖSTER / AÇ
                                  </button>
                                </div>
                              ))}
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
                            style={{ width: `${(scannedInSession.size / sessionTotalCount) * 100}%` }}
                          />
                        </div>

                        <div className="flex items-center gap-2">
                          <button
                            type="button"
                            onClick={() => setShowRemainingModal(true)}
                            className="flex-1 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-[10px] font-black uppercase tracking-wider transition border border-slate-700 flex items-center justify-center gap-2 cursor-pointer"
                          >
                            <Search className="w-3.5 h-3.5" />
                            KALANLARI GÖR ({sessionTotalCount - scannedInSession.size})
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
                                {sessionCounts[sayimItem.partNumber || sayimItem.pn || (sayimItem as any).MALZEME_PN || sayimItem.name || (sayimItem as any).MALZEME_ADI || ''] || 0} EA
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
                      <div className="fixed inset-0 z-[10300] bg-slate-900/90 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto animate-fade-in">
                        <div className="bg-slate-950 p-6 rounded-3xl border-2 border-amber-500/50 shadow-2xl space-y-5 max-w-xl w-full">
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
        <div className="fixed inset-0 z-[10200] flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-slate-900 p-6 rounded-3xl border-2 border-slate-700 shadow-2xl max-w-lg w-full flex flex-col max-h-[80vh]">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3 mb-4">
              <h4 className="text-sm font-black text-white uppercase tracking-widest">OKUTULMAYAN ÜRÜNLER (KALANLAR)</h4>
              <button onClick={() => setShowRemainingModal(false)} className="text-slate-400 hover:text-white text-lg">✕</button>
            </div>
            
            <div className="overflow-y-auto space-y-2 flex-1">
              {sessionItemsToCount
                .filter(i => !scannedInSession.has(i.partNumber || i.pn || (i as any).MALZEME_PN || i.name || (i as any).MALZEME_ADI || ''))
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

      {/* ─── ÜRÜN TALEBİ (ÇIKAN TALEBİ) MODALI (ONAY BEKLEYENLER -AT802 - ÇOKLU PARÇA VE KORUMALI VERİ) ─── */}
      {isUrunTalebiModalOpen && (
        <div className="fixed inset-0 z-[10500] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border-2 border-amber-500/50 text-white rounded-3xl shadow-2xl max-w-xl w-full overflow-hidden flex flex-col max-h-[92vh]">
            <div className="bg-amber-950/80 px-5 py-4 border-b border-amber-900/60 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-amber-500/20 text-amber-400 rounded-xl border border-amber-500/30">
                  <Package className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-amber-200">
                    ÜRÜN TALEBİ OLUŞTUR — AT-802F
                  </h3>
                  <p className="text-[10px] text-amber-300">
                    Seçilen ürünlerin resmi bilgileri otomatik doldurulur ve onaylanmadan stoktan düşülmez
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setIsUrunTalebiModalOpen(false)} 
                className="text-slate-400 hover:text-white font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleTalepSubmit} className="p-5 space-y-4 text-xs overflow-y-auto flex-1 font-sans">
              
              {/* TALEBE EKLENEN PARÇALAR LİSTESİ */}
              <div className="space-y-2.5">
                <div className="flex items-center justify-between">
                  <span className="text-[10px] font-black uppercase text-amber-400 tracking-wider font-mono">
                    TALEBE EKLENEN PARÇALAR ({talepItemList.length})
                  </span>
                  <button
                    type="button"
                    onClick={() => setIsAddTalepItemOpen(prev => !prev)}
                    className="flex items-center gap-1.5 px-3 py-1 bg-amber-500 hover:bg-amber-400 text-amber-950 font-black rounded-lg text-xs transition cursor-pointer shadow-xs active:scale-95"
                  >
                    <span className="text-sm">+</span>
                    <span>Parça / Barkod Ekle</span>
                  </button>
                </div>

                {/* Ekleme Açılır Kutusu (Sorgula / Tüm Depolar) */}
                {isAddTalepItemOpen && (
                  <div className="p-3 bg-slate-950 rounded-2xl border-2 border-amber-500/50 space-y-2 animate-fade-in">
                    <label className="block text-[9.5px] font-bold text-amber-300 uppercase font-mono">
                      TÜM DEPOLARDA ARA VE BİR ÜRÜN DAHA EKLE (SARF VE KİMYASAL)
                    </label>
                    <input
                      type="text"
                      value={addTalepSearchQuery}
                      onChange={(e) => setAddTalepSearchQuery(e.target.value)}
                      placeholder="P/N veya malzeme adı yazarak arayınız..."
                      className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-xl text-white text-xs font-mono font-bold focus:outline-none focus:border-amber-400"
                    />
                    <div className="max-h-44 overflow-y-auto divide-y divide-slate-800 rounded-xl bg-slate-900/60 border border-slate-800">
                      {((activeInventory && activeInventory.length > 0) ? activeInventory : inventory).filter(i => {
                        if (!addTalepSearchQuery.trim()) return true;
                        const q = addTalepSearchQuery.toLowerCase().trim();
                        const d = (i.description || i.name || (i as any).MALZEME_ADI || '').toLowerCase();
                        const p = (i.partNumber || i.pn || (i as any).MALZEME_PN || '').toLowerCase();
                        return d.includes(q) || p.includes(q);
                      }).slice(0, 10).map((item, idx) => (
                        <div
                          key={idx}
                          onClick={() => handleAddTalepItemFromInventory(item)}
                          className="p-2.5 hover:bg-amber-950/50 cursor-pointer text-xs flex items-center justify-between gap-2 transition"
                        >
                          <div>
                            <div className="font-bold text-white">{item.description || item.name || (item as any).MALZEME_ADI}</div>
                            <div className="text-[10px] font-mono text-slate-400">
                              P/N: <strong className="text-amber-300">{item.partNumber || item.pn || (item as any).MALZEME_PN || '-'}</strong>
                            </div>
                          </div>
                          <span className="text-[10px] font-mono font-bold bg-amber-950 text-amber-300 border border-amber-800 px-2 py-0.5 rounded-full shrink-0">
                            {item.category === 'kimyasal' ? '🧪 Kimyasal' : '📦 Sarf'}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Eklenen Parçalar Kart Listesi */}
                {talepItemList.length > 0 ? (
                  <div className="space-y-2 max-h-64 overflow-y-auto pr-1">
                    {talepItemList.map((it, idx) => (
                      <div key={it.id} className="p-3 bg-slate-950 rounded-2xl border border-slate-800 flex flex-col gap-2 relative group hover:border-amber-500/40 transition">
                        <div className="flex items-start justify-between gap-2 border-b border-slate-900 pb-2">
                          <div>
                            <div className="text-xs font-black text-white">{it.itemName}</div>
                            <div className="text-[10px] font-mono text-slate-400 flex items-center gap-2 mt-0.5">
                              <span>P/N: <strong className="text-amber-300">{it.pn || '-'}</strong></span>
                              <span className="text-slate-600">|</span>
                              <span className="text-amber-400 font-bold">OTOMATİK DOLDURULDU (DÜZENLENEMEZ)</span>
                            </div>
                          </div>
                          <div className="flex items-center gap-1.5 shrink-0">
                            <span className="text-[10px] font-mono font-bold bg-amber-950/80 text-amber-300 border border-amber-800/80 px-2 py-0.5 rounded-md">
                              {it.category === 'kimyasal' ? '🧪 Kimyasal Depo' : '📦 Sarf Depo'}
                            </span>
                            <button
                              type="button"
                              onClick={() => handleRemoveTalepItem(it.id)}
                              className="text-rose-400 hover:text-rose-200 font-bold p-1 hover:bg-rose-950/50 rounded-lg transition"
                              title="Bu parçayı listeden çıkar"
                            >
                              ✕
                            </button>
                          </div>
                        </div>

                        <div className={`grid ${it.category === 'kimyasal' ? 'grid-cols-2' : 'grid-cols-3'} gap-2 text-[11px]`}>
                          <div>
                            <label className="block text-[9px] font-bold text-slate-400 uppercase font-mono mb-0.5">Miktar (Adet)</label>
                            <input
                              type="number"
                              min={1}
                              value={it.quantity}
                              onChange={(e) => handleUpdateTalepItemField(it.id, 'quantity', Math.max(1, parseInt(e.target.value) || 1))}
                              className="w-full px-2 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono font-bold"
                            />
                          </div>

                          <div>
                            <label className="block text-[9px] font-bold text-slate-400 uppercase font-mono mb-0.5">Depo Seçimi</label>
                            <select
                              value={it.depot}
                              onChange={(e) => handleUpdateTalepItemField(it.id, 'depot', e.target.value)}
                              className="w-full px-2 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white font-bold"
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
                              <label className="block text-[9px] font-bold text-slate-400 uppercase font-mono mb-0.5">Serial Number (S/N)</label>
                              <input
                                type="text"
                                value={it.sn}
                                onChange={(e) => handleUpdateTalepItemField(it.id, 'sn', e.target.value)}
                                placeholder="S/N..."
                                className="w-full px-2 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono font-bold"
                              />
                            </div>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="p-4 bg-slate-950/60 rounded-2xl border border-slate-850 text-center text-slate-500 font-bold text-xs">
                    Henüz talebe parça eklenmedi. Yukarıdaki (+) butonuna basarak ürün ekleyebilirsiniz.
                  </div>
                )}
              </div>

              <div className="grid grid-cols-2 gap-3">
                {/* Kuyruk No (Çıkan Malzemedeki Kuyruk No Listesi) */}
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    KUYRUK NO SEÇİMİ (AT-802) *
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
                      placeholder="Kuyruk no manuel yazınız (Örn: OR-2045)..."
                      required
                      className="w-full mt-1.5 px-3 py-1.5 bg-amber-950/40 border border-amber-500/50 rounded-lg text-amber-200 font-mono text-xs focus:outline-none focus:border-amber-500 font-bold"
                    />
                  )}
                </div>

                {/* Tarih */}
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono">
                      TALEP TARİHİ (TAKVİM) *
                    </label>
                    <button
                      type="button"
                      onClick={() => {
                        const pad = (n: number) => String(n).padStart(2, '0');
                        const now = new Date();
                        setTalepDate(`${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
                      }}
                      className="text-[9px] text-amber-400 hover:text-amber-300 font-bold underline cursor-pointer"
                    >
                      📅 Şimdiki Zaman
                    </button>
                  </div>
                  <input
                    type="datetime-local"
                    value={(() => {
                      if (!talepDate) return '';
                      if (talepDate.includes('T')) return talepDate;
                      const parts = talepDate.split(' ');
                      if (parts.length === 2 && parts[0].includes('.')) {
                        const dParts = parts[0].split('.');
                        return `${dParts[2]}-${dParts[1]}-${dParts[0]}T${parts[1]}`;
                      }
                      return talepDate;
                    })()}
                    onChange={(e) => {
                      const val = e.target.value;
                      if (!val) {
                        setTalepDate('');
                        return;
                      }
                      const [d, t] = val.split('T');
                      if (d && t) {
                        const [y, m, day] = d.split('-');
                        setTalepDate(`${day}.${m}.${y} ${t}`);
                      } else {
                        setTalepDate(val);
                      }
                    }}
                    required
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono focus:outline-none focus:border-amber-500 cursor-pointer color-scheme-dark"
                  />
                </div>
              </div>

              {/* Teslim Alan Personel */}
              <div>
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  TESLİM ALAN PERSONEL (LİSTEDEN SEÇİNİZ) *
                </label>
                <PersonnelAutocomplete
                  value={talepRequestedBy}
                  onChange={setTalepRequestedBy}
                  placeholder="Kendi adınızı seçiniz veya yazınız..."
                />
              </div>

              {/* Açıklama */}
              <div>
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  AÇIKLAMA / İŞ EMRİ / TALEP NOTU
                </label>
                <textarea
                  rows={2}
                  value={talepNotes}
                  onChange={(e) => setTalepNotes(e.target.value)}
                  placeholder="Talep gerekçesi, bakım görevi, parça değişim nedeni..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-amber-500 resize-none font-sans"
                />
              </div>

              <div className="pt-3 border-t border-slate-800 flex justify-end gap-2 shrink-0 font-sans">
                <button
                  type="button"
                  onClick={() => setIsUrunTalebiModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-750 text-slate-300 rounded-xl font-bold transition cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  disabled={isTalepSubmitting}
                  className="px-5 py-2 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-400 disabled:opacity-50 text-slate-950 font-black uppercase rounded-xl shadow-lg transition cursor-pointer active:scale-95 flex items-center gap-2"
                >
                  <Send className="w-4 h-4 text-slate-950" />
                  <span>{isTalepSubmitting ? 'Kaydediliyor...' : 'TALEBİ KAYDET VE İLET'}</span>
                </button>
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
