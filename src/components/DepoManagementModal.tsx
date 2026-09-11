import React, { useState, useMemo, useRef, useEffect } from 'react';
import { 
  Boxes, 
  ArrowUpRight, 
  ArrowDownLeft, 
  ArrowLeftRight, 
  CheckCircle2, 
  Search, 
  Plus, 
  Trash2, 
  FileSpreadsheet, 
  FileText, 
  Lock, 
  Wrench, 
  X, 
  AlertCircle,
  AlertTriangle,
  Filter,
  Clock,
  Eye,
  RotateCcw,
  PlusCircle,
  ShieldCheck,
  Download,
  UploadCloud,
  Plane,
  MapPin,
  Layers,
  ChevronDown,
  RefreshCw,
  UserCheck,
  Check,
  SlidersHorizontal,
  Tag,
  Code2,
  Warehouse,
  Box,
  FlaskConical,
  Copy
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { DepoTransaction, KitComponent } from '../types';
import { cleanAndFormatDateString, unmergeAndFillWorksheet } from '../utils/driveExcelSync';

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw4kruTTc058Y9rLTyO3dKi6KloYsmdDTwV1GSiAk8ZXefyo3Z7_VDSTuurzsS9BHAQyQ/exec";

// Hava Aracı Kuyruk & Çağrı Kodları Listesi (Kullanıcı Tablosu Esas Alınmıştır)
export const AIRCRAFT_TAIL_LIST = [
  // C-650
  { unit: 'C-650', callsign: 'ORMAN 01', tailNo: 'OR 0177', fullLabel: 'ORMAN 01 (OR 0177) - C-650' },
  // B-360
  { unit: 'B-360', callsign: 'ORMAN 02', tailNo: 'OR 1839', fullLabel: 'ORMAN 02 (OR 1839) - B-360' },
  // BELL-429
  { unit: 'BELL-429', callsign: 'ORMAN 03', tailNo: 'OR 3125', fullLabel: 'ORMAN 03 (OR 3125) - BELL-429' },
  { unit: 'BELL-429', callsign: 'ORMAN 04', tailNo: 'OR 3126', fullLabel: 'ORMAN 04 (OR 3126) - BELL-429' },
  { unit: 'BELL-429', callsign: 'ORMAN 05', tailNo: 'OR 3127', fullLabel: 'ORMAN 05 (OR 3127) - BELL-429' },
  { unit: 'BELL-429', callsign: 'ORMAN 06', tailNo: 'OR 3131', fullLabel: 'ORMAN 06 (OR 3131) - BELL-429' },
  { unit: 'BELL-429', callsign: 'ORMAN 07', tailNo: 'OR 3135', fullLabel: 'ORMAN 07 (OR 3135) - BELL-429' },
  { unit: 'BELL-429', callsign: 'ORMAN 08', tailNo: 'OR 3192', fullLabel: 'ORMAN 08 (OR 3192) - BELL-429' },
  // T-70 (Sırayla OR-1018, OR-1019, OR-1020)
  { unit: 'T-70', callsign: 'ORMAN 18', tailNo: 'OR-1018', fullLabel: 'ORMAN 18 (OR-1018) - T-70' },
  { unit: 'T-70', callsign: 'ORMAN 19', tailNo: 'OR-1019', fullLabel: 'ORMAN 19 (OR-1019) - T-70' },
  { unit: 'T-70', callsign: 'ORMAN 20', tailNo: 'OR-1020', fullLabel: 'ORMAN 20 (OR-1020) - T-70' },
  // AT-802
  { unit: 'AT-802', callsign: 'ORMAN 21', tailNo: 'OR-2021', fullLabel: 'ORMAN 21 (OR-2021) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 22', tailNo: 'OR-2022', fullLabel: 'ORMAN 22 (OR-2022) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 23', tailNo: 'OR-2023', fullLabel: 'ORMAN 23 (OR-2023) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 24', tailNo: 'OR-2024', fullLabel: 'ORMAN 24 (OR-2024) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 25', tailNo: 'OR-2025', fullLabel: 'ORMAN 25 (OR-2025) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 26', tailNo: 'OR-2026', fullLabel: 'ORMAN 26 (OR-2026) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 27', tailNo: 'OR-2027', fullLabel: 'ORMAN 27 (OR-2027) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 28', tailNo: 'OR-2028', fullLabel: 'ORMAN 28 (OR-2028) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 29', tailNo: 'OR-2029', fullLabel: 'ORMAN 29 (OR-2029) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 30', tailNo: 'OR-2030', fullLabel: 'ORMAN 30 (OR-2030) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 31', tailNo: 'OR-2031', fullLabel: 'ORMAN 31 (OR-2031) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 36', tailNo: 'OR-2036', fullLabel: 'ORMAN 36 (OR-2036) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 37', tailNo: 'OR-2037', fullLabel: 'ORMAN 37 (OR-2037) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 38', tailNo: 'OR-2038', fullLabel: 'ORMAN 38 (OR-2038) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 39', tailNo: 'OR-2039', fullLabel: 'ORMAN 39 (OR-2039) - AT-802' },
  { unit: 'AT-802', callsign: 'ORMAN 40', tailNo: 'OR-2040', fullLabel: 'ORMAN 40 (OR-2040) - AT-802' }
];

// Tarih Formatlayıcı Yardımcı (G.AA.YYYY veya GG.AA.YYYY)
const getTodayFormattedDate = () => {
  const d = new Date();
  const day = d.getDate();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}.${month}.${year}`;
};

// Excel Tarih Hücresi Çözümleyici (Sayısal Seri veya Metin Tarihleri Standart GG.AA.YYYY Yapar)
const parseExcelBulkDate = (val: any): string => {
  if (val === undefined || val === null || String(val).trim() === '') {
    return getTodayFormattedDate();
  }
  const formatted = cleanAndFormatDateString(val);
  return formatted || getTodayFormattedDate();
};

// İl / Depo Bazlı Standart Lokasyonlar (Bell 429 Tek Deposu: Ankara)
export const DEPO_REGISTERED_LOCATIONS = [
  'ANKARA (BELL-429 Tek Deposu / Ana Depo)',
  'ANKARA (GÜVERCİNLİK / ETİMESGUT)',
  'ANTALYA DEPOSU',
  'İZMİR (ADNAN MENDERES / ALİAĞA)',
  'MUĞLA (MİLAS / BODRUM / DALAMAN)',
  'ÇANAKKALE (DARDANOS / HAVAALANI)',
  'ADANA DEPOSU',
  'BALIKESİR (EDREMİT) DEPOSU',
  'BURSA (YENİŞEHİR) DEPOSU',
  'KASTAMONU DEPOSU'
];

// Çıkış İşlem Türleri (Yalnızca Çıkış / Sarf işlemlerinde gösterilir)
export const CIKIS_TRANSACTION_TYPES = [
  'ANKARA ÇIKAN',
  'ANTALYA ÇIKAN',
  'İZMİR ÇIKAN',
  'MUĞLA ÇIKAN',
  'ÇANAKKALE ÇIKAN',
  'BURSA ÇIKAN',
  'ADANA ÇIKAN',
  'BALIKESİR ÇIKAN',
  'KASTAMONU ÇIKAN',
  'KARAİN ÇIKAN',
  'MİLAS ÇIKAN',
  'ÇIKIŞ',
  'DİĞER'
];

// Giriş İşlem Türleri (Yalnızca Giriş / Stok Artırma işlemlerinde gösterilir)
export const GIRIS_TRANSACTION_TYPES = [
  'ANKARA GİREN',
  'ANTALYA GİREN',
  'İZMİR GİREN',
  'MUĞLA GİREN',
  'ÇANAKKALE GİREN',
  'BURSA GİREN',
  'ADANA GİREN',
  'BALIKESİR GİREN',
  'KASTAMONU GİREN',
  'KARAİN GİREN',
  'MİLAS GİREN',
  'GELEN',
  'GİRİŞ',
  'DİĞER'
];

// Transfer İşlem Türleri (Yalnızca Depo İçi / İller Arası Transfer işlemlerinde gösterilir)
export const TRANSFER_TRANSACTION_TYPES = [
  'ANKARA TRANSFER',
  'ANTALYA TRANSFER',
  'İZMİR TRANSFER',
  'MUĞLA TRANSFER',
  'ÇANAKKALE TRANSFER',
  'BURSA TRANSFER',
  'ADANA TRANSFER',
  'BALIKESİR TRANSFER',
  'KASTAMONU TRANSFER',
  'KARAİN TRANSFER',
  'MİLAS TRANSFER',
  'TRANSFER',
  'DİĞER'
];

// Standart İşlem Türleri (Tüm Liste)
export const STANDARD_TRANSACTION_TYPES = Array.from(
  new Set([
    ...CIKIS_TRANSACTION_TYPES,
    ...GIRIS_TRANSACTION_TYPES,
    ...TRANSFER_TRANSACTION_TYPES
  ])
);

// Malzeme Türü / Kategori Listesi (Sayım ve Yönetim İçin: 2 Ana Depo)
export type MaterialCategoryType = 'TÜMÜ' | 'SARF VE PARÇA DEPO' | 'KİMYASAL DEPO';

export const detectItemCategory = (row: string[]): 'SARF VE PARÇA DEPO' | 'KİMYASAL DEPO' => {
  if (!row) return 'SARF VE PARÇA DEPO';
  
  // Kullanıcı Kuralı: Eğer Drive'da kimyasal depo Excel'i yoksa veya ürün açıkça kimyasal depo Excel'inden
  // yüklenmemişse asla otomatik kimyasal depoya atanamaz, kimyasal depo boş kalır.
  const explicitCategory = String(row[12] || row[13] || row[11] || '').trim().toLowerCase();
  if (
    explicitCategory === 'depo_kimyasal' || 
    explicitCategory === 'kimyasal_depo' || 
    explicitCategory === 'kimyasal depo'
  ) {
    return 'KİMYASAL DEPO';
  }

  const rowStr = row.join(' ').toUpperCase();
  if (rowStr.includes('##DEPO_KIMYASAL##') || rowStr.includes('[KİMYASAL DEPO EXCEL]')) {
    return 'KİMYASAL DEPO';
  }

  // Varsayılan: Sarf ve Parça Depo
  return 'SARF VE PARÇA DEPO';
};

// Transfer Kaydının Depo Stoklarında Eşleşip Düşülüp Düşülmediğini Kontrol Eden Fonksiyon
export const isTxUnmatchedWithDepo = (tx: DepoTransaction, inventory: string[][]): boolean => {
  if (!tx || !tx.itemName) return true;
  const name = tx.itemName.trim().toLowerCase();
  if (name.startsWith('malzeme #') || name === '' || name === '-') return true;

  const pn = (tx.pn || '').trim().toLowerCase();
  
  return !inventory.some(r => {
    const rName = (r[1] || '').trim().toLowerCase();
    const rPn = (r[2] || '').trim().toLowerCase();
    if (rName && rName === name) return true;
    if (rPn && rPn !== '-' && rPn === name) return true;
    if (pn && pn !== '-' && rPn && rPn === pn) return true;
    if (pn && pn !== '-' && rName && rName === pn) return true;
    
    // Alphanumeric clean match (removing dashes/spaces/commas)
    const cleanName = name.replace(/[^a-z0-9]/g, '');
    const cleanPn = pn.replace(/[^a-z0-9]/g, '');
    const cleanRName = rName.replace(/[^a-z0-9]/g, '');
    const cleanRPn = rPn.replace(/[^a-z0-9]/g, '');
    if (cleanName && cleanName.length >= 2 && (cleanName === cleanRName || cleanName === cleanRPn)) return true;
    if (cleanPn && cleanPn.length >= 2 && (cleanPn === cleanRPn || cleanPn === cleanRName)) return true;
    return false;
  });
};

interface DepoManagementModalProps {
  isOpen: boolean;
  onClose: () => void;
  subSection?: 'depo_sarf' | 'depo_kimyasal' | 'depo_all';
  depoRows?: string[][];
  at802Rows?: string[][];
  allInventory?: { unit: string; pn: string; name: string; sn: string; qty: string; loc: string; status: string }[];
  onUpdateDepoRows?: (updatedRows: string[][]) => void;
  onUpdateAt802Rows?: (updatedRows: string[][]) => void;
  transactions: DepoTransaction[];
  onAddTransaction: (tx: Omit<DepoTransaction, 'id'> | (Omit<DepoTransaction, 'id' | 'timestamp'> & { timestamp?: string })) => void;
  onAddTransactionsBatch?: (txs: (Omit<DepoTransaction, 'id'> | (Omit<DepoTransaction, 'id' | 'timestamp'> & { timestamp?: string }))[]) => void;
  onUpdateTransaction?: (tx: DepoTransaction) => void;
  onDeleteTransaction?: (txId: string) => void;
  onDeleteTransactionsBatch?: (txIds: string[]) => void;
  onUndoTransaction?: (txId: string) => void;
  certificatePdfUrl: string | null;
  onUploadCertificatePdf?: (file: File) => void;
  onSaveCertificatePdfUrl?: (url: string) => void;
  showNotification?: (msg: string) => void;
}

export const DepoManagementModal: React.FC<DepoManagementModalProps> = ({
  isOpen,
  onClose,
  subSection = 'depo_all',
  depoRows = [],
  at802Rows = [],
  allInventory,
  onUpdateDepoRows = (_updatedRows: string[][]) => {},
  onUpdateAt802Rows = (_updatedRows: string[][]) => {},
  transactions = [],
  onAddTransaction = (_tx: any) => {},
  onAddTransactionsBatch,
  onUpdateTransaction,
  onDeleteTransaction,
  onDeleteTransactionsBatch,
  onUndoTransaction = (_txId: string) => {},
  certificatePdfUrl,
  onUploadCertificatePdf = (_file: File) => {},
  onSaveCertificatePdfUrl = (_url: string) => {},
  showNotification = (_msg: string) => {},
}) => {
  const [activeTab, setActiveTab] = useState<'islemler' | 'kit' | 'sayim' | 'loglar'>('islemler');
  const bulkFileInputRef = useRef<HTMLInputElement>(null);

  // Personnel List State & Fetching
  const [personnelList, setPersonnelList] = useState<string[]>([]);
  const [isPersonnelLoading, setIsPersonnelLoading] = useState<boolean>(false);
  const [isDeliveredDropdownOpen, setIsDeliveredDropdownOpen] = useState<boolean>(false);
  const [isAcceptedDropdownOpen, setIsAcceptedDropdownOpen] = useState<boolean>(false);
  const [isSyncingTransfers, setIsSyncingTransfers] = useState<boolean>(false);
  const [isPushingTransfers, setIsPushingTransfers] = useState<boolean>(false);
  const [isProcessingBulk, setIsProcessingBulk] = useState<boolean>(false);

  // State for Editing a Transaction in Transfer Geçmişi
  const [editingTx, setEditingTx] = useState<DepoTransaction | null>(null);

  // State for Multi-Selection in Transfer Geçmişi
  const [selectedTxIds, setSelectedTxIds] = useState<string[]>([]);

  // Effective AT-802 rows for Kit creation
  const effectiveAt802Rows = useMemo(() => {
    if (at802Rows && at802Rows.length > 0) return at802Rows;
    return depoRows;
  }, [at802Rows, depoRows]);

  // Depo İşlemleri States - Tüm alanlar açılışta boş/varsayılan olarak başlar
  const [selectedOperation, setSelectedOperation] = useState<'cikis' | 'giris' | 'transfer'>('cikis');
  const [girisMode, setGirisMode] = useState<'mevcut' | 'yeni'>('mevcut');
  
  // Malzeme Arama / Otomatik Tamamlama
  const [materialSearchQuery, setMaterialSearchQuery] = useState<string>('');
  const [selectedItemIndex, setSelectedItemIndex] = useState<number>(-1);
  const [isMaterialDropdownOpen, setIsMaterialDropdownOpen] = useState<boolean>(false);

  // Form Alanları - Kullanıcı isteği doğrultusunda hiçbir şey seçili olmadan (İşlem Türü seçimsiz) başlar
  const [opQuantity, setOpQuantity] = useState<number>(1);
  const [opDate, setOpDate] = useState<string>(getTodayFormattedDate());
  const [opTransactionType, setOpTransactionType] = useState<string>('');
  const [opCustomTransactionType, setOpCustomTransactionType] = useState<string>('');
  const [opSerialNo, setOpSerialNo] = useState<string>('');
  const [opTailNo, setOpTailNo] = useState<string>('');
  const [opDeliveredTo, setOpDeliveredTo] = useState<string>('');
  const [opAcceptedBy, setOpAcceptedBy] = useState<string>('');
  const [opDepoLocation, setOpDepoLocation] = useState<string>('');
  const [opTargetLocation, setOpTargetLocation] = useState<string>('');
  const [opWorkOrder, setOpWorkOrder] = useState<string>('');
  const [opOperator, setOpOperator] = useState<string>('');
  const [opNotes, setOpNotes] = useState<string>('');

  // Transfer Geçmişi Tablosu Filtreleri
  const [txTypeFilter, setTxTypeFilter] = useState<string>('ALL');
  const [txSearchFilter, setTxSearchFilter] = useState<string>('');
  const [txAircraftTypeFilter, setTxAircraftTypeFilter] = useState<string>('ALL');
  const [txTailFilter, setTxTailFilter] = useState<string>('ALL');
  const [txStatusFilter, setTxStatusFilter] = useState<'ALL' | 'UNMATCHED' | 'MATCHED'>('ALL');

  // Transfer Düzenleme Modalı İçin Malzeme Otomatik Tamamlama
  const [editMaterialSearch, setEditMaterialSearch] = useState<string>('');
  const [isEditMaterialDropdownOpen, setIsEditMaterialDropdownOpen] = useState<boolean>(false);

  // Seçili operasyona göre (Çıkış / Giriş / Transfer) filtrelenmiş İşlem Türleri listesi
  const availableTransactionTypes = useMemo(() => {
    if (selectedOperation === 'cikis') {
      return CIKIS_TRANSACTION_TYPES;
    }
    if (selectedOperation === 'giris') {
      return GIRIS_TRANSACTION_TYPES;
    }
    return TRANSFER_TRANSACTION_TYPES;
  }, [selectedOperation]);

  // Operasyon sekmesi değiştiğinde geçersiz kalan işlem türünü sıfırla
  const handleOperationChange = (op: 'cikis' | 'giris' | 'transfer') => {
    setSelectedOperation(op);
    const targetTypes = op === 'cikis' ? CIKIS_TRANSACTION_TYPES : (op === 'giris' ? GIRIS_TRANSACTION_TYPES : TRANSFER_TRANSACTION_TYPES);
    if (opTransactionType && !targetTypes.includes(opTransactionType)) {
      setOpTransactionType('');
      setOpCustomTransactionType('');
    }
  };

  // Fetch Personnel List on Mount from yoklama & personel tables
  useEffect(() => {
    if (!isOpen) return;
    const fetchPersonnel = async () => {
      setIsPersonnelLoading(true);
      try {
        let loadedNames: string[] = [];
        const res = await fetch('/api/personnel-list');
        if (res.ok) {
          const data = await res.json();
          if (data && Array.isArray(data.personnel) && data.personnel.length > 0) {
            loadedNames = data.personnel;
          } else if (data && Array.isArray(data.data) && data.data.length > 0) {
            loadedNames = data.data;
          }
        }

        // Yerel yoklama ve görevlendirme önbelleklerinden de personelleri dahil et
        try {
          const keys = ['form_5_personel_bilgi', 'form_1_gorevlendirme', 'yoklama_listesi'];
          keys.forEach(k => {
            const raw = localStorage.getItem(k);
            if (raw) {
              const parsed = JSON.parse(raw);
              if (Array.isArray(parsed)) {
                parsed.forEach((item: any) => {
                  if (typeof item === 'string' && item.trim().length > 2) {
                    loadedNames.push(item.trim());
                  } else if (Array.isArray(item) && item[1]) {
                    const name = String(item[1]).trim();
                    if (name.length > 2 && !name.includes('ADI') && !name.includes('SOYADI')) {
                      loadedNames.push(name);
                    }
                  } else if (item && typeof item === 'object' && (item.name || item.adiSoyadi)) {
                    const name = String(item.name || item.adiSoyadi).trim();
                    if (name.length > 2) loadedNames.push(name);
                  }
                });
              }
            }
          });
        } catch (e) {
          console.warn('Local personnel storage parse error:', e);
        }

        const uniquePersonnel = Array.from(new Set(loadedNames))
          .filter(n => n && n.length >= 3 && !n.includes('TOPLAM'))
          .sort((a, b) => a.localeCompare(b, 'tr-TR'));

        if (uniquePersonnel.length > 0) {
          setPersonnelList(uniquePersonnel);
        }
      } catch (e) {
        console.warn('Personnel list fetch error:', e);
      } finally {
        setIsPersonnelLoading(false);
      }
    };
    fetchPersonnel();
  }, [isOpen]);

  // Live Sync Transfer History with Sheets
  const syncTransfersWithGoogleSheets = async () => {
    setIsSyncingTransfers(true);
    try {
      showNotification('Merkezi E-Tablo transfer geçmişi taranıyor...');
      const res = await fetch('/api/depo-transfers');
      if (res.ok) {
        const data = await res.json();
        if (data && Array.isArray(data.transfers) && data.transfers.length > 0) {
          // Merge transfers without duplicates
          data.transfers.forEach((sheetTx: any) => {
            const exists = transactions.some(t => 
              t.itemName === sheetTx.itemName && 
              t.date === sheetTx.date && 
              t.quantity === sheetTx.quantity && 
              t.tailNo === sheetTx.tailNo && 
              t.type === sheetTx.type
            );
            if (!exists) {
              onAddTransaction({
                type: sheetTx.type || 'TRANSFER',
                itemName: sheetTx.itemName,
                pn: sheetTx.pn || '-',
                sn: sheetTx.sn || '-',
                quantity: Number(sheetTx.quantity) || 1,
                date: sheetTx.date || getTodayFormattedDate(),
                timestamp: `${sheetTx.date || getTodayFormattedDate()} 00:00`,
                tailNo: sheetTx.tailNo || '-',
                operator: sheetTx.deliveredTo || '-',
                receivedBy: sheetTx.acceptedBy || '-',
                location: sheetTx.location || 'DEPO',
                sourceLocation: sheetTx.location || 'DEPO',
                targetLocation: sheetTx.location || 'DEPO',
                notes: 'E-Tablo Transfer Senkronizasyonu'
              });
            }
          });
          showNotification(`${data.transfers.length} adet transfer kaydı başarıyla senkronize edildi!`);
        } else {
          showNotification('Merkezi E-Tablo üzerinde yeni transfer kaydı bulunamadı.');
        }
      }
    } catch (e: any) {
      console.warn('Transfer sync error:', e);
      showNotification(`Eşitleme hatası: ${e.message || e}`);
    } finally {
      setIsSyncingTransfers(false);
    }
  };

  // AT-802 Transfer Geçmişini Excel (.xlsx) olarak Google Drive Klasörüne Yükleme
  const uploadAt802TransfersExcelToDrive = async (customTransfers?: DepoTransaction[]) => {
    const listToExport = customTransfers || transactions;
    if (listToExport.length === 0) {
      showNotification('Aktarılacak transfer kaydı bulunamadı.');
      return;
    }
    setIsPushingTransfers(true);
    try {
      showNotification('AT-802 Transfer Geçmişi Excel dosyası hazırlanıyor ve Drive klasörüne yükleniyor...');

      // Tablo başlıkları ve verileri
      const headers = [
        "SIRA NO", "İŞLEM TARİHİ", "İŞLEM TÜRÜ", "MALZEME ADI", "PARÇA NO (P/N)", 
        "SERİ NO (S/N)", "MİKTAR", "KUYRUK / HAVA ARACI", "ÇIKIŞ LOKASYONU", 
        "HEDEF LOKASYON / DEPO", "TESLİM EDEN / PERSONEL", "KABUL EDEN / ALAN", 
        "AÇIKLAMA / NOTLAR"
      ];

      const rows = listToExport.map((tx, idx) => [
        String(idx + 1),
        tx.date || tx.timestamp || '',
        tx.type || 'TRANSFER',
        tx.itemName || '',
        tx.pn || '-',
        tx.sn || '-',
        String(tx.quantity || 1),
        tx.tailNo || '-',
        tx.sourceLocation || tx.location || 'DEPO',
        tx.targetLocation || tx.tailNo || 'HAVA ARACI',
        tx.operator || '-',
        tx.receivedBy || '-',
        tx.notes || ''
      ]);

      const wsData = [headers, ...rows];
      const ws = XLSX.utils.aoa_to_sheet(wsData);

      // Sütun genişlikleri
      ws['!cols'] = [
        { wch: 8 }, { wch: 14 }, { wch: 16 }, { wch: 32 }, { wch: 18 },
        { wch: 16 }, { wch: 10 }, { wch: 20 }, { wch: 20 }, { wch: 24 },
        { wch: 22 }, { wch: 22 }, { wch: 35 }
      ];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "TRANSFER GEÇMİŞİ _AT-802");

      const base64Excel = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });

      // 1. Google Drive Klasörüne (1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP) Excel yükle
      await fetch('/api/upload-techizat-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: 'transfer_gecmisi_at802.xlsx',
          targetKey: 'transfer_gecmisi_at802',
          base64Data: base64Excel,
          folderId: '1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP'
        })
      });

      // 2. Canlı E-Tabloya kaydetmeyi (kullanıcı talebi doğrultusunda sadece Drive excel odaklı olacağı için) iptal ediyoruz
      /*
      const targetSpreadsheetId = "17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0";
      await fetch('/api/save-depo-transfers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          transfers: listToExport,
          spreadsheetId: targetSpreadsheetId,
          base64Data: base64Excel,
          folderId: '1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP'
        })
      });
      */

      showNotification(`✅ Toplam ${listToExport.length} transfer kaydı AT-802 Excel olarak Drive klasörüne (1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP) kaydedildi!`);
    } catch (e: any) {
      console.error('Push error:', e);
      showNotification(`Aktarım hatası: ${e.message || e}`);
    } finally {
      setIsPushingTransfers(false);
    }
  };

  // Google E-Tabloya Transfer Geçmişini Aktarma (Push to Google Sheets & Drive Excel)
  const pushTransfersToGoogleSheets = async () => {
    await uploadAt802TransfersExcelToDrive();
  };

  // Düzenleme Kaydetme & Eşleşen Malzemenin Stoğundan Otomatik Düşme
  const handleSaveEditedTransaction = (updatedTx: DepoTransaction) => {
    if (onUpdateTransaction) {
      onUpdateTransaction(updatedTx);
    }

    // Depodaki eşleşen ürünü bul ve stok düşümünü senkronize et
    let stockUpdated = false;
    if (onUpdateDepoRows && updatedTx.itemName && !updatedTx.itemName.startsWith('MALZEME #')) {
      const copyDepo = [...depoRows];
      const matchIdx = copyDepo.findIndex(d => {
        const dName = (d[1] || "").trim().toLowerCase();
        const dPn = (d[2] || "").trim().toLowerCase();
        const search = updatedTx.itemName.trim().toLowerCase();
        const searchPn = (updatedTx.pn || '').trim().toLowerCase();
        if (dName && dName === search) return true;
        if (dPn && dPn !== '-' && (dPn === search || dPn === searchPn)) return true;
        if (searchPn && searchPn !== '-' && dName === searchPn) return true;
        return false;
      });

      if (matchIdx >= 0) {
        const targetRow = [...copyDepo[matchIdx]];
        const curQty = parseFloat((targetRow[4] || "0").replace(/[^0-9.]/g, '')) || 0;
        const txQty = updatedTx.quantity || 1;
        const txType = (updatedTx.type || '').toUpperCase();

        if (txType.includes('ÇIKAN') || txType.includes('ÇIKIŞ') || txType.includes('CIKIS')) {
          targetRow[4] = `${Math.max(0, curQty - txQty)} ADET`;
          copyDepo[matchIdx] = targetRow;
          onUpdateDepoRows(copyDepo);
          stockUpdated = true;
        } else if (txType.includes('GİREN') || txType.includes('GİRİŞ') || txType.includes('GIRIS')) {
          targetRow[4] = `${curQty + txQty} ADET`;
          copyDepo[matchIdx] = targetRow;
          onUpdateDepoRows(copyDepo);
          stockUpdated = true;
        } else if (txType.includes('TRANSFER') && updatedTx.location) {
          targetRow[5] = updatedTx.location;
          copyDepo[matchIdx] = targetRow;
          onUpdateDepoRows(copyDepo);
          stockUpdated = true;
        }
      }
    }

    setEditingTx(null);
    if (stockUpdated) {
      showNotification(`✅ "${updatedTx.itemName}" transfer kaydı güncellendi ve depo stoğundan otomatik düşüldü / eşitlendi!`);
    } else {
      showNotification(`✅ "${updatedTx.itemName}" transfer kaydı başarıyla güncellendi!`);
    }
  };

  // Satır Silme
  const handleDeleteTxRow = (txId: string, itemName: string) => {
    if (window.confirm(`"${itemName}" işlem kaydını listeden silmek istediğinize emin misiniz?`)) {
      if (onDeleteTransaction) {
        onDeleteTransaction(txId);
      }
      setSelectedTxIds(prev => prev.filter(id => id !== txId));
      showNotification(`🗑️ Transfer kaydı silindi.`);
    }
  };

  // Sistemde Eşleşmeyen Toplam Kayıt Sayısı
  const totalUnmatchedCount = useMemo(() => {
    return transactions.filter(tx => isTxUnmatchedWithDepo(tx, depoRows)).length;
  }, [transactions, depoRows]);

  // Transfer Geçmişi Filtreli ve Eşleşmeyenleri En Başa Alan Listesi
  const filteredTransactions = useMemo(() => {
    const list = transactions.filter((tx) => {
      // 1. İşlem Türü Filtresi
      if (txTypeFilter !== 'ALL') {
        if (txTypeFilter === 'DİĞER') {
          const isStandard = STANDARD_TRANSACTION_TYPES.filter(t => t !== 'DİĞER').includes(tx.type);
          if (isStandard) return false;
        } else if (tx.type !== txTypeFilter) {
          return false;
        }
      }

      // 2. Hava Aracı Tipi Filtresi
      if (txAircraftTypeFilter !== 'ALL') {
        const tail = (tx.tailNo || '').toUpperCase();
        const itemInfo = `${tx.itemName} ${tx.notes || ''} ${tx.location || ''}`.toUpperCase();
        if (txAircraftTypeFilter === 'AT-802') {
          const isAt = tail.startsWith('OR-20') || itemInfo.includes('AT-802') || itemInfo.includes('AT802');
          if (!isAt) return false;
        } else if (txAircraftTypeFilter === 'BELL-429') {
          const isBell = tail.startsWith('OR 31') || tail.startsWith('OR-31') || itemInfo.includes('BELL') || itemInfo.includes('429');
          if (!isBell) return false;
        } else if (txAircraftTypeFilter === 'T-70') {
          const isT70 = tail.startsWith('OR-10') || tail.startsWith('OR 10') || itemInfo.includes('T-70') || itemInfo.includes('T70');
          if (!isT70) return false;
        } else if (txAircraftTypeFilter === 'B-360') {
          const isB360 = tail.includes('1839') || itemInfo.includes('B-360') || itemInfo.includes('B360') || itemInfo.includes('KING AIR');
          if (!isB360) return false;
        } else if (txAircraftTypeFilter === 'C-650') {
          const isC650 = tail.includes('0177') || itemInfo.includes('C-650') || itemInfo.includes('C650') || itemInfo.includes('CITATION');
          if (!isC650) return false;
        } else if (txAircraftTypeFilter === 'GSE') {
          const isGse = itemInfo.includes('YER DESTEK') || itemInfo.includes('GSE') || itemInfo.includes('YD');
          if (!isGse) return false;
        }
      }

      // 3. Kuyruk Kodu Filtresi
      if (txTailFilter !== 'ALL') {
        if ((tx.tailNo || '').trim() !== txTailFilter.trim()) return false;
      }

      // 4. Eşleşme / Düşüm Durumu Filtresi
      const isUnmatched = isTxUnmatchedWithDepo(tx, depoRows);
      if (txStatusFilter === 'UNMATCHED' && !isUnmatched) return false;
      if (txStatusFilter === 'MATCHED' && isUnmatched) return false;

      // 5. Arama Kutusu Filtresi
      if (txSearchFilter.trim()) {
        const q = txSearchFilter.trim().toLowerCase();
        const combined = `${tx.itemName} ${tx.pn} ${tx.sn} ${tx.tailNo} ${tx.operator} ${tx.receivedBy} ${tx.location} ${tx.type} ${tx.date} ${tx.notes || ''}`.toLowerCase();
        if (!combined.includes(q)) return false;
      }
      return true;
    });

    // Kullanıcı kuralı: Sistemde okunamayan/eşleşmeyen kayıtlar Kırmızı yanar söner ve listenin EN BAŞINA ATILIR!
    return list.sort((a, b) => {
      const aUnmatched = isTxUnmatchedWithDepo(a, depoRows) ? 1 : 0;
      const bUnmatched = isTxUnmatchedWithDepo(b, depoRows) ? 1 : 0;
      if (aUnmatched !== bUnmatched) {
        return bUnmatched - aUnmatched; // Eşleşmeyenler (1) önce gelir
      }
      return 0;
    });
  }, [transactions, txTypeFilter, txSearchFilter, txAircraftTypeFilter, txTailFilter, txStatusFilter, depoRows]);

  const isAllSelected = filteredTransactions.length > 0 && filteredTransactions.every(t => selectedTxIds.includes(t.id));
  const isSomeSelected = filteredTransactions.some(t => selectedTxIds.includes(t.id)) && !isAllSelected;

  const handleToggleSelectAll = () => {
    if (isAllSelected) {
      const filteredIds = new Set(filteredTransactions.map(t => t.id));
      setSelectedTxIds(prev => prev.filter(id => !filteredIds.has(id)));
    } else {
      const newIds = Array.from(new Set([...selectedTxIds, ...filteredTransactions.map(t => t.id)]));
      setSelectedTxIds(newIds);
    }
  };

  const handleToggleSelectRow = (txId: string) => {
    setSelectedTxIds(prev => 
      prev.includes(txId) ? prev.filter(id => id !== txId) : [...prev, txId]
    );
  };

  const handleDeleteSelectedTxs = () => {
    if (selectedTxIds.length === 0) return;
    if (window.confirm(`Seçilen ${selectedTxIds.length} adet transfer geçmişi kaydını kalıcı olarak silmek istediğinize emin misiniz?`)) {
      if (onDeleteTransactionsBatch) {
        onDeleteTransactionsBatch(selectedTxIds);
      } else if (onDeleteTransaction) {
        selectedTxIds.forEach(id => onDeleteTransaction(id));
      }
      showNotification(`🗑️ Seçilen ${selectedTxIds.length} adet transfer kaydı başarıyla silindi.`);
      setSelectedTxIds([]);
    }
  };

  // Yeni Ürün Girişi States
  const [newMaterialName, setNewMaterialName] = useState<string>('');
  const [newMaterialPn, setNewMaterialPn] = useState<string>('');
  const [newMaterialSn, setNewMaterialSn] = useState<string>('-');
  const [newMaterialLoc, setNewMaterialLoc] = useState<string>(DEPO_REGISTERED_LOCATIONS[0]);
  const [newMaterialStatus, setNewMaterialStatus] = useState<string>('FAAL');
  const [newMaterialExpiry, setNewMaterialExpiry] = useState<string>('-');
  const [newMaterialFirma, setNewMaterialFirma] = useState<string>('TÜBİTAK UME / Tedarikçi');

  // Sayım Modu States
  const [sayimCategoryFilter, setSayimCategoryFilter] = useState<MaterialCategoryType>('TÜMÜ');
  const [sayimLocationFilter, setSayimLocationFilter] = useState<string>('TÜMÜ');
  const [sayimCounts, setSayimCounts] = useState<{ [rowIdx: number]: number }>({});
  const [sayimSearch, setSayimSearch] = useState<string>('');

  // Kit Oluşturma States (SADECE AT-802 DEPO ÜRÜNLERİNDEN)
  const [kitName, setKitName] = useState<string>('');
  const [kitPn, setKitPn] = useState<string>('');
  const [kitQuantity, setKitQuantity] = useState<number>(1);
  const [kitLocation, setKitLocation] = useState<string>('KİT RAFI - AT-802');
  const [kitComponents, setKitComponents] = useState<KitComponent[]>([]);
  
  // Kit Parça Arama & Otomatik Doldurma
  const [compSearch, setCompSearch] = useState<string>('');
  const [isCompDropdownOpen, setIsCompDropdownOpen] = useState<boolean>(false);
  const [selectedCompRow, setSelectedCompRow] = useState<string[] | null>(null);
  const [compQtyPerKit, setCompQtyPerKit] = useState<number>(1);

  // Kit Önizleme Modalı State
  const [isKitPreviewOpen, setIsKitPreviewOpen] = useState<boolean>(false);

  // Giriş / Çıkış İşlemleri - Ön Onay ve Düzenleme Ekranı State'leri
  const [isPreConfirmOpen, setIsPreConfirmOpen] = useState<boolean>(false);
  const [pendingTxData, setPendingTxData] = useState<{
    type: string;
    selectedOperation: 'cikis' | 'giris' | 'transfer';
    itemName: string;
    pn: string;
    sn: string;
    quantity: number;
    date: string;
    tailNo: string;
    operator: string;
    receivedBy: string;
    location: string;
    sourceLocation: string;
    targetLocation: string;
    notes: string;
    currentStock: number;
    newStockAfter: number;
    targetRowIdx: number;
    isNewItem: boolean;
    newMaterialStatus?: string;
    newMaterialExpiry?: string;
    newMaterialFirma?: string;
  } | null>(null);

  // Personel filtreleme listeleri (Yazıldıkça daralan arama)
  const filteredDeliveredPersonnel = useMemo(() => {
    if (!opDeliveredTo.trim()) return personnelList.slice(0, 10);
    const q = opDeliveredTo.toLocaleLowerCase('tr-TR');
    return personnelList.filter(p => p.toLocaleLowerCase('tr-TR').includes(q)).slice(0, 15);
  }, [personnelList, opDeliveredTo]);

  const filteredAcceptedPersonnel = useMemo(() => {
    if (!opAcceptedBy.trim()) return personnelList.slice(0, 10);
    const q = opAcceptedBy.toLocaleLowerCase('tr-TR');
    return personnelList.filter(p => p.toLocaleLowerCase('tr-TR').includes(q)).slice(0, 15);
  }, [personnelList, opAcceptedBy]);

  if (!isOpen) return null;

  // Tüm kayıtlı lokasyonlar (İl bazlı + depolardaki mevcut raflar)
  const allAvailableLocations = useMemo(() => {
    const fromRows = depoRows.map(r => (r[5] || "").trim()).filter(Boolean);
    const combined = Array.from(new Set([...DEPO_REGISTERED_LOCATIONS, ...fromRows]));
    return combined;
  }, [depoRows]);

  // Malzeme arama filtreleme (Giriş / Çıkış için)
  const filteredMaterials = useMemo(() => {
    if (!materialSearchQuery.trim()) {
      return depoRows.map((row, idx) => ({ row, idx })).slice(0, 15);
    }
    const q = materialSearchQuery.toLowerCase();
    return depoRows
      .map((row, idx) => ({ row, idx }))
      .filter(({ row }) => 
        (row[1] || "").toLowerCase().includes(q) ||
        (row[2] || "").toLowerCase().includes(q) ||
        (row[3] || "").toLowerCase().includes(q) ||
        (row[5] || "").toLowerCase().includes(q)
      )
      .slice(0, 20);
  }, [depoRows, materialSearchQuery]);

  // AT-802 Parça arama filtreleme (Kit Oluşturma için SADECE AT-802)
  const filteredAt802Parts = useMemo(() => {
    if (!compSearch.trim()) {
      return effectiveAt802Rows.slice(0, 10);
    }
    const q = compSearch.toLowerCase();
    return effectiveAt802Rows
      .filter(row => 
        (row[1] || "").toLowerCase().includes(q) ||
        (row[2] || "").toLowerCase().includes(q) ||
        (row[3] || "").toLowerCase().includes(q)
      )
      .slice(0, 15);
  }, [effectiveAt802Rows, compSearch]);

  // Seçili malzemenin detayları
  const currentSelectedItem = selectedItemIndex >= 0 && selectedItemIndex < depoRows.length ? depoRows[selectedItemIndex] : null;

  // Seçili veya aranan malzeme ile eşleşen tüm depo stok satırları (Farklı Seri Numaraları ve Raflar)
  const matchingStockRows = useMemo(() => {
    if (selectedItemIndex >= 0 && selectedItemIndex < depoRows.length) {
      const baseRow = depoRows[selectedItemIndex];
      const baseName = (baseRow[1] || "").trim().toLowerCase();
      const basePn = (baseRow[2] || "").trim().toLowerCase();
      
      return depoRows
        .map((row, idx) => ({ row, idx }))
        .filter(({ row }) => {
          const rName = (row[1] || "").trim().toLowerCase();
          const rPn = (row[2] || "").trim().toLowerCase();
          if (basePn && basePn !== '-' && rPn === basePn) return true;
          if (baseName && rName === baseName) return true;
          return false;
        });
    } else if (materialSearchQuery.trim()) {
      const q = materialSearchQuery.trim().toLowerCase();
      return depoRows
        .map((row, idx) => ({ row, idx }))
        .filter(({ row }) => {
          const rName = (row[1] || "").trim().toLowerCase();
          const rPn = (row[2] || "").trim().toLowerCase();
          return (rName && rName.includes(q)) || (rPn && rPn !== '-' && rPn.includes(q));
        });
    }
    return [];
  }, [depoRows, selectedItemIndex, materialSearchQuery]);

  // Malzeme Seçimi
  const handleSelectMaterial = (idx: number) => {
    setSelectedItemIndex(idx);
    const row = depoRows[idx];
    if (row) {
      setMaterialSearchQuery(`${row[1]} [P/N: ${row[2] || '-'}]`);
      if (row[5]) setOpDepoLocation(row[5]);
      if (row[3] && row[3] !== '-') {
        setOpSerialNo(row[3]);
      } else {
        setOpSerialNo('-');
      }
    }
    setIsMaterialDropdownOpen(false);
  };

  // Depo İşlemi Ön Onay Ekranını Aç (Doğrulama ve Hazırlık)
  const handleOpenPreConfirm = () => {
    if (!opTransactionType) {
      alert("Lütfen bir İŞLEM TÜRÜ seçiniz!");
      return;
    }

    if (opTransactionType === 'DİĞER' && !opCustomTransactionType.trim()) {
      alert("Lütfen manuel işlem türü açıklamasını yazınız!");
      return;
    }

    const effectiveType = opTransactionType === 'DİĞER'
      ? (opCustomTransactionType.trim() || 'DİĞER')
      : (opTransactionType || (selectedOperation === 'cikis' ? 'ANKARA ÇIKAN' : 'ANKARA GİREN'));

    if (selectedOperation === 'giris' && girisMode === 'yeni') {
      if (!newMaterialName.trim()) {
        alert("Lütfen malzeme adını giriniz.");
        return;
      }
      if (opQuantity <= 0) {
        alert("Lütfen geçerli bir miktar (en az 1) giriniz.");
        return;
      }

      setPendingTxData({
        type: effectiveType,
        selectedOperation: 'giris',
        itemName: newMaterialName.trim(),
        pn: newMaterialPn.trim() || "-",
        sn: newMaterialSn.trim() || "-",
        quantity: opQuantity,
        date: opDate,
        tailNo: opTailNo || '-',
        operator: opDeliveredTo || opOperator,
        receivedBy: opAcceptedBy,
        location: newMaterialLoc.trim() || DEPO_REGISTERED_LOCATIONS[0],
        sourceLocation: newMaterialFirma || 'Tedarikçi',
        targetLocation: newMaterialLoc.trim() || DEPO_REGISTERED_LOCATIONS[0],
        notes: opNotes.trim(),
        currentStock: 0,
        newStockAfter: opQuantity,
        targetRowIdx: -1,
        isNewItem: true,
        newMaterialStatus,
        newMaterialExpiry,
        newMaterialFirma
      });
      setIsPreConfirmOpen(true);
      return;
    }

    if (!currentSelectedItem || selectedItemIndex < 0) {
      alert("Lütfen arama alanına yazarak işlem yapılacak bir malzeme seçiniz.");
      return;
    }

    const targetRow = currentSelectedItem;
    const itemName = targetRow[1] || "";
    const pn = targetRow[2] || "";
    const sn = opSerialNo !== '-' ? opSerialNo : (targetRow[3] || "-");
    const curQty = parseFloat((targetRow[4] || "0").replace(/[^0-9.]/g, '')) || 0;
    const curLoc = targetRow[5] || "DEPO";

    const isCikisOp = selectedOperation === 'cikis' || effectiveType.toUpperCase().includes('ÇIKAN') || effectiveType.toUpperCase().includes('ÇIKIŞ');
    const isGirisOp = selectedOperation === 'giris' || effectiveType.toUpperCase().includes('GİREN') || effectiveType.toUpperCase().includes('GİRİŞ') || effectiveType.toUpperCase().includes('GELEN');

    if (opQuantity <= 0) {
      alert("Lütfen geçerli bir miktar (en az 1) giriniz.");
      return;
    }

    if (isCikisOp && opQuantity > curQty) {
      alert(`Yetersiz stok! Depoda mevcut stok: ${curQty} adet.`);
      return;
    }

    let calculatedStockAfter = curQty;
    if (isCikisOp) {
      calculatedStockAfter = Math.max(0, curQty - opQuantity);
    } else if (isGirisOp) {
      calculatedStockAfter = curQty + opQuantity;
    }

    setPendingTxData({
      type: effectiveType,
      selectedOperation,
      itemName,
      pn,
      sn,
      quantity: opQuantity,
      date: opDate,
      tailNo: opTailNo || '-',
      operator: opDeliveredTo || opOperator,
      receivedBy: opAcceptedBy,
      location: opDepoLocation || curLoc,
      sourceLocation: curLoc,
      targetLocation: selectedOperation === 'transfer' ? (opDepoLocation || opTargetLocation) : (opTailNo || 'Hava Aracı'),
      notes: opNotes.trim(),
      currentStock: curQty,
      newStockAfter: calculatedStockAfter,
      targetRowIdx: selectedItemIndex,
      isNewItem: false
    });
    setIsPreConfirmOpen(true);
  };

  // Ön Onay Ekranından Kesin Onay ve Kaydetme (Stoktan Düş / Arttır & Google E-Tabloya Yaz)
  const handleConfirmAndSaveTransaction = async () => {
    if (!pendingTxData) return;

    const tx = pendingTxData;
    const effectiveType = tx.type.trim() || 'ANKARA ÇIKAN';
    const isCikisOp = tx.selectedOperation === 'cikis' || effectiveType.toUpperCase().includes('ÇIKAN') || effectiveType.toUpperCase().includes('ÇIKIŞ');
    const isGirisOp = tx.selectedOperation === 'giris' || effectiveType.toUpperCase().includes('GİREN') || effectiveType.toUpperCase().includes('GİRİŞ') || effectiveType.toUpperCase().includes('GELEN');
    const isTransferOp = tx.selectedOperation === 'transfer' || effectiveType.toUpperCase().includes('TRANSFER');

    if (tx.isNewItem) {
      const newRow: string[] = [
        "",
        tx.itemName.trim(),
        tx.pn.trim() || "-",
        tx.sn.trim() || "-",
        `${tx.quantity} ADET`,
        tx.location.trim() || DEPO_REGISTERED_LOCATIONS[0],
        tx.newMaterialStatus || 'FAAL',
        (tx.newMaterialExpiry && tx.newMaterialExpiry !== "-") ? "EVET" : "HAYIR",
        tx.newMaterialExpiry || '-',
        tx.newMaterialFirma || 'TÜBİTAK UME / Tedarikçi',
        tx.notes.trim() || "Yeni malzeme girişi"
      ];

      onUpdateDepoRows([...depoRows, newRow]);

      onAddTransaction({
        type: effectiveType,
        itemName: tx.itemName.trim(),
        pn: tx.pn.trim() || "-",
        sn: tx.sn.trim() || "-",
        quantity: tx.quantity,
        date: tx.date,
        timestamp: `${tx.date} ${new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' })}`,
        tailNo: tx.tailNo || '-',
        operator: tx.operator,
        receivedBy: tx.receivedBy,
        location: tx.location.trim() || 'DEPO',
        sourceLocation: tx.sourceLocation || 'Tedarikçi',
        targetLocation: tx.location.trim() || 'DEPO',
        notes: `Yeni Malzeme Girişi | Tür: ${effectiveType} | Kabul: ${tx.receivedBy} | Teslim: ${tx.operator}${tx.notes ? ' - ' + tx.notes : ''}`
      });

      showNotification(`"${tx.itemName}" (${tx.quantity} adet) başarıyla yeni malzeme olarak depoya eklendi!`);
    } else if (tx.targetRowIdx >= 0 && tx.targetRowIdx < depoRows.length) {
      const targetRow = [...depoRows[tx.targetRowIdx]];
      targetRow[1] = tx.itemName;
      targetRow[2] = tx.pn;
      targetRow[3] = tx.sn;

      if (isCikisOp) {
        const curQty = parseFloat((targetRow[4] || "0").replace(/[^0-9.]/g, '')) || 0;
        const newQty = Math.max(0, curQty - tx.quantity);
        targetRow[4] = `${newQty} ADET`;
        targetRow[5] = tx.location || targetRow[5] || "DEPO";
        const updatedRows = [...depoRows];
        updatedRows[tx.targetRowIdx] = targetRow;
        onUpdateDepoRows(updatedRows);

        onAddTransaction({
          type: effectiveType,
          itemName: tx.itemName,
          pn: tx.pn,
          sn: tx.sn,
          quantity: tx.quantity,
          date: tx.date,
          timestamp: `${tx.date} ${new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' })}`,
          tailNo: tx.tailNo || '-',
          operator: tx.operator,
          receivedBy: tx.receivedBy,
          location: tx.location,
          sourceLocation: tx.sourceLocation,
          targetLocation: tx.tailNo || 'Hava Aracı',
          notes: `${effectiveType} | Kuyruk: ${tx.tailNo} | Teslim: ${tx.operator} | Kabul: ${tx.receivedBy} | Depo: ${tx.location}${tx.notes ? ' - ' + tx.notes : ''}`
        });

        showNotification(`${tx.quantity} adet ${tx.itemName} (${tx.tailNo}) çıkışı onaylandı ve stoktan düşüldü!`);
      } else if (isGirisOp) {
        const curQty = parseFloat((targetRow[4] || "0").replace(/[^0-9.]/g, '')) || 0;
        const newQty = curQty + tx.quantity;
        targetRow[4] = `${newQty} ADET`;
        targetRow[5] = tx.location || targetRow[5] || "DEPO";
        const updatedRows = [...depoRows];
        updatedRows[tx.targetRowIdx] = targetRow;
        onUpdateDepoRows(updatedRows);

        onAddTransaction({
          type: effectiveType,
          itemName: tx.itemName,
          pn: tx.pn,
          sn: tx.sn,
          quantity: tx.quantity,
          date: tx.date,
          timestamp: `${tx.date} ${new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' })}`,
          tailNo: tx.tailNo || '-',
          operator: tx.operator,
          receivedBy: tx.receivedBy,
          location: tx.location,
          sourceLocation: 'Tedarikçi / İntikal',
          targetLocation: tx.location,
          notes: `${effectiveType} | Depoya Giriş | Kabul: ${tx.receivedBy} | Teslim: ${tx.operator}${tx.notes ? ' - ' + tx.notes : ''}`
        });

        showNotification(`${tx.quantity} adet ${tx.itemName} stok girişi onaylandı ve stok arttırıldı!`);
      } else if (isTransferOp) {
        targetRow[5] = tx.location;
        const updatedRows = [...depoRows];
        updatedRows[tx.targetRowIdx] = targetRow;
        onUpdateDepoRows(updatedRows);

        onAddTransaction({
          type: effectiveType,
          itemName: tx.itemName,
          pn: tx.pn,
          sn: tx.sn,
          quantity: tx.quantity,
          date: tx.date,
          timestamp: `${tx.date} ${new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' })}`,
          tailNo: tx.tailNo || '-',
          operator: tx.operator,
          receivedBy: tx.receivedBy,
          location: tx.location,
          sourceLocation: tx.sourceLocation,
          targetLocation: tx.location,
          notes: `Transfer (${effectiveType}): ${tx.sourceLocation} ➜ ${tx.location} | Teslim: ${tx.operator} | Kabul: ${tx.receivedBy}${tx.notes ? ' - ' + tx.notes : ''}`
        });

        showNotification(`${tx.itemName} başarıyla ${tx.location} lokasyonuna transfer edildi!`);
      }
    }

    // Google E-Tablo Online Senkronizasyon İsteği (Arka Planda Gönder)
    try {
      const payload = {
        action: 'appendDepoTransaction',
        transaction: {
          itemName: tx.itemName,
          quantity: tx.quantity,
          date: tx.date,
          type: effectiveType,
          sn: tx.sn,
          tailNo: tx.tailNo,
          operator: tx.operator,
          receivedBy: tx.receivedBy,
          location: tx.location,
          notes: tx.notes
        }
      };
      fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      }).catch(() => {});
    } catch (_err) {}

    // Formu temizle ve Ön Onay Modalı kapat
    setIsPreConfirmOpen(false);
    setPendingTxData(null);
    setMaterialSearchQuery('');
    setSelectedItemIndex(-1);
    setNewMaterialName('');
    setNewMaterialPn('');
    setOpQuantity(1);
    setOpTransactionType('');
    setOpCustomTransactionType('');
    setOpSerialNo('');
    setOpTailNo('');
    setOpDeliveredTo('');
    setOpAcceptedBy('');
    setOpDepoLocation('');
    setOpNotes('');
  };

  // Örnek Excel Şablonu İndirme (Kullanıcı Formatı: 9 Sütun)
  const downloadSampleOperationExcel = () => {
    try {
      const sampleData = [
        {
          "MALZEME ADI": "2111",
          "ADET": 3,
          "TARİH": "5.05.2023",
          "İŞLEM TÜRÜ": "ANKARA ÇIKAN",
          "SERİAL NUMBER": "-",
          "KUYRUK KODU": "OR-2024",
          "TESLİM ALAN": "T.GÜZER",
          "KABUL YAPAN": "A.ÖZMETİN",
          "DEPO YERİ": "M-70 (8BATCH PALET1 BOX2)"
        },
        {
          "MALZEME ADI": "60795-1",
          "ADET": 1,
          "TARİH": "5.05.2023",
          "İŞLEM TÜRÜ": "ANKARA ÇIKAN",
          "SERİAL NUMBER": "-",
          "KUYRUK KODU": "OR-2025",
          "TESLİM ALAN": "T.GÜZER",
          "KABUL YAPAN": "A.ÖZMETİN",
          "DEPO YERİ": "2-D (YD1 BOX41)"
        },
        {
          "MALZEME ADI": "606C86-3 , AA1E6",
          "ADET": 1,
          "TARİH": "2.06.2023",
          "İŞLEM TÜRÜ": "ANKARA ÇIKAN",
          "SERİAL NUMBER": "-",
          "KUYRUK KODU": "OR-2023",
          "TESLİM ALAN": "Y.KARADENİZ",
          "KABUL YAPAN": "T.GÜZER",
          "DEPO YERİ": "PALET 1/14/16 (8BATCH PALET4)"
        },
        {
          "MALZEME ADI": "302-246-401",
          "ADET": 1,
          "TARİH": "2.06.2023",
          "İŞLEM TÜRÜ": "ANKARA ÇIKAN",
          "SERİAL NUMBER": "-",
          "KUYRUK KODU": "OR-2023",
          "TESLİM ALAN": "Y.KARADENİZ",
          "KABUL YAPAN": "T.GÜZER",
          "DEPO YERİ": "BOX-27 (YD3 BOX60-61-62-63-13-14) (8BATCH PALET3 BOX32)"
        }
      ];

      const ws = XLSX.utils.json_to_sheet(sampleData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Depo_Transfer_Sablonu");
      XLSX.writeFile(wb, "Depo_Transfer_Islem_Sablonu.xlsx");
      showNotification("Örnek Excel şablonu (9 Sütunlu) başarıyla indirildi!");
    } catch (err) {
      alert("Şablon indirilirken hata oluştu: " + (err as Error).message);
    }
  };

  // Toplu Excel ile Depo İşlemleri Yükleme (Kullanıcı Tablosu Başlıkları: 9 Sütun - Hızlı & Donmayan Yapı)
  const handleBulkExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessingBulk(true);
    showNotification("Toplu Excel dosyası işleniyor, lütfen bekleyin...");

    try {
      const arrayBuffer = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const wsName = wb.SheetNames[0];
      const ws = wb.Sheets[wsName];

      // Birleştirilmiş hücreleri unroll yap
      unmergeAndFillWorksheet(ws);

      // 2D Array olarak oku (Başlıkları ve tüm hücreleri kayıpsız almak için)
      const sheetRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      if (!sheetRows || sheetRows.length === 0) {
        alert("Yüklenen Excel dosyasında veri bulunamadı.");
        setIsProcessingBulk(false);
        return;
      }

      // 1. Başlık Satırını Bul (Kullanıcı uyarısı: Başlıklar ilk satırdır, onu veri olarak saymadan sonraki satırları işle)
      let headerRowIdx = 0;
      for (let r = 0; r < Math.min(sheetRows.length, 5); r++) {
        const rowStr = (sheetRows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
        if (
          rowStr.includes('MALZEME') ||
          rowStr.includes('PARÇA') ||
          rowStr.includes('PARCA') ||
          rowStr.includes('ADET') ||
          rowStr.includes('TARİH') ||
          rowStr.includes('TARIH') ||
          rowStr.includes('İŞLEM') ||
          rowStr.includes('ISLEM') ||
          rowStr.includes('P/N')
        ) {
          headerRowIdx = r;
          break;
        }
      }

      const headerRow = sheetRows[headerRowIdx] || [];
      const normalizeH = (s: any) => String(s || '').toUpperCase().replace(/İ/g, 'I').replace(/Ğ/g, 'G').replace(/Ü/g, 'U').replace(/Ş/g, 'S').replace(/Ö/g, 'O').replace(/Ç/g, 'C').trim();

      // Sütun İndekslerini Dinamik Olarak Tespit Et
      let colName = -1, colQty = -1, colDate = -1, colType = -1, colSn = -1, colTail = -1, colTeslim = -1, colKabul = -1, colDepo = -1, colNotes = -1;

      headerRow.forEach((cell, idx) => {
        const h = normalizeH(cell);
        if (h.includes('MALZEME') || h.includes('PARCA') || h.includes('TANIM') || h.includes('CINS') || h.includes('ITEM') || h.includes('URUN')) {
          if (colName === -1) colName = idx;
        } else if (h.includes('ADET') || h.includes('MIKTAR') || h.includes('QTY') || h.includes('SAYI')) {
          if (colQty === -1) colQty = idx;
        } else if (h.includes('TARIH') || h.includes('DATE') || h.includes('GUN') || h.includes('ZAMAN')) {
          if (colDate === -1) colDate = idx;
        } else if (h.includes('ISLEM') || h.includes('TUR') || h.includes('HAREKET') || h.includes('DURUM')) {
          if (colType === -1) colType = idx;
        } else if (h.includes('SERIAL') || h.includes('SERI NO') || h.includes('S/N') || h.includes('SN') || h.includes('SERI')) {
          if (colSn === -1) colSn = idx;
        } else if (h.includes('KUYRUK') || h.includes('TAIL') || h.includes('UCAK') || h.includes('HELIKOPTER') || h.includes('ARAC')) {
          if (colTail === -1) colTail = idx;
        } else if (h.includes('TESLIM') || h.includes('CIKARAN') || h.includes('YAPAN') || h.includes('ALAN') || h.includes('PERSONEL')) {
          if (colTeslim === -1) colTeslim = idx;
        } else if (h.includes('KABUL') || h.includes('ONAY') || h.includes('EDEN') || h.includes('VEREN')) {
          if (colKabul === -1) colKabul = idx;
        } else if (h.includes('DEPO') || h.includes('LOKASYON') || h.includes('RAF') || h.includes('KUTU') || h.includes('PALET') || h.includes('YER')) {
          if (colDepo === -1) colDepo = idx;
        } else if (h.includes('ACIKLAMA') || h.includes('NOT') || h.includes('DETAY')) {
          if (colNotes === -1) colNotes = idx;
        }
      });

      // Default Sütun Fallbackleri (9 Sütunlu Standart Kullanıcı Formatı)
      if (colName === -1) colName = 0;
      if (colQty === -1) colQty = 1;
      if (colDate === -1) colDate = 2;
      if (colType === -1) colType = 3;
      if (colSn === -1) colSn = 4;
      if (colTail === -1) colTail = 5;
      if (colTeslim === -1) colTeslim = 6;
      if (colKabul === -1) colKabul = 7;
      if (colDepo === -1) colDepo = 8;

      let updatedDepo = [...depoRows];
      const newTxBatch: (Omit<DepoTransaction, 'id'> | (Omit<DepoTransaction, 'id' | 'timestamp'> & { timestamp?: string }))[] = [];
      let processedCount = 0;

      // Başlık satırını saymadan (headerRowIdx + 1) veri satırlarını tara
      for (let i = headerRowIdx + 1; i < sheetRows.length; i++) {
        const r = sheetRows[i];
        if (!r || r.length === 0) continue;

        // Tüm satır boş mu kontrol et
        const hasContent = r.some(c => c !== undefined && c !== null && String(c).trim() !== '');
        if (!hasContent) continue;

        // Malzeme adını oku (A sütunu veya tespit edilen Malzeme sütunu, örn: 2111, 60795-1, 606C86-3 , AA1E6)
        const valColName = r[colName] !== undefined && r[colName] !== null ? String(r[colName]).trim() : '';
        const valCol0 = r[0] !== undefined && r[0] !== null ? String(r[0]).trim() : '';
        let targetItemName = valColName || valCol0;
        
        // Tekrarlanan başlık satırı ise atla
        const normItem = normalizeH(targetItemName);
        if (normItem === 'MALZEME ADI' || normItem === 'MALZEME' || normItem === 'PARCA NO' || normItem === 'PARCA' || normItem === 'ITEM') continue;

        if (!targetItemName) {
          // Eğer ana sütun boşsa diğer sütunlardan ilk dolu metni dene
          const fallback = r.find(c => c !== undefined && c !== null && String(c).trim() !== '');
          targetItemName = fallback ? String(fallback).trim() : `MALZEME #${processedCount + 1}`;
        }

        const rawQtyCell = r[colQty];
        let rawQty = 1;
        if (typeof rawQtyCell === 'number') {
          rawQty = rawQtyCell;
        } else {
          const parsed = parseFloat(String(rawQtyCell || '1').replace(/[^0-9.]/g, ''));
          rawQty = isNaN(parsed) || parsed <= 0 ? 1 : parsed;
        }

        const rawDate = parseExcelBulkDate(r[colDate]);
        const rawType = String(r[colType] !== undefined && r[colType] !== null && String(r[colType]).trim() !== '' ? r[colType] : 'ANKARA ÇIKAN').trim();
        const rawSn = String(r[colSn] !== undefined && r[colSn] !== null && String(r[colSn]).trim() !== '' ? r[colSn] : '-').trim();
        const rawTail = String(r[colTail] !== undefined && r[colTail] !== null && String(r[colTail]).trim() !== '' ? r[colTail] : '-').trim();
        const rawTeslim = String(r[colTeslim] !== undefined && r[colTeslim] !== null && String(r[colTeslim]).trim() !== '' ? r[colTeslim] : '-').trim();
        const rawKabul = String(r[colKabul] !== undefined && r[colKabul] !== null && String(r[colKabul]).trim() !== '' ? r[colKabul] : '-').trim();
        const rawDepoYeri = String(r[colDepo] !== undefined && r[colDepo] !== null && String(r[colDepo]).trim() !== '' ? r[colDepo] : 'DEPO').trim();

        const isCikis = rawType.toUpperCase().includes("ÇIKAN") || rawType.toUpperCase().includes("ÇIKIŞ") || rawType.toUpperCase().includes("CIKIS");
        const isGiris = rawType.toUpperCase().includes("GİREN") || rawType.toUpperCase().includes("GİRİŞ") || rawType.toUpperCase().includes("GIRIS") || rawType.toUpperCase().includes("GELEN");
        const isTransfer = rawType.toUpperCase().includes("TRANSFER");

        // Depoda eşleşen satırı bul (P/N veya İsme veya Alfanumerik Temiz Metne göre)
        const matchIdx = updatedDepo.findIndex(d => {
          const dName = (d[1] || "").trim().toLowerCase();
          const dPn = (d[2] || "").trim().toLowerCase();
          const search = targetItemName.trim().toLowerCase();
          if (dName && dName === search) return true;
          if (dPn && dPn !== '-' && dPn === search) return true;
          
          const cleanSearch = search.replace(/[^a-z0-9]/g, '');
          const cleanDName = dName.replace(/[^a-z0-9]/g, '');
          const cleanDPn = dPn.replace(/[^a-z0-9]/g, '');
          if (cleanSearch && cleanSearch.length >= 2 && (cleanSearch === cleanDName || cleanSearch === cleanDPn)) return true;
          if (cleanSearch && cleanSearch.length >= 3 && (cleanDName.includes(cleanSearch) || cleanDPn.includes(cleanSearch))) return true;
          return false;
        });

        let curLoc = rawDepoYeri || "DEPO";
        let targetPn = "-";

        if (matchIdx >= 0) {
          const targetRow = [...updatedDepo[matchIdx]];
          const curQty = parseFloat((targetRow[4] || "0").replace(/[^0-9.]/g, '')) || 0;
          curLoc = targetRow[5] || rawDepoYeri || "DEPO";
          targetItemName = targetRow[1] || targetItemName;
          targetPn = targetRow[2] || "-";

          if (isCikis) {
            targetRow[4] = `${Math.max(0, curQty - rawQty)} ADET`;
            updatedDepo[matchIdx] = targetRow;
          } else if (isGiris) {
            targetRow[4] = `${curQty + rawQty} ADET`;
            updatedDepo[matchIdx] = targetRow;
          } else if (isTransfer) {
            targetRow[5] = rawDepoYeri;
            updatedDepo[matchIdx] = targetRow;
          }
        } else if (isGiris) {
          const newRow: string[] = [
            "",
            targetItemName,
            "-",
            rawSn || "-",
            `${rawQty} ADET`,
            rawDepoYeri || DEPO_REGISTERED_LOCATIONS[0],
            "FAAL",
            "HAYIR",
            "-",
            "-",
            "Toplu Excel Aktarımı",
            `Eklenme: ${rawDate}`
          ];
          updatedDepo.push(newRow);
        }

        newTxBatch.push({
          type: rawType,
          itemName: targetItemName,
          pn: targetPn,
          sn: rawSn,
          quantity: rawQty,
          date: rawDate,
          timestamp: `${rawDate} ${new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' })}`,
          tailNo: rawTail,
          operator: rawTeslim,
          receivedBy: rawKabul,
          location: rawDepoYeri,
          sourceLocation: curLoc,
          targetLocation: rawTail !== '-' ? rawTail : rawDepoYeri,
          notes: `${rawType} | Kuyruk: ${rawTail} | Teslim: ${rawTeslim} | Kabul: ${rawKabul}`
        });
        processedCount++;
      }

      onUpdateDepoRows(updatedDepo);
      if (onAddTransactionsBatch) {
        onAddTransactionsBatch(newTxBatch);
      } else {
        newTxBatch.forEach(tx => onAddTransaction(tx));
      }

      showNotification(`✅ Toplu Excel Başarılı! Toplam ${processedCount} adet işlem Transfer Geçmişi'ne eklendi.`);
      if (bulkFileInputRef.current) bulkFileInputRef.current.value = '';
      setActiveTab('loglar');
    } catch (err) {
      alert("Excel dosyası okunurken hata oluştu: " + (err as Error).message);
    } finally {
      setIsProcessingBulk(false);
    }
  };

  // Kit Bileşeni Ekleme (Seçilen AT-802 parçasını ekle)
  const handleAddSelectedKitComp = (row: string[], qty: number = 1) => {
    const pn = row[2] || row[1] || "";
    const itemName = row[1] || "";
    const curStock = parseFloat((row[4] || "0").replace(/[^0-9.]/g, '')) || 0;

    if (kitComponents.some(c => c.pn === pn && c.itemName === itemName)) {
      alert("Bu parça zaten kit listesinde mevcut.");
      return;
    }

    setKitComponents(prev => [
      ...prev,
      {
        pn,
        itemName,
        quantityPerKit: qty,
        totalNeeded: qty * kitQuantity,
        currentStock: curStock,
      }
    ]);

    setSelectedCompRow(null);
    setCompSearch('');
    setIsCompDropdownOpen(false);
    setCompQtyPerKit(1);
  };

  // "KİTİ OLUŞTUR" butonuna basınca Önizleme Ekranını Açma
  const handleOpenKitPreview = () => {
    if (!kitName.trim()) {
      alert("Lütfen Kit Adı giriniz (Örn: 100 Saatlik Bakım Kiti).");
      return;
    }
    if (kitComponents.length === 0) {
      alert("Lütfen kite en az bir parça ekleyiniz.");
      return;
    }

    // Stok Yeterlilik Kontrolü
    for (const comp of kitComponents) {
      const needed = comp.quantityPerKit * kitQuantity;
      if (needed > comp.currentStock) {
        alert(`Yetersiz stok! "${comp.itemName}" için ${needed} adet gerekiyor ancak AT-802 deposunda ${comp.currentStock} adet var.`);
        return;
      }
    }

    setIsKitPreviewOpen(true);
  };

  // Kit Önizleme Excelini İndirme
  const handleDownloadKitPreviewExcel = () => {
    try {
      const headerInfo = [
        { "PARAMETRE": "KİT ADI", "DEĞER": kitName.toUpperCase() },
        { "PARAMETRE": "KİT PARÇA NO (P/N)", "DEĞER": kitPn.toUpperCase() || `KIT-AT802-${Date.now().toString().slice(-4)}` },
        { "PARAMETRE": "ÜRETİLECEK KİT ADEDİ", "DEĞER": `${kitQuantity} KİT` },
        { "PARAMETRE": "HEDEF LOKASYON / RAF", "DEĞER": kitLocation },
        { "PARAMETRE": "KAYNAK DEPO", "DEĞER": "AT-802 SARF & PARÇA DEPOSU" },
        { "PARAMETRE": "OLUŞTURMA TARİHİ", "DEĞER": new Date().toLocaleString('tr-TR') },
      ];

      const partsTable = kitComponents.map((c, i) => ({
        "SIRA": i + 1,
        "BİLEŞEN PARÇA ADI": c.itemName,
        "PARÇA NO (P/N)": c.pn,
        "1 KİT İÇİN GEREKLİ": c.quantityPerKit,
        "TOPLAM DÜŞÜLECEK": c.quantityPerKit * kitQuantity,
        "MEVCUT DEPO STOĞU": c.currentStock,
        "KALAN DEPO STOĞU": c.currentStock - (c.quantityPerKit * kitQuantity)
      }));

      const wb = XLSX.utils.book_new();
      const wsInfo = XLSX.utils.json_to_sheet(headerInfo);
      const wsParts = XLSX.utils.json_to_sheet(partsTable);

      XLSX.utils.book_append_sheet(wb, wsInfo, "Kit_Genel_Bilgiler");
      XLSX.utils.book_append_sheet(wb, wsParts, "Kit_Bilesen_Parcalari");

      XLSX.writeFile(wb, `Kit_Onizleme_${kitName.replace(/\s+/g, '_')}_${Date.now().toString().slice(-4)}.xlsx`);
      showNotification("Kit Önizleme Exceli başarıyla indirildi!");
    } catch (err) {
      alert("Önizleme Exceli indirilirken hata: " + (err as Error).message);
    }
  };

  // Kiti Onayla, AT-802 Depoya Yeni Ürün Olarak Ekle ve Bileşen Stoklarını Düş
  const handleConfirmAndSaveKit = () => {
    // 1. Bileşen stoklarını AT-802 deposundan düş
    const updatedAt802 = effectiveAt802Rows.map(row => {
      const rowPn = row[2] || row[1] || "";
      const rowName = row[1] || "";
      const matchedComp = kitComponents.find(c => c.pn === rowPn && c.itemName === rowName);
      if (matchedComp) {
        const copy = [...row];
        const curStock = parseFloat((copy[4] || "0").replace(/[^0-9.]/g, '')) || 0;
        const deduct = matchedComp.quantityPerKit * kitQuantity;
        copy[4] = `${Math.max(0, curStock - deduct)} ADET`;
        return copy;
      }
      return row;
    });

    // 2. Yeni Kiti listenin en altına ekle
    const finalKitPn = kitPn.trim().toUpperCase() || `KIT-AT802-${Date.now().toString().slice(-4)}`;
    const finalKitSn = `KIT-SN-${Date.now().toString().slice(-6)}`;
    const newKitRow: string[] = [
      "",
      `📦 [KİT] ${kitName.trim().toUpperCase()}`,
      finalKitPn,
      finalKitSn,
      `${kitQuantity} KİT`,
      kitLocation.trim().toUpperCase(),
      "FAAL",
      "HAYIR",
      "-",
      "KİT ÜRETİMİ (ATÖLYE)",
      `İçerik: ${kitComponents.map(c => `${c.quantityPerKit * kitQuantity}x ${c.itemName}`).join(', ')}`
    ];

    const finalAt802WithKit = [...updatedAt802, newKitRow];

    if (onUpdateAt802Rows) {
      onUpdateAt802Rows(finalAt802WithKit);
    } else {
      onUpdateDepoRows(finalAt802WithKit);
    }

    // 3. Log kaydı oluştur
    onAddTransaction({
      type: 'KİT OLUŞTURMA',
      itemName: `📦 [KİT] ${kitName.toUpperCase()}`,
      pn: finalKitPn,
      sn: finalKitSn,
      sourceLocation: 'AT-802 Parça Deposu',
      targetLocation: kitLocation,
      quantity: kitQuantity,
      kitCount: kitQuantity,
      kitComponentsData: [...kitComponents],
      operator: opOperator || 'Yetkili Teknisyen (1839)',
      notes: `${kitQuantity} adet Kit üretildi. Kullanılan parçalar AT-802 stoklarından otomatik düşüldü: ${kitComponents.map(c => `${c.quantityPerKit * kitQuantity}x ${c.itemName}`).join(', ')}`
    });

    setIsKitPreviewOpen(false);
    showNotification(`✅ "${kitName}" başarıyla AT-802 deposuna yeni ürün olarak eklendi ve bileşen stokları düşüldü!`);
    
    // Formu temizle
    setKitName('');
    setKitPn('');
    setKitComponents([]);
    setActiveTab('loglar');
  };

  // Sayımı Tamamlama
  const handleCompleteSayim = () => {
    const updatedRows = [...depoRows];
    let changedCount = 0;

    Object.entries(sayimCounts).forEach(([rIdxStr, countedVal]) => {
      const rIdx = Number(rIdxStr);
      const numVal = Number(countedVal);
      if (updatedRows[rIdx]) {
        const rowCopy = [...updatedRows[rIdx]];
        const currentQty = parseFloat((rowCopy[4] || "0").replace(/[^0-9.]/g, '')) || 0;
        if (currentQty !== numVal) {
          rowCopy[4] = `${numVal} ADET`;
          updatedRows[rIdx] = rowCopy;
          changedCount++;

          onAddTransaction({
            type: numVal > currentQty ? 'GİRİŞ' : 'ÇIKIŞ',
            itemName: rowCopy[1] || "",
            pn: rowCopy[2] || "",
            sn: rowCopy[3] || "",
            sourceLocation: rowCopy[5] || "DEPO",
            targetLocation: "SAYIM DÜZELTME",
            quantity: Math.abs(numVal - currentQty),
            operator: "Yetkili Sayım Heyeti (1839)",
            notes: `Fiziki Sayım Düzeltmesi: Sistem (${currentQty}) ➜ Sayılan (${numVal})`
          });
        }
      }
    });

    onUpdateDepoRows(updatedRows);
    setSayimCounts({});
    showNotification(`Depo sayımı başarıyla tamamlandı! ${changedCount} kalemde stok güncellendi.`);
  };

  // İşlemi Geri Al (Undo)
  const handleUndoAction = (tx: DepoTransaction) => {
    if (tx.isUndone) {
      alert("Bu işlem zaten geri alınmış.");
      return;
    }

    if (!window.confirm(`"${tx.itemName}" için yapılan "${tx.type}" işlemini geri almak ve stokları eski haline getirmek istiyor musunuz?`)) {
      return;
    }

    let updatedRows = [...depoRows];

    if (tx.type === 'KİT OLUŞTURMA' && tx.kitComponentsData) {
      updatedRows = updatedRows.filter(r => !(r[2] === tx.pn && (r[1] || "").includes(tx.itemName)));
      const count = tx.kitCount || tx.quantity || 1;
      updatedRows = updatedRows.map(row => {
        const rowPn = row[2] || row[1] || "";
        const rowName = row[1] || "";
        const matchedComp = tx.kitComponentsData?.find(c => c.pn === rowPn && c.itemName === rowName);
        if (matchedComp) {
          const copy = [...row];
          const curStock = parseFloat((copy[4] || "0").replace(/[^0-9.]/g, '')) || 0;
          const refund = matchedComp.quantityPerKit * count;
          copy[4] = `${curStock + refund} ADET`;
          return copy;
        }
        return row;
      });

      if (onUpdateAt802Rows) onUpdateAt802Rows(updatedRows);
      else onUpdateDepoRows(updatedRows);

      if (onUndoTransaction) onUndoTransaction(tx.id);
      showNotification(`"${tx.itemName}" kiti silindi ve kullanılan bileşen stokları iade edildi!`);
      return;
    }

    if (tx.type === 'ÇIKIŞ') {
      updatedRows = updatedRows.map(row => {
        if ((row[2] === tx.pn && tx.pn !== "-") || row[1] === tx.itemName) {
          const copy = [...row];
          const curStock = parseFloat((copy[4] || "0").replace(/[^0-9.]/g, '')) || 0;
          copy[4] = `${curStock + tx.quantity} ADET`;
          return copy;
        }
        return row;
      });

      onUpdateDepoRows(updatedRows);
      if (onUndoTransaction) onUndoTransaction(tx.id);
      showNotification(`${tx.quantity} adet "${tx.itemName}" stoğu iade edildi!`);
      return;
    }

    if (tx.type === 'GİRİŞ') {
      updatedRows = updatedRows.map(row => {
        if ((row[2] === tx.pn && tx.pn !== "-") || row[1] === tx.itemName) {
          const copy = [...row];
          const curStock = parseFloat((copy[4] || "0").replace(/[^0-9.]/g, '')) || 0;
          copy[4] = `${Math.max(0, curStock - tx.quantity)} ADET`;
          return copy;
        }
        return row;
      });

      onUpdateDepoRows(updatedRows);
      if (onUndoTransaction) onUndoTransaction(tx.id);
      showNotification(`Giriş işlemi geri alındı: ${tx.quantity} adet "${tx.itemName}" stoktan düşüldü.`);
      return;
    }

    if (tx.type === 'TRANSFER') {
      updatedRows = updatedRows.map(row => {
        if ((row[2] === tx.pn && tx.pn !== "-") || row[1] === tx.itemName) {
          const copy = [...row];
          copy[5] = tx.sourceLocation || "DEPO";
          return copy;
        }
        return row;
      });

      onUpdateDepoRows(updatedRows);
      if (onUndoTransaction) onUndoTransaction(tx.id);
      showNotification(`Transfer geri alındı: "${tx.itemName}" lokasyonu "${tx.sourceLocation}" olarak düzeltildi.`);
      return;
    }

    if (onUndoTransaction) {
      onUndoTransaction(tx.id);
    }
  };

  // Export Depo Management Ledger to Excel (9 Sütunlu Kullanıcı Formatı - Genel & Birim Bazlı Sayfalar)
  const exportDepoLedgerToExcel = () => {
    try {
      const formatRow = (t: DepoTransaction) => ({
        "MALZEME ADI": t.itemName || "-",
        "ADET": t.quantity || 1,
        "TARİH": t.date || (t.timestamp ? t.timestamp.split(' ')[0] : getTodayFormattedDate()),
        "İŞLEM TÜRÜ": t.type || "ANKARA ÇIKAN",
        "SERİAL NUMBER": t.sn || "-",
        "KUYRUK KODU": t.tailNo || "-",
        "TESLİM ALAN": t.operator || "-",
        "KABUL YAPAN": t.receivedBy || "-",
        "DEPO YERİ": t.location || t.sourceLocation || "-"
      });

      const wb = XLSX.utils.book_new();

      // 1. Genel Sayfa ("transfer geçmişi")
      const allData = transactions.map(formatRow);
      const wsMain = XLSX.utils.json_to_sheet(allData);
      XLSX.utils.book_append_sheet(wb, wsMain, "transfer geçmişi");

      // 2. Birim Bazlı Ayrı Sayfalar
      const classify = (t: DepoTransaction): string => {
        const full = `${t.itemName || ''} ${t.pn || ''} ${t.sn || ''} ${t.tailNo || ''} ${t.location || ''} ${t.notes || ''}`.toLowerCase();
        if (full.includes('at-802') || full.includes('at802') || full.includes('or-20')) return 'at802';
        if (full.includes('bell') || full.includes('429') || full.includes('tc-zog') || full.includes('tc-zoh') || full.includes('tc-zoi')) return 'bell429';
        if (full.includes('t70') || full.includes('t-70') || full.includes('bumbi') || full.includes('helitak') || full.includes('sikorsky')) return 't70';
        if (full.includes('b360') || full.includes('b-360') || full.includes('king air') || full.includes('tc-ogm')) return 'b360';
        if (full.includes('c650') || full.includes('c-650') || full.includes('citation') || full.includes('tc-cya') || full.includes('tc-cyb')) return 'c650';
        if (full.includes('hangar')) return 'hangar';
        if (full.includes('kara')) return 'kara_araclari';
        return 'at802';
      };

      const unitSheets: { [key: string]: string } = {
        "transfer geçmişi-at802": "at802",
        "transfer geçmişi-bell429": "bell429",
        "transfer geçmişi-t70": "t70",
        "transfer geçmişi-b360": "b360",
        "transfer geçmişi-c650": "c650",
        "transfer geçmişi-hangar": "hangar",
        "transfer geçmişi-kara_araclari": "kara_araclari"
      };

      Object.entries(unitSheets).forEach(([sheetName, uKey]) => {
        const filtered = transactions.filter(t => classify(t) === uKey);
        if (filtered.length > 0) {
          const subWs = XLSX.utils.json_to_sheet(filtered.map(formatRow));
          XLSX.utils.book_append_sheet(wb, subWs, sheetName.slice(0, 31));
        }
      });

      XLSX.writeFile(wb, `Transfer_Gecmisi_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showNotification("Depo Transfer Geçmişi ('transfer geçmişi' ve birim alt sayfaları ile) başarıyla indirildi!");
    } catch (err) {
      alert("Excel aktarılırken hata: " + (err as Error).message);
    }
  };

  return (
    <>
      <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-md animate-fade-in">
        <div className="bg-white rounded-3xl shadow-2xl max-w-5xl w-full flex flex-col max-h-[92vh] overflow-hidden border border-slate-200">
        
        {/* Header */}
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-[#0b3d1d] flex items-center justify-center text-emerald-300 shadow-inner">
              <Boxes className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-black uppercase tracking-wider flex items-center gap-2">
                <span>DEPO YÖNETİM & LOG PANELİ</span>
                <span className="text-[10px] bg-emerald-950 border border-emerald-700 text-emerald-400 px-2 py-0.5 rounded-full font-mono">
                  YETKİLİ: 1839
                </span>
              </h3>
              <p className="text-[11px] text-slate-400 font-mono">
                Depo Giriş-Çıkış, Otomatik Malzeme Seçimi, AT-802 Kit Montajı ve Geri Alınabilir İşlem Logları
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white flex items-center justify-center font-bold text-sm transition-all cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Navigation Tabs */}
        <div className="flex items-center gap-2 px-6 pt-3 pb-2 bg-slate-100 border-b border-slate-200 overflow-x-auto shrink-0">
          <button
            type="button"
            onClick={() => setActiveTab('islemler')}
            className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider flex items-center gap-2 transition-all cursor-pointer ${
              activeTab === 'islemler'
                ? 'bg-[#0b3d1d] text-white shadow-md'
                : 'bg-white text-slate-600 hover:bg-slate-200/70 border border-slate-200'
            }`}
          >
            <ArrowLeftRight className="w-3.5 h-3.5" />
            <span>GİRİŞ / ÇIKIŞ İŞLEMLERİ</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('kit')}
            className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider flex items-center gap-2 transition-all cursor-pointer ${
              activeTab === 'kit'
                ? 'bg-[#0b3d1d] text-white shadow-md'
                : 'bg-white text-slate-600 hover:bg-slate-200/70 border border-slate-200'
            }`}
          >
            <Boxes className="w-3.5 h-3.5" />
            <span>KİT OLUŞTURMA</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('sayim')}
            className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider flex items-center gap-2 transition-all cursor-pointer ${
              activeTab === 'sayim'
                ? 'bg-[#0b3d1d] text-white shadow-md'
                : 'bg-white text-slate-600 hover:bg-slate-200/70 border border-slate-200'
            }`}
          >
            <CheckCircle2 className="w-3.5 h-3.5" />
            <span>LOKASYON BAZLI SAYIM</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('loglar')}
            className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider flex items-center gap-2 transition-all cursor-pointer ${
              activeTab === 'loglar'
                ? 'bg-[#0b3d1d] text-white shadow-md'
                : 'bg-white text-slate-600 hover:bg-slate-200/70 border border-slate-200'
            }`}
          >
            <Clock className="w-3.5 h-3.5" />
            <span>TRANSFER & İŞLEM GEÇMİŞİ ({transactions.length})</span>
          </button>
        </div>

        {/* Tab 1: Giriş / Çıkış İşlemleri */}
        {activeTab === 'islemler' && (
          <div className="flex-1 overflow-y-auto p-6 flex flex-col gap-5">
            
            {/* Üst Bar: Hızlı Excel Şablonu İndir & Toplu Excel Yükle */}
            <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-900 text-white p-3.5 rounded-2xl">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                <span className="text-xs font-bold">Hızlı Toplu İşlem Portalı:</span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={downloadSampleOperationExcel}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-emerald-300 border border-emerald-600/40 text-[11px] font-bold rounded-xl flex items-center gap-1.5 transition-all cursor-pointer active:scale-95"
                  title="Örnek Giriş/Çıkış Excel Şablonu İndir"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>ÖRNEK EXCEL ŞABLONU İNDİR</span>
                </button>

                <label className="px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white text-[11px] font-bold rounded-xl flex items-center gap-1.5 transition-all cursor-pointer active:scale-95 shadow-md">
                  <UploadCloud className="w-3.5 h-3.5" />
                  <span>TOPLU EXCEL İLE İŞLEM YÜKLE</span>
                  <input
                    ref={bulkFileInputRef}
                    type="file"
                    accept=".xlsx, .xls, .csv"
                    onChange={handleBulkExcelUpload}
                    className="hidden"
                  />
                </label>
              </div>
            </div>

            {/* İşlem Türü Seçimi */}
            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={() => handleOperationChange('cikis')}
                className={`flex-1 py-3 rounded-2xl text-xs font-black uppercase tracking-wider flex items-center justify-center gap-2 border transition-all cursor-pointer ${
                  selectedOperation === 'cikis'
                    ? 'bg-rose-600 text-white border-rose-700 shadow-md'
                    : 'bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200'
                }`}
              >
                <ArrowUpRight className="w-4 h-4" />
                <span>DEPO ÇIKIŞ / SARF</span>
              </button>

              <button
                type="button"
                onClick={() => handleOperationChange('giris')}
                className={`flex-1 py-3 rounded-2xl text-xs font-black uppercase tracking-wider flex items-center justify-center gap-2 border transition-all cursor-pointer ${
                  selectedOperation === 'giris'
                    ? 'bg-emerald-700 text-white border-emerald-800 shadow-md'
                    : 'bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200'
                }`}
              >
                <ArrowDownLeft className="w-4 h-4" />
                <span>DEPO GİRİŞ (STOK ARTIR / YENİ ÜRÜN)</span>
              </button>

              <button
                type="button"
                onClick={() => handleOperationChange('transfer')}
                className={`flex-1 py-3 rounded-2xl text-xs font-black uppercase tracking-wider flex items-center justify-center gap-2 border transition-all cursor-pointer ${
                  selectedOperation === 'transfer'
                    ? 'bg-indigo-600 text-white border-indigo-700 shadow-md'
                    : 'bg-slate-50 text-slate-700 hover:bg-slate-100 border-slate-200'
                }`}
              >
                <ArrowLeftRight className="w-4 h-4" />
                <span>DEPO İÇİ / İLLER ARASI TRANSFER</span>
              </button>
            </div>

            {selectedOperation === 'giris' && (
              <div className="flex items-center gap-4 p-3 bg-slate-100 rounded-2xl">
                <span className="text-xs font-bold text-slate-700">Giriş Modu:</span>
                <label className="flex items-center gap-2 text-xs font-bold text-slate-800 cursor-pointer">
                  <input
                    type="radio"
                    name="giris_mode"
                    checked={girisMode === 'mevcut'}
                    onChange={() => setGirisMode('mevcut')}
                    className="accent-emerald-700"
                  />
                  <span>Mevcut Malzeme Stok Artışı</span>
                </label>
                <label className="flex items-center gap-2 text-xs font-bold text-slate-800 cursor-pointer">
                  <input
                    type="radio"
                    name="giris_mode"
                    checked={girisMode === 'yeni'}
                    onChange={() => setGirisMode('yeni')}
                    className="accent-emerald-700"
                  />
                  <span>Yeni Malzeme Kaydı & Girişi</span>
                </label>
              </div>
            )}

            {/* Form Fields */}
            <div className="bg-slate-50/80 p-5 rounded-2xl border border-slate-200 space-y-4">
              {selectedOperation === 'giris' && girisMode === 'yeni' ? (
                <>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        YENİ MALZEME ADI <span className="text-rose-500">*</span>
                      </label>
                      <input
                        type="text"
                        placeholder="Örn: O-Ring Conta Takımı"
                        value={newMaterialName}
                        onChange={(e) => setNewMaterialName(e.target.value)}
                        className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-700"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        PARÇA NO (P/N)
                      </label>
                      <input
                        type="text"
                        placeholder="Örn: 101-380010-1"
                        value={newMaterialPn}
                        onChange={(e) => setNewMaterialPn(e.target.value)}
                        className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-700"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        GİRİŞ MİKTARI
                      </label>
                      <input
                        type="number"
                        min={1}
                        value={opQuantity}
                        onChange={(e) => setOpQuantity(Math.max(1, Number(e.target.value)))}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-700"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        HEDEF LOKASYON / RAF
                      </label>
                      <select
                        value={newMaterialLoc}
                        onChange={(e) => setNewMaterialLoc(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none"
                      >
                        {allAvailableLocations.map((loc, i) => (
                          <option key={i} value={loc}>{loc}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        TEDARİKÇİ / FİRMA
                      </label>
                      <input
                        type="text"
                        placeholder="TÜBİTAK UME vb."
                        value={newMaterialFirma}
                        onChange={(e) => setNewMaterialFirma(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        TARİH
                      </label>
                      <input
                        type="text"
                        placeholder="GG.AA.YYYY"
                        value={opDate}
                        onChange={(e) => setOpDate(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center justify-between">
                        <span>İŞLEM TÜRÜ</span>
                        {opTransactionType === 'DİĞER' && (
                          <span className="text-[9px] text-amber-800 bg-amber-100 font-bold px-1.5 py-0.5 rounded">Manuel</span>
                        )}
                        {!opTransactionType && (
                          <span className="text-[9px] text-rose-600 bg-rose-50 font-bold px-1.5 py-0.5 rounded">Seçilmedi</span>
                        )}
                      </label>
                      <select
                        value={opTransactionType}
                        onChange={(e) => setOpTransactionType(e.target.value)}
                        className={`w-full px-3.5 py-2 bg-white border-2 ${!opTransactionType ? 'border-amber-400 bg-amber-50/20' : 'border-slate-300'} rounded-xl font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]`}
                      >
                        <option value="">-- İŞLEM TÜRÜ SEÇİNİZ --</option>
                        {GIRIS_TRANSACTION_TYPES.map((t) => (
                          <option key={t} value={t}>
                            {t}
                          </option>
                        ))}
                      </select>
                      {opTransactionType === 'DİĞER' && (
                        <div className="mt-1.5 animate-fadeIn">
                          <input
                            type="text"
                            placeholder="Manuel işlem türü yazınız..."
                            value={opCustomTransactionType}
                            onChange={(e) => setOpCustomTransactionType(e.target.value)}
                            className="w-full px-3 py-1.5 bg-amber-50 border-2 border-amber-400 rounded-xl font-bold text-slate-900 placeholder:text-amber-800/50 text-xs focus:outline-none focus:border-amber-600 focus:bg-white shadow-inner"
                            autoFocus
                          />
                        </div>
                      )}
                    </div>
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        SERİAL NUMBER (S/N)
                      </label>
                      <input
                        type="text"
                        placeholder="Seri No (varsa)"
                        value={newMaterialSn}
                        onChange={(e) => setNewMaterialSn(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        KABUL YAPAN
                      </label>
                      <input
                        type="text"
                        placeholder="Personel Adı"
                        value={opAcceptedBy}
                        onChange={(e) => setOpAcceptedBy(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none"
                      />
                    </div>
                  </div>
                </>
              ) : (
                <>
                  {/* Yazdıkça Gelen Malzeme Seçimi (Autocomplete / Dynamic Search) */}
                  <div className="relative">
                    <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                      İŞLEM YAPILACAK MALZEME (YAZDIKÇA GELEN OTOMATİK DOLDURMA) <span className="text-rose-500">*</span>
                    </label>
                    
                    <div className="relative">
                      <Search className="w-4 h-4 absolute left-3.5 top-3 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Malzeme Adı veya Parça No (P/N) yazarak arayınız..."
                        value={materialSearchQuery}
                        onChange={(e) => {
                          setMaterialSearchQuery(e.target.value);
                          setIsMaterialDropdownOpen(true);
                        }}
                        onFocus={() => setIsMaterialDropdownOpen(true)}
                        className="w-full pl-10 pr-10 py-2.5 bg-white border-2 border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d] shadow-sm"
                      />
                      {materialSearchQuery && (
                        <button
                          type="button"
                          onClick={() => {
                            setMaterialSearchQuery('');
                            setSelectedItemIndex(-1);
                            setIsMaterialDropdownOpen(true);
                          }}
                          className="absolute right-3 top-3 text-slate-400 hover:text-slate-600"
                        >
                          <X className="w-4 h-4" />
                        </button>
                      )}
                    </div>

                    {/* Malzeme Öneri Açılır Listesi */}
                    {isMaterialDropdownOpen && filteredMaterials.length > 0 && (
                      <div className="absolute top-full left-0 right-0 mt-1.5 bg-white border-2 border-emerald-700/50 rounded-2xl shadow-2xl max-h-56 overflow-y-auto divide-y divide-slate-100 z-50">
                        {filteredMaterials.map(({ row, idx }) => (
                          <div
                            key={idx}
                            onClick={() => handleSelectMaterial(idx)}
                            className="p-3 hover:bg-emerald-50 cursor-pointer flex items-center justify-between transition-colors"
                          >
                            <div>
                              <div className="font-bold text-xs text-slate-900">{row[1]}</div>
                              <div className="text-[10px] font-mono text-slate-500 mt-0.5">
                                <span className="font-semibold text-emerald-900">P/N: {row[2] || '-'}</span> • Seri No: {row[3] || '-'} • Raf: {row[5] || 'DEPO'}
                              </div>
                            </div>
                            <div className="text-right">
                              <span className="px-2.5 py-1 bg-emerald-100 text-emerald-900 font-mono font-black text-xs rounded-lg">
                                Stok: {row[4] || '0'}
                              </span>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}

                    {/* Seçilen Malzemenin Bilgi Kartı */}
                    {currentSelectedItem && (
                      <div className="mt-2.5 p-3 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <CheckCircle2 className="w-4 h-4 text-emerald-700" />
                          <span className="text-xs font-black text-emerald-950">Seçili: {currentSelectedItem[1]}</span>
                          <span className="text-[11px] font-mono text-emerald-800">(P/N: {currentSelectedItem[2] || '-'})</span>
                        </div>
                        <div className="text-xs font-mono font-black text-emerald-900">
                          Mevcut Stok: {currentSelectedItem[4] || '0'} | Raf: {currentSelectedItem[5] || 'DEPO'}
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        ADET (MİKTAR)
                      </label>
                      <input
                        type="number"
                        min={1}
                        value={opQuantity}
                        onChange={(e) => setOpQuantity(Math.max(1, Number(e.target.value)))}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        TARİH
                      </label>
                      <input
                        type="text"
                        placeholder="GG.AA.YYYY"
                        value={opDate}
                        onChange={(e) => setOpDate(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center justify-between">
                        <span>İŞLEM TÜRÜ</span>
                        {opTransactionType === 'DİĞER' && (
                          <span className="text-[9px] text-amber-800 bg-amber-100 font-bold px-1.5 py-0.5 rounded">Manuel</span>
                        )}
                        {!opTransactionType && (
                          <span className="text-[9px] text-rose-600 bg-rose-50 font-bold px-1.5 py-0.5 rounded">Seçilmedi</span>
                        )}
                      </label>
                      <select
                        value={opTransactionType}
                        onChange={(e) => {
                          const val = e.target.value;
                          setOpTransactionType(val);
                          if (val.includes('ÇIKAN') || val.includes('ÇIKIŞ')) {
                            setSelectedOperation('cikis');
                          } else if (val.includes('GİREN') || val.includes('GİRİŞ') || val === 'GELEN') {
                            setSelectedOperation('giris');
                          } else if (val.includes('TRANSFER')) {
                            setSelectedOperation('transfer');
                          }
                        }}
                        className={`w-full px-3.5 py-2 bg-white border-2 ${!opTransactionType ? 'border-amber-400 bg-amber-50/20' : 'border-slate-300'} rounded-xl font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]`}
                      >
                        <option value="">-- İŞLEM TÜRÜ SEÇİNİZ --</option>
                        {availableTransactionTypes.map((t) => (
                          <option key={t} value={t}>
                            {t}
                          </option>
                        ))}
                      </select>

                      {opTransactionType === 'DİĞER' && (
                        <div className="mt-1.5 animate-fadeIn">
                          <input
                            type="text"
                            placeholder="Manuel işlem türü yazınız (Örn: ESKİŞEHİR TRANSFER, KİT vb.)..."
                            value={opCustomTransactionType}
                            onChange={(e) => setOpCustomTransactionType(e.target.value)}
                            className="w-full px-3 py-1.5 bg-amber-50 border-2 border-amber-400 rounded-xl font-bold text-slate-900 placeholder:text-amber-800/50 text-xs focus:outline-none focus:border-amber-600 focus:bg-white shadow-inner"
                            autoFocus
                          />
                        </div>
                      )}
                    </div>

                    <div>
                      {matchingStockRows.length > 1 ? (
                        <div>
                          <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center justify-between">
                            <span>SERİAL NUMBER (S/N)</span>
                            <span className="text-[9px] text-indigo-800 bg-indigo-100 font-black px-1.5 py-0.5 rounded">
                              {matchingStockRows.length} Stok / S-N Mevcut
                            </span>
                          </label>
                          <select
                            value={selectedItemIndex >= 0 ? selectedItemIndex : ''}
                            onChange={(e) => {
                              const chosenIdx = Number(e.target.value);
                              if (!isNaN(chosenIdx) && chosenIdx >= 0 && chosenIdx < depoRows.length) {
                                handleSelectMaterial(chosenIdx);
                              }
                            }}
                            className="w-full px-2.5 py-1.5 bg-indigo-50/80 border-2 border-indigo-400 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-indigo-600 focus:bg-white text-xs mb-1.5 shadow-sm"
                          >
                            <option value="">-- Seri No & Raf Seçiniz ({matchingStockRows.length} Adet) --</option>
                            {matchingStockRows.map(({ row, idx }) => {
                              const sn = (row[3] || '').trim() || '-';
                              const loc = (row[5] || '').trim() || 'DEPO';
                              const qty = (row[4] || '').trim() || '1';
                              return (
                                <option key={idx} value={idx}>
                                  {sn !== '-' ? `S/N: ${sn}` : 'Seri No Yok (-)'} | Raf: {loc} | Stok: {qty}
                                </option>
                              );
                            })}
                          </select>
                          <input
                            type="text"
                            placeholder="S/N manuel doğrula / düzenle..."
                            value={opSerialNo}
                            onChange={(e) => setOpSerialNo(e.target.value)}
                            className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-xl font-mono text-slate-900 focus:outline-none focus:border-[#0b3d1d] text-xs"
                          />
                        </div>
                      ) : (
                        <div>
                          <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                            SERİAL NUMBER
                          </label>
                          <input
                            type="text"
                            placeholder="-"
                            value={opSerialNo}
                            onChange={(e) => setOpSerialNo(e.target.value)}
                            className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-mono text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                          />
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center gap-1">
                        <Plane className="w-3.5 h-3.5 text-slate-500" />
                        <span>KUYRUK KODU</span>
                      </label>
                      <select
                        value={opTailNo}
                        onChange={(e) => setOpTailNo(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                      >
                        <option value="">- SEÇİNİZ VEYA BOŞ BIRAKINIZ -</option>
                        <option value="-">- (HAVA ARACI DIŞI / DEPO)</option>
                        {AIRCRAFT_TAIL_LIST.map((ac, i) => (
                          <option key={i} value={ac.tailNo}>
                            {ac.tailNo} ({ac.callsign} - {ac.unit})
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* TESLİM ALAN (Personel Listesinden Canlı Daralan Otomatik Tamamlama) */}
                    <div className="relative">
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center justify-between">
                        <span>TESLİM ALAN</span>
                        {isPersonnelLoading && (
                          <span className="text-[9px] text-emerald-600 animate-pulse font-normal flex items-center gap-1">
                            <RefreshCw className="w-2.5 h-2.5 animate-spin" /> Yükleniyor...
                          </span>
                        )}
                      </label>
                      <input
                        type="text"
                        placeholder="Personel Adı Yazınız..."
                        value={opDeliveredTo}
                        onFocus={() => setIsDeliveredDropdownOpen(true)}
                        onChange={(e) => {
                          setOpDeliveredTo(e.target.value);
                          setIsDeliveredDropdownOpen(true);
                        }}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                      />
                      {isDeliveredDropdownOpen && filteredDeliveredPersonnel.length > 0 && (
                        <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-300 rounded-xl shadow-2xl z-50 max-h-48 overflow-y-auto">
                          {filteredDeliveredPersonnel.map((pName, pIdx) => (
                            <button
                              key={pIdx}
                              type="button"
                              onClick={() => {
                                setOpDeliveredTo(pName);
                                setIsDeliveredDropdownOpen(false);
                              }}
                              className="w-full text-left px-3 py-2 text-xs font-bold text-slate-800 hover:bg-emerald-50 hover:text-emerald-950 border-b border-slate-100 last:border-b-0 flex items-center justify-between transition-colors"
                            >
                              <span>{pName}</span>
                              <UserCheck className="w-3.5 h-3.5 text-emerald-600 opacity-60" />
                            </button>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* KABUL YAPAN (Personel Listesinden Canlı Daralan Otomatik Tamamlama) */}
                    <div className="relative">
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center justify-between">
                        <span>KABUL YAPAN</span>
                        {isPersonnelLoading && (
                          <span className="text-[9px] text-emerald-600 animate-pulse font-normal flex items-center gap-1">
                            <RefreshCw className="w-2.5 h-2.5 animate-spin" /> Yükleniyor...
                          </span>
                        )}
                      </label>
                      <input
                        type="text"
                        placeholder="Personel Adı Yazınız..."
                        value={opAcceptedBy}
                        onFocus={() => setIsAcceptedDropdownOpen(true)}
                        onChange={(e) => {
                          setOpAcceptedBy(e.target.value);
                          setIsAcceptedDropdownOpen(true);
                        }}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                      />
                      {isAcceptedDropdownOpen && filteredAcceptedPersonnel.length > 0 && (
                        <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-300 rounded-xl shadow-2xl z-50 max-h-48 overflow-y-auto">
                          {filteredAcceptedPersonnel.map((pName, pIdx) => (
                            <button
                              key={pIdx}
                              type="button"
                              onClick={() => {
                                setOpAcceptedBy(pName);
                                setIsAcceptedDropdownOpen(false);
                              }}
                              className="w-full text-left px-3 py-2 text-xs font-bold text-slate-800 hover:bg-emerald-50 hover:text-emerald-950 border-b border-slate-100 last:border-b-0 flex items-center justify-between transition-colors"
                            >
                              <span>{pName}</span>
                              <UserCheck className="w-3.5 h-3.5 text-emerald-600 opacity-60" />
                            </button>
                          ))}
                        </div>
                      )}
                    </div>

                    <div>
                      <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                        DEPO YERİ / RAF
                      </label>
                      <input
                        type="text"
                        placeholder="Örn: M-70 (8BATCH PALET1 BOX2)"
                        value={opDepoLocation}
                        onChange={(e) => setOpDepoLocation(e.target.value)}
                        className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-medium text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                      EK AÇIKLAMA / İŞ EMRİ NO (OPSİYONEL)
                    </label>
                    <input
                      type="text"
                      placeholder="WO-2026-001 veya rutin sarf açıklaması..."
                      value={opNotes}
                      onChange={(e) => setOpNotes(e.target.value)}
                      className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-xl font-medium text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                    />
                  </div>
                </>
              )}
            </div>

            <div className="flex justify-end pt-3 border-t border-slate-200">
              <button
                type="button"
                onClick={handleOpenPreConfirm}
                className="px-6 py-3 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-all active:scale-95 cursor-pointer flex items-center gap-2"
              >
                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                <span>İŞLEMİ KAYDET VE ÖN ONAYA GÖNDER</span>
              </button>
            </div>
          </div>
        )}

        {/* Tab 2: Kit Oluşturma */}
        {activeTab === 'kit' && (
          <div className="flex-1 overflow-y-auto p-6 flex flex-col gap-5">
            <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-2xl flex items-center justify-between">
              <div>
                <h4 className="text-xs font-black text-emerald-950 uppercase tracking-wider flex items-center gap-2">
                  <Boxes className="w-4 h-4 text-emerald-700" />
                  <span>KİT OLUŞTURMA & BİLEŞEN MONTAJI</span>
                </h4>
                <p className="text-[11px] text-emerald-800 font-medium">
                  Kite eklenecek parçaları ve adetlerini belirleyiniz. "KİTİ OLUŞTUR" butonuna basıldığında detaylı önizleme ve Excel indirme ekranı açılır.
                </p>
              </div>
              <span className="px-3 py-1 bg-emerald-900 text-white font-mono text-[10px] font-black rounded-lg">
                KİT MONTAJ
              </span>
            </div>

            {/* Kit Genel Bilgileri */}
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 bg-slate-50 p-4 rounded-2xl border border-slate-200">
              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                  KİT ADI <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  placeholder="Örn: 100 Saatlik Bakım Kiti"
                  value={kitName}
                  onChange={(e) => setKitName(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                  KİT PARÇA NO (P/N)
                </label>
                <input
                  type="text"
                  placeholder="KIT-AT802-100H"
                  value={kitPn}
                  onChange={(e) => setKitPn(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-mono font-bold text-slate-900 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                  ÜRETİLECEK KİT ADEDİ
                </label>
                <input
                  type="number"
                  min={1}
                  value={kitQuantity}
                  onChange={(e) => setKitQuantity(Math.max(1, Number(e.target.value)))}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-mono font-black text-emerald-800 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                  HEDEF LOKASYON / RAF
                </label>
                <select
                  value={kitLocation}
                  onChange={(e) => setKitLocation(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none"
                >
                  {allAvailableLocations.map((loc, i) => (
                    <option key={i} value={loc}>{loc}</option>
                  ))}
                </select>
              </div>
            </div>

            {/* Parça Ekleme & Otomatik Doldurma (Sadece AT-802 Depo Ürünleri) */}
            <div className="bg-slate-50/80 p-4 rounded-2xl border border-slate-200 space-y-3">
              <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center justify-between">
                <span>AT-802 DEPOSUNDAN KİTE PARÇA EKLE</span>
                <span className="text-[10px] text-slate-500 font-mono">Yazdıkça otomatik listelenir</span>
              </h4>

              <div className="grid grid-cols-1 sm:grid-cols-12 gap-3 items-end">
                <div className="sm:col-span-7 relative">
                  <label className="block text-[10px] font-black text-slate-600 uppercase tracking-wider mb-1">
                    AT-802 PARÇA ARAMA & SEÇİMİ
                  </label>
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
                    <input
                      type="text"
                      placeholder="AT-802 parça adı veya P/N yazınız..."
                      value={compSearch}
                      onChange={(e) => {
                        setCompSearch(e.target.value);
                        setIsCompDropdownOpen(true);
                      }}
                      onFocus={() => setIsCompDropdownOpen(true)}
                      className="w-full pl-8 pr-8 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-emerald-700"
                    />
                    {compSearch && (
                      <button
                        type="button"
                        onClick={() => {
                          setCompSearch('');
                          setSelectedCompRow(null);
                        }}
                        className="absolute right-2.5 top-2.5 text-slate-400 hover:text-slate-600"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>

                  {/* Dropdown Öneriler */}
                  {isCompDropdownOpen && filteredAt802Parts.length > 0 && (
                    <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-emerald-700/40 rounded-xl shadow-xl max-h-48 overflow-y-auto divide-y divide-slate-100 z-50">
                      {filteredAt802Parts.map((row, i) => (
                        <div
                          key={i}
                          onClick={() => {
                            setSelectedCompRow(row);
                            setCompSearch(`${row[1]} [P/N: ${row[2] || '-'}]`);
                            setIsCompDropdownOpen(false);
                          }}
                          className="p-2.5 hover:bg-emerald-50 cursor-pointer flex items-center justify-between text-xs"
                        >
                          <div>
                            <div className="font-bold text-slate-900">{row[1]}</div>
                            <div className="text-[10px] font-mono text-slate-500">P/N: {row[2] || '-'} • Raf: {row[5] || 'DEPO'}</div>
                          </div>
                          <span className="px-2 py-0.5 bg-emerald-100 text-emerald-900 font-mono font-bold text-[10px] rounded">
                            Stok: {row[4] || '0'}
                          </span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                <div className="sm:col-span-3">
                  <label className="block text-[10px] font-black text-slate-600 uppercase tracking-wider mb-1">
                    1 KİT İÇİNDEKİ ADEDİ
                  </label>
                  <input
                    type="number"
                    min={1}
                    value={compQtyPerKit}
                    onChange={(e) => setCompQtyPerKit(Math.max(1, Number(e.target.value)))}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-mono font-black text-slate-900 text-center"
                  />
                </div>

                <div className="sm:col-span-2">
                  <button
                    type="button"
                    onClick={() => {
                      if (!selectedCompRow) {
                        alert("Lütfen arama kutusundan bir AT-802 parçası seçiniz.");
                        return;
                      }
                      handleAddSelectedKitComp(selectedCompRow, compQtyPerKit);
                    }}
                    className="w-full py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-black text-xs rounded-xl shadow-sm transition-all active:scale-95 cursor-pointer flex items-center justify-center gap-1"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>PARÇA EKLE</span>
                  </button>
                </div>
              </div>
            </div>

            {/* Kit İçeriğindeki Parçalar Tablosu */}
            <div className="flex flex-col gap-2">
              <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider">
                KİT İÇERİĞİNDEKİ PARÇALAR ({kitComponents.length})
              </h4>

              {kitComponents.length === 0 ? (
                <div className="p-8 border-2 border-dashed border-slate-200 rounded-2xl text-center text-slate-400 text-xs">
                  Henüz kite parça eklenmedi. Yukarıdaki AT-802 arama alanından parça seçip "PARÇA EKLE" butonuna basınız.
                </div>
              ) : (
                <div className="overflow-x-auto border border-slate-200 rounded-2xl shadow-sm">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-900 text-slate-200 uppercase tracking-wider text-[10px] font-black">
                      <tr>
                        <th className="px-4 py-2.5">PARÇA ADI</th>
                        <th className="px-4 py-2.5">P/N</th>
                        <th className="px-4 py-2.5 text-center">MEVCUT STOK</th>
                        <th className="px-4 py-2.5 text-center">1 KİT İÇİN ADET</th>
                        <th className="px-4 py-2.5 text-center">TOPLAM DÜŞÜLECEK</th>
                        <th className="px-4 py-2.5 text-center">İŞLEM</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {kitComponents.map((comp, idx) => {
                        const totalNeeded = comp.quantityPerKit * kitQuantity;
                        const isStockInsufficient = totalNeeded > comp.currentStock;

                        return (
                          <tr key={idx} className="hover:bg-slate-50">
                            <td className="px-4 py-2.5 font-bold text-slate-900">{comp.itemName}</td>
                            <td className="px-4 py-2.5 font-mono text-slate-600">{comp.pn}</td>
                            <td className="px-4 py-2.5 text-center font-mono font-bold text-slate-800">
                              {comp.currentStock}
                            </td>
                            <td className="px-4 py-2.5 text-center">
                              <input
                                type="number"
                                min={1}
                                value={comp.quantityPerKit}
                                onChange={(e) => {
                                  const val = Math.max(1, Number(e.target.value));
                                  const updated = [...kitComponents];
                                  updated[idx].quantityPerKit = val;
                                  setKitComponents(updated);
                                }}
                                className="w-16 px-2 py-1 text-center font-mono font-black bg-white border border-slate-300 rounded-lg text-xs"
                              />
                            </td>
                            <td className="px-4 py-2.5 text-center font-mono font-black">
                              <span className={isStockInsufficient ? "text-rose-600 bg-rose-50 px-2 py-0.5 rounded-full" : "text-emerald-700"}>
                                {totalNeeded} Adet {isStockInsufficient && "⚠️ Yetersiz!"}
                              </span>
                            </td>
                            <td className="px-4 py-2.5 text-center">
                              <button
                                type="button"
                                onClick={() => {
                                  setKitComponents(kitComponents.filter((_, i) => i !== idx));
                                }}
                                className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg cursor-pointer"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* KİTİ OLUŞTUR BUTONU (Sadece "KİTİ OLUŞTUR" adı) */}
            <div className="flex justify-end pt-3 border-t border-slate-200">
              <button
                type="button"
                onClick={handleOpenKitPreview}
                className="px-8 py-3.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-sm uppercase tracking-wider rounded-2xl shadow-xl transition-all active:scale-95 cursor-pointer flex items-center gap-2"
              >
                <Boxes className="w-5 h-5 text-emerald-400" />
                <span>KİTİ OLUŞTUR</span>
              </button>
            </div>
          </div>
        )}

        {/* Tab 3: Lokasyon ve Malzeme Türü Bazlı Sayım (2 Ana Depo) */}
        {activeTab === 'sayim' && (
          <div className="flex-1 overflow-y-auto p-6 flex flex-col gap-4">
            {/* 2 Ana Depo Seçim Kartları (Sarf ve Parça Depo & Kimyasal Depo) */}
            <div className="bg-slate-50/80 p-4 rounded-3xl border border-slate-200 flex flex-col gap-3">
              <div className="flex items-center justify-between flex-wrap gap-2">
                <span className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                  <Warehouse className="w-4 h-4 text-emerald-700" />
                  <span>DEPO SEÇİMİ VE LOKASYON BAZLI SAYIM:</span>
                </span>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setSayimCategoryFilter('TÜMÜ')}
                    className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                      sayimCategoryFilter === 'TÜMÜ'
                        ? 'bg-slate-900 text-white shadow-sm ring-2 ring-slate-700'
                        : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-300'
                    }`}
                  >
                    TÜM MALZEMELER ({depoRows.length})
                  </button>
                </div>
              </div>
              
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {/* 1. Sarf ve Parça Depo Kartı */}
                <div
                  onClick={() => setSayimCategoryFilter('SARF VE PARÇA DEPO')}
                  className={`border-2 rounded-3xl p-5 bg-white transition-all cursor-pointer flex items-center gap-4 relative overflow-hidden ${
                    sayimCategoryFilter === 'SARF VE PARÇA DEPO'
                      ? 'border-blue-600 ring-4 ring-blue-500/20 shadow-lg bg-blue-50/40'
                      : 'border-blue-200 hover:border-blue-400 hover:shadow-md'
                  }`}
                >
                  <div className="w-14 h-14 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center shrink-0 border border-blue-100">
                    <Box className="w-7 h-7" />
                  </div>
                  <div className="flex-1">
                    <div className="flex items-center justify-between gap-2">
                      <h4 className="text-sm font-black text-slate-900 tracking-wider uppercase">
                        SARF VE PARÇA DEPO
                      </h4>
                      <span className="px-2.5 py-0.5 bg-blue-100 text-blue-900 font-mono font-black text-[11px] rounded-full">
                        {depoRows.filter(r => detectItemCategory(r) === 'SARF VE PARÇA DEPO').length} Kalem
                      </span>
                    </div>
                    <p className="text-[11px] text-blue-700 font-semibold tracking-wide uppercase mt-0.5">
                      SARF MALZEMELERİ, YEDEK PARÇA VE MEKANİK BİLEŞENLER
                    </p>
                  </div>
                </div>

                {/* 2. Kimyasal Depo Kartı */}
                <div
                  onClick={() => setSayimCategoryFilter('KİMYASAL DEPO')}
                  className={`border-2 rounded-3xl p-5 bg-white transition-all cursor-pointer flex items-center gap-4 relative overflow-hidden ${
                    sayimCategoryFilter === 'KİMYASAL DEPO'
                      ? 'border-purple-600 ring-4 ring-purple-500/20 shadow-lg bg-purple-50/40'
                      : 'border-purple-200 hover:border-purple-400 hover:shadow-md'
                  }`}
                >
                  <div className="w-14 h-14 rounded-2xl bg-purple-50 text-purple-600 flex items-center justify-center shrink-0 border border-purple-100">
                    <FlaskConical className="w-7 h-7" />
                  </div>
                  <div className="flex-1">
                    <div className="flex items-center justify-between gap-2">
                      <h4 className="text-sm font-black text-slate-900 tracking-wider uppercase">
                        KİMYASAL DEPO
                      </h4>
                      <span className="px-2.5 py-0.5 bg-purple-100 text-purple-900 font-mono font-black text-[11px] rounded-full">
                        {depoRows.filter(r => detectItemCategory(r) === 'KİMYASAL DEPO').length} Kalem
                      </span>
                    </div>
                    <p className="text-[11px] text-purple-700 font-semibold tracking-wide uppercase mt-0.5">
                      KİMYASAL MADDELER, HAVACILIK YAĞLARI, SIVILAR, MASTİKLER VE BOYALAR
                    </p>
                  </div>
                </div>
              </div>
            </div>

            {/* Lokasyon ve Metin Filtresi */}
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 bg-slate-50 p-4 rounded-2xl border border-slate-200">
              <div className="flex items-center gap-3">
                <span className="text-xs font-black text-slate-700 uppercase tracking-wider">
                  SAYILACAK LOKASYON / RAF:
                </span>
                <select
                  value={sayimLocationFilter}
                  onChange={(e) => setSayimLocationFilter(e.target.value)}
                  className="px-3 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none"
                >
                  <option value="TÜMÜ">TÜM LOKASYONLAR</option>
                  {allAvailableLocations.map((loc, i) => (
                    <option key={i} value={loc}>{loc}</option>
                  ))}
                </select>
              </div>

              <div className="relative w-64">
                <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="Malzeme adı veya P/N filtrele..."
                  value={sayimSearch}
                  onChange={(e) => setSayimSearch(e.target.value)}
                  className="w-full pl-8 pr-3 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-medium focus:outline-none"
                />
              </div>
            </div>

            <div className="overflow-x-auto border border-slate-200 rounded-2xl shadow-sm">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-900 text-slate-200 uppercase tracking-wider text-[10px] font-black">
                  <tr>
                    <th className="px-3 py-2.5 text-center">SIRA</th>
                    <th className="px-4 py-2.5">MALZEME ADI</th>
                    <th className="px-3 py-2.5">P/N</th>
                    <th className="px-3 py-2.5 text-center">MALZEME TÜRÜ</th>
                    <th className="px-3 py-2.5">LOKASYON</th>
                    <th className="px-3 py-2.5 text-center">SİSTEMDEKİ STOK</th>
                    <th className="px-3 py-2.5 text-center">SAYILAN ADET</th>
                    <th className="px-3 py-2.5 text-center">FARK</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {depoRows
                    .map((r, originalIdx) => ({ row: r, originalIdx }))
                    .filter(({ row }) => {
                      // Malzeme türü filtresi
                      if (sayimCategoryFilter !== 'TÜMÜ') {
                        const cat = detectItemCategory(row);
                        if (cat !== sayimCategoryFilter) return false;
                      }
                      // Lokasyon filtresi
                      if (sayimLocationFilter !== 'TÜMÜ' && (row[5] || '').trim() !== sayimLocationFilter) return false;
                      // Metin arama filtresi
                      if (sayimSearch.trim()) {
                        const q = sayimSearch.toLowerCase();
                        return (row[1] || '').toLowerCase().includes(q) || (row[2] || '').toLowerCase().includes(q);
                      }
                      return true;
                    })
                    .map(({ row, originalIdx }, displayIdx) => {
                      const sysQty = parseFloat((row[4] || "0").replace(/[^0-9.]/g, '')) || 0;
                      const countedQty = sayimCounts[originalIdx] !== undefined ? sayimCounts[originalIdx] : sysQty;
                      const diff = countedQty - sysQty;
                      const cat = detectItemCategory(row);

                      return (
                        <tr key={originalIdx} className="hover:bg-slate-50">
                          <td className="px-3 py-2 text-center font-mono font-bold text-slate-500">{displayIdx + 1}</td>
                          <td className="px-4 py-2 font-bold text-slate-900">{row[1]}</td>
                          <td className="px-3 py-2 font-mono text-slate-600">{row[2] || '-'}</td>
                          <td className="px-3 py-2 text-center">
                            <span className={`px-2 py-0.5 rounded-full font-black text-[10px] ${
                              cat === 'KİMYASAL DEPO' 
                                ? 'bg-purple-100 text-purple-900 border border-purple-300' 
                                : 'bg-blue-100 text-blue-900 border border-blue-300'
                            }`}>
                              {cat}
                            </span>
                          </td>
                          <td className="px-3 py-2 font-bold text-emerald-800">{row[5] || 'DEPO'}</td>
                          <td className="px-3 py-2 text-center font-mono font-bold text-slate-800">{sysQty}</td>
                          <td className="px-3 py-2 text-center">
                            <input
                              type="number"
                              min={0}
                              value={countedQty}
                              onChange={(e) => {
                                setSayimCounts({
                                  ...sayimCounts,
                                  [originalIdx]: Number(e.target.value)
                                });
                              }}
                              className="w-20 px-2 py-1 bg-white border border-slate-300 rounded-lg text-center font-mono font-bold text-slate-900"
                            />
                          </td>
                          <td className="px-3 py-2 text-center font-mono font-black">
                            {diff === 0 ? (
                              <span className="text-slate-400">0</span>
                            ) : diff > 0 ? (
                              <span className="text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full">+{diff}</span>
                            ) : (
                              <span className="text-rose-600 bg-rose-50 px-2 py-0.5 rounded-full">{diff}</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>

            <div className="flex justify-end pt-3 border-t border-slate-200">
              <button
                type="button"
                onClick={handleCompleteSayim}
                className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all active:scale-95 cursor-pointer flex items-center gap-2"
              >
                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                <span>SAYIMI TAMAMLA VE FİZİKİ STOKLARI EŞİTLE</span>
              </button>
            </div>
          </div>
        )}

        {/* Tab 4: Transfer Geçmişi ve Yönetim Logları */}
        {activeTab === 'loglar' && (
          <div className="flex-1 overflow-y-auto p-6 flex flex-col gap-4">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
              <div>
                <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                  <Clock className="w-4 h-4 text-slate-600" />
                  <span>TRANSFER VE İŞLEM GEÇMİŞİ TABLOSU</span>
                </h4>
                <p className="text-[11px] text-slate-500">
                  Bu yönetim panelinde yapılan tüm giriş, çıkış, kit montajı ve transfer hareketleri satır satır kayıt altındadır. Düzenleme için <strong>satıra çift tıklayabilir</strong>, silebilir veya E-Tablo'ya aktarabilirsiniz.
                </p>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                {selectedTxIds.length > 0 && (
                  <div className="flex items-center gap-2 bg-rose-50 border border-rose-200 px-3 py-1.5 rounded-xl animate-fadeIn">
                    <span className="text-xs font-black text-rose-800">
                      {selectedTxIds.length} Seçili
                    </span>
                    <button
                      type="button"
                      onClick={handleDeleteSelectedTxs}
                      className="px-3 py-1 bg-rose-600 hover:bg-rose-700 text-white text-xs font-black rounded-lg flex items-center gap-1.5 shadow-sm transition-all active:scale-95 cursor-pointer"
                      title="Seçili kayıtları kalıcı olarak sil"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Seçilenleri Toplu Sil ({selectedTxIds.length})</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setSelectedTxIds([])}
                      className="px-2 py-1 bg-white hover:bg-rose-100 text-rose-700 border border-rose-300 text-[11px] font-bold rounded-lg transition-all cursor-pointer"
                    >
                      Seçimi Temizle
                    </button>
                  </div>
                )}
                <button
                  type="button"
                  disabled={isPushingTransfers || transactions.length === 0}
                  onClick={() => uploadAt802TransfersExcelToDrive()}
                  className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold rounded-xl flex items-center gap-2 shadow-sm transition-all active:scale-95 cursor-pointer disabled:opacity-50"
                  title="AT-802 Transfer geçmişini Excel olarak Drive klasörüne ve E-Tabloya aktar"
                >
                  <UploadCloud className={`w-4 h-4 ${isPushingTransfers ? 'animate-bounce' : ''}`} />
                  <span>{isPushingTransfers ? 'Drive & E-Tabloya Yükleniyor...' : 'AT-802 Transfer Geçmişini Drive\'a (Excel) Aktar'}</span>
                </button>
                <button
                  type="button"
                  disabled={isSyncingTransfers}
                  onClick={syncTransfersWithGoogleSheets}
                  className="px-4 py-2 bg-[#0b3d1d] hover:bg-[#072612] text-white text-xs font-bold rounded-xl flex items-center gap-2 shadow-sm transition-all active:scale-95 cursor-pointer disabled:opacity-50"
                  title="E-Tablo transfer geçmişini kontrol et ve yeni transferleri içeri aktar"
                >
                  <RefreshCw className={`w-4 h-4 text-emerald-400 ${isSyncingTransfers ? 'animate-spin' : ''}`} />
                  <span>{isSyncingTransfers ? 'Eşitleniyor...' : 'E-Tablo ile Eşitle'}</span>
                </button>
                <button
                  type="button"
                  onClick={exportDepoLedgerToExcel}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold rounded-xl flex items-center gap-2 shadow-sm transition-all active:scale-95 cursor-pointer"
                >
                  <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                  <span>Excel İndir</span>
                </button>
              </div>
            </div>

            {/* DÜŞÜLEMEYEN / EŞLEŞMEYEN KAYITLAR UYARI BANNERI */}
            {totalUnmatchedCount > 0 && (
              <div className="bg-rose-50 border-2 border-rose-400 p-3.5 rounded-2xl flex items-center justify-between gap-3 shadow-sm animate-pulse">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-xl bg-rose-600 text-white flex items-center justify-center shrink-0 shadow-md">
                    <AlertTriangle className="w-5 h-5 text-amber-300 animate-bounce" />
                  </div>
                  <div>
                    <div className="text-xs font-black text-rose-900 flex items-center gap-2">
                      <span>DİKKAT: {totalUnmatchedCount} ADET TRANSFER KAYDI SİSTEMDE EŞLEŞMEDİ / DÜŞÜLEMEDİ!</span>
                      <span className="px-2 py-0.5 bg-rose-600 text-white rounded-full text-[10px] font-black">EN BAŞTA LİSTELENİYOR</span>
                    </div>
                    <p className="text-[11px] text-rose-800 mt-0.5">
                      Toplu yüklenen veya adı değişen bu malzemeler stoktan otomatik düşülememiştir. Düzeltmek için ilgili satıra <strong>ÇİFT TIKLAYINIZ</strong> ve açılan pencerede depodaki doğru malzemeyi seçiniz.
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    type="button"
                    onClick={() => setTxStatusFilter(txStatusFilter === 'UNMATCHED' ? 'ALL' : 'UNMATCHED')}
                    className="px-3 py-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-black rounded-xl shadow-xs transition-all active:scale-95 cursor-pointer"
                  >
                    {txStatusFilter === 'UNMATCHED' ? 'Tüm Kayıtları Göster' : `Sadece Eşleşmeyenleri Filtrele (${totalUnmatchedCount})`}
                  </button>
                </div>
              </div>
            )}

            {/* DETAYLI FİLTRELEME VE ARAMA PANELİ */}
            {transactions.length > 0 && (
              <div className="grid grid-cols-1 sm:grid-cols-12 gap-2.5 p-3.5 bg-slate-50 border border-slate-200 rounded-2xl">
                {/* 1. Arama */}
                <div className="sm:col-span-3 relative">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    placeholder="Malzeme adı, P/N, S/N, personel ara..."
                    value={txSearchFilter}
                    onChange={(e) => setTxSearchFilter(e.target.value)}
                    className="w-full pl-9 pr-3 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-medium text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                  />
                  {txSearchFilter && (
                    <button
                      type="button"
                      onClick={() => setTxSearchFilter('')}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer text-xs"
                    >
                      ✕
                    </button>
                  )}
                </div>

                {/* 2. Hava Aracı Tipi Filtresi */}
                <div className="sm:col-span-2">
                  <select
                    value={txAircraftTypeFilter}
                    onChange={(e) => setTxAircraftTypeFilter(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                  >
                    <option value="ALL">TÜM HAVA ARAÇLARI</option>
                    <option value="AT-802">AT-802</option>
                    <option value="BELL-429">BELL-429</option>
                    <option value="T-70">T-70</option>
                    <option value="B-360">B-360</option>
                    <option value="C-650">C-650</option>
                    <option value="GSE">GSE / YER DESTEK</option>
                  </select>
                </div>

                {/* 3. Kuyruk Kodu Filtresi */}
                <div className="sm:col-span-2">
                  <select
                    value={txTailFilter}
                    onChange={(e) => setTxTailFilter(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                  >
                    <option value="ALL">TÜM KUYRUKLAR</option>
                    {Array.from(new Set(transactions.map(t => (t.tailNo || '').trim()).filter(Boolean))).map(tail => (
                      <option key={tail} value={tail}>{tail}</option>
                    ))}
                  </select>
                </div>

                {/* 4. İşlem Türü Filtresi */}
                <div className="sm:col-span-3">
                  <select
                    value={txTypeFilter}
                    onChange={(e) => setTxTypeFilter(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-[#0b3d1d]"
                  >
                    <option value="ALL">TÜM İŞLEM TÜRLERİ ({transactions.length})</option>
                    {STANDARD_TRANSACTION_TYPES.map((t) => {
                      const count = transactions.filter(x => {
                        if (t === 'DİĞER') {
                          return !STANDARD_TRANSACTION_TYPES.filter(st => st !== 'DİĞER').includes(x.type);
                        }
                        return x.type === t;
                      }).length;
                      return (
                        <option key={t} value={t}>
                          {t} {count > 0 ? `(${count})` : ''}
                        </option>
                      );
                    })}
                  </select>
                </div>

                {/* 5. Eşleşme / Düşüm Durumu Filtresi */}
                <div className="sm:col-span-2">
                  <select
                    value={txStatusFilter}
                    onChange={(e) => setTxStatusFilter(e.target.value as any)}
                    className={`w-full px-2 py-1.5 border rounded-xl text-xs font-bold focus:outline-none ${
                      txStatusFilter === 'UNMATCHED' 
                        ? 'bg-rose-100 border-rose-400 text-rose-900' 
                        : 'bg-white border-slate-300 text-slate-900'
                    }`}
                  >
                    <option value="ALL">TÜM KAYITLAR</option>
                    <option value="UNMATCHED">⚠️ DÜŞÜLEMEYENLER ({totalUnmatchedCount})</option>
                    <option value="MATCHED">✅ SİSTEME DÜŞÜLENLER</option>
                  </select>
                </div>
              </div>
            )}

            {transactions.length === 0 ? (
              <div className="p-12 text-center text-slate-400 text-xs font-medium border border-dashed border-slate-200 rounded-2xl">
                Henüz kayıtlı depo işlemi bulunmuyor. "Giriş / Çıkış" veya "Kit Oluşturma" sekmelerinden yeni hareketler gerçekleştirebilirsiniz.
              </div>
            ) : (
              <div className="overflow-x-auto border border-slate-200 rounded-2xl shadow-sm">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-900 text-slate-200 uppercase tracking-wider text-[10px] font-black select-none">
                    <tr>
                      <th className="px-3 py-3 w-10 text-center">
                        <input
                          type="checkbox"
                          checked={isAllSelected}
                          ref={(el) => {
                            if (el) el.indeterminate = isSomeSelected;
                          }}
                          onChange={handleToggleSelectAll}
                          className="w-4 h-4 rounded text-rose-600 focus:ring-rose-500 cursor-pointer accent-rose-600"
                          title="Tümünü Seç / Seçimi Kaldır"
                        />
                      </th>
                      <th className="px-3 py-3">MALZEME ADI</th>
                      <th className="px-3 py-3 text-center">ADET</th>
                      <th className="px-3 py-3 text-center">TARİH</th>
                      <th className="px-3 py-3 text-center">İŞLEM TÜRÜ</th>
                      <th className="px-3 py-3">SERİAL NUMBER</th>
                      <th className="px-3 py-3 text-center">KUYRUK KODU</th>
                      <th className="px-3 py-3">TESLİM ALAN</th>
                      <th className="px-3 py-3">KABUL YAPAN</th>
                      <th className="px-3 py-3">DEPO YERİ</th>
                      <th className="px-3 py-3 text-center">İŞLEMLER</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredTransactions.map((tx, idx) => {
                      const isSelected = selectedTxIds.includes(tx.id);
                      const isUnmatched = isTxUnmatchedWithDepo(tx, depoRows);

                      return (
                        <tr
                          key={tx.id || idx}
                          onDoubleClick={() => {
                            setEditingTx({ ...tx });
                            setEditMaterialSearch(tx.itemName || '');
                          }}
                          title="Düzenlemek için ÇİFT TIKLAYINIZ"
                          className={`transition-colors cursor-pointer select-none ${
                            isUnmatched
                              ? 'bg-rose-50/90 hover:bg-rose-100 border-l-4 border-l-rose-600 text-rose-950 font-semibold'
                              : isSelected
                              ? 'bg-rose-50/70 hover:bg-rose-100/50'
                              : 'hover:bg-slate-50'
                          } ${tx.isUndone ? 'opacity-50 bg-slate-100/60' : ''}`}
                        >
                          <td className="px-3 py-2.5 text-center" onClick={(e) => e.stopPropagation()}>
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => handleToggleSelectRow(tx.id)}
                              className="w-4 h-4 rounded text-rose-600 focus:ring-rose-500 cursor-pointer accent-rose-600"
                            />
                          </td>
                          <td className="px-3 py-2.5 font-bold text-slate-900">
                            <div className="flex items-center gap-1.5 flex-wrap">
                              <span>{tx.itemName}</span>
                              {isUnmatched && (
                                <span className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-rose-600 text-white font-black text-[9px] rounded shadow animate-pulse">
                                  <AlertTriangle className="w-2.5 h-2.5 text-amber-300" />
                                  <span>DÜŞÜLEMEDİ</span>
                                </span>
                              )}
                            </div>
                            {tx.pn && tx.pn !== '-' && (
                              <div className="text-[10px] font-mono text-slate-500">{tx.pn}</div>
                            )}
                          </td>
                          <td className="px-3 py-2.5 text-center font-mono font-black text-slate-900">
                            {tx.quantity}
                          </td>
                          <td className="px-3 py-2.5 text-center font-mono text-slate-700 whitespace-nowrap">
                            {tx.date || (tx.timestamp ? tx.timestamp.split(' ')[0] : '-')}
                          </td>
                          <td className="px-3 py-2.5 text-center">
                            <span className={`px-2 py-0.5 rounded-full font-black text-[10px] ${
                              tx.type.includes('GİRİŞ') || tx.type.includes('GİREN') || tx.type.includes('GELEN') ? 'bg-emerald-100 text-emerald-800' :
                              tx.type.includes('ÇIKIŞ') || tx.type.includes('ÇIKAN') ? 'bg-rose-100 text-rose-800' :
                              tx.type === 'KİT OLUŞTURMA' ? 'bg-purple-100 text-purple-800' :
                              tx.type.includes('TRANSFER') ? 'bg-indigo-100 text-indigo-800' :
                              'bg-amber-100 text-amber-800'
                            }`}>
                              {tx.type}
                            </span>
                          </td>
                          <td className="px-3 py-2.5 font-mono text-slate-600">{tx.sn || '-'}</td>
                          <td className="px-3 py-2.5 text-center font-mono font-bold text-emerald-900">
                            {tx.tailNo || '-'}
                          </td>
                          <td className="px-3 py-2.5 font-bold text-slate-800">{tx.operator || '-'}</td>
                          <td className="px-3 py-2.5 font-bold text-slate-800">{tx.receivedBy || '-'}</td>
                          <td className="px-3 py-2.5 text-slate-700 text-[11px] font-medium max-w-xs truncate" title={tx.location || tx.sourceLocation}>
                            {tx.location || tx.sourceLocation || 'DEPO'}
                          </td>
                          <td className="px-3 py-2.5 text-center" onClick={(e) => e.stopPropagation()}>
                            <div className="flex items-center justify-center gap-1">
                              {/* Düzenle Butonu */}
                              <button
                                type="button"
                                onClick={() => {
                                  setEditingTx({ ...tx });
                                  setEditMaterialSearch(tx.itemName || '');
                                }}
                                className="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-all active:scale-95 cursor-pointer"
                                title="Kaydı Düzenle (Çift Tıklayarak da açabilirsiniz)"
                              >
                                <Wrench className="w-3.5 h-3.5" />
                              </button>

                              {/* Geri Al Butonu */}
                              {!tx.isUndone && (
                                <button
                                  type="button"
                                  onClick={() => handleUndoAction(tx)}
                                  className="p-1 text-amber-700 hover:text-amber-900 hover:bg-amber-50 rounded-lg transition-all active:scale-95 cursor-pointer"
                                  title="İşlemi Geri Al (Stokları Eski Haline Getir)"
                                >
                                  <RotateCcw className="w-3.5 h-3.5" />
                                </button>
                              )}

                              {/* Sil Butonu */}
                              <button
                                type="button"
                                onClick={() => handleDeleteTxRow(tx.id, tx.itemName)}
                                className="p-1 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded-lg transition-all active:scale-95 cursor-pointer"
                                title="Kaydı Kalıcı Olarak Sil"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

      </div>

      {/* TRANSFER GEÇMİŞİ KAYDI DÜZENLEME MODALI (OTOMATİK TAMAMLAMA VE STOK SENKRONİZASYONLU) */}
      {editingTx && (
        <div className="fixed inset-0 z-[10000] flex items-center justify-center p-4 bg-black/80 backdrop-blur-md">
          <div className="bg-white rounded-3xl shadow-2xl max-w-lg w-full flex flex-col max-h-[90vh] overflow-hidden border-2 border-blue-600 animate-scale-up">
            <div className="px-6 py-4 bg-blue-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Wrench className="w-5 h-5 text-blue-300" />
                <div>
                  <h3 className="text-sm font-black uppercase tracking-wider">
                    TRANSFER KAYDINI DÜZENLE
                  </h3>
                  <p className="text-[10px] text-blue-200">
                    Depodaki doğru malzemeyi seçtiğinizde stoklar otomatik eşitlenir.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setEditingTx(null)}
                className="w-7 h-7 rounded-full bg-blue-950 hover:bg-blue-800 text-slate-300 flex items-center justify-center cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-6 overflow-y-auto space-y-3.5 text-xs">
              {/* Malzeme Adı ve Otomatik Tamamlama */}
              <div className="relative">
                <div className="flex items-center justify-between mb-1">
                  <label className="block text-[10px] font-bold text-slate-700 uppercase">
                    MALZEME ADI (YAZDIKÇA DEPO LİSTESİNDEN OTOMATİK DOLDURULUR)
                  </label>
                  {isTxUnmatchedWithDepo(editingTx, depoRows) ? (
                    <span className="text-[9px] font-black text-rose-600 bg-rose-50 px-1.5 py-0.5 rounded">
                      ⚠️ Depo Stoğunda Eşleşmedi
                    </span>
                  ) : (
                    <span className="text-[9px] font-black text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded">
                      ✅ Depo Stoğuyla Eşleşti
                    </span>
                  )}
                </div>

                <div className="relative">
                  <input
                    type="text"
                    value={editingTx.itemName}
                    onFocus={() => setIsEditMaterialDropdownOpen(true)}
                    onChange={(e) => {
                      const val = e.target.value;
                      setEditingTx({ ...editingTx, itemName: val });
                      setIsEditMaterialDropdownOpen(true);
                    }}
                    placeholder="Malzeme adı veya P/N yazınız..."
                    className="w-full px-3 py-2 bg-slate-50 border-2 border-blue-400 rounded-xl font-bold text-slate-900 focus:bg-white focus:border-blue-600 focus:outline-none"
                  />
                  {editingTx.itemName && (
                    <button
                      type="button"
                      onClick={() => setIsEditMaterialDropdownOpen(!isEditMaterialDropdownOpen)}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <ChevronDown className="w-4 h-4" />
                    </button>
                  )}
                </div>

                {/* Otomatik Tamamlama Açılır Listesi */}
                {isEditMaterialDropdownOpen && (
                  <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-300 rounded-xl shadow-2xl z-50 max-h-48 overflow-y-auto divide-y divide-slate-100">
                    <div className="p-2 bg-slate-100 text-[10px] font-bold text-slate-600 uppercase flex items-center justify-between">
                      <span>DEPO LİSTESİNDEKİ EŞLEŞEN MALZEMELER</span>
                      <button 
                        type="button"
                        onClick={() => setIsEditMaterialDropdownOpen(false)}
                        className="text-slate-500 hover:text-slate-800"
                      >
                        Kapat
                      </button>
                    </div>
                    {depoRows
                      .filter(r => {
                        const search = (editingTx.itemName || '').trim().toLowerCase();
                        if (!search) return true;
                        const dName = (r[1] || '').toLowerCase();
                        const dPn = (r[2] || '').toLowerCase();
                        return dName.includes(search) || dPn.includes(search);
                      })
                      .slice(0, 20)
                      .map((r, i) => (
                        <div
                          key={i}
                          onClick={() => {
                            setEditingTx({
                              ...editingTx,
                              itemName: r[1],
                              pn: r[2] && r[2] !== '-' ? r[2] : editingTx.pn,
                              location: r[5] || editingTx.location
                            });
                            setIsEditMaterialDropdownOpen(false);
                          }}
                          className="p-2 hover:bg-blue-50 cursor-pointer flex items-center justify-between text-xs transition-colors"
                        >
                          <div>
                            <div className="font-bold text-slate-900">{r[1]}</div>
                            <div className="text-[10px] text-slate-500 font-mono">P/N: {r[2] || '-'} | Raf: {r[5] || '-'}</div>
                          </div>
                          <div className="text-right">
                            <span className="font-mono font-black text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded text-[11px]">
                              {r[4] || '0 ADET'}
                            </span>
                          </div>
                        </div>
                      ))}
                  </div>
                )}
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">PARÇA NO (P/N)</label>
                  <input
                    type="text"
                    value={editingTx.pn || ''}
                    onChange={(e) => setEditingTx({ ...editingTx, pn: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">ADET / MİKTAR</label>
                  <input
                    type="number"
                    min={1}
                    value={editingTx.quantity}
                    onChange={(e) => setEditingTx({ ...editingTx, quantity: parseFloat(e.target.value) || 1 })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 focus:bg-white"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">TARİH</label>
                  <input
                    type="text"
                    value={editingTx.date}
                    onChange={(e) => setEditingTx({ ...editingTx, date: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1 flex items-center justify-between">
                    <span>İŞLEM TÜRÜ</span>
                    {!STANDARD_TRANSACTION_TYPES.filter(t => t !== 'DİĞER').includes(editingTx.type) && (
                      <span className="text-[9px] text-amber-700 bg-amber-100 font-bold px-1 rounded">Manuel</span>
                    )}
                  </label>
                  <select
                    value={STANDARD_TRANSACTION_TYPES.filter(t => t !== 'DİĞER').includes(editingTx.type) ? editingTx.type : 'DİĞER'}
                    onChange={(e) => {
                      const val = e.target.value;
                      if (val === 'DİĞER') {
                        setEditingTx({ ...editingTx, type: '' });
                      } else {
                        setEditingTx({ ...editingTx, type: val });
                      }
                    }}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 focus:bg-white mb-1.5"
                  >
                    {STANDARD_TRANSACTION_TYPES.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                  {(!STANDARD_TRANSACTION_TYPES.filter(t => t !== 'DİĞER').includes(editingTx.type) || editingTx.type === '') && (
                    <input
                      type="text"
                      placeholder="Manuel işlem türü yazınız..."
                      value={editingTx.type}
                      onChange={(e) => setEditingTx({ ...editingTx, type: e.target.value })}
                      className="w-full px-3 py-1.5 bg-amber-50 border border-amber-400 rounded-xl font-bold text-slate-900 focus:bg-white text-xs"
                      autoFocus
                    />
                  )}
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">SERİ NO (S/N)</label>
                  <input
                    type="text"
                    value={editingTx.sn || ''}
                    onChange={(e) => setEditingTx({ ...editingTx, sn: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-mono text-slate-900 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">KUYRUK KODU</label>
                  <input
                    type="text"
                    value={editingTx.tailNo || ''}
                    onChange={(e) => setEditingTx({ ...editingTx, tailNo: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 focus:bg-white"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">DEPO YERİ / LOKASYON</label>
                  <input
                    type="text"
                    value={editingTx.location || editingTx.sourceLocation || ''}
                    onChange={(e) => setEditingTx({ ...editingTx, location: e.target.value, sourceLocation: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-slate-900 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">TESLİM ALAN</label>
                  <input
                    type="text"
                    value={editingTx.operator || ''}
                    onChange={(e) => setEditingTx({ ...editingTx, operator: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 focus:bg-white"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">KABUL YAPAN</label>
                  <input
                    type="text"
                    value={editingTx.receivedBy || ''}
                    onChange={(e) => setEditingTx({ ...editingTx, receivedBy: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 focus:bg-white"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">AÇIKLAMA / NOTLAR</label>
                  <input
                    type="text"
                    value={editingTx.notes || ''}
                    onChange={(e) => setEditingTx({ ...editingTx, notes: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-slate-900 focus:bg-white"
                  />
                </div>
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-100 border-t border-slate-200 flex items-center justify-between gap-2">
              <span className="text-[11px] text-slate-500 font-medium">
                Kaydettiğinizde depo stoğu güncellenecektir.
              </span>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setEditingTx(null)}
                  className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold rounded-xl cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="button"
                  onClick={() => handleSaveEditedTransaction(editingTx)}
                  className="px-5 py-2 bg-blue-700 hover:bg-blue-800 text-white font-bold rounded-xl shadow-md cursor-pointer flex items-center gap-1.5"
                >
                  <CheckCircle2 className="w-4 h-4 text-emerald-300" />
                  <span>Değişiklikleri Kaydet & Stoktan Düş</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      </div>

      {/* KİT OLUŞTURMA ÖNİZLEME & EXCEL İNDİRME & KAYDET MODALI */}
      {isKitPreviewOpen && (
        <div className="fixed inset-0 z-[10000] flex items-center justify-center p-4 bg-black/80 backdrop-blur-md">
          <div className="bg-white rounded-3xl shadow-2xl max-w-3xl w-full flex flex-col max-h-[85vh] overflow-hidden border-2 border-emerald-600 animate-scale-up">
            
            {/* Modal Header */}
            <div className="px-6 py-4 bg-[#0b3d1d] text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Boxes className="w-5 h-5 text-emerald-300" />
                <h3 className="text-sm font-black uppercase tracking-wider">
                  KİT MONTAJ VE ONAY ÖNİZLEMESİ
                </h3>
              </div>
              <button
                onClick={() => setIsKitPreviewOpen(false)}
                className="w-7 h-7 rounded-full bg-emerald-950 hover:bg-emerald-900 text-slate-300 flex items-center justify-center"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 overflow-y-auto space-y-4">
              
              {/* Kit Bilgi Özeti */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50 p-4 rounded-2xl border border-slate-200">
                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">KİT ADI</span>
                  <span className="text-xs font-black text-slate-900">{kitName}</span>
                </div>
                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">KİT P/N</span>
                  <span className="text-xs font-mono font-bold text-slate-900">{kitPn || '-'}</span>
                </div>
                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">ÜRETİLECEK ADET</span>
                  <span className="text-xs font-mono font-black text-emerald-700">{kitQuantity} KİT</span>
                </div>
                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">HEDEF LOKASYON</span>
                  <span className="text-xs font-bold text-slate-900">{kitLocation}</span>
                </div>
              </div>

              {/* Parçalar Tablosu */}
              <div>
                <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider mb-2">
                  KİT İÇERİSİNDEKİ PARÇALAR VE STOK DÜŞÜŞ TABLOSU
                </h4>
                <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-900 text-slate-200 uppercase tracking-wider text-[10px] font-black">
                      <tr>
                        <th className="px-3 py-2">SIRA</th>
                        <th className="px-4 py-2">PARÇA ADI</th>
                        <th className="px-3 py-2">P/N</th>
                        <th className="px-3 py-2 text-center">1 KİT İÇİN</th>
                        <th className="px-3 py-2 text-center">TOPLAM KULLANILACAK</th>
                        <th className="px-3 py-2 text-center">KALAN STOK</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {kitComponents.map((c, i) => {
                        const totalDeduct = c.quantityPerKit * kitQuantity;
                        const remaining = c.currentStock - totalDeduct;
                        return (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="px-3 py-2 text-center font-mono font-bold text-slate-500">{i + 1}</td>
                            <td className="px-4 py-2 font-bold text-slate-900">{c.itemName}</td>
                            <td className="px-3 py-2 font-mono text-slate-600">{c.pn}</td>
                            <td className="px-3 py-2 text-center font-mono font-bold">{c.quantityPerKit}</td>
                            <td className="px-3 py-2 text-center font-mono font-black text-rose-700">
                              -{totalDeduct}
                            </td>
                            <td className="px-3 py-2 text-center font-mono font-black text-emerald-800">
                              {remaining}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Modal Footer: Excel İndir & Kaydet */}
            <div className="px-6 py-4 bg-slate-100 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={handleDownloadKitPreviewExcel}
                className="px-4 py-2.5 bg-emerald-700 hover:bg-emerald-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow transition-all active:scale-95 cursor-pointer flex items-center gap-2"
              >
                <Download className="w-4 h-4" />
                <span>ÖNİZLEME EXCEL İNDİR</span>
              </button>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsKitPreviewOpen(false)}
                  className="px-4 py-2.5 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs uppercase tracking-wider rounded-xl transition-all"
                >
                  Vazgeç
                </button>
                <button
                  type="button"
                  onClick={handleConfirmAndSaveKit}
                  className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-all active:scale-95 cursor-pointer flex items-center gap-2"
                >
                  <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                  <span>KAYDET VE SİSTEME YENİ ÜRÜN OLARAK EKLE</span>
                </button>
              </div>
            </div>

          </div>
        </div>
      )}

      {/* GİRİŞ / ÇIKIŞ İŞLEMLERİ - ÖN ONAY VE DÜZENLEME EKRANI */}
      {isPreConfirmOpen && pendingTxData && (
        <div className="fixed inset-0 z-[10000] flex items-center justify-center p-4 bg-black/80 backdrop-blur-md">
          <div className="bg-white rounded-3xl shadow-2xl max-w-2xl w-full flex flex-col max-h-[90vh] overflow-hidden border-2 border-emerald-600 animate-scale-up">
            
            {/* Ön Onay Header */}
            <div className="px-6 py-4 bg-gradient-to-r from-[#0b3d1d] to-[#125c2d] text-white flex items-center justify-between shadow-md">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-emerald-500/20 border border-emerald-400/40 flex items-center justify-center">
                  <ShieldCheck className="w-5 h-5 text-emerald-300" />
                </div>
                <div>
                  <h3 className="text-sm font-black uppercase tracking-wider flex items-center gap-2">
                    <span>İŞLEM ÖN ONAY VE DETAY KONTROL EKRANI</span>
                    <span className="px-2 py-0.5 bg-emerald-950/80 text-emerald-300 text-[10px] font-mono rounded">
                      YETKİLİ: 1839
                    </span>
                  </h3>
                  <p className="text-[11px] text-emerald-200">
                    Aşağıdaki bilgileri inceleyiniz ve gerekirse düzenleyiniz. Onay verdiğinizde depo stoğundan otomatik düşülecektir.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsPreConfirmOpen(false)}
                className="w-7 h-7 rounded-full bg-emerald-950/60 hover:bg-emerald-950 text-slate-300 flex items-center justify-center"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Stok Durum ve Değişim Özeti Kartı */}
            <div className="p-4 bg-slate-900 text-white border-b border-slate-800">
              <div className="grid grid-cols-3 gap-3 text-center">
                <div className="bg-slate-800/80 p-2.5 rounded-xl border border-slate-700">
                  <span className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider">İŞLEM TÜRÜ</span>
                  <span className="text-xs font-black text-amber-400 uppercase">{pendingTxData.type}</span>
                </div>

                <div className="bg-slate-800/80 p-2.5 rounded-xl border border-slate-700">
                  <span className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider">MEVCUT STOK</span>
                  <span className="text-sm font-mono font-black text-slate-200">
                    {pendingTxData.currentStock} ADET
                  </span>
                </div>

                <div className="bg-emerald-950/60 p-2.5 rounded-xl border border-emerald-700">
                  <span className="block text-[10px] font-bold text-emerald-400 uppercase tracking-wider">ONAY SONRASI STOK</span>
                  <span className="text-sm font-mono font-black text-emerald-300">
                    {pendingTxData.newStockAfter} ADET
                  </span>
                </div>
              </div>
            </div>

            {/* Ön Onay Düzenlenebilir Alanlar */}
            <div className="p-6 overflow-y-auto space-y-3.5 text-xs bg-slate-50/50">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    MALZEME ADI <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.itemName}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, itemName: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    PARÇA NO (P/N)
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.pn}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, pn: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    İŞLEM MİKTARI (ADET) <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="number"
                    min={1}
                    value={pendingTxData.quantity}
                    onChange={(e) => {
                      const qty = Math.max(1, Number(e.target.value));
                      const isCikis = pendingTxData.selectedOperation === 'cikis' || pendingTxData.type.toUpperCase().includes('ÇIKAN') || pendingTxData.type.toUpperCase().includes('ÇIKIŞ');
                      const isGiris = pendingTxData.selectedOperation === 'giris' || pendingTxData.type.toUpperCase().includes('GİREN') || pendingTxData.type.toUpperCase().includes('GİRİŞ');
                      const newAfter = isCikis 
                        ? Math.max(0, pendingTxData.currentStock - qty) 
                        : (isGiris ? pendingTxData.currentStock + qty : pendingTxData.currentStock);
                      setPendingTxData({
                        ...pendingTxData,
                        quantity: qty,
                        newStockAfter: newAfter
                      });
                    }}
                    className="w-full px-3 py-2 bg-white border border-emerald-600 rounded-xl font-mono font-black text-emerald-800 text-center text-sm focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    SERİ NO (S/N)
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.sn}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, sn: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-mono text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    İŞLEM TARİHİ
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.date}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, date: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    KUYRUK / HAVA ARACI KODU
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.tailNo}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, tailNo: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    DEPO YERİ / RAF KONUMU
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.location}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, location: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    TESLİM ALAN PERSONEL
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.operator}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, operator: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    KABUL YAPAN PERSONEL
                  </label>
                  <input
                    type="text"
                    value={pendingTxData.receivedBy}
                    onChange={(e) => setPendingTxData({ ...pendingTxData, receivedBy: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-black text-slate-700 uppercase tracking-wider mb-1">
                  AÇIKLAMA / İŞ EMRİ / NOTLAR
                </label>
                <input
                  type="text"
                  placeholder="İsteğe bağlı ek notlar..."
                  value={pendingTxData.notes}
                  onChange={(e) => setPendingTxData({ ...pendingTxData, notes: e.target.value })}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-slate-900 focus:ring-2 focus:ring-emerald-600 focus:outline-none"
                />
              </div>
            </div>

            {/* Ön Onay Footer */}
            <div className="px-6 py-4 bg-slate-100 border-t border-slate-200 flex items-center justify-between gap-3">
              <button
                type="button"
                onClick={() => setIsPreConfirmOpen(false)}
                className="px-4 py-2.5 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs uppercase tracking-wider rounded-xl transition-all cursor-pointer"
              >
                ← Düzenlemeye Dön / Vazgeç
              </button>

              <button
                type="button"
                onClick={handleConfirmAndSaveTransaction}
                className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-all active:scale-95 cursor-pointer flex items-center gap-2"
              >
                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                <span>✓ ONAYLA, STOKTAN DÜŞ VE TABLOYA KAYDET</span>
              </button>
            </div>

          </div>
        </div>
      )}
    </>
  );
};
