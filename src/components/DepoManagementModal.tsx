import React, { useState, useEffect, useMemo, useRef } from 'react';
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
  Lock, 
  Unlock,
  X, 
  Download, 
  UploadCloud, 
  Clock, 
  RefreshCw, 
  Edit2,
  Package,
  FlaskConical,
  RotateCcw,
  ChevronDown,
  Barcode,
  FileSearch,
  Wrench
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { DepoTransaction } from '../types';
import { cleanAndFormatDateString, unmergeAndFillWorksheet, isHeaderLikeRow } from '../utils/driveExcelSync';
import { YasamDestekPanel, YasamDestekRecord } from './YasamDestekPanel';
import { DepoSlipPrintModal, DepoSlipData } from './DepoSlipPrintModal';
import { DepoBelgeBulModal } from './DepoBelgeBulModal';
import { PersonnelAutocomplete } from './PersonnelAutocomplete';
import { PnAutocomplete } from './PnAutocomplete';
import { TEKNISYEN_PERSONEL_SCRIPT_URL } from '../utils/personnelData';

export { TEKNISYEN_PERSONEL_SCRIPT_URL };

export const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzhXYZBlJvYarhEYpSK_UdceV-pQwGRIHTjWAVN_UTumI7_qla7vZnAZofdJGeK0e-ZVQ/exec";
export const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

export const UNITS = [
  { id: 'all', name: 'TÜM BİRİMLER', code: 'ORTAK', password: '1839' },
  { id: 'at802', name: 'AT-802F', code: 'UÇAK', password: '802' },
  { id: 'bell429', name: 'BELL 429', code: 'HELİKOPTER', password: '429' },
  { id: 't70', name: 'T-70', code: 'HELİKOPTER', password: '70' },
  { id: 'c650', name: 'C-650', code: 'UÇAK', password: '650' },
  { id: 'b360', name: 'B-360', code: 'UÇAK', password: '360' },
  { id: 'hangar', name: 'HANGAR GENEL', code: 'GENEL', password: '1839' }
];

export const AIRCRAFT_TAILS: Record<string, string[]> = {
  at802: [
    'ORMAN 21 (OR-2021) - AT-802', 'ORMAN 22 (OR-2022) - AT-802', 'ORMAN 23 (OR-2023) - AT-802',
    'ORMAN 24 (OR-2024) - AT-802', 'ORMAN 25 (OR-2025) - AT-802', 'ORMAN 26 (OR-2026) - AT-802',
    'ORMAN 27 (OR-2027) - AT-802', 'ORMAN 28 (OR-2028) - AT-802', 'ORMAN 29 (OR-2029) - AT-802',
    'ORMAN 30 (OR-2030) - AT-802', 'ORMAN 31 (OR-2031) - AT-802', 'ORMAN 36 (OR-2036) - AT-802',
    'ORMAN 37 (OR-2037) - AT-802', 'ORMAN 38 (OR-2038) - AT-802', 'ORMAN 39 (OR-2039) - AT-802',
    'ORMAN 40 (OR-2040) - AT-802'
  ],
  bell429: [
    'ORMAN 03 (OR 3125) - BELL-429', 'ORMAN 04 (OR 3126) - BELL-429', 'ORMAN 05 (OR 3127) - BELL-429',
    'ORMAN 06 (OR 3131) - BELL-429', 'ORMAN 07 (OR 3135) - BELL-429', 'ORMAN 08 (OR 3192) - BELL-429'
  ],
  t70: [
    'ORMAN 18 (OR-1018) - T-70', 'ORMAN 19 (OR-1019) - T-70', 'ORMAN 20 (OR-1020) - T-70'
  ],
  c650: ['ORMAN 01 (OR 0177) - C-650'],
  b360: ['ORMAN 02 (OR 1839) - B-360'],
  hangar: ['GENEL HANGAR BAKIM', 'YER TESTİ / REVİZYON']
};

export const AT802_GIRIS_TYPES = ['ANKARA GİREN'];
export const AT802_CIKIS_TYPES = ['ANKARA ÇIKAN', 'KARAİN ÇIKAN', 'ÇANAKKALE ÇIKAN', 'MİLAS ÇIKAN', 'BURSA ÇIKAN', 'MUAYENE GİDEN', 'DİĞER'];
export const AT802_TRANSFER_TYPES = ['KARAİN TRANSFER', 'ÇANAKKALE TRANSFER', 'MİLAS TRANSFER', 'BURSA TRANSFER', 'ANTALYA TRANSFER', 'DEPO TRANSFER', 'DİĞER'];

export const DEFAULT_TRANSFER_TYPES: Record<string, string[]> = {
  at802: [
    'ANKARA GİREN',
    'ANKARA ÇIKAN', 'KARAİN ÇIKAN', 'ÇANAKKALE ÇIKAN', 'MİLAS ÇIKAN', 'BURSA ÇIKAN',
    'KARAİN TRANSFER', 'ÇANAKKALE TRANSFER', 'MİLAS TRANSFER', 'BURSA TRANSFER', 'ANTALYA TRANSFER', 'DEPO TRANSFER',
    'MUAYENE GİDEN', 'MUAYENE GELEN',
    'DİĞER'
  ],
  bell429: [
    'MUĞLA GİREN', 'MUĞLA ÇIKAN', 'MUĞLA TRANSFER',
    'İZMİR GİREN', 'İZMİR ÇIKAN', 'İZMİR TRANSFER',
    'ÇANAKKALE GİREN', 'ÇANAKKALE ÇIKAN', 'ÇANAKKALE TRANSFER',
    'ANTALYA GİREN', 'ANTALYA ÇIKAN', 'ANTALYA TRANSFER',
    'ANKARA GİREN', 'ANKARA ÇIKAN',
    'DİĞER'
  ],
  t70: [
    'MUĞLA/GÜVERCİNLİK GİREN', 'MUĞLA/GÜVERCİNLİK ÇIKAN', 'MUĞLA/GÜVERCİNLİK TRANSFER',
    'MUĞLA/YANIKLAR GİREN', 'MUĞLA/YANIKLAR ÇIKAN', 'MUĞLA/YANIKLAR TRANSFER',
    'ANKARA GİREN', 'ANKARA ÇIKAN',
    'ANTALYA GİREN', 'ANTALYA ÇIKAN', 'ANTALYA TRANSFER',
    'DİĞER'
  ],
  c650: [
    'ANKARA GİREN', 'ANKARA ÇIKAN',
    'ANTALYA GİREN', 'ANTALYA ÇIKAN', 'ANTALYA TRANSFER',
    'ÇANAKKALE GİREN', 'ÇANAKKALE ÇIKAN', 'ÇANAKKALE TRANSFER',
    'DİĞER'
  ],
  b360: [
    'ANKARA GİREN', 'ANKARA ÇIKAN',
    'ANTALYA GİREN', 'ANTALYA ÇIKAN', 'ANTALYA TRANSFER',
    'ÇANAKKALE GİREN', 'ÇANAKKALE ÇIKAN', 'ÇANAKKALE TRANSFER',
    'DİĞER'
  ],
  hangar: [
    'HANGAR GİREN', 'HANGAR ÇIKAN',
    'DEPO TRANSFER', 'İADE / HURDA',
    'DİĞER'
  ]
};

export interface DepoItem {
  unit: string;
  category: 'sarf' | 'kimyasal';
  description: string;
  name?: string;
  partNumber: string;
  pn?: string;
  serialAndNotes: string;
  sn?: string;
  lokasyonNo: string;
  location?: string;
  baseGelen: number;
  gelen: number;
  spainDescription?: string;
  miktarQty?: number | string;
  sozlesmeToplami?: number | string;
  piyasa?: number | string;
  spainGelen?: number | string;
  genelToplam?: number | string;
  // Computed metrics
  ankaraCikan?: number;
  ankaraMevcut?: number;
  karainTransfer?: number;
  karainCikan?: number;
  karainMevcut?: number;
  canakkaleTransfer?: number;
  canakkaleCikan?: number;
  canakkaleMevcut?: number;
  milasTransfer?: number;
  milasCikan?: number;
  milasMevcut?: number;
  bursaTransfer?: number;
  bursaCikan?: number;
  bursaMevcut?: number;
  muayeneGiden?: number;
  muayeneGelen?: number;
  muayeneToplam?: number;
  toplamStok?: number;
  hasShelfLife?: 'EVET' | 'HAYIR' | boolean | string;
  shelfLifeDate?: string;
  durumu?: string;
  tedarikci?: string;
  zone?: string;
  notes?: string;
  aciklama?: string;
}

interface DepoManagementModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialUnit?: string;
  depoRows?: string[][];
  at802Rows?: string[][];
  onUpdateDepoRows?: (updatedRows: string[][]) => void;
  transactions?: DepoTransaction[];
  onAddTransaction?: (tx: Omit<DepoTransaction, 'id'> | (Omit<DepoTransaction, 'id' | 'timestamp'> & { timestamp?: string })) => void;
  onDeleteTransaction?: (txId: string) => void;
  certificatePdfUrl?: string | null;
  showNotification?: (msg: string) => void;
}

const PaginationBar: React.FC<{
  currentPage: number;
  totalItems: number;
  pageSize: number;
  onPageChange: (page: number) => void;
  onPageSizeChange: (size: number) => void;
}> = ({
  currentPage,
  totalItems,
  pageSize,
  onPageChange,
  onPageSizeChange
}) => {
  const totalPages = pageSize > 0 ? Math.max(1, Math.ceil(totalItems / pageSize)) : 1;
  const startItem = totalItems === 0 ? 0 : pageSize > 0 ? (currentPage - 1) * pageSize + 1 : 1;
  const endItem = pageSize > 0 ? Math.min(totalItems, currentPage * pageSize) : totalItems;

  return (
    <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-2.5 bg-slate-50 border-t border-slate-200 text-xs select-none">
      <div className="flex items-center gap-2 text-slate-600 font-medium">
        <span>Görüntülenen: <strong className="text-slate-900 font-bold">{startItem}-{endItem}</strong> / Toplam <strong>{totalItems}</strong></span>
        <span className="text-slate-300">|</span>
        <label className="flex items-center gap-1.5 text-slate-500">
          <span>Sayfa Başı:</span>
          <select
            value={pageSize}
            onChange={(e) => {
              onPageSizeChange(Number(e.target.value));
              onPageChange(1);
            }}
            className="px-2 py-0.5 bg-white border border-slate-300 rounded text-slate-800 font-bold text-xs cursor-pointer focus:ring-1 focus:ring-emerald-500"
          >
            <option value={50}>50</option>
            <option value={100}>100</option>
            <option value={200}>200</option>
            <option value={-1}>Tümü</option>
          </select>
        </label>
      </div>

      {pageSize > 0 && totalPages > 1 && (
        <div className="flex items-center gap-1.5 font-mono flex-wrap">
          <button
            type="button"
            disabled={currentPage <= 1}
            onClick={() => onPageChange(currentPage - 1)}
            className="px-2.5 py-1 rounded bg-white border border-slate-300 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 cursor-pointer text-xs font-bold"
            title="Önceki Sayfa"
          >
            ‹ Önceki
          </button>

          {(() => {
            const pages: (number | string)[] = [];
            if (totalPages <= 7) {
              for (let i = 1; i <= totalPages; i++) pages.push(i);
            } else {
              pages.push(1);
              if (currentPage > 3) pages.push('...');
              const start = Math.max(2, currentPage - 1);
              const end = Math.min(totalPages - 1, currentPage + 1);
              for (let i = start; i <= end; i++) {
                if (!pages.includes(i)) pages.push(i);
              }
              if (currentPage < totalPages - 2) pages.push('...');
              if (!pages.includes(totalPages)) pages.push(totalPages);
            }
            return pages.map((p, idx) => {
              if (p === '...') {
                return <span key={`ell-${idx}`} className="px-1.5 py-1 text-slate-400 font-bold select-none text-xs">...</span>;
              }
              const num = Number(p);
              const isActive = num === currentPage;
              return (
                <button
                  key={`p-${num}`}
                  type="button"
                  onClick={() => onPageChange(num)}
                  className={`min-w-[28px] px-2 py-1 rounded text-xs font-mono font-bold transition-all cursor-pointer ${
                    isActive
                      ? 'bg-emerald-600 text-white shadow-xs font-black ring-1 ring-emerald-400'
                      : 'bg-white border border-slate-300 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  {num}
                </button>
              );
            });
          })()}

          <button
            type="button"
            disabled={currentPage >= totalPages}
            onClick={() => onPageChange(currentPage + 1)}
            className="px-2.5 py-1 rounded bg-white border border-slate-300 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 cursor-pointer text-xs font-bold"
            title="Sonraki Sayfa"
          >
            Sonraki Sayfa ›
          </button>
        </div>
      )}
    </div>
  );
};

export const DepoManagementModal: React.FC<DepoManagementModalProps> = ({
  isOpen,
  onClose,
  initialUnit = 'at802',
  transactions: initialTransactions = [],
  onAddTransaction,
  showNotification = (_msg: string) => {}
}) => {
  const [currentUnit, setCurrentUnit] = useState<string>(initialUnit || 'at802');
  const [currentDepoType, setCurrentDepoType] = useState<'sarf' | 'kimyasal' | 'yasam_destek' | 'transactions'>('sarf');
  const [isYasamDestekExportTriggered, setIsYasamDestekExportTriggered] = useState(false);
  const [unlockedUnits, setUnlockedUnits] = useState<Set<string>>(new Set(['all', initialUnit || 'at802']));

  // Yaşam Destek (AT-802) State
  const [yasamDestekRecords, setYasamDestekRecords] = useState<YasamDestekRecord[]>([]);
  const [isYasamDestekLoading, setIsYasamDestekLoading] = useState(false);

  const fetchYasamDestekData = async () => {
    try {
      setIsYasamDestekLoading(true);
      const res = await fetch('/api/yasam-destek');
      const json = await res.json();
      if (json && json.status === 'success' && Array.isArray(json.records)) {
        setYasamDestekRecords(json.records);
      }
    } catch (err) {
      console.warn('Yaşam destek fetch error:', err);
    } finally {
      setIsYasamDestekLoading(false);
    }
  };

  const handleSaveYasamDestekRecord = async (record: YasamDestekRecord): Promise<boolean> => {
    try {
      const res = await fetch('/api/yasam-destek', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ record })
      });
      const json = await res.json();
      if (json && json.status === 'success') {
        if (Array.isArray(json.records)) {
          setYasamDestekRecords(json.records);
        } else {
          fetchYasamDestekData();
        }
        return true;
      }
    } catch (err) {
      console.warn('Yaşam destek save error:', err);
    }
    return false;
  };

  const handleDeleteYasamDestekRecord = async (id: string): Promise<boolean> => {
    try {
      const res = await fetch('/api/yasam-destek/delete-record', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id })
      });
      const json = await res.json();
      if (json && json.status === 'success') {
        const updated = yasamDestekRecords.filter(r => r.id !== id);
        setYasamDestekRecords(updated);
        showNotification('🗑️ Yaşam destek kaydı silindi.');
        return true;
      }
    } catch (err) {
      console.warn('Yaşam destek delete error:', err);
    }
    return false;
  };
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [isDepoSearchPerformed, setIsDepoSearchPerformed] = useState<boolean>(false);
  const [filterRegion, setFilterRegion] = useState<string>('ALL');

  // Excel Input Ref
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Password Modal
  const [isPasswordModalOpen, setIsPasswordModalOpen] = useState<boolean>(false);
  const [pendingTargetUnit, setPendingTargetUnit] = useState<string | null>(null);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [passwordError, setPasswordError] = useState<boolean>(false);

  // Movement Modal
  const [isMovementModalOpen, setIsMovementModalOpen] = useState<boolean>(false);
  const [movementDirection, setMovementDirection] = useState<'cikis' | 'giris' | 'transfer'>('cikis');
  const [selectedItemIdx, setSelectedItemIdx] = useState<number>(-1);
  const [movementItemSearch, setMovementItemSearch] = useState<string>('');
  const [movementQty, setMovementQty] = useState<number>(1);
  const [movementDate, setMovementDate] = useState<string>('');
  const [movementType, setMovementType] = useState<string>('');
  const [movementCustomType, setMovementCustomType] = useState<string>('');
  const [movementSn, setMovementSn] = useState<string>('');
  const [movementTail, setMovementTail] = useState<string>('');
  const [movementTeslimAlan, setMovementTeslimAlan] = useState<string>('');
  const [movementKabulYapan, setMovementKabulYapan] = useState<string>('');
  const [movementDepoYeri, setMovementDepoYeri] = useState<string>('');
  const [movementNotes, setMovementNotes] = useState<string>('');
  const [isItemDropdownOpen, setIsItemDropdownOpen] = useState<boolean>(false);

  // Mobile View Specific States inside DepoManagementModal
  const [isMobile, setIsMobile] = useState<boolean>(false);
  const [mobileSearchQuery, setMobileSearchQuery] = useState<string>('');
  const [showMobileSuggestions, setShowMobileSuggestions] = useState<boolean>(false);
  const [selectedMobileItem, setSelectedMobileItem] = useState<DepoItem | null>(null);
  const [mobileCategory, setMobileCategory] = useState<'sarf' | 'kimyasal'>('sarf');

  useEffect(() => {
    const checkMobile = () => {
      setIsMobile(window.innerWidth < 768 || /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent));
    };
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);




  // Termal Barkod & Slip Basma Modal State
  const [isSlipModalOpen, setIsSlipModalOpen] = useState<boolean>(false);
  const [isGirisOnlySlip, setIsGirisOnlySlip] = useState<boolean>(false);
  const [slipModalType, setSlipModalType] = useState<'raf' | 'giris' | null>(null);
  const [slipModalData, setSlipModalData] = useState<Partial<DepoSlipData> | null>(null);

  // Belge Bul Modal State
  const [isBelgeBulModalOpen, setIsBelgeBulModalOpen] = useState<boolean>(false);

  // Depo Hareket Geçmişi Düzenleme & Şifre State
  const [selectedTxForEdit, setSelectedTxForEdit] = useState<DepoTransaction | null>(null);
  const [isTxEditModalOpen, setIsTxEditModalOpen] = useState<boolean>(false);
  const [isTxPasswordModalOpen, setIsTxPasswordModalOpen] = useState<boolean>(false);
  const [txPasswordInput, setTxPasswordInput] = useState<string>('');
  const [txPasswordError, setTxPasswordError] = useState<string>('');
  const [pendingTxToEdit, setPendingTxToEdit] = useState<DepoTransaction | null>(null);
  const [isTxSessionUnlocked, setIsTxSessionUnlocked] = useState<boolean>(false);

  // Edit form states
  const [editTxDate, setEditTxDate] = useState<string>('');
  const [editTxType, setEditTxType] = useState<string>('');
  const [editTxItemName, setEditTxItemName] = useState<string>('');
  const [editTxPn, setEditTxPn] = useState<string>('');
  const [editTxSn, setEditTxSn] = useState<string>('');
  const [editTxQuantity, setEditTxQuantity] = useState<number | string>(1);
  const [editTxTailNo, setEditTxTailNo] = useState<string>('');
  const [editTxOperator, setEditTxOperator] = useState<string>('');
  const [editTxReceivedBy, setEditTxReceivedBy] = useState<string>('');
  const [editTxLocation, setEditTxLocation] = useState<string>('');
  const [editTxNotes, setEditTxNotes] = useState<string>('');

  // New Item Modal
  const [isNewItemModalOpen, setIsNewItemModalOpen] = useState<boolean>(false);
  const [editingItemIdx, setEditingItemIdx] = useState<number>(-1);
  const [newItemUnit, setNewItemUnit] = useState<string>('at802');
  const [newItemCategory, setNewItemCategory] = useState<'sarf' | 'kimyasal'>('sarf');
  const [newItemDesc, setNewItemDesc] = useState<string>('');
  const [newItemPn, setNewItemPn] = useState<string>('');
  const [newItemLokasyon, setNewItemLokasyon] = useState<string>('');
  const [newItemSnNotes, setNewItemSnNotes] = useState<string>('');
  const [newItemGelen, setNewItemGelen] = useState<number>(1);
  const [newItemInitialRegion, setNewItemInitialRegion] = useState<string>('ankara');
  const [newItemHasShelfLife, setNewItemHasShelfLife] = useState<'HAYIR' | 'EVET'>('HAYIR');
  const [newItemShelfLifeDate, setNewItemShelfLifeDate] = useState<string>('');

  // Multi-SN & Non-catalog detection state
  const [singleSnInput, setSingleSnInput] = useState<string>('');
  const [snList, setSnList] = useState<string[]>([]);

  // Lokasyon Çoklu Filtresi (Excel Stili Checkbox Listesi - Tamam butonuna basınca filtre uygulansın)
  const [selectedLocations, setSelectedLocations] = useState<string[]>([]);
  const [pendingLocations, setPendingLocations] = useState<string[]>([]);
  const [isLocationDropdownOpen, setIsLocationDropdownOpen] = useState<boolean>(false);
  const [locationSearchTerm, setLocationSearchTerm] = useState<string>('');
  const locationDropdownRef = useRef<HTMLDivElement | null>(null);

  // Hareket Geçmişi / Envanterden Çoklu Parça Seçip Kit Oluşturma State'leri
  const [selectedTxIdsForKit, setSelectedTxIdsForKit] = useState<string[]>([]);
  const [selectedItemPnsForKit, setSelectedItemPnsForKit] = useState<string[]>([]);

  // Muayene / Bakım Gönder Modalı State'leri
  const [isMuayeneModalOpen, setIsMuayeneModalOpen] = useState<boolean>(false);
  const [muayenePn, setMuayenePn] = useState<string>('');
  const [muayeneSn, setMuayeneSn] = useState<string>('');
  const [muayeneDescription, setMuayeneDescription] = useState<string>('');
  const [muayeneQty, setMuayeneQty] = useState<number>(1);
  const [muayeneDate, setMuayeneDate] = useState<string>('');
  const [muayeneNotes, setMuayeneNotes] = useState<string>('');
  const [muayeneSourceDepot, setMuayeneSourceDepot] = useState<string>('ANKARA');
  const [muayeneOperator, setMuayeneOperator] = useState<string>('');
  const [isMuayeneDropdownOpen, setIsMuayeneDropdownOpen] = useState<boolean>(false);
  const [selectedMuayeneItem, setSelectedMuayeneItem] = useState<DepoItem | null>(null);

  // Onay Bekleyenler Modalı State'leri (Drive Excel "ONAY BEKLEYENLER -AT802" Sayfası)
  const [isOnayPasswordModalOpen, setIsOnayPasswordModalOpen] = useState<boolean>(false);
  const [isOnayBekleyenlerModalOpen, setIsOnayBekleyenlerModalOpen] = useState<boolean>(false);
  const [onayPasswordInput, setOnayPasswordInput] = useState<string>('');
  const [onayPasswordError, setOnayPasswordError] = useState<string>('');
  const [onayBekleyenlerList, setOnayBekleyenlerList] = useState<any[]>([]);
  const [isLoadingOnayList, setIsLoadingOnayList] = useState<boolean>(false);

  // Muayene Oturum Listesi State'leri (Çoklu Ekleme Açılır Kart)
  const [muayeneSessionList, setMuayeneSessionList] = useState<any[]>([]);
  const [isMuayeneSessionListOpen, setIsMuayeneSessionListOpen] = useState<boolean>(true);

  // Transfer Geri Çekme State'leri
  const [isRollbackModalOpen, setIsRollbackModalOpen] = useState<boolean>(false);
  const [rollbackTx, setRollbackTx] = useState<DepoTransaction | null>(null);
  const [rollbackQty, setRollbackQty] = useState<number>(1);
  const [rollbackReason, setRollbackReason] = useState<string>('');
  const [rollbackOperator, setRollbackOperator] = useState<string>('');

  const handleOpenRollbackModal = (tx: DepoTransaction) => {
    setRollbackTx(tx);
    setRollbackQty(Number(tx.quantity || 1));
    setRollbackReason('');
    setRollbackOperator(tx.operator || tx.teslimAlan || '');
    setIsRollbackModalOpen(true);
  };

  const handleConfirmRollback = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!rollbackTx) return;

    if (!rollbackReason.trim()) {
      alert('Lütfen transfer geri çekme nedenini yazınız!');
      return;
    }

    if (rollbackQty < 1) {
      alert('Geri çekilecek miktar en az 1 olmalıdır!');
      return;
    }

    const origQty = Number(rollbackTx.quantity || 1);
    if (rollbackQty > origQty) {
      alert(`Geri çekilecek miktar orijinal transfer miktarından (${origQty}) fazla olamaz!`);
      return;
    }

    // Determine source & target depot from tx
    const typeStr = (rollbackTx.type || rollbackTx.islemTuru || '').toUpperCase();
    let targetDepot = 'KARAİN';
    if (typeStr.includes('ÇANAKKALE')) targetDepot = 'ÇANAKKALE';
    else if (typeStr.includes('MİLAS')) targetDepot = 'MİLAS';
    else if (typeStr.includes('BURSA')) targetDepot = 'BURSA';
    else if (rollbackTx.location) {
      const locU = rollbackTx.location.toUpperCase();
      if (locU.includes('ÇANAKKALE')) targetDepot = 'ÇANAKKALE';
      else if (locU.includes('MİLAS')) targetDepot = 'MİLAS';
      else if (locU.includes('BURSA')) targetDepot = 'BURSA';
    }

    const sourceDepot = 'ANKARA';

    // Reverse stock update
    const updatedInventory = [...inventory];
    const itemIdx = updatedInventory.findIndex(i => {
      const matchPn = (i.partNumber || i.pn || '').trim().toLowerCase() === (rollbackTx.pn || '').trim().toLowerCase();
      const matchDesc = (i.description || i.name || '').trim().toLowerCase() === (rollbackTx.itemName || '').trim().toLowerCase();
      return matchPn || matchDesc;
    });

    if (itemIdx >= 0) {
      const it = { ...updatedInventory[itemIdx] };
      // Return to Ankara (source)
      it.ankaraMevcut = (it.ankaraMevcut || 0) + rollbackQty;

      // Decrease from target
      if (targetDepot.includes('KARAİN')) {
        it.karainMevcut = Math.max(0, (it.karainMevcut || 0) - rollbackQty);
      } else if (targetDepot.includes('ÇANAKKALE')) {
        it.canakkaleMevcut = Math.max(0, (it.canakkaleMevcut || 0) - rollbackQty);
      } else if (targetDepot.includes('MİLAS')) {
        it.milasMevcut = Math.max(0, (it.milasMevcut || 0) - rollbackQty);
      } else if (targetDepot.includes('BURSA')) {
        it.bursaMevcut = Math.max(0, (it.bursaMevcut || 0) - rollbackQty);
      }

      it.toplamStok = (it.ankaraMevcut || 0) + (it.karainMevcut || 0) + (it.canakkaleMevcut || 0) + (it.milasMevcut || 0) + (it.bursaMevcut || 0);
      updatedInventory[itemIdx] = it;
    }

    // New transaction for History Log
    const rollbackNewTx: DepoTransaction = {
      id: `tx_rollback_${Date.now()}`,
      timestamp: new Date().toISOString(),
      type: `${targetDepot} TRANSFER GERİ ÇEKME`,
      itemName: rollbackTx.itemName,
      pn: rollbackTx.pn || '-',
      sn: rollbackTx.sn || '-',
      quantity: rollbackQty,
      date: new Date().toLocaleString('tr-TR'),
      tailNo: rollbackTx.tailNo || '-',
      operator: rollbackOperator.trim() || 'Depo Sorumlusu',
      receivedBy: 'Ankara Merkez Depo',
      location: `${targetDepot} ➔ ANKARA (GERİ ÇEKİLDİ)`,
      unit: rollbackTx.unit || currentUnit,
      category: rollbackTx.category || 'sarf',
      sheetName: rollbackTx.sheetName || 'DEPO HAREKET GEÇMİŞİ-AT-802',
      notes: `TRANSFER GERİ ÇEKME NEDENİ: ${rollbackReason.trim()} (Orijinal Transfer ID: ${rollbackTx.id})`,
      isNewSessionTx: true
    };

    const allTransfers = [rollbackNewTx, ...transactions];
    setTransactions(allTransfers);
    setInventory(updatedInventory);

    if (onAddTransaction) onAddTransaction(rollbackNewTx);

    setIsRollbackModalOpen(false);
    setRollbackTx(null);

    showNotification(`✅ "${rollbackNewTx.itemName}" parçası (${rollbackQty} adet) ${sourceDepot} ana depoya başarıyla geri çekildi.`);
    syncDepoToExcelOnline(updatedInventory, allTransfers);
  };
  const [editingOnayId, setEditingOnayId] = useState<string | null>(null);
  const [editOnayPn, setEditOnayPn] = useState<string>('');
  const [editOnaySn, setEditOnaySn] = useState<string>('');
  const [editOnayItemName, setEditOnayItemName] = useState<string>('');
  const [editOnayQuantity, setEditOnayQuantity] = useState<number>(1);
  const [editOnayDepot, setEditOnayDepot] = useState<string>('ANKARA');
  const [editOnayTailNo, setEditOnayTailNo] = useState<string>('');
  const [editOnayRequestedBy, setEditOnayRequestedBy] = useState<string>('');
  const [editOnayDate, setEditOnayDate] = useState<string>('');
  const [editOnayNotes, setEditOnayNotes] = useState<string>('');

  const handleStartEditOnay = (item: any) => {
    setEditingOnayId(item.id);
    setEditOnayPn(item.pn || '');
    setEditOnaySn(item.sn || '');
    setEditOnayItemName(item.itemName || item.description || '');
    setEditOnayQuantity(Number(item.quantity || item.adet || 1));
    setEditOnayDepot(item.depot || 'ANKARA');
    setEditOnayTailNo(item.tailNo || '');
    setEditOnayRequestedBy(item.requestedBy || '');
    setEditOnayDate(item.date || new Date().toLocaleString('tr-TR'));
    setEditOnayNotes(item.notes || '');
  };

  const handleSaveOnayEdit = async (id: string) => {
    try {
      const res = await fetch('/api/onay-bekleyenler/guncelle', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id,
          updatedItem: {
            pn: editOnayPn,
            sn: editOnaySn,
            itemName: editOnayItemName,
            quantity: Number(editOnayQuantity),
            depot: editOnayDepot,
            tailNo: editOnayTailNo,
            requestedBy: editOnayRequestedBy,
            date: editOnayDate,
            notes: editOnayNotes
          }
        })
      });
      const data = await res.json();
      if (data && data.status === 'success') {
        showNotification('✅ Onay bekleyen talep başarıyla güncellendi.');
        setEditingOnayId(null);
        fetchOnayBekleyenler();
      } else {
        alert(data?.message || 'Güncelleme hatası!');
      }
    } catch (e: any) {
      alert('Hata: ' + e.message);
    }
  };

  // Kuyruk No "Diğer" seçimi için manuel giriş state'i
  const [customTailInput, setCustomTailInput] = useState<string>('');

  // Depo Hareket Geçmişi Silme & Şifre State
  const [pendingTxToDelete, setPendingTxToDelete] = useState<DepoTransaction | null>(null);
  const [isTxDeleteModalOpen, setIsTxDeleteModalOpen] = useState<boolean>(false);
  const [txDeletePasswordInput, setTxDeletePasswordInput] = useState<string>('');
  const [txDeletePasswordError, setTxDeletePasswordError] = useState<string>('');

  // Depoda Mevcut Ürün Satırına Tıklayınca Açılan Düzenleme & Şifre State
  const [rowEditItem, setRowEditItem] = useState<DepoItem | null>(null);
  const [isRowEditModalOpen, setIsRowEditModalOpen] = useState<boolean>(false);
  const [rowEditPn, setRowEditPn] = useState<string>('');
  const [rowEditSn, setRowEditSn] = useState<string>('');
  const [rowEditDescription, setRowEditDescription] = useState<string>('');
  const [rowEditLocation, setRowEditLocation] = useState<string>('');
  const [rowEditNotes, setRowEditNotes] = useState<string>('');
  const [rowEditHasShelfLife, setRowEditHasShelfLife] = useState<'EVET' | 'HAYIR'>('HAYIR');
  const [rowEditShelfLifeDate, setRowEditShelfLifeDate] = useState<string>('');
  const [rowEditPassword, setRowEditPassword] = useState<string>('');
  const [rowEditPasswordError, setRowEditPasswordError] = useState<string>('');

  // Depolar Arası Transfer Aşamalı Akış State'leri (İlk açılışta Transfer Merkezi hep ANKARA seçilidir)
  const [transferSourceDepot, setTransferSourceDepot] = useState<string>('ANKARA');
  const [transferTargetDepot, setTransferTargetDepot] = useState<string>('KARAİN');

  // Transactions Performance Search & Filter States
  const [txSearchInput, setTxSearchInput] = useState<string>('');
  const [txSearchActiveQuery, setTxSearchActiveQuery] = useState<string>('');
  const [txFilterType, setTxFilterType] = useState<string>('ALL');
  const [txFilterTail, setTxFilterTail] = useState<string>('ALL');
  const [txFilterStartDate, setTxFilterStartDate] = useState<string>('');
  const [txFilterEndDate, setTxFilterEndDate] = useState<string>('');
  const [isTxSearchPerformed, setIsTxSearchPerformed] = useState<boolean>(false);

  // Force reset tx search when tab/unit/modal opens
  useEffect(() => {
    setIsTxSearchPerformed(false);
  }, [currentDepoType, currentUnit, isOpen]);

  // Click-away listener for Excel location filter dropdown
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (locationDropdownRef.current && !locationDropdownRef.current.contains(e.target as Node)) {
        setIsLocationDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Fetch Onay Bekleyenler List
  const fetchOnayBekleyenler = async () => {
    try {
      setIsLoadingOnayList(true);
      const res = await fetch('/api/onay-bekleyenler');
      const data = await res.json();
      if (data && data.status === 'success' && Array.isArray(data.items)) {
        setOnayBekleyenlerList(data.items);
      }
    } catch (err) {
      console.warn("Failed to load onay bekleyenler:", err);
    } finally {
      setIsLoadingOnayList(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      fetchOnayBekleyenler();
    }
  }, [isOpen]);

  const pendingApprovalCount = useMemo(() => {
    return onayBekleyenlerList.filter(it => {
      if (it.status !== 'ONAY BEKLİYOR') return false;
      if (!currentUnit || currentUnit === 'all') return true;
      const itemUnit = (it.unit || 'at802').toLowerCase();
      const cUnit = currentUnit.toLowerCase();
      return itemUnit === cUnit || (cUnit === 'at802' && (!it.unit || itemUnit === 'at802'));
    }).length;
  }, [onayBekleyenlerList, currentUnit]);

  // Item Detail Modal States
  const [detailModalItem, setDetailModalItem] = useState<DepoItem | null>(null);
  const [detailModalColumnType, setDetailModalColumnType] = useState<string>('ALL');
  const [isDetailModalOpen, setIsDetailModalOpen] = useState<boolean>(false);
  const [isRegionalStockOpen, setIsRegionalStockOpen] = useState<boolean>(false);

  // Cloud Sync state
  const [isSyncingCloud, setIsSyncingCloud] = useState<boolean>(false);
  const [isLoadingDrive, setIsLoadingDrive] = useState<boolean>(false);

  // Bulk Excel Giriş/Çıkış State
  const [isBulkExcelModalOpen, setIsBulkExcelModalOpen] = useState<boolean>(false);
  const [isExcelTypeModalOpen, setIsExcelTypeModalOpen] = useState<boolean>(false);
  const [selectedExcelUploadCategory, setSelectedExcelUploadCategory] = useState<'sarf' | 'kimyasal' | 'yasam_destek'>('sarf');
  const [bulkExcelRows, setBulkExcelRows] = useState<Array<{
    id: number;
    pn: string;
    sn: string;
    qty: number;
    matchedItem: DepoItem | null;
    status: 'Eşleşti' | 'Eşleşmedi';
    date?: string;
    type?: string;
    tailNo?: string;
    operator?: string;
    receivedBy?: string;
  }>>([]);
  const [bulkSearchQueries, setBulkSearchQueries] = useState<Record<number, string>>({});
  const [bulkDropdownOpenRow, setBulkDropdownOpenRow] = useState<number | null>(null);

  // Inventory & Transactions State (Zero Cache - Always Live from Drive Excel)
  const [inventory, setInventory] = useState<DepoItem[]>(() => {
    try {
      const stored = localStorage.getItem('ogm_depo_inventory_v5');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed)) {
          // Clean duplicate entries and optimize memory on startup
          const seen = new Set<string>();
          const clean: DepoItem[] = [];
          parsed.forEach((item: any) => {
            if (!item) return;
            const unit = String(item.unit || 'at802').trim().toLowerCase();
            const cat = String(item.category || 'sarf').trim().toLowerCase();
            const desc = String(item.description || item.name || '').trim().toLowerCase();
            const pn = String(item.partNumber || item.pn || '').trim().toLowerCase();
            const sn = String(item.serialAndNotes || item.sn || '').trim().toLowerCase();
            const key = `${unit}_${cat}_${desc}_${pn}_${sn}`;
            if (!seen.has(key)) {
              seen.add(key);
              clean.push(item);
            }
          });
          return clean;
        }
      }
    } catch (e) {
      console.warn("Error parsing ogm_depo_inventory_v5:", e);
    }
    return [];
  });

  // Automatically save inventory changes to localStorage to prevent lost writes & keep clean
  useEffect(() => {
    try {
      if (inventory.length > 0) {
        localStorage.setItem('ogm_depo_inventory_v5', JSON.stringify(inventory));
      }
    } catch (e) {
      console.warn("Failed to persist inventory state:", e);
    }
  }, [inventory]);

  // Envanter bellekte/localStorage'da boşaldıysa diskteki kalıcı at802_sarf_data.json yedeğinden ANINDA kurtar
  useEffect(() => {
    if (inventory.length === 0) {
      fetch('/at802_sarf_data.json')
        .then(res => res.json())
        .then(data => {
          if (Array.isArray(data) && data.length > 0) {
            setInventory(prev => {
              if (prev.length > 0) return prev;
              try {
                localStorage.setItem('ogm_depo_inventory_v5', JSON.stringify(data));
              } catch {}
              return data;
            });
          }
        })
        .catch(() => {});
    }
  }, [inventory.length]);

  const isItemInCatalog = useMemo(() => {
    if (!newItemDesc.trim() && !newItemPn.trim()) return true;
    const d = newItemDesc.trim().toLowerCase();
    const p = newItemPn.trim().toLowerCase();
    return (inventory || []).some((item: any) => {
      const itemD = (item.description || item.name || '').toLowerCase();
      const itemP = (item.partNumber || item.pn || '').toLowerCase();
      return (d && itemD === d) || (p && itemP === p);
    });
  }, [newItemDesc, newItemPn, inventory]);

  const [transactions, setTransactions] = useState<DepoTransaction[]>(initialTransactions || []);
  const [isMovementSubmitting, setIsMovementSubmitting] = useState<boolean>(false);

  // Sync with initialTransactions passed from parent App.tsx (never drop newly added transactions)
  useEffect(() => {
    if (initialTransactions && initialTransactions.length > 0) {
      setTransactions(prev => {
        const getTxId = (t: DepoTransaction) => t.id || `${t.timestamp || t.date}_${t.itemName || t.pn}_${t.type || t.islemTuru}`;
        const map = new Map<string, DepoTransaction>();
        prev.forEach(t => map.set(getTxId(t), t));
        initialTransactions.forEach(t => {
          const k = getTxId(t);
          if (!map.has(k)) {
            map.set(k, t);
          }
        });
        return Array.from(map.values());
      });
    }
  }, [initialTransactions]);

  // Helper to get matching Depo Hareket Geçmişi Sheet Name for given aircraft/unit
  const getDepoSheetNameForUnit = (unitId: string) => {
    const u = (unitId || 'at802').toLowerCase();
    if (u.includes('bell')) return 'DEPO HAREKET GEÇMİŞİ-BELL 429';
    if (u.includes('t70')) return 'DEPO HAREKET GEÇMİŞİ-T-70';
    if (u.includes('360')) return 'DEPO HAREKET GEÇMİŞİ-B-360';
    if (u.includes('650')) return 'DEPO HAREKET GEÇMİŞİ-C-650';
    if (u.includes('hangar')) return 'DEPO HAREKET GEÇMİŞİ-HANGAR';
    return 'DEPO HAREKET GEÇMİŞİ-AT-802';
  };

  // Auto-fetch persistent depo transfers from server (Never delete existing records)
  useEffect(() => {
    if (!isOpen) return;
    const fetchTransfers = async () => {
      try {
        const res = await fetch('/api/get-depo-transfers');
        const data = await res.json();
        if (data && data.status === 'success' && Array.isArray(data.transactions)) {
          setTransactions(prev => {
            const getTxId = (t: any) => t.id || `${t.timestamp || t.date}_${t.itemName || t.name}_${t.type || t.islemTuru}_${t.sn || '-'}`;
            const map = new Map<string, any>();
            // Fresh server transactions first
            data.transactions.forEach((t: any) => {
              map.set(getTxId(t), t);
            });
            // Keep unsynced new session items
            prev.filter(t => (t as any).isNewSessionTx).forEach(t => {
              const k = getTxId(t);
              if (!map.has(k)) {
                map.set(k, t);
              }
            });
            return Array.from(map.values());
          });
        }
      } catch (err) {
        console.warn("Failed to load depo transfers:", err);
      }
    };
    fetchTransfers();
  }, [isOpen]);

  // Double click editing logic (Requirement 5)
  const handleTxDoubleClick = (tx: DepoTransaction) => {
    setPendingTxToEdit(tx);
    setTxPasswordInput('');
    setTxPasswordError('');
    setIsTxPasswordModalOpen(true);
  };

  const handleTxPasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!pendingTxToEdit) return;

    const u = (pendingTxToEdit.unit || 'hangar').toLowerCase();
    const unitConfig = UNITS.find(unit => unit.id === u);
    const requiredPassword = unitConfig ? unitConfig.password : '1839';

    const input = txPasswordInput.trim();
    if (
      input === requiredPassword || 
      input === '1839' || 
      input === '1234' || 
      input === '8902'
    ) {
      setEditTxDate(pendingTxToEdit.date || pendingTxToEdit.timestamp || '');
      setEditTxType(pendingTxToEdit.type || pendingTxToEdit.islemTuru || '');
      setEditTxItemName(pendingTxToEdit.itemName || pendingTxToEdit.itemDesc || pendingTxToEdit.name || '');
      setEditTxPn(pendingTxToEdit.pn || pendingTxToEdit.partNumber || '');
      setEditTxSn(pendingTxToEdit.sn || pendingTxToEdit.serialNumber || '');
      setEditTxQuantity(pendingTxToEdit.quantity !== undefined ? pendingTxToEdit.quantity : 1);
      setEditTxTailNo(pendingTxToEdit.tailNo || pendingTxToEdit.kuyrukKodu || '');
      setEditTxOperator(pendingTxToEdit.operator || pendingTxToEdit.teslimAlan || '');
      setEditTxReceivedBy(pendingTxToEdit.receivedBy || pendingTxToEdit.kabulYapan || '');
      setEditTxLocation(pendingTxToEdit.location || pendingTxToEdit.depoYeri || '');
      setEditTxNotes(pendingTxToEdit.notes || '');

      setSelectedTxForEdit(pendingTxToEdit);
      setIsTxPasswordModalOpen(false);
      setIsTxEditModalOpen(true);
    } else {
      setTxPasswordError('Hatalı Şifre! Lütfen uçak biriminin şifresini veya yetkili şifreyi giriniz.');
    }
  };

  const handleSaveTxEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedTxForEdit) return;

    let finalSn = editTxSn.trim();
    if (editTxType.includes('MUAYENEDEN DÖNEN') || editTxType.includes('MUAYENE GELEN') || editTxNotes.toLowerCase().includes('onarıl') || editTxNotes.toLowerCase().includes('muayene')) {
      if (finalSn && finalSn !== '-' && !finalSn.includes('(onarılmış)') && !finalSn.includes('(ONARILMIŞ)')) {
        finalSn = `${finalSn} (onarılmış)`;
      }
    }

    const updatedTx: DepoTransaction = {
      ...selectedTxForEdit,
      isEdit: true,
      date: editTxDate,
      type: editTxType,
      itemName: editTxItemName,
      pn: editTxPn,
      sn: finalSn,
      quantity: Number(editTxQuantity),
      tailNo: editTxTailNo,
      operator: editTxOperator,
      receivedBy: editTxReceivedBy,
      location: editTxLocation,
      notes: editTxNotes
    };

    // 1. Optimistic UI update - instantly update state, notify user, and close modal
    const updatedList = transactions.map(t => t.id === selectedTxForEdit.id ? updatedTx : t);
    setTransactions(updatedList);
    
    setIsTxEditModalOpen(false);
    setSelectedTxForEdit(null);
    setPendingTxToEdit(null);

    if (showNotification) {
      showNotification('✅ Güncelleme yerel olarak anında kaydedildi. Bulut senkronizasyonu arka planda yapılıyor...');
    }

    // 2. Perform backend API syncing asynchronously in the background
    (async () => {
      try {
        const targetSheet = getDepoSheetNameForUnit(updatedTx.unit || currentUnit);
        
        // Fast background fetch requests
        await fetch('/api/save-depo-transfers', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            spreadsheetId: '17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0',
            sheetName: targetSheet,
            transfers: [updatedTx]
          })
        });

        await fetch('/api/depo-transactions', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            transfers: updatedList
          })
        });

        console.log('Background cloud synchronization completed successfully.');
      } catch (err) {
        console.warn('Background edit sync warning:', err);
      }
    })();
  };

  // Kit Oluşturma Modalı State'leri
  const [isKitModalOpen, setIsKitModalOpen] = useState<boolean>(false);
  const [kitName, setKitName] = useState<string>('');
  const [kitCount, setKitCount] = useState<number>(1);
  const [kitItems, setKitItems] = useState<Array<{
    partNumber: string;
    description: string;
    qtyPerKit: number;
    ankaraMevcut: number;
  }>>([]);
  const [kitSearchQuery, setKitSearchQuery] = useState<string>('');
  const [isKitDropdownOpen, setIsKitDropdownOpen] = useState<boolean>(false);
  const [selectedKitItem, setSelectedKitItem] = useState<DepoItem | null>(null);
  const [kitItemQtyInput, setKitItemQtyInput] = useState<number>(1);

  // Dinamik Çıkış Kontrolü: Seçilen çıkış türüne göre ilgili deponun adını ve mevcut stok sayısını verir
  const getCikisDepotInfo = (typeStr: string, item: DepoItem | null) => {
    if (!item) return { name: 'Ankara Merkez Depo', stock: 0, code: 'ANKARA' };
    const t = (typeStr || '').toUpperCase();
    if (t.includes('KARAİN') || t.includes('KARAIN')) {
      return { name: 'Karain Depo', stock: Number(item.karainMevcut || 0), code: 'KARAİN' };
    }
    if (t.includes('ÇANAKKALE') || t.includes('CANAKKALE')) {
      return { name: 'Çanakkale Depo', stock: Number(item.canakkaleMevcut || 0), code: 'ÇANAKKALE' };
    }
    if (t.includes('MİLAS') || t.includes('MILAS')) {
      return { name: 'Milas Depo', stock: Number(item.milasMevcut || 0), code: 'MİLAS' };
    }
    if (t.includes('BURSA')) {
      return { name: 'Bursa Depo', stock: Number(item.bursaMevcut || 0), code: 'BURSA' };
    }
    return { name: 'Ankara Merkez Depo', stock: Number(item.ankaraMevcut || 0), code: 'ANKARA' };
  };

  // Her bir hava aracı için tek ve standart Sarf Depo / Kimyasal Depo Excel dosya adını belirler
  const getDepoStandardFileName = (unit: string, category: 'sarf' | 'kimyasal' | 'yasam_destek'): string => {
    if (category === 'yasam_destek') {
      return 'at-802_yasam_destek_ekipmanlari.xlsx';
    }
    const u = (unit || 'at802').toLowerCase();
    const isKimyasal = category === 'kimyasal';
    
    if (u.includes('at802') || u.includes('at-802') || u === 'at802' || u === 'at-802' || u === 'all') {
      return isKimyasal ? 'at-802_kimyasal_depo.xlsx' : 'at-802_sarf_ve_parca_deposu.xlsx';
    } else if (u.includes('bell') || u.includes('429')) {
      return isKimyasal ? 'bell-429_kimyasal_depo.xlsx' : 'bell-429_sarf_ve_parca_deposu.xlsx';
    } else if (u.includes('t70') || u.includes('t-70')) {
      return isKimyasal ? 't-70_kimyasal_depo.xlsx' : 't-70_sarf_ve_parca_deposu.xlsx';
    } else if (u.includes('b360') || u.includes('360')) {
      return isKimyasal ? 'b-360_kimyasal_depo.xlsx' : 'b-360_sarf_ve_parca_deposu.xlsx';
    } else if (u.includes('c650') || u.includes('650')) {
      return isKimyasal ? 'c-650_kimyasal_depo.xlsx' : 'c-650_sarf_ve_parca_deposu.xlsx';
    } else if (u.includes('hangar')) {
      return isKimyasal ? 'hangar_kimyasal_depo.xlsx' : 'hangar_sarf_ve_parca_deposu.xlsx';
    }
    return isKimyasal ? 'at-802_kimyasal_depo.xlsx' : 'at-802_sarf_ve_parca_deposu.xlsx';
  };

  // Otomatik Veri Senkronizasyonu: Tüm değişiklikleri anlık olarak Excel Online'a aktaran fonksiyon
  const syncDepoToExcelOnline = async (itemsToSync: DepoItem[], txsToSync: DepoTransaction[]) => {
    try {
      const unitKey = currentUnit !== 'all' ? currentUnit : 'at802';
      const catKey = currentDepoType === 'kimyasal' ? 'kimyasal' : 'sarf';
      const standardFileName = getDepoStandardFileName(unitKey, catKey);

      const wb = XLSX.utils.book_new();

      const sarfRows: (string | number)[][] = [
        [
          "DESCRIPTION", "PART NUMBER", "SERİ NUMBER - MÜKERRER NO - AÇIKLAMA", "LOKASYON NO",
          "GELEN", "TOPLAM STOK", "ANKARA ÇIKAN", "ANKARA MEVCUT",
          "KARAİN TRANSFER", "KARAİN ÇIKAN", "KARAİN MEVCUT",
          "ÇANAKKALE TRANSFER", "ÇANAKKALE ÇIKAN", "ÇANAKKALE MEVCUT",
          "MİLAS TRANSFER", "MİLAS ÇIKAN", "MİLAS MEVCUT",
          "BURSA TRANSFER", "BURSA ÇIKAN", "BURSA MEVCUT",
          "MUAYENE GİDEN", "MUAYENE GELEN", "MUAYENE TOPLAM",
          "RAF ÖMRÜ VAR MI?", "RAF ÖMRÜ BİTİŞ TARİHİ"
        ]
      ];

      itemsToSync.forEach(i => {
        const hasLife = i.hasShelfLife === 'EVET' || i.hasShelfLife === true || String(i.hasShelfLife).toUpperCase().includes('EVET');
        sarfRows.push([
          i.description || i.name || '',
          i.partNumber || i.pn || '',
          i.serialAndNotes || i.sn || '',
          i.lokasyonNo || i.location || '',
          i.gelen || 0,
          i.toplamStok || 0,
          i.ankaraCikan || 0,
          i.ankaraMevcut || 0,
          i.karainTransfer || 0,
          i.karainCikan || 0,
          i.karainMevcut || 0,
          i.canakkaleTransfer || 0,
          i.canakkaleCikan || 0,
          i.canakkaleMevcut || 0,
          i.milasTransfer || 0,
          i.milasCikan || 0,
          i.milasMevcut || 0,
          i.bursaTransfer || 0,
          i.bursaCikan || 0,
          i.bursaMevcut || 0,
          i.muayeneGiden || 0,
          i.muayeneGelen || 0,
          i.muayeneToplam || 0,
          hasLife ? 'EVET' : 'HAYIR',
          hasLife && i.shelfLifeDate && i.shelfLifeDate !== '-' ? i.shelfLifeDate : '-'
        ]);
      });
      const wsSarf = XLSX.utils.aoa_to_sheet(sarfRows);
      XLSX.utils.book_append_sheet(wb, wsSarf, catKey === 'kimyasal' ? "KIMYASAL_DEPO" : "SARF_VE_PARCA_DEPO");

      const transRows: (string | number)[][] = [
        ["MALZEME ADI", "ADET", "TARİH", "İŞLEM TÜRÜ", "SERİAL NUMBER", "KUYRUK KODU", "TESLİM ALAN", "KABUL YAPAN", "DEPO YERİ"]
      ];
      txsToSync.forEach(t => {
        transRows.push([
          t.itemName,
          t.quantity,
          t.date,
          t.type,
          t.sn,
          t.tailNo,
          t.operator,
          t.receivedBy,
          (t.category === 'kimyasal' || (t.location && t.location.toLowerCase().includes('kimya'))) ? 'KİMYASAL DEPO' : (t.location || '-')
        ]);
      });
      const wsTrans = XLSX.utils.aoa_to_sheet(transRows);
      XLSX.utils.book_append_sheet(wb, wsTrans, "DEPO_HAREKET_GECMISI");

      const base64Data = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });

      // Save to localStorage
      try {
        localStorage.setItem('ogm_depo_inventory_v5', JSON.stringify(itemsToSync));
        localStorage.setItem('excel_depo_sarf_data', JSON.stringify(sarfRows));
      } catch (e) {}

      // Upload to Drive & update Excel Online
      await fetch('/api/upload-techizat-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: standardFileName,
          targetKey: `depo_${unitKey}_${catKey}`,
          base64Data: base64Data,
          folderId: DRIVE_FOLDER_ID
        })
      });

      // Save transactions to both endpoints
      const targetSheet = getDepoSheetNameForUnit(currentUnit);
      await fetch('/api/save-depo-transfers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spreadsheetId: '17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0',
          sheetName: targetSheet,
          transfers: txsToSync
        })
      });

      await fetch('/api/depo-transactions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          transfers: txsToSync
        })
      });
    } catch (err) {
      console.warn("Excel online auto-sync warn:", err);
    }
  };

  const [customTransferTypes, setCustomTransferTypes] = useState<Record<string, string[]>>(() => {
    try {
      const saved = localStorage.getItem('ogm_depo_custom_types_v5');
      if (saved) return JSON.parse(saved);
    } catch (e) {}
    return {};
  });

  // Calculate and memoize 23-column metrics directly from the live Excel values
  const computedInventory = useMemo(() => {
    const txMapByPn: Record<string, DepoTransaction[]> = {};
    const txMapByName: Record<string, DepoTransaction[]> = {};

    transactions.forEach(t => {
      // Only map active session transactions as delta modifications
      if (!t.isNewSessionTx) return;

      const pnKey = (t.pn || '').trim().toLowerCase();
      const nameKey = (t.itemName || '').trim().toLowerCase();

      if (pnKey) {
        if (!txMapByPn[pnKey]) txMapByPn[pnKey] = [];
        txMapByPn[pnKey].push(t);
      }
      if (nameKey) {
        if (!txMapByName[nameKey]) txMapByName[nameKey] = [];
        txMapByName[nameKey].push(t);
      }
    });

    // OPTIMIZATION: Pre-index historical transactions for O(1) matching (prevents freezing)
    const normalizeTight = (s: any) => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    const histTxs = transactions.filter(t => !t.isNewSessionTx);
    
    const histMap = new Map<string, DepoTransaction[]>();

    histTxs.forEach(t => {
      const p = normalizeTight(t.pn || t.partNumber || (t as any).partNo || (t as any).pn);
      const d = normalizeTight(t.itemName || t.itemDesc || t.name || (t as any).malzemeAdi);
      
      // Index by PN
      if (p) {
        if (!histMap.has(p)) histMap.set(p, []);
        histMap.get(p)!.push(t);
      }
      // Index by Name (as it often contains PN)
      if (d) {
        if (!histMap.has(d)) histMap.set(d, []);
        histMap.get(d)!.push(t);
      }
    });

    return inventory.map(item => {
      const itemDescTight = normalizeTight(item.description || item.name || '');
      const itemPnTight = normalizeTight(item.partNumber || item.pn || '');

      const relevantTxs = new Set<DepoTransaction>();
      const itemPn = (item.partNumber || item.pn || '').trim().toLowerCase();
      const itemDesc = (item.description || item.name || '').trim().toLowerCase();

      if (itemPn && txMapByPn[itemPn]) {
        txMapByPn[itemPn].forEach(t => relevantTxs.add(t));
      }
      if (itemDesc && txMapByName[itemDesc]) {
        txMapByName[itemDesc].forEach(t => relevantTxs.add(t));
      }

      // Efficiently gather historical transactions using indexed maps
      const histSet = new Set<DepoTransaction>();
      if (itemPnTight && histMap.has(itemPnTight)) {
        histMap.get(itemPnTight)!.forEach(t => histSet.add(t));
      }
      if (itemDescTight && histMap.has(itemDescTight)) {
        histMap.get(itemDescTight)!.forEach(t => histSet.add(t));
      }
      const histForPn = Array.from(histSet);

      // Gelen: If Excel has value, use it. Otherwise count from historical transfers
      let gelen = Number(item.gelen || item.baseGelen || 0);

      // Ankara Cikan, Karain, etc.: Read directly from Excel values as the only authoritative source of historical totals!
      let ankaraCikan = Number(item.ankaraCikan || 0);
      let karainTransfer = Number(item.karainTransfer || 0);
      let karainCikan = Number(item.karainCikan || 0);
      let canakkaleTransfer = Number(item.canakkaleTransfer || 0);
      let canakkaleCikan = Number(item.canakkaleCikan || 0);
      let milasTransfer = Number(item.milasTransfer || 0);
      let milasCikan = Number(item.milasCikan || 0);
      let bursaTransfer = Number(item.bursaTransfer || 0);
      let bursaCikan = Number(item.bursaCikan || 0);
      let muayeneGiden = Number(item.muayeneGiden || 0);
      let muayeneGelen = Number(item.muayeneGelen || 0);
      let muayeneToplam = Number(item.muayeneToplam !== undefined ? item.muayeneToplam : muayeneGelen);

      // Compute stocks
      let ankaraMevcut = Number(item.ankaraMevcut !== undefined ? item.ankaraMevcut : Math.max(0, gelen - ankaraCikan - karainTransfer - canakkaleTransfer - milasTransfer - bursaTransfer - muayeneGiden));
      let karainMevcut = Number(item.karainMevcut !== undefined ? item.karainMevcut : Math.max(0, karainTransfer - karainCikan));
      let canakkaleMevcut = Number(item.canakkaleMevcut !== undefined ? item.canakkaleMevcut : Math.max(0, canakkaleTransfer - canakkaleCikan));
      let milasMevcut = Number(item.milasMevcut !== undefined ? item.milasMevcut : Math.max(0, milasTransfer - milasCikan));
      let bursaMevcut = Number(item.bursaMevcut !== undefined ? item.bursaMevcut : Math.max(0, bursaTransfer - bursaCikan));

      let toplamStok = Number(item.toplamStok !== undefined ? item.toplamStok : (ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut + muayeneToplam));

      // Now apply new active session deltas on top (preventing duplicates completely)
      if (relevantTxs.size > 0) {
        relevantTxs.forEach(t => {
          const qty = Number(t.quantity || 1);
          const type = (t.type || t.islemTuru || '').toUpperCase();

          if (type.includes('MİLAS') && type.includes('TRANSFER')) {
            milasTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            milasMevcut += qty;
          } else if (type.includes('KARAİN') && type.includes('TRANSFER')) {
            karainTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            karainMevcut += qty;
          } else if (type.includes('ÇANAKKALE') && type.includes('TRANSFER')) {
            canakkaleTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            canakkaleMevcut += qty;
          } else if (type.includes('BURSA') && type.includes('TRANSFER')) {
            bursaTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            bursaMevcut += qty;
          } else if (type.includes('ANTALYA TRANSFER') || type.includes('DEPO TRANSFER') || type.includes('DİĞER')) {
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
          } else if (type.includes('ANKARA GİREN') || type.includes('GİRİŞ') || type.includes('GİREN') || (type.includes('ANKARA') && type.includes('FAZLA TESPİT'))) {
            gelen += qty;
            ankaraMevcut += qty;
            toplamStok += qty;
          } else if (type.includes('KARAİN') && (type.includes('ÇIKAN') || type.includes('ÇIKIŞ'))) {
            karainCikan += qty;
            karainMevcut = Math.max(0, karainMevcut - qty);
            toplamStok = Math.max(0, toplamStok - qty);
          } else if (type.includes('ÇANAKKALE') && (type.includes('ÇIKAN') || type.includes('ÇIKIŞ'))) {
            canakkaleCikan += qty;
            canakkaleMevcut = Math.max(0, canakkaleMevcut - qty);
            toplamStok = Math.max(0, toplamStok - qty);
          } else if (type.includes('MİLAS') && (type.includes('ÇIKAN') || type.includes('ÇIKIŞ'))) {
            milasCikan += qty;
            milasMevcut = Math.max(0, milasMevcut - qty);
            toplamStok = Math.max(0, toplamStok - qty);
          } else if (type.includes('BURSA') && (type.includes('ÇIKAN') || type.includes('ÇIKIŞ'))) {
            bursaCikan += qty;
            bursaMevcut = Math.max(0, bursaMevcut - qty);
            toplamStok = Math.max(0, toplamStok - qty);
          } else if (type.includes('ANKARA ÇIKAN') || type.includes('ÇIKAN') || type.includes('ÇIKIŞ')) {
            ankaraCikan += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            toplamStok = Math.max(0, toplamStok - qty);
          } else if (type.includes('MUAYENE GİDEN')) {
            muayeneGiden += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
          } else if (type.includes('MUAYENE GELEN')) {
            muayeneGelen += qty;
            muayeneToplam += qty;
          }
        });

        toplamStok = ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut + muayeneToplam;
      }

      return {
        ...item,
        gelen,
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
        muayeneGiden,
        muayeneGelen,
        muayeneToplam,
        toplamStok
      };
    });
  }, [inventory, transactions]);

  const mobileSuggestions = useMemo(() => {
    if (!mobileSearchQuery.trim()) return [];
    const term = mobileSearchQuery.toLowerCase();
    return computedInventory.filter(i => {
      if (i.category !== mobileCategory) return false;
      const descMatch = (i.description || i.name || '').toLowerCase().includes(term);
      const pnMatch = (i.partNumber || i.pn || '').toLowerCase().includes(term);
      const snMatch = (i.serialAndNotes || i.sn || '').toLowerCase().includes(term);
      const locMatch = (i.lokasyonNo || i.location || '').toLowerCase().includes(term);
      return descMatch || pnMatch || snMatch || locMatch;
    }).slice(0, 100);
  }, [mobileSearchQuery, mobileCategory, computedInventory]);

  // Load fresh online Excel data directly from Google Drive (Zero Cache)
  const loadLiveDriveExcel = (forceNotify = false, forceCategory?: 'sarf' | 'kimyasal' | 'yasam_destek') => {
    const activeTab = currentDepoType;
    
    // Yaşam Destek Senkronizasyonu
    if (activeTab === 'yasam_destek' || forceCategory === 'yasam_destek') {
      setIsLoadingDrive(true);
      fetch('/api/yasam-destek/sync-from-drive', { method: 'POST' })
        .then(res => res.json())
        .then(result => {
          if (result && result.status === 'success') {
            showNotification('✅ Yaşam Destek verileri Google Drive ile senkronize edildi.');
            fetchYasamDestekData();
          } else {
            showNotification('❌ Senkronizasyon hatası: ' + (result.message || 'Bilinmeyen hata'));
          }
        })
        .catch(err => {
          showNotification('❌ Bağlantı hatası: ' + err.message);
        })
        .finally(() => setIsLoadingDrive(false));
      return;
    }

    const targetCat: 'sarf' | 'kimyasal' = (forceCategory as any) || (currentDepoType === 'kimyasal' ? 'kimyasal' : 'sarf');
    if ((currentDepoType === 'transactions') && !forceCategory) {
      return;
    }
    setIsLoadingDrive(true);
    const targetFileName = getDepoStandardFileName(currentUnit === 'all' ? 'at802' : currentUnit, targetCat);

    const uKey = (currentUnit === 'all' ? 'at802' : currentUnit).toLowerCase();
    let targetFileId = '';
    if (uKey.includes('at802') || uKey.includes('at-802') || uKey === 'at802') {
      targetFileId = targetCat === 'kimyasal' ? '1JYPCRpMSaZ8AWt7qEWnXhb4DbUKguDRK' : '11d1HAocT6gCmHaLEhFrP1oXj7U7bZMS7';
    } else if (uKey.includes('bell') || uKey.includes('429')) {
      targetFileId = targetCat === 'kimyasal' ? '' : '1tu8hDWSgIYkGn-7i_gDTC_UpUFUyaOzf';
    }

    fetch('/api/read-excel-from-drive', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileName: targetFileName, fileId: targetFileId })
    })
      .then(res => res.json())
      .then(result => {
        if (result && result.status === 'success' && result.base64) {
          try {
            const binaryString = atob(result.base64);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
              bytes[i] = binaryString.charCodeAt(i);
            }
            const workbook = XLSX.read(bytes, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            unmergeAndFillWorksheet(sheet);
            const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            if (rawRows && rawRows.length > 1) {
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

              const isKimyasalLoad = targetCat === 'kimyasal';

              let descCol = normHeader.findIndex(c => c === 'DESCRIPTION' || c.includes('DESCRIPTION') || c.includes('MALZEME'));
              let pnCol = normHeader.findIndex(c => c.includes('P/N') || c.includes('PART NUMBER') || c.includes('PART'));
              let spainDescCol = normHeader.findIndex(c => c.includes('SPAIN DISCRIPTION') || c.includes('SPAIN DESCRIPTION'));
              let miktarCol = normHeader.findIndex(c => c === 'MIKTAR' || c.includes('MIKTAR QTY') || c === 'BIRIM');
              let sozlesmeCol = normHeader.findIndex(c => c.includes('SOZLESME'));
              let piyasaCol = normHeader.findIndex(c => c.includes('PIYASA'));
              let spainGelenCol = normHeader.findIndex(c => c.includes('SPAIN GELEN'));
              let genelToplamCol = normHeader.findIndex(c => c.includes('GENEL TOPLAM'));
              let toplamStokCol = normHeader.findIndex(c => c.includes('TOPLAM STOK'));
              let ankaraCikanCol = normHeader.findIndex(c => c.includes('ANKARA CIKAN'));
              let ankaraMevcutCol = normHeader.findIndex(c => c.includes('ANKARA MEVCUT'));
              let karainTransferCol = normHeader.findIndex(c => c.includes('KARAIN TRANSFER'));
              let karainCikanCol = normHeader.findIndex(c => c.includes('KARAIN CIKAN'));
              let karainMevcutCol = normHeader.findIndex(c => c.includes('KARAIN MEVCUT'));
              let canakkaleTransferCol = normHeader.findIndex(c => c.includes('CANAKKALE TRANSFER'));
              let canakkaleCikanCol = normHeader.findIndex(c => c.includes('CANAKKALE CIKAN'));
              let canakkaleMevcutCol = normHeader.findIndex(c => c.includes('CANAKKALE MEVCUT'));
              let milasTransferCol = normHeader.findIndex(c => c.includes('MILAS TRANSFER'));
              let milasCikanCol = normHeader.findIndex(c => c.includes('MILAS CIKAN'));
              let milasMevcutCol = normHeader.findIndex(c => c.includes('MILAS MEVCUT'));
              let bursaTransferCol = normHeader.findIndex(c => c.includes('BURSA TRANSFER'));
              let bursaCikanCol = normHeader.findIndex(c => c.includes('BURSA CIKAN'));
              let bursaMevcutCol = normHeader.findIndex(c => c.includes('BURSA MEVCUT'));
              let aciklamaCol = normHeader.findIndex(c => c === 'ACIKLAMA' || c.includes('ACIKLAMA'));
              let locCol = normHeader.findIndex(c => c.includes('LOKASYON') || c.includes('RAF') || c.includes('KONUM'));
              let shelfLifeCol = normHeader.findIndex(c => c.includes('RAF OMRU VAR') || c.includes('OMRUNE TABI'));
              let shelfLifeDateCol = normHeader.findIndex(c => c.includes('RAF OMRU BITIS') || c.includes('SKT'));

              if (isKimyasalLoad) {
                if (descCol < 0) descCol = 1;
                if (pnCol < 0) pnCol = 2;
                if (spainDescCol < 0) spainDescCol = 2;
                if (miktarCol < 0) miktarCol = 3;
                if (sozlesmeCol < 0) sozlesmeCol = 4;
                if (piyasaCol < 0) piyasaCol = 5;
                if (spainGelenCol < 0) spainGelenCol = 6;
                if (genelToplamCol < 0) genelToplamCol = 7;
                if (toplamStokCol < 0) toplamStokCol = 8;
                if (ankaraCikanCol < 0) ankaraCikanCol = 9;
                if (ankaraMevcutCol < 0) ankaraMevcutCol = 10;
                if (karainTransferCol < 0) karainTransferCol = 11;
                if (karainCikanCol < 0) karainCikanCol = 12;
                if (karainMevcutCol < 0) karainMevcutCol = 13;
                if (canakkaleTransferCol < 0) canakkaleTransferCol = 14;
                if (canakkaleCikanCol < 0) canakkaleCikanCol = 15;
                if (canakkaleMevcutCol < 0) canakkaleMevcutCol = 16;
                if (milasTransferCol < 0) milasTransferCol = 17;
                if (milasCikanCol < 0) milasCikanCol = 18;
                if (milasMevcutCol < 0) milasMevcutCol = 19;
                if (bursaTransferCol < 0) bursaTransferCol = 20;
                if (bursaCikanCol < 0) bursaCikanCol = 21;
                if (bursaMevcutCol < 0) bursaMevcutCol = 22;
                if (aciklamaCol < 0) aciklamaCol = 23;
              }

              const parsedItems: DepoItem[] = [];
              const unitKey = currentUnit === 'all' ? 'at802' : currentUnit;
              for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
                const row = rawRows[r];
                if (!row || row.length === 0) continue;
                const desc = String(row[descCol >= 0 ? descCol : 1] || row[0] || '').trim();
                if (!desc || isHeaderLikeRow(row) || desc.toUpperCase() === 'DESCRIPTION' || desc.toUpperCase() === 'MALZEME ADI') continue;

                const pn = String(row[pnCol >= 0 ? pnCol : 2] !== undefined && row[pnCol >= 0 ? pnCol : 2] !== null ? row[pnCol >= 0 ? pnCol : 2] : '-').trim() || '-';
                const spainDesc = spainDescCol >= 0 ? String(row[spainDescCol] !== undefined ? row[spainDescCol] : pn).trim() : pn;

                const rawMiktarVal = String(row[miktarCol >= 0 ? miktarCol : 3] !== undefined ? row[miktarCol >= 0 ? miktarCol : 3] : 'Adet').trim();
                const miktarBirim = rawMiktarVal || 'Adet';

                const sozlesmeToplamiVal = row[sozlesmeCol >= 0 ? sozlesmeCol : 4] !== undefined ? row[sozlesmeCol >= 0 ? sozlesmeCol : 4] : '0';
                const sozlesmeToplami = String(sozlesmeToplamiVal !== undefined ? sozlesmeToplamiVal : '0').trim();

                const piyasaGelen = Number(row[piyasaCol >= 0 ? piyasaCol : 5]) || 0;
                const spainGelen = Number(row[spainGelenCol >= 0 ? spainGelenCol : 6]) || 0;
                const genelToplamVal = Number(row[genelToplamCol >= 0 ? genelToplamCol : 7]);
                const genelToplam = !isNaN(genelToplamVal) && genelToplamVal > 0 ? genelToplamVal : (piyasaGelen + spainGelen);

                const ankaraCikan = Number(row[ankaraCikanCol >= 0 ? ankaraCikanCol : 9]) || 0;
                const ankaraMevcut = Number(row[ankaraMevcutCol >= 0 ? ankaraMevcutCol : 10]) || 0;

                const karainTransfer = Number(row[karainTransferCol >= 0 ? karainTransferCol : 11]) || 0;
                const karainCikan = Number(row[karainCikanCol >= 0 ? karainCikanCol : 12]) || 0;
                const karainMevcut = Number(row[karainMevcutCol >= 0 ? karainMevcutCol : 13]) || 0;

                const canakkaleTransfer = Number(row[canakkaleTransferCol >= 0 ? canakkaleTransferCol : 14]) || 0;
                const canakkaleCikan = Number(row[canakkaleCikanCol >= 0 ? canakkaleCikanCol : 15]) || 0;
                const canakkaleMevcut = Number(row[canakkaleMevcutCol >= 0 ? canakkaleMevcutCol : 16]) || 0;

                const milasTransfer = Number(row[milasTransferCol >= 0 ? milasTransferCol : 17]) || 0;
                const milasCikan = Number(row[milasCikanCol >= 0 ? milasCikanCol : 18]) || 0;
                const milasMevcut = Number(row[milasMevcutCol >= 0 ? milasMevcutCol : 19]) || 0;

                const bursaTransfer = Number(row[bursaTransferCol >= 0 ? bursaTransferCol : 20]) || 0;
                const bursaCikan = Number(row[bursaCikanCol >= 0 ? bursaCikanCol : 21]) || 0;
                const bursaMevcut = Number(row[bursaMevcutCol >= 0 ? bursaMevcutCol : 22]) || 0;

                const aciklama = String(row[aciklamaCol >= 0 ? aciklamaCol : 23] !== undefined ? row[aciklamaCol >= 0 ? aciklamaCol : 23] : '-').trim() || '-';
                const loc = String(row[locCol >= 0 ? locCol : 3] !== undefined ? row[locCol >= 0 ? locCol : 3] : 'KİMYASAL DEPO').trim() || 'KİMYASAL DEPO';

                const computedToplam = (toplamStokCol >= 0 && Number(row[toplamStokCol]) > 0)
                  ? Number(row[toplamStokCol])
                  : ((ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut) || genelToplam);

                const durumuCol = headerRow.findIndex(c => c === 'DURUMU' || c.includes('DURUM'));
                const tedarikciCol = headerRow.findIndex(c => c === 'TEDARİKÇİ' || c === 'TEDARIKCI' || c.includes('TEDARİK') || c.includes('TEDARIK'));
                const durumu = durumuCol >= 0 ? String(row[durumuCol] || 'FAAL').trim() : 'FAAL';
                const tedarikci = tedarikciCol >= 0 ? String(row[tedarikciCol] || '-').trim() : '-';
                const hasShelfLife = shelfLifeCol >= 0 ? (String(row[shelfLifeCol] || '').toUpperCase().includes('EVET') ? 'EVET' : 'HAYIR') : 'HAYIR';
                const shelfLifeDate = shelfLifeDateCol >= 0 ? String(row[shelfLifeDateCol] || '-').trim() : '-';

                parsedItems.push({
                  unit: unitKey,
                  category: targetCat,
                  description: desc,
                  name: desc,
                  partNumber: pn,
                  pn: pn,
                  spainDescription: spainDesc,
                  miktarQty: miktarBirim,
                  sozlesmeToplami: sozlesmeToplami,
                  piyasa: piyasaGelen,
                  spainGelen: spainGelen,
                  genelToplam: genelToplam,
                  toplamStok: computedToplam,
                  ankaraCikan: ankaraCikan,
                  ankaraMevcut: ankaraMevcut,
                  karainTransfer: karainTransfer,
                  karainCikan: karainCikan,
                  karainMevcut: karainMevcut,
                  canakkaleTransfer: canakkaleTransfer,
                  canakkaleCikan: canakkaleCikan,
                  canakkaleMevcut: canakkaleMevcut,
                  milasTransfer: milasTransfer,
                  milasCikan: milasCikan,
                  milasMevcut: milasMevcut,
                  bursaTransfer: bursaTransfer,
                  bursaCikan: bursaCikan,
                  bursaMevcut: bursaMevcut,
                  serialAndNotes: aciklama,
                  sn: aciklama,
                  lokasyonNo: loc,
                  location: loc,
                  baseGelen: genelToplam,
                  gelen: genelToplam,
                  hasShelfLife: hasShelfLife,
                  shelfLifeDate: shelfLifeDate,
                  durumu: durumu,
                  tedarikci: tedarikci
                });
              }

                            if (parsedItems.length > 0) {
                setInventory(prev => {
                  const existingMap = new Map<string, DepoItem>();
                  prev.filter(i => i.unit === unitKey && i.category === targetCat).forEach(i => {
                    const key = `${(i.partNumber || i.pn || '').trim().toLowerCase()}___${(i.serialAndNotes || i.sn || '').trim().toLowerCase()}`;
                    existingMap.set(key, i);
                  });

                  const mergedItems = parsedItems.map(newItem => {
                    const key = `${(newItem.partNumber || newItem.pn || '').trim().toLowerCase()}___${(newItem.serialAndNotes || newItem.sn || '').trim().toLowerCase()}`;
                    const prevItem = existingMap.get(key);
                    if (!prevItem) return newItem;

                    // Preserve user-modified attributes if Excel file did not specify them
                    const hasShelfLife = (newItem.hasShelfLife && newItem.hasShelfLife !== 'HAYIR')
                      ? newItem.hasShelfLife
                      : (prevItem.hasShelfLife || 'HAYIR');

                    const shelfLifeDate = (newItem.shelfLifeDate && newItem.shelfLifeDate !== '-')
                      ? newItem.shelfLifeDate
                      : (prevItem.shelfLifeDate || '-');

                    const durumu = (newItem.durumu && newItem.durumu !== 'FAAL')
                      ? newItem.durumu
                      : (prevItem.durumu || 'FAAL');

                    const tedarikci = (newItem.tedarikci && newItem.tedarikci !== '-')
                      ? newItem.tedarikci
                      : (prevItem.tedarikci || '-');

                    return {
                      ...newItem,
                      hasShelfLife,
                      shelfLifeDate,
                      durumu,
                      tedarikci
                    };
                  });

                  const filtered = prev.filter(i => !(i.unit === unitKey && i.category === targetCat));
                  return [...filtered, ...mergedItems];
                });
                if (forceNotify) {
                  showNotification(`✅ Google Drive'dan ${parsedItems.length} malzeme başarıyla aktarıldı.`);
                }
              } else {
                if (forceNotify) {
                  showNotification(`⚠️ Google Drive'da canlı dosya bulunamadı veya silinmiş. Mevcut önbellekteki veriler korundu.`);
                }
              }
            }
          } catch (parseErr) {
            console.warn("Drive Excel parse error:", parseErr);
          }
        }
      })
      .catch((err) => {
        console.warn("Drive read error:", err);
        if (forceNotify) {
          showNotification("⚠️ Google Drive dosyasında erişim hatası. Mevcut önbellek verisi korundu.");
        }
      })
      .finally(() => {
        setIsLoadingDrive(false);
      });
  };

  // Always load fresh online Excel data directly from Google Drive when modal is opened or tabs switch
  useEffect(() => {
    if (!isOpen) return;

    // Her zaman transferleri çek (Detay modalı için gerekli) - Mevcut kayıtları ASLA silme
    fetch('/api/depo-transactions')
      .then(res => res.json())
      .then(result => {
        if (result && result.status === 'success' && Array.isArray(result.transfers)) {
          setTransactions(prev => {
            const getTxId = (t: any) => t.id || `${t.timestamp || t.date}_${t.itemName || t.name}_${t.type || t.islemTuru}_${t.sn || '-'}`;
            const map = new Map<string, any>();
            // Fresh server transfers first
            result.transfers.forEach((t: any) => {
              map.set(getTxId(t), t);
            });
            // Keep unsynced new session items
            prev.filter(t => (t as any).isNewSessionTx).forEach(t => {
              const k = getTxId(t);
              if (!map.has(k)) {
                map.set(k, t);
              }
            });
            return Array.from(map.values());
          });
        }
      })
      .catch(() => {});

    // Yaşam Destek kayıtlarını canlı çek
    fetchYasamDestekData();

    if (currentDepoType !== 'transactions' && currentDepoType !== 'yasam_destek') {
      const unitKey = currentUnit || 'at802';
      const loadedKey = `_ogm_depo_loaded_${unitKey}_${currentDepoType}`;
      if (!(window as any)[loadedKey] || inventory.length === 0) {
        loadLiveDriveExcel(false, currentDepoType);
        (window as any)[loadedKey] = true;
      }
    }
  }, [isOpen, currentUnit, currentDepoType]);

  // Transactions loaded strictly from Google Drive sheet (Zero Cache)

  // Pagination states for high-speed rendering
  const [sarfPage, setSarfPage] = useState<number>(1);
  const [sarfPageSize, setSarfPageSize] = useState<number>(50);
  const [kimyasalPage, setKimyasalPage] = useState<number>(1);
  const [kimyasalPageSize, setKimyasalPageSize] = useState<number>(50);
  const [txPage, setTxPage] = useState<number>(1);
  const [txPageSize, setTxPageSize] = useState<number>(50);

  useEffect(() => {
    setSarfPage(1);
    setKimyasalPage(1);
    setTxPage(1);
  }, [searchQuery, currentUnit, filterRegion, currentDepoType]);

  // Inventory index map for O(1) lookups instead of expensive findIndex in loops
  const inventoryIndexMap = useMemo(() => {
    const map = new Map<string, number>();
    inventory.forEach((item, idx) => {
      const pn = (item.partNumber || item.pn || '').trim().toLowerCase();
      const desc = (item.description || item.name || '').trim().toLowerCase();
      map.set(`${pn}___${desc}`, idx);
    });
    return map;
  }, [inventory]);

  // Unique Lokasyon No listesi (Excel filtre dropdown ve arama için)
  const uniqueLocations = useMemo(() => {
    const locSet = new Set<string>();
    computedInventory.forEach(i => {
      const loc = (i.lokasyonNo || i.location || '').trim();
      if (loc && loc !== '-') locSet.add(loc);
    });
    return Array.from(locSet).sort();
  }, [computedInventory]);

  // Excel Filtresi Dropdown içindeki arama filtresi
  const filteredLocationsForDropdown = useMemo(() => {
    if (!locationSearchTerm.trim()) return uniqueLocations;
    const q = locationSearchTerm.toLowerCase().trim();
    return uniqueLocations.filter(loc => loc.toLowerCase().includes(q));
  }, [uniqueLocations, locationSearchTerm]);

  // Helper: Herhangi bir bölge için geçmişte transfer veya hareket kaydı var mı kontrolü (Tüm Depolar İçin)
  const hadRegionHistory = (item: DepoItem, regName: string): boolean => {
    const regUpper = regName.toUpperCase().trim();
    const itemPn = (item.partNumber || item.pn || '').trim().toLowerCase();
    const itemDesc = (item.description || item.name || '').trim().toLowerCase();

    if (regUpper.includes('ANKARA')) {
      if ((Number(item.ankaraCikan) || 0) > 0 || (Number(item.gelen) || 0) > 0) return true;
    } else if (regUpper.includes('KARAİN') || regUpper.includes('KARAIN')) {
      if ((Number(item.karainTransfer) || 0) > 0 || (Number(item.karainCikan) || 0) > 0) return true;
    } else if (regUpper.includes('ÇANAKKALE') || regUpper.includes('CANAKKALE')) {
      if ((Number(item.canakkaleTransfer) || 0) > 0 || (Number(item.canakkaleCikan) || 0) > 0) return true;
    } else if (regUpper.includes('MİLAS') || regUpper.includes('MILAS')) {
      if ((Number(item.milasTransfer) || 0) > 0 || (Number(item.milasCikan) || 0) > 0) return true;
    } else if (regUpper.includes('BURSA')) {
      if ((Number(item.bursaTransfer) || 0) > 0 || (Number(item.bursaCikan) || 0) > 0) return true;
    }

    const itemLoc = (item.lokasyonNo || item.location || '').toUpperCase();
    if (itemLoc.includes(regUpper)) return true;

    return transactions.some(tx => {
      const txPn = (tx.pn || tx.partNumber || '').trim().toLowerCase();
      const txDesc = (tx.itemName || tx.itemDesc || tx.name || '').trim().toLowerCase();
      const isMatch = (itemPn && txPn && itemPn === txPn) || (itemDesc && txDesc && itemDesc === txDesc);
      if (!isMatch) return false;
      const typeStr = (tx.type || tx.islemTuru || '').toUpperCase();
      const locStr = (tx.location || tx.depoYeri || '').toUpperCase();
      const notesStr = (tx.notes || (tx as any).aciklama || '').toUpperCase();
      return typeStr.includes(regUpper) || locStr.includes(regUpper) || notesStr.includes(regUpper);
    });
  };

  // Memoized Filtered views
  const filteredSarfRows = useMemo(() => {
    return computedInventory.filter(item => {
      if (currentUnit !== 'all' && item.unit !== currentUnit) return false;
      if (item.category === 'kimyasal') return false;

      // Lokasyon No Çoklu Filtresi (Excel Stili Checkbox Listesi Görseldeki Gibi)
      if (selectedLocations.length > 0) {
        const itemLoc = (item.lokasyonNo || item.location || '').trim();
        if (!selectedLocations.includes(itemLoc)) return false;
      }

      if (filterRegion !== 'ALL') {
        const reg = filterRegion.toUpperCase().trim();
        let currentStock = 0;
        if (reg === 'ANKARA') currentStock = item.ankaraMevcut ?? 0;
        else if (reg === 'KARAİN' || reg === 'KARAIN') currentStock = item.karainMevcut ?? 0;
        else if (reg === 'ÇANAKKALE' || reg === 'CANAKKALE') currentStock = item.canakkaleMevcut ?? 0;
        else if (reg === 'MİLAS' || reg === 'MILAS') currentStock = item.milasMevcut ?? 0;
        else if (reg === 'BURSA') currentStock = item.bursaMevcut ?? 0;
        else {
          const loc = (item.lokasyonNo || item.location || '').toUpperCase();
          if (!loc.includes(reg)) return false;
          currentStock = item.toplamStok ?? 0;
        }

        if (currentStock <= 0) {
          // Kullanıcı kuralı (Tüm depolar için geçerli): Eğer mevcut 0 ise ama önceden transfer geçmişi 0'dan farklıysa filtreye dahil et!
          if (!hadRegionHistory(item, reg)) return false;
        }
      }

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const desc = (item.description || item.name || '').toLowerCase();
        const pn = (item.partNumber || item.pn || '').toLowerCase();
        const sn = (item.serialAndNotes || item.sn || '').toLowerCase();
        const loc = (item.lokasyonNo || item.location || '').toLowerCase();
        return desc.includes(q) || pn.includes(q) || sn.includes(q) || loc.includes(q);
      }
      return true;
    });
  }, [computedInventory, currentUnit, filterRegion, selectedLocations, searchQuery, transactions]);

  const filteredKimyasalRows = useMemo(() => {
    return computedInventory.filter(item => {
      if (currentUnit !== 'all' && item.unit !== currentUnit) return false;
      if (item.category !== 'kimyasal') return false;

      // Lokasyon No Çoklu Filtresi (Excel Stili Checkbox Listesi Görseldeki Gibi)
      if (selectedLocations.length > 0) {
        const itemLoc = (item.lokasyonNo || item.location || '').trim();
        if (!selectedLocations.includes(itemLoc)) return false;
      }

      if (filterRegion !== 'ALL') {
        const reg = filterRegion.toUpperCase().trim();
        let currentStock = 0;
        if (reg === 'ANKARA') currentStock = item.ankaraMevcut ?? 0;
        else if (reg === 'KARAİN' || reg === 'KARAIN') currentStock = item.karainMevcut ?? 0;
        else if (reg === 'ÇANAKKALE' || reg === 'CANAKKALE') currentStock = item.canakkaleMevcut ?? 0;
        else if (reg === 'MİLAS' || reg === 'MILAS') currentStock = item.milasMevcut ?? 0;
        else if (reg === 'BURSA') currentStock = item.bursaMevcut ?? 0;
        else {
          const loc = (item.lokasyonNo || item.location || '').toUpperCase();
          if (!loc.includes(reg)) return false;
          currentStock = item.toplamStok ?? 0;
        }

        if (currentStock <= 0) {
          // Kullanıcı kuralı (Tüm depolar için geçerli): Eğer mevcut 0 ise ama önceden transfer geçmişi 0'dan farklıysa filtreye dahil et!
          if (!hadRegionHistory(item, reg)) return false;
        }
      }

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const desc = (item.description || item.name || '').toLowerCase();
        const pn = (item.partNumber || item.pn || '').toLowerCase();
        const sn = (item.serialAndNotes || item.sn || '').toLowerCase();
        const loc = (item.lokasyonNo || item.location || '').toLowerCase();
        return desc.includes(q) || pn.includes(q) || sn.includes(q) || loc.includes(q);
      }
      return true;
    });
  }, [computedInventory, currentUnit, filterRegion, selectedLocations, searchQuery, transactions]);

  const paginatedSarfRows = useMemo(() => {
    if (sarfPageSize <= 0) return filteredSarfRows;
    const start = (sarfPage - 1) * sarfPageSize;
    return filteredSarfRows.slice(start, start + sarfPageSize);
  }, [filteredSarfRows, sarfPage, sarfPageSize]);

  const paginatedKimyasalRows = useMemo(() => {
    if (kimyasalPageSize <= 0) return filteredKimyasalRows;
    const start = (kimyasalPage - 1) * kimyasalPageSize;
    return filteredKimyasalRows.slice(start, start + kimyasalPageSize);
  }, [filteredKimyasalRows, kimyasalPage, kimyasalPageSize]);

  const filteredTransactions = useMemo(() => {
    let list = transactions;

    if (currentUnit && currentUnit !== 'all') {
      const u = currentUnit.toLowerCase().replace(/[^a-z0-9]/g, '');
      const isAt802 = u.startsWith('at802');
      list = list.filter(t => {
        const txUnit = (t.unit || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        const txSheet = (t.sheetName || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        const txTail = (t.tailNo || t.kuyrukKodu || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        
        if (isAt802) {
          if (txUnit.startsWith('at802') || txSheet.includes('at802') || txSheet.includes('at-802') || txTail.includes('at802') || txTail.includes('or20') || (!txUnit && !txSheet.includes('bell') && !txSheet.includes('t70') && !txSheet.includes('360') && !txSheet.includes('650') && !txSheet.includes('hangar'))) {
            return true;
          }
          return false;
        }

        if (txUnit && txUnit === u) return true;
        if (txSheet && txSheet.includes(u)) return true;
        if (txTail && txTail.includes(u)) return true;
        return false;
      });
    }

    // 1. İşlem Türü Filtresi (Canlı & Akıllı Eşleşme)
    if (txFilterType && txFilterType !== 'ALL') {
      const ft = txFilterType.toUpperCase().trim();
      list = list.filter(t => {
        const rawType = (t.type || t.islemTuru || '').toUpperCase().trim();
        if (!rawType) return false;
        if (rawType === ft || rawType.includes(ft)) return true;

        // Akıllı Bölge Eşleşmeleri:
        // KARAİN TRANSFER seçildiğinde "KARAİN SAYIM TESPİT TRANSFER (FAZLA TESPİT)" de dahil edilir
        if (ft === 'KARAİN TRANSFER') {
          return rawType.includes('KARAİN') && rawType.includes('TRANSFER');
        }
        // KARAİN ÇIKAN seçildiğinde "KARAİN SAYIM TESPİT ÇIKAN (EKSIK TESPİT)" de dahil edilir
        if (ft === 'KARAİN ÇIKAN') {
          return rawType.includes('KARAİN') && (rawType.includes('ÇIKAN') || rawType.includes('ÇIKIŞ'));
        }
        // ÇANAKKALE
        if (ft === 'ÇANAKKALE TRANSFER') {
          return rawType.includes('ÇANAKKALE') && rawType.includes('TRANSFER');
        }
        if (ft === 'ÇANAKKALE ÇIKAN') {
          return rawType.includes('ÇANAKKALE') && (rawType.includes('ÇIKAN') || rawType.includes('ÇIKIŞ'));
        }
        // MİLAS
        if (ft === 'MİLAS TRANSFER') {
          return rawType.includes('MİLAS') && rawType.includes('TRANSFER');
        }
        if (ft === 'MİLAS ÇIKAN') {
          return rawType.includes('MİLAS') && (rawType.includes('ÇIKAN') || rawType.includes('ÇIKIŞ'));
        }
        // BURSA
        if (ft === 'BURSA TRANSFER') {
          return rawType.includes('BURSA') && rawType.includes('TRANSFER');
        }
        if (ft === 'BURSA ÇIKAN') {
          return rawType.includes('BURSA') && (rawType.includes('ÇIKAN') || rawType.includes('ÇIKIŞ'));
        }
        // ANKARA GİREN / GELEN
        if (ft === 'ANKARA GİREN' || ft === 'GİREN') {
          return rawType.includes('ANKARA GİREN') || rawType.includes('GİRİŞ') || rawType.includes('GİREN') || (rawType.includes('ANKARA') && rawType.includes('FAZLA TESPİT'));
        }
        // ANKARA ÇIKAN
        if (ft === 'ANKARA ÇIKAN') {
          return rawType.includes('ANKARA ÇIKAN') || (rawType.includes('ANKARA') && (rawType.includes('ÇIKAN') || rawType.includes('ÇIKIŞ') || rawType.includes('EKSIK TESPİT')));
        }
        // Genel Transferler
        if (ft === 'TRANSFER') {
          return rawType.includes('TRANSFER');
        }
        // Genel Çıkanlar
        if (ft === 'ÇIKAN') {
          return rawType.includes('ÇIKAN') || rawType.includes('ÇIKIŞ');
        }
        // Sayım Farkları / Tespitleri
        if (ft === 'SAYIM FARKI') {
          return rawType.includes('SAYIM');
        }

        return false;
      });
    }

    // 2. Kuyruk No / Araç Filtresi (Kuyruk No sadece ÇIKAN işlem türünde geçerlidir)
    if (txFilterTail && txFilterTail !== 'ALL') {
      const fTail = txFilterTail.toUpperCase().replace(/\s+/, '-');
      list = list.filter(t => {
        const typeStr = (t.type || t.islemTuru || '').toUpperCase();
        const isCikis = typeStr.includes('ÇIKAN') || typeStr.includes('ÇIKIŞ') || typeStr.includes('SARF');
        if (!isCikis) return false;
        const rawTail = (t.tailNo || t.kuyrukKodu || '').toUpperCase().replace(/\s+/, '-');
        return rawTail.includes(fTail);
      });
    }

    // 3. İki Tarih Arası (Tarih Aralığı) Filtresi
    if (txFilterStartDate || txFilterEndDate) {
      list = list.filter(t => {
        const rawDate = String(t.date || t.timestamp || '').trim();
        if (!rawDate) return true;

        let txDateObj: Date | null = null;
        if (rawDate.includes('.')) {
          const parts = rawDate.split(' ')[0].split('.');
          if (parts.length === 3) {
            txDateObj = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
          }
        } else if (rawDate.includes('-')) {
          const parts = rawDate.split('T')[0].split('-');
          if (parts.length === 3) {
            txDateObj = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
          }
        }

        if (!txDateObj || isNaN(txDateObj.getTime())) return true;

        if (txFilterStartDate) {
          const startObj = new Date(txFilterStartDate);
          startObj.setHours(0, 0, 0, 0);
          if (txDateObj < startObj) return false;
        }

        if (txFilterEndDate) {
          const endObj = new Date(txFilterEndDate);
          endObj.setHours(23, 59, 59, 999);
          if (txDateObj > endObj) return false;
        }

        return true;
      });
    }

    // Apply Manual Query filter (Optional)
    if (txSearchActiveQuery.trim()) {
      const q = txSearchActiveQuery.toLowerCase().trim();
      list = list.filter(t => {
        const desc = (t.itemName || t.itemDesc || t.name || '').toLowerCase();
        const pn = (t.pn || t.partNumber || '').toLowerCase();
        const sn = (t.sn || t.serialNumber || '').toLowerCase();
        const typeStr = (t.type || t.islemTuru || '').toLowerCase();
        const tail = (t.tailNo || t.kuyrukKodu || '').toLowerCase();
        const loc = (t.location || t.depoYeri || '').toLowerCase();
        return desc.includes(q) || pn.includes(q) || sn.includes(q) || typeStr.includes(q) || tail.includes(q) || loc.includes(q);
      });
    }
    
    return list;
  }, [transactions, currentUnit, txSearchActiveQuery, txFilterType, txFilterTail, txFilterStartDate, txFilterEndDate, isTxSearchPerformed]);

  const paginatedTransactions = useMemo(() => {
    if (txPageSize <= 0) return filteredTransactions;
    const start = (txPage - 1) * txPageSize;
    return filteredTransactions.slice(start, start + txPageSize);
  }, [filteredTransactions, txPage, txPageSize]);

  const sarfCount = useMemo(() => computedInventory.filter(i => (currentUnit === 'all' || i.unit === currentUnit) && i.category === 'sarf').length, [computedInventory, currentUnit]);
  const kimyasalCount = useMemo(() => computedInventory.filter(i => (currentUnit === 'all' || i.unit === currentUnit) && i.category === 'kimyasal').length, [computedInventory, currentUnit]);
  const yasamDestekCount = useMemo(() => yasamDestekRecords.filter(r => r.durum === 'GÖREVDE').length, [yasamDestekRecords]);
  const txCount = filteredTransactions.length;

  if (!isOpen) return null;

  if (isMobile) {
    return (
      <div className="fixed inset-0 z-[10000] bg-slate-950 text-slate-100 flex flex-col font-sans overflow-hidden">
        {/* Mobile Header */}
        <header className="bg-slate-900 border-b border-slate-800 p-4 flex items-center justify-between sticky top-0 z-50 select-none">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-800 flex items-center justify-center shadow-lg border border-emerald-500/20">
              <Boxes className="w-4.5 h-4.5 text-white" />
            </div>
            <div>
              <h1 className="text-xs font-black tracking-wider text-white uppercase leading-none">OGM DEPO MOBİL</h1>
              <p className="text-[9px] text-slate-400 mt-0.5">Hızlı Malzeme &amp; Lokasyon Sorgu</p>
            </div>
          </div>
          <button 
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-full bg-slate-800 text-slate-400 hover:text-white hover:bg-slate-700 transition cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </header>

        {/* Mobile Body */}
        <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-5 max-w-md mx-auto w-full">
          {/* Step 1: Depo Category Selection */}
          <div className="flex flex-col gap-2">
            <label className="text-[10px] font-black tracking-wider text-slate-400 uppercase">
              1. İŞLEM YAPILACAK DEPO SEÇİNİZ
            </label>
            <div className="grid grid-cols-2 gap-3">
              <button
                type="button"
                onClick={() => {
                  setMobileCategory('sarf');
                  setSelectedMobileItem(null);
                  setMobileSearchQuery('');
                  setShowMobileSuggestions(false);
                }}
                className={`p-3.5 rounded-2xl flex flex-col gap-1.5 text-left border transition-all cursor-pointer ${
                  mobileCategory === 'sarf'
                    ? 'bg-gradient-to-b from-amber-600 to-amber-700 text-white border-amber-400 shadow-lg shadow-amber-950/40'
                    : 'bg-slate-900 hover:bg-slate-850 text-slate-400 border-slate-800 hover:border-slate-750'
                }`}
              >
                <Package className={`w-5 h-5 ${mobileCategory === 'sarf' ? 'text-white' : 'text-amber-500'}`} />
                <div>
                  <span className="text-xs font-black tracking-tight block uppercase leading-tight">SARF DEPOSU</span>
                  <span className={`text-[9px] block leading-none mt-0.5 ${mobileCategory === 'sarf' ? 'text-amber-100' : 'text-slate-500'}`}>
                    Yedek Parça &amp; Sarf Malzeme
                  </span>
                </div>
              </button>

              <button
                type="button"
                onClick={() => {
                  setMobileCategory('kimyasal');
                  setSelectedMobileItem(null);
                  setMobileSearchQuery('');
                  setShowMobileSuggestions(false);
                }}
                className={`p-3.5 rounded-2xl flex flex-col gap-1.5 text-left border transition-all cursor-pointer ${
                  mobileCategory === 'kimyasal'
                    ? 'bg-gradient-to-b from-emerald-650 to-emerald-750 text-white border-emerald-400 shadow-lg shadow-emerald-950/40'
                    : 'bg-slate-900 hover:bg-slate-850 text-slate-400 border-slate-800 hover:border-slate-750'
                }`}
              >
                <FlaskConical className={`w-5 h-5 ${mobileCategory === 'kimyasal' ? 'text-white' : 'text-emerald-400'}`} />
                <div>
                  <span className="text-xs font-black tracking-tight block uppercase leading-tight">KİMYASAL DEPOSU</span>
                  <span className={`text-[9px] block leading-none mt-0.5 ${mobileCategory === 'kimyasal' ? 'text-emerald-100' : 'text-slate-500'}`}>
                    Yağ / Boya / Kimyasal
                  </span>
                </div>
              </button>
            </div>
          </div>

          {/* Step 2: Search Input with Autocomplete */}
          <div className="flex flex-col gap-2 relative">
            <label className="text-[10px] font-black tracking-wider text-slate-400 uppercase">
              2. MALZEME ARA / SEÇ (OTOMATİK TAMAMLAMA)
            </label>
            <div className="relative flex items-center">
              <div className="absolute left-3.5 text-slate-400">
                <Search className="w-4 h-4" />
              </div>
              <input
                type="text"
                value={mobileSearchQuery}
                placeholder="Malzeme adı, P/N, S/N yazın..."
                onChange={(e) => {
                  setMobileSearchQuery(e.target.value);
                  setShowMobileSuggestions(true);
                  if (!e.target.value) {
                    setSelectedMobileItem(null);
                  }
                }}
                onFocus={() => setShowMobileSuggestions(true)}
                className="w-full pl-10 pr-10 py-3 rounded-xl bg-slate-900 border-2 border-slate-800 hover:border-slate-750 focus:border-emerald-500 text-sm font-medium text-white placeholder-slate-500 transition-colors focus:outline-none"
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

            {/* Autocomplete Suggestions */}
            {showMobileSuggestions && mobileSuggestions.length > 0 && (
              <div className="absolute top-[72px] left-0 w-full bg-slate-900 border border-slate-800 rounded-xl shadow-2xl max-h-60 overflow-y-auto z-50 divide-y divide-slate-850">
                {mobileSuggestions.map((item, idx) => {
                  const name = item.description || item.name || '';
                  const pn = item.partNumber || item.pn || '-';
                  const sn = item.serialAndNotes || item.sn || '-';
                  const loc = item.lokasyonNo || item.location || '-';
                  return (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => {
                        setSelectedMobileItem(item);
                        setMobileSearchQuery(name);
                        setShowMobileSuggestions(false);
                      }}
                      className="w-full px-4 py-3 text-left hover:bg-slate-850 transition-colors flex flex-col gap-1 active:bg-slate-800 cursor-pointer"
                    >
                      <span className="text-xs font-black text-slate-100 line-clamp-1">{name}</span>
                      <div className="flex flex-wrap items-center gap-x-2.5 gap-y-0.5 text-[9px] text-slate-400 font-medium">
                        <span className="font-mono bg-slate-800 px-1 py-0.2 rounded border border-slate-700/50">P/N: {pn}</span>
                        <span>S/N: {sn}</span>
                        <span className="text-amber-400 font-bold">Raf: {loc}</span>
                      </div>
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          {/* Step 3: Product Location & Details */}
          {selectedMobileItem ? (
            <div className="bg-slate-900 border border-slate-800 rounded-3xl p-4.5 flex flex-col gap-4.5 shadow-xl animate-scale-up">
              <div className="border-b border-slate-800 pb-3 flex flex-col gap-1">
                <span className="text-[9px] font-black uppercase text-emerald-400 tracking-wider font-mono">SEÇİLEN MALZEME BİLGİSİ</span>
                <h3 className="text-xs font-black text-white leading-tight">{selectedMobileItem.description || selectedMobileItem.name}</h3>
                <div className="flex gap-2 items-center mt-1 text-[10px] text-slate-400 font-mono">
                  <span>P/N: <strong>{selectedMobileItem.partNumber || selectedMobileItem.pn}</strong></span>
                  <span>|</span>
                  <span>S/N: <strong>{selectedMobileItem.serialAndNotes || selectedMobileItem.sn}</strong></span>
                </div>
              </div>

              {/* Raf / Lokasyon Kartı - En Önemlisi */}
              <div className="bg-gradient-to-r from-amber-600/10 to-amber-600/5 border-2 border-amber-500/30 rounded-2xl p-4 flex items-center gap-3.5 shadow-inner">
                <div className="w-11 h-11 rounded-xl bg-amber-500/20 border border-amber-500/30 flex items-center justify-center text-amber-400 shrink-0">
                  <Package className="w-6 h-6" />
                </div>
                <div>
                  <span className="text-[9px] font-bold text-amber-500 uppercase tracking-wider block font-mono">DEPO YERİ / RAF KONUMU</span>
                  <span className="text-lg font-black text-white tracking-wide uppercase mt-0.5 block">
                    {selectedMobileItem.lokasyonNo || selectedMobileItem.location || 'BİLİNMİYOR'}
                  </span>
                </div>
              </div>

              {/* Hangi Depolarda Olduğu (Stok Miktarları) */}
              <div className="flex flex-col gap-2">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block font-mono">MİKTARIN BULUNDUĞU DEPOLAR</span>
                <div className="bg-slate-950 rounded-2xl p-3 border border-slate-850 flex flex-col gap-2 divide-y divide-slate-900">
                  {(() => {
                    const locations = [
                      { name: 'Ankara (Merkez)', qty: selectedMobileItem.ankaraMevcut },
                      { name: 'Karain', qty: selectedMobileItem.karainMevcut },
                      { name: 'Çanakkale', qty: selectedMobileItem.canakkaleMevcut },
                      { name: 'Milas', qty: selectedMobileItem.milasMevcut },
                      { name: 'Bursa', qty: selectedMobileItem.bursaMevcut }
                    ].filter(loc => loc.qty !== undefined && Number(loc.qty) > 0);

                    if (locations.length === 0) {
                      return <div className="text-xs text-slate-500 font-bold py-1 text-center">Hiçbir depoda mevcut stok bulunmamaktadır.</div>;
                    }

                    return locations.map((loc, i) => (
                      <div key={i} className={`flex items-center justify-between py-1.5 ${i > 0 ? 'pt-2' : ''}`}>
                        <span className="text-xs font-bold text-slate-300">{loc.name}</span>
                        <span className="text-xs font-black bg-emerald-500/20 text-emerald-400 px-2.5 py-0.5 rounded-full border border-emerald-500/20 font-mono">
                          {loc.qty} ADET
                        </span>
                      </div>
                    ));
                  })()}
                </div>
              </div>

              {/* Raf Ömrü & SKT Bilgileri */}
              <div className="grid grid-cols-2 gap-3">
                <div className="bg-slate-950/60 p-3 rounded-2xl border border-slate-850">
                  <span className="text-[9px] font-bold text-slate-500 block font-mono">RAF ÖMRÜNE TABİ</span>
                  <span className={`text-xs font-black block mt-0.5 ${
                    selectedMobileItem.hasShelfLife === 'EVET' || selectedMobileItem.hasShelfLife === true || String(selectedMobileItem.hasShelfLife).toUpperCase().includes('EVET')
                      ? 'text-emerald-400'
                      : 'text-slate-400'
                  }`}>
                    {selectedMobileItem.hasShelfLife === 'EVET' || selectedMobileItem.hasShelfLife === true || String(selectedMobileItem.hasShelfLife).toUpperCase().includes('EVET') ? 'EVET' : 'HAYIR'}
                  </span>
                </div>
                <div className="bg-slate-950/60 p-3 rounded-2xl border border-slate-850">
                  <span className="text-[9px] font-bold text-slate-500 block font-mono">SKT BİTİŞ TARİHİ</span>
                  <span className="text-xs font-black block text-slate-200 mt-0.5 whitespace-pre-line leading-tight">
                    {selectedMobileItem.shelfLifeDate && selectedMobileItem.shelfLifeDate !== '-' ? selectedMobileItem.shelfLifeDate : '-'}
                  </span>
                </div>
              </div>
            </div>
          ) : (
            <div className="flex-1 flex flex-col items-center justify-center p-8 text-center bg-slate-900/20 rounded-3xl border border-slate-900 border-dashed min-h-[220px]">
              <FileSearch className="w-10 h-11 text-slate-600 mb-3" />
              <h4 className="text-xs font-bold text-slate-400 uppercase">Sorgulama Sonucu</h4>
              <p className="text-[11px] text-slate-500 mt-1 max-w-[240px]">
                Lütfen sorgulama yapmak istediğiniz depoyu seçip, malzeme adını veya P/N kodunu arama kutusuna yazınız.
              </p>
            </div>
          )}
        </div>
      </div>
    );
  }

  // Unit switching with password
  const attemptSwitchUnit = (unitId: string) => {
    if (unlockedUnits.has(unitId) || unitId === 'all') {
      setCurrentUnit(unitId);
    } else {
      setPendingTargetUnit(unitId);
      setPasswordInput('');
      setPasswordError(false);
      setIsPasswordModalOpen(true);
    }
  };

  const handlePasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!pendingTargetUnit) return;
    const target = UNITS.find(u => u.id === pendingTargetUnit);
    if (!target) return;

    if (passwordInput === target.password || passwordInput === '1839') {
      setUnlockedUnits(prev => new Set([...prev, pendingTargetUnit]));
      setCurrentUnit(pendingTargetUnit);
      setIsPasswordModalOpen(false);
      setPasswordError(false);
      setPendingTargetUnit(null);
      showNotification(`🔓 ${target.name} birimi kilidi açıldı.`);
    } else {
      setPasswordError(true);
    }
  };

  // Available types filtered by movement direction (Giriş: Ankara Giren, Çıkış: Çıkanlar, Transfer: Transferler)
  const availableMovementTypes = useMemo(() => {
    const unitKey = currentUnit === 'all' ? 'at802' : currentUnit;
    let baseList: string[] = [];
    if (unitKey === 'at802') {
      if (movementDirection === 'giris') {
        // "MALZMELE GİRİŞİ YAPILDIĞINDA MALZME GİRİŞİ SADCE ANKARDAN OLUR ANKARA GİREN OLACAK"
        baseList = [...AT802_GIRIS_TYPES];
      } else if (movementDirection === 'cikis') {
        // "ÇIKAN DA HESPİ OLANB,LİR"
        baseList = [...AT802_CIKIS_TYPES];
      } else {
        // "TARSNFERDE TÜM TARSNFER TÜRÜ OLSUN YANİ TARSNFER SEÇTİĞİMDE ÇIKAN VEYA GİRŞİŞ YAZMAZ İŞLEMELRDE"
        baseList = [...AT802_TRANSFER_TYPES];
      }
    } else {
      const def = DEFAULT_TRANSFER_TYPES[unitKey] || DEFAULT_TRANSFER_TYPES.at802;
      if (movementDirection === 'giris') {
        baseList = def.filter(t => t.includes('GİREN') || t.includes('GİRİŞ') || t.includes('GELEN'));
        if (baseList.length === 0) baseList = ['ANKARA GİREN'];
      } else if (movementDirection === 'cikis') {
        baseList = def.filter(t => t.includes('ÇIKAN') || t.includes('ÇIKIŞ') || t.includes('GİDEN'));
        if (baseList.length === 0) baseList = ['ÇIKAN'];
      } else {
        baseList = def.filter(t => t.includes('TRANSFER') && !t.includes('GİREN') && !t.includes('ÇIKAN'));
        if (baseList.length === 0) baseList = ['DEPO TRANSFER'];
      }
    }
    const customs = customTransferTypes[unitKey] || [];
    return Array.from(new Set([...baseList, ...customs]));
  }, [currentUnit, movementDirection, customTransferTypes]);

  // Movement Form
  const openMovementModalHandler = (direction: 'cikis' | 'giris' | 'transfer') => {
    setMovementDirection(direction);
    setSelectedItemIdx(-1);
    setMovementItemSearch('');
    setMovementQty(1);
    const now = new Date();
    setMovementDate(`${now.toLocaleDateString('tr-TR')} ${now.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' })}`);
    
    const unitKey = currentUnit === 'all' ? 'at802' : currentUnit;
    let allowedTypes: string[] = [];

    if (unitKey === 'at802') {
      if (direction === 'giris') {
        // Malzeme girişi SADECE Ankara'dan olur (ANKARA GİREN)
        allowedTypes = [...AT802_GIRIS_TYPES];
        setMovementDepoYeri('ANKARA MERKEZ DEPO');
      } else if (direction === 'cikis') {
        // Çıkan hepsi olabilir (Ankara, Karain, Çanakkale, Milas, Bursa Çıkan...)
        allowedTypes = [...AT802_CIKIS_TYPES];
        setMovementDepoYeri('ANKARA MERKEZ DEPO');
      } else {
        // Transferde tüm transfer türleri (Karain, Çanakkale, Milas, Bursa Transfer...) - Çıkan veya Giriş yazmaz!
        allowedTypes = [...AT802_TRANSFER_TYPES];
        setMovementDepoYeri('DEPO');
      }
    } else {
      const defaultTypes = DEFAULT_TRANSFER_TYPES[unitKey] || DEFAULT_TRANSFER_TYPES.at802;
      if (direction === 'giris') {
        allowedTypes = defaultTypes.filter(t => t.includes('GİREN') || t.includes('GİRİŞ') || t.includes('GELEN'));
        if (allowedTypes.length === 0) allowedTypes = ['ANKARA GİREN'];
        setMovementDepoYeri('ANKARA MERKEZ DEPO');
      } else if (direction === 'cikis') {
        allowedTypes = defaultTypes.filter(t => t.includes('ÇIKAN') || t.includes('ÇIKIŞ') || t.includes('GİDEN'));
        if (allowedTypes.length === 0) allowedTypes = ['ÇIKAN'];
        setMovementDepoYeri('DEPO');
      } else {
        allowedTypes = defaultTypes.filter(t => t.includes('TRANSFER') && !t.includes('GİREN') && !t.includes('ÇIKAN'));
        if (allowedTypes.length === 0) allowedTypes = ['DEPO TRANSFER'];
        setMovementDepoYeri('DEPO');
      }
    }

    const customTypes = customTransferTypes[unitKey] || [];
    const allTypes = Array.from(new Set([...allowedTypes, ...customTypes]));
    
    setMovementType(allTypes[0] || (direction === 'giris' ? 'ANKARA GİREN' : direction === 'cikis' ? 'ANKARA ÇIKAN' : 'KARAİN TRANSFER'));
    setMovementCustomType('');
    setMovementSn('');
    setMovementNotes('');
    setCustomTailInput('');

    if (direction === 'transfer') {
      // Transfer merkezi ilk açılışta hep Ankara seçilir olacak
      setTransferSourceDepot('ANKARA');
      setTransferTargetDepot('KARAİN');
      setMovementType('KARAİN TRANSFER');
    }
    
    const tails = AIRCRAFT_TAILS[unitKey] || AIRCRAFT_TAILS.hangar;
    // Malzeme girişinde kuyruk no sorulmaz / seçilmez
    setMovementTail(direction === 'giris' ? '-' : (tails[0] || '-'));
    setMovementTeslimAlan('Yetkili Teknisyen');
    setMovementKabulYapan('Depo Sorumlusu');
    setIsMovementModalOpen(true);
  };

  const handleSelectItemForMovement = (item: DepoItem, idx: number) => {
    setSelectedItemIdx(idx);
    setMovementItemSearch(`${item.description || item.name} (P/N: ${item.partNumber || item.pn || '-'})`);
    setIsItemDropdownOpen(false);

    const rawSn = String(item.serialAndNotes || item.sn || '').trim();
    const snList = rawSn.split(/[,;]/).map(s => s.trim()).filter(s => s.length > 0);
    if (snList.length > 0) {
      setMovementSn(snList[0]);
    } else {
      setMovementSn(rawSn && rawSn !== '-' ? rawSn : '');
    }
    setMovementDepoYeri(
      (item.category === 'kimyasal' || currentDepoType === 'kimyasal')
        ? 'KİMYASAL DEPO - ANKARA'
        : (item.lokasyonNo || 'ANKARA MERKEZ DEPO')
    );
  };

  const handleMovementSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isMovementSubmitting) return;
    setIsMovementSubmitting(true);

    try {
      let item = selectedItemIdx >= 0 && selectedItemIdx < computedInventory.length ? computedInventory[selectedItemIdx] : null;
      if (!item && movementItemSearch.trim()) {
        const q = movementItemSearch.toLowerCase().trim();
        item = computedInventory.find(i => {
          const d = (i.description || i.name || '').toLowerCase();
          const p = (i.partNumber || i.pn || '').toLowerCase();
          return d.includes(q) || p.includes(q);
        }) || null;
      }

      if (!item) {
        alert("Lütfen listeden geçerli bir malzeme seçiniz!");
        setIsMovementSubmitting(false);
        return;
      }

      let finalType = movementType;
      if (movementType === 'DİĞER' && movementCustomType.trim()) {
        finalType = movementCustomType.trim().toUpperCase();
        const unitKey = currentUnit === 'all' ? 'at802' : currentUnit;
        setCustomTransferTypes(prev => {
          const existing = prev[unitKey] || [];
          if (!existing.includes(finalType)) {
            const updated = { ...prev, [unitKey]: [...existing, finalType] };
            localStorage.setItem('ogm_depo_custom_types_v5', JSON.stringify(updated));
            return updated;
          }
          return prev;
        });
      }

      // 1. Depo Stok Kontrolü ve Güvenlik Kısıtlamaları
      if (movementDirection === 'transfer') {
        finalType = `${transferTargetDepot.toUpperCase()} TRANSFER`;
        const src = transferSourceDepot.toUpperCase();
        let srcStock = Number(item.ankaraMevcut || 0);
        if (src.includes('KARAİN') || src.includes('KARAIN')) srcStock = Number(item.karainMevcut || 0);
        else if (src.includes('ÇANAKKALE') || src.includes('CANAKKALE')) srcStock = Number(item.canakkaleMevcut || 0);
        else if (src.includes('MİLAS') || src.includes('MILAS')) srcStock = Number(item.milasMevcut || 0);
        else if (src.includes('BURSA')) srcStock = Number(item.bursaMevcut || 0);

        if (movementQty > srcStock) {
          alert(`❌ Hatalı Transfer Engeli: ${transferSourceDepot} deposunda mevcut stok ${srcStock} adettir. ${movementQty} adet transfer yapılamaz!`);
          setIsMovementSubmitting(false);
          return;
        }
      } else if (movementDirection === 'cikis') {
        const info = getCikisDepotInfo(finalType, item);
        if (movementQty > info.stock) {
          alert(`❌ Hatalı Çıkış Engeli: Seçilen ${info.name} stoğunda yalnızca ${info.stock} adet mevcuttur. ${movementQty} adet çıkış yapılamaz!`);
          setIsMovementSubmitting(false);
          return;
        }
      }

      // Kuyruk No: Giriş ve transferde '-', çıkışta 'Diğer' seçilmişse manuel input kullanılır
      const finalTailNo = (movementDirection === 'giris' || movementDirection === 'transfer') 
        ? '-' 
        : (movementTail === 'DIGER' ? (customTailInput.trim() || 'DİĞER') : movementTail);

      const targetSheet = getDepoSheetNameForUnit(currentUnit);

      const newTx: DepoTransaction = {
        id: `tx_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        timestamp: new Date().toISOString(),
        type: finalType,
        itemName: item.description || item.name || '',
        pn: item.partNumber || item.pn || '-',
        sn: movementSn.trim() || item.partNumber || '-',
        quantity: movementQty,
        date: movementDate,
        tailNo: finalTailNo,
        operator: movementTeslimAlan,
        location: (item.category === 'kimyasal' || currentDepoType === 'kimyasal') 
          ? (movementDirection === 'transfer' ? `KİMYASAL DEPO - ${transferSourceDepot.toUpperCase()} ➔ KİMYASAL DEPO - ${transferTargetDepot.toUpperCase()}` : (movementDepoYeri && movementDepoYeri.toUpperCase().includes('KİMYASAL') ? movementDepoYeri.toUpperCase() : `KİMYASAL DEPO - ${(movementDepoYeri || 'ANKARA').toUpperCase()}`))
          : (movementDirection === 'transfer' ? `${transferSourceDepot.toUpperCase()} ➔ ${transferTargetDepot.toUpperCase()}` : (movementDepoYeri || 'ANKARA MERKEZ DEPO')),
        depoYeri: (item.category === 'kimyasal' || currentDepoType === 'kimyasal') 
          ? (movementDirection === 'transfer' ? `KİMYASAL DEPO - ${transferSourceDepot.toUpperCase()} ➔ KİMYASAL DEPO - ${transferTargetDepot.toUpperCase()}` : (movementDepoYeri && movementDepoYeri.toUpperCase().includes('KİMYASAL') ? movementDepoYeri.toUpperCase() : `KİMYASAL DEPO - ${(movementDepoYeri || 'ANKARA').toUpperCase()}`))
          : (movementDirection === 'transfer' ? `${transferSourceDepot.toUpperCase()} ➔ ${transferTargetDepot.toUpperCase()}` : (movementDepoYeri || 'ANKARA MERKEZ DEPO')),
        unit: currentUnit,
        category: (item.category === 'kimyasal' || currentDepoType === 'kimyasal') ? 'kimyasal' : (item.category || 'sarf'),
        sheetName: targetSheet,
        notes: movementNotes.trim() || `${finalType} işlemi kaydedildi.`,
        isNewSessionTx: true
      };

      // 2. Tablodaki Depo Mevcutlarını ve Stokları Güncelle
      const targetPn = (item.partNumber || item.pn || '').trim().toLowerCase();
      const targetDesc = (item.description || item.name || '').trim().toLowerCase();
      const updatedInventory = inventory.map(invItem => {
        const invPn = (invItem.partNumber || invItem.pn || '').trim().toLowerCase();
        const invDesc = (invItem.description || invItem.name || '').trim().toLowerCase();
        if ((targetPn && invPn === targetPn) || (targetDesc && invDesc === targetDesc)) {
          let gelen = Number(invItem.gelen || invItem.baseGelen || 0);
          let ankaraCikan = Number(invItem.ankaraCikan || 0);
          let ankaraMevcut = Number(invItem.ankaraMevcut || 0);
          let karainTransfer = Number(invItem.karainTransfer || 0);
          let karainCikan = Number(invItem.karainCikan || 0);
          let karainMevcut = Number(invItem.karainMevcut || 0);
          let canakkaleTransfer = Number(invItem.canakkaleTransfer || 0);
          let canakkaleCikan = Number(invItem.canakkaleCikan || 0);
          let canakkaleMevcut = Number(invItem.canakkaleMevcut || 0);
          let milasTransfer = Number(invItem.milasTransfer || 0);
          let milasCikan = Number(invItem.milasCikan || 0);
          let milasMevcut = Number(invItem.milasMevcut || 0);
          let bursaTransfer = Number(invItem.bursaTransfer || 0);
          let bursaCikan = Number(invItem.bursaCikan || 0);
          let bursaMevcut = Number(invItem.bursaMevcut || 0);
          let muayeneToplam = Number(invItem.muayeneToplam || 0);

          if (movementDirection === 'giris') {
            gelen += movementQty;
            ankaraMevcut += movementQty;
          } else if (movementDirection === 'transfer') {
            // Çıkış yapılan depodan düş
            const src = transferSourceDepot.toUpperCase();
            if (src.includes('KARAİN') || src.includes('KARAIN')) {
              karainCikan += movementQty;
              karainMevcut = Math.max(0, karainMevcut - movementQty);
            } else if (src.includes('ÇANAKKALE')) {
              canakkaleCikan += movementQty;
              canakkaleMevcut = Math.max(0, canakkaleMevcut - movementQty);
            } else if (src.includes('MİLAS')) {
              milasCikan += movementQty;
              milasMevcut = Math.max(0, milasMevcut - movementQty);
            } else if (src.includes('BURSA')) {
              bursaCikan += movementQty;
              bursaMevcut = Math.max(0, bursaMevcut - movementQty);
            } else {
              ankaraCikan += movementQty;
              ankaraMevcut = Math.max(0, ankaraMevcut - movementQty);
            }

            // Transfer edilen hedef depoya ekle
            const tgt = transferTargetDepot.toUpperCase();
            if (tgt.includes('KARAİN') || tgt.includes('KARAIN')) {
              karainTransfer += movementQty;
              karainMevcut += movementQty;
            } else if (tgt.includes('ÇANAKKALE')) {
              canakkaleTransfer += movementQty;
              canakkaleMevcut += movementQty;
            } else if (tgt.includes('MİLAS')) {
              milasTransfer += movementQty;
              milasMevcut += movementQty;
            } else if (tgt.includes('BURSA')) {
              bursaTransfer += movementQty;
              bursaMevcut += movementQty;
            } else if (tgt.includes('ANKARA')) {
              gelen += movementQty;
              ankaraMevcut += movementQty;
            }
          } else if (movementDirection === 'cikis') {
            if (finalType.includes('KARAİN')) {
              karainCikan += movementQty;
              karainMevcut = Math.max(0, karainMevcut - movementQty);
            } else if (finalType.includes('ÇANAKKALE')) {
              canakkaleCikan += movementQty;
              canakkaleMevcut = Math.max(0, canakkaleMevcut - movementQty);
            } else if (finalType.includes('MİLAS')) {
              milasCikan += movementQty;
              milasMevcut = Math.max(0, milasMevcut - movementQty);
            } else if (finalType.includes('BURSA')) {
              bursaCikan += movementQty;
              bursaMevcut = Math.max(0, bursaMevcut - movementQty);
            } else {
              ankaraCikan += movementQty;
              ankaraMevcut = Math.max(0, ankaraMevcut - movementQty);
            }
          }

          const toplamStok = ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut + muayeneToplam;
          return {
            ...invItem,
            gelen,
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
          };
        }
        return invItem;
      });

      setInventory(updatedInventory);
      const allTransfers = [newTx, ...transactions];
      setTransactions(allTransfers);
      if (onAddTransaction) {
        onAddTransaction(newTx);
      }

      setIsMovementModalOpen(false);
      showNotification(`✅ ${finalType} işlemi kaydedildi, stoklar güncellendi ve Excel Online'a aktarıldı!`);

      // Eğer Malzeme Girişi yapıldıysa, anında Depo Girişi Slibi penceresini aç (Slip İndir & Yazdır için)
      if (movementDirection === 'giris') {
        setSlipModalType('giris');
        setSlipModalData({
          type: 'giris',
          islemTransfer: finalType,
          tarih: movementDate,
          malzemeDesc: item.description || item.name || '',
          pn: item.partNumber || item.pn || '-',
          sn: movementSn.trim() || item.partNumber || '-',
          adet: movementQty,
          depoYeri: movementDepoYeri,
          teslimAlan: movementTeslimAlan,
          kabulYapan: movementKabulYapan,
          aciklama: movementNotes.trim() || 'Malzeme kutusunda hasar yok. Sertifika ve kabul formu onaylandı. Sayım girişi tamam.'
        });
        setIsGirisOnlySlip(true);
        setIsSlipModalOpen(true);
      }

      // 3. Otomatik Veri Senkronizasyonu (Excel Online & Drive)
      await syncDepoToExcelOnline(updatedInventory, allTransfers);
    } catch (err: any) {
      alert("İşlem kaydedilemedi: " + err.message);
    } finally {
      setIsMovementSubmitting(false);
    }
  };

  // Kit Oluşturma Fonksiyonları
  const openKitModalHandler = () => {
    setKitName('');
    setKitCount(1);
    setKitItems([]);
    setKitSearchQuery('');
    setSelectedKitItem(null);
    setKitItemQtyInput(1);
    setIsKitModalOpen(true);
  };

  const handleAddKitItem = () => {
    if (!selectedKitItem) {
      alert("Lütfen sistemdeki PN listesinden geçerli bir parça seçiniz!");
      return;
    }
    const pn = (selectedKitItem.partNumber || selectedKitItem.pn || '-').trim();
    const desc = (selectedKitItem.description || selectedKitItem.name || '').trim();
    const ankaraMevcut = Number(selectedKitItem.ankaraMevcut || 0);

    const existingIdx = kitItems.findIndex(k => k.partNumber.toLowerCase() === pn.toLowerCase() && k.description.toLowerCase() === desc.toLowerCase());
    if (existingIdx >= 0) {
      const updated = [...kitItems];
      updated[existingIdx].qtyPerKit += kitItemQtyInput;
      setKitItems(updated);
    } else {
      setKitItems(prev => [...prev, {
        partNumber: pn,
        description: desc,
        qtyPerKit: kitItemQtyInput,
        ankaraMevcut: ankaraMevcut
      }]);
    }

    setSelectedKitItem(null);
    setKitSearchQuery('');
    setKitItemQtyInput(1);
    setIsKitDropdownOpen(false);
  };

  const handleRemoveKitItem = (idx: number) => {
    setKitItems(prev => prev.filter((_, i) => i !== idx));
  };

  const handleCreateKitFromSelection = () => {
    const selectedList: { partNumber: string; description: string; qtyPerKit: number; ankaraMevcut: number }[] = [];

    if (currentDepoType === 'transactions' && selectedTxIdsForKit.length > 0) {
      selectedTxIdsForKit.forEach(id => {
        const tx = transactions.find(t => t.id === id);
        if (tx) {
          const pn = (tx.pn || tx.partNumber || '-').trim();
          const desc = (tx.itemName || tx.itemDesc || tx.name || '').trim();
          const invMatch = computedInventory.find(i => (i.partNumber || i.pn || '').toLowerCase() === pn.toLowerCase() || (i.description || i.name || '').toLowerCase() === desc.toLowerCase());
          const curStock = invMatch ? Number(invMatch.ankaraMevcut || invMatch.toplamStok || 0) : 10;
          if (!selectedList.some(k => k.partNumber.toLowerCase() === pn.toLowerCase() && k.description.toLowerCase() === desc.toLowerCase())) {
            selectedList.push({
              partNumber: pn,
              description: desc,
              qtyPerKit: 1,
              ankaraMevcut: curStock
            });
          }
        }
      });
    } else if (selectedItemPnsForKit.length > 0) {
      selectedItemPnsForKit.forEach(key => {
        const item = computedInventory.find(i => `${(i.partNumber || i.pn || '').trim()}_${(i.description || i.name || '').trim()}` === key);
        if (item) {
          const pn = (item.partNumber || item.pn || '-').trim();
          const desc = (item.description || item.name || '').trim();
          const curStock = Number(item.ankaraMevcut || item.toplamStok || 0);
          if (!selectedList.some(k => k.partNumber.toLowerCase() === pn.toLowerCase() && k.description.toLowerCase() === desc.toLowerCase())) {
            selectedList.push({
              partNumber: pn,
              description: desc,
              qtyPerKit: 1,
              ankaraMevcut: curStock
            });
          }
        }
      });
    }

    if (selectedList.length === 0) {
      alert("Lütfen kiti oluşturacak parçaları tablodan kutucukları işaretleyerek seçiniz!");
      return;
    }

    setKitName('');
    setKitCount(1);
    setKitItems(selectedList);
    setKitSearchQuery('');
    setSelectedKitItem(null);
    setKitItemQtyInput(1);
    setIsKitModalOpen(true);
  };

  const handleCreateKitSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!kitName.trim()) {
      alert("Lütfen kit adını giriniz!");
      return;
    }
    if (kitCount < 1) {
      alert("Lütfen en az 1 kit adedi giriniz!");
      return;
    }
    if (kitItems.length === 0) {
      alert("Lütfen kiti oluşturacak en az bir parça ekleyiniz!");
      return;
    }

    // Stok Yeterlilik Kontrolü: Formül: (Kit İçi Parça Adedi) × (Toplam Üretilecek Kit Sayısı)
    for (const ki of kitItems) {
      const required = ki.qtyPerKit * kitCount;
      if (ki.ankaraMevcut < required) {
        alert(`❌ Yetersiz Stok! "${ki.description}" (P/N: ${ki.partNumber}) için gereken miktar: ${required}, ancak Ankara deposundaki mevcut: ${ki.ankaraMevcut}`);
        return;
      }
    }

    const unitKey = currentUnit !== 'all' ? currentUnit : 'at802';
    const updatedInventory = [...inventory];
    const newTransactions: DepoTransaction[] = [];
    const dateStr = new Date().toLocaleDateString('tr-TR');

    // 1. Her parçadan (Kit İçi Adet x Kit Sayısı) kadar SADECE Ankara deposundan düş
    kitItems.forEach(ki => {
      const deductQty = ki.qtyPerKit * kitCount;
      const targetItemIdx = updatedInventory.findIndex(i => {
        const itemPn = (i.partNumber || i.pn || '').trim().toLowerCase();
        const searchPn = ki.partNumber.trim().toLowerCase();
        if (searchPn && searchPn !== '-' && itemPn === searchPn) return true;
        return (i.description || i.name || '').trim().toLowerCase() === ki.description.trim().toLowerCase();
      });

      if (targetItemIdx >= 0) {
        const it = { ...updatedInventory[targetItemIdx] };
        it.ankaraCikan = (it.ankaraCikan || 0) + deductQty;
        it.ankaraMevcut = Math.max(0, (it.ankaraMevcut || 0) - deductQty);
        // Sadece Ankara stoğundan düşülür; diğer tüm depolar korunur
        it.toplamStok = (it.ankaraMevcut || 0) + (it.karainMevcut || 0) + (it.canakkaleMevcut || 0) + (it.milasMevcut || 0) + (it.bursaMevcut || 0) + (it.muayeneToplam || 0);
        updatedInventory[targetItemIdx] = it;
      }

      // Depo Hareket Geçmişine parça düşüş kaydını ekle (Kullanıcı kuralı: ÇIKAN (KİT OLUŞTURMA) olarak eklenecek)
      newTransactions.push({
        id: `tx_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        timestamp: new Date().toISOString(),
        type: 'ÇIKAN (KİT OLUŞTURMA)',
        itemName: ki.description,
        pn: ki.partNumber,
        sn: '-',
        quantity: deductQty,
        date: dateStr,
        tailNo: '-',
        operator: 'Kit Montaj',
        receivedBy: 'Kit Üretimi',
        location: 'ANKARA MERKEZ DEPO',
        sheetName: 'DEPO HAREKET GEÇMİŞİ-AT-802',
        notes: `ÇIKAN (KİT OLUŞTURMA): "${kitName.trim()}" kiti (${kitCount} adet) için ilgili depodan ${deductQty} adet parça düşüldü.`,
        isNewSessionTx: true
      });
    });

    // 2. Oluşturulan Kiti tablonun en altına yeni ürün olarak ekle (Kullanıcı kuralı: kitin pn no kısmı ve lokasyon no sisteme eklenirken boş olacak)
    const newKitRow: DepoItem = {
      unit: unitKey,
      category: 'sarf',
      description: kitName.trim().toUpperCase(),
      name: kitName.trim().toUpperCase(),
      partNumber: '',
      pn: '',
      serialAndNotes: `KİT İÇERİĞİ: ${kitItems.map(k => `${k.partNumber} (${k.qtyPerKit} ad.)`).join(', ')}`,
      sn: `LOT-${new Date().getFullYear()}`,
      lokasyonNo: '',
      location: '',
      baseGelen: kitCount,
      gelen: kitCount,
      toplamStok: kitCount,
      ankaraCikan: 0,
      ankaraMevcut: kitCount, // Ankara Depo Mevcudu: Üretilen toplam kit adedi
      karainTransfer: 0,
      karainCikan: 0,
      karainMevcut: 0,      // Diğer Depo Mevcutları: 0 (Sıfır)
      canakkaleTransfer: 0,
      canakkaleCikan: 0,
      canakkaleMevcut: 0,
      milasTransfer: 0,
      milasCikan: 0,
      milasMevcut: 0,
      bursaTransfer: 0,
      bursaCikan: 0,
      bursaMevcut: 0,
      muayeneGiden: 0,
      muayeneGelen: 0,
      muayeneToplam: 0
    };

    updatedInventory.push(newKitRow);

    // Depo Hareket Geçmişine kit giriş kaydını ekle
    newTransactions.push({
      id: `tx_${Date.now()}_kit`,
      timestamp: new Date().toISOString(),
      type: 'ANKARA GİREN',
      itemName: newKitRow.description,
      pn: newKitRow.partNumber,
      sn: newKitRow.sn || '-',
      quantity: kitCount,
      date: dateStr,
      tailNo: '-',
      operator: 'Kit Montaj',
      receivedBy: 'Depo Sorumlusu',
      location: 'ANKARA MERKEZ DEPO',
      sheetName: 'DEPO HAREKET GEÇMİŞİ-AT-802',
      notes: `KİT OLUŞTURULDU: "${kitName.trim()}" kiti (${kitCount} adet) üretildi ve Ankara Merkez Depo stoğuna eklendi.`,
      isNewSessionTx: true
    });

    const allTransfers = [...newTransactions, ...transactions];
    setTransactions(allTransfers);
    setInventory(updatedInventory);

    if (onAddTransaction) {
      newTransactions.forEach(t => onAddTransaction(t));
    }

    setIsKitModalOpen(false);
    showNotification(`✅ "${kitName.trim()}" kiti (${kitCount} adet) başarıyla üretildi, stoklar düşüldü ve Excel Online'a aktarıldı!`);

    // 3. Otomatik Excel Online ve Drive Senkronizasyonu
    syncDepoToExcelOnline(updatedInventory, allTransfers);
  };

  // ─── MUAYENE / BAKIM GÖNDER FONKSİYONLARI ───
  const openMuayeneModalHandler = () => {
    setMuayenePn('');
    setMuayeneSn('');
    setMuayeneDescription('');
    setMuayeneQty(1);
    setMuayeneNotes('');
    setMuayeneSourceDepot('ANKARA');
    setMuayeneOperator('');
    setSelectedMuayeneItem(null);
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    setMuayeneDate(`${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
    setIsMuayeneModalOpen(true);
  };

  const handleSelectMuayeneItem = (item: DepoItem) => {
    setSelectedMuayeneItem(item);
    setMuayenePn(item.partNumber || item.pn || '');
    setMuayeneDescription(item.description || item.name || '');
    const snVal = item.serialAndNotes || item.sn || '';
    setMuayeneSn(snVal !== '-' ? snVal : '');
    setIsMuayeneDropdownOpen(false);
  };

  const handleMuayeneSubmit = async (e: React.FormEvent, keepOpen: boolean = false) => {
    e.preventDefault();
    if (!selectedMuayeneItem && !muayenePn.trim()) {
      alert("Lütfen muayene/bakıma gönderilecek parçayı seçiniz!");
      return;
    }
    if (muayeneQty < 1) {
      alert("Lütfen en az 1 adet miktar giriniz!");
      return;
    }

    const targetItem = selectedMuayeneItem || computedInventory.find(i => {
      const p = (i.partNumber || i.pn || '').trim().toLowerCase();
      return p && p === muayenePn.trim().toLowerCase();
    });

    if (!targetItem) {
      alert("Seçilen parça envanterde bulunamadı!");
      return;
    }

    // Depo stoğu kontrolü
    const src = muayeneSourceDepot.toUpperCase();
    let srcStock = Number(targetItem.ankaraMevcut || 0);
    if (src.includes('KARAİN') || src.includes('KARAIN')) srcStock = Number(targetItem.karainMevcut || 0);
    else if (src.includes('ÇANAKKALE')) srcStock = Number(targetItem.canakkaleMevcut || 0);
    else if (src.includes('MİLAS')) srcStock = Number(targetItem.milasMevcut || 0);
    else if (src.includes('BURSA')) srcStock = Number(targetItem.bursaMevcut || 0);

    if (srcStock < muayeneQty) {
      alert(`❌ Yetersiz Stok! ${muayeneSourceDepot} deposundaki mevcut stok: ${srcStock}, gönderilmek istenen: ${muayeneQty}`);
      return;
    }

    const updatedInventory = [...inventory];
    const itemIdx = updatedInventory.findIndex(i => {
      const matchPn = (i.partNumber || i.pn || '').trim().toLowerCase() === (targetItem.partNumber || targetItem.pn || '').trim().toLowerCase();
      const matchDesc = (i.description || i.name || '').trim().toLowerCase() === (targetItem.description || targetItem.name || '').trim().toLowerCase();
      return matchPn && matchDesc;
    });

    if (itemIdx >= 0) {
      const it = { ...updatedInventory[itemIdx] };
      if (src === 'ANKARA') {
        it.ankaraCikan = (it.ankaraCikan || 0) + muayeneQty;
        it.ankaraMevcut = Math.max(0, (it.ankaraMevcut || 0) - muayeneQty);
      } else if (src.includes('KARAİN') || src.includes('KARAIN')) {
        it.karainCikan = (it.karainCikan || 0) + muayeneQty;
        it.karainMevcut = Math.max(0, (it.karainMevcut || 0) - muayeneQty);
      } else if (src.includes('ÇANAKKALE')) {
        it.canakkaleCikan = (it.canakkaleCikan || 0) + muayeneQty;
        it.canakkaleMevcut = Math.max(0, (it.canakkaleMevcut || 0) - muayeneQty);
      } else if (src.includes('MİLAS')) {
        it.milasCikan = (it.milasCikan || 0) + muayeneQty;
        it.milasMevcut = Math.max(0, (it.milasMevcut || 0) - muayeneQty);
      } else if (src.includes('BURSA')) {
        it.bursaCikan = (it.bursaCikan || 0) + muayeneQty;
        it.bursaMevcut = Math.max(0, (it.bursaMevcut || 0) - muayeneQty);
      }

      it.muayeneGiden = (it.muayeneGiden || 0) + muayeneQty;
      it.muayeneToplam = (it.muayeneToplam || 0) + muayeneQty;
      it.toplamStok = (it.ankaraMevcut || 0) + (it.karainMevcut || 0) + (it.canakkaleMevcut || 0) + (it.milasMevcut || 0) + (it.bursaMevcut || 0) + (it.muayeneToplam || 0);
      updatedInventory[itemIdx] = it;
    }

    const targetSheet = getDepoSheetNameForUnit(currentUnit !== 'all' ? currentUnit : targetItem.unit || 'at802');
    const newTx: DepoTransaction = {
      id: `tx_muayene_${Date.now()}`,
      timestamp: new Date().toISOString(),
      type: 'MUAYENE GİDEN',
      itemName: muayeneDescription.trim() || targetItem.description,
      pn: muayenePn.trim() || targetItem.partNumber || '-',
      sn: muayeneSn.trim() || targetItem.serialAndNotes || '-',
      quantity: muayeneQty,
      date: muayeneDate.trim() || new Date().toLocaleString('tr-TR'),
      tailNo: '-',
      operator: muayeneOperator.trim() || 'Muayene Birimi',
      receivedBy: 'Muayene / Bakım Atölyesi',
      location: `${muayeneSourceDepot} ➔ MUAYENE / BAKIM`,
      unit: targetItem.unit || currentUnit,
      category: targetItem.category || 'sarf',
      sheetName: targetSheet,
      notes: muayeneNotes.trim() || 'Parça muayene ve dış bakıma sevk edildi.',
      isNewSessionTx: true
    };

    const allTransfers = [newTx, ...transactions];
    setTransactions(allTransfers);
    setInventory(updatedInventory);
    setMuayeneSessionList(prev => [newTx, ...prev]);

    if (onAddTransaction) {
      onAddTransaction(newTx);
    }

    showNotification(`✅ "${newTx.itemName}" parçası (${muayeneQty} adet) başarıyla muayene/bakıma sevk edildi.`);
    syncDepoToExcelOnline(updatedInventory, allTransfers);

    if (keepOpen) {
      setMuayenePn('');
      setMuayeneSn('');
      setMuayeneDescription('');
      setMuayeneQty(1);
      setSelectedMuayeneItem(null);
    } else {
      setIsMuayeneModalOpen(false);
    }
  };

  // ─── ONAY BEKLEYENLER (APPROVAL) FONKSİYONLARI ───
  const handleOnayPasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const authorized = ['802', '1839', '1234', '8902', '1923', '2024'];
    if (authorized.includes(onayPasswordInput.trim())) {
      setIsOnayPasswordModalOpen(false);
      setIsOnayBekleyenlerModalOpen(true);
      fetchOnayBekleyenler();
    } else {
      setOnayPasswordError('Hatalı Şifre! Yetkili şifresini (1839) giriniz.');
    }
  };

  const handleApproveRequest = async (reqItem: any) => {
    try {
      const res = await fetch('/api/onay-bekleyenler/onayla', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id: reqItem.id,
          password: onayPasswordInput || '1839',
          approver: 'Depo Sorumlusu'
        })
      });
      const data = await res.json();
      if (data && data.status === 'success') {
        showNotification(`✅ Talep (${reqItem.id}) onaylandı ve malzeme çıkışı gerçekleştirildi.`);
        
        // Optimistic stock deduction in frontend
        const targetDepot = (reqItem.depot || 'ANKARA').toUpperCase();
        const reqQty = Number(reqItem.quantity || reqItem.adet || 1);
        const normalizeStr = (s?: string) => String(s || '').replace(/["' \-_]/g, '').toLowerCase();

        const reqPnNorm = normalizeStr(reqItem.pn);
        const reqDescNorm = normalizeStr(reqItem.itemName || reqItem.description);

        const notesLower = (reqItem.notes || '').toLowerCase();
        const isSayim = reqItem.category === 'sayim' || notesLower.includes('sayım') || notesLower.includes('sayim') || notesLower.includes('fazla') || notesLower.includes('fark');
        const isFromAnkara = (notesLower.includes('ankara') || notesLower.includes('ankardan')) && !targetDepot.includes('ANKARA');

        let matchedAny = false;
        const updatedInventory = inventory.map(item => {
          const itemPnNorm = normalizeStr(item.partNumber || item.pn);
          const itemDescNorm = normalizeStr(item.description || item.name);

          const matchPn = reqPnNorm && itemPnNorm && (itemPnNorm === reqPnNorm || itemPnNorm.includes(reqPnNorm) || reqPnNorm.includes(itemPnNorm));
          const matchDesc = reqDescNorm && itemDescNorm && (itemDescNorm === reqDescNorm || itemDescNorm.includes(reqDescNorm) || reqDescNorm.includes(itemDescNorm));

          if (!matchedAny && (matchPn || matchDesc)) {
            matchedAny = true;
            const it = { ...item };
            if (isSayim) {
              // Sayım Fazlası / Transfer Kaydı
              if (targetDepot === 'ANKARA') {
                it.gelen = (it.gelen || 0) + reqQty;
                it.ankaraMevcut = (it.ankaraMevcut || 0) + reqQty;
              } else if (targetDepot.includes('KARAİN') || targetDepot.includes('KARAIN')) {
                it.karainTransfer = (it.karainTransfer || 0) + reqQty;
                it.karainMevcut = (it.karainMevcut || 0) + reqQty;
              } else if (targetDepot.includes('ÇANAKKALE')) {
                it.canakkaleTransfer = (it.canakkaleTransfer || 0) + reqQty;
                it.canakkaleMevcut = (it.canakkaleMevcut || 0) + reqQty;
              } else if (targetDepot.includes('MİLAS')) {
                it.milasTransfer = (it.milasTransfer || 0) + reqQty;
                it.milasMevcut = (it.milasMevcut || 0) + reqQty;
              } else if (targetDepot.includes('BURSA')) {
                it.bursaTransfer = (it.bursaTransfer || 0) + reqQty;
                it.bursaMevcut = (it.bursaMevcut || 0) + reqQty;
              }
            } else {
              // Normal Talep Çıkışı
              if (targetDepot === 'ANKARA') {
                it.ankaraCikan = (it.ankaraCikan || 0) + reqQty;
                it.ankaraMevcut = Math.max(0, (it.ankaraMevcut || 0) - reqQty);
              } else if (targetDepot.includes('KARAİN') || targetDepot.includes('KARAIN')) {
                it.karainCikan = (it.karainCikan || 0) + reqQty;
                it.karainMevcut = Math.max(0, (it.karainMevcut || 0) - reqQty);
              } else if (targetDepot.includes('ÇANAKKALE')) {
                it.canakkaleCikan = (it.canakkaleCikan || 0) + reqQty;
                it.canakkaleMevcut = Math.max(0, (it.canakkaleMevcut || 0) - reqQty);
              } else if (targetDepot.includes('MİLAS')) {
                it.milasCikan = (it.milasCikan || 0) + reqQty;
                it.milasMevcut = Math.max(0, (it.milasMevcut || 0) - reqQty);
              } else if (targetDepot.includes('BURSA')) {
                it.bursaCikan = (it.bursaCikan || 0) + reqQty;
                it.bursaMevcut = Math.max(0, (it.bursaMevcut || 0) - reqQty);
              }
            }
            it.toplamStok = (it.ankaraMevcut || 0) + (it.karainMevcut || 0) + (it.canakkaleMevcut || 0) + (it.milasMevcut || 0) + (it.bursaMevcut || 0) + (it.muayeneToplam || 0);
            return it;
          }
          return item;
        });

        setInventory(updatedInventory);

        const newApprovedTx: DepoTransaction = data.approvedTx || {
          id: `tx_approved_${Date.now()}`,
          timestamp: new Date().toISOString(),
          date: new Date().toLocaleString('tr-TR'),
          type: `${targetDepot} ÇIKAN`,
          itemName: reqItem.itemName || reqItem.description || 'Malzeme',
          pn: reqItem.pn || '-',
          sn: reqItem.sn || '-',
          quantity: reqQty,
          tailNo: reqItem.tailNo || '-',
          operator: reqItem.requestedBy || 'Personel',
          receivedBy: 'Depo Sorumlusu',
          location: targetDepot === 'ANKARA' ? 'ANKARA MERKEZ DEPO' : `${targetDepot} DEPOSU`,
          unit: currentUnit !== 'all' ? currentUnit : 'at802',
          category: reqItem.category || 'sarf',
          sheetName: getDepoSheetNameForUnit(currentUnit),
          notes: `ONAYLANAN TALEP (${reqItem.id}): ${reqItem.notes || ''}`
        };

        const allTransfers = [newApprovedTx, ...transactions];
        setTransactions(allTransfers);
        if (onAddTransaction) onAddTransaction(newApprovedTx);
        syncDepoToExcelOnline(updatedInventory, allTransfers);

        fetchOnayBekleyenler();
      } else {
        alert("Onaylanamadı: " + (data.message || 'Bilinmeyen hata'));
      }
    } catch (err: any) {
      alert("Hata: " + err.message);
    }
  };

  const handleRejectRequest = async (reqItem: any) => {
    const reason = prompt("Lütfen ret nedenini belirtiniz:") || "Depo yetkilisi tarafından reddedildi.";
    try {
      const res = await fetch('/api/onay-bekleyenler/reddet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id: reqItem.id,
          password: onayPasswordInput || '1839',
          rejector: 'Depo Sorumlusu',
          reason
        })
      });
      const data = await res.json();
      if (data && data.status === 'success') {
        showNotification(`Talep (${reqItem.id}) reddedildi.`);
        fetchOnayBekleyenler();
      }
    } catch (err: any) {
      alert("Hata: " + err.message);
    }
};
// FİLTRELENMİŞ EXCEL İNDİR - YALNIZCA SEÇİLİ BÖLGE SÜTUNLARINI VE FİLTRELENMİŞ SATIRLARI İNDİRİR
const handleDownloadFilteredExcel = () => {
if (currentDepoType === 'yasam_destek') {
  setIsYasamDestekExportTriggered(true);
  return;
}
const activeUnit = (currentUnit || 'at802').toUpperCase();
const reg = (filterRegion || 'ALL').toUpperCase().trim();
const isKimyasal = currentDepoType === 'kimyasal';
const rowsToExport = isKimyasal ? filteredKimyasalRows : filteredSarfRows;
if (rowsToExport.length === 0) {
alert("İndirilecek filtrelenmiş veri bulunamadı. Lütfen filtrelerinizi kontrol ediniz.");
return;
}
const aoaData: (string | number)[][] = [];
if (isKimyasal) {
// Kimyasal Depo Başlıkları (Yalnızca seçili bölge sütunlarını içerir)
const headerRow: string[] = [
"No",
"Description",
"P/N - SPAIN DISCRIPTION",
"MİKTAR",
"QTY SÖZLEŞME TOPLAMI",
"PİYASA GELEN",
"SPAIN GELEN",
"GENEL TOPLAM",
"TOPLAM STOK"
];
if (reg === 'ALL' || reg === 'ANKARA') {
headerRow.push("ANKARA ÇIKAN", "ANKARA MEVCUT");
}
if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
headerRow.push("KARAİN TRANSFER", "KARAİN ÇIKAN", "KARAİN MEVCUT");
}
if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
headerRow.push("ÇANAKKALE TRANSFER", "ÇANAKKALE ÇIKAN", "ÇANAKKALE MEVCUT");
}
if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
headerRow.push("MİLAS TRANSFER", "MİLAS ÇIKAN", "MİLAS MEVCUT");
}
if (reg === 'ALL' || reg === 'BURSA') {
headerRow.push("BURSA TRANSFER", "BURSA ÇIKAN", "BURSA MEVCUT");
}
headerRow.push("AÇIKLAMA", "RAF ÖMRÜ VAR MI?", "RAF ÖMRÜ BİTİŞ TARİHİ");
aoaData.push(headerRow);
// Veri Satırları
rowsToExport.forEach((item, idx) => {
const row: (string | number)[] = [
idx + 1,
item.description || item.name || '',
item.spainDescription || item.partNumber || item.pn || '-',
item.miktarQty || 'Adet',
item.sozlesmeToplami || '0',
item.piyasa || 0,
item.spainGelen || 0,
item.genelToplam || 0,
item.toplamStok || 0
];
if (reg === 'ALL' || reg === 'ANKARA') {
row.push(item.ankaraCikan || 0, item.ankaraMevcut || 0);
}
if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
row.push(item.karainTransfer || 0, item.karainCikan || 0, item.karainMevcut || 0);
}
if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
row.push(item.canakkaleTransfer || 0, item.canakkaleCikan || 0, item.canakkaleMevcut || 0);
}
if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
row.push(item.milasTransfer || 0, item.milasCikan || 0, item.milasMevcut || 0);
}
if (reg === 'ALL' || reg === 'BURSA') {
row.push(item.bursaTransfer || 0, item.bursaCikan || 0, item.bursaMevcut || 0);
}
const hasLife = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
row.push(
item.serialAndNotes || item.sn || '-',
hasLife ? 'EVET' : 'HAYIR',
hasLife && item.shelfLifeDate && item.shelfLifeDate !== '-' ? item.shelfLifeDate : '-'
);
aoaData.push(row);
});
} else {
// Sarf & Parça Deposu Başlıkları (Yalnızca seçili bölge sütunlarını içerir)
const headerRow: string[] = [
"SIRA NO",
"DESCRIPTION",
"PART NUMBER",
"SERİ NUMBER - MÜKERRER NO - AÇIKLAMA",
"LOKASYON NO",
"GELEN",
"TOPLAM STOK"
];
if (reg === 'ALL' || reg === 'ANKARA') {
headerRow.push("ANKARA ÇIKAN", "ANKARA MEVCUT");
}
if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
headerRow.push("KARAİN TRANSFER", "KARAİN ÇIKAN", "KARAİN MEVCUT");
}
if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
headerRow.push("ÇANAKKALE TRANSFER", "ÇANAKKALE ÇIKAN", "ÇANAKKALE MEVCUT");
}
if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
headerRow.push("MİLAS TRANSFER", "MİLAS ÇIKAN", "MİLAS MEVCUT");
}
if (reg === 'ALL' || reg === 'BURSA') {
headerRow.push("BURSA TRANSFER", "BURSA ÇIKAN", "BURSA MEVCUT");
}
if (reg === 'ALL') {
headerRow.push("MUAYENE GİDEN", "MUAYENE GELEN", "MUAYENE TOPLAM");
}
headerRow.push("RAF ÖMRÜ VAR MI?", "RAF ÖMRÜ BİTİŞ TARİHİ");
aoaData.push(headerRow);
// Veri Satırları
rowsToExport.forEach((item, idx) => {
const row: (string | number)[] = [
idx + 1,
item.description || item.name || '',
item.partNumber || item.pn || '-',
item.serialAndNotes || item.sn || '-',
item.lokasyonNo || item.location || '-',
item.gelen || 0,
item.toplamStok || 0
];
if (reg === 'ALL' || reg === 'ANKARA') {
row.push(item.ankaraCikan || 0, item.ankaraMevcut || 0);
}
if (reg === 'ALL' || reg === 'KARAİN' || reg === 'KARAIN') {
row.push(item.karainTransfer || 0, item.karainCikan || 0, item.karainMevcut || 0);
}
if (reg === 'ALL' || reg === 'ÇANAKKALE' || reg === 'CANAKKALE') {
row.push(item.canakkaleTransfer || 0, item.canakkaleCikan || 0, item.canakkaleMevcut || 0);
}
if (reg === 'ALL' || reg === 'MİLAS' || reg === 'MILAS') {
row.push(item.milasTransfer || 0, item.milasCikan || 0, item.milasMevcut || 0);
}
if (reg === 'ALL' || reg === 'BURSA') {
row.push(item.bursaTransfer || 0, item.bursaCikan || 0, item.bursaMevcut || 0);
}
if (reg === 'ALL') {
row.push(item.muayeneGiden || 0, item.muayeneGelen || 0, item.muayeneToplam || 0);
}
const hasLife = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
row.push(
hasLife ? 'EVET' : 'HAYIR',
hasLife && item.shelfLifeDate && item.shelfLifeDate !== '-' ? item.shelfLifeDate : '-'
);

        aoaData.push(row);
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(aoaData);
    const wb = XLSX.utils.book_new();
    const sheetName = isKimyasal ? "KIMYASAL_FILTRELENMIS" : "SARF_FILTRELENMIS";
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    let fileName = `${activeUnit}_${currentDepoType.toUpperCase()}`;
    if (reg !== 'ALL') {
      fileName += `_${reg}`;
    }
    if (selectedLocations.length > 0) {
      fileName += `_LOKASYON_FILTRELENMIS`;
    }
    fileName += `_${new Date().toISOString().slice(0, 10)}.xlsx`;

    XLSX.writeFile(wb, fileName);
    showNotification(`✅ ${rowsToExport.length} filtrelenmiş malzeme "${fileName}" olarak bilgisayarınıza indirildi.`);
  };

  // EXCEL YÜKLE - MEVCUT VERİ YÜKLEME GÜNCELLEME İLE BİREBİR UYUMLU VE DRIVE ENTEGRASYONU
  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    showNotification("Excel dosyası çözümleniyor ve Google Drive'a aktarılıyor...");

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

      // YAŞAM DESTEK EXCEL YÜKLEME VE SERİ NO ÇÖZÜMLEME MODÜLÜ
      if (selectedExcelUploadCategory === 'yasam_destek') {
        const canYelegiParsed: any[] = [];
        const spareAirParsed: any[] = [];
        const helmetKitParsed: any[] = [];

        workbook.SheetNames.forEach(sName => {
          const ws = workbook.Sheets[sName];
          if (!ws) return;
          unmergeAndFillWorksheet(ws);
          const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (!rows || rows.length < 1) return;

          const sUpper = sName.toUpperCase().replace(/İ/g, 'I');
          const isCanYelegi = sUpper.includes('CAN YELEG') || sUpper.includes('LIFE VEST') || sUpper.includes('YELEK');
          const isSpareAir = sUpper.includes('SPARE AIR') || sUpper.includes('HEED') || sUpper.includes('YEDEK HAVA');
          const isHelmetKit = sUpper.includes('HELMET') || sUpper.includes('DAVID CLARK') || sUpper.includes('KASK') || sUpper.includes('DOLGU');

          // Başlık satırını bul
          let hIdx = 0;
          for (let r = 0; r < Math.min(10, rows.length); r++) {
            const rowStr = (rows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
            if (rowStr.includes('MALZEME') || rowStr.includes('PARÇA') || rowStr.includes('SERİ') || rowStr.includes('S. NO') || rowStr.includes('P/N') || rowStr.includes('S/N') || rowStr.includes('DESCRIPTION')) {
              hIdx = r;
              break;
            }
          }

          const hRow = (rows[hIdx] || []).map(c => String(c || '').toUpperCase().trim().replace(/İ/g, 'I').replace(/Ç/g, 'C').replace(/Ğ/g, 'G').replace(/Ö/g, 'O').replace(/Ş/g, 'S').replace(/Ü/g, 'U'));

          // 1. CAN YELEĞİ STOK
          if (isCanYelegi || (!isSpareAir && !isHelmetKit && hRow.some(c => c.includes('DIS NO') || c.includes('DIŞ NO')))) {
            const sNoCol = hRow.findIndex(c => c.includes('S. NO') || c === 'NO' || c === 'SIRA');
            const malzCol = hRow.findIndex(c => c.includes('MALZEME') || c.includes('DESCRIPTION'));
            const pnCol = hRow.findIndex(c => c.includes('PARCA') || c.includes('P/N'));
            const snCol = hRow.findIndex(c => c.includes('SERI') || c.includes('S/N'));
            const disNoCol = hRow.findIndex(c => c.includes('DIS NO') || c.includes('DIŞ NO'));
            const acikCol = hRow.findIndex(c => c.includes('ACIKLAMA') || c.includes('DURUM') || c.includes('GOREV'));

            for (let i = hIdx + 1; i < rows.length; i++) {
              const r = rows[i];
              if (!r || r.every(cell => !cell || String(cell).trim() === '')) continue;
              const sn = String(r[snCol >= 0 ? snCol : 3] || '').trim();
              const malz = String(r[malzCol >= 0 ? malzCol : 1] || 'LIFE WEST - CAN YELEĞİ').trim();
              if (!sn && !malz) continue;

              const acik = String(r[acikCol >= 0 ? acikCol : 5] || '').trim();
              const disNo = String(r[disNoCol >= 0 ? disNoCol : 4] || '-').trim();
              const pn = String(r[pnCol >= 0 ? pnCol : 2] || 'S-7200-511').trim();

              let durum = 'GÖREV BÖLGESİNDE';
              let color = '#92d050';
              const upperAcik = (acik + ' ' + malz).toUpperCase().replace(/İ/g, 'I');
              if (upperAcik.includes('KAYIT SILME') || upperAcik.includes('SİL')) {
                durum = 'KAYIT SİLME';
                color = '#ff0000';
              } else if (upperAcik.includes('BAKIM') || upperAcik.includes('DOLUM') || upperAcik.includes('ARIZA')) {
                durum = 'BAKIMA GİDECEK';
                color = '#00b0f0';
              } else if (upperAcik.includes('ONARIM')) {
                durum = 'ONARIMDA';
                color = '#ffff00';
              } else if (upperAcik.includes('DEPODA') || upperAcik.includes('BOSTA') || upperAcik.includes('FAAL')) {
                durum = 'BOŞTA';
                color = '#92d050';
              }

              canYelegiParsed.push({
                sNo: Number(r[sNoCol >= 0 ? sNoCol : 0]) || canYelegiParsed.length + 1,
                malzemeAdi: malz || 'LIFE WEST - CAN YELEĞİ',
                parcaNo: pn || 'S-7200-511',
                seriNo: sn || '-',
                disNo: disNo || '-',
                aciklamalar: acik,
                durum,
                color
              });
            }
          } 
          // 2. SPARE AIR STOK
          else if (isSpareAir || (!isCanYelegi && !isHelmetKit && hRow.some(c => c.includes('HEED') || c.includes('SPARE')))) {
            const sNoCol = hRow.findIndex(c => c.includes('S. NO') || c === 'NO' || c === 'SIRA');
            const malzCol = hRow.findIndex(c => c.includes('MALZEME') || c.includes('DESCRIPTION'));
            const pnCol = hRow.findIndex(c => c.includes('PARCA') || c.includes('P/N'));
            const snCol = hRow.findIndex(c => c.includes('SERI') || c.includes('S/N'));
            const acikCol = hRow.findIndex(c => c.includes('ACIKLAMA') || c.includes('DURUM') || c.includes('GOREV'));

            for (let i = hIdx + 1; i < rows.length; i++) {
              const r = rows[i];
              if (!r || r.every(cell => !cell || String(cell).trim() === '')) continue;
              const sn = String(r[snCol >= 0 ? snCol : 3] || '').trim();
              const malz = String(r[malzCol >= 0 ? malzCol : 1] || 'SPARE AIR (HEED 3)').trim();
              if (!sn && !malz) continue;

              const acik = String(r[acikCol >= 0 ? acikCol : 4] || '').trim();
              const pn = String(r[pnCol >= 0 ? pnCol : 2] || '175-001-CE').trim();

              let durum = 'GÖREV BÖLGESİNDE';
              let color = '#92d050';
              const upperAcik = (acik + ' ' + malz).toUpperCase().replace(/İ/g, 'I');
              if (upperAcik.includes('KAYIT SILME') || upperAcik.includes('SİL')) {
                durum = 'KAYIT SİLME';
                color = '#ff0000';
              } else if (upperAcik.includes('BAKIM') || upperAcik.includes('DOLUM') || upperAcik.includes('ARIZA')) {
                durum = 'BAKIMA GİDECEK';
                color = '#00b0f0';
              } else if (upperAcik.includes('ONARIM')) {
                durum = 'ONARIMDA';
                color = '#ffff00';
              } else if (upperAcik.includes('DEPODA') || upperAcik.includes('BOSTA') || upperAcik.includes('FAAL')) {
                durum = 'BOŞTA';
                color = '#92d050';
              }

              spareAirParsed.push({
                sNo: Number(r[sNoCol >= 0 ? sNoCol : 0]) || spareAirParsed.length + 1,
                malzemeAdi: malz || 'SPARE AIR (HEED 3)',
                parcaNo: pn || '175-001-CE',
                seriNo: sn || '-',
                aciklamalar: acik,
                durum,
                color
              });
            }
          } 
          // 3. HELMET KIT VE DAVID CLARK KULAK VE DOLGU
          else if (isHelmetKit || hRow.some(c => c.includes('HELMET') || c.includes('DAVID CLARK') || c.includes('TARIH'))) {
            const sNoCol = hRow.findIndex(c => c.includes('S. NO') || c === 'NO' || c === 'SIRA');
            const acikCol = hRow.findIndex(c => c.includes('ACIKLAMA') || c.includes('PERSONEL') || c.includes('PILOT'));
            const tarihCol = hRow.findIndex(c => c.includes('TARIH') || c.includes('DATE'));

            for (let i = hIdx + 1; i < rows.length; i++) {
              const r = rows[i];
              if (!r || r.every(cell => !cell || String(cell).trim() === '')) continue;
              const acik = String(r[acikCol >= 0 ? acikCol : 7] || r[4] || r[1] || '').trim();
              if (!acik && !r[1]) continue;

              const malz1 = String(r[1] || 'HELMET KIT').trim();
              const pn1 = String(r[2] || '18852G-01').trim();
              const sn1 = String(r[3] || '-').trim();
              const malz2 = String(r[4] || 'DAVID CLARK').trim();
              const pn2 = String(r[5] || 'H-10-13X').trim();
              const sn2 = String(r[6] || '-').trim();
              const tarih = String(r[tarihCol >= 0 ? tarihCol : 8] || '-').trim();

              let durum = 'GÖREV BÖLGESİNDE';
              let color = '#92d050';
              const upperAcik = (acik + ' ' + malz1).toUpperCase().replace(/İ/g, 'I');
              if (upperAcik.includes('DEPODA') || upperAcik.includes('BOSTA')) {
                durum = 'BOŞTA';
                color = '#92d050';
              }

              helmetKitParsed.push({
                sNo: Number(r[sNoCol >= 0 ? sNoCol : 0]) || helmetKitParsed.length + 1,
                malzemeAdi1: malz1 || 'HELMET KIT',
                parcaNo1: pn1 || '18852G-01',
                seriNo1: sn1 || '-',
                malzemeAdi2: malz2 || 'DAVID CLARK',
                parcaNo2: pn2 || 'H-10-13X',
                seriNo2: sn2 || '-',
                aciklamalar: acik,
                tarih: tarih || '-',
                durum,
                color
              });
            }
          }
        });

        // Convert array buffer to base64 & upload to Drive
        let binary = '';
        const bytes = new Uint8Array(arrayBuffer);
        for (let i = 0; i < bytes.byteLength; i++) {
          binary += String.fromCharCode(bytes[i]);
        }
        const base64Data = btoa(binary);

        try {
          await fetch('/api/upload-techizat-excel', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              fileName: 'at-802_yasam_destek_ekipmanlari.xlsx',
              targetKey: 'depo_at802_yasam_destek',
              base64Data: base64Data,
              folderId: DRIVE_FOLDER_ID
            })
          });
        } catch (e) {}

        // Backend'e aktar ve Google E-Tablo senkronizasyonunu tetikle
        await fetch('/api/yasam-destek/upload-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            canYelegi: canYelegiParsed,
            spareAir: spareAirParsed,
            helmetKit: helmetKitParsed,
            base64Data
          })
        });

        await fetchYasamDestekData();
        setCurrentDepoType('yasam_destek');
        showNotification(`✅ Yaşam Destek Exceli Başarıyla Yüklendi! (Can Yeleği: ${canYelegiParsed.length}, Spare Air: ${spareAirParsed.length}, Helmet Kit: ${helmetKitParsed.length} adet serno ve stok aktarıldı)`);
        if (fileInputRef.current) fileInputRef.current.value = '';
        return;
      }

      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Birleştirilmiş hücreleri unmerge yap ve değerleri yay
      unmergeAndFillWorksheet(sheet);

      const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!rawRows || rawRows.length < 2) {
        alert("Excel dosyasında veri bulunamadı.");
        return;
      }

      // Başlık satırını bul
      let headerRowIdx = 0;
      for (let r = 0; r < Math.min(15, rawRows.length); r++) {
        const rowStr = (rawRows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
        if (
          rowStr.includes('MALZEME') ||
          rowStr.includes('DESCRIPTION') ||
          rowStr.includes('PARÇA') ||
          rowStr.includes('P/N') ||
          rowStr.includes('PART')
        ) {
          headerRowIdx = r;
          break;
        }
      }

      const headerRow = (rawRows[headerRowIdx] || []).map(c => String(c || '').toUpperCase().trim());
      const normHeader = headerRow.map(c => c.replace(/Ç/g, 'C').replace(/Ğ/g, 'G').replace(/İ/g, 'I').replace(/Ö/g, 'O').replace(/Ş/g, 'S').replace(/Ü/g, 'U'));

      const isKimyasalUpload = selectedExcelUploadCategory === 'kimyasal';

      let descCol = normHeader.findIndex(c => c === 'DESCRIPTION' || c.includes('DESCRIPTION') || c.includes('MALZEME'));
      let pnCol = normHeader.findIndex(c => c.includes('P/N') || c.includes('PART NUMBER') || c.includes('PART'));
      let spainDescCol = normHeader.findIndex(c => c.includes('SPAIN DISCRIPTION') || c.includes('SPAIN DESCRIPTION'));
      let miktarCol = normHeader.findIndex(c => c === 'MIKTAR' || c.includes('MIKTAR QTY') || c === 'BIRIM');
      let sozlesmeCol = normHeader.findIndex(c => c.includes('SOZLESME'));
      let piyasaCol = normHeader.findIndex(c => c.includes('PIYASA'));
      let spainGelenCol = normHeader.findIndex(c => c.includes('SPAIN GELEN'));
      let genelToplamCol = normHeader.findIndex(c => c.includes('GENEL TOPLAM'));
      let toplamStokCol = normHeader.findIndex(c => c.includes('TOPLAM STOK'));
      let ankaraCikanCol = normHeader.findIndex(c => c.includes('ANKARA CIKAN'));
      let ankaraMevcutCol = normHeader.findIndex(c => c.includes('ANKARA MEVCUT'));
      let karainTransferCol = normHeader.findIndex(c => c.includes('KARAIN TRANSFER'));
      let karainCikanCol = normHeader.findIndex(c => c.includes('KARAIN CIKAN'));
      let karainMevcutCol = normHeader.findIndex(c => c.includes('KARAIN MEVCUT'));
      let canakkaleTransferCol = normHeader.findIndex(c => c.includes('CANAKKALE TRANSFER'));
      let canakkaleCikanCol = normHeader.findIndex(c => c.includes('CANAKKALE CIKAN'));
      let canakkaleMevcutCol = normHeader.findIndex(c => c.includes('CANAKKALE MEVCUT'));
      let milasTransferCol = normHeader.findIndex(c => c.includes('MILAS TRANSFER'));
      let milasCikanCol = normHeader.findIndex(c => c.includes('MILAS CIKAN'));
      let milasMevcutCol = normHeader.findIndex(c => c.includes('MILAS MEVCUT'));
      let bursaTransferCol = normHeader.findIndex(c => c.includes('BURSA TRANSFER'));
      let bursaCikanCol = normHeader.findIndex(c => c.includes('BURSA CIKAN'));
      let bursaMevcutCol = normHeader.findIndex(c => c.includes('BURSA MEVCUT'));
      let aciklamaCol = normHeader.findIndex(c => c === 'ACIKLAMA' || c.includes('ACIKLAMA'));
      let locCol = normHeader.findIndex(c => c.includes('LOKASYON') || c.includes('RAF') || c.includes('KONUM'));
      let shelfLifeCol = normHeader.findIndex(c => c.includes('RAF OMRU VAR') || c.includes('OMRUNE TABI'));
      let shelfLifeDateCol = normHeader.findIndex(c => c.includes('RAF OMRU BITIS') || c.includes('SKT'));

      // If Kimyasal Upload, apply exact 24-column Kimyasal positional fallbacks if header index missing:
      if (isKimyasalUpload) {
        if (descCol < 0) descCol = 1;
        if (pnCol < 0) pnCol = 2;
        if (spainDescCol < 0) spainDescCol = 2;
        if (miktarCol < 0) miktarCol = 3;
        if (sozlesmeCol < 0) sozlesmeCol = 4;
        if (piyasaCol < 0) piyasaCol = 5;
        if (spainGelenCol < 0) spainGelenCol = 6;
        if (genelToplamCol < 0) genelToplamCol = 7;
        if (toplamStokCol < 0) toplamStokCol = 8;
        if (ankaraCikanCol < 0) ankaraCikanCol = 9;
        if (ankaraMevcutCol < 0) ankaraMevcutCol = 10;
        if (karainTransferCol < 0) karainTransferCol = 11;
        if (karainCikanCol < 0) karainCikanCol = 12;
        if (karainMevcutCol < 0) karainMevcutCol = 13;
        if (canakkaleTransferCol < 0) canakkaleTransferCol = 14;
        if (canakkaleCikanCol < 0) canakkaleCikanCol = 15;
        if (canakkaleMevcutCol < 0) canakkaleMevcutCol = 16;
        if (milasTransferCol < 0) milasTransferCol = 17;
        if (milasCikanCol < 0) milasCikanCol = 18;
        if (milasMevcutCol < 0) milasMevcutCol = 19;
        if (bursaTransferCol < 0) bursaTransferCol = 20;
        if (bursaCikanCol < 0) bursaCikanCol = 21;
        if (bursaMevcutCol < 0) bursaMevcutCol = 22;
        if (aciklamaCol < 0) aciklamaCol = 23;
      } else {
        if (descCol < 0) descCol = 0;
        if (pnCol < 0) pnCol = 1;
      }

      const unitKey = currentUnit !== 'all' ? currentUnit : 'at802';
      const parsedItems: DepoItem[] = [];

      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;
        const desc = String(row[descCol >= 0 ? descCol : 1] || row[0] || '').trim();
        if (!desc || isHeaderLikeRow(row) || desc.toUpperCase() === 'DESCRIPTION' || desc.toUpperCase() === 'MALZEME ADI') continue;

        const pn = String(row[pnCol >= 0 ? pnCol : 2] !== undefined && row[pnCol >= 0 ? pnCol : 2] !== null ? row[pnCol >= 0 ? pnCol : 2] : '-').trim() || '-';
        const spainDesc = spainDescCol >= 0 ? String(row[spainDescCol] !== undefined ? row[spainDescCol] : pn).trim() : pn;

        // MİKTAR column contains text like "Adet", "Lt", "Kg"
        const rawMiktarVal = String(row[miktarCol >= 0 ? miktarCol : 3] !== undefined ? row[miktarCol >= 0 ? miktarCol : 3] : 'Adet').trim();
        const miktarBirim = rawMiktarVal || 'Adet';

        const sozlesmeToplamiVal = row[sozlesmeCol >= 0 ? sozlesmeCol : 4] !== undefined ? row[sozlesmeCol >= 0 ? sozlesmeCol : 4] : '0';
        const sozlesmeToplami = String(sozlesmeToplamiVal !== undefined ? sozlesmeToplamiVal : '0').trim();

        const piyasaGelen = Number(row[piyasaCol >= 0 ? piyasaCol : 5]) || 0;
        const spainGelen = Number(row[spainGelenCol >= 0 ? spainGelenCol : 6]) || 0;
        const genelToplamVal = Number(row[genelToplamCol >= 0 ? genelToplamCol : 7]);
        const genelToplam = !isNaN(genelToplamVal) && genelToplamVal > 0 ? genelToplamVal : (piyasaGelen + spainGelen);

        const ankaraCikan = Number(row[ankaraCikanCol >= 0 ? ankaraCikanCol : 9]) || 0;
        const ankaraMevcut = Number(row[ankaraMevcutCol >= 0 ? ankaraMevcutCol : 10]) || 0;

        const karainTransfer = Number(row[karainTransferCol >= 0 ? karainTransferCol : 11]) || 0;
        const karainCikan = Number(row[karainCikanCol >= 0 ? karainCikanCol : 12]) || 0;
        const karainMevcut = Number(row[karainMevcutCol >= 0 ? karainMevcutCol : 13]) || 0;

        const canakkaleTransfer = Number(row[canakkaleTransferCol >= 0 ? canakkaleTransferCol : 14]) || 0;
        const canakkaleCikan = Number(row[canakkaleCikanCol >= 0 ? canakkaleCikanCol : 15]) || 0;
        const canakkaleMevcut = Number(row[canakkaleMevcutCol >= 0 ? canakkaleMevcutCol : 16]) || 0;

        const milasTransfer = Number(row[milasTransferCol >= 0 ? milasTransferCol : 17]) || 0;
        const milasCikan = Number(row[milasCikanCol >= 0 ? milasCikanCol : 18]) || 0;
        const milasMevcut = Number(row[milasMevcutCol >= 0 ? milasMevcutCol : 19]) || 0;

        const bursaTransfer = Number(row[bursaTransferCol >= 0 ? bursaTransferCol : 20]) || 0;
        const bursaCikan = Number(row[bursaCikanCol >= 0 ? bursaCikanCol : 21]) || 0;
        const bursaMevcut = Number(row[bursaMevcutCol >= 0 ? bursaMevcutCol : 22]) || 0;

        const aciklama = String(row[aciklamaCol >= 0 ? aciklamaCol : 23] !== undefined ? row[aciklamaCol >= 0 ? aciklamaCol : 23] : '-').trim() || '-';
        const loc = String(row[locCol >= 0 ? locCol : 3] !== undefined ? row[locCol >= 0 ? locCol : 3] : 'KİMYASAL DEPO').trim() || 'KİMYASAL DEPO';

        const computedToplam = (toplamStokCol >= 0 && Number(row[toplamStokCol]) > 0)
          ? Number(row[toplamStokCol])
          : ((ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut) || genelToplam);

        parsedItems.push({
          unit: unitKey,
          category: selectedExcelUploadCategory,
          description: desc,
          name: desc,
          partNumber: pn,
          pn: pn,
          spainDescription: spainDesc,
          miktarQty: miktarBirim,
          sozlesmeToplami: sozlesmeToplami,
          piyasa: piyasaGelen,
          spainGelen: spainGelen,
          genelToplam: genelToplam,
          toplamStok: computedToplam,
          ankaraCikan: ankaraCikan,
          ankaraMevcut: ankaraMevcut,
          karainTransfer: karainTransfer,
          karainCikan: karainCikan,
          karainMevcut: karainMevcut,
          canakkaleTransfer: canakkaleTransfer,
          canakkaleCikan: canakkaleCikan,
          canakkaleMevcut: canakkaleMevcut,
          milasTransfer: milasTransfer,
          milasCikan: milasCikan,
          milasMevcut: milasMevcut,
          bursaTransfer: bursaTransfer,
          bursaCikan: bursaCikan,
          bursaMevcut: bursaMevcut,
          serialAndNotes: aciklama,
          sn: aciklama,
          lokasyonNo: loc,
          location: loc,
          baseGelen: genelToplam,
          gelen: genelToplam,
          hasShelfLife: shelfLifeCol >= 0 ? (String(row[shelfLifeCol] || '').toUpperCase().includes('EVET') ? 'EVET' : 'HAYIR') : 'HAYIR',
          shelfLifeDate: shelfLifeDateCol >= 0 ? String(row[shelfLifeDateCol] || '-').trim() : '-'
        });
      }

      if (parsedItems.length > 0) {
        setInventory(prev => {
          const filtered = prev.filter(i => !(i.unit === unitKey && i.category === selectedExcelUploadCategory));
          const updated = [...parsedItems, ...filtered];
          localStorage.setItem('ogm_depo_inventory_v5', JSON.stringify(updated));
          return updated;
        });
      }

      // Convert array buffer to base64 & upload to Drive (Mevcut Veri Güncelleme Mantığı)
      let binary = '';
      const bytes = new Uint8Array(arrayBuffer);
      for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64Data = btoa(binary);

      const standardFileName = getDepoStandardFileName(unitKey, selectedExcelUploadCategory);

      try {
        await fetch('/api/upload-techizat-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: standardFileName,
            targetKey: `depo_${unitKey}_${currentDepoType}`,
            base64Data: base64Data,
            folderId: DRIVE_FOLDER_ID
          })
        });
      } catch (driveErr) {
        console.warn('Drive yükleme uyarısı:', driveErr);
      }

      showNotification(`✅ "${standardFileName}" Google Drive'a kaydedildi ve ${parsedItems.length} kayıt envantere aktarıldı!`);
      if (fileInputRef.current) fileInputRef.current.value = '';
    } catch (err) {
      alert("Excel yükleme hatası: " + (err as Error).message);
    }
  };

  // =========================================================================
  // EXCEL ŞABLONU İLE TOPLU GİRİŞ / ÇIKIŞ MODÜLÜ METOTLARI
  // =========================================================================

  // 1. Örnek Şablon İndir (.xlsx) - Visual Standard Headers
  const handleDownloadBulkTemplate = () => {
    try {
      const headers = [
        ["MALZEME ADI", "ADET", "TARİH", "İŞLEM TÜRÜ", "SERİAL NUMBER", "KUYRUK KODU", "TESLİM ALAN", "KABUL YAPAN"]
      ];
      const examples = [
        ["101-50272", 2, new Date().toLocaleDateString('tr-TR'), "KARAİN TRANSFER", "MS28775-272", "ORMAN 21 (OR-2021) - AT-802", "Yetkili Teknisyen", "Depo Sorumlusu"]
      ];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([...headers, ...examples]);
      XLSX.utils.book_append_sheet(wb, ws, "SABLON");
      XLSX.writeFile(wb, "OGM_Toplu_Islem_Sablonu.xlsx");
      showNotification("✅ Örnek Toplu İşlem Şablonu indirildi.");
    } catch (err: any) {
      alert("Şablon indirme hatası: " + err.message);
    }
  };

  // 2. Excel Dosyası Çözümleme ve Otomatik Akıllı Eşleştirme (PN ve SN Taraması)
  const handleBulkExcelParse = async (file: File) => {
    if (!file) return;
    showNotification("Toplu işlem dosyası çözümleniyor...");
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      if (!rawRows || rawRows.length < 2) {
        alert("Excel dosyasında veri bulunamadı.");
        return;
      }

      // Sütun indekslerini tespit et
      let headerRowIdx = 0;
      for (let r = 0; r < Math.min(10, rawRows.length); r++) {
        const rowStr = (rawRows[r] || []).map(c => String(c || '').toUpperCase().trim()).join(' ');
        if (rowStr.includes('PN') || rowStr.includes('SN') || rowStr.includes('PART') || rowStr.includes('ADET') || rowStr.includes('MİKTAR') || rowStr.includes('MALZEME') || rowStr.includes('SERIAL')) {
          headerRowIdx = r;
          break;
        }
      }

      const headerRow = (rawRows[headerRowIdx] || []).map(c => String(c || '').toUpperCase().trim());
      const pnCol = headerRow.findIndex(c => c === 'PN' || c.includes('PART') || c.includes('P/N') || c === 'MALZEME ADI' || c === 'MALZEME');
      const snCol = headerRow.findIndex(c => c === 'SN' || c.includes('SERİ') || c.includes('SERI') || c.includes('S/N') || c === 'SERIAL NUMBER' || c === 'SERİAL NUMBER');
      const qtyCol = headerRow.findIndex(c => c.includes('ADET') || c.includes('MİKTAR') || c.includes('MIKTAR') || c.includes('QTY') || c.includes('SAYI') || c.includes('KAYIT') || c === 'ADET/MİKTAR');
      
      const tarihCol = headerRow.findIndex(c => c === 'TARİH' || c === 'TARIH');
      const islemCol = headerRow.findIndex(c => c === 'İŞLEM TÜRÜ' || c === 'ISLEM TURU' || c === 'İŞLEM' || c === 'ISLEM');
      const kuyrukCol = headerRow.findIndex(c => c === 'KUYRUK KODU' || c === 'KUYRUK');
      const teslimCol = headerRow.findIndex(c => c === 'TESLİM ALAN' || c === 'TESLIM ALAN');
      const kabulCol = headerRow.findIndex(c => c === 'KABUL YAPAN' || c === 'KABUL');

      const resolvedPnCol = pnCol >= 0 ? pnCol : 0;
      const resolvedSnCol = snCol >= 0 ? snCol : (snCol === -1 && pnCol === 0 ? 4 : 1);
      const resolvedQtyCol = qtyCol >= 0 ? qtyCol : (qtyCol === -1 && pnCol === 0 ? 1 : 2);

      const parsed: any[] = [];
      let rowId = 1;

      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;
        const pnVal = String(row[resolvedPnCol] !== undefined ? row[resolvedPnCol] : '').trim();
        if (!pnVal || pnVal.toUpperCase() === 'PN' || pnVal.toUpperCase() === 'PART NUMBER' || pnVal.toUpperCase() === 'MALZEME ADI') continue;

        const snVal = String(row[resolvedSnCol] !== undefined ? row[resolvedSnCol] : '-').trim() || '-';
        const qtyVal = parseInt(String(row[resolvedQtyCol] || '1').replace(/[^0-9]/g, '')) || 1;

        const dateVal = tarihCol >= 0 && row[tarihCol] ? String(row[tarihCol]).trim() : '';
        const typeVal = islemCol >= 0 && row[islemCol] ? String(row[islemCol]).trim() : '';
        const tailVal = kuyrukCol >= 0 && row[kuyrukCol] ? String(row[kuyrukCol]).trim() : '';
        const teslimVal = teslimCol >= 0 && row[teslimCol] ? String(row[teslimCol]).trim() : '';
        const kabulVal = kabulCol >= 0 && row[kabulCol] ? String(row[kabulCol]).trim() : '';

        // Akıllı Eşleştirme Taraması: computedInventory içinden eşleştir
        // Sadece geçerli unit ve kategoriyle sınırla
        const unitKey = currentUnit !== 'all' ? currentUnit : 'at802';
        const categoryKey = currentDepoType === 'kimyasal' ? 'kimyasal' : 'sarf';

        const possibleMatches = computedInventory.filter(i => i.unit === unitKey && i.category === categoryKey);

        // PN ve SN tam uyan
        let matched = possibleMatches.find(i => {
          const itemPn = (i.partNumber || i.pn || '').trim().toLowerCase();
          const itemSn = (i.serialAndNotes || i.sn || '').trim().toLowerCase();
          return itemPn === pnVal.toLowerCase() && itemSn === snVal.toLowerCase();
        }) || null;

        // Bulunamazsa sadece PN uyan
        if (!matched) {
          matched = possibleMatches.find(i => {
            const itemPn = (i.partNumber || i.pn || '').trim().toLowerCase();
            return itemPn === pnVal.toLowerCase();
          }) || null;
        }

        parsed.push({
          id: rowId++,
          pn: pnVal,
          sn: snVal,
          qty: qtyVal,
          matchedItem: matched,
          status: matched ? 'Eşleşti' : 'Eşleşmedi',
          date: dateVal,
          type: typeVal,
          tailNo: tailVal,
          operator: teslimVal,
          receivedBy: kabulVal
        });
      }

      setBulkExcelRows(parsed);
      // Arama kutuları için ilk değerleri ata
      const initialQueries: Record<number, string> = {};
      parsed.forEach(p => {
        initialQueries[p.id] = p.matchedItem ? `${p.matchedItem.description} (P/N: ${p.matchedItem.partNumber})` : '';
      });
      setBulkSearchQueries(initialQueries);

      showNotification(`📊 Toplu işlem için ${parsed.length} satır çözümlendi.`);
    } catch (err) {
      alert("Excel okuma hatası: " + err.message);
    }
  };

  // 3. Manuel Düzeltme ve Eşleştirme Değişimi
  const handleBulkRowMatchSelect = (rowId: number, item: DepoItem) => {
    setBulkExcelRows(prev => prev.map(row => {
      if (row.id === rowId) {
        return {
          ...row,
          matchedItem: item,
          status: 'Eşleşti'
        };
      }
      return row;
    }));
    setBulkSearchQueries(prev => ({
      ...prev,
      [rowId]: `${item.description || item.name} (P/N: ${item.partNumber || item.pn || '-'})`
    }));
    setBulkDropdownOpenRow(null);
  };

  // 4. Toplu Onay ve Kaydet Mekanizması (Giriş veya Çıkış olarak işler)
  const handleSaveBulkExcelTransactions = async () => {
    const unmergedRows = bulkExcelRows.filter(r => !r.matchedItem);
    if (unmergedRows.length > 0) {
      alert(`⚠️ Hata: Henüz eşleşmeyen ${unmergedRows.length} kayıt bulunuyor. Lütfen listeden manuel arama yaparak eşleştiriniz.`);
      return;
    }

    const updatedInventory = [...inventory];
    const newTransactions: DepoTransaction[] = [];
    const dateStr = new Date().toLocaleDateString('tr-TR');
    const targetSheet = getDepoSheetNameForUnit(currentUnit);

    for (const row of bulkExcelRows) {
      const item = row.matchedItem!;
      const qty = row.qty;

      // Stok Limit Kontrolleri (Çıkış İşlemlerinde)
      if (movementDirection === 'cikis') {
        const info = getCikisDepotInfo(movementType, item);
        if (qty > info.stock) {
          alert(`❌ Toplu Çıkış Engeli: "${item.description}" malzemesi için talep edilen adet: ${qty}, ancak ${info.name} stoğu: ${info.stock}. Lütfen adedi güncelleyin!`);
          return;
        }
      }

      // Envanter Güncelleme
      const targetPn = (item.partNumber || item.pn || '').trim().toLowerCase();
      const targetDesc = (item.description || item.name || '').trim().toLowerCase();

      const idx = updatedInventory.findIndex(invItem => {
        const invPn = (invItem.partNumber || invItem.pn || '').trim().toLowerCase();
        const invDesc = (invItem.description || invItem.name || '').trim().toLowerCase();
        return (targetPn && invPn === targetPn) || (targetDesc && invDesc === targetDesc);
      });

      if (idx >= 0) {
        const invItem = { ...updatedInventory[idx] };
        let gelen = Number(invItem.gelen || invItem.baseGelen || 0);
        let ankaraCikan = Number(invItem.ankaraCikan || 0);
        let ankaraMevcut = Number(invItem.ankaraMevcut || 0);
        let karainTransfer = Number(invItem.karainTransfer || 0);
        let karainCikan = Number(invItem.karainCikan || 0);
        let karainMevcut = Number(invItem.karainMevcut || 0);
        let canakkaleTransfer = Number(invItem.canakkaleTransfer || 0);
        let canakkaleCikan = Number(invItem.canakkaleCikan || 0);
        let canakkaleMevcut = Number(invItem.canakkaleMevcut || 0);
        let milasTransfer = Number(invItem.milasTransfer || 0);
        let milasCikan = Number(invItem.milasCikan || 0);
        let milasMevcut = Number(invItem.milasMevcut || 0);
        let bursaTransfer = Number(invItem.bursaTransfer || 0);
        let bursaCikan = Number(invItem.bursaCikan || 0);
        let bursaMevcut = Number(invItem.bursaMevcut || 0);
        let muayeneToplam = Number(invItem.muayeneToplam || 0);

        // Determine type and direction dynamically from standard column data or fallback
        const resolvedType = (row.type || movementType || '').toUpperCase().trim();
        const isRowGiris = resolvedType.includes('GİREN') || resolvedType.includes('GİRİŞ') || resolvedType.includes('GELEN');
        const isRowTransfer = resolvedType.includes('TRANSFER');

        if (isRowGiris) {
          gelen += qty;
          ankaraMevcut += qty;
        } else if (isRowTransfer) {
          if (resolvedType.includes('KARAİN')) {
            karainTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            karainMevcut += qty;
          } else if (resolvedType.includes('ÇANAKKALE')) {
            canakkaleTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            canakkaleMevcut += qty;
          } else if (resolvedType.includes('MİLAS')) {
            milasTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            milasMevcut += qty;
          } else if (resolvedType.includes('BURSA')) {
            bursaTransfer += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
            bursaMevcut += qty;
          } else {
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
          }
        } else {
          // Çıkış
          if (resolvedType.includes('KARAİN')) {
            karainCikan += qty;
            karainMevcut = Math.max(0, karainMevcut - qty);
          } else if (resolvedType.includes('ÇANAKKALE')) {
            canakkaleCikan += qty;
            canakkaleMevcut = Math.max(0, canakkaleMevcut - qty);
          } else if (resolvedType.includes('MİLAS')) {
            milasCikan += qty;
            milasMevcut = Math.max(0, milasMevcut - qty);
          } else if (resolvedType.includes('BURSA')) {
            bursaCikan += qty;
            bursaMevcut = Math.max(0, bursaMevcut - qty);
          } else {
            ankaraCikan += qty;
            ankaraMevcut = Math.max(0, ankaraMevcut - qty);
          }
        }

        const toplamStok = ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut + muayeneToplam;
        updatedInventory[idx] = {
          ...invItem,
          gelen,
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
        };
      }

      const resolvedDateVal = row.date || dateStr;
      const resolvedTypeVal = row.type || movementType;
      const resolvedTailVal = row.tailNo || (movementDirection === 'giris' ? '-' : movementTail);
      const resolvedOperatorVal = row.operator || movementTeslimAlan;
      const resolvedReceivedByVal = row.receivedBy || movementKabulYapan;

      // Depo Hareket Geçmişi kaydı oluştur
      newTransactions.push({
        id: `tx_bulk_${Date.now()}_${row.id}_${Math.random().toString(36).substr(2, 4)}`,
        timestamp: new Date().toISOString(),
        type: resolvedTypeVal,
        itemName: item.description || item.name || '',
        pn: item.partNumber || item.pn || '-',
        sn: row.sn || item.serialAndNotes || '-',
        quantity: qty,
        date: resolvedDateVal,
        tailNo: resolvedTailVal,
        operator: resolvedOperatorVal,
        receivedBy: resolvedReceivedByVal,
        location: movementDepoYeri,
        unit: currentUnit,
        category: item.category || (currentDepoType === 'kimyasal' ? 'kimyasal' : 'sarf'),
        sheetName: targetSheet,
        notes: `Toplu Excel ${movementDirection === 'giris' ? 'Girişi' : 'Çıkışı'} ile eklendi.`,
        isNewSessionTx: true
      });
    }

    setInventory(updatedInventory);
    const allTransfers = [...newTransactions, ...transactions];
    setTransactions(allTransfers);

    if (onAddTransaction) {
      newTransactions.forEach(t => onAddTransaction(t));
    }

    setIsBulkExcelModalOpen(false);
    setIsMovementModalOpen(false);
    showNotification(`✅ Toplu ${movementDirection === 'giris' ? 'Giriş' : 'Çıkış'} işlemi çözümlendi, ${bulkExcelRows.length} parça başarıyla işlendi ve Excel Online'a aktarıldı!`);

    // Otomatik Excel Online ve Drive Senkronizasyonu
    syncDepoToExcelOnline(updatedInventory, allTransfers);
  };

  // Excel İndir - Web HTML Tasarımını ve Renklerini Birebir Koruyarak Filtrelenmiş Tüm Satırları Excel'e Çevirir
  const handleExportToExcel = () => {
    try {
      const isKimya = currentDepoType === 'kimyasal';
      const isTx = currentDepoType === 'transactions';

      // 1. Dışa aktarılacak filtrelenmiş verileri al (Sayfalama kısıtlaması olmadan tüm filtrelenmiş veriler)
      let exportRowsCount = 0;
      let title = "";
      let htmlTable = "";

      if (isTx) {
        title = `${currentUnit.toUpperCase()}_Depo_Hareket_Gecmisi`;
        const list = filteredTransactions;
        exportRowsCount = list.length;

        htmlTable = `
          <table border="1" style="border-collapse:collapse; font-family:Calibri,'Segoe UI',Arial,sans-serif; font-size:11pt; width:100%;">
            <thead>
              <tr style="background-color:#94a3b8; color:#0f172a; font-weight:bold; height:32px; text-align:center;">
                <th style="border:1px solid #64748b; padding:6px; background-color:#cbd5e1;">SIRA</th>
                <th style="border:1px solid #64748b; padding:6px; text-align:left; background-color:#e2e8f0;">MALZEME ADI</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#f3e8ff;">KATEGORİ</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">ADET</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">TARİH</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">İŞLEM TÜRÜ</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">SERİAL NUMBER</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">KUYRUK KODU</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">TESLİM ALAN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">KABUL YAPAN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#fef08a;">DEPO YERİ</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#e2e8f0;">SAYFA / BİRİM</th>
              </tr>
            </thead>
            <tbody>
              ${list.map((t, idx) => {
                const locNorm = String(t.location || t.depoYeri || '').toUpperCase().replace(/İ/g, 'I');
                const sheetNorm = String(t.sheetName || '').toUpperCase().replace(/İ/g, 'I');
                const nameNorm = String(t.itemName || t.itemDesc || t.name || '').toUpperCase().replace(/İ/g, 'I');
                const pnNorm = String(t.pn || t.partNumber || '').toUpperCase().replace(/İ/g, 'I');
                const isItemKimya = (t.category === 'kimyasal') || 
                                   locNorm.includes('KIMYA') || 
                                   sheetNorm.includes('KIMYA') ||
                                   nameNorm.includes('AEROSHELL') || 
                                   nameNorm.includes('GREASE') || 
                                   nameNorm.includes('ENGINE OIL') || 
                                   nameNorm.includes('OIL') || 
                                   nameNorm.includes('FLUID') || 
                                   nameNorm.includes('HYDRAULIC') || 
                                   nameNorm.includes('BOYA') || 
                                   nameNorm.includes('YAG') ||
                                   pnNorm.includes('MIL-PRF') ||
                                   pnNorm.includes('MIL-H-') ||
                                   pnNorm.includes('MIL-G-');
                const resolvedLocation = isItemKimya ? 'KİMYASAL DEPO' : (t.location || t.depoYeri || '-');
                const isGiris = (t.type || '').includes('GİREN') || (t.type || '').includes('GİRİŞ') || (t.type || '').includes('GELEN') || (t.type || '').includes('YENİ ÜRÜN');
                const isTransfer = (t.type || '').includes('TRANSFER');
                const typeBg = isTransfer ? '#e0e7ff' : (isGiris ? '#dcfce7' : '#ffe4e6');
                const typeColor = isTransfer ? '#312e81' : (isGiris ? '#14532d' : '#881337');

                return `
                  <tr style="height:26px;">
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:#f8fafc;">${idx + 1}</td>
                    <td style="border:1px solid #cbd5e1; text-align:left; font-weight:bold; padding-left:6px; max-width:280px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${t.itemName || t.itemDesc || t.name || '-'}">${t.itemName || t.itemDesc || t.name || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${isItemKimya ? '#f3e8ff; color:#6b21a8;' : '#eff6ff; color:#1e40af;'}">${isItemKimya ? '🧪 KİMYASAL' : '📦 SARF/PARÇA'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold;">${t.quantity !== undefined ? t.quantity : (t.adet || 1)}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center;">${t.date || t.timestamp || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${typeBg}; color:${typeColor}">${t.type || t.islemTuru || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-family:Consolas,monospace;">${t.sn || t.serialNumber || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold;">${t.tailNo || t.kuyrukKodu || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:left; padding-left:6px;">${t.operator || t.teslimAlan || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:left; padding-left:6px;">${t.receivedBy || t.kabulYapan || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:#fef9c3;">${resolvedLocation}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-family:Consolas,monospace;">${t.sheetName || 'DEPO HAREKET GEÇMİŞİ'}</td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        `;
      } else {
        // SARF veya KİMYASAL DEPOSU (23 Sütunlu Canlı Web HTML Tasarımı)
        const targetList = isKimya ? filteredKimyasalRows : filteredSarfRows;
        exportRowsCount = targetList.length;
        title = isKimya ? `${currentUnit.toUpperCase()}_Kimyasal_Depo` : `${currentUnit.toUpperCase()}_Sarf_ve_Parca_Deposu`;

        htmlTable = `
          <table border="1" style="border-collapse:collapse; font-family:Calibri,'Segoe UI',Arial,sans-serif; font-size:10pt; width:100%;">
            <thead>
              <tr style="height:34px; font-weight:bold; text-align:center;">
                <th style="border:1px solid #64748b; padding:6px; background-color:#f1f5f9; color:#0f172a; min-width:45px;">SIRA NO</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#f1f5f9; color:#0f172a; text-align:left; min-width:220px;">DESCRIPTION</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#f1f5f9; color:#0f172a; text-align:left; font-family:Consolas,monospace; min-width:140px;">PART NUMBER</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#f1f5f9; color:#0f172a; text-align:left; min-width:180px;">SERİ NUMBER - MÜKERRER NO - AÇIKLAMA</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ffff00; color:#000000; font-weight:bold; min-width:110px;">LOKASYON NO</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#d9d9d9; color:#000000; min-width:65px;">GELEN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#00ffff; color:#000000; font-weight:bold; min-width:85px;">TOPLAM STOK</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff9900; color:#000000; min-width:75px;">ANKARA ÇIKAN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff9900; color:#000000; font-weight:bold; min-width:75px;">ANKARA MEVCUT</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ffff00; color:#000000; min-width:75px;">KARAİN TRANSFER</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ffff00; color:#000000; min-width:70px;">KARAİN ÇIKAN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ffff00; color:#000000; font-weight:bold; min-width:70px;">KARAİN MEVCUT</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#92d050; color:#000000; min-width:75px;">ÇANAKKALE TRANSFER</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#92d050; color:#000000; min-width:70px;">ÇANAKKALE ÇIKAN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#92d050; color:#000000; font-weight:bold; min-width:70px;">ÇANAKKALE MEVCUT</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff9900; color:#000000; min-width:75px;">MİLAS TRANSFER</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff9900; color:#000000; min-width:70px;">MİLAS ÇIKAN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff9900; color:#000000; font-weight:bold; min-width:70px;">MİLAS MEVCUT</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff00ff; color:#000000; min-width:75px;">BURSA TRANSFER</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff00ff; color:#000000; min-width:70px;">BURSA ÇIKAN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ff00ff; color:#000000; font-weight:bold; min-width:70px;">BURSA MEVCUT</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#bfbfbf; color:#000000; min-width:70px;">MUAYENE GİDEN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#bfbfbf; color:#000000; min-width:70px;">MUAYENE GELEN</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#bfbfbf; color:#000000; font-weight:bold; min-width:70px;">MUAYENE TOPLAM</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#d1fae5; color:#064e3b; font-weight:bold; min-width:95px;">RAF ÖMRÜ VAR MI?</th>
                <th style="border:1px solid #64748b; padding:6px; background-color:#ecfdf5; color:#064e3b; font-weight:bold; min-width:110px;">RAF ÖMRÜ BİTİŞ TARİHİ</th>
              </tr>
            </thead>
            <tbody>
              ${targetList.map((item, idx) => {
                const hasLife = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
                const lifeDate = hasLife && item.shelfLifeDate && item.shelfLifeDate !== '-' ? item.shelfLifeDate : '-';
                const ankaraZero = (item.ankaraMevcut || 0) <= 0 && hadRegionHistory(item, 'ANKARA');
                const karainZero = (item.karainMevcut || 0) <= 0 && hadRegionHistory(item, 'KARAİN');
                const canakkaleZero = (item.canakkaleMevcut || 0) <= 0 && hadRegionHistory(item, 'ÇANAKKALE');
                const milasZero = (item.milasMevcut || 0) <= 0 && hadRegionHistory(item, 'MİLAS');
                const bursaZero = (item.bursaMevcut || 0) <= 0 && hadRegionHistory(item, 'BURSA');

                return `
                  <tr style="height:26px;">
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:#f1f5f9;">${idx + 1}</td>
                    <td style="border:1px solid #cbd5e1; text-align:left; font-weight:bold; padding-left:6px;">${item.description || item.name || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:left; font-family:Consolas,monospace; font-weight:bold; padding-left:6px;">${item.partNumber || item.pn || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:left; padding-left:6px;">${item.serialAndNotes || item.sn || '-'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:#fef9c3;">${item.lokasyonNo || item.location || (isKimya ? 'KİMYASAL DEPO' : '-')}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#f1f5f9;">${item.gelen || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:#cffafe;">${item.toplamStok || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#fffbeb;">${item.ankaraCikan || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${ankaraZero ? '#e11d48; color:#ffffff;' : '#fef3c7; color:#78350f;'}">${item.ankaraMevcut || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#fef9c3;">${item.karainTransfer || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#fef9c3;">${item.karainCikan || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${karainZero ? '#e11d48; color:#ffffff;' : '#fef08a; color:#713f12;'}">${item.karainMevcut || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#f0fdf4;">${item.canakkaleTransfer || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#f0fdf4;">${item.canakkaleCikan || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${canakkaleZero ? '#e11d48; color:#ffffff;' : '#bbf7d0; color:#14532d;'}">${item.canakkaleMevcut || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#fff7ed;">${item.milasTransfer || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#fff7ed;">${item.milasCikan || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${milasZero ? '#e11d48; color:#ffffff;' : '#fed7aa; color:#7c2d12;'}">${item.milasMevcut || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#fdf4ff;">${item.bursaTransfer || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#fdf4ff;">${item.bursaCikan || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${bursaZero ? '#e11d48; color:#ffffff;' : '#f5d0fe; color:#701a75;'}">${item.bursaMevcut || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#f1f5f9;">${item.muayeneGiden || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; background-color:#f1f5f9;">${item.muayeneGelen || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:#e2e8f0;">${item.muayeneToplam || 0}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-weight:bold; background-color:${hasLife ? '#fef3c7; color:#92400e;' : '#f8fafc; color:#64748b;'}">${hasLife ? 'EVET' : 'HAYIR'}</td>
                    <td style="border:1px solid #cbd5e1; text-align:center; font-family:Consolas,monospace; font-weight:bold; color:#065f46;">${lifeDate}</td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        `;
      }

      // 2. Tam Web HTML Excel Belgesi oluşturma (Microsoft Excel & LibreOffice 100% renk ve tasarım korumalı açar)
      const fullHtmlDoc = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
          <!--[if gte mso 9]>
          <xml>
            <x:ExcelWorkbook>
              <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                  <x:Name>${title.slice(0, 31)}</x:Name>
                  <x:WorksheetOptions>
                    <x:DisplayGridlines/>
                  </x:WorksheetOptions>
                </x:ExcelWorksheet>
              </x:ExcelWorksheets>
            </x:ExcelWorkbook>
          </xml>
          <![endif]-->
          <style>
            table { border-collapse: collapse; font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 10pt; }
            th, td { border: 1px solid #cbd5e1; padding: 6px 8px; vertical-align: middle; }
          </style>
        </head>
        <body>
          <div style="margin-bottom:12px; font-family:sans-serif;">
            <h2 style="margin:0; color:#0f172a; font-size:16pt;">OGM ENVENTER VE DEPO YÖNETİM SİSTEMİ</h2>
            <p style="margin:4px 0 0 0; color:#475569; font-size:10pt;">
              Rapor: <strong>${title.replace(/_/g, ' ')}</strong> | İndirme Tarihi: <strong>${new Date().toLocaleString('tr-TR')}</strong> | Toplam Filtrelenmiş Kayıt: <strong>${exportRowsCount} adet</strong>
            </p>
          </div>
          ${htmlTable}
        </body>
        </html>
      `;

      const blob = new Blob([fullHtmlDoc], { type: 'application/vnd.ms-excel;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${title}_${new Date().toISOString().slice(0, 10)}.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      showNotification(`✅ ${exportRowsCount} filtrelenmiş kaydın tamamı orijinal renk ve tasarımlarıyla Excel (.xls) olarak indirildi!`);
    } catch (e: any) {
      alert("Excel İndirme Hatası: " + e.message);
    }
  };

  // E-Tabloya Kaydet
  const handleSyncWithGoogleSheet = async () => {
    setIsSyncingCloud(true);
    showNotification("Merkezi Google E-Tabloya veriler aktarılıyor...");
    try {
      const at802Items = inventory.filter(i => i.unit === 'at802' || i.unit === 'all');
      await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'saveAt802SarfDepo', items: at802Items })
      });

      await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'saveDepoTransfers', transfers: transactions })
      });

      showNotification("✅ Canlı E-Tablo başarıyla senkronize edildi!");
    } catch (err) {
      showNotification("⚠️ E-Tablo senkronizasyonunda uyarı oluştu.");
    } finally {
      setIsSyncingCloud(false);
    }
  };

  // New / Edit Item Submit - Updates state and syncs directly to online Excel in Google Drive
  const handleSaveItem = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newItemDesc.trim()) {
      alert("Lütfen malzeme adı giriniz!");
      return;
    }

    const itemObj: DepoItem = {
      unit: newItemUnit,
      category: newItemCategory,
      description: newItemDesc.trim(),
      name: newItemDesc.trim(),
      partNumber: newItemPn.trim() || '-',
      pn: newItemPn.trim() || '-',
      lokasyonNo: newItemCategory === 'kimyasal' ? (newItemLokasyon.trim() || 'KİMYASAL DEPO') : (newItemLokasyon.trim() || 'ANKARA RAF A-1'),
      location: newItemCategory === 'kimyasal' ? (newItemLokasyon.trim() || 'KİMYASAL DEPO') : (newItemLokasyon.trim() || 'ANKARA RAF A-1'),
      serialAndNotes: newItemSnNotes.trim() || '-',
      sn: newItemSnNotes.trim() || '-',
      baseGelen: newItemGelen,
      gelen: newItemGelen,
      hasShelfLife: newItemHasShelfLife,
      shelfLifeDate: newItemHasShelfLife === 'EVET' ? (newItemShelfLifeDate.trim() || '-') : '-'
    };

    const oldItem = editingItemIdx >= 0 ? inventory[editingItemIdx] : null;

    if (editingItemIdx >= 0) {
      if (!window.confirm("Yapılan değişiklikler kaydedilsin mi?")) return;
      setInventory(prev => {
        const updated = [...prev];
        updated[editingItemIdx] = itemObj;
        return updated;
      });
      showNotification("✅ Veri kaydedildi.");
    } else {
      setInventory(prev => [itemObj, ...prev]);
      
      // Otomatik Depo Hareket Geçmişi kaydı oluştur (Kullanıcı Kuralı: Yeni ürün eklendiğinde otomatik hareket kaydı eklenir)
      const pad = (n: number) => String(n).padStart(2, '0');
      const nowD = new Date();
      const formattedDate = `${pad(nowD.getDate())}.${pad(nowD.getMonth() + 1)}.${nowD.getFullYear()}`;
      const targetSheet = getDepoSheetNameForUnit(newItemUnit);
      const newTx: DepoTransaction = {
        id: 'tx_add_' + Date.now(),
        timestamp: `${formattedDate} ${pad(nowD.getHours())}:${pad(nowD.getMinutes())}`,
        date: formattedDate,
        type: 'GİRİŞ',
        itemDesc: itemObj.description,
        itemName: itemObj.description,
        name: itemObj.description,
        partNumber: itemObj.partNumber,
        pn: itemObj.partNumber,
        serialNumber: itemObj.serialAndNotes,
        sn: itemObj.serialAndNotes,
        quantity: itemObj.gelen || 1,
        adet: itemObj.gelen || 1,
        location: newItemCategory === 'kimyasal' ? 'KİMYASAL DEPO' : (itemObj.lokasyonNo || 'ANKARA DEPO'),
        depoYeri: newItemCategory === 'kimyasal' ? 'KİMYASAL DEPO' : (itemObj.lokasyonNo || 'ANKARA DEPO'),
        islemTuru: 'Yeni Ürün Kaydı',
        kuyrukKodu: '-',
        tailNo: '-',
        teslimAlan: 'SİSTEM',
        kabulYapan: 'DEPO YÖNETİCİSİ',
        operator: 'DEPO YÖNETİCİSİ',
        receivedBy: 'SİSTEM',
        unit: newItemUnit,
        category: newItemCategory,
        sheetName: targetSheet,
        isNewSessionTx: true
      };

      setTransactions(prev => [newTx, ...prev]);
      if (onAddTransaction) {
        onAddTransaction(newTx);
      }

      // Google E-Tablo ilgili Hareket Geçmişi sayfasına canlı kaydet
      fetch('/api/save-depo-transfers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          transfers: [newTx],
          sheetName: targetSheet,
          spreadsheetId: '17ScGYYx0erzDwHDk6RGiHOdJATdfmmExXFBY39dXpF0'
        })
      }).catch(e => console.warn("Save new item transfer log warn:", e));

      showNotification("✅ Veri kaydedildi.");
    }
    setIsNewItemModalOpen(false);

    // Call /api/update-depo-excel-item so online Excel in Drive & disk is updated in place
    const targetFileName = getDepoStandardFileName(newItemUnit, newItemCategory);

    try {
      const res = await fetch('/api/update-depo-excel-item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: targetFileName,
          unit: newItemUnit,
          category: newItemCategory,
          oldItem: oldItem,
          newItem: itemObj,
          action: editingItemIdx >= 0 ? 'update' : 'add',
          folderId: DRIVE_FOLDER_ID
        })
      });
      const data = await res.json();
      if (data && data.status === 'success') {
        showNotification(`✅ Veri kaydedildi.`);
      }
    } catch (err) {
      console.warn("Excel update warn:", err);
    }
  };

  // Depoda Mevcut Ürün Satırına Tıklandığında Düzenleme Açma Fonksiyonu
  const handleRowClick = (item: DepoItem) => {
    setRowEditItem(item);
    setRowEditPn(item.partNumber || item.pn || '');
    setRowEditSn(item.serialAndNotes || item.sn || '');
    setRowEditDescription(item.description || item.name || '');
    setRowEditLocation(item.lokasyonNo || item.location || '');
    setRowEditNotes(item.notes || (item as any).aciklama || '');
    const hasLife = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
    setRowEditHasShelfLife(hasLife ? 'EVET' : 'HAYIR');
    setRowEditShelfLifeDate(item.shelfLifeDate && item.shelfLifeDate !== '-' ? item.shelfLifeDate : '');
    setRowEditPassword('');
    setRowEditPasswordError('');
    setIsRowEditModalOpen(true);
  };

  // Satır Düzenleme Kaydetme (Şifre Korumalı)
  const handleSaveRowEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!rowEditItem) return;

    const u = (rowEditItem.unit || currentUnit || 'at802').toLowerCase();
    const unitConfig = UNITS.find(unit => unit.id === u);
    const requiredPassword = unitConfig ? unitConfig.password : '1839';
    const input = rowEditPassword.trim();

    if (
      input !== requiredPassword &&
      input !== '1839' &&
      input !== '1234' &&
      input !== '8902'
    ) {
      setRowEditPasswordError('Hatalı Şifre! Düzenleme yetkisi için lütfen birim şifresini veya yetkili şifreyi (1839) giriniz.');
      return;
    }

    const updatedItem: DepoItem = {
      ...rowEditItem,
      description: rowEditDescription.trim(),
      name: rowEditDescription.trim(),
      partNumber: rowEditPn.trim() || '-',
      pn: rowEditPn.trim() || '-',
      serialAndNotes: rowEditSn.trim() || '-',
      sn: rowEditSn.trim() || '-',
      lokasyonNo: rowEditLocation.trim() || '-',
      location: rowEditLocation.trim() || '-',
      notes: rowEditNotes.trim() || '-',
      hasShelfLife: rowEditHasShelfLife,
      shelfLifeDate: rowEditHasShelfLife === 'EVET' ? (rowEditShelfLifeDate.trim() || '-') : '-'
    };

    const oldTarget = rowEditItem;
    setInventory(prev => prev.map(i => {
      const isMatch = (i.partNumber || i.pn || '').trim().toLowerCase() === (oldTarget.partNumber || oldTarget.pn || '').trim().toLowerCase() &&
                      (i.description || i.name || '').trim().toLowerCase() === (oldTarget.description || oldTarget.name || '').trim().toLowerCase();
      return isMatch ? updatedItem : i;
    }));

    setIsRowEditModalOpen(false);
    setRowEditItem(null);
    showNotification('✅ Malzeme bilgileri başarıyla güncellendi.');

    // Online Excel ve yerel ayna senkronizasyonu
    const targetFileName = getDepoStandardFileName(updatedItem.unit || currentUnit, updatedItem.category || 'sarf');
    try {
      await fetch('/api/update-depo-excel-item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: targetFileName,
          unit: updatedItem.unit || currentUnit,
          category: updatedItem.category || 'sarf',
          oldItem: oldTarget,
          newItem: updatedItem,
          action: 'update',
          folderId: DRIVE_FOLDER_ID
        })
      });
    } catch (err) {
      console.warn("Row edit sync warn:", err);
    }
  };

  // Depo Hareket Geçmişi Satırı Silme Fonksiyonları (Şifre Korumalı)
  const handleOpenDeleteTx = (t: DepoTransaction) => {
    setPendingTxToDelete(t);
    setTxDeletePasswordInput('');
    setTxDeletePasswordError('');
    setIsTxDeleteModalOpen(true);
  };

  const handleConfirmDeleteTx = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!pendingTxToDelete) return;

    const u = (pendingTxToDelete.unit || currentUnit || 'hangar').toLowerCase();
    const unitConfig = UNITS.find(unit => unit.id === u);
    const requiredPassword = unitConfig ? unitConfig.password : '1839';
    const input = txDeletePasswordInput.trim();

    if (
      input === requiredPassword ||
      input === '1839' ||
      input === '1234' ||
      input === '8902'
    ) {
      const updated = transactions.filter(t => t.id !== pendingTxToDelete.id);
      setTransactions(updated);
      setIsTxDeleteModalOpen(false);
      setPendingTxToDelete(null);

      showNotification('🗑️ Hareket kaydı silindi. Bulut senkronizasyonu yapılıyor...');

      try {
        const res = await fetch('/api/delete-depo-transaction', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            id: pendingTxToDelete.id,
            tx: pendingTxToDelete,
            password: input,
            unit: u
          })
        });
        const data = await res.json();
        if (data.status === 'success') {
          showNotification('🗑️ Hareket kaydı sistemden ve Google Sheets tablosundan kalıcı olarak silindi.');
        } else {
          showNotification('⚠️ ' + (data.message || 'Silme işlemi tamamlanamadı.'));
        }
      } catch (err) {
        console.warn("Delete tx sync warn:", err);
      }
    } else {
      setTxDeletePasswordError('Hatalı Şifre! Lütfen uçak biriminin şifresini veya yetkili şifreyi (1839) giriniz.');
    }
  };

  const handleOpenEditItem = (item: DepoItem, idx?: number) => {
    let targetIdx = idx !== undefined && idx >= 0 ? idx : inventory.findIndex(i => {
      const matchPn = (i.partNumber || i.pn || '').trim().toLowerCase() === (item.partNumber || item.pn || '').trim().toLowerCase();
      const matchDesc = (i.description || i.name || '').trim().toLowerCase() === (item.description || item.name || '').trim().toLowerCase();
      return matchPn && matchDesc;
    });
    setEditingItemIdx(targetIdx);
    setNewItemUnit(item.unit || 'at802');
    setNewItemCategory(item.category || 'sarf');
    setNewItemDesc(item.description || item.name || '');
    setNewItemPn(item.partNumber || item.pn || '');
    setNewItemLokasyon(item.lokasyonNo || item.location || '');
    const rawSnStr = item.serialAndNotes || item.sn || '';
    setNewItemSnNotes(rawSnStr);
    if (rawSnStr && rawSnStr !== '-') {
      const parts = rawSnStr.split(/[,/;]+/).map(s => s.trim()).filter(Boolean);
      setSnList(parts);
    } else {
      setSnList([]);
    }
    setSingleSnInput('');
    setNewItemGelen(item.gelen || 1);
    const hasLife = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
    setNewItemHasShelfLife(hasLife ? 'EVET' : 'HAYIR');
    setNewItemShelfLifeDate(item.shelfLifeDate && item.shelfLifeDate !== '-' ? item.shelfLifeDate : '');
    setIsNewItemModalOpen(true);
  };

  const handleDeleteItem = async (targetItem: DepoItem | number) => {
    let itemToDelete: DepoItem | null = null;
    if (typeof targetItem === 'number') {
      itemToDelete = inventory[targetItem] || null;
    } else {
      itemToDelete = targetItem;
    }

    if (!itemToDelete) return;

    const itemDesc = itemToDelete.description || itemToDelete.name || 'Malzeme';
    const itemPn = itemToDelete.partNumber || itemToDelete.pn || '-';

    if (!window.confirm(`⚠️ DİKKAT: "${itemDesc}" (P/N: ${itemPn}) malzemesini silmek istediğinize emin misiniz?`)) {
      return;
    }

    const descLower = itemDesc.trim().toLowerCase();
    const pnLower = itemPn.trim().toLowerCase();

    setInventory(prev => prev.filter(i => {
      const p = (i.partNumber || i.pn || '').trim().toLowerCase();
      const d = (i.description || i.name || '').trim().toLowerCase();
      return !(p === pnLower && d === descLower);
    }));

    showNotification("🗑️ Veri silindi.");

    const targetFileName = getDepoStandardFileName(itemToDelete.unit || currentUnit, itemToDelete.category || 'sarf');

    try {
      await fetch('/api/update-depo-excel-item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: targetFileName,
          unit: itemToDelete.unit || currentUnit,
          category: itemToDelete.category || 'sarf',
          oldItem: itemToDelete,
          action: 'delete',
          folderId: DRIVE_FOLDER_ID
        })
      });
      showNotification("🗑️ Veri silindi.");
    } catch (e) {
      showNotification("🗑️ Veri silindi.");
    }
  };

  // Helper to match a transaction against a product item precisely using PN and Description
  const matchTxToItem = (tx: DepoTransaction, item: DepoItem) => {
    const normalizeTight = (s: any) => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    
    const iPn = normalizeTight(item.partNumber || item.pn);
    const iName = normalizeTight(item.description || item.name);

    const tPn = normalizeTight(tx.pn || tx.partNumber || (tx as any).partNo || (tx as any).pn);
    const tName = normalizeTight(tx.itemName || tx.itemDesc || tx.name || (tx as any).malzemeAdi);

    if (!iPn && !iName) return false;
    if (!tPn && !tName) return false;

    // Direct matches
    const pnMatch = iPn && tPn && iPn === tPn;
    const nameMatch = iName && tName && (iName === tName || iName.includes(tName) || tName.includes(iName));
    
    // Cross matches (very common in their data where P/N is written in Name field)
    const pnVsNameMatch = iPn && tName && iPn === tName;
    const nameVsPnMatch = iName && tPn && iName === tPn;

    return pnMatch || nameMatch || pnVsNameMatch || nameVsPnMatch;
  };

  // Helper to render hoverable table cells that display a "Detay Gör" action overlay on hover
  const renderHoverableCell = (val: number | string, item: DepoItem, bgClass: string, columnType = 'ALL', borderClass = "border-r border-slate-200") => {
    return (
      <td className={`py-2 px-1.5 text-center font-bold relative group cursor-pointer ${bgClass} ${borderClass}`}>
        <span className="group-hover:opacity-10 transition duration-150">{val}</span>
        <button
          type="button"
          onClick={() => {
            setDetailModalItem(item);
            setDetailModalColumnType(columnType);
            setIsDetailModalOpen(true);
          }}
          className="absolute inset-0 flex items-center justify-center bg-slate-900/90 text-white text-[9px] font-extrabold uppercase rounded opacity-0 group-hover:opacity-100 transition duration-150 cursor-pointer z-10 whitespace-nowrap px-1 font-sans"
        >
          Detay Gör
        </button>
      </td>
    );
  };

  // Reusable Depo Hareket Geçmişi Section (Covering both Sarf and Kimyasal depots, filtered by aircraft)
  const renderTransactionsCard = (isMainView = false) => {
    // Collect distinct tails for dropdown (Requirement: Kuyruk No sadece ÇIKAN işlem türünde olur ve OR- ile başlar)
    const distinctTails = Array.from(new Set(
      transactions
        .filter(t => {
          const typeStr = (t.type || t.islemTuru || '').toUpperCase();
          return typeStr.includes('ÇIKAN') || typeStr.includes('ÇIKIŞ') || typeStr.includes('SARF');
        })
        .map(t => {
          const raw = (t.tailNo || t.kuyrukKodu || '').trim().toUpperCase();
          const match = raw.match(/OR[-\s]?\d+/i);
          if (match) return match[0].toUpperCase().replace(/\s+/, '-');
          if (raw.startsWith('OR-')) return raw;
          return '';
        })
        .filter(t => t && t.startsWith('OR-'))
    )).sort();

    // CANLI İŞLEM TÜRLERİ: Excel ve veritabanındaki tüm mevcut işlem türlerini dinamik olarak tara ve sırala
    const baseTxTypes = [
      'ANKARA GİREN', 'ANKARA ÇIKAN',
      'KARAİN TRANSFER', 'KARAİN ÇIKAN',
      'ÇANAKKALE TRANSFER', 'ÇANAKKALE ÇIKAN',
      'MİLAS TRANSFER', 'MİLAS ÇIKAN',
      'BURSA TRANSFER', 'BURSA ÇIKAN',
      'MUAYENE GİDEN', 'MUAYENE GELEN'
    ];
    const liveTypes = transactions
      .map(t => (t.islemTuru || t.type || '').trim())
      .filter(Boolean);
    const distinctTxTypes = Array.from(new Set([...baseTxTypes, ...liveTypes])).sort((a, b) => a.localeCompare(b, 'tr-TR'));

    const isFilterActive = Boolean(
      txSearchActiveQuery.trim() || 
      (txFilterType && txFilterType !== 'ALL') || 
      (txFilterTail && txFilterTail !== 'ALL') || 
      txFilterStartDate || 
      txFilterEndDate
    );

    return (
      <div className={`bg-white rounded-2xl shadow-xs border border-slate-200 overflow-hidden ${!isMainView ? 'mt-6' : ''}`}>
        <div className="p-3.5 bg-gradient-to-r from-slate-900 via-slate-850 to-slate-900 text-white flex flex-wrap items-center justify-between gap-3 select-none">
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-emerald-500/20 text-emerald-400 rounded-lg border border-emerald-500/30">
              <RotateCcw className="w-4 h-4" />
            </div>
            <div>
              <div className="font-extrabold text-white text-xs uppercase font-mono tracking-wider flex items-center gap-2">
                <span>DEPO HAREKET GEÇMİŞİ</span>
                <span className="text-[10px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 px-2 py-0.5 rounded-full">
                  {currentUnit === 'all' ? 'TÜM FİLO' : currentUnit.toUpperCase()}
                </span>
                <span className="text-[10px] font-mono text-emerald-400 bg-black/40 px-2 py-0.5 rounded-md border border-emerald-900/50">
                  {filteredTransactions.length} Kayıt
                </span>
              </div>
              <span className="text-[10px] text-slate-300 font-sans">
                Giriş, Çıkış, Bölge Transferleri, Sayım Farkları ve Kimyasal Depo Hareketleri
              </span>
            </div>
          </div>

          {/* Quick aircraft/unit switcher inside history */}
          <div className="flex items-center gap-1.5 flex-wrap">
            {UNITS.map(u => {
              const isSelected = currentUnit === u.id;
              return (
                <button
                  key={u.id}
                  type="button"
                  onClick={() => attemptSwitchUnit(u.id)}
                  className={`px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider transition cursor-pointer ${
                    isSelected 
                      ? 'bg-emerald-500 text-white shadow-sm' 
                      : 'bg-slate-800 text-slate-300 hover:text-white hover:bg-slate-700 border border-slate-700 font-sans'
                  }`}
                >
                  {u.name}
                </button>
              );
            })}
          </div>
        </div>

        {/* 1. GELİŞMİŞ FİLTRELEME ÇUBUĞU (ARA BUTONUNUN ÜSTÜNDE) */}
        <div className="p-3 bg-slate-100/90 border-b border-slate-200 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-2.5 text-xs font-sans">
          
          {/* İŞLEM TÜRÜ */}
          <div>
            <label className="block text-[9.5px] font-black text-slate-600 uppercase font-mono mb-1">
              İŞLEM TÜRÜ
            </label>
            <select
              value={txFilterType}
              onChange={(e) => setTxFilterType(e.target.value)}
              className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-800 text-xs focus:outline-none focus:border-emerald-600 shadow-2xs"
            >
              <option value="ALL">Tüm İşlem Türleri</option>
              <optgroup label="── GENEL FİLTRELER ──">
                <option value="GİREN">Giriş / Giren (Tümü)</option>
                <option value="ÇIKAN">Çıkış / Çıkan (Tümü)</option>
                <option value="TRANSFER">Transferler (Tümü)</option>
                <option value="SAYIM FARKI">Sayım Farkları / Tespitleri (Tümü)</option>
              </optgroup>
              <optgroup label="── CANLI / EXCEL MEVCUT İŞLEM TÜRLERİ ──">
                {distinctTxTypes.map((type, idx) => (
                  <option key={idx} value={type}>{type}</option>
                ))}
              </optgroup>
            </select>
          </div>

          {/* KUYRUK NO / ARAÇ */}
          <div>
            <label className="block text-[9.5px] font-black text-slate-600 uppercase font-mono mb-1">
              KUYRUK NO (Sadece Çıkanlar: OR-..)
            </label>
            <select
              value={txFilterTail}
              onChange={(e) => setTxFilterTail(e.target.value)}
              className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-800 text-xs focus:outline-none focus:border-emerald-600 shadow-2xs"
            >
              <option value="ALL">Tüm Kuyruklar</option>
              {distinctTails.map((tail, idx) => (
                <option key={idx} value={tail}>{tail}</option>
              ))}
            </select>
          </div>

          {/* BAŞLANGIÇ TARİHİ */}
          <div>
            <label className="block text-[9.5px] font-black text-slate-600 uppercase font-mono mb-1">
              BAŞLANGIÇ TARİHİ
            </label>
            <input
              type="date"
              value={txFilterStartDate}
              onChange={(e) => setTxFilterStartDate(e.target.value)}
              className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-800 text-xs focus:outline-none focus:border-emerald-600 shadow-2xs cursor-pointer"
            />
          </div>

          {/* BİTİŞ TARİHİ */}
          <div>
            <label className="block text-[9.5px] font-black text-slate-600 uppercase font-mono mb-1">
              BİTİŞ TARİHİ
            </label>
            <input
              type="date"
              value={txFilterEndDate}
              onChange={(e) => setTxFilterEndDate(e.target.value)}
              className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-800 text-xs focus:outline-none focus:border-emerald-600 shadow-2xs cursor-pointer"
            />
          </div>

        </div>

        {/* 2. MANUEL ARAMA VE ARA BUTONU BÖLÜMÜ */}
        <div className="p-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center gap-2.5 font-sans">
          <div className="relative flex-1 min-w-[260px]">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400 pointer-events-none text-xs">🔍</span>
            <input
              type="text"
              placeholder="PN, Malzeme Adı, Seri No veya Lokasyon ile ara..."
              value={txSearchInput}
              onChange={(e) => {
                setTxSearchInput(e.target.value);
                if (!e.target.value.trim()) {
                  setTxSearchActiveQuery('');
                }
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  setIsTxSearchPerformed(true);
                  setTxSearchActiveQuery(txSearchInput);
                }
              }}
              className="w-full pl-8 pr-3 py-1.5 bg-white text-slate-800 text-xs font-bold rounded-xl border border-slate-300 focus:outline-none focus:ring-2 focus:ring-emerald-500 placeholder:text-slate-400 font-sans"
            />
          </div>
          <button
            type="button"
            onClick={() => {
              setIsTxSearchPerformed(true);
              setTxSearchActiveQuery(txSearchInput);
            }}
            className="px-5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black uppercase tracking-wider rounded-xl transition cursor-pointer active:scale-95 shadow-sm font-sans"
          >
            ARA
          </button>
          {isFilterActive && (
            <button
              type="button"
              onClick={() => {
                setTxSearchInput('');
                setTxSearchActiveQuery('');
                setTxFilterType('ALL');
                setTxFilterTail('ALL');
                setTxFilterStartDate('');
                setTxFilterEndDate('');
                setIsTxSearchPerformed(false);
              }}
              className="px-3 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-bold rounded-xl transition cursor-pointer font-sans"
            >
              Filtreleri Sıfırla
            </button>
          )}
        </div>

        {!isTxSearchPerformed ? (
          <div className="py-14 text-center text-slate-500 flex flex-col items-center justify-center bg-white border-b border-slate-200 select-none">
            <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center mb-2.5 text-xl">🔍</div>
            <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider font-sans">
              Depo Hareket Geçmişi Arama & Filtreleme
            </h4>
            <p className="text-[11px] text-slate-500 mt-1 max-w-md px-4 font-sans leading-relaxed">
              Depo hareket geçmişi veritabanı oldukça büyüktür. Arama yapıp sistemi hızlandırmak için yukarıdaki filtreleri ayarlayın veya arama kutusuna yazıp <strong>ARA</strong> butonuna basın!
            </p>
            <button
              type="button"
              onClick={() => {
                setIsTxSearchPerformed(true);
                setTxSearchActiveQuery(' ');
              }}
              className="mt-4 px-5 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-black uppercase rounded-xl transition cursor-pointer active:scale-95 shadow-xs font-sans"
            >
              Tüm Depo Hareket Geçmişini Listele
            </button>
          </div>
        ) : filteredTransactions.length === 0 ? (
          <div className="py-14 text-center text-slate-500 flex flex-col items-center justify-center bg-white border-b border-slate-200 select-none">
            <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center mb-2 text-lg">⚠️</div>
            <h4 className="text-xs font-bold text-slate-800 uppercase font-sans">Eşleşen Kayıt Bulunamadı</h4>
            <p className="text-[11px] text-slate-500 mt-1">Girdiğiniz kriterlere uygun depo hareketi bulunamadı.</p>
          </div>
        ) : (
          <>
            <div className="overflow-x-auto max-h-[72vh] overflow-y-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-[#b0b0b0] sticky top-0 z-20 border-b-2 border-slate-400 text-slate-800 font-black text-[11px] uppercase tracking-wider text-center select-none font-sans">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[50px]">SIRA</th>
                    <th className="py-2.5 px-4 border-r border-slate-300 text-left w-[260px] max-w-[280px]">MALZEME ADI</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[90px]">KATEGORİ</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[65px]">ADET</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[95px]">TARİH</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[130px]">İŞLEM TÜRÜ</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 font-mono min-w-[120px]">SERİAL NUMBER</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[130px]">KUYRUK KODU</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px]">TESLİM ALAN</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px]">KABUL YAPAN</th>
                    <th className="py-2.5 px-3 border-r border-slate-300 min-w-[130px]">DEPO YERİ</th>
                    <th className="py-2.5 px-4 border-r border-slate-300 min-w-[160px]">SAYFA / BİRİM</th>
                    <th className="py-2.5 px-3 min-w-[130px] bg-slate-700 text-white font-black">İŞLEMLER</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                  {paginatedTransactions.map((t, i) => {
                    const typeStr = (t.type || t.islemTuru || '').toUpperCase();
                    const isTransfer = typeStr.includes('TRANSFER');
                    const isGiris = typeStr.includes('GİREN') || typeStr.includes('GİRİŞ') || typeStr.includes('GELEN') || typeStr.includes('YENİ ÜRÜN');
                    const locNorm = String(t.location || t.depoYeri || '').toUpperCase().replace(/İ/g, 'I');
                    const sheetNorm = String(t.sheetName || '').toUpperCase().replace(/İ/g, 'I');
                    const nameNorm = String(t.itemName || t.itemDesc || t.name || '').toUpperCase().replace(/İ/g, 'I');
                    const pnNorm = String(t.pn || t.partNumber || '').toUpperCase().replace(/İ/g, 'I');
                    const isKimyasal = (t.category === 'kimyasal') || 
                                       locNorm.includes('KIMYA') || 
                                       sheetNorm.includes('KIMYA') ||
                                       nameNorm.includes('AEROSHELL') || 
                                       nameNorm.includes('GREASE') || 
                                       nameNorm.includes('ENGINE OIL') || 
                                       nameNorm.includes('OIL') || 
                                       nameNorm.includes('FLUID') || 
                                       nameNorm.includes('HYDRAULIC') || 
                                       nameNorm.includes('BOYA') || 
                                       nameNorm.includes('YAG') ||
                                       pnNorm.includes('MIL-PRF') ||
                                       pnNorm.includes('MIL-H-') ||
                                       pnNorm.includes('MIL-G-');
                    const globalIdx = (txPage - 1) * txPageSize + i + 1;

                    return (
                      <tr 
                        key={i} 
                        onDoubleClick={() => handleTxDoubleClick(t)}
                        className="hover:bg-slate-50 border-b border-slate-200 text-xs text-center cursor-pointer transition select-none"
                        title="DÜZENLEMEK İÇİN ÇİFT TIKLAYIN (Şifre korumalı)"
                      >
                        <td className="py-2 px-2 text-center font-mono font-bold text-slate-500 border-r border-slate-200">{globalIdx}</td>
                        <td className="py-2 px-3 font-bold text-slate-900 border-r border-slate-200 text-left font-sans max-w-[280px]">
                          <div 
                            className="truncate overflow-hidden text-ellipsis whitespace-nowrap cursor-pointer hover:text-blue-600 transition" 
                            title={`${t.itemName || t.itemDesc || t.name || '-'} (Detay için çift tıklayın)`}
                          >
                            {t.itemName || t.itemDesc || t.name || '-'}
                          </div>
                        </td>
                        <td className="py-2 px-3 text-center border-r border-slate-200 font-sans">
                          <span className={`inline-block px-2 py-0.5 rounded-full text-[10px] font-black uppercase ${
                            isKimyasal ? 'bg-purple-100 text-purple-900 border border-purple-300' : 'bg-blue-50 text-blue-900 border border-blue-200'
                          }`}>
                            {isKimyasal ? '🧪 KİMYASAL' : '📦 SARF/PARÇA'}
                          </span>
                        </td>
                        <td className={`py-2 px-3 text-center font-mono font-black border-r border-slate-200 ${
                          isGiris ? 'text-emerald-700' : (isTransfer ? 'text-indigo-700' : 'text-rose-700')
                        }`}>
                          {t.quantity !== undefined ? t.quantity : (t.adet || 1)}
                        </td>
                        <td className="py-2 px-3 font-mono text-slate-600 border-r border-slate-200 text-center">{t.date || t.timestamp || '-'}</td>
                        <td className="py-2 px-3 border-r border-slate-200 text-center font-sans">
                          <span className={`inline-block px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider ${
                            isTransfer 
                              ? 'bg-indigo-100 text-indigo-900 border border-indigo-300' 
                              : isGiris 
                              ? 'bg-emerald-100 text-emerald-900 border border-emerald-300' 
                              : 'bg-rose-100 text-rose-900 border border-rose-300'
                          }`}>
                            {t.type || t.islemTuru || 'TRANSFER'}
                          </span>
                        </td>
                        <td className="py-2 px-3 font-mono font-bold text-slate-700 border-r border-slate-200 text-center">{t.sn || t.serialNumber || '-'}</td>
                        <td className="py-2 px-3 font-semibold text-slate-700 border-r border-slate-200 text-center font-sans">{t.tailNo || t.kuyrukKodu || '-'}</td>
                        <td className="py-2 px-3 text-slate-700 border-r border-slate-200 font-sans text-left">{t.operator || t.teslimAlan || '-'}</td>
                        <td className="py-2 px-3 text-slate-700 border-r border-slate-200 font-sans text-left">{t.receivedBy || t.kabulYapan || '-'}</td>
                        <td className="py-2 px-3 font-semibold text-slate-800 border-r border-slate-200 text-left font-sans">
                          {isKimyasal ? 'KİMYASAL DEPO' : (t.location || t.depoYeri || '-')}
                        </td>
                        <td className="py-2 px-4 text-center font-sans border-r border-slate-200">
                          <span className="inline-block px-2 py-0.5 rounded-md bg-slate-100 text-slate-800 border border-slate-300 font-mono text-[10px] font-bold">
                            {t.sheetName || (currentUnit === 'all' ? 'DEPO HAREKET GEÇMİŞİ' : `DEPO HAREKET GEÇMİŞİ-${currentUnit.toUpperCase()}`)}
                          </span>
                        </td>
                        <td className="py-2 px-2 text-center bg-slate-50">
                          <div className="flex items-center justify-center gap-1.5" onClick={(e) => e.stopPropagation()}>
                            {isTransfer && (
                              <button
                                type="button"
                                onClick={() => handleOpenRollbackModal(t)}
                                className="px-2 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-[10px] font-black uppercase flex items-center gap-1 cursor-pointer transition shadow-2xs font-sans"
                                title="Bu transferi geri çekip ana stoğa iade edin"
                              >
                                <RotateCcw className="w-3 h-3" />
                                <span>Geri Çek</span>
                              </button>
                            )}
                            <button
                              type="button"
                              onClick={() => handleTxDoubleClick(t)}
                              className="px-2 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-lg text-[10px] font-black uppercase flex items-center gap-1 cursor-pointer transition shadow-2xs font-sans"
                              title="Bu hareket kaydını düzenle (Şifre korumalı)"
                            >
                              <Edit2 className="w-3 h-3" />
                              <span>Düzenle</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => handleOpenDeleteTx(t)}
                              className="px-2 py-1 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-[10px] font-black uppercase flex items-center gap-1 cursor-pointer transition shadow-2xs font-sans"
                              title="Bu hareket kaydını sistemden ve Excel'den sil (Şifre korumalı)"
                            >
                              <Trash2 className="w-3 h-3" />
                              <span>Sil</span>
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <PaginationBar
              currentPage={txPage}
              totalItems={filteredTransactions.length}
              pageSize={txPageSize}
              onPageChange={setTxPage}
              onPageSizeChange={setTxPageSize}
            />

            {filteredTransactions.length === 0 && (
              <div className="py-8 text-center bg-slate-50 font-sans">
                <div className="w-9 h-9 rounded-full bg-slate-200 flex items-center justify-center mx-auto mb-1.5 text-base">🔄</div>
                <h4 className="text-xs font-bold text-slate-700 uppercase">Eşleşen Hareket Kaydı Bulunmuyor</h4>
                <p className="text-[11px] text-slate-500 mt-0.5">Aradığınız kritere uygun herhangi bir malzeme girişi, çıkışı veya transfer hareketi bulunamadı.</p>
              </div>
            )}
          </>
        )}
      </div>
    );
  };

  return (
    <div className="fixed inset-0 z-[9999] flex flex-col bg-slate-100 text-slate-800 overflow-y-auto animate-fade-in">
      
      {/* 1. TOP HEADER - GÖRSEL VE TASARIMIN KALBİ */}
      <header className="bg-[#0b3d1d] text-white shadow-md sticky top-0 z-40 border-b border-emerald-500/30 shrink-0">
        <div className="max-w-7xl mx-auto px-4 py-2.5 flex flex-wrap items-center justify-between gap-3">
          
          {/* Logo & Başlık */}
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-white/10 flex items-center justify-center border border-white/20 text-lg shadow-inner">
              📦
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-sm sm:text-base font-black tracking-wider uppercase text-white">
                  OGM HAVACILIK DEPO PORTALI
                </h1>
                <span className="text-[10px] bg-emerald-500/20 text-emerald-300 font-mono font-bold px-2 py-0.5 rounded-full border border-emerald-400/30">
                  CANLI STOK
                </span>
              </div>
              <p className="text-[11px] text-emerald-200/80">
                Merkez Ankara & Bölge Transfer Takip Sistemi
              </p>
            </div>
          </div>

          {/* Quick Actions */}
          <div className="flex items-center gap-2 flex-wrap">

            <button 
              type="button"
              onClick={handleExportToExcel}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-800 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold transition shadow-xs border border-emerald-600/40 cursor-pointer active:scale-95"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Excel İndir</span>
            </button>

            {/* EXCEL YÜKLE - DEPO TÜRÜ SEÇTİREN MODAL */}
            <button
              type="button"
              onClick={() => setIsExcelTypeModalOpen(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white/10 hover:bg-white/20 text-white rounded-lg text-xs font-bold transition shadow-xs border border-white/20 cursor-pointer active:scale-95"
            >
              <UploadCloud className="w-3.5 h-3.5 text-emerald-300" />
              <span>Excel Yükle</span>
            </button>
            <input 
              ref={fileInputRef}
              type="file" 
              accept=".xlsx, .xls, .csv" 
              className="hidden" 
              onChange={handleExcelUpload} 
            />

            <button 
              type="button"
              onClick={() => {
                setEditingItemIdx(-1);
                setNewItemUnit(currentUnit !== 'all' ? currentUnit : 'at802');
                setNewItemCategory(currentDepoType === 'kimyasal' ? 'kimyasal' : 'sarf');
                setNewItemDesc('');
                setNewItemPn('');
                setNewItemLokasyon('');
                setNewItemSnNotes('');
                setNewItemGelen(1);
                setIsNewItemModalOpen(true);
              }}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-500 hover:bg-amber-400 text-amber-950 font-black rounded-lg text-xs uppercase tracking-wider transition shadow-xs cursor-pointer active:scale-95"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Yeni Malzeme</span>
            </button>

            <button 
              type="button"
              onClick={onClose}
              className="flex items-center gap-1 px-3 py-1.5 bg-rose-700 hover:bg-rose-600 text-white font-bold rounded-lg text-xs transition cursor-pointer active:scale-95 ml-2"
              title="Portaldan Çıkış Yap"
            >
              <X className="w-3.5 h-3.5" />
              <span>Kapat</span>
            </button>
          </div>
        </div>

        {/* 2. FLEET UNIT TABS WITH PASSWORD PROTOCOL */}
        <div className="bg-[#082a14] px-4 py-1.5 border-t border-white/10 overflow-x-auto">
          <div className="max-w-7xl mx-auto flex items-center justify-between gap-3">
            <div className="flex items-center gap-1.5 overflow-x-auto py-0.5">
              {UNITS.map(u => {
                const isActive = currentUnit === u.id;
                const isUnlocked = unlockedUnits.has(u.id);
                const count = computedInventory.filter(i => u.id === 'all' || i.unit === u.id).length;

                return (
                  <button
                    key={u.id}
                    type="button"
                    onClick={() => attemptSwitchUnit(u.id)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-black uppercase tracking-wider transition whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                      isActive ? 'bg-emerald-500 text-white shadow-md' : 'text-emerald-100/70 hover:text-white hover:bg-white/10'
                    }`}
                  >
                    <span>{isUnlocked ? '🔓' : '🔒'}</span>
                    <span>{u.name}</span>
                    <span className={`text-[10px] font-mono px-1.5 py-0.2 rounded-full ${isActive ? 'bg-emerald-700 text-white' : 'bg-white/10 text-emerald-300'}`}>
                      {count}
                    </span>
                  </button>
                );
              })}
            </div>
            

          </div>
        </div>
      </header>

      {/* 3. SUB NAVIGATION & ACTION BUTTONS */}
      <section className="bg-white border-b border-slate-200 shadow-xs shrink-0">
        <div className="max-w-7xl mx-auto px-4 py-2.5 flex flex-wrap items-center justify-between gap-4">
          
          {/* Sekmeler */}
          <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200">
            <button
              type="button"
              onClick={() => setCurrentDepoType('sarf')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition cursor-pointer ${
                currentDepoType === 'sarf'
                  ? 'bg-[#0b3d1d] text-white shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              📦 Sarf & Parça Depo ({sarfCount})
            </button>

            <button
              type="button"
              onClick={() => setCurrentDepoType('kimyasal')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer ${
                currentDepoType === 'kimyasal'
                  ? 'bg-[#0b3d1d] text-white shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              🧪 Kimyasal Depo ({kimyasalCount})
            </button>

            {/* SADECE AT-802 İÇİN: YAŞAM DESTEK (Kimyasal Depo ile Depo Hareket Geçmişi Arasında) */}
            {(currentUnit === 'all' || currentUnit.toLowerCase().startsWith('at802') || currentUnit.toLowerCase().startsWith('at-802')) && (
              <button
                type="button"
                onClick={() => setCurrentDepoType('yasam_destek')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                  currentDepoType === 'yasam_destek'
                    ? 'bg-[#0b3d1d] text-white shadow-xs ring-1 ring-emerald-400'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <span>🦺 Yaşam Destek ({yasamDestekCount})</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => setCurrentDepoType('transactions')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer ${
                currentDepoType === 'transactions'
                  ? 'bg-[#0b3d1d] text-white shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              🔄 Depo Hareket Geçmişi ({txCount})
            </button>
          </div>

          {/* Aksiyon Butonları (Görsel 1 ile birebir aynı) */}
          <div className="flex items-center gap-2.5">
            <button 
              type="button"
              onClick={() => openMovementModalHandler('cikis')}
              className="flex items-center gap-2 px-3.5 py-1.5 bg-rose-50/90 hover:bg-rose-100 text-[#a0153e] border border-rose-300 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95"
            >
              <span className="w-4 h-4 rounded bg-blue-500 text-white flex items-center justify-center text-[10px] font-bold">↗</span>
              <span>Malzeme Çıkışı (Sarf)</span>
            </button>

            <button 
              type="button"
              onClick={() => openMovementModalHandler('giris')}
              className="flex items-center gap-2 px-3.5 py-1.5 bg-emerald-50/90 hover:bg-emerald-100 text-[#006a4e] border border-emerald-300 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95"
            >
              <span className="w-4 h-4 rounded bg-blue-500 text-white flex items-center justify-center text-[10px] font-bold">↙</span>
              <span>Malzeme Girişi</span>
            </button>

            <button 
              type="button"
              onClick={() => openMovementModalHandler('transfer')}
              className="flex items-center gap-2 px-3.5 py-1.5 bg-indigo-50/90 hover:bg-indigo-100 text-indigo-900 border border-indigo-300 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95"
            >
              <ArrowLeftRight className="w-3.5 h-3.5 text-indigo-700" />
              <span>Depolar Arası Transfer</span>
            </button>

            <button 
              type="button"
              onClick={openKitModalHandler}
              className="flex items-center gap-2 px-3.5 py-1.5 bg-amber-50/90 hover:bg-amber-100 text-amber-900 border border-amber-300 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95"
              title="Yeni Kit Oluştur ve Malzeme Stoklarından Otomatik Düş"
            >
              <Package className="w-3.5 h-3.5 text-amber-700" />
              <span>Kit Oluştur</span>
            </button>

            {/* KİT OLUŞTUR İLE BARKODA BAS ARASINDAKİ MUAYENE/BAKIM GÖNDER BUTONU */}
            <button 
              type="button"
              onClick={openMuayeneModalHandler}
              className="flex items-center gap-2 px-3.5 py-1.5 bg-blue-50/90 hover:bg-blue-100 text-blue-950 border border-blue-300 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95"
              title="Malzemeyi Muayene veya Dış Bakıma Gönder (Mevcuttan Düş ve Kaydet)"
            >
              <Wrench className="w-3.5 h-3.5 text-blue-700" />
              <span>Muayene/Bakım Gönder</span>
            </button>

            <button 
              type="button"
              onClick={() => {
                setSlipModalType(null);
                setSlipModalData(null);
                setIsSlipModalOpen(true);
              }}
              className="flex items-center gap-2 px-3.5 py-1.5 bg-slate-900 hover:bg-slate-800 text-white border border-slate-700 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95 font-sans"
              title="Raf Lokasyon Slibi veya Depo Girişi Slibi Yazdır / İndir"
            >
              <Barcode className="w-3.5 h-3.5 text-emerald-400" />
              <span>Barkod Bas</span>
            </button>

            <button 
              type="button"
              onClick={() => setIsBelgeBulModalOpen(true)}
              className="flex items-center gap-2 px-3.5 py-1.5 bg-indigo-900 hover:bg-indigo-800 text-white border border-indigo-700 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95"
              title="Malzeme Belgeleri, Sertifikalar, Form-1, CoC ve EBYS Evraklarını Bul"
            >
              <FileSearch className="w-3.5 h-3.5 text-indigo-400" />
              <span>Belge Bul</span>
            </button>

            {/* BELGE BUL'UN YANINDA ONAY BEKLEYENLER BUTONU */}
            <button 
              type="button"
              onClick={() => {
                setOnayPasswordInput('');
                setOnayPasswordError('');
                setIsOnayPasswordModalOpen(true);
              }}
              className="relative flex items-center gap-2 px-3.5 py-1.5 bg-amber-600 hover:bg-amber-500 text-white border border-amber-500 rounded-xl text-xs font-extrabold transition shadow-xs cursor-pointer active:scale-95 font-sans"
              title="Barkod Okuyucu / Portal üzerinden gelen AT-802 çıkış taleplerini görüntüle ve onayla (Şifre Korumalı)"
            >
              <Clock className="w-3.5 h-3.5 text-amber-200" />
              <span>Onay Bekleyenler</span>
              {pendingApprovalCount > 0 && (
                <span className="w-5 h-5 rounded-full bg-rose-600 text-white text-[10px] font-black flex items-center justify-center animate-bounce shadow-xs">
                  {pendingApprovalCount}
                </span>
              )}
            </button>
          </div>
        </div>
      </section>

      {/* 4. MAIN CONTAINER & VIEWS */}
      <main className="max-w-7xl mx-auto px-4 py-4 flex-1 w-full space-y-4">
        
        {/* ARAMA VE BÖLGE FİLTRE ÇUBUĞU */}
        <div className="bg-white p-3 rounded-2xl shadow-xs border border-slate-200 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3 flex-1 min-w-[280px]">
            <div className="relative flex-1 flex items-center gap-2">
              <div className="relative flex-1">
                <input 
                  type="text" 
                  value={searchQuery}
                  onChange={(e) => {
                    setSearchQuery(e.target.value);
                    if (!e.target.value.trim()) {
                      setIsDepoSearchPerformed(false);
                    }
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      setIsDepoSearchPerformed(true);
                    }
                  }}
                  placeholder="Malzeme Adı (Description), Part Number (P/N), Seri No veya Lokasyon ara..." 
                  className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:border-emerald-600 transition"
                />
                <Search className="w-4 h-4 absolute left-3 top-2.5 text-slate-400" />
              </div>
              <button 
                type="button"
                onClick={() => setIsDepoSearchPerformed(true)}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-black uppercase rounded-xl transition shadow-sm cursor-pointer active:scale-95 shrink-0 font-sans"
              >
                ARA
              </button>
            </div>

            {currentDepoType !== 'transactions' && (
              <div className="flex items-center gap-2 flex-wrap">
                <select 
                  value={filterRegion}
                  onChange={(e) => {
                    setFilterRegion(e.target.value);
                    setIsDepoSearchPerformed(true);
                  }}
                  className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 focus:outline-none cursor-pointer font-sans"
                >
                  <option value="ALL">Tüm Bölgeler / Lokasyonlar</option>
                  <option value="ANKARA">Ankara (Merkez)</option>
                  <option value="KARAİN">Karain</option>
                  <option value="ÇANAKKALE">Çanakkale</option>
                  <option value="MİLAS">Milas</option>
                  <option value="BURSA">Bursa</option>
                </select>

                {/* LOKASYON NO ÇOKLU SEÇİM (EXCEL STİLİ CHECKBOX FİLTRESİ) */}
                <div className="relative" ref={locationDropdownRef}>
                  <button
                    type="button"
                    onClick={() => {
                      if (!isLocationDropdownOpen) {
                        setPendingLocations(selectedLocations.length === 0 ? [...uniqueLocations] : [...selectedLocations]);
                      }
                      setIsLocationDropdownOpen(!isLocationDropdownOpen);
                    }}
                    className="bg-slate-50 border border-slate-200 hover:border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 focus:outline-none flex items-center justify-between gap-2 min-w-[200px] cursor-pointer font-sans"
                    title="Excel gibi birden fazla lokasyonu seçerek filtreleyin"
                  >
                    <div className="flex items-center gap-1.5 truncate">
                      <span>📍</span>
                      <span className="truncate">
                        {selectedLocations.length === 0 ? 'Lokasyon Seç: (Tümü)' : `Lokasyon: (${selectedLocations.length} Seçili)`}
                      </span>
                    </div>
                    <ChevronDown className={`w-3.5 h-3.5 text-slate-400 transition-transform ${isLocationDropdownOpen ? 'rotate-180' : ''}`} />
                  </button>

                  {isLocationDropdownOpen && (
                    <div className="absolute top-full left-0 mt-1 w-80 bg-white border-2 border-slate-300 rounded-2xl shadow-2xl z-50 p-2.5 text-xs font-sans animate-fade-in">
                      {/* Metin Filtreleri Arama Kutusu */}
                      <div className="relative mb-2">
                        <input
                          type="text"
                          value={locationSearchTerm}
                          onChange={(e) => setLocationSearchTerm(e.target.value)}
                          placeholder="Metin Filtreleri (Lokasyon Ara)..."
                          className="w-full pl-7 pr-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:border-emerald-600"
                          autoFocus
                        />
                        <Search className="w-3.5 h-3.5 absolute left-2 top-2 text-slate-400" />
                      </div>

                      {/* (Tümünü Seç) / (Tümünü Kaldır) Buton Çubuğu */}
                      <div className="border-b border-slate-200 pb-2 mb-2 flex items-center justify-between gap-2">
                        <button
                          type="button"
                          onClick={() => setPendingLocations([...uniqueLocations])}
                          className="px-2.5 py-1 bg-emerald-100 hover:bg-emerald-200 text-emerald-900 rounded-lg text-[11px] font-bold transition cursor-pointer"
                        >
                          ☑ Tümünü Seç
                        </button>
                        <button
                          type="button"
                          onClick={() => setPendingLocations([])}
                          className="px-2.5 py-1 bg-rose-100 hover:bg-rose-200 text-rose-900 rounded-lg text-[11px] font-bold transition cursor-pointer"
                        >
                          ☒ Tümünü Kaldır
                        </button>
                      </div>

                      {/* Checkbox Lokasyon Listesi (Geçici seçme yapılsın, Tamam butonuna basılana kadar filtre işlemesin) */}
                      <div className="max-h-60 overflow-y-auto space-y-0.5 divide-y divide-slate-50 pr-1">
                        {filteredLocationsForDropdown.length === 0 ? (
                          <div className="p-3 text-center text-slate-400 text-[11px]">
                            Eşleşen lokasyon bulunamadı.
                          </div>
                        ) : (
                          filteredLocationsForDropdown.map((loc, idx) => {
                            const isChecked = pendingLocations.includes(loc);
                            return (
                              <label key={idx} className="flex items-center gap-2 px-2 py-1 hover:bg-emerald-50 rounded-lg cursor-pointer text-slate-700 text-[11px] font-mono select-none">
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => {
                                    if (isChecked) {
                                      setPendingLocations(pendingLocations.filter(l => l !== loc));
                                    } else {
                                      setPendingLocations([...pendingLocations, loc]);
                                    }
                                  }}
                                  className="w-4 h-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                                />
                                <span className="truncate" title={loc}>{loc}</span>
                              </label>
                            );
                          })
                        )}
                      </div>

                      {/* Alt Aksiyon Butonları */}
                      <div className="pt-2 mt-2 border-t border-slate-200 flex items-center justify-between text-[11px]">
                        <button
                          type="button"
                          onClick={() => setPendingLocations([])}
                          className="text-slate-500 hover:text-slate-900 font-bold px-2 py-1 rounded hover:bg-slate-100"
                        >
                          Temizle
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            if (pendingLocations.length === 0 || pendingLocations.length === uniqueLocations.length) {
                              setSelectedLocations([]);
                            } else {
                              setSelectedLocations(pendingLocations);
                            }
                            setIsDepoSearchPerformed(true);
                            setIsLocationDropdownOpen(false);
                          }}
                          className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-lg cursor-pointer shadow-xs"
                        >
                          Tamam (Filtrele)
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          <div className="flex items-center gap-3 text-xs text-slate-600 font-medium font-sans">
            <span>
              Kayıt: <strong className="text-slate-900 font-bold">
                {currentDepoType === 'sarf' ? (isDepoSearchPerformed ? filteredSarfRows.length : 0) : (currentDepoType === 'kimyasal' ? (isDepoSearchPerformed ? filteredKimyasalRows.length : 0) : transactions.length)}
              </strong> / {computedInventory.filter(i => currentUnit === 'all' || i.unit === currentUnit).length}
            </span>
<button 
type="button"
onClick={handleDownloadFilteredExcel}
className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white rounded-xl text-xs font-black transition shadow-xs cursor-pointer active:scale-95 font-sans"
title="Ekrandaki filtrelenmiş malzemeleri ve seçili bölgeye ait sütunları Excel olarak bilgisayarınıza indirir"
>
<Download className="w-3.5 h-3.5 text-emerald-200" />
<span>Filtrelenmiş Excel İndir</span>
</button>
<button 
type="button"
onClick={() => loadLiveDriveExcel(true)}
disabled={isLoadingDrive}
className="flex items-center gap-1.5 px-3 py-1.5 bg-[#0b3d1d] hover:bg-emerald-900 disabled:opacity-50 text-white rounded-xl text-xs font-bold transition shadow-xs cursor-pointer active:scale-95"
title="Google Drive üzerindeki Excel dosyasından anında canlı verileri tazeler"
>
<RefreshCw className={`w-3.5 h-3.5 ${isLoadingDrive ? 'animate-spin' : ''}`} />
<span>{isLoadingDrive ? 'Drive Bağlanıyor...' : 'Drive Canlı Senkronize Et'}</span>
</button>
            <button 
              type="button"
              onClick={() => { setSearchQuery(''); setFilterRegion('ALL'); setSelectedLocations([]); setIsDepoSearchPerformed(false); }}
              className="px-2.5 py-1 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg cursor-pointer font-sans"
            >
              Sıfırla
            </button>
          </div>
        </div>

        {/* VIEW 1: SARF & PARÇA DEPOSU (28 SÜTUNLU LİVE MATRIX) */}
        {currentDepoType === 'sarf' && (
          <div className="space-y-4 font-sans">
            {!isDepoSearchPerformed && !searchQuery.trim() ? (
              <div className="py-14 text-center text-slate-500 flex flex-col items-center justify-center bg-white border border-slate-200 rounded-2xl shadow-xs select-none">
                <div className="w-12 h-12 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center mb-2.5 text-xl font-black">🔍</div>
                <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider font-sans">
                  {currentUnit.toUpperCase()} SARF & PARÇA DEPOSU
                </h4>
                <p className="text-[11px] text-slate-500 mt-1 max-w-md px-4 font-sans leading-relaxed">
                  Depo stok veritabanında <strong>{computedInventory.filter(i => (currentUnit === 'all' || i.unit === currentUnit) && (i.category === 'sarf' || !i.category)).length} kalem</strong> malzeme bulunmaktadır. Sistemi dondurmamak için yukarıdaki kutuya arama yazıp <strong>ARA</strong> butonuna basın veya aşağıdaki butonla tümünü listeleyin.
                </p>
                <button
                  type="button"
                  onClick={() => setIsDepoSearchPerformed(true)}
                  className="mt-4 px-5 py-2.5 bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-black uppercase tracking-wider rounded-xl transition cursor-pointer active:scale-95 shadow-sm font-sans flex items-center gap-2"
                >
                  <span>Tüm Malzemeleri Listele ({computedInventory.filter(i => (currentUnit === 'all' || i.unit === currentUnit) && (i.category === 'sarf' || !i.category)).length} Kalem)</span>
                </button>
              </div>
            ) : (
            <div className="bg-white rounded-2xl shadow-xs border border-slate-200 overflow-hidden">
              
              {/* Tablo Bilgi & Renk Açıklaması */}
              <div className="px-4 py-2 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between text-xs">
                <div className="flex items-center gap-2">
                  <span className="font-extrabold text-[#0b3d1d] uppercase font-mono tracking-wider">
                    {currentUnit.toUpperCase()} SARF & PARÇA DEPO
                  </span>
                  <span className="text-[11px] text-emerald-800 font-bold bg-emerald-100 px-2 py-0.5 rounded-md">
                    {isLoadingDrive ? '⏳ Drive Excel Verileri Yükleniyor...' : `🟢 Canlı Drive Excel (${inventory.length} Malzeme)`}
                  </span>
                  <span className="text-[11px] text-slate-500 font-medium">
                    · Ankara Merkez Depo Dağıtımlı Canlı Stok
                  </span>
                </div>
                <div className="flex items-center gap-2 text-[11px]">
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#ffff00] border border-slate-300 inline-block rounded-xs"></span> Lokasyon/Karain</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#00ffff] border border-slate-300 inline-block rounded-xs"></span> Toplam</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#ff9900] border border-slate-300 inline-block rounded-xs"></span> Ankara/Milas</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#92d050] border border-slate-300 inline-block rounded-xs"></span> Çanakkale</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#ff00ff] border border-slate-300 inline-block rounded-xs"></span> Bursa</span>
                </div>
              </div>

              <div className="overflow-x-auto max-h-[72vh]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 z-20 text-[11px] font-black uppercase text-center tracking-wider border-b-2 border-slate-300 select-none">
                    <tr>
                      <th className="py-2.5 px-2 bg-slate-100 text-slate-800 border-r border-slate-300 min-w-[65px] text-center sticky left-0 z-30 font-black">SIRA NO</th>
                      <th className="py-2.5 px-3 bg-white text-slate-800 border-r border-slate-300 min-w-[240px] text-left sticky left-[65px] z-30">DESCRIPTION</th>
                      <th className="py-2.5 px-3 bg-white text-slate-800 border-r border-slate-300 min-w-[140px] font-mono sticky left-[305px] z-30 text-left">PART NUMBER</th>
                      <th className="py-2.5 px-3 bg-white text-slate-800 border-r border-slate-300 min-w-[200px] text-left">SERİ NUMBER - MÜKERRER NO - AÇIKLAMA</th>
                      <th className="py-2.5 px-2 bg-[#ffff00] text-black border-r border-slate-300 min-w-[100px]">LOKASYON NO</th>
                      <th className="py-2.5 px-2 bg-[#d9d9d9] text-black border-r border-slate-300 min-w-[70px]">GELEN</th>
                      <th className="py-2.5 px-2 bg-[#00ffff] text-black border-r border-slate-300 min-w-[85px] font-mono">TOPLAM STOK</th>
                      
                      {/* ANKARA (Turuncu) */}
                      {(filterRegion === 'ALL' || filterRegion === 'ANKARA') && (
                        <>
                          <th className="py-2.5 px-2 bg-[#ff9900] text-black border-r border-slate-300 min-w-[75px]">ANKARA ÇIKAN</th>
                          <th className="py-2.5 px-2 bg-[#ff9900] text-black border-r border-slate-300 min-w-[75px]">ANKARA MEVCUT</th>
                        </>
                      )}

                      {/* KARAİN (Sarı) */}
                      {(filterRegion === 'ALL' || filterRegion === 'KARAİN') && (
                        <>
                          <th className="py-2.5 px-2 bg-[#ffff00] text-black border-r border-slate-300 min-w-[75px]">KARAİN TRANSFER</th>
                          <th className="py-2.5 px-2 bg-[#ffff00] text-black border-r border-slate-300 min-w-[70px]">KARAİN ÇIKAN</th>
                          <th className="py-2.5 px-2 bg-[#ffff00] text-black border-r border-slate-300 min-w-[70px]">KARAİN MEVCUT</th>
                        </>
                      )}

                      {/* ÇANAKKALE (Yeşil) */}
                      {(filterRegion === 'ALL' || filterRegion === 'ÇANAKKALE') && (
                        <>
                          <th className="py-2.5 px-2 bg-[#92d050] text-black border-r border-slate-300 min-w-[75px]">ÇANAKKALE TRANSFER</th>
                          <th className="py-2.5 px-2 bg-[#92d050] text-black border-r border-slate-300 min-w-[70px]">ÇANAKKALE ÇIKAN</th>
                          <th className="py-2.5 px-2 bg-[#92d050] text-black border-r border-slate-300 min-w-[70px]">ÇANAKKALE MEVCUT</th>
                        </>
                      )}

                      {/* MİLAS (Turuncu) */}
                      {(filterRegion === 'ALL' || filterRegion === 'MİLAS') && (
                        <>
                          <th className="py-2.5 px-2 bg-[#ff9900] text-black border-r border-slate-300 min-w-[75px]">MİLAS TRANSFER</th>
                          <th className="py-2.5 px-2 bg-[#ff9900] text-black border-r border-slate-300 min-w-[70px]">MİLAS ÇIKAN</th>
                          <th className="py-2.5 px-2 bg-[#ff9900] text-black border-r border-slate-300 min-w-[70px]">MİLAS MEVCUT</th>
                        </>
                      )}

                      {/* BURSA (Pembe) */}
                      {(filterRegion === 'ALL' || filterRegion === 'BURSA') && (
                        <>
                          <th className="py-2.5 px-2 bg-[#ff00ff] text-black border-r border-slate-300 min-w-[75px]">BURSA TRANSFER</th>
                          <th className="py-2.5 px-2 bg-[#ff00ff] text-black border-r border-slate-300 min-w-[70px]">BURSA ÇIKAN</th>
                          <th className="py-2.5 px-2 bg-[#ff00ff] text-black border-r border-slate-300 min-w-[70px]">BURSA MEVCUT</th>
                        </>
                      )}

                      {/* MUAYENE (Gri) */}
                      {filterRegion === 'ALL' && (
                        <>
                          <th className="py-2.5 px-2 bg-[#bfbfbf] text-black border-r border-slate-300 min-w-[70px]">MUAYENE GİDEN</th>
                          <th className="py-2.5 px-2 bg-[#bfbfbf] text-black border-r border-slate-300 min-w-[70px]">MUAYENE GELEN</th>
                          <th className="py-2.5 px-2 bg-[#bfbfbf] text-black border-r border-slate-300 min-w-[70px]">MUAYENE TOPLAM</th>
                        </>
                      )}

                      {/* RAF ÖMRÜ SÜTUNLARI */}
                      <th className="py-2.5 px-2 bg-emerald-100 text-emerald-950 border-r border-slate-300 min-w-[95px] font-black">RAF ÖMRÜ VAR MI?</th>
                      <th className="py-2.5 px-2 bg-emerald-50 text-emerald-950 border-r border-slate-300 min-w-[110px] font-black">RAF ÖMRÜ BİTİŞ TARİHİ</th>

                      {/* İŞLEM */}
                      <th className="py-2.5 px-2 bg-slate-200 text-slate-800 min-w-[80px]">İŞLEM</th>
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-200 font-medium">
                    {paginatedSarfRows.map((item, idx) => {
                      const key = `${(item.partNumber || item.pn || '').trim().toLowerCase()}___${(item.description || item.name || '').trim().toLowerCase()}`;
                      const globalIdx = inventoryIndexMap.get(key) ?? -1;
                      const siraNo = (sarfPage - 1) * (sarfPageSize > 0 ? sarfPageSize : 0) + idx + 1;
                      return (
                        <tr 
                          key={idx} 
                          onClick={() => handleRowClick(item)}
                          className="hover:bg-emerald-50/80 cursor-pointer border-b border-slate-200 text-center font-mono text-[11px] transition select-none group"
                          title="Malzeme bilgilerini (P/N, S/N, Tanım, Lokasyon, Açıklama, Raf Ömrü) düzenlemek için tıklayınız (Şifre Korumalı)"
                        >
                          <td className="py-2 px-2 bg-slate-100 font-mono font-bold text-slate-700 border-r border-slate-200 sticky left-0 z-20 text-center">
                            {siraNo}
                          </td>
                          <td className="py-2 px-3 text-left font-sans font-bold text-slate-900 border-r border-slate-200 sticky left-[65px] bg-white z-20 relative group">
                            <div className="truncate max-w-[230px] group-hover:opacity-10 transition duration-150" title={item.description || item.name || '-'}>
                              {item.description || item.name || '-'}
                            </div>
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setSlipModalType('raf');
                                setSlipModalData({
                                  type: 'raf',
                                  lokasyonKodu: item.lokasyonNo || item.location || 'A-01-RAF-1',
                                  bolgeZone: item.category === 'kimyasal' ? 'KİMYASAL DEPOSU' : 'SARF MALZEME DEPOSU',
                                  rafTipi: 'KÜÇÜK PARÇA GÖZÜ',
                                  tanimIcerik: `${item.description || item.name} GÖZÜ`,
                                  guncelSayim: item.ankaraMevcut ?? item.toplamStok ?? 0,
                                  malzemeDesc: item.description || item.name || '',
                                  pn: item.partNumber || item.pn || '-',
                                  sn: item.serialAndNotes || item.sn || '-',
                                  depoYeri: item.lokasyonNo || item.location || 'ANKARA MERKEZ DEP'
                                });
                                setIsSlipModalOpen(true);
                              }}
                              className="absolute inset-0 flex items-center justify-center bg-slate-900/90 text-white text-[9px] font-extrabold uppercase rounded opacity-0 group-hover:opacity-100 transition duration-150 cursor-pointer z-20 whitespace-nowrap px-1 font-sans gap-1"
                              title="Bu malzeme için hızlıca Raf Slibi / Barkod Bas"
                            >
                              <Barcode className="w-3 h-3 text-emerald-400" />
                              <span>Slip Bas</span>
                            </button>
                          </td>
                          <td className="py-2 px-3 font-mono font-bold text-slate-800 border-r border-slate-200 sticky left-[305px] bg-white z-20 text-left">
                            {item.partNumber || item.pn || '-'}
                          </td>
                          <td className="py-2 px-3 text-left font-sans text-slate-600 border-r border-slate-200">
                            <div className="truncate max-w-[190px]" title={item.serialAndNotes || item.sn || '-'}>
                              {item.serialAndNotes || item.sn || '-'}
                            </div>
                          </td>
                          <td className="py-2 px-2 bg-yellow-50 font-bold text-slate-800 border-r border-slate-200">{item.lokasyonNo || item.location || '-'}</td>
                          {renderHoverableCell(item.gelen || 0, item, "bg-slate-100 text-slate-700", "GELEN")}
                          <td className="py-2 px-2 bg-cyan-100 font-black text-cyan-950 border-r border-slate-200 text-xs">{item.toplamStok || 0}</td>
                          {(filterRegion === 'ALL' || filterRegion === 'ANKARA') && (
                            <>
                              {renderHoverableCell(item.ankaraCikan || 0, item, "bg-amber-50 text-slate-600", "ANKARA ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.ankaraMevcut || 0) <= 0 && hadRegionHistory(item, 'ANKARA')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-amber-100 font-black text-amber-900'
                              }`}>
                                {item.ankaraMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'KARAİN') && (
                            <>
                              {renderHoverableCell(item.karainTransfer || 0, item, "bg-yellow-50 text-slate-600", "KARAİN TRANSFER")}
                              {renderHoverableCell(item.karainCikan || 0, item, "bg-yellow-50 text-slate-600", "KARAİN ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.karainMevcut || 0) <= 0 && hadRegionHistory(item, 'KARAİN')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-yellow-100 font-black text-yellow-950'
                              }`}>
                                {item.karainMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'ÇANAKKALE') && (
                            <>
                              {renderHoverableCell(item.canakkaleTransfer || 0, item, "bg-emerald-50 text-slate-600", "ÇANAKKALE TRANSFER")}
                              {renderHoverableCell(item.canakkaleCikan || 0, item, "bg-emerald-50 text-slate-600", "ÇANAKKALE ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.canakkaleMevcut || 0) <= 0 && hadRegionHistory(item, 'ÇANAKKALE')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-emerald-100 font-black text-emerald-900'
                              }`}>
                                {item.canakkaleMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'MİLAS') && (
                            <>
                              {renderHoverableCell(item.milasTransfer || 0, item, "bg-orange-50 text-slate-600", "MİLAS TRANSFER")}
                              {renderHoverableCell(item.milasCikan || 0, item, "bg-orange-50 text-slate-600", "MİLAS ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.milasMevcut || 0) <= 0 && hadRegionHistory(item, 'MİLAS')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-orange-100 font-black text-orange-950'
                              }`}>
                                {item.milasMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'BURSA') && (
                            <>
                              {renderHoverableCell(item.bursaTransfer || 0, item, "bg-fuchsia-50 text-slate-600", "BURSA TRANSFER")}
                              {renderHoverableCell(item.bursaCikan || 0, item, "bg-fuchsia-50 text-slate-600", "BURSA ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.bursaMevcut || 0) <= 0 && hadRegionHistory(item, 'BURSA')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-fuchsia-100 font-black text-fuchsia-950'
                              }`}>
                                {item.bursaMevcut || 0}
                              </td>
                            </>
                          )}
                          {filterRegion === 'ALL' && (
                            <>
                              {renderHoverableCell(item.muayeneGiden || 0, item, "bg-slate-50 text-slate-600", "MUAYENE GİDEN")}
                              {renderHoverableCell(item.muayeneGelen || 0, item, "bg-slate-50 text-slate-600", "MUAYENE GELEN")}
                              <td className="py-2 px-1.5 bg-slate-100 font-black text-slate-800 border-r border-slate-200">{item.muayeneToplam || 0}</td>
                            </>
                          )}

                          {/* RAF ÖMRÜ SÜTUNLARI */}
                          <td className="py-2 px-2 text-center border-r border-slate-200">
                            <span className={`inline-block px-2 py-0.5 rounded-full text-[10px] font-black uppercase ${
                              item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET')
                                ? 'bg-amber-100 text-amber-900 border border-amber-300'
                                : 'bg-slate-100 text-slate-600'
                            }`}>
                              {item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET') ? 'EVET' : 'HAYIR'}
                            </span>
                          </td>
                          <td className="py-2 px-2 text-center border-r border-slate-200 font-mono font-bold text-slate-700">
                            {(item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET')) && item.shelfLifeDate && item.shelfLifeDate !== '-'
                              ? item.shelfLifeDate
                              : '-'}
                          </td>

                          <td className="py-2 px-2 text-center bg-slate-50">
                            <div className="flex items-center justify-center gap-1">
                              <button 
                                type="button"
                                onClick={() => handleOpenEditItem(item)} 
                                className="p-1 hover:bg-slate-200 rounded text-slate-600 cursor-pointer" 
                                title="Düzenle"
                              >
                                ✏️
                              </button>
                              <button 
                                type="button"
                                onClick={() => handleDeleteItem(item)} 
                                className="p-1 hover:bg-rose-100 rounded text-rose-600 cursor-pointer" 
                                title="Sil"
                              >
                                🗑️
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <PaginationBar
                currentPage={sarfPage}
                totalItems={filteredSarfRows.length}
                pageSize={sarfPageSize}
                onPageChange={setSarfPage}
                onPageSizeChange={setSarfPageSize}
              />

              {filteredSarfRows.length === 0 && (
                <div className="py-14 text-center">
                  <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center mx-auto mb-2 text-xl">📦</div>
                  <h3 className="text-xs font-bold text-slate-700 uppercase">Kayıtlı Malzeme Bulunamadı</h3>
                  <p className="text-[11px] text-slate-500 mt-0.5">Yukarıdaki "+ Yeni Malzeme" veya "Excel Yükle" ile ekleyebilirsiniz.</p>
                </div>
              )}
            </div>
            )}
          </div>
        )}

        {/* VIEW 2: KİMYASAL DEPO */}
        {currentDepoType === 'kimyasal' && (
          <div className="space-y-4 font-sans">
            {!isDepoSearchPerformed && !searchQuery.trim() ? (
              <div className="py-14 text-center text-slate-500 flex flex-col items-center justify-center bg-white border border-slate-200 rounded-2xl shadow-xs select-none">
                <div className="w-12 h-12 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center mb-2.5 text-xl font-black">🧪</div>
                <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider font-sans">
                  {currentUnit.toUpperCase()} KİMYASAL DEPOSU
                </h4>
                <p className="text-[11px] text-slate-500 mt-1 max-w-md px-4 font-sans leading-relaxed">
                  Kimyasal depo veritabanında <strong>{computedInventory.filter(i => (currentUnit === 'all' || i.unit === currentUnit) && i.category === 'kimyasal').length} kalem</strong> ürün bulunmaktadır. Sistemi dondurmamak için yukarıdaki kutuya arama yazıp <strong>ARA</strong> butonuna basın veya aşağıdaki butonla tümünü listeleyin.
                </p>
                <button
                  type="button"
                  onClick={() => setIsDepoSearchPerformed(true)}
                  className="mt-4 px-5 py-2.5 bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-black uppercase tracking-wider rounded-xl transition cursor-pointer active:scale-95 shadow-sm font-sans flex items-center gap-2 font-bold"
                >
                  <span>Tüm Kimyasalları Listele ({computedInventory.filter(i => (currentUnit === 'all' || i.unit === currentUnit) && i.category === 'kimyasal').length} Kalem)</span>
                </button>
              </div>
            ) : (
            <div className="bg-white rounded-2xl shadow-xs border border-slate-200 p-4">
              <div className="flex flex-wrap items-center justify-between gap-3 mb-2 px-1">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-black text-slate-800 uppercase tracking-wider font-mono">
                    🧪 {currentUnit.toUpperCase()} KİMYASAL VE BOYA/YAĞ ENVANTERİ
                  </span>
                  <span className="text-[11px] text-slate-500 font-medium">
                    · Kimyasal Depo Dağıtımlı Canlı Stok
                  </span>
                </div>
                <div className="flex items-center gap-2 text-[11px]">
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#ffff00] border border-slate-300 inline-block rounded-xs"></span> Lokasyon/Karain</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#00ffff] border border-slate-300 inline-block rounded-xs"></span> Toplam</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#ff9900] border border-slate-300 inline-block rounded-xs"></span> Ankara/Milas</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#92d050] border border-slate-300 inline-block rounded-xs"></span> Çanakkale</span>
                  <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 bg-[#ff00ff] border border-slate-300 inline-block rounded-xs"></span> Bursa</span>
                </div>
              </div>

              <div className="overflow-x-auto max-h-[72vh]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 z-20 text-[11px] font-black uppercase text-center tracking-wider border-b-2 border-slate-300 select-none">
                    <tr className="border-b-2 border-slate-900">
                      <th className="py-3 px-2 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[55px] text-center sticky left-0 z-30 font-black">No</th>
                      <th className="py-3 px-3 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[220px] text-left sticky left-[55px] z-30 font-black">Description</th>
                      <th className="py-3 px-3 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[170px] font-mono sticky left-[275px] z-30 text-left font-black">P/N - SPAIN DISCRIPTION</th>
                      <th className="py-3 px-2 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[85px] font-black">MİKTAR</th>
                      <th className="py-3 px-2 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[110px] font-black">QTY SÖZLEŞME TOPLAMI</th>
                      <th className="py-3 px-2 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[95px] font-black">PİYASA GELEN</th>
                      <th className="py-3 px-2 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[95px] font-black">SPAIN GELEN</th>
                      <th className="py-3 px-2 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[100px] font-black">GENEL TOPLAM</th>
                      <th className="py-3 px-2 bg-[#00ffff] text-black border-r border-slate-800 min-w-[90px] font-black">TOPLAM STOK</th>
                      
                      {/* ANKARA (Turuncu - Görsel 1 ile Birebir) */}
                      {(filterRegion === 'ALL' || filterRegion === 'ANKARA') && (
                        <>
                          <th className="py-3 px-2 bg-[#ffa500] text-black border-r border-slate-800 min-w-[80px] font-black">ANKARA ÇIKAN</th>
                          <th className="py-3 px-2 bg-[#ffa500] text-black border-r border-slate-800 min-w-[80px] font-black">ANKARA MEVCUT</th>
                        </>
                      )}

                      {/* KARAİN (Sarı - Görsel 1 ile Birebir) */}
                      {(filterRegion === 'ALL' || filterRegion === 'KARAİN') && (
                        <>
                          <th className="py-3 px-2 bg-[#ffff00] text-black border-r border-slate-800 min-w-[80px] font-black">KARAİN TRANSFER</th>
                          <th className="py-3 px-2 bg-[#ffff00] text-black border-r border-slate-800 min-w-[75px] font-black">KARAİN ÇIKAN</th>
                          <th className="py-3 px-2 bg-[#ffff00] text-black border-r border-slate-800 min-w-[75px] font-black">KARAİN MEVCUT</th>
                        </>
                      )}

                      {/* ÇANAKKALE (Açık Yeşil - Görsel 1 ile Birebir) */}
                      {(filterRegion === 'ALL' || filterRegion === 'ÇANAKKALE') && (
                        <>
                          <th className="py-3 px-2 bg-[#90ee90] text-black border-r border-slate-800 min-w-[80px] font-black">ÇANAKKALE TRANSFER</th>
                          <th className="py-3 px-2 bg-[#90ee90] text-black border-r border-slate-800 min-w-[75px] font-black">ÇANAKKALE ÇIKAN</th>
                          <th className="py-3 px-2 bg-[#90ee90] text-black border-r border-slate-800 min-w-[75px] font-black">ÇANAKKALE MEVCUT</th>
                        </>
                      )}

                      {/* MİLAS (Koyu Turuncu - Görsel 1 ile Birebir) */}
                      {(filterRegion === 'ALL' || filterRegion === 'MİLAS') && (
                        <>
                          <th className="py-3 px-2 bg-[#ff8c00] text-black border-r border-slate-800 min-w-[80px] font-black">MİLAS TRANSFER</th>
                          <th className="py-3 px-2 bg-[#ff8c00] text-black border-r border-slate-800 min-w-[75px] font-black">MİLAS ÇIKAN</th>
                          <th className="py-3 px-2 bg-[#ff8c00] text-black border-r border-slate-800 min-w-[75px] font-black">MİLAS MEVCUT</th>
                        </>
                      )}

                      {/* BURSA (Mavi - Görsel 1 ile Birebir) */}
                      {(filterRegion === 'ALL' || filterRegion === 'BURSA') && (
                        <>
                          <th className="py-3 px-2 bg-[#4682b4] text-white border-r border-slate-800 min-w-[80px] font-black">BURSA TRANSFER</th>
                          <th className="py-3 px-2 bg-[#4682b4] text-white border-r border-slate-800 min-w-[75px] font-black">BURSA ÇIKAN</th>
                          <th className="py-3 px-2 bg-[#4682b4] text-white border-r border-slate-800 min-w-[75px] font-black">BURSA MEVCUT</th>
                        </>
                      )}

                      <th className="py-3 px-3 bg-[#b0b0b0] text-slate-950 border-r border-slate-800 min-w-[150px] text-left font-black">AÇIKLAMA</th>

                      {/* RAF ÖMRÜ SÜTUNLARI */}
                      <th className="py-2.5 px-2 bg-emerald-100 text-emerald-950 border-r border-slate-300 min-w-[95px] font-black">RAF ÖMRÜ VAR MI?</th>
                      <th className="py-2.5 px-2 bg-emerald-50 text-emerald-950 border-r border-slate-300 min-w-[110px] font-black">RAF ÖMRÜ BİTİŞ TARİHİ</th>

                      {/* İŞLEM */}
                      <th className="py-2.5 px-2 bg-slate-200 text-slate-800 min-w-[80px]">İŞLEM</th>
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-200 font-medium">
                    {paginatedKimyasalRows.map((item, idx) => {
                      const siraNo = (kimyasalPage - 1) * (kimyasalPageSize > 0 ? kimyasalPageSize : 0) + idx + 1;
                      return (
                        <tr 
                          key={idx} 
                          onClick={() => handleRowClick(item)}
                          className="hover:bg-purple-50/80 cursor-pointer border-b border-slate-200 text-center font-mono text-[11px] transition select-none group"
                          title="Malzeme bilgilerini (P/N, S/N, Tanım, Lokasyon, Açıklama, Raf Ömrü) düzenlemek için tıklayınız (Şifre Korumalı)"
                        >
                          <td className="py-2 px-2 bg-slate-100 font-mono font-bold text-slate-700 border-r border-slate-200 sticky left-0 z-20 text-center">
                            {siraNo}
                          </td>
                          <td className="py-2 px-3 text-left font-sans font-bold text-slate-900 border-r border-slate-200 sticky left-[55px] bg-white z-20 relative group">
                            <div className="truncate max-w-[210px] group-hover:opacity-10 transition duration-150" title={item.description || item.name || '-'}>
                              {item.description || item.name || '-'}
                            </div>
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setSlipModalType('raf');
                                setSlipModalData({
                                  type: 'raf',
                                  lokasyonKodu: item.lokasyonNo || item.location || 'KİMYASAL DEPO',
                                  bolgeZone: 'KİMYASAL DEPOSU',
                                  rafTipi: 'KİMYASAL RAF / SIVI GÖZÜ',
                                  tanimIcerik: `${item.description || item.name} GÖZÜ`,
                                  guncelSayim: item.ankaraMevcut ?? item.gelen ?? item.toplamStok ?? 0,
                                  malzemeDesc: item.description || item.name || '',
                                  pn: item.partNumber || item.pn || '-',
                                  sn: item.serialAndNotes || item.sn || '-',
                                  depoYeri: 'KİMYASAL DEPO'
                                });
                                setIsSlipModalOpen(true);
                              }}
                              className="absolute inset-0 flex items-center justify-center bg-slate-900/90 text-white text-[9px] font-extrabold uppercase rounded opacity-0 group-hover:opacity-100 transition duration-150 cursor-pointer z-20 whitespace-nowrap px-1 font-sans gap-1"
                              title="Bu kimyasal malzeme için hızlıca Raf Slibi / Barkod Bas"
                            >
                              <Barcode className="w-3 h-3 text-emerald-400" />
                              <span>Slip Bas</span>
                            </button>
                          </td>
                          <td className="py-2 px-3 font-mono font-bold text-slate-800 border-r border-slate-200 sticky left-[275px] bg-white z-20 text-left">
                            {item.partNumber && item.partNumber !== '-' ? item.partNumber : (item.spainDescription || item.pn || '-')}
                          </td>
                          <td className="py-2 px-2 bg-amber-50/60 font-bold text-slate-800 border-r border-slate-200">
                            {item.miktarQty || 'Adet'}
                          </td>
                          <td className="py-2 px-2 text-slate-700 border-r border-slate-200 font-mono">
                            {item.sozlesmeToplami || '0'}
                          </td>
                          {renderHoverableCell(item.piyasa || 0, item, "bg-slate-100 text-slate-700", "PİYASA GELEN")}
                          {renderHoverableCell(item.spainGelen || 0, item, "bg-emerald-50/70 text-emerald-950", "SPAIN GELEN")}
                          <td className="py-2 px-2 bg-indigo-50/70 font-bold text-indigo-950 border-r border-slate-200 font-mono">
                            {item.genelToplam !== undefined ? item.genelToplam : (item.toplamStok || 0)}
                          </td>
                          <td className="py-2 px-2 bg-cyan-100 font-black text-cyan-950 border-r border-slate-200 text-xs font-mono">{item.toplamStok || 0}</td>
                          {(filterRegion === 'ALL' || filterRegion === 'ANKARA') && (
                            <>
                              {renderHoverableCell(item.ankaraCikan || 0, item, "bg-amber-50 text-slate-600", "ANKARA ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.ankaraMevcut || 0) <= 0 && hadRegionHistory(item, 'ANKARA')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-amber-100 font-black text-amber-900'
                              }`}>
                                {item.ankaraMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'KARAİN') && (
                            <>
                              {renderHoverableCell(item.karainTransfer || 0, item, "bg-yellow-50 text-slate-600", "KARAİN TRANSFER")}
                              {renderHoverableCell(item.karainCikan || 0, item, "bg-yellow-50 text-slate-600", "KARAİN ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.karainMevcut || 0) <= 0 && hadRegionHistory(item, 'KARAİN')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-yellow-100 font-black text-yellow-950'
                              }`}>
                                {item.karainMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'ÇANAKKALE') && (
                            <>
                              {renderHoverableCell(item.canakkaleTransfer || 0, item, "bg-emerald-50 text-slate-600", "ÇANAKKALE TRANSFER")}
                              {renderHoverableCell(item.canakkaleCikan || 0, item, "bg-emerald-50 text-slate-600", "ÇANAKKALE ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.canakkaleMevcut || 0) <= 0 && hadRegionHistory(item, 'ÇANAKKALE')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-emerald-100 font-black text-emerald-900'
                              }`}>
                                {item.canakkaleMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'MİLAS') && (
                            <>
                              {renderHoverableCell(item.milasTransfer || 0, item, "bg-orange-50 text-slate-600", "MİLAS TRANSFER")}
                              {renderHoverableCell(item.milasCikan || 0, item, "bg-orange-50 text-slate-600", "MİLAS ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.milasMevcut || 0) <= 0 && hadRegionHistory(item, 'MİLAS')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-orange-100 font-black text-orange-950'
                              }`}>
                                {item.milasMevcut || 0}
                              </td>
                            </>
                          )}
                          {(filterRegion === 'ALL' || filterRegion === 'BURSA') && (
                            <>
                              {renderHoverableCell(item.bursaTransfer || 0, item, "bg-fuchsia-50 text-slate-600", "BURSA TRANSFER")}
                              {renderHoverableCell(item.bursaCikan || 0, item, "bg-fuchsia-50 text-slate-600", "BURSA ÇIKAN")}
                              <td className={`py-2 px-1.5 border-r border-slate-200 text-center font-black ${
                                (item.bursaMevcut || 0) <= 0 && hadRegionHistory(item, 'BURSA')
                                  ? 'bg-rose-600 text-white animate-pulse shadow-xs font-black'
                                  : 'bg-fuchsia-100 font-black text-fuchsia-950'
                              }`}>
                                {item.bursaMevcut || 0}
                              </td>
                            </>
                          )}

                          {/* AÇIKLAMA */}
                          <td className="py-2 px-3 text-left font-sans text-slate-600 border-r border-slate-200">
                            <div className="truncate max-w-[170px]" title={item.aciklama || item.serialAndNotes || item.notes || '-'}>
                              {item.aciklama || item.serialAndNotes || item.notes || '-'}
                            </div>
                          </td>

                          {/* RAF ÖMRÜ SÜTUNLARI */}
                          <td className="py-2 px-2 text-center border-r border-slate-200">
                            <span className={`inline-block px-2 py-0.5 rounded-full text-[10px] font-black uppercase ${
                              item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET')
                                ? 'bg-amber-100 text-amber-900 border border-amber-300'
                                : 'bg-slate-100 text-slate-600'
                            }`}>
                              {item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET') ? 'EVET' : 'HAYIR'}
                            </span>
                          </td>
                          <td className="py-2 px-2 text-center border-r border-slate-200 font-mono font-bold text-slate-700">
                            {(item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET')) && item.shelfLifeDate && item.shelfLifeDate !== '-'
                              ? item.shelfLifeDate
                              : '-'}
                          </td>

                          <td className="py-2 px-2 text-center bg-slate-50">
                            <div className="flex items-center justify-center gap-1">
                              <button 
                                type="button"
                                onClick={() => handleOpenEditItem(item)} 
                                className="p-1 hover:bg-slate-200 rounded text-slate-600 cursor-pointer" 
                                title="Düzenle"
                              >
                                ✏️
                              </button>
                              <button 
                                type="button"
                                onClick={() => handleDeleteItem(item)} 
                                className="p-1 hover:bg-rose-100 rounded text-rose-600 cursor-pointer" 
                                title="Sil"
                              >
                                🗑️
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <PaginationBar
                currentPage={kimyasalPage}
                totalItems={filteredKimyasalRows.length}
                pageSize={kimyasalPageSize}
                onPageChange={setKimyasalPage}
                onPageSizeChange={setKimyasalPageSize}
              />

              {filteredKimyasalRows.length === 0 && (
                <div className="py-12 text-center text-xs text-slate-500">
                  Bu birime ait kimyasal depo kaydı bulunamadı.
                </div>
              )}
            </div>
            )}

            {/* DEPO HAREKET GEÇMİŞİ (Kimyasal depoda da her zaman görünür - REMOVED) */}
          </div>
        )}

        {/* VIEW 2.5: YAŞAM DESTEK (AT-802 ÖZEL) */}
        {currentDepoType === 'yasam_destek' && (
          <div className="space-y-4">
            <YasamDestekPanel
              records={yasamDestekRecords}
              isLoading={isYasamDestekLoading}
              onRefresh={fetchYasamDestekData}
              onSaveRecord={handleSaveYasamDestekRecord}
              onDeleteRecord={handleDeleteYasamDestekRecord}
              showNotification={showNotification}
              triggerExport={isYasamDestekExportTriggered}
              onExportComplete={() => setIsYasamDestekExportTriggered(false)}
            />
          </div>
        )}

        {/* VIEW 3: DEPO HAREKET GEÇMİŞİ (Dedicated Tab) */}
        {currentDepoType === 'transactions' && (
          <div className="space-y-4">
            {renderTransactionsCard(true)}
          </div>
        )}

      </main>

      {/* ========================================================================= */}
      {/* 5. MODALLER */}
      {/* ========================================================================= */}

      {/* BİRİM ŞİFRE KORUMA MODALI */}
      {isPasswordModalOpen && (
        <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full border border-slate-200 overflow-hidden">
            <div className="bg-[#0b3d1d] text-white p-4 text-center border-b border-emerald-500/30">
              <div className="w-10 h-10 rounded-xl bg-white/10 flex items-center justify-center mx-auto mb-2 text-xl">🔒</div>
              <h3 className="text-xs font-black uppercase tracking-wider text-white">Birim Güvenlik Girişi</h3>
              <p className="text-[10px] text-emerald-200/80">Teknik yayınlar ve filo şifreleme protokolü</p>
            </div>

            <form onSubmit={handlePasswordSubmit} className="p-4 space-y-3">
              <div>
                <label className="block text-[11px] font-bold text-slate-700 uppercase font-mono mb-1">ŞİFRE</label>
                <input 
                  type="password" 
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  required 
                  placeholder="Birim şifresini giriniz..." 
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-mono tracking-widest text-center font-black focus:outline-none focus:border-[#0b3d1d]" 
                />
                {passwordError && (
                  <div className="text-xs text-rose-600 font-bold mt-1 text-center">❌ Hatalı şifre girdiniz!</div>
                )}
              </div>

              <div className="bg-slate-50 p-2 rounded-xl border border-slate-200 text-[10px] text-slate-500 space-y-0.5">
                <div className="flex justify-between"><span>AT-802: <strong>802</strong></span><span>BELL 429: <strong>429</strong></span><span>T-70: <strong>70</strong></span></div>
                <div className="flex justify-between pt-1 border-t border-slate-200"><span>C-650: <strong>650</strong></span><span>B-360: <strong>360</strong></span><span>Hangar: <strong>1839</strong></span></div>
              </div>

              <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                <button type="button" onClick={() => setIsPasswordModalOpen(false)} className="px-3 py-1.5 bg-slate-100 text-slate-700 text-xs font-bold rounded-lg cursor-pointer">İptal</button>
                <button type="submit" className="px-4 py-1.5 bg-[#0b3d1d] text-white text-xs font-black uppercase rounded-lg shadow-xs cursor-pointer">Kilidi Aç</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* HAREKET FORMU (ÇIKIŞ, GİRİŞ, TRANSFER) */}
      {isMovementModalOpen && (() => {
        const selectedItemForForm = selectedItemIdx >= 0 && selectedItemIdx < computedInventory.length
          ? computedInventory[selectedItemIdx]
          : (movementItemSearch.trim()
              ? computedInventory.find(i => {
                  const q = movementItemSearch.toLowerCase().trim();
                  const d = (i.description || i.name || '').toLowerCase();
                  const p = (i.partNumber || i.pn || '').toLowerCase();
                  return d.includes(q) || p.includes(q);
                }) || null
              : null);

        const selectedCikisDepotInfo = getCikisDepotInfo(movementType, selectedItemForForm);
        const isTransferOverStock = movementDirection === 'transfer' && selectedItemForForm && movementQty > Number(selectedItemForForm.ankaraMevcut || 0);
        const isCikisOverStock = movementDirection === 'cikis' && selectedItemForForm && movementQty > selectedCikisDepotInfo.stock;

        return (
          <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full border border-slate-200 overflow-hidden">
              <div className={`px-5 py-3.5 flex items-center justify-between text-white ${
                movementDirection === 'cikis' ? 'bg-[#a0153e]' : (movementDirection === 'giris' ? 'bg-[#006a4e]' : 'bg-indigo-700')
              }`}>
                <h3 className="text-xs font-black uppercase tracking-wider">
                  {movementDirection === 'cikis' ? '↗ MALZEME ÇIKIŞI (SARF / TÜKETİM)' : (movementDirection === 'giris' ? '↙ MALZEME GİRİŞİ (ANKARA GİREN)' : '⇄ DEPOLAR ARASI TRANSFER')}
                </h3>
                <button type="button" onClick={() => setIsMovementModalOpen(false)} className="text-white hover:text-slate-300 font-bold cursor-pointer">✕</button>
              </div>

              <form onSubmit={handleMovementSubmit} className="p-5 space-y-3.5 text-xs">
                {movementDirection === 'transfer' ? (
                  /* ─── AŞAMALI DEPOLAR ARASI TRANSFER FORMU ─── */
                  <div className="space-y-3.5">
                    {/* AŞAMA 1: TRANSFER MERKEZİ (ÇIKIŞ / NEREDEN) - İLK AÇILIŞTA HEP ANKARA SEÇİLİDİR */}
                    <div className="p-3 bg-indigo-50/70 border border-indigo-200 rounded-xl">
                      <label className="block font-black text-indigo-950 uppercase font-mono mb-1.5 flex items-center gap-2">
                        <span className="w-5 h-5 rounded-full bg-indigo-700 text-white flex items-center justify-center text-[10px] font-black shrink-0">1</span>
                        <span>TRANSFER MERKEZİ (ÇIKIŞ DEPOSU / NEREDEN) *</span>
                      </label>
                      <select 
                        value={transferSourceDepot}
                        onChange={(e) => {
                          const val = e.target.value;
                          setTransferSourceDepot(val);
                          if (transferTargetDepot === val) {
                            setTransferTargetDepot(val === 'ANKARA' ? 'KARAİN' : 'ANKARA');
                          }
                        }}
                        className="w-full px-3 py-2 bg-white border-2 border-indigo-400 rounded-xl font-black text-indigo-950 text-xs focus:outline-none focus:border-indigo-600 shadow-xs"
                      >
                        <option value="ANKARA">{currentDepoType === 'kimyasal' || selectedItemForForm?.category === 'kimyasal' ? 'KİMYASAL DEPO - ANKARA' : 'ANKARA (MERKEZ DEPO)'}</option>
                        <option value="KARAİN">{currentDepoType === 'kimyasal' || selectedItemForForm?.category === 'kimyasal' ? 'KİMYASAL DEPO - KARAİN' : 'KARAİN DEPOSU'}</option>
                        <option value="ÇANAKKALE">{currentDepoType === 'kimyasal' || selectedItemForForm?.category === 'kimyasal' ? 'KİMYASAL DEPO - ÇANAKKALE' : 'ÇANAKKALE DEPOSU'}</option>
                        <option value="MİLAS">{currentDepoType === 'kimyasal' || selectedItemForForm?.category === 'kimyasal' ? 'KİMYASAL DEPO - MİLAS' : 'MİLAS DEPOSU'}</option>
                        <option value="BURSA">{currentDepoType === 'kimyasal' || selectedItemForForm?.category === 'kimyasal' ? 'KİMYASAL DEPO - BURSA' : 'BURSA DEPOSU'}</option>
                      </select>
                      <span className="text-[10px] text-indigo-600 font-sans mt-1 block">
                        ℹ️ Malzemenin çıkış yapılacağı ve stoktan düşüleceği başlangıç merkezidir.
                      </span>
                    </div>

                    {/* AŞAMA 2: PN VE MALZEME SEÇİMİ */}
                    <div className="relative p-3 bg-slate-50 border border-slate-200 rounded-xl">
                      <label className="block font-black text-slate-800 uppercase font-mono mb-1.5 flex items-center gap-2">
                        <span className="w-5 h-5 rounded-full bg-indigo-700 text-white flex items-center justify-center text-[10px] font-black shrink-0">2</span>
                        <span>PARÇA / P/N SEÇİMİ (MALZEME) *</span>
                      </label>
                      <div className="relative">
                        <input 
                          type="text"
                          value={movementItemSearch}
                          onChange={(e) => {
                            setMovementItemSearch(e.target.value);
                            setIsItemDropdownOpen(true);
                          }}
                          onFocus={() => setIsItemDropdownOpen(true)}
                          required
                          placeholder="P/N veya Malzeme adı arayın (Örn: 013-00066, ENCODER)..."
                          className="w-full pl-8 pr-3 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-none focus:border-indigo-600 transition"
                        />
                        <Search className="w-4 h-4 absolute left-2.5 top-2.5 text-slate-400" />
                      </div>

                      {isItemDropdownOpen && (
                        <div className="absolute top-full left-3 right-3 z-50 mt-1 bg-white border-2 border-indigo-600 rounded-xl shadow-2xl max-h-52 overflow-y-auto divide-y divide-slate-100">
                          {computedInventory.filter(i => {
                            if (!movementItemSearch.trim()) return true;
                            const q = movementItemSearch.toLowerCase().trim();
                            const d = (i.description || i.name || '').toLowerCase();
                            const p = (i.partNumber || i.pn || '').toLowerCase();
                            return d.includes(q) || p.includes(q);
                          }).slice(0, 10).map((m, idx) => (
                            <div 
                              key={idx} 
                              onClick={() => handleSelectItemForMovement(m, idx)}
                              className="p-2.5 hover:bg-indigo-50 cursor-pointer text-xs flex items-center justify-between gap-2 transition"
                            >
                              <div>
                                <div className="font-bold text-slate-900">{m.description || m.name}</div>
                                <div className="text-[11px] font-mono text-slate-500">
                                  P/N: <strong className="text-indigo-700">{m.partNumber || m.pn || '-'}</strong> | S/N: {m.serialAndNotes || m.sn || '-'}
                                </div>
                              </div>
                              <span className="text-[10px] font-mono font-bold bg-indigo-100 text-indigo-950 px-2 py-0.5 rounded-full shrink-0">
                                {transferSourceDepot} Mevcudu: {
                                  transferSourceDepot === 'ANKARA' ? (m.ankaraMevcut || 0) :
                                  transferSourceDepot === 'KARAİN' ? (m.karainMevcut || 0) :
                                  transferSourceDepot === 'ÇANAKKALE' ? (m.canakkaleMevcut || 0) :
                                  transferSourceDepot === 'MİLAS' ? (m.milasMevcut || 0) :
                                  transferSourceDepot === 'BURSA' ? (m.bursaMevcut || 0) : (m.toplamStok || 0)
                                }
                              </span>
                            </div>
                          ))}
                        </div>
                      )}

                      {/* Seçilen Malzeme Kartı */}
                      {selectedItemForForm && (() => {
                        const src = transferSourceDepot.toUpperCase();
                        let srcStock = Number(selectedItemForForm.ankaraMevcut || 0);
                        if (src.includes('KARAİN') || src.includes('KARAIN')) srcStock = Number(selectedItemForForm.karainMevcut || 0);
                        else if (src.includes('ÇANAKKALE')) srcStock = Number(selectedItemForForm.canakkaleMevcut || 0);
                        else if (src.includes('MİLAS')) srcStock = Number(selectedItemForForm.milasMevcut || 0);
                        else if (src.includes('BURSA')) srcStock = Number(selectedItemForForm.bursaMevcut || 0);
                        const isOver = movementQty > srcStock;

                        return (
                          <div className={`mt-2 p-2.5 rounded-xl border text-xs flex items-center justify-between ${
                            isOver ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-indigo-50 border-indigo-200 text-indigo-900'
                          }`}>
                            <div>
                              <div className="font-bold flex items-center gap-1.5">
                                <span>📦 {selectedItemForForm.description || selectedItemForForm.name}</span>
                              </div>
                              <div className="text-[11px] font-mono text-slate-600 mt-0.5">
                                P/N: <strong>{selectedItemForForm.partNumber || selectedItemForForm.pn || '-'}</strong> | S/N: {selectedItemForForm.serialAndNotes || selectedItemForForm.sn || '-'}
                              </div>
                            </div>
                            <div className="text-right shrink-0 ml-3">
                              <div className="text-[10px] font-mono font-bold uppercase tracking-wider text-indigo-700">{transferSourceDepot} Stoğu</div>
                              <div className="text-sm font-black font-mono text-indigo-950">{srcStock} ADET</div>
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    {/* AŞAMA 3 & 4: ADET SEÇİMİ VE TARİH SEÇİMİ */}
                    <div className="grid grid-cols-2 gap-3">
                      <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                        <label className="block font-black text-slate-800 uppercase font-mono mb-1.5 flex items-center gap-1.5">
                          <span className="w-5 h-5 rounded-full bg-indigo-700 text-white flex items-center justify-center text-[10px] font-black shrink-0">3</span>
                          <span>ADET (MİKTAR) *</span>
                        </label>
                        <input 
                          type="number" 
                          min={1} 
                          value={movementQty}
                          onChange={(e) => setMovementQty(Math.max(1, parseInt(e.target.value) || 1))}
                          required 
                          className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl font-black font-mono text-slate-900" 
                        />
                      </div>

                      <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                        <label className="block font-black text-slate-800 uppercase font-mono mb-1.5 flex items-center justify-between">
                          <span className="flex items-center gap-1.5">
                            <span className="w-5 h-5 rounded-full bg-indigo-700 text-white flex items-center justify-center text-[10px] font-black shrink-0">4</span>
                            <span>TARİH SEÇİMİ *</span>
                          </span>
                          <span className="text-[10px] text-indigo-700 font-sans cursor-pointer font-bold hover:underline" onClick={() => {
                            const now = new Date();
                            const pad = (n: number) => String(n).padStart(2, '0');
                            setMovementDate(`${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
                          }}>📅 Bugün</span>
                        </label>
                        <div className="flex items-center gap-1.5">
                          <input 
                            type="text" 
                            value={movementDate}
                            onChange={(e) => setMovementDate(e.target.value)}
                            placeholder="GG.AA.YYYY SS:DK"
                            className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl font-mono text-slate-700 text-xs font-bold" 
                          />
                          <input 
                            type="date" 
                            onChange={(e) => {
                              if (e.target.value) {
                                const [yyyy, mm, dd] = e.target.value.split('-');
                                const now = new Date();
                                const pad = (n: number) => String(n).padStart(2, '0');
                                setMovementDate(`${dd}.${mm}.${yyyy} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
                              }
                            }}
                            className="p-2 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-xl text-xs cursor-pointer shrink-0" 
                            title="Takvimden Tarih Seç"
                          />
                        </div>
                      </div>
                    </div>

                    {/* AŞAMA 5: TESLİM ALAN SEÇİMİ (KABUL YAPAN YOK, DEPO YERİ YOK) */}
                    <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                      <label className="block font-black text-slate-800 uppercase font-mono mb-1.5 flex items-center gap-2">
                        <span className="w-5 h-5 rounded-full bg-indigo-700 text-white flex items-center justify-center text-[10px] font-black shrink-0">5</span>
                        <span>TESLİM ALAN PERSONEL *</span>
                      </label>
                      <PersonnelAutocomplete 
                        value={movementTeslimAlan}
                        onChange={setMovementTeslimAlan}
                        placeholder="Teslim alan personeli seçiniz veya yazınız"
                      />
                    </div>

                    {/* AŞAMA 6: TRANSFER EDİLEN MERKEZ (HEDEF / NEREYE) */}
                    <div className="p-3 bg-indigo-50/70 border border-indigo-200 rounded-xl">
                      <label className="block font-black text-indigo-950 uppercase font-mono mb-1.5 flex items-center gap-2">
                        <span className="w-5 h-5 rounded-full bg-indigo-700 text-white flex items-center justify-center text-[10px] font-black shrink-0">6</span>
                        <span>TRANSFER EDİLEN MERKEZ (HEDEF DEPO / NEREYE) *</span>
                      </label>
                      <select 
                        value={transferTargetDepot}
                        onChange={(e) => {
                          setTransferTargetDepot(e.target.value);
                          setMovementType(`${e.target.value} TRANSFER`);
                        }}
                        className="w-full px-3 py-2 bg-white border-2 border-indigo-400 rounded-xl font-black text-indigo-950 text-xs focus:outline-none focus:border-indigo-600 shadow-xs"
                      >
                        {['ANKARA', 'KARAİN', 'ÇANAKKALE', 'MİLAS', 'BURSA']
                          .filter(depot => depot !== transferSourceDepot)
                          .map((d, i) => (
                            <option key={i} value={d}>
                              {currentDepoType === 'kimyasal' || selectedItemForForm?.category === 'kimyasal'
                                ? `KİMYASAL DEPO - ${d}`
                                : (d === 'ANKARA' ? 'ANKARA (MERKEZ DEPO)' : `${d} DEPOSU`)}
                            </option>
                          ))}
                      </select>
                    </div>

                    {/* İsteğe Bağlı Transfer Notu */}
                    <div>
                      <label className="block text-[10px] font-bold text-slate-700 uppercase font-mono mb-1">
                        Açıklama / Transfer Notu (İsteğe Bağlı)
                      </label>
                      <textarea 
                        rows={2}
                        value={movementNotes}
                        onChange={(e) => setMovementNotes(e.target.value)}
                        placeholder={`${transferSourceDepot} ➔ ${transferTargetDepot} transfer açıklaması...`}
                        className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-sans focus:outline-none focus:border-indigo-600 resize-none"
                      />
                    </div>
                  </div>
                ) : (
                  /* ─── NORMAL GİRİŞ & ÇIKIŞ FORMU ─── */
                  <>
                    {/* MALZEME ARAMA & AUTOCOMPLETE */}
                    <div className="relative">
                      <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                        MALZEME (DESCRIPTION / P/N) *
                      </label>
                      <div className="relative">
                        <input 
                          type="text"
                          value={movementItemSearch}
                          onChange={(e) => {
                            setMovementItemSearch(e.target.value);
                            setIsItemDropdownOpen(true);
                          }}
                          onFocus={() => setIsItemDropdownOpen(true)}
                          required
                          placeholder="P/N veya Malzeme adı yazın (Örn: ENCODER, 013-00066-00)..."
                          className="w-full pl-8 pr-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-none focus:border-emerald-600 transition"
                        />
                        <Search className="w-4 h-4 absolute left-2.5 top-2.5 text-slate-400" />
                      </div>

                      {isItemDropdownOpen && (
                        <div className="absolute top-full left-0 right-0 z-50 mt-1 bg-white border-2 border-emerald-600 rounded-xl shadow-2xl max-h-52 overflow-y-auto divide-y divide-slate-100">
                          {computedInventory.filter(i => {
                            if (!movementItemSearch.trim()) return true;
                            const q = movementItemSearch.toLowerCase().trim();
                            const d = (i.description || i.name || '').toLowerCase();
                            const p = (i.partNumber || i.pn || '').toLowerCase();
                            return d.includes(q) || p.includes(q);
                          }).slice(0, 10).map((m, idx) => (
                            <div 
                              key={idx} 
                              onClick={() => handleSelectItemForMovement(m, idx)}
                              className="p-2.5 hover:bg-emerald-50 cursor-pointer text-xs flex items-center justify-between gap-2 transition"
                            >
                              <div>
                                <div className="font-bold text-slate-900">{m.description || m.name}</div>
                                <div className="text-[11px] font-mono text-slate-500">
                                  P/N: <strong className="text-emerald-700">{m.partNumber || m.pn || '-'}</strong> | S/N: {m.serialAndNotes || m.sn || '-'}
                                </div>
                              </div>
                              <span className="text-[10px] font-mono font-bold bg-amber-100 text-amber-950 px-2 py-0.5 rounded-full shrink-0">
                                Ankara Depo Mevcudu: {m.ankaraMevcut || 0}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}

                      {/* Seçilen Malzeme ve Depo Stok Kartı */}
                      {selectedItemForForm && (
                        <div className={`mt-2 p-2.5 rounded-xl border text-xs flex items-center justify-between ${
                          movementDirection === 'cikis'
                            ? (isCikisOverStock ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-amber-50 border-amber-200 text-amber-900')
                            : 'bg-emerald-50 border-emerald-200 text-emerald-900'
                        }`}>
                          <div>
                            <div className="font-bold flex items-center gap-1.5">
                              <span>📦 {selectedItemForForm.description || selectedItemForForm.name}</span>
                            </div>
                            <div className="text-[11px] font-mono text-slate-600 mt-0.5">
                              P/N: <strong>{selectedItemForForm.partNumber || selectedItemForForm.pn || '-'}</strong>
                            </div>
                          </div>

                          <div className="text-right shrink-0 ml-3">
                            {movementDirection === 'cikis' ? (
                              <>
                                <div className="text-[10px] font-mono font-bold uppercase tracking-wider text-amber-800">
                                  {selectedCikisDepotInfo.name} Mevcudu
                                </div>
                                <div className="text-sm font-black font-mono text-amber-950">
                                  {selectedCikisDepotInfo.stock} ADET
                                </div>
                              </>
                            ) : (
                              <>
                                <div className="text-[10px] font-mono font-bold uppercase tracking-wider text-emerald-700">Giriş Deposu</div>
                                <div className="text-sm font-black font-mono text-emerald-950">ANKARA MERKEZ DEPO</div>
                              </>
                            )}
                          </div>
                        </div>
                      )}

                      {isCikisOverStock && (
                        <div className="mt-1.5 p-2 bg-rose-100 border border-rose-300 rounded-lg text-rose-800 font-bold text-[11px] flex items-center gap-1.5">
                          <span>⚠️ Hatalı Çıkış Engeli: {selectedCikisDepotInfo.name}'nda yalnızca {selectedCikisDepotInfo.stock} adet bulunmaktadır. {movementQty} adet çıkış yapılamaz!</span>
                        </div>
                      )}
                    </div>

                    {/* Miktar ve Tarih */}
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block font-bold text-slate-700 uppercase font-mono mb-1">ADET (MİKTAR) *</label>
                        <input 
                          type="number" 
                          min={1} 
                          value={movementQty}
                          onChange={(e) => setMovementQty(Math.max(1, parseInt(e.target.value) || 1))}
                          required 
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold font-mono" 
                        />
                      </div>
                      <div>
                        <label className="block font-bold text-slate-700 uppercase font-mono mb-1 flex items-center justify-between">
                          <span>TARİH</span>
                          <span className="text-[10px] text-emerald-700 font-sans cursor-pointer font-bold hover:underline" onClick={() => {
                            const now = new Date();
                            const pad = (n: number) => String(n).padStart(2, '0');
                            setMovementDate(`${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
                          }}>📅 Bugün</span>
                        </label>
                        <div className="flex items-center gap-1.5">
                          <input 
                            type="text" 
                            value={movementDate}
                            onChange={(e) => setMovementDate(e.target.value)}
                            placeholder="GG.AA.YYYY SS:DK"
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-mono text-slate-700 text-xs font-bold" 
                          />
                          <input 
                            type="date" 
                            onChange={(e) => {
                              if (e.target.value) {
                                const [yyyy, mm, dd] = e.target.value.split('-');
                                const now = new Date();
                                const pad = (n: number) => String(n).padStart(2, '0');
                                setMovementDate(`${dd}.${mm}.${yyyy} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
                              }
                            }}
                            className="p-2 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-xl text-xs cursor-pointer shrink-0" 
                            title="Takvimden Tarih Seç"
                          />
                        </div>
                      </div>
                    </div>

                    {/* İŞLEM TÜRÜ / TRANSFER TÜRÜ */}
                    <div>
                      <label className="block font-bold text-slate-700 uppercase font-mono mb-1">İŞLEM / TRANSFER TÜRÜ *</label>
                      <select 
                        value={movementType}
                        onChange={(e) => setMovementType(e.target.value)}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-800 focus:outline-none focus:border-emerald-600"
                      >
                        {availableMovementTypes.map((t, i) => (
                          <option key={i} value={t}>{t}</option>
                        ))}
                      </select>

                      {movementType === 'DİĞER' && (
                        <div className="mt-2 p-2.5 bg-amber-50 border border-amber-200 rounded-xl">
                          <label className="block text-[10px] font-bold text-amber-900 uppercase font-mono mb-1">Yeni Transfer Türü Yazın:</label>
                          <input 
                            type="text" 
                            value={movementCustomType}
                            onChange={(e) => setMovementCustomType(e.target.value)}
                            placeholder="Örn: ESKİŞEHİR TRANSFER..." 
                            className="w-full px-2.5 py-1.5 bg-white border border-amber-300 rounded-lg text-xs font-bold text-amber-950 uppercase" 
                          />
                        </div>
                      )}
                    </div>

                    {/* SERİ NO VE KUYRUK KODU (KUYRUK NO SEÇİMİNDE 'DİĞER' VE MANUEL YAZMA DESTEĞİ) */}
                    <div className={`grid ${movementDirection === 'cikis' ? 'grid-cols-2' : 'grid-cols-1'} gap-3`}>
                      <div>
                        <label className="block font-bold text-slate-700 uppercase font-mono mb-1">SERİAL NUMBER (S/N)</label>
                        <input 
                          type="text" 
                          value={movementSn}
                          onChange={(e) => setMovementSn(e.target.value)}
                          placeholder="S/N..." 
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-mono text-xs font-bold" 
                        />
                      </div>

                      {movementDirection === 'cikis' && (
                        <div>
                          <label className="block font-bold text-slate-700 uppercase font-mono mb-1">KUYRUK KODU / ARAÇ</label>
                          <select 
                            value={movementTail}
                            onChange={(e) => {
                              setMovementTail(e.target.value);
                              if (e.target.value !== 'DIGER') setCustomTailInput('');
                            }}
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-semibold"
                          >
                            {(AIRCRAFT_TAILS[currentUnit === 'all' ? 'at802' : currentUnit] || AIRCRAFT_TAILS.hangar).map((t, i) => (
                              <option key={i} value={t}>{t}</option>
                            ))}
                            <option value="DIGER">Diğer (Manuel Giriş)...</option>
                          </select>

                          {movementTail === 'DIGER' && (
                            <div className="mt-1.5">
                              <input 
                                type="text"
                                value={customTailInput}
                                onChange={(e) => setCustomTailInput(e.target.value)}
                                required
                                placeholder="Kuyruk No manuel yazınız (Örn: OR-2045)..."
                                className="w-full px-2.5 py-1.5 bg-amber-50 border-2 border-amber-400 rounded-xl text-xs font-black text-amber-950 font-mono focus:outline-none"
                              />
                            </div>
                          )}
                        </div>
                      )}
                    </div>

                    {/* Teslim Alan & Kabul Yapan & Depo Yeri */}
                    <div className="grid grid-cols-3 gap-2.5">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-700 uppercase font-mono mb-1">Teslim Alan</label>
                        <PersonnelAutocomplete 
                          value={movementTeslimAlan}
                          onChange={setMovementTeslimAlan}
                          placeholder="Ad Soyad seç/yaz"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-700 uppercase font-mono mb-1">Kabul Yapan</label>
                        <PersonnelAutocomplete 
                          value={movementKabulYapan}
                          onChange={setMovementKabulYapan}
                          placeholder="Yetkili seç/yaz"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-700 uppercase font-mono mb-1">Depo Yeri</label>
                        <input 
                          type="text" 
                          value={movementDepoYeri}
                          onChange={(e) => setMovementDepoYeri(e.target.value)}
                          placeholder="Raf / Bölge" 
                          className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold" 
                        />
                      </div>
                    </div>

                    {/* AÇIKLAMA KISMI (Malzeme Girişi & Hareket Notları) */}
                    <div>
                      <label className="block text-[10px] font-bold text-slate-700 uppercase font-mono mb-1 flex items-center justify-between">
                        <span>AÇIKLAMA {movementDirection === 'giris' && <span className="text-emerald-700 font-semibold">(Kabul / Giriş Notu)</span>}</span>
                      </label>
                      <textarea 
                        rows={2}
                        value={movementNotes}
                        onChange={(e) => setMovementNotes(e.target.value)}
                        placeholder={movementDirection === 'giris' ? "Malzeme kutusunda hasar yok. Sertifika ve kabul formu onaylandı. Sayım girişi tamam..." : "Hareket açıklaması / not..."}
                        className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-sans focus:outline-none focus:border-emerald-600 resize-none"
                      />
                    </div>
                  </>
                )}

                <div className="pt-3 border-t border-slate-200 flex justify-end gap-2 flex-wrap items-center">
                  {(movementDirection === 'giris' || movementDirection === 'cikis') && (
                    <button
                      type="button"
                      onClick={() => {
                        setBulkExcelRows([]);
                        setIsBulkExcelModalOpen(true);
                      }}
                      className="mr-auto flex items-center gap-1.5 px-3 py-2 bg-amber-500 hover:bg-amber-400 text-amber-950 rounded-xl text-xs font-black uppercase tracking-wider transition shadow-sm cursor-pointer"
                    >
                      <FileSpreadsheet className="w-3.5 h-3.5" />
                      <span>Excel ile Toplu Giriş/Çıkış Yap</span>
                    </button>
                  )}
                  <button 
                    type="button" 
                    onClick={() => setIsMovementModalOpen(false)} 
                    className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-xl cursor-pointer"
                  >
                    İptal
                  </button>
                  <button 
                    type="submit" 
                    disabled={Boolean(isTransferOverStock || isCikisOverStock)}
                    className={`px-5 py-2 font-black uppercase rounded-xl shadow-xs text-white cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed ${
                      movementDirection === 'cikis' ? 'bg-[#a0153e] hover:bg-[#851233]' : (movementDirection === 'giris' ? 'bg-[#006a4e] hover:bg-[#00523c]' : 'bg-indigo-700 hover:bg-indigo-800')
                    }`}
                  >
                    {movementDirection === 'cikis' ? 'Çıkışı Onayla' : (movementDirection === 'giris' ? 'Girişi Onayla' : 'Transferi Onayla')}
                  </button>
                </div>

              </form>
            </div>
          </div>
        );
      })()}

            {/* EXCEL YÜKLEME TÜRÜ SEÇİM MODALI */}
      {isExcelTypeModalOpen && (
        <div className="fixed inset-0 z-[10500] bg-black/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full border border-slate-200 overflow-hidden text-slate-800">
            <div className="bg-[#0b3d1d] text-white px-5 py-3.5 flex items-center justify-between">
              <h3 className="text-xs font-black uppercase tracking-wider flex items-center gap-2">
                <span>📊 Excel Yükleme Adımı</span>
              </h3>
              <button type="button" onClick={() => setIsExcelTypeModalOpen(false)} className="text-white hover:text-slate-300 font-bold cursor-pointer">✕</button>
            </div>
            <div className="p-5 space-y-4">
              <p className="text-xs font-bold text-slate-700">
                Lütfen yükleyeceğiniz Excel dosyasının hangi depoya ait olduğunu seçiniz:
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <button
                  type="button"
                  onClick={() => {
                    setSelectedExcelUploadCategory('sarf');
                    setIsExcelTypeModalOpen(false);
                    setTimeout(() => fileInputRef.current?.click(), 100);
                  }}
                  className="p-4 rounded-xl border-2 border-slate-200 hover:border-emerald-600 hover:bg-emerald-50 text-left transition cursor-pointer group"
                >
                  <div className="text-2xl mb-1">📦</div>
                  <div className="font-black text-xs text-slate-900 group-hover:text-emerald-800">Sarf Parça Deposu</div>
                  <div className="text-[10px] text-slate-500 font-medium">Sarf & Mekanik Parçalar</div>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setSelectedExcelUploadCategory('kimyasal');
                    setIsExcelTypeModalOpen(false);
                    setTimeout(() => fileInputRef.current?.click(), 100);
                  }}
                  className="p-4 rounded-xl border-2 border-slate-200 hover:border-amber-500 hover:bg-amber-50 text-left transition cursor-pointer group"
                >
                  <div className="text-2xl mb-1">🧪</div>
                  <div className="font-black text-xs text-slate-900 group-hover:text-amber-800">Kimyasal Depo</div>
                  <div className="text-[10px] text-slate-500 font-medium">Sıvı & Kimyasal Parçalar</div>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setSelectedExcelUploadCategory('yasam_destek');
                    setIsExcelTypeModalOpen(false);
                    setTimeout(() => fileInputRef.current?.click(), 100);
                  }}
                  className="p-4 rounded-xl border-2 border-slate-200 hover:border-purple-600 hover:bg-purple-50 text-left transition cursor-pointer group"
                >
                  <div className="text-2xl mb-1">🦺</div>
                  <div className="font-black text-xs text-slate-900 group-hover:text-purple-800">Yaşam Destek</div>
                  <div className="text-[10px] text-slate-500 font-medium">Can Yeleği, Spare Air, Helmet Kit & Seri Nolar</div>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* YENİ MALZEME KAYDETME & DÜZENLEME MODALI */}
      {isNewItemModalOpen && (
        <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full border border-slate-200 overflow-hidden">
            <div className="bg-[#0b3d1d] text-white px-5 py-3.5 flex items-center justify-between border-b border-emerald-500/30">
              <h3 className="text-xs font-black uppercase tracking-wider">
                {editingItemIdx >= 0 ? '✏️ Malzeme Kaydını Düzenle' : '📦 Yeni Malzeme Ekle'}
              </h3>
              <button type="button" onClick={() => setIsNewItemModalOpen(false)} className="text-white hover:text-slate-300 font-bold cursor-pointer">✕</button>
            </div>

            <form onSubmit={handleSaveItem} className="p-5 space-y-3 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">Hava Aracı / Birim</label>
                  <select 
                    value={newItemUnit}
                    onChange={(e) => setNewItemUnit(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                  >
                    <option value="at802">AT-802F</option>
                    <option value="bell429">BELL 429</option>
                    <option value="t70">T-70</option>
                    <option value="c650">C-650</option>
                    <option value="b360">B-360 King Air</option>
                    <option value="hangar">HANGAR GENEL</option>
                  </select>
                </div>
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">Kategori</label>
                  <select 
                    value={newItemCategory}
                    onChange={(e) => setNewItemCategory(e.target.value as any)}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                  >
                    <option value="sarf">Sarf & Mekanik Parça</option>
                    <option value="kimyasal">Kimyasal / Sıvı</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">Description (Malzeme Adı) *</label>
                <input 
                  type="text" 
                  value={newItemDesc}
                  onChange={(e) => setNewItemDesc(e.target.value)}
                  required 
                  placeholder="Malzeme adı ve tanımı..." 
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-medium" 
                />
              </div>

              {!isItemInCatalog && (newItemDesc.trim() || newItemPn.trim()) && (
                <div className="p-2.5 bg-amber-50 border-2 border-amber-400/80 rounded-xl flex items-center justify-between text-xs font-bold text-amber-950 shadow-xs animate-pulse font-sans">
                  <div className="flex items-center gap-2">
                    <span className="text-base">⚠️</span>
                    <span>Sistem Kataloğunda Bulunmayan Yeni Ürün Kaydı (Katalog Dışı)</span>
                  </div>
                  <span className="px-2 py-0.5 bg-amber-200 text-amber-900 rounded font-mono text-[10px] uppercase font-bold">Listede Yok</span>
                </div>
              )}

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">Part Number (P/N)</label>
                  <input 
                    type="text" 
                    value={newItemPn}
                    onChange={(e) => setNewItemPn(e.target.value)}
                    placeholder="Parça numarası..." 
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-mono font-bold" 
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">Lokasyon No</label>
                  <input 
                    type="text" 
                    value={newItemLokasyon}
                    onChange={(e) => setNewItemLokasyon(e.target.value)}
                    placeholder="Örn: KARAİN RAF B-2" 
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl" 
                  />
                </div>
              </div>

              {/* SERİ NO (S/N) EKLEME - ÇOKLU S/N DESTEĞİ VE + BUTONU */}
              <div className="space-y-1.5 p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="block font-bold text-slate-800 uppercase font-mono text-xs flex items-center justify-between">
                  <span>Seri Numaraları (S/N Ekleme)</span>
                  <span className="text-[10px] text-slate-500 font-sans font-normal">
                    Birden fazla S/N varsa yazıp <strong>+ Ekle</strong> butonuna basınız
                  </span>
                </label>
                
                <div className="flex items-center gap-2">
                  <input 
                    type="text" 
                    value={singleSnInput}
                    onChange={(e) => setSingleSnInput(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        if (singleSnInput.trim()) {
                          const val = singleSnInput.trim();
                          if (!snList.includes(val)) {
                            const updated = [...snList, val];
                            setSnList(updated);
                            setNewItemSnNotes(updated.join(', '));
                            setNewItemGelen(updated.length);
                          }
                          setSingleSnInput('');
                        }
                      }
                    }}
                    placeholder="Seri No giriniz (Örn: SN-9011)..." 
                    className="flex-1 px-3 py-1.5 bg-white border border-slate-300 rounded-lg font-mono text-xs focus:outline-none focus:border-emerald-500 font-bold" 
                  />
                  <button
                    type="button"
                    onClick={() => {
                      if (singleSnInput.trim()) {
                        const val = singleSnInput.trim();
                        if (!snList.includes(val)) {
                          const updated = [...snList, val];
                          setSnList(updated);
                          setNewItemSnNotes(updated.join(', '));
                          setNewItemGelen(updated.length);
                        }
                        setSingleSnInput('');
                      }
                    }}
                    className="px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white font-bold rounded-lg text-xs transition active:scale-95 cursor-pointer flex items-center gap-1 shadow-xs"
                  >
                    <span className="text-sm font-black">+</span>
                    <span>S/N Ekle</span>
                  </button>
                </div>

                {/* S/N Badge Tags */}
                {snList.length > 0 ? (
                  <div className="flex flex-wrap gap-1.5 pt-1.5 max-h-24 overflow-y-auto">
                    {snList.map((s, sIdx) => (
                      <span key={sIdx} className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-slate-900 text-emerald-400 text-xs font-mono font-bold border border-slate-700 shadow-xs">
                        <span>{s}</span>
                        <button
                          type="button"
                          onClick={() => {
                            const updated = snList.filter((_, i) => i !== sIdx);
                            setSnList(updated);
                            setNewItemSnNotes(updated.join(', '));
                            if (updated.length > 0) setNewItemGelen(updated.length);
                          }}
                          className="text-rose-400 hover:text-rose-200 font-black text-xs cursor-pointer ml-1"
                        >
                          ✕
                        </button>
                      </span>
                    ))}
                  </div>
                ) : (
                  <input 
                    type="text" 
                    value={newItemSnNotes}
                    onChange={(e) => setNewItemSnNotes(e.target.value)}
                    placeholder="Veya elle toplu S/N / Açıklama yazın (Örn: 15003, 14986)..." 
                    className="w-full px-2.5 py-1 bg-white border border-slate-300 rounded-lg text-xs font-mono text-slate-700" 
                  />
                )}
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 uppercase font-mono mb-1">Gelen Miktar</label>
                  <input 
                    type="number" 
                    value={newItemGelen}
                    onChange={(e) => setNewItemGelen(Math.max(1, parseInt(e.target.value) || 1))}
                    min={1} 
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-mono font-bold" 
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 uppercase font-mono mb-1">İlk Depo</label>
                  <select 
                    value={newItemInitialRegion}
                    onChange={(e) => setNewItemInitialRegion(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold"
                  >
                    <option value="ankara">Ankara (Merkez)</option>
                    <option value="karain">Karain</option>
                    <option value="canakkale">Çanakkale</option>
                    <option value="milas">Milas</option>
                    <option value="bursa">Bursa</option>
                  </select>
                </div>
              </div>

              {/* RAF ÖMRÜ ALANLARI */}
              <div className="grid grid-cols-2 gap-3 p-3 bg-emerald-50/70 rounded-xl border border-emerald-200">
                <div>
                  <label className="block text-[10px] font-black text-emerald-950 uppercase font-mono mb-1">
                    Raf Ömrü Var mı?
                  </label>
                  <select 
                    value={newItemHasShelfLife}
                    onChange={(e) => setNewItemHasShelfLife(e.target.value as any)}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-bold text-xs"
                  >
                    <option value="HAYIR">HAYIR</option>
                    <option value="EVET">EVET</option>
                  </select>
                </div>
                 <div>
                  <label className="block text-[10px] font-black text-emerald-950 uppercase font-mono mb-1">
                    Raf Ömrü Bitiş Tarihi (Yeni satır ile Birden Fazla Parti)
                  </label>
                  <textarea 
                    rows={2}
                    disabled={newItemHasShelfLife !== 'EVET'}
                    value={newItemShelfLifeDate}
                    onChange={(e) => setNewItemShelfLifeDate(e.target.value)}
                    placeholder="Örn: 30.12.2026&#10;Birden fazla parti varsa her birini yeni satıra yazınız."
                    className={`w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-bold font-mono whitespace-pre ${
                      newItemHasShelfLife !== 'EVET' ? 'bg-slate-100 text-slate-400 cursor-not-allowed font-sans' : 'text-slate-800'
                    }`}
                  />
                  {newItemHasShelfLife === 'EVET' && (
                    <div className="mt-1 flex items-center gap-1.5">
                      <input 
                        type="date"
                        onChange={(e) => {
                          if (e.target.value) {
                            const [y, m, d] = e.target.value.split('-');
                            const formatted = `${d}.${m}.${y}`;
                            setNewItemShelfLifeDate(prev => {
                              const trimmed = prev.trim();
                              if (trimmed === '-' || !trimmed) return formatted;
                              return `${trimmed}\n${formatted}`;
                            });
                          }
                        }}
                        className="px-1.5 py-0.5 border border-slate-300 rounded text-[10px] font-bold cursor-pointer bg-white text-slate-800"
                      />
                      <span className="text-[9px] text-emerald-800 font-bold select-none">Tarih Seç ve Parti Ekle</span>
                    </div>
                  )}
                </div>
              </div>

              <div className="pt-3 border-t border-slate-200 flex justify-end gap-2">
                <button type="button" onClick={() => setIsNewItemModalOpen(false)} className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-xl cursor-pointer">İptal</button>
                <button type="submit" className="px-5 py-2 bg-[#0b3d1d] hover:bg-[#072a14] text-white font-black uppercase rounded-xl shadow-xs cursor-pointer">Kaydet</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* EXCEL ŞABLONU İLE TOPLU GİRİŞ / ÇIKIŞ ALT MODALI */}
      {isBulkExcelModalOpen && (
        <div className="fixed inset-0 z-[10010] bg-black/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl max-w-4xl w-full border border-slate-200 overflow-hidden flex flex-col max-h-[95vh] animate-scale-up">
            
            {/* Header */}
            <div className="bg-gradient-to-r from-amber-600 to-amber-700 text-white px-6 py-4 flex items-center justify-between border-b border-amber-800">
              <div className="flex items-center gap-2.5">
                <FileSpreadsheet className="w-6 h-6 text-amber-100" />
                <div>
                  <h3 className="text-sm font-black uppercase tracking-wider">Excel Şablonu ile Toplu Malzeme {movementDirection === 'giris' ? 'Girişi' : 'Çıkışı'}</h3>
                  <p className="text-[10px] text-amber-100/90 font-medium">Parça numaraları (PN) ve seri numaraları (SN) otomatik taranır ve eşleştirilir</p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setIsBulkExcelModalOpen(false)} 
                className="text-white hover:text-amber-200 font-extrabold text-sm cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="p-6 overflow-y-auto space-y-4 flex-1 text-xs">
              
              {/* 1. Step: Download template or upload */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                
                {/* Download Template Box */}
                <div className="p-4 bg-amber-50/50 rounded-2xl border border-amber-200 flex flex-col justify-between">
                  <div>
                    <h4 className="font-extrabold text-amber-950 uppercase tracking-wider text-[11px] mb-1">1. Format Şablonunu İndirin</h4>
                    <p className="text-[11px] text-slate-600 leading-relaxed mb-3">
                      Sisteme toplu giriş/çıkış verisi yüklemek için aşağıdaki format şablonunu indirin ve verilerinizi bu formata göre hazırlayın. Şablonda <strong>Description (Açıklama)</strong> sütunu yoktur, ürün tanımları otomatik eşleşecektir.
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={handleDownloadBulkTemplate}
                    className="flex items-center justify-center gap-2 w-full py-2 bg-amber-600 hover:bg-amber-500 text-white text-xs font-black uppercase tracking-wider rounded-xl transition shadow-xs cursor-pointer active:scale-95"
                  >
                    <Download className="w-4 h-4" />
                    <span>Örnek Şablon İndir (.xlsx)</span>
                  </button>
                </div>

                {/* Upload Excel Box */}
                <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200 flex flex-col justify-between">
                  <h4 className="font-extrabold text-slate-800 uppercase tracking-wider text-[11px] mb-1">2. Hazırladığınız Şablonu Yükleyin</h4>
                  <div className="mt-2 border-2 border-dashed border-slate-300 hover:border-amber-500 hover:bg-amber-50/10 rounded-2xl p-6 text-center cursor-pointer transition relative">
                    <input
                      type="file"
                      accept=".xlsx, .xls"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) handleBulkExcelParse(file);
                      }}
                      className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                    />
                    <UploadCloud className="w-8 h-8 text-slate-400 mx-auto mb-2" />
                    <span className="text-xs font-bold text-slate-700 block">Dosya Seçin veya Sürükleyin</span>
                    <span className="text-[10px] text-slate-400 block mt-1">Sadece .xlsx, .xls şablon dosyaları</span>
                  </div>
                </div>

              </div>

              {/* 2. Step: Parsed Items Table with Autocomplete matching */}
              {bulkExcelRows.length > 0 && (
                <div className="space-y-2 pt-2">
                  <div className="flex items-center justify-between">
                    <h4 className="font-black text-slate-800 uppercase tracking-wider text-[11px]">
                      Yüklenen Ürün Verileri ve Eşleşme Statüleri ({bulkExcelRows.length} Kayıt)
                    </h4>
                    <span className="text-[10px] text-slate-500 font-medium">
                      Eşleşmeyen veya yanlış eşleşen kayıtları yanlarındaki arama kutularından düzeltebilirsiniz.
                    </span>
                  </div>

                  <div className="border border-slate-200 rounded-2xl overflow-hidden shadow-xs">
                    <div className="max-h-60 overflow-y-auto">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead className="bg-slate-100 sticky top-0 z-10 text-slate-700 font-bold border-b border-slate-200 text-[10px] uppercase font-mono">
                          <tr>
                            <th className="p-2.5 border-r border-slate-200 text-center w-12">No</th>
                            <th className="p-2.5 border-r border-slate-200 w-36">Part Number (PN)</th>
                            <th className="p-2.5 border-r border-slate-200 w-28 text-center">Serial (SN)</th>
                            <th className="p-2.5 border-r border-slate-200 text-center w-20">Adet</th>
                            <th className="p-2.5 border-r border-slate-200 w-24 text-center">Eşleşme Durumu</th>
                            <th className="p-2.5">Sistem Ürün Tanımı / Manuel Düzeltme Arama</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200 font-medium">
                          {bulkExcelRows.map((row, rIdx) => {
                            const isMatched = row.status === 'Eşleşti';
                            const isDropdownOpen = bulkDropdownOpenRow === row.id;
                            
                            return (
                              <tr key={row.id} className={`hover:bg-slate-50 ${!isMatched ? 'bg-rose-50/40' : ''}`}>
                                <td className="p-2 border-r border-slate-200 text-center font-mono font-bold text-slate-500">{row.id}</td>
                                <td className="p-2 border-r border-slate-200 font-mono font-bold text-slate-800">{row.pn}</td>
                                <td className="p-2 border-r border-slate-200 font-mono text-center text-slate-600">{row.sn}</td>
                                <td className="p-2 border-r border-slate-200 text-center font-mono font-black text-slate-900">{row.qty}</td>
                                <td className="p-2 border-r border-slate-200 text-center">
                                  <span className={`inline-block px-2 py-0.5 rounded-full text-[9px] font-black uppercase ${
                                    isMatched 
                                      ? 'bg-emerald-100 text-emerald-900 border border-emerald-300' 
                                      : 'bg-rose-100 text-rose-900 border border-rose-300'
                                  }`}>
                                    {row.status}
                                  </span>
                                </td>
                                <td className="p-2 relative">
                                  <div className="flex items-center gap-2">
                                    <div className="relative flex-1">
                                      <input
                                        type="text"
                                        value={bulkSearchQueries[row.id] || ''}
                                        onChange={(e) => {
                                          const val = e.target.value;
                                          setBulkSearchQueries(prev => ({ ...prev, [row.id]: val }));
                                          setBulkDropdownOpenRow(row.id);
                                        }}
                                        onFocus={() => setBulkDropdownOpenRow(row.id)}
                                        placeholder="Ürün adı veya P/N ile sistemden manuel eşleştir..."
                                        className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-bold text-slate-800 focus:outline-none focus:border-amber-600"
                                      />
                                      {bulkSearchQueries[row.id] && (
                                        <button 
                                          type="button"
                                          onClick={() => {
                                            setBulkSearchQueries(prev => ({ ...prev, [row.id]: '' }));
                                            setBulkExcelRows(prevRows => prevRows.map(r => r.id === row.id ? { ...r, matchedItem: null, status: 'Eşleşmedi' } : r));
                                          }}
                                          className="absolute right-2 top-2 text-slate-400 hover:text-slate-600"
                                        >
                                          ✕
                                        </button>
                                      )}
                                    </div>
                                    
                                    {isMatched && (
                                      <span className="text-[10px] text-slate-500 shrink-0 font-bold bg-slate-100 px-2 py-1 rounded-md border border-slate-200 max-w-[200px] truncate" title={row.matchedItem?.description}>
                                        {row.matchedItem?.description}
                                      </span>
                                    )}
                                  </div>

                                  {/* Row dropdown results */}
                                  {isDropdownOpen && (
                                    <div className="absolute left-2 right-2 top-full z-[10020] mt-1 bg-white border-2 border-amber-600 rounded-xl shadow-2xl max-h-40 overflow-y-auto divide-y divide-slate-100">
                                      {computedInventory.filter(i => {
                                        const query = (bulkSearchQueries[row.id] || '').toLowerCase().trim();
                                        const unitKey = currentUnit !== 'all' ? currentUnit : 'at802';
                                        const categoryKey = currentDepoType === 'kimyasal' ? 'kimyasal' : 'sarf';
                                        if (i.unit !== unitKey || i.category !== categoryKey) return false;
                                        if (!query) return true;

                                        const desc = (i.description || i.name || '').toLowerCase();
                                        const pn = (i.partNumber || i.pn || '').toLowerCase();
                                        return desc.includes(query) || pn.includes(query);
                                      }).slice(0, 10).map((m, mIdx) => (
                                        <div
                                          key={mIdx}
                                          onClick={() => handleBulkRowMatchSelect(row.id, m)}
                                          className="p-2 hover:bg-amber-50 cursor-pointer text-xs flex items-center justify-between gap-1.5"
                                        >
                                          <div>
                                            <div className="font-bold text-slate-900">{m.description || m.name}</div>
                                            <div className="text-[10px] font-mono text-slate-500">P/N: <strong className="text-amber-800">{m.partNumber || m.pn}</strong></div>
                                          </div>
                                          <span className="text-[10px] font-mono font-bold bg-amber-100 text-amber-950 px-2 py-0.5 rounded-full shrink-0">
                                            Stok: {m.ankaraMevcut || 0}
                                          </span>
                                        </div>
                                      ))}
                                    </div>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

            </div>

            {/* Footer */}
            <div className="bg-slate-50 px-6 py-4 border-t border-slate-200 flex items-center justify-between shrink-0">
              <div className="text-[11px] text-slate-500 font-medium">
                {bulkExcelRows.length > 0 && (
                  <span>
                    Toplam: <strong>{bulkExcelRows.length}</strong> kalem | Eşleşen: <strong className="text-emerald-700">{bulkExcelRows.filter(r => r.status === 'Eşleşti').length}</strong> | Kalan: <strong className="text-rose-600">{bulkExcelRows.filter(r => r.status !== 'Eşleşti').length}</strong>
                  </span>
                )}
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsBulkExcelModalOpen(false)}
                  className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-xl cursor-pointer hover:bg-slate-200"
                >
                  Geri Dön
                </button>
                <button
                  type="button"
                  disabled={bulkExcelRows.length === 0}
                  onClick={handleSaveBulkExcelTransactions}
                  className="px-5 py-2 bg-amber-600 hover:bg-amber-500 text-white font-black uppercase rounded-xl shadow-md transition disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  İşlemi Onayla ve Toplu Kaydet
                </button>
              </div>
            </div>

          </div>
        </div>
      )}

      {/* KİT OLUŞTURMA MODALI */}
      {isKitModalOpen && (
        <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="bg-amber-600 text-white px-5 py-3.5 flex items-center justify-between border-b border-amber-700">
              <div className="flex items-center gap-2">
                <Package className="w-5 h-5 text-amber-100" />
                <div>
                  <h3 className="text-xs font-black uppercase tracking-wider">Depo Kit Oluşturma Modülü</h3>
                  <p className="text-[10px] text-amber-100 font-medium">Reçeteli parça düşüşü ve otomatik stok ataması</p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setIsKitModalOpen(false)} 
                className="text-white hover:text-amber-200 font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleCreateKitSubmit} className="p-5 overflow-y-auto space-y-4 text-xs">
              
              {/* Kit Adı ve Toplam Üretim Sayısı */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="sm:col-span-2">
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    KİT ADI *
                  </label>
                  <input 
                    type="text" 
                    value={kitName}
                    onChange={(e) => setKitName(e.target.value)}
                    required
                    placeholder="Örn: AT-802 100 SAAT BAKIM KİTİ" 
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-none focus:border-amber-600 transition uppercase" 
                  />
                </div>

                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    TOPLAM ÜRETİLECEK KİT *
                  </label>
                  <input 
                    type="number" 
                    min={1} 
                    value={kitCount}
                    onChange={(e) => setKitCount(Math.max(1, parseInt(e.target.value) || 1))}
                    required 
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold font-mono text-center text-amber-950 focus:outline-none focus:border-amber-600" 
                  />
                </div>
              </div>

              {/* PN Listesinden Parça Seçimi ve Kit İçi Adet Ekleme */}
              <div className="bg-slate-50 p-3.5 rounded-2xl border border-slate-200 space-y-3">
                <h4 className="font-bold text-slate-800 text-[11px] uppercase tracking-wider flex items-center gap-1.5">
                  <span>➕ Kit İçeriğine Parça Ekle (PN Listesinden)</span>
                </h4>

                <div className="grid grid-cols-1 sm:grid-cols-12 gap-2.5">
                  <div className="sm:col-span-8 relative">
                    <label className="block text-[10px] font-bold text-slate-600 uppercase font-mono mb-1">
                      P/N VEYA MALZEME ADI
                    </label>
                    <div className="relative">
                      <input 
                        type="text"
                        value={kitSearchQuery}
                        onChange={(e) => {
                          setKitSearchQuery(e.target.value);
                          setIsKitDropdownOpen(true);
                        }}
                        onFocus={() => setIsKitDropdownOpen(true)}
                        placeholder="Parça ara (P/N veya isim yazın)..."
                        className="w-full pl-8 pr-3 py-2 bg-white border border-slate-300 rounded-xl font-bold text-slate-800 text-xs focus:outline-none focus:border-amber-600"
                      />
                      <Search className="w-4 h-4 absolute left-2.5 top-2.5 text-slate-400" />
                    </div>

                    {isKitDropdownOpen && (
                      <div className="absolute top-full left-0 right-0 z-50 mt-1 bg-white border-2 border-amber-600 rounded-xl shadow-2xl max-h-48 overflow-y-auto divide-y divide-slate-100">
                        {computedInventory.filter(i => {
                          if (!kitSearchQuery.trim()) return true;
                          const q = kitSearchQuery.toLowerCase().trim();
                          const d = (i.description || i.name || '').toLowerCase();
                          const p = (i.partNumber || i.pn || '').toLowerCase();
                          return d.includes(q) || p.includes(q);
                        }).slice(0, 10).map((m, idx) => (
                          <div 
                            key={idx} 
                            onClick={() => {
                              setSelectedKitItem(m);
                              setKitSearchQuery(`${m.description || m.name} (P/N: ${m.partNumber || m.pn || '-'})`);
                              setIsKitDropdownOpen(false);
                            }}
                            className="p-2 hover:bg-amber-50 cursor-pointer text-xs flex items-center justify-between gap-2"
                          >
                            <div>
                              <div className="font-bold text-slate-900">{m.description || m.name}</div>
                              <div className="text-[10px] font-mono text-slate-500">
                                P/N: <strong className="text-amber-800">{m.partNumber || m.pn || '-'}</strong>
                              </div>
                            </div>
                            <span className="text-[10px] font-mono font-bold bg-amber-100 text-amber-950 px-2 py-0.5 rounded-full shrink-0">
                              Ankara: {m.ankaraMevcut || 0} Adet
                            </span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  <div className="sm:col-span-2">
                    <label className="block text-[10px] font-bold text-slate-600 uppercase font-mono mb-1">
                      KİT İÇİ ADET
                    </label>
                    <input 
                      type="number"
                      min={1}
                      value={kitItemQtyInput}
                      onChange={(e) => setKitItemQtyInput(Math.max(1, parseInt(e.target.value) || 1))}
                      className="w-full px-2.5 py-2 bg-white border border-slate-300 rounded-xl font-bold font-mono text-center text-xs"
                    />
                  </div>

                  <div className="sm:col-span-2 flex items-end">
                    <button
                      type="button"
                      onClick={handleAddKitItem}
                      className="w-full py-2 bg-amber-600 hover:bg-amber-500 text-white font-bold rounded-xl text-xs transition cursor-pointer shadow-xs"
                    >
                      Ekle
                    </button>
                  </div>
                </div>

                {selectedKitItem && (
                  <div className="text-[11px] text-amber-900 bg-amber-50 p-2 rounded-lg border border-amber-200 flex justify-between items-center">
                    <span>Seçilen: <strong>{selectedKitItem.description || selectedKitItem.name}</strong></span>
                    <span>Ankara Depo Mevcudu: <strong>{selectedKitItem.ankaraMevcut || 0} Adet</strong></span>
                  </div>
                )}
              </div>

              {/* Alt Alta Eklenen Parçalar Tablosu ve Formül Hesaplaması */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <label className="font-bold text-slate-700 uppercase font-mono text-[11px]">
                    KİT İÇERİK LİSTESİ ({kitItems.length} Kalem)
                  </label>
                  <span className="text-[10px] text-slate-500">
                    Formül: (Kit İçi Adet) × (Toplam Üretilecek Kit Sayısı: {kitCount})
                  </span>
                </div>

                {kitItems.length === 0 ? (
                  <div className="p-6 border-2 border-dashed border-slate-200 rounded-2xl text-center text-slate-400">
                    Henüz parça eklenmedi. Yukarıdaki arama kutusundan parçaları seçip adet belirterek alt alta ekleyiniz.
                  </div>
                ) : (
                  <div className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead className="bg-slate-100 text-slate-700 text-[10px] font-mono uppercase font-bold border-b border-slate-200">
                        <tr>
                          <th className="p-2 border-r border-slate-200">PARÇA (P/N & MALZEME ADI)</th>
                          <th className="p-2 text-center border-r border-slate-200 w-24">1 KİT İÇİ</th>
                          <th className="p-2 text-center border-r border-slate-200 w-28 bg-amber-50 text-amber-950 font-black">TOPLAM DÜŞECEK</th>
                          <th className="p-2 text-center border-r border-slate-200 w-24">ANKARA MEVCUT</th>
                          <th className="p-2 text-center border-r border-slate-200 w-24">KALAN STOK</th>
                          <th className="p-2 text-center w-12">SİL</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 font-medium">
                        {kitItems.map((ki, idx) => {
                          const totalDeduct = ki.qtyPerKit * kitCount;
                          const remaining = ki.ankaraMevcut - totalDeduct;
                          const isShortage = remaining < 0;

                          return (
                            <tr key={idx} className={isShortage ? 'bg-rose-50' : 'hover:bg-slate-50'}>
                              <td className="p-2 border-r border-slate-200">
                                <div className="font-bold text-slate-900">{ki.description}</div>
                                <div className="text-[10px] font-mono text-slate-500">P/N: {ki.partNumber}</div>
                              </td>
                              <td className="p-2 text-center font-mono font-bold text-slate-700 border-r border-slate-200">
                                {ki.qtyPerKit}
                              </td>
                              <td className="p-2 text-center font-mono font-black text-amber-900 bg-amber-50/50 border-r border-slate-200 text-xs">
                                {totalDeduct}
                              </td>
                              <td className="p-2 text-center font-mono font-bold text-slate-800 border-r border-slate-200">
                                {ki.ankaraMevcut}
                              </td>
                              <td className={`p-2 text-center font-mono font-bold border-r border-slate-200 ${isShortage ? 'text-rose-600 font-black' : 'text-emerald-700'}`}>
                                {isShortage ? `❌ ${remaining} (Yetersiz)` : remaining}
                              </td>
                              <td className="p-2 text-center">
                                <button
                                  type="button"
                                  onClick={() => handleRemoveKitItem(idx)}
                                  className="text-rose-600 hover:text-rose-800 font-bold p-1 cursor-pointer"
                                  title="Parçayı Listeden Kaldır"
                                >
                                  ✕
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

              {/* Bilgilendirme Kartı */}
              <div className="p-3 bg-amber-50/60 border border-amber-200 rounded-xl text-[11px] text-amber-950 space-y-1">
                <div className="font-bold flex items-center gap-1">
                  <span>💡 Otomatik Stok Ataması Kuralı:</span>
                </div>
                <div>• Oluşturulan Kit, girilen <strong>"{kitName || 'Kit Adı'}"</strong> ile tablonun en altına yeni bir ürün kalemi olarak eklenecektir.</div>
                <div>• Ankara Depo Mevcuduna üretilen toplam kit adedi (<strong>{kitCount}</strong>) yazılacak; diğer tüm depolar (Karain, Çanakkale vb.) otomatik olarak <strong>0 (Sıfır)</strong> atanacaktır.</div>
                <div>• Yapılan tüm değişiklikler ve parça düşüşleri bağlı bulunan Excel Online dosyasına anlık ve otomatik olarak yansıtılacaktır.</div>
              </div>

              {/* Butonlar */}
              <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setIsKitModalOpen(false)}
                  className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-xl cursor-pointer font-sans"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  disabled={kitItems.length === 0 || kitItems.some(ki => (ki.qtyPerKit * kitCount) > ki.ankaraMevcut)}
                  className="px-5 py-2 bg-amber-600 hover:bg-amber-500 disabled:opacity-50 disabled:cursor-not-allowed font-black uppercase text-white rounded-xl shadow-xs cursor-pointer flex items-center gap-2 font-sans"
                >
                  <Package className="w-4 h-4" />
                  <span>Kiti Oluştur ve Stoğu Güncelle</span>
                </button>
              </div>

            </form>
          </div>
        </div>
      )}

      {/* ÜRÜNE ÖZEL DETAY GÖR POP-UP MODALI */}
      {isDetailModalOpen && detailModalItem && (
        <div className="fixed inset-0 z-[10100] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-3xl w-full border border-slate-200 overflow-hidden flex flex-col max-h-[85vh]">
            
            {/* Header */}
            <div className="bg-slate-900 text-white px-5 py-4 flex items-center justify-between border-b border-slate-800">
              <div className="flex items-center gap-3">
                <div className="p-1.5 bg-emerald-500/20 text-emerald-400 rounded-lg border border-emerald-500/30 animate-pulse">
                  <RotateCcw className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs font-black uppercase tracking-wider text-slate-100">MALZEME HAREKET VE DAĞILIM DETAYLARI</h3>
                  <p className="text-[11px] text-slate-300 font-bold font-sans mt-0.5">
                    {detailModalItem.description || detailModalItem.name} 
                    <span className="text-emerald-400 font-mono ml-2">P/N: {detailModalItem.partNumber || detailModalItem.pn || '-'}</span>
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => {
                  setIsDetailModalOpen(false);
                  setDetailModalItem(null);
                }} 
                className="text-slate-400 hover:text-white font-extrabold text-sm transition p-1 cursor-pointer"
              >
                ✕
              </button>
            </div>

            {/* Content */}
            <div className="p-5 overflow-y-auto space-y-4">
              
              {/* Stats Summary Cards */}
              <div className="grid grid-cols-2 sm:grid-cols-2 gap-3">
                {(() => {
                  const list = transactions.filter(t => matchTxToItem(t, detailModalItem));
                  const isGelen = detailModalColumnType === 'GELEN';
                  
                  // Helper to get stats for a specific type/region
                  const getStatsForType = (typeFilter: string | null) => {
                    let filtered = list;
                    if (typeFilter === 'GELEN' || typeFilter === 'ANKARA GİREN') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('GİRİŞ') || u.includes('GELEN') || u.includes('GİREN') || (u.includes('ANKARA') && u.includes('FAZLA TESPİT'));
                        });
                    } else if (typeFilter === 'KARAİN TRANSFER') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('KARAİN') && u.includes('TRANSFER');
                        });
                    } else if (typeFilter === 'KARAİN ÇIKAN') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('KARAİN') && (u.includes('ÇIKAN') || u.includes('ÇIKIŞ'));
                        });
                    } else if (typeFilter === 'ÇANAKKALE TRANSFER') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('ÇANAKKALE') && u.includes('TRANSFER');
                        });
                    } else if (typeFilter === 'ÇANAKKALE ÇIKAN') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('ÇANAKKALE') && (u.includes('ÇIKAN') || u.includes('ÇIKIŞ'));
                        });
                    } else if (typeFilter === 'MİLAS TRANSFER') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('MİLAS') && u.includes('TRANSFER');
                        });
                    } else if (typeFilter === 'MİLAS ÇIKAN') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('MİLAS') && (u.includes('ÇIKAN') || u.includes('ÇIKIŞ'));
                        });
                    } else if (typeFilter === 'BURSA TRANSFER') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('BURSA') && u.includes('TRANSFER');
                        });
                    } else if (typeFilter === 'BURSA ÇIKAN') {
                        filtered = list.filter(t => {
                            const u = (t.type || t.islemTuru || '').toUpperCase();
                            return u.includes('BURSA') && (u.includes('ÇIKAN') || u.includes('ÇIKIŞ'));
                        });
                    } else if (typeFilter && typeFilter !== 'ALL') {
                        filtered = list.filter(t => (t.type || t.islemTuru || '').toUpperCase().includes(typeFilter.toUpperCase()));
                    }

                    // Eğer spesifik bir filtre (örneğin KARAİN TRANSFER) varsa, o filtrelenmiş listenin toplamını almalıyız.
                    // Eğer filtre yoksa (ALL), o zaman geleneksel "ÇIKAN" toplamını göstermeye devam ederiz.
                    const cikan = (typeFilter && typeFilter !== 'ALL')
                      ? filtered.reduce((sum, t) => sum + Number(t.quantity || t.adet || 0), 0)
                      : filtered.filter(t => (t.type || t.islemTuru || '').toUpperCase().includes('ÇIKAN')).reduce((sum, t) => sum + Number(t.quantity || t.adet || 0), 0);
                    
                    let mevcut = 0;
                    if (detailModalColumnType === 'ALL') mevcut = detailModalItem.toplamStok || 0;
                    else if (detailModalColumnType.includes('ANKARA')) mevcut = detailModalItem.ankaraMevcut || 0;
                    else if (detailModalColumnType.includes('KARAİN')) mevcut = detailModalItem.karainMevcut || 0;
                    else if (detailModalColumnType.includes('ÇANAKKALE')) mevcut = detailModalItem.canakkaleMevcut || 0;
                    else if (detailModalColumnType.includes('MİLAS')) mevcut = detailModalItem.milasMevcut || 0;
                    else if (detailModalColumnType.includes('BURSA')) mevcut = detailModalItem.bursaMevcut || 0;
                    else mevcut = detailModalItem.toplamStok || 0;

                    return { cikan, mevcut };
                  };

                  const stats = getStatsForType(detailModalColumnType === 'ALL' ? null : detailModalColumnType);
                  
                  // Label Logic
                  let leftLabel = 'TOPLAM ÇIKAN';
                  let rightLabel = 'TOPLAM STOK';
                  let leftValue = stats.cikan;
                  let rightValue = detailModalItem.toplamStok || 0;

                  if (isGelen) {
                      leftLabel = 'TOPLAM GELEN';
                      leftValue = detailModalItem.gelen !== undefined && detailModalItem.gelen !== null ? detailModalItem.gelen : stats.cikan;
                  } else if (detailModalColumnType !== 'ALL') {
                      leftLabel = detailModalColumnType;
                      const region = detailModalColumnType.split(' ')[0];
                      rightLabel = `${region} MEVCUT`;
                      rightValue = stats.mevcut;

                      if (detailModalColumnType === 'ANKARA ÇIKAN') {
                          leftValue = (detailModalItem.ankaraCikan !== undefined && detailModalItem.ankaraCikan !== null && detailModalItem.ankaraCikan > 0)
                            ? detailModalItem.ankaraCikan
                            : stats.cikan;
                      } else if (detailModalColumnType === 'KARAİN TRANSFER') {
                          leftValue = detailModalItem.karainTransfer || stats.cikan;
                      } else if (detailModalColumnType === 'KARAİN ÇIKAN') {
                          leftValue = detailModalItem.karainCikan || stats.cikan;
                      } else if (detailModalColumnType === 'ÇANAKKALE TRANSFER') {
                          leftValue = detailModalItem.canakkaleTransfer || stats.cikan;
                      } else if (detailModalColumnType === 'ÇANAKKALE ÇIKAN') {
                          leftValue = detailModalItem.canakkaleCikan || stats.cikan;
                      } else if (detailModalColumnType === 'MİLAS TRANSFER') {
                          leftValue = detailModalItem.milasTransfer || stats.cikan;
                      } else if (detailModalColumnType === 'MİLAS ÇIKAN') {
                          leftValue = detailModalItem.milasCikan || stats.cikan;
                      } else if (detailModalColumnType === 'BURSA TRANSFER') {
                          leftValue = detailModalItem.bursaTransfer || stats.cikan;
                      } else if (detailModalColumnType === 'BURSA ÇIKAN') {
                          leftValue = detailModalItem.bursaCikan || stats.cikan;
                      } else {
                          leftValue = stats.cikan;
                      }
                  }

                  return (
                    <>
                        <div className="bg-slate-100 border border-slate-200 p-4 rounded-xl shadow-xs">
                            <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider font-mono">
                                {leftLabel}
                            </div>
                            <div className="text-2xl font-black text-slate-800 font-mono mt-0.5">
                                {leftValue}
                            </div>
                        </div>
                        <div className="bg-slate-100 border border-slate-200 p-4 rounded-xl shadow-xs">
                            <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider font-mono">
                                {rightLabel}
                            </div>
                            <div className="text-2xl font-black text-slate-800 font-mono mt-0.5">
                                {rightValue}
                            </div>
                        </div>
                    </>
                  );
                })()}
              </div>

              {/* STOK DURUMU (BÖLGESEL) */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                <button 
                  type="button"
                  onClick={() => setIsRegionalStockOpen(!isRegionalStockOpen)}
                  className="w-full flex items-center justify-between p-4 bg-white hover:bg-slate-50 transition-colors cursor-pointer group"
                >
                  <div className="flex items-center gap-2">
                    <Package className="w-4 h-4 text-slate-600 group-hover:text-emerald-600 transition-colors" />
                    <h4 className="text-[10px] font-black text-slate-700 uppercase font-mono tracking-wider">GÜNCEL STOK DURUMU (BÖLGESEL)</h4>
                  </div>
                  <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform duration-300 ${isRegionalStockOpen ? 'rotate-180' : ''}`} />
                </button>
                
                {isRegionalStockOpen && (
                  <div className="p-4 pt-0 space-y-3 animate-slide-down">
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                      {/* Ankara */}
                      <div className="p-2.5 bg-white border border-[#ff9900] rounded-lg shadow-xs flex flex-col min-h-[70px]">
                        <div className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">ANKARA</div>
                        <div className="flex flex-col mt-1">
                          <div className="text-xl font-black text-slate-900 font-mono">
                            {detailModalItem.ankaraMevcut || 0}
                          </div>
                          <div className="mt-1 text-[10px] text-black font-black bg-yellow-400 px-1.5 py-1 rounded border border-yellow-500 shadow-sm leading-tight text-center">
                            {detailModalItem.lokasyonNo || '-'}
                          </div>
                        </div>
                      </div>
                      {/* Milas */}
                      <div className="p-2.5 bg-white border border-[#ff9900] rounded-lg shadow-xs flex flex-col min-h-[70px]">
                        <div className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">MİLAS</div>
                        <div className="text-xl font-black text-slate-900 font-mono mt-1">
                          {detailModalItem.milasMevcut || 0}
                        </div>
                      </div>
                      {/* Karain */}
                      <div className="p-2.5 bg-white border border-[#ffff00] rounded-lg shadow-xs flex flex-col min-h-[70px]">
                        <div className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">KARAİN</div>
                        <div className="text-xl font-black text-slate-900 font-mono mt-1">
                          {detailModalItem.karainMevcut || 0}
                        </div>
                      </div>
                      {/* Çanakkale */}
                      <div className="p-2.5 bg-white border border-[#92d050] rounded-lg shadow-xs flex flex-col min-h-[70px]">
                        <div className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">ÇANAKKALE</div>
                        <div className="text-xl font-black text-slate-900 font-mono mt-1">
                          {detailModalItem.canakkaleMevcut || 0}
                        </div>
                      </div>
                      {/* Bursa */}
                      <div className="p-2.5 bg-white border border-[#ff00ff] rounded-lg shadow-xs flex flex-col min-h-[70px]">
                        <div className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">BURSA</div>
                        <div className="text-xl font-black text-slate-900 font-mono mt-1">
                          {detailModalItem.bursaMevcut || 0}
                        </div>
                      </div>
                      {/* Muayene */}
                      <div className="p-2.5 bg-white border border-[#bfbfbf] rounded-lg shadow-xs flex flex-col min-h-[70px]">
                        <div className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">MUAYENE/DİĞER</div>
                        <div className="text-xl font-black text-slate-900 font-mono mt-1">
                          {detailModalItem.muayeneToplam || 0}
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Transactions list of this item */}
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <h4 className="text-[10px] font-black text-slate-700 uppercase font-mono tracking-wider">
                    {detailModalColumnType !== 'ALL' ? `${detailModalColumnType} Hareketleri` : 'Tüm Hareket Geçmişi'}
                  </h4>
                </div>

                {(() => {
                  const list = transactions.filter(t => matchTxToItem(t, detailModalItem));
                  
                  // Filter list based on the context (the column clicked)
                  const filteredList = list.filter(t => {
                    if (detailModalColumnType === 'ALL') return true;
                    
                    const typeStr = (t.type || t.islemTuru || '').toUpperCase();
                    const filterUpper = detailModalColumnType.toUpperCase();
                    
                    if (filterUpper === 'GELEN' || filterUpper === 'ANKARA GİREN') {
                      return typeStr.includes('GİRİŞ') || typeStr.includes('GELEN') || typeStr.includes('GİREN') || (typeStr.includes('ANKARA') && typeStr.includes('FAZLA TESPİT'));
                    }
                    if (filterUpper === 'KARAİN TRANSFER') {
                      return typeStr.includes('KARAİN') && typeStr.includes('TRANSFER');
                    }
                    if (filterUpper === 'KARAİN ÇIKAN') {
                      return typeStr.includes('KARAİN') && (typeStr.includes('ÇIKAN') || typeStr.includes('ÇIKIŞ'));
                    }
                    if (filterUpper === 'ÇANAKKALE TRANSFER') {
                      return typeStr.includes('ÇANAKKALE') && typeStr.includes('TRANSFER');
                    }
                    if (filterUpper === 'ÇANAKKALE ÇIKAN') {
                      return typeStr.includes('ÇANAKKALE') && (typeStr.includes('ÇIKAN') || typeStr.includes('ÇIKIŞ'));
                    }
                    if (filterUpper === 'MİLAS TRANSFER') {
                      return typeStr.includes('MİLAS') && typeStr.includes('TRANSFER');
                    }
                    if (filterUpper === 'MİLAS ÇIKAN') {
                      return typeStr.includes('MİLAS') && (typeStr.includes('ÇIKAN') || typeStr.includes('ÇIKIŞ'));
                    }
                    if (filterUpper === 'BURSA TRANSFER') {
                      return typeStr.includes('BURSA') && typeStr.includes('TRANSFER');
                    }
                    if (filterUpper === 'BURSA ÇIKAN') {
                      return typeStr.includes('BURSA') && (typeStr.includes('ÇIKAN') || typeStr.includes('ÇIKIŞ'));
                    }
                    
                    return typeStr.includes(filterUpper);
                  });
                  
                  if (filteredList.length === 0) {
                    return (
                      <div className="py-10 text-center bg-slate-50 rounded-xl border border-dashed border-slate-200">
                        <div className="text-xl mb-1">📋</div>
                        <p className="text-[11px] font-bold text-slate-500 uppercase tracking-tight">Bu kritere uygun hareket kaydı bulunamadı.</p>
                      </div>
                    );
                  }

                  return (
                    <div className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead className="bg-slate-100 text-slate-700 font-mono text-[10px] uppercase font-bold border-b border-slate-200 text-center">
                          <tr>
                            <th className="py-2 px-2 border-r border-slate-200">SIRA</th>
                            <th className="py-2 px-3 border-r border-slate-200 text-left">İŞLEM TÜRÜ</th>
                            <th className="py-2 px-2 border-r border-slate-200">ADET</th>
                            <th className="py-2 px-2 border-r border-slate-200">TARİH</th>
                            <th className="py-2 px-3 border-r border-slate-200 text-left">DEPO / BÖLGE</th>
                            <th className="py-2 px-3 border-r border-slate-200 text-left">TESLİM ALAN / KABUL</th>
                            <th className="py-2 px-2">KUYRUK KODU</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200 font-medium text-center text-[11px]">
                          {filteredList.map((t, i) => {
                            const typeStr = (t.type || t.islemTuru || '').toUpperCase();
                            const isTransfer = typeStr.includes('TRANSFER');
                            const isGiris = typeStr.includes('GİREN') || typeStr.includes('GİRİŞ') || typeStr.includes('GELEN') || typeStr.includes('YENİ ÜRÜN');

                            return (
                              <tr key={i} className="hover:bg-slate-50">
                                <td className="py-2 px-2 font-mono font-bold text-slate-500 border-r border-slate-200">{i + 1}</td>
                                <td className="py-2 px-3 border-r border-slate-200 text-left font-sans">
                                  <span className={`inline-block px-2 py-0.5 rounded text-[9px] font-black uppercase tracking-wider ${
                                    isTransfer 
                                      ? 'bg-indigo-100 text-indigo-900 border border-indigo-200' 
                                      : isGiris 
                                      ? 'bg-emerald-100 text-emerald-900 border border-emerald-200' 
                                      : 'bg-rose-100 text-rose-900 border border-rose-200'
                                  }`}>
                                    {t.type || t.islemTuru || 'TRANSFER'}
                                  </span>
                                </td>
                                <td className={`py-2 px-2 font-mono font-black border-r border-slate-200 ${
                                  isGiris ? 'text-emerald-700' : (isTransfer ? 'text-indigo-700' : 'text-rose-700')
                                }`}>
                                  {t.quantity || t.adet || 1}
                                </td>
                                <td className="py-2 px-2 font-mono text-slate-600 border-r border-slate-200">{t.date || t.timestamp || '-'}</td>
                                <td className="py-2 px-3 font-semibold text-slate-800 border-r border-slate-200 text-left font-sans">{t.location || t.depoYeri || '-'}</td>
                                <td className="py-2 px-3 text-slate-700 border-r border-slate-200 text-left font-sans">
                                  <div className="truncate max-w-[120px]" title={`Alan: ${t.operator || '-'} | Kabul: ${t.receivedBy || '-'}`}>
                                    {t.operator || '-'} / {t.receivedBy || '-'}
                                  </div>
                                </td>
                                <td className="py-2 px-2 font-semibold text-slate-600 font-mono">{t.tailNo || t.kuyrukKodu || '-'}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  );
                })()}
              </div>

            </div>

            {/* Footer */}
            <div className="bg-slate-50 px-6 py-4 border-t border-slate-200 flex items-center justify-end font-sans">
              <button
                type="button"
                onClick={() => {
                  setIsDetailModalOpen(false);
                  setDetailModalItem(null);
                }}
                className="px-5 py-2 bg-slate-900 hover:bg-slate-800 text-white font-extrabold uppercase tracking-wider rounded-xl cursor-pointer transition shadow-sm font-sans"
              >
                Kapat
              </button>
            </div>

          </div>
        </div>
      )}

      {/* TERMAL BARKOD & SLİP BASMA MODALI (RAF SLİBİ / DEPO GİRİŞ SLİBİ) */}
      <DepoSlipPrintModal
        isOpen={isSlipModalOpen}
        onClose={() => {
          setIsSlipModalOpen(false);
          setSlipModalType(null);
          setSlipModalData(null);
          setIsGirisOnlySlip(false);
        }}
        initialType={slipModalType}
        initialData={slipModalData}
        isGirisOnly={isGirisOnlySlip}
        inventory={computedInventory}
        transactions={transactions}
        showNotification={showNotification}
      />

      {/* DEPO BELGE & SERTİFİKA BULUCU MODAL */}
      <DepoBelgeBulModal
        isOpen={isBelgeBulModalOpen}
        onClose={() => setIsBelgeBulModalOpen(false)}
        inventory={computedInventory}
        showNotification={showNotification}
        currentUnit={currentUnit}
      />

      {/* TRANSACTION PASSWORD PROMPT MODAL */}
      {isTxPasswordModalOpen && (
        <div className="fixed inset-0 z-[10500] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl shadow-2xl max-w-sm w-full p-6 space-y-4 text-white">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-amber-500/20 text-amber-400 rounded-xl border border-amber-500/30">
                <Lock className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-100">
                  DÜZENLEME GÜVENLİK ŞİFRESİ
                </h3>
                <p className="text-[10px] text-slate-400">
                  Bu işlem yetkili personel şifresi gerektirmektedir.
                </p>
              </div>
            </div>

            <form onSubmit={handleTxPasswordSubmit} className="space-y-3">
              <div>
                <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">
                  BİRİM DÜZENLEME ŞİFRESİ:
                </label>
                <input
                  type="password"
                  value={txPasswordInput}
                  onChange={(e) => {
                    setTxPasswordInput(e.target.value);
                    setTxPasswordError('');
                  }}
                  autoFocus
                  placeholder="Şifre..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono text-center text-sm focus:outline-none focus:border-amber-500"
                />
              </div>

              {txPasswordError && (
                <div className="text-[11px] text-rose-400 font-bold bg-rose-950/40 p-2 rounded-xl border border-rose-800/50">
                  {txPasswordError}
                </div>
              )}

              <div className="flex items-center justify-end gap-2 pt-1">
                <button
                  type="button"
                  onClick={() => {
                    setIsTxPasswordModalOpen(false);
                    setPendingTxToEdit(null);
                  }}
                  className="px-4 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-300 text-xs font-bold rounded-xl cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-1.5 bg-amber-600 hover:bg-amber-500 text-slate-950 text-xs font-black uppercase rounded-xl cursor-pointer"
                >
                  Şifreyi Doğrula
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* TRANSACTION EDIT MODAL */}
      {isTxEditModalOpen && selectedTxForEdit && (
        <div className="fixed inset-0 z-[10500] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-700 text-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[92vh]">
            
            {/* Header */}
            <div className="bg-slate-950 px-5 py-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-blue-500/20 text-blue-400 rounded-xl border border-blue-500/30">
                  <Edit2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-100">
                    DEPO HAREKET KAYDINI DÜZENLE
                  </h3>
                  <p className="text-[10px] text-slate-400">
                    Seçilen hareket kaydına ait verileri anlık olarak güncelleyin.
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setIsTxEditModalOpen(false)} 
                className="text-slate-400 hover:text-white font-bold"
              >
                ✕
              </button>
            </div>

            {/* Form Body */}
            <form onSubmit={handleSaveTxEdit} className="p-5 space-y-3.5 text-xs overflow-y-auto flex-1 font-sans">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Tarih</label>
                  <input
                    type="text"
                    value={editTxDate}
                    onChange={(e) => setEditTxDate(e.target.value)}
                    className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">İşlem Türü *</label>
                  <select
                    value={editTxType}
                    onChange={(e) => setEditTxType(e.target.value)}
                    className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white font-bold focus:outline-none"
                  >
                    <option value="ANKARA GİREN">ANKARA GİREN</option>
                    <option value="ANKARA ÇIKAN">ANKARA ÇIKAN</option>
                    <option value="KARAİN TRANSFER">KARAİN TRANSFER</option>
                    <option value="KARAİN ÇIKAN">KARAİN ÇIKAN</option>
                    <option value="ÇANAKKALE TRANSFER">ÇANAKKALE TRANSFER</option>
                    <option value="ÇANAKKALE ÇIKAN">ÇANAKKALE ÇIKAN</option>
                    <option value="MİLAS TRANSFER">MİLAS TRANSFER</option>
                    <option value="MİLAS ÇIKAN">MİLAS ÇIKAN</option>
                    <option value="BURSA TRANSFER">BURSA TRANSFER</option>
                    <option value="BURSA ÇIKAN">BURSA ÇIKAN</option>
                    <option value="MUAYENE GİDEN">MUAYENE GİDEN</option>
                    <option value="MUAYENEDEN DÖNEN">MUAYENEDEN DÖNEN</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Malzeme Tanımı (Description) — [SİSTEMDEN OTOMATİK GELİR]</label>
                <input
                  type="text"
                  value={editTxItemName}
                  readOnly
                  disabled
                  placeholder="P/N seçildiğinde otomatik gelir..."
                  className="w-full px-3 py-1.5 bg-slate-950/60 border border-slate-800 rounded-xl text-slate-400 font-bold focus:outline-none cursor-not-allowed opacity-80"
                />
              </div>

              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Part Number (P/N) *</label>
                  <PnAutocomplete
                    value={editTxPn}
                    inventory={computedInventory}
                    placeholder="P/N yazınız..."
                    onChange={(pVal, matched) => {
                      setEditTxPn(pVal);
                      if (matched) {
                        setEditTxItemName(matched.description || matched.name || '');
                        if (matched.serialAndNotes || matched.sn) {
                          setEditTxSn(matched.serialAndNotes || matched.sn || '-');
                        }
                      }
                    }}
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Serial Number (S/N)</label>
                  <input
                    type="text"
                    value={editTxSn}
                    onChange={(e) => setEditTxSn(e.target.value)}
                    placeholder="S/N..."
                    className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Miktar (Adet)</label>
                  <input
                    type="number"
                    value={editTxQuantity}
                    onChange={(e) => setEditTxQuantity(e.target.value)}
                    className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Teslim Alan Personel *</label>
                  <PersonnelAutocomplete
                    value={editTxOperator}
                    onChange={setEditTxOperator}
                    placeholder="Personel seçiniz"
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Kabul Yapan Personel *</label>
                  <PersonnelAutocomplete
                    value={editTxReceivedBy}
                    onChange={setEditTxReceivedBy}
                    placeholder="Personel seçiniz"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Kuyruk No *</label>
                  <select
                    value={editTxTailNo}
                    onChange={(e) => setEditTxTailNo(e.target.value)}
                    className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none"
                  >
                    <option value="-">-- Seçiniz --</option>
                    <option value="TC-HEV">TC-HEV</option>
                    <option value="TC-HEY">TC-HEY</option>
                    <option value="TC-HEZ">TC-HEZ</option>
                    <option value="TC-HKA">TC-HKA</option>
                    <option value="TC-HKB">TC-HKB</option>
                    <option value="TC-HKC">TC-HKC</option>
                    <option value="TC-HKD">TC-HKD</option>
                    <option value="TC-HKE">TC-HKE</option>
                    <option value="TC-HKF">TC-HKF</option>
                    <option value="TC-HKG">TC-HKG</option>
                    <option value="TC-HKH">TC-HKH</option>
                    <option value="TC-HKI">TC-HKI</option>
                    <option value="TC-HKJ">TC-HKJ</option>
                    <option value="TC-HKK">TC-HKK</option>
                    <option value="TC-HKL">TC-HKL</option>
                    <option value="TC-HKM">TC-HKM</option>
                    <option value="TC-HKN">TC-HKN</option>
                    <option value="TC-HKO">TC-HKO</option>
                    <option value="TC-HKP">TC-HKP</option>
                    <option value="TC-HKR">TC-HKR</option>
                    <option value="TC-HKS">TC-HKS</option>
                    <option value="TC-HKT">TC-HKT</option>
                    <option value="TC-HKU">TC-HKU</option>
                    <option value="TC-HKV">TC-HKV</option>
                    <option value="TC-HKY">TC-HKY</option>
                    <option value="TC-HKZ">TC-HKZ</option>
                    <option value="STOK / ATÖLYE">STOK / ATÖLYE</option>
                    <option value="GENEL">GENEL</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Depo Yeri / Lokasyon *</label>
                  <select
                    value={editTxLocation}
                    onChange={(e) => setEditTxLocation(e.target.value)}
                    className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white font-bold focus:outline-none"
                  >
                    <option value="ANKARA">ANKARA (MERKEZ DEPO)</option>
                    <option value="KARAİN">KARAİN DEPOSU</option>
                    <option value="ÇANAKKALE">ÇANAKKALE DEPOSU</option>
                    <option value="MİLAS">MİLAS DEPOSU</option>
                    <option value="BURSA">BURSA DEPOSU</option>
                    <option value="KİMYASAL DEPO">KİMYASAL DEPO</option>
                    <option value="SARF VE PARÇA DEPOSU">SARF VE PARÇA DEPOSU</option>
                    <option value="MUAYENE / BAKIM">MUAYENE / BAKIM</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">Açıklama / Notlar</label>
                <textarea
                  rows={2}
                  value={editTxNotes}
                  onChange={(e) => setEditTxNotes(e.target.value)}
                  className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white focus:outline-none resize-none font-sans"
                />
              </div>

              <div className="pt-3 border-t border-slate-800 flex justify-end gap-2 shrink-0 font-sans">
                <button
                  type="button"
                  onClick={() => setIsTxEditModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-750 text-white rounded-xl font-bold transition cursor-pointer font-sans"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black uppercase rounded-xl shadow transition cursor-pointer font-sans"
                >
                  KAYDET
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ─── DEPO HAREKET GEÇMİŞİ SİLME MODALI (ŞİFRE KORUMALI) ─── */}
      {isTxDeleteModalOpen && pendingTxToDelete && (
        <div className="fixed inset-0 z-[10550] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-rose-500/40 text-white rounded-3xl shadow-2xl max-w-md w-full overflow-hidden flex flex-col">
            <div className="bg-rose-950/70 px-5 py-4 border-b border-rose-900/50 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-rose-500/20 text-rose-400 rounded-xl border border-rose-500/30">
                  <Trash2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-rose-100">
                    DEPO HAREKET KAYDINI SİL
                  </h3>
                  <p className="text-[10px] text-rose-300">
                    Sistemden ve Excel sayfasından kalıcı olarak silme
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => {
                  setIsTxDeleteModalOpen(false);
                  setPendingTxToDelete(null);
                }} 
                className="text-slate-400 hover:text-white font-bold"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleConfirmDeleteTx} className="p-5 space-y-4 text-xs font-sans">
              <div className="p-3 bg-slate-950 border border-slate-800 rounded-2xl space-y-1.5 text-xs">
                <div className="text-slate-300 font-bold">
                  {pendingTxToDelete.itemName || pendingTxToDelete.itemDesc || pendingTxToDelete.name || '-'}
                </div>
                <div className="text-[11px] text-slate-400 font-mono">
                  P/N: <strong className="text-white">{pendingTxToDelete.pn || pendingTxToDelete.partNumber || '-'}</strong> | 
                  Adet: <strong className="text-emerald-400">{pendingTxToDelete.quantity ?? pendingTxToDelete.adet ?? 1}</strong> | 
                  Tarih: <strong className="text-slate-300">{pendingTxToDelete.date || pendingTxToDelete.timestamp || '-'}</strong>
                </div>
                <div className="text-[11px] text-indigo-300 font-mono">
                  İşlem Türü: {pendingTxToDelete.type || pendingTxToDelete.islemTuru || '-'} | Teslim: {pendingTxToDelete.operator || pendingTxToDelete.teslimAlan || '-'}
                </div>
              </div>

              <div className="p-3 bg-rose-500/10 border border-rose-500/30 rounded-xl text-[11px] text-rose-300 leading-relaxed">
                ⚠️ Bu hareket kaydı veritabanından ve Google Drive Excel sayfasındaki satırdan kalıcı olarak silinecektir. İşlemi onaylamak için yetkili şifresini giriniz.
              </div>

              <div>
                <label className="block text-[10px] font-black text-slate-300 uppercase font-mono mb-1">
                  YETKİLİ ŞİFRESİ (1839 veya Birim Şifresi) *
                </label>
                <input
                  type="password"
                  value={txDeletePasswordInput}
                  onChange={(e) => {
                    setTxDeletePasswordInput(e.target.value);
                    setTxDeletePasswordError('');
                  }}
                  autoFocus
                  required
                  placeholder="Şifreyi giriniz..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono focus:outline-none focus:border-rose-500"
                />
                {txDeletePasswordError && (
                  <p className="mt-1 text-[11px] font-bold text-rose-400 font-sans">
                    {txDeletePasswordError}
                  </p>
                )}
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-800">
                <button
                  type="button"
                  onClick={() => {
                    setIsTxDeleteModalOpen(false);
                    setPendingTxToDelete(null);
                  }}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-750 text-slate-300 text-xs font-bold rounded-xl cursor-pointer"
                >
                  Vazgeç
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-rose-600 hover:bg-rose-500 text-white text-xs font-black uppercase rounded-xl cursor-pointer shadow-lg active:scale-95 transition"
                >
                  ŞİFRE İLE SİL
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ─── DEPO MEVCUT ÜRÜN SATIR DÜZENLEME MODALI (ŞİFRE KORUMALI) ─── */}
      {isRowEditModalOpen && rowEditItem && (
        <div className="fixed inset-0 z-[10550] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-emerald-500/40 text-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[92vh]">
            <div className="bg-slate-950 px-5 py-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-emerald-500/20 text-emerald-400 rounded-xl border border-emerald-500/30">
                  <Edit2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-emerald-100">
                    DEPO MEVCUT MALZEME BİLGİLERİNİ DÜZENLE
                  </h3>
                  <p className="text-[10px] text-slate-400">
                    Seçilen satırın bilgilerini Excel ve sistemde güncelleyin (Şifre Korumalı)
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => {
                  setIsRowEditModalOpen(false);
                  setRowEditItem(null);
                }} 
                className="text-slate-400 hover:text-white font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleSaveRowEdit} className="p-5 space-y-3.5 text-xs overflow-y-auto flex-1 font-sans">
              <div>
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  DESCRIPTION (MALZEME ADI / TANIMI) *
                </label>
                <input
                  type="text"
                  value={rowEditDescription}
                  onChange={(e) => setRowEditDescription(e.target.value)}
                  required
                  placeholder="Malzeme adı..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-bold focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    PART NUMBER (P/N)
                  </label>
                  <input
                    type="text"
                    value={rowEditPn}
                    onChange={(e) => setRowEditPn(e.target.value)}
                    placeholder="P/N..."
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none focus:border-emerald-500"
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    SERIAL NUMBER (S/N)
                  </label>
                  <input
                    type="text"
                    value={rowEditSn}
                    onChange={(e) => setRowEditSn(e.target.value)}
                    placeholder="S/N..."
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  LOKASYON NO
                </label>
                <input
                  type="text"
                  list="depo-lokasyon-datalist"
                  value={rowEditLocation}
                  onChange={(e) => setRowEditLocation(e.target.value)}
                  placeholder="Lokasyon Kodu (Örn: A-01-R3)..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div>
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  AÇIKLAMA / NOTLAR
                </label>
                <textarea
                  rows={2}
                  value={rowEditNotes}
                  onChange={(e) => setRowEditNotes(e.target.value)}
                  placeholder="İlgili satır açıklaması..."
                  className="w-full px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-emerald-500 resize-none font-sans"
                />
              </div>

              <div className="grid grid-cols-2 gap-3 p-3 bg-slate-950 border border-slate-800 rounded-2xl">
                <div>
                  <label className="block text-[9.5px] font-black text-amber-300 uppercase font-mono mb-1">
                    RAF ÖMRÜ VAR MI?
                  </label>
                  <select
                    value={rowEditHasShelfLife}
                    onChange={(e) => setRowEditHasShelfLife(e.target.value as 'EVET' | 'HAYIR')}
                    className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-xl text-white font-black focus:outline-none"
                  >
                    <option value="HAYIR">HAYIR</option>
                    <option value="EVET">EVET</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-amber-300 uppercase font-mono mb-1">
                    RAF ÖMRÜ BİTİŞ TARİHİ
                  </label>
                  <input
                    type="date"
                    disabled={rowEditHasShelfLife !== 'EVET'}
                    value={rowEditShelfLifeDate}
                    onChange={(e) => setRowEditShelfLifeDate(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-900 border border-slate-700 rounded-xl text-white font-mono focus:outline-none disabled:opacity-40"
                  />
                </div>
              </div>

              <div className="pt-2">
                <label className="block text-[10px] font-black text-emerald-400 uppercase font-mono mb-1">
                  YETKİLİ ŞİFRESİ (1839 veya Birim Şifresi) *
                </label>
                <input
                  type="password"
                  value={rowEditPassword}
                  onChange={(e) => {
                    setRowEditPassword(e.target.value);
                    setRowEditPasswordError('');
                  }}
                  required
                  placeholder="Düzenlemeyi onaylamak için şifre giriniz..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono focus:outline-none focus:border-emerald-500"
                />
                {rowEditPasswordError && (
                  <p className="mt-1 text-[11px] font-bold text-rose-400 font-sans">
                    {rowEditPasswordError}
                  </p>
                )}
              </div>

              <div className="pt-3 border-t border-slate-800 flex justify-end gap-2 shrink-0 font-sans">
                <button
                  type="button"
                  onClick={() => {
                    setIsRowEditModalOpen(false);
                    setRowEditItem(null);
                  }}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-750 text-white rounded-xl font-bold transition cursor-pointer font-sans"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black uppercase rounded-xl shadow transition cursor-pointer font-sans"
                >
                  ŞİFRE İLE KAYDET
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ─── MUAYENE / BAKIM GÖNDER MODALI ─── */}
      {isMuayeneModalOpen && (
        <div className="fixed inset-0 z-[10550] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-blue-500/40 text-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[92vh]">
            <div className="bg-slate-950 px-5 py-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-blue-500/20 text-blue-400 rounded-xl border border-blue-500/30">
                  <Wrench className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-blue-100">
                    PARÇAYI MUAYENE / BAKIMA GÖNDER
                  </h3>
                  <p className="text-[10px] text-slate-400">
                    Mevcut depodan düşülerek muayene/bakım takibine sevk edilir
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setIsMuayeneModalOpen(false)} 
                className="text-slate-400 hover:text-white font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleMuayeneSubmit} className="p-5 space-y-3.5 text-xs overflow-y-auto flex-1 font-sans">
              {/* Eklenen Parçalar Akordiyon Kartı (Birden Fazla Parça Eklendikçe Büyür) */}
              {muayeneSessionList.length > 0 && (
                <div className="bg-slate-950 border border-blue-500/50 rounded-2xl overflow-hidden shadow-lg mb-1 animate-slide-down">
                  <button
                    type="button"
                    onClick={() => setIsMuayeneSessionListOpen(!isMuayeneSessionListOpen)}
                    className="w-full px-4 py-2.5 bg-blue-950/80 hover:bg-blue-900/70 flex items-center justify-between transition text-xs font-bold text-blue-200 cursor-pointer"
                  >
                    <div className="flex items-center gap-2">
                      <Wrench className="w-4 h-4 text-blue-400" />
                      <span>📦 Sevk Edilen Parçalar ({muayeneSessionList.length} Kalem)</span>
                    </div>
                    <ChevronDown className={`w-4 h-4 text-blue-300 transition-transform duration-200 ${isMuayeneSessionListOpen ? 'rotate-180' : ''}`} />
                  </button>

                  {isMuayeneSessionListOpen && (
                    <div className="p-3 space-y-2 max-h-36 overflow-y-auto divide-y divide-slate-800/80">
                      {muayeneSessionList.map((item, idx) => (
                        <div key={idx} className="pt-2 text-xs flex items-center justify-between gap-2">
                          <div className="min-w-0 flex-1">
                            <div className="font-bold text-white truncate">{item.itemName}</div>
                            <div className="text-[10px] font-mono text-slate-400 truncate">
                              P/N: <strong className="text-blue-300">{item.pn}</strong> | S/N: {item.sn} | Çıkış: {item.location}
                            </div>
                          </div>
                          <span className="font-mono font-black text-emerald-400 text-xs bg-emerald-950 px-2 py-0.5 rounded border border-emerald-800 shrink-0">
                            {item.quantity} Adet
                          </span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* PN Arama & Seçim */}
              <div className="relative">
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  PARÇA NO (P/N) SEÇİMİ *
                </label>
                <div className="relative">
                  <input
                    type="text"
                    value={muayenePn}
                    onChange={(e) => {
                      setMuayenePn(e.target.value);
                      setIsMuayeneDropdownOpen(true);
                    }}
                    onFocus={() => setIsMuayeneDropdownOpen(true)}
                    required
                    placeholder="P/N veya Malzeme adı arayınız..."
                    className="w-full pl-8 pr-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none focus:border-blue-500"
                  />
                  <Search className="w-4 h-4 absolute left-2.5 top-2.5 text-slate-500" />
                </div>

                {isMuayeneDropdownOpen && (
                  <div className="absolute top-full left-0 right-0 z-50 mt-1 bg-slate-950 border-2 border-blue-500/60 rounded-xl shadow-2xl max-h-48 overflow-y-auto divide-y divide-slate-800">
                    {computedInventory.filter(i => {
                      if (!muayenePn.trim()) return true;
                      const q = muayenePn.toLowerCase().trim();
                      const d = (i.description || i.name || '').toLowerCase();
                      const p = (i.partNumber || i.pn || '').toLowerCase();
                      return d.includes(q) || p.includes(q);
                    }).slice(0, 10).map((item, idx) => (
                      <div
                        key={idx}
                        onClick={() => handleSelectMuayeneItem(item)}
                        className="p-2.5 hover:bg-blue-900/40 cursor-pointer text-xs flex items-center justify-between gap-2"
                      >
                        <div>
                          <div className="font-bold text-white">{item.description || item.name}</div>
                          <div className="text-[11px] font-mono text-slate-400">
                            P/N: <strong className="text-blue-300">{item.partNumber || item.pn || '-'}</strong> | S/N: {item.serialAndNotes || item.sn || '-'}
                          </div>
                        </div>
                        <span className="text-[10px] font-mono font-bold bg-blue-950 text-blue-300 border border-blue-800 px-2 py-0.5 rounded-full shrink-0">
                          Stok: {item.ankaraMevcut || item.toplamStok || 0}
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Malzeme Tanımı & SN */}
              <div>
                <label className="block text-[9.5px] font-black text-slate-400 uppercase font-mono mb-1">
                  MALZEME TANIMI (DESCRIPTION) — [SİSTEMDEN OTOMATİK GELİR]
                </label>
                <input
                  type="text"
                  value={muayeneDescription}
                  readOnly
                  disabled
                  placeholder="P/N seçildiğinde otomatik gelir..."
                  className="w-full px-3 py-2 bg-slate-950/60 border border-slate-800 rounded-xl text-slate-400 font-bold focus:outline-none cursor-not-allowed opacity-80 font-sans"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    SERİAL NUMBER (S/N)
                  </label>
                  <input
                    type="text"
                    value={muayeneSn}
                    onChange={(e) => setMuayeneSn(e.target.value)}
                    placeholder="S/N (Satırda varsa otomatik gelir)..."
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-bold focus:outline-none focus:border-blue-500"
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    ÇIKIŞ YAPILACAK DEPO *
                  </label>
                  <select
                    value={muayeneSourceDepot}
                    onChange={(e) => setMuayeneSourceDepot(e.target.value)}
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-bold focus:outline-none"
                  >
                    <option value="ANKARA">ANKARA (MERKEZ DEPO)</option>
                    <option value="KARAİN">KARAİN DEPOSU</option>
                    <option value="ÇANAKKALE">ÇANAKKALE DEPOSU</option>
                    <option value="MİLAS">MİLAS DEPOSU</option>
                    <option value="BURSA">BURSA DEPOSU</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    MİKTAR (ADET) *
                  </label>
                  <input
                    type="number"
                    min={1}
                    value={muayeneQty}
                    onChange={(e) => setMuayeneQty(Math.max(1, parseInt(e.target.value) || 1))}
                    required
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-black focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    MUAYENE GİDİŞ TARİHİ *
                  </label>
                  <input
                    type="text"
                    value={muayeneDate}
                    onChange={(e) => setMuayeneDate(e.target.value)}
                    placeholder="GG.AA.YYYY"
                    required
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono focus:outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  TESLİM EDEN / SEVK EDEN PERSONEL
                </label>
                <PersonnelAutocomplete
                  value={muayeneOperator}
                  onChange={setMuayeneOperator}
                  placeholder="Personel seçiniz veya yazınız"
                />
              </div>

              <div>
                <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                  AÇIKLAMA SÜTUNU / BAKIM DETAYLARI
                </label>
                <textarea
                  rows={2}
                  value={muayeneNotes}
                  onChange={(e) => setMuayeneNotes(e.target.value)}
                  placeholder="Muayene / dış bakım sevk nedeni, firma veya atölye bilgisi..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-blue-500 resize-none font-sans"
                />
              </div>

              <div className="pt-3 border-t border-slate-800 flex items-center justify-end gap-2 shrink-0 font-sans flex-wrap">
                <button
                  type="button"
                  onClick={() => setIsMuayeneModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-750 text-white rounded-xl font-bold transition cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="button"
                  onClick={(e) => handleMuayeneSubmit(e, true)}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-black uppercase rounded-xl shadow transition cursor-pointer flex items-center gap-1.5"
                  title="Mevcut pencereyi kapatmadan başka parça ekleyin"
                >
                  <span>➕ BİR PARÇA DAHA EKLE</span>
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-black uppercase rounded-xl shadow transition cursor-pointer"
                >
                  MUAYENEYE GÖNDER VE KAYDET
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ─── ONAY BEKLEYENLER ŞİFRE MODALI ─── */}
      {isOnayPasswordModalOpen && (
        <div className="fixed inset-0 z-[10550] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-amber-500/40 text-white rounded-3xl shadow-2xl max-w-sm w-full overflow-hidden">
            <div className="bg-amber-950/70 px-5 py-4 border-b border-amber-900/50 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Clock className="w-5 h-5 text-amber-400" />
                <h3 className="text-xs font-black uppercase tracking-wider text-amber-200">
                  ONAY BEKLEYENLER YETKİ GİRİŞİ
                </h3>
              </div>
              <button 
                type="button" 
                onClick={() => setIsOnayPasswordModalOpen(false)} 
                className="text-slate-400 hover:text-white font-bold"
              >
                ✕
              </button>
            </div>
            <form onSubmit={handleOnayPasswordSubmit} className="p-5 space-y-4 text-xs font-sans">
              <p className="text-[11px] text-slate-300">
                Depo çıkış taleplerini görüntülemek ve onaylamak için yetkili şifresini (1839) giriniz.
              </p>
              <div>
                <input
                  type="password"
                  value={onayPasswordInput}
                  onChange={(e) => {
                    setOnayPasswordInput(e.target.value);
                    setOnayPasswordError('');
                  }}
                  autoFocus
                  required
                  placeholder="Yetkili şifresini giriniz..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono focus:outline-none focus:border-amber-500"
                />
                {onayPasswordError && (
                  <p className="mt-1 text-[11px] font-bold text-rose-400">{onayPasswordError}</p>
                )}
              </div>
              <div className="flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setIsOnayPasswordModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-750 text-slate-300 rounded-xl font-bold cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-amber-600 hover:bg-amber-500 text-slate-950 font-black uppercase rounded-xl cursor-pointer"
                >
                  GİRİŞ YAP
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ─── ONAY BEKLEYENLER LİSTESİ MODALI (ONAY BEKLEYENLER -AT802) ─── */}
      {isOnayBekleyenlerModalOpen && (
        <div className="fixed inset-0 z-[10550] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-amber-500/40 text-white rounded-3xl shadow-2xl max-w-4xl w-full overflow-hidden flex flex-col max-h-[92vh]">
            <div className="bg-slate-950 px-5 py-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-amber-500/20 text-amber-400 rounded-xl border border-amber-500/30">
                  <Clock className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-amber-200">
                    ONAY BEKLEYEN ÇIKIŞ TALEPLERİ — {currentUnit === 'all' ? 'TÜM BİRİMLER' : currentUnit.toUpperCase()}
                  </h3>
                  <p className="text-[10px] text-slate-400">
                    Drive Excel 'ONAY BEKLEYENLER' sayfası ile senkronize canlı talep listesi
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={fetchOnayBekleyenler}
                  className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs flex items-center gap-1 cursor-pointer"
                  title="Yenile"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Yenile</span>
                </button>
                <button 
                  type="button" 
                  onClick={() => setIsOnayBekleyenlerModalOpen(false)} 
                  className="text-slate-400 hover:text-white font-bold cursor-pointer"
                >
                  ✕
                </button>
              </div>
            </div>

            <div className="p-5 overflow-y-auto flex-1 space-y-3 font-sans text-xs">
              {(() => {
                const filteredOnayList = onayBekleyenlerList.filter(reqItem => {
                  if (!currentUnit || currentUnit === 'all') return true;
                  const itemUnit = (reqItem.unit || 'at802').toLowerCase();
                  const cUnit = currentUnit.toLowerCase();
                  return itemUnit === cUnit || (cUnit === 'at802' && (!reqItem.unit || itemUnit === 'at802'));
                });

                if (isLoadingOnayList) {
                  return (
                    <div className="py-12 text-center text-slate-400">
                      <div className="animate-spin text-2xl mb-2">⏳</div>
                      <p>Talepler Drive Excel tablosundan yükleniyor...</p>
                    </div>
                  );
                }

                if (filteredOnayList.length === 0) {
                  return (
                    <div className="py-12 text-center text-slate-400 border border-dashed border-slate-800 rounded-2xl">
                      <div className="text-3xl mb-2">✅</div>
                      <h4 className="font-bold text-white text-sm">Onay Bekleyen Talep Yok</h4>
                      <p className="text-xs text-slate-500 mt-1">
                        {currentUnit === 'all' ? 'Tüm birimler' : currentUnit.toUpperCase()} için henüz onay bekleyen talep bulunmuyor.
                      </p>
                    </div>
                  );
                }

                return (
                  <div className="space-y-3">
                    {filteredOnayList.map((reqItem, idx) => {
                    const isPending = reqItem.status === 'ONAY BEKLİYOR';
                    const isEditing = editingOnayId === reqItem.id;

                    if (isEditing) {
                      return (
                        <div key={idx} className="p-4 rounded-2xl border border-blue-500/60 bg-slate-950 space-y-3">
                          <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                            <span className="font-mono font-bold text-amber-400 text-xs">DÜZENLEME: {reqItem.id}</span>
                            <button
                              type="button"
                              onClick={() => setEditingOnayId(null)}
                              className="text-slate-400 hover:text-white font-bold"
                            >
                              ✕ İptal
                            </button>
                          </div>

                          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                            <div>
                              <label className="block text-[9px] text-slate-400 font-bold mb-1">PART NUMBER (P/N) *</label>
                              <PnAutocomplete
                                value={editOnayPn}
                                inventory={computedInventory}
                                placeholder="P/N yazınız..."
                                onChange={(pVal, matched) => {
                                  setEditOnayPn(pVal);
                                  if (matched) {
                                    setEditOnayItemName(matched.description || matched.name || '');
                                    if (matched.serialAndNotes || matched.sn) {
                                      setEditOnaySn(matched.serialAndNotes || matched.sn || '-');
                                    }
                                  }
                                }}
                              />
                            </div>
                            <div>
                              <label className="block text-[9px] text-slate-400 font-bold mb-1">SERIAL NUMBER (S/N)</label>
                              <input
                                type="text"
                                value={editOnaySn}
                                onChange={(e) => setEditOnaySn(e.target.value)}
                                className="w-full px-2.5 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono"
                              />
                            </div>
                            <div>
                              <label className="block text-[9px] text-slate-400 font-bold mb-1">MİKTAR (ADET)</label>
                              <input
                                type="number"
                                min={1}
                                value={editOnayQuantity}
                                onChange={(e) => setEditOnayQuantity(Number(e.target.value))}
                                className="w-full px-2.5 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono font-bold"
                              />
                            </div>
                          </div>

                          <div>
                            <label className="block text-[9px] text-slate-400 font-bold mb-1">MALZEME ADI / TANIMI — [SİSTEMDEN OTOMATİK GELİR]</label>
                            <input
                              type="text"
                              value={editOnayItemName}
                              readOnly
                              disabled
                              placeholder="P/N seçildiğinde otomatik gelir..."
                              className="w-full px-2.5 py-1 bg-slate-900/60 border border-slate-800 rounded-lg text-slate-400 font-bold cursor-not-allowed opacity-80"
                            />
                          </div>

                          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                            <div>
                              <label className="block text-[9px] text-slate-400 font-bold mb-1">TALEP DEPOSU</label>
                              <select
                                value={editOnayDepot}
                                onChange={(e) => setEditOnayDepot(e.target.value)}
                                className="w-full px-2.5 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white font-bold"
                              >
                                <option value="ANKARA">ANKARA (MERKEZ DEPO)</option>
                                <option value="KARAİN">KARAİN DEPOSU</option>
                                <option value="ÇANAKKALE">ÇANAKKALE DEPOSU</option>
                                <option value="MİLAS">MİLAS DEPOSU</option>
                                <option value="BURSA">BURSA DEPOSU</option>
                                <option value="KİMYASAL DEPO">KİMYASAL DEPO</option>
                                <option value="SARF VE PARÇA DEPOSU">SARF VE PARÇA DEPOSU</option>
                              </select>
                            </div>

                            <div>
                              <label className="block text-[9px] text-slate-400 font-bold mb-1">KUYRUK NO</label>
                              <select
                                value={editOnayTailNo}
                                onChange={(e) => setEditOnayTailNo(e.target.value)}
                                className="w-full px-2.5 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white font-mono font-bold"
                              >
                                <option value="-">-- Seçiniz --</option>
                                <option value="TC-HEV">TC-HEV</option>
                                <option value="TC-HEY">TC-HEY</option>
                                <option value="TC-HEZ">TC-HEZ</option>
                                <option value="TC-HKA">TC-HKA</option>
                                <option value="TC-HKB">TC-HKB</option>
                                <option value="TC-HKC">TC-HKC</option>
                                <option value="TC-HKD">TC-HKD</option>
                                <option value="TC-HKE">TC-HKE</option>
                                <option value="TC-HKF">TC-HKF</option>
                                <option value="TC-HKG">TC-HKG</option>
                                <option value="TC-HKH">TC-HKH</option>
                                <option value="TC-HKI">TC-HKI</option>
                                <option value="TC-HKJ">TC-HKJ</option>
                                <option value="TC-HKK">TC-HKK</option>
                                <option value="TC-HKL">TC-HKL</option>
                                <option value="TC-HKM">TC-HKM</option>
                                <option value="TC-HKN">TC-HKN</option>
                                <option value="TC-HKO">TC-HKO</option>
                                <option value="TC-HKP">TC-HKP</option>
                                <option value="TC-HKR">TC-HKR</option>
                                <option value="TC-HKS">TC-HKS</option>
                                <option value="TC-HKT">TC-HKT</option>
                                <option value="TC-HKU">TC-HKU</option>
                                <option value="TC-HKV">TC-HKV</option>
                                <option value="TC-HKY">TC-HKY</option>
                                <option value="TC-HKZ">TC-HKZ</option>
                                <option value="STOK / ATÖLYE">STOK / ATÖLYE</option>
                                <option value="GENEL">GENEL</option>
                              </select>
                            </div>

                            <div>
                              <label className="block text-[9px] text-slate-400 font-bold mb-1">TESLİM ALACAK PERSONEL</label>
                              <PersonnelAutocomplete
                                value={editOnayRequestedBy}
                                onChange={setEditOnayRequestedBy}
                                placeholder="Personel seçiniz"
                              />
                            </div>
                          </div>

                          <div>
                            <label className="block text-[9px] text-slate-400 font-bold mb-1">AÇIKLAMA / NOTLAR</label>
                            <input
                              type="text"
                              value={editOnayNotes}
                              onChange={(e) => setEditOnayNotes(e.target.value)}
                              className="w-full px-2.5 py-1 bg-slate-900 border border-slate-700 rounded-lg text-white"
                            />
                          </div>

                          <div className="flex justify-end gap-2 pt-2 border-t border-slate-800">
                            <button
                              type="button"
                              onClick={() => setEditingOnayId(null)}
                              className="px-3 py-1 bg-slate-800 text-slate-300 rounded-lg font-bold hover:bg-slate-700"
                            >
                              İptal
                            </button>
                            <button
                              type="button"
                              onClick={() => handleSaveOnayEdit(reqItem.id)}
                              className="px-4 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-black uppercase shadow"
                            >
                              ✓ GÜNCELLE VE KAYDET
                            </button>
                          </div>
                        </div>
                      );
                    }

                    return (
                      <div
                        key={idx}
                        className={`p-4 rounded-2xl border transition ${
                          isPending 
                            ? 'bg-slate-950 border-amber-500/40 hover:border-amber-400' 
                            : reqItem.status === 'ONAYLANDI'
                            ? 'bg-slate-950/60 border-emerald-500/30 opacity-75'
                            : 'bg-slate-950/60 border-rose-500/30 opacity-60'
                        }`}
                      >
                        <div className="flex flex-wrap items-start justify-between gap-3">
                          <div className="space-y-1">
                            <div className="flex items-center gap-2">
                              <span className="font-mono font-bold text-amber-400 text-xs">{reqItem.id}</span>
                              <span className={`px-2 py-0.5 rounded text-[10px] font-black uppercase ${
                                isPending ? 'bg-amber-500/20 text-amber-300 border border-amber-500/30 animate-pulse' :
                                reqItem.status === 'ONAYLANDI' ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30' :
                                'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                              }`}>
                                {reqItem.status}
                              </span>
                              <span className="text-[10px] text-slate-500 font-mono">📅 {reqItem.date}</span>
                              <span className="text-[10px] bg-slate-800 text-slate-300 px-1.5 py-0.5 rounded font-mono">
                                {reqItem.category === 'kimyasal' ? '🧪 KİMYASAL' : '📦 SARF/PARÇA'}
                              </span>
                            </div>

                            <h4 className="text-sm font-bold text-white mt-1">
                              {reqItem.itemName || reqItem.description}
                            </h4>

                            <div className="text-[11px] font-mono text-slate-400 flex flex-wrap gap-x-4 gap-y-1">
                              <span>P/N: <strong className="text-white">{reqItem.pn || '-'}</strong></span>
                              <span>S/N: <strong className="text-slate-300">{reqItem.sn || '-'}</strong></span>
                              <span>Talep Deposu: <strong className="text-amber-300">{reqItem.depot || 'ANKARA'}</strong></span>
                              <span>Adet: <strong className="text-emerald-400">{reqItem.quantity || reqItem.adet || 1}</strong></span>
                              <span>Kuyruk No: <strong className="text-indigo-300">{reqItem.tailNo || '-'}</strong></span>
                              <span>Teslim Alacak: <strong className="text-white">{reqItem.requestedBy || '-'}</strong></span>
                            </div>

                            {reqItem.notes && (
                              <p className="text-[11px] text-slate-400 italic mt-1 bg-slate-900/80 p-2 rounded-lg border border-slate-800">
                                💬 {reqItem.notes}
                              </p>
                            )}
                          </div>

                          <div className="flex items-center gap-2 shrink-0">
                            {isPending && (
                              <button
                                type="button"
                                onClick={() => handleStartEditOnay(reqItem)}
                                className="px-3 py-1.5 bg-blue-900/60 hover:bg-blue-800 text-blue-200 border border-blue-700/50 rounded-xl font-bold cursor-pointer transition text-xs flex items-center gap-1"
                              >
                                <Edit2 className="w-3.5 h-3.5" />
                                <span>Düzenle</span>
                              </button>
                            )}
                            {isPending && (
                              <button
                                type="button"
                                onClick={() => handleRejectRequest(reqItem)}
                                className="px-3 py-1.5 bg-rose-900/60 hover:bg-rose-800 text-rose-200 border border-rose-700/50 rounded-xl font-bold cursor-pointer transition text-xs"
                              >
                                Reddet
                              </button>
                            )}
                            {isPending && (
                              <button
                                type="button"
                                onClick={() => handleApproveRequest(reqItem)}
                                className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl font-black uppercase cursor-pointer shadow-md transition text-xs"
                              >
                                ✓ ONAYLA VE ÇIKIŞ YAP
                              </button>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              );
            })()}
            </div>
          </div>
        </div>
      )}

      {/* ─── TRANSFER GERİ ÇEKME MODALI ─── */}
      {isRollbackModalOpen && rollbackTx && (
        <div className="fixed inset-0 z-[10600] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-indigo-500/50 text-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden">
            <div className="bg-slate-950 px-5 py-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-indigo-500/20 text-indigo-400 rounded-xl border border-indigo-500/30">
                  <RotateCcw className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-indigo-200">
                    TRANSFER GERİ ÇEKME (STOĞA İADE)
                  </h3>
                  <p className="text-[10px] text-slate-400">
                    Geri çekilen miktar orijinal çıkış yapılan depoya iade edilir ve kayıt altına alınır.
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setIsRollbackModalOpen(false)} 
                className="text-slate-400 hover:text-white font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleConfirmRollback} className="p-5 space-y-4 text-xs font-sans">
              <div className="bg-slate-950 p-3.5 rounded-2xl border border-slate-800 space-y-2">
                <div className="text-[10px] font-bold text-indigo-400 uppercase font-mono">
                  MEVCUT TRANSFER BİLGİLERİ
                </div>
                <h4 className="text-sm font-bold text-white">{rollbackTx.itemName}</h4>
                <div className="text-[11px] font-mono text-slate-400 flex flex-wrap gap-x-4 gap-y-1">
                  <span>P/N: <strong className="text-white">{rollbackTx.pn || '-'}</strong></span>
                  <span>S/N: <strong>{rollbackTx.sn || '-'}</strong></span>
                  <span>Transfer Tipi: <strong className="text-amber-300">{rollbackTx.type || rollbackTx.islemTuru}</strong></span>
                  <span>Orijinal Miktar: <strong className="text-emerald-400">{rollbackTx.quantity} Adet</strong></span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    GERİ ÇEKİLECEK ADET *
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={Number(rollbackTx.quantity || 1)}
                    value={rollbackQty}
                    onChange={(e) => setRollbackQty(Math.max(1, Math.min(Number(rollbackTx.quantity || 1), parseInt(e.target.value) || 1)))}
                    required
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white font-mono font-black text-sm focus:outline-none focus:border-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-[9.5px] font-black text-slate-300 uppercase font-mono mb-1">
                    İŞLEMİ YAPAN PERSONEL *
                  </label>
                  <PersonnelAutocomplete
                    value={rollbackOperator}
                    onChange={setRollbackOperator}
                    placeholder="Personel seçiniz"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[9.5px] font-black text-rose-300 uppercase font-mono mb-1">
                  GERİ ÇEKME NEDENİ (ZORUNLU) *
                </label>
                <textarea
                  rows={2}
                  value={rollbackReason}
                  onChange={(e) => setRollbackReason(e.target.value)}
                  required
                  placeholder="Transferin neden geri çekildiğini açıklayınız (Örn: Hatalı transfer, atölye geri iadesi, görev iptali vb.)..."
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-white focus:outline-none focus:border-indigo-500 resize-none font-sans"
                />
              </div>

              <div className="pt-3 border-t border-slate-800 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setIsRollbackModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-750 text-white rounded-xl font-bold transition cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-black uppercase rounded-xl shadow-lg transition cursor-pointer"
                >
                  ↩️ TRANSFERİ GERİ ÇEK VE STOĞA İADE ET
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  );
};
