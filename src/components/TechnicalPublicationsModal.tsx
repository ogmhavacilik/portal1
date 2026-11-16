import React, { useState, useEffect, useRef, useMemo } from 'react';
import {
  X,
  BookOpen,
  Search,
  Upload,
  Plus,
  Trash2,
  Edit3,
  Lock,
  Unlock,
  CheckCircle,
  AlertCircle,
  FileText,
  ExternalLink,
  ChevronDown,
  ChevronUp,
  Download,
  Archive,
  ArrowRight,
  Sparkles,
  Calendar,
  AlertTriangle,
  ArrowLeft,
  Eye,
  Wrench,
  Package,
  Layers,
  FileSpreadsheet,
  Check,
  Clipboard,
  RotateCw,
  Cloud,
  CheckCheck,
  ShieldCheck,
  Plane,
  CheckCircle2,
  RefreshCw,
  Camera,
  Crop,
  Printer,
  Globe,
  Image as ImageIcon,
  Link2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { GOOGLE_SCRIPT_URL, fileToBase64 } from '../App';
import { cachePdfBase64, getCachedPdfBlobUrl } from '../utils/pdfCache';
import { unmergeAndFillWorksheet, isHeaderLikeRow } from '../utils/driveExcelSync';
import { fetchExcelFromDrive as remoteFetchExcelFromDrive, DRIVE_FOLDER_ID, DIRECT_GAS_URLS } from '../utils/remoteDriveSync';
import JSZip from 'jszip';

export const OGM_BULUT_AT802_URL = 'https://bulut.ogm.gov.tr/app/tr-TR/Klasor/Paylas/OGMDrive/0b7eb51b-6174-4314-89bc-c26d74d78acc';
export const OGM_EXTENSION_STORE_URL = 'https://chromewebstore.google.com/detail/iframe-unlocker-%E2%80%93-remove/jkijbninecngmcipaijcnibjhbejaoch?pli=1';

export const OGM_WIKIMEDIA_LOGO_URL = 'https://upload.wikimedia.org/wikipedia/tr/d/d4/Orman_Genel_M%C3%BCd%C3%BCrl%C3%BC%C4%9F%C3%BC_logo.png?utm_source=tr.wikipedia.org&utm_campaign=index&utm_content=original';

// Resmi OGM dairesel amblem SVG fallback veri URI'si
const OGM_LEFT_SEAL_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120"><circle cx="60" cy="60" r="56" fill="#f0fdf4" stroke="#166534" stroke-width="5"/><circle cx="60" cy="60" r="46" fill="none" stroke="#15803d" stroke-width="1.5" stroke-dasharray="3 2"/><path d="M60 26 L44 56 H54 L40 80 H80 L66 56 H76 Z" fill="#15803d"/><rect x="56" y="80" width="8" height="12" fill="#78350f"/><text x="60" y="20" text-anchor="middle" font-size="8.5" font-weight="bold" fill="#14532d" font-family="Arial,sans-serif">ORMAN GENEL MÜDÜRLÜĞÜ</text><text x="60" y="106" text-anchor="middle" font-size="7.5" font-weight="bold" fill="#14532d" font-family="Arial,sans-serif">T.C. TARIM VE ORMAN BAKANLIĞI</text></svg>`;
const OGM_LEFT_SEAL_DATA_URI = `data:image/svg+xml;utf8,${encodeURIComponent(OGM_LEFT_SEAL_SVG)}`;
const OGM_RIGHT_SEAL_DATA_URI = OGM_LEFT_SEAL_DATA_URI;

export interface ZeroReportRow {
  siraNo: number;
  parcaAdi: string;
  parcaNo: string;
  alternatifPn: string;
  istenenMiktar: string;
  oncelik: 'Yüksek' | 'Normal' | 'Düşük' | '';
}

export interface ZeroReportFormState {
  havaAraciTipi: string;
  selectedTypes: {
    at802: boolean;
    t70: boolean;
    bell429: boolean;
    b360: boolean;
  };
  rows: ZeroReportRow[];
  ekNotlar: string;
  talepYapanAdSoyad: string;
  date: string;
}

export interface TechPublication {
  id: string;
  unit: 'at802' | 'bell429' | 't70' | 't70_bumbi' | 't70_helitak' | 'c650' | 'b360' | 'hangar' | string;
  unitKey?: string;
  unitLabel: string;
  category?: string; // 'AMM' | 'IPC' | 'CMM' | 'ŞEMA' | 'EL KİTABI' | 'STANDART' | 'DİĞER'
  revision?: string; // e.g. "Rev. 01"
  section: string; // e.g. "Bölüm 1 - Gövde", "IPC", "AMM", "Hidrolik"
  title: string;
  fileName: string;
  fileSize?: string;
  uploadDate: string; // "2026-08-20" or "20.08.2026 14:30"
  driveFileId?: string;
  viewUrl?: string;
  base64Data?: string;
  notes?: string;
}

export interface DepotItem {
  unitKey: string;
  unitTitle: string;
  section: string;
  siraNo: string;
  adi: string;
  parcaNo: string;
  seriNo: string;
  miktar: string;
  miktarNum: number;
  yer: string;
  durumu: string;
  sonKontrol: string;
  gelecekKontrol: string;
  firma: string;
  aciklama: string;
  // Regional stock fields for AT-802
  ankaraMevcut?: number;
  milasMevcut?: number;
  karainMevcut?: number;
  canakkaleMevcut?: number;
  bursaMevcut?: number;
  toplamStok?: number;
  isSarfItem?: boolean;
}

// Global in-memory cache for fast instant PDF loading
const pdfBase64Cache = new Map<string, string>();

interface TechnicalPublicationsModalProps {
  isOpen: boolean;
  onClose: () => void;
  allUnitData?: {
    at802?: any[];
    bell429?: any[];
    t70?: any[];
    t70_bumbi?: any[];
    t70_helitak?: any[];
    c650?: any[];
    b360?: any[];
    hangar?: any[];
  };
  onNavigateToDepo?: (unitKey: string, section: string) => void;
  onNavigateToEquipment?: (type: string, label: string, section?: string) => void;
}

// Şifre eşleştirmeleri:
// AT-802: 802
// T-70: 70
// BELL 429: 429
// C 650: 650
// B 360: 360
// HANGAR: 1839
const UNIT_PASSWORDS: Record<string, string> = {
  'at802': '802',
  'at-802': '802',
  'bell429': '429',
  'bell 429': '429',
  't70': '70',
  't-70': '70',
  't70_bumbi': '70',
  't70_helitak': '70',
  'c650': '650',
  'c 650': '650',
  'b360': '360',
  'b 360': '360',
  'hangar': '1839'
};

const UNIT_FOLDER_OPTIONS = [
  { key: 'at802', label: 'AT-802F', type: 'UÇAK', sub: 'Air Tractor Yangın Söndürme Uçağı', tag: 'AT-', badgeBg: 'bg-amber-600 text-white', password: '802' },
  { key: 'bell429', label: 'BELL 429', type: 'HELİKOPTER', sub: 'Bell 429 Keşif & Yangın Helikopteri', tag: 'BEL', badgeBg: 'bg-teal-700 text-white', password: '429' },
  { key: 't70', label: 'T-70', type: 'HELİKOPTER', sub: 'T-70 Genel Maksat Yangın Helikopteri', tag: 'T-7', badgeBg: 'bg-emerald-800 text-white', password: '70' },
  { key: 'c650', label: 'C-650', type: 'UÇAK', sub: 'Cessna Citation Sovereign Uçağı', tag: 'C-6', badgeBg: 'bg-blue-600 text-white', password: '650' },
  { key: 'b360', label: 'B-360', type: 'UÇAK', sub: 'Beechcraft King Air 360 Uçağı', tag: 'B-3', badgeBg: 'bg-indigo-600 text-white', password: '360' },
  { key: 'hangar', label: 'HANGAR & GENEL', type: 'GENEL', sub: 'Hangar Standartları ve El Kitapları', tag: 'HAN', badgeBg: 'bg-slate-700 text-white', password: '1839' }
];

const CATEGORY_TABS = [
  'TÜM KATEGORİLER',
  'AMM',
  'IPC',
  'CMM',
  'ŞEMA',
  'EL KİTABI',
  'STANDART'
];

const SECTION_SUGGESTIONS = [
  'IPC - Parça Kataloğu (Illustrated Parts Catalog)',
  'AMM - Uçak Bakım El Kitabı (Aircraft Maintenance Manual)',
  'CMM - Komponent Bakım El Kitabı (Component Maint. Manual)',
  'WDM - Kablolama ve Şemalar (Wiring Diagram)',
  'SB / AD - Servis Bültenleri ve Direktifler',
  'Bölüm 1 - Gövde ve Yapısal (Airframe)',
  'Bölüm 2 - Hidrolik ve Pnömatik',
  'Bölüm 3 - Motor ve Pervane (Powerplant)',
  'Bölüm 4 - Elektrik ve Aviyonik',
  'Bölüm 5 - Yangın Görev Donanımı / Kit',
  'Genel Teknik Döküman'
];

export const TechnicalPublicationsModal: React.FC<TechnicalPublicationsModalProps> = ({
  isOpen,
  onClose,
  allUnitData,
  onNavigateToDepo
}) => {
  // Publications storage
  const [publications, setPublications] = useState<TechPublication[]>([]);
  const [selectedUnitFilter, setSelectedUnitFilter] = useState<string>('');
  const [selectedCategoryFilter, setSelectedCategoryFilter] = useState<string>('TÜM KATEGORİLER');
  const [pubSearchQuery, setPubSearchQuery] = useState<string>('');

  // Unit PIN / Password Protection State - None unlocked by default
  const [unlockedUnits, setUnlockedUnits] = useState<Set<string>>(new Set<string>());
  const [unitPasswordModalOpen, setUnitPasswordModalOpen] = useState<boolean>(false);
  const [targetUnitToUnlock, setTargetUnitToUnlock] = useState<string | null>(null);
  const [enteredUnitPassword, setEnteredUnitPassword] = useState<string>('');
  const [unitPasswordError, setUnitPasswordError] = useState<string>('');

  // Active viewing publication (Viewer mode)
  const [activePub, setActivePub] = useState<TechPublication | null>(null);
  const [isViewerActive, setIsViewerActive] = useState<boolean>(false);

  // Upload modal state
  const [isUploadModalOpen, setIsUploadModalOpen] = useState<boolean>(false);
  const [uploadUnit, setUploadUnit] = useState<string>('at802');
  const [uploadCategory, setUploadCategory] = useState<string>('IPC');
  const [uploadRevision, setUploadRevision] = useState<string>(`Rev. ${new Date().toLocaleDateString('tr-TR')}`);
  const [uploadSection, setUploadSection] = useState<string>('IPC - Parça Kataloğu (Illustrated Parts Catalog)');
  const [uploadCustomSection, setUploadCustomSection] = useState<string>('');
  const [uploadRevisionDate, setUploadRevisionDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [successCount, setSuccessCount] = useState<number>(0);
  const [failCount, setFailCount] = useState<number>(0);
  const [selectedFiles, setSelectedFiles] = useState<{
    file: File;
    customTitle: string;
    customSection: string;
    customUnit: string;
    customCategory: string;
    customRevision: string;
    progress?: number;
    status?: 'pending' | 'uploading' | 'success' | 'error';
  }[]>([]);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [uploadProgress, setUploadProgress] = useState<string>('');

  // Password verification modal
  const [passwordModalOpen, setPasswordModalOpen] = useState<boolean>(false);
  const [passwordTargetAction, setPasswordTargetAction] = useState<'edit' | 'delete' | null>(null);
  const [passwordTargetPub, setPasswordTargetPub] = useState<TechPublication | null>(null);
  const [enteredPassword, setEnteredPassword] = useState<string>('');
  const [passwordError, setPasswordError] = useState<string>('');

  // Edit publication modal
  const [editModalOpen, setEditModalOpen] = useState<boolean>(false);
  const [editingPub, setEditingPub] = useState<TechPublication | null>(null);
  const [editTitle, setEditTitle] = useState<string>('');
  const [editSection, setEditSection] = useState<string>('');
  const [editUnit, setEditUnit] = useState<string>('at802');
  const [editCategory, setEditCategory] = useState<string>('IPC');
  const [editRevision, setEditRevision] = useState<string>('Rev. 01');
  const [editNotes, setEditNotes] = useState<string>('');

  // Depo live search, Autocomplete & Drawer state
  const [depoSearchQuery, setDepoSearchQuery] = useState<string>('');
  const [isDepoAutocompleteOpen, setIsDepoAutocompleteOpen] = useState<boolean>(false);
  const [autocompleteActiveIdx, setAutocompleteActiveIdx] = useState<number>(-1);
  const viewerSearchWrapperRef = useRef<HTMLDivElement>(null);
  const globalSearchWrapperRef = useRef<HTMLDivElement>(null);
  const [liveSarfDepoItems, setLiveSarfDepoItems] = useState<any[]>([]);
  const [isDepotLoading, setIsDepotLoading] = useState<boolean>(false);

  const [isStockDrawerOpen, setIsStockDrawerOpen] = useState<boolean>(false);
  const [drawerHeight, setDrawerHeight] = useState<number>(240);
  const isDraggingRef = useRef<boolean>(false);
  const startYRef = useRef<number>(0);
  const startHeightRef = useRef<number>(240);

  // Zero Report Modal & 2-page form state
  const [zeroReportModalOpen, setZeroReportModalOpen] = useState<boolean>(false);
  const [zeroReportData, setZeroReportData] = useState<{
    query: string;
    targetUnit: string;
    date: string;
    items: DepotItem[];
  } | null>(null);
  const [zeroReportForm, setZeroReportForm] = useState<ZeroReportFormState>({
    havaAraciTipi: 'AT-802F',
    selectedTypes: { at802: true, t70: false, bell429: false, b360: false },
    rows: Array.from({ length: 20 }, (_, i) => ({
      siraNo: i + 1,
      parcaAdi: '',
      parcaNo: '',
      alternatifPn: '',
      istenenMiktar: '',
      oncelik: ''
    })),
    ekNotlar: '',
    talepYapanAdSoyad: '',
    date: ''
  });
  const [capturedCatalogImage, setCapturedCatalogImage] = useState<string>('');
  const [circularOgmLogoDataUri, setCircularOgmLogoDataUri] = useState<string>(OGM_WIKIMEDIA_LOGO_URL);

  // Render Wikimedia OGM logo into a true circular transparent-corner PNG Data URI (so Word .doc & print match preview)
  useEffect(() => {
    let cancelled = false;
    const renderCircularLogo = (src: string, fallbackSrc?: string) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        if (cancelled) return;
        try {
          const size = 220;
          const canvas = document.createElement('canvas');
          canvas.width = size;
          canvas.height = size;
          const ctx = canvas.getContext('2d');
          if (!ctx) return;
          ctx.clearRect(0, 0, size, size);
          ctx.save();
          ctx.beginPath();
          ctx.arc(size / 2, size / 2, size / 2 - 2, 0, Math.PI * 2);
          ctx.closePath();
          ctx.clip();
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, size, size);
          const scale = Math.max(size / img.width, size / img.height);
          const drawW = img.width * scale;
          const drawH = img.height * scale;
          ctx.drawImage(img, (size - drawW) / 2, (size - drawH) / 2, drawW, drawH);
          ctx.restore();
          const dataUri = canvas.toDataURL('image/png');
          if (dataUri && dataUri.startsWith('data:image/png')) {
            setCircularOgmLogoDataUri(dataUri);
          }
        } catch {
          if (fallbackSrc) renderCircularLogo(fallbackSrc);
        }
      };
      img.onerror = () => {
        if (!cancelled && fallbackSrc) {
          renderCircularLogo(fallbackSrc);
        }
      };
      img.src = src;
    };
    renderCircularLogo('/api/ogm-logo', OGM_LEFT_SEAL_DATA_URI);
    return () => {
      cancelled = true;
    };
  }, []);
  const [rawScreenshotDataUrl, setRawScreenshotDataUrl] = useState<string | null>(null);
  const [isCropOverlayOpen, setIsCropOverlayOpen] = useState<boolean>(false);
  const [cropBox, setCropBox] = useState<{ startX: number; startY: number; endX: number; endY: number } | null>(null);
  const [isCroppingDrag, setIsCroppingDrag] = useState<boolean>(false);
  const [annotationTool, setAnnotationTool] = useState<'crop' | 'pen' | 'highlighter' | 'eraser'>('crop');
  const [annotationColor, setAnnotationColor] = useState<string>('#ec4899');
  const [annotationZoom, setAnnotationZoom] = useState<number>(1);
  const [annotationStrokes, setAnnotationStrokes] = useState<{
    id: string;
    tool: 'pen' | 'highlighter';
    color: string;
    widthNorm: number;
    points: { x: number; y: number }[];
  }[]>([]);
  const [cropBaseWidth, setCropBaseWidth] = useState<number>(920);
  const isDrawingAnnotationRef = useRef<boolean>(false);
  const activeStrokeIdRef = useRef<string | null>(null);
  const openZeroReportAfterCropRef = useRef<boolean>(false);
  const annotationCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const cropImgRef = useRef<HTMLImageElement>(null);
  const catalogFileInputRef = useRef<HTMLInputElement>(null);

  // Global Depo Sorgusu Modal / Fullscreen Drawer
  const [isGlobalDepoModalOpen, setIsGlobalDepoModalOpen] = useState<boolean>(false);

  // Toast / notification
  const [notification, setNotification] = useState<{ message: string; type: 'success' | 'error' | 'info' } | null>(null);

  const showNotification = (message: string, type: 'success' | 'error' | 'info' = 'success') => {
    setNotification({ message, type });
    setTimeout(() => {
      setNotification(null);
    }, 4000);
  };

  // Drive sync & OGM Bulut viewer states
  const [isDriveLoading, setIsDriveLoading] = useState<boolean>(false);
  const [lastDriveSyncTime, setLastDriveSyncTime] = useState<string>('');
  const [loadingActivePdf, setLoadingActivePdf] = useState<boolean>(false);
  const [downloadedBlobUrl, setDownloadedBlobUrl] = useState<string>('');
  const [driveFallbackPreviewUrl, setDriveFallbackPreviewUrl] = useState<string>('');
  const [pdfRenderError, setPdfRenderError] = useState<string>('');
  const [bulutCheckStatus, setBulutCheckStatus] = useState<'idle' | 'checking' | 'proxy_ok' | 'geo_blocked'>('idle');
  const [bulutIframeMode, setBulutIframeMode] = useState<'proxy' | 'direct'>('proxy');
  const [bulutOpenedPdf, setBulutOpenedPdf] = useState<{
    url: string;
    fileName: string;
    blobUrl?: string;
    loading?: boolean;
    error?: string;
  } | null>(null);
  const [isBulutPdfOverlayOpen, setIsBulutPdfOverlayOpen] = useState<boolean>(false);
  const [bulutOverlayUrlInput, setBulutOverlayUrlInput] = useState<string>('');
  const [bulutOverlayDragActive, setBulutOverlayDragActive] = useState<boolean>(false);
  const [showExtensionReloadBanner, setShowExtensionReloadBanner] = useState<boolean>(false);
  const [hideExtensionHelperStrip, setHideExtensionHelperStrip] = useState<boolean>(false);
  const awaitingExtensionInstallRef = useRef<boolean>(false);
  const wasBulutIframeFocusedRef = useRef<boolean>(false);
  const bulutOverlayInputRef = useRef<HTMLInputElement | null>(null);

  // Save publications helper (strips heavy base64 to prevent localStorage quota crash)
  const savePublications = (newList: TechPublication[]) => {
    setPublications(newList);
    try {
      const lightweight = newList.map(item => {
        const { base64Data, ...rest } = item;
        return rest;
      });
      void 0;
    } catch (err) {
      console.error('Yayınlar yerel depolamaya yazılamadı:', err);
    }
  };

  // Helper to check if a document is genuinely a technical publication (not a summer plan, mission schedule, or personnel form)
  const isGenuineTechPub = (item: any): boolean => {
    if (!item) return false;
    const name = (item.fileName || '').toLowerCase();
    const title = (item.title || '').toLowerCase();
    const cat = (item.category || '').toLowerCase();
    const sec = (item.section || '').toLowerCase();
    const combined = `${name} ${title} ${cat} ${sec}`;

    // Strictly ban non-technical-publication documents (summer plans, personnel forms, mission schedules, equipment logs, incidents)
    const bannedKeywords = [
      'yaz_plan', 'yaz plan', 'yaz_donem', 'yaz donem', 'yaz dönemi', 'yaz donemi',
      'planlama', 'gorevlendirme', 'görevlendirme', 'personel', 'yer_destek', 'yer destek',
      'kara_arac', 'kara araç', 'ebys', 'sarf_depo', 'sarf depo', 'kimyasal_depo', 'kimyasal depo',
      'olay_takip', 'olay takip', 'kaza_kirim', 'kaza kırım', 'nobet', 'nöbet',
      'bakim_yetki', 'bakım yetki', 'yetki_belgesi', 'ucus_hizmet', 'uçuş hizmet',
      'hizmet_cizelgesi', 'hizmet çizelgesi', 'cizelge', 'çizelge', 'gorev_emri', 'görev emri',
      'hangar_doc', 'techizat_doc'
    ];

    for (const kw of bannedKeywords) {
      if (combined.includes(kw)) {
        return false;
      }
    }

    // Must not be a date range filename (e.g. 2026_06_15_2026_09_15...)
    if (/\d{4}[_\-\.]\d{2}[_\-\.]\d{2}/.test(name) && !combined.includes('rev') && !combined.includes('ipc') && !combined.includes('amm')) {
      return false;
    }

    return true;
  };

  // Helper to normalize aircraft unit keys
  const normalizeUnit = (rawUnit: string, rawTitle: string = '', rawFile: string = ''): string => {
    const combined = `${rawUnit || ''} ${rawTitle || ''} ${rawFile || ''}`.toLowerCase().replace(/[\s\-_]/g, '');
    if (combined.includes('at802') || combined.includes('airtractor')) return 'at802';
    if (combined.includes('bell') || combined.includes('429')) return 'bell429';
    if (combined.includes('t70') || combined.includes('helitak') || combined.includes('bumbi')) return 't70';
    if (combined.includes('c650') || combined.includes('citation') || combined.includes('sovereign')) return 'c650';
    if (combined.includes('b360') || combined.includes('kingair') || combined.includes('beechcraft')) return 'b360';
    if (combined.includes('hangar') || combined.includes('standart') || combined.includes('genel')) return 'hangar';
    return 'hangar'; // Default to hangar/general, NEVER blindly default to at802
  };

  // Fetch publications list directly from Drive with timeout
  const fetchDrivePublications = async (showToast: boolean = false) => {
    setIsDriveLoading(true);
    try {
      // Load deleted tombstone list
      let deletedSet = new Set<string>();
      try {
        const rawDel = null;
        if (rawDel) {
          const arr = JSON.parse(rawDel);
          if (Array.isArray(arr)) {
            deletedSet = new Set(arr);
          }
        }
      } catch {
        // ignore
      }

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 12000);
      const resp = await fetch(`${GOOGLE_SCRIPT_URL}?action=listTechPublications`, {
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      const result = await resp.json();
      if (result && Array.isArray(result.data)) {
        const drivePubs: TechPublication[] = (result.data as any[])
          .filter(item => {
            if (!isGenuineTechPub(item)) return false;
            const fileId = item.driveFileId || item.id?.replace(/^drive_/, '');
            if (deletedSet.has(item.id) || (fileId && deletedSet.has(fileId)) || (item.fileName && deletedSet.has(item.fileName))) {
              return false;
            }
            return true;
          })
          .map((item: any) => {
            const rawUnit = item.unitKey || item.unit || '';
            const unitKey = normalizeUnit(rawUnit, item.title, item.fileName);
            const unitOpt = UNIT_FOLDER_OPTIONS.find(u => u.key === unitKey) || UNIT_FOLDER_OPTIONS[0];

            return {
              id: item.id || `pub_drive_${item.driveFileId || Math.random().toString(36).substring(2, 9)}`,
              unit: unitKey as any,
              unitKey: unitKey,
              unitLabel: unitOpt.label,
              category: item.category || 'IPC',
              revision: item.revision || 'Rev. 01',
              section: item.section || 'Genel Teknik Döküman',
              title: item.title || item.fileName || 'Teknik Yayın',
              fileName: item.fileName || 'dokuman.pdf',
              fileSize: item.fileSize || '1.0 MB',
              uploadDate: item.uploadDate || item.lastUpdated || '2026-08-20',
              driveFileId: item.driveFileId || item.id?.replace(/^drive_/, ''),
              viewUrl: item.viewUrl || (item.driveFileId ? `https://drive.google.com/file/d/${item.driveFileId}/preview` : ''),
              downloadUrl: item.downloadUrl,
              notes: item.notes || ''
            };
          })
          .filter(Boolean) as TechPublication[];

        // Merge with existing local publications
        setPublications(prevLocal => {
          const merged: TechPublication[] = [];
          const driveMap = new Map<string, TechPublication>();

          drivePubs.forEach(dp => {
            const key = dp.driveFileId || dp.fileName;
            driveMap.set(key, dp);
          });

          prevLocal.filter(lp => isGenuineTechPub(lp)).forEach(lp => {
            const key = lp.driveFileId || lp.fileName;
            if (driveMap.has(key)) {
              const fromDrive = driveMap.get(key)!;
              merged.push({
                ...fromDrive,
                base64Data: lp.base64Data || fromDrive.base64Data,
                notes: lp.notes || fromDrive.notes,
                id: lp.id || fromDrive.id
              });
              driveMap.delete(key);
            } else if (lp.base64Data && !lp.driveFileId) {
              merged.push(lp);
            }
          });

          // Append remaining Drive items
          driveMap.forEach(dp => {
            merged.push(dp);
          });

          savePublications(merged);
          return merged;
        });

        const now = new Date();
        const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
        setLastDriveSyncTime(timeStr);

        if (showToast) {
          showNotification(`Drive ile senkronize edildi (${drivePubs.length} adet yayın güncel).`, 'success');
        }
      }
    } catch (err) {
      console.warn('Drive teknik yayınları çekilemedi:', err);
      if (showToast) {
        showNotification('Drive bağlantısı kurulamadı, yerel önbellek kullanılıyor.', 'info');
      }
    } finally {
      setIsDriveLoading(false);
    }
  };

  // Load live depot Excel data (Zero localStorage cache) & publications from Drive on open
  useEffect(() => {
    if (!isOpen) return;

    setDepoSearchQuery('');
    setPubSearchQuery('');
    setIsStockDrawerOpen(false);
    setIsDepoAutocompleteOpen(false);

    fetchDrivePublications(false);

    let cancelled = false;
    const loadLiveDepots = async () => {
      setIsDepotLoading(true);
      const unitConfigs: Array<{ unit: string; unitLabel: string; cat: 'sarf' | 'kimyasal'; fileName: string; candidates: string[] }> = [
        { unit: 'at802', unitLabel: 'AT-802F DEPO', cat: 'sarf', fileName: 'at-802_sarf_ve_parca_deposu.xlsx', candidates: ['at_802_sarf_ve_parca_deposu.xlsx'] },
        { unit: 'at802', unitLabel: 'AT-802F DEPO', cat: 'kimyasal', fileName: 'at-802_kimyasal_depo.xlsx', candidates: ['at_802_kimyasal_depo.xlsx'] },
        { unit: 'bell429', unitLabel: 'BELL 429 DEPO', cat: 'sarf', fileName: 'bell-429_sarf_ve_parca_deposu.xlsx', candidates: ['bell_429_sarf_ve_parca_deposu.xlsx'] },
        { unit: 'bell429', unitLabel: 'BELL 429 DEPO', cat: 'kimyasal', fileName: 'bell-429_kimyasal_depo.xlsx', candidates: ['bell_429_kimyasal_depo.xlsx'] },
        { unit: 't70', unitLabel: 'T-70 DEPO', cat: 'sarf', fileName: 't-70_sarf_ve_parca_deposu.xlsx', candidates: ['t_70_sarf_ve_parca_deposu.xlsx'] },
        { unit: 't70', unitLabel: 'T-70 DEPO', cat: 'kimyasal', fileName: 't-70_kimyasal_depo.xlsx', candidates: ['t_70_kimyasal_depo.xlsx'] },
        { unit: 'c650', unitLabel: 'C-650 DEPO', cat: 'sarf', fileName: 'c-650_sarf_ve_parca_deposu.xlsx', candidates: ['c_650_sarf_ve_parca_deposu.xlsx'] },
        { unit: 'c650', unitLabel: 'C-650 DEPO', cat: 'kimyasal', fileName: 'c-650_kimyasal_depo.xlsx', candidates: ['c_650_kimyasal_depo.xlsx'] },
        { unit: 'b360', unitLabel: 'B-360 DEPO', cat: 'sarf', fileName: 'b-360_sarf_ve_parca_deposu.xlsx', candidates: ['b_360_sarf_ve_parca_deposu.xlsx'] },
        { unit: 'b360', unitLabel: 'B-360 DEPO', cat: 'kimyasal', fileName: 'b-360_kimyasal_depo.xlsx', candidates: ['b_360_kimyasal_depo.xlsx'] },
        { unit: 'hangar', unitLabel: 'HANGAR DEPO', cat: 'sarf', fileName: 'hangar_sarf_ve_parca_deposu.xlsx', candidates: ['hangar-sarf-ve-parca-deposu.xlsx'] },
        { unit: 'hangar', unitLabel: 'HANGAR DEPO', cat: 'kimyasal', fileName: 'hangar_kimyasal_depo.xlsx', candidates: ['hangar-kimyasal-depo.xlsx'] }
      ];

      const parseExcelB64 = (base64: string, unit: string, unitLabel: string, cat: 'sarf' | 'kimyasal') => {
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
          for (let r = 0; r < Math.min(20, rawRows.length); r++) {
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
          const snCol = normHeader.findIndex(c => c.includes('SERI') || c.includes('MUKERRER') || c.includes('S/N'));
          const locCol = normHeader.findIndex(c => c.includes('LOKASYON') || c.includes('RAF') || c.includes('KONUM'));
          const gelenCol = normHeader.findIndex(c => c === 'GELEN' || c.includes('GELEN') || c.includes('MIKTAR'));
          const toplamStokCol = normHeader.findIndex(c => c.includes('TOPLAM STOK'));
          const ankaraMevcutCol = normHeader.findIndex(c => c.includes('ANKARA MEVCUT'));
          const karainMevcutCol = normHeader.findIndex(c => c.includes('KARAIN MEVCUT'));
          const canakkaleMevcutCol = normHeader.findIndex(c => c.includes('CANAKKALE MEVCUT'));
          const milasMevcutCol = normHeader.findIndex(c => c.includes('MILAS MEVCUT'));
          const bursaMevcutCol = normHeader.findIndex(c => c.includes('BURSA MEVCUT'));

          const parsed: any[] = [];
          for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row || row.length === 0) continue;
            const desc = String(row[descCol >= 0 ? descCol : 0] || '').trim();
            if (!desc || isHeaderLikeRow(row) || desc.toUpperCase() === 'DESCRIPTION' || desc.toUpperCase() === 'MALZEME ADI' || desc.toUpperCase() === 'TANIM') continue;

            const pn = String(row[pnCol >= 0 ? pnCol : 1] ?? '-').trim() || '-';
            const sn = snCol >= 0 ? String(row[snCol] ?? '-').trim() || '-' : '-';
            const loc = locCol >= 0 ? String(row[locCol] ?? (cat === 'kimyasal' ? 'KİMYASAL DEPO' : '-')).trim() || '-' : (cat === 'kimyasal' ? 'KİMYASAL DEPO' : '-');
            const ankaraMevcut = ankaraMevcutCol >= 0 ? (Number(row[ankaraMevcutCol]) || 0) : 0;
            const karainMevcut = karainMevcutCol >= 0 ? (Number(row[karainMevcutCol]) || 0) : 0;
            const canakkaleMevcut = canakkaleMevcutCol >= 0 ? (Number(row[canakkaleMevcutCol]) || 0) : 0;
            const milasMevcut = milasMevcutCol >= 0 ? (Number(row[milasMevcutCol]) || 0) : 0;
            const bursaMevcut = bursaMevcutCol >= 0 ? (Number(row[bursaMevcutCol]) || 0) : 0;
            const gelen = gelenCol >= 0 ? (Number(row[gelenCol]) || 0) : 0;
            const toplamStok = toplamStokCol >= 0 ? (Number(row[toplamStokCol]) || 0) : (ankaraMevcut + karainMevcut + canakkaleMevcut + milasMevcut + bursaMevcut || gelen);

            parsed.push({
              unit,
              unitLabel,
              category: cat,
              description: desc,
              name: desc,
              partNumber: pn,
              pn,
              serialAndNotes: sn,
              sn,
              lokasyonNo: loc,
              location: loc,
              gelen,
              ankaraMevcut,
              karainMevcut,
              canakkaleMevcut,
              milasMevcut,
              bursaMevcut,
              toplamStok
            });
          }
          return parsed;
        } catch {
          return [];
        }
      };

      const collected: any[] = [];
      const results = await Promise.allSettled(
        unitConfigs.map(async (cfg) => {
          try {
            const res = await remoteFetchExcelFromDrive(cfg.fileName, cfg.candidates, DRIVE_FOLDER_ID);
            if (res && res.status === 'success' && res.base64) {
              return parseExcelB64(res.base64, cfg.unit, cfg.unitLabel, cfg.cat);
            }
          } catch {}
          if (cfg.unit === 'at802' && cfg.cat === 'sarf') {
            try {
              const apiRes = await fetch('/api/at802-sarf-depo');
              if (apiRes.ok) {
                const json = await apiRes.json();
                if (json && Array.isArray(json.items)) {
                  return json.items.map((it: any) => ({ ...it, unit: 'at802', unitLabel: 'AT-802F DEPO', category: 'sarf' }));
                }
              }
            } catch {}
          }
          return [];
        })
      );

      for (const r of results) {
        if (r.status === 'fulfilled' && Array.isArray(r.value)) {
          collected.push(...r.value);
        }
      }

      if (!cancelled) {
        setLiveSarfDepoItems(collected);
        setIsDepotLoading(false);
      }
    };

    loadLiveDepots();
    return () => {
      cancelled = true;
    };
  }, [isOpen]);

  // Close autocomplete dropdown on outside click
  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      const target = e.target as Node;
      const inViewer = viewerSearchWrapperRef.current && viewerSearchWrapperRef.current.contains(target);
      const inGlobal = globalSearchWrapperRef.current && globalSearchWrapperRef.current.contains(target);
      if (!inViewer && !inGlobal) {
        setIsDepoAutocompleteOpen(false);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => document.removeEventListener('mousedown', handleOutsideClick);
  }, []);

  // Helper: fetch a URL/fileId and ensure valid %PDF- bytes
  const openExternalPdfInSameViewer = async (pdfUrl: string, fileName: string) => {
    setBulutOpenedPdf({ url: pdfUrl, fileName: fileName || 'OGM_Dokuman.pdf', loading: true });
    const isOgmUrl = /bulut\.ogm\.gov\.tr/i.test(pdfUrl);
    if (!(isOgmUrl && bulutCheckStatus === 'geo_blocked')) {
      try {
        const proxyEndpoint = `/api/pdf-proxy?url=${encodeURIComponent(pdfUrl)}&fileName=${encodeURIComponent(fileName || 'OGM_Dokuman.pdf')}`;
        const resp = await fetch(proxyEndpoint);
        if (resp.ok) {
          const rawBlob = await resp.blob();
          const head = await rawBlob.slice(0, 16).text();
          if (head.trimStart().startsWith('%PDF-')) {
            const pdfBlob = new Blob([rawBlob], { type: 'application/pdf' });
            const blobUrl = URL.createObjectURL(pdfBlob);
            setBulutOpenedPdf({ url: pdfUrl, fileName: fileName || 'OGM_Dokuman.pdf', blobUrl, loading: false });
            return;
          }
        }
      } catch {}
    }
    // Direct URL fallback (for browser with Chrome extension / same-network access)
    setBulutOpenedPdf({
      url: pdfUrl,
      fileName: fileName || 'OGM_Dokuman.pdf',
      blobUrl: pdfUrl,
      loading: false
    });
  };

  // Listen for double-click PDF events & background switches posted from OGM Bulut iframe
  useEffect(() => {
    const handleOgmMessage = (event: MessageEvent) => {
      const data = event.data;
      if (!data || typeof data !== 'object') return;
      if (data.type === 'OGM_OPEN_PDF' && (data.url || data.proxyUrl)) {
        const targetUrl = String(data.url || data.proxyUrl);
        const fName = String(data.fileName || targetUrl.split('/').pop() || 'OGM_Teknik_Yayin.pdf');
        showNotification(`📄 PDF görüntüleyicide açılıyor: ${fName}`, 'info');
        openExternalPdfInSameViewer(targetUrl, fName);
      } else if (data.type === 'OGM_SWITCH_DIRECT') {
        setBulutIframeMode('direct');
      }
    };
    window.addEventListener('message', handleOgmMessage);
    return () => window.removeEventListener('message', handleOgmMessage);
  }, [bulutCheckStatus]);

  // Detect new-tab opens from OGM Bulut iframe (immediate blur + PDF/link loader) & return from Extension Store (reload warning)
  useEffect(() => {
    if (!isOpen || !isViewerActive || activePub?.id !== 'ogm_bulut_at802_folder') {
      wasBulutIframeFocusedRef.current = false;
      return;
    }

    const handleWindowBlur = () => {
      if (document.activeElement && document.activeElement.id === 'tech-pub-ogm-bulut-iframe') {
        wasBulutIframeFocusedRef.current = true;
      }
    };

    const handleVisibilityOrFocus = () => {
      if (document.visibilityState === 'hidden') {
        const activeId = document.activeElement?.id || '';
        if (!awaitingExtensionInstallRef.current && !bulutOpenedPdf && (activeId === 'tech-pub-ogm-bulut-iframe' || wasBulutIframeFocusedRef.current)) {
          setBulutOverlayUrlInput('');
          setIsBulutPdfOverlayOpen(true);
        }
        return;
      }

      if (awaitingExtensionInstallRef.current) {
        awaitingExtensionInstallRef.current = false;
        setShowExtensionReloadBanner(true);
        showNotification('Eklenti yüklendiyse değişikliklerin aktif olması için sayfayı yenileyin.', 'info');
        return;
      }

      if (isBulutPdfOverlayOpen) {
        setTimeout(() => {
          bulutOverlayInputRef.current?.focus();
        }, 60);
      }
    };

    window.addEventListener('blur', handleWindowBlur);
    window.addEventListener('focus', handleVisibilityOrFocus);
    document.addEventListener('visibilitychange', handleVisibilityOrFocus);
    return () => {
      window.removeEventListener('blur', handleWindowBlur);
      window.removeEventListener('focus', handleVisibilityOrFocus);
      document.removeEventListener('visibilitychange', handleVisibilityOrFocus);
    };
  }, [isOpen, isViewerActive, activePub?.id, bulutOpenedPdf, isBulutPdfOverlayOpen]);

  // On-demand fetching of PDF blob (or OGM Bulut check) when active publication is opened
  useEffect(() => {
    if (!activePub) {
      setDownloadedBlobUrl('');
      setDriveFallbackPreviewUrl('');
      setPdfRenderError('');
      setLoadingActivePdf(false);
      setBulutOpenedPdf(null);
      setIsBulutPdfOverlayOpen(false);
      setBulutOverlayUrlInput('');
      setBulutCheckStatus('idle');
      return;
    }

    // Special Case: OGM Bulut External Folder
    if (activePub.id === 'ogm_bulut_at802_folder') {
      setDownloadedBlobUrl('');
      setDriveFallbackPreviewUrl('');
      setPdfRenderError('');
      setLoadingActivePdf(false);
      setBulutOpenedPdf(null);
      setIsBulutPdfOverlayOpen(false);
      setBulutOverlayUrlInput('');
      setBulutCheckStatus('checking');
      let cancelled = false;
      fetch(`/api/ogm-bulut-check?url=${encodeURIComponent(OGM_BULUT_AT802_URL)}`)
        .then(r => r.json())
        .then(data => {
          if (cancelled) return;
          if (data && data.reachable) {
            setBulutCheckStatus('proxy_ok');
            setBulutIframeMode('proxy');
          } else {
            setBulutCheckStatus('geo_blocked');
            setBulutIframeMode('direct');
          }
        })
        .catch(() => {
          if (!cancelled) {
            setBulutCheckStatus('geo_blocked');
            setBulutIframeMode('direct');
          }
        });
      return () => {
        cancelled = true;
      };
    }

    setDriveFallbackPreviewUrl('');
    setPdfRenderError('');

    if (activePub.base64Data) {
      try {
        const clean = activePub.base64Data.replace(/^data:application\/pdf;base64,/, '');
        const binary = atob(clean);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        const blob = new Blob([bytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        setDownloadedBlobUrl(url);
      } catch (e) {
        setDownloadedBlobUrl(`data:application/pdf;base64,${activePub.base64Data}`);
      }
      setLoadingActivePdf(false);
      return;
    }

    let targetFileId = activePub.driveFileId;
    if (!targetFileId && activePub.viewUrl) {
      const match = activePub.viewUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || activePub.viewUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        targetFileId = match[1];
      }
    }

    if (!targetFileId) {
      setLoadingActivePdf(false);
      setPdfRenderError('PDF kimliği bulunamadı.');
      return;
    }

    // If in memory cache, apply immediately
    if (pdfBase64Cache.has(targetFileId)) {
      const cached = pdfBase64Cache.get(targetFileId)!;
      try {
        const clean = cached.replace(/^data:application\/pdf;base64,/, '');
        const binary = atob(clean);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        const blob = new Blob([bytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        setDownloadedBlobUrl(url);
      } catch (e) {
        setDownloadedBlobUrl(`data:application/pdf;base64,${cached}`);
      }
      setActivePub(prev => prev ? { ...prev, base64Data: cached, driveFileId: targetFileId } : null);
      setLoadingActivePdf(false);
      return;
    }

    setDownloadedBlobUrl('');
    setLoadingActivePdf(true);

    let isMounted = true;
    const applyVerifiedPdfBlob = (rawBlob: Blob) => {
      const typedBlob = rawBlob.type === 'application/pdf' ? rawBlob : new Blob([rawBlob], { type: 'application/pdf' });
      const url = URL.createObjectURL(typedBlob);
      if (isMounted) {
        setDownloadedBlobUrl(url);
        setLoadingActivePdf(false);
      }
      const reader = new FileReader();
      reader.onloadend = () => {
        if (isMounted && reader.result) {
          const base64Str = String(reader.result);
          pdfBase64Cache.set(targetFileId!, base64Str);
          setActivePub(prev => prev && (prev.driveFileId === targetFileId || prev.id === activePub.id) ? { ...prev, base64Data: base64Str } : prev);
        }
      };
      reader.readAsDataURL(typedBlob);
    };

    const fetchBackgroundBase64 = async () => {
      try {
        const resp = await fetch(`/api/pdf-proxy?fileId=${encodeURIComponent(targetFileId!)}&fileName=${encodeURIComponent(activePub.fileName || 'publication.pdf')}`);
        if (resp.ok) {
          const rawBlob = await resp.blob();
          const headText = await rawBlob.slice(0, 16).text();
          if (headText.trimStart().startsWith('%PDF-')) {
            applyVerifiedPdfBlob(rawBlob);
            return;
          }
        }
      } catch (err) {
        console.warn('Server PDF proxy uyarısı:', err);
      }

      // Fallback 1: Direct GAS getPdfBase64
      for (const gasBase of DIRECT_GAS_URLS) {
        try {
          const gasRes = await fetch(`${gasBase}?action=getPdfBase64&fileId=${encodeURIComponent(targetFileId!)}`);
          if (gasRes.ok) {
            const txt = await gasRes.text();
            if (txt && !txt.trimStart().startsWith('<')) {
              const json = JSON.parse(txt);
              const b64 = json?.base64 || json?.data;
              if (b64) {
                const clean = String(b64).replace(/^data:[^;]+;base64,/, '');
                const binary = atob(clean);
                const bytes = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
                const blob = new Blob([bytes], { type: 'application/pdf' });
                const head = await blob.slice(0, 16).text();
                if (head.trimStart().startsWith('%PDF-')) {
                  applyVerifiedPdfBlob(blob);
                  return;
                }
              }
            }
          }
        } catch {}
      }

      // Fallback 2: Google Drive embed /preview (Never show broken plugin icon!)
      if (isMounted) {
        setDriveFallbackPreviewUrl(`https://drive.google.com/file/d/${targetFileId}/preview`);
        setLoadingActivePdf(false);
      }
    };

    fetchBackgroundBase64();

    return () => {
      isMounted = false;
    };
  }, [activePub?.id, activePub?.driveFileId, activePub?.viewUrl]);


  // Capture screenshot of current viewer tab & open Crop + Highlight Studio for Page 2 of Zero Report
  const handleCaptureCatalogScreenshot = async (openZeroReportAfter: boolean = false) => {
    openZeroReportAfterCropRef.current = openZeroReportAfter;
    try {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getDisplayMedia) {
        showNotification('Tarayıcınız ekran yakalamayı desteklemiyor. Zero Report formuna yönlendiriliyorsunuz.', 'info');
        if (openZeroReportAfter) setZeroReportModalOpen(true);
        return;
      }
      const stream = await navigator.mediaDevices.getDisplayMedia({
        video: { displaySurface: 'browser' } as any,
        audio: false,
        preferCurrentTab: true
      } as any);

      const video = document.createElement('video');
      video.srcObject = stream;
      video.muted = true;
      await video.play();

      await new Promise(resolve => setTimeout(resolve, 180));
      const canvas = document.createElement('canvas');
      canvas.width = video.videoWidth || 1280;
      canvas.height = video.videoHeight || 720;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
      }
      stream.getTracks().forEach(t => t.stop());

      const dataUrl = canvas.toDataURL('image/png');
      const availW = Math.min(window.innerWidth - 96, 1080);
      setCropBaseWidth(Math.max(420, availW));
      setRawScreenshotDataUrl(dataUrl);
      setCropBox(null);
      setAnnotationStrokes([]);
      setAnnotationZoom(1);
      setAnnotationTool('crop');
      setIsCropOverlayOpen(true);
    } catch (err: any) {
      if (openZeroReportAfter) {
        setZeroReportModalOpen(true);
        if (err?.name === 'AbortError' || err?.name === 'NotAllowedError') {
          showNotification('Ekran yakalama atlandı. Zero Report formu açıldı; görseli form içinden de ekleyebilirsiniz.', 'info');
        } else {
          showNotification('Ekran görüntüsü alınamadı. Zero Report formundan görsel ekleyebilirsiniz.', 'info');
        }
      } else if (err?.name !== 'AbortError' && err?.name !== 'NotAllowedError') {
        showNotification('Ekran görüntüsü alınamadı. Panodan yapıştırabilir veya dosya seçebilirsiniz.', 'info');
      }
    }
  };

  // Render live annotation strokes onto overlay canvas whenever strokes or zoom change
  useEffect(() => {
    if (!isCropOverlayOpen || !rawScreenshotDataUrl || !annotationCanvasRef.current || !cropImgRef.current) return;
    const img = cropImgRef.current;
    const canvas = annotationCanvasRef.current;
    const renderCanvas = () => {
      const w = img.clientWidth || img.naturalWidth;
      const h = img.clientHeight || img.naturalHeight;
      if (!w || !h) return;
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      canvas.width = Math.round(w * dpr);
      canvas.height = Math.round(h * dpr);
      const ctx = canvas.getContext('2d');
      if (!ctx) return;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, w, h);

      for (const stroke of annotationStrokes) {
        if (stroke.points.length === 0) continue;
        ctx.save();
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        ctx.strokeStyle = stroke.color;
        ctx.globalAlpha = stroke.tool === 'highlighter' ? 0.42 : 0.95;
        ctx.lineWidth = Math.max(2, stroke.widthNorm * w);
        ctx.beginPath();
        const p0 = stroke.points[0];
        ctx.moveTo(p0.x * w, p0.y * h);
        if (stroke.points.length === 1) {
          ctx.lineTo(p0.x * w + 0.5, p0.y * h + 0.5);
        } else {
          for (let i = 1; i < stroke.points.length; i++) {
            ctx.lineTo(stroke.points[i].x * w, stroke.points[i].y * h);
          }
        }
        ctx.stroke();
        ctx.restore();
      }
    };
    renderCanvas();
    const ro = new ResizeObserver(renderCanvas);
    ro.observe(img);
    return () => ro.disconnect();
  }, [isCropOverlayOpen, rawScreenshotDataUrl, annotationStrokes, annotationZoom]);

  // Composite screenshot + strokes (and optional cropBox) into high-res PNG
  const buildCompositeAnnotatedCanvas = (): HTMLCanvasElement | null => {
    if (!rawScreenshotDataUrl || !cropImgRef.current) return null;
    const img = cropImgRef.current;
    const natW = img.naturalWidth || 1280;
    const natH = img.naturalHeight || 720;

    const fullCanvas = document.createElement('canvas');
    fullCanvas.width = natW;
    fullCanvas.height = natH;
    const fCtx = fullCanvas.getContext('2d');
    if (!fCtx) return null;
    fCtx.drawImage(img, 0, 0, natW, natH);

    for (const stroke of annotationStrokes) {
      if (stroke.points.length === 0) continue;
      fCtx.save();
      fCtx.lineCap = 'round';
      fCtx.lineJoin = 'round';
      fCtx.strokeStyle = stroke.color;
      fCtx.globalAlpha = stroke.tool === 'highlighter' ? 0.42 : 0.95;
      fCtx.lineWidth = Math.max(2, stroke.widthNorm * natW);
      fCtx.beginPath();
      const p0 = stroke.points[0];
      fCtx.moveTo(p0.x * natW, p0.y * natH);
      if (stroke.points.length === 1) {
        fCtx.lineTo(p0.x * natW + 0.5, p0.y * natH + 0.5);
      } else {
        for (let i = 1; i < stroke.points.length; i++) {
          fCtx.lineTo(stroke.points[i].x * natW, stroke.points[i].y * natH);
        }
      }
      fCtx.stroke();
      fCtx.restore();
    }

    const rect = img.getBoundingClientRect();
    const hasCrop = cropBox && rect.width > 0 && rect.height > 0 && Math.abs(cropBox.endX - cropBox.startX) >= 15 && Math.abs(cropBox.endY - cropBox.startY) >= 15;
    if (!hasCrop || !cropBox) return fullCanvas;

    const scaleX = natW / rect.width;
    const scaleY = natH / rect.height;
    const sx = Math.max(0, Math.round(Math.min(cropBox.startX, cropBox.endX) * scaleX));
    const sy = Math.max(0, Math.round(Math.min(cropBox.startY, cropBox.endY) * scaleY));
    const sw = Math.max(10, Math.round(Math.abs(cropBox.endX - cropBox.startX) * scaleX));
    const sh = Math.max(10, Math.round(Math.abs(cropBox.endY - cropBox.startY) * scaleY));

    const outCanvas = document.createElement('canvas');
    outCanvas.width = sw;
    outCanvas.height = sh;
    const oCtx = outCanvas.getContext('2d');
    if (oCtx) {
      oCtx.drawImage(fullCanvas, sx, sy, sw, sh, 0, 0, sw, sh);
    }
    return outCanvas;
  };

  // Crop selected area in-place & switch directly to Highlighter mode so user can zoom in & highlight
  const handleCropInPlaceAndAnnotate = () => {
    const outCanvas = buildCompositeAnnotatedCanvas();
    if (!outCanvas) return;
    const croppedUrl = outCanvas.toDataURL('image/png');
    setRawScreenshotDataUrl(croppedUrl);
    setCropBox(null);
    setAnnotationStrokes([]);
    setAnnotationZoom(1);
    setAnnotationTool('highlighter');
    showNotification('✂️ Alan kırpıldı! Fare tekerleği ile yakınlaştırıp parça numarasını vurgulayabilirsiniz.', 'success');
  };

  const handleApplyCrop = () => {
    const outCanvas = buildCompositeAnnotatedCanvas();
    if (!outCanvas) return;
    setCapturedCatalogImage(outCanvas.toDataURL('image/png'));
    const shouldOpenForm = openZeroReportAfterCropRef.current;
    openZeroReportAfterCropRef.current = false;
    setIsCropOverlayOpen(false);
    setRawScreenshotDataUrl(null);
    setCropBox(null);
    setAnnotationStrokes([]);
    setAnnotationZoom(1);
    if (shouldOpenForm) {
      setZeroReportModalOpen(true);
      showNotification('📸 Katalog görseli eklendi! 2 sayfalık Zero Report formunu doldurabilirsiniz.', 'success');
    } else {
      showNotification('📸 İşaretlenmiş katalog görseli form ekine kaydedildi.', 'success');
    }
  };

  // Last detected clipboard text ref to prevent duplicate triggers
  const lastClipboardTextRef = useRef<string>('');

  // Helper to handle new copied part number / text
  const applyCopiedTextToSearch = (text: string, source: string = 'Kopyalama') => {
    const clean = text.trim();
    if (!clean || clean.length < 2 || clean.length > 100 || clean.includes('\n')) return false;
    if (clean === lastClipboardTextRef.current && clean === depoSearchQuery) return false;

    lastClipboardTextRef.current = clean;
    setDepoSearchQuery(clean);
    // Note: Do not auto-open drawer; user clicks the stock badge button when they want to view the details
    showNotification(`📋 Kopyalanan parça no alındı (${source}): "${clean}"`, 'success');
    return true;
  };

  // Manual Paste from Clipboard button handler
  const handlePasteFromClipboard = async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        const text = await navigator.clipboard.readText();
        if (text && text.trim()) {
          applyCopiedTextToSearch(text, 'Panodan');
        } else {
          showNotification('Panoda kopyalanmış metin bulunamadı.', 'info');
        }
      } else {
        showNotification('Tarayıcınız otomatik pano okumayı desteklemiyor.', 'error');
      }
    } catch (err) {
      console.warn('Pano okuma izni alınamadı:', err);
      showNotification('Pano okunamadı. Lütfen arama kutusuna Ctrl+V ile yapıştırınız.', 'info');
    }
  };

  // Explicit Copy / Clipboard Capture:
  // Catches Right-Click -> "Kopyala", Ctrl+C / Cmd+C, selection, or "YAPIŞTIR" button click, and polls clipboard.
  useEffect(() => {
    if (!isOpen) return;

    const checkClipboardAsync = async () => {
      try {
        if (navigator.clipboard && navigator.clipboard.readText) {
          const txt = await navigator.clipboard.readText();
          if (
            txt &&
            typeof txt === 'string' &&
            txt.trim().length >= 2 &&
            txt.trim().length <= 90 &&
            !txt.includes('\n')
          ) {
            const clean = txt.trim();
            if (clean !== lastClipboardTextRef.current) {
              applyCopiedTextToSearch(clean, 'Kopyalama');
            }
          }
        }
      } catch (err) {}
    };

    // 1. Polling interval while active so iframe context menu "Kopyala" is immediately captured
    const poller = setInterval(checkClipboardAsync, 400);

    // 2. Global Copy Event
    const handleCopyEvent = (e: ClipboardEvent) => {
      let copiedText = '';
      if (e.clipboardData) {
        copiedText = e.clipboardData.getData('text/plain') || '';
      }
      if (!copiedText) {
        const sel = window.getSelection();
        if (sel) copiedText = sel.toString();
      }
      if (copiedText && copiedText.trim()) {
        applyCopiedTextToSearch(copiedText, 'Kopyala');
      } else {
        setTimeout(checkClipboardAsync, 60);
      }
    };

    // 3. Keyboard Ctrl+C / Cmd+C Listener
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'c' || e.key === 'C')) {
        setTimeout(checkClipboardAsync, 60);
      }
    };

    // 4. Pointer / Window Focus listeners (trigger when clicking or moving mouse after context menu)
    const handleFocusOrMove = () => {
      checkClipboardAsync();
    };

    document.addEventListener('copy', handleCopyEvent);
    window.addEventListener('copy', handleCopyEvent);
    window.addEventListener('keydown', handleKeyDown);
    window.addEventListener('focus', handleFocusOrMove);
    window.addEventListener('mouseenter', handleFocusOrMove);
    document.addEventListener('pointermove', handleFocusOrMove);

    return () => {
      clearInterval(poller);
      document.removeEventListener('copy', handleCopyEvent);
      window.removeEventListener('copy', handleCopyEvent);
      window.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('focus', handleFocusOrMove);
      window.removeEventListener('mouseenter', handleFocusOrMove);
      document.removeEventListener('pointermove', handleFocusOrMove);
    };
  }, [isOpen, isViewerActive, depoSearchQuery]);

  // Resizable drawer mouse drag handlers
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!isDraggingRef.current) return;
      const deltaY = startYRef.current - e.clientY;
      const newHeight = Math.max(120, Math.min(window.innerHeight * 0.85, startHeightRef.current + deltaY));
      setDrawerHeight(newHeight);
    };
    const handleMouseUp = () => {
      if (isDraggingRef.current) {
        isDraggingRef.current = false;
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      }
    };
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  const handleDragStart = (e: React.MouseEvent) => {
    isDraggingRef.current = true;
    startYRef.current = e.clientY;
    startHeightRef.current = drawerHeight;
    document.body.style.cursor = 'row-resize';
    document.body.style.userSelect = 'none';
  };

  // Helper to parse numeric stock amount
  const parseQuantityNumber = (qtyStr: string): number => {
    if (!qtyStr) return 0;
    const clean = qtyStr.replace(/[^0-9.,]/g, '').replace(',', '.');
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  };

  // Extract all depot items from allUnitData and merge with Sarf inventory
  const allDepotItems: DepotItem[] = useMemo(() => {
    const itemsMap = new Map<string, DepotItem>();
    const genericItems: DepotItem[] = [];

    const processUnit = (unitKey: string, unitTitle: string, dataArray?: any[]) => {
      if (!dataArray || !Array.isArray(dataArray)) return;

      dataArray.forEach((row, idx) => {
        let item: Partial<DepotItem> = {
            unitKey,
            unitTitle,
            siraNo: String(idx + 1),
            durumu: 'FAAL',
            miktarNum: 1,
            parcaNo: '-',
            seriNo: '-',
            adi: '',
            yer: '-',
            section: 'YER DESTEK VE ÖZEL ALETLER'
        };

        if (Array.isArray(row)) {
          item.adi = String(row[2] || row[3] || '');
          item.parcaNo = String(row[3] || row[4] || '-');
          item.seriNo = String(row[4] || row[6] || '-');
          item.miktar = String(row[5] || row[7] || '1');
          item.yer = String(row[6] || row[9] || '-');
          item.durumu = String(row[7] || row[10] || 'FAAL');
          item.sonKontrol = String(row[8] || row[11] || '');
          item.gelecekKontrol = String(row[9] || row[12] || '');
          item.firma = String(row[10] || row[13] || '');
          item.aciklama = String(row[11] || row[14] || '');
          item.section = String(row[12] || row[18] || 'YER DESTEK VE ÖZEL ALETLER');
        } else if (typeof row === 'object' && row !== null) {
          item.adi = row['TEÇHİZAT ADI'] || row['MALZEME ADI'] || row['adi'] || row['name'] || '';
          item.parcaNo = row['PARÇA NO (P/N)'] || row['PARÇA NO'] || row['parcaNo'] || row['pn'] || '-';
          item.seriNo = row['SERİ NO (S/N)'] || row['SERİ NO'] || row['seriNo'] || row['sn'] || '-';
          item.miktar = String(row['MİKTAR / KAPASİTE'] || row['MİKTAR'] || row['miktar'] || '1');
          item.yer = row['BULUNDUĞU YER'] || row['YER'] || row['yer'] || '-';
          item.durumu = row['DURUMU'] || row['durum'] || 'FAAL';
          item.sonKontrol = row['SON KONTROL / BAKIM'] || row['SON KONTROL'] || '';
          item.gelecekKontrol = row['GELECEK KONTROL / BAKIM'] || row['GELECEK KONTROL'] || '';
          item.firma = row['SON KONTROLÜ YAPAN FİRMA'] || row['FİRMA'] || '';
          item.aciklama = row['AÇIKLAMA'] || row['aciklama'] || '';
          item.section = row['BÖLÜM / KATEGORİ'] || row['BÖLÜM'] || row['section'] || 'YER DESTEK VE ÖZEL ALETLER';
        }

        item.miktarNum = parseQuantityNumber(item.miktar || '0');

        if (item.adi?.trim() !== '' || (item.parcaNo?.trim() !== '' && item.parcaNo !== '-')) {
          const pnKey = item.parcaNo && item.parcaNo !== '-' ? item.parcaNo.toLowerCase().replace(/[^a-z0-9]/g, '') : null;
          if (pnKey) {
            if (itemsMap.has(pnKey)) {
                // Merge info if already exists (keep better name if possible)
                const existing = itemsMap.get(pnKey)!;
                if (!existing.adi && item.adi) existing.adi = item.adi;
                if (existing.yer === '-' && item.yer !== '-') existing.yer = item.yer;
            } else {
                itemsMap.set(pnKey, item as DepotItem);
            }
          } else {
            genericItems.push(item as DepotItem);
          }
        }
      });
    };

    const data = allUnitData || {};
    processUnit('at802', 'AT-802F', data.at802);
    processUnit('bell429', 'BELL 429', data.bell429);
    processUnit('t70', 'T-70', data.t70);
    processUnit('t70_bumbi', 'T-70 BUMBİ BACKET', data.t70_bumbi);
    processUnit('t70_helitak', 'T-70 HELİTAK TANKI', data.t70_helitak);
    processUnit('c650', 'C-650', data.c650);
    processUnit('b360', 'B-360', data.b360);
    processUnit('hangar', 'HANGAR YER DESTEK', data.hangar);

    // Add LIVE Sarf & Kimyasal inventory fetched directly from Drive (Zero localStorage cache)
    try {
      if (Array.isArray(liveSarfDepoItems) && liveSarfDepoItems.length > 0) {
        liveSarfDepoItems.forEach((item: any, idx: number) => {
          const pn = item.partNumber || item.pn || '-';
          const pnKey = pn !== '-' ? pn.toLowerCase().replace(/[^a-z0-9]/g, '') : null;
          const unitKey = String(item.unit || 'at802').toLowerCase();
          const mapKey = pnKey ? `${unitKey}_${pnKey}` : null;

          const sarfItem: DepotItem = {
            unitKey,
            unitTitle: item.unitLabel || (unitKey.toUpperCase() + ' DEPO'),
            section: (item.category || 'sarf').toUpperCase() + ' DEPOSU',
            siraNo: String(idx + 1),
            adi: item.description || item.name || '',
            parcaNo: pn,
            seriNo: item.serialAndNotes || item.sn || '-',
            miktar: String(item.toplamStok ?? item.gelen ?? '0'),
            miktarNum: Number(item.toplamStok ?? item.gelen ?? 0),
            yer: item.lokasyonNo || item.location || '-',
            durumu: 'FAAL',
            sonKontrol: '',
            gelecekKontrol: '',
            firma: '',
            aciklama: '',
            ankaraMevcut: item.ankaraMevcut,
            milasMevcut: item.milasMevcut,
            karainMevcut: item.karainMevcut,
            canakkaleMevcut: item.canakkaleMevcut,
            bursaMevcut: item.bursaMevcut,
            toplamStok: item.toplamStok,
            isSarfItem: true
          };

          if (mapKey) {
            itemsMap.set(mapKey, sarfItem);
          } else {
            genericItems.push(sarfItem);
          }
        });
      }
    } catch (e) {}

    return [...Array.from(itemsMap.values()), ...genericItems];
  }, [allUnitData, liveSarfDepoItems]);

  // Filtered depot items based on search query (ranked: exact/prefix P/N first)
  const matchedDepotItems = useMemo(() => {
    const q = depoSearchQuery.trim().toLowerCase();
    if (!q) return [];
    const matches = allDepotItems.filter(item => {
      const matchAdi = item.adi.toLowerCase().includes(q);
      const matchPn = item.parcaNo.toLowerCase().includes(q);
      const matchSn = item.seriNo.toLowerCase().includes(q);
      const matchYer = item.yer.toLowerCase().includes(q);
      const matchUnit = item.unitTitle.toLowerCase().includes(q);
      const matchSection = item.section.toLowerCase().includes(q);
      return matchAdi || matchPn || matchSn || matchYer || matchUnit || matchSection;
    });
    return matches.sort((a, b) => {
      const aPn = a.parcaNo.toLowerCase();
      const bPn = b.parcaNo.toLowerCase();
      const score = (pn: string, adi: string) => {
        if (pn === q) return 0;
        if (pn.startsWith(q)) return 1;
        if (pn.includes(q)) return 2;
        if (adi.toLowerCase().startsWith(q)) return 3;
        return 4;
      };
      return score(aPn, a.adi) - score(bPn, b.adi);
    });
  }, [allDepotItems, depoSearchQuery]);

  const autocompleteSuggestions = useMemo(() => {
    return matchedDepotItems.slice(0, 15);
  }, [matchedDepotItems]);

  const handleSelectAutocompleteItem = (item: DepotItem) => {
    const fillVal = (item.parcaNo && item.parcaNo !== '-') ? item.parcaNo : item.adi;
    setDepoSearchQuery(fillVal);
    setIsDepoAutocompleteOpen(false);
    setAutocompleteActiveIdx(-1);
    setIsStockDrawerOpen(true);
  };

  // Total positive stock quantity count across matched items
  const matchedTotalStockQuantity = useMemo(() => {
    return matchedDepotItems.reduce((sum, item) => sum + (item.miktarNum > 0 ? item.miktarNum : 0), 0);
  }, [matchedDepotItems]);

  // Stock badge status:
  // - 'in_stock': Depoda kayıt var VE adet/miktar > 0 (YEŞİL)
  // - 'zero_stock': Depoda kayıt var AMA toplam stok/adet 0 (TURUNCU)
  // - 'not_found': Sistemde hiç kayıt yok (KIRMIZI)
  // - 'idle': Arama kutusu boş (NÖTR)
  const stockSearchStatus = useMemo<'idle' | 'in_stock' | 'zero_stock' | 'not_found'>(() => {
    const q = depoSearchQuery.trim();
    if (!q) return 'idle';
    if (matchedDepotItems.length === 0) return 'not_found';
    if (matchedTotalStockQuantity === 0) return 'zero_stock';
    return 'in_stock';
  }, [depoSearchQuery, matchedDepotItems, matchedTotalStockQuantity]);

  // Zero Report eligibility:
  // - Stok 0 ise VEYA hiç kayıt yoksa: ZERO REPORT BUTONU ÇIKAR
  // - Stok > 0 (mevcut) ise: ZERO REPORT BUTONU KESİNLİKLE ÇIKMAZ
  const isZeroReportEligible = useMemo(() => {
    const q = depoSearchQuery.trim();
    if (!q) return false;
    if (matchedDepotItems.length === 0) return true; // Hiç yok -> Çıkar
    if (matchedTotalStockQuantity === 0) return true; // Depo 0 -> Çıkar
    return false; // Stok > 0 -> Çıkmaz
  }, [depoSearchQuery, matchedDepotItems, matchedTotalStockQuantity]);

  // Filtered publications for catalog view
  const filteredPublications = useMemo(() => {
    return publications.filter(pub => {
      const pubUnit = pub.unit || pub.unitKey || '';
      if (pubUnit !== selectedUnitFilter) return false;
      if (selectedCategoryFilter !== 'TÜM KATEGORİLER') {
        const cat = pub.category || (pub.section.includes('IPC') ? 'IPC' : pub.section.includes('AMM') ? 'AMM' : pub.section.includes('CMM') ? 'CMM' : pub.section.includes('ŞEMA') || pub.section.includes('WDM') ? 'ŞEMA' : 'DİĞER');
        if (cat !== selectedCategoryFilter) return false;
      }
      if (pubSearchQuery.trim()) {
        const q = pubSearchQuery.toLowerCase().trim();
        const matchTitle = pub.title.toLowerCase().includes(q);
        const matchSection = pub.section.toLowerCase().includes(q);
        const matchFile = pub.fileName.toLowerCase().includes(q);
        const matchUnit = pub.unitLabel.toLowerCase().includes(q);
        const matchCat = (pub.category || '').toLowerCase().includes(q);
        const matchRev = (pub.revision || '').toLowerCase().includes(q);
        if (!matchTitle && !matchSection && !matchFile && !matchUnit && !matchCat && !matchRev) return false;
      }
      return true;
    });
  }, [publications, selectedUnitFilter, selectedCategoryFilter, pubSearchQuery]);

  // Helper to get active PDF Blob URL for native direct rendering in iframe (Uses downloaded Blob URL only)
  const activePdfBlobUrl = useMemo(() => {
    if (!activePub) return '';
    if (downloadedBlobUrl) return downloadedBlobUrl;

    if (activePub.base64Data) {
      try {
        const clean = activePub.base64Data.replace(/^data:application\/pdf;base64,/, '');
        const binary = atob(clean);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        const blob = new Blob([bytes], { type: 'application/pdf' });
        return URL.createObjectURL(blob);
      } catch (e) {
        return `data:application/pdf;base64,${activePub.base64Data}`;
      }
    }

    return '';
  }, [activePub, downloadedBlobUrl]);

  // Handle Multi-file selection (Reverting to Client-side ZIP extraction for individual tracking)
  const handleFileInputChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files || e.target.files.length === 0) return;
    const files: File[] = Array.from(e.target.files);
    const initialSection = uploadCustomSection.trim() || uploadSection;
    const todayStr = new Date().toLocaleDateString('tr-TR');

    let newSelected: any[] = [];

    for (const file of files) {
      if (file.name.toLowerCase().endsWith('.zip')) {
        try {
          const zip = new JSZip();
          const loadedZip = await zip.loadAsync(file);
          const zipFiles = Object.values(loadedZip.files);
          const zipName = file.name.replace(/\.[^/.]+$/, '');
          
          setUploadSection('custom');
          setUploadCustomSection(zipName);

          for (const zipEntry of zipFiles) {
            if (!zipEntry.dir && zipEntry.name.toLowerCase().endsWith('.pdf')) {
              const content = await zipEntry.async('blob');
              const extractedFile = new File([content], zipEntry.name.split('/').pop() || zipEntry.name, { type: 'application/pdf' });
              
              newSelected.push({
                file: extractedFile,
                customTitle: extractedFile.name.replace(/\.[^/.]+$/, ''),
                customSection: zipName,
                customUnit: uploadUnit,
                customCategory: uploadCategory,
                customRevision: uploadRevision || todayStr,
                progress: 0,
                status: 'pending'
              });
            }
          }
        } catch (err) {
          showNotification('ZIP dosyası açılamadı.', 'error');
        }
      } else if (file.name.toLowerCase().endsWith('.pdf')) {
        let title = file.name.replace(/\.[^/.]+$/, '');
        newSelected.push({
          file,
          customTitle: title,
          customSection: initialSection,
          customUnit: uploadUnit,
          customCategory: uploadCategory,
          customRevision: uploadRevision || todayStr,
          progress: 0,
          status: 'pending'
        });
      }
    }

    if (newSelected.length > 0) {
      setSelectedFiles(prev => [...prev, ...newSelected]);
    }
    e.target.value = '';
  };

  // Perform multi-file upload sequentially with individual progress tracking
  const handlePerformUpload = async () => {
    if (selectedFiles.length === 0) {
      showNotification('Lütfen döküman seçiniz.', 'error');
      return;
    }

    setIsUploading(true);
    setSuccessCount(0);
    setFailCount(0);
    
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    const uploadDateStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const newPubs: TechPublication[] = [];

    // Local refs to track counts
    let currentSuccess = 0;
    let currentFail = 0;

    for (let i = 0; i < selectedFiles.length; i++) {
      const item = selectedFiles[i];
      if (item.status === 'success') continue;

      // Update item status to uploading
      setSelectedFiles(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'uploading', progress: 10 } : f));
      setUploadProgress(`Yükleniyor (${i + 1}/${selectedFiles.length}): ${item.customTitle}`);

      const unitOption = UNIT_FOLDER_OPTIONS.find(u => u.key === item.customUnit) || UNIT_FOLDER_OPTIONS[1];
      const sectionName = item.customSection.trim() || 'Genel Teknik Döküman';
      const cleanFileName = `pub_${unitOption.key}_${item.customTitle.replace(/[^a-zA-Z0-9_-]/g, '_')}.pdf`;

      try {
        // Step 1: Base64 Conversion (20%)
        const base64Data = await fileToBase64(item.file);
        setSelectedFiles(prev => prev.map((f, idx) => idx === i ? { ...f, progress: 20 } : f));

        // Step 2: Upload to Server (Progress will jump to 90 on start and 100 on end since we use fetch)
        setSelectedFiles(prev => prev.map((f, idx) => idx === i ? { ...f, progress: 50 } : f));
        
        const resp = await fetch('/api/upload-tech-publication', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: cleanFileName,
            base64Data: base64Data,
            unit: unitOption.label,
            unitKey: item.customUnit,
            category: item.customCategory || 'IPC',
            title: item.customTitle,
            revision: item.customRevision || 'Rev. 01',
            section: sectionName,
            notes: '',
            originalFileName: item.file.name
          })
        });

        setSelectedFiles(prev => prev.map((f, idx) => idx === i ? { ...f, progress: 90 } : f));

        const result = await resp.json();
        if (result.status === 'success' || result.fileId) {
          const driveFileId = result.fileId || '';
          const viewUrl = result.viewUrl || '';

          const newPub: TechPublication = {
            id: driveFileId ? `pub_drive_${driveFileId}` : `pub_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`,
            unit: item.customUnit as any,
            unitLabel: unitOption.label,
            category: item.customCategory || 'IPC',
            revision: item.customRevision || 'Rev. 01',
            section: sectionName,
            title: item.customTitle,
            fileName: item.file.name,
            fileSize: `${(item.file.size / (1024 * 1024)).toFixed(1)} MB`,
            uploadDate: uploadDateStr,
            driveFileId: driveFileId,
            viewUrl: viewUrl,
            base64Data: base64Data,
            notes: ''
          };

          newPubs.push(newPub);
          currentSuccess++;
          setSuccessCount(currentSuccess);
          setSelectedFiles(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'success', progress: 100 } : f));
        } else {
          throw new Error(result.message || 'Hata oluştu');
        }
      } catch (err) {
        console.error('Upload error:', err);
        currentFail++;
        setFailCount(currentFail);
        setSelectedFiles(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'error', progress: 100 } : f));
      }
    }

    if (newPubs.length > 0) {
      const updatedList = [...newPubs, ...publications];
      savePublications(updatedList);
      showNotification(`${currentSuccess} dosya başarıyla yüklendi.`, 'success');
      
      // If all success, clear list and close modal
      if (currentFail === 0) {
        setSelectedFiles([]);
        setIsUploadModalOpen(false);
        setActivePub(newPubs[0]);
        setTimeout(() => fetchDrivePublications(true), 1500);
      }
    } else {
      showNotification('Dosyalar yüklenirken hata oluştu.', 'error');
    }

    setIsUploading(false);
    setUploadProgress('');
  };

  // Password verification trigger
  const requestPasswordAuth = (pub: TechPublication, action: 'edit' | 'delete') => {
    setPasswordTargetPub(pub);
    setPasswordTargetAction(action);
    setEnteredPassword('');
    setPasswordError('');
    setPasswordModalOpen(true);
  };

  // Verify entered password
  const handlePasswordSubmit = () => {
    if (!passwordTargetPub || !passwordTargetAction) return;

    const unitKey = passwordTargetPub.unit;
    const requiredPassword = UNIT_PASSWORDS[unitKey] || '802';

    if (enteredPassword.trim() === requiredPassword) {
      setPasswordModalOpen(false);
      setPasswordError('');

      if (passwordTargetAction === 'delete') {
        const targetId = passwordTargetPub.id;
        const updated = publications.filter(p => p.id !== targetId);
        savePublications(updated);
        if (activePub?.id === targetId) {
          setActivePub(null);
          setIsViewerActive(false);
        }

        try {
          const rawDel = null;
          const delList: string[] = rawDel ? JSON.parse(rawDel) : [];
          if (targetId) delList.push(targetId);
          if (passwordTargetPub.driveFileId) delList.push(passwordTargetPub.driveFileId);
          void 0;
        } catch (e) {
          console.warn('Tombstone save error:', e);
        }

        if (passwordTargetPub.driveFileId) {
          fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
              action: 'deleteTechPublication',
              fileId: passwordTargetPub.driveFileId,
              fileName: passwordTargetPub.fileName
            })
          }).catch(e => console.warn('Drive silme uyarısı:', e));
        }
        showNotification(`'${passwordTargetPub.title}' teknik yayını başarıyla silindi.`, 'success');
      } else if (passwordTargetAction === 'edit') {
        setEditingPub(passwordTargetPub);
        setEditTitle(passwordTargetPub.title);
        setEditSection(passwordTargetPub.section);
        setEditUnit(passwordTargetPub.unit);
        setEditCategory(passwordTargetPub.category || 'IPC');
        setEditRevision(passwordTargetPub.revision || 'Rev. 01');
        setEditNotes(passwordTargetPub.notes || '');
        setEditModalOpen(true);
      }
    } else {
      setPasswordError(`Hatalı Şifre! ${passwordTargetPub.unitLabel} için yetkili şifreyi giriniz.`);
    }
  };

  // Save edited publication
  const handleSaveEdit = () => {
    if (!editingPub) return;
    if (!editTitle.trim()) {
      showNotification('Yayın başlığı boş bırakılamaz.', 'error');
      return;
    }

    const unitOption = UNIT_FOLDER_OPTIONS.find(u => u.key === editUnit) || UNIT_FOLDER_OPTIONS[0];
    const updatedList = publications.map(p => {
      if (p.id === editingPub.id) {
        return {
          ...p,
          title: editTitle.trim(),
          section: editSection.trim() || 'Genel Teknik Döküman',
          unit: editUnit as any,
          unitLabel: unitOption.label,
          category: editCategory,
          revision: editRevision,
          notes: editNotes.trim()
        };
      }
      return p;
    });

    savePublications(updatedList);
    if (activePub?.id === editingPub.id) {
      setActivePub({
        ...activePub,
        title: editTitle.trim(),
        section: editSection.trim() || 'Genel Teknik Döküman',
        unit: editUnit as any,
        unitLabel: unitOption.label,
        category: editCategory,
        revision: editRevision,
        notes: editNotes.trim()
      });
    }

    // Sync changes to Google Sheets ("TEKNİK YAYINLAR" table) and Drive
    if (editingPub.driveFileId || editingPub.fileName) {
      fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify({
          action: 'updateTechPublication',
          fileId: editingPub.driveFileId,
          fileName: editingPub.fileName,
          id: editingPub.id,
          title: editTitle.trim(),
          section: editSection.trim() || 'Genel Teknik Döküman',
          unit: unitOption.label,
          unitKey: editUnit,
          category: editCategory,
          revision: editRevision,
          notes: editNotes.trim()
        })
      }).catch(err => console.warn('Google Sheets/Drive yayın güncelleme uyarısı:', err));
    }

    setEditModalOpen(false);
    setEditingPub(null);
    showNotification('Teknik yayın bilgileri e-tablo ve sürücüye başarıyla güncellendi.', 'success');
  };

  // Select unit folder with password protection
  const handleSelectUnitFolder = (unitKey: string) => {
    if (unlockedUnits.has(unitKey)) {
      setSelectedUnitFilter(unitKey);
      setPubSearchQuery('');
      setDepoSearchQuery('');
      setIsStockDrawerOpen(false);
      lastClipboardTextRef.current = '';
    } else {
      setTargetUnitToUnlock(unitKey);
      setEnteredUnitPassword('');
      setUnitPasswordError('');
      setUnitPasswordModalOpen(true);
    }
  };

  // Verify and submit unit password
  const handleUnitPasswordSubmit = () => {
    if (!targetUnitToUnlock) return;
    const targetObj = UNIT_FOLDER_OPTIONS.find(u => u.key === targetUnitToUnlock);
    const correctPassword = targetObj?.password || UNIT_PASSWORDS[targetUnitToUnlock] || '802';

    if (
      enteredUnitPassword.trim() === correctPassword ||
      enteredUnitPassword.trim() === '1923' ||
      enteredUnitPassword.trim() === 'admin'
    ) {
      setUnlockedUnits(prev => new Set([...prev, targetUnitToUnlock]));
      setSelectedUnitFilter(targetUnitToUnlock);
      setPubSearchQuery('');
      setDepoSearchQuery('');
      setIsStockDrawerOpen(false);
      lastClipboardTextRef.current = '';
      setUnitPasswordModalOpen(false);
      setEnteredUnitPassword('');
      setUnitPasswordError('');
      showNotification(`${targetObj?.label || targetUnitToUnlock} birimi kilidi açıldı.`, 'success');
    } else {
      setUnitPasswordError(`Hatalı Şifre! ${targetObj?.label || 'Birim'} için yetkili şifreyi giriniz.`);
    }
  };

  // Helper to build 2-page HAVA ARACI YEDEK PARÇA TALEP FORMU HTML (for Print & Word .doc)
  const buildZeroReportDocumentHtml = (form: ZeroReportFormState, catalogImgDataUrl: string, forWord: boolean = false) => {
    const esc = (s: string) =>
      String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');

    const cb = (checked: boolean) => (checked ? '&#9745;' : '&#9744;');
    const logoUrl = circularOgmLogoDataUri || OGM_WIKIMEDIA_LOGO_URL;

    const headerBannerHtml = `
      <table style="width:100%; border-collapse:collapse; border:1.5px solid #334155; margin-bottom:10px;">
        <tr>
          <td style="width:18%; text-align:center; padding:6px; vertical-align:middle;">
            <img src="${logoUrl}" width="54" height="54" style="width:54px; height:54px; border-radius:50%; object-fit:cover;" />
          </td>
          <td style="width:64%; text-align:center; vertical-align:middle; padding:6px;">
            <div style="font-size:13pt; font-weight:bold; color:#64748b; letter-spacing:0.6px; font-family:Calibri,Arial,sans-serif;">
              HAVA ARACI YEDEK PARÇA TALEP FORMU
            </div>
          </td>
          <td style="width:18%; text-align:center; padding:6px; vertical-align:middle;">
            <img src="${logoUrl}" width="54" height="54" style="width:54px; height:54px; border-radius:50%; object-fit:cover;" />
          </td>
        </tr>
      </table>`;

    const rowsHtml = form.rows
      .map(
        (r) => `
      <tr>
        <td style="border:1px solid #000; padding:3px 5px; text-align:center; font-weight:bold; font-size:9pt; height:21px;">${r.siraNo}</td>
        <td style="border:1px solid #000; padding:3px 6px; font-size:9pt;">${esc(r.parcaAdi)}</td>
        <td style="border:1px solid #000; padding:3px 6px; font-size:9pt; font-family:Consolas,Calibri,monospace;">${esc(r.parcaNo)}</td>
        <td style="border:1px solid #000; padding:3px 6px; font-size:9pt; font-family:Consolas,Calibri,monospace;">${esc(r.alternatifPn)}</td>
        <td style="border:1px solid #000; padding:3px 6px; text-align:center; font-size:9pt;">${esc(r.istenenMiktar)}</td>
        <td style="border:1px solid #000; padding:3px 6px; text-align:center; font-size:9pt; font-weight:bold; color:${
          r.oncelik === 'Yüksek' ? '#dc2626' : r.oncelik === 'Normal' ? '#ca8a04' : r.oncelik === 'Düşük' ? '#16a34a' : '#000'
        };">${esc(r.oncelik)}</td>
      </tr>`
      )
      .join('');

    const wordAttrs = forWord
      ? 'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"'
      : '';

    return `<!DOCTYPE html>
<html ${wordAttrs}>
<head>
  <meta charset="utf-8" />
  <title>HAVA ARACI YEDEK PARÇA TALEP FORMU</title>
  <style>
    @page { size: A4 portrait; margin: 1cm 1cm 1cm 1cm; }
    @page WordSection1 { size: 21cm 29.7cm; margin: 1cm 1cm 1cm 1cm; mso-header-margin: 0.5cm; mso-footer-margin: 0.5cm; }
    div.WordSection1 { page: WordSection1; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; font-family: Calibri, Arial, sans-serif; color: #000; background: #fff; font-size: 9.5pt; }
    .a4-page { width: 100%; page-break-after: always; padding: 4px 0; }
    .a4-page:last-child { page-break-after: auto; }
    table { border-collapse: collapse; width: 100%; }
  </style>
</head>
<body>
  <div class="WordSection1">
  <div class="a4-page">
    ${headerBannerHtml}

    <table style="width:100%; border-collapse:collapse; border:1.5px solid #000; margin-bottom:14px;">
      <tr>
        <td style="width:62%; background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">1. Hava Aracı Bilgileri</td>
        <td style="width:38%; background:#d9d9d9; border:1px solid #000; padding:4px 8px;"></td>
      </tr>
      <tr>
        <td style="border:1px solid #000; padding:8px 10px; vertical-align:top; height:72px; font-size:9.5pt;">
          1. Hava Aracı Tipi : <strong>${esc(form.havaAraciTipi)}</strong>
        </td>
        <td style="border:1px solid #000; padding:6px 12px; vertical-align:middle; font-size:9.5pt; line-height:1.45;">
          <div>${cb(form.selectedTypes.at802)} AT-802</div>
          <div>${cb(form.selectedTypes.t70)} T-70</div>
          <div>${cb(form.selectedTypes.bell429)} Bell-429</div>
          <div>${cb(form.selectedTypes.b360)} B-360</div>
        </td>
      </tr>
    </table>

    <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
      <tr>
        <td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">2-Talep Edilen Malzemeler</td>
      </tr>
      <tr style="background:#f3f4f6;">
        <th style="border:1px solid #000; padding:5px 4px; width:7%; text-align:center; font-size:8.5pt;">Sıra<br/>No</th>
        <th style="border:1px solid #000; padding:5px 6px; width:24%; text-align:left; font-size:8.5pt;">Parça Adı/Tanımı</th>
        <th style="border:1px solid #000; padding:5px 6px; width:19%; text-align:left; font-size:8.5pt;">Parça Numarası<br/>(Part Number)</th>
        <th style="border:1px solid #000; padding:5px 6px; width:20%; text-align:left; font-size:8.5pt;">Alternatif Parça<br/>Numaraları<br/>(Eğer biliniyorsa)</th>
        <th style="border:1px solid #000; padding:5px 4px; width:14%; text-align:center; font-size:8.5pt;">İstenen Miktar</th>
        <th style="border:1px solid #000; padding:5px 4px; width:16%; text-align:center; font-size:8.5pt;">Parçanın Öncelik<br/>Derecesi</th>
      </tr>
      ${rowsHtml}
      <tr>
        <td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; text-align:center; font-weight:bold; font-size:9pt;">
          Parcanın Öncelik Derecesi kısmına yazılacaklar
        </td>
      </tr>
      <tr>
        <td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:9pt;">
          <span style="color:#dc2626; font-weight:bold;">Yüksek</span> – Kısa Sürede Gerekli (Arıza vb)
        </td>
      </tr>
      <tr>
        <td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:9pt;">
          <span style="color:#eab308; font-weight:bold;">Normal</span> – Planlı Bakım İçin
        </td>
      </tr>
      <tr>
        <td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:9pt;">
          <span style="color:#16a34a; font-weight:bold;">Düşük</span> – Depo Miktarını Tamamlamak İçin
        </td>
      </tr>
      <tr>
        <td colspan="6" style="border:1px solid #000; padding:5px 8px; font-size:8.5pt; font-style:italic;">
          *Ek satırlara ihtiyaç duyulması halinde bu bölümü çoğaltabilirsiniz.*
        </td>
      </tr>
    </table>
  </div>

  ${forWord ? '<br clear="all" style="page-break-before:always;" />' : ''}

  <div class="a4-page">
    ${headerBannerHtml}

    <table style="width:100%; border-collapse:collapse; border:1.5px solid #000; margin-bottom:14px;">
      <tr>
        <td style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">3. Parça Kataloğu Görseli</td>
      </tr>
      <tr>
        <td style="border:1px solid #000; padding:10px; vertical-align:top; height:430px;">
          <div style="font-size:9.5pt; margin-bottom:10px; line-height:1.4;">
            Lütfen istenen parçanın katalogdan alınmış görselini formun ekine ekleyiniz.<br/>
            Doğru parçanın tespit edilmesine yardımcı olacaktır.
          </div>
          ${
            catalogImgDataUrl
              ? `<div style="text-align:center; margin-top:8px;"><img src="${catalogImgDataUrl}" style="max-width:100%; max-height:355px; border:1px solid #cbd5e1; object-fit:contain;" /></div>`
              : ''
          }
        </td>
      </tr>
    </table>

    <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
      <tr>
        <td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">4. Ek Notlar/Açıklamalar</td>
      </tr>
      <tr>
        <td colspan="2" style="border:1px solid #000; padding:8px 10px; vertical-align:top; height:160px;">
          <div style="font-size:9.5pt; margin-bottom:6px;">Parçanın özel gereksinimlerini belirten ek bilgiler var ise:</div>
          <div style="font-size:9.5pt; white-space:pre-wrap; margin-top:6px;">${esc(form.ekNotlar)}</div>
        </td>
      </tr>
      <tr>
        <td style="border:1px solid #000; padding:8px 10px; width:50%; font-size:9.5pt; height:36px;">
          Talep Yapan Personel Adı Soyadı:
        </td>
        <td style="border:1px solid #000; padding:8px 10px; width:50%; font-size:9.5pt; font-weight:bold;">
          ${esc(form.talepYapanAdSoyad)}
        </td>
      </tr>
      <tr>
        <td style="border:1px solid #000; padding:8px 10px; width:50%; font-size:9.5pt; height:44px;">
          Talep Sahibi İmza:
        </td>
        <td style="border:1px solid #000; padding:8px 10px; width:50%;"></td>
      </tr>
    </table>
  </div>
  </div>
</body>
</html>`;
  };

  const handlePrintZeroReport = () => {
    const html = buildZeroReportDocumentHtml(zeroReportForm, capturedCatalogImage, false);
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);
    const doc = iframe.contentDocument || iframe.contentWindow?.document;
    if (doc) {
      doc.open();
      doc.write(html);
      doc.close();
      setTimeout(() => {
        try {
          iframe.contentWindow?.focus();
          iframe.contentWindow?.print();
        } finally {
          setTimeout(() => {
            if (document.body.contains(iframe)) document.body.removeChild(iframe);
          }, 1500);
        }
      }, 250);
    }
  };

  const handleDownloadZeroReportWord = () => {
    const html = buildZeroReportDocumentHtml(zeroReportForm, capturedCatalogImage, true);
    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const pnTag = (zeroReportForm.rows[0]?.parcaNo || 'Talep').replace(/[^a-zA-Z0-9_-]/g, '_');
    a.download = `OGM_Hava_Araci_Yedek_Parca_Talep_Formu_${pnTag}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showNotification('📄 HAVA ARACI YEDEK PARÇA TALEP FORMU Word (.doc) olarak indirildi!', 'success');
  };

  // Step-by-step Zero Report launcher: Step 1 -> Capture & Annotate screenshot, Step 2 -> Open 2-page form
  const handleOpenZeroReport = () => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    const dateStr = `${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;

    const activeUnitKey = String(activePub?.unitKey || activePub?.unit || selectedUnitFilter || 'at802').toLowerCase();
    const unitLabel = activePub?.unitLabel || UNIT_FOLDER_OPTIONS.find(u => u.key === activeUnitKey)?.label || 'AT-802F';

    const selectedTypes = {
      at802: activeUnitKey.includes('at802'),
      t70: activeUnitKey.includes('t70'),
      bell429: activeUnitKey.includes('bell') || activeUnitKey.includes('429'),
      b360: activeUnitKey.includes('b360')
    };

    const seedRows: ZeroReportRow[] = Array.from({ length: 20 }, (_, i) => ({
      siraNo: i + 1,
      parcaAdi: '',
      parcaNo: '',
      alternatifPn: '',
      istenenMiktar: '',
      oncelik: ''
    }));

    if (matchedDepotItems.length > 0) {
      matchedDepotItems.slice(0, 20).forEach((it, idx) => {
        seedRows[idx] = {
          siraNo: idx + 1,
          parcaAdi: it.adi || '',
          parcaNo: it.parcaNo !== '-' ? it.parcaNo : depoSearchQuery.trim(),
          alternatifPn: '',
          istenenMiktar: '1',
          oncelik: 'Yüksek'
        };
      });
    } else if (depoSearchQuery.trim()) {
      seedRows[0] = {
        siraNo: 1,
        parcaAdi: '',
        parcaNo: depoSearchQuery.trim(),
        alternatifPn: '',
        istenenMiktar: '1',
        oncelik: 'Yüksek'
      };
    }

    setZeroReportForm({
      havaAraciTipi: unitLabel,
      selectedTypes,
      rows: seedRows,
      ekNotlar: '',
      talepYapanAdSoyad: '',
      date: dateStr
    });

    setZeroReportData({
      query: depoSearchQuery.trim(),
      targetUnit: unitLabel,
      date: dateStr,
      items: matchedDepotItems
    });
    setIsDepoAutocompleteOpen(false);
    // Start Step 1 automatically: Screen capture -> Crop & Highlight -> Form
    handleCaptureCatalogScreenshot(true);
  };

  // Helper to open a publication in Viewer Mode
  const openPublicationInViewer = (pub: TechPublication) => {
    setDepoSearchQuery('');
    setPubSearchQuery('');
    setIsStockDrawerOpen(false);
    setIsDepoAutocompleteOpen(false);
    setBulutOpenedPdf(null);
    lastClipboardTextRef.current = '';
    setActivePub(pub);
    setIsViewerActive(true);
  };

  // Helper to open OGM Bulut live AT-802F folder in Viewer Mode
  const handleOpenOgmBulutFolderInViewer = () => {
    openPublicationInViewer({
      id: 'ogm_bulut_at802_folder',
      unit: 'at802',
      unitKey: 'at802',
      unitLabel: 'AT-802F',
      category: 'CANLI KLASÖR',
      revision: 'OGM Bulut',
      section: 'AT-802F OGM BULUT AÇIK BAĞLANTI KAYNAĞI',
      title: 'AT-802F OGM BULUT TEKNİK YAYINLAR KLASÖRÜ (OGMDrive)',
      fileName: 'OGMDrive_AT802_Teknik_Yayinlar',
      fileSize: 'Canlı Arşiv',
      uploadDate: 'Canlı',
      viewUrl: OGM_BULUT_AT802_URL,
      notes: 'OGM Bulut açık bağlantı kaynağı'
    });
  };

  if (!isOpen) return null;

  return (
    <div id="tech-pubs-modal-container" className="fixed inset-0 z-[9999] bg-[#020d07] flex flex-col font-sans select-text text-slate-100">
      
      {/* Toast Notification */}
      <AnimatePresence>
        {notification && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className={`fixed top-4 right-8 z-[10000] px-5 py-3 rounded-xl shadow-2xl flex items-center gap-3 border text-sm font-semibold ${
              notification.type === 'success'
                ? 'bg-emerald-900/90 text-emerald-200 border-emerald-500/50'
                : notification.type === 'error'
                ? 'bg-rose-900/90 text-rose-200 border-rose-500/50'
                : 'bg-sky-900/90 text-sky-200 border-sky-500/50'
            }`}
          >
            {notification.type === 'success' && <CheckCircle className="w-5 h-5 text-emerald-400" />}
            {notification.type === 'error' && <AlertCircle className="w-5 h-5 text-rose-400" />}
            {notification.type === 'info' && <Sparkles className="w-5 h-5 text-sky-400" />}
            <span>{notification.message}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* ========================================================================= */}
      {/* VIEW MODE 1: DIRECT IFRAME VIEWER (IMAGE 2 STYLE)                         */}
      {/* When isViewerActive && activePub is true                                  */}
      {/* ========================================================================= */}
      {isViewerActive && activePub ? (
        <div id="tech-pub-direct-viewer-screen" className="flex-1 flex flex-col bg-slate-950 overflow-hidden relative">
          
          {/* Top Main Navigation Header (Clean & Minimalist like Görevlendirme Çizelgeleri) */}
          <div className="bg-[#0b1329] border-b border-slate-800 px-5 py-2.5 flex items-center justify-between gap-4 select-none shadow-md shrink-0 z-30">
            {/* Back button & Title */}
            <div className="flex items-center gap-3">
              <button
                id="btn-back-to-catalog"
                onClick={() => {
                  if (activePub?.id === 'ogm_bulut_at802_folder' && (bulutOpenedPdf || isBulutPdfOverlayOpen)) {
                    setBulutOpenedPdf(null);
                    setIsBulutPdfOverlayOpen(false);
                    setBulutOverlayUrlInput('');
                    setIsStockDrawerOpen(false);
                    wasBulutIframeFocusedRef.current = false;
                    return;
                  }
                  setIsViewerActive(false);
                  setDepoSearchQuery('');
                  setPubSearchQuery('');
                  setIsStockDrawerOpen(false);
                  lastClipboardTextRef.current = '';
                }}
                className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-white rounded-lg flex items-center gap-1.5 text-xs font-black uppercase tracking-wider transition-all cursor-pointer border border-slate-700 active:scale-95"
                title={activePub?.id === 'ogm_bulut_at802_folder' && (bulutOpenedPdf || isBulutPdfOverlayOpen) ? 'OGM Bulut linkine geri dön' : 'Teknik yayın listesine dön'}
              >
                <ArrowLeft className="w-4 h-4" />
                <span>GERİ</span>
              </button>
              <div className="text-left">
                <h2 className="text-xs font-black text-slate-100 uppercase tracking-wider">
                  TEKNİK YAYINLAR & DOKÜMANTASYON &bull; {activePub.unitLabel}
                </h2>
              </div>
            </div>

            {/* Right Action Buttons */}
            <div className="flex items-center gap-2">
              {(activePdfBlobUrl || bulutOpenedPdf?.blobUrl) && (
                <a
                  id="btn-viewer-download-pdf"
                  href={bulutOpenedPdf?.blobUrl || activePdfBlobUrl}
                  download={bulutOpenedPdf?.fileName || activePub.fileName}
                  className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white rounded-lg flex items-center gap-1.5 text-[10px] font-black uppercase tracking-wider transition-all cursor-pointer shadow-sm"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>PLAN PDF İNDİR</span>
                </a>
              )}

              <button
                id="btn-viewer-close-all"
                onClick={onClose}
                className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg ml-1"
                title="Kapat"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          </div>

          {/* Sub-Header: File Title & Live Depo Search Bar (Üstteki Arama Kısmı) */}
          <div className="bg-slate-900 border-b border-slate-800 px-5 py-2 flex items-center justify-between gap-4 shrink-0 flex-wrap z-30 relative">
            {/* Left: Document File Name + Back to OGM Bulut button when viewing a PDF opened from OGM Bulut */}
            <div className="flex items-center gap-2 text-slate-300 min-w-0">
              {activePub.id === 'ogm_bulut_at802_folder' && bulutOpenedPdf && (
                <button
                  type="button"
                  onClick={() => setBulutOpenedPdf(null)}
                  className="px-2.5 py-1 rounded-lg bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 text-[10px] font-black uppercase flex items-center gap-1 cursor-pointer shrink-0"
                >
                  <ArrowLeft className="w-3 h-3" />
                  <span>OGM Bulut Klasörüne Dön</span>
                </button>
              )}
              <span className="font-mono text-xs font-black text-white uppercase tracking-wide truncate max-w-[280px]">
                {(bulutOpenedPdf?.fileName || activePub.fileName).toUpperCase()}
              </span>
              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 shrink-0">
                {bulutOpenedPdf ? 'OGM BULUT PDF' : (activePub.category || 'IPC')}
              </span>
              {isDepotLoading && (
                <span className="text-[10px] text-emerald-400 font-semibold animate-pulse">
                  Depo verileri yükleniyor...
                </span>
              )}
            </div>

            {/* Center / Right: Live Depo Search Bar + Autocomplete Dropdown */}
            <div className="flex items-center gap-2.5 flex-1 max-w-2xl justify-end">
              <div ref={viewerSearchWrapperRef} className="relative flex-1 max-w-lg flex items-center gap-2">
                <div className="relative flex-1">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    id="input-viewer-depo-search"
                    type="text"
                    value={depoSearchQuery}
                    onFocus={() => {
                      if (depoSearchQuery.trim().length > 0) setIsDepoAutocompleteOpen(true);
                    }}
                    onChange={e => {
                      const val = e.target.value;
                      setDepoSearchQuery(val);
                      setAutocompleteActiveIdx(-1);
                      setIsDepoAutocompleteOpen(val.trim().length > 0);
                    }}
                    onKeyDown={e => {
                      if (e.key === 'ArrowDown' && autocompleteSuggestions.length > 0) {
                        e.preventDefault();
                        setIsDepoAutocompleteOpen(true);
                        setAutocompleteActiveIdx(prev => Math.min(prev + 1, autocompleteSuggestions.length - 1));
                      } else if (e.key === 'ArrowUp' && autocompleteSuggestions.length > 0) {
                        e.preventDefault();
                        setAutocompleteActiveIdx(prev => Math.max(prev - 1, 0));
                      } else if (e.key === 'Enter') {
                        if (isDepoAutocompleteOpen && autocompleteActiveIdx >= 0 && autocompleteSuggestions[autocompleteActiveIdx]) {
                          e.preventDefault();
                          handleSelectAutocompleteItem(autocompleteSuggestions[autocompleteActiveIdx]);
                        } else if (depoSearchQuery.trim()) {
                          setIsDepoAutocompleteOpen(false);
                          setIsStockDrawerOpen(true);
                        }
                      } else if (e.key === 'Escape') {
                        setIsDepoAutocompleteOpen(false);
                      }
                    }}
                    placeholder="Parça No (P/N), Seri No veya Malzeme Adı ile Canlı Depo Ara..."
                    className="w-full bg-slate-950 border border-slate-700 rounded-xl pl-9 pr-14 py-1.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-emerald-500 transition-colors"
                    autoComplete="off"
                  />
                  <div className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center gap-1">
                    {depoSearchQuery && (
                      <button
                        type="button"
                        onClick={() => {
                          setDepoSearchQuery('');
                          setIsDepoAutocompleteOpen(false);
                        }}
                        className="p-1 text-slate-400 hover:text-white"
                        title="Aramayı Temizle"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>

                  {/* Live Autocomplete Dropdown */}
                  {isDepoAutocompleteOpen && depoSearchQuery.trim().length > 0 && (
                    <div
                      id="viewer-depo-autocomplete-dropdown"
                      className="absolute left-0 right-0 top-full mt-1.5 bg-slate-900 border border-emerald-500/40 rounded-2xl shadow-2xl overflow-hidden z-[100] max-h-80 overflow-y-auto"
                    >
                      {autocompleteSuggestions.length === 0 ? (
                        <div className="p-3 text-xs text-slate-400 flex items-center justify-between gap-2">
                          <span>&ldquo;{depoSearchQuery}&rdquo; için depoda eşleşen kayıt bulunamadı.</span>
                          <button
                            type="button"
                            onClick={() => {
                              setIsDepoAutocompleteOpen(false);
                              handleOpenZeroReport();
                            }}
                            className="px-2.5 py-1 rounded-lg bg-amber-500 text-slate-950 font-black text-[10px] shrink-0 cursor-pointer"
                          >
                            ZERO REPORT
                          </button>
                        </div>
                      ) : (
                        autocompleteSuggestions.map((item, idx) => {
                          const qty = item.toplamStok ?? item.miktarNum ?? 0;
                          const isActive = idx === autocompleteActiveIdx;
                          return (
                            <button
                              key={`${item.unitKey}_${item.parcaNo}_${idx}`}
                              type="button"
                              onClick={() => handleSelectAutocompleteItem(item)}
                              className={`w-full px-3.5 py-2 text-left border-b border-slate-800/80 last:border-b-0 transition-colors flex items-center justify-between gap-3 cursor-pointer ${
                                isActive ? 'bg-emerald-900/50' : 'hover:bg-slate-800/90'
                              }`}
                            >
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center gap-1.5 mb-0.5">
                                  <span className="px-1.5 py-0.5 rounded bg-emerald-500/15 text-emerald-300 border border-emerald-500/30 text-[9px] font-black uppercase shrink-0">
                                    {item.unitTitle}
                                  </span>
                                  <span className="text-xs font-bold text-white truncate">{item.adi || item.parcaNo}</span>
                                </div>
                                <div className="text-[10px] font-mono text-slate-400 flex items-center gap-2.5 flex-wrap">
                                  <span>P/N: <strong className="text-emerald-300">{item.parcaNo}</strong></span>
                                  {item.seriNo && item.seriNo !== '-' && <span>S/N: {item.seriNo}</span>}
                                  {item.yer && item.yer !== '-' && <span>Konum: <strong className="text-amber-300">{item.yer}</strong></span>}
                                </div>
                              </div>
                              <span
                                className={`px-2 py-0.5 rounded-lg text-[10px] font-black font-mono shrink-0 ${
                                  qty > 0
                                    ? 'bg-emerald-950 text-emerald-300 border border-emerald-700/60'
                                    : 'bg-amber-950 text-amber-300 border border-amber-700/60'
                                }`}
                              >
                                Stok: {qty}
                              </span>
                            </button>
                          );
                        })
                      )}
                    </div>
                  )}
                </div>

                <button
                  type="button"
                  onClick={() => {
                    setIsDepoAutocompleteOpen(false);
                    setIsStockDrawerOpen(true);
                  }}
                  className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-black uppercase rounded-xl transition shadow-sm cursor-pointer active:scale-95 shrink-0"
                >
                  ARA
                </button>
              </div>

              {/* Matched count indicator badge & drawer toggle */}
              <button
                id="btn-toggle-stock-drawer"
                type="button"
                onClick={() => setIsStockDrawerOpen(!isStockDrawerOpen)}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 border shadow-sm cursor-pointer select-none ${
                  stockSearchStatus === 'in_stock'
                    ? 'bg-emerald-950/80 border-emerald-500/80 text-emerald-300 hover:bg-emerald-900/80 shadow-emerald-950/40'
                    : stockSearchStatus === 'zero_stock'
                    ? 'bg-amber-950/80 border-amber-500/80 text-amber-300 hover:bg-amber-900/80 shadow-amber-950/40'
                    : stockSearchStatus === 'not_found'
                    ? 'bg-rose-950/80 border-rose-500/80 text-rose-300 hover:bg-rose-900/80 shadow-rose-950/40'
                    : 'bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700'
                }`}
              >
                <Package className={`w-3.5 h-3.5 ${
                  stockSearchStatus === 'in_stock'
                    ? 'text-emerald-400'
                    : stockSearchStatus === 'zero_stock'
                    ? 'text-amber-400'
                    : stockSearchStatus === 'not_found'
                    ? 'text-rose-400'
                    : 'text-slate-400'
                }`} />
                <span>
                  {stockSearchStatus === 'in_stock'
                    ? `${matchedDepotItems.length} Stok Kaydı`
                    : stockSearchStatus === 'zero_stock'
                    ? `${matchedDepotItems.length} Stok Kaydı (Stok: 0)`
                    : stockSearchStatus === 'not_found'
                    ? '0 Stok Kaydı (Kayıt Yok)'
                    : 'Stok Kaydı'}
                </span>
                {isStockDrawerOpen ? <ChevronDown className="w-3 h-3" /> : <ChevronUp className="w-3 h-3" />}
              </button>

              {/* Zero Report Button */}
              {isZeroReportEligible && (
                <button
                  id="btn-viewer-zero-report"
                  type="button"
                  onClick={handleOpenZeroReport}
                  className="px-3 py-1.5 rounded-xl bg-amber-500 hover:bg-amber-400 active:scale-95 text-slate-950 font-black text-xs transition-all flex items-center gap-1.5 shadow-md animate-pulse cursor-pointer shrink-0"
                  title="Depoda bulunmayan veya stoğu 0 olan parça için Zero / İhtiyaç Raporu oluştur"
                >
                  <AlertTriangle className="w-3.5 h-3.5" />
                  <span>ZERO REPORT OLUŞTUR</span>
                </button>
              )}
            </div>
          </div>

          {/* Main Area: OGM Bulut Unlocker Iframe OR Direct Verified PDF Viewer */}
          <div className="flex-1 bg-slate-950 overflow-hidden relative flex flex-col">
            {activePub.id === 'ogm_bulut_at802_folder' ? (
              <>
                {/* Return-from-Store Page Reload Banner */}
                {showExtensionReloadBanner && (
                  <div
                    id="ogm-extension-reload-banner"
                    className="bg-amber-500 text-slate-950 px-5 py-2.5 flex items-center justify-between gap-3 shrink-0 z-40 shadow-lg"
                  >
                    <div className="flex items-center gap-2.5 text-xs font-black">
                      <AlertTriangle className="w-4 h-4 shrink-0" />
                      <span>Eklenti kurulumunu tamamladıysanız bağlantının açılması için lütfen sayfayı yenileyin.</span>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <button
                        id="btn-reload-after-extension"
                        type="button"
                        onClick={() => window.location.reload()}
                        className="px-3.5 py-1.5 rounded-lg bg-slate-950 hover:bg-slate-800 text-white text-xs font-black uppercase flex items-center gap-1.5 cursor-pointer shadow-sm"
                      >
                        <RotateCw className="w-3.5 h-3.5" />
                        <span>Sayfayı Yenile</span>
                      </button>
                      <button
                        type="button"
                        onClick={() => setShowExtensionReloadBanner(false)}
                        className="p-1 rounded hover:bg-amber-600/30 text-slate-900 cursor-pointer"
                        title="Kapat"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                )}

                {bulutOpenedPdf && (
                  bulutOpenedPdf.loading ? (
                    <div className="flex-1 w-full flex flex-col items-center justify-center p-8 text-center bg-slate-950 text-slate-300">
                      <RefreshCw className="w-10 h-10 mb-3 text-emerald-400 animate-spin" />
                      <p className="text-sm font-bold text-white mb-1">OGM Bulut PDF Dökümanı Yükleniyor...</p>
                      <p className="text-xs text-slate-400">{bulutOpenedPdf.fileName}</p>
                    </div>
                  ) : (
                    <iframe
                      src={(() => {
                        const raw = bulutOpenedPdf.blobUrl || bulutOpenedPdf.url;
                        return raw.includes('#') ? raw : `${raw}#toolbar=1&view=FitH`;
                      })()}
                      className="flex-1 w-full h-full border-0 bg-slate-950"
                      title={bulutOpenedPdf.fileName}
                      id="tech-pub-bulut-pdf-iframe"
                    />
                  )
                )}

                <div className={bulutOpenedPdf ? 'hidden' : 'flex-1 flex flex-col overflow-hidden relative'}>
                  {/* Connection-Refused Helper Strip (Eklenti Yükle) */}
                  {bulutCheckStatus === 'geo_blocked' && !hideExtensionHelperStrip && (
                    <div className="bg-slate-900/95 border-b border-slate-800 px-4 py-1.5 flex items-center justify-between gap-3 text-xs shrink-0 flex-wrap">
                      <div className="flex items-center gap-2 text-slate-300">
                        <AlertCircle className="w-3.5 h-3.5 text-amber-400 shrink-0" />
                        <span>Bağlantı reddedildi hatası alıyorsanız eklentiyi yükleyip sayfayı yenileyin:</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <a
                          id="btn-install-iframe-extension"
                          href={OGM_EXTENSION_STORE_URL}
                          target="_blank"
                          rel="noopener noreferrer"
                          onClick={() => {
                            awaitingExtensionInstallRef.current = true;
                          }}
                          className="px-3 py-1 rounded-lg bg-amber-500 hover:bg-amber-400 text-slate-950 text-[11px] font-black uppercase flex items-center gap-1.5 cursor-pointer shadow-sm"
                        >
                          <ExternalLink className="w-3.5 h-3.5" />
                          <span>Eklenti Yükle</span>
                        </a>
                        <button
                          type="button"
                          onClick={() => setIsBulutPdfOverlayOpen(true)}
                          className="px-2.5 py-1 rounded-lg bg-emerald-700 hover:bg-emerald-600 text-white text-[11px] font-bold flex items-center gap-1 cursor-pointer"
                        >
                          <Upload className="w-3 h-3" />
                          <span>PDF / Bağlantı Aç</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => setHideExtensionHelperStrip(true)}
                          className="p-1 text-slate-400 hover:text-white cursor-pointer"
                          title="Gizle"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  )}

                  {/* Embedded OGM Bulut Iframe */}
                  <div className="flex-1 relative bg-white overflow-hidden">
                    {bulutCheckStatus === 'checking' ? (
                      <div className="h-full w-full flex flex-col items-center justify-center p-8 text-center bg-slate-950 text-slate-300">
                        <RefreshCw className="w-10 h-10 mb-3 text-emerald-400 animate-spin" />
                        <p className="text-sm font-bold text-white mb-1">OGM Bulut Arşivi Yükleniyor...</p>
                        <p className="text-xs text-slate-400">AT-802F teknik doküman klasörü hazırlanıyor.</p>
                      </div>
                    ) : (
                      <iframe
                        key={`${bulutIframeMode}_${bulutCheckStatus}`}
                        src={
                          bulutIframeMode === 'proxy'
                            ? `/api/ogm-bulut-proxy?url=${encodeURIComponent(OGM_BULUT_AT802_URL)}`
                            : OGM_BULUT_AT802_URL
                        }
                        onError={() => setHideExtensionHelperStrip(false)}
                        className={`w-full h-full border-0 bg-white transition-all duration-200 ${
                          isBulutPdfOverlayOpen ? 'blur-md scale-[1.01] pointer-events-none select-none' : ''
                        }`}
                        title="AT-802F OGM Bulut Teknik Yayınlar"
                        id="tech-pub-ogm-bulut-iframe"
                        allow="clipboard-read; clipboard-write"
                      />
                    )}

                    {/* Blurred PDF / Link Loader Overlay (triggered automatically when new tab opens) */}
                    {isBulutPdfOverlayOpen && (
                      <div
                        id="ogm-bulut-pdf-blur-overlay"
                        className="absolute inset-0 z-30 bg-slate-950/75 backdrop-blur-md flex items-center justify-center p-6"
                      >
                        <div className="max-w-lg w-full bg-[#071a10] border border-emerald-500/40 rounded-2xl p-6 shadow-2xl text-slate-100 space-y-5">
                          <div className="flex items-start justify-between gap-3">
                            <div>
                              <h3 className="text-base font-black text-white">PDF / Bağlantı Açma Alanı</h3>
                              <p className="text-xs text-slate-400 mt-1">
                                Yeni sekmede açılan PDF bağlantısını yapıştırıp <strong className="text-emerald-300">Bağlantıyı Aç</strong>&apos;a basın veya indirdiğiniz PDF dosyasını yükleyin.
                              </p>
                            </div>
                            <button
                              type="button"
                              onClick={() => {
                                setIsBulutPdfOverlayOpen(false);
                                setBulutOverlayUrlInput('');
                                wasBulutIframeFocusedRef.current = false;
                              }}
                              className="p-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-400 hover:text-white cursor-pointer"
                              title="OGM Bulut Klasörüne Dön"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>

                          {/* URL Input + Bağlantıyı Aç */}
                          <div className="space-y-2">
                            <label className="block text-[11px] font-black uppercase tracking-wider text-emerald-300">
                              PDF Bağlantısı (URL)
                            </label>
                            <div className="flex items-center gap-2">
                              <input
                                ref={bulutOverlayInputRef}
                                id="input-bulut-overlay-url"
                                type="text"
                                value={bulutOverlayUrlInput}
                                onChange={e => setBulutOverlayUrlInput(e.target.value)}
                                onKeyDown={e => {
                                  if (e.key === 'Enter' && bulutOverlayUrlInput.trim()) {
                                    const u = bulutOverlayUrlInput.trim();
                                    setIsBulutPdfOverlayOpen(false);
                                    openExternalPdfInSameViewer(u, decodeURIComponent(u.split('?')[0].split('/').pop() || 'OGM_Dokuman.pdf'));
                                  }
                                }}
                                placeholder="https://bulut.ogm.gov.tr/..."
                                className="flex-1 bg-slate-950 border border-slate-700 focus:border-emerald-500 rounded-xl px-3.5 py-2.5 text-xs text-white placeholder-slate-500 outline-none"
                              />
                              <button
                                id="btn-bulut-overlay-open-url"
                                type="button"
                                disabled={!bulutOverlayUrlInput.trim()}
                                onClick={() => {
                                  const u = bulutOverlayUrlInput.trim();
                                  if (!u) return;
                                  setIsBulutPdfOverlayOpen(false);
                                  openExternalPdfInSameViewer(u, decodeURIComponent(u.split('?')[0].split('/').pop() || 'OGM_Dokuman.pdf'));
                                }}
                                className="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs font-black uppercase flex items-center gap-1.5 cursor-pointer shrink-0 shadow-md"
                              >
                                <Link2 className="w-3.5 h-3.5" />
                                <span>Bağlantıyı Aç</span>
                              </button>
                            </div>
                          </div>

                          {/* PDF File Dropzone */}
                          <div
                            onDragOver={e => {
                              e.preventDefault();
                              setBulutOverlayDragActive(true);
                            }}
                            onDragLeave={() => setBulutOverlayDragActive(false)}
                            onDrop={e => {
                              e.preventDefault();
                              setBulutOverlayDragActive(false);
                              const f = e.dataTransfer.files?.[0];
                              if (f) {
                                const url = URL.createObjectURL(f);
                                setBulutOpenedPdf({ url, blobUrl: url, fileName: f.name, loading: false });
                                setIsBulutPdfOverlayOpen(false);
                                setBulutOverlayUrlInput('');
                              }
                            }}
                            className={`border-2 border-dashed rounded-xl p-5 text-center transition-colors ${
                              bulutOverlayDragActive
                                ? 'border-emerald-400 bg-emerald-500/15'
                                : 'border-slate-700 bg-slate-950/60 hover:border-emerald-500/50'
                            }`}
                          >
                            <Upload className="w-7 h-7 text-emerald-400 mx-auto mb-2" />
                            <p className="text-xs font-bold text-white mb-1">PDF Dosyasını Buraya Sürükleyin</p>
                            <p className="text-[11px] text-slate-400 mb-3">veya bilgisayarınızdan indirdiğiniz PDF dosyasını seçin</p>
                            <label className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 text-xs font-black uppercase cursor-pointer shadow-sm">
                              <FileText className="w-3.5 h-3.5" />
                              <span>PDF Yükle / Seç</span>
                              <input
                                type="file"
                                accept=".pdf,application/pdf"
                                className="hidden"
                                onChange={e => {
                                  const f = e.target.files?.[0];
                                  if (f) {
                                    const url = URL.createObjectURL(f);
                                    setBulutOpenedPdf({ url, blobUrl: url, fileName: f.name, loading: false });
                                    setIsBulutPdfOverlayOpen(false);
                                    setBulutOverlayUrlInput('');
                                  }
                                  e.target.value = '';
                                }}
                              />
                            </label>
                          </div>

                          {/* Return to Bulut Link */}
                          <div className="pt-1 flex items-center justify-end border-t border-slate-800/80">
                            <button
                              id="btn-bulut-overlay-back"
                              type="button"
                              onClick={() => {
                                setIsBulutPdfOverlayOpen(false);
                                setBulutOverlayUrlInput('');
                                wasBulutIframeFocusedRef.current = false;
                              }}
                              className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-xs font-black uppercase flex items-center gap-1.5 cursor-pointer"
                            >
                              <ArrowLeft className="w-3.5 h-3.5" />
                              <span>Geri Dön (OGM Bulut Linki)</span>
                            </button>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </>
            ) : loadingActivePdf ? (
              <div className="h-full w-full flex flex-col items-center justify-center p-8 text-center bg-slate-950 text-slate-300 select-none">
                <RefreshCw className="w-10 h-10 mb-3 text-emerald-400 animate-spin" />
                <p className="text-sm font-bold text-white mb-1 font-sans">PDF Dökümanı Arka Planda İndiriliyor...</p>
                <p className="text-xs text-slate-400 font-sans max-w-md">Metin seçimi ve parça no kopyalama için doğrudan tarayıcı PDF görüntüleyicisine aktarılıyor.</p>
              </div>
            ) : activePdfBlobUrl ? (
              <iframe
                src={`${activePdfBlobUrl}#toolbar=1&view=FitH`}
                className="w-full h-full border-0 bg-slate-950"
                title={activePub.title}
                id="tech-pub-iframe"
              />
            ) : driveFallbackPreviewUrl ? (
              <iframe
                src={driveFallbackPreviewUrl}
                className="w-full h-full border-0 bg-slate-950"
                title={activePub.title}
                id="tech-pub-iframe"
                allow="autoplay"
              />
            ) : (
              <div className="h-full w-full flex flex-col items-center justify-center p-8 text-center bg-slate-950 text-slate-400 select-none">
                <FileText className="w-12 h-12 mb-3 text-slate-600" />
                <p className="text-xs font-mono text-slate-300">{pdfRenderError || 'Döküman indirilemedi veya erişim sağlanamadı.'}</p>
              </div>
            )}
          </div>

          {/* Bottom Resizable Live Depo Search Drawer in Viewer */}
          {isStockDrawerOpen && (
            <div
              id="tech-pubs-viewer-bottom-drawer"
              style={{ height: drawerHeight }}
              className="bg-slate-900/95 border-t border-slate-700/80 flex flex-col shadow-2xl transition-[height] duration-75 relative shrink-0 z-30"
            >
              {/* Resize Handle */}
              <div
                onMouseDown={handleDragStart}
                className="h-2 w-full bg-slate-800 hover:bg-emerald-500/50 cursor-row-resize flex items-center justify-center transition-colors group"
                title="Yüksekliği ayarlamak için yukarı/aşağı sürükleyiniz"
              >
                <div className="w-12 h-1 rounded-full bg-slate-600 group-hover:bg-white transition-colors" />
              </div>

              {/* Drawer Header */}
              <div className="px-5 py-2 bg-slate-950 flex items-center justify-between border-b border-slate-800 shrink-0">
                <div className="flex items-center gap-2 text-xs font-bold">
                  <Archive className={`w-4 h-4 ${
                    stockSearchStatus === 'in_stock'
                      ? 'text-emerald-400'
                      : stockSearchStatus === 'zero_stock'
                      ? 'text-amber-400'
                      : stockSearchStatus === 'not_found'
                      ? 'text-rose-400'
                      : 'text-emerald-400'
                  }`} />
                  <span className="text-white">CANLI DEPO STOK LİSTESİ &bull; {matchedDepotItems.length} Eşleşme</span>
                  {stockSearchStatus === 'in_stock' && (
                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                      MEVCUT (STOK &gt; 0)
                    </span>
                  )}
                  {stockSearchStatus === 'zero_stock' && (
                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                      STOK: 0 (TÜKENMİŞ)
                    </span>
                  )}
                  {stockSearchStatus === 'not_found' && (
                    <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-rose-500/20 text-rose-300 border border-rose-500/30">
                      SİSTEMDE KAYIT YOK
                    </span>
                  )}
                </div>
                <div className="flex items-center gap-2">
                  {isZeroReportEligible && (
                    <button
                      type="button"
                      onClick={handleOpenZeroReport}
                      className="px-2.5 py-1 rounded-lg bg-amber-500/20 text-amber-300 border border-amber-500/40 text-[11px] font-bold hover:bg-amber-500/30 transition-all flex items-center gap-1 cursor-pointer"
                    >
                      <AlertTriangle className="w-3 h-3" />
                      <span>Zero Report Oluştur</span>
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() => setIsStockDrawerOpen(false)}
                    className="p-1 text-slate-400 hover:text-white rounded cursor-pointer"
                    title="Kapat"
                  >
                    <ChevronDown className="w-4 h-4" />
                  </button>
                </div>
              </div>

              {/* Drawer Content Table */}
              <div className="flex-1 overflow-y-auto p-3">
                {matchedDepotItems.length === 0 ? (
                  <div className="h-full flex flex-col items-center justify-center text-center p-6 text-slate-400">
                    <Package className="w-8 h-8 text-slate-600 mb-2" />
                    <p className="text-xs font-semibold text-slate-300">
                      {depoSearchQuery.trim()
                        ? `'${depoSearchQuery}' kriterine uygun parça bulunamadı.`
                        : 'Arama kutusuna parça numarası girerek veya PDF üzerinden metin seçerek depolarda aratabilirsiniz.'}
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2.5">
                    {matchedDepotItems.map((item, i) => (
                      <div
                        key={i}
                        className="p-2.5 rounded-xl bg-slate-950/80 border border-slate-800 hover:border-emerald-500/40 transition-all"
                      >
                        <div className="flex items-start justify-between gap-2 mb-1">
                          <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-emerald-500/10 text-emerald-300 border border-emerald-500/30">
                            {item.unitTitle}
                          </span>
                          <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${
                            (item.toplamStok > 0 || item.miktarNum > 0) ? 'bg-emerald-950 text-emerald-400 border border-emerald-800' : 'bg-rose-950 text-rose-400 border border-rose-800'
                          }`}>
                            {(item.toplamStok > 0 || item.miktarNum > 0) ? 'VAR' : 'YOK'}
                          </span>
                        </div>
                        <h4 className="text-xs font-bold text-white truncate mb-1">{item.adi}</h4>
                        <div className="text-[11px] font-mono text-slate-300 space-y-0.5">
                          <div className="flex justify-between">
                            <span className="text-slate-500">P/N:</span>
                            <span className="text-emerald-300 font-bold">{item.parcaNo}</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-slate-500">S/N:</span>
                            <span>{item.seriNo}</span>
                          </div>
                          
                          {item.isSarfItem ? (
                            <div className="mt-2 pt-2 border-t border-slate-800 space-y-1">
                                <div className="flex flex-col bg-yellow-400 text-black px-1.5 py-1 rounded font-black text-center shadow-sm border border-yellow-500">
                                    <div className="text-[10px] uppercase border-b border-black/10 pb-0.5 mb-1">ANKARA MEVCUT</div>
                                    <div className="text-sm">{item.ankaraMevcut || 0}</div>
                                    <div className="text-[9px] mt-1 pt-1 border-t border-black/10 font-bold italic">
                                      {item.yer || '-'}
                                    </div>
                                </div>
                                <div className="flex justify-between items-center">
                                    <span>MİLAS MEVCUT:</span>
                                    <span>{item.milasMevcut || 0}</span>
                                </div>
                                <div className="flex justify-between items-center">
                                    <span>KARAİN MEVCUT:</span>
                                    <span>{item.karainMevcut || 0}</span>
                                </div>
                                <div className="flex justify-between items-center">
                                    <span>ÇANAKKALE MEVCUT:</span>
                                    <span>{item.canakkaleMevcut || 0}</span>
                                </div>
                                <div className="flex justify-between items-center">
                                    <span>BURSA MEVCUT:</span>
                                    <span>{item.bursaMevcut || 0}</span>
                                </div>
                                <div className="flex justify-between items-center pt-1 border-t border-slate-800 text-cyan-300 font-black">
                                    <span>TOPLAM STOK:</span>
                                    <span>{item.toplamStok || 0}</span>
                                </div>
                            </div>
                          ) : (
                            <>
                                <div className="flex justify-between">
                                    <span className="text-slate-500">Yer / Raf:</span>
                                    <span className="text-amber-300">{item.yer}</span>
                                </div>
                                <div className="flex justify-between items-center">
                                    <span className="text-slate-500">Miktar / Stok:</span>
                                    <span className={`font-bold px-1.5 py-0.5 rounded text-[11px] ${
                                    item.miktarNum > 0
                                        ? 'text-emerald-300 bg-emerald-950/80 border border-emerald-800/60'
                                        : 'text-amber-300 bg-amber-950/80 border border-amber-800/60'
                                    }`}>{item.miktar}</span>
                                </div>
                            </>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

        </div>
      ) : (
        /* ========================================================================= */
        /* VIEW MODE 2: CATALOG & DOCUMENT POOL (IMAGE 1 EXACT DESIGN)              */
        /* ========================================================================= */
        <div id="tech-pubs-catalog-screen" className="flex-1 flex flex-col bg-[#061e12] overflow-hidden select-text">
          
          {/* TOP GREEN BANNER (IMAGE 1 HEADER) */}
          <div id="tech-pubs-main-header" className="bg-[#0b3d1d] border-b border-[#0f4d25] px-6 py-3.5 flex items-center justify-between gap-4 shadow-xl shrink-0">
            {/* Left: Icon & Title */}
            <div className="flex items-center gap-3.5">
              <div className="w-10 h-10 rounded-full bg-emerald-900/60 border border-emerald-400/30 flex items-center justify-center text-emerald-300 shadow-inner">
                <BookOpen className="w-5 h-5" />
              </div>
              <div>
                <h1 className="text-base font-black text-white tracking-wide flex items-center gap-2">
                  <span>📖 TEKNİK YAYINLAR HAVUZU & DOKÜMANTASYON</span>
                </h1>
                <p className="text-[11px] font-medium text-emerald-200/80">
                  Uçak & Helikopter AMM, IPC, CMM, Şema ve Bültenler &bull; Anlık Depo Stok Taraması
                </p>
              </div>
            </div>

            {/* Right Action Buttons (Image 1) */}
            <div className="flex items-center gap-2.5">
              {/* BIRIM DEGISTIR BUTTON */}
              {selectedUnitFilter && unlockedUnits.has(selectedUnitFilter) && (
                <button
                  id="btn-switch-unit"
                  onClick={() => {
                    setSelectedUnitFilter('');
                    setTargetUnitToUnlock(null);
                    setEnteredUnitPassword('');
                    setUnitPasswordError('');
                  }}
                  className="px-3.5 py-1.5 rounded-full bg-emerald-900/80 hover:bg-emerald-800 border border-emerald-500/50 text-emerald-200 text-xs font-bold transition-all flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-95"
                  title="Başka bir hava aracı veya birim seç"
                >
                  <Plane className="w-3.5 h-3.5 text-emerald-400" />
                  <span>BİRİM DEĞİŞTİR</span>
                </button>
              )}

              {/* DRIVE SYNC & REFRESH BUTTON */}
              <button
                id="btn-sync-drive-pubs"
                onClick={() => fetchDrivePublications(true)}
                disabled={isDriveLoading}
                className="px-3 py-1.5 rounded-full bg-[#052813] hover:bg-[#07381b] border border-emerald-500/40 flex items-center gap-2 text-xs font-bold text-emerald-300 transition-all cursor-pointer disabled:opacity-60 shadow-sm"
                title={lastDriveSyncTime ? `Son Senkronizasyon: ${lastDriveSyncTime}. Yenilemek için tıklayın.` : "Drive'dan Yayınları Yenile"}
              >
                <RotateCw className={`w-3.5 h-3.5 text-emerald-400 ${isDriveLoading ? 'animate-spin' : ''}`} />
                <span>{isDriveLoading ? "VERİ YÜKLENİYOR..." : (lastDriveSyncTime ? `DRIVE SENKRON (${lastDriveSyncTime})` : "DRIVE SENKRON")}</span>
              </button>

              {/* CANLI DEPO ENTEGRE badge */}
              <div className="px-3 py-1.5 rounded-full bg-[#052813] border border-emerald-500/30 flex items-center gap-2 text-xs font-bold text-emerald-300">
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
                <span>CANLI DEPO ENTEGRE</span>
              </div>

              {/* TÜM DEPOLARDA SORGULA button (Gold/Amber theme) */}
              <button
                id="btn-global-depo-query"
                onClick={() => setIsGlobalDepoModalOpen(true)}
                className="px-4 py-2 rounded-full bg-[#523e02]/80 hover:bg-[#684f02] border border-[#eab308]/60 text-[#fef08a] text-xs font-black transition-all flex items-center gap-2 shadow-md cursor-pointer active:scale-95"
              >
                <span>📦 TÜM DEPOLARDA SORGULA</span>
              </button>

              {/* + PDF / YAYIN YÜKLE button (Bright Emerald) */}
              <button
                id="btn-open-upload-modal"
                onClick={() => setIsUploadModalOpen(true)}
                className="px-4 py-2 rounded-full bg-[#009b4d] hover:bg-[#00b359] text-white text-xs font-black transition-all flex items-center gap-2 shadow-lg cursor-pointer active:scale-95"
              >
                <Upload className="w-4 h-4" />
                <span>+ PDF / YAYIN YÜKLE</span>
              </button>

              {/* Close Button */}
              <button
                id="btn-close-tech-pubs-modal"
                onClick={onClose}
                className="p-2 rounded-full bg-emerald-950/80 hover:bg-emerald-900 text-emerald-300 hover:text-white transition-all ml-1 border border-emerald-800"
                title="Kapat"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
          </div>

          {/* MAIN BODY: SIDEBAR + CONTENT AREA (IMAGE 1) */}
          <div className="flex-1 flex overflow-hidden">
            
            {/* LEFT SIDEBAR: BİRİM KLASÖRLERİ */}
            <div id="tech-pubs-birim-sidebar" className="w-80 bg-[#092b17] border-r border-[#0f4624] flex flex-col shrink-0">
              
              {/* Sidebar Header */}
              <div className="px-5 py-3.5 border-b border-[#0f4624] flex items-center justify-between text-white">
                <span className="text-xs font-black tracking-wider uppercase flex items-center gap-2">
                  <span>📁 BİRİM KLASÖRLERİ</span>
                </span>
              </div>

              {/* Aircraft Folder Items List */}
              <div className="flex-1 overflow-y-auto p-3 space-y-1.5">
                {UNIT_FOLDER_OPTIONS.map(folder => {
                  const isSelected = selectedUnitFilter === folder.key;
                  const isUnlocked = unlockedUnits.has(folder.key);
                  const count = publications.filter(p => (p.unit || p.unitKey) === folder.key).length;

                  return (
                    <button
                      key={folder.key}
                      id={`btn-folder-${folder.key}`}
                      onClick={() => handleSelectUnitFolder(folder.key)}
                      className={`w-full p-2.5 rounded-2xl transition-all flex items-center justify-between text-left group cursor-pointer ${
                        isSelected
                          ? 'bg-[#0f4624] border border-emerald-500/60 shadow-md text-white'
                          : 'bg-[#0b331b]/60 hover:bg-[#0d3b1f] border border-transparent text-slate-200'
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        {/* Folder Badge / Tag Icon */}
                        {folder.tag ? (
                          <div className={`w-9 h-9 rounded-xl ${folder.badgeBg} font-black text-xs flex items-center justify-center font-mono shrink-0 shadow-sm`}>
                            {folder.tag}
                          </div>
                        ) : (
                          <div className="w-9 h-9 rounded-xl bg-emerald-500/20 text-emerald-300 flex items-center justify-center shrink-0 border border-emerald-500/30">
                            <Layers className="w-5 h-5" />
                          </div>
                        )}

                        <div className="min-w-0">
                          <div className="flex items-center gap-1.5">
                            <span className="text-xs font-black text-white group-hover:text-emerald-300 transition-colors">
                              {folder.label}
                            </span>
                            {folder.type && (
                              <span className="text-[9px] font-black px-1.5 py-0.2 rounded bg-slate-900/60 text-slate-300 border border-slate-700/60">
                                {folder.type}
                              </span>
                            )}
                          </div>
                          <p className="text-[10px] text-emerald-200/60 truncate max-w-[145px]">
                            {folder.sub}
                          </p>
                        </div>
                      </div>

                      {/* Count Badge & Lock Status */}
                      <div className="flex items-center gap-1.5 shrink-0">
                        <span className={`px-2 py-0.5 rounded-full text-xs font-black font-mono shrink-0 ${
                          isSelected ? 'bg-emerald-500 text-slate-950' : 'bg-slate-900/80 text-emerald-300'
                        }`}>
                          {count}
                        </span>
                        {isUnlocked ? (
                          <Unlock className="w-3.5 h-3.5 text-emerald-400 opacity-80" />
                        ) : (
                          <Lock className="w-3.5 h-3.5 text-amber-400 opacity-80" />
                        )}
                      </div>
                    </button>
                  );
                })}
              </div>

              {/* Bottom Usage Hint Card (Image 1) */}
              <div className="p-3 m-3 rounded-2xl bg-[#072413] border border-emerald-500/20 text-emerald-200/90 text-[11px] leading-relaxed">
                <div className="flex items-start gap-2">
                  <span className="text-emerald-400 font-bold text-sm">ℹ️</span>
                  <div>
                    <strong className="text-white font-bold">Kullanım:</strong> Dokümana çift tıklayarak PDF okuyucusunu açabilir; metin veya P/N seçerek depoda arayabilirsiniz.
                  </div>
                </div>
              </div>

            </div>

            {/* RIGHT MAIN CONTENT: SEARCH & DOCUMENT CARDS (IMAGE 1) */}
            <div id="tech-pubs-main-catalog-pane" className="flex-1 bg-[#f4f7f5] text-slate-900 flex flex-col overflow-hidden">
              
              {/* Search & Category Filter Bar */}
              <div className="p-4 bg-white border-b border-slate-200 flex items-center justify-between gap-4 shrink-0 flex-wrap shadow-xs">
                {/* Search input */}
                <div className="relative flex-1 min-w-[280px] max-w-lg">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                  <input
                    id="input-catalog-search"
                    type="text"
                    value={pubSearchQuery}
                    onChange={e => setPubSearchQuery(e.target.value)}
                    placeholder="Yayın adı, dosya adı veya revizyon ara..."
                    className="w-full bg-slate-50 border border-slate-300 rounded-full pl-10 pr-8 py-2 text-xs font-medium text-slate-800 placeholder-slate-400 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-500/10 transition-all"
                  />
                  {pubSearchQuery && (
                    <button
                      onClick={() => setPubSearchQuery('')}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  )}
                </div>

                {/* Category Pills (Image 1) */}
                <div className="flex items-center gap-1.5 flex-wrap">
                  {CATEGORY_TABS.map(tab => (
                    <button
                      key={tab}
                      onClick={() => {
                        setSelectedCategoryFilter(tab);
                        setPubSearchQuery('');
                      }}
                      className={`px-3 py-1.5 rounded-full text-xs font-black transition-all cursor-pointer ${
                        selectedCategoryFilter === tab
                          ? 'bg-[#153422] text-white shadow-sm'
                          : 'bg-slate-100 hover:bg-slate-200 text-slate-600'
                      }`}
                    >
                      {tab}
                    </button>
                  ))}
                </div>
              </div>

              {/* Catalog Section Header */}
              <div className="px-6 py-3 bg-slate-100/80 border-b border-slate-200 flex items-center justify-between gap-4 shrink-0">
                <div className="flex items-center gap-2.5">
                  <span className="text-xs font-black text-slate-800 uppercase tracking-wider">
                    {UNIT_FOLDER_OPTIONS.find(u => u.key === selectedUnitFilter)?.label || 'AT-802F'} TEKNİK YAYINLARI
                  </span>
                  <span className="px-2 py-0.5 rounded-md bg-emerald-100 text-emerald-800 text-[10px] font-black uppercase">
                    {filteredPublications.length} DOKÜMAN
                  </span>
                </div>
                <span className="text-[11px] text-slate-500 italic">
                  * Görüntülemek için dokümana çift tıklayınız
                </span>
              </div>

              {/* Publications Cards List (Image 1) */}
              <div className="flex-1 overflow-y-auto p-6 space-y-3">
                {/* AT-802F OGM BULUT AÇIK BAĞLANTI KAYNAĞI SATIRI */}
                {selectedUnitFilter === 'at802' && (
                  <div
                    id="pub-card-ogm-bulut-at802"
                    onDoubleClick={handleOpenOgmBulutFolderInViewer}
                    onClick={handleOpenOgmBulutFolderInViewer}
                    className="p-4 rounded-2xl bg-gradient-to-r from-[#0b3d1d] to-[#12542a] border-2 border-emerald-500/70 hover:border-amber-400 hover:shadow-lg transition-all flex items-center justify-between gap-4 group cursor-pointer text-white"
                  >
                    <div className="flex items-center gap-4 min-w-0">
                      <div className="relative shrink-0">
                        <div className="w-12 h-12 rounded-2xl bg-emerald-400/20 border border-emerald-300/40 text-emerald-200 flex items-center justify-center shadow-sm">
                          <Globe className="w-6 h-6" />
                        </div>
                        <div className="absolute -bottom-1 -right-1 w-5 h-5 rounded-full bg-amber-400 text-slate-950 flex items-center justify-center shadow-xs">
                          <Link2 className="w-3 h-3" />
                        </div>
                      </div>

                      <div className="min-w-0">
                        <div className="flex items-center gap-1.5 mb-1 flex-wrap">
                          <span className="px-2 py-0.5 rounded-md bg-amber-500 text-slate-950 font-black text-[10px] uppercase">
                            AT-802F
                          </span>
                          <span className="px-2 py-0.5 rounded-md bg-emerald-400/20 text-emerald-200 font-black text-[10px] uppercase border border-emerald-300/40">
                            AÇIK BAĞLANTI KAYNAĞI
                          </span>
                          <span className="px-2 py-0.5 rounded-md bg-sky-500/20 text-sky-200 font-bold text-[10px] uppercase border border-sky-400/40">
                            OGM BULUT CANLI ARŞİV
                          </span>
                        </div>
                        <h3 className="text-sm font-black text-white group-hover:text-amber-300 transition-colors truncate">
                          AT-802F OGM BULUT TEKNİK YAYINLARI (CANLI KLASÖR)
                        </h3>
                        <p className="text-[11px] text-emerald-200/80 font-mono truncate">
                          {OGM_BULUT_AT802_URL}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleOpenOgmBulutFolderInViewer();
                        }}
                        className="px-4 py-2 rounded-full bg-amber-500 hover:bg-amber-400 text-slate-950 text-xs font-black transition-all flex items-center gap-1.5 shadow-md cursor-pointer active:scale-95"
                      >
                        <Eye className="w-4 h-4" />
                        <span>KLASÖRÜ AÇ & ARA</span>
                      </button>
                    </div>
                  </div>
                )}

                {filteredPublications.length === 0 ? (
                  <div className="h-full flex flex-col items-center justify-center text-center p-8 text-slate-500">
                    <div className="w-16 h-16 rounded-3xl bg-slate-200 flex items-center justify-center text-slate-400 mb-3">
                      <Archive className="w-8 h-8" />
                    </div>
                    <h3 className="text-sm font-bold text-slate-700 mb-1">Döküman Bulunamadı</h3>
                    <p className="text-xs text-slate-500 max-w-sm mb-4 leading-relaxed">
                      Seçili filtre veya arama kriterine uygun teknik yayın bulunmuyor. Yeni döküman eklemek için yukarıdaki butonu kullanabilirsiniz.
                    </p>
                    <button
                      onClick={() => setIsUploadModalOpen(true)}
                      className="px-5 py-2.5 rounded-full bg-[#009b4d] hover:bg-[#00b359] text-white text-xs font-black transition-all flex items-center gap-2 shadow-md"
                    >
                      <Plus className="w-4 h-4" />
                      <span>+ PDF / Yayın Yükle</span>
                    </button>
                  </div>
                ) : (
                  filteredPublications.map(pub => (
                    <div
                      key={pub.id}
                      id={`pub-card-${pub.id}`}
                      onDoubleClick={() => openPublicationInViewer(pub)}
                      className="p-4 rounded-2xl bg-white border border-slate-200 hover:border-emerald-500 hover:shadow-md transition-all flex items-center justify-between gap-4 group cursor-pointer"
                    >
                      {/* Left: Book Icon with Tool Tag */}
                      <div className="flex items-center gap-4 min-w-0">
                        <div className="relative shrink-0">
                          <div className="w-12 h-12 rounded-2xl bg-[#0b3d1d] text-emerald-300 flex items-center justify-center shadow-sm">
                            <BookOpen className="w-6 h-6" />
                          </div>
                          <div className="absolute -bottom-1 -right-1 w-5 h-5 rounded-full bg-amber-500 text-slate-950 flex items-center justify-center shadow-xs">
                            <Wrench className="w-3 h-3" />
                          </div>
                        </div>

                        {/* Middle: Badges, Title & File Info */}
                        <div className="min-w-0">
                          <div className="flex items-center gap-1.5 mb-1 flex-wrap">
                            <span className="px-2 py-0.5 rounded-md bg-amber-600 text-white font-black text-[10px] uppercase">
                              {pub.unitLabel}
                            </span>
                            <span className="px-2 py-0.5 rounded-md bg-emerald-100 text-emerald-800 font-black text-[10px] uppercase border border-emerald-200">
                              {pub.category || 'IPC'}
                            </span>
                            <span className="px-2 py-0.5 rounded-md bg-slate-100 text-slate-600 font-bold text-[10px] uppercase border border-slate-200">
                              {pub.revision || 'Rev. 01'}
                            </span>
                          </div>

                          <h3 className="text-sm font-black text-slate-900 group-hover:text-emerald-800 transition-colors truncate">
                            {pub.title}
                          </h3>

                          <p className="text-[11px] text-slate-500 font-medium">
                            {pub.fileName} &bull; {pub.fileSize || '11.1 MB'} &bull; Yükleme: {pub.uploadDate || '2026-08-20'}
                          </p>
                        </div>
                      </div>

                      {/* Right: Actions */}
                      <div className="flex items-center gap-2 shrink-0">
                        {/* GÖRÜNTÜLE & ARA button (Image 1) */}
                        <button
                          id={`btn-view-pub-${pub.id}`}
                          onClick={() => openPublicationInViewer(pub)}
                          className="px-4 py-2 rounded-full bg-[#0b3d1d] hover:bg-[#0f4d25] text-white text-xs font-black transition-all flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-95"
                        >
                          <Eye className="w-4 h-4" />
                          <span>👁️ GÖRÜNTÜLE & ARA</span>
                        </button>

                        {/* Edit Button (Password protected) */}
                        <button
                          id={`btn-edit-pub-${pub.id}`}
                          onClick={(e) => {
                            e.stopPropagation();
                            requestPasswordAuth(pub, 'edit');
                          }}
                          className="p-2 rounded-full hover:bg-slate-100 text-slate-400 hover:text-amber-600 transition-colors"
                          title="Düzenle (Şifreli)"
                        >
                          <Edit3 className="w-4 h-4" />
                        </button>

                        {/* Delete Button (Password protected) */}
                        <button
                          id={`btn-delete-pub-${pub.id}`}
                          onClick={(e) => {
                            e.stopPropagation();
                            requestPasswordAuth(pub, 'delete');
                          }}
                          className="p-2 rounded-full hover:bg-slate-100 text-slate-400 hover:text-rose-600 transition-colors"
                          title="Sil (Şifreli)"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>

            </div>

          </div>

        </div>
      )}

      {/* ========================================================================= */}
      {/* GLOBAL DEPO SORGUSU MODAL                                                 */}
      {/* ========================================================================= */}
      {isGlobalDepoModalOpen && (
        <div className="fixed inset-0 z-[10001] bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl w-full max-w-4xl max-h-[85vh] flex flex-col shadow-2xl overflow-hidden">
            <div className="px-6 py-4 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-amber-500/20 text-amber-300 border border-amber-500/40 flex items-center justify-center">
                  <Archive className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-white uppercase">TÜM DEPOLARDA CANLI STOK SORGUSU</h3>
                  <p className="text-xs text-slate-400">Tüm filo ve hangar depolarında anlık parça / seri numarası taraması</p>
                </div>
              </div>
              <button
                onClick={() => setIsGlobalDepoModalOpen(false)}
                className="p-2 rounded-xl text-slate-400 hover:text-white hover:bg-slate-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-4 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center gap-3 relative z-30">
              <div ref={globalSearchWrapperRef} className="relative flex-1 min-w-[260px]">
                <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={depoSearchQuery}
                  onFocus={() => {
                    if (depoSearchQuery.trim().length > 0) setIsDepoAutocompleteOpen(true);
                  }}
                  onChange={e => {
                    const val = e.target.value;
                    setDepoSearchQuery(val);
                    setAutocompleteActiveIdx(-1);
                    setIsDepoAutocompleteOpen(val.trim().length > 0);
                  }}
                  placeholder="Parça No (P/N), Seri No veya Malzeme Adı giriniz..."
                  className="w-full bg-slate-950 border border-slate-700 rounded-2xl pl-10 pr-4 py-2.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-emerald-500"
                  autoFocus
                  autoComplete="off"
                />
                {isDepoAutocompleteOpen && depoSearchQuery.trim().length > 0 && autocompleteSuggestions.length > 0 && (
                  <div className="absolute left-0 right-0 top-full mt-1.5 bg-slate-900 border border-emerald-500/40 rounded-2xl shadow-2xl overflow-hidden z-[100] max-h-64 overflow-y-auto">
                    {autocompleteSuggestions.map((item, idx) => {
                      const qty = item.toplamStok ?? item.miktarNum ?? 0;
                      return (
                        <button
                          key={`global_${item.unitKey}_${item.parcaNo}_${idx}`}
                          type="button"
                          onClick={() => handleSelectAutocompleteItem(item)}
                          className="w-full px-3.5 py-2 text-left border-b border-slate-800/80 last:border-b-0 hover:bg-slate-800/90 transition-colors flex items-center justify-between gap-3 cursor-pointer"
                        >
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-1.5">
                              <span className="px-1.5 py-0.5 rounded bg-emerald-500/15 text-emerald-300 text-[9px] font-black uppercase">
                                {item.unitTitle}
                              </span>
                              <span className="text-xs font-bold text-white truncate">{item.adi || item.parcaNo}</span>
                            </div>
                            <div className="text-[10px] font-mono text-slate-400">
                              P/N: <strong className="text-emerald-300">{item.parcaNo}</strong> &bull; Konum: {item.yer}
                            </div>
                          </div>
                          <span className="px-2 py-0.5 rounded-lg text-[10px] font-black font-mono bg-slate-950 text-emerald-300 border border-emerald-700/60">
                            Stok: {qty}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Status Badge */}
              <div className={`px-3 py-2 rounded-2xl text-xs font-bold flex items-center gap-1.5 border shadow-sm ${
                stockSearchStatus === 'in_stock'
                  ? 'bg-emerald-950/80 border-emerald-500/80 text-emerald-300'
                  : stockSearchStatus === 'zero_stock'
                  ? 'bg-amber-950/80 border-amber-500/80 text-amber-300'
                  : stockSearchStatus === 'not_found'
                  ? 'bg-rose-950/80 border-rose-500/80 text-rose-300'
                  : 'bg-slate-800 border-slate-700 text-slate-300'
              }`}>
                <Package className={`w-3.5 h-3.5 ${
                  stockSearchStatus === 'in_stock'
                    ? 'text-emerald-400'
                    : stockSearchStatus === 'zero_stock'
                    ? 'text-amber-400'
                    : stockSearchStatus === 'not_found'
                    ? 'text-rose-400'
                    : 'text-slate-400'
                }`} />
                <span>
                  {stockSearchStatus === 'in_stock'
                    ? `${matchedDepotItems.length} Stok Kaydı (Mevcut: ${matchedTotalStockQuantity})`
                    : stockSearchStatus === 'zero_stock'
                    ? `${matchedDepotItems.length} Stok Kaydı (Stok: 0)`
                    : stockSearchStatus === 'not_found'
                    ? '0 Stok Kaydı (Kayıt Yok)'
                    : 'Stok Kaydı'}
                </span>
              </div>

              {/* Zero Report Button (Appears only when eligible: 0 stock or not found) */}
              {isZeroReportEligible && (
                <button
                  type="button"
                  onClick={handleOpenZeroReport}
                  className="px-3 py-2 rounded-2xl bg-amber-500 hover:bg-amber-400 active:scale-95 text-slate-950 font-black text-xs transition-all flex items-center gap-1.5 shadow-md animate-pulse cursor-pointer shrink-0"
                  title="Depoda bulunmayan veya stoğu 0 olan parça için Zero / İhtiyaç Raporu oluştur"
                >
                  <AlertTriangle className="w-3.5 h-3.5" />
                  <span>ZERO REPORT OLUŞTUR</span>
                </button>
              )}
            </div>

            <div className="flex-1 overflow-y-auto p-4">
              {matchedDepotItems.length === 0 ? (
                <div className="text-center p-8 text-slate-400">
                  <Package className="w-10 h-10 text-slate-600 mx-auto mb-2" />
                  <p className="text-xs font-semibold text-slate-300">
                    {depoSearchQuery.trim() ? `'${depoSearchQuery}' ile eşleşen parça bulunamadı.` : 'Aramak istediğiniz parça veya seri numarasını yazınız.'}
                  </p>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {matchedDepotItems.map((item, i) => (
                    <div key={i} className="p-3 rounded-2xl bg-slate-950 border border-slate-800">
                      <div className="flex justify-between items-start gap-2 mb-1">
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                          {item.unitTitle}
                        </span>
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-slate-800 text-white">
                          Stok: {item.miktar}
                        </span>
                      </div>
                      <h4 className="text-xs font-bold text-white mb-1">{item.adi}</h4>
                      <div className="text-[11px] text-slate-400 space-y-0.5">
                        <div><strong className="text-slate-300">P/N:</strong> {item.parcaNo}</div>
                        <div><strong className="text-slate-300">S/N:</strong> {item.seriNo}</div>
                        
                        {item.isSarfItem ? (
                          <div className="mt-2 pt-2 border-t border-slate-800 space-y-1">
                            <div className="flex justify-between items-center text-orange-300 font-bold">
                                <span>ANKARA:</span>
                                <span>{item.ankaraMevcut || 0} ({item.yer})</span>
                            </div>
                            <div className="flex justify-between items-center">
                                <span>MİLAS:</span>
                                <span>{item.milasMevcut || 0}</span>
                            </div>
                            <div className="flex justify-between items-center">
                                <span>KARAİN:</span>
                                <span>{item.karainMevcut || 0}</span>
                            </div>
                            <div className="flex justify-between items-center">
                                <span>ÇANAKKALE:</span>
                                <span>{item.canakkaleMevcut || 0}</span>
                            </div>
                            <div className="flex justify-between items-center">
                                <span>BURSA:</span>
                                <span>{item.bursaMevcut || 0}</span>
                            </div>
                          </div>
                        ) : (
                          <div><strong className="text-slate-300">Konum:</strong> {item.yer}</div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* UPLOAD MODAL (+ PDF / YAYIN YÜKLE)                                        */}
      {/* ========================================================================= */}
      {isUploadModalOpen && (
        <div className="fixed inset-0 z-[10002] bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl w-full max-w-2xl max-h-[90vh] flex flex-col shadow-2xl overflow-hidden text-slate-100">
            <div className="px-6 py-4 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 flex items-center justify-center">
                  <Upload className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-white uppercase">+ TEKNİK YAYIN & DÖKÜMAN YÜKLE</h3>
                  <p className="text-xs text-slate-400">PDF kataloglarını, IPC ve AMM dökümanlarını sisteme yükleyin</p>
                </div>
              </div>
              <button
                onClick={() => setIsUploadModalOpen(false)}
                className="text-slate-400 hover:text-white p-2 rounded-xl hover:bg-slate-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 overflow-y-auto space-y-5 flex-1">
              {/* Unit & Category Selector */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1">Hedef Hava Aracı</label>
                  <select
                    value={uploadUnit}
                    onChange={e => setUploadUnit(e.target.value)}
                    disabled={isUploading}
                    className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none cursor-pointer disabled:opacity-50"
                  >
                    {UNIT_FOLDER_OPTIONS.filter(u => u.key !== 'all').map(u => (
                      <option key={u.key} value={u.key}>{u.label}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1">Yayın Türü / Kategori</label>
                  <select
                    value={uploadCategory}
                    onChange={e => setUploadCategory(e.target.value)}
                    disabled={isUploading}
                    className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none cursor-pointer disabled:opacity-50"
                  >
                    <option value="IPC">IPC (Parça Kataloğu)</option>
                    <option value="AMM">AMM (Bakım El Kitabı)</option>
                    <option value="CMM">CMM (Komponent Bakım)</option>
                    <option value="ŞEMA">ŞEMA / WDM</option>
                    <option value="EL KİTABI">EL KİTABI</option>
                    <option value="STANDART">STANDART</option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1">Rev. Tarihi</label>
                    <input
                      type="date"
                      value={uploadRevisionDate}
                      disabled={isUploading}
                      onChange={e => {
                        const date = e.target.value;
                        setUploadRevisionDate(date);
                        if (date) {
                          const d = new Date(date);
                          const revLabel = `Rev. ${d.toLocaleDateString('tr-TR')}`;
                          setUploadRevision(revLabel);
                          setSelectedFiles(prev => prev.map(f => ({ ...f, customRevision: revLabel })));
                        }
                      }}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl px-2 py-2 text-[10px] text-white outline-none cursor-pointer disabled:opacity-50"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1">Etiket</label>
                    <input
                      type="text"
                      value={uploadRevision}
                      disabled={isUploading}
                      onChange={e => {
                        const val = e.target.value;
                        setUploadRevision(val);
                        setSelectedFiles(prev => prev.map(f => ({ ...f, customRevision: val })));
                      }}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl px-2 py-2 text-[10px] text-white outline-none disabled:opacity-50"
                    />
                  </div>
                </div>
              </div>

              {/* Section selector (Folder Name) */}
              <div className="p-4 bg-slate-950/50 border border-slate-800 rounded-2xl space-y-3">
                <div>
                  <label className="block text-xs font-bold text-emerald-400 mb-1.5 uppercase tracking-wider flex items-center gap-2">
                    <Layers className="w-3.5 h-3.5" />
                    Klasör / Bölüm Adı
                  </label>
                  <select
                    value={uploadSection}
                    disabled={isUploading}
                    onChange={e => {
                      const val = e.target.value;
                      setUploadSection(val);
                      if (val !== 'custom') {
                        setUploadCustomSection('');
                        setSelectedFiles(prev => prev.map(f => ({ ...f, customSection: val })));
                      }
                    }}
                    className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none cursor-pointer disabled:opacity-50"
                  >
                    {SECTION_SUGGESTIONS.map(sec => (
                      <option key={sec} value={sec}>{sec}</option>
                    ))}
                    <option value="custom">+ Yeni Klasör / Özel Bölüm</option>
                  </select>
                </div>

                {(uploadSection === 'custom' || uploadSection === '') && (
                  <motion.div initial={{ opacity: 0, y: -10 }} animate={{ opacity: 1, y: 0 }}>
                    <label className="block text-xs font-semibold text-slate-400 mb-1">Klasör Adı Girin</label>
                    <input
                      type="text"
                      value={uploadCustomSection}
                      disabled={isUploading}
                      onChange={e => {
                        const val = e.target.value;
                        setUploadCustomSection(val);
                        setSelectedFiles(prev => prev.map(f => ({ ...f, customSection: val })));
                      }}
                      placeholder="Örn: PART CATALOG OCAK 2026"
                      className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none disabled:opacity-50"
                    />
                  </motion.div>
                )}
              </div>

              {/* File Dropzone */}
              {!isUploading && (
                <div className="border-2 border-dashed border-slate-700/80 hover:border-emerald-500/60 rounded-2xl p-6 text-center bg-slate-950/40 transition-colors relative cursor-pointer group">
                  <input
                    type="file"
                    multiple
                    accept=".pdf,application/pdf,.zip"
                    onChange={handleFileInputChange}
                    className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                  />
                  <div className="w-12 h-12 rounded-2xl bg-slate-900 border border-slate-800 flex items-center justify-center text-emerald-400 mx-auto mb-2 group-hover:scale-105 transition-transform">
                    <FileText className="w-6 h-6" />
                  </div>
                  <h4 className="text-xs font-bold text-white mb-1">PDF / ZIP Dökümanlarını Seçiniz veya Sürükleyiniz</h4>
                  <p className="text-[11px] text-slate-400">ZIP dosyaları otomatik olarak dökümanlara ayrılacaktır</p>
                </div>
              )}

              {/* Selected files list with Individual Progress Tracking */}
              {selectedFiles.length > 0 && (
                <div className="space-y-2">
                  <div className="flex items-center justify-between text-xs text-slate-400 px-1">
                    <div className="flex items-center gap-2">
                      <span className="font-bold">Seçilen Yayınlar ({selectedFiles.length})</span>
                      {isUploading && (
                        <span className="text-emerald-400 animate-pulse text-[10px] font-mono">
                          ({successCount + failCount}/{selectedFiles.length} Tamamlandı)
                        </span>
                      )}
                    </div>
                    {!isUploading && (
                      <button
                        onClick={() => setSelectedFiles([])}
                        className="text-rose-400 hover:text-rose-300 text-[11px]"
                      >
                        Listeyi Temizle
                      </button>
                    )}
                  </div>
                  <div className="space-y-2 max-h-64 overflow-y-auto pr-1">
                    {selectedFiles.map((item, idx) => (
                      <div 
                        key={idx} 
                        className={`p-3 rounded-xl border transition-all ${
                          item.status === 'uploading' ? 'bg-emerald-500/10 border-emerald-500/50 ring-1 ring-emerald-500/20' :
                          item.status === 'success' ? 'bg-slate-950/40 border-emerald-500/30 opacity-70' :
                          item.status === 'error' ? 'bg-rose-500/10 border-rose-500/50' :
                          'bg-slate-950 border-slate-800'
                        }`}
                      >
                        <div className="flex items-center gap-3">
                          {item.status === 'success' ? (
                            <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
                          ) : item.status === 'error' ? (
                            <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0" />
                          ) : item.status === 'uploading' ? (
                            <RefreshCw className="w-4 h-4 text-emerald-400 shrink-0 animate-spin" />
                          ) : (
                            <FileText className="w-4 h-4 text-slate-400 shrink-0" />
                          )}
                          
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center justify-between gap-2 mb-1">
                              <span className="text-xs text-white font-medium truncate">
                                {item.customTitle}
                              </span>
                              {item.progress !== undefined && (item.status === 'uploading' || item.status === 'success') && (
                                <span className={`text-[10px] font-mono font-bold ${item.status === 'success' ? 'text-emerald-400' : 'text-slate-400'}`}>
                                  %{item.progress}
                                </span>
                              )}
                            </div>
                            
                            {/* Individual Progress Bar */}
                            {(item.status === 'uploading' || item.status === 'success' || item.status === 'error') && (
                              <div className="w-full h-1 bg-slate-900 rounded-full overflow-hidden">
                                <motion.div 
                                  className={`h-full ${item.status === 'error' ? 'bg-rose-500' : 'bg-emerald-500'}`}
                                  initial={{ width: 0 }}
                                  animate={{ width: `${item.progress || 0}%` }}
                                  transition={{ duration: 0.3 }}
                                />
                              </div>
                            )}
                          </div>

                          {!isUploading && item.status !== 'success' && (
                            <button
                              onClick={() => setSelectedFiles(prev => prev.filter((_, i) => i !== idx))}
                              className="p-1.5 rounded-lg text-slate-500 hover:text-rose-400 hover:bg-rose-400/10"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            <div className="px-6 py-4 border-t border-slate-800 bg-slate-950 flex items-center justify-end gap-3">
              <button
                onClick={() => setIsUploadModalOpen(false)}
                disabled={isUploading}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold"
              >
                İptal
              </button>
              <button
                id="btn-confirm-upload"
                onClick={handlePerformUpload}
                disabled={isUploading || selectedFiles.length === 0}
                className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white text-xs font-bold flex items-center gap-2 shadow-lg"
              >
                {isUploading ? (
                  <>
                    <Sparkles className="w-4 h-4 animate-spin" />
                    <span>Yükleniyor...</span>
                  </>
                ) : (
                  <>
                    <Upload className="w-4 h-4" />
                    <span>Yayınları Kaydet</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* PASSWORD AUTH MODAL                                                       */}
      {/* ========================================================================= */}
      {passwordModalOpen && passwordTargetPub && (
        <div className="fixed inset-0 z-[10003] bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl w-full max-w-sm p-6 shadow-2xl text-center text-slate-100">
            <div className="w-12 h-12 rounded-2xl bg-amber-500/20 text-amber-400 border border-amber-500/40 flex items-center justify-center mx-auto mb-3">
              <Lock className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-black text-white uppercase mb-1">
              GÜVENLİK ŞİFRESİ GİRİNİZ
            </h3>
            <p className="text-xs text-slate-400 mb-4 leading-relaxed">
              <strong className="text-white">{passwordTargetPub.unitLabel}</strong> birimine ait yayını {passwordTargetAction === 'delete' ? 'silmek' : 'düzenlemek'} için şifreyi giriniz.
            </p>
            <input
              type="password"
              value={enteredPassword}
              onChange={e => setEnteredPassword(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handlePasswordSubmit()}
              placeholder="Şifre"
              className="w-full bg-slate-950 border border-slate-700 rounded-2xl px-4 py-2.5 text-center text-sm font-mono text-white focus:border-emerald-500 outline-none mb-2 tracking-widest"
              autoFocus
            />
            {passwordError && (
              <p className="text-xs text-rose-400 font-bold mb-3">{passwordError}</p>
            )}
            <div className="flex items-center gap-2 mt-4">
              <button
                onClick={() => setPasswordModalOpen(false)}
                className="flex-1 px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-300"
              >
                İptal
              </button>
              <button
                onClick={handlePasswordSubmit}
                className="flex-1 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white shadow-md"
              >
                Doğrula & Devam Et
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* UNIT ACCESS PIN / PASSWORD MODAL & INITIAL UNIT SELECTOR                   */}
      {/* ========================================================================= */}
      {(!selectedUnitFilter || !unlockedUnits.has(selectedUnitFilter)) && !isViewerActive && (
        <div className="fixed inset-0 z-[10002] bg-slate-950/85 backdrop-blur-md flex items-center justify-center p-4 sm:p-6 animate-in fade-in duration-200">
          <div className="bg-gradient-to-b from-[#092b17] to-[#04150b] border border-emerald-500/40 rounded-3xl w-full max-w-3xl p-6 sm:p-8 shadow-2xl text-slate-100 flex flex-col relative overflow-hidden">
            {/* Header */}
            <div className="flex items-center justify-between pb-4 border-b border-emerald-500/20 mb-6">
              <div className="flex items-center gap-3">
                <div className="w-12 h-12 rounded-2xl bg-emerald-500/20 border border-emerald-400/40 text-emerald-300 flex items-center justify-center shadow-inner">
                  <BookOpen className="w-6 h-6" />
                </div>
                <div>
                  <h2 className="text-base sm:text-lg font-black text-white tracking-wide flex items-center gap-2 uppercase">
                    <span>TEKNİK YAYINLAR &bull; BİRİM SEÇİNİZ</span>
                  </h2>
                  <p className="text-xs text-emerald-300/80 font-medium">
                    İncelemek istediğiniz hava aracını veya birimi seçiniz ve birim yetkili şifresini giriniz.
                  </p>
                </div>
              </div>
              <button
                onClick={onClose}
                className="p-2 rounded-full bg-emerald-950/80 hover:bg-emerald-900 text-emerald-300 hover:text-white transition-all border border-emerald-800 cursor-pointer"
                title="Kapat"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Units Grid */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3.5 mb-6">
              {UNIT_FOLDER_OPTIONS.map(folder => {
                const isSelected = targetUnitToUnlock === folder.key;
                const isUnlocked = unlockedUnits.has(folder.key);
                const count = publications.filter(p => (p.unit || p.unitKey) === folder.key).length;

                return (
                  <button
                    key={folder.key}
                    type="button"
                    onClick={() => {
                      if (isUnlocked) {
                        setSelectedUnitFilter(folder.key);
                        setPubSearchQuery('');
                        setDepoSearchQuery('');
                      } else {
                        setTargetUnitToUnlock(folder.key);
                        setEnteredUnitPassword('');
                        setUnitPasswordError('');
                      }
                    }}
                    className={`p-4 rounded-2xl border text-left transition-all relative flex flex-col justify-between cursor-pointer ${
                      isSelected
                        ? 'bg-emerald-800/80 border-emerald-400 ring-2 ring-emerald-400/50 shadow-xl'
                        : isUnlocked
                        ? 'bg-[#083019] hover:bg-[#0c4022] border-emerald-500/40 text-slate-200'
                        : 'bg-[#062212]/90 hover:bg-[#0a331c] border-emerald-900/60 text-slate-300'
                    }`}
                  >
                    <div className="flex items-center justify-between mb-3">
                      <div className={`w-10 h-10 rounded-xl ${folder.badgeBg} font-black text-xs flex items-center justify-center font-mono shadow-sm`}>
                        {folder.tag}
                      </div>
                      <div className="flex items-center gap-1.5">
                        <span className="px-2 py-0.5 rounded-full text-[11px] font-black font-mono bg-slate-950/80 text-emerald-300 border border-emerald-500/30">
                          {count} PDF
                        </span>
                        {isUnlocked ? (
                          <Unlock className="w-4 h-4 text-emerald-400" />
                        ) : (
                          <Lock className="w-4 h-4 text-amber-400" />
                        )}
                      </div>
                    </div>

                    <div>
                      <div className="flex items-center gap-2">
                        <span className="text-sm font-black text-white">{folder.label}</span>
                        <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-slate-900/70 text-slate-300 border border-slate-700/60">
                          {folder.type}
                        </span>
                      </div>
                      <p className="text-[11px] text-emerald-200/70 mt-1 leading-snug line-clamp-2">
                        {folder.sub}
                      </p>
                    </div>
                  </button>
                );
              })}
            </div>

            {/* Password input section if a unit is selected */}
            {targetUnitToUnlock && (
              <div className="bg-slate-950/90 border border-emerald-500/40 rounded-2xl p-4 sm:p-5 shadow-inner animate-in fade-in zoom-in-95">
                <div className="flex flex-col sm:flex-row items-center gap-3">
                  <div className="flex items-center gap-2.5 text-left w-full sm:w-auto">
                    <div className="w-9 h-9 rounded-xl bg-amber-500/20 text-amber-400 border border-amber-500/40 flex items-center justify-center shrink-0">
                      <Lock className="w-4 h-4" />
                    </div>
                    <div>
                      <span className="text-xs font-black text-white uppercase block">
                        {UNIT_FOLDER_OPTIONS.find(u => u.key === targetUnitToUnlock)?.label} ŞİFRESİ
                      </span>
                      <span className="text-[11px] text-slate-400">
                        Bu birimin teknik yayınlarını açmak için şifreyi giriniz
                      </span>
                    </div>
                  </div>

                  <div className="flex-1 w-full flex items-center gap-2">
                    <input
                      type="password"
                      value={enteredUnitPassword}
                      onChange={e => setEnteredUnitPassword(e.target.value)}
                      onKeyDown={e => e.key === 'Enter' && handleUnitPasswordSubmit()}
                      placeholder="Birim Yetkili Şifresi"
                      className="flex-1 bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-center text-sm font-mono text-white focus:border-emerald-500 outline-none tracking-widest"
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={handleUnitPasswordSubmit}
                      className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-black text-white transition-all shadow-md cursor-pointer active:scale-95 shrink-0"
                    >
                      KİLİDİ AÇ
                    </button>
                  </div>
                </div>
                {unitPasswordError && (
                  <p className="text-xs text-rose-400 font-bold mt-2 text-center sm:text-right">{unitPasswordError}</p>
                )}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* UNIT ACCESS PIN / PASSWORD MODAL (From Sidebar folder click)              */}
      {/* ========================================================================= */}
      {unitPasswordModalOpen && targetUnitToUnlock && selectedUnitFilter && (
        <div className="fixed inset-0 z-[10003] bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl w-full max-w-sm p-6 shadow-2xl text-center text-slate-100 animate-in fade-in zoom-in-95">
            <div className="w-12 h-12 rounded-2xl bg-amber-500/20 text-amber-400 border border-amber-500/40 flex items-center justify-center mx-auto mb-3">
              <Lock className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-black text-white uppercase mb-1">
              BİRİM GİRİŞ ŞİFRESİ
            </h3>
            <p className="text-xs text-slate-400 mb-4 leading-relaxed">
              <strong className="text-white">{UNIT_FOLDER_OPTIONS.find(u => u.key === targetUnitToUnlock)?.label || targetUnitToUnlock}</strong> birimine ait teknik yayınları görüntülemek için yetkili şifreyi giriniz.
            </p>
            <input
              type="password"
              value={enteredUnitPassword}
              onChange={e => setEnteredUnitPassword(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handleUnitPasswordSubmit()}
              placeholder="Birim Şifresi"
              className="w-full bg-slate-950 border border-slate-700 rounded-2xl px-4 py-2.5 text-center text-sm font-mono text-white focus:border-emerald-500 outline-none mb-2 tracking-widest"
              autoFocus
            />
            {unitPasswordError && (
              <p className="text-xs text-rose-400 font-bold mb-3">{unitPasswordError}</p>
            )}
            <div className="flex items-center gap-2 mt-4">
              <button
                onClick={() => {
                  setUnitPasswordModalOpen(false);
                  setEnteredUnitPassword('');
                  setUnitPasswordError('');
                }}
                className="flex-1 px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-300 cursor-pointer"
              >
                İptal
              </button>
              <button
                onClick={handleUnitPasswordSubmit}
                className="flex-1 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white shadow-md cursor-pointer"
              >
                Giriş Yap
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* EDIT MODAL                                                                */}
      {/* ========================================================================= */}
      {editModalOpen && editingPub && (
        <div className="fixed inset-0 z-[10003] bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl w-full max-w-lg p-6 shadow-2xl text-slate-100">
            <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2.5">
                <Edit3 className="w-5 h-5 text-amber-400" />
                <h3 className="text-sm font-black text-white uppercase">YAYIN BİLGİLERİNİ DÜZENLE</h3>
              </div>
              <button onClick={() => setEditModalOpen(false)} className="text-slate-400 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">Yayın Başlığı</label>
                <input
                  type="text"
                  value={editTitle}
                  onChange={e => setEditTitle(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1">Kategori</label>
                  <select
                    value={editCategory}
                    onChange={e => setEditCategory(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none"
                  >
                    <option value="IPC">IPC</option>
                    <option value="AMM">AMM</option>
                    <option value="CMM">CMM</option>
                    <option value="ŞEMA">ŞEMA</option>
                    <option value="EL KİTABI">EL KİTABI</option>
                    <option value="STANDART">STANDART</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1">Revizyon</label>
                  <input
                    type="text"
                    value={editRevision}
                    onChange={e => setEditRevision(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-300 mb-1">Bölüm / Kategori</label>
                <input
                  type="text"
                  value={editSection}
                  onChange={e => setEditSection(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white focus:border-emerald-500 outline-none"
                />
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 mt-6 pt-3 border-t border-slate-800">
              <button
                onClick={() => setEditModalOpen(false)}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-300"
              >
                İptal
              </button>
              <button
                onClick={handleSaveEdit}
                className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white shadow-md"
              >
                Kaydet
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* SCREENSHOT CROP & ANNOTATION STUDIO (STEP 1 FOR ZERO REPORT CATALOG)      */}
      {/* ========================================================================= */}
      {isCropOverlayOpen && rawScreenshotDataUrl && (
        <div className="fixed inset-0 z-[10006] bg-slate-950/95 backdrop-blur-md flex flex-col p-3 sm:p-4 select-none">
          {/* Top Toolbar: Pen, Highlighter, Eraser, Colors, Mouse Wheel Zoom, Crop */}
          <div className="flex items-center justify-between gap-3 flex-wrap px-4 py-2.5 bg-slate-900 border border-slate-700 rounded-2xl mb-3 shrink-0 shadow-xl">
            <div className="flex items-center gap-3 flex-wrap">
              <div className="flex items-center gap-2 pr-3 border-r border-slate-700">
                <div className="w-8 h-8 rounded-xl bg-amber-500/20 text-amber-300 border border-amber-500/40 flex items-center justify-center">
                  <Crop className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-xs font-black text-white uppercase tracking-wide">ADIM 1: GÖRSEL KIRPMA & VURGULAMA</h3>
                  <p className="text-[10px] text-slate-400">Fare tekerleği ile yakınlaştırın, kalem/vurgulayıcı ile parça yerini işaretleyin</p>
                </div>
              </div>

              {/* Tool Buttons: Crop | Pen | Highlighter | Eraser (matching reference icons) */}
              <div className="flex items-center gap-1.5 bg-slate-950 p-1 rounded-xl border border-slate-800">
                <button
                  type="button"
                  onClick={() => setAnnotationTool('crop')}
                  className={`px-2.5 py-1.5 rounded-lg text-[11px] font-black flex items-center gap-1.5 cursor-pointer transition-all ${
                    annotationTool === 'crop'
                      ? 'bg-amber-500 text-slate-950 shadow-sm'
                      : 'text-slate-300 hover:bg-slate-800'
                  }`}
                  title="Kırpma Alanı Seç"
                >
                  <Crop className="w-3.5 h-3.5" />
                  <span>Kırp</span>
                </button>

                {/* Pen Tool (Kalem) */}
                <button
                  type="button"
                  onClick={() => setAnnotationTool('pen')}
                  className={`px-2.5 py-1.5 rounded-lg text-[11px] font-black flex items-center gap-1.5 cursor-pointer transition-all ${
                    annotationTool === 'pen'
                      ? 'bg-sky-500/20 text-white border border-sky-400 shadow-sm'
                      : 'text-slate-300 hover:bg-slate-800 border border-transparent'
                  }`}
                  title="Kalem (Çizim / İşaret)"
                >
                  <svg viewBox="0 0 24 24" className="w-4 h-4" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M12 19l7-7-4-4-7 7-2 6 6-2z" />
                    <path d="M15 8l4 4" stroke={annotationColor} strokeWidth="3" />
                  </svg>
                  <span>Kalem</span>
                </button>

                {/* Highlighter Tool (Fosforlu Vurgulayıcı) */}
                <button
                  type="button"
                  onClick={() => setAnnotationTool('highlighter')}
                  className={`px-2.5 py-1.5 rounded-lg text-[11px] font-black flex items-center gap-1.5 cursor-pointer transition-all ${
                    annotationTool === 'highlighter'
                      ? 'bg-pink-500/25 text-white border border-pink-400 shadow-sm'
                      : 'text-slate-300 hover:bg-slate-800 border border-transparent'
                  }`}
                  title="Fosforlu Vurgulayıcı (Parça No / Şekil Vurgula)"
                >
                  <svg viewBox="0 0 24 24" className="w-4 h-4" fill="none">
                    <rect x="6" y="4" width="12" height="5" rx="1" fill={annotationColor} stroke="currentColor" strokeWidth="1.5" />
                    <path d="M8 9h8l-1 6-5 3-1-3L8 9z" fill={annotationColor} fillOpacity="0.65" stroke="currentColor" strokeWidth="1.5" />
                  </svg>
                  <span>Vurgula</span>
                </button>

                {/* Eraser Tool (Silgi) */}
                <button
                  type="button"
                  onClick={() => setAnnotationTool('eraser')}
                  className={`px-2.5 py-1.5 rounded-lg text-[11px] font-black flex items-center gap-1.5 cursor-pointer transition-all ${
                    annotationTool === 'eraser'
                      ? 'bg-rose-500/25 text-white border border-rose-400 shadow-sm'
                      : 'text-slate-300 hover:bg-slate-800 border border-transparent'
                  }`}
                  title="Silgi (Çizilen İşareti Sil)"
                >
                  <svg viewBox="0 0 24 24" className="w-4 h-4" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M7 21h10" />
                    <path d="M5.5 13.5l8.5-8.5a2 2 0 0 1 2.8 0l2.7 2.7a2 2 0 0 1 0 2.8L11 19H7l-1.5-1.5a2.8 2.8 0 0 1 0-4z" />
                  </svg>
                  <span>Silgi</span>
                </button>
              </div>

              {/* Color Swatches */}
              <div className="flex items-center gap-1.5 bg-slate-950 px-2.5 py-1.5 rounded-xl border border-slate-800">
                {['#ec4899', '#facc15', '#ef4444', '#22c55e', '#38bdf8'].map(c => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => {
                      setAnnotationColor(c);
                      if (annotationTool === 'crop' || annotationTool === 'eraser') {
                        setAnnotationTool('highlighter');
                      }
                    }}
                    className={`w-5 h-5 rounded-full cursor-pointer transition-transform ${
                      annotationColor === c ? 'ring-2 ring-white scale-110' : 'opacity-75 hover:opacity-100'
                    }`}
                    style={{ backgroundColor: c }}
                    title="Vurgulama Rengi"
                  />
                ))}
              </div>

              {/* Zoom Controls + Mouse Wheel Support */}
              <div className="flex items-center gap-1 bg-slate-950 px-2 py-1 rounded-xl border border-slate-800 text-xs font-mono">
                <button
                  type="button"
                  onClick={() => setAnnotationZoom(z => Math.max(0.5, Number((z - 0.25).toFixed(2))))}
                  className="px-2 py-1 rounded hover:bg-slate-800 text-slate-300 font-black cursor-pointer"
                  title="Uzaklaştır"
                >
                  -
                </button>
                <button
                  type="button"
                  onClick={() => setAnnotationZoom(1)}
                  className="px-2 py-1 rounded hover:bg-slate-800 text-emerald-300 font-bold cursor-pointer min-w-[52px] text-center"
                  title="Yakınlaştırmayı Sıfırla (%100)"
                >
                  %{Math.round(annotationZoom * 100)}
                </button>
                <button
                  type="button"
                  onClick={() => setAnnotationZoom(z => Math.min(5, Number((z + 0.25).toFixed(2))))}
                  className="px-2 py-1 rounded hover:bg-slate-800 text-slate-300 font-black cursor-pointer"
                  title="Yakınlaştır"
                >
                  +
                </button>
              </div>

              {/* Undo / Clear Strokes */}
              {annotationStrokes.length > 0 && (
                <div className="flex items-center gap-1">
                  <button
                    type="button"
                    onClick={() => setAnnotationStrokes(prev => prev.slice(0, -1))}
                    className="px-2.5 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-[11px] font-bold text-slate-200 cursor-pointer"
                  >
                    Geri Al
                  </button>
                  <button
                    type="button"
                    onClick={() => setAnnotationStrokes([])}
                    className="px-2.5 py-1.5 rounded-xl bg-slate-800 hover:bg-rose-900/60 text-[11px] font-bold text-rose-300 cursor-pointer"
                  >
                    Temizle
                  </button>
                </div>
              )}
            </div>

            {/* Right Finish / Crop Action Buttons */}
            <div className="flex items-center gap-2 shrink-0">
              {cropBox && Math.abs(cropBox.endX - cropBox.startX) >= 15 && Math.abs(cropBox.endY - cropBox.startY) >= 15 && (
                <button
                  type="button"
                  onClick={handleCropInPlaceAndAnnotate}
                  className="px-3.5 py-2 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 text-xs font-black flex items-center gap-1.5 shadow-md cursor-pointer"
                >
                  <Crop className="w-3.5 h-3.5" />
                  <span>Seçili Alanı Kırp & Vurgula</span>
                </button>
              )}
              <button
                type="button"
                onClick={() => {
                  setIsCropOverlayOpen(false);
                  setRawScreenshotDataUrl(null);
                  setCropBox(null);
                  setAnnotationStrokes([]);
                  setAnnotationZoom(1);
                  if (openZeroReportAfterCropRef.current) {
                    openZeroReportAfterCropRef.current = false;
                    setZeroReportModalOpen(true);
                  }
                }}
                className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-bold text-slate-300 cursor-pointer"
              >
                {openZeroReportAfterCropRef.current ? 'Görselsiz Devam Et' : 'İptal'}
              </button>
              <button
                type="button"
                onClick={handleApplyCrop}
                className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-black text-white flex items-center gap-1.5 shadow-lg cursor-pointer"
              >
                <Check className="w-4 h-4" />
                <span>
                  {openZeroReportAfterCropRef.current
                    ? 'Onayla & Zero Report Formunu Aç'
                    : cropBox
                    ? 'Seçili Alanı Kaydet'
                    : 'Görseli Kaydet'}
                </span>
              </button>
            </div>
          </div>

          {/* Interactive Zoomable Viewport (Mouse Wheel Zoom + Draw / Highlight / Erase / Crop) */}
          <div
            onWheel={e => {
              e.preventDefault();
              const delta = e.deltaY < 0 ? 0.2 : -0.2;
              setAnnotationZoom(z => Math.max(0.5, Math.min(5, Number((z + delta).toFixed(2)))));
            }}
            className="flex-1 overflow-auto bg-slate-900/70 rounded-2xl border border-slate-800 p-4 flex items-start justify-center"
          >
            <div
              className={`relative inline-block select-none my-auto ${
                annotationTool === 'crop'
                  ? 'cursor-crosshair'
                  : annotationTool === 'eraser'
                  ? 'cursor-cell'
                  : 'cursor-crosshair'
              }`}
              style={{ width: `${Math.round(cropBaseWidth * annotationZoom)}px`, maxWidth: 'none' }}
              onMouseDown={e => {
                const rect = e.currentTarget.getBoundingClientRect();
                if (!rect.width || !rect.height) return;
                const px = Math.max(0, Math.min(rect.width, e.clientX - rect.left));
                const py = Math.max(0, Math.min(rect.height, e.clientY - rect.top));
                const nx = px / rect.width;
                const ny = py / rect.height;

                if (annotationTool === 'crop') {
                  setIsCroppingDrag(true);
                  setCropBox({ startX: px, startY: py, endX: px, endY: py });
                  return;
                }

                if (annotationTool === 'eraser') {
                  isDrawingAnnotationRef.current = true;
                  const threshold = 0.025 / Math.max(1, annotationZoom * 0.6);
                  setAnnotationStrokes(prev =>
                    prev.filter(st => !st.points.some(pt => Math.hypot(pt.x - nx, pt.y - ny) < threshold))
                  );
                  return;
                }

                const newId = `st_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
                activeStrokeIdRef.current = newId;
                isDrawingAnnotationRef.current = true;
                const widthNorm = annotationTool === 'highlighter' ? 0.024 : 0.0045;
                setAnnotationStrokes(prev => [
                  ...prev,
                  {
                    id: newId,
                    tool: annotationTool,
                    color: annotationColor,
                    widthNorm,
                    points: [{ x: nx, y: ny }]
                  }
                ]);
              }}
              onMouseMove={e => {
                const rect = e.currentTarget.getBoundingClientRect();
                if (!rect.width || !rect.height) return;
                const px = Math.max(0, Math.min(rect.width, e.clientX - rect.left));
                const py = Math.max(0, Math.min(rect.height, e.clientY - rect.top));
                const nx = px / rect.width;
                const ny = py / rect.height;

                if (annotationTool === 'crop') {
                  if (!isCroppingDrag) return;
                  setCropBox(prev => (prev ? { ...prev, endX: px, endY: py } : null));
                  return;
                }

                if (!isDrawingAnnotationRef.current) return;

                if (annotationTool === 'eraser') {
                  const threshold = 0.025 / Math.max(1, annotationZoom * 0.6);
                  setAnnotationStrokes(prev =>
                    prev.filter(st => !st.points.some(pt => Math.hypot(pt.x - nx, pt.y - ny) < threshold))
                  );
                  return;
                }

                const activeId = activeStrokeIdRef.current;
                if (!activeId) return;
                setAnnotationStrokes(prev =>
                  prev.map(st => (st.id === activeId ? { ...st, points: [...st.points, { x: nx, y: ny }] } : st))
                );
              }}
              onMouseUp={() => {
                setIsCroppingDrag(false);
                isDrawingAnnotationRef.current = false;
                activeStrokeIdRef.current = null;
              }}
              onMouseLeave={() => {
                setIsCroppingDrag(false);
                isDrawingAnnotationRef.current = false;
                activeStrokeIdRef.current = null;
              }}
            >
              <img
                ref={cropImgRef}
                src={rawScreenshotDataUrl}
                alt="Ekran Görüntüsü"
                className="w-full h-auto block rounded-lg shadow-2xl pointer-events-none select-none"
                draggable={false}
              />
              <canvas
                ref={annotationCanvasRef}
                className="absolute inset-0 w-full h-full pointer-events-none rounded-lg"
              />
              {cropBox && (
                <div
                  className="absolute border-2 border-amber-400 bg-amber-400/15 shadow-[0_0_0_9999px_rgba(2,6,23,0.55)] pointer-events-none"
                  style={{
                    left: Math.min(cropBox.startX, cropBox.endX),
                    top: Math.min(cropBox.startY, cropBox.endY),
                    width: Math.abs(cropBox.endX - cropBox.startX),
                    height: Math.abs(cropBox.endY - cropBox.startY)
                  }}
                />
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* ZERO REPORT / HAVA ARACI YEDEK PARÇA TALEP FORMU (EXACT 2-PAGE OGM FORM)  */}
      {/* ========================================================================= */}
      {zeroReportModalOpen && zeroReportData && (
        <div
          className="fixed inset-0 z-[10004] bg-slate-950/90 backdrop-blur-sm flex items-center justify-center p-2 sm:p-4"
          onPaste={e => {
            const items = e.clipboardData?.items;
            if (!items) return;
            for (let i = 0; i < items.length; i++) {
              if (items[i].type.startsWith('image/')) {
                const file = items[i].getAsFile();
                if (file) {
                  const reader = new FileReader();
                  reader.onload = ev => {
                    if (ev.target?.result) {
                      setCapturedCatalogImage(String(ev.target.result));
                      showNotification('📋 Panodaki katalog görseli formun 3. bölümüne eklendi!', 'success');
                    }
                  };
                  reader.readAsDataURL(file);
                }
              }
            }
          }}
        >
          <div className="bg-slate-900 border border-amber-500/40 rounded-3xl w-full max-w-[1180px] h-[94vh] flex flex-col shadow-2xl overflow-hidden text-slate-100">
            {/* Modal Top Toolbar */}
            <div className="px-5 py-3.5 bg-slate-950 border-b border-slate-800 flex items-center justify-between gap-3 flex-wrap shrink-0">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-amber-500/20 text-amber-300 border border-amber-500/40 flex items-center justify-center">
                  <FileText className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-amber-300 uppercase">HAVA ARACI YEDEK PARÇA TALEP FORMU (ZERO REPORT)</h3>
                  <p className="text-[11px] text-slate-400">
                    2 Sayfalık Resmi OGM Talep Formu &bull; Alanlara tıklayarak düzenleyebilir, Katalog Görseli ekleyip Yazdır / Word (.doc) indirebilirsiniz
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2 flex-wrap">
                <button
                  type="button"
                  onClick={handleDownloadZeroReportWord}
                  className="px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-500 text-white text-xs font-black flex items-center gap-1.5 shadow-md cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>Word (.doc) İndir</span>
                </button>
                <button
                  type="button"
                  onClick={handlePrintZeroReport}
                  className="px-4 py-2 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-black text-xs flex items-center gap-1.5 shadow-md cursor-pointer"
                >
                  <Printer className="w-4 h-4" />
                  <span>Yazdır / PDF Kaydet</span>
                </button>
                <button
                  type="button"
                  onClick={() => setZeroReportModalOpen(false)}
                  className="p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white cursor-pointer"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            {/* 2-Page Live Editable Sheet Canvas */}
            <div className="flex-1 overflow-y-auto bg-slate-300 p-4 sm:p-6">
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 max-w-[1120px] mx-auto items-start">

                {/* ==================== SAYFA 1 ==================== */}
                <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 min-h-[760px] flex flex-col justify-between font-sans">
                  <div>
                    <div className="text-[10px] font-bold text-slate-400 uppercase mb-1 flex justify-between">
                      <span>Sayfa 1 / 2</span>
                      <span>Tarih: {zeroReportForm.date}</span>
                    </div>

                    {/* Top Header Box (Logos + Title) */}
                    <div className="border border-slate-700 px-3 py-2 flex items-center justify-between mb-4">
                      <img src={circularOgmLogoDataUri} alt="OGM Amblem" className="w-12 h-12 rounded-full object-cover shrink-0 bg-white border border-slate-200" />
                      <div className="text-center font-bold text-slate-500 tracking-wide text-sm sm:text-base px-2">
                        HAVA ARACI YEDEK PARÇA TALEP FORMU
                      </div>
                      <img src={circularOgmLogoDataUri} alt="OGM Amblem" className="w-12 h-12 rounded-full object-cover shrink-0 bg-white border border-slate-200" />
                    </div>

                    {/* 1. Hava Aracı Bilgileri */}
                    <table className="w-full border-collapse border border-black mb-4 text-xs">
                      <thead>
                        <tr className="bg-[#d9d9d9]">
                          <th className="border border-black px-2 py-1 text-left font-bold w-[60%]">1. Hava Aracı Bilgileri</th>
                          <th className="border border-black px-2 py-1 text-left font-bold w-[40%]"></th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td className="border border-black px-2.5 py-2 align-top">
                            <div className="flex items-center gap-1.5">
                              <span className="shrink-0">1. Hava Aracı Tipi :</span>
                              <input
                                type="text"
                                value={zeroReportForm.havaAraciTipi}
                                onChange={e => setZeroReportForm(prev => ({ ...prev, havaAraciTipi: e.target.value }))}
                                className="w-full border-b border-dotted border-slate-400 focus:border-blue-600 outline-none font-bold text-xs px-1 bg-transparent"
                              />
                            </div>
                          </td>
                          <td className="border border-black px-3 py-1.5 text-[11px] space-y-0.5">
                            {[
                              { key: 'at802', label: 'AT-802' },
                              { key: 't70', label: 'T-70' },
                              { key: 'bell429', label: 'Bell-429' },
                              { key: 'b360', label: 'B-360' }
                            ].map(opt => (
                              <label key={opt.key} className="flex items-center gap-1.5 cursor-pointer select-none">
                                <input
                                  type="checkbox"
                                  checked={(zeroReportForm.selectedTypes as any)[opt.key]}
                                  onChange={e =>
                                    setZeroReportForm(prev => ({
                                      ...prev,
                                      selectedTypes: { ...prev.selectedTypes, [opt.key]: e.target.checked }
                                    }))
                                  }
                                  className="w-3.5 h-3.5 accent-emerald-700 cursor-pointer"
                                />
                                <span>{opt.label}</span>
                              </label>
                            ))}
                          </td>
                        </tr>
                      </tbody>
                    </table>

                    {/* 2-Talep Edilen Malzemeler (20 Rows) */}
                    <table className="w-full border-collapse border border-black text-[11px]">
                      <thead>
                        <tr className="bg-[#d9d9d9]">
                          <th colSpan={6} className="border border-black px-2 py-1 text-left font-bold text-xs">
                            2-Talep Edilen Malzemeler
                          </th>
                        </tr>
                        <tr className="bg-[#f3f4f6]">
                          <th className="border border-black px-1 py-1 text-center font-bold w-[7%]">Sıra<br/>No</th>
                          <th className="border border-black px-1.5 py-1 text-left font-bold w-[23%]">Parça Adı/Tanımı</th>
                          <th className="border border-black px-1.5 py-1 text-left font-bold w-[20%]">Parça Numarası<br/>(Part Number)</th>
                          <th className="border border-black px-1.5 py-1 text-left font-bold w-[20%]">Alternatif Parça Numaraları<br/>(Eğer biliniyorsa)</th>
                          <th className="border border-black px-1 py-1 text-center font-bold w-[13%]">İstenen Miktar</th>
                          <th className="border border-black px-1 py-1 text-center font-bold w-[17%]">Parçanın Öncelik Derecesi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {zeroReportForm.rows.map((row, rIdx) => (
                          <tr key={row.siraNo} className="hover:bg-amber-50/40">
                            <td className="border border-black px-1 py-0.5 text-center font-bold">{row.siraNo}</td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={row.parcaAdi}
                                onChange={e => {
                                  const v = e.target.value;
                                  setZeroReportForm(prev => ({
                                    ...prev,
                                    rows: prev.rows.map((r, i) => (i === rIdx ? { ...r, parcaAdi: v } : r))
                                  }));
                                }}
                                className="w-full px-1.5 py-0.5 text-[10.5px] bg-transparent outline-none focus:bg-blue-50"
                              />
                            </td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={row.parcaNo}
                                onChange={e => {
                                  const v = e.target.value;
                                  setZeroReportForm(prev => ({
                                    ...prev,
                                    rows: prev.rows.map((r, i) => (i === rIdx ? { ...r, parcaNo: v } : r))
                                  }));
                                }}
                                className="w-full px-1.5 py-0.5 text-[10.5px] font-mono font-bold bg-transparent outline-none focus:bg-blue-50"
                              />
                            </td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={row.alternatifPn}
                                onChange={e => {
                                  const v = e.target.value;
                                  setZeroReportForm(prev => ({
                                    ...prev,
                                    rows: prev.rows.map((r, i) => (i === rIdx ? { ...r, alternatifPn: v } : r))
                                  }));
                                }}
                                className="w-full px-1.5 py-0.5 text-[10.5px] font-mono bg-transparent outline-none focus:bg-blue-50"
                              />
                            </td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={row.istenenMiktar}
                                onChange={e => {
                                  const v = e.target.value;
                                  setZeroReportForm(prev => ({
                                    ...prev,
                                    rows: prev.rows.map((r, i) => (i === rIdx ? { ...r, istenenMiktar: v } : r))
                                  }));
                                }}
                                className="w-full px-1 py-0.5 text-center text-[10.5px] font-bold bg-transparent outline-none focus:bg-blue-50"
                              />
                            </td>
                            <td className="border border-black p-0">
                              <select
                                value={row.oncelik}
                                onChange={e => {
                                  const v = e.target.value as ZeroReportRow['oncelik'];
                                  setZeroReportForm(prev => ({
                                    ...prev,
                                    rows: prev.rows.map((r, i) => (i === rIdx ? { ...r, oncelik: v } : r))
                                  }));
                                }}
                                className={`w-full px-1 py-0.5 text-center text-[10.5px] font-bold bg-transparent outline-none cursor-pointer ${
                                  row.oncelik === 'Yüksek'
                                    ? 'text-red-600'
                                    : row.oncelik === 'Normal'
                                    ? 'text-amber-600'
                                    : row.oncelik === 'Düşük'
                                    ? 'text-emerald-700'
                                    : ''
                                }`}
                              >
                                <option value="">-</option>
                                <option value="Yüksek">Yüksek</option>
                                <option value="Normal">Normal</option>
                                <option value="Düşük">Düşük</option>
                              </select>
                            </td>
                          </tr>
                        ))}
                        <tr className="bg-[#d9d9d9]">
                          <td colSpan={6} className="border border-black px-2 py-0.5 text-center font-bold text-[10.5px]">
                            Parcanın Öncelik Derecesi kısmına yazılacaklar
                          </td>
                        </tr>
                        <tr>
                          <td colSpan={6} className="border border-black px-2 py-1 text-[10.5px]">
                            <span className="text-red-600 font-bold">Yüksek</span> – Kısa Sürede Gerekli (Arıza vb)
                          </td>
                        </tr>
                        <tr>
                          <td colSpan={6} className="border border-black px-2 py-1 text-[10.5px]">
                            <span className="text-amber-500 font-bold">Normal</span> – Planlı Bakım İçin
                          </td>
                        </tr>
                        <tr>
                          <td colSpan={6} className="border border-black px-2 py-1 text-[10.5px]">
                            <span className="text-emerald-600 font-bold">Düşük</span> – Depo Miktarını Tamamlamak İçin
                          </td>
                        </tr>
                        <tr>
                          <td colSpan={6} className="border border-black px-2 py-1 text-[10px] italic">
                            *Ek satırlara ihtiyaç duyulması halinde bu bölümü çoğaltabilirsiniz.*
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* ==================== SAYFA 2 ==================== */}
                <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 min-h-[760px] flex flex-col justify-between font-sans">
                  <div>
                    <div className="text-[10px] font-bold text-slate-400 uppercase mb-1 flex justify-between">
                      <span>Sayfa 2 / 2</span>
                      <span>Parça Kataloğu &amp; Notlar</span>
                    </div>

                    {/* Top Header Box (Logos + Title repeated on Page 2) */}
                    <div className="border border-slate-700 px-3 py-2 flex items-center justify-between mb-4">
                      <img src={circularOgmLogoDataUri} alt="OGM Amblem" className="w-12 h-12 rounded-full object-cover shrink-0 bg-white border border-slate-200" />
                      <div className="text-center font-bold text-slate-500 tracking-wide text-sm sm:text-base px-2">
                        HAVA ARACI YEDEK PARÇA TALEP FORMU
                      </div>
                      <img src={circularOgmLogoDataUri} alt="OGM Amblem" className="w-12 h-12 rounded-full object-cover shrink-0 bg-white border border-slate-200" />
                    </div>

                    {/* 3. Parça Kataloğu Görseli */}
                    <table className="w-full border-collapse border border-black mb-4 text-xs">
                      <thead>
                        <tr className="bg-[#d9d9d9]">
                          <th className="border border-black px-2.5 py-1 text-left font-bold">3. Parça Kataloğu Görseli</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td className="border border-black p-3 align-top h-[410px] relative">
                            <p className="text-[11px] leading-snug mb-2">
                              Lütfen istenen parçanın katalogdan alınmış görselini formun ekine ekleyiniz.<br/>
                              Doğru parçanın tespit edilmesine yardımcı olacaktır.
                            </p>

                            {capturedCatalogImage ? (
                              <div className="relative group mt-2 flex flex-col items-center">
                                <img
                                  src={capturedCatalogImage}
                                  alt="Parça Kataloğu Görseli"
                                  className="max-w-full max-h-[310px] object-contain border border-slate-300 shadow-xs"
                                />
                                <div className="mt-2 flex items-center gap-2">
                                  <button
                                    type="button"
                                    onClick={() => handleCaptureCatalogScreenshot(true)}
                                    className="px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-white text-[10px] font-bold flex items-center gap-1 cursor-pointer"
                                  >
                                    <Camera className="w-3 h-3" />
                                    <span>Yeniden Yakala</span>
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setCapturedCatalogImage('')}
                                    className="px-2.5 py-1 rounded-lg bg-rose-600 hover:bg-rose-500 text-white text-[10px] font-bold cursor-pointer"
                                  >
                                    Görseli Kaldır
                                  </button>
                                </div>
                              </div>
                            ) : (
                              <div className="h-[310px] border-2 border-dashed border-slate-300 rounded-xl flex flex-col items-center justify-center text-center p-4 bg-slate-50/70">
                                <ImageIcon className="w-9 h-9 text-slate-400 mb-2" />
                                <p className="text-xs font-bold text-slate-600 mb-1">
                                  Henüz Katalog Görseli Eklenmedi
                                </p>
                                <p className="text-[10.5px] text-slate-500 max-w-xs mb-3">
                                  Ekran yakala butonu ile PDF sayfadan kırpabilir, <strong>Ctrl+V</strong> ile yapıştırabilir veya dosya seçebilirsiniz.
                                </p>
                                <div className="flex items-center gap-2 flex-wrap justify-center">
                                  <button
                                    type="button"
                                    onClick={() => handleCaptureCatalogScreenshot(true)}
                                    className="px-3 py-1.5 rounded-lg bg-amber-500 hover:bg-amber-400 text-slate-950 font-black text-[11px] flex items-center gap-1 cursor-pointer shadow-xs"
                                  >
                                    <Camera className="w-3.5 h-3.5" />
                                    <span>PDF Sayfadan Ekran Yakala</span>
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => catalogFileInputRef.current?.click()}
                                    className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-white font-bold text-[11px] flex items-center gap-1 cursor-pointer"
                                  >
                                    <Upload className="w-3.5 h-3.5" />
                                    <span>Görsel Yükle</span>
                                  </button>
                                  <input
                                    ref={catalogFileInputRef}
                                    type="file"
                                    accept="image/*"
                                    className="hidden"
                                    onChange={e => {
                                      const f = e.target.files?.[0];
                                      if (f) {
                                        const reader = new FileReader();
                                        reader.onload = ev => {
                                          if (ev.target?.result) setCapturedCatalogImage(String(ev.target.result));
                                        };
                                        reader.readAsDataURL(f);
                                      }
                                      e.target.value = '';
                                    }}
                                  />
                                </div>
                              </div>
                            )}
                          </td>
                        </tr>
                      </tbody>
                    </table>

                    {/* 4. Ek Notlar/Açıklamalar + İmza Bloğu */}
                    <table className="w-full border-collapse border border-black text-xs">
                      <thead>
                        <tr className="bg-[#d9d9d9]">
                          <th colSpan={2} className="border border-black px-2.5 py-1 text-left font-bold">
                            4. Ek Notlar/Açıklamalar
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td colSpan={2} className="border border-black p-2.5 align-top">
                            <div className="text-[11px] mb-1.5">Parçanın özel gereksinimlerini belirten ek bilgiler var ise:</div>
                            <textarea
                              rows={5}
                              value={zeroReportForm.ekNotlar}
                              onChange={e => setZeroReportForm(prev => ({ ...prev, ekNotlar: e.target.value }))}
                              placeholder="Varsa özel sertifika, hızlı tedarik veya teknik notları buraya yazınız..."
                              className="w-full text-xs bg-transparent outline-none resize-none focus:bg-blue-50/40 p-1"
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="border border-black px-2.5 py-2 w-[50%] text-[11px]">
                            Talep Yapan Personel Adı Soyadı:
                          </td>
                          <td className="border border-black p-1 w-[50%]">
                            <input
                              type="text"
                              value={zeroReportForm.talepYapanAdSoyad}
                              onChange={e => setZeroReportForm(prev => ({ ...prev, talepYapanAdSoyad: e.target.value }))}
                              placeholder="Ad Soyad giriniz..."
                              className="w-full px-2 py-1 text-xs font-bold bg-transparent outline-none focus:bg-blue-50"
                            />
                          </td>
                        </tr>
                        <tr>
                          <td className="border border-black px-2.5 py-3 w-[50%] text-[11px]">
                            Talep Sahibi İmza:
                          </td>
                          <td className="border border-black px-2.5 py-3 w-[50%]"></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};
