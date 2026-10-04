import React, { useState, useEffect, useRef, useMemo } from 'react';
import html2canvas from 'html2canvas';
import { 
  Printer, 
  Download, 
  Copy, 
  Check, 
  X, 
  Barcode, 
  QrCode, 
  FileText, 
  Search, 
  Layers, 
  ArrowDownLeft, 
  Package,
  Calendar,
  User,
  MapPin,
  FileSpreadsheet,
  Bluetooth,
  Usb,
  Radio,
  CheckCircle2,
  AlertCircle,
  Wrench
} from 'lucide-react';
import JsBarcode from 'jsbarcode';
import QRCode from 'qrcode';
import { 
  Document as DocxDocument, 
  Packer as DocxPacker, 
  Paragraph as DocxParagraph, 
  Table as DocxTable, 
  TableRow as DocxTableRow, 
  TableCell as DocxTableCell, 
  WidthType, 
  AlignmentType, 
  TextRun as DocxTextRun, 
  ImageRun as DocxImageRun, 
  BorderStyle, 
  HeadingLevel,
  ShadingType
} from 'docx';
import { DepoItem } from './DepoManagementModal';
import { DepoTransaction } from '../types';

export interface DepoSlipData {
  type: 'raf' | 'giris' | 'etiket';
  
  // Raf Slibi Alanları
  lokasyonKodu?: string;
  bolgeZone?: string;
  rafTipi?: string;
  tanimIcerik?: string;
  guncelSayim?: string | number;
  birim?: string;
  sonSayimTarih?: string;
  sayimPersonel?: string;
  
  // Malzeme Giriş Slibi Alanları
  islemTransfer?: string;
  tarih?: string;
  malzemeDesc?: string;
  pn?: string;
  sn?: string;
  adet?: number | string;
  depoYeri?: string;
  teslimAlan?: string;
  kabulYapan?: string;
  aciklama?: string;
}

interface DepoSlipPrintModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialType?: 'raf' | 'giris' | 'etiket' | null;
  initialData?: (Partial<DepoSlipData> & { imageUrl?: string }) | null;
  inventory?: DepoItem[];
  techizatList?: any[];
  mode?: 'depo' | 'techizat' | 'giris_only';
  isGirisOnly?: boolean;
  transactions?: DepoTransaction[];
  showNotification?: (msg: string) => void;
}

// Transliterate Turkish characters and clean strings for CODE128 barcode compatibility
function sanitizeBarcode(val: string): string {
  if (!val) return 'OGM12345';
  const trMap: Record<string, string> = {
    'ç': 'c', 'Ç': 'C',
    'ğ': 'g', 'Ğ': 'G',
    'ı': 'i', 'İ': 'I',
    'ö': 'o', 'Ö': 'O',
    'ş': 's', 'Ş': 'S',
    'ü': 'u', 'Ü': 'U'
  };
  const clean = val.replace(/[çÇğĞıİöÖşŞüÜ]/g, m => trMap[m] || m)
    .replace(/[^A-Za-z0-9\-_. /]/g, '')
    .trim();
  return clean || 'OGM12345';
}

// Extract direct viewable CDN URL for Google Drive image links
export function getDirectDriveImageUrl(url: string): string {
  if (!url || typeof url !== 'string') return '';
  const trimmed = url.trim();
  if (!trimmed) return '';

  if (trimmed.startsWith('data:image')) return trimmed;

  let fileId = '';
  const fileDMatch = trimmed.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileDMatch && fileDMatch[1]) {
    fileId = fileDMatch[1];
  } else {
    const idParamMatch = trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (idParamMatch && idParamMatch[1]) {
      fileId = idParamMatch[1];
    } else if (/^[a-zA-Z0-9_-]{25,}$/.test(trimmed)) {
      fileId = trimmed;
    }
  }

  if (fileId) {
    return `https://lh3.googleusercontent.com/d/${fileId}=w1000`;
  }

  return trimmed;
}

// Fallback 1x1 transparent PNG data URL
const FALLBACK_1X1_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

// Convert OKLCH colors to RGB/RGBA dynamically to prevent html2canvas crash on Tailwind v4
function oklchToRgb(l: number, c: number, h: number): [number, number, number] {
  const hRad = (h * Math.PI) / 180;
  const a = c * Math.cos(hRad);
  const b = c * Math.sin(hRad);

  const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
  const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
  const s_ = l - 0.0894841775 * a - 1.2914855480 * b;

  const l_3 = l_ * l_ * l_;
  const m_3 = m_ * m_ * m_;
  const s_3 = s_ * s_ * s_;

  const rL = +4.0767416621 * l_3 - 3.3077115913 * m_3 + 0.2309699292 * s_3;
  const gL = -1.2684380046 * l_3 + 2.6097574011 * m_3 - 0.3413193965 * s_3;
  const bL = -0.0041960863 * l_3 - 0.7034186147 * m_3 + 1.7076147010 * s_3;

  const fn = (v: number) => {
    const lim = 0.0031308;
    return v > lim ? 1.055 * Math.pow(v, 1 / 2.4) - 0.055 : 12.92 * v;
  };

  const outR = Math.max(0, Math.min(255, Math.round(fn(rL) * 255)));
  const outG = Math.max(0, Math.min(255, Math.round(fn(gL) * 255)));
  const outB = Math.max(0, Math.min(255, Math.round(fn(bL) * 255)));

  return [outR, outG, outB];
}

function parseOklchStringAndConvert(oklchStr: string): string {
  const regex = /oklch\(\s*([0-9.-]+%?)\s+([0-9.-]+%?)\s+([0-9.-]+%?)(?:\s*\/\s*([0-9.-]+%?))?\s*\)/gi;
  return oklchStr.replace(regex, (match, lStr, cStr, hStr, aStr) => {
    try {
      let L = lStr.endsWith('%') ? parseFloat(lStr) / 100 : parseFloat(lStr);
      let C = cStr.endsWith('%') ? parseFloat(cStr) / 100 : parseFloat(cStr);
      let H = hStr.endsWith('%') ? parseFloat(hStr) : parseFloat(hStr);
      
      if (isNaN(L)) L = 0;
      if (isNaN(C)) C = 0;
      if (isNaN(H)) H = 0;

      const [r, g, b] = oklchToRgb(L, C, H);
      
      if (aStr) {
        let A = aStr.endsWith('%') ? parseFloat(aStr) / 100 : parseFloat(aStr);
        if (isNaN(A)) A = 1;
        return `rgba(${r}, ${g}, ${b}, ${A})`;
      }
      return `rgb(${r}, ${g}, ${b})`;
    } catch (e) {
      return 'rgb(0, 0, 0)';
    }
  });
}

// Helper to generate Base64 Barcode Data URL locally using an offscreen canvas
function generateLocalBarcodeDataUrl(text: string): string {
  const safeText = sanitizeBarcode(text);
  try {
    const canvas = document.createElement('canvas');
    JsBarcode(canvas, safeText, {
      format: 'CODE128',
      width: 2,
      height: 45,
      displayValue: false,
      margin: 4,
      background: '#ffffff',
      lineColor: '#000000'
    });
    return canvas.toDataURL('image/png');
  } catch (e) {
    console.warn('Local barcode render error:', e);
    // Draw simple programmatic barcode fallback on canvas if JsBarcode fails
    try {
      const canvas = document.createElement('canvas');
      canvas.width = 240;
      canvas.height = 45;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, 240, 45);
        ctx.fillStyle = '#000000';
        let x = 12;
        for (let i = 0; i < safeText.length; i++) {
          const code = safeText.charCodeAt(i);
          const w = (code % 3) + 1;
          ctx.fillRect(x, 4, w, 37);
          x += w + 2;
        }
        return canvas.toDataURL('image/png');
      }
    } catch {}
    return FALLBACK_1X1_PNG;
  }
}

// Helper to generate Base64 QR Code Data URL locally
async function generateLocalQrDataUrl(text: string): Promise<string> {
  try {
    return await QRCode.toDataURL(text || 'OGM-HAVACILIK', {
      width: 150,
      margin: 1,
      color: {
        dark: '#000000',
        light: '#ffffff'
      }
    });
  } catch (e) {
    console.warn('Local QR render error:', e);
    return FALLBACK_1X1_PNG;
  }
}

// Convert dataURL to Uint8Array for native docx OpenXML image embedding
function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  if (!dataUrl) return new Uint8Array(0);
  try {
    const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  } catch (e) {
    console.warn('dataUrlToUint8Array error:', e);
    return new Uint8Array(0);
  }
}

export const DepoSlipPrintModal: React.FC<DepoSlipPrintModalProps> = ({
  isOpen,
  onClose,
  initialType = null,
  initialData = null,
  inventory = [],
  techizatList = [],
  mode = 'depo',
  isGirisOnly = false,
  transactions = [],
  showNotification
}) => {
  const isGirisOnlyMode = mode === 'giris_only' || isGirisOnly;
  const isTechizatMode = mode === 'techizat' || initialType === 'etiket';
  const [selectedSlipType, setSelectedSlipType] = useState<'raf' | 'giris' | 'etiket'>(
    isTechizatMode ? 'etiket' : (isGirisOnlyMode || initialType === 'giris' ? 'giris' : 'raf')
  );
  const [imageUrlInput, setImageUrlInput] = useState<string>('');
  const [activeTab, setActiveTab] = useState<'preview' | 'ascii'>('preview');
  const [orientation, setOrientation] = useState<'dikey' | 'yatay'>('yatay');
  const [copied, setCopied] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [isAutocompleteOpen, setIsAutocompleteOpen] = useState<boolean>(false);

  // Web Bluetooth & Web USB & Web Serial Printer Connection States
  const [btDevice, setBtDevice] = useState<any>(null);
  const [usbDevice, setUsbDevice] = useState<any>(null);
  const [serialPort, setSerialPort] = useState<any>(null);
  const [printerConnectionStatus, setPrinterConnectionStatus] = useState<string>('Termal Cihaz Seçilmedi (Tarayıcı Standart)');
  const [isDeviceMenuOpen, setIsDeviceMenuOpen] = useState(false);

  // Bulk Print State (Tüm Depo Etiketleri / Envanter Toplu Bas)
  const [isBulkPrint, setIsBulkPrint] = useState(false);
  const [bulkExportProgress, setBulkExportProgress] = useState<number | null>(null);

  // Form Fields State for Raf Slibi / Teçhizat Etiketi
  const [lokasyonKodu, setLokasyonKodu] = useState('A-04-R2-BIN-08');
  const [bolgeZone, setBolgeZone] = useState('BÖLGE A - SARF MALZEME');
  const [rafTipi, setRafTipi] = useState('KÜÇÜK PARÇA GÖZÜ');
  const [tanimIcerik, setTanimIcerik] = useState('O-RING / SEAL / GASKET GÖZÜ');
  const [guncelSayim, setGuncelSayim] = useState('120');
  const [sonSayimTarih, setSonSayimTarih] = useState('27.09.2026');
  const [sayimPersonel, setSayimPersonel] = useState('4082');

  // Form Fields State for Giriş Slibi
  const [islemTransfer, setIslemTransfer] = useState('ANKARA GİREN');
  const [tarih, setTarih] = useState('27.09.2026 14:02');
  const [malzemeDesc, setMalzemeDesc] = useState('ENCODER');
  const [pn, setPn] = useState('013-00066-00');
  const [sn, setSn] = useState('SN-99482014');
  const [adet, setAdet] = useState<number | string>('1');
  const [depoYeri, setDepoYeri] = useState('ANKARA MERKEZ DEP');
  const [teslimAlan, setTeslimAlan] = useState('Yetkili Teknisyen');
  const [kabulYapan, setKabulYapan] = useState('Depo Sorumlusu');
  const [aciklama, setAciklama] = useState('Kutuda hasar yok. Onaylandı.');

  // Toplu Etiket Basım Filtre State (Sarf vs Kimyasal, Ankara Depo vs, Stoku 0 olanlar haric)
  const [bulkDepoCategory, setBulkDepoCategory] = useState<'sarf' | 'kimyasal'>('sarf');
  const [bulkSelectedDepo, setBulkSelectedDepo] = useState<string>('ANKARA DEPO');

  const filteredBulkItems = useMemo(() => {
    if (!inventory || inventory.length === 0) return [];
    return inventory.filter((item: any) => {
      // 1. Category check
      const cat = (item.category || item.kategori || item.DEPO_TIPI || '').toLowerCase();
      const isKimyasal = cat.includes('kimyasal') || cat.includes('chemical') || (item.description || item.name || '').toLowerCase().includes('kimyasal');
      if (bulkDepoCategory === 'kimyasal' && !isKimyasal) return false;
      if (bulkDepoCategory === 'sarf' && isKimyasal) return false;

      // 2. Specific Depo check
      const loc = (item.lokasyonNo || item.location || item.DEPO_YERI || '').toUpperCase();
      if (bulkSelectedDepo !== 'TÜM DEPOLAR' && bulkSelectedDepo !== 'TÜMÜ') {
        const targetDepoName = bulkSelectedDepo.replace(' DEPO', '').trim().toUpperCase();
        if (loc && !loc.includes(targetDepoName) && !loc.includes(bulkSelectedDepo)) {
          if (targetDepoName === 'ANKARA' && (item.ankaraMevcut === 0 || item.ankaraMevcut === undefined)) {
            return false;
          }
        }
      }

      // 3. Exclude 0 quantity items (0 HARİÇ HEPSİNİ BASAR!)
      const qty = item.ankaraMevcut ?? item.toplamStok ?? item.quantity ?? item.guncelSayim ?? item.ADET ?? 0;
      if (Number(qty) <= 0) return false;

      return true;
    });
  }, [inventory, bulkDepoCategory, bulkSelectedDepo]);

  // Dynamic Locations and S/N dropdowns for selected equipment
  const matchingTechizatItems = useMemo(() => {
    const list = isTechizatMode && techizatList && techizatList.length > 0 ? techizatList : inventory;
    if (!tanimIcerik.trim() && !pn.trim()) return list;
    const qDesc = tanimIcerik.toLowerCase().trim();
    const qPn = pn.toLowerCase().trim();
    return list.filter((item: any) => {
      const d = (item.description || item.name || item.ALET_ADI || item.TECHIZAT_ADI || item.tanim || '').toLowerCase();
      const p = (item.partNumber || item.pn || item.PARCA_NO || item.pnNo || '').toLowerCase();
      return (qDesc && d.includes(qDesc)) || (qPn && p.includes(qPn));
    });
  }, [tanimIcerik, pn, techizatList, inventory, isTechizatMode]);

  // Helper to extract clean individual location names from concatenated or delimited strings
  const extractIndividualLocations = (raw: string): string[] => {
    if (!raw || raw === '-' || raw === '--') return [];
    const result: string[] = [];
    const knownCities = ['ANKARA', 'ANTALYA', 'ÇANAKKALE', 'CANAKKALE', 'MİLAS', 'MILAS', 'BURSA', 'KARAİN', 'KARAIN', 'İZMİR', 'IZMIR', 'MUĞLA', 'MUGLA', 'GÜVERCİNLİK', 'GUVERCINLIK', 'HANGAR', 'GARAJ', 'DEPO'];

    const chunks = raw.split(/[,/;\r\n\t]+/);
    chunks.forEach(chunk => {
      const trimmed = chunk.trim();
      if (!trimmed || trimmed === '-' || trimmed === '--') return;

      const words = trimmed.split(/\s+/);
      const matched = words.filter(w => knownCities.includes(w.toUpperCase()));

      if (matched.length >= 2) {
        words.forEach(w => {
          const cleanW = w.trim().toUpperCase();
          if (cleanW && cleanW.length >= 3 && cleanW !== 'VE' && cleanW !== 'İLE') {
            result.push(cleanW);
          }
        });
      } else {
        result.push(trimmed);
      }
    });

    return Array.from(new Set(result));
  };

  const availableLocations = useMemo(() => {
    const locs = new Set<string>();
    matchingTechizatItems.forEach((item: any) => {
      const rawLoc = item.lokasyonNo || item.location || item.DEPO_YERI || item.HANGAR_LOKASYON || item.unit || '';
      extractIndividualLocations(rawLoc).forEach(l => locs.add(l));

      if (item.ankaraMevcut > 0) locs.add('ANKARA HANGAR');
      if (item.karainMevcut > 0) locs.add('KARAİN HANGAR');
      if (item.canakkaleMevcut > 0) locs.add('ÇANAKKALE HANGAR');
      if (item.milasMevcut > 0) locs.add('MİLAS HANGAR');
      if (item.bursaMevcut > 0) locs.add('BURSA HANGAR');
    });

    if (locs.size === 0 && lokasyonKodu) {
      extractIndividualLocations(lokasyonKodu).forEach(l => locs.add(l));
    }

    if (locs.size === 0) {
      ['ANKARA HANGAR', 'KARAİN HANGAR', 'ÇANAKKALE HANGAR', 'MİLAS HANGAR', 'BURSA HANGAR'].forEach(h => locs.add(h));
    }

    return Array.from(locs);
  }, [matchingTechizatItems, lokasyonKodu]);

  const uniqueExcelDepotNames = useMemo(() => {
    const set = new Set<string>();
    set.add('ANKARA DEPO');
    (inventory || []).forEach((item: any) => {
      const loc = (item.lokasyonNo || item.location || item.DEPO_YERI || '').toUpperCase().trim();
      if (loc.includes('ANKARA')) set.add('ANKARA DEPO');
      if (loc.includes('KARAİN') || loc.includes('KARAIN')) set.add('KARAİN DEPO');
      if (loc.includes('ÇANAKKALE') || loc.includes('CANAKKALE')) set.add('ÇANAKKALE DEPO');
      if (loc.includes('MİLAS') || loc.includes('MILAS')) set.add('MİLAS DEPO');
      if (loc.includes('BURSA')) set.add('BURSA DEPO');
      if (loc.includes('GÜVERCİNLİK') || loc.includes('GUVERCINLIK')) set.add('GÜVERCİNLİK DEPO');
      if (Number(item.ankaraMevcut) > 0) set.add('ANKARA DEPO');
      if (Number(item.karainMevcut) > 0) set.add('KARAİN DEPO');
      if (Number(item.canakkaleMevcut) > 0) set.add('ÇANAKKALE DEPO');
      if (Number(item.milasMevcut) > 0) set.add('MİLAS DEPO');
      if (Number(item.bursaMevcut) > 0) set.add('BURSA DEPO');
    });
    set.add('TÜM DEPOLAR');
    return Array.from(set);
  }, [inventory]);

  // Kalibrasyon / Bakım Tarihi State
  const [kalibrasyonTarihi, setKalibrasyonTarihi] = useState<string>('');

  useEffect(() => {
    if (!matchingTechizatItems || matchingTechizatItems.length === 0) return;
    const item = matchingTechizatItems[0] as any;
    const dateVal = item.sonBakimTarihi || item.kalibrasyonTarihi || item.gelecekBakimTarihi || item['SON KONTROL / KALİBRASYON / BAKIM'] || item['GELECEK KONTROL / KALİBRASYON / BAKIM'] || item.lastControlDate || '';
    if (dateVal) {
      setKalibrasyonTarihi(dateVal);
    }
  }, [matchingTechizatItems]);

  const availableSerialNumbers = useMemo(() => {
    const sns = new Set<string>();
    if (sn && sn !== '-') sns.add(sn);
    
    const locFiltered = matchingTechizatItems.filter((item: any) => {
      const loc = item.lokasyonNo || item.location || item.DEPO_YERI || item.HANGAR_LOKASYON || '';
      return !lokasyonKodu || !loc || loc.toLowerCase() === lokasyonKodu.toLowerCase();
    });

    const targetList = locFiltered.length > 0 ? locFiltered : matchingTechizatItems;

    targetList.forEach((item: any) => {
      const rawSn = item.serialAndNotes || item.sn || item.SERI_NO || item.snNo || '';
      if (rawSn && rawSn !== '-') {
        const parts = rawSn.split(/[,/;]+/);
        parts.forEach((p: string) => {
          const clean = p.trim();
          if (clean && clean !== '-') sns.add(clean);
        });
      }
    });

    return Array.from(sns);
  }, [matchingTechizatItems, lokasyonKodu, sn]);

  // Generated ASCII, Barcode DataUrl and QR
  const [qrDataUrl, setQrDataUrl] = useState<string>('');
  const [barcodeDataUrl, setBarcodeDataUrl] = useState<string>('');
  const [asciiQrString, setAsciiQrString] = useState<string>('');
  const barcodeSvgRef = useRef<SVGSVGElement | null>(null);

  // Sync initialData when modal opens
  useEffect(() => {
    if (!isOpen) return;

    if (isTechizatMode || (initialType as string) === 'etiket') {
      setSelectedSlipType('etiket');
    } else if (initialType === 'giris') {
      setSelectedSlipType('giris');
    } else {
      setSelectedSlipType('raf');
    }

    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    const defaultDateStr = `${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;

    if (initialData) {
      if (initialData.imageUrl) setImageUrlInput(initialData.imageUrl);

      if ((initialData.type as string) === 'etiket' || isTechizatMode) {
        setSelectedSlipType('etiket');
        if (initialData.tanimIcerik || initialData.malzemeDesc) {
          const title = initialData.tanimIcerik || initialData.malzemeDesc || '';
          setTanimIcerik(title);
          setMalzemeDesc(title);
        }
        if (initialData.pn) setPn(initialData.pn);
        if (initialData.sn) setSn(initialData.sn);
        if (initialData.lokasyonKodu || initialData.depoYeri) {
          const loc = initialData.lokasyonKodu || initialData.depoYeri || '';
          setLokasyonKodu(loc);
          setDepoYeri(loc);
        }
        if (initialData.guncelSayim !== undefined || initialData.adet !== undefined) {
          const val = initialData.guncelSayim || initialData.adet || 1;
          setGuncelSayim(String(val));
          setAdet(val);
        }
      } else if (initialData.type === 'giris' || initialType === 'giris') {
        setSelectedSlipType('giris');
        if (initialData.islemTransfer) setIslemTransfer(initialData.islemTransfer);
        if (initialData.tarih) setTarih(initialData.tarih);
        else setTarih(defaultDateStr);
        if (initialData.malzemeDesc) setMalzemeDesc(initialData.malzemeDesc);
        if (initialData.pn) setPn(initialData.pn);
        if (initialData.sn) setSn(initialData.sn);
        if (initialData.adet !== undefined) setAdet(initialData.adet);
        if (initialData.depoYeri) setDepoYeri(initialData.depoYeri);
        if (initialData.teslimAlan) setTeslimAlan(initialData.teslimAlan);
        if (initialData.kabulYapan) setKabulYapan(initialData.kabulYapan);
        if (initialData.aciklama !== undefined) setAciklama(initialData.aciklama);
      } else if (initialData.type === 'raf' || initialType === 'raf') {
        setSelectedSlipType('raf');
        if (initialData.lokasyonKodu) setLokasyonKodu(initialData.lokasyonKodu);
        if (initialData.bolgeZone) setBolgeZone(initialData.bolgeZone);
        if (initialData.rafTipi) setRafTipi(initialData.rafTipi);
        if (initialData.tanimIcerik) setTanimIcerik(initialData.tanimIcerik);
        if (initialData.guncelSayim !== undefined) setGuncelSayim(String(initialData.guncelSayim));
        if (initialData.sonSayimTarih) setSonSayimTarih(initialData.sonSayimTarih);
        if (initialData.sayimPersonel) setSayimPersonel(initialData.sayimPersonel);
      }
    } else if (isTechizatMode && techizatList && techizatList.length > 0) {
      const first = techizatList[0];
      const desc = first.description || first.name || first.ALET_ADI || first.TECHIZAT_ADI || first.tanim || 'TEÇHİZAT';
      const itemPn = first.partNumber || first.pn || first.PARCA_NO || first.pnNo || '-';
      const itemSn = first.serialAndNotes || first.sn || first.SERI_NO || first.snNo || '-';
      const loc = first.lokasyonNo || first.location || first.DEPO_YERI || first.HANGAR_LOKASYON || 'ANKARA HANGAR';
      const qty = first.ankaraMevcut ?? first.gelen ?? first.toplamStok ?? first.ADET ?? first.MEVCUT ?? 1;
      const img = getDirectDriveImageUrl(first.imageUrl || first.fotoUrl || first.DriveFoto || first.foto || first.GÖRSEL || first.FOTOĞRAF || '');

      setTanimIcerik(desc);
      setMalzemeDesc(desc);
      setPn(itemPn);
      setSn(itemSn);
      setLokasyonKodu(loc);
      setDepoYeri(loc);
      setGuncelSayim(String(qty));
      setAdet(qty);
      setImageUrlInput(img);
    } else if (inventory.length > 0 && !initialData) {
      const first = inventory[0];
      setMalzemeDesc(first.description || first.name || 'ENCODER');
      setPn(first.partNumber || first.pn || '013-00066-00');
      setSn(first.serialAndNotes && first.serialAndNotes !== '-' ? first.serialAndNotes : 'SN-99482014');
      setLokasyonKodu(first.lokasyonNo || first.location || 'A-04-R2-BIN-08');
      setDepoYeri(first.lokasyonNo || 'ANKARA MERKEZ DEP');
      setTanimIcerik(`${first.description || first.name} GÖZÜ`);
      setGuncelSayim(String(first.ankaraMevcut || first.toplamStok || 120));
      setTarih(defaultDateStr);
    }
  }, [isOpen, initialType, initialData, inventory, techizatList, isTechizatMode]);

  // Derived Barcode and QR Values
  const barcodeValue = selectedSlipType === 'raf'
    ? `LOC-${(lokasyonKodu || 'A04R2BIN08').replace(/[^a-zA-Z0-9]/g, '').toUpperCase()}`
    : `${(pn || '013-00066-00').trim()}-${(sn && sn !== '-' ? sn : 'GENEL').replace(/[^a-zA-Z0-9]/g, '')}`;

  const qrContent = selectedSlipType === 'raf'
    ? `OGM-DEPO-RAF|LOC:${lokasyonKodu}|ZONE:${bolgeZone}|QTY:${guncelSayim}|DATE:${sonSayimTarih}`
    : `OGM-DEPO-GIRIS|PN:${pn}|SN:${sn}|QTY:${adet}|DATE:${tarih}|LOC:${depoYeri}`;

  // Generate Barcode SVG, DataURL & QR Code whenever values change
  useEffect(() => {
    if (!isOpen || isBulkPrint) return;

    // 1. Generate DataURL barcode so preview always works reliably via <img>
    try {
      const bData = generateLocalBarcodeDataUrl(barcodeValue || 'OGM12345');
      setBarcodeDataUrl(bData);
    } catch (e) {
      console.warn('Barcode data url error:', e);
    }

    // 2. Generate Code-128 via JsBarcode in SVG ref
    if (barcodeSvgRef.current) {
      try {
        const safeVal = sanitizeBarcode(barcodeValue);
        JsBarcode(barcodeSvgRef.current, safeVal, {
          format: 'CODE128',
          width: orientation === 'yatay' ? 1.0 : 1.2,
          height: orientation === 'yatay' ? 24 : 34,
          displayValue: false,
          margin: 0,
          background: '#ffffff',
          lineColor: '#000000'
        });
      } catch (e) {
        console.warn('JsBarcode render error:', e);
      }
    }

    // 3. Generate graphical QR DataURL
    QRCode.toDataURL(qrContent, {
      width: orientation === 'yatay' ? 65 : 85,
      margin: 1,
      color: {
        dark: '#000000',
        light: '#ffffff'
      }
    }).then(url => {
      setQrDataUrl(url);
    }).catch(err => {
      console.warn('QR code generation error:', err);
    });

    // 4. Generate UTF-8 Half-Block ASCII QR Code
    QRCode.toString(qrContent, {
      type: 'utf8',
      margin: 0
    }).then(ascii => {
      setAsciiQrString(ascii);
    }).catch(err => {
      console.warn('ASCII QR error:', err);
    });
  }, [isOpen, selectedSlipType, barcodeValue, qrContent, orientation, isBulkPrint]);

  // Handler to pick an inventory item to autofill
  const handleSelectInventoryItem = (item: DepoItem) => {
    setSearchQuery('');
    if (selectedSlipType === 'raf') {
      setLokasyonKodu(item.lokasyonNo || item.location || 'A-01-RAF-1');
      setBolgeZone(item.category === 'kimyasal' ? 'KİMYASAL DEPOSU' : 'SARF MALZEME DEPOSU');
      setTanimIcerik(`${item.description || item.name} GÖZÜ`);
      setGuncelSayim(String(item.ankaraMevcut ?? item.toplamStok ?? 1));
      const now = new Date();
      setSonSayimTarih(`${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`);
    } else {
      setMalzemeDesc(item.description || item.name || '');
      setPn(item.partNumber || item.pn || '-');
      setSn(item.serialAndNotes && item.serialAndNotes !== '-' ? item.serialAndNotes : 'SN-99482');
      setDepoYeri(item.lokasyonNo || 'ANKARA MERKEZ DEP');
      setAdet(1);
    }
  };

  // Build compact ASCII Slip String
  const generateAsciiSlip = (): string => {
    const LINE = '==========================================';
    const SUBLINE = '------------------------------------------';

    if (orientation === 'yatay') {
      if (selectedSlipType === 'raf') {
        return `${LINE}
 OGM HAVACILIK  •  RAF LOKASYON ETİKETİ
${LINE}
LOKASYON: ${lokasyonKodu.padEnd(14)} BÖLGE: ${bolgeZone.slice(0, 14)}
RAF TİPİ: ${rafTipi.padEnd(14)} SAYIM: ${guncelSayim} EA
TANIM   : ${tanimIcerik.slice(0, 30)}
${SUBLINE}
BARKOD: *${barcodeValue}*
${LINE}`;
      } else {
        return `${LINE}
 OGM HAVACILIK  •  MALZEME GİRİŞ ETİKETİ
${LINE}
İŞLEM: ${islemTransfer.padEnd(17)} TARİH: ${tarih}
MALZ : ${malzemeDesc.slice(0, 32)}
P/N  : ${pn.padEnd(18)} S/N: ${sn}
ADET : ${adet} EA              DEP: ${depoYeri}
${SUBLINE}
BARKOD: *${barcodeValue}*
${LINE}`;
      }
    } else {
      return `${LINE}
             OGM HAVACILIK
${selectedSlipType === 'raf' ? '      RAF LOKASYON VE SAYIM SLİBİ' : '         MALZEME GİRİŞ SLİBİ'}
${LINE}
TARİH: ${new Date().toLocaleString('tr-TR')}
${SUBLINE}
${selectedSlipType === 'raf' ? `LOKASYON : ${lokasyonKodu}
BÖLGE    : ${bolgeZone}
RAF TİPİ : ${rafTipi}
İÇERİK   : ${tanimIcerik}
SAYIM    : ${guncelSayim} EA` : `İŞLEM    : ${islemTransfer}
TANIM    : ${malzemeDesc}
P/N      : ${pn}
S/N      : ${sn}
ADET     : ${adet} EA
DEPO     : ${depoYeri}`}
${SUBLINE}
*${barcodeValue}*
${LINE}`;
    }
  };

  // Download TXT Slip File
  const handleDownloadTxt = () => {
    const content = generateAsciiSlip();
    const fileName = selectedSlipType === 'raf'
      ? `OGM_Raf_Slibi_76mm_${lokasyonKodu.replace(/[^a-zA-Z0-9_-]/g, '_')}.txt`
      : `OGM_Giris_Slibi_76mm_${(pn || 'MALZEME').replace(/[^a-zA-Z0-9_-]/g, '_')}.txt`;

    const blob = new Blob(['\ufeff' + content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    if (showNotification) {
      showNotification(`✅ 76mm ${fileName} indirildi.`);
    }
  };

  // Helper to build a 100% compliant OpenXML docx TableCell with embedded QR and 1D Barcode images (NO nested tables, NO nested paragraphs)
  // Helper to build a 100% compliant OpenXML docx TableCell children with embedded QR and 1D Barcode images
  const generateLabelChildren = (
    title: string,
    fields: { label: string; value: string; isBold?: boolean }[],
    qrBase64: string,
    barcodeBase64: string,
    barcodeVal: string
  ): DocxParagraph[] => {
    const qrBytes = dataUrlToUint8Array(qrBase64);
    const barcodeBytes = dataUrlToUint8Array(barcodeBase64);

    const cellChildren: DocxParagraph[] = [];

    // 1. Header Banner
    cellChildren.push(
      new DocxParagraph({
        alignment: AlignmentType.CENTER,
        shading: { type: ShadingType.CLEAR, fill: '0f172a' },
        spacing: { before: 80, after: 80 },
        children: [
          new DocxTextRun({ text: title, bold: true, color: 'FFFFFF', size: 19, font: 'Calibri' })
        ]
      })
    );

    // 2. Field Paragraphs
    fields.forEach(f => {
      cellChildren.push(
        new DocxParagraph({
          spacing: { before: 20, after: 20 },
          children: [
            new DocxTextRun({ text: `${f.label}: `, bold: true, size: 17, font: 'Calibri' }),
            new DocxTextRun({ text: f.value, bold: !!f.isBold, size: 17, font: 'Calibri' })
          ]
        })
      );
    });

    // 3. QR Code Image (Standalone Paragraph)
    if (qrBytes.length > 0) {
      cellChildren.push(
        new DocxParagraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 20 },
          children: [
            new DocxImageRun({
              data: qrBytes,
              transformation: { width: 75, height: 75 },
              type: 'png'
            } as any)
          ]
        })
      );
      cellChildren.push(
        new DocxParagraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 40 },
          children: [
            new DocxTextRun({ text: '[HIZLI SORGULAMA KAREKODU]', size: 13, color: '64748b', bold: true, font: 'Calibri' })
          ]
        })
      );
    }

      // 4. Barcode Image (Standalone Paragraph)
      if (barcodeBytes.length > 0) {
        cellChildren.push(
          new DocxParagraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 40, after: 10 },
            children: [
              new DocxImageRun({
                data: barcodeBytes,
                transformation: { width: 210, height: 36 },
                type: 'png'
              } as any)
            ]
          })
        );
      }

    return cellChildren;
  };

  // Helper to build a 100% compliant OpenXML docx TableCell with embedded QR and 1D Barcode images
  const createDocxLabelCell = (
    title: string,
    fields: { label: string; value: string; isBold?: boolean }[],
    qrBase64: string,
    barcodeBase64: string,
    barcodeVal: string,
    cellWidthPct: number = 50,
    borderStyle: any = BorderStyle.SINGLE,
    borderColor: string = '0f172a'
  ): DocxTableCell => {
    const children = generateLabelChildren(title, fields, qrBase64, barcodeBase64, barcodeVal);

    return new DocxTableCell({
      width: { size: cellWidthPct, type: WidthType.PERCENTAGE },
      margins: { top: 120, bottom: 120, left: 140, right: 140 },
      borders: {
        top: { style: borderStyle, size: 12, color: borderColor },
        bottom: { style: borderStyle, size: 12, color: borderColor },
        left: { style: borderStyle, size: 12, color: borderColor },
        right: { style: borderStyle, size: 12, color: borderColor }
      },
      children: children
    });
  };

  // Robust Native OpenXML MS Word (.docx) Export with 100% Turkish Support and True Embedded Images
  const handleDownloadWord = async () => {
    setBulkExportProgress(1);
    try {
      if (!isBulkPrint) {
        // Build the single slip using native Word elements for perfect resolution and Office 2007 compatibility
        const barcodeBase64 = generateLocalBarcodeDataUrl(barcodeValue);
        const qrBase64 = await generateLocalQrDataUrl(qrContent);

        const fields = selectedSlipType === 'raf' ? [
          { label: 'LOKASYON', value: lokasyonKodu, isBold: true },
          { label: 'BÖLGE', value: bolgeZone },
          { label: 'İÇERİK', value: tanimIcerik, isBold: true },
          { label: 'SAYIM', value: `${guncelSayim} EA`, isBold: true },
          { label: 'TARİH', value: sonSayimTarih }
        ] : [
          { label: 'İŞLEM', value: islemTransfer, isBold: true },
          { label: 'MALZEME', value: malzemeDesc, isBold: true },
          { label: 'P/N', value: pn, isBold: true },
          { label: 'S/N', value: sn },
          { label: 'ADET', value: `${adet} EA`, isBold: true },
          { label: 'TARİH', value: tarih }
        ];

        const slipCell = createDocxLabelCell(
          selectedSlipType === 'raf' ? 'OGM HAVACILIK - RAF SLİBİ' : 'OGM HAVACILIK - GİRİŞ SLİBİ',
          fields,
          qrBase64,
          barcodeBase64,
          barcodeValue,
          100
        );

        const doc = new DocxDocument({
          creator: "OGM Depo Sistemi",
          description: "Depo Malzeme Slibi",
          compatibility: { version: 12 },
          sections: [
            {
              properties: {
                page: {
                  margin: { top: 400, bottom: 400, left: 400, right: 400 }
                }
              },
              children: [
                new DocxTable({
                  width: { size: 4535, type: WidthType.DXA }, // Exact ~80mm width for termal compatibility
                  alignment: AlignmentType.CENTER,
                  rows: [new DocxTableRow({ children: [slipCell] })]
                }),
                new DocxParagraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 100 },
                  children: [
                    new DocxTextRun({
                      text: 'OGM Orman Havacılık - 80mm Termal Çıktı Standartı',
                      size: 12,
                      color: '64748b',
                      font: 'Calibri'
                    })
                  ]
                })
              ]
            }
          ]
        });

        const docxBlob = await DocxPacker.toBlob(doc);
        const url = URL.createObjectURL(docxBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `OGM_Slip_${Date.now()}.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        if (showNotification) {
          showNotification('✅ Word belgesi yerel Docx elementleri ile (Office 2007 uyumlu) oluşturuldu ve indirildi.');
        }
        return;
      }

      // ----------------------------------------------------
      // BULK PRINT BRANCH (Multiple Raf Labels - Table-based A4 layout)
      // ----------------------------------------------------
      let docTitle = 'OGM HAVACILIK - RAF LOKASYON VE BARKOD ETİKETLERİ (A4 TOPLU BASIM)';
      const tableRows: DocxTableRow[] = [];
      const itemsToExport = inventory.slice(0, 500);
      const totalItems = itemsToExport.length;

      for (let i = 0; i < totalItems; i += 2) {
        setBulkExportProgress(Math.round(((i + 1) / totalItems) * 100));
        await new Promise(r => setTimeout(r, 5));

        const item1 = itemsToExport[i];
        const item2 = itemsToExport[i + 1];

        const buildCellFromItem = async (item: DepoItem | undefined) => {
          if (!item) {
            return new DocxTableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
                bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
                left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
                right: { style: BorderStyle.NONE, size: 0, color: 'auto' }
              },
              children: [new DocxParagraph({})]
            });
          }

          const itemPn = item.partNumber || item.pn || '-';
          const itemSn = item.serialAndNotes || item.sn || '-';
          const itemLoc = item.lokasyonNo || item.location || 'A-01';
          const itemQty = item.ankaraMevcut || item.toplamStok || 0;
          const itemDesc = item.description || item.name || '-';
          const hasShelf = item.hasShelfLife === 'EVET' || item.hasShelfLife === true || String(item.hasShelfLife).toUpperCase().includes('EVET');
          const shelfDate = item.shelfLifeDate || '-';

          const bValue = `LOC-${itemLoc.replace(/[^a-zA-Z0-9]/g, '').toUpperCase()}`;
          const qContent = `OGM-DEPO-RAF|LOC:${itemLoc}|PN:${itemPn}|QTY:${itemQty}|DATE:${new Date().toLocaleDateString('tr-TR')}`;

          const barcodeBase64 = generateLocalBarcodeDataUrl(bValue);
          const qrBase64 = await generateLocalQrDataUrl(qContent);

          const fields = [
            { label: 'LOKASYON', value: itemLoc, isBold: true },
            { label: 'MALZEME', value: itemDesc },
            { label: 'P/N', value: itemPn, isBold: true },
            { label: 'S/N', value: itemSn },
            { label: 'STOK', value: `${itemQty} Adet`, isBold: true },
            { label: 'RAF ÖMRÜ', value: hasShelf ? `EVET (${shelfDate})` : 'HAYIR' },
            { label: 'TARİH', value: new Date().toLocaleDateString('tr-TR') }
          ];

          return createDocxLabelCell('OGM HAVACILIK - RAF LOKASYON ETİKETİ', fields, qrBase64, barcodeBase64, bValue, 50, BorderStyle.DASHED, 'CBD5E1');
        };

        const c1 = await buildCellFromItem(item1);
        const c2 = await buildCellFromItem(item2);

        tableRows.push(
          new DocxTableRow({
            children: [c1, c2]
          })
        );
      }

      const docxTable = new DocxTable({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
          bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
          left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
          right: { style: BorderStyle.NONE, size: 0, color: 'auto' }
        },
        rows: tableRows
      });

      const doc = new DocxDocument({
        compatibility: { version: 12 },
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: 720,
                  bottom: 720,
                  left: 720,
                  right: 720
                }
              }
            },
            children: [
              new DocxParagraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 120 },
                children: [
                  new DocxTextRun({
                    text: docTitle,
                    bold: true,
                    size: 18,
                    font: 'Calibri'
                  })
                ]
              }),
              new DocxParagraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 240 },
                children: [
                  new DocxTextRun({
                    text: `Oluşturulma Tarihi: ${new Date().toLocaleString('tr-TR')} | OGM Orman Havacılık Depo Yönetimi`,
                    size: 14,
                    color: '334155',
                    font: 'Calibri'
                  }),
                  new DocxTextRun({
                    text: '(Standart Microsoft Word .docx OpenXML Formatı)',
                    size: 12,
                    color: '64748b',
                    break: 1,
                    font: 'Calibri'
                  })
                ]
              }),
              docxTable
            ]
          }
        ]
      });

      const docxBlob = await DocxPacker.toBlob(doc);
      const url = URL.createObjectURL(docxBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `OGM_Havacilik_Toplu_Barkod_Word_${Date.now()}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      if (showNotification) {
        showNotification('🚀 Toplu barkod etiketleri MS Word (.docx) formatında başarıyla indirildi.');
      }
    } catch (e: any) {
      console.warn("Word export error:", e);
      alert('Word dosyası oluşturulurken hata: ' + e.message);
    } finally {
      setBulkExportProgress(null);
    }
  };

  // Bluetooth Printer Connection Handler
  const connectBluetoothPrinter = async () => {
    try {
      const nav = navigator as any;
      if (!nav.bluetooth) {
        alert("Tarayıcınız Web Bluetooth desteklemiyor. Lütfen Google Chrome veya Edge kullanın.");
        return;
      }
      const device = await nav.bluetooth.requestDevice({
        acceptAllDevices: true,
        optionalServices: ['00001101-0000-1000-8000-00805f9b34fb', '000018f0-0000-1000-8000-00805f9b34fb']
      });
      setBtDevice(device);
      setUsbDevice(null);
      setSerialPort(null);
      setPrinterConnectionStatus(`🟢 Bluetooth: ${device.name || 'Termal Cihaz'}`);
      if (showNotification) {
        showNotification(`✅ ${device.name || 'Termal Yazıcı'} Bluetooth üzerinden bağlandı!`);
      }
    } catch (err: any) {
      console.warn("Bluetooth printer connection error:", err);
    }
  };

  // USB Printer Connection Handler
  const connectUsbPrinter = async () => {
    try {
      const nav = navigator as any;
      if (!nav.usb) {
        alert("Tarayıcınız Web USB desteklemiyor. Lütfen Google Chrome veya Edge kullanın.");
        return;
      }
      const device = await nav.usb.requestDevice({
        filters: []
      });
      setUsbDevice(device);
      setBtDevice(null);
      setSerialPort(null);
      setPrinterConnectionStatus(`🟢 USB: ${device.productName || 'Termal Cihaz'}`);
      if (showNotification) {
        showNotification(`✅ ${device.productName || 'USB Yazıcı'} USB üzerinden bağlandı!`);
      }
    } catch (err: any) {
      console.warn("USB printer connection error:", err);
    }
  };

  // Web Serial Printer Connection Handler (for COM / USB-Serial thermal printers)
  const connectSerialPrinter = async () => {
    try {
      const nav = navigator as any;
      if (!nav.serial) {
        alert("Tarayıcınız Web Serial desteklemiyor. Lütfen Google Chrome veya Edge kullanın.");
        return;
      }
      const port = await nav.serial.requestPort();
      await port.open({ baudRate: 9600 });
      setSerialPort(port);
      setBtDevice(null);
      setUsbDevice(null);
      setPrinterConnectionStatus(`🟢 COM / Seri Port Termal Cihaz`);
      if (showNotification) {
        showNotification(`✅ COM Seri Port Termal Yazıcı bağlandı!`);
      }
    } catch (err: any) {
      console.warn("Serial printer connection error:", err);
    }
  };

  const sendToBluetoothPrinter = async (text: string) => {
    if (!btDevice) return;
    try {
      if (showNotification) showNotification("⏳ Bluetooth yazıcıya veri gönderiliyor...");
      const server = await btDevice.gatt.connect();
      let service;
      try {
        service = await server.getPrimaryService('00001101-0000-1000-8000-00805f9b34fb');
      } catch (sErr) {
        const services = await server.getPrimaryServices();
        if (services.length > 0) service = services[0];
        else throw new Error("Bluetooth servis bulunamadı.");
      }
      
      const characteristics = await service.getCharacteristics();
      const writeChar = characteristics.find(c => c.properties.write || c.properties.writeWithoutResponse);
      if (!writeChar) {
        throw new Error("Yazıcı yazma özelliği (characteristic) desteklemiyor.");
      }

      const encoder = new TextEncoder();
      const printBytes = encoder.encode(text + "\n\n\n\n\x1d\x56\x41\x03");
      
      const chunkSize = 20;
      for (let i = 0; i < printBytes.length; i += chunkSize) {
        const chunk = printBytes.slice(i, i + chunkSize);
        await writeChar.writeValue(chunk);
      }
      
      if (showNotification) showNotification("✅ Çıktı Bluetooth yazıcıya başarıyla iletildi!");
    } catch (err: any) {
      console.error("Bluetooth print error:", err);
      window.print();
    }
  };

  const sendToUsbPrinter = async (text: string) => {
    if (!usbDevice) return;
    try {
      if (showNotification) showNotification("⏳ USB yazıcıya veri gönderiliyor...");
      await usbDevice.open();
      await usbDevice.selectConfiguration(1);
      
      const iface = usbDevice.configuration.interfaces[0];
      await usbDevice.claimInterface(iface.interfaceNumber);
      
      const alt = iface.alternates[0];
      const endpoint = alt.endpoints.find((e: any) => e.direction === 'out' && e.type === 'bulk');
      if (!endpoint) {
        throw new Error("Yazıcı BULK OUT endpoint'i desteklemiyor.");
      }

      const encoder = new TextEncoder();
      const printBytes = encoder.encode(text + "\n\n\n\n\x1d\x56\x41\x03");
      
      await usbDevice.transferOut(endpoint.endpointNumber, printBytes);
      if (showNotification) showNotification("✅ Çıktı USB yazıcıya başarıyla iletildi!");
    } catch (err: any) {
      console.error("USB print error:", err);
      window.print();
    }
  };

  const sendToSerialPrinter = async (text: string) => {
    if (!serialPort) return;
    try {
      if (showNotification) showNotification("⏳ Seri Port yazıcıya veri gönderiliyor...");
      const textEncoder = new TextEncoderStream();
      const writableStreamClosed = textEncoder.readable.pipeTo(serialPort.writable);
      const writer = textEncoder.writable.getWriter();
      await writer.write(text + "\n\n\n\n\x1d\x56\x41\x03");
      writer.releaseLock();
      if (showNotification) showNotification("✅ Çıktı Seri Port termal cihaza iletildi!");
    } catch (err: any) {
      console.error("Serial print error:", err);
      window.print();
    }
  };

  const handlePrint = () => {
    const slipText = generateAsciiSlip();
    if (btDevice) {
      sendToBluetoothPrinter(slipText);
    } else if (usbDevice) {
      sendToUsbPrinter(slipText);
    } else if (serialPort) {
      sendToSerialPrinter(slipText);
    } else {
      // Single Slip printing via dynamic Iframe for perfect receipt layout on mobile & PC
      const slipEl = document.getElementById('ogm-printable-slip');
      if (!slipEl) {
        window.print();
        return;
      }

      try {
        const iframe = document.createElement('iframe');
        iframe.style.position = 'fixed';
        iframe.style.right = '0';
        iframe.style.bottom = '0';
        iframe.style.width = '0';
        iframe.style.height = '0';
        iframe.style.border = '0';
        document.body.appendChild(iframe);

        const doc = iframe.contentWindow?.document || iframe.contentDocument;
        if (doc) {
          doc.open();
          doc.write(`
            <html>
              <head>
                <title>OGM Orman Havacılık Depo Slibi</title>
                <style>
                  @import url('https://fonts.googleapis.com/css2?family=Courier+Prime:wght@400;700&display=swap');
                  body {
                    margin: 0;
                    padding: 8px;
                    background-color: white;
                    color: black;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                  }
                  @media print {
                    @page {
                      size: 80mm auto; /* Perfect standard thermal receipt width */
                      margin: 0;
                    }
                    body {
                      padding: 4mm;
                    }
                  }
                  /* Include the exact visual slip styling we need */
                  .text-center { text-align: center; }
                  .font-mono { font-family: 'Courier New', Courier, monospace; }
                  .font-bold { font-weight: bold; }
                  .border-dashed { border-style: dashed; }
                  .border-b { border-bottom-width: 1px; }
                  .border-slate-800 { border-color: #1e293b; }
                  .py-2 { padding-top: 0.5rem; padding-bottom: 0.5rem; }
                  .my-2 { margin-top: 0.5rem; margin-bottom: 0.5rem; }
                  .space-y-1 > * + * { margin-top: 0.25rem; }
                  .flex { display: flex; }
                  .justify-between { justify-content: space-between; }
                  .items-center { align-items: center; }
                  .w-full { width: 100%; }
                  .h-10 { height: 2.5rem; }
                  .text-xs { font-size: 0.75rem; }
                  .text-[10px] { font-size: 10px; }
                  .text-sm { font-size: 0.875rem; }
                  .leading-tight { line-height: 1.25; }
                </style>
              </head>
              <body>
                ${slipEl.innerHTML}
                <script type="text/javascript">
                  window.onload = function() {
                    window.focus();
                    window.print();
                    setTimeout(function() {
                      window.parent.document.body.removeChild(iframe);
                    }, 1000);
                  }
                </script>
              </body>
            </html>
          `);
          doc.close();
        }
      } catch (e) {
        console.warn("Iframe print error:", e);
        window.print();
      }
    }
  };

  // Robust HTML A4 Printing (Replaces Word for bulk printing)
  const handlePrintA4HTML = async () => {
    const itemsToExport = isBulkPrint ? (filteredBulkItems.length > 0 ? filteredBulkItems : inventory.filter((i: any) => Number(i.ankaraMevcut ?? i.toplamStok ?? i.quantity ?? 0) > 0)) : [];
    
    // Create printable HTML content
    let htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8" />
          <title>OGM Havacılık A4 Toplu Etiket Basımı</title>
          <style>
            @media print {
              @page { size: A4 portrait; margin: 5mm; }
              body { margin: 0; }
              .no-print { display: none; }
            }
            body { 
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
              color: black;
              background: white;
              margin: 5mm;
              padding: 0;
            }
            .grid {
              display: grid;
              grid-template-columns: repeat(2, 95mm);
              gap: 4mm;
              width: 100%;
              justify-content: start;
              align-items: start;
            }
            .label-box {
              width: 95mm;
              height: 52mm;
              box-sizing: border-box;
              border: 1.5px dashed #333;
              border-radius: 4px;
              padding: 3mm;
              display: flex;
              flex-direction: row;
              justify-content: space-between;
              align-items: stretch;
              background: #fff;
              page-break-inside: avoid;
            }
            .header-vertical {
              writing-mode: vertical-rl;
              transform: rotate(180deg);
              font-size: 8px;
              font-weight: 800;
              color: #94a3b8;
              text-transform: uppercase;
              letter-spacing: 0.5px;
              white-space: nowrap;
              padding-right: 2mm;
              border-right: 1px solid #e2e8f0;
              margin-right: 2.5mm;
              display: flex;
              align-items: center;
              justify-content: center;
            }
            .content-area {
              flex: 1;
              display: flex;
              flex-direction: column;
              justify-content: space-between;
              font-size: 9px;
              padding-right: 2mm;
            }
            .row {
              display: flex;
              align-items: baseline;
              gap: 3px;
              line-height: 1.25;
            }
            .lbl {
              font-weight: 800;
              color: #0f172a;
              font-size: 9px;
              min-width: 28px;
              white-space: nowrap;
            }
            .val {
              font-weight: 700;
              color: #0f172a;
              font-size: 9.5px;
              word-break: break-word;
            }
            .barcode-vertical-section {
              display: flex;
              flex-direction: column;
              align-items: center;
              justify-content: space-between;
              width: 28mm;
              margin-left: 1.5mm;
              padding-left: 1.5mm;
              border-left: 1px dashed #cbd5e1;
            }
            .qr-code-img {
              width: 14mm;
              height: 14mm;
              object-fit: contain;
            }
            .barcode-container {
              display: flex;
              flex-direction: column;
              align-items: center;
              justify-content: center;
              margin-top: 1mm;
              width: 100%;
            }
            .barcode-img-vert {
              width: 100%;
              height: 18mm;
              object-fit: fill;
            }
            .barcode-txt {
              font-size: 5.5px;
              font-family: monospace;
              font-weight: bold;
              color: #0f172a;
              white-space: nowrap;
              margin-top: 1px;
            }
          </style>
        </head>
        <body>
          <div class="grid">
    `;

    const currentDateStr = new Date().toLocaleDateString('tr-TR');

    if (!isBulkPrint) {
      // Single Label
      const bImg = generateLocalBarcodeDataUrl(barcodeValue);
      const qImg = await generateLocalQrDataUrl(qrContent);
      
      if (selectedSlipType === 'etiket') {
        htmlContent += `
          <div class="label-box" style="border: 2px solid #0f172a; border-radius: 10px; padding: 8px; background: #fff; width: 100%;">
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0f172a; padding-bottom: 4px; margin-bottom: 6px;">
              <strong style="font-size: 11px; font-family: sans-serif;">OGM HAVACILIK</strong>
              <span style="font-size: 8px; font-weight: bold; background: #fef3c7; color: #78350f; border: 1px solid #fde68a; padding: 2px 5px; border-radius: 4px; font-family: monospace;">TEÇHİZAT ETİKETİ</span>
            </div>
            <div style="display: flex; gap: 8px; align-items: center;">
              <div style="flex: 1; font-size: 9px; line-height: 1.35; font-family: monospace;">
                <div><b>TNM:</b> <span style="color: #0f172a; font-weight: bold;">${tanimIcerik || malzemeDesc}</span></div>
                <div><b>P/N:</b> ${pn && pn.trim() !== '' && pn.trim() !== '-' ? pn : '<span style="color:#94a3b8; font-style:italic;">-</span>'}</div>
                ${sn && sn.trim() !== '' && sn.trim() !== '-' ? `<div><b>S/N:</b> <span style="font-weight: bold;">${sn}</span></div>` : ''}
                <div><b>LOK:</b> ${lokasyonKodu || depoYeri}</div>
                ${kalibrasyonTarihi && kalibrasyonTarihi.trim() !== '' && kalibrasyonTarihi.trim() !== '-' ? `<div><b>KAL/BAK:</b> <span style="font-weight: bold; color: #065f46;">${kalibrasyonTarihi}</span></div>` : ''}
                <div style="font-size: 7px; color: #64748b; margin-top: 2px; font-family: sans-serif;">Tarih: ${currentDateStr}</div>
              </div>
              <div style="width: 65px; height: 65px; border: 2px solid #cbd5e1; border-radius: 8px; overflow: hidden; background: #f8fafc; display: flex; align-items: center; justify-content: center; shrink: 0;">
                ${imageUrlInput ? `<img src="${imageUrlInput}" style="width: 100%; height: 100%; object-fit: cover;" />` : `<span style="font-size: 24px;">🔧</span>`}
              </div>
            </div>
            <div style="border-top: 1px solid #e2e8f0; margin-top: 6px; padding-top: 4px; text-align: center;">
              <img src="${qImg}" style="width: 48px; height: 48px; border: 1px solid #cbd5e1; padding: 2px; border-radius: 6px; background: #fff;" />
            </div>
          </div>
        `;
      } else {
        const fields = selectedSlipType === 'raf' ? [
          { l: 'LOK:', v: lokasyonKodu },
          { l: 'BÖL:', v: bolgeZone },
          { l: 'TNM:', v: tanimIcerik },
          { l: 'STK:', v: `${guncelSayim} EA` },
          { l: 'Tarih:', v: currentDateStr }
        ] : [
          { l: 'İŞLEM:', v: islemTransfer },
          { l: 'MLZ:', v: malzemeDesc },
          { l: 'P/N:', v: pn },
          { l: 'ADET:', v: `${adet} EA` },
          { l: 'Tarih:', v: currentDateStr }
        ];

        htmlContent += `
          <div class="label-box">
            <div class="header-vertical">OGM HAVACILIK - ${selectedSlipType === 'raf' ? 'RAF SLİBİ' : 'GİRİŞ SLİBİ'}</div>
            <div class="content-area">
              ${fields.map(f => `<div class="row"><span class="lbl">${f.l}</span><span class="val">${f.v}</span></div>`).join('')}
            </div>
            <div class="barcode-vertical-section">
              <img src="${qImg}" class="qr-code-img" />
              <div class="barcode-container">
                <img src="${bImg}" class="barcode-img-vert" />
              </div>
            </div>
          </div>
        `;
      }
    } else {
      // Bulk Labels
      for (const item of itemsToExport) {
        const itemPn = item.partNumber || item.pn || '-';
        const itemLoc = item.lokasyonNo || item.location || 'A-01';
        const bValue = `LOC-${itemLoc.replace(/[^a-zA-Z0-9]/g, '').toUpperCase()}`;
        const qContent = `OGM-DEPO-RAF|LOC:${itemLoc}|PN:${itemPn}|QTY:${item.ankaraMevcut || 0}`;
        
        const bImg = generateLocalBarcodeDataUrl(bValue);
        const qImg = await generateLocalQrDataUrl(qContent);

        htmlContent += `
          <div class="label-box">
            <div class="header-vertical">OGM HAVACILIK - RAF SLİBİ</div>
            <div class="content-area">
              <div class="row"><span class="lbl">LOK:</span><span class="val">${itemLoc}</span></div>
              <div class="row"><span class="lbl">BÖL:</span><span class="val">${item.zone || 'GENEL'}</span></div>
              <div class="row"><span class="lbl">TNM:</span><span class="val">${item.description || item.name || '-'}</span></div>
              <div class="row"><span class="lbl">STK:</span><span class="val">${item.ankaraMevcut || 0} EA</span></div>
              <div class="row"><span class="lbl">Tarih:</span><span class="val">${currentDateStr}</span></div>
            </div>
            <div class="barcode-vertical-section">
              <img src="${qImg}" class="qr-code-img" />
              <div class="barcode-container">
                <img src="${bImg}" class="barcode-img-vert" />
              </div>
            </div>
          </div>
        `;
      }
    }

    htmlContent += `
          </div>
        </body>
      </html>
    `;

    const printWin = window.open('', '_blank');
    if (printWin) {
      printWin.document.write(htmlContent);
      printWin.document.close();
      printWin.focus();
      setTimeout(() => {
        printWin.print();
      }, 500);
    }
  };

  const handleCopyText = () => {
    const text = generateAsciiSlip();
    navigator.clipboard.writeText(text).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
      if (showNotification) {
        showNotification('📋 76mm Slip metni kopyalandı.');
      }
    });
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[10200] bg-black/75 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
      <div className="bg-slate-900 text-slate-100 rounded-3xl border border-slate-700 shadow-2xl max-w-5xl w-full max-h-[92vh] flex flex-col overflow-hidden">
        
        {/* MODAL HEADER */}
        <div className="bg-slate-950 px-5 py-3.5 border-b border-slate-800 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-emerald-500/20 text-emerald-400 rounded-xl border border-emerald-500/30">
              <Printer className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-extrabold text-white tracking-wide flex items-center gap-2">
                <span>OGM DEPO TERMAL SLİP VE BARKOD YAZICI</span>
                <span className="text-[10px] bg-emerald-500/20 text-emerald-300 font-mono px-2 py-0.5 rounded-full border border-emerald-500/30 font-bold">
                  76mm / 80mm Standart
                </span>
              </h2>
              <p className="text-[11px] text-slate-400">
                Karekod (QR) ve 1D Barkodlu termal etiket basımı, Word (.doc) aktarımı ve doğrudan cihaz iletişimi
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {/* Direct Hardware Connection Toggle */}
            <button
              type="button"
              onClick={() => setIsDeviceMenuOpen(!isDeviceMenuOpen)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-bold border border-slate-600 transition cursor-pointer"
            >
              <Radio className="w-3.5 h-3.5 text-emerald-400 animate-pulse" />
              <span className="truncate max-w-[150px]">{printerConnectionStatus}</span>
            </button>

            <button
              type="button"
              onClick={onClose}
              className="p-1.5 hover:bg-slate-800 text-slate-400 hover:text-white rounded-xl transition-colors cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* DEVICE CHOOSER DROPDOWN */}
        {isDeviceMenuOpen && (
          <div className="bg-slate-950/95 border-b border-slate-800 p-3 flex items-center justify-between gap-3 text-xs flex-wrap animate-fade-in">
            <div className="text-slate-300 font-bold flex items-center gap-1.5">
              <span>Termal Cihaz Eşleştir:</span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={connectBluetoothPrinter}
                className="flex items-center gap-1.5 px-3 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-bold transition cursor-pointer"
              >
                <Bluetooth className="w-3.5 h-3.5" />
                <span>Bluetooth Cihaz Bul</span>
              </button>
              <button
                type="button"
                onClick={connectUsbPrinter}
                className="flex items-center gap-1.5 px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg font-bold transition cursor-pointer"
              >
                <Usb className="w-3.5 h-3.5" />
                <span>USB Termal Cihaz Bul</span>
              </button>
              <button
                type="button"
                onClick={connectSerialPrinter}
                className="flex items-center gap-1.5 px-3 py-1 bg-teal-600 hover:bg-teal-500 text-white rounded-lg font-bold transition cursor-pointer"
              >
                <Printer className="w-3.5 h-3.5" />
                <span>COM / Seri Port</span>
              </button>
            </div>
          </div>
        )}

        {/* MODAL CONTENT */}
        <div className="p-4 sm:p-6 overflow-y-auto flex-1 space-y-5">
          
          {/* TOP SELECTION BAR */}
          <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-950 p-3 rounded-2xl border border-slate-800">
            <div className="flex items-center gap-2">
              {initialType === 'giris' ? (
                /* MALZEME GİRİŞİ SONRASI YÖNLENDİRME - SADECE MALZEME GİRİŞ SLİBİ GÖRÜNÜR */
                <button
                  type="button"
                  className="px-4 py-2 rounded-xl text-xs font-bold bg-emerald-600 text-white shadow-md flex items-center gap-2 cursor-default font-sans"
                >
                  <ArrowDownLeft className="w-4 h-4 text-emerald-200" />
                  <span>Malzeme Giriş Slibi</span>
                </button>
              ) : isTechizatMode ? (
                <>
                  <button
                    type="button"
                    onClick={() => { setSelectedSlipType('etiket'); setIsBulkPrint(false); }}
                    className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                      selectedSlipType === 'etiket' && !isBulkPrint
                        ? 'bg-amber-600 text-white shadow-md'
                        : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    }`}
                  >
                    <Wrench className="w-3.5 h-3.5 text-amber-300" />
                    <span>1. Tek Teçhizat / Alet Etiketi Bas</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => { setIsBulkPrint(true); }}
                    className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                      isBulkPrint
                        ? 'bg-blue-600 text-white shadow-md'
                        : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    }`}
                  >
                    <Layers className="w-3.5 h-3.5" />
                    <span>2. Tüm Teçhizat &amp; Alet Etiketlerini Bas ({(techizatList && techizatList.length > 0 ? techizatList : inventory).length} Adet)</span>
                  </button>
                </>
              ) : isGirisOnlyMode ? (
                <>
                  <button
                    type="button"
                    onClick={() => { setSelectedSlipType('giris'); setIsBulkPrint(false); }}
                    className="px-3.5 py-2 rounded-xl text-xs font-bold bg-emerald-600 text-white shadow-md flex items-center gap-2 cursor-default"
                  >
                    <ArrowDownLeft className="w-3.5 h-3.5" />
                    <span>Malzeme Giriş Slibi</span>
                  </button>
                </>
              ) : (
                <>
                  <button
                    type="button"
                    onClick={() => { setSelectedSlipType('raf'); setIsBulkPrint(false); }}
                    className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                      selectedSlipType === 'raf' && !isBulkPrint
                        ? 'bg-emerald-600 text-white shadow-md'
                        : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    }`}
                  >
                    <MapPin className="w-3.5 h-3.5" />
                    <span>1. Raf Lokasyon Slibi</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => { setSelectedSlipType('giris'); setIsBulkPrint(false); }}
                    className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                      selectedSlipType === 'giris' && !isBulkPrint
                        ? 'bg-emerald-600 text-white shadow-md'
                        : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    }`}
                  >
                    <ArrowDownLeft className="w-3.5 h-3.5" />
                    <span>2. Malzeme Giriş Slibi</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => { setIsBulkPrint(true); }}
                    className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                      isBulkPrint
                        ? 'bg-blue-600 text-white shadow-md'
                        : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    }`}
                  >
                    <Layers className="w-3.5 h-3.5" />
                    <span>3. A4 Toplu Depo Etiket Basımı ({inventory.length} Adet)</span>
                  </button>
                </>
              )}
            </div>

            {/* QUICK AUTOFILL SEARCH */}
            {!isBulkPrint && (
              <div className="relative flex-1 max-w-xs">
                <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-500" />
                <input
                  type="text"
                  placeholder={isTechizatMode ? "Teçhizat adından otomatik doldur..." : "Envanterden otomatik doldur..."}
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl pl-8 pr-3 py-1.5 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-amber-500 font-medium"
                />
                {searchQuery && (
                  <div className="absolute left-0 right-0 top-full mt-1 bg-slate-950 border border-slate-700 rounded-xl max-h-48 overflow-y-auto z-50 shadow-2xl p-1 text-xs">
                    {(isTechizatMode && techizatList && techizatList.length > 0 ? techizatList : inventory)
                      .filter((i: any) => {
                        const name = (i.description || i.name || i.ALET_ADI || i.TECHIZAT_ADI || i.tanim || '').toLowerCase();
                        const partN = (i.partNumber || i.pn || i.PARCA_NO || i.pnNo || '').toLowerCase();
                        const loc = (i.lokasyonNo || i.location || i.DEPO_YERI || i.HANGAR_LOKASYON || '').toLowerCase();
                        const q = searchQuery.toLowerCase();
                        return name.includes(q) || partN.includes(q) || loc.includes(q);
                      })
                      .slice(0, 8)
                      .map((item: any, idx: number) => {
                        const itemDesc = item.description || item.name || item.ALET_ADI || item.TECHIZAT_ADI || item.tanim || 'TEÇHİZAT';
                        const itemPn = item.partNumber || item.pn || item.PARCA_NO || item.pnNo || '-';
                        const itemLoc = item.lokasyonNo || item.location || item.DEPO_YERI || item.HANGAR_LOKASYON || 'HANGAR';
                        const itemImg = getDirectDriveImageUrl(item.imageUrl || item.fotoUrl || item.DriveFoto || item.foto || item.GÖRSEL || item.FOTOĞRAF || '');

                        return (
                          <div
                            key={idx}
                            onClick={() => {
                              setTanimIcerik(itemDesc);
                              setMalzemeDesc(itemDesc);
                              setPn(itemPn);
                              setSn(item.serialAndNotes || item.sn || item.SERI_NO || item.snNo || '-');
                              setLokasyonKodu(itemLoc);
                              setDepoYeri(itemLoc);
                              setGuncelSayim(item.ankaraMevcut ?? item.gelen ?? item.toplamStok ?? item.ADET ?? item.MEVCUT ?? 1);
                              setAdet(item.ankaraMevcut ?? item.gelen ?? item.toplamStok ?? item.ADET ?? item.MEVCUT ?? 1);
                              setImageUrlInput(itemImg);
                              setSearchQuery('');
                            }}
                            className="p-2 hover:bg-amber-950/40 rounded-lg cursor-pointer flex justify-between items-center text-slate-300 hover:text-white"
                          >
                            <div className="truncate pr-2">
                              <strong className="text-amber-400">{itemPn}</strong> - {itemDesc}
                            </div>
                            <span className="text-[10px] bg-slate-800 px-1.5 py-0.5 rounded text-slate-400 shrink-0 font-mono">
                              {itemLoc}
                            </span>
                          </div>
                        );
                      })}
                  </div>
                )}
              </div>
            )}
          </div>

          {/* MAIN GRID */}
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
            
            {/* LEFT FORM EDITORS */}
            <div className="lg:col-span-5 bg-slate-950 p-4 rounded-2xl border border-slate-800 space-y-3 text-xs">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <span className="font-extrabold text-slate-300 uppercase tracking-wider font-sans">
                  {isBulkPrint ? 'Toplu Etiket Parametreleri' : selectedSlipType === 'etiket' ? 'Teçhizat & Alet Etiketi Parametreleri' : selectedSlipType === 'raf' ? 'Raf Slibi Parametreleri' : 'Giriş Slibi Parametreleri'}
                </span>
              </div>

              {isBulkPrint ? (
                <div className="space-y-3 text-slate-300">
                  <div>
                    <label className="block text-[10px] font-bold text-amber-400 uppercase font-mono mb-1">
                      1. Depo Kategori Seçimi
                    </label>
                    <div className="grid grid-cols-2 gap-2">
                      <button
                        type="button"
                        onClick={() => setBulkDepoCategory('sarf')}
                        className={`px-3 py-2 rounded-xl text-xs font-bold transition flex items-center justify-center gap-1 cursor-pointer border ${
                          bulkDepoCategory === 'sarf'
                            ? 'bg-amber-600 text-white border-amber-400 shadow-md'
                            : 'bg-slate-900 text-slate-400 border-slate-800 hover:bg-slate-800'
                        }`}
                      >
                        📦 Sarf Parça Depo
                      </button>
                      <button
                        type="button"
                        onClick={() => setBulkDepoCategory('kimyasal')}
                        className={`px-3 py-2 rounded-xl text-xs font-bold transition flex items-center justify-center gap-1 cursor-pointer border ${
                          bulkDepoCategory === 'kimyasal'
                            ? 'bg-purple-600 text-white border-purple-400 shadow-md'
                            : 'bg-slate-900 text-slate-400 border-slate-800 hover:bg-slate-800'
                        }`}
                      >
                        🧪 Kimyasal Depo
                      </button>
                    </div>
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-amber-400 uppercase font-mono mb-1">
                      2. Hangi Depo? (Lokasyon Süzgeci)
                    </label>
                    <select
                      value={bulkSelectedDepo}
                      onChange={(e) => setBulkSelectedDepo(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-white font-bold text-xs focus:border-amber-500 focus:outline-none cursor-pointer"
                    >
                      <option value="ANKARA DEPO">🏢 Ankara Depo</option>
                      <option value="ESKİŞEHİR DEPO">🏬 Eskişehir Depo</option>
                      <option value="İZMİR DEPO">🏬 İzmir Depo</option>
                      <option value="GÜVERCİNLİK DEPO">🏬 Güvercinlik Depo</option>
                      <option value="TÜM DEPOLAR">🌐 Tüm Depolar</option>
                    </select>
                  </div>

                  <div className="p-3 bg-emerald-950/40 border border-emerald-700/50 rounded-xl text-emerald-300 text-[11px] space-y-1.5 font-sans">
                    <div className="font-extrabold text-xs text-emerald-200 flex items-center gap-1.5">
                      <span>🏷️</span> Yazdırılacak Etiket Sayısı: <span className="text-amber-300 font-mono text-sm underline">{filteredBulkItems.length} Adet</span>
                    </div>
                    <div className="text-[10px] text-emerald-400">
                      ✓ Stoku 0 olan ürünler otomatik elenmiştir (0 hariç hepsi basılır).
                    </div>
                    <div className="text-[10px] text-slate-400">
                      ✓ A4 formatında QR karekod ve lokasyon detaylarıyla hazırlanır.
                    </div>
                  </div>
                </div>
              ) : selectedSlipType === 'etiket' ? (
                <div className="space-y-2.5">
                  <div className="relative">
                    <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1 flex items-center justify-between">
                      <span>Teçhizat / Alet Tanımı</span>
                      <span className="text-[9px] text-amber-400 font-sans font-extrabold flex items-center gap-1">
                        <span>✨</span> Otomatik Doldurma Açık
                      </span>
                    </label>
                    <input
                      type="text"
                      value={tanimIcerik}
                      onFocus={() => setIsAutocompleteOpen(true)}
                      onChange={e => {
                        setTanimIcerik(e.target.value);
                        setMalzemeDesc(e.target.value);
                        setIsAutocompleteOpen(true);
                      }}
                      placeholder="Alet veya teçhizat adı yazın (Örn: A, Mule, Jack...)..."
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-bold focus:border-amber-500 focus:outline-none"
                    />

                    {/* Autocomplete Dropdown List */}
                    {isAutocompleteOpen && (
                      <div className="absolute left-0 right-0 top-full mt-1 bg-slate-950 border-2 border-amber-500 rounded-xl max-h-56 overflow-y-auto z-[100] shadow-2xl p-1 text-xs divide-y divide-slate-800">
                        {(isTechizatMode && techizatList && techizatList.length > 0 ? techizatList : inventory)
                          .filter((item: any) => {
                            if (!tanimIcerik.trim()) return true;
                            const q = tanimIcerik.toLowerCase().trim();
                            const name = (item.description || item.name || item.ALET_ADI || item.TECHIZAT_ADI || item.tanim || '').toLowerCase();
                            const partN = (item.partNumber || item.pn || item.PARCA_NO || item.pnNo || '').toLowerCase();
                            const sn = (item.serialAndNotes || item.sn || item.SERI_NO || item.snNo || '').toLowerCase();
                            return name.includes(q) || partN.includes(q) || sn.includes(q);
                          })
                          .slice(0, 10)
                          .map((item: any, idx: number) => {
                            const itemDesc = item.description || item.name || item.ALET_ADI || item.TECHIZAT_ADI || item.tanim || 'TEÇHİZAT';
                            const itemPn = item.partNumber || item.pn || item.PARCA_NO || item.pnNo || '-';
                            const itemSn = item.serialAndNotes || item.sn || item.SERI_NO || item.snNo || '-';
                            const itemLoc = item.lokasyonNo || item.location || item.DEPO_YERI || item.HANGAR_LOKASYON || 'HANGAR';
                            const itemQty = item.ankaraMevcut ?? item.gelen ?? item.toplamStok ?? item.ADET ?? item.MEVCUT ?? 1;
                            const itemImg = getDirectDriveImageUrl(item.imageUrl || item.fotoUrl || item.DriveFoto || item.foto || item.GÖRSEL || item.FOTOĞRAF || '');

                            return (
                              <div
                                key={idx}
                                onClick={() => {
                                  setTanimIcerik(itemDesc);
                                  setMalzemeDesc(itemDesc);
                                  setPn(itemPn);
                                  setSn(itemSn);
                                  setLokasyonKodu(itemLoc);
                                  setDepoYeri(itemLoc);
                                  setGuncelSayim(String(itemQty));
                                  setAdet(itemQty);
                                  setImageUrlInput(itemImg);
                                  setIsAutocompleteOpen(false);
                                }}
                                className="p-2 hover:bg-amber-950/60 rounded-lg cursor-pointer flex justify-between items-center text-slate-200 hover:text-white transition"
                              >
                                <div>
                                  <strong className="text-amber-400 font-sans block text-xs">{itemDesc}</strong>
                                  <span className="text-[10px] text-slate-400 font-mono">
                                    P/N: <span className="text-slate-200">{itemPn}</span> | S/N: {itemSn}
                                  </span>
                                </div>
                                <div className="text-right shrink-0 font-mono text-[10px] pl-2">
                                  <span className="bg-slate-800 text-slate-300 px-1.5 py-0.5 rounded block mb-0.5 font-bold">
                                    {itemLoc}
                                  </span>
                                  <span className="text-emerald-400 font-bold">{itemQty} ADET</span>
                                </div>
                              </div>
                            );
                          })}
                        
                        <div className="p-1.5 text-[10px] text-center text-slate-400 font-sans bg-slate-900/80 rounded-b-lg font-bold flex items-center justify-between">
                          <span>Listedeki alete tıklayarak otomatik doldurun</span>
                          <button 
                            type="button" 
                            onClick={(e) => { e.stopPropagation(); setIsAutocompleteOpen(false); }}
                            className="text-amber-400 hover:underline px-1"
                          >
                            ✕ Kapat
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">P/N (Parça No)</label>
                      <input
                        type="text"
                        value={pn}
                        onChange={e => setPn(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-amber-300 font-mono font-bold focus:border-amber-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1 flex items-center justify-between">
                        <span>S/N (Seri No)</span>
                        {availableSerialNumbers.length > 0 && (
                          <span className="text-[9px] text-red-400 font-sans font-bold bg-red-950/80 border border-red-500/50 px-1.5 py-0.5 rounded animate-pulse">
                            ⚠️ Seri No Seçin!
                          </span>
                        )}
                      </label>
                      {availableSerialNumbers.length > 0 ? (
                        <select
                          value={sn}
                          onChange={e => setSn(e.target.value)}
                          className={`w-full border rounded-lg px-2 py-1.5 font-mono font-bold text-xs focus:outline-none cursor-pointer transition-all ${
                            sn === '-' || !sn
                              ? 'bg-red-950/80 border-red-500 text-red-200 ring-2 ring-red-500/60 animate-pulse'
                              : 'bg-slate-900 border-slate-700 text-amber-300'
                          }`}
                        >
                          <option value="-">⚠️ Lütfen Seri No Seçiniz (-)</option>
                          {availableSerialNumbers.map((sNum, idx) => (
                            <option key={idx} value={sNum}>{sNum}</option>
                          ))}
                        </select>
                      ) : (
                        <input
                          type="text"
                          value={sn}
                          onChange={e => setSn(e.target.value)}
                          placeholder="SN Giriniz..."
                          className={`w-full border rounded-lg px-2.5 py-1.5 font-mono focus:outline-none transition-all ${
                            sn === '-' || !sn
                              ? 'bg-red-950/60 border-red-500/80 text-red-200 placeholder-red-400'
                              : 'bg-slate-900 border-slate-700 text-white'
                          }`}
                        />
                      )}
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">
                        Lokasyon / Hangar {availableLocations.length > 1 ? '(Listeden Seç)' : '(Sabit)'}
                      </label>
                      {availableLocations.length > 1 ? (
                        <select
                          value={lokasyonKodu}
                          onChange={e => {
                            setLokasyonKodu(e.target.value);
                            setDepoYeri(e.target.value);
                          }}
                          className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-white font-bold text-xs focus:border-amber-500 focus:outline-none cursor-pointer font-sans"
                        >
                          {availableLocations.map((loc, idx) => (
                            <option key={idx} value={loc}>{loc}</option>
                          ))}
                        </select>
                      ) : (
                        <div className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1.5 text-amber-300 font-bold text-xs font-mono">
                          {availableLocations[0] || lokasyonKodu || 'ANKARA HANGAR'}
                        </div>
                      )}
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">
                        Son Kontrol / Kalibrasyon Tarihi
                      </label>
                      <input
                        type="text"
                        value={kalibrasyonTarihi}
                        onChange={e => setKalibrasyonTarihi(e.target.value)}
                        placeholder="Örn: 15.08.2026..."
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-emerald-400 font-bold focus:border-amber-500 focus:outline-none font-mono"
                      />
                    </div>
                  </div>
                </div>
              ) : selectedSlipType === 'raf' ? (
                <div className="space-y-2.5">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Lokasyon Kodu</label>
                    <input
                      type="text"
                      value={lokasyonKodu}
                      onChange={e => setLokasyonKodu(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono font-bold focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Depo Bölgesi / Zone</label>
                    <input
                      type="text"
                      value={bolgeZone}
                      onChange={e => setBolgeZone(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Tanım / İçerik</label>
                    <input
                      type="text"
                      value={tanimIcerik}
                      onChange={e => setTanimIcerik(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Güncel Sayım</label>
                      <input
                        type="text"
                        value={guncelSayim}
                        onChange={e => setGuncelSayim(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-emerald-400 font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Sayım Tarihi</label>
                      <div className="flex items-center gap-1.5">
                        <input
                          type="text"
                          value={sonSayimTarih}
                          onChange={e => setSonSayimTarih(e.target.value)}
                          className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white focus:border-emerald-500 focus:outline-none"
                        />
                        <input
                          type="date"
                          onChange={(e) => {
                            if (e.target.value) {
                              const [yyyy, mm, dd] = e.target.value.split('-');
                              setSonSayimTarih(`${dd}.${mm}.${yyyy}`);
                            }
                          }}
                          className="p-1.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg text-xs cursor-pointer text-white shrink-0"
                          title="Takvimden Tarih Seç"
                        />
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="space-y-2.5">
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">İşlem Türü</label>
                      <input
                        type="text"
                        value={islemTransfer}
                        onChange={e => setIslemTransfer(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">İşlem Tarihi</label>
                      <div className="flex items-center gap-1.5">
                        <input
                          type="text"
                          value={tarih}
                          onChange={e => setTarih(e.target.value)}
                          className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono text-xs focus:border-emerald-500 focus:outline-none"
                        />
                        <input
                          type="date"
                          onChange={(e) => {
                            if (e.target.value) {
                              const [yyyy, mm, dd] = e.target.value.split('-');
                              const now = new Date();
                              const pad = (n: number) => String(n).padStart(2, '0');
                              setTarih(`${dd}.${mm}.${yyyy} ${pad(now.getHours())}:${pad(now.getMinutes())}`);
                            }
                          }}
                          className="p-1.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg text-xs cursor-pointer text-white shrink-0"
                          title="Takvimden Tarih Seç"
                        />
                      </div>
                    </div>
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Malzeme Tanımı</label>
                    <input
                      type="text"
                      value={malzemeDesc}
                      onChange={e => setMalzemeDesc(e.target.value)}
                      className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white focus:border-emerald-500 focus:outline-none"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">P/N</label>
                      <input
                        type="text"
                        value={pn}
                        onChange={e => setPn(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">S/N</label>
                      <input
                        type="text"
                        value={sn}
                        onChange={e => setSn(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Miktar / Adet</label>
                      <input
                        type="text"
                        value={adet}
                        onChange={e => setAdet(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-emerald-400 font-bold focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-bold text-slate-400 uppercase font-mono mb-1">Depo Lokasyonu</label>
                      <input
                        type="text"
                        value={depoYeri}
                        onChange={e => setDepoYeri(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white focus:border-emerald-500 focus:outline-none"
                      />
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* RIGHT PREVIEW COLUMN */}
            <div className="lg:col-span-7 flex flex-col items-center justify-center space-y-4">
              
              {/* TOP ACTIONS */}
              <div className="w-full flex items-center justify-between gap-2 flex-wrap bg-slate-950 p-2.5 rounded-xl border border-slate-800">
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setOrientation('yatay')}
                    className={`px-2.5 py-1 rounded-lg text-xs font-bold transition cursor-pointer ${
                      orientation === 'yatay' ? 'bg-blue-600 text-white' : 'bg-slate-800 text-slate-400 hover:text-white'
                    }`}
                  >
                    Yatay (Kompakt 76mm)
                  </button>
                  <button
                    type="button"
                    onClick={() => setOrientation('dikey')}
                    className={`px-2.5 py-1 rounded-lg text-xs font-bold transition cursor-pointer ${
                      orientation === 'dikey' ? 'bg-blue-600 text-white' : 'bg-slate-800 text-slate-400 hover:text-white'
                    }`}
                  >
                    Dikey Standart
                  </button>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={handleCopyText}
                    className="flex items-center gap-1 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-bold transition cursor-pointer"
                    title="Metni Kopyala"
                  >
                    {copied ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                    <span>Kopyala</span>
                  </button>

                  <button
                    type="button"
                    onClick={handlePrint}
                    className="flex items-center gap-1 px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-black transition cursor-pointer shadow-xs active:scale-95"
                  >
                    <Printer className="w-3.5 h-3.5 text-emerald-100" />
                    <span>YAZDIR (TERMAL)</span>
                  </button>

                  <button
                    type="button"
                    onClick={handlePrintA4HTML}
                    className="flex items-center gap-1 px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-black transition cursor-pointer shadow-xs active:scale-95"
                  >
                    <Printer className="w-3.5 h-3.5 text-indigo-100" />
                    <span>A4 YAZDIR (HTML)</span>
                  </button>

                  <button
                    type="button"
                    onClick={handleDownloadWord}
                    disabled={bulkExportProgress !== null}
                    className="flex items-center gap-1 px-3 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-black transition cursor-pointer shadow-xs active:scale-95"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" />
                    <span>{bulkExportProgress !== null ? `Aktarılıyor %${bulkExportProgress}` : 'WORD\'E AKTAR'}</span>
                  </button>
                </div>
              </div>

              {/* PRINTABLE SLIP (Shortened Horizontal Width max-w-[270px]) */}
              <div 
                id="ogm-printable-slip"
                className={`bg-white text-slate-900 rounded-2xl shadow-2xl border-2 border-slate-300 font-mono leading-tight select-text transition-all ${
                  orientation === 'yatay' 
                    ? 'w-full max-w-[270px] p-3 text-[9px]' 
                    : 'w-full max-w-[290px] p-3.5 text-[9.5px]'
                }`}
              >
                {selectedSlipType === 'etiket' ? (
                  /* TEÇHİZAT & ÖZEL ALET ETİKETİ (ÇERÇEVELİ GÖRSEL VE KAREKODLU, BARKODSUZ) */
                  <div className="space-y-1.5">
                    <div className="flex items-center justify-between border-b-2 border-slate-900 pb-0.5">
                      <span className="text-[10px] font-black uppercase tracking-tight text-slate-950 font-sans">
                        OGM HAVACILIK
                      </span>
                      <span className="text-[8px] font-extrabold text-amber-900 bg-amber-100 border border-amber-300 px-1.5 py-0.5 rounded font-mono">
                        TEÇHİZAT ETİKETİ
                      </span>
                    </div>

                    <div className="grid grid-cols-12 gap-1.5 items-center pt-1">
                      {/* Left: Text Info */}
                      <div className="col-span-8 space-y-0.5 text-[8.5px]">
                        <div className="truncate"><strong className="text-slate-500">TNM:</strong> <span className="font-bold text-slate-950">{tanimIcerik || malzemeDesc}</span></div>
                        <div className="truncate">
                          <strong className="text-slate-500">P/N:</strong>{' '}
                          {pn && pn.trim() !== '' && pn.trim() !== '-' ? (
                            <span className="font-black font-mono text-slate-900">{pn}</span>
                          ) : (
                            <span className="text-slate-400 italic font-mono">-</span>
                          )}
                        </div>
                        {sn && sn.trim() !== '' && sn.trim() !== '-' && (
                          <div className="truncate">
                            <strong className="text-slate-500">S/N:</strong>{' '}
                            <span className="font-mono font-bold text-amber-950">{sn}</span>
                          </div>
                        )}
                        <div className="truncate"><strong className="text-slate-500">LOK:</strong> <span className="font-black text-slate-900">{lokasyonKodu || depoYeri}</span></div>
                        {kalibrasyonTarihi && kalibrasyonTarihi.trim() !== '' && kalibrasyonTarihi.trim() !== '-' && (
                          <div className="truncate"><strong className="text-slate-500">KAL/BAK:</strong> <span className="font-bold text-emerald-900 font-mono">{kalibrasyonTarihi}</span></div>
                        )}
                        <div className="text-[7px] text-slate-400 font-sans">Tarih: {sonSayimTarih || tarih}</div>
                      </div>

                      {/* Right: Square Framed Photo Box */}
                      <div className="col-span-4 flex items-center justify-center">
                        <div className="w-16 h-16 border-2 border-slate-300 rounded-xl overflow-hidden bg-slate-50 flex items-center justify-center p-0.5 shadow-xs">
                          {imageUrlInput ? (
                            <img src={getDirectDriveImageUrl(imageUrlInput)} alt="Teçhizat Görseli" className="w-full h-full object-cover rounded-lg" />
                          ) : (
                            <Wrench className="w-8 h-8 text-slate-400" />
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Bottom: Centered QR Code (No 1D Barcode) */}
                    <div className="border-t border-slate-200 pt-1 flex flex-col items-center justify-center">
                      {qrDataUrl && (
                        <img src={qrDataUrl} alt="QR Code" className="w-14 h-14 border border-slate-300 p-0.5 bg-white rounded-lg" />
                      )}
                      <div className="text-[6.5px] font-bold tracking-widest text-slate-600 font-mono mt-0.5">
                        [HIZLI SORGULAMA KAREKODU]
                      </div>
                    </div>
                  </div>
                ) : orientation === 'yatay' ? (
                  /* KISALTILMIŞ YATAY KOMPAKT TERMAL SLİP */
                  <div className="space-y-1.5">
                    <div className="flex items-center justify-between border-b-2 border-slate-900 pb-0.5">
                      <span className="text-[10px] font-black uppercase tracking-tight text-slate-950 font-sans">
                        OGM HAVACILIK
                      </span>
                      <span className="text-[8px] font-bold text-slate-700 font-mono">
                        {selectedSlipType === 'raf' ? 'RAF SLİBİ' : 'GİRİŞ SLİBİ'}
                      </span>
                    </div>

                    <div className="grid grid-cols-12 gap-1.5 items-start">
                      {/* Left: Data */}
                      <div className="col-span-8 space-y-0.5 text-[8.5px]">
                        {selectedSlipType === 'raf' ? (
                          <>
                            <div className="truncate"><strong className="text-slate-500">LOK:</strong> <span className="font-black text-slate-950">{lokasyonKodu}</span></div>
                            <div className="truncate"><strong className="text-slate-500">BÖL:</strong> <span>{bolgeZone}</span></div>
                            <div className="truncate"><strong className="text-slate-500">TNM:</strong> <span className="font-bold text-slate-900">{tanimIcerik}</span></div>
                            <div><strong className="text-slate-500">STK:</strong> <span className="font-black text-emerald-800">{guncelSayim} EA</span></div>
                            <div className="text-[7px] text-slate-400 font-sans">Tarih: {sonSayimTarih}</div>
                          </>
                        ) : (
                          <>
                            <div className="truncate"><strong className="text-slate-500">İŞL:</strong> <span className="font-black text-slate-950">{islemTransfer}</span></div>
                            <div className="truncate"><strong className="text-slate-500">MALZ:</strong> <span className="font-bold text-slate-900">{malzemeDesc}</span></div>
                            <div className="truncate"><strong className="text-slate-500">P/N:</strong> <span className="font-black font-mono">{pn}</span></div>
                            <div className="truncate"><strong className="text-slate-500">S/N:</strong> <span className="font-mono text-slate-800">{sn}</span></div>
                            <div><strong className="text-slate-500">MİKT:</strong> <span className="font-black text-emerald-800">{adet} EA</span></div>
                            <div className="text-[7px] text-slate-400 font-sans">Giriş: {tarih}</div>
                          </>
                        )}
                      </div>

                      {/* Right: QR */}
                      <div className="col-span-4 flex flex-col items-center justify-center border-l border-slate-200 pl-1">
                        {qrDataUrl && (
                          <img src={qrDataUrl} alt="QR Code" className="w-12 h-12 border border-slate-200 p-0.5 bg-white mb-0.5" />
                        )}
                        <div className="text-[6.5px] font-bold text-center tracking-tighter text-slate-600 font-mono">
                          [HIZLI QR]
                        </div>
                      </div>
                    </div>

                    {/* Bottom: Barcode */}
                    <div className="border-t border-slate-200 pt-1 flex flex-col items-center">
                      {barcodeDataUrl ? (
                        <img src={barcodeDataUrl} alt="Barcode" className="max-w-full h-6 object-contain" />
                      ) : (
                        <svg ref={barcodeSvgRef} className="max-w-full h-6" />
                      )}
                      <div className="text-[7.5px] font-bold tracking-widest text-slate-800 font-mono mt-0.5">
                        *{barcodeValue}*
                      </div>
                    </div>
                  </div>
                ) : (
                  /* DİKEY SLİP */
                  <div className="space-y-1.5">
                    <div className="text-center py-0.5 border-b border-slate-900 pb-1">
                      <div className="text-[11px] font-black tracking-wider text-slate-950 uppercase font-sans">
                        OGM HAVACILIK
                      </div>
                      <div className="text-[8.5px] font-bold text-slate-700 uppercase">
                        {selectedSlipType === 'raf' ? 'RAF LOKASYON VE SAYIM SLİBİ' : 'MALZEME GİRİŞ SLİBİ'}
                      </div>
                    </div>

                    {selectedSlipType === 'raf' ? (
                      <div className="space-y-0.5 text-[9px] py-1">
                        <div className="flex justify-between">
                          <span className="font-bold text-slate-600">LOKASYON:</span>
                          <span className="font-black text-slate-950">{lokasyonKodu}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="font-bold text-slate-600">BÖLGE:</span>
                          <span className="text-slate-800 text-right truncate max-w-[150px]">{bolgeZone}</span>
                        </div>
                        <div>
                          <span className="font-bold text-slate-600">İÇERİK:</span>{' '}
                          <span className="font-bold text-slate-900">{tanimIcerik}</span>
                        </div>
                        <div className="flex justify-between items-center pt-0.5">
                          <span className="font-bold text-slate-600">GÜNCEL SAYIM:</span>
                          <span className="font-black text-emerald-800 text-[10px]">{guncelSayim} EA</span>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-0.5 text-[9px] py-1">
                        <div className="flex justify-between">
                          <span className="font-bold text-slate-600">İŞLEM:</span>
                          <span className="font-black text-slate-950">{islemTransfer}</span>
                        </div>
                        <div>
                          <span className="font-bold text-slate-600">MALZEME:</span>{' '}
                          <span className="font-bold text-slate-900 truncate">{malzemeDesc}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="font-bold text-slate-600">P/N:</span>
                          <span className="font-black font-mono">{pn}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="font-bold text-slate-600">S/N:</span>
                          <span className="font-mono text-slate-800">{sn}</span>
                        </div>
                        <div className="flex justify-between items-center">
                          <span className="font-bold text-slate-600">MİKTAR:</span>
                          <span className="font-black text-emerald-800 text-[10px]">{adet} EA</span>
                        </div>
                        <div className="flex justify-between text-[8px]">
                          <span className="font-bold text-slate-600">DEPO:</span>
                          <span className="text-slate-800">{depoYeri}</span>
                        </div>
                      </div>
                    )}

                    {/* 1D Barcode */}
                    <div className="border-t border-dashed border-slate-300 pt-1 text-center flex flex-col items-center">
                      {barcodeDataUrl ? (
                        <img src={barcodeDataUrl} alt="Barcode" className="max-w-full h-7 object-contain" />
                      ) : (
                        <svg ref={barcodeSvgRef} className="max-w-full h-7" />
                      )}
                      <div className="text-[7.5px] font-bold tracking-widest text-slate-800 font-mono mt-0.5">
                        *{barcodeValue}*
                      </div>
                    </div>

                    {/* QR Code */}
                    <div className="border-t border-dashed border-slate-300 pt-1 text-center flex flex-col items-center">
                      {qrDataUrl && (
                        <img src={qrDataUrl} alt="QR Code" className="w-14 h-14 border border-slate-200 p-0.5 bg-white" />
                      )}
                      <div className="text-[7px] font-bold text-slate-600 mt-0.5">
                        [HIZLI SORGULAMA KAREKODU]
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>

          </div>

        </div>

        {/* MODAL FOOTER */}
        <div className="bg-slate-950 px-5 py-3 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400">
          <div>
            <span className="text-emerald-400 font-bold">✓ Termal Yazıcı, Bluetooth, USB ve Word .doc tam uyumlu.</span>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="px-5 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold transition cursor-pointer"
          >
            Kapat
          </button>
        </div>

      </div>
    </div>
  );
};
