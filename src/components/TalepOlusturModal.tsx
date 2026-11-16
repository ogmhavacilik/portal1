import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  X,
  FileText,
  Download,
  Printer,
  CheckCircle2,
  ArrowLeft,
  Plus,
  Trash2,
  Search,
  Send,
  ClipboardList,
  Wrench,
  FlaskConical,
  Package,
  Gauge,
  Hammer,
  Truck,
  Check,
  Image as ImageIcon,
  Archive,
  Clock,
  ZoomIn,
  ZoomOut,
  Maximize2,
  Minimize2,
  ChevronLeft,
  ChevronRight,
  RefreshCw,
  GitBranch,
  Upload,
  ExternalLink,
  FileCheck,
  AlertTriangle,
  Copy,
  Edit3
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { useCircularOgmLogo, OGM_WIKIMEDIA_LOGO_URL } from '../utils/ogmLogo';
import { fetchExcelFromDrive as remoteFetchExcelFromDrive, DRIVE_FOLDER_ID } from '../utils/remoteDriveSync';
import { unmergeAndFillWorksheet, isHeaderLikeRow } from '../utils/driveExcelSync';
import { PersonnelAutocomplete } from './PersonnelAutocomplete';
import { AIRCRAFT_TAILS } from './DepoManagementModal';
import { getStoredPersonnelList, syncPersonnelListFromScript } from '../utils/personnelData';
import {
  calculateDaysDiff,
  getRowSection,
  isHeaderLikeRow as isGunTakipHeaderRow
} from './GunTakipModal';

export type RequestFormType =
  | 'parca_onarim'
  | 'bakim_sarf'
  | 'parca_talep'
  | 'ozel_alet_kalibrasyon'
  | 'ozel_alet_alim_imal'
  | 'hizmet_talep'
  | 'arac_techizat_hizmet';

interface RequestTypeOption {
  id: RequestFormType;
  num: number;
  label: string;
  formTitle: string;
  desc: string;
  icon: React.ReactNode;
  badgeColor: string;
}

const REQUEST_TYPES: RequestTypeOption[] = [
  {
    id: 'parca_onarim',
    num: 1,
    label: 'HAVA ARACI PARÇA ONARIM TALEP',
    formTitle: 'HAVA ARACI PARÇA ONARIM TALEP FORMU',
    desc: 'Depodan parça ve S/N seçimi ile 2 sayfalık arızalı parça onarım formu',
    icon: <Wrench className="w-5 h-5" />,
    badgeColor: 'bg-amber-600 text-white'
  },
  {
    id: 'bakim_sarf',
    num: 2,
    label: 'HAVA ARACI BAKIM İÇİN GEREKLİ SARF MALZEME TALEP',
    formTitle: 'HAVA ARACI MADENİ YAĞ, GAZ VE KİMYASAL MALZEME TALEP FORMU',
    desc: 'Kimyasal & madeni yağ deposundan otomatik aktarımlı talep formu',
    icon: <FlaskConical className="w-5 h-5" />,
    badgeColor: 'bg-indigo-600 text-white'
  },
  {
    id: 'parca_talep',
    num: 3,
    label: 'HAVA ARACI PARÇA TALEP',
    formTitle: 'HAVA ARACI YEDEK PARÇA TALEP FORMU',
    desc: '2 sayfalık resmi Zero Report yedek parça & katalog görseli talep formu',
    icon: <Package className="w-5 h-5" />,
    badgeColor: 'bg-emerald-700 text-white'
  },
  {
    id: 'ozel_alet_kalibrasyon',
    num: 4,
    label: 'ÖZEL ALET KALİBRASYON TALEP',
    formTitle: 'ÖZEL ALET KALİBRASYON TALEP FORMU',
    desc: 'Birim özel alet envanterinden P/N, S/N ve kalibrasyon tarihli çoklu seçim',
    icon: <Gauge className="w-5 h-5" />,
    badgeColor: 'bg-cyan-700 text-white'
  },
  {
    id: 'ozel_alet_alim_imal',
    num: 5,
    label: 'ÖZEL ALET ALIM İMAL ÜRETİM TALEP',
    formTitle: 'ÖZEL ALET VE YER DESTEK MALZEMELERİ ALIM/ İMAL/ ÜRETİM TALEP FORMU',
    desc: '3 sayfalık kapsamlı teknik şartname, malzeme türü ve çizim/katalog formu',
    icon: <Hammer className="w-5 h-5" />,
    badgeColor: 'bg-purple-700 text-white'
  },
  {
    id: 'hizmet_talep',
    num: 6,
    label: 'HİZMET TALEP',
    formTitle: 'HİZMET TALEP FORMU',
    desc: 'Hava aracı / parça bakım, revizyon, kaplama, boyama, NDT hizmet talebi',
    icon: <ClipboardList className="w-5 h-5" />,
    badgeColor: 'bg-rose-700 text-white'
  },
  {
    id: 'arac_techizat_hizmet',
    num: 7,
    label: 'ARAÇ TEÇHİZAT HİZMET TALEP',
    formTitle: 'HİZMET TALEP FORMU',
    desc: 'Kara araçları ve yer destek teçhizatı bilgilerini otomatik çeken hizmet formu',
    icon: <Truck className="w-5 h-5" />,
    badgeColor: 'bg-slate-700 text-white'
  }
];

const AIRCRAFT_UNITS = [
  { key: 'at802', label: 'AT-802' },
  { key: 't70', label: 'T-70' },
  { key: 'bell429', label: 'Bell-429' },
  { key: 'b360', label: 'B-360' },
  { key: 'c650', label: 'C-650' },
  { key: 'hangar', label: 'Hangar / Genel' }
];

interface TalepOlusturModalProps {
  isOpen: boolean;
  onClose: () => void;
  showNotification: (msg: string, type?: 'success' | 'error' | 'info') => void;
  unitDataMap: {
    at802?: string[][];
    at802_ozel_alet?: string[][];
    bell429?: string[][];
    t70?: string[][];
    t70_bumbi_backet?: string[][];
    t70_helitak?: string[][];
    b360?: string[][];
    c650?: string[][];
    hangar?: string[][];
    kara_araclari?: string[][];
  };
}

interface UnifiedDepotItem {
  unit: string;
  unitLabel: string;
  category: 'sarf' | 'kimyasal';
  name: string;
  pn: string;
  sn: string;
  unitMeasure?: string;
}

interface UnifiedToolItem {
  unit: string;
  unitLabel: string;
  name: string;
  pn: string;
  sn: string;
  location: string;
  kalibreBitis: string;
  aciklama: string;
}

interface UnifiedVehicleItem {
  name: string;
  plate: string;
  brandModel: string;
  sn: string;
  typeDesc: string;
  location: string;
}

interface OzelAletAlertItem {
  id: string;
  unitKey: 'at802' | 't70' | 'bell429' | 'b360';
  unitLabel: string;
  techName: string;
  partNo: string;
  seriNo: string;
  location: string;
  sonTarih: string;
  gelecekTarih: string;
  daysDiff: number;
  aciklama: string;
  isNew: boolean;
}

const OGM_WET_SIGNED_UPLOAD_URL = 'https://bulut.ogm.gov.tr/SARTNAMELER';
const OGM_WET_SIGNED_VIEW_URL =
  'https://bulut.ogm.gov.tr/app/tr-TR/Klasor/Paylas/OGMDrive/4324a6cd-7f8a-4c6b-86bc-92e0990d26c3';

interface SubmittedTalepRecord {
  id: string;
  talepNo: string;
  formType: RequestFormType;
  formTypeLabel: string;
  formTitle: string;
  unit: string;
  unitLabel: string;
  tailNo: string;
  summary: string;
  requestedBy: string;
  date: string;
  timestamp: string;
  status: string;
  wetSignedDocUploaded?: boolean;
  wetSignedClickedAt?: string;
  wetSignedUploadedAt?: string;
  wetSignedDocUrl?: string;
  formData?: any;
}

export interface FormImageItem {
  id: string;
  dataUrl: string;
  scale: number; // 25..100 (% width)
  title?: string;
}

const createImageId = () => `img_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;

const normalizeImageList = (
  list?: FormImageItem[],
  legacy1?: string,
  legacy2?: string
): FormImageItem[] => {
  if (Array.isArray(list) && list.length > 0) {
    return list.map(it => ({
      id: it.id || createImageId(),
      dataUrl: it.dataUrl,
      scale: typeof it.scale === 'number' && it.scale >= 25 && it.scale <= 100 ? it.scale : 50,
      title: it.title
    }));
  }
  const out: FormImageItem[] = [];
  if (legacy1) out.push({ id: createImageId(), dataUrl: legacy1, scale: 50, title: 'Katalog / Görsel 1' });
  if (legacy2) out.push({ id: createImageId(), dataUrl: legacy2, scale: 50, title: 'Durum / Görsel 2' });
  return out;
};

const optimizeImageFileToDataUrl = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Dosya okunamadı'));
    reader.onload = ev => {
      const rawUrl = String(ev.target?.result || '');
      if (!rawUrl) {
        reject(new Error('Boş görsel verisi'));
        return;
      }
      const img = new Image();
      img.onload = () => {
        const maxDim = 1600;
        let { width, height } = img;
        if (width <= maxDim && height <= maxDim && file.size < 650 * 1024) {
          resolve(rawUrl);
          return;
        }
        if (width > maxDim || height > maxDim) {
          const ratio = Math.min(maxDim / width, maxDim / height);
          width = Math.max(1, Math.round(width * ratio));
          height = Math.max(1, Math.round(height * ratio));
        }
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(rawUrl);
          return;
        }
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);
        const outUrl = canvas.toDataURL('image/jpeg', 0.86);
        resolve(outUrl);
      };
      img.onerror = () => resolve(rawUrl);
      img.src = rawUrl;
    };
    reader.readAsDataURL(file);
  });
};

const SCALE_PRESETS = [
  { label: 'Küçük (%35)', value: 35 },
  { label: 'Orta (%50)', value: 50 },
  { label: 'Büyük (%75)', value: 75 },
  { label: 'Tam (%100)', value: 100 }
];

interface FormImageGalleryProps {
  images: FormImageItem[];
  onChange: (next: FormImageItem[]) => void;
  defaultScale: number;
  onDefaultScaleChange: (scale: number) => void;
  primaryLabel: string;
  secondaryLabel?: string;
  compact?: boolean;
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void;
}

const FormImageGallery: React.FC<FormImageGalleryProps> = ({
  images,
  onChange,
  defaultScale,
  onDefaultScaleChange,
  primaryLabel,
  secondaryLabel,
  compact = false,
  onNotify
}) => {
  const [isExpandedArea, setIsExpandedArea] = useState(false);
  const imagesRef = useRef<FormImageItem[]>(images);
  imagesRef.current = images;

  const handleFiles = async (files: File[] | FileList | null, titlePrefix?: string) => {
    const fileArray = Array.isArray(files) ? files : files ? Array.from(files) : [];
    if (fileArray.length === 0) return;

    const imageFiles = fileArray.filter(
      f =>
        f.type.startsWith('image/') ||
        /\.(png|jpe?g|webp|gif|bmp|svg|heic|jfif)$/i.test(f.name || '')
    );
    if (imageFiles.length === 0) return;

    const added: FormImageItem[] = [];
    for (const f of imageFiles) {
      try {
        const dataUrl = await optimizeImageFileToDataUrl(f);
        added.push({
          id: createImageId(),
          dataUrl,
          scale: defaultScale,
          title: titlePrefix || f.name
        });
      } catch {}
    }
    if (added.length > 0) {
      const merged = [...imagesRef.current, ...added];
      imagesRef.current = merged;
      onChange(merged);
      onNotify?.(`🖼️ ${added.length} görsel optimize edilip form ekine eklendi.`, 'success');
    }
  };

  const moveImage = (idx: number, dir: -1 | 1) => {
    const targetIdx = idx + dir;
    if (targetIdx < 0 || targetIdx >= images.length) return;
    const next = [...images];
    const tmp = next[idx];
    next[idx] = next[targetIdx];
    next[targetIdx] = tmp;
    onChange(next);
  };

  const adjustScale = (idx: number, delta: number) => {
    const next = images.map((it, i) =>
      i === idx ? { ...it, scale: Math.min(100, Math.max(25, (it.scale || 50) + delta)) } : it
    );
    onChange(next);
  };

  const toggleFullScale = (idx: number) => {
    const next = images.map((it, i) =>
      i === idx ? { ...it, scale: it.scale >= 95 ? 50 : 100 } : it
    );
    onChange(next);
  };

  const removeImage = (idx: number) => {
    onChange(images.filter((_, i) => i !== idx));
  };

  return (
    <div
      className="mt-2 space-y-2"
      onPaste={async e => {
        const items = e.clipboardData?.items;
        if (!items) return;
        const pasted: File[] = [];
        for (let i = 0; i < items.length; i++) {
          if (items[i].type.startsWith('image/')) {
            const f = items[i].getAsFile();
            if (f) pasted.push(f);
          }
        }
        if (pasted.length > 0) {
          e.stopPropagation();
          await handleFiles(pasted, 'Yapıştırılan Görsel');
        }
      }}
    >
      <div className="flex flex-wrap items-center justify-between gap-1.5 bg-slate-50 border border-slate-200 rounded-lg px-2 py-1.5">
        <div className="flex flex-wrap items-center gap-1.5">
          <label className="px-2.5 py-1 bg-white hover:bg-emerald-50 text-slate-800 hover:text-emerald-800 border border-slate-300 hover:border-emerald-400 rounded text-[10.5px] font-bold cursor-pointer inline-flex items-center gap-1 shadow-2xs">
            <ImageIcon className="w-3.5 h-3.5 text-emerald-600" />
            <span>{primaryLabel}</span>
            <input
              type="file"
              accept="image/*"
              multiple
              className="hidden"
              onChange={e => {
                const picked = e.target.files ? Array.from(e.target.files) : [];
                e.target.value = '';
                handleFiles(picked, primaryLabel);
              }}
            />
          </label>
          {secondaryLabel && (
            <label className="px-2.5 py-1 bg-white hover:bg-blue-50 text-slate-800 hover:text-blue-800 border border-slate-300 hover:border-blue-400 rounded text-[10.5px] font-bold cursor-pointer inline-flex items-center gap-1 shadow-2xs">
              <ImageIcon className="w-3.5 h-3.5 text-blue-600" />
              <span>{secondaryLabel}</span>
              <input
                type="file"
                accept="image/*"
                multiple
                className="hidden"
                onChange={e => {
                  const picked = e.target.files ? Array.from(e.target.files) : [];
                  e.target.value = '';
                  handleFiles(picked, secondaryLabel);
                }}
              />
            </label>
          )}
        </div>

        <div className="flex items-center gap-1.5">
          <span className="text-[10px] text-slate-500 font-semibold">Başlangıç:</span>
          <select
            value={defaultScale}
            onChange={e => onDefaultScaleChange(Number(e.target.value))}
            className="px-1.5 py-0.5 bg-white border border-slate-300 rounded text-[10px] font-bold text-slate-700 outline-none cursor-pointer"
          >
            {SCALE_PRESETS.map(p => (
              <option key={p.value} value={p.value}>
                {p.label}
              </option>
            ))}
          </select>
          {images.length > 0 && (
            <button
              type="button"
              onClick={() => setIsExpandedArea(v => !v)}
              title={isExpandedArea ? 'Fotoğraf alanını standart boyuta getir' : 'Fotoğraf alanını genişlet'}
              className="px-2 py-0.5 bg-white hover:bg-slate-200 border border-slate-300 rounded text-[10px] font-bold text-slate-700 cursor-pointer inline-flex items-center gap-1"
            >
              {isExpandedArea ? <Minimize2 className="w-3 h-3" /> : <Maximize2 className="w-3 h-3" />}
              <span>{isExpandedArea ? 'Daralt' : 'Alanı Büyüt'}</span>
            </button>
          )}
        </div>
      </div>

      {images.length > 0 && (
        <div
          className={`grid ${isExpandedArea ? 'grid-cols-1 gap-3 min-h-[280px]' : compact ? 'grid-cols-2 gap-2' : 'grid-cols-2 gap-2.5'} p-2 bg-slate-50/70 border border-slate-200 rounded-lg transition-all`}
        >
          {images.map((img, idx) => {
            const imgScale = img.scale || 50;
            const spanFull = isExpandedArea || imgScale >= 70;
            const maxPx = Math.round((isExpandedArea ? 200 : 120) + (imgScale - 25) * 4.2);
            return (
              <div
                key={img.id}
                className={`group relative bg-white border border-slate-300 rounded-lg p-1.5 flex flex-col items-center justify-center overflow-hidden transition-all ${
                  spanFull ? 'col-span-2' : 'col-span-1'
                }`}
              >
                <div className="w-full flex items-center justify-between text-[9.5px] font-mono text-slate-400 mb-1 px-0.5">
                  <span>
                    #{idx + 1} {img.title ? `• ${img.title}` : ''}
                  </span>
                  <span className="font-bold text-slate-600">%{imgScale}</span>
                </div>

                <img
                  src={img.dataUrl}
                  alt={img.title || `Görsel ${idx + 1}`}
                  style={{
                    width: `${imgScale}%`,
                    maxHeight: `${maxPx}px`
                  }}
                  className="object-contain mx-auto transition-all duration-150 rounded"
                />

                {/* Hover Overlay Toolbar: Reorder Left/Right, Scale +/-, Full/Medium, Delete */}
                <div className="absolute inset-x-1.5 bottom-1.5 opacity-0 group-hover:opacity-100 transition-opacity duration-150 flex items-center justify-center pointer-events-none group-hover:pointer-events-auto">
                  <div className="flex items-center gap-1 bg-slate-900/90 text-white px-2 py-1 rounded-lg shadow-lg border border-slate-700 text-[10px] font-bold">
                    <button
                      type="button"
                      disabled={idx === 0}
                      onClick={() => moveImage(idx, -1)}
                      title="Sola / Önceki Sıraya Kaydır"
                      className="p-0.5 hover:bg-slate-700 disabled:opacity-30 rounded cursor-pointer"
                    >
                      <ChevronLeft className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      disabled={idx === images.length - 1}
                      onClick={() => moveImage(idx, 1)}
                      title="Sağa / Sonraki Sıraya Kaydır"
                      className="p-0.5 hover:bg-slate-700 disabled:opacity-30 rounded cursor-pointer"
                    >
                      <ChevronRight className="w-3.5 h-3.5" />
                    </button>
                    <span className="w-px h-3.5 bg-slate-700 mx-0.5" />
                    <button
                      type="button"
                      disabled={imgScale <= 25}
                      onClick={() => adjustScale(idx, -15)}
                      title="Ölçeği Küçült"
                      className="p-0.5 hover:bg-slate-700 disabled:opacity-30 rounded cursor-pointer"
                    >
                      <ZoomOut className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      disabled={imgScale >= 100}
                      onClick={() => adjustScale(idx, 15)}
                      title="Ölçeği Büyüt"
                      className="px-1 py-0.5 hover:bg-emerald-600 disabled:opacity-30 rounded cursor-pointer inline-flex items-center gap-0.5"
                    >
                      <ZoomIn className="w-3.5 h-3.5" />
                      <span>Büyüt</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => toggleFullScale(idx)}
                      title={imgScale >= 95 ? 'Orta Boyuta Döndür' : 'Tam Genişlik Yap'}
                      className="p-0.5 hover:bg-blue-600 rounded cursor-pointer"
                    >
                      {imgScale >= 95 ? <Minimize2 className="w-3 h-3" /> : <Maximize2 className="w-3 h-3" />}
                    </button>
                    <span className="w-px h-3.5 bg-slate-700 mx-0.5" />
                    <button
                      type="button"
                      onClick={() => removeImage(idx)}
                      title="Görseli Kaldır"
                      className="px-1 py-0.5 bg-rose-600 hover:bg-rose-500 rounded cursor-pointer"
                    >
                      ✕
                    </button>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

const getTodayTr = () => {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`;
};

const buildDefaultFormState = () => ({
  // Common
  selectedUnit: '',
  havaAraciTipi: '',
  kuyrukTescilNo: '',
  selectedTypes: { at802: false, t70: false, bell429: false, b360: false },
  departmanDiger: '',
  talepYapanAdSoyad: '',
  date: getTodayTr(),
  ekNotlar: '',

  // Type 1 & 6 & 7: Arızalı Parça / Araç / Hizmet
  parcaAdi: '',
  parcaNumarasi: '',
  seriNumarasi: '',
  arizaKisaTanimi: '',
  arizaDetayliAciklama: '',
  arizaGorsel1: '',
  arizaGorsel2: '',
  arizaGorselleri: [] as FormImageItem[],
  islemler: {
    onarim: false,
    revizyon: false,
    kaplama: false,
    boyama: false,
    kaynak: false,
    ndt: false
  },
  islemDiger: '',
  arizaGecmisi: '' as 'Var' | 'Yok' | '',
  oncekiArizaTarihi: '',
  oncekiArizaNedeni: '',

  // Type 7 specific
  aracTechizatTipi: '',
  aracTescilNo: '',
  aracAdiTanimi: '',
  aracMarkaModel: '',

  // Type 2: Bakım Sarf / Kimyasal (18 rows default)
  kimyasalRows: Array.from({ length: 18 }, (_, i) => ({
    siraNo: i + 1,
    malzemeKodu: '',
    malzemeTanimi: '',
    birim: '',
    miktar: '',
    aciklama: ''
  })),

  // Type 3: Yedek Parça Talep / Zero Report (20 rows default)
  parcaTalepRows: Array.from({ length: 20 }, (_, i) => ({
    siraNo: i + 1,
    parcaAdi: '',
    parcaNo: '',
    alternatifPn: '',
    istenenMiktar: '',
    oncelik: '' as 'Yüksek' | 'Normal' | 'Düşük' | ''
  })),
  katalogGorseli: '',
  katalogGorselleri: [] as FormImageItem[],

  // Type 4: Özel Alet Kalibrasyon (28 rows default)
  kalibrasyonTuru: { ilk: false, periyodik: false, onarimSonrasi: false },
  kalibrasyonDiger: '',
  talepEdilenSonTarih: '',
  oncelikSeviyesi: '' as 'Yüksek (Acil)' | 'Normal' | 'Düşük' | '',
  kalibrasyonRows: Array.from({ length: 28 }, (_, i) => ({
    siraNo: i + 1,
    malzemeAdi: '',
    parcaNo: '',
    seriNo: '',
    tanimi: '',
    kalibreBitisTarihi: ''
  })),

  // Type 5: Özel Alet Alım/İmal/Üretim
  kullanimAmaci: '',
  gorevBakimlar: '',
  kullanimSikligi: '' as 'Sürekli' | 'Dönemsel' | 'Acil Durumlarda' | '',
  calismaOrtami: '' as 'Açık Alan' | 'Kapalı Alan' | 'Deniz / Su Alanı' | '',
  uzunluk: '',
  genislik: '',
  yukseklik: '',
  agirlik: '',
  isiDayaniklilik: '',
  darbeDayaniklilik: '' as 'Hafif' | 'Orta' | 'Yüksek' | '',
  suGecirmezlik: '' as 'Evet' | 'Hayır' | '',
  tozGecirmezlik: '' as 'Evet' | 'Hayır' | '',
  kimyasalDayaniklilik: '' as 'Evet' | 'Hayır' | '',
  uvKoruma: '' as 'Evet' | 'Hayır' | '',
  metaller: {
    celik: false,
    paslanmazCelik: false,
    aluminyum: false,
    bakir: false,
    titanyum: false,
    pirinc: false,
    dokmeDemir: false
  },
  metallerDiger: '',
  ametaller: {
    plastik: false,
    kaucuk: false,
    seramik: false,
    cam: false,
    karbonFiber: false,
    kompozit: false
  },
  ametallerDiger: '',
  kaplama: '',
  boyaRenkKodu: '',
  cilalama: '',
  yuzeyDiger: '',
  korozyonDayaniklilik: '' as 'Evet' | 'Hayır' | '',
  enerjiKaynagi: {
    manuel: false,
    pnomatik: false,
    hidrolik: false,
    elektrik: false
  },
  bataryaBilgisi: '',
  calismaVoltaji: '',
  gucTuketimi: '',
  standartlar: { ce: false, iso: false },
  standartDiger: '',
  cizimNotlari: '',
  cizimGorseli: '',
  cizimGorselleri: [] as FormImageItem[],
  ornekUrunLinkNotlari: '',
  ornekUrunGorseli: '',
  ornekUrunGorselleri: [] as FormImageItem[]
});

export const TalepOlusturModal: React.FC<TalepOlusturModalProps> = ({
  isOpen,
  onClose,
  showNotification,
  unitDataMap
}) => {
  const circularOgmLogoDataUri = useCircularOgmLogo();

  const [activeMainTab, setActiveMainTab] = useState<'create' | 'archive' | 'approved_archive'>('create');
  const [selectedType, setSelectedType] = useState<RequestFormType | null>(null);
  const [formState, setFormState] = useState(buildDefaultFormState);

  // Helper Autocomplete search states
  const [helperQuery, setHelperQuery] = useState('');
  const [isHelperDropdownOpen, setIsHelperDropdownOpen] = useState(false);

  // Live Depo inventory loaded on open
  const [depotItems, setDepotItems] = useState<UnifiedDepotItem[]>([]);
  const [isLoadingDepots, setIsLoadingDepots] = useState(false);

  // Personnel & inline form-cell autocomplete states
  const [personnelList, setPersonnelList] = useState<string[]>([]);
  const [activeCellAutocomplete, setActiveCellAutocomplete] = useState<string | null>(null);

  // Archive ("ONAY BEKLEYEN TALEPLER") & Taskline timeline state
  const [submittedList, setSubmittedList] = useState<SubmittedTalepRecord[]>([]);
  const [isLoadingArchive, setIsLoadingArchive] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccessModal, setSaveSuccessModal] = useState<{
    talepNo: string;
    formTitle: string;
    record?: SubmittedTalepRecord | null;
  } | null>(null);
  const [defaultImageScale, setDefaultImageScale] = useState<number>(50);
  const [pendingWetSignedWarning, setPendingWetSignedWarning] = useState<{
    record: SubmittedTalepRecord;
    clickedAt: string;
  } | null>(null);
  const [promptBelgeTalepNo, setPromptBelgeTalepNo] = useState<string | null>(null);
  const [isWetPdfConfirmedChecked, setIsWetPdfConfirmedChecked] = useState<boolean>(true);
  const [isUpdatingBelge, setIsUpdatingBelge] = useState<boolean>(false);
  const [tasklineFlowModal, setTasklineFlowModal] = useState<{
    talepNo: string;
    loading: boolean;
    error?: string;
    matched?: boolean;
    demand?: any;
    history: Array<{ date: string; description: string; isSupportNote?: boolean; isKalanIsler?: boolean }>;
  } | null>(null);

  // Live Gün Takip scan data & Özel Alet Kalibrasyon ikazlı ürün seçim modalı
  const [liveGunTakipData, setLiveGunTakipData] = useState<Record<string, any[][]>>({});
  const [liveGunTakipLoaded, setLiveGunTakipLoaded] = useState(false);
  const [kalibrasyonAlertModal, setKalibrasyonAlertModal] = useState<{
    unitKeys: Array<'at802' | 't70' | 'bell429' | 'b360'>;
    unitTitle: string;
    loading: boolean;
    items: OzelAletAlertItem[];
  } | null>(null);
  const [kalibrasyonAlertFilter, setKalibrasyonAlertFilter] = useState<'all' | 'overdue' | 'upcoming'>('all');
  const [kalibrasyonAlertSearch, setKalibrasyonAlertSearch] = useState('');
  const [kalibrasyonAlertSelectedIds, setKalibrasyonAlertSelectedIds] = useState<Record<string, boolean>>({});

  const normalizeStatusTr = (val?: string) =>
    String(val || '')
      .trim()
      .toLocaleUpperCase('tr-TR')
      .replace(/\s+/g, ' ');

  const isApprovedStatus = (val?: string) => normalizeStatusTr(val) === 'ONAYLANDI';

  const isCompletedStatus = (val?: string) => {
    const n = normalizeStatusTr(val);
    return n === 'İŞLEM TAMAMLANDI' || n === 'ISLEM TAMAMLANDI';
  };

  const pendingSubmittedList = useMemo(
    () => submittedList.filter(rec => !isApprovedStatus(rec.status) && !isCompletedStatus(rec.status)),
    [submittedList]
  );

  const approvedArchiveList = useMemo(
    () => submittedList.filter(rec => isApprovedStatus(rec.status)),
    [submittedList]
  );

  const formatUploadTimestampTr = (): string => {
    try {
      const parts = new Intl.DateTimeFormat('tr-TR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: 'Europe/Istanbul'
      }).formatToParts(new Date());
      const get = (t: string) => parts.find(p => p.type === t)?.value || '00';
      return `${get('day')}.${get('month')}.${get('year')} ${get('hour')}:${get('minute')}`;
    } catch {
      const d = new Date();
      const pad = (n: number) => String(n).padStart(2, '0');
      return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
  };

  const handleClickBelgeYukle = (rec: SubmittedTalepRecord) => {
    const clickedAt = formatUploadTimestampTr();
    setPendingWetSignedWarning({ record: rec, clickedAt });
  };

  const handleConfirmWarningAndOpenUpload = async () => {
    if (!pendingWetSignedWarning) return;
    const { record, clickedAt } = pendingWetSignedWarning;
    setPendingWetSignedWarning(null);
    window.open(OGM_WET_SIGNED_UPLOAD_URL, '_blank', 'noopener,noreferrer');
    setPromptBelgeTalepNo(record.talepNo);
    setIsWetPdfConfirmedChecked(true);

    // Tıklama zamanını (GÜN.AY.YIL SAAT:DK) hemen veri olarak kaydet
    setSubmittedList(prev =>
      prev.map(it =>
        String(it.talepNo || it.id) === String(record.talepNo || record.id)
          ? { ...it, wetSignedClickedAt: clickedAt }
          : it
      )
    );
    try {
      await fetch('/api/onay-bekleyen-talepler/belge', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          talepNo: record.talepNo,
          wetSignedClickedAt: clickedAt
        })
      });
    } catch {}
    showNotification(
      `📄 Şartnameler / Islak İmzalı Evrak yükleme sayfası yeni sekmede açıldı (${clickedAt}). PDF yükledikten sonra "TAMAM" butonuna basınız.`,
      'info'
    );
  };

  const handleConfirmWetSignedUpload = async (rec: SubmittedTalepRecord) => {
    const finalTimestamp = rec.wetSignedClickedAt || formatUploadTimestampTr();
    setIsUpdatingBelge(true);
    try {
      const res = await fetch('/api/onay-bekleyen-talepler/belge', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          talepNo: rec.talepNo,
          wetSignedDocUploaded: true,
          wetSignedClickedAt: finalTimestamp,
          wetSignedUploadedAt: finalTimestamp,
          wetSignedDocUrl: OGM_WET_SIGNED_VIEW_URL
        })
      });
      if (res.ok) {
        const json = await res.json();
        if (Array.isArray(json.items)) {
          setSubmittedList(json.items);
        }
      } else {
        setSubmittedList(prev =>
          prev.map(it =>
            String(it.talepNo || it.id) === String(rec.talepNo || rec.id)
              ? {
                  ...it,
                  wetSignedDocUploaded: true,
                  wetSignedClickedAt: finalTimestamp,
                  wetSignedUploadedAt: finalTimestamp,
                  wetSignedDocUrl: OGM_WET_SIGNED_VIEW_URL
                }
              : it
          )
        );
      }
      setPromptBelgeTalepNo(null);
      showNotification(`✅ Islak imzalı belge kaydedildi (${finalTimestamp}). Artık "BELGE GÖR" ile görüntüleyebilirsiniz.`, 'success');
    } catch {
      showNotification('Belge durumu güncellenirken bağlantı hatası oluştu.', 'error');
    } finally {
      setIsUpdatingBelge(false);
    }
  };

  const formatTasklineDateTr = (isoStr: string): string => {
    if (!isoStr) return '-';
    const ts = Date.parse(isoStr);
    if (isNaN(ts)) return isoStr;
    try {
      return new Date(ts)
        .toLocaleString('tr-TR', {
          day: '2-digit',
          month: 'long',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          timeZone: 'Europe/Istanbul'
        })
        .toLocaleUpperCase('tr-TR');
    } catch {
      return isoStr;
    }
  };

  const handleOpenTasklineFlow = async (talepNo: string, searchCode?: string) => {
    const code = (searchCode ?? talepNo).trim();
    if (!code) return;
    setTasklineFlowModal({
      talepNo: code,
      loading: true,
      history: []
    });
    try {
      const res = await fetch(`/api/taskline-flow?talepNo=${encodeURIComponent(code)}&_t=${Date.now()}`);
      const json = await res.json();
      if (res.ok && json.status === 'success') {
        setTasklineFlowModal({
          talepNo: code,
          loading: false,
          matched: Boolean(json.matched),
          demand: json.demand || null,
          history: Array.isArray(json.history) ? json.history : []
        });
      } else {
        setTasklineFlowModal({
          talepNo: code,
          loading: false,
          matched: false,
          error: json?.message || 'Taskline online bağlantısı kurulamadı.',
          history: []
        });
      }
    } catch (err: any) {
      setTasklineFlowModal({
        talepNo: code,
        loading: false,
        matched: false,
        error: err?.message || 'Taskline bağlantı hatası.',
        history: []
      });
    }
  };

  useEffect(() => {
    if (!isOpen) {
      setActiveCellAutocomplete(null);
      return;
    }
    setPersonnelList(getStoredPersonnelList());
    syncPersonnelListFromScript().then(list => {
      if (list && list.length > 0) setPersonnelList(list);
    });
  }, [isOpen]);

  // Fetch archive from backend
  const fetchSubmittedArchive = async () => {
    setIsLoadingArchive(true);
    try {
      const res = await fetch(`/api/onay-bekleyen-talepler?_t=${Date.now()}`);
      if (res.ok) {
        const json = await res.json();
        if (json && Array.isArray(json.items)) {
          setSubmittedList(json.items);
        }
      }
    } catch {} finally {
      setIsLoadingArchive(false);
    }
  };

  // Load live Depo (Sarf + Kimyasal) items when modal opens
  useEffect(() => {
    if (!isOpen) return;
    fetchSubmittedArchive();

    let cancelled = false;
    const loadDepots = async () => {
      setIsLoadingDepots(true);
      const targets = [
        { unit: 'at802', unitLabel: 'AT-802', cat: 'sarf' as const, file: 'depo_at802_sarf.xlsx' },
        { unit: 'at802', unitLabel: 'AT-802', cat: 'kimyasal' as const, file: 'depo_at802_kimyasal.xlsx' },
        { unit: 'bell429', unitLabel: 'Bell-429', cat: 'sarf' as const, file: 'depo_bell429_sarf.xlsx' },
        { unit: 'bell429', unitLabel: 'Bell-429', cat: 'kimyasal' as const, file: 'depo_bell429_kimyasal.xlsx' },
        { unit: 't70', unitLabel: 'T-70', cat: 'sarf' as const, file: 'depo_t70_sarf.xlsx' },
        { unit: 't70', unitLabel: 'T-70', cat: 'kimyasal' as const, file: 'depo_t70_kimyasal.xlsx' },
        { unit: 'b360', unitLabel: 'B-360', cat: 'sarf' as const, file: 'depo_b360_sarf.xlsx' },
        { unit: 'b360', unitLabel: 'B-360', cat: 'kimyasal' as const, file: 'depo_b360_kimyasal.xlsx' },
        { unit: 'c650', unitLabel: 'C-650', cat: 'sarf' as const, file: 'depo_c650_sarf.xlsx' },
        { unit: 'hangar', unitLabel: 'Hangar', cat: 'sarf' as const, file: 'depo_hangar_sarf.xlsx' }
      ];

      const parseExcelB64 = (b64: string, unit: string, unitLabel: string, cat: 'sarf' | 'kimyasal'): UnifiedDepotItem[] => {
        try {
          const bin = atob(b64);
          const bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          const wb = XLSX.read(bytes, { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          if (!ws) return [];
          unmergeAndFillWorksheet(ws);
          const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (!rawRows || rawRows.length <= 1) return [];

          let headerIdx = 0;
          for (let r = 0; r < Math.min(20, rawRows.length); r++) {
            const s = (rawRows[r] || []).map(c => String(c || '').toUpperCase()).join(' ');
            if (s.includes('MALZEME') || s.includes('DESCRIPTION') || s.includes('PARÇA') || s.includes('P/N')) {
              headerIdx = r;
              break;
            }
          }
          const hRow = (rawRows[headerIdx] || []).map(c =>
            String(c || '')
              .toUpperCase()
              .trim()
              .replace(/Ç/g, 'C')
              .replace(/Ğ/g, 'G')
              .replace(/İ/g, 'I')
              .replace(/Ö/g, 'O')
              .replace(/Ş/g, 'S')
              .replace(/Ü/g, 'U')
          );
          const descCol = hRow.findIndex(c => c === 'DESCRIPTION' || c.includes('DESCRIPTION') || c.includes('MALZEME') || c === 'TANIM');
          const pnCol = hRow.findIndex(c => c.includes('P/N') || c.includes('PART NUMBER') || c.includes('PARCA NO') || c.includes('PART'));
          const snCol = hRow.findIndex(c => c.includes('SERI') || c.includes('S/N'));
          const birimCol = hRow.findIndex(c => c.includes('BIRIM') || c.includes('UNIT'));

          const out: UnifiedDepotItem[] = [];
          for (let r = headerIdx + 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row || row.length === 0) continue;
            const name = String(row[descCol >= 0 ? descCol : 0] || '').trim();
            if (!name || isHeaderLikeRow(row)) continue;
            const pn = String(row[pnCol >= 0 ? pnCol : 1] ?? '-').trim();
            const snRaw = snCol >= 0 ? String(row[snCol] ?? '').trim() : '';
            const sn = snRaw && snRaw !== '-' && snRaw !== 'YOK' ? snRaw : '';
            const unitMeasure = birimCol >= 0 ? String(row[birimCol] ?? '').trim() : '';
            out.push({ unit, unitLabel, category: cat, name, pn, sn, unitMeasure });
          }
          return out;
        } catch {
          return [];
        }
      };

      const collected: UnifiedDepotItem[] = [];
      const results = await Promise.allSettled(
        targets.map(async (t) => {
          try {
            const res = await remoteFetchExcelFromDrive(t.file, [t.file], DRIVE_FOLDER_ID);
            if (res && res.status === 'success' && res.base64) {
              return parseExcelB64(res.base64, t.unit, t.unitLabel, t.cat);
            }
          } catch {}
          if (t.unit === 'at802' && t.cat === 'sarf') {
            try {
              const apiRes = await fetch('/api/at802-sarf-depo');
              if (apiRes.ok) {
                const json = await apiRes.json();
                if (json && Array.isArray(json.items)) {
                  return json.items.map((it: any) => ({
                    unit: 'at802',
                    unitLabel: 'AT-802',
                    category: 'sarf' as const,
                    name: String(it.description || it.name || '').trim(),
                    pn: String(it.partNumber || it.pn || '-').trim(),
                    sn: String(it.serialAndNotes || it.sn || '').trim()
                  }));
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
        setDepotItems(collected);
        setIsLoadingDepots(false);
      }
    };

    loadDepots();

    // Gün Takip canlı tarama verisini arka planda çek (Özel Alet Kalibrasyon ikaz penceresi için)
    fetch('/api/gun-takip-live-scan')
      .then(res => (res.ok ? res.json() : null))
      .then(json => {
        if (!cancelled && json && json.status === 'success' && json.data) {
          setLiveGunTakipData(json.data);
          setLiveGunTakipLoaded(true);
        } else if (!cancelled) {
          setLiveGunTakipLoaded(true);
        }
      })
      .catch(() => {
        if (!cancelled) setLiveGunTakipLoaded(true);
      });

    return () => {
      cancelled = true;
    };
  }, [isOpen]);

  // Unified Özel Alet & Teçhizat items from unitDataMap (liveGunTakipData fallback ile)
  const toolItems: UnifiedToolItem[] = useMemo(() => {
    const out: UnifiedToolItem[] = [];
    const pickRows = (k: keyof typeof unitDataMap) => {
      const propRows = (unitDataMap as any)[k];
      if (Array.isArray(propRows) && propRows.length > 0) return propRows;
      const liveRows = (liveGunTakipData as any)[k];
      return Array.isArray(liveRows) ? liveRows : [];
    };
    const mapSources: Array<{ key: string; label: string; rows?: string[][] }> = [
      { key: 'at802', label: 'AT-802', rows: pickRows('at802_ozel_alet') },
      { key: 'at802', label: 'AT-802 Yer Destek', rows: pickRows('at802') },
      { key: 'bell429', label: 'Bell-429', rows: pickRows('bell429') },
      { key: 't70', label: 'T-70', rows: pickRows('t70') },
      { key: 'b360', label: 'B-360', rows: pickRows('b360') },
      { key: 'c650', label: 'C-650', rows: pickRows('c650') },
      { key: 'hangar', label: 'Hangar', rows: pickRows('hangar') }
    ];
    mapSources.forEach(src => {
      (src.rows || []).forEach(r => {
        if (!Array.isArray(r) || r.length < 2) return;
        const name = String(r[1] || '').trim();
        if (!name || isGunTakipHeaderRow(r)) return;
        const isBell = src.key === 'bell429';
        out.push({
          unit: src.key,
          unitLabel: src.label,
          name,
          pn: String(r[2] || '').trim(),
          sn: String(r[3] || '').trim(),
          location: String(r[5] || '').trim(),
          kalibreBitis: String(isBell ? (r[8] || '') : (r[9] || '')).trim(),
          aciklama: String(isBell ? (r[10] || r[11] || '') : (r[11] || '')).trim()
        });
      });
    });
    return out;
  }, [unitDataMap, liveGunTakipData]);

  // Unified Kara Araçları items (11-column Kara Araçları layout: [0: sira, 1: plaka, 2: marka, 3: model, 4: yakit, 5: km, 6: durum, 7: bakimTarihi, 8: sonrakiBakim, 9: firma, 10: aciklama])
  const karaAracItems: UnifiedVehicleItem[] = useMemo(() => {
    const out: UnifiedVehicleItem[] = [];
    (unitDataMap.kara_araclari || []).forEach(r => {
      if (!Array.isArray(r) || r.length < 2) return;
      const plateRaw = String(r[1] || '').trim();
      const plate = plateRaw === '0' ? '-' : plateRaw;
      const marka = String(r[2] || '').trim();
      const model = String(r[3] || '').trim();
      const yakit = String(r[4] || '').trim();
      const km = String(r[5] || '').trim();
      const aciklama = String(r[10] || r[11] || '').trim();
      const displayName = [marka !== '-' ? marka : '', model !== '-' ? model : ''].filter(Boolean).join(' ') || plate;
      if (!displayName && !plate) return;
      out.push({
        name: displayName,
        plate,
        brandModel: [marka !== '-' ? marka : '', model !== '-' ? model : '', yakit !== '-' ? yakit : ''].filter(Boolean).join(' / '),
        sn: km && km !== '-' ? `${km} KM` : '-',
        typeDesc: aciklama || displayName,
        location: yakit || '-'
      });
    });
    return out;
  }, [unitDataMap.kara_araclari]);

  // GunTakipModal.tsx ile birebir aynı Özel Alet 90 gün & günü geçmiş tarama mantığı
  const scanExpiringOzelAletForUnits = (
    unitKeys: Array<'at802' | 't70' | 'bell429' | 'b360'>,
    overrideData?: Record<string, any[][]>
  ): OzelAletAlertItem[] => {
    const unitLabelMap: Record<'at802' | 't70' | 'bell429' | 'b360', string> = {
      at802: 'AT-802',
      t70: 'T-70',
      bell429: 'Bell-429',
      b360: 'B-360'
    };
    const srcLive = overrideData || liveGunTakipData;
    const collected: OzelAletAlertItem[] = [];

    unitKeys.forEach(uKey => {
      let rawRows: any[][] = [];
      if (uKey === 'at802') {
        const liveOzel = (srcLive as any)?.at802_ozel_alet;
        const propOzel = unitDataMap.at802_ozel_alet;
        rawRows =
          Array.isArray(liveOzel) && liveOzel.length > 0
            ? liveOzel
            : Array.isArray(propOzel) && propOzel.length > 0
              ? propOzel
              : ((srcLive as any)?.at802 || unitDataMap.at802 || []);
      } else {
        const liveRows = (srcLive as any)?.[uKey];
        const propRows = (unitDataMap as any)?.[uKey];
        rawRows = Array.isArray(liveRows) && liveRows.length > 0 ? liveRows : (Array.isArray(propRows) ? propRows : []);
      }

      if (!Array.isArray(rawRows) || rawRows.length === 0) return;

      rawRows.forEach((row, rIdx) => {
        if (!Array.isArray(row) || isGunTakipHeaderRow(row)) return;
        const isAt802OzelSource =
          uKey === 'at802' &&
          ((Array.isArray((srcLive as any)?.at802_ozel_alet) && (srcLive as any).at802_ozel_alet.length > 0) ||
            (Array.isArray(unitDataMap.at802_ozel_alet) && unitDataMap.at802_ozel_alet.length > 0));
        const sec = isAt802OzelSource ? 'ozel_alet' : getRowSection(row);
        if (sec !== 'ozel_alet') return;

        let techName = '';
        let partNo = '';
        let seriNo = '';
        let location = '';
        let sonTarih = '';
        let gelecekTarih = '';
        let aciklama = '';
        let lastMailSentDate = '';
        let isMuaf = false;

        if (uKey === 'bell429') {
          techName = String(row[1] || 'Teçhizat').trim();
          partNo = String(row[2] || '-').trim();
          seriNo = String(row[3] || '-').trim();
          location = String(row[5] || '-').trim();
          sonTarih = String(row[7] || '').trim();
          gelecekTarih = String(row[8] || '').trim();
          aciklama = String(row[10] || row[11] || '').trim();
          lastMailSentDate = String(row[11] || '').trim();
        } else {
          techName = String(row[1] || 'Teçhizat').trim();
          partNo = String(row[2] || '-').trim();
          seriNo = String(row[3] || '-').trim();
          location = String(row[5] || '-').trim();
          const kalibTabi = String(row[7] || '').trim().toUpperCase();
          if (kalibTabi === 'HAYIR' || kalibTabi.includes('MUAF')) isMuaf = true;
          sonTarih = String(row[8] || '').trim();
          gelecekTarih = String(row[9] || '').trim();
          aciklama = String(row[11] || '').trim();
          lastMailSentDate = String(row[12] || row[11] || '').trim();
        }

        if (isMuaf || !gelecekTarih) return;

        const gelecekLines = String(gelecekTarih).split(/\r?\n/);
        const sonLines = String(sonTarih || '').split(/\r?\n/);
        const seriLines = String(seriNo || '').split(/\r?\n/);
        const locLines = String(location || '').split(/\r?\n/);
        const isMultiLine = gelecekLines.length > 1;
        const isNew = !lastMailSentDate || lastMailSentDate.trim() === '' || lastMailSentDate.trim() === '-';

        if (isMultiLine) {
          gelecekLines.forEach((rawLine, lIdx) => {
            const cleanLineDate = rawLine.trim();
            if (!cleanLineDate || /^x+$/i.test(cleanLineDate) || cleanLineDate === '-') return;
            const lineDaysDiff = calculateDaysDiff(cleanLineDate);
            if (lineDaysDiff === null || lineDaysDiff > 90) return;

            const alignedSeri = (seriLines[lIdx] ?? (seriLines.length === 1 ? seriLines[0] : '')).trim() || '-';
            const alignedLoc = (locLines[lIdx] ?? (locLines.length === 1 ? locLines[0] : '')).trim() || '-';
            const alignedSon = (sonLines[lIdx] ?? (sonLines.length === 1 ? sonLines[0] : '')).trim() || '-';

            collected.push({
              id: `${uKey}_${rIdx}_${lIdx}`,
              unitKey: uKey,
              unitLabel: unitLabelMap[uKey],
              techName: techName || 'Özel Alet',
              partNo: partNo || '-',
              seriNo: alignedSeri,
              location: alignedLoc,
              sonTarih: alignedSon,
              gelecekTarih: cleanLineDate,
              daysDiff: lineDaysDiff,
              aciklama,
              isNew
            });
          });
        } else {
          const daysDiff = calculateDaysDiff(gelecekTarih);
          if (daysDiff === null || daysDiff > 90) return;

          collected.push({
            id: `${uKey}_${rIdx}_0`,
            unitKey: uKey,
            unitLabel: unitLabelMap[uKey],
            techName: techName || 'Özel Alet',
            partNo: partNo || '-',
            seriNo: seriNo || '-',
            location: location || '-',
            sonTarih: sonTarih || '-',
            gelecekTarih,
            daysDiff,
            aciklama,
            isNew
          });
        }
      });
    });

    collected.sort((a, b) => a.daysDiff - b.daysDiff);
    return collected;
  };

  const openKalibrasyonAlertForUnits = async (
    unitKeys: Array<'at802' | 't70' | 'bell429' | 'b360'>,
    silentIfEmpty = true
  ) => {
    if (unitKeys.length === 0) return;
    const unitLabelMap: Record<'at802' | 't70' | 'bell429' | 'b360', string> = {
      at802: 'AT-802',
      t70: 'T-70',
      bell429: 'Bell-429',
      b360: 'B-360'
    };
    const unitTitle = unitKeys.map(k => unitLabelMap[k]).join(', ');
    let items = scanExpiringOzelAletForUnits(unitKeys);

    if (items.length === 0 && !liveGunTakipLoaded) {
      setKalibrasyonAlertFilter('all');
      setKalibrasyonAlertSearch('');
      setKalibrasyonAlertSelectedIds({});
      setKalibrasyonAlertModal({ unitKeys, unitTitle, loading: true, items: [] });
      try {
        const res = await fetch('/api/gun-takip-live-scan');
        if (res.ok) {
          const json = await res.json();
          if (json?.status === 'success' && json.data) {
            setLiveGunTakipData(json.data);
            setLiveGunTakipLoaded(true);
            items = scanExpiringOzelAletForUnits(unitKeys, json.data);
          }
        }
      } catch {}
      if (items.length === 0 && silentIfEmpty) {
        setKalibrasyonAlertModal(null);
        return;
      }
      setKalibrasyonAlertModal({ unitKeys, unitTitle, loading: false, items });
      return;
    }

    if (items.length === 0 && silentIfEmpty) {
      return;
    }

    setKalibrasyonAlertFilter('all');
    setKalibrasyonAlertSearch('');
    setKalibrasyonAlertSelectedIds({});
    setKalibrasyonAlertModal({ unitKeys, unitTitle, loading: false, items });
  };

  const handleTransferSelectedKalibrasyonAlerts = () => {
    if (!kalibrasyonAlertModal) return;
    const chosen = kalibrasyonAlertModal.items.filter(it => kalibrasyonAlertSelectedIds[it.id]);
    if (chosen.length === 0) {
      showNotification('Lütfen en az bir ikazlı özel alet işaretleyiniz.', 'info');
      return;
    }

    let addedCount = 0;
    let skippedCount = 0;

    setFormState(prev => {
      const rows = [...prev.kalibrasyonRows];
      const existingSignatures = new Set(
        rows
          .filter(r => r.malzemeAdi || r.parcaNo || r.seriNo)
          .map(r =>
            `${r.malzemeAdi.trim().toUpperCase()}__${r.parcaNo.trim().toUpperCase()}__${r.seriNo.trim().toUpperCase()}__${r.kalibreBitisTarihi.trim()}`
          )
      );

      chosen.forEach(item => {
        const cleanPn = item.partNo === '-' ? '' : item.partNo;
        const cleanSn = item.seriNo === '-' ? '' : item.seriNo;
        const sig = `${item.techName.trim().toUpperCase()}__${cleanPn.trim().toUpperCase()}__${cleanSn.trim().toUpperCase()}__${item.gelecekTarih.trim()}`;
        if (existingSignatures.has(sig)) {
          skippedCount++;
          return;
        }
        existingSignatures.add(sig);

        const emptyIdx = rows.findIndex(r => !r.malzemeAdi && !r.parcaNo && !r.seriNo);
        const newRow = {
          siraNo: emptyIdx >= 0 ? rows[emptyIdx].siraNo : rows.length + 1,
          malzemeAdi: item.techName,
          parcaNo: cleanPn,
          seriNo: cleanSn,
          tanimi: item.techName,
          kalibreBitisTarihi: item.gelecekTarih
        };
        if (emptyIdx >= 0) {
          rows[emptyIdx] = newRow;
        } else {
          rows.push(newRow);
        }
        addedCount++;
      });

      return { ...prev, kalibrasyonRows: rows };
    });

    setKalibrasyonAlertModal(null);
    if (addedCount > 0) {
      showNotification(
        `✅ ${addedCount} adet ikazlı özel alet kalibrasyon tablosuna aktarıldı.${skippedCount > 0 ? ` (${skippedCount} kayıt zaten mevcuttu)` : ''}`,
        'success'
      );
    } else {
      showNotification('Seçilen ürünler zaten kalibrasyon tablosunda mevcut.', 'info');
    }
  };

  // Tail number suggestions based on selected aircraft type(s)
  const matchingTailSuggestions = useMemo(() => {
    const checkedKeys = Object.entries(formState.selectedTypes)
      .filter(([, v]) => v)
      .map(([k]) => k);
    let pool: string[] = [];
    if (checkedKeys.length > 0) {
      checkedKeys.forEach(k => {
        if (AIRCRAFT_TAILS[k]) pool.push(...AIRCRAFT_TAILS[k]);
      });
    } else if (formState.selectedUnit && AIRCRAFT_TAILS[formState.selectedUnit]) {
      pool = [...AIRCRAFT_TAILS[formState.selectedUnit]];
    } else {
      Object.values(AIRCRAFT_TAILS).forEach(arr => pool.push(...arr));
    }
    const q = formState.kuyrukTescilNo.trim().toLowerCase();
    return pool.filter(t => !q || t.toLowerCase().includes(q));
  }, [formState.selectedTypes, formState.selectedUnit, formState.kuyrukTescilNo]);

  // Helper to toggle aircraft type checkbox inside ANY form and sync havaAraciTipi + selectedUnit
  const handleToggleAircraftType = (key: 'at802' | 't70' | 'bell429' | 'b360', checked: boolean, singleSelect: boolean = false) => {
    const labelsMap: Record<'at802' | 't70' | 'bell429' | 'b360', string> = {
      at802: 'AT-802',
      t70: 'T-70',
      bell429: 'Bell-429',
      b360: 'B-360'
    };
    setFormState(prev => {
      const nextTypes = singleSelect
        ? { at802: false, t70: false, bell429: false, b360: false, [key]: checked }
        : { ...prev.selectedTypes, [key]: checked };
      const checkedKeys = (['at802', 't70', 'bell429', 'b360'] as const).filter(k => nextTypes[k]);
      const nextLabel = checkedKeys.map(k => labelsMap[k]).join(', ');
      const nextUnit = checkedKeys.length === 1 ? checkedKeys[0] : (checked ? key : (checkedKeys[0] || ''));
      const nextTailList = nextUnit ? (AIRCRAFT_TAILS[nextUnit] || []) : [];
      const tailBelongsToNew = nextTailList.includes(prev.kuyrukTescilNo);
      return {
        ...prev,
        selectedTypes: nextTypes,
        havaAraciTipi: nextLabel,
        selectedUnit: checkedKeys.length > 0 ? nextUnit : '',
        kuyrukTescilNo: checkedKeys.length === 0 ? '' : (singleSelect && !tailBelongsToNew ? '' : prev.kuyrukTescilNo)
      };
    });

    if (checked && selectedType === 'ozel_alet_kalibrasyon') {
      openKalibrasyonAlertForUnits([key], true);
    }
  };

  const getFilteredDepotSuggestions = (query: string, categoryPrefer?: 'sarf' | 'kimyasal') => {
    const q = query.trim().toLowerCase();
    const checkedUnits = Object.entries(formState.selectedTypes)
      .filter(([, v]) => v)
      .map(([k]) => k);
    let pool = depotItems;
    if (categoryPrefer === 'kimyasal') {
      const chem = depotItems.filter(d => d.category === 'kimyasal');
      if (chem.length > 0) pool = chem;
    }
    const unitFiltered =
      checkedUnits.length > 0
        ? pool.filter(d => checkedUnits.includes(d.unit))
        : formState.selectedUnit && formState.selectedUnit !== 'all'
          ? pool.filter(d => d.unit === formState.selectedUnit)
          : pool;
    const effectivePool = unitFiltered.length > 0 ? unitFiltered : pool;
    return effectivePool
      .filter(d => !q || `${d.name} ${d.pn} ${d.sn}`.toLowerCase().includes(q))
      .slice(0, 15);
  };

  const getFilteredToolSuggestions = (query: string) => {
    const q = query.trim().toLowerCase();
    const checkedUnits = Object.entries(formState.selectedTypes)
      .filter(([, v]) => v)
      .map(([k]) => k);
    const unitFiltered = checkedUnits.length > 0 ? toolItems.filter(t => checkedUnits.includes(t.unit)) : toolItems;
    const effectivePool = unitFiltered.length > 0 ? unitFiltered : toolItems;
    return effectivePool
      .filter(t => !q || `${t.name} ${t.pn} ${t.sn} ${t.location}`.toLowerCase().includes(q))
      .slice(0, 15);
  };

  const getFilteredVehicleSuggestions = (query: string) => {
    const q = query.trim().toLowerCase();
    return karaAracItems
      .filter(v => !q || `${v.name} ${v.plate} ${v.brandModel} ${v.sn} ${v.typeDesc}`.toLowerCase().includes(q))
      .slice(0, 15);
  };

  const getFilteredPersonnelSuggestions = (query: string) => {
    const q = query.trim().toLowerCase();
    if (!q) return personnelList.slice(0, 12);
    return personnelList.filter(p => p.toLowerCase().includes(q)).slice(0, 12);
  };

  // Filtered helper suggestions based on current selectedType
  const filteredHelperResults = useMemo(() => {
    const q = helperQuery.trim().toLowerCase();
    if (!selectedType) return [];
    if (selectedType === 'arac_techizat_hizmet') {
      return karaAracItems
        .filter(v => !q || `${v.name} ${v.plate} ${v.brandModel} ${v.sn}`.toLowerCase().includes(q))
        .slice(0, 25);
    }
    if (selectedType === 'ozel_alet_kalibrasyon') {
      const checkedUnits = Object.entries(formState.selectedTypes)
        .filter(([, v]) => v)
        .map(([k]) => k);
      return toolItems
        .filter(t => (checkedUnits.length === 0 || checkedUnits.includes(t.unit)))
        .filter(t => !q || `${t.name} ${t.pn} ${t.sn} ${t.location}`.toLowerCase().includes(q))
        .slice(0, 25);
    }
    if (selectedType === 'bakim_sarf') {
      const checkedUnits = Object.entries(formState.selectedTypes)
        .filter(([, v]) => v)
        .map(([k]) => k);
      const chemFirst = depotItems.filter(d => d.category === 'kimyasal');
      const pool = chemFirst.length > 0 ? chemFirst : depotItems;
      return pool
        .filter(d => (checkedUnits.length === 0 || checkedUnits.includes(d.unit)))
        .filter(d => !q || `${d.name} ${d.pn} ${d.sn}`.toLowerCase().includes(q))
        .slice(0, 25);
    }
    // parca_onarim, parca_talep, hizmet_talep
    const unitKey = formState.selectedUnit;
    return depotItems
      .filter(d => !unitKey || unitKey === 'all' || d.unit === unitKey)
      .filter(d => !q || `${d.name} ${d.pn} ${d.sn}`.toLowerCase().includes(q))
      .slice(0, 25);
  }, [selectedType, helperQuery, formState.selectedUnit, formState.selectedTypes, depotItems, toolItems, karaAracItems]);

  const handleSelectHelperItem = (item: any) => {
    if (!selectedType) return;
    if (selectedType === 'parca_onarim' || selectedType === 'hizmet_talep') {
      const d = item as UnifiedDepotItem;
      setFormState(prev => ({
        ...prev,
        havaAraciTipi: d.unitLabel || prev.havaAraciTipi,
        parcaAdi: d.name,
        parcaNumarasi: d.pn === '-' ? '' : d.pn,
        seriNumarasi: d.sn === '-' ? '' : d.sn
      }));
      showNotification(`✅ "${d.name}" bilgileri forma aktarıldı.`, 'success');
    } else if (selectedType === 'bakim_sarf') {
      const d = item as UnifiedDepotItem;
      setFormState(prev => {
        const rows = [...prev.kimyasalRows];
        const emptyIdx = rows.findIndex(r => !r.malzemeTanimi && !r.malzemeKodu);
        const newRow = {
          siraNo: emptyIdx >= 0 ? rows[emptyIdx].siraNo : rows.length + 1,
          malzemeKodu: d.pn === '-' ? '' : d.pn,
          malzemeTanimi: d.name,
          birim: d.unitMeasure || 'Adet',
          miktar: '1',
          aciklama: d.unitLabel
        };
        if (emptyIdx >= 0) rows[emptyIdx] = newRow;
        else rows.push(newRow);
        return { ...prev, kimyasalRows: rows };
      });
      showNotification(`✅ "${d.name}" kimyasal/sarf tablosuna eklendi.`, 'success');
    } else if (selectedType === 'parca_talep') {
      const d = item as UnifiedDepotItem;
      setFormState(prev => {
        const rows = [...prev.parcaTalepRows];
        const emptyIdx = rows.findIndex(r => !r.parcaAdi && !r.parcaNo);
        const newRow = {
          siraNo: emptyIdx >= 0 ? rows[emptyIdx].siraNo : rows.length + 1,
          parcaAdi: d.name,
          parcaNo: d.pn === '-' ? '' : d.pn,
          alternatifPn: '',
          istenenMiktar: '1',
          oncelik: 'Yüksek' as const
        };
        if (emptyIdx >= 0) rows[emptyIdx] = newRow;
        else rows.push(newRow);
        return { ...prev, parcaTalepRows: rows };
      });
      showNotification(`✅ "${d.name}" parça listesine eklendi.`, 'success');
    } else if (selectedType === 'ozel_alet_kalibrasyon') {
      const t = item as UnifiedToolItem;
      setFormState(prev => {
        const rows = [...prev.kalibrasyonRows];
        const emptyIdx = rows.findIndex(r => !r.malzemeAdi && !r.seriNo);
        const newRow = {
          siraNo: emptyIdx >= 0 ? rows[emptyIdx].siraNo : rows.length + 1,
          malzemeAdi: t.name,
          parcaNo: t.pn === '-' ? '' : t.pn,
          seriNo: t.sn === '-' ? '' : t.sn,
          tanimi: t.name,
          kalibreBitisTarihi: t.kalibreBitis || ''
        };
        if (emptyIdx >= 0) rows[emptyIdx] = newRow;
        else rows.push(newRow);
        return { ...prev, kalibrasyonRows: rows };
      });
      showNotification(`✅ "${t.name}" kalibrasyon listesine eklendi.`, 'success');
    } else if (selectedType === 'arac_techizat_hizmet') {
      const v = item as UnifiedVehicleItem;
      setFormState(prev => ({
        ...prev,
        aracTechizatTipi: v.typeDesc || v.name,
        aracTescilNo: v.plate && v.plate !== '-' ? v.plate : '-----------',
        aracAdiTanimi: v.name,
        aracMarkaModel: v.brandModel,
        seriNumarasi: v.sn && v.sn !== '-' ? v.sn : '-----------'
      }));
      showNotification(`✅ "${v.name}" kara aracı bilgileri forma dolduruldu.`, 'success');
    }
    setIsHelperDropdownOpen(false);
    setHelperQuery('');
  };

  const renderHtmlImageList = (imgs: FormImageItem[], maxContainerPx = 580) => {
    if (!imgs || imgs.length === 0) return '';
    return `<div style="margin-top:8px; text-align:center;">
      ${imgs
        .map(img => {
          const pct = typeof img.scale === 'number' ? img.scale : 50;
          const pxWidth = Math.round(maxContainerPx * (pct / 100));
          return `<div style="margin:6px auto; text-align:center;">
            <img src="${img.dataUrl}" width="${pxWidth}" style="width:${pct}%; max-width:100%; height:auto; border:1px solid #cbd5e1;" />
          </div>`;
        })
        .join('')}
    </div>`;
  };

  // Build authentic Word (.doc) / Print HTML with exact 1cm margins and Circular OGM Logos
  const buildDocumentHtml = (
    formType: RequestFormType,
    f: typeof formState,
    forWord: boolean = false,
    talepNoBadge?: string
  ): string => {
    const esc = (s: string) =>
      String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    const cb = (checked: boolean) => (checked ? '&#9745;' : '&#9744;');
    const meta = REQUEST_TYPES.find(t => t.id === formType)!;
    const logoUrl = circularOgmLogoDataUri || OGM_WIKIMEDIA_LOGO_URL;

    const headerBanner = (title: string) => `
      <table style="width:100%; border-collapse:collapse; border:1.2px solid #475569; margin-bottom:10px;">
        <tr>
          <td style="width:16%; text-align:center; padding:5px; vertical-align:middle;">
            <img src="${logoUrl}" width="52" height="52" style="width:52px; height:52px; border-radius:50%; object-fit:cover;" />
          </td>
          <td style="width:68%; text-align:center; vertical-align:middle; padding:6px;">
            <div style="font-size:12pt; font-weight:bold; color:#64748b; letter-spacing:0.4px; font-family:Calibri,Arial,sans-serif;">
              ${esc(title)}
            </div>
            ${talepNoBadge ? `<div style="font-size:8.5pt; color:#0f172a; margin-top:2px; font-weight:bold;">TALEP NO: ${esc(talepNoBadge)} | TARİH: ${esc(f.date)}</div>` : ''}
          </td>
          <td style="width:16%; text-align:center; padding:5px; vertical-align:middle;">
            <img src="${logoUrl}" width="52" height="52" style="width:52px; height:52px; border-radius:50%; object-fit:cover;" />
          </td>
        </tr>
      </table>`;

    const signatureFooter = `
      <tr>
        <td style="border:1px solid #000; padding:6px 8px; width:36%; font-size:9.5pt;">Talep Yapan Personel Adı Soyadı:</td>
        <td style="border:1px solid #000; padding:6px 8px; width:64%; font-size:9.5pt; font-weight:bold;">${esc(f.talepYapanAdSoyad)}</td>
      </tr>
      <tr>
        <td style="border:1px solid #000; padding:6px 8px; width:36%; font-size:9.5pt; height:48px;">Talep Sahibi İmza:</td>
        <td style="border:1px solid #000; padding:6px 8px; width:64%;"></td>
      </tr>`;

    let bodyHtml = '';
    const pageBreak = forWord ? '<br clear="all" style="page-break-before:always;" />' : '';

    if (formType === 'parca_onarim' || formType === 'hizmet_talep' || formType === 'arac_techizat_hizmet') {
      const isVeh = formType === 'arac_techizat_hizmet';
      const sec1Title = isVeh ? 'Bölüm 1: Araç/Teçhizat Bilgileri' : 'Bölüm 1: Hava Aracı Bilgileri';
      const sec2Title = isVeh ? 'Bölüm 2: Arızalı Parça Bilgileri' : 'Bölüm 2: Arızalı Parça Bilgileri';
      const sec6Title = isVeh ? 'Bölüm 6: Varsa Araç/Teçhizat Arıza Geçmişi:' : 'Bölüm 6: Varsa Parçanın Arıza Geçmişi:';

      const sec1Rows = isVeh
        ? `<tr><td style="border:1px solid #000; padding:5px 8px; width:36%; font-size:9.5pt;">Araç/Teçhizat Tipi:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.aracTechizatTipi)}</td></tr>
           <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">Varsa Aracın Tescil No:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.aracTescilNo)}</td></tr>`
        : `<tr><td style="border:1px solid #000; padding:5px 8px; width:36%; font-size:9.5pt;">Hava Aracı Tipi:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.havaAraciTipi)} &nbsp;&nbsp; ${cb(f.selectedTypes.at802)} AT-802 &nbsp; ${cb(f.selectedTypes.t70)} T-70 &nbsp; ${cb(f.selectedTypes.bell429)} Bell-429 &nbsp; ${cb(f.selectedTypes.b360)} B-360</td></tr>
           <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">Hava Aracı Kuyruk Tescil No:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.kuyrukTescilNo)}</td></tr>`;

      const sec2Rows = isVeh
        ? `<tr><td style="border:1px solid #000; padding:5px 8px; width:36%; font-size:9.5pt;">Araç/Teçhizat Adı / Tanımı:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.aracAdiTanimi)}</td></tr>
           <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">Araç/Teçhizat Marka/Model:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.aracMarkaModel)}</td></tr>
           <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">(Varsa) Seri Numarası:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.seriNumarasi)}</td></tr>`
        : `<tr><td style="border:1px solid #000; padding:5px 8px; width:36%; font-size:9.5pt;">Parça Adı / Tanımı:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.parcaAdi)}</td></tr>
           <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">Parça Numarası:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.parcaNumarasi)}</td></tr>
           <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">(Varsa) Seri Numarası:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9.5pt;">${esc(f.seriNumarasi)}</td></tr>`;

      bodyHtml = `
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">${sec1Title}</td></tr>
          ${sec1Rows}
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">${sec2Title}</td></tr>
          ${sec2Rows}
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 3: Arıza Tanımı Ve Detayları</td></tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:6px 8px; vertical-align:top; height:64px; font-size:9pt;">
              <div>Arızanın Kısa Tanımı: (Örn: Aşınma, Çatlak, Fonksiyon Bozukluğu, Korozyon Vs.)</div>
              <div style="margin-top:6px; font-weight:bold; font-size:9.5pt; white-space:pre-wrap;">${esc(f.arizaKisaTanimi)}</div>
            </td>
          </tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:6px 8px; vertical-align:top; height:260px; font-size:9pt;">
              <div>Arızanın Detaylı Açıklaması: (Arızanın Nasıl Fark Edildiği, Etkilediği Sistemler, Gözlemler Vb.)</div>
              <div style="margin-top:8px; font-size:9.5pt; white-space:pre-wrap;">${esc(f.arizaDetayliAciklama)}</div>
              ${(() => {
                const arizaImgs = normalizeImageList(f.arizaGorselleri, f.arizaGorsel1, f.arizaGorsel2);
                return arizaImgs.length > 0
                  ? renderHtmlImageList(arizaImgs, 560)
                  : `<div style="margin-top:75px; text-align:center; font-style:italic; font-weight:bold;">(Arızalı Parçanın Katalogdan Alınmış Fotoğrafı Ve Güncel Durum Fotoğrafı Ekine Eklenmelidir.)</div>`;
              })()}
            </td>
          </tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 4: Talep Edilen İşlem</td></tr>
          <tr><td colspan="2" style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Lütfen Talep Edilen İşlemi İşaretleyiniz:</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:6px 8px; vertical-align:top; font-size:9.5pt;">Talep Edilen İşlem</td>
            <td style="border:1px solid #000; padding:6px 10px; font-size:9pt; line-height:1.45;">
              <div>${cb(f.islemler.onarim)} Onarım</div>
              <div>${cb(f.islemler.revizyon)} Revizyon</div>
              <div>${cb(f.islemler.kaplama)} Kaplama</div>
              <div>${cb(f.islemler.boyama)} Boyama</div>
              <div>${cb(f.islemler.kaynak)} Kaynak</div>
              <div>${cb(f.islemler.ndt)} Tahribatsız Muayene (Ndt)</div>
            </td>
          </tr>
          <tr>
            <td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">Diğer (Lütfen Belirtiniz):</td>
            <td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.islemDiger)}</td>
          </tr>
          <tr>
            <td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">
              ${sec6Title} &nbsp;&nbsp; ${cb(f.arizaGecmisi === 'Var')} Var &nbsp;&nbsp; ${cb(f.arizaGecmisi === 'Yok')} Yok
            </td>
          </tr>
          <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">Varsa Önceki Arıza Tarihi:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.oncekiArizaTarihi)}</td></tr>
          <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">Varsa Önceki Arıza Nedeni:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.oncekiArizaNedeni)}</td></tr>
        </table>
      </div>
      ${pageBreak}
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 7: Ek Notlar / Özel İstekler</td></tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:8px 10px; vertical-align:top; height:150px; font-size:9.5pt;">
              <div>Öncelik, Teslim Süresi, Özel Talimatlar Veya Dikkat Edilmesi Gereken Diğer Hususlar.</div>
              <div style="margin-top:8px; white-space:pre-wrap;">${esc(f.ekNotlar)}</div>
            </td>
          </tr>
          ${signatureFooter}
        </table>
      </div>`;
    } else if (formType === 'bakim_sarf') {
      const kRowsHtml = f.kimyasalRows
        .map(
          r => `
        <tr>
          <td style="border:1px solid #000; padding:3px 4px; text-align:center; font-size:9pt; height:21px;">${r.siraNo}</td>
          <td style="border:1px solid #000; padding:3px 6px; font-size:9pt; font-family:Consolas,Calibri,monospace;">${esc(r.malzemeKodu)}</td>
          <td style="border:1px solid #000; padding:3px 6px; font-size:9pt;">${esc(r.malzemeTanimi)}</td>
          <td style="border:1px solid #000; padding:3px 6px; font-size:9pt; text-align:center;">${esc(r.birim)}</td>
          <td style="border:1px solid #000; padding:3px 6px; font-size:9pt; text-align:center;">${esc(r.miktar)}</td>
          <td style="border:1px solid #000; padding:3px 6px; font-size:9pt;">${esc(r.aciklama)}</td>
        </tr>`
        )
        .join('');

      bodyHtml = `
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 1: Hava Aracı Birimi</td></tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:6px 8px; font-size:9.5pt; vertical-align:top;">Hava Aracı Tipi:</td>
            <td colspan="4" style="border:1px solid #000; padding:6px 10px; font-size:9pt; line-height:1.4;">
              <div>${cb(f.selectedTypes.at802)} AT-802</div>
              <div>${cb(f.selectedTypes.t70)} T-70</div>
              <div>${cb(f.selectedTypes.bell429)} Bell-429</div>
              <div>${cb(f.selectedTypes.b360)} B-360</div>
            </td>
          </tr>
          <tr><td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 2: Talep Edilen Malzemeler</td></tr>
          <tr style="background:#f3f4f6;">
            <th style="border:1px solid #000; padding:4px; width:8%; font-size:8.5pt;">Sıra No</th>
            <th style="border:1px solid #000; padding:4px; width:24%; font-size:8.5pt;">Malzeme Kodu / Parça Numarası</th>
            <th style="border:1px solid #000; padding:4px; width:24%; font-size:8.5pt;">Malzeme Tanımı</th>
            <th style="border:1px solid #000; padding:4px; width:20%; font-size:8.5pt;">Birim (Adet, Lt, Kg, Cc)</th>
            <th style="border:1px solid #000; padding:4px; width:9%; font-size:8.5pt;">Miktar</th>
            <th style="border:1px solid #000; padding:4px; width:15%; font-size:8.5pt;">Açıklama</th>
          </tr>
          ${kRowsHtml}
          <tr>
            <td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:8px; font-size:8.5pt; font-style:italic; line-height:1.35;">
              <strong>*** ÖNEMLİ NOT:***</strong><br/>
              Bu Form, Hava Aracı Bakım Ve Onarım Faaliyetlerinde kullanılan Sarf Malzemelerinin Talep Edilmesi Amacıyla Hazırlanmıştır. Talep Edebileceğiniz Ürün Çeşitliliği; <strong>Motor Yağı Ve Klima Yağı</strong> Gibi Sıvı Malzemelerden, <strong>Gres Ve Hidrolik</strong> Sıvılara, Ayrıca <strong>Nitrojen, Klima Gazı</strong> Gibi Özel Gaz Çeşitlerine Ve Diğer bakım esnasında kullanılacak Parça temizlik mayisi, Alkol okil asidi vs. gibi Kimyasalları Kapsamaktadır.
            </td>
          </tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:6px 8px; font-size:9.5pt;">Talep Yapan Personel Adı Soyadı:</td>
            <td colspan="4" style="border:1px solid #000; padding:6px 8px; font-size:9.5pt; font-weight:bold;">${esc(f.talepYapanAdSoyad)}</td>
          </tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:6px 8px; font-size:9.5pt; height:48px;">Talep Sahibi İmza:</td>
            <td colspan="4" style="border:1px solid #000; padding:6px 8px;"></td>
          </tr>
        </table>
      </div>`;
    } else if (formType === 'parca_talep') {
      const pRowsHtml = f.parcaTalepRows
        .map(
          r => `
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

      bodyHtml = `
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000; margin-bottom:12px;">
          <tr>
            <td style="width:62%; background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">1. Hava Aracı Bilgileri</td>
            <td style="width:38%; background:#d9d9d9; border:1px solid #000; padding:4px 8px;"></td>
          </tr>
          <tr>
            <td style="border:1px solid #000; padding:8px 10px; vertical-align:top; font-size:9.5pt;">1. Hava Aracı Tipi : <strong>${esc(f.havaAraciTipi)}</strong></td>
            <td style="border:1px solid #000; padding:6px 12px; font-size:9.5pt; line-height:1.45;">
              <div>${cb(f.selectedTypes.at802)} AT-802</div>
              <div>${cb(f.selectedTypes.t70)} T-70</div>
              <div>${cb(f.selectedTypes.bell429)} Bell-429</div>
              <div>${cb(f.selectedTypes.b360)} B-360</div>
            </td>
          </tr>
        </table>
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">2-Talep Edilen Malzemeler</td></tr>
          <tr style="background:#f3f4f6;">
            <th style="border:1px solid #000; padding:4px; width:7%; font-size:8.5pt;">Sıra No</th>
            <th style="border:1px solid #000; padding:4px; width:24%; font-size:8.5pt; text-align:left;">Parça Adı/Tanımı</th>
            <th style="border:1px solid #000; padding:4px; width:19%; font-size:8.5pt; text-align:left;">Parça Numarası (Part Number)</th>
            <th style="border:1px solid #000; padding:4px; width:20%; font-size:8.5pt; text-align:left;">Alternatif Parça Numaraları</th>
            <th style="border:1px solid #000; padding:4px; width:14%; font-size:8.5pt;">İstenen Miktar</th>
            <th style="border:1px solid #000; padding:4px; width:16%; font-size:8.5pt;">Parçanın Öncelik Derecesi</th>
          </tr>
          ${pRowsHtml}
          <tr><td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; text-align:center; font-weight:bold; font-size:9pt;">Parcanın Öncelik Derecesi kısmına yazılacaklar</td></tr>
          <tr><td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:9pt;"><span style="color:#dc2626; font-weight:bold;">Yüksek</span> – Kısa Sürede Gerekli (Arıza vb)</td></tr>
          <tr><td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:9pt;"><span style="color:#eab308; font-weight:bold;">Normal</span> – Planlı Bakım İçin</td></tr>
          <tr><td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:9pt;"><span style="color:#16a34a; font-weight:bold;">Düşük</span> – Depo Miktarını Tamamlamak İçin</td></tr>
          <tr><td colspan="6" style="border:1px solid #000; padding:5px 8px; font-size:8.5pt; font-style:italic;">*Ek satırlara ihtiyaç duyulması halinde bu bölümü çoğaltabilirsiniz.*</td></tr>
        </table>
      </div>
      ${pageBreak}
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000; margin-bottom:12px;">
          <tr><td style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">3. Parça Kataloğu Görseli</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:10px; vertical-align:top; height:400px;">
              <div style="font-size:9.5pt; margin-bottom:10px;">Lütfen istenen parçanın katalogdan alınmış görselini formun ekine ekleyiniz.<br/>Doğru parçanın tespit edilmesine yardımcı olacaktır.</div>
              ${(() => {
                const katImgs = normalizeImageList(f.katalogGorselleri, f.katalogGorseli);
                return katImgs.length > 0 ? renderHtmlImageList(katImgs, 580) : '';
              })()}
            </td>
          </tr>
        </table>
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">4. Ek Notlar/Açıklamalar</td></tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:8px 10px; vertical-align:top; height:140px; font-size:9.5pt;">
              <div>Parçanın özel gereksinimlerini belirten ek bilgiler var ise:</div>
              <div style="margin-top:6px; white-space:pre-wrap;">${esc(f.ekNotlar)}</div>
            </td>
          </tr>
          ${signatureFooter}
        </table>
      </div>`;
    } else if (formType === 'ozel_alet_kalibrasyon') {
      const calRowsHtml = f.kalibrasyonRows
        .map(
          r => `
        <tr>
          <td style="border:1px solid #000; padding:2.5px 4px; text-align:center; font-size:8.5pt; height:19px;">${r.siraNo}</td>
          <td style="border:1px solid #000; padding:2.5px 5px; font-size:8.5pt;">${esc(r.malzemeAdi)}</td>
          <td style="border:1px solid #000; padding:2.5px 5px; font-size:8.5pt; font-family:Consolas,monospace;">${esc(r.parcaNo)}</td>
          <td style="border:1px solid #000; padding:2.5px 5px; font-size:8.5pt; font-family:Consolas,monospace;">${esc(r.seriNo)}</td>
          <td style="border:1px solid #000; padding:2.5px 5px; font-size:8.5pt;">${esc(r.tanimi)}</td>
          <td style="border:1px solid #000; padding:2.5px 5px; text-align:center; font-size:8.5pt;">${esc(r.kalibreBitisTarihi)}</td>
        </tr>`
        )
        .join('');

      bodyHtml = `
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 1: Hava Aracı Bilgileri</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:6px 8px; width:45%; font-size:9.5pt; vertical-align:top;">Departman / Bölüm:</td>
            <td style="border:1px solid #000; padding:6px 10px; font-size:9pt; line-height:1.45;">
              <div>${cb(f.selectedTypes.at802)} AT-802</div>
              <div>${cb(f.selectedTypes.t70)} T-70</div>
              <div>${cb(f.selectedTypes.bell429)} Bell-429</div>
              <div>${cb(f.selectedTypes.b360)} B-360</div>
            </td>
          </tr>
          <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">Diğer:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.departmanDiger)}</td></tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 2: Kalibrasyon Gereklilikleri</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:6px 8px; font-size:9.5pt; vertical-align:top;">Kalibrasyon Türü:</td>
            <td style="border:1px solid #000; padding:6px 10px; font-size:9pt; line-height:1.45;">
              <div>${cb(f.kalibrasyonTuru.ilk)} İlk Kalibrasyon</div>
              <div>${cb(f.kalibrasyonTuru.periyodik)} Periyodik Kalibrasyon</div>
              <div>${cb(f.kalibrasyonTuru.onarimSonrasi)} Onarım Sonrası Kalibrasyon</div>
            </td>
          </tr>
          <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">Diğer (Lütfen Belirtiniz):</td><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.kalibrasyonDiger)}</td></tr>
          <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">Talep Edilen Son Tarih: [GG/AA/YYYY]</td><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.talepEdilenSonTarih)}</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:6px 8px; font-size:9.5pt; vertical-align:top;">Öncelik Seviyesi:</td>
            <td style="border:1px solid #000; padding:6px 10px; font-size:9pt; line-height:1.45;">
              <div>${cb(f.oncelikSeviyesi === 'Yüksek (Acil)')} Yüksek (Acil)</div>
              <div>${cb(f.oncelikSeviyesi === 'Normal')} Normal</div>
              <div>${cb(f.oncelikSeviyesi === 'Düşük')} Düşük</div>
            </td>
          </tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 3: Ek Notlar / Dikkat Edilmesi Gerekenler</td></tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:8px; vertical-align:top; height:320px; font-size:9pt;">
              <div>(Kalibrasyon İşlemiyle İlgili Özel Talimatlar, Ölçüm Noktaları, Raporlama Formatı Veya Diğer Önemli Bilgiler.)</div>
              <div style="margin-top:8px; white-space:pre-wrap; font-size:9.5pt;">${esc(f.ekNotlar)}</div>
            </td>
          </tr>
        </table>
      </div>
      ${pageBreak}
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="6" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">Bölüm 4: Kalibrasyon Yapılacak Malzeme Listesi</td></tr>
          <tr><td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Aşağıdaki tabloya kalibrasyon talep edilen tüm malzemelerin bilgilerini giriniz.</td></tr>
          <tr style="background:#f3f4f6;">
            <th style="border:1px solid #000; padding:4px; width:8%; font-size:8pt;">Sıra No</th>
            <th style="border:1px solid #000; padding:4px; width:20%; font-size:8pt;">Malzeme Adı</th>
            <th style="border:1px solid #000; padding:4px; width:16%; font-size:8pt;">Parça Numarası</th>
            <th style="border:1px solid #000; padding:4px; width:16%; font-size:8pt;">Seri Numarası (S/N)</th>
            <th style="border:1px solid #000; padding:4px; width:22%; font-size:8pt;">Malzemenin Tanımı</th>
            <th style="border:1px solid #000; padding:4px; width:18%; font-size:8pt;">Kalibre Bitiş Tarihi</th>
          </tr>
          ${calRowsHtml}
          <tr><td colspan="6" style="border:1px solid #000; padding:4px 8px; font-size:8.5pt; font-style:italic;">*Ek satırlara ihtiyaç duyulması halinde bu bölümü çoğaltabilirsiniz.*</td></tr>
          <tr>
            <td colspan="3" style="border:1px solid #000; padding:6px 8px; font-size:9.5pt;">Talep Yapan Personel Adı Soyadı:</td>
            <td colspan="3" style="border:1px solid #000; padding:6px 8px; font-size:9.5pt; font-weight:bold;">${esc(f.talepYapanAdSoyad)}</td>
          </tr>
          <tr>
            <td colspan="3" style="border:1px solid #000; padding:6px 8px; font-size:9.5pt; height:44px;">Talep Sahibi İmza:</td>
            <td colspan="3" style="border:1px solid #000; padding:6px 8px;"></td>
          </tr>
        </table>
      </div>`;
    } else if (formType === 'ozel_alet_alim_imal') {
      bodyHtml = `
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000; margin-bottom:10px;">
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">1. Hava Aracı Bilgileri</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:6px 8px; width:42%; font-size:9pt; vertical-align:top;">Hava Aracı Tipi:</td>
            <td style="border:1px solid #000; padding:6px 10px; font-size:9pt; line-height:1.35;">
              <div>${cb(f.selectedTypes.at802)} AT-802</div>
              <div>${cb(f.selectedTypes.t70)} T-70</div>
              <div>${cb(f.selectedTypes.bell429)} Bell-429</div>
              <div>${cb(f.selectedTypes.b360)} B-360</div>
            </td>
          </tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">2. Genel Özellikler</td></tr>
          <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">A) Aletin Kullanım Amacı:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.kullanimAmaci)}</td></tr>
          <tr><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">B) Hangi Görevlerde/Bakımlarda Kullanılacağı:</td><td style="border:1px solid #000; padding:5px 8px; font-size:9pt;">${esc(f.gorevBakimlar)}</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:5px 8px; font-size:9pt; vertical-align:top;">C) Kullanım Sıklığı:</td>
            <td style="border:1px solid #000; padding:5px 10px; font-size:9pt; line-height:1.35;">
              <div>${cb(f.kullanimSikligi === 'Sürekli')} Sürekli</div>
              <div>${cb(f.kullanimSikligi === 'Dönemsel')} Dönemsel</div>
              <div>${cb(f.kullanimSikligi === 'Acil Durumlarda')} Acil Durumlarda</div>
            </td>
          </tr>
          <tr>
            <td style="border:1px solid #000; padding:5px 8px; font-size:9pt; vertical-align:top;">D) Çalışma Ortamı:</td>
            <td style="border:1px solid #000; padding:5px 10px; font-size:9pt; line-height:1.35;">
              <div>${cb(f.calismaOrtami === 'Açık Alan')} Açık Alan</div>
              <div>${cb(f.calismaOrtami === 'Kapalı Alan')} Kapalı Alan</div>
              <div>${cb(f.calismaOrtami === 'Deniz / Su Alanı')} Deniz / Su Alanı</div>
            </td>
          </tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">3. Teknik Özellikler</td></tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9pt;">A) Boyut ve Ağırlık</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Uzunluk:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.uzunluk)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Genişlik:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.genislik)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Yükseklik:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.yukseklik)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Ağırlık:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.agirlik)}</td></tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9pt;">B) Dayanıklılık ve Koruma</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Isıya Dayanıklılık ( °C'ye Kadar):</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.isiDayaniklilik)}</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:4px 8px; font-size:9pt; vertical-align:top;">Darbelere Karşı Dayanıklılık:</td>
            <td style="border:1px solid #000; padding:4px 10px; font-size:9pt;">
              <div>${cb(f.darbeDayaniklilik === 'Hafif')} Hafif</div>
              <div>${cb(f.darbeDayaniklilik === 'Orta')} Orta</div>
              <div>${cb(f.darbeDayaniklilik === 'Yüksek')} Yüksek</div>
            </td>
          </tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Su Geçirmezlik:</td><td style="border:1px solid #000; padding:4px 10px; font-size:9pt;">${cb(f.suGecirmezlik === 'Evet')} Evet &nbsp; ${cb(f.suGecirmezlik === 'Hayır')} Hayır</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Toz Geçirmezlik:</td><td style="border:1px solid #000; padding:4px 10px; font-size:9pt;">${cb(f.tozGecirmezlik === 'Evet')} Evet &nbsp; ${cb(f.tozGecirmezlik === 'Hayır')} Hayır</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Kimyasal Maddelere Dayanıklılık:</td><td style="border:1px solid #000; padding:4px 10px; font-size:9pt;">${cb(f.kimyasalDayaniklilik === 'Evet')} Evet &nbsp; ${cb(f.kimyasalDayaniklilik === 'Hayır')} Hayır</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">UV Işınlarına Karşı Koruma:</td><td style="border:1px solid #000; padding:4px 10px; font-size:9pt;">${cb(f.uvKoruma === 'Evet')} Evet &nbsp; ${cb(f.uvKoruma === 'Hayır')} Hayır</td></tr>
        </table>
      </div>
      ${pageBreak}
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">C) Malzeme Türü</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:5px 8px; width:42%; font-size:9pt; vertical-align:top;">Metaller:</td>
            <td style="border:1px solid #000; padding:5px 10px; font-size:8.5pt; line-height:1.35;">
              <div>${cb(f.metaller.celik)} Çelik</div>
              <div>${cb(f.metaller.paslanmazCelik)} Paslanmaz Çelik</div>
              <div>${cb(f.metaller.aluminyum)} Alüminyum</div>
              <div>${cb(f.metaller.bakir)} Bakır</div>
              <div>${cb(f.metaller.titanyum)} Titanyum</div>
              <div>${cb(f.metaller.pirinc)} Pirinç</div>
              <div>${cb(f.metaller.dokmeDemir)} Dökme Demir</div>
            </td>
          </tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Diğer:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.metallerDiger)}</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:5px 8px; font-size:9pt; vertical-align:top;">Ametaller:</td>
            <td style="border:1px solid #000; padding:5px 10px; font-size:8.5pt; line-height:1.35;">
              <div>${cb(f.ametaller.plastik)} Plastik (ABS, PVC, PP)</div>
              <div>${cb(f.ametaller.kaucuk)} Kauçuk</div>
              <div>${cb(f.ametaller.seramik)} Seramik</div>
              <div>${cb(f.ametaller.cam)} Cam</div>
              <div>${cb(f.ametaller.karbonFiber)} Karbon Fiber</div>
              <div>${cb(f.ametaller.kompozit)} Kompozit Malzemeler</div>
            </td>
          </tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Diğer:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.ametallerDiger)}</td></tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">D) Yüzey İşlemi ve Kaplama</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Kaplama:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.kaplama)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Boya ve Renk Kodu:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.boyaRenkKodu)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Cilalama:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.cilalama)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Diğer:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.yuzeyDiger)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Korozyona Karşı Dayanıklılık:</td><td style="border:1px solid #000; padding:4px 10px; font-size:9pt;">${cb(f.korozyonDayaniklilik === 'Evet')} Evet &nbsp; ${cb(f.korozyonDayaniklilik === 'Hayır')} Hayır</td></tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">E) Enerji ve Çalışma Mekanizması</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:5px 8px; font-size:9pt; vertical-align:top;">Enerji Kaynağı:</td>
            <td style="border:1px solid #000; padding:5px 10px; font-size:8.5pt; line-height:1.35;">
              <div>${cb(f.enerjiKaynagi.manuel)} Manuel Çalışma</div>
              <div>${cb(f.enerjiKaynagi.pnomatik)} Pnömatik</div>
              <div>${cb(f.enerjiKaynagi.hidrolik)} Hidrolik</div>
              <div>${cb(f.enerjiKaynagi.elektrik)} Elektrik</div>
            </td>
          </tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Varsa Batarya (Türü ve Kapasitesi):</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.bataryaBilgisi)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Varsa Çalışma Voltajı (V):</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.calismaVoltaji)}</td></tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Varsa Güç Tüketimi (Watt):</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.gucTuketimi)}</td></tr>
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">F) Güvenlik ve Standartlar</td></tr>
          <tr>
            <td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Uluslararası Standartlara Uygunluk (Varsa):</td>
            <td style="border:1px solid #000; padding:4px 10px; font-size:9pt;">${cb(f.standartlar.ce)} CE &nbsp;&nbsp; ${cb(f.standartlar.iso)} ISO</td>
          </tr>
          <tr><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">Diğer:</td><td style="border:1px solid #000; padding:4px 8px; font-size:9pt;">${esc(f.standartDiger)}</td></tr>
        </table>
      </div>
      ${pageBreak}
      <div class="a4-page">
        ${headerBanner(meta.formTitle)}
        <table style="width:100%; border-collapse:collapse; border:1.5px solid #000;">
          <tr><td colspan="2" style="background:#d9d9d9; border:1px solid #000; padding:4px 8px; font-weight:bold; font-size:9.5pt;">4. Çizim Alanı</td></tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:8px; vertical-align:top; height:320px; font-size:9pt;">
              <div><strong>Formu Dolduran Kişi, FORMDA OLMAYAN FAKAT İLAVE ETMEK İSTEDİĞİ TÜM ÖZELLİKLERİ YAZABİLİR.</strong> Aşağıdaki Alanı Tasarım Ve Çizim İçin Kullanabilir.</div>
              <div style="margin-top:8px; white-space:pre-wrap; font-size:9.5pt;">${esc(f.cizimNotlari)}</div>
              ${(() => {
                const cizimImgs = normalizeImageList(f.cizimGorselleri, f.cizimGorseli);
                return cizimImgs.length > 0 ? renderHtmlImageList(cizimImgs, 560) : '';
              })()}
            </td>
          </tr>
          <tr>
            <td colspan="2" style="border:1px solid #000; padding:8px; vertical-align:top; height:240px; font-size:9pt;">
              <div>Lütfen İstenen Ürünün Varsa Kataloğundan Alınmış Fotoğrafı Veya Örnek Ürün Linklerini Formun Ekine Ekleyiniz.</div>
              <div style="margin-top:8px; white-space:pre-wrap; font-size:9.5pt;">${esc(f.ornekUrunLinkNotlari)}</div>
              ${(() => {
                const ornekImgs = normalizeImageList(f.ornekUrunGorselleri, f.ornekUrunGorseli);
                return ornekImgs.length > 0 ? renderHtmlImageList(ornekImgs, 560) : '';
              })()}
            </td>
          </tr>
          ${signatureFooter}
        </table>
      </div>`;
    }

    const wordAttrs = forWord
      ? 'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"'
      : '';

    return `<!DOCTYPE html>
<html ${wordAttrs}>
<head>
  <meta charset="utf-8" />
  <title>${esc(meta.formTitle)}</title>
  <style>
    @page { size: A4 portrait; margin: 1cm 1cm 1cm 1cm; }
    @page WordSection1 { size: 21cm 29.7cm; margin: 1cm 1cm 1cm 1cm; mso-header-margin: 0.5cm; mso-footer-margin: 0.5cm; }
    div.WordSection1 { page: WordSection1; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; font-family: Calibri, Arial, sans-serif; color: #000; background: #fff; font-size: 9.5pt; }
    .a4-page { width: 100%; page-break-after: always; padding: 2px 0; }
    .a4-page:last-child { page-break-after: auto; }
    table { border-collapse: collapse; width: 100%; }
  </style>
</head>
<body>
  <div class="WordSection1">${bodyHtml}</div>
</body>
</html>`;
  };

  const triggerDownloadWord = (formType: RequestFormType, data: typeof formState, talepNo?: string) => {
    const html = buildDocumentHtml(formType, data, true, talepNo);
    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const prefix = talepNo ? `${talepNo}_` : '';
    const slug = formType.toUpperCase();
    a.download = `OGM_${prefix}${slug}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showNotification('📄 Word (.doc) belgesi 1cm sayfa kenar boşlukları ile indirildi.', 'success');
  };

  const triggerPrintPdf = (formType: RequestFormType, data: typeof formState, talepNo?: string) => {
    const html = buildDocumentHtml(formType, data, false, talepNo);
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

  const buildSummaryText = (formType: RequestFormType, f: typeof formState): string => {
    if (formType === 'bakim_sarf') {
      const filled = f.kimyasalRows.filter(r => r.malzemeTanimi || r.malzemeKodu);
      return filled.map(r => `${r.malzemeTanimi || r.malzemeKodu} (${r.miktar || 1} ${r.birim || 'Adet'})`).join(', ') || 'Sarf/Kimyasal Talebi';
    }
    if (formType === 'parca_talep') {
      const filled = f.parcaTalepRows.filter(r => r.parcaAdi || r.parcaNo);
      return filled.map(r => `${r.parcaAdi || r.parcaNo} (${r.istenenMiktar || 1})`).join(', ') || 'Yedek Parça Talebi';
    }
    if (formType === 'ozel_alet_kalibrasyon') {
      const filled = f.kalibrasyonRows.filter(r => r.malzemeAdi || r.seriNo);
      return filled.map(r => `${r.malzemeAdi || r.parcaNo} [S/N: ${r.seriNo || '-'}]`).join(', ') || 'Kalibrasyon Talebi';
    }
    if (formType === 'ozel_alet_alim_imal') {
      return f.kullanimAmaci || f.gorevBakimlar || 'Özel Alet Alım / İmal / Üretim Talebi';
    }
    if (formType === 'arac_techizat_hizmet') {
      return `${f.aracAdiTanimi || f.aracTechizatTipi || 'Araç/Teçhizat'} — ${f.arizaKisaTanimi || 'Hizmet Talebi'}`;
    }
    return `${f.parcaAdi || f.parcaNumarasi || 'Parça'} — ${f.arizaKisaTanimi || 'Talep'}`;
  };

  const handleSaveAndSubmit = async () => {
    if (!selectedType) return;
    if (!formState.talepYapanAdSoyad.trim()) {
      showNotification('⚠️ Lütfen formun altındaki "Talep Yapan Personel Adı Soyadı" alanını doldurunuz.', 'error');
      return;
    }

    const meta = REQUEST_TYPES.find(t => t.id === selectedType)!;
    const summary = buildSummaryText(selectedType, formState);
    const chosenUnits = Object.entries(formState.selectedTypes)
      .filter(([, v]) => v)
      .map(([k]) => k.toUpperCase())
      .join(', ');

    setIsSaving(true);
    try {
      const res = await fetch('/api/onay-bekleyen-talepler', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          formType: selectedType,
          formTypeLabel: meta.label,
          formTitle: meta.formTitle,
          unit: formState.selectedUnit || 'genel',
          unitLabel: formState.havaAraciTipi || chosenUnits || formState.aracTechizatTipi || 'Genel',
          tailNo: formState.kuyrukTescilNo || formState.aracTescilNo || '-',
          summary,
          requestedBy: formState.talepYapanAdSoyad.trim(),
          date: formState.date || getTodayTr(),
          formData: formState
        })
      });
      const json = await res.json();
      if (res.ok && json.status === 'success') {
        const talepNo = json.talepNo || `${new Date().getFullYear()}-0001`;
        await fetchSubmittedArchive();
        triggerDownloadWord(selectedType, formState, talepNo);
        setSaveSuccessModal({ talepNo, formTitle: meta.formTitle, record: json.request || null });
        showNotification(`✅ Talebiniz sisteme işlendi (Talep No: ${talepNo}).`, 'success');
      } else {
        showNotification(json?.message || 'Talep kaydedilemedi.', 'error');
      }
    } catch (err: any) {
      showNotification('Sunucu bağlantı hatası: ' + (err?.message || err), 'error');
    } finally {
      setIsSaving(false);
    }
  };

  if (!isOpen) return null;

  const currentTypeMeta = selectedType ? REQUEST_TYPES.find(t => t.id === selectedType) : null;

  return (
    <div
      className="fixed inset-0 z-[10002] bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4 animate-fade-in"
      onClick={() => setActiveCellAutocomplete(null)}
    >
      <div
        className="bg-white border border-slate-200 rounded-3xl w-full max-w-[1280px] h-[95vh] flex flex-col shadow-2xl overflow-hidden text-slate-900"
        onClick={e => e.stopPropagation()}
      >
        {/* Modal Top Header */}
        <div className="px-5 py-3.5 bg-white border-b border-slate-200 flex items-center justify-between gap-3 flex-wrap shrink-0">
          <div className="flex items-center gap-3">
            {selectedType && activeMainTab === 'create' && (
              <button
                type="button"
                onClick={() => {
                  setSelectedType(null);
                  setActiveCellAutocomplete(null);
                }}
                className="p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 cursor-pointer flex items-center gap-1 text-xs font-bold"
              >
                <ArrowLeft className="w-4 h-4" />
                <span className="hidden sm:inline">Talep Türleri</span>
              </button>
            )}
            <div className="w-10 h-10 rounded-2xl bg-emerald-50 text-emerald-700 border border-emerald-200 flex items-center justify-center">
              <FileText className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm sm:text-base font-black text-slate-900 uppercase tracking-wide">
                {selectedType && activeMainTab === 'create' ? currentTypeMeta?.label : '10. TALEP OLUŞTUR'}
              </h2>
              <p className="text-[11px] text-slate-500">
                Resmi OGM Havacılık Word/PDF Talep Formları &bull; Canlı Önizleme, Otomatik Veri Aktarımı ve Onay Takibi
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            {/* Switch between Create, Pending Approvals & Approved Archive */}
            <div className="flex items-center p-1 bg-slate-100 border border-slate-200 rounded-xl flex-wrap gap-0.5">
              <button
                type="button"
                onClick={() => setActiveMainTab('create')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer transition-all ${
                  activeMainTab === 'create' ? 'bg-emerald-600 text-white shadow-xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Yeni Talep Oluştur
              </button>
              <button
                type="button"
                onClick={() => {
                  setActiveMainTab('archive');
                  fetchSubmittedArchive();
                }}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer transition-all flex items-center gap-1.5 ${
                  activeMainTab === 'archive' ? 'bg-amber-500 text-white font-black shadow-xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Clock className="w-3.5 h-3.5" />
                <span>Onay Bekleyen Talepler</span>
                <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-mono ${activeMainTab === 'archive' ? 'bg-white/25 text-white' : 'bg-slate-200 text-slate-700'}`}>
                  {pendingSubmittedList.length}
                </span>
              </button>
              {approvedArchiveList.length > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    setActiveMainTab('approved_archive');
                    fetchSubmittedArchive();
                  }}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold cursor-pointer transition-all flex items-center gap-1.5 ${
                    activeMainTab === 'approved_archive'
                      ? 'bg-emerald-700 text-white font-black shadow-xs'
                      : 'text-emerald-800 hover:bg-emerald-50'
                  }`}
                >
                  <Archive className="w-3.5 h-3.5" />
                  <span>Onaylanan Talepler Arşivi</span>
                  <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-mono ${activeMainTab === 'approved_archive' ? 'bg-white/25 text-white' : 'bg-emerald-100 text-emerald-800'}`}>
                    {approvedArchiveList.length}
                  </span>
                </button>
              )}
            </div>

            {selectedType && activeMainTab === 'create' && (
              <>
                <button
                  type="button"
                  onClick={() => triggerDownloadWord(selectedType, formState)}
                  className="px-3.5 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-black flex items-center gap-1.5 cursor-pointer shadow-sm"
                >
                  <Download className="w-4 h-4" />
                  <span>Word (.doc) İndir</span>
                </button>
                <button
                  type="button"
                  onClick={() => triggerPrintPdf(selectedType, formState)}
                  className="px-3.5 py-2 rounded-xl bg-slate-700 hover:bg-slate-800 text-white text-xs font-black flex items-center gap-1.5 cursor-pointer shadow-sm"
                >
                  <Printer className="w-4 h-4" />
                  <span>Yazdır / PDF</span>
                </button>
                <button
                  type="button"
                  disabled={isSaving}
                  onClick={handleSaveAndSubmit}
                  className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-black flex items-center gap-1.5 cursor-pointer shadow-md"
                >
                  <Send className="w-4 h-4" />
                  <span>{isSaving ? 'Kaydediliyor...' : 'Kaydet & Sisteme İlet'}</span>
                </button>
              </>
            )}

            <button
              type="button"
              onClick={onClose}
              className="p-2 rounded-xl bg-slate-100 hover:bg-rose-50 text-slate-600 hover:text-rose-600 border border-slate-200 cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* BODY */}
        {activeMainTab === 'archive' || activeMainTab === 'approved_archive' ? (
          /* ONAY BEKLEYEN TALEPLER / ONAYLANAN TALEPLER ARŞİVİ TAB */
          (() => {
            const isApprovedTab = activeMainTab === 'approved_archive';
            const displayRows = isApprovedTab ? approvedArchiveList : pendingSubmittedList;
            return (
              <div className="flex-1 overflow-y-auto p-5 bg-slate-50 space-y-4">
                <div className="flex items-center justify-between flex-wrap gap-2">
                  <div>
                    <h3 className={`text-sm font-black uppercase ${isApprovedTab ? 'text-emerald-800' : 'text-amber-800'}`}>
                      {isApprovedTab
                        ? 'ONAYLANMIŞ TALEPLER ARŞİVİ (ARŞİV ŞARTNAMELER LİSTESİ)'
                        : 'ONAY BEKLEYEN TALEPLER LİSTESİ (YIL - SIRA NO)'}
                    </h3>
                    <p className="text-[11px] text-slate-600">
                      {isApprovedTab
                        ? 'Durumu "ONAYLANDI" olan talepler arşivlenir. Talep numarasına tıklayarak Taskline online akış geçmişini görebilir, ıslak imzalı PDF belgelerini yükleyip görüntüleyebilirsiniz.'
                        : 'Kaydedilen her form otomatik YYYY-SIRA NO numarası alır. Islak imzalı PDF evraklarınızı "BELGE YÜKLE" ile Arşiv Şartnameler alanına yükleyebilirsiniz.'}
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => fetchSubmittedArchive()}
                      className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-slate-100 border border-slate-200 text-slate-700 text-[11px] font-bold flex items-center gap-1 cursor-pointer"
                    >
                      <RefreshCw className={`w-3.5 h-3.5 ${isLoadingArchive ? 'animate-spin' : ''}`} />
                      <span>Yenile</span>
                    </button>
                  </div>
                </div>

                {displayRows.length === 0 ? (
                  <div className="p-12 border-2 border-dashed border-slate-300 bg-white rounded-2xl text-center text-slate-500 text-xs">
                    {isApprovedTab
                      ? 'Henüz durumu "ONAYLANDI" olarak arşivlenmiş talep bulunmamaktadır.'
                      : 'Onay bekleyen talep bulunmamaktadır. "Yeni Talep Oluştur" sekmesinden form oluşturabilirsiniz.'}
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white shadow-xs">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-100 border-b border-slate-200 text-[10px] uppercase font-mono text-slate-600">
                          <th className="p-3">Talep No</th>
                          <th className="p-3">Tarih</th>
                          <th className="p-3">Talep Türü</th>
                          <th className="p-3">Birim / Araç</th>
                          <th className="p-3">Özet</th>
                          <th className="p-3">Talep Eden</th>
                          <th className="p-3">Durum</th>
                          <th className="p-3">Belge (Islak İmzalı)</th>
                          <th className="p-3 text-right">Form / İşlem</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {displayRows.map(rec => {
                          const approved = isApprovedStatus(rec.status);
                          const isPromptActive = promptBelgeTalepNo === rec.talepNo;
                          const displayDocTs = rec.wetSignedUploadedAt || rec.wetSignedClickedAt || '';
                          const viewUrl = rec.wetSignedDocUrl || OGM_WET_SIGNED_VIEW_URL;
                          return (
                            <tr key={rec.talepNo || rec.id} className="hover:bg-slate-50">
                              <td className="p-3 font-mono font-black whitespace-nowrap">
                                {approved ? (
                                  <button
                                    type="button"
                                    onClick={() => handleOpenTasklineFlow(rec.talepNo)}
                                    title="Taskline Online Akış Geçmişini Görüntüle"
                                    className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-blue-50 hover:bg-blue-600 text-blue-700 hover:text-white border border-blue-200 cursor-pointer transition-colors"
                                  >
                                    <GitBranch className="w-3.5 h-3.5" />
                                    <span>{rec.talepNo}</span>
                                  </button>
                                ) : (
                                  <span className="text-emerald-700">{rec.talepNo}</span>
                                )}
                              </td>
                              <td className="p-3 font-mono text-slate-600 whitespace-nowrap">{rec.date}</td>
                              <td className="p-3 font-bold text-slate-900">{rec.formTypeLabel}</td>
                              <td className="p-3 text-slate-700">{rec.unitLabel}</td>
                              <td className="p-3 text-slate-600 max-w-xs truncate" title={rec.summary}>
                                {rec.summary}
                              </td>
                              <td className="p-3 text-slate-900 font-semibold">{rec.requestedBy}</td>
                              <td className="p-3">
                                <span
                                  className={`px-2 py-0.5 rounded-full text-[10px] font-black border whitespace-nowrap ${
                                    approved
                                      ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                                      : 'bg-amber-50 text-amber-800 border-amber-300'
                                  }`}
                                >
                                  {rec.status}
                                </span>
                              </td>
                              <td className="p-3 whitespace-nowrap">
                                {isPromptActive ? (
                                  <div className="inline-flex flex-col gap-1.5 p-2 rounded-xl bg-amber-50 border border-amber-300 shadow-xs min-w-[165px]">
                                    <label className="flex items-center gap-1.5 text-[11px] font-bold text-amber-950 cursor-pointer select-none">
                                      <input
                                        type="checkbox"
                                        checked={isWetPdfConfirmedChecked}
                                        onChange={e => setIsWetPdfConfirmedChecked(e.target.checked)}
                                        className="w-3.5 h-3.5 text-emerald-600 rounded border-amber-400 cursor-pointer"
                                      />
                                      <span>PDF YÜKLENDİ Mİ?</span>
                                    </label>
                                    {displayDocTs && (
                                      <div className="text-[10px] font-mono font-bold text-blue-700">
                                        {displayDocTs}
                                      </div>
                                    )}
                                    <div className="flex items-center gap-1.5">
                                      <button
                                        type="button"
                                        disabled={!isWetPdfConfirmedChecked || isUpdatingBelge}
                                        onClick={() => handleConfirmWetSignedUpload(rec)}
                                        className="flex-1 py-1 px-2.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 text-white font-black text-[10px] font-mono cursor-pointer"
                                      >
                                        TAMAM
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => setPromptBelgeTalepNo(null)}
                                        className="py-1 px-2 rounded-lg bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-[10px] cursor-pointer"
                                      >
                                        İptal
                                      </button>
                                    </div>
                                  </div>
                                ) : rec.wetSignedDocUploaded ? (
                                  <div className="flex flex-col items-start gap-1">
                                    <div className="flex items-center gap-1">
                                      <button
                                        type="button"
                                        onClick={() => window.open(viewUrl, '_blank', 'noopener,noreferrer')}
                                        title={`Islak İmzalı Belgeyi Görüntüle (${viewUrl})`}
                                        className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-bold text-[11px] cursor-pointer shadow-xs"
                                      >
                                        <FileCheck className="w-3.5 h-3.5 text-blue-100" />
                                        <span>BELGE GÖR</span>
                                        <ExternalLink className="w-3 h-3 text-white/80" />
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => handleClickBelgeYukle(rec)}
                                        title="Islak İmzalı Belgeyi Yeniden Yükle"
                                        className="p-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-900 cursor-pointer"
                                      >
                                        <Edit3 className="w-3.5 h-3.5" />
                                      </button>
                                    </div>
                                    {displayDocTs && (
                                      <span className="text-[10.5px] font-mono font-bold text-blue-700">
                                        {displayDocTs}
                                      </span>
                                    )}
                                  </div>
                                ) : (
                                  <div className="flex flex-col items-start gap-1">
                                    <button
                                      type="button"
                                      onClick={() => handleClickBelgeYukle(rec)}
                                      title="Islak İmzalı PDF Yükle (https://bulut.ogm.gov.tr/SARTNAMELER)"
                                      className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-amber-500 hover:bg-amber-600 text-white font-bold text-[11px] cursor-pointer shadow-xs"
                                    >
                                      <Upload className="w-3.5 h-3.5" />
                                      <span>BELGE YÜKLE</span>
                                    </button>
                                    {rec.wetSignedClickedAt && (
                                      <span className="text-[10px] font-mono font-bold text-blue-700">
                                        {rec.wetSignedClickedAt}
                                      </span>
                                    )}
                                  </div>
                                )}
                              </td>
                              <td className="p-3 text-right whitespace-nowrap">
                                {rec.formData && (
                                  <div className="inline-flex items-center gap-1.5">
                                    <button
                                      type="button"
                                      onClick={() => triggerDownloadWord(rec.formType, rec.formData, rec.talepNo)}
                                      className="px-2.5 py-1 rounded-lg bg-blue-50 hover:bg-blue-600 text-blue-700 hover:text-white border border-blue-200 text-[10px] font-bold cursor-pointer"
                                    >
                                      Word İndir
                                    </button>
                                    {!approved && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const defaults = buildDefaultFormState();
                                          const fd = rec.formData || {};
                                          setFormState({
                                            ...defaults,
                                            ...fd,
                                            arizaGorselleri: normalizeImageList(fd.arizaGorselleri, fd.arizaGorsel1, fd.arizaGorsel2),
                                            katalogGorselleri: normalizeImageList(fd.katalogGorselleri, fd.katalogGorseli),
                                            cizimGorselleri: normalizeImageList(fd.cizimGorselleri, fd.cizimGorseli),
                                            ornekUrunGorselleri: normalizeImageList(fd.ornekUrunGorselleri, fd.ornekUrunGorseli)
                                          });
                                          setSelectedType(rec.formType);
                                          setActiveMainTab('create');
                                        }}
                                        className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-200 text-[10px] font-bold cursor-pointer"
                                      >
                                        Aç / Düzenle
                                      </button>
                                    )}
                                  </div>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })()
        ) : !selectedType ? (
          /* STEP 1: CHECKBOX / TICK LIST OF 7 REQUEST TYPES */
          <div className="flex-1 overflow-y-auto p-6 bg-slate-50">
            <div className="max-w-4xl mx-auto space-y-5">
              <div className="text-center space-y-1">
                <h3 className="text-base sm:text-lg font-black text-slate-900 uppercase tracking-wide">
                  TALEP TÜRÜNÜ SEÇİNİZ (7 FORM TİPİ)
                </h3>
                <p className="text-xs text-slate-600">
                  Oluşturmak istediğiniz talep türünü işaretleyin. Form üzerindeki hücrelere yazdıkça ilgili depo, kuyruk no, özel alet, kara aracı ve personel bilgileri otomatik tamamlanacaktır.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                {REQUEST_TYPES.map(opt => (
                  <button
                    key={opt.id}
                    type="button"
                    onClick={() => {
                      setFormState(buildDefaultFormState());
                      setSelectedType(opt.id);
                      setHelperQuery('');
                      setIsHelperDropdownOpen(false);
                      setActiveCellAutocomplete(null);
                    }}
                    className="group p-4 rounded-2xl bg-white hover:bg-emerald-50/40 border border-slate-200 hover:border-emerald-500 text-left flex items-start gap-3.5 transition-all cursor-pointer shadow-xs hover:shadow-md"
                  >
                    <div className="mt-0.5 w-6 h-6 rounded-md border-2 border-slate-300 group-hover:border-emerald-600 group-hover:bg-emerald-50 flex items-center justify-center shrink-0 transition-colors">
                      <Check className="w-3.5 h-3.5 text-emerald-600 opacity-0 group-hover:opacity-100 transition-opacity" />
                    </div>
                    <div className={`w-10 h-10 rounded-xl ${opt.badgeColor} flex items-center justify-center shrink-0 shadow-xs`}>
                      {opt.icon}
                    </div>
                    <div className="min-w-0 flex-1">
                      <div className="text-xs font-black text-slate-900 group-hover:text-emerald-800 uppercase leading-snug">
                        {opt.num}. {opt.label}
                      </div>
                      <div className="text-[11px] text-slate-600 mt-1 leading-relaxed">{opt.desc}</div>
                    </div>
                  </button>
                ))}
              </div>
            </div>
          </div>
        ) : (
          /* STEP 2: AUTOCOMPLETE HELPER BAR + LIVE EDITABLE WORD PREVIEW CANVAS */
          <div className="flex-1 flex flex-col overflow-hidden bg-white">
            {/* Smart Auto-fill & Helper Toolbar */}
            <div className="px-5 py-3 bg-white border-b border-slate-200 flex flex-wrap items-center gap-3 shrink-0">
              {(selectedType === 'parca_onarim' || selectedType === 'parca_talep' || selectedType === 'hizmet_talep') && (
                <div className="flex items-center gap-2">
                  <span className="text-[10px] font-black text-slate-600 uppercase font-mono">Birim:</span>
                  <select
                    value={formState.selectedUnit}
                    onChange={e => {
                      const u = e.target.value;
                      const lbl = u ? (AIRCRAFT_UNITS.find(a => a.key === u)?.label || u.toUpperCase()) : '';
                      setFormState(prev => ({
                        ...prev,
                        selectedUnit: u,
                        havaAraciTipi: lbl,
                        selectedTypes: {
                          at802: u === 'at802',
                          t70: u === 't70',
                          bell429: u === 'bell429',
                          b360: u === 'b360'
                        }
                      }));
                    }}
                    className="px-2.5 py-1.5 rounded-xl bg-slate-50 border border-slate-300 text-xs font-bold text-slate-900"
                  >
                    <option value="">Seçiniz...</option>
                    {AIRCRAFT_UNITS.map(u => (
                      <option key={u.key} value={u.key}>
                        {u.label}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {selectedType === 'ozel_alet_kalibrasyon' && (() => {
                const checkedUnits = (['at802', 't70', 'bell429', 'b360'] as const).filter(
                  k => (formState.selectedTypes as any)[k]
                );
                const alertCount = checkedUnits.length > 0 ? scanExpiringOzelAletForUnits(checkedUnits).length : 0;
                return (
                  <button
                    type="button"
                    onClick={() => {
                      const targets = checkedUnits.length > 0 ? checkedUnits : (['at802', 't70', 'bell429', 'b360'] as const);
                      openKalibrasyonAlertForUnits([...targets], false);
                    }}
                    className={`px-3 py-1.5 rounded-xl text-xs font-black flex items-center gap-1.5 cursor-pointer border transition-all shrink-0 ${
                      alertCount > 0
                        ? 'bg-rose-50 hover:bg-rose-100 text-rose-700 border-rose-300 shadow-xs'
                        : 'bg-amber-50 hover:bg-amber-100 text-amber-800 border-amber-300'
                    }`}
                  >
                    <AlertTriangle className="w-3.5 h-3.5" />
                    <span>Gün Takip İkazlı Ürünleri Seç</span>
                    {alertCount > 0 && (
                      <span className="px-1.5 py-0.5 rounded-full bg-rose-600 text-white text-[10px] font-black leading-none">
                        {alertCount}
                      </span>
                    )}
                  </button>
                );
              })()}

              {selectedType !== 'ozel_alet_alim_imal' && (
                <div className="relative flex-1 min-w-[260px] max-w-xl">
                  <Search className="w-4 h-4 text-emerald-600 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={helperQuery}
                    onFocus={() => setIsHelperDropdownOpen(true)}
                    onChange={e => {
                      setHelperQuery(e.target.value);
                      setIsHelperDropdownOpen(true);
                    }}
                    placeholder={
                      selectedType === 'arac_techizat_hizmet'
                        ? 'Kara Araçlarından Seç (Plaka, Ad, Marka/Model ara)...'
                        : selectedType === 'ozel_alet_kalibrasyon'
                          ? 'Özel Alet / Teçhizat Seç (Çoklu eklemek için tıklayın)...'
                          : selectedType === 'bakim_sarf'
                            ? 'Kimyasal / Madeni Yağ Deposundan Ara ve Tabloya Ekle...'
                            : 'Depo Envanterinden Parça Ara (Ad veya P/N ile otomatik doldur)...'
                    }
                    className="w-full pl-9 pr-3 py-1.5 rounded-xl bg-slate-50 border border-emerald-500/40 text-xs text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-emerald-600 focus:bg-white"
                  />
                  {isHelperDropdownOpen && (
                    <div className="absolute left-0 right-0 top-full mt-1 bg-white border border-emerald-500/50 rounded-xl shadow-2xl max-h-60 overflow-y-auto z-50 divide-y divide-slate-100">
                      {isLoadingDepots && selectedType !== 'ozel_alet_kalibrasyon' && selectedType !== 'arac_techizat_hizmet' ? (
                        <div className="p-3 text-xs text-slate-500 text-center">Depo verileri yükleniyor...</div>
                      ) : filteredHelperResults.length === 0 ? (
                        <div className="p-3 text-xs text-slate-500 text-center">Eşleşen kayıt bulunamadı. Forma doğrudan yazabilirsiniz.</div>
                      ) : (
                        filteredHelperResults.map((it: any, idx: number) => (
                          <button
                            key={idx}
                            type="button"
                            onClick={() => handleSelectHelperItem(it)}
                            className="w-full p-2.5 hover:bg-emerald-50 text-left flex items-center justify-between gap-2 cursor-pointer"
                          >
                            <div className="min-w-0">
                              <div className="text-xs font-bold text-slate-900 truncate">
                                {it.name}
                                {it.plate ? ` (${it.plate})` : ''}
                              </div>
                              <div className="text-[10px] font-mono text-slate-500 truncate">
                                {it.brandModel ? `Model: ${it.brandModel}` : `P/N: ${it.pn || '-'}`}
                                {it.sn ? ` • S/N: ${it.sn}` : ''}
                                {it.kalibreBitis ? ` • Kalibre Bitiş: ${it.kalibreBitis}` : ''}
                              </div>
                            </div>
                            <span className="text-[10px] font-black px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200 shrink-0">
                              + Aktar
                            </span>
                          </button>
                        ))
                      )}
                    </div>
                  )}
                </div>
              )}

              {/* Personnel Auto-fill Helper */}
              <div className="flex items-center gap-2 min-w-[230px]">
                <span className="text-[10px] font-black text-slate-600 uppercase font-mono shrink-0">Talep Eden:</span>
                <div className="w-48 text-slate-900">
                  <PersonnelAutocomplete
                    value={formState.talepYapanAdSoyad}
                    onChange={val => setFormState(prev => ({ ...prev, talepYapanAdSoyad: val }))}
                    placeholder="Personel seç / yaz..."
                  />
                </div>
              </div>
            </div>

            {/* LIVE WORD PREVIEW SHEET(S) */}
            <div
              className="flex-1 overflow-y-auto bg-slate-100 p-4 sm:p-6"
              onClick={() => {
                if (isHelperDropdownOpen) setIsHelperDropdownOpen(false);
                if (activeCellAutocomplete) setActiveCellAutocomplete(null);
              }}
              onPaste={async e => {
                const items = e.clipboardData?.items;
                if (!items) return;
                const pastedFiles: File[] = [];
                for (let i = 0; i < items.length; i++) {
                  if (items[i].type.startsWith('image/')) {
                    const file = items[i].getAsFile();
                    if (file) pastedFiles.push(file);
                  }
                }
                if (pastedFiles.length === 0) return;
                const newImgs: FormImageItem[] = [];
                for (const file of pastedFiles) {
                  try {
                    const dataUrl = await optimizeImageFileToDataUrl(file);
                    newImgs.push({
                      id: createImageId(),
                      dataUrl,
                      scale: defaultImageScale,
                      title: 'Yapıştırılan Görsel'
                    });
                  } catch {}
                }
                if (newImgs.length === 0) return;
                if (selectedType === 'parca_talep') {
                  setFormState(prev => ({ ...prev, katalogGorselleri: [...prev.katalogGorselleri, ...newImgs] }));
                } else if (selectedType === 'ozel_alet_alim_imal') {
                  setFormState(prev => ({ ...prev, cizimGorselleri: [...prev.cizimGorselleri, ...newImgs] }));
                } else {
                  setFormState(prev => ({ ...prev, arizaGorselleri: [...prev.arizaGorselleri, ...newImgs] }));
                }
                showNotification(`📋 ${newImgs.length} görsel optimize edilip form ekine yapıştırıldı!`, 'success');
              }}
            >
              <div className="max-w-[1140px] mx-auto space-y-6">
                {/* ==================== PAGE 1 & 2 / 3 RENDERERS BY FORM TYPE ==================== */}
                {(selectedType === 'parca_onarim' ||
                  selectedType === 'hizmet_talep' ||
                  selectedType === 'arac_techizat_hizmet') && (() => {
                  const isVeh = selectedType === 'arac_techizat_hizmet';
                  return (
                    <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 items-start">
                      {/* PAGE 1 */}
                      <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 font-sans text-xs">
                        <div className="border border-slate-600 px-3 py-2 flex items-center justify-between mb-3">
                          <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                          <div className="text-center font-bold text-slate-500 text-xs sm:text-sm uppercase">
                            {currentTypeMeta?.formTitle}
                          </div>
                          <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                        </div>

                        <table className="w-full border-collapse border border-black text-xs">
                          <tbody>
                            <tr className="bg-[#d9d9d9]">
                              <td colSpan={2} className="border border-black px-2 py-1 font-bold">
                                {isVeh ? 'Bölüm 1: Araç/Teçhizat Bilgileri' : 'Bölüm 1: Hava Aracı Bilgileri'}
                              </td>
                            </tr>
                            {isVeh ? (
                              <>
                                <tr>
                                  <td className="border border-black px-2 py-1 w-[38%]">Araç/Teçhizat Tipi:</td>
                                  <td className="border border-black p-0 relative">
                                    <input
                                      type="text"
                                      value={formState.aracTechizatTipi}
                                      onFocus={() => setActiveCellAutocomplete('vehType')}
                                      onChange={e => {
                                        setFormState(p => ({ ...p, aracTechizatTipi: e.target.value }));
                                        setActiveCellAutocomplete('vehType');
                                      }}
                                      placeholder="Kara aracı veya ekipman tipi..."
                                      className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                                    />
                                    {activeCellAutocomplete === 'vehType' && getFilteredVehicleSuggestions(formState.aracTechizatTipi).length > 0 && (
                                      <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                        {getFilteredVehicleSuggestions(formState.aracTechizatTipi).map((v, i) => (
                                          <button
                                            key={i}
                                            type="button"
                                            onMouseDown={e => {
                                              e.preventDefault();
                                              setFormState(p => ({
                                                ...p,
                                                aracTechizatTipi: v.typeDesc || v.name,
                                                aracTescilNo: v.plate && v.plate !== '-' ? v.plate : '-----------',
                                                aracAdiTanimi: v.name,
                                                aracMarkaModel: v.brandModel,
                                                seriNumarasi: v.sn && v.sn !== '-' ? v.sn : '-----------'
                                              }));
                                              setActiveCellAutocomplete(null);
                                            }}
                                            className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                          >
                                            <div className="font-bold text-slate-900">{v.name} {v.plate ? `(${v.plate})` : ''}</div>
                                            <div className="text-[10px] text-slate-500 font-mono">{v.brandModel || v.typeDesc} {v.sn ? `• S/N: ${v.sn}` : ''}</div>
                                          </button>
                                        ))}
                                      </div>
                                    )}
                                  </td>
                                </tr>
                                <tr>
                                  <td className="border border-black px-2 py-1">Varsa Aracın Tescil No:</td>
                                  <td className="border border-black p-0 relative">
                                    <input
                                      type="text"
                                      value={formState.aracTescilNo}
                                      onFocus={() => setActiveCellAutocomplete('vehPlate')}
                                      onChange={e => {
                                        setFormState(p => ({ ...p, aracTescilNo: e.target.value }));
                                        setActiveCellAutocomplete('vehPlate');
                                      }}
                                      placeholder="Opsiyonel (Plaka / Tescil)..."
                                      className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                                    />
                                    {activeCellAutocomplete === 'vehPlate' && getFilteredVehicleSuggestions(formState.aracTescilNo).length > 0 && (
                                      <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                        {getFilteredVehicleSuggestions(formState.aracTescilNo).map((v, i) => (
                                          <button
                                            key={i}
                                            type="button"
                                            onMouseDown={e => {
                                              e.preventDefault();
                                              setFormState(p => ({
                                                ...p,
                                                aracTechizatTipi: v.typeDesc || v.name,
                                                aracTescilNo: v.plate && v.plate !== '-' ? v.plate : '-----------',
                                                aracAdiTanimi: v.name,
                                                aracMarkaModel: v.brandModel,
                                                seriNumarasi: v.sn && v.sn !== '-' ? v.sn : '-----------'
                                              }));
                                              setActiveCellAutocomplete(null);
                                            }}
                                            className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                          >
                                            <div className="font-bold text-slate-900">{v.plate || v.name} — {v.name}</div>
                                            <div className="text-[10px] text-slate-500 font-mono">{v.brandModel} {v.sn ? `• S/N: ${v.sn}` : ''}</div>
                                          </button>
                                        ))}
                                      </div>
                                    )}
                                  </td>
                                </tr>
                              </>
                            ) : (
                              <>
                                <tr>
                                  <td className="border border-black px-2 py-1.5 w-[38%] align-top">
                                    <div>Hava Aracı Tipi:</div>
                                    <div className="font-bold text-emerald-800 mt-0.5">{formState.havaAraciTipi || '-'}</div>
                                  </td>
                                  <td className="border border-black px-2.5 py-1.5 space-y-0.5 text-[11px]">
                                    <div className="grid grid-cols-2 gap-x-2 gap-y-0.5">
                                      {[
                                        { key: 'at802' as const, label: 'AT-802' },
                                        { key: 't70' as const, label: 'T-70' },
                                        { key: 'bell429' as const, label: 'Bell-429' },
                                        { key: 'b360' as const, label: 'B-360' }
                                      ].map(opt => (
                                        <label key={opt.key} className="flex items-center gap-1.5 cursor-pointer select-none">
                                          <input
                                            type="checkbox"
                                            checked={(formState.selectedTypes as any)[opt.key]}
                                            onChange={e => handleToggleAircraftType(opt.key, e.target.checked, true)}
                                            className="accent-blue-600"
                                          />
                                          <span className="font-semibold">{opt.label}</span>
                                        </label>
                                      ))}
                                    </div>
                                  </td>
                                </tr>
                                <tr>
                                  <td className="border border-black px-2 py-1">Hava Aracı Kuyruk Tescil No:</td>
                                  <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                    <input
                                      type="text"
                                      value={formState.kuyrukTescilNo}
                                      onFocus={() => setActiveCellAutocomplete('tailNo')}
                                      onChange={e => {
                                        setFormState(p => ({ ...p, kuyrukTescilNo: e.target.value }));
                                        setActiveCellAutocomplete('tailNo');
                                      }}
                                      placeholder="Kuyruk no seçiniz veya yazınız..."
                                      className="w-full px-2 py-1 outline-none focus:bg-blue-50 font-semibold"
                                    />
                                    {activeCellAutocomplete === 'tailNo' && matchingTailSuggestions.length > 0 && (
                                      <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-blue-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                        {matchingTailSuggestions.map((tail, i) => (
                                          <button
                                            key={i}
                                            type="button"
                                            onMouseDown={e => {
                                              e.preventDefault();
                                              setFormState(p => ({ ...p, kuyrukTescilNo: tail }));
                                              setActiveCellAutocomplete(null);
                                            }}
                                            className="w-full px-2.5 py-1.5 text-left hover:bg-blue-50 text-[11px] font-bold text-slate-900 cursor-pointer"
                                          >
                                            {tail}
                                          </button>
                                        ))}
                                      </div>
                                    )}
                                  </td>
                                </tr>
                              </>
                            )}

                            <tr className="bg-[#d9d9d9]">
                              <td colSpan={2} className="border border-black px-2 py-1 font-bold">
                                Bölüm 2: Arızalı Parça Bilgileri
                              </td>
                            </tr>
                            {isVeh ? (
                              <>
                                <tr>
                                  <td className="border border-black px-2 py-1">Araç/Teçhizat Adı / Tanımı:</td>
                                  <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                    <input
                                      type="text"
                                      value={formState.aracAdiTanimi}
                                      onFocus={() => setActiveCellAutocomplete('vehName')}
                                      onChange={e => {
                                        setFormState(p => ({ ...p, aracAdiTanimi: e.target.value }));
                                        setActiveCellAutocomplete('vehName');
                                      }}
                                      className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                                    />
                                    {activeCellAutocomplete === 'vehName' && getFilteredVehicleSuggestions(formState.aracAdiTanimi).length > 0 && (
                                      <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                        {getFilteredVehicleSuggestions(formState.aracAdiTanimi).map((v, i) => (
                                          <button
                                            key={i}
                                            type="button"
                                            onMouseDown={e => {
                                              e.preventDefault();
                                              setFormState(p => ({
                                                ...p,
                                                aracTechizatTipi: v.typeDesc || v.name,
                                                aracTescilNo: v.plate && v.plate !== '-' ? v.plate : '-----------',
                                                aracAdiTanimi: v.name,
                                                aracMarkaModel: v.brandModel,
                                                seriNumarasi: v.sn && v.sn !== '-' ? v.sn : '-----------'
                                              }));
                                              setActiveCellAutocomplete(null);
                                            }}
                                            className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                          >
                                            <div className="font-bold text-slate-900">{v.name} {v.plate ? `(${v.plate})` : ''}</div>
                                            <div className="text-[10px] text-slate-500 font-mono">{v.brandModel}</div>
                                          </button>
                                        ))}
                                      </div>
                                    )}
                                  </td>
                                </tr>
                                <tr>
                                  <td className="border border-black px-2 py-1">Araç/Teçhizat Marka/Model:</td>
                                  <td className="border border-black p-0">
                                    <input
                                      type="text"
                                      value={formState.aracMarkaModel}
                                      onChange={e => setFormState(p => ({ ...p, aracMarkaModel: e.target.value }))}
                                      className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                                    />
                                  </td>
                                </tr>
                              </>
                            ) : (
                              <>
                                <tr>
                                  <td className="border border-black px-2 py-1">Parça Adı / Tanımı:</td>
                                  <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                    <input
                                      type="text"
                                      value={formState.parcaAdi}
                                      onFocus={() => setActiveCellAutocomplete('partName')}
                                      onChange={e => {
                                        setFormState(p => ({ ...p, parcaAdi: e.target.value }));
                                        setActiveCellAutocomplete('partName');
                                      }}
                                      placeholder="Parça adı yazarak seçebilirsiniz..."
                                      className="w-full px-2 py-1 outline-none focus:bg-blue-50 font-semibold"
                                    />
                                    {activeCellAutocomplete === 'partName' && getFilteredDepotSuggestions(formState.parcaAdi).length > 0 && (
                                      <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                        {getFilteredDepotSuggestions(formState.parcaAdi).map((d, i) => (
                                          <button
                                            key={i}
                                            type="button"
                                            onMouseDown={e => {
                                              e.preventDefault();
                                              setFormState(p => ({
                                                ...p,
                                                parcaAdi: d.name,
                                                parcaNumarasi: d.pn === '-' ? '' : d.pn,
                                                seriNumarasi: d.sn === '-' ? '' : d.sn
                                              }));
                                              setActiveCellAutocomplete(null);
                                            }}
                                            className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                          >
                                            <div className="font-bold text-slate-900">{d.name}</div>
                                            <div className="text-[10px] text-slate-500 font-mono">
                                              P/N: {d.pn || '-'} {d.sn ? `• S/N: ${d.sn}` : ''} ({d.unitLabel})
                                            </div>
                                          </button>
                                        ))}
                                      </div>
                                    )}
                                  </td>
                                </tr>
                                <tr>
                                  <td className="border border-black px-2 py-1">Parça Numarası:</td>
                                  <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                    <input
                                      type="text"
                                      value={formState.parcaNumarasi}
                                      onFocus={() => setActiveCellAutocomplete('partPn')}
                                      onChange={e => {
                                        setFormState(p => ({ ...p, parcaNumarasi: e.target.value }));
                                        setActiveCellAutocomplete('partPn');
                                      }}
                                      placeholder="P/N yazarak seçebilirsiniz..."
                                      className="w-full px-2 py-1 outline-none focus:bg-blue-50 font-mono"
                                    />
                                    {activeCellAutocomplete === 'partPn' && getFilteredDepotSuggestions(formState.parcaNumarasi).length > 0 && (
                                      <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                        {getFilteredDepotSuggestions(formState.parcaNumarasi).map((d, i) => (
                                          <button
                                            key={i}
                                            type="button"
                                            onMouseDown={e => {
                                              e.preventDefault();
                                              setFormState(p => ({
                                                ...p,
                                                parcaAdi: d.name,
                                                parcaNumarasi: d.pn === '-' ? '' : d.pn,
                                                seriNumarasi: d.sn === '-' ? '' : d.sn
                                              }));
                                              setActiveCellAutocomplete(null);
                                            }}
                                            className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                          >
                                            <div className="font-bold text-slate-900 font-mono">{d.pn} — <span className="font-sans">{d.name}</span></div>
                                            <div className="text-[10px] text-slate-500 font-mono">{d.sn ? `S/N: ${d.sn} • ` : ''}{d.unitLabel}</div>
                                          </button>
                                        ))}
                                      </div>
                                    )}
                                  </td>
                                </tr>
                              </>
                            )}
                            <tr>
                              <td className="border border-black px-2 py-1">(Varsa) Seri Numarası:</td>
                              <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                <input
                                  type="text"
                                  value={formState.seriNumarasi}
                                  onFocus={() => setActiveCellAutocomplete('partSn')}
                                  onChange={e => {
                                    setFormState(p => ({ ...p, seriNumarasi: e.target.value }));
                                    setActiveCellAutocomplete('partSn');
                                  }}
                                  className="w-full px-2 py-1 outline-none focus:bg-blue-50 font-mono"
                                />
                                {!isVeh && activeCellAutocomplete === 'partSn' && getFilteredDepotSuggestions(formState.seriNumarasi).length > 0 && (
                                  <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                    {getFilteredDepotSuggestions(formState.seriNumarasi).map((d, i) => (
                                      <button
                                        key={i}
                                        type="button"
                                        onMouseDown={e => {
                                          e.preventDefault();
                                          setFormState(p => ({
                                            ...p,
                                            parcaAdi: d.name,
                                            parcaNumarasi: d.pn === '-' ? '' : d.pn,
                                            seriNumarasi: d.sn === '-' ? '' : d.sn
                                          }));
                                          setActiveCellAutocomplete(null);
                                        }}
                                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                      >
                                        <div className="font-bold text-slate-900">{d.name}</div>
                                        <div className="text-[10px] text-slate-500 font-mono">P/N: {d.pn || '-'} {d.sn ? `• S/N: ${d.sn}` : ''}</div>
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </td>
                            </tr>

                            <tr className="bg-[#d9d9d9]">
                              <td colSpan={2} className="border border-black px-2 py-1 font-bold">
                                Bölüm 3: Arıza Tanımı Ve Detayları
                              </td>
                            </tr>
                            <tr>
                              <td colSpan={2} className="border border-black p-2">
                                <div className="text-[11px] text-slate-600 mb-1">
                                  Arızanın Kısa Tanımı: (Örn: Aşınma, Çatlak, Fonksiyon Bozukluğu, Korozyon Vs.)
                                </div>
                                <input
                                  type="text"
                                  value={formState.arizaKisaTanimi}
                                  onChange={e => setFormState(p => ({ ...p, arizaKisaTanimi: e.target.value }))}
                                  placeholder="Kısa arıza tanımı giriniz..."
                                  className="w-full px-2 py-1 bg-slate-50 border border-slate-200 rounded outline-none focus:border-blue-500 font-semibold"
                                />
                              </td>
                            </tr>
                            <tr>
                              <td colSpan={2} className="border border-black p-2">
                                <div className="text-[11px] text-slate-600 mb-1">
                                  Arızanın Detaylı Açıklaması: (Arızanın Nasıl Fark Edildiği, Etkilediği Sistemler, Gözlemler Vb.)
                                </div>
                                <textarea
                                  rows={5}
                                  value={formState.arizaDetayliAciklama}
                                  onChange={e => setFormState(p => ({ ...p, arizaDetayliAciklama: e.target.value }))}
                                  placeholder="Detaylı arıza açıklaması..."
                                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded outline-none focus:border-blue-500 resize-none"
                                />
                                <div className="mt-2 text-center italic font-semibold text-[11px] text-slate-700">
                                  (Arızalı Parçanın Katalogdan Alınmış Fotoğrafı Ve Güncel Durum Fotoğrafı Ekine Eklenmelidir.)
                                </div>
                                <FormImageGallery
                                  images={formState.arizaGorselleri}
                                  onChange={next => setFormState(p => ({ ...p, arizaGorselleri: next }))}
                                  defaultScale={defaultImageScale}
                                  onDefaultScaleChange={setDefaultImageScale}
                                  primaryLabel="Katalog Fotoğrafı Ekle"
                                  secondaryLabel="Güncel Durum Fotoğrafı Ekle"
                                  onNotify={showNotification}
                                />
                              </td>
                            </tr>

                            <tr className="bg-[#d9d9d9]">
                              <td colSpan={2} className="border border-black px-2 py-1 font-bold">
                                Bölüm 4: Talep Edilen İşlem
                              </td>
                            </tr>
                            <tr>
                              <td colSpan={2} className="border border-black px-2 py-1 text-[11px]">
                                Lütfen Talep Edilen İşlemi İşaretleyiniz:
                              </td>
                            </tr>
                            <tr>
                              <td className="border border-black px-2 py-2 align-top">Talep Edilen İşlem</td>
                              <td className="border border-black px-3 py-1.5 space-y-0.5 text-[11px]">
                                {[
                                  { k: 'onarim', l: 'Onarım' },
                                  { k: 'revizyon', l: 'Revizyon' },
                                  { k: 'kaplama', l: 'Kaplama' },
                                  { k: 'boyama', l: 'Boyama' },
                                  { k: 'kaynak', l: 'Kaynak' },
                                  { k: 'ndt', l: 'Tahribatsız Muayene (Ndt)' }
                                ].map(item => (
                                  <label key={item.k} className="flex items-center gap-1.5 cursor-pointer">
                                    <input
                                      type="checkbox"
                                      checked={(formState.islemler as any)[item.k]}
                                      onChange={e =>
                                        setFormState(p => ({
                                          ...p,
                                          islemler: { ...p.islemler, [item.k]: e.target.checked }
                                        }))
                                      }
                                      className="accent-emerald-700"
                                    />
                                    <span>{item.l}</span>
                                  </label>
                                ))}
                              </td>
                            </tr>
                            <tr>
                              <td className="border border-black px-2 py-1">Diğer (Lütfen Belirtiniz):</td>
                              <td className="border border-black p-0">
                                <input
                                  type="text"
                                  value={formState.islemDiger}
                                  onChange={e => setFormState(p => ({ ...p, islemDiger: e.target.value }))}
                                  className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                                />
                              </td>
                            </tr>

                            <tr className="bg-[#d9d9d9]">
                              <td colSpan={2} className="border border-black px-2 py-1 font-bold">
                                <div className="flex items-center gap-4">
                                  <span>{isVeh ? 'Bölüm 6: Varsa Araç/Teçhizat Arıza Geçmişi:' : 'Bölüm 6: Varsa Parçanın Arıza Geçmişi:'}</span>
                                  <label className="inline-flex items-center gap-1 cursor-pointer">
                                    <input
                                      type="checkbox"
                                      checked={formState.arizaGecmisi === 'Var'}
                                      onChange={() => setFormState(p => ({ ...p, arizaGecmisi: p.arizaGecmisi === 'Var' ? '' : 'Var' }))}
                                    />
                                    <span>Var</span>
                                  </label>
                                  <label className="inline-flex items-center gap-1 cursor-pointer">
                                    <input
                                      type="checkbox"
                                      checked={formState.arizaGecmisi === 'Yok'}
                                      onChange={() => setFormState(p => ({ ...p, arizaGecmisi: p.arizaGecmisi === 'Yok' ? '' : 'Yok' }))}
                                    />
                                    <span>Yok</span>
                                  </label>
                                </div>
                              </td>
                            </tr>
                            <tr>
                              <td className="border border-black px-2 py-1">Varsa Önceki Arıza Tarihi:</td>
                              <td className="border border-black p-0">
                                <input
                                  type="text"
                                  value={formState.oncekiArizaTarihi}
                                  onChange={e => setFormState(p => ({ ...p, oncekiArizaTarihi: e.target.value }))}
                                  placeholder="GG.AA.YYYY"
                                  className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                                />
                              </td>
                            </tr>
                            <tr>
                              <td className="border border-black px-2 py-1">Varsa Önceki Arıza Nedeni:</td>
                              <td className="border border-black p-0">
                                <input
                                  type="text"
                                  value={formState.oncekiArizaNedeni}
                                  onChange={e => setFormState(p => ({ ...p, oncekiArizaNedeni: e.target.value }))}
                                  className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                                />
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>

                      {/* PAGE 2 */}
                      <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 font-sans text-xs min-h-[700px] flex flex-col justify-between">
                        <div>
                          <div className="border border-slate-600 px-3 py-2 flex items-center justify-between mb-3">
                            <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                            <div className="text-center font-bold text-slate-500 text-xs sm:text-sm uppercase">
                              {currentTypeMeta?.formTitle}
                            </div>
                            <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                          </div>

                          <table className="w-full border-collapse border border-black text-xs">
                            <tbody>
                              <tr className="bg-[#d9d9d9]">
                                <td colSpan={2} className="border border-black px-2 py-1 font-bold">
                                  Bölüm 7: Ek Notlar / Özel İstekler
                                </td>
                              </tr>
                              <tr>
                                <td colSpan={2} className="border border-black p-2">
                                  <div className="text-[11px] text-slate-600 mb-1">
                                    Öncelik, Teslim Süresi, Özel Talimatlar Veya Dikkat Edilmesi Gereken Diğer Hususlar.
                                  </div>
                                  <textarea
                                    rows={6}
                                    value={formState.ekNotlar}
                                    onChange={e => setFormState(p => ({ ...p, ekNotlar: e.target.value }))}
                                    placeholder="Ek notlarınızı yazabilirsiniz..."
                                    className="w-full p-2 bg-slate-50 border border-slate-200 rounded outline-none focus:border-blue-500 resize-none"
                                  />
                                </td>
                              </tr>
                              <tr>
                                <td className="border border-black px-2 py-2 w-[38%]">Talep Yapan Personel Adı Soyadı:</td>
                                <td className="border border-black p-1 relative" onClick={e => e.stopPropagation()}>
                                  <input
                                    type="text"
                                    value={formState.talepYapanAdSoyad}
                                    onFocus={() => setActiveCellAutocomplete('personnel')}
                                    onChange={e => {
                                      setFormState(p => ({ ...p, talepYapanAdSoyad: e.target.value }));
                                      setActiveCellAutocomplete('personnel');
                                    }}
                                    placeholder="Personel Adı Soyadı yazınız..."
                                    className="w-full px-2 py-1 font-bold outline-none focus:bg-blue-50"
                                  />
                                  {activeCellAutocomplete === 'personnel' && getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).length > 0 && (
                                    <div className="absolute left-0 right-0 bottom-full mb-1 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                      {getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).map((person, i) => (
                                        <button
                                          key={i}
                                          type="button"
                                          onMouseDown={e => {
                                            e.preventDefault();
                                            setFormState(p => ({ ...p, talepYapanAdSoyad: person }));
                                            setActiveCellAutocomplete(null);
                                          }}
                                          className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-xs font-bold text-slate-900 cursor-pointer"
                                        >
                                          {person}
                                        </button>
                                      ))}
                                    </div>
                                  )}
                                </td>
                              </tr>
                              <tr>
                                <td className="border border-black px-2 py-4 w-[38%]">Talep Sahibi İmza:</td>
                                <td className="border border-black px-2 py-4"></td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                        <div className="text-[10px] text-slate-400 text-center pt-4">Sayfa 2 / 2</div>
                      </div>
                    </div>
                  );
                })()}

                {/* TYPE 2: BAKIM SARF / MADENİ YAĞ, GAZ VE KİMYASAL */}
                {selectedType === 'bakim_sarf' && (
                  <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 max-w-[780px] mx-auto font-sans text-xs">
                    <div className="border border-slate-600 px-3 py-2 flex items-center justify-between mb-3">
                      <img src={circularOgmLogoDataUri} alt="OGM" className="w-12 h-12 rounded-full object-cover shrink-0" />
                      <div className="text-center font-bold text-slate-500 text-xs sm:text-sm uppercase px-2">
                        {currentTypeMeta?.formTitle}
                      </div>
                      <img src={circularOgmLogoDataUri} alt="OGM" className="w-12 h-12 rounded-full object-cover shrink-0" />
                    </div>

                    <table className="w-full border-collapse border border-black text-xs">
                      <tbody>
                        <tr className="bg-[#d9d9d9]">
                          <td colSpan={6} className="border border-black px-2 py-1 font-bold">
                            Bölüm 1: Hava Aracı Birimi
                          </td>
                        </tr>
                        <tr>
                          <td colSpan={2} className="border border-black px-2 py-2 align-top">
                            Hava Aracı Tipi:
                          </td>
                          <td colSpan={4} className="border border-black px-3 py-1.5 space-y-0.5 text-[11px]">
                            {[
                              { key: 'at802' as const, label: 'AT-802' },
                              { key: 't70' as const, label: 'T-70' },
                              { key: 'bell429' as const, label: 'Bell-429' },
                              { key: 'b360' as const, label: 'B-360' }
                            ].map(opt => (
                              <label key={opt.key} className="flex items-center gap-1.5 cursor-pointer">
                                <input
                                  type="checkbox"
                                  checked={(formState.selectedTypes as any)[opt.key]}
                                  onChange={e => handleToggleAircraftType(opt.key, e.target.checked, false)}
                                />
                                <span>{opt.label}</span>
                              </label>
                            ))}
                          </td>
                        </tr>
                        <tr className="bg-[#d9d9d9]">
                          <td colSpan={6} className="border border-black px-2 py-1 font-bold">
                            Bölüm 2: Talep Edilen Malzemeler
                          </td>
                        </tr>
                        <tr className="bg-[#f3f4f6] text-[10.5px]">
                          <th className="border border-black px-1 py-1 w-[7%]">Sıra No</th>
                          <th className="border border-black px-1.5 py-1 w-[23%]">Malzeme Kodu / Parça Numarası</th>
                          <th className="border border-black px-1.5 py-1 w-[24%]">Malzeme Tanımı</th>
                          <th className="border border-black px-1 py-1 w-[18%]">Birim (Adet, Lt, Kg, Cc)</th>
                          <th className="border border-black px-1 py-1 w-[10%]">Miktar</th>
                          <th className="border border-black px-1.5 py-1 w-[18%]">Açıklama</th>
                        </tr>
                        {formState.kimyasalRows.map((r, rIdx) => (
                          <tr key={r.siraNo} className="hover:bg-amber-50/40">
                            <td className="border border-black text-center font-bold px-1">{r.siraNo}</td>
                            <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                              <input
                                type="text"
                                value={r.malzemeKodu}
                                onFocus={() => setActiveCellAutocomplete(`kim_code_${rIdx}`)}
                                onChange={e => {
                                  const v = e.target.value;
                                  setFormState(p => ({
                                    ...p,
                                    kimyasalRows: p.kimyasalRows.map((it, i) => (i === rIdx ? { ...it, malzemeKodu: v } : it))
                                  }));
                                  setActiveCellAutocomplete(`kim_code_${rIdx}`);
                                }}
                                className="w-full px-1.5 py-0.5 font-mono text-[11px] outline-none focus:bg-blue-50"
                              />
                              {activeCellAutocomplete === `kim_code_${rIdx}` && getFilteredDepotSuggestions(r.malzemeKodu, 'kimyasal').length > 0 && (
                                <div className="absolute left-0 top-full mt-0.5 w-72 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-44 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                  {getFilteredDepotSuggestions(r.malzemeKodu, 'kimyasal').map((d, idx) => (
                                    <button
                                      key={idx}
                                      type="button"
                                      onMouseDown={e => {
                                        e.preventDefault();
                                        setFormState(p => ({
                                          ...p,
                                          kimyasalRows: p.kimyasalRows.map((it, i) =>
                                            i === rIdx
                                              ? {
                                                  ...it,
                                                  malzemeKodu: d.pn === '-' ? '' : d.pn,
                                                  malzemeTanimi: d.name,
                                                  birim: d.unitMeasure || it.birim || 'Adet',
                                                  miktar: it.miktar || '1',
                                                  aciklama: it.aciklama || d.unitLabel
                                                }
                                              : it
                                          )
                                        }));
                                        setActiveCellAutocomplete(null);
                                      }}
                                      className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                    >
                                      <div className="font-bold text-slate-900">{d.name}</div>
                                      <div className="text-[10px] font-mono text-slate-500">P/N: {d.pn || '-'} • {d.unitLabel}</div>
                                    </button>
                                  ))}
                                </div>
                              )}
                            </td>
                            <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                              <input
                                type="text"
                                value={r.malzemeTanimi}
                                onFocus={() => setActiveCellAutocomplete(`kim_name_${rIdx}`)}
                                onChange={e => {
                                  const v = e.target.value;
                                  setFormState(p => ({
                                    ...p,
                                    kimyasalRows: p.kimyasalRows.map((it, i) => (i === rIdx ? { ...it, malzemeTanimi: v } : it))
                                  }));
                                  setActiveCellAutocomplete(`kim_name_${rIdx}`);
                                }}
                                className="w-full px-1.5 py-0.5 text-[11px] outline-none focus:bg-blue-50"
                              />
                              {activeCellAutocomplete === `kim_name_${rIdx}` && getFilteredDepotSuggestions(r.malzemeTanimi, 'kimyasal').length > 0 && (
                                <div className="absolute left-0 top-full mt-0.5 w-72 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-44 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                  {getFilteredDepotSuggestions(r.malzemeTanimi, 'kimyasal').map((d, idx) => (
                                    <button
                                      key={idx}
                                      type="button"
                                      onMouseDown={e => {
                                        e.preventDefault();
                                        setFormState(p => ({
                                          ...p,
                                          kimyasalRows: p.kimyasalRows.map((it, i) =>
                                            i === rIdx
                                              ? {
                                                  ...it,
                                                  malzemeKodu: d.pn === '-' ? '' : d.pn,
                                                  malzemeTanimi: d.name,
                                                  birim: d.unitMeasure || it.birim || 'Adet',
                                                  miktar: it.miktar || '1',
                                                  aciklama: it.aciklama || d.unitLabel
                                                }
                                              : it
                                          )
                                        }));
                                        setActiveCellAutocomplete(null);
                                      }}
                                      className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                    >
                                      <div className="font-bold text-slate-900">{d.name}</div>
                                      <div className="text-[10px] font-mono text-slate-500">P/N: {d.pn || '-'} • {d.unitLabel}</div>
                                    </button>
                                  ))}
                                </div>
                              )}
                            </td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={r.birim}
                                onChange={e =>
                                  setFormState(p => ({
                                    ...p,
                                    kimyasalRows: p.kimyasalRows.map((it, i) => (i === rIdx ? { ...it, birim: e.target.value } : it))
                                  }))
                                }
                                placeholder="Adet/Lt..."
                                className="w-full px-1.5 py-0.5 text-center text-[11px] outline-none focus:bg-blue-50"
                              />
                            </td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={r.miktar}
                                onChange={e =>
                                  setFormState(p => ({
                                    ...p,
                                    kimyasalRows: p.kimyasalRows.map((it, i) => (i === rIdx ? { ...it, miktar: e.target.value } : it))
                                  }))
                                }
                                className="w-full px-1 py-0.5 text-center text-[11px] outline-none focus:bg-blue-50"
                              />
                            </td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={r.aciklama}
                                onChange={e =>
                                  setFormState(p => ({
                                    ...p,
                                    kimyasalRows: p.kimyasalRows.map((it, i) => (i === rIdx ? { ...it, aciklama: e.target.value } : it))
                                  }))
                                }
                                className="w-full px-1.5 py-0.5 text-[11px] outline-none focus:bg-blue-50"
                              />
                            </td>
                          </tr>
                        ))}
                        <tr>
                          <td colSpan={6} className="border border-black p-1 text-right bg-slate-50">
                            <button
                              type="button"
                              onClick={() =>
                                setFormState(p => ({
                                  ...p,
                                  kimyasalRows: [
                                    ...p.kimyasalRows,
                                    { siraNo: p.kimyasalRows.length + 1, malzemeKodu: '', malzemeTanimi: '', birim: '', miktar: '', aciklama: '' }
                                  ]
                                }))
                              }
                              className="px-2 py-0.5 bg-slate-200 hover:bg-slate-300 rounded text-[10px] font-bold cursor-pointer inline-flex items-center gap-1"
                            >
                              <Plus className="w-3 h-3" /> Satır Ekle
                            </button>
                          </td>
                        </tr>
                        <tr className="bg-[#d9d9d9]">
                          <td colSpan={6} className="border border-black p-2.5 text-[10.5px] italic leading-relaxed">
                            <strong>*** ÖNEMLİ NOT:***</strong>
                            <br />
                            Bu Form, Hava Aracı Bakım Ve Onarım Faaliyetlerinde kullanılan Sarf Malzemelerinin Talep Edilmesi Amacıyla Hazırlanmıştır. Talep Edebileceğiniz Ürün Çeşitliliği; <strong>Motor Yağı Ve Klima Yağı</strong> Gibi Sıvı Malzemelerden, <strong>Gres Ve Hidrolik</strong> Sıvılara, Ayrıca <strong>Nitrojen, Klima Gazı</strong> Gibi Özel Gaz Çeşitlerine Ve Diğer bakım esnasında kullanılacak Parça temizlik mayisi, Alkol okil asidi vs. gibi Kimyasalları Kapsamaktadır.
                          </td>
                        </tr>
                        <tr>
                          <td colSpan={2} className="border border-black px-2 py-2">Talep Yapan Personel Adı Soyadı:</td>
                          <td colSpan={4} className="border border-black p-1 relative" onClick={e => e.stopPropagation()}>
                            <input
                              type="text"
                              value={formState.talepYapanAdSoyad}
                              onFocus={() => setActiveCellAutocomplete('personnel')}
                              onChange={e => {
                                setFormState(p => ({ ...p, talepYapanAdSoyad: e.target.value }));
                                setActiveCellAutocomplete('personnel');
                              }}
                              placeholder="Personel Adı Soyadı yazınız..."
                              className="w-full px-2 py-1 font-bold outline-none focus:bg-blue-50"
                            />
                            {activeCellAutocomplete === 'personnel' && getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).length > 0 && (
                              <div className="absolute left-0 right-0 bottom-full mb-1 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                {getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).map((person, i) => (
                                  <button
                                    key={i}
                                    type="button"
                                    onMouseDown={e => {
                                      e.preventDefault();
                                      setFormState(p => ({ ...p, talepYapanAdSoyad: person }));
                                      setActiveCellAutocomplete(null);
                                    }}
                                    className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-xs font-bold text-slate-900 cursor-pointer"
                                  >
                                    {person}
                                  </button>
                                ))}
                              </div>
                            )}
                          </td>
                        </tr>
                        <tr>
                          <td colSpan={2} className="border border-black px-2 py-4">Talep Sahibi İmza:</td>
                          <td colSpan={4} className="border border-black px-2 py-4"></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}

                {/* TYPE 3: HAVA ARACI PARÇA TALEP (ZERO REPORT FORMAT) */}
                {selectedType === 'parca_talep' && (
                  <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 items-start">
                    <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 font-sans text-xs">
                      <div className="border border-slate-700 px-3 py-2 flex items-center justify-between mb-3">
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                        <div className="text-center font-bold text-slate-500 text-xs sm:text-sm uppercase">
                          HAVA ARACI YEDEK PARÇA TALEP FORMU
                        </div>
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                      </div>

                      <table className="w-full border-collapse border border-black mb-3 text-xs">
                        <thead>
                          <tr className="bg-[#d9d9d9]">
                            <th className="border border-black px-2 py-1 text-left font-bold w-[60%]">1. Hava Aracı Bilgileri</th>
                            <th className="border border-black px-2 py-1 text-left font-bold w-[40%]"></th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td className="border border-black px-2 py-2 align-top">
                              <div className="flex items-center gap-1">
                                <span>1. Hava Aracı Tipi :</span>
                                <input
                                  type="text"
                                  value={formState.havaAraciTipi}
                                  onChange={e => setFormState(p => ({ ...p, havaAraciTipi: e.target.value }))}
                                  className="flex-1 border-b border-dotted border-slate-400 outline-none font-bold px-1"
                                />
                              </div>
                            </td>
                            <td className="border border-black px-3 py-1 text-[11px] space-y-0.5">
                              {[
                                { key: 'at802' as const, label: 'AT-802' },
                                { key: 't70' as const, label: 'T-70' },
                                { key: 'bell429' as const, label: 'Bell-429' },
                                { key: 'b360' as const, label: 'B-360' }
                              ].map(opt => (
                                <label key={opt.key} className="flex items-center gap-1.5 cursor-pointer">
                                  <input
                                    type="checkbox"
                                    checked={(formState.selectedTypes as any)[opt.key]}
                                    onChange={e => handleToggleAircraftType(opt.key, e.target.checked, false)}
                                  />
                                  <span>{opt.label}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                        </tbody>
                      </table>

                      <table className="w-full border-collapse border border-black text-[11px]">
                        <thead>
                          <tr className="bg-[#d9d9d9]">
                            <th colSpan={6} className="border border-black px-2 py-1 text-left font-bold">
                              2-Talep Edilen Malzemeler
                            </th>
                          </tr>
                          <tr className="bg-[#f3f4f6]">
                            <th className="border border-black px-1 py-1 w-[7%]">Sıra<br/>No</th>
                            <th className="border border-black px-1 py-1 text-left w-[24%]">Parça Adı/Tanımı</th>
                            <th className="border border-black px-1 py-1 text-left w-[19%]">Parça Numarası</th>
                            <th className="border border-black px-1 py-1 text-left w-[20%]">Alternatif P/N</th>
                            <th className="border border-black px-1 py-1 w-[13%]">Miktar</th>
                            <th className="border border-black px-1 py-1 w-[17%]">Öncelik</th>
                          </tr>
                        </thead>
                        <tbody>
                          {formState.parcaTalepRows.map((r, rIdx) => (
                            <tr key={r.siraNo} className="hover:bg-amber-50/40">
                              <td className="border border-black text-center font-bold px-1">{r.siraNo}</td>
                              <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                <input
                                  type="text"
                                  value={r.parcaAdi}
                                  onFocus={() => setActiveCellAutocomplete(`parca_name_${rIdx}`)}
                                  onChange={e => {
                                    const v = e.target.value;
                                    setFormState(p => ({
                                      ...p,
                                      parcaTalepRows: p.parcaTalepRows.map((it, i) => (i === rIdx ? { ...it, parcaAdi: v } : it))
                                    }));
                                    setActiveCellAutocomplete(`parca_name_${rIdx}`);
                                  }}
                                  className="w-full px-1 py-0.5 text-[10.5px] outline-none focus:bg-blue-50"
                                />
                                {activeCellAutocomplete === `parca_name_${rIdx}` && getFilteredDepotSuggestions(r.parcaAdi).length > 0 && (
                                  <div className="absolute left-0 top-full mt-0.5 w-72 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-44 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                    {getFilteredDepotSuggestions(r.parcaAdi).map((d, idx) => (
                                      <button
                                        key={idx}
                                        type="button"
                                        onMouseDown={e => {
                                          e.preventDefault();
                                          setFormState(p => ({
                                            ...p,
                                            parcaTalepRows: p.parcaTalepRows.map((it, i) =>
                                              i === rIdx
                                                ? {
                                                    ...it,
                                                    parcaAdi: d.name,
                                                    parcaNo: d.pn === '-' ? '' : d.pn,
                                                    istenenMiktar: it.istenenMiktar || '1',
                                                    oncelik: it.oncelik || 'Yüksek'
                                                  }
                                                : it
                                            )
                                          }));
                                          setActiveCellAutocomplete(null);
                                        }}
                                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                      >
                                        <div className="font-bold text-slate-900">{d.name}</div>
                                        <div className="text-[10px] font-mono text-slate-500">P/N: {d.pn || '-'} • {d.unitLabel}</div>
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </td>
                              <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                <input
                                  type="text"
                                  value={r.parcaNo}
                                  onFocus={() => setActiveCellAutocomplete(`parca_pn_${rIdx}`)}
                                  onChange={e => {
                                    const v = e.target.value;
                                    setFormState(p => ({
                                      ...p,
                                      parcaTalepRows: p.parcaTalepRows.map((it, i) => (i === rIdx ? { ...it, parcaNo: v } : it))
                                    }));
                                    setActiveCellAutocomplete(`parca_pn_${rIdx}`);
                                  }}
                                  className="w-full px-1 py-0.5 text-[10.5px] font-mono outline-none focus:bg-blue-50"
                                />
                                {activeCellAutocomplete === `parca_pn_${rIdx}` && getFilteredDepotSuggestions(r.parcaNo).length > 0 && (
                                  <div className="absolute left-0 top-full mt-0.5 w-72 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-44 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                    {getFilteredDepotSuggestions(r.parcaNo).map((d, idx) => (
                                      <button
                                        key={idx}
                                        type="button"
                                        onMouseDown={e => {
                                          e.preventDefault();
                                          setFormState(p => ({
                                            ...p,
                                            parcaTalepRows: p.parcaTalepRows.map((it, i) =>
                                              i === rIdx
                                                ? {
                                                    ...it,
                                                    parcaAdi: d.name,
                                                    parcaNo: d.pn === '-' ? '' : d.pn,
                                                    istenenMiktar: it.istenenMiktar || '1',
                                                    oncelik: it.oncelik || 'Yüksek'
                                                  }
                                                : it
                                            )
                                          }));
                                          setActiveCellAutocomplete(null);
                                        }}
                                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                      >
                                        <div className="font-bold text-slate-900 font-mono">{d.pn} — <span className="font-sans">{d.name}</span></div>
                                        <div className="text-[10px] text-slate-500 font-mono">{d.unitLabel}</div>
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </td>
                              <td className="border border-black p-0">
                                <input
                                  type="text"
                                  value={r.alternatifPn}
                                  onChange={e =>
                                    setFormState(p => ({
                                      ...p,
                                      parcaTalepRows: p.parcaTalepRows.map((it, i) => (i === rIdx ? { ...it, alternatifPn: e.target.value } : it))
                                    }))
                                  }
                                  className="w-full px-1 py-0.5 text-[10.5px] font-mono outline-none focus:bg-blue-50"
                                />
                              </td>
                              <td className="border border-black p-0">
                                <input
                                  type="text"
                                  value={r.istenenMiktar}
                                  onChange={e =>
                                    setFormState(p => ({
                                      ...p,
                                      parcaTalepRows: p.parcaTalepRows.map((it, i) => (i === rIdx ? { ...it, istenenMiktar: e.target.value } : it))
                                    }))
                                  }
                                  className="w-full px-1 py-0.5 text-center text-[10.5px] outline-none focus:bg-blue-50"
                                />
                              </td>
                              <td className="border border-black p-0">
                                <select
                                  value={r.oncelik}
                                  onChange={e =>
                                    setFormState(p => ({
                                      ...p,
                                      parcaTalepRows: p.parcaTalepRows.map((it, i) => (i === rIdx ? { ...it, oncelik: e.target.value as any } : it))
                                    }))
                                  }
                                  className="w-full px-1 py-0.5 text-[10.5px] font-bold bg-transparent outline-none"
                                >
                                  <option value=""></option>
                                  <option value="Yüksek">Yüksek</option>
                                  <option value="Normal">Normal</option>
                                  <option value="Düşük">Düşük</option>
                                </select>
                              </td>
                            </tr>
                          ))}
                          <tr>
                            <td colSpan={6} className="bg-[#d9d9d9] border border-black px-2 py-1 text-center font-bold">
                              Parcanın Öncelik Derecesi kısmına yazılacaklar
                            </td>
                          </tr>
                          <tr>
                            <td colSpan={6} className="border border-black px-2 py-0.5">
                              <span className="text-red-600 font-bold">Yüksek</span> – Kısa Sürede Gerekli (Arıza vb)
                            </td>
                          </tr>
                          <tr>
                            <td colSpan={6} className="border border-black px-2 py-0.5">
                              <span className="text-yellow-600 font-bold">Normal</span> – Planlı Bakım İçin
                            </td>
                          </tr>
                          <tr>
                            <td colSpan={6} className="border border-black px-2 py-0.5">
                              <span className="text-green-600 font-bold">Düşük</span> – Depo Miktarını Tamamlamak İçin
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 font-sans text-xs min-h-[700px] flex flex-col justify-between">
                      <div>
                        <div className="border border-slate-700 px-3 py-2 flex items-center justify-between mb-3">
                          <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                          <div className="text-center font-bold text-slate-500 text-xs sm:text-sm uppercase">
                            HAVA ARACI YEDEK PARÇA TALEP FORMU
                          </div>
                          <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                        </div>

                        <table className="w-full border-collapse border border-black mb-3 text-xs">
                          <tbody>
                            <tr className="bg-[#d9d9d9]">
                              <td className="border border-black px-2 py-1 font-bold">3. Parça Kataloğu Görseli</td>
                            </tr>
                            <tr>
                              <td className="border border-black p-3 min-h-[300px]">
                                <div className="mb-2">
                                  Lütfen istenen parçanın katalogdan alınmış görselini formun ekine ekleyiniz. (Birden fazla dosya yükleyebilir veya Ctrl+V ile yapıştırabilirsiniz)
                                </div>
                                <FormImageGallery
                                  images={formState.katalogGorselleri}
                                  onChange={next => setFormState(p => ({ ...p, katalogGorselleri: next }))}
                                  defaultScale={defaultImageScale}
                                  onDefaultScaleChange={setDefaultImageScale}
                                  primaryLabel="Katalog Görseli Seç (Çoklu)"
                                  onNotify={showNotification}
                                />
                              </td>
                            </tr>
                          </tbody>
                        </table>

                        <table className="w-full border-collapse border border-black text-xs">
                          <tbody>
                            <tr className="bg-[#d9d9d9]">
                              <td colSpan={2} className="border border-black px-2 py-1 font-bold">4. Ek Notlar/Açıklamalar</td>
                            </tr>
                            <tr>
                              <td colSpan={2} className="border border-black p-2">
                                <textarea
                                  rows={5}
                                  value={formState.ekNotlar}
                                  onChange={e => setFormState(p => ({ ...p, ekNotlar: e.target.value }))}
                                  placeholder="Parçanın özel gereksinimlerini belirten ek bilgiler..."
                                  className="w-full p-1.5 bg-slate-50 border border-slate-200 rounded outline-none resize-none"
                                />
                              </td>
                            </tr>
                            <tr>
                              <td className="border border-black px-2 py-2 w-[45%]">Talep Yapan Personel Adı Soyadı:</td>
                              <td className="border border-black p-1 relative" onClick={e => e.stopPropagation()}>
                                <input
                                  type="text"
                                  value={formState.talepYapanAdSoyad}
                                  onFocus={() => setActiveCellAutocomplete('personnel')}
                                  onChange={e => {
                                    setFormState(p => ({ ...p, talepYapanAdSoyad: e.target.value }));
                                    setActiveCellAutocomplete('personnel');
                                  }}
                                  placeholder="Personel Adı Soyadı yazınız..."
                                  className="w-full px-2 py-1 font-bold outline-none focus:bg-blue-50"
                                />
                                {activeCellAutocomplete === 'personnel' && getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).length > 0 && (
                                  <div className="absolute left-0 right-0 bottom-full mb-1 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                    {getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).map((person, i) => (
                                      <button
                                        key={i}
                                        type="button"
                                        onMouseDown={e => {
                                          e.preventDefault();
                                          setFormState(p => ({ ...p, talepYapanAdSoyad: person }));
                                          setActiveCellAutocomplete(null);
                                        }}
                                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-xs font-bold text-slate-900 cursor-pointer"
                                      >
                                        {person}
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </td>
                            </tr>
                            <tr>
                              <td className="border border-black px-2 py-4 w-[45%]">Talep Sahibi İmza:</td>
                              <td className="border border-black px-2 py-4"></td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}

                {/* TYPE 4: ÖZEL ALET KALİBRASYON TALEP FORMU */}
                {selectedType === 'ozel_alet_kalibrasyon' && (
                  <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 items-start">
                    <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 font-sans text-xs">
                      <div className="border border-slate-600 px-3 py-2 flex items-center justify-between mb-3">
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                        <div className="text-center font-bold text-slate-500 text-xs sm:text-sm uppercase">
                          {currentTypeMeta?.formTitle}
                        </div>
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                      </div>

                      <table className="w-full border-collapse border border-black text-xs">
                        <tbody>
                          <tr className="bg-[#d9d9d9]">
                            <td colSpan={2} className="border border-black px-2 py-1 font-bold">Bölüm 1: Hava Aracı Bilgileri</td>
                          </tr>
                          <tr>
                            <td className="border border-black px-2 py-2 align-top w-[44%]">Departman / Bölüm:</td>
                            <td className="border border-black px-3 py-1.5 space-y-0.5 text-[11px]">
                              {[
                                { key: 'at802' as const, label: 'AT-802' },
                                { key: 't70' as const, label: 'T-70' },
                                { key: 'bell429' as const, label: 'Bell-429' },
                                { key: 'b360' as const, label: 'B-360' }
                              ].map(opt => (
                                <label key={opt.key} className="flex items-center gap-1.5 cursor-pointer">
                                  <input
                                    type="checkbox"
                                    checked={(formState.selectedTypes as any)[opt.key]}
                                    onChange={e => handleToggleAircraftType(opt.key, e.target.checked, false)}
                                  />
                                  <span>{opt.label}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr>
                            <td className="border border-black px-2 py-1">Diğer:</td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={formState.departmanDiger}
                                onChange={e => setFormState(p => ({ ...p, departmanDiger: e.target.value }))}
                                className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                              />
                            </td>
                          </tr>
                          <tr className="bg-[#d9d9d9]">
                            <td colSpan={2} className="border border-black px-2 py-1 font-bold">Bölüm 2: Kalibrasyon Gereklilikleri</td>
                          </tr>
                          <tr>
                            <td className="border border-black px-2 py-2 align-top">Kalibrasyon Türü:</td>
                            <td className="border border-black px-3 py-1.5 space-y-0.5 text-[11px]">
                              {[
                                { k: 'ilk', l: 'İlk Kalibrasyon' },
                                { k: 'periyodik', l: 'Periyodik Kalibrasyon' },
                                { k: 'onarimSonrasi', l: 'Onarım Sonrası Kalibrasyon' }
                              ].map(opt => (
                                <label key={opt.k} className="flex items-center gap-1.5 cursor-pointer">
                                  <input
                                    type="checkbox"
                                    checked={(formState.kalibrasyonTuru as any)[opt.k]}
                                    onChange={e =>
                                      setFormState(p => ({
                                        ...p,
                                        kalibrasyonTuru: { ...p.kalibrasyonTuru, [opt.k]: e.target.checked }
                                      }))
                                    }
                                  />
                                  <span>{opt.l}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr>
                            <td className="border border-black px-2 py-1">Diğer (Lütfen Belirtiniz):</td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={formState.kalibrasyonDiger}
                                onChange={e => setFormState(p => ({ ...p, kalibrasyonDiger: e.target.value }))}
                                className="w-full px-2 py-1 outline-none focus:bg-blue-50"
                              />
                            </td>
                          </tr>
                          <tr>
                            <td className="border border-black px-2 py-1">Talep Edilen Son Tarih: [GG/AA/YYYY]</td>
                            <td className="border border-black p-0">
                              <input
                                type="text"
                                value={formState.talepEdilenSonTarih}
                                onChange={e => setFormState(p => ({ ...p, talepEdilenSonTarih: e.target.value }))}
                                placeholder="GG/AA/YYYY"
                                className="w-full px-2 py-1 outline-none focus:bg-blue-50 font-mono"
                              />
                            </td>
                          </tr>
                          <tr>
                            <td className="border border-black px-2 py-2 align-top">Öncelik Seviyesi:</td>
                            <td className="border border-black px-3 py-1.5 space-y-0.5 text-[11px]">
                              {(['Yüksek (Acil)', 'Normal', 'Düşük'] as const).map(lvl => (
                                <label key={lvl} className="flex items-center gap-1.5 cursor-pointer">
                                  <input
                                    type="checkbox"
                                    checked={formState.oncelikSeviyesi === lvl}
                                    onChange={() => setFormState(p => ({ ...p, oncelikSeviyesi: p.oncelikSeviyesi === lvl ? '' : lvl }))}
                                  />
                                  <span>{lvl}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr className="bg-[#d9d9d9]">
                            <td colSpan={2} className="border border-black px-2 py-1 font-bold">
                              Bölüm 3: Ek Notlar / Dikkat Edilmesi Gerekenler
                            </td>
                          </tr>
                          <tr>
                            <td colSpan={2} className="border border-black p-2">
                              <div className="text-[11px] text-slate-600 mb-1">
                                (Kalibrasyon İşlemiyle İlgili Özel Talimatlar, Ölçüm Noktaları, Raporlama Formatı Veya Diğer Önemli Bilgiler.)
                              </div>
                              <textarea
                                rows={8}
                                value={formState.ekNotlar}
                                onChange={e => setFormState(p => ({ ...p, ekNotlar: e.target.value }))}
                                className="w-full p-2 bg-slate-50 border border-slate-200 rounded outline-none resize-none"
                              />
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    <div className="bg-white text-black shadow-2xl border border-slate-400 p-6 font-sans text-xs">
                      <div className="border border-slate-600 px-3 py-2 flex items-center justify-between mb-3">
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                        <div className="text-center font-bold text-slate-500 text-xs sm:text-sm uppercase">
                          {currentTypeMeta?.formTitle}
                        </div>
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-11 h-11 rounded-full object-cover shrink-0" />
                      </div>

                      <table className="w-full border-collapse border border-black text-[10.5px]">
                        <thead>
                          <tr className="bg-[#d9d9d9]">
                            <th colSpan={6} className="border border-black px-2 py-1 text-left font-bold text-xs">
                              Bölüm 4: Kalibrasyon Yapılacak Malzeme Listesi
                            </th>
                          </tr>
                          <tr>
                            <td colSpan={6} className="border border-black px-2 py-1 text-[11px]">
                              Aşağıdaki tabloya kalibrasyon talep edilen tüm malzemelerin bilgilerini giriniz.
                            </td>
                          </tr>
                          <tr className="bg-[#f3f4f6]">
                            <th className="border border-black px-1 py-1 w-[7%]">Sıra No</th>
                            <th className="border border-black px-1 py-1 w-[21%]">Malzeme Adı</th>
                            <th className="border border-black px-1 py-1 w-[16%]">Parça Numarası</th>
                            <th className="border border-black px-1 py-1 w-[16%]">Seri Numarası (S/N)</th>
                            <th className="border border-black px-1 py-1 w-[22%]">Malzemenin Tanımı</th>
                            <th className="border border-black px-1 py-1 w-[18%]">Kalibre Bitiş Tarihi</th>
                          </tr>
                        </thead>
                        <tbody>
                          {formState.kalibrasyonRows.map((r, rIdx) => (
                            <tr key={r.siraNo} className="hover:bg-cyan-50/40">
                              <td className="border border-black text-center font-bold px-1">{r.siraNo}</td>
                              <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                <input
                                  type="text"
                                  value={r.malzemeAdi}
                                  onFocus={() => setActiveCellAutocomplete(`kal_name_${rIdx}`)}
                                  onChange={e => {
                                    const v = e.target.value;
                                    setFormState(p => ({
                                      ...p,
                                      kalibrasyonRows: p.kalibrasyonRows.map((it, i) => (i === rIdx ? { ...it, malzemeAdi: v, tanimi: v } : it))
                                    }));
                                    setActiveCellAutocomplete(`kal_name_${rIdx}`);
                                  }}
                                  className="w-full px-1 py-0.5 outline-none focus:bg-blue-50"
                                />
                                {activeCellAutocomplete === `kal_name_${rIdx}` && getFilteredToolSuggestions(r.malzemeAdi).length > 0 && (
                                  <div className="absolute left-0 top-full mt-0.5 w-72 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-44 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                    {getFilteredToolSuggestions(r.malzemeAdi).map((t, idx) => (
                                      <button
                                        key={idx}
                                        type="button"
                                        onMouseDown={e => {
                                          e.preventDefault();
                                          setFormState(p => ({
                                            ...p,
                                            kalibrasyonRows: p.kalibrasyonRows.map((it, i) =>
                                              i === rIdx
                                                ? {
                                                    ...it,
                                                    malzemeAdi: t.name,
                                                    parcaNo: t.pn === '-' ? '' : t.pn,
                                                    seriNo: t.sn === '-' ? '' : t.sn,
                                                    tanimi: t.name,
                                                    kalibreBitisTarihi: t.kalibreBitis || ''
                                                  }
                                                : it
                                            )
                                          }));
                                          setActiveCellAutocomplete(null);
                                        }}
                                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                      >
                                        <div className="font-bold text-slate-900">{t.name}</div>
                                        <div className="text-[10px] font-mono text-slate-500">P/N: {t.pn || '-'} • S/N: {t.sn || '-'} {t.kalibreBitis ? `• Kalibre: ${t.kalibreBitis}` : ''}</div>
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </td>
                              <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                <input
                                  type="text"
                                  value={r.parcaNo}
                                  onFocus={() => setActiveCellAutocomplete(`kal_pn_${rIdx}`)}
                                  onChange={e => {
                                    const v = e.target.value;
                                    setFormState(p => ({
                                      ...p,
                                      kalibrasyonRows: p.kalibrasyonRows.map((it, i) => (i === rIdx ? { ...it, parcaNo: v } : it))
                                    }));
                                    setActiveCellAutocomplete(`kal_pn_${rIdx}`);
                                  }}
                                  className="w-full px-1 py-0.5 font-mono outline-none focus:bg-blue-50"
                                />
                                {activeCellAutocomplete === `kal_pn_${rIdx}` && getFilteredToolSuggestions(r.parcaNo).length > 0 && (
                                  <div className="absolute left-0 top-full mt-0.5 w-72 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-44 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                    {getFilteredToolSuggestions(r.parcaNo).map((t, idx) => (
                                      <button
                                        key={idx}
                                        type="button"
                                        onMouseDown={e => {
                                          e.preventDefault();
                                          setFormState(p => ({
                                            ...p,
                                            kalibrasyonRows: p.kalibrasyonRows.map((it, i) =>
                                              i === rIdx
                                                ? {
                                                    ...it,
                                                    malzemeAdi: t.name,
                                                    parcaNo: t.pn === '-' ? '' : t.pn,
                                                    seriNo: t.sn === '-' ? '' : t.sn,
                                                    tanimi: t.name,
                                                    kalibreBitisTarihi: t.kalibreBitis || ''
                                                  }
                                                : it
                                            )
                                          }));
                                          setActiveCellAutocomplete(null);
                                        }}
                                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                      >
                                        <div className="font-bold text-slate-900">{t.name}</div>
                                        <div className="text-[10px] font-mono text-slate-500">P/N: {t.pn || '-'} • S/N: {t.sn || '-'}</div>
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </td>
                              <td className="border border-black p-0 relative" onClick={e => e.stopPropagation()}>
                                <input
                                  type="text"
                                  value={r.seriNo}
                                  onFocus={() => setActiveCellAutocomplete(`kal_sn_${rIdx}`)}
                                  onChange={e => {
                                    const v = e.target.value;
                                    setFormState(p => ({
                                      ...p,
                                      kalibrasyonRows: p.kalibrasyonRows.map((it, i) => (i === rIdx ? { ...it, seriNo: v } : it))
                                    }));
                                    setActiveCellAutocomplete(`kal_sn_${rIdx}`);
                                  }}
                                  className="w-full px-1 py-0.5 font-mono outline-none focus:bg-blue-50"
                                />
                                {activeCellAutocomplete === `kal_sn_${rIdx}` && getFilteredToolSuggestions(r.seriNo).length > 0 && (
                                  <div className="absolute left-0 top-full mt-0.5 w-72 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-44 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                    {getFilteredToolSuggestions(r.seriNo).map((t, idx) => (
                                      <button
                                        key={idx}
                                        type="button"
                                        onMouseDown={e => {
                                          e.preventDefault();
                                          setFormState(p => ({
                                            ...p,
                                            kalibrasyonRows: p.kalibrasyonRows.map((it, i) =>
                                              i === rIdx
                                                ? {
                                                    ...it,
                                                    malzemeAdi: t.name,
                                                    parcaNo: t.pn === '-' ? '' : t.pn,
                                                    seriNo: t.sn === '-' ? '' : t.sn,
                                                    tanimi: t.name,
                                                    kalibreBitisTarihi: t.kalibreBitis || ''
                                                  }
                                                : it
                                            )
                                          }));
                                          setActiveCellAutocomplete(null);
                                        }}
                                        className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-[11px] cursor-pointer"
                                      >
                                        <div className="font-bold text-slate-900">{t.name}</div>
                                        <div className="text-[10px] font-mono text-slate-500">S/N: {t.sn || '-'} • Kalibre: {t.kalibreBitis || '-'}</div>
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </td>
                              <td className="border border-black p-0">
                                <input
                                  type="text"
                                  value={r.tanimi}
                                  onChange={e =>
                                    setFormState(p => ({
                                      ...p,
                                      kalibrasyonRows: p.kalibrasyonRows.map((it, i) => (i === rIdx ? { ...it, tanimi: e.target.value } : it))
                                    }))
                                  }
                                  className="w-full px-1 py-0.5 outline-none focus:bg-blue-50"
                                />
                              </td>
                              <td className="border border-black p-0">
                                <input
                                  type="text"
                                  value={r.kalibreBitisTarihi}
                                  onChange={e =>
                                    setFormState(p => ({
                                      ...p,
                                      kalibrasyonRows: p.kalibrasyonRows.map((it, i) => (i === rIdx ? { ...it, kalibreBitisTarihi: e.target.value } : it))
                                    }))
                                  }
                                  className="w-full px-1 py-0.5 text-center outline-none focus:bg-blue-50"
                                />
                              </td>
                            </tr>
                          ))}
                          <tr>
                            <td colSpan={6} className="border border-black px-2 py-1 italic text-[10px]">
                              *Ek satırlara ihtiyaç duyulması halinde bu bölümü çoğaltabilirsiniz.*
                            </td>
                          </tr>
                          <tr>
                            <td colSpan={3} className="border border-black px-2 py-1.5 text-xs">Talep Yapan Personel Adı Soyadı:</td>
                            <td colSpan={3} className="border border-black p-1 relative" onClick={e => e.stopPropagation()}>
                              <input
                                type="text"
                                value={formState.talepYapanAdSoyad}
                                onFocus={() => setActiveCellAutocomplete('personnel')}
                                onChange={e => {
                                  setFormState(p => ({ ...p, talepYapanAdSoyad: e.target.value }));
                                  setActiveCellAutocomplete('personnel');
                                }}
                                placeholder="Personel Adı Soyadı yazınız..."
                                className="w-full px-2 py-0.5 font-bold outline-none focus:bg-blue-50 text-xs"
                              />
                              {activeCellAutocomplete === 'personnel' && getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).length > 0 && (
                                <div className="absolute left-0 right-0 bottom-full mb-1 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                  {getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).map((person, i) => (
                                    <button
                                      key={i}
                                      type="button"
                                      onMouseDown={e => {
                                        e.preventDefault();
                                        setFormState(p => ({ ...p, talepYapanAdSoyad: person }));
                                        setActiveCellAutocomplete(null);
                                      }}
                                      className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-xs font-bold text-slate-900 cursor-pointer"
                                    >
                                      {person}
                                    </button>
                                  ))}
                                </div>
                              )}
                            </td>
                          </tr>
                          <tr>
                            <td colSpan={3} className="border border-black px-2 py-3 text-xs">Talep Sahibi İmza:</td>
                            <td colSpan={3} className="border border-black px-2 py-3"></td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* TYPE 5: ÖZEL ALET ALIM İMAL ÜRETİM (3 PAGES: PAGE 1 & 2 SIDE BY SIDE, PAGE 3 BELOW) */}
                {selectedType === 'ozel_alet_alim_imal' && (
                  <div className="space-y-6">
                    <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 items-start">
                      {/* PAGE 1 */}
                      <div className="bg-white text-black shadow-2xl border border-slate-400 p-5 font-sans text-[11px]">
                        <div className="border border-slate-600 px-2 py-1.5 flex items-center justify-between mb-2">
                          <img src={circularOgmLogoDataUri} alt="OGM" className="w-10 h-10 rounded-full object-cover shrink-0" />
                          <div className="text-center font-bold text-slate-500 text-[10.5px] uppercase px-1">
                            {currentTypeMeta?.formTitle}
                          </div>
                          <img src={circularOgmLogoDataUri} alt="OGM" className="w-10 h-10 rounded-full object-cover shrink-0" />
                        </div>
                        <table className="w-full border-collapse border border-black">
                          <tbody>
                            <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">1. Hava Aracı Bilgileri</td></tr>
                            <tr>
                              <td className="border border-black px-1.5 py-1 w-[44%]">Hava Aracı Tipi:</td>
                              <td className="border border-black px-2 py-1 space-y-0.5">
                                {(['at802', 't70', 'bell429', 'b360'] as const).map(k => (
                                  <label key={k} className="flex items-center gap-1 cursor-pointer">
                                    <input
                                      type="checkbox"
                                      checked={(formState.selectedTypes as any)[k]}
                                      onChange={e => handleToggleAircraftType(k, e.target.checked, false)}
                                    />
                                    <span>{k === 'at802' ? 'AT-802' : k === 't70' ? 'T-70' : k === 'bell429' ? 'Bell-429' : 'B-360'}</span>
                                  </label>
                                ))}
                              </td>
                            </tr>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">2. Genel Özellikler</td></tr>
                          <tr>
                            <td className="border border-black px-1.5 py-1">A) Aletin Kullanım Amacı:</td>
                            <td className="border border-black p-0"><input type="text" value={formState.kullanimAmaci} onChange={e => setFormState(p => ({ ...p, kullanimAmaci: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td>
                          </tr>
                          <tr>
                            <td className="border border-black px-1.5 py-1">B) Hangi Görevlerde/Bakımlarda Kullanılacağı:</td>
                            <td className="border border-black p-0"><input type="text" value={formState.gorevBakimlar} onChange={e => setFormState(p => ({ ...p, gorevBakimlar: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td>
                          </tr>
                          <tr>
                            <td className="border border-black px-1.5 py-1">C) Kullanım Sıklığı:</td>
                            <td className="border border-black px-2 py-1 space-y-0.5">
                              {(['Sürekli', 'Dönemsel', 'Acil Durumlarda'] as const).map(v => (
                                <label key={v} className="flex items-center gap-1 cursor-pointer">
                                  <input type="checkbox" checked={formState.kullanimSikligi === v} onChange={() => setFormState(p => ({ ...p, kullanimSikligi: p.kullanimSikligi === v ? '' : v }))} />
                                  <span>{v}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr>
                            <td className="border border-black px-1.5 py-1">D) Çalışma Ortamı:</td>
                            <td className="border border-black px-2 py-1 space-y-0.5">
                              {(['Açık Alan', 'Kapalı Alan', 'Deniz / Su Alanı'] as const).map(v => (
                                <label key={v} className="flex items-center gap-1 cursor-pointer">
                                  <input type="checkbox" checked={formState.calismaOrtami === v} onChange={() => setFormState(p => ({ ...p, calismaOrtami: p.calismaOrtami === v ? '' : v }))} />
                                  <span>{v}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">3. Teknik Özellikler</td></tr>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">A) Boyut ve Ağırlık</td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Uzunluk:</td><td className="border border-black p-0"><input type="text" value={formState.uzunluk} onChange={e => setFormState(p => ({ ...p, uzunluk: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Genişlik:</td><td className="border border-black p-0"><input type="text" value={formState.genislik} onChange={e => setFormState(p => ({ ...p, genislik: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Yükseklik:</td><td className="border border-black p-0"><input type="text" value={formState.yukseklik} onChange={e => setFormState(p => ({ ...p, yukseklik: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Ağırlık:</td><td className="border border-black p-0"><input type="text" value={formState.agirlik} onChange={e => setFormState(p => ({ ...p, agirlik: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">B) Dayanıklılık ve Koruma</td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Isıya Dayanıklılık ( °C'ye Kadar):</td><td className="border border-black p-0"><input type="text" value={formState.isiDayaniklilik} onChange={e => setFormState(p => ({ ...p, isiDayaniklilik: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr>
                            <td className="border border-black px-1.5 py-0.5">Darbelere Karşı Dayanıklılık:</td>
                            <td className="border border-black px-2 py-0.5 flex gap-2">
                              {(['Hafif', 'Orta', 'Yüksek'] as const).map(v => (
                                <label key={v} className="flex items-center gap-1 cursor-pointer">
                                  <input type="checkbox" checked={formState.darbeDayaniklilik === v} onChange={() => setFormState(p => ({ ...p, darbeDayaniklilik: p.darbeDayaniklilik === v ? '' : v }))} />
                                  <span>{v}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          {[
                            { field: 'suGecirmezlik', label: 'Su Geçirmezlik:' },
                            { field: 'tozGecirmezlik', label: 'Toz Geçirmezlik:' },
                            { field: 'kimyasalDayaniklilik', label: 'Kimyasal Maddelere Dayanıklılık:' },
                            { field: 'uvKoruma', label: 'UV Işınlarına Karşı Koruma:' }
                          ].map(row => (
                            <tr key={row.field}>
                              <td className="border border-black px-1.5 py-0.5">{row.label}</td>
                              <td className="border border-black px-2 py-0.5 flex gap-3">
                                {(['Evet', 'Hayır'] as const).map(v => (
                                  <label key={v} className="flex items-center gap-1 cursor-pointer">
                                    <input
                                      type="checkbox"
                                      checked={(formState as any)[row.field] === v}
                                      onChange={() => setFormState(p => ({ ...p, [row.field]: (p as any)[row.field] === v ? '' : v }))}
                                    />
                                    <span>{v}</span>
                                  </label>
                                ))}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {/* PAGE 2 */}
                    <div className="bg-white text-black shadow-2xl border border-slate-400 p-4 font-sans text-[11px]">
                      <div className="border border-slate-600 px-2 py-1.5 flex items-center justify-between mb-2">
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-9 h-9 rounded-full object-cover shrink-0" />
                        <div className="text-center font-bold text-slate-500 text-[10px] uppercase px-1">
                          {currentTypeMeta?.formTitle}
                        </div>
                        <img src={circularOgmLogoDataUri} alt="OGM" className="w-9 h-9 rounded-full object-cover shrink-0" />
                      </div>
                      <table className="w-full border-collapse border border-black">
                        <tbody>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">C) Malzeme Türü</td></tr>
                          <tr>
                            <td className="border border-black px-1.5 py-1 w-[42%] align-top">Metaller:</td>
                            <td className="border border-black px-2 py-1 space-y-0.5">
                              {[
                                { k: 'celik', l: 'Çelik' },
                                { k: 'paslanmazCelik', l: 'Paslanmaz Çelik' },
                                { k: 'aluminyum', l: 'Alüminyum' },
                                { k: 'bakir', l: 'Bakır' },
                                { k: 'titanyum', l: 'Titanyum' },
                                { k: 'pirinc', l: 'Pirinç' },
                                { k: 'dokmeDemir', l: 'Dökme Demir' }
                              ].map(m => (
                                <label key={m.k} className="flex items-center gap-1 cursor-pointer">
                                  <input type="checkbox" checked={(formState.metaller as any)[m.k]} onChange={e => setFormState(p => ({ ...p, metaller: { ...p.metaller, [m.k]: e.target.checked } }))} />
                                  <span>{m.l}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Diğer:</td><td className="border border-black p-0"><input type="text" value={formState.metallerDiger} onChange={e => setFormState(p => ({ ...p, metallerDiger: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr>
                            <td className="border border-black px-1.5 py-1 align-top">Ametaller:</td>
                            <td className="border border-black px-2 py-1 space-y-0.5">
                              {[
                                { k: 'plastik', l: 'Plastik (ABS, PVC, PP)' },
                                { k: 'kaucuk', l: 'Kauçuk' },
                                { k: 'seramik', l: 'Seramik' },
                                { k: 'cam', l: 'Cam' },
                                { k: 'karbonFiber', l: 'Karbon Fiber' },
                                { k: 'kompozit', l: 'Kompozit Malzemeler' }
                              ].map(m => (
                                <label key={m.k} className="flex items-center gap-1 cursor-pointer">
                                  <input type="checkbox" checked={(formState.ametaller as any)[m.k]} onChange={e => setFormState(p => ({ ...p, ametaller: { ...p.ametaller, [m.k]: e.target.checked } }))} />
                                  <span>{m.l}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Diğer:</td><td className="border border-black p-0"><input type="text" value={formState.ametallerDiger} onChange={e => setFormState(p => ({ ...p, ametallerDiger: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">D) Yüzey İşlemi ve Kaplama</td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Kaplama:</td><td className="border border-black p-0"><input type="text" value={formState.kaplama} onChange={e => setFormState(p => ({ ...p, kaplama: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Boya ve Renk Kodu:</td><td className="border border-black p-0"><input type="text" value={formState.boyaRenkKodu} onChange={e => setFormState(p => ({ ...p, boyaRenkKodu: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Cilalama:</td><td className="border border-black p-0"><input type="text" value={formState.cilalama} onChange={e => setFormState(p => ({ ...p, cilalama: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Diğer:</td><td className="border border-black p-0"><input type="text" value={formState.yuzeyDiger} onChange={e => setFormState(p => ({ ...p, yuzeyDiger: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr>
                            <td className="border border-black px-1.5 py-0.5">Korozyona Karşı Dayanıklılık:</td>
                            <td className="border border-black px-2 py-0.5 flex gap-3">
                              {(['Evet', 'Hayır'] as const).map(v => (
                                <label key={v} className="flex items-center gap-1 cursor-pointer">
                                  <input type="checkbox" checked={formState.korozyonDayaniklilik === v} onChange={() => setFormState(p => ({ ...p, korozyonDayaniklilik: p.korozyonDayaniklilik === v ? '' : v }))} />
                                  <span>{v}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">E) Enerji ve Çalışma Mekanizması</td></tr>
                          <tr>
                            <td className="border border-black px-1.5 py-1 align-top">Enerji Kaynağı:</td>
                            <td className="border border-black px-2 py-1 space-y-0.5">
                              {[
                                { k: 'manuel', l: 'Manuel Çalışma' },
                                { k: 'pnomatik', l: 'Pnömatik' },
                                { k: 'hidrolik', l: 'Hidrolik' },
                                { k: 'elektrik', l: 'Elektrik' }
                              ].map(eOpt => (
                                <label key={eOpt.k} className="flex items-center gap-1 cursor-pointer">
                                  <input type="checkbox" checked={(formState.enerjiKaynagi as any)[eOpt.k]} onChange={e => setFormState(p => ({ ...p, enerjiKaynagi: { ...p.enerjiKaynagi, [eOpt.k]: e.target.checked } }))} />
                                  <span>{eOpt.l}</span>
                                </label>
                              ))}
                            </td>
                          </tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Varsa Batarya (Türü ve Kapasitesi):</td><td className="border border-black p-0"><input type="text" value={formState.bataryaBilgisi} onChange={e => setFormState(p => ({ ...p, bataryaBilgisi: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Varsa Çalışma Voltajı (V):</td><td className="border border-black p-0"><input type="text" value={formState.calismaVoltaji} onChange={e => setFormState(p => ({ ...p, calismaVoltaji: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Varsa Güç Tüketimi (Watt):</td><td className="border border-black p-0"><input type="text" value={formState.gucTuketimi} onChange={e => setFormState(p => ({ ...p, gucTuketimi: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                          <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">F) Güvenlik ve Standartlar</td></tr>
                          <tr>
                            <td className="border border-black px-1.5 py-0.5">Uluslararası Standartlara Uygunluk:</td>
                            <td className="border border-black px-2 py-0.5 flex gap-3">
                              <label className="flex items-center gap-1 cursor-pointer">
                                <input type="checkbox" checked={formState.standartlar.ce} onChange={e => setFormState(p => ({ ...p, standartlar: { ...p.standartlar, ce: e.target.checked } }))} />
                                <span>CE</span>
                              </label>
                              <label className="flex items-center gap-1 cursor-pointer">
                                <input type="checkbox" checked={formState.standartlar.iso} onChange={e => setFormState(p => ({ ...p, standartlar: { ...p.standartlar, iso: e.target.checked } }))} />
                                <span>ISO</span>
                              </label>
                            </td>
                          </tr>
                          <tr><td className="border border-black px-1.5 py-0.5">Diğer:</td><td className="border border-black p-0"><input type="text" value={formState.standartDiger} onChange={e => setFormState(p => ({ ...p, standartDiger: e.target.value }))} className="w-full px-1 py-0.5 outline-none" /></td></tr>
                        </tbody>
                      </table>
                    </div>

                    </div>

                    {/* PAGE 3 BELOW PAGE 1 & 2 */}
                    <div className="max-w-[565px] mx-auto w-full">
                      <div className="bg-white text-black shadow-2xl border border-slate-400 p-5 font-sans text-[11px] flex flex-col justify-between min-h-[650px]">
                        <div>
                          <div className="border border-slate-600 px-2 py-1.5 flex items-center justify-between mb-2">
                            <img src={circularOgmLogoDataUri} alt="OGM" className="w-10 h-10 rounded-full object-cover shrink-0" />
                            <div className="text-center font-bold text-slate-500 text-[10.5px] uppercase px-1">
                              {currentTypeMeta?.formTitle}
                            </div>
                            <img src={circularOgmLogoDataUri} alt="OGM" className="w-10 h-10 rounded-full object-cover shrink-0" />
                          </div>
                          <table className="w-full border-collapse border border-black">
                            <tbody>
                              <tr className="bg-[#d9d9d9]"><td colSpan={2} className="border border-black px-1.5 py-0.5 font-bold">4. Çizim Alanı</td></tr>
                              <tr>
                                <td colSpan={2} className="border border-black p-2">
                                  <div className="text-[10px] mb-1">
                                    <strong>Formu Dolduran Kişi, FORMDA OLMAYAN FAKAT İLAVE ETMEK İSTEDİĞİ TÜM ÖZELLİKLERİ YAZABİLİR.</strong> Aşağıdaki Alanı Tasarım Ve Çizim İçin Kullanabilir.
                                  </div>
                                  <textarea
                                    rows={5}
                                    value={formState.cizimNotlari}
                                    onChange={e => setFormState(p => ({ ...p, cizimNotlari: e.target.value }))}
                                    placeholder="Tasarım notları veya çizim açıklaması..."
                                    className="w-full p-1 bg-slate-50 border border-slate-200 rounded outline-none resize-none"
                                  />
                                  <FormImageGallery
                                    images={formState.cizimGorselleri}
                                    onChange={next => setFormState(p => ({ ...p, cizimGorselleri: next }))}
                                    defaultScale={defaultImageScale}
                                    onDefaultScaleChange={setDefaultImageScale}
                                    primaryLabel="Tasarım / Çizim Görseli Yükle (Çoklu)"
                                    compact
                                    onNotify={showNotification}
                                  />
                                </td>
                              </tr>
                              <tr>
                                <td colSpan={2} className="border border-black p-2">
                                  <div className="text-[10px] mb-1">
                                    Lütfen İstenen Ürünün Varsa Kataloğundan Alınmış Fotoğrafı Veya Örnek Ürün Linklerini Formun Ekine Ekleyiniz.
                                  </div>
                                  <textarea
                                    rows={3}
                                    value={formState.ornekUrunLinkNotlari}
                                    onChange={e => setFormState(p => ({ ...p, ornekUrunLinkNotlari: e.target.value }))}
                                    placeholder="Örnek ürün linkleri / not..."
                                    className="w-full p-1 bg-slate-50 border border-slate-200 rounded outline-none resize-none"
                                  />
                                  <FormImageGallery
                                    images={formState.ornekUrunGorselleri}
                                    onChange={next => setFormState(p => ({ ...p, ornekUrunGorselleri: next }))}
                                    defaultScale={defaultImageScale}
                                    onDefaultScaleChange={setDefaultImageScale}
                                    primaryLabel="Örnek Ürün Fotoğrafı Yükle (Çoklu)"
                                    compact
                                    onNotify={showNotification}
                                  />
                                </td>
                              </tr>
                              <tr>
                                <td className="border border-black px-1.5 py-1.5 w-[42%]">Talep Yapan Personel Adı Soyadı:</td>
                                <td className="border border-black p-0.5 relative" onClick={e => e.stopPropagation()}>
                                  <input
                                    type="text"
                                    value={formState.talepYapanAdSoyad}
                                    onFocus={() => setActiveCellAutocomplete('personnel')}
                                    onChange={e => {
                                      setFormState(p => ({ ...p, talepYapanAdSoyad: e.target.value }));
                                      setActiveCellAutocomplete('personnel');
                                    }}
                                    placeholder="Personel Adı Soyadı yazınız..."
                                    className="w-full px-1 py-0.5 font-bold outline-none focus:bg-blue-50"
                                  />
                                  {activeCellAutocomplete === 'personnel' && getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).length > 0 && (
                                    <div className="absolute left-0 right-0 bottom-full mb-1 bg-white border border-emerald-500 rounded-lg shadow-xl max-h-48 overflow-y-auto z-[10500] divide-y divide-slate-100">
                                      {getFilteredPersonnelSuggestions(formState.talepYapanAdSoyad).map((person, i) => (
                                        <button
                                          key={i}
                                          type="button"
                                          onMouseDown={e => {
                                            e.preventDefault();
                                            setFormState(p => ({ ...p, talepYapanAdSoyad: person }));
                                            setActiveCellAutocomplete(null);
                                          }}
                                          className="w-full px-2.5 py-1.5 text-left hover:bg-emerald-50 text-xs font-bold text-slate-900 cursor-pointer"
                                        >
                                          {person}
                                        </button>
                                      ))}
                                    </div>
                                  )}
                                </td>
                              </tr>
                              <tr>
                                <td className="border border-black px-1.5 py-3">Talep Sahibi İmza:</td>
                                <td className="border border-black px-1.5 py-3"></td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* SAVED CONFIRMATION MODAL WITH TALEP NO (YYYY-SIRA NO) */}
      {saveSuccessModal && (
        <div className="fixed inset-0 z-[10008] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-white text-slate-800 w-full max-w-md rounded-2xl shadow-2xl border-2 border-emerald-500 overflow-hidden">
            <div className="px-5 py-4 bg-emerald-50 border-b border-emerald-200 flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0">
                <CheckCircle2 className="w-6 h-6" />
              </div>
              <div>
                <h4 className="text-sm font-black text-emerald-950 uppercase">TALEBİNİZ SİSTEME İŞLENDİ</h4>
                <div className="text-xs font-mono font-bold text-emerald-700">
                  Talep No: {saveSuccessModal.talepNo}
                </div>
              </div>
            </div>
            <div className="p-6 space-y-4">
              <p className="text-xs sm:text-sm font-bold text-slate-700 leading-relaxed">
                Talebiniz <strong>"ONAY BEKLEYEN TALEPLER"</strong> sistemine <strong>{saveSuccessModal.talepNo}</strong> sıra numarası ile kaydedildi.
              </p>
              <div className="p-3.5 rounded-xl bg-amber-50 border border-amber-300 text-amber-950 text-xs font-extrabold leading-relaxed">
                📄 İndirilen Word / PDF belgesini ıslak imzalayıp PDF olarak taradıktan sonra <strong>"Onay Bekleyen Talepler"</strong> listesinde yer alan <strong>"BELGE YÜKLE"</strong> butonundan yükleyebilirsiniz.
              </div>
              <div className="flex items-center justify-end gap-2 flex-wrap pt-2">
                {selectedType && (
                  <button
                    type="button"
                    onClick={() => triggerPrintPdf(selectedType, formState, saveSuccessModal.talepNo)}
                    className="px-3.5 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-800 text-xs font-bold cursor-pointer flex items-center gap-1.5"
                  >
                    <Printer className="w-4 h-4" />
                    <span>Yazdır / PDF</span>
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => {
                    const targetRec =
                      saveSuccessModal.record ||
                      submittedList.find(it => it.talepNo === saveSuccessModal.talepNo) ||
                      ({
                        id: saveSuccessModal.talepNo,
                        talepNo: saveSuccessModal.talepNo,
                        formType: selectedType || 'parca_talep',
                        formTypeLabel: '',
                        formTitle: saveSuccessModal.formTitle,
                        unit: formState.selectedUnit || 'genel',
                        unitLabel: formState.havaAraciTipi || 'Genel',
                        tailNo: formState.kuyrukTescilNo || '-',
                        summary: '',
                        requestedBy: formState.talepYapanAdSoyad || '',
                        date: formState.date || getTodayTr(),
                        timestamp: new Date().toISOString(),
                        status: 'ONAY BEKLİYOR'
                      } as SubmittedTalepRecord);
                    setSaveSuccessModal(null);
                    setActiveMainTab('archive');
                    handleClickBelgeYukle(targetRec);
                  }}
                  className="px-4 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-black cursor-pointer flex items-center gap-1.5 shadow-sm"
                >
                  <Upload className="w-4 h-4" />
                  <span>Islak İmzalı Belge Yükle</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setSaveSuccessModal(null);
                    setActiveMainTab('archive');
                  }}
                  className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black uppercase cursor-pointer"
                >
                  Tamam
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* WET-SIGNED PDF UPLOAD WARNING MODAL (BEFORE OPENING bulut.ogm.gov.tr/SARTNAMELER) */}
      {pendingWetSignedWarning && (
        <div
          className="fixed inset-0 z-[10010] flex items-center justify-center p-4 bg-slate-900/65 backdrop-blur-sm animate-fade-in"
          onClick={() => setPendingWetSignedWarning(null)}
        >
          <div
            className="bg-white text-slate-900 w-full max-w-lg rounded-2xl shadow-2xl border-2 border-amber-400 overflow-hidden"
            onClick={e => e.stopPropagation()}
          >
            <div className="px-5 py-4 bg-amber-50 border-b border-amber-200 flex items-center justify-between gap-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-amber-500 text-white flex items-center justify-center shadow-xs shrink-0">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="text-sm font-black text-amber-950 uppercase tracking-wide">
                    ISLAK İMZALI BELGE YÜKLEME UYARISI
                  </h4>
                  <div className="text-[11px] font-mono font-bold text-blue-700">
                    Kayıt Zamanı: {pendingWetSignedWarning.clickedAt}
                  </div>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setPendingWetSignedWarning(null)}
                className="text-slate-400 hover:text-slate-700 font-bold text-sm p-1 cursor-pointer"
              >
                ✕
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div className="p-4 rounded-xl bg-amber-50/90 border border-amber-300 text-amber-950 text-xs sm:text-[13px] font-extrabold leading-relaxed space-y-2">
                <p>
                  LÜTFEN TARANMIŞ ISLAK İMZALI PDF BELGESİNİ MUTLAKA <span className="underline">TARİH VE AÇIKLAMA</span> BİLGİLERİ İLE YÜKLEYİNİZ.
                </p>
                <p>
                  DOSYA ADI KISMINA TALEBİ KISA VE ÖZETLEYİCİ BİR BAŞLIK YAZINIZ.
                </p>
              </div>
              <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                <div className="flex items-center justify-between">
                  <span className="text-[10px] font-black uppercase text-slate-500 font-mono">
                    ÖNERİLEN KISA ÖZETLİ DOSYA ADI / AÇIKLAMA
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      const suggested = `${pendingWetSignedWarning.clickedAt.split(' ')[0]} - ${pendingWetSignedWarning.record.talepNo} - ${pendingWetSignedWarning.record.unitLabel} ${pendingWetSignedWarning.record.formTypeLabel} - ${pendingWetSignedWarning.record.requestedBy}`;
                      navigator.clipboard?.writeText(suggested).catch(() => {});
                      showNotification('Önerilen başlık panoya kopyalandı.', 'info');
                    }}
                    className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-white hover:bg-slate-100 border border-slate-300 text-[10px] font-bold text-slate-700 cursor-pointer"
                  >
                    <Copy className="w-3 h-3" />
                    <span>Kopyala</span>
                  </button>
                </div>
                <div className="text-xs font-mono font-bold text-slate-800 break-words">
                  {pendingWetSignedWarning.clickedAt.split(' ')[0]} - {pendingWetSignedWarning.record.talepNo} - {pendingWetSignedWarning.record.unitLabel} {pendingWetSignedWarning.record.formTypeLabel} - {pendingWetSignedWarning.record.requestedBy}
                </div>
              </div>
              <div className="flex items-center gap-2.5 pt-1">
                <button
                  type="button"
                  onClick={() => setPendingWetSignedWarning(null)}
                  className="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="button"
                  onClick={handleConfirmWarningAndOpenUpload}
                  className="flex-1 py-2.5 px-5 rounded-xl bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider cursor-pointer shadow-md flex items-center justify-center gap-2 active:scale-95 transition-all"
                >
                  <CheckCircle2 className="w-4 h-4" />
                  <span>OKUDUM, ANLADIM (YÜKLEMEYE GİT)</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* TASKLINE ONLINE AKIŞ GEÇMİŞİ ZAMAN ÇİZELGESİ MODALI */}
      {tasklineFlowModal && (
        <div
          className="fixed inset-0 z-[10009] flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs animate-fade-in"
          onClick={() => setTasklineFlowModal(null)}
        >
          <div
            className="bg-white text-slate-900 w-full max-w-2xl max-h-[86vh] rounded-3xl shadow-2xl border border-slate-200 flex flex-col overflow-hidden"
            onClick={e => e.stopPropagation()}
          >
            <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-3">
              <div className="flex items-center gap-3 min-w-0">
                <div className="w-10 h-10 rounded-2xl bg-blue-100 text-blue-700 border border-blue-200 flex items-center justify-center shrink-0">
                  <GitBranch className="w-5 h-5" />
                </div>
                <div className="min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <h4 className="text-sm font-black text-slate-900 uppercase tracking-wide">
                      TASKLINE AKIŞ GEÇMİŞİ
                    </h4>
                    <span className="px-2 py-0.5 rounded-full text-[11px] font-mono font-black bg-blue-600 text-white">
                      {tasklineFlowModal.talepNo}
                    </span>
                    {tasklineFlowModal.demand?.status && (
                      <span className="px-2 py-0.5 rounded-full text-[10px] font-black bg-emerald-50 text-emerald-800 border border-emerald-200">
                        {tasklineFlowModal.demand.status}
                      </span>
                    )}
                  </div>
                  <p className="text-[11px] text-slate-500 truncate mt-0.5">
                    {tasklineFlowModal.demand?.title
                      ? `${tasklineFlowModal.demand.title}${tasklineFlowModal.demand.requestingUnit ? ` • ${tasklineFlowModal.demand.requestingUnit}` : ''}`
                      : 'İşlemdeki Talepler & Tamamlanan Talepler (H Sütunu / K Sütunu)'}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-1.5 shrink-0">
                <button
                  type="button"
                  disabled={tasklineFlowModal.loading}
                  onClick={() => handleOpenTasklineFlow(tasklineFlowModal.talepNo)}
                  title="Taskline Verilerini Yenile"
                  className="p-2 rounded-xl bg-white hover:bg-slate-100 text-slate-600 border border-slate-200 cursor-pointer disabled:opacity-50"
                >
                  <RefreshCw className={`w-4 h-4 ${tasklineFlowModal.loading ? 'animate-spin' : ''}`} />
                </button>
                <button
                  type="button"
                  onClick={() => setTasklineFlowModal(null)}
                  className="p-2 rounded-xl bg-white hover:bg-rose-50 text-slate-600 hover:text-rose-600 border border-slate-200 cursor-pointer"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-y-auto p-6 bg-slate-50/70">
              {tasklineFlowModal.loading ? (
                <div className="py-16 flex flex-col items-center justify-center gap-3 text-slate-500 text-xs font-bold">
                  <RefreshCw className="w-7 h-7 text-blue-600 animate-spin" />
                  <span>Taskline sistemi ile canlı bağlantı kuruluyor...</span>
                </div>
              ) : tasklineFlowModal.error ? (
                <div className="p-5 rounded-2xl bg-rose-50 border border-rose-200 text-rose-800 text-xs font-semibold text-center">
                  {tasklineFlowModal.error}
                </div>
              ) : !tasklineFlowModal.matched ? (
                <div className="p-8 rounded-2xl bg-white border border-slate-200 text-center space-y-2">
                  <div className="text-sm font-black text-slate-800">
                    "{tasklineFlowModal.talepNo}" Numarası Taskline H Sütununda Bulunamadı
                  </div>
                  <p className="text-xs text-slate-500 max-w-md mx-auto leading-relaxed">
                    Taskline sistemindeki "İşlemdeki Talepler" ve "Tamamlanan Talepler" sayfalarının H sütununda (EBYS Numarası) bu kod ile eşleşen bir kayıt henüz bulunmamaktadır.
                  </p>
                </div>
              ) : tasklineFlowModal.history.length === 0 ? (
                <div className="p-8 rounded-2xl bg-white border border-slate-200 text-center space-y-1.5">
                  <div className="text-sm font-black text-slate-800">Kayıt Bulundu, Akış Geçmişi Henüz Boş</div>
                  <p className="text-xs text-slate-500">
                    Talep kaydı eşleşti ancak K sütununda henüz akış geçmişi adımı yer almıyor.
                  </p>
                </div>
              ) : (
                <div className="relative pl-6 ml-2 border-l-2 border-blue-200 space-y-4">
                  {tasklineFlowModal.history.map((item, idx) => (
                    <div key={idx} className="relative">
                      <span className="absolute -left-[31px] top-4 w-3.5 h-3.5 rounded-full bg-blue-600 ring-4 ring-blue-100 border-2 border-white shadow-xs" />
                      <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs hover:shadow-md transition-shadow">
                        <div className="flex items-center justify-between gap-2 mb-1.5">
                          <span className="text-[11px] font-black text-blue-700 tracking-wide uppercase">
                            {formatTasklineDateTr(item.date)}
                          </span>
                          {item.isSupportNote && (
                            <span className="px-2 py-0.5 rounded-full text-[9.5px] font-bold bg-amber-50 text-amber-800 border border-amber-200">
                              Destek Notu
                            </span>
                          )}
                        </div>
                        <p className="text-xs sm:text-[13px] text-slate-800 font-medium leading-relaxed whitespace-pre-wrap">
                          {item.description}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ÖZEL ALET KALİBRASYON — GÜN TAKİP İKAZLI ÜRÜN SEÇİM PENCERESİ */}
      {kalibrasyonAlertModal && (() => {
        const allItems = kalibrasyonAlertModal.items;
        const overdueCount = allItems.filter(i => i.daysDiff <= 0).length;
        const upcomingCount = allItems.filter(i => i.daysDiff > 0).length;
        const q = kalibrasyonAlertSearch.trim().toLowerCase();

        const filteredItems = allItems.filter(it => {
          if (kalibrasyonAlertFilter === 'overdue' && it.daysDiff > 0) return false;
          if (kalibrasyonAlertFilter === 'upcoming' && it.daysDiff <= 0) return false;
          if (!q) return true;
          return `${it.techName} ${it.partNo} ${it.seriNo} ${it.location} ${it.unitLabel} ${it.gelecekTarih}`
            .toLowerCase()
            .includes(q);
        });

        const selectedCount = Object.values(kalibrasyonAlertSelectedIds).filter(Boolean).length;
        const existingFormSigs = new Set(
          formState.kalibrasyonRows
            .filter(r => r.malzemeAdi || r.parcaNo || r.seriNo)
            .map(r =>
              `${r.malzemeAdi.trim().toUpperCase()}__${r.parcaNo.trim().toUpperCase()}__${r.seriNo.trim().toUpperCase()}__${r.kalibreBitisTarihi.trim()}`
            )
        );

        return (
          <div
            className="fixed inset-0 z-[10012] flex items-center justify-center p-3 sm:p-5 bg-slate-900/60 backdrop-blur-xs animate-fade-in"
            onClick={() => setKalibrasyonAlertModal(null)}
          >
            <div
              className="bg-white text-slate-900 w-full max-w-6xl max-h-[90vh] rounded-3xl shadow-2xl border border-slate-200 flex flex-col overflow-hidden"
              onClick={e => e.stopPropagation()}
            >
              {/* Modal Header */}
              <div className="px-6 py-4 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-rose-50 border border-rose-200 text-rose-600 flex items-center justify-center shrink-0">
                    <AlertTriangle className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <h4 className="text-sm sm:text-base font-black text-slate-900 uppercase tracking-wide">
                        GÜN TAKİP İKAZLI ÖZEL ALETLER
                      </h4>
                      <span className="px-2.5 py-0.5 rounded-full text-[11px] font-black bg-slate-900 text-white">
                        {kalibrasyonAlertModal.unitTitle}
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-500 mt-0.5">
                      Kalibrasyon süresi dolmuş veya 90 gün içinde dolacak özel aletler listelenmektedir. Başına tik attığınız ürünler kalibrasyon talep tablosuna otomatik aktarılır.
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setKalibrasyonAlertModal(null)}
                  className="p-2 rounded-xl bg-slate-100 hover:bg-rose-50 text-slate-500 hover:text-rose-600 cursor-pointer"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              {/* Filter Tabs & Search Bar (Matching GunTakipModal Design) */}
              <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setKalibrasyonAlertFilter('all')}
                    className={`px-3.5 py-1.5 rounded-full text-[11px] font-black tracking-wide transition-all cursor-pointer border ${
                      kalibrasyonAlertFilter === 'all'
                        ? 'bg-slate-900 text-white border-slate-900 shadow-xs'
                        : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                    }`}
                  >
                    TÜMÜ ({allItems.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setKalibrasyonAlertFilter('overdue')}
                    className={`px-3.5 py-1.5 rounded-full text-[11px] font-black tracking-wide transition-all cursor-pointer border inline-flex items-center gap-1.5 ${
                      kalibrasyonAlertFilter === 'overdue'
                        ? 'bg-rose-600 text-white border-rose-600 shadow-xs'
                        : 'bg-white text-rose-700 border-rose-200 hover:bg-rose-50'
                    }`}
                  >
                    <span className={`w-2 h-2 rounded-full ${kalibrasyonAlertFilter === 'overdue' ? 'bg-white' : 'bg-rose-500'}`} />
                    GÜNÜ GEÇENLER ({overdueCount})
                  </button>
                  <button
                    type="button"
                    onClick={() => setKalibrasyonAlertFilter('upcoming')}
                    className={`px-3.5 py-1.5 rounded-full text-[11px] font-black tracking-wide transition-all cursor-pointer border inline-flex items-center gap-1.5 ${
                      kalibrasyonAlertFilter === 'upcoming'
                        ? 'bg-amber-500 text-white border-amber-500 shadow-xs'
                        : 'bg-white text-amber-800 border-amber-300 hover:bg-amber-50'
                    }`}
                  >
                    <span className={`w-2 h-2 rounded-full ${kalibrasyonAlertFilter === 'upcoming' ? 'bg-white' : 'bg-amber-500'}`} />
                    YAKLAŞANLAR ({upcomingCount})
                  </button>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <div className="relative min-w-[230px]">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={kalibrasyonAlertSearch}
                      onChange={e => setKalibrasyonAlertSearch(e.target.value)}
                      placeholder="Malzeme, P/N, S/N veya Yer Ara..."
                      className="w-full pl-8 pr-3 py-1.5 rounded-full bg-white border border-slate-300 text-xs text-slate-800 placeholder:text-slate-400 focus:outline-none focus:border-slate-500"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      const next: Record<string, boolean> = { ...kalibrasyonAlertSelectedIds };
                      filteredItems.forEach(it => {
                        next[it.id] = true;
                      });
                      setKalibrasyonAlertSelectedIds(next);
                    }}
                    className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-300 text-[11px] font-bold cursor-pointer"
                  >
                    Tüm Görünenleri Seç
                  </button>
                  {overdueCount > 0 && (
                    <button
                      type="button"
                      onClick={() => {
                        const next: Record<string, boolean> = { ...kalibrasyonAlertSelectedIds };
                        allItems.filter(it => it.daysDiff <= 0).forEach(it => {
                          next[it.id] = true;
                        });
                        setKalibrasyonAlertSelectedIds(next);
                      }}
                      className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-rose-50 text-rose-700 border border-rose-300 text-[11px] font-bold cursor-pointer"
                    >
                      Günü Geçenleri Seç
                    </button>
                  )}
                  {selectedCount > 0 && (
                    <button
                      type="button"
                      onClick={() => setKalibrasyonAlertSelectedIds({})}
                      className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-slate-100 text-slate-600 border border-slate-300 text-[11px] font-bold cursor-pointer"
                    >
                      Temizle
                    </button>
                  )}
                </div>
              </div>

              {/* Table Body */}
              <div className="flex-1 overflow-auto bg-white">
                {kalibrasyonAlertModal.loading ? (
                  <div className="py-16 flex flex-col items-center justify-center gap-3 text-slate-500 text-xs font-bold">
                    <RefreshCw className="w-7 h-7 text-emerald-600 animate-spin" />
                    <span>Gün Takip & Özel Alet kalibrasyon verileri taranıyor...</span>
                  </div>
                ) : filteredItems.length === 0 ? (
                  <div className="py-16 text-center space-y-2">
                    <CheckCircle2 className="w-10 h-10 text-emerald-500 mx-auto" />
                    <div className="text-sm font-black text-slate-800">
                      {allItems.length === 0
                        ? `${kalibrasyonAlertModal.unitTitle} için günü geçmiş veya 90 gün içinde dolacak özel alet bulunamadı.`
                        : 'Arama / filtre kriterine uyan ikazlı ürün bulunamadı.'}
                    </div>
                  </div>
                ) : (
                  <table className="w-full border-collapse text-left text-[11.5px]">
                    <thead className="sticky top-0 z-10 bg-[#0b132b] text-white uppercase text-[10px] font-black tracking-wider">
                      <tr>
                        <th className="px-3 py-2.5 text-center w-10">
                          <input
                            type="checkbox"
                            checked={filteredItems.length > 0 && filteredItems.every(it => kalibrasyonAlertSelectedIds[it.id])}
                            onChange={e => {
                              const checked = e.target.checked;
                              setKalibrasyonAlertSelectedIds(prev => {
                                const next = { ...prev };
                                filteredItems.forEach(it => {
                                  next[it.id] = checked;
                                });
                                return next;
                              });
                            }}
                            className="w-4 h-4 accent-emerald-500 cursor-pointer"
                          />
                        </th>
                        <th className="px-2.5 py-2.5 text-center w-10">#</th>
                        <th className="px-3 py-2.5">TEÇHİZAT / MALZEME ADI</th>
                        <th className="px-3 py-2.5">P/N VEYA SİCİL NO</th>
                        <th className="px-3 py-2.5">SERİ NO (S/N)</th>
                        <th className="px-3 py-2.5 text-center">NORM KAPSAMI</th>
                        <th className="px-3 py-2.5">BULUNDUĞU YER</th>
                        <th className="px-3 py-2.5 text-center">SON İŞLEM</th>
                        <th className="px-3 py-2.5 text-center">GELECEK TARİH</th>
                        <th className="px-3 py-2.5 text-center">KALAN GÜN</th>
                        <th className="px-3 py-2.5 text-center">BİLDİRİM DURUMU</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {filteredItems.map((it, idx) => {
                        const isOverdue = it.daysDiff <= 0;
                        const isSelected = Boolean(kalibrasyonAlertSelectedIds[it.id]);
                        const cleanPn = it.partNo === '-' ? '' : it.partNo;
                        const cleanSn = it.seriNo === '-' ? '' : it.seriNo;
                        const sig = `${it.techName.trim().toUpperCase()}__${cleanPn.trim().toUpperCase()}__${cleanSn.trim().toUpperCase()}__${it.gelecekTarih.trim()}`;
                        const alreadyInForm = existingFormSigs.has(sig);

                        return (
                          <tr
                            key={it.id}
                            onClick={() =>
                              setKalibrasyonAlertSelectedIds(prev => ({
                                ...prev,
                                [it.id]: !prev[it.id]
                              }))
                            }
                            className={`cursor-pointer transition-colors border-l-4 ${
                              isOverdue
                                ? `${isSelected ? 'bg-rose-100/90' : 'bg-rose-50/75 hover:bg-rose-100/60'} border-l-rose-600`
                                : `${isSelected ? 'bg-amber-100/90' : 'bg-amber-50/75 hover:bg-amber-100/60'} border-l-amber-500`
                            }`}
                          >
                            <td className="px-3 py-2.5 text-center" onClick={e => e.stopPropagation()}>
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={e =>
                                  setKalibrasyonAlertSelectedIds(prev => ({
                                    ...prev,
                                    [it.id]: e.target.checked
                                  }))
                                }
                                className="w-4 h-4 accent-emerald-600 cursor-pointer"
                              />
                            </td>
                            <td className="px-2.5 py-2.5 text-center font-mono font-bold text-slate-500">
                              {idx + 1}
                            </td>
                            <td className="px-3 py-2.5 font-extrabold text-slate-900">
                              <div>{it.techName}</div>
                              {alreadyInForm && (
                                <span className="inline-block mt-0.5 px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 border border-emerald-300 text-[9.5px] font-black">
                                  ✓ Formda Mevcut
                                </span>
                              )}
                            </td>
                            <td className="px-3 py-2.5 font-mono text-slate-700">{it.partNo || '-'}</td>
                            <td className="px-3 py-2.5 font-mono font-semibold text-slate-800">{it.seriNo || '-'}</td>
                            <td className="px-3 py-2.5 text-center">
                              <span className="inline-block px-2 py-0.5 rounded-md bg-white border border-slate-300 text-[9.5px] font-black text-slate-700 leading-tight">
                                ÖZEL ALET
                              </span>
                            </td>
                            <td className="px-3 py-2.5 font-bold text-slate-700">{it.location || '-'}</td>
                            <td className="px-3 py-2.5 text-center font-mono text-slate-500">{it.sonTarih || '-'}</td>
                            <td className="px-3 py-2.5 text-center font-mono font-black text-slate-900">{it.gelecekTarih}</td>
                            <td className="px-3 py-2.5 text-center font-black leading-tight">
                              {isOverdue ? (
                                <span className="text-rose-600">
                                  {it.daysDiff === 0 ? 'BUGÜN DOLUYOR' : `${Math.abs(it.daysDiff)} GÜN GEÇTİ`}
                                </span>
                              ) : (
                                <span className="text-amber-700">{it.daysDiff} GÜN KALDI</span>
                              )}
                            </td>
                            <td className="px-3 py-2.5 text-center">
                              {isOverdue ? (
                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-rose-500 text-white text-[9.5px] font-black shadow-xs">
                                  OTOMATİK BİLDİRİMDE
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-amber-100 text-amber-900 border border-amber-300 text-[9.5px] font-black">
                                  TAKİPTE (YAKLAŞAN)
                                </span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Footer */}
              <div className="px-6 py-3.5 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
                <div className="text-xs font-bold text-slate-600">
                  Görüntülenen <span className="text-slate-900 font-black">{filteredItems.length}</span> Kalem •{' '}
                  Seçilen: <span className="text-emerald-700 font-black">{selectedCount}</span> Ürün
                </div>
                <div className="flex items-center gap-2.5">
                  <button
                    type="button"
                    onClick={() => setKalibrasyonAlertModal(null)}
                    className="px-4 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-800 text-xs font-bold cursor-pointer"
                  >
                    Kapat
                  </button>
                  <button
                    type="button"
                    disabled={selectedCount === 0}
                    onClick={handleTransferSelectedKalibrasyonAlerts}
                    className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-45 disabled:cursor-not-allowed text-white text-xs font-black uppercase tracking-wide cursor-pointer shadow-sm flex items-center gap-1.5"
                  >
                    <CheckCircle2 className="w-4 h-4" />
                    <span>Seçilenleri Forma Aktar ({selectedCount})</span>
                  </button>
                </div>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
};
