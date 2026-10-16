import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import {
  Camera,
  X,
  Play,
  Pause,
  RotateCcw,
  SkipBack,
  SkipForward,
  Clock,
  Gauge,
  Flame,
  AlertTriangle,
  CheckCircle2,
  Sliders,
  TrendingUp,
  Activity,
  ZoomIn,
  ZoomOut,
  Maximize2,
  Crosshair,
  Info
} from 'lucide-react';
import { YELLOW_FOCUS_COLUMNS, YellowFocusColumnConfig, ParsedMvpRow, ExceedanceEpisode } from './MvpFlightDataAnalysisModal';
import html2canvas from 'html2canvas-pro';

interface FlightEngineChartModalProps {
  isOpen: boolean;
  onClose: () => void;
  rows: ParsedMvpRow[];
  headers: string[];
  episodes: ExceedanceEpisode[];
  initialParamKey?: string;
  tailLabel: string;
  motorSn: string;
}

// Parametre eşik aralıkları ve limit sınırları: En üstte Kırmızı, Sonra Sarı, Sonra Yeşil
interface LimitBands {
  minY: number;
  maxY: number;
  unit: string;
  bands: {
    label: string;
    level: 'GREEN' | 'YELLOW' | 'RED';
    min: number;
    max: number;
    color: string;
    fillColor: string;
  }[];
}

const PARAM_LIMITS_MAP: Record<string, LimitBands> = {
  ng: {
    minY: 50,
    maxY: 110,
    unit: '%',
    bands: [
      { label: 'Normal / Yeşil (< %104)', level: 'GREEN', min: 0, max: 104.0, color: '#16a34a', fillColor: 'rgba(34, 197, 94, 0.10)' },
      { label: 'Sarı İkaz (%104 - %105,5)', level: 'YELLOW', min: 104.0, max: 105.5, color: '#ca8a04', fillColor: 'rgba(234, 179, 8, 0.18)' },
      { label: 'Kritik Limit / Kırmızı (> %105,5)', level: 'RED', min: 105.5, max: 115.0, color: '#dc2626', fillColor: 'rgba(239, 68, 68, 0.15)' }
    ]
  },
  itt: {
    minY: 400,
    maxY: 1100,
    unit: '°C',
    bands: [
      { label: 'Normal / Yeşil (< 870 °C)', level: 'GREEN', min: 0, max: 870.0, color: '#16a34a', fillColor: 'rgba(34, 197, 94, 0.10)' },
      { label: 'Sarı İkaz (870 - 1000 °C)', level: 'YELLOW', min: 870.0, max: 1000.0, color: '#ca8a04', fillColor: 'rgba(234, 179, 8, 0.18)' },
      { label: 'Kritik Limit / Kırmızı (> 1000 °C)', level: 'RED', min: 1000.0, max: 1200.0, color: '#dc2626', fillColor: 'rgba(239, 68, 68, 0.15)' }
    ]
  },
  np: {
    minY: 800,
    maxY: 2100,
    unit: 'RPM',
    bands: [
      { label: 'Normal / Yeşil (< 1700 RPM)', level: 'GREEN', min: 0, max: 1700.0, color: '#16a34a', fillColor: 'rgba(34, 197, 94, 0.10)' },
      { label: 'Sarı İkaz (1700 - 1870 RPM)', level: 'YELLOW', min: 1700.0, max: 1870.0, color: '#ca8a04', fillColor: 'rgba(234, 179, 8, 0.18)' },
      { label: 'Kritik Limit / Kırmızı (> 1870 RPM)', level: 'RED', min: 1870.0, max: 2200.0, color: '#dc2626', fillColor: 'rgba(239, 68, 68, 0.15)' }
    ]
  },
  torq: {
    minY: 0,
    maxY: 7200,
    unit: 'LB/FT',
    bands: [
      { label: 'Normal / Yeşil (< 5252 LB/FT)', level: 'GREEN', min: 0, max: 5252.0, color: '#16a34a', fillColor: 'rgba(34, 197, 94, 0.10)' },
      { label: 'Sarı İkaz (5252 - 6092 LB/FT)', level: 'YELLOW', min: 5252.0, max: 6092.0, color: '#ca8a04', fillColor: 'rgba(234, 179, 8, 0.18)' },
      { label: 'Kritik Limit / Kırmızı (> 6092 LB/FT)', level: 'RED', min: 6092.0, max: 7500.0, color: '#dc2626', fillColor: 'rgba(239, 68, 68, 0.15)' }
    ]
  },
  fuel_p: {
    minY: 0,
    maxY: 35,
    unit: 'PSI',
    bands: [
      { label: 'Normal / Yeşil (10 - 20 PSI)', level: 'GREEN', min: 0, max: 20.0, color: '#16a34a', fillColor: 'rgba(34, 197, 94, 0.10)' },
      { label: 'Sarı İkaz (20 - 25 PSI)', level: 'YELLOW', min: 20.0, max: 25.0, color: '#ca8a04', fillColor: 'rgba(234, 179, 8, 0.18)' },
      { label: 'Kritik Limit / Kırmızı (> 25 PSI)', level: 'RED', min: 25.0, max: 40.0, color: '#dc2626', fillColor: 'rgba(239, 68, 68, 0.15)' }
    ]
  },
  oil_p: {
    minY: 40,
    maxY: 160,
    unit: 'PSI',
    bands: [
      { label: 'Normal / Yeşil (90 - 120 PSI)', level: 'GREEN', min: 0, max: 120.0, color: '#16a34a', fillColor: 'rgba(34, 197, 94, 0.10)' },
      { label: 'Sarı İkaz (120 - 135 PSI)', level: 'YELLOW', min: 120.0, max: 135.0, color: '#ca8a04', fillColor: 'rgba(234, 179, 8, 0.18)' },
      { label: 'Kritik Limit / Kırmızı (> 135 PSI)', level: 'RED', min: 135.0, max: 180.0, color: '#dc2626', fillColor: 'rgba(239, 68, 68, 0.15)' }
    ]
  },
  oil_t: {
    minY: 20,
    maxY: 130,
    unit: '°C',
    bands: [
      { label: 'Normal / Yeşil (< 105 °C)', level: 'GREEN', min: 0, max: 105.0, color: '#16a34a', fillColor: 'rgba(34, 197, 94, 0.10)' },
      { label: 'Sarı İkaz (105 - 110 °C)', level: 'YELLOW', min: 105.0, max: 110.0, color: '#ca8a04', fillColor: 'rgba(234, 179, 8, 0.18)' },
      { label: 'Kritik Limit / Kırmızı (> 110 °C)', level: 'RED', min: 110.0, max: 140.0, color: '#dc2626', fillColor: 'rgba(239, 68, 68, 0.15)' }
    ]
  }
};

/**
 * Örnek AT-802F Uçuş Verisi Üreteci (X Ekseni: 15:59:58.35 Zaman Formatı)
 */
export const generateSampleFlightData = (paramKey: string) => {
  const baseSeconds = 15 * 3600 + 59 * 60 + 58.35; // 15:59:58.35
  const count = 350;
  const list = [];

  for (let i = 0; i < count; i++) {
    const timeSec = baseSeconds + i * 0.1;
    const h = Math.floor(timeSec / 3600);
    const m = Math.floor((timeSec % 3600) / 60);
    const s = Math.floor(timeSec % 60);
    const cs = Math.floor((timeSec % 1) * 100);
    const timeStr = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}.${String(cs).padStart(2, '0')}`;

    const p = i / count;
    let val = 0;
    let severity: 'NORMAL' | 'YELLOW' | 'RED' = 'NORMAL';

    if (paramKey === 'ng') {
      if (p < 0.2) val = 62.5 + Math.sin(i * 0.1) * 0.3;
      else if (p < 0.45) val = 62.5 + ((p - 0.2) / 0.25) * 40.5;
      else if (p < 0.58) { val = 104.1 + ((p - 0.45) / 0.13) * 1.3; severity = 'YELLOW'; }
      else if (p < 0.65) { val = 105.3 + Math.sin(((p - 0.58) / 0.07) * Math.PI) * 0.5; severity = val > 105.5 ? 'RED' : 'YELLOW'; }
      else if (p < 0.8) { val = 104.8 - ((p - 0.65) / 0.15) * 4.2; severity = val > 104.0 ? 'YELLOW' : 'NORMAL'; }
      else val = 98.4 + Math.sin(i * 0.05) * 0.4;
    } else if (paramKey === 'itt') {
      if (p < 0.2) val = 580 + Math.sin(i * 0.1) * 5;
      else if (p < 0.45) val = 580 + ((p - 0.2) / 0.25) * 280;
      else if (p < 0.58) { val = 875 + ((p - 0.45) / 0.13) * 110; severity = 'YELLOW'; }
      else if (p < 0.65) { val = 995 + Math.sin(((p - 0.58) / 0.07) * Math.PI) * 28; severity = val > 1000 ? 'RED' : 'YELLOW'; }
      else if (p < 0.8) { val = 980 - ((p - 0.65) / 0.15) * 160; severity = val > 870 ? 'YELLOW' : 'NORMAL'; }
      else val = 790 + Math.sin(i * 0.05) * 8;
    } else if (paramKey === 'np') {
      if (p < 0.2) val = 1050 + Math.sin(i * 0.1) * 12;
      else if (p < 0.45) val = 1050 + ((p - 0.2) / 0.25) * 620;
      else if (p < 0.58) { val = 1710 + ((p - 0.45) / 0.13) * 140; severity = 'YELLOW'; }
      else if (p < 0.65) { val = 1860 + Math.sin(((p - 0.58) / 0.07) * Math.PI) * 55; severity = val > 1870 ? 'RED' : 'YELLOW'; }
      else if (p < 0.8) { val = 1850 - ((p - 0.65) / 0.15) * 160; severity = val > 1700 ? 'YELLOW' : 'NORMAL'; }
      else val = 1680 + Math.sin(i * 0.05) * 10;
    } else if (paramKey === 'torq') {
      if (p < 0.2) val = 1100 + Math.sin(i * 0.1) * 35;
      else if (p < 0.45) val = 1100 + ((p - 0.2) / 0.25) * 4100;
      else if (p < 0.58) { val = 5300 + ((p - 0.45) / 0.13) * 700; severity = 'YELLOW'; }
      else if (p < 0.65) { val = 6050 + Math.sin(((p - 0.58) / 0.07) * Math.PI) * 230; severity = val > 6092 ? 'RED' : 'YELLOW'; }
      else if (p < 0.8) { val = 5900 - ((p - 0.65) / 0.15) * 900; severity = val > 5252 ? 'YELLOW' : 'NORMAL'; }
      else val = 4750 + Math.sin(i * 0.05) * 35;
    } else if (paramKey === 'fuel_p') {
      if (p < 0.45) val = 16.5 + Math.sin(i * 0.2) * 1.2;
      else if (p < 0.58) { val = 21.0 + Math.sin(i * 0.3) * 1.8; severity = 'YELLOW'; }
      else if (p < 0.65) { val = 24.5 + Math.sin(((p - 0.58) / 0.07) * Math.PI) * 2.2; severity = val > 25 ? 'RED' : 'YELLOW'; }
      else val = 18.5 + Math.sin(i * 0.1) * 1.0;
    } else if (paramKey === 'oil_p') {
      if (p < 0.45) val = 95 + Math.sin(i * 0.2) * 7;
      else if (p < 0.58) { val = 122 + Math.sin(i * 0.3) * 7; severity = 'YELLOW'; }
      else if (p < 0.65) { val = 132 + Math.sin(((p - 0.58) / 0.07) * Math.PI) * 7; severity = val > 135 ? 'RED' : 'YELLOW'; }
      else val = 108 + Math.sin(i * 0.1) * 4;
    } else if (paramKey === 'oil_t') {
      if (p < 0.2) val = 70 + Math.sin(i * 0.1) * 1.5;
      else if (p < 0.45) val = 70 + ((p - 0.2) / 0.25) * 32;
      else if (p < 0.58) { val = 105.5 + ((p - 0.45) / 0.13) * 3.5; severity = 'YELLOW'; }
      else if (p < 0.65) { val = 109.5 + Math.sin(((p - 0.58) / 0.07) * Math.PI) * 2.8; severity = val > 110 ? 'RED' : 'YELLOW'; }
      else if (p < 0.8) { val = 109.0 - ((p - 0.65) / 0.15) * 8; severity = val > 105 ? 'YELLOW' : 'NORMAL'; }
      else val = 94 + Math.sin(i * 0.05) * 1.5;
    }

    list.push({
      originalIndex: i + 1,
      index: i,
      timeStr,
      timeSec,
      val: parseFloat(val.toFixed(1)),
      severity
    });
  }
  return list;
};

/**
 * AT-802F Yangın Söndürme Uçağı SVG İmleci
 */
export const At802fPlaneIcon = ({ className = 'w-9 h-9', heading = 0 }: { className?: string; heading?: number }) => (
  <svg
    viewBox="0 0 64 64"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
    className={className}
    style={{ transform: `rotate(${heading}deg)` }}
  >
    <filter id="planeShadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="2" stdDeviation="2" floodColor="#000" floodOpacity="0.4" />
    </filter>
    {/* Gövde - Sarı AT-802 Gövdesi */}
    <g filter="url(#planeShadow)">
      {/* Kanatlar */}
      <path
        d="M 6 36 L 28 32 L 32 30 L 36 32 L 58 36 L 56 40 L 35 36 L 33 52 L 42 56 L 41 58 L 32 56 L 23 58 L 22 56 L 31 52 L 29 36 L 8 40 Z"
        fill="#facc15"
        stroke="#854d0e"
        strokeWidth="1.5"
      />
      {/* Kanat Uçları Kırmızı Şerit */}
      <polygon points="6,36 12,35 13,39 8,40" fill="#dc2626" />
      <polygon points="58,36 52,35 51,39 56,40" fill="#dc2626" />
      {/* Gövde Ana Hat */}
      <path
        d="M 32 6 C 30 6 28 9 28 16 L 29 32 L 29 50 L 32 58 L 35 50 L 35 32 L 36 16 C 36 9 34 6 32 6 Z"
        fill="#fef08a"
        stroke="#854d0e"
        strokeWidth="1.5"
      />
      {/* Kokpit Camı */}
      <path
        d="M 30 18 C 30 15 34 15 34 18 L 34 26 C 34 28 30 28 30 26 Z"
        fill="#0284c7"
        stroke="#0369a1"
        strokeWidth="1"
      />
      {/* Pervane Burun Konisi (PT6A Motor) */}
      <ellipse cx="32" cy="7" rx="3.5" ry="2.5" fill="#1e293b" />
      <line x1="24" y1="7" x2="40" y2="7" stroke="#cbd5e1" strokeWidth="1.5" strokeLinecap="round" />
      {/* Su Tankı Kapağı / Yangın Hopper */}
      <rect x="30" y="32" width="4" height="8" rx="1" fill="#ea580c" />
    </g>
  </svg>
);

export const FlightEngineChartModal: React.FC<FlightEngineChartModalProps> = ({
  isOpen,
  onClose,
  rows,
  headers,
  episodes,
  initialParamKey = 'ng',
  tailLabel,
  motorSn
}) => {
  const [activeParamKey, setActiveParamKey] = useState<string>(initialParamKey);
  const [isPlaying, setIsPlaying] = useState<boolean>(false);
  const [playSpeed, setPlaySpeed] = useState<number>(2); // 1x, 2x, 5x, 10x, 25x
  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [currentEpisodeIndex, setCurrentEpisodeIndex] = useState<number>(0);

  // Yakınlaştırma (Zoom) & Kaydırma (Pan) State'leri: [startIndex, endIndex]
  const [zoomRange, setZoomRange] = useState<[number, number] | null>(null);
  const [hoverIndex, setHoverIndex] = useState<number | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);

  const containerRef = useRef<HTMLDivElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);
  const animationFrameRef = useRef<number | null>(null);
  const lastTimeRef = useRef<number | null>(null);
  const dragStartRef = useRef<{ clientX: number; startIndex: number; endIndex: number } | null>(null);

  // Initial param key güncellendiğinde senkronize et
  useEffect(() => {
    if (initialParamKey) {
      setActiveParamKey(initialParamKey);
      setZoomRange(null); // Parametre değiştiğinde tam görünüm
    }
  }, [initialParamKey]);

  // Ekran görüntüsü al
  const handleScreenshot = async () => {
    if (containerRef.current) {
      const canvas = await html2canvas(containerRef.current);
      const dataUrl = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      link.href = dataUrl;
      link.download = `ucuş-analizi-${new Date().getTime()}.png`;
      link.click();
    }
  };



  // Parametrenin Excel sütun indeksini bul
  const paramColIdx = useMemo(() => {
    if (!headers || headers.length === 0) return 2;
    const cfg = YELLOW_FOCUS_COLUMNS.find(c => c.key === activeParamKey);
    if (!cfg) return 2;

    const idx = headers.findIndex((h, i) => cfg.matchPattern.test(h) || i === cfg.expectedIndex);
    return idx >= 0 ? idx : cfg.expectedIndex;
  }, [headers, activeParamKey]);

  // Tüm satırlardan nümerik değer ve zaman dizisi oluştur
  const isUsingSampleData = !rows || rows.length === 0;

  const chartData = useMemo(() => {
    if (isUsingSampleData) {
      return generateSampleFlightData(activeParamKey);
    }
    
    const raw = rows.map((r, i) => {
      const cell = r.cells[paramColIdx] || '0';
      const clean = parseFloat(String(cell).replace(',', '.').replace(/[^\d.-]/g, '')) || 0;
      return {
        originalIndex: r.originalIndex,
        index: i,
        timeStr: r.timeStr || `Satır ${r.originalIndex}`,
        timeSec: r.timeSec != null ? r.timeSec : i,
        val: clean,
        severity: r.severity
      };
    });

    return raw;
  }, [rows, paramColIdx, activeParamKey, isUsingSampleData]);

  // Görüntülenen X penceresi sınırları [viewStart, viewEnd]
  const [viewStart, viewEnd] = useMemo(() => {
    if (!chartData || chartData.length === 0) return [0, 0];
    if (!zoomRange) return [0, chartData.length - 1];
    const s = Math.max(0, Math.min(chartData.length - 5, Math.round(zoomRange[0])));
    const e = Math.min(chartData.length - 1, Math.max(s + 4, Math.round(zoomRange[1])));
    return [s, e];
  }, [chartData, zoomRange]);

  const visibleCount = viewEnd - viewStart + 1;
  const zoomFactor = chartData.length > 0 ? (chartData.length / Math.max(1, visibleCount)).toFixed(1) : '1.0';
  const isZoomed = zoomRange !== null && visibleCount < chartData.length - 2;

  // İlgili parametreye ait limit bandı konfigürasyonu
  const limitConfig = useMemo(() => {
    return PARAM_LIMITS_MAP[activeParamKey] || PARAM_LIMITS_MAP['ng'];
  }, [activeParamKey]);

  // Verinin gerçek min ve max değerlerini hesapla
  const { dataMin, dataMax, peakItem } = useMemo(() => {
    if (chartData.length === 0) return { dataMin: 0, dataMax: 100, peakItem: null };
    let min = Infinity;
    let max = -Infinity;
    let peak = chartData[0];

    chartData.forEach(d => {
      if (d.val < min) min = d.val;
      if (d.val > max) {
        max = d.val;
        peak = d;
      }
    });

    return {
      dataMin: min === Infinity ? 0 : min,
      dataMax: max === -Infinity ? 100 : max,
      peakItem: peak
    };
  }, [chartData]);

  // Y Ekseninin sınırları
  const yDomainMin = useMemo(() => {
    return Math.min(limitConfig.minY, Math.floor(dataMin * 0.95));
  }, [limitConfig.minY, dataMin]);

  const yDomainMax = useMemo(() => {
    return Math.max(limitConfig.maxY, Math.ceil(dataMax * 1.05));
  }, [limitConfig.maxY, dataMax]);

  // Aktif parametreye ait aşım olayları (episodes)
  const activeEpisodes = useMemo(() => {
    if (!isUsingSampleData && episodes && episodes.length > 0) {
      return episodes
        .filter(e => e.colKey === activeParamKey)
        .sort((a, b) => a.startRow - b.startRow);
    }
    if (isUsingSampleData) {
      return [
        {
          id: 'sample-yellow-1',
          colKey: activeParamKey,
          colName: YELLOW_FOCUS_COLUMNS.find(c => c.key === activeParamKey)?.name || activeParamKey.toUpperCase(),
          level: 'YELLOW' as const,
          startRow: 160,
          endRow: 230,
          startTimeStr: '16:00:14.35',
          endTimeStr: '16:00:21.35',
          durationFormatted: '7.0 saniye',
          peakValueFormatted: activeParamKey === 'ng' ? '%105.1' : 'Sarı İkaz'
        },
        {
          id: 'sample-red-1',
          colKey: activeParamKey,
          colName: YELLOW_FOCUS_COLUMNS.find(c => c.key === activeParamKey)?.name || activeParamKey.toUpperCase(),
          level: 'RED' as const,
          startRow: 205,
          endRow: 228,
          startTimeStr: '16:00:18.85',
          endTimeStr: '16:00:21.15',
          durationFormatted: '2.3 saniye',
          peakValueFormatted: activeParamKey === 'ng' ? '%105.7' : 'Kritik Tepe'
        }
      ];
    }
    return [];
  }, [episodes, activeParamKey, isUsingSampleData]);

  // SVG Görsel Alan Boyutları
  const svgWidth = 1000;
  const svgHeight = 440;
  const padding = { top: 30, right: 40, bottom: 50, left: 75 };
  const plotWidth = svgWidth - padding.left - padding.right;
  const plotHeight = svgHeight - padding.top - padding.bottom;

  // Koordinat Dönüşüm Fonksiyonları (Zoomed Windowing)
  const getX = useCallback((idx: number) => {
    const range = viewEnd - viewStart || 1;
    const ratio = (idx - viewStart) / range;
    return padding.left + ratio * plotWidth;
  }, [viewStart, viewEnd, plotWidth, padding.left]);

  const getY = useCallback((val: number) => {
    const range = yDomainMax - yDomainMin || 1;
    const clamped = Math.max(yDomainMin, Math.min(yDomainMax, val));
    return padding.top + plotHeight - ((clamped - yDomainMin) / range) * plotHeight;
  }, [yDomainMin, yDomainMax, plotHeight, padding.top]);

  // SVG Polyline / Path noktaları (Görünür aralığa göre yüksek çözünürlüklü)
  const { pathD, visibleSlice } = useMemo(() => {
    if (chartData.length === 0) return { pathD: '', visibleSlice: [] };

    const start = Math.max(0, viewStart - 2);
    const end = Math.min(chartData.length - 1, viewEnd + 2);
    const slice = chartData.slice(start, end + 1);

    // Zoom durumuna göre çizgi noktası
    const maxRenderPoints = 1200;
    const step = Math.max(1, Math.floor(slice.length / maxRenderPoints));
    const points: { x: number; y: number; item: typeof chartData[0]; idx: number }[] = [];

    for (let i = 0; i < slice.length; i += step) {
      const item = slice[i];
      points.push({
        x: getX(item.index),
        y: getY(item.val),
        item,
        idx: item.index
      });
    }

    // Son noktayı ekle
    if (points[points.length - 1]?.idx !== slice[slice.length - 1]?.index) {
      const last = slice[slice.length - 1];
      points.push({
        x: getX(last.index),
        y: getY(last.val),
        item: last,
        idx: last.index
      });
    }

    const d = points.reduce((acc, p, idx) => {
      return idx === 0 ? `M ${p.x.toFixed(1)} ${p.y.toFixed(1)}` : `${acc} L ${p.x.toFixed(1)} ${p.y.toFixed(1)}`;
    }, '');

    return { pathD: d, visibleSlice: slice };
  }, [chartData, viewStart, viewEnd, getX, getY]);

  // Anlık Uçak İmleç Noktası
  const currentItem = chartData[currentIndex] || chartData[0];
  const currentPlaneX = getX(currentIndex);
  const currentPlaneY = currentItem ? getY(currentItem.val) : getY(0);

  // Uçağın uçuş açısı (slope / heading)
  const currentHeading = useMemo(() => {
    if (!chartData || chartData.length < 2) return 0;
    
    // Yatay uçuş zorlaması (eğer dikey değer sıfırdan farklıysa burnu 90 derece sağa döndür)
    if (currentItem.val !== 0) return 90;

    const nextIdx = Math.min(chartData.length - 1, currentIndex + 3);
    const prevIdx = Math.max(0, currentIndex - 3);
    if (nextIdx === prevIdx) return 0;

    const dx = getX(nextIdx) - getX(prevIdx);
    const dy = getY(chartData[nextIdx].val) - getY(chartData[prevIdx].val);
    const angleRad = Math.atan2(dy, dx);
    const angleDeg = (angleRad * 180) / Math.PI;

    return Math.max(-45, Math.min(45, angleDeg));
  }, [chartData, currentIndex, getX, getY, currentItem.val]);

  // Oynatma Sırasında Otomatik Takip (Uçak görünür alanın dışına çıkarsa görünümü kaydır)
  useEffect(() => {
    if (isPlaying && zoomRange) {
      const windowSize = viewEnd - viewStart;
      if (currentIndex >= viewEnd) {
        const nextStart = Math.min(chartData.length - 1 - windowSize, currentIndex - Math.floor(windowSize * 0.2));
        const nextEnd = nextStart + windowSize;
        setZoomRange([nextStart, nextEnd]);
      } else if (currentIndex < viewStart) {
        const nextStart = Math.max(0, currentIndex - Math.floor(windowSize * 0.1));
        const nextEnd = nextStart + windowSize;
        setZoomRange([nextStart, nextEnd]);
      }
    }
  }, [currentIndex, isPlaying, zoomRange, viewStart, viewEnd, chartData.length]);

  // Oynatma Animasyon Döngüsü (requestAnimationFrame)
  useEffect(() => {
    if (!isPlaying) {
      if (animationFrameRef.current) cancelAnimationFrame(animationFrameRef.current);
      lastTimeRef.current = null;
      return;
    }

    const animate = (time: number) => {
      if (lastTimeRef.current != null) {
        const deltaMs = time - lastTimeRef.current;
        const advance = (deltaMs / 1000) * 15 * playSpeed;
        
        setCurrentIndex(prev => {
          const next = prev + advance;
          if (next >= chartData.length - 1) {
            setIsPlaying(false);
            return chartData.length - 1;
          }
          return Math.floor(next);
        });
      }
      lastTimeRef.current = time;
      animationFrameRef.current = requestAnimationFrame(animate);
    };

    animationFrameRef.current = requestAnimationFrame(animate);

    return () => {
      if (animationFrameRef.current) cancelAnimationFrame(animationFrameRef.current);
    };
  }, [isPlaying, playSpeed, chartData.length]);

  // =========================================================================
  // FARE İLE YAKINLAŞMA (ZOOM TO MOUSE POSITION) & SÜRÜKLEME (PAN)
  // =========================================================================

  const handleWheel = useCallback((e: React.WheelEvent<SVGSVGElement>) => {
    e.preventDefault();
    if (chartData.length <= 5) return;

    const rect = e.currentTarget.getBoundingClientRect();
    const mouseX = e.clientX - rect.left;
    const plotLeftPx = (padding.left / svgWidth) * rect.width;
    const plotWidthPx = (plotWidth / svgWidth) * rect.width;
    const ratio = Math.max(0, Math.min(1, (mouseX - plotLeftPx) / plotWidthPx));

    const currentWindow = viewEnd - viewStart;
    const centerIdx = viewStart + ratio * currentWindow;

    // deltaY < 0 ise yakınlaş (zoom in), deltaY > 0 ise uzaklaş (zoom out)
    const factor = e.deltaY < 0 ? 0.1 : 10.0;
    let newWindow = currentWindow * factor;

    // Minimum 10 nokta, maksimum tüm veri
    newWindow = Math.max(10, Math.min(chartData.length - 1, newWindow));

    let newStart = Math.round(centerIdx - ratio * newWindow);
    let newEnd = Math.round(newStart + newWindow);

    if (newStart < 0) {
      newStart = 0;
      newEnd = Math.min(chartData.length - 1, Math.round(newWindow));
    }
    if (newEnd > chartData.length - 1) {
      newEnd = chartData.length - 1;
      newStart = Math.max(0, Math.round(newEnd - newWindow));
    }

    if (newWindow >= chartData.length - 1) {
      setZoomRange(null); // Tam uçuş
    } else {
      setZoomRange([newStart, newEnd]);
    }
  }, [chartData.length, viewStart, viewEnd, padding.left, plotWidth, svgWidth]);

  // Fare Sürükleme Başlangıcı
  const handleMouseDown = (e: React.MouseEvent<SVGSVGElement>) => {
    if (e.button !== 0) return;
    setIsDragging(true);
    dragStartRef.current = {
      clientX: e.clientX,
      startIndex: viewStart,
      endIndex: viewEnd
    };
  };

  // Fare Hareketi: Hem Pan Kaydırma Hem de İmleç Detay Gösterimi (Crosshair)
  const handleMouseMove = (e: React.MouseEvent<SVGSVGElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const mouseX = e.clientX - rect.left;
    const plotLeftPx = (padding.left / svgWidth) * rect.width;
    const plotWidthPx = (plotWidth / svgWidth) * rect.width;
    const ratio = Math.max(0, Math.min(1, (mouseX - plotLeftPx) / plotWidthPx));
    const hIdx = Math.round(viewStart + ratio * (viewEnd - viewStart));

    if (hIdx >= 0 && hIdx < chartData.length) {
      setHoverIndex(hIdx);
    }

    if (isDragging && dragStartRef.current) {
      const deltaX = e.clientX - dragStartRef.current.clientX;
      const windowSize = dragStartRef.current.endIndex - dragStartRef.current.startIndex;
      const pointsPerPixel = windowSize / plotWidthPx;
      const shift = Math.round(-deltaX * pointsPerPixel);

      let nextStart = dragStartRef.current.startIndex + shift;
      let nextEnd = dragStartRef.current.endIndex + shift;

      if (nextStart < 0) {
        nextStart = 0;
        nextEnd = windowSize;
      }
      if (nextEnd > chartData.length - 1) {
        nextEnd = chartData.length - 1;
        nextStart = Math.max(0, nextEnd - windowSize);
      }
      setZoomRange([nextStart, nextEnd]);
    }
  };

  const handleMouseUp = () => {
    setIsDragging(false);
    dragStartRef.current = null;
  };

  const handleMouseLeave = () => {
    setIsDragging(false);
    dragStartRef.current = null;
    setHoverIndex(null);
  };

  // Çift Tıklama: Çift tıklanan alana 3 kat yakınlaş veya sıfırla
  const handleDoubleClick = (e: React.MouseEvent<SVGSVGElement>) => {
    if (isZoomed) {
      setZoomRange(null); // Sıfırla
    } else {
      const rect = e.currentTarget.getBoundingClientRect();
      const mouseX = e.clientX - rect.left;
      const plotLeftPx = (padding.left / svgWidth) * rect.width;
      const plotWidthPx = (plotWidth / svgWidth) * rect.width;
      const ratio = Math.max(0, Math.min(1, (mouseX - plotLeftPx) / plotWidthPx));
      const centerIdx = Math.round(ratio * (chartData.length - 1));
      const targetWindow = Math.max(15, Math.floor(chartData.length * 0.25));
      const newStart = Math.max(0, centerIdx - Math.floor(targetWindow / 2));
      const newEnd = Math.min(chartData.length - 1, newStart + targetWindow);
      setZoomRange([newStart, newEnd]);
    }
  };

  // Butonla Manuel Yakınlaştırma / Uzaklaştırma
  const handleZoomInBtn = () => {
    const currentWindow = viewEnd - viewStart;
    const centerIdx = Math.round((viewStart + viewEnd) / 2);
    const newWindow = Math.max(10, Math.round(currentWindow * 0.6));
    const newStart = Math.max(0, centerIdx - Math.floor(newWindow / 2));
    const newEnd = Math.min(chartData.length - 1, newStart + newWindow);
    setZoomRange([newStart, newEnd]);
  };

  const handleZoomOutBtn = () => {
    const currentWindow = viewEnd - viewStart;
    const centerIdx = Math.round((viewStart + viewEnd) / 2);
    const newWindow = Math.min(chartData.length - 1, Math.round(currentWindow * 1.6));
    const newStart = Math.max(0, centerIdx - Math.floor(newWindow / 2));
    const newEnd = Math.min(chartData.length - 1, newStart + newWindow);
    if (newWindow >= chartData.length - 1) {
      setZoomRange(null);
    } else {
      setZoomRange([newStart, newEnd]);
    }
  };

  // Aşım / İkaz Odağına Otomatik Yakınlaş
  const handleFocusOnExceedance = () => {
    if (activeEpisodes.length === 0) return;
    const nextIndex = (currentEpisodeIndex + 1) % activeEpisodes.length;
    setCurrentEpisodeIndex(nextIndex);
    const ep = activeEpisodes[nextIndex];
    const margin = Math.max(15, Math.round((ep.endRow - ep.startRow) * 1.5));
    const start = Math.max(0, ep.startRow - margin);
    const end = Math.min(chartData.length - 1, ep.endRow + margin);
    setZoomRange([start, end]);
    setCurrentIndex(ep.startRow);
  };

  // Tıklamayla Uçak İmlecini Taşıma (Seek)
  const handleSvgClick = (e: React.MouseEvent<SVGSVGElement>) => {
    if (isDragging) return;
    const rect = e.currentTarget.getBoundingClientRect();
    const clickX = e.clientX - rect.left;
    const plotLeftPx = (padding.left / svgWidth) * rect.width;
    const plotWidthPx = (plotWidth / svgWidth) * rect.width;
    const ratio = Math.max(0, Math.min(1, (clickX - plotLeftPx) / plotWidthPx));
    const targetIdx = Math.round(viewStart + ratio * (viewEnd - viewStart));
    if (targetIdx >= 0 && targetIdx < chartData.length) {
      setCurrentIndex(targetIdx);
    }
  };

  // Sonraki / Önceki İkaza Atlama
  const jumpToNextEpisode = () => {
    if (activeEpisodes.length === 0) return;
    const nextEp = activeEpisodes.find(ep => ep.startRow > (currentItem?.originalIndex || 0));
    if (nextEp) {
      const idx = chartData.findIndex(d => d.originalIndex >= nextEp.startRow);
      if (idx >= 0) {
        setCurrentIndex(idx);
        // Eğer görünür alanda değilse kaydır
        if (idx < viewStart || idx > viewEnd) {
          const w = viewEnd - viewStart;
          setZoomRange([Math.max(0, idx - Math.floor(w * 0.2)), Math.min(chartData.length - 1, idx + Math.floor(w * 0.8))]);
        }
      }
    } else {
      const idx = chartData.findIndex(d => d.originalIndex >= activeEpisodes[0].startRow);
      if (idx >= 0) setCurrentIndex(idx);
    }
  };

  const jumpToPrevEpisode = () => {
    if (activeEpisodes.length === 0) return;
    const reversed = [...activeEpisodes].reverse();
    const prevEp = reversed.find(ep => ep.endRow < (currentItem?.originalIndex || 0));
    if (prevEp) {
      const idx = chartData.findIndex(d => d.originalIndex >= prevEp.startRow);
      if (idx >= 0) {
        setCurrentIndex(idx);
        if (idx < viewStart || idx > viewEnd) {
          const w = viewEnd - viewStart;
          setZoomRange([Math.max(0, idx - Math.floor(w * 0.2)), Math.min(chartData.length - 1, idx + Math.floor(w * 0.8))]);
        }
      }
    } else {
      const last = activeEpisodes[activeEpisodes.length - 1];
      const idx = chartData.findIndex(d => d.originalIndex >= last.startRow);
      if (idx >= 0) setCurrentIndex(idx);
    }
  };

  if (!isOpen) return null;

  const currentStatus = currentItem
    ? currentItem.val >= (limitConfig.bands.find(b => b.level === 'RED')?.min || 9999)
      ? { label: '🔴 KRİTİK LİMİT AŞIMI', color: 'bg-red-600 text-white border-red-700 shadow-sm' }
      : currentItem.val >= (limitConfig.bands.find(b => b.level === 'YELLOW')?.min || 9999)
      ? { label: '🟡 SARI İKAZ BÖLGESİ', color: 'bg-amber-400 text-slate-950 border-amber-500 shadow-sm' }
      : { label: '🟢 NORMAL ÇALIŞMA', color: 'bg-emerald-600 text-white border-emerald-700 shadow-sm' }
    : { label: 'NORMAL', color: 'bg-slate-200 text-slate-700' };

  // Görünür X ekseni zaman tıkları (6-8 adet temiz aralık)
  const xTicks = [0, 0.16, 0.33, 0.5, 0.66, 0.83, 1].map((r) => {
    const idx = Math.min(chartData.length - 1, Math.round(viewStart + r * (viewEnd - viewStart)));
    const item = chartData[idx];
    return {
      idx,
      xPos: getX(idx),
      timeStr: item ? item.timeStr : ''
    };
  });

  const activeHoverItem = hoverIndex !== null ? chartData[hoverIndex] : null;

  return (
    <div className="fixed inset-0 z-[12000] flex items-center justify-center p-2 sm:p-4 bg-slate-950/70 backdrop-blur-md animate-fade-in font-sans">
      <div className="bg-white text-slate-800 w-full max-w-7xl h-[95vh] rounded-3xl shadow-2xl border border-slate-300 flex flex-col overflow-hidden animate-scale-up">
        
        {/* ========================================================================= */}
        {/* ÜST BAŞLIK BÖLÜMÜ (BEYAZ TEMA & TEMİZ HAVACILIK TASARIMI) */}
        {/* ========================================================================= */}
        <div className="px-6 py-3.5 bg-gradient-to-r from-slate-100 via-white to-slate-100 border-b border-slate-200 flex items-center justify-between shrink-0 shadow-xs">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-amber-50 border border-amber-300 flex items-center justify-center text-amber-600 shadow-sm">
              <At802fPlaneIcon className="w-7 h-7" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-black tracking-widest text-amber-900 uppercase px-2 py-0.5 rounded bg-amber-100 border border-amber-300">
                  AT-802F UÇUŞ GRAFİĞİ & OYNATICI
                </span>
                <span className="text-xs text-slate-600 font-mono font-bold">
                  {tailLabel} • Motor S/N: {motorSn}
                </span>
                {isUsingSampleData && (
                  <span className="text-[10px] bg-blue-50 text-blue-800 border border-blue-300 px-2 py-0.5 rounded font-mono font-bold flex items-center gap-1 shadow-2xs">
                    <span>Örnek Demo Uçuş Verisi (15:59:58.35)</span>
                  </span>
                )}
              </div>
              <h2 className="text-base sm:text-lg font-black tracking-tight text-slate-900 mt-0.5 flex items-center gap-2">
                <span>Motor Parametre Zaman-Grafik Analizi</span>
                <span className="text-xs font-mono font-bold text-slate-500 font-normal">
                  (X: Uçuş Zamanı, Y: Eşik Limitleri)
                </span>
              </h2>
            </div>
          </div>

          <div className="flex items-center gap-3">
            {/* Anlık Durum Rozeti */}
            <div className={`px-3 py-1 rounded-xl text-xs font-black uppercase tracking-wider border flex items-center gap-1.5 ${currentStatus.color}`}>
              <Activity className="w-3.5 h-3.5 animate-pulse" />
              <span>{currentStatus.label}</span>
            </div>

            {/* Kapat Butonu */}
            <button
              type="button"
              onClick={() => {
                setIsPlaying(false);
                onClose();
              }}
              className="w-9 h-9 rounded-xl bg-slate-100 hover:bg-red-600 hover:text-white text-slate-600 transition-all flex items-center justify-center cursor-pointer border border-slate-300 shadow-2xs"
              title="Grafiği Kapat"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* 7 PARAMETRE SEKMELERİ (HIZLI GEÇİŞ) */}
        <div className="px-6 py-2 bg-slate-50 border-b border-slate-200 flex items-center gap-2 overflow-x-auto shrink-0 select-none">
          <span className="text-xs font-bold text-slate-600 flex items-center gap-1 mr-1">
            <Sliders className="w-3.5 h-3.5 text-amber-600" />
            <span>Parametre:</span>
          </span>
          {YELLOW_FOCUS_COLUMNS.map(col => {
            const isSelected = activeParamKey === col.key;
            const epsCount = episodes.filter(e => e.colKey === col.key).length;
            return (
              <button
                key={col.key}
                type="button"
                onClick={() => {
                  setActiveParamKey(col.key);
                  setZoomRange(null);
                }}
                className={`px-3 py-1.5 rounded-xl font-bold text-xs font-mono transition-all flex items-center gap-1.5 cursor-pointer border ${
                  isSelected
                    ? 'bg-amber-500 text-slate-950 font-black border-amber-600 shadow-sm scale-102 ring-2 ring-amber-300'
                    : 'bg-white hover:bg-slate-100 text-slate-700 border-slate-300 shadow-2xs'
                }`}
              >
                <span className="text-amber-800">[{col.excelLetter}]</span>
                <span>{col.name}</span>
                {epsCount > 0 && (
                  <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-black ${
                    isSelected ? 'bg-amber-950 text-amber-200' : 'bg-red-600 text-white'
                  }`}>
                    {epsCount}
                  </span>
                )}
              </button>
            );
          })}
        </div>

        {/* YAKINLAŞTIRMA (ZOOM) VE ANALİZ ARAÇ ÇUBUĞU */}
        <div className="px-6 py-1.5 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-2 shrink-0 select-none text-xs">
          <div className="flex items-center gap-2">
            <span className="font-bold text-slate-700 flex items-center gap-1">
              <ZoomIn className="w-4 h-4 text-sky-600" />
              <span>Yakınlaştırma:</span>
            </span>

            <button
              type="button"
              onClick={handleZoomInBtn}
              className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold border border-slate-300 flex items-center gap-1 cursor-pointer transition-all active:scale-95 shadow-2xs"
              title="Grafiğe Yakınlaş (+)"
            >
              <ZoomIn className="w-3.5 h-3.5 text-slate-700" />
              <span>Yakınlaş (+)</span>
            </button>

            <button
              type="button"
              onClick={handleZoomOutBtn}
              className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold border border-slate-300 flex items-center gap-1 cursor-pointer transition-all active:scale-95 shadow-2xs"
              title="Grafikten Uzaklaş (-)"
            >
              <ZoomOut className="w-3.5 h-3.5 text-slate-700" />
              <span>Uzaklaş (-)</span>
            </button>

            {isZoomed && (
              <button
                type="button"
                onClick={() => setZoomRange(null)}
                className="px-2.5 py-1 rounded-lg bg-amber-100 hover:bg-amber-200 text-amber-900 font-black border border-amber-300 flex items-center gap-1 cursor-pointer transition-all active:scale-95 shadow-2xs"
                title="Tüm Uçuş Görünümüne Sıfırla"
              >
                <Maximize2 className="w-3.5 h-3.5 text-amber-800" />
                <span>↺ Tam Uçuş (Sıfırla)</span>
              </button>
            )}

            {activeEpisodes.length > 0 && (
              <button
                type="button"
                onClick={handleFocusOnExceedance}
                className="px-2.5 py-1 rounded-lg bg-yellow-100 hover:bg-yellow-200 text-yellow-900 font-bold border border-yellow-300 flex items-center gap-1 cursor-pointer transition-all active:scale-95 shadow-2xs"
                title="Aşım veya İkaz Olayına Otomatik Yakınlaş"
              >
                <Flame className="w-3.5 h-3.5 text-yellow-600" />
                <span>⚡ İkaz Odağı</span>
              </button>
            )}

            <span className="font-mono text-[11px] text-slate-500 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
              Büyütme: <strong className="text-slate-800">{zoomFactor}x</strong> ({visibleCount} Satır Görünür)
            </span>
          </div>

          <div className="flex items-center gap-1 text-[11px] text-slate-500 font-medium">
            <Info className="w-3.5 h-3.5 text-sky-600 shrink-0" />
            <span>
              <strong>Fare Tekerleği</strong> ile imlecin olduğu noktaya anında yakınlaşabilir, fareyle basılı tutarak sağa/sola kaydırabilirsiniz.
            </span>
          </div>
        </div>

        {/* ========================================================================= */}
        {/* ANA GRAFİK GÖVDE ALANI (BEYAZ ARKA PLAN & AT-802F İMLECİ) */}
        {/* ========================================================================= */}
        <div
          ref={containerRef}
          className="flex-1 p-3 bg-slate-50 flex flex-col justify-center items-center relative overflow-hidden select-none"
        >
          {/* Üst Bilgi Başlığı (Seçili Parametre, Birim ve Eşik Renk İpuçları) */}
          <div className="w-full max-w-5xl flex items-center justify-between mb-1.5 px-2 text-xs font-mono">
            <div className="flex items-center gap-2">
              <span className="text-slate-900 font-black text-sm">
                {YELLOW_FOCUS_COLUMNS.find(c => c.key === activeParamKey)?.name}
              </span>
              <span className="text-slate-500 font-bold">
                ({limitConfig.unit})
              </span>
              <span className="text-slate-300">|</span>
              <span className="text-slate-700">
                Uçuş Tepe Değeri: <strong className="text-red-700 font-black">{peakItem ? `${peakItem.val} ${limitConfig.unit}` : '-'}</strong> (Zaman: {peakItem?.timeStr})
              </span>
            </div>

            <div className="flex items-center gap-3">
              {limitConfig.bands.map((band, bIdx) => (
                <div key={bIdx} className="flex items-center gap-1.5 text-[11px] font-bold">
                  <span className="w-3 h-3 rounded-sm border shadow-2xs" style={{ backgroundColor: band.color }} />
                  <span className="text-slate-700">{band.label}</span>
                </div>
              ))}
            </div>
          </div>

          {/* SVG İnteraktif Grafik Çizimi (BEYAZ ARKA PLAN) */}
          <div className="w-full max-w-5xl bg-white rounded-2xl border-2 border-slate-300 p-2 shadow-sm relative flex-1 flex flex-col justify-center">
            <svg
              ref={svgRef}
              viewBox={`0 0 ${svgWidth} ${svgHeight}`}
              className="w-full h-full cursor-crosshair overflow-hidden"
              onWheel={handleWheel}
              onMouseDown={handleMouseDown}
              onMouseMove={handleMouseMove}
              onMouseUp={handleMouseUp}
              onMouseLeave={handleMouseLeave}
              onDoubleClick={handleDoubleClick}
              onClick={handleSvgClick}
            >
              {/* Tanımlar (Kırpma Alanı, Gradientler ve Gölgeler) */}
              <defs>
                <clipPath id="plotClip">
                  <rect
                    x={padding.left}
                    y={padding.top}
                    width={plotWidth}
                    height={plotHeight}
                  />
                </clipPath>

                <linearGradient id="lineGradient" x1="0" y1="1" x2="0" y2="0">
                  <stop offset="0%" stopColor="#16a34a" />
                  <stop offset="65%" stopColor="#ca8a04" />
                  <stop offset="100%" stopColor="#dc2626" />
                </linearGradient>

                <filter id="softShadow" x="-20%" y="-20%" width="140%" height="140%">
                  <feDropShadow dx="0" dy="2" stdDeviation="3" floodColor="#000" floodOpacity="0.15" />
                </filter>
              </defs>

              {/* BEYAZ GRAFİK TABANI */}
              <rect
                x={padding.left}
                y={padding.top}
                width={plotWidth}
                height={plotHeight}
                fill="#ffffff"
              />

              {/* 1. LİMİT ARALIKLARI RENKLİ BANTLARI (EN ÜSTTE KIRMIZI, ORTADA SARI, ALTA YEŞİL) */}
              {limitConfig.bands.map((band, bIdx) => {
                const isRed = band.level === 'RED';
                const isGreen = band.level === 'GREEN';
                const rawTopY = getY(band.max);
                const rawBottomY = getY(band.min);
                const bandTopY = isRed ? padding.top : Math.max(padding.top, Math.min(rawTopY, rawBottomY));
                const bandBottomY = isGreen ? padding.top + plotHeight : Math.min(padding.top + plotHeight, Math.max(rawTopY, rawBottomY));
                const rectHeight = Math.max(0, bandBottomY - bandTopY);

                return (
                  <g key={`band-${bIdx}`}>
                    <rect
                      x={padding.left}
                      y={bandTopY}
                      width={plotWidth}
                      height={rectHeight}
                      fill={band.fillColor}
                    />
                    {/* Eşik Çizgisi */}
                    <line
                      x1={padding.left}
                      y1={bandBottomY}
                      x2={padding.left + plotWidth}
                      y2={bandBottomY}
                      stroke={band.color}
                      strokeWidth="1.5"
                      strokeDasharray="4 4"
                      opacity="0.85"
                    />
                    {/* Eşik Değer Etiketi */}
                    <text
                      x={padding.left - 8}
                      y={bandBottomY + 4}
                      fill={band.color}
                      fontSize="10"
                      fontWeight="bold"
                      textAnchor="end"
                      fontFamily="monospace"
                    >
                      {band.min} {limitConfig.unit}
                    </text>
                  </g>
                );
              })}

              {/* 2. DİKEY IZGARA ÇİZGİLERİ (ZAMAN TIKLARI BOYUNCA) */}
              {xTicks.map((tick, tIdx) => (
                <line
                  key={`vgrid-${tIdx}`}
                  x1={tick.xPos}
                  y1={padding.top}
                  x2={tick.xPos}
                  y2={padding.top + plotHeight}
                  stroke="#f1f5f9"
                  strokeWidth="1"
                />
              ))}

              {/* 3. AŞIM EPİZOTLARI İŞARETLERİ (Kırpma Alanı İçinde) */}
              <g clipPath="url(#plotClip)">
                {activeEpisodes.map((ep, eIdx) => {
                  const startIdx = chartData.findIndex(d => d.originalIndex >= ep.startRow);
                  const endIdx = chartData.findIndex(d => d.originalIndex >= ep.endRow);
                  if (startIdx < 0 || endIdx < 0) return null;

                  const epX1 = getX(startIdx);
                  const epX2 = getX(endIdx);
                  const epWidth = Math.max(4, epX2 - epX1);
                  const isRedEp = ep.level === 'RED';
                  const sameLevelEpisodes = activeEpisodes.filter(e => e.level === ep.level);
                  const epNumber = sameLevelEpisodes.indexOf(ep) + 1;

                  return (
                    <g key={`ep-rect-${eIdx}`}>
                      {visibleCount < 100 && (
                        <>
                          <line x1={epX1} y1={padding.top} x2={epX1} y2={padding.top + plotHeight} stroke={isRedEp ? '#dc2626' : '#ca8a04'} strokeWidth="1.5" strokeDasharray="3 3" />
                          <line x1={epX2} y1={padding.top} x2={epX2} y2={padding.top + plotHeight} stroke={isRedEp ? '#dc2626' : '#ca8a04'} strokeWidth="1.5" strokeDasharray="3 3" />
                        </>
                      )}
                      {/* Olay Başlığı */}
                      <text
                        x={epX1 + 4}
                        y={padding.top + 14}
                        fill={isRedEp ? '#991b1b' : '#854d0e'}
                        fontSize="9"
                        fontWeight="900"
                        fontFamily="monospace"
                      >
                        {isRedEp ? '🔴 KIRMIZI' : '🟡 SARI'} #{epNumber}
                      </text>
                    </g>
                  );
                })}

                {/* 4. VERİ ÇİZGİSİ (Uçuş Eğrisi) */}
                {pathD && (
                  <path
                    d={pathD}
                    fill="none"
                    stroke="url(#lineGradient)"
                    strokeWidth="2.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                )}

                {/* 5. YAKINLAŞILDIĞINDA TEKİL VERİ NOKTALARI (Points - Sıkışıklığı Önleyen Detay) */}
                {visibleCount <= 120 && visibleSlice.map((item) => {
                  const x = getX(item.index);
                  const y = getY(item.val);
                  const isRed = item.severity === 'RED';
                  const isYellow = item.severity === 'YELLOW';
                  const isSelectedPlane = item.index === currentIndex;
                  const isHovered = item.index === hoverIndex;

                  return (
                    <circle
                      key={`pt-${item.index}`}
                      cx={x}
                      cy={y}
                      r={isSelectedPlane ? 6 : isHovered ? 5 : visibleCount <= 40 ? 4 : 2.5}
                      fill={isRed ? '#dc2626' : isYellow ? '#ca8a04' : '#0284c7'}
                      stroke="#ffffff"
                      strokeWidth={isSelectedPlane || isHovered ? 2 : 1}
                    />
                  );
                })}

                {/* 6. FARE GEZİNME VE İNCELEME DİKEY ÇİZGİSİ (Crosshair) */}
                {activeHoverItem && (
                  <g>
                    <line
                      x1={getX(activeHoverItem.index)}
                      y1={padding.top}
                      x2={getX(activeHoverItem.index)}
                      y2={padding.top + plotHeight}
                      stroke="#0284c7"
                      strokeWidth="1.5"
                      strokeDasharray="3 3"
                    />
                    <circle
                      cx={getX(activeHoverItem.index)}
                      cy={getY(activeHoverItem.val)}
                      r="5"
                      fill="#0284c7"
                      stroke="#ffffff"
                      strokeWidth="2"
                    />
                  </g>
                )}

                {/* 7. ANLIK OYNATMA ZAMAN ÇİZGİSİ */}
                <line
                  x1={currentPlaneX}
                  y1={padding.top}
                  x2={currentPlaneX}
                  y2={padding.top + plotHeight}
                  stroke="#ea580c"
                  strokeWidth="1.5"
                  strokeDasharray="2 2"
                  opacity="0.85"
                />

                {/* 8. GRAFİK İMLECİ: AT-802F UÇAK İKONU */}
                <g
                  transform={`translate(${currentPlaneX - 20}, ${currentPlaneY - 20})`}
                  className="transition-transform duration-75"
                >
                  <circle cx="20" cy="20" r="16" fill="rgba(245, 158, 11, 0.2)" stroke="#ea580c" strokeWidth="1" className="animate-ping" />
                  <foreignObject width="40" height="40">
                    <At802fPlaneIcon className="w-10 h-10 drop-shadow-md" heading={currentHeading} />
                  </foreignObject>
                </g>
              </g>

              {/* 9. Y EKSENİ VE X EKSENİ ANA ÇİZGİLERİ (Sadece yüksek yakınlaştırmada görünür) */}
              <line
                x1={padding.left}
                y1={padding.top}
                x2={padding.left}
                y2={padding.top + plotHeight}
                stroke="#64748b"
                strokeWidth="1.5"
                opacity={visibleCount < 100 ? 1 : 0.2}
              />
              <line
                x1={padding.left}
                y1={padding.top + plotHeight}
                x2={padding.left + plotWidth}
                y2={padding.top + plotHeight}
                stroke="#64748b"
                strokeWidth="1.5"
                opacity={visibleCount < 100 ? 1 : 0.2}
              />

              {/* 10. X EKSENİ ZAMAN TIKLARI VE ETİKETLERİ */}
              {xTicks.map((tick, tIdx) => (
                <g key={`xtick-${tIdx}`}>
                  <line
                    x1={tick.xPos}
                    y1={padding.top + plotHeight}
                    x2={tick.xPos}
                    y2={padding.top + plotHeight + 6}
                    stroke="#64748b"
                    strokeWidth="1.5"
                  />
                  <text
                    x={tick.xPos}
                    y={padding.top + plotHeight + 20}
                    fill="#334155"
                    fontSize="10"
                    fontWeight="bold"
                    textAnchor={tIdx === 0 ? 'start' : tIdx === xTicks.length - 1 ? 'end' : 'middle'}
                    fontFamily="monospace"
                  >
                    {tick.timeStr}
                  </text>
                </g>
              ))}


            </svg>
          </div>

          {/* MİNİ GENEL BAKIŞ HARİTASI (TIMELINE SCRUBBER & MINI-MAP) */}
          <div className="w-full max-w-5xl mt-2 flex items-center gap-2">
            <span className="text-[10px] font-mono font-bold text-slate-500 shrink-0">
              Uçuş Zaman Tüneli:
            </span>
            <div className="flex-1 relative h-4 bg-slate-200 rounded-md overflow-hidden border border-slate-300">
              {/* Olayların mini haritadaki yerleri */}
              {activeEpisodes.map((ep, idx) => {
                const sRatio = Math.max(0, ep.startRow / (chartData.length || 1));
                const eRatio = Math.min(1, ep.endRow / (chartData.length || 1));
                return (
                  <div
                    key={`mini-ep-${idx}`}
                    style={{ left: `${sRatio * 100}%`, width: `${Math.max(1, (eRatio - sRatio) * 100)}%` }}
                    className={`absolute top-0 bottom-0 ${ep.level === 'RED' ? 'bg-red-500' : 'bg-amber-400'} opacity-70`}
                  />
                );
              })}

              {/* Görünür Yakınlaştırılmış Pencere Kutusu */}
              <div
                style={{
                  left: `${(viewStart / (chartData.length - 1 || 1)) * 100}%`,
                  width: `${(visibleCount / (chartData.length || 1)) * 100}%`
                }}
                className="absolute top-0 bottom-0 bg-sky-500/30 border-2 border-sky-600 rounded-xs"
                title="Şu anda ekranda görüntülenen zaman aralığı"
              />

              {/* Anlık Uçak Konumu İşareti */}
              <div
                style={{ left: `${(currentIndex / (chartData.length - 1 || 1)) * 100}%` }}
                className="absolute top-0 bottom-0 w-1 bg-amber-600 shadow-xs"
              />
            </div>
            <span className="text-[10px] font-mono text-slate-500 font-bold shrink-0">
              {chartData[chartData.length - 1]?.timeStr || ''}
            </span>
          </div>
        </div>

        {/* ========================================================================= */}
        {/* ALT KONTROL VE OYNATMA PANELİ (PLAYBACK CONTROLLER - BEYAZ/AÇIK TEMA) */}
        {/* ========================================================================= */}
        <div className="p-3.5 bg-slate-50 border-t border-slate-200 shrink-0">
          <div className="max-w-5xl mx-auto flex flex-col gap-2.5">
            
            {/* Süre Çizelgesi (Scrubber Slider) */}
            <div className="flex items-center gap-3">
              <span className="text-xs font-mono font-bold text-slate-600 w-24 text-right">
                {chartData[0]?.timeStr || '00:00:00'}
              </span>

              <div className="flex-1 relative flex items-center">
                <input
                  type="range"
                  min="0"
                  max={Math.max(0, chartData.length - 1)}
                  value={currentIndex}
                  onChange={e => {
                    setCurrentIndex(parseInt(e.target.value, 10));
                  }}
                  className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-amber-500 focus:outline-none"
                />
              </div>

              <span className="text-xs font-mono font-bold text-slate-600 w-24">
                {chartData[chartData.length - 1]?.timeStr || '00:00:00'}
              </span>
            </div>

            {/* Oynatma Butonları, İkaza Atlama ve Hız Seçimi */}
            <div className="flex flex-wrap items-center justify-between gap-3 pt-0.5">
              
              {/* Sol: Olaylara Atlama Butonları */}
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={jumpToPrevEpisode}
                  disabled={activeEpisodes.length === 0}
                  className="px-3 py-1.5 rounded-xl bg-white hover:bg-slate-100 disabled:opacity-40 text-slate-800 text-xs font-bold transition-all flex items-center gap-1 cursor-pointer border border-slate-300 shadow-2xs"
                  title="Önceki sarı/kırmızı ikaz aralığına atla"
                >
                  <SkipBack className="w-3.5 h-3.5 text-amber-600" />
                  <span>Önceki İkaz</span>
                </button>

                <button
                  type="button"
                  onClick={jumpToNextEpisode}
                  disabled={activeEpisodes.length === 0}
                  className="px-3 py-1.5 rounded-xl bg-white hover:bg-slate-100 disabled:opacity-40 text-slate-800 text-xs font-bold transition-all flex items-center gap-1 cursor-pointer border border-slate-300 shadow-2xs"
                  title="Sonraki sarı/kırmızı ikaz aralığına atla"
                >
                  <span>Sonraki İkaz</span>
                  <SkipForward className="w-3.5 h-3.5 text-amber-600" />
                </button>

                <span className="text-[11px] font-mono text-slate-600 ml-2">
                  Toplam İkaz: <strong className="text-amber-700">{activeEpisodes.length} Adet</strong>
                </span>
              </div>

              {/* Orta: Ana Oynat / Durdur / Başa Sar Butonları */}
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setCurrentIndex(0)}
                  className="p-2 rounded-xl bg-white hover:bg-slate-100 text-slate-700 transition-all cursor-pointer border border-slate-300 shadow-2xs"
                  title="Başa Sar"
                >
                  <RotateCcw className="w-4 h-4" />
                </button>

                <button
                  type="button"
                  onClick={() => setIsPlaying(!isPlaying)}
                  className="px-6 py-2 rounded-2xl bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 text-slate-950 font-black text-sm uppercase tracking-wider shadow-md hover:scale-102 active:scale-95 transition-all flex items-center gap-2 cursor-pointer border border-amber-600"
                >
                  {isPlaying ? (
                    <>
                      <Pause className="w-4 h-4 fill-slate-950" />
                      <span>Durdur</span>
                    </>
                  ) : (
                    <>
                      <Play className="w-4 h-4 fill-slate-950" />
                      <span>Uçuşu Oynat</span>
                    </>
                  )}
                </button>
              </div>

              {/* Sağ: Oynatma Hızı ve Süre Sayacı */}
              <div className="flex items-center gap-3 font-mono text-xs">
                {/* Hız Butonları */}
                <div className="flex items-center bg-white border border-slate-300 rounded-xl p-0.5 shadow-2xs">
                  {[1, 2, 5, 10, 25].map(spd => (
                    <button
                      key={spd}
                      type="button"
                      onClick={() => setPlaySpeed(spd)}
                      className={`px-2 py-1 rounded-lg font-bold text-[11px] transition-all cursor-pointer ${
                        playSpeed === spd
                          ? 'bg-amber-500 text-slate-950 font-black shadow-2xs'
                          : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                      }`}
                    >
                      {spd}x
                    </button>
                  ))}
                </div>

                {/* Zaman ve Satır Sayacı */}
                <div className="flex items-center gap-2">
                  <div className="bg-white border border-slate-300 px-3 py-1.5 rounded-xl text-slate-800 flex items-center gap-2 shadow-2xs">
                    <Clock className="w-3.5 h-3.5 text-amber-600" />
                    <span className="font-bold text-slate-900">
                      {(hoverIndex !== null ? chartData[hoverIndex]?.timeStr : currentItem?.timeStr) || '-'}
                    </span>
                    <span className="text-slate-400">|</span>
                    <span className="text-[11px] text-slate-600">
                      Satır #{(hoverIndex !== null ? chartData[hoverIndex]?.originalIndex : currentItem?.originalIndex) || 0} / #{chartData.length}
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={handleScreenshot}
                    className="p-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-600 transition-colors border border-slate-300"
                    title="Ekran Görüntüsü Al"
                  >
                    <Camera className="w-5 h-5" />
                  </button>
                </div>
              </div>

            </div>
          </div>
        </div>

      </div>
    </div>
  );
};
