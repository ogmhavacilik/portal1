import React, { useState, useEffect, useRef } from 'react';
import { removeBackground } from '@imgly/background-removal';
import { 
  Camera, 
  Upload, 
  Check, 
  X,
  Maximize2,
  Trash2,
  Lock,
  Unlock
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { CachedDriveImage } from './CachedDriveImage';

interface ImageEditorAndRetoucherProps {
  imageKey: string;
  currentImageUrl: string | null;
  hasImage: boolean;
  isImageUpdateUnlocked: boolean;
  isUploadingToDrive: boolean;
  onSaveImage: (base64Data: string, mimeType: string) => Promise<void>;
  onRemoveImage: () => void;
  onUnlockImageUpdate: () => void;
  onLockImageUpdate: () => void;
  onSaveImageUrl?: (url: string) => void;
  partName: string;
  manufacturer: string;
}

export const ImageEditorAndRetoucher: React.FC<ImageEditorAndRetoucherProps> = ({
  imageKey,
  currentImageUrl,
  hasImage,
  isImageUpdateUnlocked,
  isUploadingToDrive,
  onSaveImage,
  onRemoveImage,
  onUnlockImageUpdate,
  onLockImageUpdate,
  partName,
  manufacturer
}) => {
  const [statusText, setStatusText] = useState<string>('');
  const [resultImageSrc, setResultImageSrc] = useState<string | null>(null);
  const [processedBlob, setProcessedBlob] = useState<Blob | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [imageScale, setImageScale] = useState<number>(1);
  const [isFullScreen, setIsFullScreen] = useState<boolean>(false);

  // Fullscreen Zoom & Pan States
  const [fsScale, setFsScale] = useState<number>(1);
  const [fsOffset, setFsOffset] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [isDraggingFs, setIsDraggingFs] = useState<boolean>(false);
  const [dragStart, setDragStart] = useState<{ x: number; y: number }>({ x: 0, y: 0 });

  const [selectedBgColor, setSelectedBgColor] = useState<string>('transparent');

  const BG_COLOR_OPTIONS = [
    { id: 'transparent', label: 'Şeffaf (PNG)', color: 'transparent', border: 'border-slate-300' },
    { id: '#ffffff', label: 'Saf Beyaz', color: '#ffffff', border: 'border-slate-300' },
    { id: '#f8fafc', label: 'Stüdyo Beyazı', color: '#f8fafc', border: 'border-slate-300' },
    { id: '#f1f5f9', label: 'Açık Gri', color: '#f1f5f9', border: 'border-slate-300' },
    { id: '#e2e8f0', label: 'Gümüş Gri', color: '#e2e8f0', border: 'border-slate-300' },
    { id: '#64748b', label: 'Füme', color: '#64748b', border: 'border-slate-400' },
    { id: '#1e293b', label: 'Koyu Antrasit', color: '#1e293b', border: 'border-slate-600' },
    { id: '#000000', label: 'Saf Siyah', color: '#000000', border: 'border-black' },
  ];

  // Handle photo capture or file selection
  const handleFileProcess = async (file: File) => {
    if (!file) return;

    setStatusText('Arka plan temizleniyor, lütfen bekleyin...');
    setResultImageSrc(null);
    setProcessedBlob(null);
    setIsProcessing(true);
    setSelectedBgColor('transparent');

    try {
      // 1. İstemci tarafında AI modeliyle arka planı sil (Ön bellekte işlenir)
      const removeFn = (window as any).imglyRemoveBackground || removeBackground;
      const blob = await removeFn(file, {
        model: 'isnet_fp16', // Hızlı ve kaliteli model seçeneği
        progress: (key: string, current: number, total: number) => {
          if (total > 0) {
            const percent = Math.round((current / total) * 100);
            setStatusText(`Model yükleniyor: %${percent}`);
          }
        }
      });

      // 2. İşlenen PNG Blob verisini URL'ye dönüştür
      const imageObjectURL = URL.createObjectURL(blob);

      // 3. Ekranda göster
      setResultImageSrc(imageObjectURL);
      setProcessedBlob(blob);
      setStatusText('İşlem tamamlandı!');
    } catch (error: any) {
      console.error('Arka plan temizleme hatası:', error);
      setStatusText('Bir hata oluştu, lütfen tekrar deneyin.');
    } finally {
      setIsProcessing(false);
    }
  };

  // Convert Blob to base64 and save to cloud (with selected background color if chosen)
  const handleSaveToCloud = async () => {
    if (!processedBlob) return;

    let finalBlob = processedBlob;

    // Apply background color to image canvas if a color is selected
    if (selectedBgColor !== 'transparent') {
      try {
        const img = new Image();
        const objectUrl = URL.createObjectURL(processedBlob);
        await new Promise((resolve, reject) => {
          img.onload = resolve;
          img.onerror = reject;
          img.src = objectUrl;
        });

        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.fillStyle = selectedBgColor;
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(img, 0, 0);

          const tempBlob = await new Promise<Blob | null>((resolve) => {
            canvas.toBlob((b) => resolve(b), 'image/png');
          });
          if (tempBlob) {
            finalBlob = tempBlob;
          }
        }
        URL.revokeObjectURL(objectUrl);
      } catch (err) {
        console.error("Canvas background fill error:", err);
      }
    }

    const reader = new FileReader();
    reader.onloadend = async () => {
      const base64Data = (reader.result as string).split(',')[1];
      if (base64Data) {
        await onSaveImage(base64Data, 'image/png');
        setResultImageSrc(null);
        setProcessedBlob(null);
        setStatusText('');
        setSelectedBgColor('transparent');
      }
    };
    reader.readAsDataURL(finalBlob);
  };

  return (
    <div className="flex flex-col gap-5 w-full select-text text-left">
      {/* Header Title */}
      <div className="flex items-center justify-between border-b border-slate-200/80 pb-3">
        <div>
          <h3 className="text-sm font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
            📷 Mobil Otomatik Arka Plan Temizleme
          </h3>
          <p className="text-[10px] font-bold text-slate-500">
            Kamerayla çekilen nesne arka planı yapay zeka ile şeffaf PNG yapılır.
          </p>
        </div>
      </div>

      {/* Main Container / Result Viewport */}
      <div className="w-full aspect-square bg-slate-50 border-2 border-slate-200 rounded-[2rem] overflow-hidden relative flex flex-col items-center justify-center p-4 shadow-inner group">
        {isProcessing ? (
          <div className="flex flex-col items-center justify-center gap-3 p-6 text-center animate-pulse">
            <div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin" />
            <p className="text-xs font-black text-blue-800 uppercase tracking-widest">ARKA PLAN TEMİZLENİYOR...</p>
            <p className="text-[11px] font-extrabold text-blue-600">{statusText}</p>
          </div>
        ) : resultImageSrc ? (
          <div className="w-full h-full flex flex-col items-center justify-center gap-2 relative">
            {/* İşlenmiş Şeffaf (PNG) Görsel */}
            <img
              id="resultImage"
              src={resultImageSrc}
              alt="Temizlenmiş Görsel"
              style={{
                maxWidth: '100%',
                maxHeight: '100%',
                borderRadius: '8px',
                border: '2px dashed #ccc',
                display: 'block',
                backgroundColor: selectedBgColor !== 'transparent' ? selectedBgColor : 'transparent',
                backgroundImage:
                  selectedBgColor === 'transparent'
                    ? 'repeating-linear-gradient(45deg, #ccc 25%, transparent 25%, transparent 75%, #ccc 75%, #ccc), repeating-linear-gradient(45deg, #ccc 25%, #fff 25%, #fff 75%, #ccc 75%, #ccc)'
                    : 'none',
                backgroundPosition: '0 0, 9px 9px',
                backgroundSize: '18px 18px',
                objectFit: 'contain',
                transform: `scale(${imageScale})`
              }}
            />
            <button
              type="button"
              onClick={() => {
                setFsScale(1);
                setFsOffset({ x: 0, y: 0 });
                setIsFullScreen(true);
              }}
              className="absolute top-2 right-2 w-9 h-9 bg-black/60 hover:bg-black/80 backdrop-blur-sm text-white rounded-xl flex items-center justify-center shadow-lg cursor-pointer transition-all active:scale-95 z-10 border border-white/10"
              title="Tam Ekran Görüntüle"
            >
              <Maximize2 className="w-4 h-4 text-emerald-400" />
            </button>
          </div>
        ) : currentImageUrl ? (
          <div className="w-full h-full relative flex items-center justify-center overflow-hidden">
            <CachedDriveImage
              src={currentImageUrl}
              alt="Teçhizat Görseli"
              className="w-full h-full object-contain rounded-2xl"
              style={{ transform: `scale(${imageScale})` }}
              referrerPolicy="no-referrer"
            />
            <button
              type="button"
              onClick={() => {
                setFsScale(1);
                setFsOffset({ x: 0, y: 0 });
                setIsFullScreen(true);
              }}
              className="absolute top-2 right-2 w-9 h-9 bg-black/60 hover:bg-black/80 backdrop-blur-sm text-white rounded-xl flex items-center justify-center shadow-lg cursor-pointer transition-all active:scale-95 z-10 border border-white/10"
              title="Tam Ekran Görüntüle"
            >
              <Maximize2 className="w-4 h-4 text-emerald-400" />
            </button>
          </div>
        ) : (
          <div className="text-center p-6 text-slate-400 flex flex-col items-center">
            <span className="text-5xl mb-3 block select-none">📸</span>
            <p className="text-xs font-black uppercase text-slate-500 mb-1">Görsel Bulunmamaktadır</p>
            <p className="text-[10px] text-slate-400 leading-relaxed max-w-xs text-center">
              Aşağıdaki butonla fotoğraınızı çekip arka planını otomatik temizleyebilirsiniz.
            </p>
          </div>
        )}

        {/* Zoom Controls */}
        {(resultImageSrc || currentImageUrl) && !isProcessing && (
          <div className="absolute bottom-3 left-1/2 -translate-x-1/2 bg-black/60 backdrop-blur-md rounded-full px-3 py-1 flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity z-10">
            <button
              type="button"
              onClick={() => setImageScale((prev) => Math.min(prev + 0.25, 3))}
              className="w-6 h-6 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
            >
              ＋
            </button>
            <button
              type="button"
              onClick={() => setImageScale((prev) => Math.max(prev - 0.25, 0.5))}
              className="w-6 h-6 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
            >
              －
            </button>
            <button
              type="button"
              onClick={() => setImageScale(1)}
              className="px-2 py-0.5 bg-white/20 hover:bg-white/40 text-white rounded-full text-[9px] font-bold cursor-pointer transition-colors"
            >
              SIFIRLA
            </button>
          </div>
        )}
      </div>

      {/* Status Div */}
      <div id="status" className="my-1 font-bold text-[#555] text-xs text-center min-h-[20px]">
        {statusText}
      </div>

      {/* Arka Plan Renk Seçenekleri (Arka plan temizlendikten sonra açılır) */}
      {resultImageSrc && (
        <div className="flex flex-col gap-2.5 bg-white border border-slate-200 rounded-2xl p-4 shadow-sm animate-fade-in">
          <div className="flex items-center justify-between border-b border-slate-100 pb-2">
            <span className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
              🎨 Arka Plan Renk Seçenekleri
            </span>
            <span className="text-[10px] font-bold text-blue-600 bg-blue-50 px-2.5 py-0.5 rounded-full">
              {BG_COLOR_OPTIONS.find((o) => o.color === selectedBgColor)?.label}
            </span>
          </div>
          <div className="grid grid-cols-4 sm:grid-cols-8 gap-2">
            {BG_COLOR_OPTIONS.map((opt) => {
              const isSelected = selectedBgColor === opt.color;
              return (
                <button
                  key={opt.id}
                  type="button"
                  onClick={() => setSelectedBgColor(opt.color)}
                  className={`flex flex-col items-center justify-center p-2 rounded-xl border-2 transition-all cursor-pointer relative ${
                    isSelected
                      ? 'border-blue-600 ring-2 ring-blue-500/20 bg-blue-50/50 scale-105 shadow-sm'
                      : 'border-slate-100 hover:border-slate-300 bg-slate-50/50'
                  }`}
                  title={opt.label}
                >
                  <div
                    className={`w-7 h-7 rounded-lg border shadow-inner ${opt.border} flex items-center justify-center overflow-hidden`}
                    style={{
                      backgroundColor: opt.color !== 'transparent' ? opt.color : undefined,
                      backgroundImage:
                        opt.color === 'transparent'
                          ? 'repeating-linear-gradient(45deg, #ccc 25%, transparent 25%, transparent 75%, #ccc 75%, #ccc), repeating-linear-gradient(45deg, #ccc 25%, #fff 25%, #fff 75%, #ccc 75%, #ccc)'
                          : 'none',
                      backgroundSize: '8px 8px'
                    }}
                  >
                    {isSelected && (
                      <Check className={`w-4 h-4 ${['#ffffff', '#f8fafc', '#f1f5f9', '#e2e8f0', 'transparent'].includes(opt.color) ? 'text-blue-600' : 'text-white'}`} />
                    )}
                  </div>
                  <span className="text-[9px] font-extrabold text-slate-700 truncate max-w-full mt-1">
                    {opt.label.split(' ')[0]}
                  </span>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* Controls Section */}
      <div className="flex flex-col gap-4 border border-slate-100 bg-slate-50/50 rounded-[2rem] p-5">
        {/* Permission and Lock Indicators */}
        <div className="flex justify-between items-center">
          <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
            {isImageUpdateUnlocked ? (
              <>
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                <span className="text-emerald-800 font-extrabold text-[10px]">GÖRSEL GÜNCELLEME YETKİSİ AKTİF</span>
              </>
            ) : (
              <>
                <span className="w-2 h-2 rounded-full bg-slate-400" />
                <span className="text-slate-500 font-extrabold text-[10px]">GÖRSEL KİLİDİ AKTİF</span>
              </>
            )}
          </span>
          <div className="flex items-center gap-3">
            {hasImage && !resultImageSrc && (
              <button
                type="button"
                onClick={onRemoveImage}
                className="text-[10px] font-extrabold text-red-500 hover:text-red-700 hover:underline cursor-pointer flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> Görseli Kaldır
              </button>
            )}
            {isImageUpdateUnlocked ? (
              <button
                type="button"
                onClick={onLockImageUpdate}
                className="text-[10px] font-bold text-slate-400 hover:text-slate-600 hover:underline cursor-pointer flex items-center gap-1"
              >
                <Lock className="w-3 h-3" /> Kilitle
              </button>
            ) : (
              <button
                type="button"
                onClick={onUnlockImageUpdate}
                className="text-[10px] font-bold text-emerald-600 hover:text-emerald-800 hover:underline cursor-pointer flex items-center gap-1"
              >
                <Unlock className="w-3 h-3" /> Kilidi Aç
              </button>
            )}
          </div>
        </div>

        {/* Mobil Kamerayı Doğrudan Açan Buton */}
        <div className="flex flex-col gap-2.5">
          <label className="btn-camera bg-blue-600 hover:bg-blue-700 text-white font-extrabold text-xs uppercase px-5 py-3.5 rounded-xl cursor-pointer flex items-center justify-center gap-2 transition-all shadow-md active:scale-95 text-center">
            📷 Fotoğraf Çek
            <input
              type="file"
              id="cameraInput"
              accept="image/*"
              capture="environment"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) handleFileProcess(file);
              }}
              className="hidden"
            />
          </label>

          {/* Dosyadan Seç Alternative */}
          <button
            type="button"
            onClick={() => document.getElementById('galleryFileInput')?.click()}
            className="py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-extrabold text-[11px] uppercase rounded-xl transition-all cursor-pointer flex items-center justify-center gap-2"
          >
            <Upload className="w-3.5 h-3.5" />
            📂 Dosyadan Fotoğraf Seç
          </button>
          <input
            id="galleryFileInput"
            type="file"
            accept="image/*"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) handleFileProcess(file);
            }}
            className="hidden"
          />
        </div>

        {/* Action Buttons for Processed Result */}
        {resultImageSrc && processedBlob && (
          <div className="flex gap-2.5 mt-2">
            <button
              type="button"
              disabled={isUploadingToDrive || isProcessing}
              onClick={handleSaveToCloud}
              className={`flex-1 py-3 bg-emerald-700 hover:bg-emerald-800 text-white font-extrabold text-[11px] uppercase tracking-wider rounded-xl transition-all cursor-pointer shadow-md text-center active:scale-95 flex items-center justify-center gap-1.5 ${
                isUploadingToDrive || isProcessing ? 'opacity-50 cursor-not-allowed' : ''
              }`}
            >
              <Check className="w-4 h-4" />
              BULUTA GÖNDER VE GÜNCELLE
            </button>
            <button
              type="button"
              onClick={() => {
                setResultImageSrc(null);
                setProcessedBlob(null);
                setStatusText('');
              }}
              className="px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 font-extrabold text-[11px] uppercase tracking-wider rounded-xl transition-all cursor-pointer"
            >
              İPTAL
            </button>
          </div>
        )}
      </div>

      {/* Fullscreen Modal */}
      <AnimatePresence>
        {isFullScreen && (resultImageSrc || currentImageUrl) && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/95 backdrop-blur-md z-[9999] flex items-center justify-center p-4 overflow-hidden"
            onWheel={(e) => {
              e.preventDefault();
              const zoomIntensity = 0.15;
              setFsScale((prev) => {
                const next = prev + (e.deltaY < 0 ? zoomIntensity : -zoomIntensity);
                const bounded = Math.max(0.8, Math.min(5, next));
                if (bounded <= 1) {
                  setFsOffset({ x: 0, y: 0 });
                }
                return bounded;
              });
            }}
          >
            <button
              onClick={() => {
                setIsFullScreen(false);
                setFsScale(1);
                setFsOffset({ x: 0, y: 0 });
              }}
              className="absolute top-6 right-6 text-white hover:text-red-400 bg-white/10 hover:bg-white/20 p-2.5 rounded-full transition-all z-20 cursor-pointer"
              title="Kapat"
            >
              <X className="w-6 h-6" />
            </button>

            <div
              className="relative max-w-5xl max-h-[90vh] w-full h-full flex items-center justify-center overflow-hidden"
              onMouseDown={(e) => {
                if (fsScale > 1) {
                  setIsDraggingFs(true);
                  setDragStart({ x: e.clientX - fsOffset.x, y: e.clientY - fsOffset.y });
                }
              }}
              onMouseMove={(e) => {
                if (isDraggingFs && fsScale > 1) {
                  setFsOffset({
                    x: e.clientX - dragStart.x,
                    y: e.clientY - dragStart.y
                  });
                }
              }}
              onMouseUp={() => setIsDraggingFs(false)}
              onMouseLeave={() => setIsDraggingFs(false)}
            >
              <CachedDriveImage
                src={resultImageSrc || currentImageUrl}
                alt="Teçhizat Görseli Tam Ekran"
                className="max-w-full max-h-full object-contain rounded-2xl select-none transition-transform duration-75 ease-out"
                style={{
                  transform: `translate(${fsOffset.x}px, ${fsOffset.y}px) scale(${fsScale})`,
                  cursor: fsScale > 1 ? (isDraggingFs ? 'grabbing' : 'grab') : 'default'
                }}
                referrerPolicy="no-referrer"
              />
            </div>

            {/* Floating Zoom Controls */}
            <div className="absolute bottom-6 left-1/2 -translate-x-1/2 bg-white/10 backdrop-blur-md rounded-full px-4 py-2 flex items-center gap-3 border border-white/20 text-white select-none z-30">
              <button
                type="button"
                onClick={() => {
                  setFsScale((prev) => {
                    const next = Math.max(0.8, prev - 0.2);
                    if (next <= 1) setFsOffset({ x: 0, y: 0 });
                    return next;
                  });
                }}
                className="w-8 h-8 bg-white/10 hover:bg-white/20 active:scale-95 text-white rounded-full flex items-center justify-center font-bold cursor-pointer transition-all"
                title="Uzaklaştır"
              >
                －
              </button>
              <span className="text-xs font-mono font-bold min-w-14 text-center">
                % {Math.round(fsScale * 100)}
              </span>
              <button
                type="button"
                onClick={() => setFsScale((prev) => Math.min(5, prev + 0.2))}
                className="w-8 h-8 bg-white/10 hover:bg-white/20 active:scale-95 text-white rounded-full flex items-center justify-center font-bold cursor-pointer transition-all"
                title="Yakınlaştır"
              >
                ＋
              </button>
              <button
                type="button"
                onClick={() => {
                  setFsScale(1);
                  setFsOffset({ x: 0, y: 0 });
                }}
                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white rounded-xl text-[10px] font-black cursor-pointer transition-all"
              >
                SIFIRLA
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
};
