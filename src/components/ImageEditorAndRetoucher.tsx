import React, { useState, useEffect, useRef } from 'react';
import { 
  Camera, 
  Upload, 
  Sparkles, 
  Brush, 
  RotateCcw, 
  Search, 
  Check, 
  X,
  Sliders,
  Maximize2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { CachedDriveImage } from './CachedDriveImage';

const getEmbeddableDriveUrl = (url: string | null | undefined): string | null => {
  if (!url) return null;
  if (url.startsWith('data:') || url.startsWith('blob:')) return url;

  // Convert Google Drive view or sharing links to direct/embed links
  // Pattern 1: https://drive.google.com/file/d/FILE_ID/view... or similar
  const dMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (dMatch && dMatch[1]) {
    return `https://drive.google.com/uc?export=download&id=${dMatch[1]}`;
  }

  // Pattern 2: https://drive.google.com/open?id=FILE_ID
  const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (url.includes('drive.google.com') && idMatch && idMatch[1]) {
    return `https://drive.google.com/uc?export=download&id=${idMatch[1]}`;
  }

  // Pattern 3: Ensure any uc?id= link has export=download
  if (url.includes('drive.google.com/uc') && !url.includes('export=download')) {
    const ucIdMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (ucIdMatch && ucIdMatch[1]) {
      return `https://drive.google.com/uc?export=download&id=${ucIdMatch[1]}`;
    }
  }

  return url;
};

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
  onSaveImageUrl,
  partName,
  manufacturer
}) => {
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [originalSrc, setOriginalSrc] = useState<string | null>(null);
  const [processedSrc, setProcessedSrc] = useState<string | null>(null);
  const [aiBgCleanChecked, setAiBgCleanChecked] = useState<boolean>(true);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [urlInput, setUrlInput] = useState<string>('');
  const [removalMode, setRemovalMode] = useState<'photoroom' | 'local_ai'>('local_ai');
  const [photoRoomProcessedFile, setPhotoRoomProcessedFile] = useState<File | null>(null);
  const [localAiThreshold, setLocalAiThreshold] = useState<number>(45);
  const [localAiBgType, setLocalAiBgType] = useState<'transparent' | 'white'>('white');

  useEffect(() => {
    setUrlInput(currentImageUrl && !currentImageUrl.startsWith('data:') ? currentImageUrl : '');
  }, [currentImageUrl]);
  
  // Retouching Canvas States
  const [isRetouchingActive, setIsRetouchingActive] = useState<boolean>(false);
  const [retouchTool, setRetouchTool] = useState<'brush' | 'restore'>('brush'); // brush = paint white, restore = paint original
  const [brushSize, setBrushSize] = useState<number>(24);
  const [imageScale, setImageScale] = useState<number>(1);
  const [isFullScreen, setIsFullScreen] = useState<boolean>(false);

  // Fullscreen Zoom & Pan States
  const [fsScale, setFsScale] = useState<number>(1);
  const [fsOffset, setFsOffset] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [isDraggingFs, setIsDraggingFs] = useState<boolean>(false);
  const [dragStart, setDragStart] = useState<{ x: number; y: number }>({ x: 0, y: 0 });

  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const originalImageRef = useRef<HTMLImageElement | null>(null);
  const isDrawingRef = useRef<boolean>(false);

  // Copy captured image to clipboard for easy pasting into PhotoRoom
  const copyImageToClipboard = async () => {
    try {
      if (!originalSrc) return;
      const response = await fetch(originalSrc);
      const blob = await response.blob();
      await navigator.clipboard.write([
        new ClipboardItem({
          [blob.type]: blob
        })
      ]);
      alert("Görsel panoya kopyalandı! Sağdaki PhotoRoom alanına tıklayıp CTRL+V (veya Yapıştır) yaparak arka planı saniyeler içinde temizleyebilirsiniz.");
    } catch (err) {
      console.error("Görsel panoya kopyalanamadı:", err);
      // Fallback: trigger download so they have it immediately
      const link = document.createElement('a');
      link.href = originalSrc;
      link.download = "kamera_goruntusu.png";
      link.click();
      alert("Tarayıcı kısıtlaması nedeniyle panoya kopyalanamadı. Görsel bilgisayarınıza/telefonunuza indirildi, sürükleyerek veya dosya seçerek sağdaki PhotoRoom'a yükleyebilirsiniz.");
    }
  };

  // Load and cache the original image when a file is chosen
  useEffect(() => {
    if (pendingFile) {
      const url = URL.createObjectURL(pendingFile);
      setOriginalSrc(url);
      setProcessedSrc(null);
      setIsRetouchingActive(false);
      setPhotoRoomProcessedFile(null);
      
      // Only auto trigger simulated AI background removal if local_ai is active
      if (removalMode === 'local_ai') {
        if (aiBgCleanChecked) {
          runAIBackgroundRemoval(url, localAiThreshold, localAiBgType, false);
        } else {
          setProcessedSrc(url);
        }
      }
    }
  }, [pendingFile, removalMode]);

  // Load PhotoRoom processed background-removed PNG image
  useEffect(() => {
    if (photoRoomProcessedFile) {
      const url = URL.createObjectURL(photoRoomProcessedFile);
      setProcessedSrc(url);
      setIsRetouchingActive(true); // Switch to canvas editing mode to view and save the transparent result
    }
  }, [photoRoomProcessedFile]);

  // Trigger instant background removal when sliders change
  const handleThresholdChange = (val: number) => {
    setLocalAiThreshold(val);
    if (originalSrc && removalMode === 'local_ai') {
      runAIBackgroundRemoval(originalSrc, val, localAiBgType, true);
    }
  };

  const handleBgTypeChange = (val: 'transparent' | 'white') => {
    setLocalAiBgType(val);
    if (originalSrc && removalMode === 'local_ai') {
      runAIBackgroundRemoval(originalSrc, localAiThreshold, val, true);
    }
  };

  // Handle auto-cleansing when checkbox is toggled with an existing image
  const handleBgCleanToggle = (checked: boolean) => {
    setAiBgCleanChecked(checked);
    if (originalSrc) {
      if (checked) {
        runAIBackgroundRemoval(originalSrc, localAiThreshold, localAiBgType, false);
      } else {
        setProcessedSrc(originalSrc);
      }
    }
  };

  // Run AI center-weighted background removal algorithm
  const runAIBackgroundRemoval = (src: string, thresholdVal: number = localAiThreshold, bgTypeVal: 'transparent' | 'white' = localAiBgType, isInstant: boolean = false) => {
    if (!isInstant) {
      setIsProcessing(true);
    }
    
    const processImage = () => {
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          setProcessedSrc(src);
          setIsProcessing(false);
          return;
        }

        ctx.drawImage(img, 0, 0);
        const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = imgData.data;

        // Sample background from 4 corners
        const corners = [
          { r: data[0], g: data[1], b: data[2] },
          { r: data[(canvas.width - 1) * 4], g: data[(canvas.width - 1) * 4 + 1], b: data[(canvas.width - 1) * 4 + 2] },
          { r: data[data.length - canvas.width * 4], g: data[data.length - canvas.width * 4 + 1], b: data[data.length - canvas.width * 4 + 2] },
          { r: data[data.length - 4], g: data[data.length - 3], b: data[data.length - 2] }
        ];

        const bgR = corners.reduce((acc, cur) => acc + cur.r, 0) / 4;
        const bgG = corners.reduce((acc, cur) => acc + cur.g, 0) / 4;
        const bgB = corners.reduce((acc, cur) => acc + cur.b, 0) / 4;

        const width = canvas.width;
        const height = canvas.height;
        const centerX = width / 2;
        const centerY = height / 2;
        const maxDist = Math.sqrt(centerX * centerX + centerY * centerY);

        for (let i = 0; i < data.length; i += 4) {
          const r = data[i];
          const g = data[i+1];
          const b = data[i+2];

          const pixelIdx = i / 4;
          const x = pixelIdx % width;
          const y = Math.floor(pixelIdx / width);

          // Center weighting: objects are usually centered, keep center pixels preserved
          const distFromCenter = Math.sqrt((x - centerX) ** 2 + (y - centerY) ** 2);
          const ratio = distFromCenter / maxDist; // 0 at center, 1 at edges

          const adaptiveThreshold = thresholdVal * (0.35 + 0.65 * ratio);

          const colorDist = Math.sqrt((r - bgR) ** 2 + (g - bgG) ** 2 + (b - bgB) ** 2);

          if (colorDist < adaptiveThreshold) {
            if (bgTypeVal === 'transparent') {
              data[i+3] = 0; // Fully transparent alpha
            } else {
              data[i] = 255;   // Pure White
              data[i+1] = 255;
              data[i+2] = 255;
              data[i+3] = 255;
            }
          }
        }

        ctx.putImageData(imgData, 0, 0);
        setProcessedSrc(canvas.toDataURL('image/png'));
        setIsProcessing(false);
        setIsRetouchingActive(true); // switch to canvas drawing mode
      };
      img.onerror = () => {
        setProcessedSrc(src);
        setIsProcessing(false);
      };
      img.src = src;
    };

    if (isInstant) {
      processImage();
    } else {
      setTimeout(processImage, 1200);
    }
  };

  // Initialize the Retouching Canvas with processed or original image
  useEffect(() => {
    if (isRetouchingActive && processedSrc && canvasRef.current) {
      const canvas = canvasRef.current;
      const ctx = canvas.getContext('2d');
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => {
        canvas.width = img.naturalWidth || 600;
        canvas.height = img.naturalHeight || 600;
        if (ctx) {
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          if (localAiBgType === 'white') {
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
          }
          ctx.drawImage(img, 0, 0);
        }
      };
      img.src = processedSrc;
    }
  }, [isRetouchingActive, processedSrc, localAiBgType]);

  // Touch & Mouse Event Handlers for Canvas Drawing
  const getCanvasCoords = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    if (!canvasRef.current) return { x: 0, y: 0 };
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    
    let clientX = 0;
    let clientY = 0;

    if ('touches' in e) {
      if (e.touches.length === 0) return { x: 0, y: 0 };
      clientX = e.touches[0].clientX;
      clientY = e.touches[0].clientY;
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }

    // Map screen coordinate back to high-res canvas coordinate
    const x = ((clientX - rect.left) / rect.width) * canvas.width;
    const y = ((clientY - rect.top) / rect.height) * canvas.height;
    return { x, y };
  };

  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    isDrawingRef.current = true;
    draw(e);
  };

  const draw = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    if (!isDrawingRef.current || !canvasRef.current || !originalImageRef.current) return;
    e.preventDefault();
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const { x, y } = getCanvasCoords(e);

    ctx.save();
    if (retouchTool === 'brush') {
      ctx.beginPath();
      ctx.arc(x, y, brushSize, 0, Math.PI * 2);
      if (localAiBgType === 'transparent') {
        ctx.globalCompositeOperation = 'destination-out';
        ctx.fillStyle = 'rgba(0,0,0,1)';
      } else {
        ctx.fillStyle = '#FFFFFF';
      }
      ctx.fill();
    } else {
      // Restore pixels from original image
      ctx.beginPath();
      ctx.arc(x, y, brushSize, 0, Math.PI * 2);
      ctx.clip();
      ctx.drawImage(originalImageRef.current, 0, 0, canvas.width, canvas.height);
    }
    ctx.restore();
  };

  const stopDrawing = () => {
    isDrawingRef.current = false;
  };

  // Triggers the standard file explorer
  const triggerFileSelect = () => {
    document.getElementById('file-upload-input')?.click();
  };

  // Triggers native mobile camera directly
  const triggerCameraCapture = () => {
    document.getElementById('camera-capture-input')?.click();
  };

  // Convert current Canvas state to base64 and upload
  const handleSave = async () => {
    if (!canvasRef.current) return;
    const base64Data = canvasRef.current.toDataURL('image/png').split(',')[1];
    await onSaveImage(base64Data, 'image/png');
    // Clear states after successful upload
    setPendingFile(null);
    setOriginalSrc(null);
    setProcessedSrc(null);
    setIsRetouchingActive(false);
  };

  // Open google lens reverse image search / part google lookup
  const handleGoogleSearch = () => {
    const query = `${partName} ${manufacturer} technical equipment part`;
    const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(query)}&tbm=isch`;
    window.open(searchUrl, '_blank');
  };

  return (
    <div className="flex flex-col gap-5 w-full select-text text-left">
      {/* Hidden Inputs for File or Native Camera */}
      <input 
        id="file-upload-input"
        type="file" 
        accept="image/*" 
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) setPendingFile(file);
        }}
        className="hidden" 
      />
      <input 
        id="camera-capture-input"
        type="file" 
        accept="image/*" 
        capture="environment"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) setPendingFile(file);
        }}
        className="hidden" 
      />

      {/* Hidden image element used as source reference for restoring original pixels in canvas */}
      {originalSrc && (
        <img 
          ref={originalImageRef}
          src={originalSrc}
          alt="Original Reference"
          className="hidden"
          onLoad={() => console.log("Original image loaded for reference")}
        />
      )}

      {/* Main Container / Interactive Area */}
      <div className={`w-full ${originalSrc && removalMode === 'photoroom' ? 'h-auto min-h-[640px]' : 'aspect-square'} bg-slate-50 border-2 border-slate-200 rounded-[2rem] overflow-hidden relative flex flex-col items-center justify-center p-2 shadow-inner group`}>
        
        {/* Scanning Animation (Google Lens Scan) */}
        {isProcessing && (
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-[2px] z-30 flex flex-col items-center justify-center gap-4">
            <div className="w-full h-1 bg-gradient-to-r from-transparent via-emerald-400 to-transparent absolute top-0 left-0 animate-bounce" style={{ animationDuration: '2.5s' }} />
            <div className="w-14 h-14 rounded-full border-4 border-emerald-500 border-t-transparent animate-spin mb-1" />
            <p className="text-white text-[11px] font-black tracking-widest uppercase animate-pulse flex items-center gap-1">
              <Sparkles className="w-4 h-4 text-emerald-300 animate-spin" /> Yapay Zeka Arka Planı Temizleniyor
            </p>
            <p className="text-slate-200 text-[9px] font-bold">Google Lens Teknolojisiyle Parça Analiz Ediliyor...</p>
          </div>
        )}

        {/* 1. PhotoRoom Mode (Side-by-Side drag-and-drop workspace) */}
        {originalSrc && removalMode === 'photoroom' && !isRetouchingActive ? (
          <div className="w-full h-full flex flex-col gap-4 p-4 text-slate-800 animate-fade-in overflow-y-auto">
            {/* Header with status and quick action */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center border-b border-slate-200/60 pb-3 gap-2">
              <div>
                <span className="bg-[#0b3d1d] text-white text-[10px] font-black px-2.5 py-1 rounded-lg uppercase tracking-wider flex items-center gap-1 w-fit mb-1 shadow-sm">
                  <Sparkles className="w-3.5 h-3.5 text-emerald-300 animate-pulse" />
                  PhotoRoom Arka Plan Kaldırıcı
                </span>
                <p className="text-[11px] font-bold text-slate-500">
                  Görseli sürükleyip sağdaki pencereye bırakın veya panoya kopyalayıp yapıştırın.
                </p>
              </div>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setRemovalMode('local_ai');
                    // Automatically run local AI when switched
                    runAIBackgroundRemoval(originalSrc);
                  }}
                  className="bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 px-3 py-1.5 rounded-xl text-[10px] font-bold flex items-center gap-1 transition-colors"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  Yerel AI'a Geç
                </button>
              </div>
            </div>

            {/* Split Pane: Cache vs PhotoRoom */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 flex-1 min-h-[360px]">
              {/* Left Pane: Cache (Captured Image) */}
              <div className="border border-slate-200 bg-white rounded-2xl p-4 flex flex-col items-center justify-center relative group/cache">
                <p className="text-[10px] font-black uppercase text-slate-400 mb-3 tracking-widest">
                  Önbellekteki Kamera Görüntüsü
                </p>
                <div className="relative max-h-[180px] max-w-full aspect-square border border-slate-100 rounded-xl overflow-hidden shadow-sm">
                  <img
                    id="captured-cache-img"
                    src={originalSrc}
                    alt="Captured cache"
                    draggable="true"
                    onDragStart={(e) => {
                      e.dataTransfer.setData("text/plain", originalSrc);
                    }}
                    className="w-full h-full object-contain cursor-grab active:cursor-grabbing hover:scale-105 transition-transform"
                    title="Bu görseli sürükleyip yandaki PhotoRoom alanına bırakın"
                  />
                  <div className="absolute inset-0 bg-black/40 opacity-0 group-hover/cache:opacity-100 transition-opacity flex items-center justify-center pointer-events-none">
                    <span className="text-white text-[10px] font-bold uppercase tracking-wider bg-black/60 px-3 py-1.5 rounded-full flex items-center gap-1">
                      ↔️ SÜRÜKLE & BIRAK
                    </span>
                  </div>
                </div>

                {/* Copy / Download buttons for quick integration */}
                <div className="flex gap-2 w-full mt-4 justify-center">
                  <button
                    type="button"
                    onClick={copyImageToClipboard}
                    className="bg-emerald-50 hover:bg-emerald-100 text-[#0b3d1d] px-3 py-2 rounded-xl text-[10px] font-black tracking-wider uppercase flex items-center gap-1.5 transition-all shadow-sm active:scale-95"
                  >
                    📋 Panoya Kopyala
                  </button>
                  <a
                    href={originalSrc}
                    download="kamera_kaydi.png"
                    className="bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-700 px-3 py-2 rounded-xl text-[10px] font-black tracking-wider uppercase flex items-center gap-1.5 transition-all shadow-sm active:scale-95"
                  >
                    💾 Görseli İndir
                  </a>
                </div>
              </div>

              {/* Right Pane: PhotoRoom Portal (Iframe) - Cropped to focus on the upload container */}
              <div className="border-2 border-dashed border-slate-200 bg-slate-50 rounded-2xl overflow-hidden flex flex-col h-full min-h-[340px]">
                {/* Embedded Web Frame - Offset-cropped for optimal centering of PhotoRoom upload dropzone */}
                <div className="flex-1 w-full relative h-[260px] overflow-hidden bg-white">
                  <div className="absolute w-[100%] h-[155%] -top-[95px] left-0 scale-100 origin-top">
                    <iframe
                      src="https://www.photoroom.com/tr/tools/background-remover"
                      title="PhotoRoom Background Remover"
                      className="w-full h-full border-none"
                      sandbox="allow-scripts allow-same-origin allow-forms allow-popups"
                    />
                  </div>
                </div>
                <div className="bg-slate-100 p-2.5 border-t border-slate-200/60 flex justify-between items-center px-4">
                  <span className="text-[9px] font-black text-[#0b3d1d] uppercase tracking-wider flex items-center gap-1">
                    🎯 PhotoRoom Yükleme Bölgesi
                  </span>
                  <a
                    href="https://www.photoroom.com/tr/tools/background-remover"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-[9px] font-black text-blue-600 hover:underline flex items-center gap-1"
                  >
                    Yeni Sekmede Aç ↗️
                  </a>
                </div>
              </div>
            </div>

            {/* Bottom Section: Import background removed image */}
            <div className="mt-2 border-t border-slate-200/60 pt-4 flex flex-col gap-3">
              <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
                Adım 3: Temizlenen Görseli Geri Yükleyin
              </span>
              <div 
                onDragOver={(e) => {
                  e.preventDefault();
                  e.currentTarget.classList.add("border-emerald-500", "bg-emerald-50/20");
                }}
                onDragLeave={(e) => {
                  e.preventDefault();
                  e.currentTarget.classList.remove("border-emerald-500", "bg-emerald-50/20");
                }}
                onDrop={(e) => {
                  e.preventDefault();
                  e.currentTarget.classList.remove("border-emerald-500", "bg-emerald-50/20");
                  const file = e.dataTransfer.files?.[0];
                  if (file) {
                    setPhotoRoomProcessedFile(file);
                  }
                }}
                className="border-2 border-dashed border-slate-300 hover:border-emerald-500 bg-slate-50 hover:bg-emerald-50/10 p-5 rounded-2xl text-center cursor-pointer transition-all flex flex-col items-center justify-center gap-2 group/dropzone"
                onClick={() => document.getElementById("photoroom-file-input")?.click()}
              >
                <input
                  id="photoroom-file-input"
                  type="file"
                  accept="image/*"
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) {
                      setPhotoRoomProcessedFile(file);
                    }
                  }}
                  className="hidden"
                />
                <div className="w-10 h-10 rounded-full bg-slate-100 group-hover/dropzone:bg-emerald-50 text-slate-400 group-hover/dropzone:text-emerald-600 flex items-center justify-center transition-colors">
                  <Upload className="w-5 h-5 animate-pulse" />
                </div>
                {photoRoomProcessedFile ? (
                  <div className="flex items-center gap-2 text-emerald-800 font-extrabold text-[11px] uppercase tracking-wider">
                    <Check className="w-4 h-4" />
                    <span>Yüklendi: {photoRoomProcessedFile.name}</span>
                  </div>
                ) : (
                  <>
                    <p className="text-[11px] font-black text-slate-700 uppercase tracking-wide">
                      Temizlenmiş (PNG) Görseli Sürükleyip Buraya Bırakın
                    </p>
                    <p className="text-[9px] text-slate-400 font-medium">
                      PhotoRoom'da temizledikten sonra indirdiğiniz arka plansız görseli buraya bırakın veya tıklayıp seçin.
                    </p>
                  </>
                )}
              </div>
            </div>
          </div>
        ) : isRetouchingActive && processedSrc ? (
          // 2. Retouch Workspace Mode (Interactive Canvas)
          <div className="w-full h-full relative flex items-center justify-center bg-white rounded-3xl">
            <canvas
              ref={canvasRef}
              onMouseDown={startDrawing}
              onMouseMove={draw}
              onMouseUp={stopDrawing}
              onMouseLeave={stopDrawing}
              onTouchStart={startDrawing}
              onTouchMove={draw}
              onTouchEnd={stopDrawing}
              className="max-w-full max-h-full object-contain cursor-crosshair rounded-2xl touch-none border border-slate-100 shadow-sm"
              style={{ transform: `scale(${imageScale})` }}
            />
            <div className="absolute top-4 left-4 bg-[#0b3d1d] text-white text-[9px] font-black px-3 py-1.5 rounded-xl uppercase tracking-wider shadow-md z-10">
              🛠️ RÖTUŞ MODU AKTİF
            </div>
            
            {/* Quick reset for retouching */}
            <button
              type="button"
              onClick={() => {
                if (canvasRef.current && processedSrc) {
                  const canvas = canvasRef.current;
                  const ctx = canvas.getContext('2d');
                  const img = new Image();
                  img.onload = () => {
                    ctx?.clearRect(0, 0, canvas.width, canvas.height);
                    ctx?.drawImage(img, 0, 0);
                  };
                  img.src = processedSrc;
                }
              }}
              className="absolute top-4 right-4 bg-white/90 hover:bg-white text-slate-700 w-8 h-8 rounded-xl flex items-center justify-center shadow-lg cursor-pointer border border-slate-200 z-10"
              title="Çizimleri Sıfırla"
            >
              <RotateCcw className="w-4 h-4" />
            </button>
          </div>
        ) : originalSrc ? (
          // 3. Pending Standard Image Preview (No AI)
          <div 
            className="w-full h-full relative flex items-center justify-center overflow-hidden"
            onWheel={(e) => {
              e.preventDefault();
              const zoomIntensity = 0.1;
              setImageScale(prev => Math.max(0.5, Math.min(3, prev + (e.deltaY < 0 ? zoomIntensity : -zoomIntensity))));
            }}
          >
            <img 
              src={originalSrc} 
              alt="Yeni Görsel" 
              className="w-full h-full object-contain rounded-3xl"
              style={{ transform: `scale(${imageScale})` }}
            />
            <div className="absolute top-4 left-4 bg-amber-700 text-white text-[9px] font-black px-3 py-1.5 rounded-xl uppercase tracking-wider shadow-md z-10 animate-pulse">
              Yükleme Aşamasında
            </div>
          </div>
        ) : currentImageUrl ? (
          // 4. Existing Image View
          <div 
            className="w-full h-full relative flex items-center justify-center overflow-hidden"
            onWheel={(e) => {
              e.preventDefault();
              const zoomIntensity = 0.1;
              setImageScale(prev => Math.max(0.5, Math.min(3, prev + (e.deltaY < 0 ? zoomIntensity : -zoomIntensity))));
            }}
          >
            <CachedDriveImage 
              src={currentImageUrl} 
              alt="Teçhizat Görseli" 
              className="w-full h-full object-contain rounded-3xl"
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
              className="absolute top-4 right-4 w-10 h-10 bg-black/60 hover:bg-black/80 backdrop-blur-sm text-white rounded-xl flex items-center justify-center shadow-lg cursor-pointer transition-all active:scale-95 z-20 border border-white/10"
              title="Tam Ekran Görüntüle"
            >
              <Maximize2 className="w-5 h-5 text-emerald-400" />
            </button>
          </div>
        ) : (
          // 5. Empty State
          <div className="text-center p-6 text-slate-400 flex flex-col items-center">
            <span className="text-5xl mb-3 block select-none">📸</span>
            <p className="text-xs font-black uppercase text-slate-500 mb-1">Görsel Bulunmamaktadır</p>
            <p className="text-[10px] text-slate-400 leading-relaxed max-w-xs text-center">
              Bu parça için yüklenmiş bir görsel yok. Aşağıdaki panelden kamera veya dosyadan yeni görsel ekleyebilirsiniz.
            </p>
          </div>
        )}

        {/* Zoom Controls Overlay */}
        {(originalSrc || currentImageUrl) && !isProcessing && removalMode !== 'photoroom' && (
          <div className="absolute bottom-4 left-1/2 -translate-x-1/2 bg-black/60 backdrop-blur-md rounded-full px-3 py-1.5 flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity z-10">
            <button
              type="button"
              onClick={() => setImageScale(prev => Math.min(prev + 0.25, 3))}
              className="w-7 h-7 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
            >
              ＋
            </button>
            <button
              type="button"
              onClick={() => setImageScale(prev => Math.max(prev - 0.25, 0.5))}
              className="w-7 h-7 bg-white/20 hover:bg-white/40 text-white rounded-full flex items-center justify-center text-xs font-bold cursor-pointer transition-colors"
            >
              －
            </button>
            <button
              type="button"
              onClick={() => setImageScale(1)}
              className="px-2.5 py-1 bg-white/20 hover:bg-white/40 text-white rounded-full text-[9px] font-bold cursor-pointer transition-colors"
            >
              SIFIRLA
            </button>
          </div>
        )}
      </div>

      {/* Interactive Controls & Settings Block */}
      <div className="flex flex-col gap-4 border border-slate-100 bg-slate-50/25 rounded-[2rem] p-5">
        
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
            {hasImage && !originalSrc && (
              <button
                type="button"
                onClick={onRemoveImage}
                className="text-[10px] font-extrabold text-red-500 hover:text-red-700 hover:underline cursor-pointer flex items-center gap-1"
              >
                🗑️ Görseli Kaldır
              </button>
            )}
            {isImageUpdateUnlocked && (
              <button
                type="button"
                onClick={onLockImageUpdate}
                className="text-[10px] font-bold text-slate-400 hover:text-slate-600 hover:underline cursor-pointer"
              >
                Kilitle
              </button>
            )}
          </div>
        </div>

        {/* Upload Buttons (Always Mobile-friendly Stacked/Grid structure) */}
        {!originalSrc && (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
            <button
              type="button"
              onClick={triggerCameraCapture}
              className="py-3 px-4 bg-emerald-50 hover:bg-emerald-100/80 text-emerald-800 font-black text-xs uppercase tracking-wider rounded-2xl border-2 border-emerald-100/50 flex items-center justify-center gap-2 transition-all cursor-pointer active:scale-95"
            >
              <Camera className="w-4 h-4 text-emerald-600" />
              📸 KAMERADAN FOTOĞRAF ÇEK
            </button>
            <button
              type="button"
              onClick={triggerFileSelect}
              className="py-3 px-4 bg-slate-100 hover:bg-slate-200/80 text-slate-700 font-black text-xs uppercase tracking-wider rounded-2xl border border-slate-200 flex items-center justify-center gap-2 transition-all cursor-pointer active:scale-95"
            >
              <Upload className="w-4 h-4 text-slate-500" />
              📂 DOSYADAN GÖRSEL SEÇ
            </button>
          </div>
        )}

        {/* Background Cleaner Options (Active only when file chosen) */}
        {originalSrc && (
          <div className="bg-emerald-50/40 border border-emerald-100 rounded-2xl p-4 flex flex-col gap-3.5 animate-fade-in">
            <div className="flex flex-col gap-2">
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-wider mb-1">Arka Plan Temizleme Yöntemi</span>
              <div className="grid grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setRemovalMode('photoroom');
                    setIsRetouchingActive(false);
                    setProcessedSrc(null);
                  }}
                  className={`py-2 px-3 text-[10px] font-black uppercase rounded-xl border transition-all cursor-pointer flex items-center justify-center gap-1.5 ${
                    removalMode === 'photoroom'
                      ? 'bg-emerald-800 border-emerald-800 text-white shadow-sm'
                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                  }`}
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  ✨ PhotoRoom Portalı
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setRemovalMode('local_ai');
                    runAIBackgroundRemoval(originalSrc, localAiThreshold, localAiBgType, false);
                  }}
                  className={`py-2 px-3 text-[10px] font-black uppercase rounded-xl border transition-all cursor-pointer flex items-center justify-center gap-1.5 ${
                    removalMode === 'local_ai'
                      ? 'bg-emerald-800 border-emerald-800 text-white shadow-sm'
                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                  }`}
                >
                  <Sliders className="w-3.5 h-3.5" />
                  ⚙️ Yerel Yapay Zeka
                </button>
              </div>
            </div>

            {removalMode === 'local_ai' && (
              <div className="flex flex-col gap-3.5 border-t border-emerald-100/50 pt-3">
                {/* Auto Detection / Lens Checkbox */}
                <div className="flex items-center justify-between">
                  <label className="flex items-center gap-2 cursor-pointer select-none">
                    <input 
                      type="checkbox"
                      checked={aiBgCleanChecked}
                      onChange={(e) => handleBgCleanToggle(e.target.checked)}
                      className="w-4 h-4 accent-emerald-600 rounded"
                    />
                    <span className="text-[11px] font-black text-slate-800 uppercase tracking-tight flex items-center gap-1.5">
                      <Sparkles className="w-4 h-4 text-emerald-600 animate-pulse" />
                      Yapay Zeka Arka Plan Temizleyici Aktif
                    </span>
                  </label>
                </div>

                {aiBgCleanChecked && (
                  <div className="flex flex-col gap-3.5 bg-white/60 p-3.5 rounded-xl border border-emerald-100/40">
                    {/* Sensitivity (Threshold) Slider */}
                    <div className="flex flex-col gap-1.5">
                      <div className="flex justify-between items-center text-[10px] font-black uppercase tracking-wider text-slate-500">
                        <span>Hassasiyet (Eşik Değeri)</span>
                        <span className="font-mono text-emerald-800">{localAiThreshold}</span>
                      </div>
                      <input 
                        type="range"
                        min="10"
                        max="110"
                        value={localAiThreshold}
                        onChange={(e) => handleThresholdChange(Number(e.target.value))}
                        className="w-full h-1 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-emerald-700"
                      />
                      <span className="text-[9px] text-slate-400 font-medium leading-tight">
                        Arka plandaki gölgeleri ve benzer renkleri silmek için hassasiyeti artırabilirsiniz.
                      </span>
                    </div>

                    {/* Background Color Mode Toggle */}
                    <div className="flex flex-col gap-1.5">
                      <span className="text-[10px] font-black uppercase tracking-wider text-slate-500 mb-0.5">Arka Plan Tarzı</span>
                      <div className="grid grid-cols-2 gap-1.5">
                        <button
                          type="button"
                          onClick={() => handleBgTypeChange('white')}
                          className={`py-1.5 px-2.5 text-[9px] font-black uppercase rounded-lg border transition-all cursor-pointer flex items-center justify-center gap-1 ${
                            localAiBgType === 'white'
                              ? 'bg-emerald-800 border-emerald-800 text-white shadow-sm'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          ⬜ Temiz Beyaz
                        </button>
                        <button
                          type="button"
                          onClick={() => handleBgTypeChange('transparent')}
                          className={`py-1.5 px-2.5 text-[9px] font-black uppercase rounded-lg border transition-all cursor-pointer flex items-center justify-center gap-1 ${
                            localAiBgType === 'transparent'
                              ? 'bg-emerald-800 border-emerald-800 text-white shadow-sm'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          🏁 Şeffaf (PNG)
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Canvas Retouching Brush Settings */}
            {isRetouchingActive && (
              <div className="flex flex-col gap-3 pt-3 border-t border-emerald-100/50">
                <div className="flex items-center justify-between">
                  <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">🛠️ MANUEL RÖTUŞ ARAÇLARI</span>
                  <div className="flex gap-1.5">
                    <button
                      type="button"
                      onClick={() => setRetouchTool('brush')}
                      className={`px-2.5 py-1.5 text-[9px] font-black rounded-lg transition-all cursor-pointer flex items-center gap-1 ${
                        retouchTool === 'brush' 
                          ? 'bg-emerald-700 text-white shadow-sm' 
                          : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      <Brush className="w-3.5 h-3.5" />
                      SİLİCİ (BEYAZ)
                    </button>
                    <button
                      type="button"
                      onClick={() => setRetouchTool('restore')}
                      className={`px-2.5 py-1.5 text-[9px] font-black rounded-lg transition-all cursor-pointer flex items-center gap-1 ${
                        retouchTool === 'restore' 
                          ? 'bg-emerald-700 text-white shadow-sm' 
                          : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      <RotateCcw className="w-3.5 h-3.5" />
                      GERİ YÜKLE
                    </button>
                  </div>
                </div>

                <div className="flex items-center gap-3">
                  <Sliders className="w-4 h-4 text-slate-400" />
                  <span className="text-[10px] font-bold text-slate-500 min-w-16">Fırça Boyutu:</span>
                  <input 
                    type="range"
                    min="8"
                    max="64"
                    value={brushSize}
                    onChange={(e) => setBrushSize(Number(e.target.value))}
                    className="flex-1 h-1 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-emerald-700"
                  />
                  <span className="text-[10px] font-mono font-black text-slate-700 bg-white border border-slate-200 px-2 py-0.5 rounded-md">
                    {brushSize}px
                  </span>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Save & Cancel Operations for Selected Photo */}
        {originalSrc && (
          <div className="flex gap-2.5">
            <button
              type="button"
              disabled={isUploadingToDrive || isProcessing}
              onClick={handleSave}
              className={`flex-1 py-3 bg-emerald-700 hover:bg-emerald-800 text-white font-extrabold text-[11px] uppercase tracking-widest rounded-2xl transition-all cursor-pointer shadow-md text-center active:scale-95 flex items-center justify-center gap-1.5 ${
                (isUploadingToDrive || isProcessing) ? 'opacity-50 cursor-not-allowed' : ''
              }`}
            >
              <Check className="w-4 h-4" />
              BULUTA GÖNDER VE GÜNCELLE
            </button>
            <button
              type="button"
              onClick={() => {
                setPendingFile(null);
                setOriginalSrc(null);
                setProcessedSrc(null);
                setIsRetouchingActive(false);
              }}
              className="px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 font-extrabold text-[11px] uppercase tracking-widest rounded-2xl transition-all cursor-pointer"
            >
              İPTAL
            </button>
          </div>
        )}
      </div>

      {/* Fullscreen Image Overlay */}
      <AnimatePresence>
        {isFullScreen && (currentImageUrl || originalSrc) && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/95 backdrop-blur-md z-[9999] flex items-center justify-center p-4 overflow-hidden"
            onWheel={(e) => {
              e.preventDefault();
              const zoomIntensity = 0.15;
              setFsScale(prev => {
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
              onTouchStart={(e) => {
                if (fsScale > 1 && e.touches.length === 1) {
                  setIsDraggingFs(true);
                  setDragStart({
                    x: e.touches[0].clientX - fsOffset.x,
                    y: e.touches[0].clientY - fsOffset.y
                  });
                }
              }}
              onTouchMove={(e) => {
                if (isDraggingFs && fsScale > 1 && e.touches.length === 1) {
                  setFsOffset({
                    x: e.touches[0].clientX - dragStart.x,
                    y: e.touches[0].clientY - dragStart.y
                  });
                }
              }}
              onTouchEnd={() => setIsDraggingFs(false)}
            >
              <CachedDriveImage 
                src={originalSrc || currentImageUrl} 
                alt="Teçhizat Görseli Tam Ekran" 
                className="max-w-full max-h-full object-contain rounded-2xl select-none transition-transform duration-75 ease-out"
                style={{ 
                  transform: `translate(${fsOffset.x}px, ${fsOffset.y}px) scale(${fsScale})`,
                  cursor: fsScale > 1 ? (isDraggingFs ? 'grabbing' : 'grab') : 'default'
                }}
                referrerPolicy="no-referrer"
              />
            </div>

            {/* Floating Fullscreen Zoom Controls */}
            <div className="absolute bottom-6 left-1/2 -translate-x-1/2 bg-white/10 backdrop-blur-md rounded-full px-4 py-2 flex items-center gap-3 border border-white/20 text-white select-none z-30">
              <button
                type="button"
                onClick={() => {
                  setFsScale(prev => {
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
                onClick={() => setFsScale(prev => Math.min(5, prev + 0.2))}
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
