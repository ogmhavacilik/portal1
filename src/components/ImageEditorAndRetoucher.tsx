import React, { useState, useEffect, useRef } from 'react';
import { removeBackground, preload } from '@imgly/background-removal';
import { 
  Camera, 
  Upload, 
  Check, 
  X,
  Maximize2,
  Trash2,
  Lock,
  Unlock,
  ChevronDown,
  ChevronUp,
  Wand2,
  ClipboardPaste
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { CachedDriveImage, getEmbeddableDriveUrl } from './CachedDriveImage';

// Helper function to prepare high-quality optimized image for AI processing
// Preserves crisp object details and sharp edges needed for @imgly/background-removal
const prepareOptimizedImageForAI = async (file: File, maxDimension: number = 1600): Promise<Blob | File> => {
  return new Promise((resolve) => {
    // If file is already smaller than 2MB, pass directly for maximum quality and object detection precision
    if (file.size < 2 * 1024 * 1024) {
      resolve(file);
      return;
    }
    const img = new Image();
    const objectUrl = URL.createObjectURL(file);
    img.onload = () => {
      URL.revokeObjectURL(objectUrl);
      let { width, height } = img;
      if (width <= maxDimension && height <= maxDimension) {
        resolve(file);
        return;
      }
      if (width > height) {
        height = Math.round((height * maxDimension) / width);
        width = maxDimension;
      } else {
        width = Math.round((width * maxDimension) / height);
        height = maxDimension;
      }
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(img, 0, 0, width, height);
        canvas.toBlob((blob) => {
          canvas.width = 0;
          canvas.height = 0;
          if (blob) {
            resolve(new File([blob], file.name || 'optimized.png', { type: 'image/png' }));
          } else {
            resolve(file);
          }
        }, 'image/png');
      } else {
        resolve(file);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(file);
    };
    img.src = objectUrl;
  });
};

// Helper function to detect if an image file is already a pre-cleaned cutout with transparent background
const checkIfImageHasTransparency = async (file: File): Promise<boolean> => {
  const isPngOrWebp =
    file.type === 'image/png' ||
    file.type === 'image/webp' ||
    file.name.toLowerCase().endsWith('.png') ||
    file.name.toLowerCase().endsWith('.webp');

  if (!isPngOrWebp) return false;

  return new Promise((resolve) => {
    const img = new Image();
    const objectUrl = URL.createObjectURL(file);
    img.onload = () => {
      URL.revokeObjectURL(objectUrl);
      try {
        const canvas = document.createElement('canvas');
        const maxDim = 200;
        let w = img.width;
        let h = img.height;
        if (w > maxDim || h > maxDim) {
          if (w > h) {
            h = Math.round((h * maxDim) / w);
            w = maxDim;
          } else {
            w = Math.round((w * maxDim) / h);
            h = maxDim;
          }
        }
        canvas.width = Math.max(1, w);
        canvas.height = Math.max(1, h);
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(false);
          return;
        }
        ctx.drawImage(img, 0, 0, w, h);
        const imageData = ctx.getImageData(0, 0, w, h);
        const data = imageData.data;
        let fullyTransparentCount = 0;
        const totalPixels = data.length / 4;
        for (let i = 3; i < data.length; i += 4) {
          if (data[i] < 30) {
            fullyTransparentCount++;
          }
        }
        // Require at least 8% of total pixels to be fully transparent background
        if (fullyTransparentCount / totalPixels > 0.08) {
          resolve(true);
        } else {
          resolve(false);
        }
      } catch (err) {
        resolve(false);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(false);
    };
    img.src = objectUrl;
  });
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
  partName,
  manufacturer
}) => {
  const [statusText, setStatusText] = useState<string>('');
  const [resultImageSrc, setResultImageSrc] = useState<string | null>(null);
  const [processedBlob, setProcessedBlob] = useState<Blob | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [imageScale, setImageScale] = useState<number>(1);
  const [isFullScreen, setIsFullScreen] = useState<boolean>(false);
  const [isDropHover, setIsDropHover] = useState<boolean>(false);

  // Fullscreen Zoom & Pan States
  const [fsScale, setFsScale] = useState<number>(1);
  const [fsOffset, setFsOffset] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [isDraggingFs, setIsDraggingFs] = useState<boolean>(false);
  const [dragStart, setDragStart] = useState<{ x: number; y: number }>({ x: 0, y: 0 });

  // Touch-up / Eraser Brush & Filter States
  const [isCleanToolsOpen, setIsCleanToolsOpen] = useState<boolean>(false);
  const [isEraserMode, setIsEraserMode] = useState<boolean>(false);
  const [retouchTool, setRetouchTool] = useState<'eraser' | 'restore'>('eraser');
  const [showFadedBackground, setShowFadedBackground] = useState<boolean>(true);
  const [rawOriginalImageSrc, setRawOriginalImageSrc] = useState<string | null>(null);
  const originalImgRef = useRef<HTMLImageElement | null>(null);
  const lastCoordsRef = useRef<{ x: number; y: number } | null>(null);

  const [brushSize, setBrushSize] = useState<number>(25);
  const [isDrawing, setIsDrawing] = useState<boolean>(false);
  const [undoHistory, setUndoHistory] = useState<string[]>([]);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [eraserCursor, setEraserCursor] = useState<{ x: number; y: number; displayDiameter: number } | null>(null);

  const [selectedBgColor, setSelectedBgColor] = useState<string>('transparent');
  const [progressPercent, setProgressPercent] = useState<number>(0);

  // Load original raw image into HTMLImageElement ref for Kurtar/Onar restore tool
  useEffect(() => {
    const srcToUse = rawOriginalImageSrc || currentImageUrl;
    if (srcToUse) {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        originalImgRef.current = img;
      };
      img.onerror = () => {
        const img2 = new Image();
        img2.onload = () => {
          originalImgRef.current = img2;
        };
        img2.src = srcToUse;
      };
      img.src = srcToUse;
    } else {
      originalImgRef.current = null;
    }
  }, [rawOriginalImageSrc, currentImageUrl]);

  // Handle Ctrl+V global clipboard paste
  useEffect(() => {
    const handlePaste = (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (items) {
        for (let i = 0; i < items.length; i++) {
          const item = items[i];
          if (item.type.startsWith('image/') || item.kind === 'file') {
            const file = item.getAsFile();
            if (file) {
              e.preventDefault();
              setStatusText('📋 Panodan görsel yapıştırıldı, arka plan temizleniyor...');
              handleFileProcess(file);
              return;
            }
          }
        }
      }
    };

    window.addEventListener('paste', handlePaste);
    return () => {
      window.removeEventListener('paste', handlePaste);
    };
  }, [isProcessing]);

  // Read image directly from browser clipboard on button click
  const handlePasteFromClipboard = async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          const imageType = item.types.find((t) => t.startsWith('image/'));
          if (imageType) {
            const blob = await item.getType(imageType);
            const file = new File([blob], `pasted_image_${Date.now()}.png`, { type: imageType });
            setStatusText('📋 Panodan görsel alındı, arka plan temizleniyor...');
            handleFileProcess(file);
            return;
          }
        }
      }
    } catch (err) {
      console.warn("Clipboard direct access restricted or error:", err);
    }
    setStatusText('💡 Panodaki (Word/Excel/Ekran Alıntısı) görseli yapıştırmak için klavyeden Ctrl + V yapabilirsiniz.');
  };

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

  // Background preload of AI model weights on component mount
  useEffect(() => {
    try {
      if (typeof preload === 'function') {
        preload({ model: 'isnet_fp16' }).catch(() => {});
      }
    } catch (e) {
      // Ignore preload errors
    }
  }, []);

  // Helper function to safely convert remote Drive image to a local same-origin Blob URL for canvas editing
  const initializeRetouchCanvas = async () => {
    if (resultImageSrc) return;
    if (!currentImageUrl) return;

    setStatusText('Görsel düzenleme paneline yükleniyor...');
    try {
      const embedUrl = getEmbeddableDriveUrl(currentImageUrl) || currentImageUrl;
      const res = await fetch(embedUrl);
      if (res.ok) {
        const blob = await res.blob();
        const objectUrl = URL.createObjectURL(blob);
        setResultImageSrc(objectUrl);
        setProcessedBlob(blob);
      } else {
        setResultImageSrc(currentImageUrl);
      }
    } catch (e) {
      setResultImageSrc(currentImageUrl);
    }
  };

  // Load resultImageSrc into canvas for interactive erasing and retouching
  useEffect(() => {
    if (resultImageSrc && canvasRef.current) {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        const canvas = canvasRef.current;
        if (!canvas) return;
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          ctx.globalCompositeOperation = 'source-over';
          ctx.drawImage(img, 0, 0);
          setUndoHistory((prev) => (prev.length === 0 ? [canvas.toDataURL()] : prev));
        }
      };
      img.onerror = () => {
        const img2 = new Image();
        img2.onload = () => {
          const canvas = canvasRef.current;
          if (!canvas) return;
          canvas.width = img2.width;
          canvas.height = img2.height;
          const ctx = canvas.getContext('2d');
          if (ctx) {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.globalCompositeOperation = 'source-over';
            ctx.drawImage(img2, 0, 0);
            setUndoHistory((prev) => (prev.length === 0 ? [canvas.toDataURL()] : prev));
          }
        };
        img2.src = resultImageSrc;
      };
      img.src = resultImageSrc;
    }
  }, [resultImageSrc, isFullScreen]);

  // Save Canvas to Blob & resultImageSrc
  const updateStateFromCanvas = async () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const dataUrl = canvas.toDataURL('image/png');
    setResultImageSrc(dataUrl);
    canvas.toBlob((blob) => {
      if (blob) setProcessedBlob(blob);
    }, 'image/png');
  };

  // Erase stroke handlers
  const getCanvasCoords = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    let clientX = 0;
    let clientY = 0;

    if ('touches' in e) {
      if (e.touches.length > 0) {
        clientX = e.touches[0].clientX;
        clientY = e.touches[0].clientY;
      }
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }

    return {
      x: (clientX - rect.left) * scaleX,
      y: (clientY - rect.top) * scaleY
    };
  };

  const updateEraserCursorPosition = (e: React.MouseEvent<HTMLElement> | React.TouchEvent<HTMLElement>) => {
    if (!isEraserMode || !containerRef.current || !canvasRef.current) {
      setEraserCursor(null);
      return;
    }
    const containerRect = containerRef.current.getBoundingClientRect();
    const canvasRect = canvasRef.current.getBoundingClientRect();

    let clientX = 0;
    let clientY = 0;

    if ('touches' in e) {
      if (e.touches && e.touches.length > 0) {
        clientX = e.touches[0].clientX;
        clientY = e.touches[0].clientY;
      } else {
        setEraserCursor(null);
        return;
      }
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }

    const x = clientX - containerRect.left;
    const y = clientY - containerRect.top;

    const scaleRatio = canvasRect.width / (canvasRef.current.width || 1);
    const displayDiameter = Math.max(6, brushSize * scaleRatio);

    setEraserCursor({ x, y, displayDiameter });
  };

  const paintRestoreCircle = (ctx: CanvasRenderingContext2D, x: number, y: number, canvasWidth: number, canvasHeight: number) => {
    if (!originalImgRef.current || !originalImgRef.current.complete) return;
    ctx.save();
    ctx.beginPath();
    ctx.arc(x, y, brushSize / 2, 0, Math.PI * 2, false);
    ctx.clip();
    ctx.globalCompositeOperation = 'source-over';
    ctx.drawImage(originalImgRef.current, 0, 0, canvasWidth, canvasHeight);
    ctx.restore();
  };

  const paintRestoreLine = (
    ctx: CanvasRenderingContext2D,
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    canvasWidth: number,
    canvasHeight: number
  ) => {
    if (!originalImgRef.current || !originalImgRef.current.complete) return;
    const radius = brushSize / 2;
    const dx = x2 - x1;
    const dy = y2 - y1;
    const distance = Math.hypot(dx, dy);
    const steps = Math.max(1, Math.ceil(distance / Math.max(2, radius / 3)));

    for (let i = 0; i <= steps; i++) {
      const currX = x1 + (dx * i) / steps;
      const currY = y1 + (dy * i) / steps;
      paintRestoreCircle(ctx, currX, currY, canvasWidth, canvasHeight);
    }
  };

  const startErasing = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    if (!isEraserMode) return;
    updateEraserCursorPosition(e);
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    setUndoHistory((prev) => [...prev.slice(-15), canvas.toDataURL()]);

    setIsDrawing(true);
    const coords = getCanvasCoords(e);
    lastCoordsRef.current = coords;

    if (retouchTool === 'eraser') {
      ctx.globalCompositeOperation = 'destination-out';
      ctx.beginPath();
      ctx.arc(coords.x, coords.y, brushSize / 2, 0, Math.PI * 2, false);
      ctx.fill();
    } else if (retouchTool === 'restore') {
      paintRestoreCircle(ctx, coords.x, coords.y, canvas.width, canvas.height);
    }
  };

  const drawErase = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    if (!isEraserMode) return;
    updateEraserCursorPosition(e);
    if (!isDrawing) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const coords = getCanvasCoords(e);
    const prevCoords = lastCoordsRef.current || coords;

    if (retouchTool === 'eraser') {
      ctx.globalCompositeOperation = 'destination-out';
      ctx.beginPath();
      ctx.lineWidth = brushSize;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.moveTo(prevCoords.x, prevCoords.y);
      ctx.lineTo(coords.x, coords.y);
      ctx.stroke();
    } else if (retouchTool === 'restore') {
      paintRestoreLine(ctx, prevCoords.x, prevCoords.y, coords.x, coords.y, canvas.width, canvas.height);
    }

    lastCoordsRef.current = coords;
  };

  const stopErasing = () => {
    setEraserCursor(null);
    lastCoordsRef.current = null;
    if (isDrawing) {
      setIsDrawing(false);
      updateStateFromCanvas();
    }
  };

  const handleUndoErase = () => {
    if (undoHistory.length <= 1) return;
    const previous = undoHistory[undoHistory.length - 2];
    setUndoHistory((prev) => prev.slice(0, prev.length - 1));

    const img = new Image();
    img.onload = () => {
      const canvas = canvasRef.current;
      if (!canvas) return;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.globalCompositeOperation = 'source-over';
        ctx.drawImage(img, 0, 0);
        updateStateFromCanvas();
      }
    };
    img.src = previous;
  };

  // Automatic filter to clean paper textures, faint lines & handwriting artifacts
  const handleAutoCleanPaperLines = async () => {
    if (!processedBlob) return;
    setStatusText('Kağıt ve çizgi lekeleri otomatik temizleniyor...');
    setIsProcessing(true);

    try {
      const img = new Image();
      const objectUrl = URL.createObjectURL(processedBlob);
      await new Promise((resolve) => {
        img.onload = resolve;
        img.src = objectUrl;
      });

      const canvas = document.createElement('canvas');
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(img, 0, 0);
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = imageData.data;

        for (let i = 0; i < data.length; i += 4) {
          const r = data[i];
          const g = data[i + 1];
          const b = data[i + 2];
          const a = data[i + 3];

          if (a > 0) {
            // Semi-transparent edge noise / paper shadow
            if (a < 190 && r > 110 && g > 110 && b > 110) {
              data[i + 3] = 0;
            }
            // Light paper background noise
            else if (r > 185 && g > 185 && b > 185) {
              data[i + 3] = 0;
            }
            // Thin grey line noise
            else if (r > 130 && g > 130 && b > 130 && Math.abs(r - g) < 25 && Math.abs(g - b) < 25 && a < 240) {
              data[i + 3] = 0;
            }
          }
        }

        ctx.putImageData(imageData, 0, 0);
        const cleanedBlob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
        if (cleanedBlob) {
          const newUrl = URL.createObjectURL(cleanedBlob);
          setResultImageSrc(newUrl);
          setProcessedBlob(cleanedBlob);

          if (canvasRef.current) {
            const mainCtx = canvasRef.current.getContext('2d');
            if (mainCtx) {
              mainCtx.clearRect(0, 0, canvas.width, canvas.height);
              mainCtx.drawImage(canvas, 0, 0);
            }
          }
        }
        URL.revokeObjectURL(objectUrl);
        setStatusText('Kağıt ve çizgi lekeleri başarıyla temizlendi!');
      }
    } catch (e) {
      console.error(e);
      setStatusText('Temizleme tamamlandı.');
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle photo capture or file selection
  const handleFileProcess = async (file: File) => {
    if (!file) return;

    try {
      const rawUrl = URL.createObjectURL(file);
      setRawOriginalImageSrc(rawUrl);
    } catch (e) {
      console.warn("Could not create object URL for raw file", e);
    }

    // Auto-detect if image is already transparent / background removed
    const isAlreadyTransparent = await checkIfImageHasTransparency(file);
    if (isAlreadyTransparent) {
      const rawUrl = URL.createObjectURL(file);
      setResultImageSrc(rawUrl);
      setProcessedBlob(file);
      setStatusText('⚡ Arka planı zaten temizlenmiş / şeffaf görsel algılandı. Doğrudan rötuş paneline yüklendi.');
      setProgressPercent(100);
      setIsProcessing(false);
      setIsCleanToolsOpen(true);
      return;
    }

    let currentP = 15;
    setProgressPercent(15);
    setStatusText('Görsel işleniyor (%15)...');
    setResultImageSrc(null);
    setProcessedBlob(null);
    setIsProcessing(true);
    setSelectedBgColor('transparent');

    const progressTimer = setInterval(() => {
      if (currentP < 95) {
        currentP = Math.min(95, currentP + (currentP < 50 ? 6 : 3));
        setProgressPercent(currentP);
        setStatusText(`Arka Plan Analizi: %${currentP}`);
      }
    }, 100);

    try {
      // 1. Dev fotoğrafı AI segmentasyonu için optimum yüksek çözünürlüklü (maks. 1600px) formata getir
      const processedFile = await prepareOptimizedImageForAI(file, 1600);
      currentP = Math.max(currentP, 35);
      setProgressPercent(currentP);
      setStatusText(`AI Modeli işliyor (%${currentP})...`);

      // 2. İstemci tarafında hafif ve hızlı AI modeliyle arka planı sil
      const removeFn = (window as any).imglyRemoveBackground || removeBackground;
      const blob = await removeFn(processedFile, {
        model: 'isnet_fp16', // Hızlı ve yüksek kaliteli AI modeli
        progress: (key: string, current: number, total: number) => {
          if (total > 0) {
            const calculatedPercent = Math.min(95, Math.max(35, Math.round(35 + (current / total) * 60)));
            if (calculatedPercent > currentP) {
              currentP = calculatedPercent;
              setProgressPercent(currentP);
              setStatusText(`Arka Plan Analizi: %${currentP}`);
            }
          }
        }
      });

      clearInterval(progressTimer);

      // Instant 100% completion
      setProgressPercent(100);
      setStatusText('İşlem tamamlandı! (%100)');

      // 3. İşlenen PNG Blob verisini URL'ye dönüştür
      const imageObjectURL = URL.createObjectURL(blob);

      // 4. Ekranda göster
      setResultImageSrc(imageObjectURL);
      setProcessedBlob(blob);
    } catch (error: any) {
      clearInterval(progressTimer);
      console.error('Arka plan temizleme hatası:', error);
      setStatusText('Bir hata oluştu, lütfen tekrar deneyin.');
      setProgressPercent(0);
    } finally {
      clearInterval(progressTimer);
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
            Görsel yapıştırın (Ctrl+V), seçin veya çekin; arka plan anında şeffaf yapılır.
          </p>
        </div>
      </div>

      {/* Main Container / Result Viewport */}
      <div 
        onDragOver={(e) => { e.preventDefault(); setIsDropHover(true); }}
        onDragLeave={(e) => { e.preventDefault(); setIsDropHover(false); }}
        onDrop={(e) => {
          e.preventDefault();
          setIsDropHover(false);
          const file = e.dataTransfer.files?.[0];
          if (file) handleFileProcess(file);
        }}
        className={`w-full max-h-[260px] sm:max-h-[320px] aspect-[4/3] rounded-[2rem] overflow-hidden relative flex flex-col items-center justify-center p-2 shadow-inner group transition-all ${
          isDropHover 
            ? 'border-4 border-dashed border-emerald-500 bg-emerald-50/90 scale-[1.01] ring-4 ring-emerald-500/20' 
            : 'bg-slate-50 border-2 border-slate-200'
        }`}
      >
        {isProcessing ? (
          <div className="flex flex-col items-center justify-center gap-3 p-6 text-center w-full max-w-sm">
            <div className="relative flex items-center justify-center">
              <div className="w-16 h-16 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin" />
              <span className="absolute text-sm font-black font-mono text-blue-700">
                %{progressPercent}
              </span>
            </div>

            <div className="flex flex-col items-center gap-1 w-full">
              <span className="text-xs font-black text-blue-900 uppercase tracking-wider">
                ARKA PLAN TEMİZLENİYOR (%{progressPercent})
              </span>
              
              {/* Progress Bar */}
              <div className="w-full max-w-xs bg-slate-200 h-3.5 rounded-full overflow-hidden p-0.5 shadow-inner border border-slate-300/80 my-1">
                <div
                  className="bg-gradient-to-r from-blue-500 via-indigo-600 to-blue-600 h-full rounded-full transition-all duration-200 shadow-sm"
                  style={{ width: `${Math.max(5, progressPercent)}%` }}
                />
              </div>

              <p className="text-[11px] font-bold text-slate-700">{statusText}</p>
            </div>
          </div>
        ) : resultImageSrc ? (
          <div className="w-full h-full flex flex-col items-center justify-center gap-2 relative">
            {/* Canvas / Image Display with Interactive Eraser and Kurtar/Restore Mode */}
            <div
              ref={containerRef}
              className="relative max-w-full max-h-full flex items-center justify-center overflow-hidden rounded-lg"
              onMouseEnter={updateEraserCursorPosition}
              onMouseMove={updateEraserCursorPosition}
              onMouseLeave={() => setEraserCursor(null)}
              onTouchMove={updateEraserCursorPosition}
            >
              {/* Faded Original Background Image for Silgi & Kurtar guidance */}
              {(rawOriginalImageSrc || currentImageUrl) && showFadedBackground && isEraserMode && (
                <CachedDriveImage
                  src={rawOriginalImageSrc || currentImageUrl}
                  alt=""
                  className="pointer-events-none absolute inset-0 w-full h-full object-contain opacity-35 select-none z-0 rounded-lg"
                  style={{
                    transform: `scale(${imageScale})`
                  }}
                />
              )}

              <canvas
                ref={canvasRef}
                onMouseDown={startErasing}
                onMouseMove={drawErase}
                onMouseUp={stopErasing}
                onMouseLeave={stopErasing}
                onTouchStart={startErasing}
                onTouchMove={drawErase}
                onTouchEnd={stopErasing}
                className="relative z-10"
                style={{
                  maxWidth: '100%',
                  maxHeight: '100%',
                  borderRadius: '8px',
                  border: '2px dashed #ccc',
                  display: 'block',
                  backgroundColor: selectedBgColor !== 'transparent' ? selectedBgColor : 'transparent',
                  backgroundImage:
                    selectedBgColor === 'transparent' && !(showFadedBackground && isEraserMode)
                      ? 'repeating-linear-gradient(45deg, #ccc 25%, transparent 25%, transparent 75%, #ccc 75%, #ccc), repeating-linear-gradient(45deg, #ccc 25%, #fff 25%, #fff 75%, #ccc 75%, #ccc)'
                      : 'none',
                  backgroundPosition: '0 0, 9px 9px',
                  backgroundSize: '18px 18px',
                  objectFit: 'contain',
                  transform: `scale(${imageScale})`,
                  cursor: isEraserMode ? 'none' : 'default',
                  touchAction: isEraserMode ? 'none' : 'auto'
                }}
              />

              {/* Dynamic Round Brush Cursor Overlay */}
              {isEraserMode && eraserCursor && (
                <div
                  className={`pointer-events-none absolute rounded-full border-2 shadow-lg z-30 transition-none ${
                    retouchTool === 'eraser'
                      ? 'border-red-500 bg-red-500/25'
                      : 'border-emerald-500 bg-emerald-50/25 border-emerald-500'
                  }`}
                  style={{
                    left: `${eraserCursor.x}px`,
                    top: `${eraserCursor.y}px`,
                    width: `${eraserCursor.displayDiameter}px`,
                    height: `${eraserCursor.displayDiameter}px`,
                    transform: 'translate(-50%, -50%)',
                  }}
                >
                  <div
                    className={`w-1.5 h-1.5 rounded-full absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 shadow-sm ${
                      retouchTool === 'eraser' ? 'bg-red-600' : 'bg-emerald-600'
                    }`}
                  />
                </div>
              )}

              {isEraserMode && (
                <div
                  className={`absolute top-2 left-2 text-white text-[10px] font-black px-2.5 py-1 rounded-full shadow backdrop-blur-sm pointer-events-none animate-pulse z-20 ${
                    retouchTool === 'eraser' ? 'bg-red-600/90' : 'bg-emerald-700/90'
                  }`}
                >
                  {retouchTool === 'eraser' ? `🧹 SİLGİ AKTİF (${brushSize}px)` : `🖌️ KURTAR / ONAR AKTİF (${brushSize}px)`}
                </div>
              )}
            </div>

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
            <div className="absolute bottom-2 left-1/2 -translate-x-1/2 flex items-center gap-2 z-10">
              <button
                type="button"
                onClick={() => {
                  initializeRetouchCanvas();
                  setIsCleanToolsOpen(true);
                  setIsEraserMode(true);
                }}
                className="px-3 py-1.5 bg-amber-600/90 hover:bg-amber-600 backdrop-blur-sm text-white text-[10px] font-black rounded-xl shadow-md transition-all cursor-pointer flex items-center gap-1 active:scale-95"
              >
                🧹 Silgi & Rötuş Aç
              </button>
            </div>
            <button
              type="button"
              onClick={() => {
                setFsScale(1);
                setFsOffset({ x: 0, y: 0 });
                setIsFullScreen(true);
              }}
              className="absolute top-2 right-2 w-9 h-9 bg-black/60 hover:bg-black/80 backdrop-blur-sm text-white rounded-xl flex items-center justify-center shadow-lg cursor-pointer transition-all active:scale-95 z-10 border border-white/10"
              title="Tam Ekran Görüntüle & Rötuş Yap"
            >
              <Maximize2 className="w-4 h-4 text-emerald-400" />
            </button>
          </div>
        ) : (
          <div className="text-center p-5 text-slate-400 flex flex-col items-center">
            <span className="text-4xl mb-2 block select-none">📸</span>
            <p className="text-xs font-black uppercase text-slate-700 mb-1">Görsel Bulunmamaktadır</p>
            <p className="text-[11px] font-bold text-slate-500 leading-relaxed max-w-sm text-center">
              Excel, Word veya ekran görüntüsünden kopyalanan görselleri <strong className="text-emerald-800 bg-emerald-100 px-1.5 py-0.5 rounded border border-emerald-300">Ctrl + V</strong> ile doğrudan buraya yapıştırabilirsiniz.
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

      {/* Yazı/Kağıt Lekesi Temizleme ve Manuel Rötuş Silgisi Paneli (Katlanabilir Accordion) */}
      {resultImageSrc && (
        <div className="flex flex-col bg-amber-50/70 border border-amber-200/80 rounded-2xl overflow-hidden shadow-sm transition-all animate-fade-in">
          {/* Accordion Toggle Header */}
          <button
            type="button"
            onClick={() => {
              const nextState = !isCleanToolsOpen;
              setIsCleanToolsOpen(nextState);
              if (!nextState && isEraserMode) {
                setIsEraserMode(false);
              }
            }}
            className="w-full flex items-center justify-between p-3.5 bg-amber-100/50 hover:bg-amber-100/80 transition-colors cursor-pointer text-left select-none"
          >
            <div className="flex items-center gap-2">
              <span className="text-base">🧹</span>
              <div className="flex flex-col">
                <span className="text-xs font-black text-amber-900 uppercase tracking-wider flex items-center gap-1.5">
                  Yazı, Çizgi & Leke Temizleme
                  {isEraserMode && (
                    <span className="text-[9px] font-black bg-red-600 text-white px-2 py-0.5 rounded-full animate-pulse">
                      SİLGİ AÇIK
                    </span>
                  )}
                </span>
                <span className="text-[10px] text-amber-800 font-medium">
                  {isCleanToolsOpen ? 'Araçları gizlemek için tıklayın' : 'Açmak ve hassas rötuş yapmak için tıklayın'}
                </span>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-extrabold text-amber-800 bg-amber-200/60 px-2.5 py-1 rounded-lg flex items-center gap-1">
                {isCleanToolsOpen ? (
                  <>GİZLE <ChevronUp className="w-3.5 h-3.5" /></>
                ) : (
                  <>AÇ <ChevronDown className="w-3.5 h-3.5" /></>
                )}
              </span>
            </div>
          </button>

          {/* Expanded Tools Body */}
          {isCleanToolsOpen && (
            <div className="p-3.5 pt-2 flex flex-col gap-3 border-t border-amber-200/50 bg-amber-50/40 animate-fade-in">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                {/* Otomatik Temizle Button */}
                <button
                  type="button"
                  onClick={handleAutoCleanPaperLines}
                  disabled={isProcessing}
                  className="py-2.5 px-3 bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 text-white font-extrabold text-[11px] uppercase rounded-xl transition-all shadow-sm flex items-center justify-center gap-1.5 cursor-pointer active:scale-95"
                >
                  <Wand2 className="w-3.5 h-3.5" /> Otomatik Leke & Yazıları Sil
                </button>

                {/* Manuel Silgi Toggle Button */}
                <button
                  type="button"
                  onClick={() => setIsEraserMode(!isEraserMode)}
                  className={`py-2.5 px-3 font-extrabold text-[11px] uppercase rounded-xl transition-all shadow-sm flex items-center justify-center gap-1.5 cursor-pointer active:scale-95 ${
                    isEraserMode
                      ? 'bg-red-600 text-white ring-2 ring-red-400'
                      : 'bg-slate-800 hover:bg-slate-900 text-white'
                  }`}
                >
                  {isEraserMode ? '⏹️ Silgiyi Kapat' : '🖌️ Manuel Silgi Fırçası Aç'}
                </button>
              </div>

              {/* Manuel Silgi & Kurtar/Onar Ayarları */}
              {isEraserMode && (
                <div className="flex flex-col gap-3 bg-white/95 border border-amber-200 p-3.5 rounded-xl shadow-inner animate-fade-in">
                  {/* Tool selection bar: Silgi vs Kurtar */}
                  <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2 border-b border-slate-100 pb-2.5">
                    <span className="text-[10px] font-black text-slate-700 uppercase">
                      🛠️ Rötuş Aracı Seçimi:
                    </span>
                    <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl">
                      <button
                        type="button"
                        onClick={() => setRetouchTool('eraser')}
                        className={`flex-1 sm:flex-initial text-[10px] font-extrabold px-3 py-1.5 rounded-lg transition-all flex items-center justify-center gap-1 cursor-pointer ${
                          retouchTool === 'eraser'
                            ? 'bg-red-600 text-white shadow-sm'
                            : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200'
                        }`}
                      >
                        🔴 Silgi (Sil)
                      </button>
                      <button
                        type="button"
                        onClick={() => setRetouchTool('restore')}
                        className={`flex-1 sm:flex-initial text-[10px] font-extrabold px-3 py-1.5 rounded-lg transition-all flex items-center justify-center gap-1 cursor-pointer ${
                          retouchTool === 'restore'
                            ? 'bg-emerald-700 text-white shadow-sm'
                            : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200'
                        }`}
                      >
                        🟢 Kurtar / Onar (Geri Getir)
                      </button>
                    </div>
                  </div>

                  <div className="flex items-center justify-between">
                    <span className="text-[10px] font-black text-slate-700 uppercase flex items-center gap-1.5">
                      {retouchTool === 'eraser' ? '🔴 Silgi Fırça Boyutu:' : '🟢 Kurtarma Fırça Boyutu:'}
                      <strong className={`font-mono text-xs ${retouchTool === 'eraser' ? 'text-red-600' : 'text-emerald-700'}`}>
                        {brushSize}px
                      </strong>
                    </span>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setShowFadedBackground(!showFadedBackground)}
                        className={`text-[9px] font-black px-2 py-1 rounded-lg border transition-all flex items-center gap-1 cursor-pointer ${
                          showFadedBackground
                            ? 'bg-amber-100 border-amber-300 text-amber-900'
                            : 'bg-slate-100 border-slate-200 text-slate-500'
                        }`}
                        title="Rötuş yaparken orijinal görseli arka planda silik görünümde açıp kapatır"
                      >
                        👁️ {showFadedBackground ? 'Silik Arka Plan Açık' : 'Silik Arka Plan Kapalı'}
                      </button>
                      <button
                        type="button"
                        onClick={handleUndoErase}
                        disabled={undoHistory.length <= 1}
                        className={`text-[10px] font-extrabold px-2.5 py-1 rounded-lg transition-all flex items-center gap-1 ${
                          undoHistory.length > 1
                            ? 'bg-slate-200 hover:bg-slate-300 text-slate-800 cursor-pointer'
                            : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                        }`}
                      >
                        ↩️ Geri Al
                      </button>
                    </div>
                  </div>

                  <div className="flex items-center gap-3">
                    <input
                      type="range"
                      min="5"
                      max="120"
                      value={brushSize}
                      onChange={(e) => setBrushSize(Number(e.target.value))}
                      className={`w-full cursor-pointer ${retouchTool === 'eraser' ? 'accent-red-600' : 'accent-emerald-700'}`}
                    />
                    {/* Live Resizing Circular Brush Preview */}
                    <div
                      className="flex items-center justify-center w-11 h-11 bg-slate-100 rounded-xl border border-slate-200 shrink-0 overflow-hidden shadow-inner"
                      title="Fırça Yuvarlaklığı & Boyut Önizleme"
                    >
                      <div
                        className={`rounded-full transition-all duration-75 flex items-center justify-center shadow-sm ${
                          retouchTool === 'eraser' ? 'bg-red-600/80 border border-red-800' : 'bg-emerald-600/80 border border-emerald-800'
                        }`}
                        style={{
                          width: `${Math.max(6, Math.min(38, (brushSize / 120) * 38))}px`,
                          height: `${Math.max(6, Math.min(38, (brushSize / 120) * 38))}px`,
                        }}
                      >
                        <div className="w-1 h-1 bg-white rounded-full" />
                      </div>
                    </div>
                  </div>

                  <p className="text-[10px] text-slate-600 font-medium italic leading-relaxed">
                    {retouchTool === 'eraser' ? (
                      <>💡 <strong>Silgi Modu:</strong> Arka planda silik gözüken orijinal fotoğrafa bakarak fuzuli leke, çizgi veya kalıntıları kolayca silin.</>
                    ) : (
                      <>💡 <strong>Kurtar / Onar Modu:</strong> AI modelinin yanlışlıkla sıldığı teçhizat parçalarının üzerine sürerek orijinal fotoğraftan geriye boyayın!</>
                    )}
                  </p>
                </div>
              )}
            </div>
          )}
        </div>
      )}

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
            className="fixed inset-0 bg-black/95 backdrop-blur-md z-[9999] flex flex-col items-center justify-between p-4 overflow-hidden"
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
            {/* Top Close & Info Header */}
            <div className="w-full flex items-center justify-between z-30 px-2 pt-1">
              <div className="flex items-center gap-2">
                <span className="text-sm font-black text-white uppercase tracking-wider flex items-center gap-1.5 bg-white/10 px-3 py-1.5 rounded-full border border-white/15">
                  🔍 Tam Ekran Görsel & Rötuş
                </span>
                {isEraserMode && (
                  <span
                    className={`text-[10px] font-black px-3 py-1 rounded-full shadow animate-pulse ${
                      retouchTool === 'eraser' ? 'bg-red-600 text-white' : 'bg-emerald-600 text-white'
                    }`}
                  >
                    {retouchTool === 'eraser' ? `🧹 SİLGİ AKTİF (${brushSize}px)` : `🖌️ KURTAR / ONAR AKTİF (${brushSize}px)`}
                  </span>
                )}
              </div>

              <button
                type="button"
                onClick={() => {
                  setIsFullScreen(false);
                  setFsScale(1);
                  setFsOffset({ x: 0, y: 0 });
                }}
                className="text-white hover:text-red-400 bg-white/10 hover:bg-white/20 p-2.5 rounded-full transition-all cursor-pointer border border-white/20"
                title="Kapat"
              >
                <X className="w-6 h-6" />
              </button>
            </div>

            {/* Main Interactive Canvas / Image Display in Fullscreen */}
            <div
              className="relative max-w-5xl max-h-[72vh] w-full h-full flex items-center justify-center overflow-hidden my-auto rounded-2xl bg-slate-950/80 border border-slate-800 p-2"
              onMouseDown={(e) => {
                if (!isEraserMode && fsScale > 1) {
                  setIsDraggingFs(true);
                  setDragStart({ x: e.clientX - fsOffset.x, y: e.clientY - fsOffset.y });
                }
              }}
              onMouseMove={(e) => {
                if (!isEraserMode && isDraggingFs && fsScale > 1) {
                  setFsOffset({
                    x: e.clientX - dragStart.x,
                    y: e.clientY - dragStart.y
                  });
                }
              }}
              onMouseUp={() => setIsDraggingFs(false)}
              onMouseLeave={() => setIsDraggingFs(false)}
            >
              {resultImageSrc ? (
                <div
                  ref={containerRef}
                  className="relative max-w-full max-h-full flex items-center justify-center overflow-hidden rounded-xl"
                  onMouseEnter={updateEraserCursorPosition}
                  onMouseMove={updateEraserCursorPosition}
                  onMouseLeave={() => setEraserCursor(null)}
                >
                  {(rawOriginalImageSrc || currentImageUrl) && showFadedBackground && isEraserMode && (
                    <CachedDriveImage
                      src={rawOriginalImageSrc || currentImageUrl}
                      alt=""
                      className="pointer-events-none absolute inset-0 w-full h-full object-contain opacity-35 select-none z-0 rounded-xl"
                      style={{
                        transform: `translate(${fsOffset.x}px, ${fsOffset.y}px) scale(${fsScale})`
                      }}
                    />
                  )}

                  <canvas
                    ref={canvasRef}
                    onMouseDown={startErasing}
                    onMouseMove={(e) => {
                      updateEraserCursorPosition(e);
                      drawErase(e);
                    }}
                    onMouseUp={stopErasing}
                    onTouchStart={startErasing}
                    onTouchMove={drawErase}
                    onTouchEnd={stopErasing}
                    className="relative z-10 rounded-xl"
                    style={{
                      maxWidth: '100%',
                      maxHeight: '100%',
                      objectFit: 'contain',
                      display: 'block',
                      backgroundColor: selectedBgColor !== 'transparent' ? selectedBgColor : 'transparent',
                      backgroundImage:
                        selectedBgColor === 'transparent' && !(showFadedBackground && isEraserMode)
                          ? 'repeating-linear-gradient(45deg, #ccc 25%, transparent 25%, transparent 75%, #ccc 75%, #ccc), repeating-linear-gradient(45deg, #ccc 25%, #fff 25%, #fff 75%, #ccc 75%, #ccc)'
                          : 'none',
                      backgroundPosition: '0 0, 9px 9px',
                      backgroundSize: '18px 18px',
                      cursor: isEraserMode ? 'crosshair' : (fsScale > 1 ? (isDraggingFs ? 'grabbing' : 'grab') : 'default'),
                      transform: `translate(${fsOffset.x}px, ${fsOffset.y}px) scale(${fsScale})`
                    }}
                  />

                  {isEraserMode && eraserCursor && (
                    <div
                      className={`pointer-events-none absolute rounded-full border-2 shadow-lg z-30 transition-none ${
                        retouchTool === 'eraser'
                          ? 'border-red-500 bg-red-500/25'
                          : 'border-emerald-500 bg-emerald-50/25 border-emerald-500'
                      }`}
                      style={{
                        left: `${eraserCursor.x}px`,
                        top: `${eraserCursor.y}px`,
                        width: `${eraserCursor.displayDiameter}px`,
                        height: `${eraserCursor.displayDiameter}px`,
                        transform: 'translate(-50%, -50%)',
                      }}
                    >
                      <div
                        className={`w-1.5 h-1.5 rounded-full absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 shadow-sm ${
                          retouchTool === 'eraser' ? 'bg-red-600' : 'bg-emerald-600'
                        }`}
                      />
                    </div>
                  )}
                </div>
              ) : (
                <CachedDriveImage
                  src={currentImageUrl}
                  alt="Teçhizat Görseli Tam Ekran"
                  className="max-w-full max-h-full object-contain rounded-2xl select-none transition-transform duration-75 ease-out"
                  style={{
                    transform: `translate(${fsOffset.x}px, ${fsOffset.y}px) scale(${fsScale})`,
                    cursor: fsScale > 1 ? (isDraggingFs ? 'grabbing' : 'grab') : 'default'
                  }}
                  referrerPolicy="no-referrer"
                />
              )}
            </div>

            {/* Floating Control & Retouch Toolbar at the Bottom */}
            <div className="w-full max-w-4xl bg-slate-900/90 backdrop-blur-md rounded-2xl p-3 border border-slate-700/80 text-white shadow-2xl flex flex-col gap-2 z-30">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-800 pb-2">
                {/* Tool Selection Buttons */}
                <div className="flex items-center gap-1.5 bg-slate-800/80 p-1 rounded-xl">
                  <button
                    type="button"
                    onClick={() => {
                      if (!resultImageSrc && currentImageUrl) {
                        setResultImageSrc(currentImageUrl);
                      }
                      setRetouchTool('eraser');
                      setIsEraserMode(true);
                      setIsCleanToolsOpen(true);
                    }}
                    className={`px-3 py-1.5 rounded-lg text-[10px] font-black transition-all flex items-center gap-1 cursor-pointer ${
                      isEraserMode && retouchTool === 'eraser'
                        ? 'bg-red-600 text-white shadow'
                        : 'text-slate-300 hover:text-white hover:bg-slate-700/50'
                    }`}
                  >
                    🔴 Silgi (Sil)
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      if (!resultImageSrc && currentImageUrl) {
                        setResultImageSrc(currentImageUrl);
                      }
                      setRetouchTool('restore');
                      setIsEraserMode(true);
                      setIsCleanToolsOpen(true);
                    }}
                    className={`px-3 py-1.5 rounded-lg text-[10px] font-black transition-all flex items-center gap-1 cursor-pointer ${
                      isEraserMode && retouchTool === 'restore'
                        ? 'bg-emerald-600 text-white shadow'
                        : 'text-slate-300 hover:text-white hover:bg-slate-700/50'
                    }`}
                  >
                    🟢 Kurtar / Onar
                  </button>

                  <button
                    type="button"
                    onClick={() => setIsEraserMode(false)}
                    className={`px-3 py-1.5 rounded-lg text-[10px] font-black transition-all flex items-center gap-1 cursor-pointer ${
                      !isEraserMode
                        ? 'bg-blue-600 text-white shadow'
                        : 'text-slate-300 hover:text-white hover:bg-slate-700/50'
                    }`}
                  >
                    🔍 Gezin / Yakınlaştır
                  </button>
                </div>

                {/* Auxiliary Controls */}
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setShowFadedBackground(!showFadedBackground)}
                    className={`text-[9px] font-black px-2.5 py-1.5 rounded-xl border transition-all flex items-center gap-1 cursor-pointer ${
                      showFadedBackground
                        ? 'bg-amber-500/20 border-amber-500/50 text-amber-300'
                        : 'bg-slate-800 border-slate-700 text-slate-400'
                    }`}
                  >
                    👁️ {showFadedBackground ? 'Silik Arka Plan Açık' : 'Silik Arka Plan Kapalı'}
                  </button>

                  <button
                    type="button"
                    onClick={handleUndoErase}
                    disabled={undoHistory.length <= 1}
                    className={`text-[10px] font-black px-2.5 py-1.5 rounded-xl transition-all flex items-center gap-1 ${
                      undoHistory.length > 1
                        ? 'bg-slate-800 hover:bg-slate-700 text-white cursor-pointer'
                        : 'bg-slate-800/40 text-slate-600 cursor-not-allowed'
                    }`}
                  >
                    ↩️ Geri Al
                  </button>

                  {resultImageSrc && processedBlob && (
                    <button
                      type="button"
                      onClick={() => {
                        setIsFullScreen(false);
                        handleSaveToCloud();
                      }}
                      className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-black uppercase rounded-xl transition-all cursor-pointer shadow-md active:scale-95"
                    >
                      ☁️ Buluta Kaydet
                    </button>
                  )}
                </div>
              </div>

              {/* Brush size slider when Silgi or Kurtar is active */}
              {isEraserMode && (
                <div className="flex items-center gap-3 bg-slate-800/60 p-2 rounded-xl">
                  <span className="text-[10px] font-black text-slate-300 whitespace-nowrap">
                    Fırça Boyutu: <strong className={retouchTool === 'eraser' ? 'text-red-400' : 'text-emerald-400'}>{brushSize}px</strong>
                  </span>
                  <input
                    type="range"
                    min="5"
                    max="120"
                    value={brushSize}
                    onChange={(e) => setBrushSize(Number(e.target.value))}
                    className={`w-full cursor-pointer ${retouchTool === 'eraser' ? 'accent-red-500' : 'accent-emerald-500'}`}
                  />
                  <div className="flex items-center justify-center w-8 h-8 bg-slate-900 rounded-lg border border-slate-700 shrink-0 overflow-hidden">
                    <div
                      className={`rounded-full transition-all duration-75 ${
                        retouchTool === 'eraser' ? 'bg-red-500' : 'bg-emerald-500'
                      }`}
                      style={{
                        width: `${Math.max(4, Math.min(26, (brushSize / 120) * 26))}px`,
                        height: `${Math.max(4, Math.min(26, (brushSize / 120) * 26))}px`,
                      }}
                    />
                  </div>
                </div>
              )}

              {/* Zoom Controls */}
              <div className="flex items-center justify-between text-xs text-slate-400 px-1 pt-0.5">
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setFsScale((prev) => {
                        const next = Math.max(0.8, prev - 0.2);
                        if (next <= 1) setFsOffset({ x: 0, y: 0 });
                        return next;
                      });
                    }}
                    className="w-7 h-7 bg-slate-800 hover:bg-slate-700 text-white rounded-lg flex items-center justify-center font-bold cursor-pointer transition-all"
                    title="Uzaklaştır"
                  >
                    －
                  </button>
                  <span className="text-[10px] font-mono font-bold text-slate-300 min-w-10 text-center">
                    % {Math.round(fsScale * 100)}
                  </span>
                  <button
                    type="button"
                    onClick={() => setFsScale((prev) => Math.min(5, prev + 0.2))}
                    className="w-7 h-7 bg-slate-800 hover:bg-slate-700 text-white rounded-lg flex items-center justify-center font-bold cursor-pointer transition-all"
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
                    className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-[9px] font-black cursor-pointer transition-all"
                  >
                    Sıfırla
                  </button>
                </div>

                {!resultImageSrc && currentImageUrl && (
                  <button
                    type="button"
                    onClick={() => {
                      fetch(currentImageUrl)
                        .then((res) => res.blob())
                        .then((blob) => {
                          const file = new File([blob], 'current.png', { type: blob.type || 'image/png' });
                          handleFileProcess(file);
                        });
                    }}
                    className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-[10px] font-black rounded-lg transition-all cursor-pointer shadow active:scale-95"
                  >
                    ✨ AI ile Arka Plan Temizle
                  </button>
                )}
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
};
