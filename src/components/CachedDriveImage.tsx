import React, { useState, useEffect } from 'react';
import { GOOGLE_SCRIPT_URL } from '../App';

export const getEmbeddableDriveUrl = (url: string | null | undefined): string | null => {
  if (!url) return null;
  if (url.startsWith('data:') || url.startsWith('blob:')) return url;

  const fileId = getDriveFileId(url);
  if (fileId) {
    return `https://drive.google.com/thumbnail?id=${fileId}&sz=w1200`;
  }

  return url;
};

const getDriveFileId = (url: string | null | undefined): string | null => {
  if (!url) return null;
  if (url.startsWith('data:') || url.startsWith('blob:')) return null;

  const dMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (dMatch && dMatch[1]) return dMatch[1];

  const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) return idMatch[1];

  const lhMatch = url.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/);
  if (lhMatch && lhMatch[1]) return lhMatch[1];

  return null;
};

const getMimeTypeFromName = (name: string): string => {
  if (!name) return "image/png";
  const ext = name.split('.').pop()?.toLowerCase();
  if (ext === 'jpg' || ext === 'jpeg') return "image/jpeg";
  if (ext === 'gif') return "image/gif";
  if (ext === 'webp') return "image/webp";
  return "image/png";
};

// Global memory cache to prevent duplicate loads during session
const memoryCache = new Map<string, string>();

export const fetchDriveImageAsBase64 = async (src: string | null | undefined): Promise<string | null> => {
  if (!src) return null;
  if (src.startsWith('data:')) return src;

  const fileId = getDriveFileId(src);
  if (fileId) {
    const cacheKey = `cached_drive_img_${fileId}`;
    const cachedData = memoryCache.get(fileId) || localStorage.getItem(cacheKey) || sessionStorage.getItem(cacheKey);
    if (cachedData && cachedData.startsWith('data:')) return cachedData;

    try {
      const fetchUrl = `${GOOGLE_SCRIPT_URL}?action=getPdfBase64&fileId=${fileId}`;
      const res = await fetch(fetchUrl);
      if (res.ok) {
        const data = await res.json();
        if (data && data.base64) {
          const mimeType = getMimeTypeFromName(data.name || "");
          const dataUrl = `data:${mimeType};base64,${data.base64}`;
          memoryCache.set(fileId, dataUrl);
          try {
            localStorage.setItem(cacheKey, dataUrl);
            sessionStorage.setItem(cacheKey, dataUrl);
          } catch (e) {}
          return dataUrl;
        }
      }
    } catch (e) {
      console.warn("fetchDriveImageAsBase64 failed via Apps Script:", e);
    }
  }

  // Fallback: load image via DOM Image & Canvas to convert to base64 in HD
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    const embedUrl = getEmbeddableDriveUrl(src) || src;
    img.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        const maxDim = 1200; // HD resolution
        let width = img.width;
        let height = img.height;
        if (width > maxDim || height > maxDim) {
          if (width > height) {
            height = Math.round((height * maxDim) / width);
            width = maxDim;
          } else {
            width = Math.round((width * maxDim) / height);
            height = maxDim;
          }
        }
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.drawImage(img, 0, 0, width, height);
          const dataUrl = canvas.toDataURL("image/jpeg", 0.90);
          if (fileId) {
            memoryCache.set(fileId, dataUrl);
            try {
              localStorage.setItem(`cached_drive_img_${fileId}`, dataUrl);
              sessionStorage.setItem(`cached_drive_img_${fileId}`, dataUrl);
            } catch (e) {}
          }
          resolve(dataUrl);
          return;
        }
      } catch (e) {
        console.warn("Canvas export failed:", e);
      }
      resolve(null);
    };
    img.onerror = () => {
      resolve(null);
    };
    img.src = embedUrl;
  });
};

interface CachedDriveImageProps extends React.ImgHTMLAttributes<HTMLImageElement> {
  src: string | null | undefined;
  fallbackToDirect?: boolean;
}

export const CachedDriveImage: React.FC<CachedDriveImageProps> = ({ 
  src, 
  alt, 
  className, 
  style, 
  referrerPolicy = 'no-referrer', 
  fallbackToDirect = true,
  ...props 
}) => {
  const [cachedUrl, setCachedUrl] = useState<string | null>(() => {
    if (!src) return null;
    if (src.startsWith('data:') || src.startsWith('blob:')) return src;
    const fileId = getDriveFileId(src);
    if (fileId) {
      const cacheKey = `cached_drive_img_${fileId}`;
      const saved = memoryCache.get(fileId) || localStorage.getItem(cacheKey) || sessionStorage.getItem(cacheKey);
      if (saved) return saved;
      return `https://lh3.googleusercontent.com/d/${fileId}`;
    }
    return src;
  });
  const [isLoading, setIsLoading] = useState<boolean>(false);

  useEffect(() => {
    if (!src) {
      setCachedUrl(null);
      return;
    }

    // If it's already a base64 or object URL, use it directly
    if (src.startsWith('data:') || src.startsWith('blob:')) {
      setCachedUrl(src);
      return;
    }

    const fileId = getDriveFileId(src);
    if (!fileId) {
      // Not a Drive URL, fallback to regular URL loading
      setCachedUrl(src);
      return;
    }

    // Try memory, localStorage and sessionStorage cache first
    const cacheKey = `cached_drive_img_${fileId}`;
    const cachedData = memoryCache.get(fileId) || localStorage.getItem(cacheKey) || sessionStorage.getItem(cacheKey);
    if (cachedData) {
      if (!memoryCache.has(fileId)) {
        memoryCache.set(fileId, cachedData);
      }
      setCachedUrl(cachedData);
      return;
    }

    let active = true;
    setIsLoading(true);

    // Call Google Apps Script Web App to get image as base64 (CORS safe, bypasses Drive URL blocking)
    const fetchUrl = `${GOOGLE_SCRIPT_URL}?action=getPdfBase64&fileId=${fileId}`;

    fetch(fetchUrl)
      .then(res => {
        if (!res.ok) throw new Error("Network response was not ok " + res.status);
        return res.json();
      })
      .then(data => {
        if (!active) return;
        if (data && data.base64) {
          const mimeType = getMimeTypeFromName(data.name || "");
          const dataUrl = `data:${mimeType};base64,${data.base64}`;
          
          // Cache in memory and localStorage
          memoryCache.set(fileId, dataUrl);
          try {
            localStorage.setItem(cacheKey, dataUrl);
            sessionStorage.setItem(cacheKey, dataUrl);
          } catch (e) {
            console.warn("Could not save to storage:", e);
          }

          setCachedUrl(dataUrl);
        } else {
          throw new Error(data.message || "Failed to retrieve valid image base64");
        }
        setIsLoading(false);
      })
      .catch(err => {
        console.warn("Background fetch from Apps Script failed, falling back to direct Drive link:", err);
        if (!active) return;
        const embedUrl = `https://lh3.googleusercontent.com/d/${fileId}` || getEmbeddableDriveUrl(src);
        setCachedUrl(embedUrl);
        setIsLoading(false);
      });

    return () => {
      active = false;
    };
  }, [src]);

  if (isLoading && !cachedUrl) {
    return (
      <div className="w-full h-full flex flex-col items-center justify-center bg-slate-50 border border-slate-100 rounded-3xl p-4 min-h-[200px]">
        <div className="w-6 h-6 rounded-full border-2 border-emerald-600 border-t-transparent animate-spin mb-1.5" />
        <p className="text-[9px] font-black text-slate-500 uppercase tracking-widest animate-pulse">DRIVE ÖNBELLEK ALINIYOR...</p>
      </div>
    );
  }

  return (
    <img
      src={cachedUrl || undefined}
      alt={alt}
      className={className}
      style={style}
      referrerPolicy={referrerPolicy}
      onError={(e) => {
        const fileId = getDriveFileId(src);
        if (fileId && cachedUrl?.includes('thumbnail')) {
          setCachedUrl(`https://lh3.googleusercontent.com/d/${fileId}`);
        } else if (fileId && cachedUrl?.includes('lh3.googleusercontent.com')) {
          setCachedUrl(`https://drive.google.com/uc?export=view&id=${fileId}`);
        }
      }}
      {...props}
    />
  );
};
