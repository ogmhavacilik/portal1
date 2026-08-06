import React, { useState, useEffect } from 'react';
import { GOOGLE_SCRIPT_URL } from '../App';

export const getEmbeddableDriveUrl = (url: string | null | undefined): string | null => {
  if (!url) return null;
  if (url.startsWith('data:') || url.startsWith('blob:')) return url;

  const dMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (dMatch && dMatch[1]) {
    return `https://drive.google.com/uc?export=download&id=${dMatch[1]}`;
  }

  const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (url.includes('drive.google.com') && idMatch && idMatch[1]) {
    return `https://drive.google.com/uc?export=download&id=${idMatch[1]}`;
  }

  if (url.includes('drive.google.com/uc') && !url.includes('export=download')) {
    const ucIdMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (ucIdMatch && ucIdMatch[1]) {
      return `https://drive.google.com/uc?export=download&id=${ucIdMatch[1]}`;
    }
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
  const [cachedUrl, setCachedUrl] = useState<string | null>(null);
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

    // Try memory and sessionStorage cache first
    const cacheKey = `cached_drive_img_${fileId}`;
    const cachedData = memoryCache.get(fileId) || sessionStorage.getItem(cacheKey);
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
          
          // Cache in memory and sessionStorage
          memoryCache.set(fileId, dataUrl);
          try {
            sessionStorage.setItem(cacheKey, dataUrl);
          } catch (e) {
            console.warn("Could not save to sessionStorage:", e);
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
        const embedUrl = getEmbeddableDriveUrl(src);
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
      {...props}
    />
  );
};
