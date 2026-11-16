import { useEffect, useState } from 'react';

export const OGM_WIKIMEDIA_LOGO_URL =
  'https://upload.wikimedia.org/wikipedia/tr/d/d4/Orman_Genel_M%C3%BCd%C3%BCrl%C3%BC%C4%9F%C3%BC_logo.png?utm_source=tr.wikipedia.org&utm_campaign=index&utm_content=original';

export const OGM_FALLBACK_SEAL_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120"><circle cx="60" cy="60" r="56" fill="#f0fdf4" stroke="#166534" stroke-width="5"/><circle cx="60" cy="60" r="46" fill="none" stroke="#15803d" stroke-width="1.5" stroke-dasharray="3 2"/><path d="M60 26 L44 56 H54 L40 80 H80 L66 56 H76 Z" fill="#15803d"/><rect x="56" y="80" width="8" height="12" fill="#78350f"/><text x="60" y="20" text-anchor="middle" font-size="8.5" font-weight="bold" fill="#14532d" font-family="Arial,sans-serif">ORMAN GENEL MÜDÜRLÜĞÜ</text><text x="60" y="106" text-anchor="middle" font-size="7.5" font-weight="bold" fill="#14532d" font-family="Arial,sans-serif">T.C. TARIM VE ORMAN BAKANLIĞI</text></svg>`;
export const OGM_FALLBACK_SEAL_DATA_URI = `data:image/svg+xml;utf8,${encodeURIComponent(OGM_FALLBACK_SEAL_SVG)}`;

let cachedCircularLogo: string | null = null;

export const useCircularOgmLogo = (): string => {
  const [logoUri, setLogoUri] = useState<string>(cachedCircularLogo || OGM_WIKIMEDIA_LOGO_URL);

  useEffect(() => {
    if (cachedCircularLogo) {
      setLogoUri(cachedCircularLogo);
      return;
    }
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
            cachedCircularLogo = dataUri;
            setLogoUri(dataUri);
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
    renderCircularLogo(OGM_WIKIMEDIA_LOGO_URL, OGM_FALLBACK_SEAL_DATA_URI);
    return () => {
      cancelled = true;
    };
  }, []);

  return logoUri;
};
