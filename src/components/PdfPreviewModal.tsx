import React, { useState, useEffect, useRef } from 'react';
import { FileText, Download, ExternalLink, X, Loader2, Image as ImageIcon, Archive } from 'lucide-react';
import { HangarPdfDoc, getHangarPdfDocById, getFileCategory, getDocMimeType } from '../utils/hangarPdfStorage';
import { GOOGLE_SCRIPT_URL } from '../App';
import { getCachedPdfBlobUrl, cachePdfBase64 } from '../utils/pdfCache';

interface PdfPreviewModalProps {
  preview: HangarPdfDoc | null;
  onClose: () => void;
}

export const PdfPreviewModal: React.FC<PdfPreviewModalProps> = ({ preview, onClose }) => {
  const [blobUrl, setBlobUrl] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string>('');
  const currentBlobRef = useRef<string>('');

  const fileCategory = preview ? getFileCategory(preview.fileName) : 'pdf';
  const isImage = fileCategory === 'image';
  const isArchive = fileCategory === 'archive';

  useEffect(() => {
    let isCancelled = false;

    // Cleanup previous blob
    if (currentBlobRef.current && currentBlobRef.current.startsWith('blob:')) {
      try {
        URL.revokeObjectURL(currentBlobRef.current);
      } catch (e) {
        // ignore
      }
      currentBlobRef.current = '';
    }

    setBlobUrl('');
    setIsLoading(true);
    setErrorMessage('');

    if (!preview) {
      setIsLoading(false);
      return;
    }

    const docMime = getDocMimeType(preview.fileName, preview.fileData);

    const loadPdfData = async () => {
      try {
        let rawData = preview.fileData || '';

        // 1. If fileData is truncated or missing, load full doc from IndexedDB
        if (!rawData || rawData.includes('...[IDB]') || rawData.length < 500) {
          try {
            const fullDoc = await getHangarPdfDocById(preview.id);
            if (fullDoc && fullDoc.fileData && !fullDoc.fileData.includes('...[IDB]')) {
              rawData = fullDoc.fileData;
            }
          } catch (err) {
            console.warn('IndexedDB full fetch error:', err);
          }
        }

        // 2. Direct Blob or HTTP URL
        if (rawData && (rawData.startsWith('blob:') || rawData.startsWith('http://') || rawData.startsWith('https://'))) {
          if (!isCancelled) {
            setBlobUrl(rawData);
            setIsLoading(false);
          }
          return;
        }

        // 3. Fast Data URL to Blob conversion using native fetch
        if (rawData && rawData.startsWith('data:')) {
          try {
            const res = await fetch(rawData);
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            currentBlobRef.current = url;
            if (!isCancelled) {
              setBlobUrl(url);
              setIsLoading(false);
            }
            return;
          } catch (e) {
            console.warn('Native fetch blob conversion failed, trying Uint8Array fallback:', e);
            // Fallback Uint8Array parsing
            const base64Index = rawData.indexOf(';base64,');
            const base64 = base64Index !== -1 ? rawData.substring(base64Index + 8) : rawData;
            const binary = atob(base64);
            const len = binary.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
              bytes[i] = binary.charCodeAt(i);
            }
            const blob = new Blob([bytes], { type: docMime });
            const url = URL.createObjectURL(blob);
            currentBlobRef.current = url;
            if (!isCancelled) {
              setBlobUrl(url);
              setIsLoading(false);
            }
            return;
          }
        }

        // 4. Raw base64 (starts with JVBER or other base64 headers)
        if (rawData && (rawData.startsWith('JVBER') || rawData.startsWith('/9j/') || rawData.startsWith('iVBOR') || rawData.startsWith('UEsDB') || rawData.length > 200)) {
          try {
            const cleanBase64 = rawData.replace(/^data:[^;]+;base64,/, '').trim();
            const dataUri = `data:${docMime};base64,${cleanBase64}`;
            const res = await fetch(dataUri);
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            currentBlobRef.current = url;
            if (!isCancelled) {
              setBlobUrl(url);
              setIsLoading(false);
            }
            return;
          } catch (e) {
            console.warn('Raw base64 conversion failed:', e);
          }
        }

        // 5. If no valid fileData but driveFileId exists, fetch via fast server proxy or Google Apps Script
        const fileId = preview.driveFileId || (preview.driveUrl ? preview.driveUrl.match(/\/d\/([a-zA-Z0-9_-]+)/)?.[1] : null);
        if (fileId) {
          // Check IndexedDB cache first
          const cachedBlob = await getCachedPdfBlobUrl(fileId);
          if (cachedBlob && !isCancelled) {
            currentBlobRef.current = cachedBlob;
            setBlobUrl(cachedBlob);
            setIsLoading(false);
            return;
          }

          // Try direct fast proxy from internal server
          try {
            const proxyRes = await fetch(`/api/pdf-proxy?fileId=${encodeURIComponent(fileId)}&fileName=${encodeURIComponent(preview.fileName || 'Belge')}`);
            if (proxyRes.ok) {
              const blob = await proxyRes.blob();
              const url = URL.createObjectURL(blob);
              currentBlobRef.current = url;
              if (!isCancelled) {
                setBlobUrl(url);
                setIsLoading(false);
              }
              return;
            }
          } catch (proxyErr) {
            console.warn('Proxy fetch failed, falling back to GAS:', proxyErr);
          }

          // Fetch from Google Apps Script fallback
          try {
            const gasUrl = `${GOOGLE_SCRIPT_URL}?action=getPdfBase64&fileId=${encodeURIComponent(fileId)}`;
            const response = await fetch(gasUrl);
            if (response.ok) {
              const data = await response.json();
              const base64 = data?.base64 || data?.data;
              if (base64) {
                await cachePdfBase64(fileId, base64);
                const dataUri = `data:${docMime};base64,${base64}`;
                const res = await fetch(dataUri);
                const blob = await res.blob();
                const url = URL.createObjectURL(blob);
                currentBlobRef.current = url;
                if (!isCancelled) {
                  setBlobUrl(url);
                  setIsLoading(false);
                }
                return;
              }
            }
          } catch (gasErr) {
            console.warn('GAS fetch error:', gasErr);
          }
        }

        // 6. Drive URL fallback (Google Drive preview link)
        if (preview.driveUrl) {
          if (!isCancelled) {
            setBlobUrl(preview.driveUrl);
            setIsLoading(false);
          }
          return;
        }

        if (!isCancelled) {
          setErrorMessage('Belge içeriği bulunamadı veya henüz yüklenmedi.');
          setIsLoading(false);
        }
      } catch (err: any) {
        console.error('Belge Önizleme yükleme hatası:', err);
        if (!isCancelled) {
          setErrorMessage(err.message || 'Önizleme oluşturulamadı.');
          setIsLoading(false);
        }
      }
    };

    loadPdfData();

    return () => {
      isCancelled = true;
      if (currentBlobRef.current && currentBlobRef.current.startsWith('blob:')) {
        try {
          URL.revokeObjectURL(currentBlobRef.current);
        } catch (e) {
          // ignore
        }
      }
    };
  }, [preview]);

  if (!preview) return null;

  const displayUrl = blobUrl || preview.driveUrl || '';

  const handleOpenNewTab = () => {
    if (blobUrl) {
      window.open(blobUrl, '_blank');
    } else if (preview.driveUrl) {
      window.open(preview.driveUrl, '_blank');
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4 bg-slate-950/85 backdrop-blur-sm animate-fade-in">
      <div className="bg-white w-full max-w-6xl h-[94vh] rounded-2xl shadow-2xl flex flex-col overflow-hidden border border-slate-700">
        {/* Modal Header */}
        <div className="px-5 py-3 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800 shrink-0">
          <div className="flex items-center gap-3 truncate mr-4">
            <div className={`w-8 h-8 rounded-xl flex items-center justify-center shrink-0 shadow-md ${
              isImage ? 'bg-blue-600 text-white shadow-blue-600/30' :
              isArchive ? 'bg-indigo-600 text-white shadow-indigo-600/30' :
              'bg-rose-600 text-white shadow-rose-600/30'
            }`}>
              {isImage ? <ImageIcon className="w-4 h-4" /> : isArchive ? <Archive className="w-4 h-4" /> : <FileText className="w-4 h-4" />}
            </div>
            <div className="truncate">
              <div className="flex items-center gap-2">
                <h4 className="text-xs sm:text-sm font-black truncate text-white">{preview.fileName}</h4>
                <span className={`text-[9px] font-mono px-1.5 py-0.5 rounded font-black uppercase ${
                  isImage ? 'bg-blue-500/20 text-blue-300 border border-blue-500/30' :
                  isArchive ? 'bg-indigo-500/20 text-indigo-300 border border-indigo-500/30' :
                  'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                }`}>
                  {preview.fileName?.split('.').pop()?.toUpperCase() || 'BELGE'}
                </span>
              </div>
              <p className="text-[10px] sm:text-[11px] text-slate-400 font-mono flex items-center gap-2 mt-0.5">
                <span>Tür: <strong className="text-rose-400 font-bold">{preview.docType}</strong></span>
                <span>•</span>
                <span>Firma: <strong className="text-slate-200">{preview.firma || "-"}</strong></span>
                <span>•</span>
                <span>Boyut: <strong className="text-slate-200">{preview.fileSize || "-"}</strong></span>
                <span>•</span>
                <span>Tarih: <strong className="text-slate-200">{preview.uploadDate || preview.uploadedAt || "-"}</strong></span>
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            {displayUrl && (
              <button
                type="button"
                onClick={handleOpenNewTab}
                className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white font-bold text-xs rounded-xl transition-all flex items-center gap-1.5 border border-slate-700 cursor-pointer"
                title="Yeni Sekmede Aç"
              >
                <ExternalLink className="w-3.5 h-3.5" />
                <span className="hidden sm:inline">Yeni Sekmede Aç</span>
              </button>
            )}
            {displayUrl && (
              <a
                href={displayUrl}
                download={preview.fileName}
                className={`px-3.5 py-1.5 text-white font-bold text-xs rounded-xl transition-all flex items-center gap-1.5 shadow-sm cursor-pointer ${
                  isArchive ? 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-600/30' :
                  isImage ? 'bg-blue-600 hover:bg-blue-700 shadow-blue-600/30' :
                  'bg-emerald-600 hover:bg-emerald-700'
                }`}
                title="İndir"
              >
                <Download className="w-3.5 h-3.5" />
                <span>İndir</span>
              </a>
            )}
            <button
              type="button"
              onClick={onClose}
              className="w-8 h-8 rounded-full bg-slate-800 hover:bg-rose-600 text-slate-300 hover:text-white flex items-center justify-center font-bold text-sm transition-all cursor-pointer"
              title="Kapat"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Modal Viewer Body */}
        <div className="flex-1 bg-slate-900 relative flex flex-col overflow-hidden">
          {!isLoading && displayUrl ? (
            isImage ? (
              /* IMAGE PREVIEW */
              <div className="w-full h-full flex items-center justify-center p-4 bg-slate-950 overflow-auto">
                <img
                  src={displayUrl}
                  alt={preview.fileName}
                  className="max-h-full max-w-full object-contain rounded-lg shadow-2xl border border-slate-800"
                />
              </div>
            ) : isArchive ? (
              /* ZIP / RAR ARCHIVE PREVIEW CARD */
              <div className="w-full h-full flex flex-col items-center justify-center p-8 text-center bg-slate-950 text-white">
                <div className="w-20 h-20 rounded-3xl bg-indigo-900/40 border border-indigo-700/50 flex items-center justify-center text-indigo-400 mb-4 shadow-xl">
                  <Archive className="w-10 h-10 text-indigo-400" />
                </div>
                <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-indigo-950 border border-indigo-800 text-indigo-300 text-xs font-mono font-bold uppercase mb-2">
                  {preview.fileName?.toLowerCase().endsWith('.rar') ? 'RAR Sıkıştırılmış Arşiv Dosyası' : 'ZIP Sıkıştırılmış Arşiv Dosyası'}
                </div>
                <h4 className="text-base font-bold text-white mb-2 max-w-md">{preview.fileName}</h4>
                <p className="text-xs text-slate-400 max-w-md mb-6 leading-relaxed">
                  Bu belge sıkıştırılmış bir arşiv dosyasıdır (.zip / .rar). Dosya içeriğine ulaşmak için bilgisayarınıza indirebilir veya Google Drive üzerinden açabilirsiniz.
                </p>
                <div className="flex flex-wrap items-center justify-center gap-3">
                  <a
                    href={displayUrl}
                    download={preview.fileName}
                    className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl shadow-lg shadow-indigo-600/30 flex items-center gap-2 transition-all cursor-pointer"
                  >
                    <Download className="w-4 h-4" />
                    <span>Arşiv Dosyasını İndir ({preview.fileSize || 'İndir'})</span>
                  </a>
                  {preview.driveUrl && (
                    <a
                      href={preview.driveUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="px-5 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white font-bold text-xs rounded-xl border border-slate-700 flex items-center gap-2 transition-all"
                    >
                      <ExternalLink className="w-4 h-4" />
                      <span>Google Drive'da Aç</span>
                    </a>
                  )}
                </div>
              </div>
            ) : (
              /* PDF / STANDARD DOCUMENT IFRAME VIEWER */
              <iframe
                src={`${displayUrl}#toolbar=1&navpanes=0&view=FitH`}
                className="w-full h-full border-0 bg-slate-900"
                title={preview.fileName}
              />
            )
          ) : isLoading ? (
            <div className="w-full h-full flex flex-col items-center justify-center p-8 text-center bg-slate-950 text-white">
              <Loader2 className="w-10 h-10 text-rose-500 animate-spin mb-3" />
              <h5 className="text-sm font-bold text-slate-200 mb-1">{preview.fileName}</h5>
              <p className="text-xs text-rose-400 max-w-md mb-2 font-mono font-bold tracking-wider uppercase">
                BEKLEYİNİZ...
              </p>
              <p className="text-[11px] text-slate-400">
                Belge hazırlanıyor, bekleyiniz.
              </p>
            </div>
          ) : (
            <div className="w-full h-full flex flex-col items-center justify-center p-8 text-center bg-slate-950 text-white">
              <FileText className="w-12 h-12 text-slate-600 mb-3" />
              <h5 className="text-sm font-bold text-slate-300 mb-2">{preview.fileName}</h5>
              <p className="text-xs text-slate-400 max-w-md mb-4 font-mono">
                {errorMessage || "Önizleme oluşturulamadı."}
              </p>
              <div className="flex items-center gap-3">
                {preview.driveUrl && (
                  <a
                    href={preview.driveUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-xl flex items-center gap-2"
                  >
                    <ExternalLink className="w-3.5 h-3.5" />
                    <span>Drive Üzerinden Aç</span>
                  </a>
                )}
                <button
                  type="button"
                  onClick={onClose}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold text-xs rounded-xl"
                >
                  Kapat
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
