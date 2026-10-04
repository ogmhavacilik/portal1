import React, { useState, useEffect } from 'react';
import { 
  FileSearch, 
  Search, 
  FileText, 
  Upload, 
  X, 
  CheckCircle2, 
  RefreshCw,
  FolderOpen,
  Eye,
  ExternalLink,
  ChevronRight,
  Database,
  Layers,
  ArrowLeft,
  Check
} from 'lucide-react';
import { DepoItem } from './DepoManagementModal';

interface SearchResultItem {
  page: number;
  matches: number;
  snippet: string;
  viewUrl: string;
}

interface DepoBelgeBulModalProps {
  isOpen: boolean;
  onClose: () => void;
  inventory?: DepoItem[];
  selectedItem?: DepoItem | null;
  showNotification?: (msg: string) => void;
  currentUnit?: string;
}

const UNIT_NAMES_MAP: Record<string, string> = {
  all: 'TÜM BİRİMLER',
  at802: 'AT-802',
  bell429: 'BELL 429',
  t70: 'T-70',
  c650: 'C-650',
  b360: 'B-360',
  hangar: 'HANGAR'
};

export const DepoBelgeBulModal: React.FC<DepoBelgeBulModalProps> = ({
  isOpen,
  onClose,
  inventory = [],
  selectedItem = null,
  showNotification,
  currentUnit = 'at802'
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [hasSearched, setHasSearched] = useState(false);
  const [isSearching, setIsSearching] = useState(false);
  const [searchResults, setSearchResults] = useState<SearchResultItem[]>([]);
  const [targetDocName, setTargetDocName] = useState('SERTİFİKA TARAMA_AT802.pdf');
  const [totalPages, setTotalPages] = useState<number>(0);
  const [totalMatches, setTotalMatches] = useState<number>(0);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // In-modal PDF Page Preview
  const [previewPage, setPreviewPage] = useState<number | null>(null);

  // Veri Güncelleme Modal & Upload State
  const [isUpdateModalOpen, setIsUpdateModalOpen] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [targetUploadName, setTargetUploadName] = useState('');
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);

  const cleanUnit = (currentUnit || 'at802').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  const defaultTargetName = `SERTİFİKA TARAMA_${cleanUnit}.pdf`;
  const DRIVE_FOLDER_ID = "1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP";

  // Pre-fill search if opened for a specific item, but do NOT execute search until user clicks ARA
  useEffect(() => {
    if (!isOpen) {
      setHasSearched(false);
      setSearchResults([]);
      setPreviewPage(null);
      setErrorMessage(null);
      return;
    }

    // Set default target name based on unit
    setTargetUploadName(defaultTargetName);

    // Fetch initial document metadata for this specific unit
    fetch(`/api/sertifika-tarama/info?unit=${cleanUnit}`)
      .then(res => res.json())
      .then(data => {
        if (data && data.fileName) {
          setTargetDocName(data.fileName);
          setTotalPages(data.totalPages || 0);
        }
      })
      .catch(() => {});

    if (selectedItem) {
      const pn = selectedItem.partNumber || selectedItem.pn || '';
      const queryStr = pn || selectedItem.description || '';
      setSearchQuery(queryStr);
      // Başlangıçta sonuçlar gösterilmez, sadece arama kutusuna yazılır
      setHasSearched(false);
      setSearchResults([]);
    } else {
      setHasSearched(false);
      setSearchResults([]);
    }
  }, [isOpen, selectedItem, currentUnit]);

  // Execute Background PDF Page Search on Google Drive Master "SERTİFİKA TARAMA" Document
  const handleExecuteSearch = async (queryToSearch?: string) => {
    const term = (queryToSearch !== undefined ? queryToSearch : searchQuery).trim();
    if (!term) {
      if (showNotification) showNotification('⚠️ Lütfen aramak istediğiniz bir kelime veya parça no giriniz.');
      return;
    }

    setIsSearching(true);
    setErrorMessage(null);
    setPreviewPage(null);

    try {
      const resp = await fetch(`/api/sertifika-tarama/search?q=${encodeURIComponent(term)}&unit=${cleanUnit}`);
      if (!resp.ok) {
        throw new Error(`Arama servisi yanıt vermedi (${resp.status})`);
      }

      const data = await resp.json();
      if (data.status === 'success') {
        setHasSearched(true);
        setSearchResults(data.results || []);
        setTargetDocName(data.targetDocName || defaultTargetName);
        setTotalPages(data.totalPages || 0);
        setTotalMatches(data.totalMatches || 0);

        if ((data.results || []).length > 0) {
          if (showNotification) {
            showNotification(`✅ "${term}" master belgede ${data.totalMatches} kez bulundu (${data.results.length} sayfa).`);
          }
        }
      } else {
        throw new Error(data.message || 'Arama yapılamadı.');
      }
    } catch (err: any) {
      console.error('PDF Search error:', err);
      setErrorMessage(err.message || 'Arama sırasında bir hata oluştu.');
      setHasSearched(true);
      setSearchResults([]);
    } finally {
      setIsSearching(false);
    }
  };

  // Upload or Update Master "SERTİFİKA TARAMA" PDF
  const handleUploadMasterPdf = async () => {
    if (!selectedFile) {
      alert('Lütfen yüklenecek PDF dosyasını seçiniz.');
      return;
    }

    setIsUploading(true);
    setUploadProgress(10);

    try {
      const reader = new FileReader();
      reader.onload = async () => {
        try {
          const base64Data = reader.result as string;
          setUploadProgress(40);

          const finalDocName = targetUploadName.trim() || defaultTargetName;

          const resp = await fetch('/api/sertifika-tarama/upload', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              fileName: finalDocName,
              base64Data: base64Data,
              unit: cleanUnit
            })
          });

          setUploadProgress(85);

          if (!resp.ok) {
            throw new Error(`Yükleme başarısız oldu (${resp.status})`);
          }

          const resJson = await resp.json();
          setUploadProgress(100);

          if (resJson.status === 'success') {
            setTargetDocName(resJson.targetDocName || finalDocName);
            setTotalPages(resJson.totalPages || 0);
            setIsUpdateModalOpen(false);
            setSelectedFile(null);

            if (showNotification) {
              showNotification(`✅ "${finalDocName}" Google Drive'a başarıyla yüklendi ve indekslendi!`);
            }

            // If there was an active search query, re-run search on new index
            if (searchQuery.trim()) {
              handleExecuteSearch();
            }
          } else {
            throw new Error(resJson.message || 'Yükleme tamamlanamadı.');
          }
        } catch (e: any) {
          alert('Hata: ' + e.message);
        } finally {
          setIsUploading(false);
          setUploadProgress(0);
        }
      };

      reader.readAsDataURL(selectedFile);
    } catch (err: any) {
      alert('Dosya okunamadı: ' + err.message);
      setIsUploading(false);
      setUploadProgress(0);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[10200] bg-black/80 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
      <div className="bg-slate-900 text-slate-100 rounded-3xl border border-slate-700 shadow-2xl max-w-5xl w-full max-h-[92vh] flex flex-col overflow-hidden">
        
        {/* MODAL HEADER */}
        <div className="bg-slate-950 px-6 py-4 border-b border-slate-800 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-blue-500/20 text-blue-400 rounded-2xl border border-blue-500/30">
              <FileSearch className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-sm font-black text-white tracking-wide">
                  BELGE BUL & SERTİFİKA METİN TARAMA
                </h2>
                <span className="text-[10px] bg-blue-500/20 text-blue-300 font-mono px-2.5 py-0.5 rounded-full border border-blue-500/30 font-bold uppercase">
                  {UNIT_NAMES_MAP[currentUnit] || currentUnit.toUpperCase()}
                </span>
              </div>
              <p className="text-[11px] text-slate-400">
                Google Drive master sertifika PDF'i içinde arka planda sayfa ve metin taraması
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2.5">
            {/* PORTAL STYLE "VERİ GÜNCELLE" BUTTON */}
            <button
              type="button"
              onClick={() => {
                setTargetUploadName(defaultTargetName);
                setIsUpdateModalOpen(true);
              }}
              className="flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black transition cursor-pointer shadow-md active:scale-95 border border-emerald-400/40"
              title="Google Drive'daki Master Sertifika PDF'ini Güncelle"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Veri Güncelle</span>
            </button>

            {/* CLOSE BUTTON */}
            <button
              type="button"
              onClick={onClose}
              className="p-1.5 hover:bg-slate-800 text-slate-400 hover:text-white rounded-xl transition-colors cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* TARGET MASTER DOCUMENT BANNER */}
        <div className="bg-slate-950/80 px-6 py-2.5 border-b border-slate-800/80 flex items-center justify-between text-xs text-slate-400 shrink-0">
          <div className="flex items-center gap-2">
            <span className="text-blue-400 font-bold flex items-center gap-1">
              <Database className="w-3.5 h-3.5" />
              <span>Hedef Master PDF:</span>
            </span>
            <span className="font-mono text-slate-200 font-bold bg-slate-800/90 px-2 py-0.5 rounded-lg border border-slate-700">
              {targetDocName}
            </span>
            {totalPages > 0 && (
              <span className="text-[11px] text-slate-400 font-medium">
                ({totalPages} Sayfa İndekslendi)
              </span>
            )}
          </div>
          <div className="hidden sm:flex items-center gap-1.5 text-[11px] text-slate-500 font-medium">
            <span>Yalnızca bu master PDF içinde aranır</span>
          </div>
        </div>

        {/* MODAL MAIN CONTENT */}
        <div className="p-6 overflow-y-auto flex-1 space-y-5">
          
          {/* GOOGLE-STYLE "BELGE BUL" BRANDING LOGO */}
          <div className="flex flex-col items-center justify-center py-6 text-center select-none bg-slate-950/40 rounded-3xl border border-slate-800/60 p-5">
            <div className="flex items-center gap-0.5 text-4xl sm:text-5xl font-black tracking-tight font-sans drop-shadow-sm">
              <span className="text-blue-500">B</span>
              <span className="text-red-500">e</span>
              <span className="text-yellow-500">l</span>
              <span className="text-green-500">G</span>
              <span className="text-blue-500">e</span>
              <span className="text-slate-100 mx-1"> </span>
              <span className="text-red-500">B</span>
              <span className="text-green-500">u</span>
              <span className="text-yellow-500">l</span>
            </div>
            <div className="text-[9px] sm:text-[10px] font-black tracking-[0.3em] text-slate-400 mt-2 uppercase">
              ORMAN HAVACILIK
            </div>
            <div className="text-[10px] text-slate-500 font-mono mt-1.5">
              Aktif Aranan Belge: <span className="text-blue-400 font-bold">{targetDocName}</span>
            </div>

            {/* SEARCH INPUT BAR */}
            <form
              onSubmit={(e) => {
                e.preventDefault();
                handleExecuteSearch();
              }}
              className="w-full max-w-2xl mt-6 space-y-4"
            >
              <div className="relative">
                <Search className="w-5 h-5 absolute left-4.5 top-3.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="Aranacak P/N, Seri No, Form-1 No veya kelime/rakam girin (örn: EK9052, FILTER)..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded-full pl-12 pr-12 py-3.5 text-xs sm:text-sm text-slate-100 placeholder-slate-500 focus:outline-none focus:border-blue-500 font-medium transition shadow-lg focus:ring-4 focus:ring-blue-500/15"
                />
                {searchQuery && (
                  <button
                    type="button"
                    onClick={() => {
                      setSearchQuery('');
                      setHasSearched(false);
                      setSearchResults([]);
                    }}
                    className="absolute right-4 top-3.5 text-slate-500 hover:text-slate-300 p-0.5"
                  >
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>

              <div className="flex justify-center">
                <button
                  type="submit"
                  disabled={isSearching || !searchQuery.trim()}
                  className="px-6 py-2.5 bg-slate-800 hover:bg-slate-700 active:bg-slate-600 disabled:opacity-50 disabled:cursor-not-allowed text-slate-200 border border-slate-700 rounded-lg text-xs font-bold transition shadow-sm cursor-pointer shrink-0 min-w-[150px]"
                >
                  {isSearching ? 'Belgeler Taranıyor...' : 'Belgelerde Ara'}
                </button>
              </div>
            </form>
          </div>

          {/* QUICK SUGGESTIONS FROM SELECTED ITEM OR COMMON AIR TRACTOR PARTS */}
            {selectedItem && (
              <div className="flex items-center gap-2 pt-1 text-[11px] text-slate-400 flex-wrap">
                <span className="font-bold text-slate-500">Seçili Malzeme Hızlı Ara:</span>
                {(selectedItem.partNumber || selectedItem.pn) && (
                  <button
                    type="button"
                    onClick={() => {
                      const pn = selectedItem.partNumber || selectedItem.pn || '';
                      setSearchQuery(pn);
                      handleExecuteSearch(pn);
                    }}
                    className="px-2 py-0.5 bg-slate-800 hover:bg-slate-700 text-blue-400 rounded font-mono text-[10px] border border-slate-700 transition cursor-pointer"
                  >
                    P/N: {selectedItem.partNumber || selectedItem.pn}
                  </button>
                )}
                {selectedItem.description && (
                  <button
                    type="button"
                    onClick={() => {
                      const desc = selectedItem.description || '';
                      setSearchQuery(desc);
                      handleExecuteSearch(desc);
                    }}
                    className="px-2 py-0.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded text-[10px] border border-slate-700 transition cursor-pointer truncate max-w-[200px]"
                  >
                    {selectedItem.description}
                  </button>
                )}
              </div>
            )}

          {/* MAIN DISPLAY AREA */}
          {previewPage !== null ? (
            /* IN-MODAL EMBEDDED PDF VIEWER SHOWING EXACT PAGE */
            <div className="bg-slate-950 rounded-2xl border border-slate-800 p-4 space-y-3">
              <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={() => setPreviewPage(null)}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition cursor-pointer"
                  >
                    <ArrowLeft className="w-3.5 h-3.5" />
                    <span>Sonuç Listesine Dön</span>
                  </button>
                  <div className="text-xs">
                    <span className="font-bold text-white">SAYFA {previewPage}</span>
                    <span className="text-slate-400 ml-2 font-mono text-[11px]">({targetDocName})</span>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <a
                    href={`/api/sertifika-tarama/pdf?unit=${cleanUnit}#page=${previewPage}`}
                    target="_blank"
                    rel="noreferrer"
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600/30 hover:bg-blue-600/50 text-blue-300 border border-blue-500/40 rounded-xl text-xs font-bold transition"
                  >
                    <ExternalLink className="w-3.5 h-3.5" />
                    <span>Yeni Sekmede Aç</span>
                  </a>
                </div>
              </div>
 
              {/* Embedded Chrome PDF Iframe */}
              <div className="w-full h-[520px] rounded-xl overflow-hidden border border-slate-800 bg-slate-900">
                <iframe
                  src={`/api/sertifika-tarama/pdf?unit=${cleanUnit}#page=${previewPage}`}
                  title={`Sertifika Sayfa ${previewPage}`}
                  className="w-full h-full border-none"
                />
              </div>
            </div>
          ) : isSearching ? (
            /* SEARCHING ANIMATION */
            <div className="bg-slate-950/70 rounded-2xl border border-slate-800 p-12 flex flex-col items-center justify-center space-y-4 text-center">
              <div className="relative">
                <div className="w-16 h-16 rounded-full border-4 border-blue-500/20 border-t-blue-500 animate-spin flex items-center justify-center"></div>
                <Search className="w-6 h-6 text-blue-400 absolute inset-0 m-auto" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-white">Master Belge Taranıyor...</h3>
                <p className="text-xs text-slate-400 max-w-md">
                  Google Drive üzerindeki <span className="text-blue-400 font-mono font-bold">{targetDocName}</span> taranıyor ve sayfalar indeksleniyor...
                </p>
              </div>
            </div>
          ) : !hasSearched ? (
            /* BAŞLANGIÇTA SONUÇ GÖSTERİLMEZ - CLEAN LANDING STATE */
            <div className="bg-slate-950/50 rounded-2xl border border-slate-800/80 p-10 flex flex-col items-center justify-center space-y-4 text-center">
              <div className="p-4 bg-slate-900 rounded-3xl border border-slate-800 text-slate-400 shadow-md">
                <FileSearch className="w-10 h-10 text-blue-400" />
              </div>
              <div className="space-y-1.5 max-w-lg">
                <h3 className="text-sm font-extrabold text-white">
                  Aramak İstediğiniz Kelimeyi veya Parça Numarasını Girin
                </h3>
                <p className="text-xs text-slate-400 leading-relaxed">
                  Başlangıçta gereksiz liste gösterilmez. Yukarıdaki arama kutusuna parça numarası (P/N), seri no, form no veya herhangi bir metin girip <span className="text-blue-400 font-bold">"ARA"</span> butonuna bastığınızda, arka planda Google Drive master belgesi taranarak tam eşleşen sayfalar listelenir.
                </p>
              </div>
              <div className="flex items-center gap-2 pt-2">
                <span className="text-[11px] text-slate-500 font-mono bg-slate-900 px-3 py-1 rounded-full border border-slate-800">
                  🎯 Hedef Belge: {targetDocName}
                </span>
              </div>
            </div>
          ) : searchResults.length > 0 ? (
            /* ARAMA SONUÇLARI LİSTESİ */
            <div className="space-y-3">
              <div className="flex items-center justify-between px-1">
                <div className="text-xs text-slate-300 font-bold flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                  <span>
                    "{searchQuery}" için <strong className="text-emerald-400">{totalMatches} eşleşme</strong> bulundu ({searchResults.length} sayfada)
                  </span>
                </div>
                <span className="text-[11px] text-slate-500">
                  Açmak istediğiniz sayfa kartına tıklayınız
                </span>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                {searchResults.map((item, idx) => (
                  <div
                    key={idx}
                    onClick={() => setPreviewPage(item.page)}
                    className="bg-slate-950 hover:bg-slate-900/90 border border-slate-800 hover:border-blue-500/60 rounded-2xl p-4 transition-all duration-150 cursor-pointer shadow-md group flex flex-col justify-between space-y-3"
                  >
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <span className="px-2.5 py-1 bg-blue-600/20 text-blue-400 border border-blue-500/30 rounded-lg text-xs font-black font-mono">
                            📄 SAYFA {item.page}
                          </span>
                          <span className="text-[11px] bg-slate-800 text-slate-300 font-bold px-2 py-0.5 rounded-full">
                            {item.matches} kez geçiyor
                          </span>
                        </div>
                        <Eye className="w-4 h-4 text-slate-500 group-hover:text-blue-400 transition" />
                      </div>

                      <div className="text-xs text-slate-300 font-sans leading-relaxed line-clamp-3 bg-slate-900/70 p-2.5 rounded-xl border border-slate-800/80">
                        {item.snippet}
                      </div>
                    </div>

                    <div className="flex items-center justify-between pt-1 border-t border-slate-800/80 text-[11px]">
                      <span className="text-slate-500 font-mono">{targetDocName}</span>
                      <span className="text-blue-400 group-hover:text-blue-300 font-bold flex items-center gap-1">
                        <span>Sayfayı Aç</span>
                        <ChevronRight className="w-3.5 h-3.5" />
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            /* BULUNAMADI STATE */
            <div className="bg-slate-950/70 rounded-2xl border border-slate-800 p-10 flex flex-col items-center justify-center space-y-3 text-center">
              <div className="p-3 bg-slate-900 rounded-2xl border border-slate-800 text-amber-400">
                <FileText className="w-8 h-8" />
              </div>
              <h3 className="text-sm font-bold text-white">Eşleşen Sayfa Bulunamadı</h3>
              <p className="text-xs text-slate-400 max-w-md">
                "{searchQuery}" terimi master <strong className="text-slate-200">{targetDocName}</strong> belgesi içinde bulunamadı. Lütfen kelimeyi kontrol edin veya "Veri Güncelle" ile yeni bir master PDF yükleyin.
              </p>
            </div>
          )}

        </div>

        {/* MODAL FOOTER */}
        <div className="bg-slate-950 px-6 py-3 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400 shrink-0">
          <div className="flex items-center gap-2">
            <span className="text-emerald-400 font-bold">✓ Google Drive PDF Tarama & Sayfa İndeksi Aktif</span>
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

      {/* VERİ GÜNCELLE MODALI (MASTER PDF YÜKLEME VE İNDEKSLENME) */}
      {isUpdateModalOpen && (
        <div className="fixed inset-0 z-[10500] bg-black/85 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl shadow-2xl max-w-md w-full p-6 space-y-5 text-white animate-fade-in">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-emerald-500/20 text-emerald-400 rounded-xl border border-emerald-500/30">
                  <RefreshCw className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-white">MASTER PDF VERİ GÜNCELLE</h3>
                  <p className="text-[11px] text-slate-400">Google Drive Sertifika Master Belgesini Yenile</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  if (!isUploading) {
                    setIsUpdateModalOpen(false);
                    setSelectedFile(null);
                  }
                }}
                className="text-slate-400 hover:text-white p-1"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3.5 text-xs">
              {/* TARGET FILENAME */}
              <div className="space-y-1">
                <label className="text-slate-300 font-bold">
                  Google Drive Hedef Belge Adı:
                </label>
                <input
                  type="text"
                  value={targetUploadName}
                  onChange={(e) => setTargetUploadName(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white font-mono focus:outline-none focus:border-emerald-500 font-bold"
                  placeholder="SERTİFİKA TARAMA_AT802.pdf"
                />
                <p className="text-[10px] text-slate-500">
                  Sistem otomatik olarak bu dosya adını Google Drive klasörüne yazar.
                </p>
              </div>

              {/* TARGET DRIVE FOLDER */}
              <div className="bg-slate-950 p-2.5 rounded-xl border border-slate-800 space-y-1">
                <div className="text-[10px] text-slate-400 font-medium">Hedef Google Drive Klasörü:</div>
                <div className="text-[11px] font-mono text-emerald-400 font-bold truncate">
                  1HQR_NYKhHQGA7_2W3nArI9pCh-LJasTP
                </div>
              </div>

              {/* FILE PICKER */}
              <div className="space-y-1.5">
                <label className="text-slate-300 font-bold">Yeni Master PDF Seç:</label>
                <input
                  type="file"
                  accept="application/pdf"
                  disabled={isUploading}
                  onChange={(e) => {
                    if (e.target.files && e.target.files[0]) {
                      const f = e.target.files[0];
                      setSelectedFile(f);
                    }
                  }}
                  className="w-full text-xs text-slate-400 file:mr-3 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-bold file:bg-slate-800 file:text-white hover:file:bg-slate-700 cursor-pointer border border-slate-700 rounded-xl p-1 bg-slate-950"
                />
                {selectedFile && (
                  <div className="text-[11px] text-emerald-400 font-medium flex items-center gap-1.5 pt-1">
                    <Check className="w-3.5 h-3.5" />
                    <span>Seçildi: {selectedFile.name} ({(selectedFile.size / (1024 * 1024)).toFixed(2)} MB)</span>
                  </div>
                )}
              </div>

              {/* PERCENTAGE PROGRESS BAR */}
              {isUploading && (
                <div className="space-y-1.5 bg-slate-950 p-3 rounded-xl border border-slate-800">
                  <div className="flex items-center justify-between text-xs">
                    <span className="text-slate-300 font-bold">Drive'a Yükleniyor...</span>
                    <span className="text-emerald-400 font-black font-mono">%{uploadProgress}</span>
                  </div>
                  <div className="w-full bg-slate-800 h-2.5 rounded-full overflow-hidden">
                    <div
                      className="bg-emerald-500 h-full transition-all duration-300 rounded-full"
                      style={{ width: `${uploadProgress}%` }}
                    />
                  </div>
                  <div className="text-[10px] text-slate-500 truncate">
                    Hedef: {targetUploadName}
                  </div>
                </div>
              )}
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
              <button
                type="button"
                disabled={isUploading}
                onClick={() => {
                  setIsUpdateModalOpen(false);
                  setSelectedFile(null);
                }}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                Vazgeç
              </button>
              <button
                type="button"
                disabled={isUploading || !selectedFile}
                onClick={handleUploadMasterPdf}
                className="flex items-center gap-2 px-5 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 disabled:cursor-not-allowed text-white rounded-xl text-xs font-black transition cursor-pointer shadow-md active:scale-95"
              >
                {isUploading ? (
                  <>
                    <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    <span>Yükleniyor %{uploadProgress}</span>
                  </>
                ) : (
                  <>
                    <Upload className="w-3.5 h-3.5" />
                    <span>Drive'a Yükle ve İndeksle</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};
