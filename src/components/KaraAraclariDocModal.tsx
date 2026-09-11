import React, { useState } from 'react';
import { FileText, Plus, Trash2, Eye, Download, X, Car, Image as ImageIcon, Archive } from 'lucide-react';
import { VehicleDocument } from '../types';
import { getFileCategory } from '../utils/hangarPdfStorage';

interface KaraAraclariDocModalProps {
  isOpen: boolean;
  onClose: () => void;
  plate?: string;
  vehiclePlate?: string;
  vehicleName: string;
  documents: VehicleDocument[];
  onUploadDoc?: (doc: Omit<VehicleDocument, 'id' | 'uploadDate'>) => void;
  onUploadDocument?: (doc: VehicleDocument) => void;
  onDeleteDoc?: (docId: string) => void;
  onDeleteDocument?: (docId: string) => void;
}

export const KaraAraclariDocModal: React.FC<KaraAraclariDocModalProps> = ({
  isOpen,
  onClose,
  plate,
  vehiclePlate,
  vehicleName,
  documents,
  onUploadDoc,
  onUploadDocument,
  onDeleteDoc,
  onDeleteDocument,
}) => {
  const currentPlate = plate || vehiclePlate || '';
  const [docType, setDocType] = useState<'ruhsat' | 'muayene' | 'sigorta_kasko' | 'diger'>('ruhsat');
  const [title, setTitle] = useState<string>('');
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [previewDoc, setPreviewDoc] = useState<{ url: string; name: string } | null>(null);

  if (!isOpen) return null;

  const vehicleDocs = documents.filter(d => (d.plate || d.vehiclePlate || '').toUpperCase() === currentPlate.toUpperCase());

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      setSelectedFile(file);
      if (!title.trim()) {
        setTitle(file.name.replace(/\.[^/.]+$/, ''));
      }
    }
  };

  const handleUpload = () => {
    if (!selectedFile) {
      alert("Lütfen yüklenecek PDF, JPEG, ZIP veya RAR dosyasını seçiniz.");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const base64Data = reader.result as string;
      const newDocData = {
        plate: currentPlate,
        vehiclePlate: currentPlate,
        docType,
        title: title.trim() || selectedFile.name,
        fileName: selectedFile.name,
        fileSize: `${(selectedFile.size / 1024).toFixed(1)} KB`,
        fileDataUrl: base64Data
      };

      if (onUploadDoc) {
        onUploadDoc(newDocData);
      } else if (onUploadDocument) {
        onUploadDocument({
          ...newDocData,
          id: `veh_doc_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
          uploadDate: new Date().toLocaleDateString('tr-TR')
        });
      }

      setSelectedFile(null);
      setTitle('');
    };
    reader.readAsDataURL(selectedFile);
  };

  const handleDelete = (docId: string) => {
    if (onDeleteDoc) {
      onDeleteDoc(docId);
    } else if (onDeleteDocument) {
      onDeleteDocument(docId);
    }
  };

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-md animate-fade-in">
      <div className="bg-white rounded-3xl shadow-2xl max-w-3xl w-full max-h-[92vh] flex flex-col overflow-hidden border border-slate-200">
        
        {/* Header */}
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-emerald-700/80 flex items-center justify-center text-white shadow-inner">
              <Car className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-black uppercase tracking-wider flex items-center gap-2">
                ARAÇ BELGELERİ & DÖKÜMAN YÖNETİMİ
              </h3>
              <p className="text-[11px] text-slate-400 font-mono">
                Plaka: <span className="text-emerald-400 font-bold">{plate}</span> • {vehicleName}
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center font-bold text-sm transition-all cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-6 overflow-y-auto flex flex-col gap-6">
          {/* Upload Box */}
          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 flex flex-col gap-3">
            <div className="flex items-center justify-between">
              <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
                <Plus className="w-4 h-4 text-emerald-700" />
                <span>YENİ ARAÇ BELGESİ / DÖKÜMAN YÜKLE</span>
              </h4>
              <span className="text-[10px] font-mono font-bold bg-slate-200 text-slate-700 px-2 py-0.5 rounded">
                PDF • JPEG • ZIP • RAR
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label className="block text-[10px] font-black text-slate-600 uppercase mb-1">
                  Belge Türü
                </label>
                <select
                  value={docType}
                  onChange={(e) => setDocType(e.target.value as any)}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none"
                >
                  <option value="ruhsat">🚗 Araç Ruhsatı</option>
                  <option value="muayene">🛠️ TÜVTÜRK Muayene Raporu</option>
                  <option value="sigorta_kasko">📜 Sigorta / Kasko Poliçesi</option>
                  <option value="diger">📁 Diğer Teknik Belge / Arşiv</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-black text-slate-600 uppercase mb-1">
                  Belge Başlığı / Açıklama
                </label>
                <input
                  type="text"
                  placeholder="Örn: 2026 Yılı Muayene Raporu"
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[10px] font-black text-slate-600 uppercase mb-1">
                  Dosya Seç (PDF, JPEG, ZIP, RAR)
                </label>
                <input
                  type="file"
                  accept=".pdf,application/pdf,image/jpeg,image/png,image/jpg,.jpg,.jpeg,.png,.zip,application/zip,application/x-zip-compressed,.rar,application/x-rar-compressed,application/octet-stream"
                  onChange={handleFileChange}
                  className="w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-xl file:border-0 file:text-[10px] file:font-black file:bg-[#0b3d1d] file:text-white cursor-pointer"
                />
              </div>
            </div>

            <div className="flex justify-end pt-1">
              <button
                type="button"
                onClick={handleUpload}
                disabled={!selectedFile}
                className="px-5 py-2 bg-[#0b3d1d] hover:bg-[#072612] disabled:bg-slate-300 text-white font-black text-xs uppercase rounded-xl transition-all shadow-sm cursor-pointer disabled:cursor-not-allowed flex items-center gap-2"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Belgeyi Sisteme Yükle</span>
              </button>
            </div>
          </div>

          {/* Document Preview Drawer if open */}
          {previewDoc && (
            <div className="border border-slate-300 rounded-2xl overflow-hidden h-80 bg-slate-900 relative shadow-inner flex flex-col">
              <div className="flex items-center justify-between px-4 py-2 bg-slate-800 text-white text-xs border-b border-slate-700">
                <span className="font-bold truncate max-w-sm">{previewDoc.name}</span>
                <button
                  onClick={() => setPreviewDoc(null)}
                  className="p-1 hover:bg-slate-700 rounded text-slate-300 hover:text-white cursor-pointer"
                >
                  ✕ Kapat
                </button>
              </div>
              <div className="flex-1 overflow-auto flex items-center justify-center p-2">
                {(() => {
                  const cat = getFileCategory(previewDoc.name);
                  if (cat === 'image') {
                    return (
                      <img
                        src={previewDoc.url}
                        alt={previewDoc.name}
                        className="max-h-full max-w-full object-contain"
                      />
                    );
                  }
                  if (cat === 'archive') {
                    return (
                      <div className="text-center p-6 flex flex-col items-center gap-3 text-white">
                        <Archive className="w-12 h-12 text-indigo-400" />
                        <h4 className="font-bold text-sm">{previewDoc.name}</h4>
                        <p className="text-xs text-slate-400">Arşiv dosyaları doğrudan indirilebilir.</p>
                        <a
                          href={previewDoc.url}
                          download={previewDoc.name}
                          className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-2"
                        >
                          <Download className="w-4 h-4" />
                          <span>Arşivi İndir</span>
                        </a>
                      </div>
                    );
                  }
                  return (
                    <iframe
                      src={previewDoc.url}
                      className="w-full h-full border-0"
                      title="Araç Belgesi Önizleme"
                    />
                  );
                })()}
              </div>
            </div>
          )}

          {/* List of Existing Documents */}
          <div className="flex flex-col gap-2">
            <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider">
              KAYITLI BELGELER ({vehicleDocs.length})
            </h4>

            {vehicleDocs.length === 0 ? (
              <div className="p-8 text-center text-slate-400 text-xs border border-dashed border-slate-200 rounded-2xl">
                Bu araca ait henüz yüklenmiş bir belge bulunmamaktadır.
              </div>
            ) : (
              <div className="divide-y divide-slate-100 border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                {vehicleDocs.map((doc) => {
                  const cat = getFileCategory(doc.fileName || doc.title);
                  const isImg = cat === 'image';
                  const isArch = cat === 'archive';
                  const ext = (doc.fileName || doc.title)?.split('.').pop()?.toUpperCase() || 'BELGE';

                  return (
                    <div key={doc.id} className="p-3.5 bg-white flex items-center justify-between hover:bg-slate-50 transition-colors">
                      <div className="flex items-center gap-3">
                        <div className={`p-2 rounded-xl border ${
                          isImg ? 'bg-blue-50 text-blue-700 border-blue-200' :
                          isArch ? 'bg-indigo-50 text-indigo-700 border-indigo-200' :
                          'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        }`}>
                          {isImg ? <ImageIcon className="w-5 h-5" /> : isArch ? <Archive className="w-5 h-5" /> : <FileText className="w-5 h-5" />}
                        </div>
                        <div>
                          <div className="flex items-center gap-1.5">
                            <h5 className="font-bold text-xs text-slate-900">{doc.title}</h5>
                            <span className={`text-[9px] font-mono px-1 py-0.2 rounded font-bold uppercase ${
                              isImg ? 'bg-blue-100 text-blue-800' :
                              isArch ? 'bg-indigo-100 text-indigo-800' :
                              'bg-emerald-100 text-emerald-800'
                            }`}>
                              {ext}
                            </span>
                          </div>
                          <p className="text-[10px] text-slate-500 font-mono">
                            {doc.fileName} • {doc.fileSize || "Belirtilmemiş"} • {doc.uploadDate}
                          </p>
                        </div>
                      </div>

                      <div className="flex items-center gap-2">
                        {doc.fileDataUrl && (
                          <>
                            <button
                              type="button"
                              onClick={() => setPreviewDoc({ url: doc.fileDataUrl!, name: doc.fileName || doc.title })}
                              className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl flex items-center gap-1.5 transition-all cursor-pointer"
                            >
                              <Eye className="w-3.5 h-3.5" />
                              <span>Gör</span>
                            </button>
                            <a
                              href={doc.fileDataUrl}
                              download={doc.fileName || doc.title}
                              className="px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-bold text-xs rounded-xl flex items-center gap-1.5 transition-all cursor-pointer"
                            >
                              <Download className="w-3.5 h-3.5" />
                              <span>İndir</span>
                            </a>
                          </>
                        )}
                        <button
                          type="button"
                          onClick={() => handleDelete(doc.id)}
                          className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-xl transition-all cursor-pointer"
                          title="Belgeyi Sil"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
