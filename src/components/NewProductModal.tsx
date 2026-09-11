import React, { useState, useEffect } from 'react';
import { PlusCircle, X, ShieldAlert, Sparkles } from 'lucide-react';

interface NewProductModalProps {
  isOpen: boolean;
  onClose: () => void;
  activeUnit?: string;
  defaultUnit?: string;
  activeSection?: string;
  existingDataCount?: number;
  onSave: (unit: string, newRow: string[]) => void;
}

export const NewProductModal: React.FC<NewProductModalProps> = ({
  isOpen,
  onClose,
  activeUnit,
  defaultUnit,
  activeSection,
  existingDataCount = 0,
  onSave,
}) => {
  const initialUnit = activeUnit || defaultUnit || 'at802';
  const [selectedUnit, setSelectedUnit] = useState<string>(initialUnit === 'all' ? 'at802' : initialUnit);
  const [selectedKategori, setSelectedKategori] = useState<string>(
    activeSection === 'depo_sarf' ? 'depo_sarf' :
    activeSection === 'depo_kimyasal' ? 'depo_kimyasal' :
    activeSection === 'ozel_alet' ? 'ozel_alet' :
    'yer_destek'
  );

  const isKaraAraci = selectedUnit === 'kara_araclari';
  const isDepo = selectedKategori === 'depo_sarf' || selectedKategori === 'depo_kimyasal';

  // Otomatik Sıra No hesaplama
  const [siraNo, setSiraNo] = useState<string>('');
  const [name, setName] = useState<string>('');
  const [marka, setMarka] = useState<string>('');
  const [pn, setPn] = useState<string>('');
  const [sn, setSn] = useState<string>('-');
  const [miktar, setMiktar] = useState<string>('1 ADET');
  const [yer, setYer] = useState<string>('Hangar / Takımhane');
  const [durum, setDurum] = useState<string>('FAAL');
  const [kalibTabi, setKalibTabi] = useState<string>('EVET');
  const [sonBakim, setSonBakim] = useState<string>('');
  const [gelecekBakim, setGelecekBakim] = useState<string>('');
  const [firma, setFirma] = useState<string>('-');
  const [aciklama, setAciklama] = useState<string>('');
  const [error, setError] = useState<string>('');

  // Modal açıldığında veya birim/kategori değiştiğinde otomatik en son sıradan sonraki numarayı belirle
  useEffect(() => {
    try {
      let targetKey = selectedUnit;
      if (selectedKategori === 'depo_sarf' || selectedKategori === 'depo_kimyasal') {
        targetKey = `${selectedUnit}_${selectedKategori}`;
      }
      
      let count = existingDataCount || 0;
      const rawStored = localStorage.getItem(`excel_techizat_${targetKey}_data`) || localStorage.getItem(`excel_techizat_${selectedUnit}_data`);
      if (rawStored) {
        const parsed = JSON.parse(rawStored);
        if (Array.isArray(parsed)) {
          count = parsed.length;
        }
      }
      const nextNum = (count > 0 ? count + 1 : 1);
      setSiraNo(String(nextNum));
    } catch (e) {
      setSiraNo(String(existingDataCount ? existingDataCount + 1 : 1));
    }
  }, [selectedUnit, selectedKategori, existingDataCount, isOpen]);

  // Kategoriye göre varsayılan yer / miktar güncellemeleri
  useEffect(() => {
    if (selectedUnit === 'kara_araclari') {
      setYer('ANKARA');
      setMiktar('0');
      setKalibTabi('EVET');
    } else if (selectedKategori === 'depo_sarf' || selectedKategori === 'depo_kimyasal') {
      setYer('DEPO / RAF');
      setMiktar('1 ADET');
      setKalibTabi('HAYIR');
    } else {
      setYer('Hangar / Takımhane');
      setMiktar('1 ADET');
      setKalibTabi('EVET');
    }
  }, [selectedUnit, selectedKategori]);

  if (!isOpen) return null;

  const handleSave = () => {
    if (!name.trim()) {
      setError(isKaraAraci ? 'Lütfen Araç Plakası / Tanımını giriniz.' : isDepo ? 'Lütfen Malzeme / Parça Adını giriniz.' : 'Lütfen Teçhizat Adı / Tanımını giriniz.');
      return;
    }

    let constructedRow: string[] = [];

    if (isKaraAraci) {
      // Kara Araçları formatı: SIRA NO, ARAÇ PLAKASI, MARKA, PARÇA NO (P/N), BULUNDUĞU YER, SON KM, DURUMU, KALİBRASYONA TABİ, SON BAKIM, GELECEK BAKIM, FİRMA, AÇIKLAMA, MAİL
      constructedRow = [
        siraNo.trim() || "1",
        name.trim().toUpperCase(),
        marka.trim().toUpperCase() || "-",
        pn.trim().toUpperCase() || "-",
        yer.trim().toUpperCase() || "ANKARA",
        miktar.trim() || "0",
        durum.trim().toUpperCase(),
        kalibTabi.trim().toUpperCase(),
        sonBakim.trim() || "-",
        gelecekBakim.trim() || "-",
        firma.trim() || "-",
        aciklama.trim() || "-",
        ""
      ];
    } else if (isDepo) {
      // Depo formatı: SIRA NO, MALZEME ADI, P/N, S/N, MİKTAR, RAF, DURUM, ÖMÜRLÜ, ÖMÜR TARİHİ, TEDARİKÇİ, AÇIKLAMA
      constructedRow = [
        siraNo.trim() || "1",
        name.trim(),
        pn.trim() || "-",
        sn.trim() || "-",
        miktar.trim() || "1 ADET",
        yer.trim() || "DEPO",
        durum.trim().toUpperCase(),
        kalibTabi === "EVET" ? "EVET" : "HAYIR",
        gelecekBakim.trim() || "-",
        firma.trim() || "-",
        aciklama.trim() || "-"
      ];
    } else {
      // Standart Teçhizat formatı
      constructedRow = [
        siraNo.trim() || "1",
        name.trim(),
        pn.trim() || "-",
        sn.trim() || "-",
        miktar.trim() || "1 ADET",
        yer.trim() || "Hangar / Takımhane",
        durum.trim().toUpperCase(),
        kalibTabi.trim().toUpperCase(),
        sonBakim.trim() || "-",
        gelecekBakim.trim() || "-",
        firma.trim() || "-",
        aciklama.trim() || "-",
        ""
      ];
    }

    let targetUnitKey = selectedUnit;
    if (isDepo) {
      targetUnitKey = `${selectedUnit}_${selectedKategori}`;
    }

    onSave(targetUnitKey, constructedRow);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm animate-fade-in">
      <div className="bg-white rounded-3xl p-6 sm:p-8 max-w-2xl w-full shadow-2xl border border-slate-100 flex flex-col gap-5 max-h-[92vh] overflow-y-auto">
        
        {/* Header */}
        <div className="flex items-start justify-between border-b border-slate-100 pb-4">
          <div className="flex items-center gap-3.5">
            <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-800 flex items-center justify-center shrink-0 shadow-inner">
              <PlusCircle className="w-6 h-6 text-emerald-700" />
            </div>
            <div>
              <h3 className="text-base font-black text-slate-900 uppercase tracking-wider">
                YENİ ÜRÜN / TEÇHİZAT KAYDI EKLE
              </h3>
              <p className="text-xs text-slate-500 font-medium mt-0.5">
                Seçtiğiniz kategoriye göre dinamik olarak yapılandırılır, en son sıraya eklenir ve bulutla senkronize edilir.
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-400 hover:text-slate-600 flex items-center justify-center font-bold text-sm transition-all cursor-pointer shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {error && (
          <div className="p-3 bg-rose-50 border border-rose-200 rounded-xl text-xs font-bold text-rose-700 flex items-center gap-2">
            <ShieldAlert className="w-4 h-4 shrink-0" />
            <span>{error}</span>
          </div>
        )}

        {/* Form Container */}
        <div className="bg-slate-50/70 p-5 rounded-2xl border border-slate-200/80 space-y-4 text-xs font-medium">
          
          {/* Row 1: Hava Aracı / Birim & Kategori / Bölüm */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
            <div>
              <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center gap-1.5">
                <span>✈️</span> HAVA ARACI / BİRİM:
              </label>
              <select
                value={selectedUnit}
                onChange={(e) => setSelectedUnit(e.target.value)}
                className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
              >
                <option value="at802">AIR TRACTOR (AT-802)</option>
                <option value="bell429">BELL 429 HELİKOPTER</option>
                <option value="t70">T-70 SİKORSKY HELİKOPTER</option>
                <option value="t70_bumbi_backet">T-70 BUMBİ BUCKET</option>
                <option value="t70_helitak">T-70 HELİTAK</option>
                <option value="b360">BEECHCRAFT B-360</option>
                <option value="c650">CESSNA CITATION C-650</option>
                <option value="hangar">HANGAR GENEL TEÇHİZAT</option>
                <option value="kara_araclari">🚗 KARA ARAÇLARI FİLOSU</option>
              </select>
            </div>

            <div>
              <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center gap-1.5">
                <span>🏷️</span> KATEGORİ / BÖLÜM:
              </label>
              <select
                value={selectedKategori}
                onChange={(e) => setSelectedKategori(e.target.value)}
                disabled={isKaraAraci}
                className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm disabled:bg-slate-100 disabled:text-slate-400"
              >
                {isKaraAraci ? (
                  <option value="kara_araclari">🚗 KARA ARAÇ TAKİP</option>
                ) : (
                  <>
                    <option value="yer_destek">YER DESTEK TEÇHİZATLARI</option>
                    <option value="ozel_alet">ÖZEL ALETLER</option>
                    <option value="depo_sarf">📦 SARF VE PARÇA DEPOSU</option>
                    <option value="depo_kimyasal">🧪 KİMYASAL VE YAĞ DEPOSU</option>
                  </>
                )}
              </select>
            </div>
          </div>

          {/* Row 2: Sıra No & Başlık/Tanım */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5">
            <div>
              <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center gap-1">
                <span>🔢</span> SIRA NO:
              </label>
              <input
                type="text"
                value={siraNo}
                onChange={(e) => setSiraNo(e.target.value)}
                placeholder="Örn: 1"
                className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-black text-emerald-800 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
              />
            </div>

            <div className="sm:col-span-2">
              <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                {isKaraAraci ? "ARAÇ PLAKASI / TANIMI" : isDepo ? "MALZEME / PARÇA ADI" : "TEÇHİZAT ADI / TANIMI"} <span className="text-rose-500">*</span>
              </label>
              <input
                type="text"
                placeholder={isKaraAraci ? "Örn: 06 CUK 695" : isDepo ? "Örn: O-Ring Conta Takımı / Mobil Jet Oil II" : "Örn: Tork Anahtarı 1/2 inç"}
                value={name}
                onChange={(e) => { setName(e.target.value); setError(''); }}
                className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
              />
            </div>
          </div>

          {/* DİNAMİK ALANLAR - KARA ARAÇLARI ÖZEL */}
          {isKaraAraci && (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    MARKA
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: Ford, Toyota, Isuzu, Mercedes"
                    value={marka}
                    onChange={(e) => setMarka(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    PARÇA NO (P/N) / MODEL <span className="text-slate-400 font-normal">(İsteğe Bağlı / Boş)</span>
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: Ranger 4x4 / Boş bırakılabilir"
                    value={pn}
                    onChange={(e) => setPn(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    BULUNDUĞU YER / ŞEHİR
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: ANKARA, ANTALYA, İZMİR"
                    value={yer}
                    onChange={(e) => setYer(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    SON KM Sİ / SAYAÇ
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: 125400"
                    value={miktar}
                    onChange={(e) => setMiktar(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    DURUMU
                  </label>
                  <select
                    value={durum}
                    onChange={(e) => setDurum(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  >
                    <option value="FAAL">🟢 FAAL</option>
                    <option value="GAYRİ FAAL">🔴 GAYRİ FAAL</option>
                    <option value="BAKIMDA">🟠 BAKIMDA / SERVİSTE</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    MUAYENE / BAKIMA TABİ Mİ?
                  </label>
                  <select
                    value={kalibTabi}
                    onChange={(e) => setKalibTabi(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  >
                    <option value="EVET">EVET</option>
                    <option value="HAYIR">HAYIR</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    SON MUAYENE / BAKIM TARİHİ
                  </label>
                  <input
                    type="text"
                    placeholder="gg.aa.yyyy"
                    value={sonBakim}
                    onChange={(e) => setSonBakim(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    GELECEK MUAYENE / BAKIM TARİHİ
                  </label>
                  <input
                    type="text"
                    placeholder="gg.aa.yyyy"
                    value={gelecekBakim}
                    onChange={(e) => setGelecekBakim(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    YETKİLİ SERVİS / FİRMA
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: Ford Otosan / Yetkili Servis"
                    value={firma}
                    onChange={(e) => setFirma(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    AÇIKLAMA / NOT
                  </label>
                  <input
                    type="text"
                    placeholder="Araç ile ilgili notlar..."
                    value={aciklama}
                    onChange={(e) => setAciklama(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-medium text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>
            </>
          )}

          {/* DİNAMİK ALANLAR - DEPO SARF / KİMYASAL ÖZEL */}
          {isDepo && (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    PARÇA NO (P/N)
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: 101-380010-1"
                    value={pn}
                    onChange={(e) => setPn(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    SERİ NO (S/N) / LOT NO
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: LOT-2026-B veya -"
                    value={sn}
                    onChange={(e) => setSn(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    MİKTAR (ADET / KUTU / LT)
                  </label>
                  <input
                    type="text"
                    placeholder="1 ADET"
                    value={miktar}
                    onChange={(e) => setMiktar(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    BULUNDUĞU YER / RAF
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: M-70 (8BATCH PALET1 BOX2)"
                    value={yer}
                    onChange={(e) => setYer(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    DURUMU
                  </label>
                  <select
                    value={durum}
                    onChange={(e) => setDurum(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  >
                    <option value="FAAL">🟢 FAAL / KULLANILABİLİR</option>
                    <option value="GAYRİ FAAL">🔴 GAYRİ FAAL</option>
                    <option value="RAF ÖMRÜ DOLDU">🟠 RAF ÖMRÜ DOLDU</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    ÖMÜRLÜ PARÇA MI? (RAF ÖMRÜ)
                  </label>
                  <select
                    value={kalibTabi}
                    onChange={(e) => setKalibTabi(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  >
                    <option value="HAYIR">HAYIR (SÜRESİZ)</option>
                    <option value="EVET">EVET (ÖMÜRLÜ / SKT VAR)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center gap-1.5">
                    <span>📅</span> ÖMÜR BİTİŞ / SKT TARİHİ
                  </label>
                  <input
                    type="text"
                    placeholder="gg.aa.yyyy veya -"
                    value={gelecekBakim}
                    onChange={(e) => setGelecekBakim(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    TEDARİK EDİLEN FİRMA
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: TÜBİTAK UME / TUSAŞ / Tedarikçi"
                    value={firma}
                    onChange={(e) => setFirma(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                  AÇIKLAMA / NOT
                </label>
                <input
                  type="text"
                  placeholder="İsteğe bağlı depo notu..."
                  value={aciklama}
                  onChange={(e) => setAciklama(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-medium text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                />
              </div>
            </>
          )}

          {/* DİNAMİK ALANLAR - STANDART TEÇHİZAT & ÖZEL ALET ÖZEL */}
          {!isKaraAraci && !isDepo && (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    PARÇA NO (P/N) / MODEL
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: 101-380010-1"
                    value={pn}
                    onChange={(e) => setPn(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    SERİ NO (S/N)
                  </label>
                  <input
                    type="text"
                    placeholder="-"
                    value={sn}
                    onChange={(e) => setSn(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    MİKTAR / KAPASİTE
                  </label>
                  <input
                    type="text"
                    placeholder="1 ADET"
                    value={miktar}
                    onChange={(e) => setMiktar(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    BULUNDUĞU YER
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: Hangar / Takımhane"
                    value={yer}
                    onChange={(e) => setYer(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    DURUMU
                  </label>
                  <select
                    value={durum}
                    onChange={(e) => setDurum(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  >
                    <option value="FAAL">🟢 FAAL</option>
                    <option value="GAYRİ FAAL">🔴 GAYRİ FAAL</option>
                    <option value="KISMEN FAAL">🟡 KISMEN FAAL</option>
                    <option value="BAKIMDA">🟠 BAKIMDA</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    {isKaraAraci || selectedUnit === 'hangar' ? 'BAKIMA TABİ Mİ?' : 'KALİBRASYONA / BAKIMA TABİ Mİ?'}
                  </label>
                  <select
                    value={kalibTabi}
                    onChange={(e) => setKalibTabi(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  >
                    <option value="EVET">EVET</option>
                    <option value="HAYIR">HAYIR</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center gap-1.5">
                    <span>📅</span> SON KONTROL / KALİBRASYON TARİHİ
                  </label>
                  <input
                    type="text"
                    placeholder="gg.aa.yyyy"
                    value={sonBakim}
                    onChange={(e) => setSonBakim(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1 flex items-center gap-1.5">
                    <span>📅</span> GELECEK KONTROL / KALİBRASYON / BAKIM
                  </label>
                  <input
                    type="text"
                    placeholder="gg.aa.yyyy"
                    value={gelecekBakim}
                    onChange={(e) => setGelecekBakim(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-mono font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    KONTROLÜ YAPAN FİRMA / TEDARİKÇİ
                  </label>
                  <input
                    type="text"
                    placeholder="-"
                    value={firma}
                    onChange={(e) => setFirma(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-black text-slate-700 uppercase tracking-wider mb-1">
                    AÇIKLAMA / NOT
                  </label>
                  <input
                    type="text"
                    placeholder="İsteğe bağlı ek açıklama..."
                    value={aciklama}
                    onChange={(e) => setAciklama(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl font-medium text-slate-900 focus:outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-600/10 shadow-sm"
                  />
                </div>
              </div>
            </>
          )}

        </div>

        {/* Footer */}
        <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
          <button
            type="button"
            onClick={onClose}
            className="px-6 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl cursor-pointer transition-all"
          >
            İptal
          </button>
          <button
            type="button"
            onClick={handleSave}
            className="px-6 py-2.5 bg-[#0b3d1d] hover:bg-[#072612] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg shadow-emerald-950/20 transition-all active:scale-95 cursor-pointer flex items-center gap-2"
          >
            <PlusCircle className="w-4 h-4 text-emerald-400" />
            <span>KAYDET VE EKLE</span>
          </button>
        </div>

      </div>
    </div>
  );
};
