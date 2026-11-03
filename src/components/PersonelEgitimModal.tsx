import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  GraduationCap,
  Search,
  X,
  Download,
  Plus,
  Upload,
  ExternalLink,
  CheckCircle2,
  FileCheck,
  FilePlus,
  RefreshCw,
  Building2,
  Calendar,
  User,
  BookOpen,
  Trash2,
  Edit2,
  Check,
  FileSpreadsheet,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { DEFAULT_TEKNISYEN_PERSONEL_NAMES } from '../utils/personnelData';
import { cleanAndFormatDateString, unmergeAndFillWorksheet } from '../utils/driveExcelSync';

export interface TrainingItem {
  id: string;
  egitimAdi: string;
  egitimTarihi: string;
  egitimVeren: string;
}

export interface PersonelEgitimPerson {
  id: string;
  siraNo: number;
  adiSoyadi: string;
  trainings: TrainingItem[];
  hasPdf: boolean;
  linkTitle?: string;
}

const STORAGE_KEY = 'ogm_personel_egitim_listesi_v2';
const DRIVE_FILE_NAME = 'personel_egitim_listesi.xlsx';

interface PersonelEgitimModalProps {
  isOpen: boolean;
  onClose: () => void;
  showNotification?: (message: string) => void;
}

export const PersonelEgitimModal: React.FC<PersonelEgitimModalProps> = ({
  isOpen,
  onClose,
  showNotification
}) => {
  // "başlangıç her zaman sıfır olacak" - No fake data. Start strictly empty or from user uploaded/saved data.
  const [people, setPeople] = useState<PersonelEgitimPerson[]>([]);

  const [isLoadingDrive, setIsLoadingDrive] = useState(false);
  const [syncStatusMessage, setSyncStatusMessage] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [activeMatchIdx, setActiveMatchIdx] = useState(0);

  // New Person Modal
  const [isAddPersonModalOpen, setIsAddPersonModalOpen] = useState(false);
  const [newPersonName, setNewPersonName] = useState('');
  const [newInitialTrainingName, setNewInitialTrainingName] = useState('');
  const [newInitialTrainingDate, setNewInitialTrainingDate] = useState(() => {
    return new Date().toLocaleDateString('tr-TR');
  });
  const [newInitialTrainingProvider, setNewInitialTrainingProvider] = useState('');

  // Inline "Eğitim Ekle" State for a specific person
  const [inlineAddingPersonId, setInlineAddingPersonId] = useState<string | null>(null);
  const [inlineTrainingName, setInlineTrainingName] = useState('');
  const [inlineTrainingDate, setInlineTrainingDate] = useState(() => new Date().toLocaleDateString('tr-TR'));
  const [inlineTrainingProvider, setInlineTrainingProvider] = useState('');

  // Prompt state for "PDF Yüklediniz mi?" with checkbox & EVET button
  const [promptPersonId, setPromptPersonId] = useState<string | null>(null);
  const [isPdfConfirmedChecked, setIsPdfConfirmedChecked] = useState<boolean>(true);

  // Accordion state for "Diğer Eğitimleri Gör" dropdown
  const [expandedPersonIds, setExpandedPersonIds] = useState<Set<string>>(new Set());

  const toggleExpandPerson = (personId: string) => {
    setExpandedPersonIds(prev => {
      const next = new Set(prev);
      if (next.has(personId)) {
        next.delete(personId);
      } else {
        next.add(personId);
      }
      return next;
    });
  };

  const expandAll = () => {
    setExpandedPersonIds(new Set(people.map(p => p.id)));
  };

  const collapseAll = () => {
    setExpandedPersonIds(new Set());
  };

  // File input ref for "VERİ EKLE (EXCEL YÜKLE)"
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Save to backend/Drive on change removed to rely on explicit upload actions
  useEffect(() => {
  }, [people]);

  // Load live Excel from Google Drive / Local on opening if empty
  const fetchLiveExcelFromDrive = async (isManual = false, customMsg = '') => {
    setIsLoadingDrive(true);
    setSyncStatusMessage(customMsg || (isManual ? "Google Drive Exceli Aranıyor..." : "Veriler Yükleniyor..."));
    if (showNotification && isManual) {
      showNotification("🔄 Google Drive'daki Personel Eğitim Excel dosyası aranıyor...");
    }

    try {
      const res = await fetch('/api/read-excel-from-drive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileName: DRIVE_FILE_NAME })
      });

      if (res.ok) {
        const json = await res.json();
        console.log("DEBUG [Frontend]: Backend response received:", json);
        if (json && json.status === 'success' && json.base64) {
          parseExcelAndSetPeople(json.base64, false);
          if (showNotification) {
            showNotification("✅ Google Drive'dan Personel Eğitim Listesi başarıyla yüklendi.");
          }
          return;
        }
      }

      } catch (err) {
        console.warn("Drive Excel fetch failed:", err);
      } finally {
        setIsLoadingDrive(false);
        setSyncStatusMessage('');
      }
  };

  useEffect(() => {
    if (isOpen) {
      fetchLiveExcelFromDrive(false);
    }
  }, [isOpen]);

  // Helper to parse base64 Excel and populate people
  const parseExcelAndSetPeople = (base64: string, syncToDrive = false) => {
    try {
      const binary = atob(base64);
      const len = binary.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      const wb = XLSX.read(bytes, { type: 'array' });
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];

      // Unmerge cells so every row has complete data
      unmergeAndFillWorksheet(sheet);

      const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!rawRows || rawRows.length < 2) return;

      // Find header row (must contain ADI/PERSONEL/SOYADI and EĞİTİM/SIRA with >= 2 columns)
      let headerRowIdx = -1;
      for (let r = 0; r < Math.min(15, rawRows.length); r++) {
        const rowArr = (rawRows[r] || []).map(c => String(c || '').toUpperCase().trim());
        const rowStr = rowArr.join(' ');
        const nonEmp = rowArr.filter(c => c.length > 0);
        if (nonEmp.length >= 2 && (rowStr.includes('ADI') || rowStr.includes('SOYADI') || rowStr.includes('PERSONEL')) && (rowStr.includes('EĞİTİM') || rowStr.includes('EGITIM') || rowStr.includes('SIRA') || rowStr.includes('TARİH') || rowStr.includes('TARIH'))) {
          headerRowIdx = r;
          break;
        }
      }

      if (headerRowIdx < 0) {
        for (let r = 0; r < Math.min(15, rawRows.length); r++) {
          const nonEmp = (rawRows[r] || []).filter(c => String(c || '').trim().length > 0);
          if (nonEmp.length >= 3) {
            headerRowIdx = r;
            break;
          }
        }
      }
      if (headerRowIdx < 0) headerRowIdx = 0;

      const normHeader = (rawRows[headerRowIdx] || []).map(c =>
        String(c || '').toUpperCase().trim()
          .replace(/Ç/g, 'C').replace(/Ğ/g, 'G').replace(/İ/g, 'I').replace(/I/g, 'I').replace(/Ö/g, 'O').replace(/Ş/g, 'S').replace(/Ü/g, 'U')
      );

      let siraCol = normHeader.findIndex(c => c.includes('SIRA') || c === 'NO');
      let adCol = normHeader.findIndex(c => c.includes('ADI') || c.includes('SOYADI') || c.includes('PERSONEL'));
      let egtCol = normHeader.findIndex(c => c.includes('EGITIM') || c.includes('KURS') || c.includes('KONU'));
      let tarCol = normHeader.findIndex(c => c.includes('TARIH') || c.includes('DONEM'));
      let kurumCol = normHeader.findIndex(c => c.includes('KURUM') || c.includes('VEREN') || c.includes('EGITMEN'));
      let linkCol = normHeader.findIndex(c => c.includes('LINK') || c.includes('BELGE') || c.includes('DURUM') || c.includes('SERTIFIKA'));

      if (siraCol < 0) siraCol = 0;
      if (adCol < 0) adCol = 1;
      if (egtCol < 0) egtCol = 2;
      if (tarCol < 0) tarCol = 3;
      if (kurumCol < 0) kurumCol = 4;
      if (linkCol < 0) linkCol = 5;

      const parsedPeople: PersonelEgitimPerson[] = [];
      let lastPerson: PersonelEgitimPerson | null = null;

      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;

        const rawSira = String(siraCol >= 0 && row[siraCol] !== undefined ? row[siraCol] : row[0] || '').trim();
        const rawAd = String(adCol >= 0 && row[adCol] !== undefined ? row[adCol] : row[1] || '').trim();
        const rawEgt = String(egtCol >= 0 && row[egtCol] !== undefined ? row[egtCol] : row[2] || '').trim();
        const rawTar = String(tarCol >= 0 && row[tarCol] !== undefined ? row[tarCol] : row[3] || '').trim();
        const rawKurum = String(kurumCol >= 0 && row[kurumCol] !== undefined ? row[kurumCol] : row[4] || '').trim();
        const rawLink = String(linkCol >= 0 && row[linkCol] !== undefined ? row[linkCol] : row[5] || '').trim();

        const adUpper = rawAd.toUpperCase();
        
        let person: PersonelEgitimPerson | null = null;
        
        if (rawAd && adUpper !== 'ADI SOYADI' && adUpper !== 'PERSONEL' && adUpper !== 'ADI' && !adUpper.includes('ORMAN GENEL') && !adUpper.includes('PERSONEL EĞİTİM') && !adUpper.includes('PERSONEL EGITIM')) {
          console.log("Searching for:", adUpper, "in parsedPeople count:", parsedPeople.length);
          person = parsedPeople.find(p => p.adiSoyadi.toUpperCase() === adUpper) || null;
          if (!person) {
            const numSira = parseInt(rawSira) || (parsedPeople.length + 1);
            const hasPdf = rawLink.toUpperCase().includes('GÖR') || rawLink.toUpperCase().includes('GOR') || rawLink.toUpperCase().includes('SERTİFİKA') || rawLink.toUpperCase().includes('EVET');
            person = {
              id: `person-${Date.now()}-${parsedPeople.length + 1}`,
              siraNo: numSira,
              adiSoyadi: adUpper,
              trainings: [],
              hasPdf: hasPdf,
              linkTitle: rawLink || `${adUpper} EĞİTİM SERTİFİKALARI`
            };
            parsedPeople.push(person);
          }
          lastPerson = person;
        } else {
          person = lastPerson;
        }

        if (person && rawEgt && rawEgt !== '-' && rawEgt.toUpperCase() !== 'EĞİTİMİN ADI' && rawEgt.toUpperCase() !== 'EGITIMIN ADI') {
          const egtUpper = rawEgt.toUpperCase();
          const egtTarVal = cleanAndFormatDateString(rawTar) || rawTar || '-';
          const egtKurumVal = (rawKurum || 'OGM HAVACILIK').toUpperCase();

          if (!person.trainings.some(t => t.egitimAdi === egtUpper && t.egitimTarihi === egtTarVal)) {
            person.trainings.push({
              id: `tr-${Date.now()}-${r}`,
              egitimAdi: egtUpper,
              egitimTarihi: egtTarVal,
              egitimVeren: egtKurumVal
            });
          }
        }
      }

      if (parsedPeople.length > 0) {
        setPeople(parsedPeople);
        if (syncToDrive) {
          uploadExcelToGoogleDrive(base64);
        }
      } else {
        alert("Yüklenen Excel dosyasında geçerli personel eğitimi verisi okunamadı. Lütfen sütun başlıklarını kontrol ediniz.");
      }
    } catch (err) {
      console.error("Excel parse error:", err);
      alert("Excel dosyası okunurken hata oluştu: " + err);
    }
  };

  // Upload an Excel file via "VERİ EKLE (EXCEL YÜKLE)"
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (showNotification) {
      showNotification(`📤 "${file.name}" taranıyor ve Google Drive'a kaydediliyor...`);
    }

    try {
      const arrayBuffer = await file.arrayBuffer();
      const bytes = new Uint8Array(arrayBuffer);
      let binary = '';
      for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64 = btoa(binary);

      parseExcelAndSetPeople(base64, true);

      if (showNotification) {
        showNotification(`✅ Excel başarıyla yüklendi, kalıcı veritabanı güncellendi ve Google Drive'a yedeklendi.`);
      }
    } catch (err: any) {
      alert("Excel yükleme hatası: " + err.message);
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // Upload generated or updated Excel to Google Drive & local disk
  const uploadExcelToGoogleDrive = async (base64Data?: string, actionMsg: string = "Excel Senkronize Ediliyor...") => {
    setIsLoadingDrive(true);
    setSyncStatusMessage(actionMsg);
    try {
      let b64 = base64Data;
      if (!b64) {
        b64 = buildExcelBase64FromPeople(people);
      }

      // Save to primary personel_egitim_listesi.xlsx
      await fetch('/api/upload-techizat-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: DRIVE_FILE_NAME,
          base64Data: b64,
          targetKey: 'personel_egitim_listesi'
        })
      });

      // Immediately re-fetch to ensure UI matches the actual Drive file content and wait until complete
      await fetchLiveExcelFromDrive(false, actionMsg + " (Güncel Veri Çekiliyor)");
    } catch (e) {
      console.warn("Drive sync failed:", e);
    } finally {
      setIsLoadingDrive(false);
      setSyncStatusMessage('');
    }
  };

  // Build Excel base64 from people
  const buildExcelBase64FromPeople = (peopleList: PersonelEgitimPerson[]): string => {
    const aoa: any[][] = [
      ['SIRA NO', 'ADI SOYADI', 'EĞİTİMİN ADI', 'EĞİTİM TARİHİ', 'EĞİTİM VEREN KURUM/PERSONEL', 'LİNK']
    ];
    const merges: XLSX.Range[] = [];
    let currentRow = 1;

    peopleList.forEach(p => {
      const docText = p.hasPdf ? `BELGE GÖR (https://bulut.ogm.gov.tr/BELGEGOR)` : `BELGE YÜKLE (https://bulut.ogm.gov.tr/BELGEYUKLE)`;
      if (p.trainings.length === 0) {
        aoa.push([p.siraNo, p.adiSoyadi, '-', '-', '-', docText]);
        currentRow++;
      } else {
        const startRow = currentRow;
        const count = p.trainings.length;
        p.trainings.forEach((t, tIdx) => {
          aoa.push([
            tIdx === 0 ? p.siraNo : '',
            tIdx === 0 ? p.adiSoyadi : '',
            t.egitimAdi,
            t.egitimTarihi,
            t.egitimVeren,
            tIdx === 0 ? docText : ''
          ]);
          currentRow++;
        });
        if (count > 1) {
          merges.push({ s: { r: startRow, c: 0 }, e: { r: startRow + count - 1, c: 0 } });
          merges.push({ s: { r: startRow, c: 1 }, e: { r: startRow + count - 1, c: 1 } });
          merges.push({ s: { r: startRow, c: 5 }, e: { r: startRow + count - 1, c: 5 } });
        }
      }
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (merges.length > 0) {
      ws['!merges'] = merges;
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Personel_Egitim_Listesi');
    const out = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
    return out;
  };

  // Add a new Person
  const handleSaveNewPerson = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPersonName.trim()) {
      alert("Lütfen personelin adını ve soyadını giriniz.");
      return;
    }

    const nextSiraNo = people.length > 0 ? Math.max(...people.map(p => p.siraNo)) + 1 : 1;
    const initialTrainings: TrainingItem[] = [];

    if (newInitialTrainingName.trim()) {
      initialTrainings.push({
        id: `tr-${Date.now()}-1`,
        egitimAdi: newInitialTrainingName.trim().toUpperCase(),
        egitimTarihi: newInitialTrainingDate.trim() || new Date().toLocaleDateString('tr-TR'),
        egitimVeren: (newInitialTrainingProvider.trim() || 'OGM HAVACILIK').toUpperCase()
      });
    }

    const newPerson: PersonelEgitimPerson = {
      id: `person-${Date.now()}`,
      siraNo: nextSiraNo,
      adiSoyadi: newPersonName.trim().toUpperCase(),
      trainings: initialTrainings,
      hasPdf: false,
      linkTitle: `${newPersonName.trim().toUpperCase()} EĞİTİM SERTİFİKALARI`
    };

    const updated = [...people, newPerson];
    setPeople(updated);
    setIsAddPersonModalOpen(false);
    setNewPersonName('');
    setNewInitialTrainingName('');
    setNewInitialTrainingProvider('');

    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Yeni Personel Ekleniyor ve Excel'e Yazılıyor...");

    if (showNotification) {
      showNotification(`✅ "${newPerson.adiSoyadi}" listeye eklendi.`);
    }
  };

  // Inline "Eğitim Ekle" for a person:
  // "kişinin en son eğitimamdı kısmına eğitim ekl ebuton olur ona basınca otomatik olarak eğitimi ekler"
  const handleAddInlineTraining = (personId: string) => {
    if (!inlineTrainingName.trim()) {
      alert("Lütfen eğitim adını giriniz.");
      return;
    }

    const newTraining: TrainingItem = {
      id: `tr-${Date.now()}`,
      egitimAdi: inlineTrainingName.trim().toUpperCase(),
      egitimTarihi: inlineTrainingDate.trim() || new Date().toLocaleDateString('tr-TR'),
      egitimVeren: (inlineTrainingProvider.trim() || 'OGM HAVACILIK').toUpperCase()
    };

    const updated = people.map(p => {
      if (p.id === personId) {
        return {
          ...p,
          trainings: [...p.trainings, newTraining]
        };
      }
      return p;
    });

    setPeople(updated);
    setInlineAddingPersonId(null);
    setInlineTrainingName('');
    setInlineTrainingProvider('');

    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Yeni Eğitim Ekleniyor ve Excel'e Yazılıyor...");

    if (showNotification) {
      showNotification(`✅ Eğitim başarıyla eklendi.`);
    }
  };

  // Delete training from a person
  const handleDeleteTraining = (personId: string, trainingId: string) => {
    if (!confirm("Bu eğitimi listeden silmek istediğinize emin misiniz?")) return;
    const updated = people.map(p => {
      if (p.id === personId) {
        return {
          ...p,
          trainings: p.trainings.filter(t => t.id !== trainingId)
        };
      }
      return p;
    });
    setPeople(updated);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Eğitim Siliniyor ve Excel Güncelleniyor...");
  };

  // Delete entire person
  const handleDeletePerson = (personId: string) => {
    if (!confirm("Bu personeli ve tüm eğitimlerini silmek istediğinize emin misiniz?")) return;
    const updated = people.filter(p => p.id !== personId);
    setPeople(updated);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Personel Kaydı Siliniyor ve Excel Güncelleniyor...");
  };

  // Handle "Belge Ekle" click:
  // "LİNK SUTUNA BELGE EKLE BASINCA YENİ SEKMEDE 'https://bulut.ogm.gov.tr/BELGEYUKLE' BUNU AÇAR BASAR BASMAZ LİNK AÇILDIKTAN SONRA İLGİLİ PERSONELİN SATIRINDA SORU GELİR TİKLİ TİKLİ SORUDA 'PDF YÜKLEDİNİZ Mİ?' TİKLİ EVET BUTONUNA BASARSA O KİŞİNİN OLDUĞU YERDE BELGE GÖR BUTONU OLUR"
  const handleBelgeEkleClick = (person: PersonelEgitimPerson) => {
    window.open('https://bulut.ogm.gov.tr/BELGEYUKLE', '_blank');
    setPromptPersonId(person.id);
    setIsPdfConfirmedChecked(true);

    if (showNotification) {
      showNotification(`📄 Belge yükleme sayfası yeni sekmede açıldı. Lütfen yükleme onayını veriniz.`);
    }
  };

  // Confirm PDF upload -> switches to "Belge Gör"
  const handleConfirmPdfUpload = (personId: string) => {
    const updated = people.map(p => (p.id === personId ? { ...p, hasPdf: true } : p));
    setPeople(updated);
    setPromptPersonId(null);
    uploadExcelToGoogleDrive(buildExcelBase64FromPeople(updated), "Belge Onayı Excel'e İşleniyor...");

    if (showNotification) {
      showNotification('✅ Belge onaylandı. Artık "Belge Gör" ile açılabilir.');
    }
  };

  // Handle "Belge Gör" click:
  // "BU SEFER YÜKLE YERİNE BELGE GÖR BUTONU DA 'https://bulut.ogm.gov.tr/BELGEGOR' BU OLUR BU DA YENİ SEKMEDE AÇILIR."
  const handleBelgeGorClick = () => {
    window.open('https://bulut.ogm.gov.tr/BELGEGOR', '_blank');
  };

  // Filtered People
  const filteredPeople = useMemo(() => {
    if (!searchQuery.trim()) return people;
    const q = searchQuery.toLowerCase().trim();
    return people.filter(p =>
      String(p.siraNo).includes(q) ||
      p.adiSoyadi.toLowerCase().includes(q) ||
      p.trainings.some(t =>
        t.egitimAdi.toLowerCase().includes(q) ||
        t.egitimTarihi.toLowerCase().includes(q) ||
        t.egitimVeren.toLowerCase().includes(q)
      )
    );
  }, [people, searchQuery]);

  // Search match positions for highlighting
  const searchMatches = useMemo(() => {
    if (!searchQuery.trim()) return [];
    const q = searchQuery.toLowerCase().trim();
    const list: { personId: string; field: string }[] = [];
    filteredPeople.forEach(p => {
      if (p.adiSoyadi.toLowerCase().includes(q)) list.push({ personId: p.id, field: 'adiSoyadi' });
      p.trainings.forEach(t => {
        if (t.egitimAdi.toLowerCase().includes(q)) list.push({ personId: p.id, field: 'egitimAdi' });
        if (t.egitimVeren.toLowerCase().includes(q)) list.push({ personId: p.id, field: 'egitimVeren' });
      });
    });
    return list;
  }, [filteredPeople, searchQuery]);

  // EXCEL İNDİR:
  // "excel indirde 'belge durumu kısmna köprü yap eğer yükleme belge gör varsa köprü de belge gör olsun bağlantı o line açsın"
  const handleDownloadExcel = () => {
    try {
      if (filteredPeople.length === 0) {
        alert("İndirilecek personel eğitimi verisi bulunamadı.");
        return;
      }

      // Generate HTML-based Excel with XML hyperlink support, light-blue ADI SOYADI background (#dbeafe), and BELGE DURUMU köprü
      let tableRowsHtml = '';

      filteredPeople.forEach(p => {
        const rowCount = Math.max(1, p.trainings.length);
        const linkHref = p.hasPdf ? 'https://bulut.ogm.gov.tr/BELGEGOR' : 'https://bulut.ogm.gov.tr/BELGEYUKLE';
        const linkText = p.hasPdf ? `${p.adiSoyadi} EĞİTİM SERTİFİKALARI (BELGE GÖR)` : 'BELGE YÜKLE';
        const linkColor = p.hasPdf ? '#1d4ed8' : '#b45309';

        if (p.trainings.length === 0) {
          tableRowsHtml += `
            <tr>
              <td style="text-align: center; border: 1px solid #cbd5e1; font-weight: bold; background-color: #f1f5f9;">${p.siraNo}</td>
              <td style="border: 1px solid #cbd5e1; font-weight: bold; background-color: #dbeafe; padding: 6px;">${p.adiSoyadi}</td>
              <td style="border: 1px solid #cbd5e1; padding: 6px;">-</td>
              <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">-</td>
              <td style="border: 1px solid #cbd5e1; padding: 6px;">-</td>
              <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">
                <a href="${linkHref}" style="color: ${linkColor}; text-decoration: underline; font-weight: bold;">${linkText}</a>
              </td>
            </tr>
          `;
        } else {
          p.trainings.forEach((t, idx) => {
            if (idx === 0) {
              tableRowsHtml += `
                <tr>
                  <td rowspan="${rowCount}" style="text-align: center; vertical-align: middle; border: 1px solid #cbd5e1; font-weight: bold; background-color: #f1f5f9;">${p.siraNo}</td>
                  <td rowspan="${rowCount}" style="vertical-align: middle; border: 1px solid #cbd5e1; font-weight: bold; background-color: #dbeafe; padding: 8px;">${p.adiSoyadi}</td>
                  <td style="border: 1px solid #cbd5e1; padding: 6px; font-weight: 500;">${t.egitimAdi}</td>
                  <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">${t.egitimTarihi}</td>
                  <td style="border: 1px solid #cbd5e1; padding: 6px;">${t.egitimVeren}</td>
                  <td rowspan="${rowCount}" style="vertical-align: middle; text-align: center; border: 1px solid #cbd5e1; padding: 8px;">
                    <a href="${linkHref}" style="color: ${linkColor}; text-decoration: underline; font-weight: bold;">${linkText}</a>
                  </td>
                </tr>
              `;
            } else {
              tableRowsHtml += `
                <tr>
                  <td style="border: 1px solid #cbd5e1; padding: 6px; font-weight: 500;">${t.egitimAdi}</td>
                  <td style="border: 1px solid #cbd5e1; text-align: center; padding: 6px;">${t.egitimTarihi}</td>
                  <td style="border: 1px solid #cbd5e1; padding: 6px;">${t.egitimVeren}</td>
                </tr>
              `;
            }
          });
        }
      });

      const html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <!--[if gte mso 9]>
          <xml>
            <x:ExcelWorkbook>
              <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                  <x:Name>Personel Egitim Listesi</x:Name>
                  <x:WorksheetOptions>
                    <x:DisplayGridlines/>
                  </x:WorksheetOptions>
                </x:ExcelWorksheet>
              </x:ExcelWorksheets>
            </x:ExcelWorkbook>
          </xml>
          <![endif]-->
          <style>
            table { border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
            th { background-color: #0b3d1d; color: #ffffff; font-weight: bold; padding: 10px; border: 1px solid #000000; text-align: left; }
            td { padding: 6px; border: 1px solid #cbd5e1; vertical-align: middle; }
          </style>
        </head>
        <body>
          <h2 style="color: #0b3d1d; font-family: Calibri, Arial, sans-serif; margin-bottom: 8px;">
            ORMAN GENEL MÜDÜRLÜĞÜ - PERSONEL EĞİTİM LİSTESİ
          </h2>
          <table border="1">
            <thead>
              <tr style="background-color: #0b3d1d; color: #ffffff;">
                <th style="text-align: center; width: 60px;">SIRA NO</th>
                <th style="width: 220px;">ADI SOYADI</th>
                <th style="width: 380px;">EĞİTİMİN ADI</th>
                <th style="text-align: center; width: 160px;">EĞİTİM TARİHİ</th>
                <th style="width: 300px;">EĞİTİM VEREN KURUM/PERSONEL</th>
                <th style="text-align: center; width: 260px;">LİNK</th>
              </tr>
            </thead>
            <tbody>
              ${tableRowsHtml}
            </tbody>
          </table>
        </body>
        </html>
      `;

      const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `personel_egitim_listesi_${new Date().toISOString().slice(0, 10)}.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      if (showNotification) {
        showNotification('✅ Personel Eğitim Listesi köprüleri ile birlikte Excel formatında indirildi.');
      }
    } catch (e: any) {
      alert("Excel indirme hatası: " + e.message);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[600] bg-slate-950 flex flex-col overflow-hidden text-white animate-fade-in select-none">
      {/* 1. TOP HEADER - Styled exactly like Personel Bilgi Çizelgesi */}
      <div className="shrink-0 bg-slate-900 border-b border-slate-800 px-6 py-4 flex flex-col lg:flex-row items-center justify-between gap-4 shadow-md select-none">
        
        {/* Brand & Badge */}
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-2xl bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400 shadow-inner">
            <GraduationCap className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
              <h2 className="text-sm font-black text-white tracking-widest uppercase font-mono">
                PERSONEL EĞİTİM LİSTESİ
              </h2>
            </div>
            
          </div>
        </div>

        {/* Center: Search Box */}
        <div className="flex items-center gap-3 w-full max-w-xl">
          <div className="flex items-center gap-3 bg-slate-950 px-4 py-2.5 rounded-2xl border border-slate-800 shadow-inner w-full">
            <Search className="w-4 h-4 text-emerald-400 shrink-0" />
            <input
              type="text"
              placeholder="PERSONEL ADI, EĞİTİM ADI VEYA KURUM ARA..."
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setActiveMatchIdx(0);
              }}
              className="bg-transparent text-slate-100 text-xs font-extrabold outline-none border-none placeholder-slate-600 w-full uppercase"
            />
            {searchQuery && (
              <button
                onClick={() => {
                  setSearchQuery('');
                  setActiveMatchIdx(0);
                }}
                className="text-slate-500 hover:text-slate-300 transition-colors cursor-pointer"
                title="Aramayı Temizle"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          {/* Search Match Nav */}
          {searchQuery && searchMatches.length > 0 && (
            <div className="flex items-center gap-1.5 bg-emerald-950/70 border border-emerald-700/60 px-3 py-1.5 rounded-xl font-mono text-[10px] text-emerald-300 shrink-0">
              <button
                type="button"
                onClick={() => setActiveMatchIdx(prev => (prev > 0 ? prev - 1 : searchMatches.length - 1))}
                className="hover:text-white cursor-pointer px-1 font-bold"
              >
                ◀
              </button>
              <span className="font-extrabold tracking-widest">
                {activeMatchIdx + 1}/{searchMatches.length}
              </span>
              <button
                type="button"
                onClick={() => setActiveMatchIdx(prev => (prev < searchMatches.length - 1 ? prev + 1 : 0))}
                className="hover:text-white cursor-pointer px-1 font-bold"
              >
                ▶
              </button>
            </div>
          )}
        </div>

        {/* Right Actions: VERİ EKLE (EXCEL YÜKLE) + YENİ PERSONEL + EXCEL İNDİR */}
        <div className="flex items-center gap-2.5 shrink-0 flex-wrap">
          
          {/* Hidden File Input for Excel Upload */}
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls"
            className="hidden"
          />

          {/* VERİ EKLE (EXCEL YÜKLE) BUTTON */}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-black font-mono text-xs rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-md hover:scale-[1.02] active:scale-95"
            title="Google Drive'a ve sisteme Excel dosyası yükle"
          >
            <Upload className="w-4 h-4 text-blue-200" />
            <span>VERİ EKLE (EXCEL YÜKLE)</span>
          </button>

          {/* YENİ PERSONEL EKLE */}
          <button
            onClick={() => setIsAddPersonModalOpen(true)}
            className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black font-mono text-xs rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-md hover:scale-[1.02] active:scale-95"
            title="Sıfırdan Yeni Personel ve Eğitimi Ekle"
          >
            <Plus className="w-4 h-4" />
            <span>YENİ PERSONEL EKLE</span>
          </button>

          {/* DRIVE'DAN YENİLE */}
          <button
            onClick={() => fetchLiveExcelFromDrive(true)}
            disabled={isLoadingDrive}
            className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-xl transition-all border border-slate-700 cursor-pointer"
            title="Google Drive'dan verileri canlı tazele"
          >
            <RefreshCw className={`w-4 h-4 text-emerald-400 ${isLoadingDrive ? 'animate-spin' : ''}`} />
          </button>

          {/* EXCEL İNDİR (Köprülü) */}
          <button
            onClick={handleDownloadExcel}
            className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white font-extrabold font-mono text-xs rounded-xl flex items-center gap-2 transition-all cursor-pointer border border-slate-700 shadow-sm active:scale-95"
            title="Mevcut Tabloyu Köprüler ile Excel Olarak İndir"
          >
            <Download className="w-4 h-4 text-emerald-400" />
            <span>EXCEL İNDİR</span>
          </button>

          {/* Kapat Butonu */}
          <button
            onClick={onClose}
            className="p-2 bg-slate-800 hover:bg-red-950 hover:text-red-400 text-slate-400 rounded-xl transition-all border border-slate-700 cursor-pointer ml-1"
            title="Kapat"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

      </div>

      {/* 2. TABLE BODY CONTAINER - Grouped Multi-row layout matching image_802de070.png */}
      <div className="flex-1 overflow-auto p-4 md:p-6 bg-slate-950 flex justify-center scrollbar-thin">
        <div className="w-full bg-white text-slate-800 rounded-2xl shadow-2xl p-6 flex flex-col border border-slate-200 overflow-x-auto">
          
          {/* Info Banner */}
          <div className="mb-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 bg-emerald-50/80 border border-emerald-200/80 px-4 py-2.5 rounded-xl text-xs text-emerald-950">
            <div className="flex items-center gap-2">
              <BookOpen className="w-4 h-4 text-emerald-700 shrink-0" />
              <span className="font-semibold">
                Her personelin aldığı tüm eğitimler liste halinde toplanmıştır. Birden fazla eğitimi olan personeller için &quot;Diğer Eğitimleri Gör&quot; butonuna basarak tüm eğitim geçmişini aşağı doğru açabilirsiniz.
              </span>
            </div>
            <div className="flex items-center gap-3 font-mono text-[11px] text-emerald-800 shrink-0">
              <div className="flex items-center gap-1.5">
                <span className="font-bold">Toplam Personel:</span>
                <span className="bg-emerald-200/80 px-2 py-0.5 rounded-full font-black text-emerald-950">{people.length}</span>
              </div>
              <div className="flex items-center gap-1 bg-white/80 p-0.5 rounded-lg border border-emerald-200 shadow-xs">
                <button
                  type="button"
                  onClick={expandAll}
                  className="px-2 py-1 text-[10px] font-bold text-emerald-900 hover:bg-emerald-100 rounded cursor-pointer transition-colors"
                  title="Tüm personellerin eğitimlerini aşağı doğru aç"
                >
                  Tümünü Aç
                </button>
                <span className="text-emerald-300">|</span>
                <button
                  type="button"
                  onClick={collapseAll}
                  className="px-2 py-1 text-[10px] font-bold text-slate-600 hover:bg-slate-100 rounded cursor-pointer transition-colors"
                  title="Tüm açık eğitim pencerelerini kapat"
                >
                  Tümünü Kapat
                </button>
              </div>
            </div>
          </div>

          {/* TABLE CONTAINER WITH FLU OVERLAY */}
          <div className="relative flex-1 overflow-x-auto w-full rounded-xl border border-slate-300 min-h-[420px]">
            {/* FLU / LOADING OVERLAY */}
            {isLoadingDrive && (
              <div className="absolute inset-0 z-40 bg-white/80 backdrop-blur-md flex flex-col items-center justify-center gap-3 transition-all duration-300">
                <div className="relative flex items-center justify-center">
                  <div className="w-14 h-14 rounded-full border-4 border-emerald-200 border-t-emerald-600 animate-spin" />
                  <FileSpreadsheet className="w-6 h-6 text-emerald-700 absolute" />
                </div>
                <div className="text-center space-y-1">
                  <div className="text-sm font-black text-slate-900 tracking-wide font-mono">
                    {syncStatusMessage || "Veriler Yükleniyor..."}
                  </div>
                  <p className="text-xs font-semibold text-slate-500">
                    Excel dosyası işleniyor ve senkronize ediliyor, lütfen bekleyiniz...
                  </p>
                </div>
              </div>
            )}

            <table className={`table-auto w-full border-collapse text-left text-xs text-slate-800 transition duration-300 ${isLoadingDrive ? 'filter blur-xs pointer-events-none select-none opacity-50' : ''}`}>
              <thead>
                <tr className="bg-slate-900 text-white font-extrabold uppercase tracking-wider text-[11px] select-none">
                  <th className="p-3 border-b-4 border-r border-slate-600 text-center whitespace-nowrap w-16">
                    SIRA NO
                  </th>
                  <th className="p-3 border-b-4 border-r border-slate-600 text-left whitespace-nowrap min-w-[200px]">
                    ADI SOYADI
                  </th>
                  <th className="p-3 border-b-4 border-r border-slate-600 text-left whitespace-nowrap min-w-[340px]">
                    EĞİTİMİN ADI
                  </th>
                  <th className="p-3 border-b-4 border-r border-slate-600 text-center whitespace-nowrap min-w-[140px]">
                    EĞİTİM TARİHİ
                  </th>
                  <th className="p-3 border-b-4 border-r border-slate-600 text-left whitespace-nowrap min-w-[280px]">
                    EĞİTİM VEREN KURUM/PERSONEL
                  </th>
                  <th className="p-3 border-b-4 border-slate-600 text-center whitespace-nowrap min-w-[260px]">
                    LİNK
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300">
                {!isLoadingDrive && filteredPeople.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="text-center py-20 text-slate-400 font-mono text-xs uppercase tracking-wider">
                      <div className="flex flex-col items-center justify-center gap-3">
                        <FileSpreadsheet className="w-12 h-12 text-slate-300 animate-pulse" />
                        <span className="font-bold text-slate-600">Henüz personel eğitimi verisi bulunmamaktadır.</span>
                        <div className="flex items-center gap-3 mt-2">
                          <button
                            type="button"
                            onClick={() => fileInputRef.current?.click()}
                            className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow cursor-pointer active:scale-95"
                          >
                            <Upload className="w-4 h-4" />
                            <span>VERİ EKLE (EXCEL YÜKLE)</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => setIsAddPersonModalOpen(true)}
                            className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow cursor-pointer active:scale-95"
                          >
                            <Plus className="w-4 h-4" />
                            <span>YENİ PERSONEL EKLE</span>
                          </button>
                        </div>
                      </div>
                    </td>
                  </tr>
                ) : (
                  filteredPeople.map((person) => {
                    const isPromptActive = promptPersonId === person.id;
                    const isInlineAdding = inlineAddingPersonId === person.id;
                    const trainingsCount = person.trainings.length;
                    const isExpanded = expandedPersonIds.has(person.id);
                    const primaryTraining = trainingsCount > 0 ? person.trainings[0] : null;

                    return (
                      <React.Fragment key={person.id}>
                        {/* MAIN PERSON ROW */}
                        <tr className="border-b-2 border-slate-300 hover:bg-slate-50/90 transition-colors">
                          {/* 1. SIRA NO with '+' button */}
                          <td className="p-3 border-r border-slate-300 font-mono font-bold text-center text-slate-700 bg-slate-100/70 align-middle w-16">
                            <div className="flex flex-col items-center justify-center gap-1.5">
                              <span className="text-slate-900 font-extrabold text-sm">{person.siraNo}</span>
                              <button
                                type="button"
                                onClick={() => {
                                  setInlineAddingPersonId(person.id);
                                  setInlineTrainingName('');
                                  setInlineTrainingDate(new Date().toLocaleDateString('tr-TR'));
                                  setInlineTrainingProvider('');
                                }}
                                className="w-6 h-6 rounded-full bg-emerald-600 hover:bg-emerald-500 text-white flex items-center justify-center cursor-pointer shadow-xs transition-all active:scale-90"
                                title="Bu personele yeni eğitim ekle (+)"
                              >
                                <Plus className="w-3.5 h-3.5 stroke-[3]" />
                              </button>
                            </div>
                          </td>

                          {/* 2. ADI SOYADI */}
                          <td className="p-3.5 border-r border-slate-300 font-bold text-slate-900 bg-[#dbeafe] align-middle min-w-[200px]">
                            <div className="flex flex-col gap-1">
                              <span className="font-extrabold text-sm tracking-wide text-slate-900">
                                {person.adiSoyadi}
                              </span>
                              <div className="flex items-center justify-between gap-1 mt-1 text-[10px]">
                                <span className="font-mono text-slate-600 bg-blue-100/80 px-2 py-0.5 rounded-md font-semibold">
                                  {trainingsCount} Eğitim
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleDeletePerson(person.id)}
                                  className="text-red-600 hover:text-red-800 hover:underline flex items-center gap-0.5 cursor-pointer"
                                  title="Personeli Sil"
                                >
                                  <Trash2 className="w-3 h-3" />
                                  <span>Sil</span>
                                </button>
                              </div>
                            </div>
                          </td>

                          {/* 3. EĞİTİMİN ADI + DİĞER EĞİTİMLERİ GÖR AŞAĞI AÇILIR BUTON */}
                          <td className="p-3 border-r border-slate-200 font-semibold text-slate-800 align-middle">
                            {primaryTraining ? (
                              <div className="flex flex-col gap-1.5">
                                <div className="flex items-center justify-between gap-2">
                                  <span className="font-bold text-slate-900 text-xs">{primaryTraining.egitimAdi}</span>
                                  <button
                                    type="button"
                                    onClick={() => handleDeleteTraining(person.id, primaryTraining.id)}
                                    className="text-slate-300 hover:text-red-500 opacity-80 hover:opacity-100 transition-opacity p-0.5 cursor-pointer shrink-0"
                                    title="Bu Eğitimi Sil"
                                  >
                                    <X className="w-3.5 h-3.5" />
                                  </button>
                                </div>

                                {trainingsCount > 1 && (
                                  <button
                                    type="button"
                                    onClick={() => toggleExpandPerson(person.id)}
                                    className="self-start inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 font-mono text-[11px] font-bold transition-all cursor-pointer shadow-xs active:scale-95"
                                  >
                                    {isExpanded ? (
                                      <>
                                        <ChevronUp className="w-3.5 h-3.5 text-emerald-600" />
                                        <span>Diğer Eğitimleri Gizle</span>
                                      </>
                                    ) : (
                                      <>
                                        <ChevronDown className="w-3.5 h-3.5 text-emerald-600" />
                                        <span>Diğer Eğitimleri Gör ({trainingsCount - 1} Eğitim Daha)</span>
                                      </>
                                    )}
                                  </button>
                                )}
                              </div>
                            ) : (
                              <span className="text-slate-400 italic text-xs">Henüz eğitim kaydı bulunmuyor</span>
                            )}
                          </td>

                          {/* 4. EĞİTİM TARİHİ */}
                          <td className="p-3 border-r border-slate-200 text-center font-mono font-bold text-slate-700 whitespace-nowrap align-middle min-w-[140px]">
                            {primaryTraining ? (
                              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-slate-100 text-slate-800 border border-slate-200 text-[11px]">
                                <Calendar className="w-3 h-3 text-slate-400" />
                                {primaryTraining.egitimTarihi}
                              </span>
                            ) : (
                              <span>-</span>
                            )}
                          </td>

                          {/* 5. EĞİTİM VEREN KURUM/PERSONEL */}
                          <td className="p-3 border-r border-slate-200 font-medium text-slate-700 whitespace-nowrap align-middle min-w-[280px]">
                            {primaryTraining ? (
                              <div className="flex items-center gap-1.5 text-xs">
                                <Building2 className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                                <span>{primaryTraining.egitimVeren}</span>
                              </div>
                            ) : (
                              <span>-</span>
                            )}
                          </td>

                          {/* 6. LİNK */}
                          <td className="p-3 text-center align-middle whitespace-nowrap bg-white min-w-[240px]">
                            {isPromptActive ? (
                              <div className="inline-flex flex-col items-center gap-2 bg-amber-50 border-2 border-amber-400/90 p-3 rounded-2xl shadow-lg text-amber-950 max-w-[240px]">
                                <label className="flex items-center gap-1.5 cursor-pointer text-[11px] font-bold">
                                  <input
                                    type="checkbox"
                                    checked={isPdfConfirmedChecked}
                                    onChange={(e) => setIsPdfConfirmedChecked(e.target.checked)}
                                    className="w-4 h-4 text-emerald-600 rounded border-amber-300 focus:ring-emerald-500 cursor-pointer"
                                  />
                                  <span className="font-black text-amber-900">PDF Yüklediniz mi?</span>
                                </label>
                                <div className="flex items-center gap-2">
                                  <button
                                    type="button"
                                    disabled={!isPdfConfirmedChecked}
                                    onClick={() => handleConfirmPdfUpload(person.id)}
                                    className={`px-3 py-1.5 text-white font-black text-[11px] rounded-lg transition-all flex items-center gap-1 shadow cursor-pointer ${
                                      isPdfConfirmedChecked
                                        ? 'bg-emerald-600 hover:bg-emerald-500 active:scale-95'
                                        : 'bg-slate-300 cursor-not-allowed opacity-60'
                                    }`}
                                  >
                                    <CheckCircle2 className="w-3.5 h-3.5" />
                                    <span>EVET</span>
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setPromptPersonId(null)}
                                    className="p-1.5 text-slate-400 hover:text-slate-600 rounded-lg hover:bg-amber-100 transition-colors cursor-pointer"
                                  >
                                    <X className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </div>
                            ) : person.hasPdf ? (
                              <div className="inline-flex flex-col items-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={handleBelgeGorClick}
                                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black text-xs rounded-xl flex items-center gap-2 transition-all shadow-md hover:shadow-lg cursor-pointer active:scale-95 group"
                                >
                                  <FileCheck className="w-4 h-4 text-emerald-200 group-hover:scale-110 transition-transform" />
                                  <span>BELGE GÖR</span>
                                  <ExternalLink className="w-3.5 h-3.5 text-emerald-300 opacity-90" />
                                </button>
                                <span className="text-[10px] text-blue-600 font-semibold underline cursor-pointer hover:text-blue-800 transition-colors" onClick={handleBelgeGorClick}>
                                  {person.linkTitle && !person.linkTitle.includes('http') && !person.linkTitle.includes('BELGE') ? person.linkTitle : `${person.adiSoyadi} EĞİTİM SERTİFİKALARI`}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleBelgeEkleClick(person)}
                                  className="text-[10px] text-slate-400 hover:text-slate-700 flex items-center gap-1 mt-1 cursor-pointer"
                                >
                                  <RefreshCw className="w-2.5 h-2.5" />
                                  <span>Belgeyi Değiştir</span>
                                </button>
                              </div>
                            ) : (
                              <div className="inline-flex flex-col items-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => handleBelgeEkleClick(person)}
                                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-black text-xs rounded-xl flex items-center gap-2 transition-all shadow-md hover:shadow-lg cursor-pointer active:scale-95 group"
                                >
                                  <FilePlus className="w-4 h-4 text-blue-200 group-hover:scale-110 transition-transform" />
                                  <span>BELGE EKLE</span>
                                  <ExternalLink className="w-3.5 h-3.5 text-blue-300 opacity-90" />
                                </button>
                                <span className="text-[9px] text-amber-700 font-medium">Belge Bekleniyor</span>
                              </div>
                            )}
                          </td>
                        </tr>

                        {/* EXPANDABLE ACCORDION ROW FOR OTHER TRAININGS ("aşağı doğru açma yap") */}
                        {isExpanded && trainingsCount > 1 && (
                          <tr className="bg-slate-50/95 border-b-4 border-slate-800">
                            <td colSpan={6} className="p-4 pl-12 pr-6">
                              <div className="bg-white rounded-2xl border-2 border-emerald-500/40 shadow-lg p-4">
                                <div className="flex items-center justify-between pb-3 mb-3 border-b border-slate-200">
                                  <div className="flex items-center gap-2.5">
                                    <div className="w-8 h-8 rounded-xl bg-emerald-50 border border-emerald-300 flex items-center justify-center text-emerald-700">
                                      <GraduationCap className="w-4 h-4" />
                                    </div>
                                    <div>
                                      <h4 className="text-xs font-black uppercase text-slate-900 tracking-wider font-mono">
                                        {person.adiSoyadi} • TÜM EĞİTİM GEÇMİŞİ
                                      </h4>
                                      <span className="text-[10px] text-slate-500 font-mono">
                                        Bu personele ait toplam {trainingsCount} adet kayıtlı eğitim bulunmaktadır
                                      </span>
                                    </div>
                                  </div>

                                  <div className="flex items-center gap-2">
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setInlineAddingPersonId(person.id);
                                        setInlineTrainingName('');
                                        setInlineTrainingDate(new Date().toLocaleDateString('tr-TR'));
                                        setInlineTrainingProvider('');
                                      }}
                                      className="px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-black font-mono text-[11px] flex items-center gap-1.5 shadow-sm cursor-pointer active:scale-95"
                                    >
                                      <Plus className="w-3.5 h-3.5 stroke-[3]" />
                                      <span>Yeni Eğitim Ekle</span>
                                    </button>

                                    <button
                                      type="button"
                                      onClick={() => toggleExpandPerson(person.id)}
                                      className="p-1.5 text-slate-400 hover:text-slate-700 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
                                      title="Kapat"
                                    >
                                      <ChevronUp className="w-4 h-4" />
                                    </button>
                                  </div>
                                </div>

                                {/* Mini Table */}
                                <div className="overflow-x-auto rounded-xl border border-slate-200">
                                  <table className="w-full text-xs text-left">
                                    <thead>
                                      <tr className="bg-slate-100 text-slate-700 font-extrabold uppercase text-[10px] tracking-wider border-b border-slate-200">
                                        <th className="p-2.5 text-center w-12 font-mono">#</th>
                                        <th className="p-2.5 text-left">EĞİTİMİN ADI</th>
                                        <th className="p-2.5 text-center w-40">EĞİTİM TARİHİ</th>
                                        <th className="p-2.5 text-left w-72">EĞİTİM VEREN KURUM / PERSONEL</th>
                                        <th className="p-2.5 text-center w-16">İŞLEM</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100">
                                      {person.trainings.map((t, idx) => (
                                        <tr key={t.id} className="hover:bg-emerald-50/40 transition-colors">
                                          <td className="p-2.5 text-center font-mono font-bold text-slate-500 bg-slate-50/50">
                                            {idx + 1}
                                          </td>
                                          <td className="p-2.5 font-bold text-slate-900">
                                            {t.egitimAdi}
                                          </td>
                                          <td className="p-2.5 text-center font-mono font-bold text-slate-700">
                                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-slate-100 text-slate-800 border border-slate-200 text-[10px]">
                                              <Calendar className="w-3 h-3 text-slate-400" />
                                              {t.egitimTarihi}
                                            </span>
                                          </td>
                                          <td className="p-2.5 text-slate-700 font-medium">
                                            <div className="flex items-center gap-1.5">
                                              <Building2 className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                                              <span>{t.egitimVeren}</span>
                                            </div>
                                          </td>
                                          <td className="p-2.5 text-center">
                                            <button
                                              type="button"
                                              onClick={() => handleDeleteTraining(person.id, t.id)}
                                              className="p-1 text-slate-400 hover:text-red-600 rounded-md hover:bg-red-50 transition cursor-pointer"
                                              title="Bu eğitimi sil"
                                            >
                                              <Trash2 className="w-3.5 h-3.5" />
                                            </button>
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}

                        {/* INLINE TRAINING INPUT ROW */}
                        {isInlineAdding && (
                          <tr className="border-b-2 border-emerald-400 bg-emerald-50/80">
                            <td className="p-3 text-center font-mono font-bold text-emerald-800 bg-emerald-100/50">
                              +
                            </td>
                            <td className="p-3 font-bold text-emerald-950 font-mono text-xs">
                              {person.adiSoyadi} (Yeni Eğitim)
                            </td>
                            <td className="p-3 border-r border-slate-200">
                              <datalist id="training-list">
                                {Array.from(new Set(people.flatMap(p => p.trainings.map(t => t.egitimAdi)))).map((t, i) => (
                                  <option key={i} value={t} />
                                ))}
                              </datalist>
                              <input
                                type="text"
                                autoFocus
                                list="training-list"
                                placeholder="Eğitimin Adını Yazınız..."
                                value={inlineTrainingName}
                                onChange={(e) => setInlineTrainingName(e.target.value)}
                                className="w-full px-3 py-2 bg-white border-2 border-emerald-500 rounded-xl text-xs font-bold uppercase text-slate-900 outline-none shadow-xs"
                              />
                            </td>
                            <td className="p-3 border-r border-slate-200 text-center">
                              <input
                                type="text"
                                placeholder="GG.AA.YYYY"
                                value={inlineTrainingDate}
                                onChange={(e) => setInlineTrainingDate(e.target.value)}
                                className="w-32 px-3 py-2 bg-white border-2 border-emerald-500 rounded-xl text-xs font-mono font-bold text-slate-900 text-center outline-none shadow-xs"
                              />
                            </td>
                            <td className="p-3 border-r border-slate-200">
                              <div className="flex items-center gap-2">
                                <input
                                  type="text"
                                  placeholder="Veren Kurum / Personel"
                                  value={inlineTrainingProvider}
                                  onChange={(e) => setInlineTrainingProvider(e.target.value)}
                                  className="flex-1 px-3 py-2 bg-white border-2 border-emerald-500 rounded-xl text-xs font-bold uppercase text-slate-900 outline-none shadow-xs"
                                />
                                <button
                                  type="button"
                                  onClick={() => handleAddInlineTraining(person.id)}
                                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black text-xs rounded-xl flex items-center gap-1 shadow cursor-pointer active:scale-95"
                                >
                                  <Check className="w-3.5 h-3.5 stroke-[3]" />
                                  <span>Ekle</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setInlineAddingPersonId(null)}
                                  className="p-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl cursor-pointer"
                                  title="Vazgeç"
                                >
                                  <X className="w-4 h-4" />
                                </button>
                              </div>
                            </td>
                            <td className="p-3 text-center bg-white">
                              <span className="text-[10px] text-slate-400 font-mono">Kaydedildiğinde listeye eklenir</span>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Table Footer */}
          <div className="mt-4 pt-3 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between text-[11px] text-slate-500 font-mono">
            <span>
              Resmi Havacılık Personel Eğitim Portalı • Son Senkronizasyon: {new Date().toLocaleDateString('tr-TR')}
            </span>
            <span className="font-bold text-emerald-800">
              OGM Yangınla Mücadele Havacılık Teknik Şube
            </span>
          </div>

        </div>
      </div>

      {/* 3. YENİ PERSONEL EKLEME MODALI */}
      {isAddPersonModalOpen && (
        <div className="fixed inset-0 z-[700] bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white text-slate-900 rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden animate-scale-up">
            
            {/* Header */}
            <div className="bg-slate-900 text-white px-6 py-4 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <GraduationCap className="w-5 h-5 text-emerald-400" />
                <h3 className="font-black text-sm uppercase tracking-wider font-mono">
                  Yeni Personel ve Eğitim Kaydı
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setIsAddPersonModalOpen(false)}
                className="text-slate-400 hover:text-white p-1 rounded-lg"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Form */}
            <form onSubmit={handleSaveNewPerson} className="p-6 space-y-4 text-xs">
              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                  Personel Adı Soyadı *
                </label>
                <div className="relative">
                  <input
                    type="text"
                    required
                    list="personnel-suggestions"
                    placeholder="Örn: MEHMET SEFA YÜCEDAĞ"
                    value={newPersonName}
                    onChange={(e) => setNewPersonName(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none focus:ring-2 focus:ring-emerald-500/30"
                  />
                  <datalist id="personnel-suggestions">
                    {DEFAULT_TEKNISYEN_PERSONEL_NAMES.map((name, i) => (
                      <option key={i} value={name} />
                    ))}
                  </datalist>
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                  İlk Eğitimin Adı (Opsiyonel)
                </label>
                <input
                  type="text"
                  list="training-list"
                  placeholder="Örn: BELL 429 (PWC PW207D1/D2) EASA B1.3"
                  value={newInitialTrainingName}
                  onChange={(e) => setNewInitialTrainingName(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none focus:ring-2 focus:ring-emerald-500/30"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    Eğitim Tarihi
                  </label>
                  <input
                    type="text"
                    placeholder="GG.AA.YYYY"
                    value={newInitialTrainingDate}
                    onChange={(e) => setNewInitialTrainingDate(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold text-slate-800 outline-none focus:ring-2 focus:ring-emerald-500/30 font-mono"
                  />
                </div>

                <div>
                  <label className="block font-bold text-slate-700 uppercase font-mono mb-1">
                    Eğitim Veren Kurum / Personel
                  </label>
                  <input
                    type="text"
                    placeholder="Örn: THE BELL TRAINING ACADEMY"
                    value={newInitialTrainingProvider}
                    onChange={(e) => setNewInitialTrainingProvider(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold uppercase text-slate-800 outline-none focus:ring-2 focus:ring-emerald-500/30"
                  />
                </div>
              </div>

              {/* Action Buttons */}
              <div className="pt-4 border-t border-slate-100 flex items-center justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setIsAddPersonModalOpen(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl"
                >
                  Vazgeç
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-xl shadow cursor-pointer active:scale-95"
                >
                  Kaydet
                </button>
              </div>
            </form>

          </div>
        </div>
      )}

    </div>
  );
};
