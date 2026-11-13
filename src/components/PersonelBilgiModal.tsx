import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Users,
  Search,
  X,
  Download,
  Plus,
  RefreshCw,
  Trash2,
  Edit2,
  Check,
  ShieldCheck,
  Lock,
  Calendar,
  FileSpreadsheet,
  FileText,
  Mail,
  Clock,
  AlertTriangle
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import { cleanAndFormatDateString } from '../utils/driveExcelSync';
import { fetchExcelFromDrive, uploadExcelToDrive, DRIVE_FOLDER_ID } from '../utils/remoteDriveSync';

export interface PersonelBilgiRecord {
  id: string;
  siraNo: number;
  adiSoyadi: string;
  tcKimlik: string;
  sicilNo: string;
  kadroUnvani: string;
  dogumTarihi: string;
  gorevYeri: string;
  telefonNo: string;
  kanGrubu: string;
  adresBilgisi: string;
  yakininAdiSoyadi: string;
  esTelefonNumaralari: string;
  periyodikMuayeneTarihi?: string;
}


const DRIVE_FILE_NAME = 'personel_bilgi_cizelgesi.xlsx';
const REQUIRED_PIN = '1839';

// Date parser helper to JS Date
function parseTrDate(dateStr?: string): Date | null {
  if (!dateStr || dateStr === '-' || dateStr.trim() === '') return null;
  const parts = dateStr.trim().split(/[\.\/-]/);
  if (parts.length === 3) {
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    let y = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
      return new Date(y, m, d);
    }
  }
  return null;
}

// Expiry calculator for Periyodik Muayene (5 Years)
export function getExpiryStatus(dateStr?: string, yearsValid: number = 5) {
  if (!dateStr || dateStr === '-' || dateStr.trim() === '') {
    return { hasDate: false, isExpired: false, isWarning: false, daysLeft: 9999, expiryDateStr: '-' };
  }
  const dateObj = parseTrDate(dateStr);
  if (!dateObj) {
    return { hasDate: false, isExpired: false, isWarning: false, daysLeft: 9999, expiryDateStr: '-' };
  }

  const expiryDate = new Date(dateObj);
  expiryDate.setFullYear(expiryDate.getFullYear() + yearsValid);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const diffTime = expiryDate.getTime() - today.getTime();
  const daysLeft = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  const expiryDateStr = `${String(expiryDate.getDate()).padStart(2, '0')}.${String(expiryDate.getMonth() + 1).padStart(2, '0')}.${expiryDate.getFullYear()}`;

  const isExpired = daysLeft <= 0;
  const isWarning = daysLeft > 0 && daysLeft <= 30;

  return {
    hasDate: true,
    isExpired,
    isWarning,
    daysLeft,
    expiryDateStr
  };
}

export function getNextMonthFirstDayStr(): string {
  const today = new Date();
  let nextMonth = today.getMonth() + 1;
  let year = today.getFullYear();
  if (nextMonth > 11) {
    nextMonth = 0;
    year += 1;
  }
  const firstDay = new Date(year, nextMonth, 1);
  return `${String(firstDay.getDate()).padStart(2, '0')}.${String(firstDay.getMonth() + 1).padStart(2, '0')}.${firstDay.getFullYear()}`;
}

interface PersonelBilgiModalProps {
  isOpen: boolean;
  onClose: () => void;
  showNotification?: (message: string) => void;
}

export const PersonelBilgiModal: React.FC<PersonelBilgiModalProps> = ({
  isOpen,
  onClose,
  showNotification
}) => {
  const [people, setPeople] = useState<PersonelBilgiRecord[]>([]);

  const [isLoadingDrive, setIsLoadingDrive] = useState(false);
  const [syncStatusMessage, setSyncStatusMessage] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedKadroFilter, setSelectedKadroFilter] = useState('');

  // Security Passcode (PIN 1839) Protection
  const [isAuthorized, setIsAuthorized] = useState(false);
  const [isPinModalOpen, setIsPinModalOpen] = useState(false);
  const [pinInput, setPinInput] = useState('');
  const [pinError, setPinError] = useState(false);
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null);

  // Test Mail Modal State
  const [isTestMailModalOpen, setIsTestMailModalOpen] = useState(false);
  const [testMailPinInput, setTestMailPinInput] = useState('');
  const [testMailPinError, setTestMailPinError] = useState(false);
  const [testMailAuthorized, setTestMailAuthorized] = useState(false);
  const [isSimulatingMail, setIsSimulatingMail] = useState(false);

  // Add / Edit Personnel Modal
  const [isFormModalOpen, setIsFormModalOpen] = useState(false);
  const [formMode, setFormMode] = useState<'add' | 'edit'>('add');
  const [editingPersonId, setEditingPersonId] = useState<string | null>(null);

  // Delete Personnel In-App Confirmation Modal
  const [personToDelete, setPersonToDelete] = useState<PersonelBilgiRecord | null>(null);
  const [deletePinInput, setDeletePinInput] = useState('');
  const [deletePinError, setDeletePinError] = useState(false);

  // Form Fields
  const [formData, setFormData] = useState({
    siraNo: 1,
    adiSoyadi: '',
    tcKimlik: '',
    sicilNo: '',
    kadroUnvani: '',
    dogumTarihi: '',
    gorevYeri: '',
    telefonNo: '',
    kanGrubu: '',
    adresBilgisi: '',
    yakininAdiSoyadi: '',
    esTelefonNumaralari: '',
    periyodikMuayeneTarihi: ''
  });

  // Selected Detail Modal
  const [selectedDetailPerson, setSelectedDetailPerson] = useState<PersonelBilgiRecord | null>(null);



  // Non-blocking asynchronous live Excel fetch
  const fetchLiveExcelFromDrive = async (isManual = false) => {
    if (isManual) {
      setIsLoadingDrive(true);
      setSyncStatusMessage("Google Drive E-Tablo Aranıyor...");
    }

    try {
      const result = await fetchExcelFromDrive(
        DRIVE_FILE_NAME,
        ['personel_bilgi_cizelgesi.xlsx', 'personel_bilgi.xlsx'],
        DRIVE_FOLDER_ID
      );

      if (result && result.status === 'success' && result.base64) {
        const success = parseExcelAndSetPeople(result.base64);
        if (success && showNotification && isManual) {
          showNotification("✅ Google Drive'dan Personel Bilgi Çizelgesi canlı olarak yüklendi.");
        } else if (!success) {
           console.warn("Drive fetch success but parse failed");
           if (showNotification) showNotification("⚠️ Veri alındı ancak işlenemedi (başlık satırı bulunamadı).");
        }
      } else {
        console.warn("Drive fetch returned no data or failed");
        if (showNotification) showNotification("⚠️ Google Drive'dan veri çekilemedi.");
      }
    } catch (err) {
      console.warn("Drive Excel fetch warning:", err);
      if (showNotification) showNotification("⚠️ Google Drive bağlantısı başarısız.");
    } finally {
      setIsLoadingDrive(false);
      setSyncStatusMessage('');
    }
  };

  useEffect(() => {
    if (isOpen) {
      setIsLoadingDrive(true); // Ensure it's blurred immediately
      setSyncStatusMessage("Veriler güncelleniyor...");
      setTimeout(() => {
        fetchLiveExcelFromDrive(false);
      }, 50);
    }
  }, [isOpen]);

  // Helper to auto sync technicians to Personel Egitim
  const autoSyncTechnicianToPersonelEgitim = async (adiSoyadi: string, kadroUnvani: string) => {
    const norm = (kadroUnvani || '').toUpperCase();
    const isTech = norm.includes('TEKNİSYEN') || norm.includes('TEKNİKER') || norm.includes('TEKNISYEN') || norm.includes('TEKNIKER') || norm.includes('UÇUŞ') || norm.includes('UCUS');
    if (!isTech || !adiSoyadi) return;

    try {
      const res = await fetchExcelFromDrive('personel_egitim_listesi.xlsx', ['personel_egitim_listesi.xlsx'], DRIVE_FOLDER_ID);
      if (res && res.status === 'success' && res.base64) {
        const binary = atob(res.base64);
        const len = binary.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
        const wb = XLSX.read(bytes, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        if (!ws) return;

        const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        const exists = rawRows.some(r => (r[1] || '').toString().toUpperCase().trim() === adiSoyadi.toUpperCase().trim());
        if (!exists) {
          const lastSira = rawRows.length > 1 ? (parseInt(rawRows[rawRows.length - 1][0]) || rawRows.length) : 1;
          rawRows.push([
            lastSira + 1,
            adiSoyadi.toUpperCase().trim(),
            '-',
            '-',
            '-',
            '-',
            '-',
            'BELGE YÜKLE (https://bulut.ogm.gov.tr/BELGEYUKLE)',
            '-'
          ]);
          const newWs = XLSX.utils.aoa_to_sheet(rawRows);
          wb.Sheets[wb.SheetNames[0]] = newWs;
          const outBase64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
          await uploadExcelToDrive('personel_egitim_listesi.xlsx', outBase64, 'personel_egitim_listesi', DRIVE_FOLDER_ID);
        }
      }
    } catch (err) {
      console.warn("Auto-sync technician warning:", err);
    }
  };

  // Parse Excel Base64
  const parseExcelAndSetPeople = (base64: string): boolean => {
    try {
      const binary = atob(base64);
      const len = binary.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      const wb = XLSX.read(bytes, { type: 'array' });
      const firstSheetName = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheetName];
      if (!ws) return false;

      const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!rawRows || rawRows.length === 0) return false;

      let headerIdx = 0;
      for (let r = 0; r < Math.min(10, rawRows.length); r++) {
        const rowStr = (rawRows[r] || []).map(c => String(c || '').toLowerCase()).join(' ');
        if (rowStr.includes('adı') || rowStr.includes('soyadi') || rowStr.includes('sicil') || rowStr.includes('sıra')) {
          headerIdx = r;
          break;
        }
      }

      const headers = (rawRows[headerIdx] || []).map(c => String(c || '').trim());
      const findCol = (kw: string[]) => {
        return headers.findIndex(h => {
          const l = h.toLowerCase();
          return kw.some(k => l.includes(k));
        });
      };

      const cSira = findCol(['sıra', 'sira', 'no']);
      const cAdi = findCol(['adı', 'adi', 'ad soyad', 'isim']);
      const cTc = findCol(['tc', 't.c', 'kimlik']);
      const cSicil = findCol(['sicil']);
      const cKadro = findCol(['kadro', 'unvan', 'görev', 'gorev']);
      const cDogum = findCol(['doğum', 'dogum']);
      const cGorevYeri = findCol(['görev yeri', 'gorev yeri', 'yeri']);
      const cTel = findCol(['telefon', 'tel no', 'cep']);
      const cKan = findCol(['kan']);
      const cAdres = findCol(['adres']);
      const cYakin = findCol(['yakın', 'yakin', 'es adı', 'eş adı']);
      const cYakinTel = findCol(['eş telefon', 'es telefon', 'yakın tel', 'yakin tel']);
      const cMuayene = findCol(['muayene', 'periyodik', 'ogm']);

      const parsed: PersonelBilgiRecord[] = [];

      for (let r = headerIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;

        const rawAdi = String(row[cAdi >= 0 ? cAdi : 1] || '').trim();
        if (!rawAdi || rawAdi.toLowerCase() === 'adı soyadı' || rawAdi.toLowerCase().includes('toplam')) continue;

        const rawSira = row[cSira >= 0 ? cSira : 0];
        const numSira = Number(rawSira) || (parsed.length + 1);

        parsed.push({
          id: `pb_${r}_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
          siraNo: numSira,
          adiSoyadi: rawAdi,
          tcKimlik: String(row[cTc >= 0 ? cTc : 2] || '').trim(),
          sicilNo: String(row[cSicil >= 0 ? cSicil : 3] || '').trim(),
          kadroUnvani: String(row[cKadro >= 0 ? cKadro : 4] || '').trim(),
          dogumTarihi: cleanAndFormatDateString(String(row[cDogum >= 0 ? cDogum : 5] || '')),
          gorevYeri: String(row[cGorevYeri >= 0 ? cGorevYeri : 6] || '').trim(),
          telefonNo: String(row[cTel >= 0 ? cTel : 7] || '').trim(),
          kanGrubu: String(row[cKan >= 0 ? cKan : 8] || '').trim(),
          adresBilgisi: String(row[cAdres >= 0 ? cAdres : 9] || '').trim(),
          yakininAdiSoyadi: String(row[cYakin >= 0 ? cYakin : 10] || '').trim(),
          esTelefonNumaralari: String(row[cYakinTel >= 0 ? cYakinTel : 11] || '').trim(),
          periyodikMuayeneTarihi: (() => {
            const rawVal = String(row[cMuayene >= 0 ? cMuayene : 12] || '');
            const lowerVal = rawVal.toLowerCase().trim();
            if (!rawVal || lowerVal === 'hayır' || lowerVal === 'hayir' || lowerVal === '-' || lowerVal === 'null' || lowerVal === 'undefined') return '-';
            const val = cleanAndFormatDateString(rawVal);
            return (val && val.toLowerCase() !== 'hayır' && val.toLowerCase() !== 'hayir' && val !== '-') ? val : '-';
          })()
        });
      }

      if (parsed.length > 0) {
        setPeople(parsed);
        return true;
      }
      return false;
    } catch (e) {
      console.error('Excel parse error:', e);
      return false;
    }
  };

  // Build Excel Base64
  const buildExcelBase64FromPeople = (peopleList: PersonelBilgiRecord[]): string => {
    const aoa: any[][] = [
      [
        'Sıra No',
        'Adı Soyadı',
        'T.C. Kimlik',
        'Sicil No',
        'Kadro Unvanı / Görevi',
        'Doğum Tarihi',
        'Görev Yeri',
        'Telefon No',
        'Kan Grubu',
        'Adres Bilgisi',
        'Yakının Adı-Soyadı',
        'Eş Telefon Numaraları',
        'PERİYODİK MUAYENE (OGM) TARİHİ'
      ]
    ];

    peopleList.forEach((p, idx) => {
      aoa.push([
        p.siraNo || (idx + 1),
        p.adiSoyadi || '',
        p.tcKimlik || '',
        p.sicilNo || '',
        p.kadroUnvani || '',
        p.dogumTarihi || '',
        p.gorevYeri || '',
        p.telefonNo || '',
        p.kanGrubu || '',
        p.adresBilgisi || '',
        p.yakininAdiSoyadi || '',
        p.esTelefonNumaralari || '',
        (p.periyodikMuayeneTarihi && p.periyodikMuayeneTarihi.toLowerCase() !== 'hayır' && p.periyodikMuayeneTarihi.toLowerCase() !== 'hayir' && p.periyodikMuayeneTarihi !== '-') ? p.periyodikMuayeneTarihi : '-'
      ]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    return XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
  };

  const syncToGoogleDrive = async (updatedPeople: PersonelBilgiRecord[], actionMsg: string) => {
    setIsLoadingDrive(true);
    setSyncStatusMessage(actionMsg);
    try {
      const b64 = buildExcelBase64FromPeople(updatedPeople);
      await uploadExcelToDrive(DRIVE_FILE_NAME, b64, 'personel_bilgi');
      if (showNotification) {
        showNotification("✅ Personel bilgileri Google Drive'a canlı kaydedildi.");
      }
    } catch (e) {
      console.warn("Drive sync error:", e);
    } finally {
      setIsLoadingDrive(false);
      setSyncStatusMessage('');
    }
  };

  // Unique Kadro List for Filter
  const uniqueKadroList = useMemo(() => {
    const set = new Set<string>();
    people.forEach(p => {
      if (p.kadroUnvani && p.kadroUnvani.trim()) {
        set.add(p.kadroUnvani.trim().toUpperCase());
      }
    });
    return Array.from(set).sort();
  }, [people]);

  // Filtered People
  const filteredPeople = useMemo(() => {
    return people.filter(p => {
      const matchesSearch =
        !searchQuery.trim() ||
        p.adiSoyadi.toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.sicilNo.toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.tcKimlik.toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.telefonNo.toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.gorevYeri.toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.kadroUnvani.toLowerCase().includes(searchQuery.toLowerCase());

      const matchesKadro =
        !selectedKadroFilter ||
        p.kadroUnvani.trim().toUpperCase() === selectedKadroFilter.toUpperCase();

      return matchesSearch && matchesKadro;
    });
  }, [people, searchQuery, selectedKadroFilter]);

  // Expiration Warning Calculation for Periyodik Muayene
  const expiringRecords = useMemo(() => {
    const list: { person: PersonelBilgiRecord; daysLeft: number; expiryDateStr: string }[] = [];
    people.forEach(p => {
      const status = getExpiryStatus(p.periyodikMuayeneTarihi, 5);
      if (status.hasDate && (status.isWarning || status.isExpired)) {
        list.push({
          person: p,
          daysLeft: status.daysLeft,
          expiryDateStr: status.expiryDateStr
        });
      }
    });
    return list;
  }, [people]);

  // Security Verification Guard (Protects Add, Edit, Delete with PIN 1839)
  const requirePin = (action: () => void) => {
    if (isAuthorized) {
      action();
      return;
    }
    setPendingAction(() => action);
    setPinInput('');
    setPinError(false);
    setIsPinModalOpen(true);
  };

  const handleVerifyPin = () => {
    if (pinInput.trim() === REQUIRED_PIN) {
      setIsAuthorized(true);
      setIsPinModalOpen(false);
      setPinError(false);
      if (pendingAction) {
        pendingAction();
        setPendingAction(null);
      }
    } else {
      setPinError(true);
    }
  };

  const handleOpenAddModal = () => {
    requirePin(() => {
      setFormMode('add');
      setEditingPersonId(null);
      const nextSira = people.length > 0 ? Math.max(...people.map(p => Number(p.siraNo) || 0)) + 1 : 1;
      setFormData({
        siraNo: nextSira,
        adiSoyadi: '',
        tcKimlik: '',
        sicilNo: '',
        kadroUnvani: '',
        dogumTarihi: '',
        gorevYeri: '',
        telefonNo: '',
        kanGrubu: '',
        adresBilgisi: '',
        yakininAdiSoyadi: '',
        esTelefonNumaralari: '',
        periyodikMuayeneTarihi: '-'
      });
      setIsFormModalOpen(true);
    });
  };

  const handleOpenEditModal = (person: PersonelBilgiRecord) => {
    requirePin(() => {
      setFormMode('edit');
      setEditingPersonId(person.id);
      setFormData({
        siraNo: person.siraNo,
        adiSoyadi: person.adiSoyadi,
        tcKimlik: person.tcKimlik,
        sicilNo: person.sicilNo,
        kadroUnvani: person.kadroUnvani,
        dogumTarihi: person.dogumTarihi,
        gorevYeri: person.gorevYeri,
        telefonNo: person.telefonNo,
        kanGrubu: person.kanGrubu,
        adresBilgisi: person.adresBilgisi,
        yakininAdiSoyadi: person.yakininAdiSoyadi,
        esTelefonNumaralari: person.esTelefonNumaralari,
        periyodikMuayeneTarihi: person.periyodikMuayeneTarihi || '-'
      });
      setIsFormModalOpen(true);
    });
  };

  const handleSaveForm = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.adiSoyadi.trim()) {
      alert("Lütfen personelin adını ve soyadını giriniz.");
      return;
    }

    let updatedList: PersonelBilgiRecord[];

    if (formMode === 'add') {
      const newPerson: PersonelBilgiRecord = {
        id: `pb_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
        siraNo: Number(formData.siraNo) || (people.length + 1),
        adiSoyadi: formData.adiSoyadi.trim().toUpperCase(),
        tcKimlik: formData.tcKimlik.trim(),
        sicilNo: formData.sicilNo.trim(),
        kadroUnvani: formData.kadroUnvani.trim(),
        dogumTarihi: formData.dogumTarihi.trim(),
        gorevYeri: formData.gorevYeri.trim().toUpperCase(),
        telefonNo: formData.telefonNo.trim(),
        kanGrubu: formData.kanGrubu.trim().toUpperCase(),
        adresBilgisi: formData.adresBilgisi.trim(),
        yakininAdiSoyadi: formData.yakininAdiSoyadi.trim(),
        esTelefonNumaralari: formData.esTelefonNumaralari.trim(),
        periyodikMuayeneTarihi: cleanAndFormatDateString(formData.periyodikMuayeneTarihi) || formData.periyodikMuayeneTarihi || '-'
      };
      updatedList = [...people, newPerson];
    } else {
      updatedList = people.map(p => {
        if (p.id === editingPersonId) {
          return {
            ...p,
            siraNo: Number(formData.siraNo) || p.siraNo,
            adiSoyadi: formData.adiSoyadi.trim().toUpperCase(),
            tcKimlik: formData.tcKimlik.trim(),
            sicilNo: formData.sicilNo.trim(),
            kadroUnvani: formData.kadroUnvani.trim(),
            dogumTarihi: formData.dogumTarihi.trim(),
            gorevYeri: formData.gorevYeri.trim().toUpperCase(),
            telefonNo: formData.telefonNo.trim(),
            kanGrubu: formData.kanGrubu.trim().toUpperCase(),
            adresBilgisi: formData.adresBilgisi.trim(),
            yakininAdiSoyadi: formData.yakininAdiSoyadi.trim(),
            esTelefonNumaralari: formData.esTelefonNumaralari.trim(),
            periyodikMuayeneTarihi: cleanAndFormatDateString(formData.periyodikMuayeneTarihi) || formData.periyodikMuayeneTarihi || '-'
          };
        }
        return p;
      });
    }

    updatedList.sort((a, b) => Number(a.siraNo) - Number(b.siraNo));
    setPeople(updatedList);
    setIsFormModalOpen(false);

    syncToGoogleDrive(
      updatedList,
      formMode === 'add' ? 'Yeni Personel Google Drive E-Tabloya Ekleniyor...' : 'Personel Güncellemesi Drive\'a İşleniyor...'
    );

    autoSyncTechnicianToPersonelEgitim(formData.adiSoyadi, formData.kadroUnvani);
  };

  const handleDeletePerson = (person: PersonelBilgiRecord) => {
    setPersonToDelete(person);
    setDeletePinInput('');
    setDeletePinError(false);
  };

  const handleConfirmDeletePerson = () => {
    if (!personToDelete) return;
    if (deletePinInput.trim() !== REQUIRED_PIN && !isAuthorized) {
      setDeletePinError(true);
      return;
    }

    const targetName = personToDelete.adiSoyadi;
    const targetId = personToDelete.id;
    const updatedList = people.filter(p => p.id !== targetId);
    setPeople(updatedList);
    setPersonToDelete(null);
    setDeletePinInput('');
    setDeletePinError(false);
    setIsAuthorized(true);

    if (showNotification) {
      showNotification(`🗑️ "${targetName}" personeli silindi. Google Drive güncelleniyor...`);
    }

    syncToGoogleDrive(updatedList, `"${targetName}" Personel Kaydı Siliniyor...`);
  };

  // Test Mail PIN Handler
  const handleVerifyTestMailPin = (e: React.FormEvent) => {
    e.preventDefault();
    if (testMailPinInput.trim() === REQUIRED_PIN) {
      setTestMailAuthorized(true);
      setTestMailPinError(false);
      setTestMailPinInput('');
    } else {
      setTestMailPinError(true);
    }
  };

  const handleSimulateSendMail = () => {
    setIsSimulatingMail(true);
    setTimeout(() => {
      setIsSimulatingMail(false);
      setIsTestMailModalOpen(false);
      setTestMailAuthorized(false);
      alert(`✅ Test maili simüle edildi!\nOtomatik Gönderim Tarihi: ${getNextMonthFirstDayStr()}\nAlıcı: ormanhavacilik.bakimsube@gmail.com\nToplam Uyarı Sayısı: ${expiringRecords.length}`);
    }, 1000);
  };

  // Excel Download
  const handleDownloadExcel = () => {
    try {
      if (filteredPeople.length === 0) {
        alert("İndirilecek personel verisi bulunamadı.");
        return;
      }

      let rowsHtml = '';
      filteredPeople.forEach((p, idx) => {
        rowsHtml += `
          <tr>
            <td style="text-align: center; border: 1px solid #cbd5e1; font-weight: bold; background-color: #f1f5f9;">${p.siraNo || (idx + 1)}</td>
            <td style="border: 1px solid #cbd5e1; font-weight: bold; background-color: #dbeafe; padding: 6px;">${p.adiSoyadi}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center;">${p.tcKimlik || '-'}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center;">${p.sicilNo || '-'}</td>
            <td style="border: 1px solid #cbd5e1;">${p.kadroUnvani || '-'}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center;">${p.dogumTarihi || '-'}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center;">${p.gorevYeri || '-'}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center;">${p.telefonNo || '-'}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center;">${p.kanGrubu || '-'}</td>
            <td style="border: 1px solid #cbd5e1;">${p.adresBilgisi || '-'}</td>
            <td style="border: 1px solid #cbd5e1;">${p.yakininAdiSoyadi || '-'}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center;">${p.esTelefonNumaralari || '-'}</td>
            <td style="border: 1px solid #cbd5e1; text-align: center; font-weight: bold;">${p.periyodikMuayeneTarihi || '-'}</td>
          </tr>
        `;
      });

      const html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="utf-8">
          <style>
            table { border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
            th { background-color: #090f1e; color: #ffffff; font-weight: bold; padding: 10px; border: 1px solid #000000; text-align: left; }
            td { padding: 6px; border: 1px solid #cbd5e1; vertical-align: middle; }
          </style>
        </head>
        <body>
          <h2 style="color: #0b3d1d; font-family: Calibri, Arial, sans-serif; margin-bottom: 8px;">
            ORMAN GENEL MÜDÜRLÜĞÜ - PERSONEL BİLGİ SİSTEMİ
          </h2>
          <table border="1">
            <thead>
              <tr style="background-color: #090f1e; color: #ffffff;">
                <th style="text-align: center; width: 50px;">SIRA</th>
                <th style="width: 200px;">ADI SOYADI</th>
                <th style="text-align: center; width: 130px;">T.C. KİMLİK</th>
                <th style="text-align: center; width: 90px;">SİCİL</th>
                <th style="width: 180px;">KADRO UNV.</th>
                <th style="text-align: center; width: 120px;">DOĞUM T.</th>
                <th style="text-align: center; width: 120px;">GÖREV YERİ</th>
                <th style="text-align: center; width: 140px;">TELEFON NO</th>
                <th style="text-align: center; width: 70px;">KAN</th>
                <th style="width: 240px;">ADRES BİLGİSİ</th>
                <th style="width: 180px;">YAKININ ADI</th>
                <th style="text-align: center; width: 140px;">YAKIN TEL NO</th>
                <th style="text-align: center; width: 180px;">PERİYODİK MUAYENE (OGM) TARİHİ</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>
        </body>
        </html>
      `;

      const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `personel_bilgi_cizelgesi_${new Date().toISOString().slice(0, 10)}.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      if (showNotification) {
        showNotification("📥 Personel Bilgi Çizelgesi Excel olarak indirildi.");
      }
    } catch (e: any) {
      alert("Excel indirme hatası: " + e.message);
    }
  };

  // PDF Export
  const handleDownloadPdf = () => {
    try {
      if (filteredPeople.length === 0) {
        alert("Yazdırılacak personel verisi bulunamadı.");
        return;
      }

      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4'
      });

      doc.setFontSize(14);
      doc.text("ORMAN GENEL MÜDÜRLÜĞÜ - PERSONEL BİLGİ SİSTEMİ", 14, 15);
      doc.setFontSize(9);
      doc.text(`Tarih: ${new Date().toLocaleDateString('tr-TR')} | Toplam Personel: ${filteredPeople.length}`, 14, 22);

      let y = 30;
      doc.setFontSize(8);
      doc.setFillColor(9, 15, 30);
      doc.rect(14, y, 269, 8, 'F');
      doc.setTextColor(255, 255, 255);
      doc.text("SIRA", 16, y + 5.5);
      doc.text("ADI SOYADI", 28, y + 5.5);
      doc.text("SİCİL", 85, y + 5.5);
      doc.text("KADRO UNVANI", 105, y + 5.5);
      doc.text("GÖREV YERİ", 155, y + 5.5);
      doc.text("TELEFON", 195, y + 5.5);
      doc.text("PERİYODİK MUAYENE (OGM)", 230, y + 5.5);

      y += 8;
      doc.setTextColor(0, 0, 0);

      filteredPeople.forEach((p, idx) => {
        if (y > 185) {
          doc.addPage();
          y = 15;
        }
        if (idx % 2 === 0) {
          doc.setFillColor(245, 247, 250);
          doc.rect(14, y, 269, 7, 'F');
        }
        doc.text(String(p.siraNo || (idx + 1)), 16, y + 5);
        doc.text(p.adiSoyadi.slice(0, 28), 28, y + 5);
        doc.text(p.sicilNo || '-', 85, y + 5);
        doc.text((p.kadroUnvani || '-').slice(0, 22), 105, y + 5);
        doc.text((p.gorevYeri || '-').slice(0, 18), 155, y + 5);
        doc.text(p.telefonNo || '-', 195, y + 5);
        doc.text(p.periyodikMuayeneTarihi || '-', 230, y + 5);

        y += 7;
      });

      doc.save("personel_bilgi_cizelgesi.pdf");
      if (showNotification) {
        showNotification("📥 Personel Bilgi Çizelgesi PDF olarak indirildi.");
      }
    } catch (e: any) {
      alert("PDF indirme hatası: " + e.message);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[600] flex items-center justify-center bg-black/85 backdrop-blur-md p-2 sm:p-4 animate-fade-in font-sans">
      <div className="relative w-full max-w-[99vw] 2xl:max-w-[1780px] h-[95vh] bg-[#080d1a] border border-slate-800 rounded-3xl shadow-2xl flex flex-col overflow-hidden text-slate-100">
        
        {/* TOP BAR - STYLED EXACTLY MATCHING image_b60e11d6.png */}
        <div className="shrink-0 bg-[#090d18] border-b border-slate-800/80 px-6 py-4 flex flex-col lg:flex-row items-center justify-between gap-4 select-none">
          
          {/* Left Title: PERSONEL | BİLGİ SİSTEMİ */}
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-1.5 font-mono text-sm font-black tracking-wider">
              <span className="text-cyan-400">PERSONEL</span>
              <span className="text-cyan-400 font-bold">|</span>
              <span className="text-white">BİLGİ SİSTEMİ</span>
            </div>
            {expiringRecords.length > 0 && (
              <span className="text-[10px] bg-amber-500/10 text-amber-400 border border-amber-500/30 font-bold px-2.5 py-0.5 rounded-full flex items-center gap-1">
                <AlertTriangle className="w-3 h-3 text-amber-400" />
                {expiringRecords.length} Personel Muayene Yaklaştı
              </span>
            )}
          </div>

          {/* Search Input: Q ARA (İSİM, SİCİL...) */}
          <div className="flex items-center gap-2 bg-slate-950 border border-slate-700/80 px-4 py-2 rounded-full min-w-[240px]">
            <Search className="w-4 h-4 text-cyan-400 shrink-0" />
            <input
              type="text"
              placeholder="ARA (İSİM, SİCİL...)"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="bg-transparent text-white font-bold text-xs outline-none uppercase w-full placeholder-slate-400 font-mono"
            />
            {searchQuery && (
              <button onClick={() => setSearchQuery('')} className="text-slate-400 hover:text-white cursor-pointer">
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {/* Kadro Filter Dropdown */}
          <div className="flex items-center gap-1.5 bg-slate-950 border border-slate-700/80 px-4 py-2 rounded-full font-mono text-xs">
            <span className="text-slate-400 font-bold">KADRO:</span>
            <select
              value={selectedKadroFilter}
              onChange={(e) => setSelectedKadroFilter(e.target.value)}
              className="bg-transparent text-white font-black outline-none uppercase cursor-pointer"
            >
              <option value="" className="bg-slate-900">HEPSİ / TÜMÜ</option>
              {uniqueKadroList.map(k => (
                <option key={k} value={k} className="bg-slate-900">{k.toUpperCase()}</option>
              ))}
            </select>
          </div>

          {/* Right Action Buttons Matching image_b60e11d6.png */}
          <div className="flex items-center gap-2 flex-wrap">
            {/* + + YENİ PERSONEL EKLE */}
            <button
              onClick={handleOpenAddModal}
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black font-mono text-xs rounded-full flex items-center gap-1.5 transition-all cursor-pointer shadow-md active:scale-95"
            >
              <Plus className="w-4 h-4 stroke-[3]" />
              <span>+ YENİ PERSONEL EKLE</span>
            </button>

            {/* DRİVE'DAN VERİLERİ YENİLE */}
            <button
              onClick={() => fetchLiveExcelFromDrive(true)}
              disabled={isLoadingDrive}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-black font-mono text-xs rounded-full flex items-center gap-1.5 transition-all cursor-pointer shadow-md active:scale-95"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoadingDrive ? 'animate-spin' : ''}`} />
              <span>DRİVE'DAN VERİLERİ YENİLE</span>
            </button>

            {/* EXCEL İNDİR */}
            <button
              onClick={handleDownloadExcel}
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black font-mono text-xs rounded-full flex items-center gap-1.5 transition-all cursor-pointer shadow-md active:scale-95"
            >
              <Download className="w-3.5 h-3.5" />
              <span>EXCEL İNDİR</span>
            </button>

            {/* PDF İNDİR */}
            <button
              onClick={handleDownloadPdf}
              className="px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white font-black font-mono text-xs rounded-full flex items-center gap-1.5 transition-all cursor-pointer shadow-md active:scale-95"
            >
              <FileText className="w-3.5 h-3.5" />
              <span>PDF İNDİR</span>
            </button>

            {/* TEST MAİLİ GÖNDER */}
            <button
              onClick={() => {
                setIsTestMailModalOpen(true);
                setTestMailAuthorized(false);
                setTestMailPinError(false);
              }}
              className="px-4 py-2 bg-amber-600 hover:bg-amber-500 text-white font-black font-mono text-xs rounded-full flex items-center gap-1.5 transition-all cursor-pointer shadow-md active:scale-95"
            >
              <Mail className="w-3.5 h-3.5" />
              <span>TEST MAİLİ GÖNDER</span>
            </button>

            {/* YAZDIR / PDF YAP */}
            <button
              onClick={handleDownloadPdf}
              className="px-4 py-2 border-2 border-purple-500 hover:bg-purple-600/20 text-purple-300 font-black font-mono text-xs rounded-full flex items-center gap-1.5 transition-all cursor-pointer shadow-sm active:scale-95"
            >
              <FileText className="w-3.5 h-3.5 text-purple-400" />
              <span>YAZDIR / PDF YAP</span>
            </button>

            {/* Kapat Button */}
            <button
              onClick={onClose}
              className="p-2 rounded-full bg-slate-800 hover:bg-red-600 text-slate-300 hover:text-white transition cursor-pointer ml-1"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

        </div>

        {/* CENTRAL WHITE CARD CONTAINER MATCHING image_b60e11d6.png */}
        <div className="flex-1 overflow-auto p-4 sm:p-6 bg-[#080d1a] flex flex-col justify-center items-center">
          <div className="w-full max-w-[1720px] h-full bg-white rounded-[2.5rem] shadow-2xl p-6 sm:p-8 flex flex-col overflow-hidden border border-slate-200 text-slate-800 relative">
            
            {/* Table Container with Blur while loading */}
            <div className={`flex-1 overflow-auto rounded-2xl border border-slate-300 shadow-inner relative transition-all duration-300 ${isLoadingDrive ? 'filter blur-sm select-none pointer-events-none' : ''}`}>
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-[#090f1e] text-white font-black uppercase tracking-wider text-[11px] font-mono select-none">
                    <th className="p-3 text-center border-r border-slate-700 w-12">SIRA</th>
                    <th className="p-3 text-left border-r border-slate-700 min-w-[180px]">ADI SOYADI</th>
                    <th className="p-3 text-center border-r border-slate-700 min-w-[120px]">T.C. KİMLİK</th>
                    <th className="p-3 text-center border-r border-slate-700 w-20">SİCİL</th>
                    <th className="p-3 text-left border-r border-slate-700 min-w-[140px]">KADRO UNV.</th>
                    <th className="p-3 text-center border-r border-slate-700 min-w-[110px]">DOĞUM T.</th>
                    <th className="p-3 text-center border-r border-slate-700 min-w-[110px]">GÖREV YERİ</th>
                    <th className="p-3 text-center border-r border-slate-700 min-w-[130px]">TELEFON NO</th>
                    <th className="p-3 text-center border-r border-slate-700 w-16">KAN</th>
                    <th className="p-3 text-left border-r border-slate-700 min-w-[200px]">ADRES BİLGİSİ</th>
                    <th className="p-3 text-left border-r border-slate-700 min-w-[150px]">YAKININ ADI</th>
                    <th className="p-3 text-center border-r border-slate-700 min-w-[130px]">YAKIN TEL NO</th>
                    <th className="p-3 text-center border-r border-slate-700 min-w-[190px] text-amber-300 font-mono">
                      PERİYODİK MUAYENE (OGM) TARİHİ
                    </th>
                    <th className="p-3 text-center w-24">İŞLEM</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {filteredPeople.length === 0 ? (
                    <tr>
                      <td colSpan={14} className="text-center py-20 text-slate-400 font-mono text-xs uppercase tracking-wider">
                        <span>🚫 Arama kriterlerine uygun personel kaydı bulunamadı.</span>
                      </td>
                    </tr>
                  ) : (
                    filteredPeople.map((person, idx) => {
                      const isRowOne = idx === 0;
                      const muayeneStatus = getExpiryStatus(person.periyodikMuayeneTarihi, 5);

                      return (
                        <tr
                          key={person.id}
                          onDoubleClick={() => setSelectedDetailPerson(person)}
                          className={`transition-colors font-semibold text-xs ${
                            isRowOne
                              ? 'bg-[#3b82f6] text-white font-bold'
                              : 'bg-white hover:bg-slate-100 text-slate-800'
                          }`}
                        >
                          <td className="p-3 text-center font-bold font-mono border-r border-slate-200">
                            {person.siraNo}
                          </td>
                          <td className="p-3 font-extrabold border-r border-slate-200">
                            {person.adiSoyadi}
                          </td>
                          <td className="p-3 text-center font-mono border-r border-slate-200">
                            {person.tcKimlik || '-'}
                          </td>
                          <td className="p-3 text-center font-mono border-r border-slate-200">
                            {person.sicilNo || '-'}
                          </td>
                          <td className="p-3 border-r border-slate-200">
                            {person.kadroUnvani || '-'}
                          </td>
                          <td className="p-3 text-center font-mono border-r border-slate-200 whitespace-nowrap">
                            {person.dogumTarihi || '-'}
                          </td>
                          <td className="p-3 text-center font-bold border-r border-slate-200">
                            {person.gorevYeri || '-'}
                          </td>
                          <td className="p-3 text-center font-mono border-r border-slate-200 whitespace-nowrap">
                            {person.telefonNo || '-'}
                          </td>
                          <td className="p-3 text-center font-mono border-r border-slate-200">
                            {person.kanGrubu || '-'}
                          </td>
                          <td className="p-3 border-r border-slate-200 text-[11px] truncate max-w-[220px]">
                            {person.adresBilgisi || '-'}
                          </td>
                          <td className="p-3 border-r border-slate-200 font-medium">
                            {person.yakininAdiSoyadi || '-'}
                          </td>
                          <td className="p-3 text-center font-mono border-r border-slate-200 whitespace-nowrap">
                            {person.esTelefonNumaralari || '-'}
                          </td>
                          {/* PERİYODİK MUAYENE (OGM) TARİHİ SÜTUNU */}
                          <td className="p-3 text-center font-mono font-bold border-r border-slate-200 whitespace-nowrap">
                            {person.periyodikMuayeneTarihi && person.periyodikMuayeneTarihi !== '-' && person.periyodikMuayeneTarihi.toLowerCase() !== 'hayır' ? (
                              <div className="flex flex-col items-center justify-center gap-0.5">
                                <span className={`px-2 py-0.5 rounded-full border text-[10px] font-bold ${
                                  muayeneStatus.isExpired
                                    ? 'bg-red-100 text-red-800 border-red-300'
                                    : muayeneStatus.isWarning
                                    ? 'bg-amber-100 text-amber-900 border-amber-300'
                                    : isRowOne ? 'bg-white/20 text-white border-white/40' : 'bg-slate-100 text-slate-800 border-slate-300'
                                }`}>
                                  {person.periyodikMuayeneTarihi}
                                </span>
                                {muayeneStatus.isWarning && (
                                  <span className="text-[8px] font-black text-amber-800 bg-amber-50 px-1 rounded border border-amber-300">
                                    ⚠️ {muayeneStatus.daysLeft} Gün
                                  </span>
                                )}
                                {muayeneStatus.isExpired && (
                                  <span className="text-[8px] font-black text-red-700 bg-red-50 px-1 rounded border border-red-300">
                                    🚨 DOLDU
                                  </span>
                                )}
                              </div>
                            ) : (
                              <span className="opacity-50">-</span>
                            )}
                          </td>
                          {/* İŞLEM SÜTUNU */}
                          <td className="p-3 text-center whitespace-nowrap">
                            <div className="flex items-center justify-center gap-1.5">
                              <button
                                onClick={() => handleOpenEditModal(person)}
                                className={`p-1.5 rounded-xl border transition cursor-pointer shadow-xs ${
                                  isRowOne
                                    ? 'bg-white/20 border-white/40 text-white hover:bg-white/30'
                                    : 'bg-blue-50 border-blue-200 text-blue-600 hover:bg-blue-100'
                                }`}
                                title="Düzenle (Şifre: 1839)"
                              >
                                <Edit2 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => handleDeletePerson(person)}
                                className={`p-1.5 rounded-xl border transition cursor-pointer shadow-xs ${
                                  isRowOne
                                    ? 'bg-red-500/80 border-red-300 text-white hover:bg-red-600'
                                    : 'bg-red-50 border-red-200 text-red-600 hover:bg-red-100'
                                }`}
                                title="Sil (Şifre: 1839)"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Footer inside Card matching image_b60e11d6.png */}
            <div className="mt-4 pt-3 border-t border-slate-200 flex items-center justify-between text-[10px] text-slate-400 font-mono uppercase tracking-widest font-semibold">
              <span>İMZA / ONAY: HAVACILIK DAİRESİ BAŞKANLIĞI</span>
              <span>SAYFA 1 / 1</span>
              <span>PORTAL SÜRÜM: V3.8-E (EXCEL-PDF)</span>
            </div>

            {/* Loading Overlay with Blur indicator */}
            {isLoadingDrive && (
              <div className="absolute inset-0 z-40 bg-white/80 backdrop-blur-md flex flex-col items-center justify-center gap-3 text-slate-800 rounded-[2.5rem]">
                <div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin" />
                <div className="text-center px-4">
                  <div className="font-mono text-sm font-black text-slate-900 tracking-wider uppercase">
                    {syncStatusMessage || "VERİLER GÜNCELLENİYOR..."}
                  </div>
                  <div className="text-xs text-slate-500 font-semibold mt-1">
                    Güncel personel verileri çekilene kadar lütfen bekleyiniz...
                  </div>
                </div>
              </div>
            )}

          </div>
        </div>

      </div>

      {/* PERSONEL SİLME ONAY MODALI */}
      {personToDelete && (
        <div className="fixed inset-0 z-[750] bg-black/90 backdrop-blur-xl flex items-center justify-center p-4 animate-fade-in font-sans">
          <div className="bg-[#0b1329] border border-red-500/60 text-white rounded-[2.5rem] shadow-2xl p-8 max-w-md w-full space-y-6 animate-scale-up border-t-8 border-t-red-600">
            <div className="flex flex-col items-center text-center gap-4">
              <div className="w-20 h-20 rounded-3xl bg-red-600 text-white flex items-center justify-center shadow-lg shadow-red-950/40 animate-bounce">
                <Trash2 className="w-10 h-10" />
              </div>
              <div>
                <h2 className="text-4xl font-black text-red-500 uppercase tracking-tighter mb-1">
                  SİLİNSİN Mİ?
                </h2>
                <p className="text-slate-300 font-bold text-lg px-2">
                  <span className="text-white underline underline-offset-4 decoration-red-500 decoration-4">{personToDelete.adiSoyadi}</span>
                </p>
                <p className="text-[11px] text-slate-400 mt-2 font-mono uppercase tracking-widest">
                  İsimli personeli sistemden kalıcı olarak silmek istediğinize emin misiniz?
                </p>
              </div>
            </div>

            <div className="p-4 bg-slate-950/80 rounded-2xl border border-slate-800/80 space-y-2 font-mono text-[10px]">
              <div className="flex items-center justify-between border-b border-slate-800 pb-1.5 text-slate-400 uppercase">
                <span>T.C. KİMLİK:</span>
                <span className="text-slate-200 font-black">{personToDelete.tcKimlik || '-'}</span>
              </div>
              <div className="flex items-center justify-between border-b border-slate-800 pb-1.5 text-slate-400 uppercase">
                <span>SİCİL NO:</span>
                <span className="text-slate-200 font-black">{personToDelete.sicilNo || '-'}</span>
              </div>
              <div className="flex items-center justify-between text-slate-400 uppercase">
                <span>KADRO:</span>
                <span className="text-slate-200 font-black truncate max-w-[200px]">{personToDelete.kadroUnvani || '-'}</span>
              </div>
            </div>

            <div className="bg-slate-900/80 p-6 rounded-3xl border-2 border-slate-800 space-y-3 shadow-inner">
              <label className="block text-[11px] font-mono font-black text-red-400 mb-1.5 uppercase text-center tracking-widest">
                İŞLEMİ ONAYLAMAK İÇİN ŞİFREYİ GİRİN:
              </label>
              <input
                type="password"
                autoFocus
                placeholder="****"
                value={deletePinInput}
                onChange={(e) => {
                  setDeletePinInput(e.target.value);
                  setDeletePinError(false);
                }}
                onKeyDown={(e) => e.key === 'Enter' && handleConfirmDeletePerson()}
                className="w-full px-4 py-4 bg-black border-2 border-slate-700 focus:border-red-600 rounded-2xl font-mono text-3xl font-black text-center tracking-[0.6em] outline-none text-white transition-all shadow-2xl"
              />
              {deletePinError && (
                <span className="text-red-500 text-[10px] font-black mt-1 block text-center uppercase animate-pulse">
                  ❌ HATALI ŞİFRE! (YETKİLİ: 1839)
                </span>
              )}
            </div>

            <div className="grid grid-cols-2 gap-4">
              <button
                type="button"
                onClick={() => {
                  setPersonToDelete(null);
                  setDeletePinInput('');
                  setDeletePinError(false);
                }}
                className="py-4 bg-slate-800 hover:bg-slate-700 text-slate-200 font-black rounded-2xl text-xs cursor-pointer transition active:scale-95 uppercase tracking-widest"
              >
                VAZGEÇ
              </button>
              <button
                type="button"
                onClick={handleConfirmDeletePerson}
                className="py-4 bg-red-600 hover:bg-red-500 text-white font-black rounded-2xl text-xs cursor-pointer transition shadow-xl shadow-red-950/20 active:scale-95 uppercase tracking-widest flex items-center justify-center gap-2"
              >
                <span>EVET, SİL</span>
                <Check className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* SECURITY PIN MODAL (1839) */}
      {isPinModalOpen && (
        <div className="fixed inset-0 z-[700] bg-black/80 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 text-white rounded-3xl shadow-2xl p-6 max-w-sm w-full space-y-4 animate-scale-up">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-2xl bg-blue-500/10 border border-blue-500/20 flex items-center justify-center text-blue-400">
                <Lock className="w-5 h-5" />
              </div>
              <div>
                <h3 className="font-extrabold text-sm uppercase tracking-wider font-mono">
                  YETKİLİ GİRİŞİ (1839)
                </h3>
                <p className="text-[11px] text-slate-400">
                  İşleme devam etmek için şifrenizi giriniz.
                </p>
              </div>
            </div>

            <div>
              <input
                type="password"
                autoFocus
                placeholder="****"
                value={pinInput}
                onChange={(e) => setPinInput(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleVerifyPin()}
                className="w-full px-4 py-3 bg-slate-950 border border-slate-800 rounded-2xl font-mono text-lg font-bold text-center tracking-widest outline-none focus:border-blue-500"
              />
              {pinError && (
                <span className="text-red-400 text-xs font-bold mt-1 block text-center">Hatalı Şifre! Lütfen 1839 giriniz.</span>
              )}
            </div>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setIsPinModalOpen(false)}
                className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-bold rounded-xl cursor-pointer"
              >
                Vazgeç
              </button>
              <button
                type="button"
                onClick={handleVerifyPin}
                className="px-5 py-2.5 bg-blue-600 hover:bg-blue-500 text-white font-black text-xs rounded-xl shadow cursor-pointer"
              >
                GİRİŞ YAP
              </button>
            </div>
          </div>
        </div>
      )}

      {/* TEST MAİLİ MODALI */}
      {isTestMailModalOpen && (
        <div className="fixed inset-0 z-[750] bg-black/80 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 text-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden animate-scale-up">
            <div className="bg-gradient-to-r from-amber-950 via-slate-900 to-amber-950 p-5 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-amber-500/20 border border-amber-500/40 flex items-center justify-center text-amber-400">
                  <Mail className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-sm uppercase tracking-wider font-mono text-amber-400">
                    PERİYODİK MUAYENE TEST MAİLİ
                  </h3>
                  <p className="text-[11px] text-slate-400 font-mono">
                    Periyodik Muayene (OGM - 5 Yıl) Otomatik Takip Servisi
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsTestMailModalOpen(false)}
                className="p-1.5 text-slate-400 hover:text-white rounded-xl hover:bg-slate-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {!testMailAuthorized ? (
              <form onSubmit={handleVerifyTestMailPin} className="p-6 space-y-4">
                <div className="flex items-center gap-2 bg-amber-500/10 border border-amber-500/20 p-3 rounded-2xl text-amber-300 text-xs">
                  <Lock className="w-4 h-4 shrink-0 text-amber-400" />
                  <span>Test maili yetkilendirme doğrulama adımı. Lütfen yetkili şifresini giriniz.</span>
                </div>

                <div>
                  <label className="block text-xs font-bold uppercase font-mono text-slate-300 mb-2">
                    YETKİLİ ŞİFRESİ (1839)
                  </label>
                  <input
                    type="password"
                    autoFocus
                    placeholder="****"
                    value={testMailPinInput}
                    onChange={(e) => setTestMailPinInput(e.target.value)}
                    className="w-full px-4 py-3 bg-slate-950 border border-slate-800 rounded-2xl font-mono text-lg font-bold text-center tracking-widest outline-none focus:border-amber-500"
                  />
                  {testMailPinError && (
                    <span className="text-red-400 text-xs font-bold mt-1 block text-center">Hatalı Şifre! Lütfen 1839 giriniz.</span>
                  )}
                </div>

                <div className="flex items-center justify-end gap-2 pt-2">
                  <button
                    type="button"
                    onClick={() => setIsTestMailModalOpen(false)}
                    className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-bold rounded-xl"
                  >
                    Vazgeç
                  </button>
                  <button
                    type="submit"
                    className="px-5 py-2.5 bg-amber-600 hover:bg-amber-500 text-white font-black text-xs rounded-xl shadow cursor-pointer"
                  >
                    ŞİFREYİ DOĞRULA
                  </button>
                </div>
              </form>
            ) : (
              <div className="p-6 space-y-4 text-xs">
                <div className="bg-slate-950 border border-slate-800 p-4 rounded-2xl space-y-2">
                  <div className="flex items-center justify-between text-[11px] font-mono border-b border-slate-800 pb-2">
                    <span className="text-slate-400">HEDEF E-POSTA:</span>
                    <span className="font-bold text-emerald-400">ormanhavacilik.bakimsube@gmail.com</span>
                  </div>
                  <div className="flex items-center justify-between text-[11px] font-mono border-b border-slate-800 pb-2">
                    <span className="text-slate-400">ZAMANLANMIŞ MAİL TARİHİ:</span>
                    <span className="font-bold text-amber-400 font-mono flex items-center gap-1">
                      <Clock className="w-3.5 h-3.5 text-amber-400" />
                      {getNextMonthFirstDayStr()} (BİR SONRAKİ AYIN İLK GÜNÜ)
                    </span>
                  </div>
                  <div className="flex items-center justify-between text-[11px] font-mono">
                    <span className="text-slate-400">UYARISI YAKLAŞAN / DOLAN PERSONEL:</span>
                    <span className="font-extrabold text-white bg-amber-500/20 px-2 py-0.5 rounded border border-amber-500/30 font-mono">
                      {expiringRecords.length} Personel
                    </span>
                  </div>
                </div>

                {expiringRecords.length > 0 ? (
                  <div className="max-h-48 overflow-y-auto space-y-2 pr-1">
                    {expiringRecords.map((item, idx) => (
                      <div key={idx} className="bg-slate-950 border border-slate-800 p-2.5 rounded-xl flex items-center justify-between text-[11px]">
                        <div>
                          <span className="font-bold text-white block">{item.person.adiSoyadi}</span>
                          <span className="text-[10px] text-amber-400 font-mono">
                            Periyodik Muayene (OGM - 5 Yıl) • Bitiş: {item.expiryDateStr}
                          </span>
                        </div>
                        <span className="font-bold font-mono text-amber-300 bg-amber-950/80 px-2 py-1 rounded border border-amber-700/50">
                          {item.daysLeft <= 0 ? 'Süre Doldu' : `${item.daysLeft} Gün Kaldı`}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="bg-emerald-950/40 border border-emerald-800/40 p-3 rounded-2xl text-emerald-300 text-center font-bold">
                    ✅ Süresi 1 ay kalan veya dolan herhangi bir periyodik muayene kaydı bulunmamaktadır.
                  </div>
                )}

                <div className="pt-2 flex items-center justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setIsTestMailModalOpen(false)}
                    className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold rounded-xl"
                  >
                    Kapat
                  </button>
                  <button
                    type="button"
                    disabled={isSimulatingMail}
                    onClick={handleSimulateSendMail}
                    className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-xl shadow cursor-pointer flex items-center gap-2 active:scale-95"
                  >
                    {isSimulatingMail ? (
                      <>
                        <RefreshCw className="w-4 h-4 animate-spin" />
                        <span>SİMÜLE EDİLİYOR...</span>
                      </>
                    ) : (
                      <>
                        <Mail className="w-4 h-4" />
                        <span>ŞİMDİ SİMÜLE ET & GÖNDER</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* FORM MODAL (ADD / EDIT) */}
      {isFormModalOpen && (
        <div className="fixed inset-0 z-[700] bg-black/80 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 text-white rounded-3xl shadow-2xl max-w-2xl w-full overflow-hidden animate-scale-up">
            <div className="bg-gradient-to-r from-slate-950 via-slate-900 to-slate-950 p-5 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center text-emerald-400">
                  <Users className="w-5 h-5" />
                </div>
                <h3 className="font-extrabold text-sm uppercase tracking-wider font-mono">
                  {formMode === 'add' ? 'YENİ PERSONEL EKLE' : 'PERSONEL BİLGİLERİNİ DÜZENLE'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setIsFormModalOpen(false)}
                className="p-1.5 text-slate-400 hover:text-white rounded-xl hover:bg-slate-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveForm} className="p-6 space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">SIRA NO</label>
                  <input
                    type="number"
                    value={formData.siraNo}
                    onChange={(e) => setFormData({ ...formData, siraNo: parseInt(e.target.value) || 1 })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold font-mono outline-none text-white focus:border-emerald-500"
                  />
                </div>
                <div className="col-span-2">
                  <label className="block text-slate-400 font-mono font-bold mb-1">ADI SOYADI *</label>
                  <input
                    type="text"
                    required
                    value={formData.adiSoyadi}
                    onChange={(e) => setFormData({ ...formData, adiSoyadi: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold uppercase outline-none text-white focus:border-emerald-500"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">T.C. KİMLİK</label>
                  <input
                    type="text"
                    value={formData.tcKimlik}
                    onChange={(e) => setFormData({ ...formData, tcKimlik: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold font-mono outline-none text-white focus:border-emerald-500"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">SİCİL NO</label>
                  <input
                    type="text"
                    value={formData.sicilNo}
                    onChange={(e) => setFormData({ ...formData, sicilNo: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold font-mono outline-none text-white focus:border-emerald-500"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">KADRO UNVANI</label>
                  <input
                    type="text"
                    value={formData.kadroUnvani}
                    onChange={(e) => setFormData({ ...formData, kadroUnvani: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold outline-none text-white focus:border-emerald-500"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">DOĞUM TARİHİ</label>
                  <input
                    type="text"
                    placeholder="20 Temmuz 1975"
                    value={formData.dogumTarihi}
                    onChange={(e) => setFormData({ ...formData, dogumTarihi: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold outline-none text-white focus:border-emerald-500 font-mono"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">GÖREV YERİ</label>
                  <input
                    type="text"
                    placeholder="OGM / HANGAR..."
                    value={formData.gorevYeri}
                    onChange={(e) => setFormData({ ...formData, gorevYeri: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold uppercase outline-none text-white focus:border-emerald-500"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">TELEFON NO</label>
                  <input
                    type="text"
                    placeholder="532 230 37 37"
                    value={formData.telefonNo}
                    onChange={(e) => setFormData({ ...formData, telefonNo: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold font-mono outline-none text-white focus:border-emerald-500"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">KAN GRUBU</label>
                  <input
                    type="text"
                    placeholder="A Rh (+)"
                    value={formData.kanGrubu}
                    onChange={(e) => setFormData({ ...formData, kanGrubu: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold outline-none text-white focus:border-emerald-500 font-mono"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">YAKININ ADI</label>
                  <input
                    type="text"
                    value={formData.yakininAdiSoyadi}
                    onChange={(e) => setFormData({ ...formData, yakininAdiSoyadi: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold outline-none text-white focus:border-emerald-500"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 font-mono font-bold mb-1">YAKIN TEL NO</label>
                  <input
                    type="text"
                    value={formData.esTelefonNumaralari}
                    onChange={(e) => setFormData({ ...formData, esTelefonNumaralari: e.target.value })}
                    className="w-full px-3.5 py-2.5 bg-slate-950 border border-slate-800 rounded-xl font-bold font-mono outline-none text-white focus:border-emerald-500"
                  />
                </div>
              </div>

              {/* PERİYODİK MUAYENE (OGM) TARİHİ INPUT */}
              <div>
                <label className="block text-amber-400 font-mono font-bold mb-1">
                  PERİYODİK MUAYENE (OGM) TARİHİ (5 YIL GEÇERLİ)
                </label>
                <input
                  type="text"
                  placeholder="GG.AA.YYYY"
                  value={formData.periyodikMuayeneTarihi}
                  onChange={(e) => setFormData({ ...formData, periyodikMuayeneTarihi: e.target.value })}
                  className="w-full px-3.5 py-2.5 bg-slate-950 border border-amber-500/50 rounded-xl font-bold text-amber-300 font-mono outline-none focus:border-amber-400"
                />
              </div>

              <div>
                <label className="block text-slate-400 font-mono font-bold mb-1">ADRES BİLGİSİ</label>
                <textarea
                  rows={2}
                  value={formData.adresBilgisi}
                  onChange={(e) => setFormData({ ...formData, adresBilgisi: e.target.value })}
                  className="w-full px-3.5 py-2 bg-slate-950 border border-slate-800 rounded-xl font-bold text-xs outline-none text-white focus:border-emerald-500 resize-none"
                />
              </div>

              <div className="pt-3 border-t border-slate-800 flex items-center justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setIsFormModalOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold rounded-xl"
                >
                  Vazgeç
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-xl shadow cursor-pointer active:scale-95"
                >
                  KAYDET & DRİVE'A İŞLE
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* DETAIL MODAL (DOUBLE CLICK) */}
      {selectedDetailPerson && (
        <div className="fixed inset-0 z-[700] bg-black/80 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 text-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden animate-scale-up">
            <div className="bg-gradient-to-r from-blue-950 via-slate-900 to-blue-950 p-5 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-blue-500/20 border border-blue-500/40 flex items-center justify-center text-blue-400">
                  <Users className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-sm uppercase tracking-wider font-mono text-white">
                    {selectedDetailPerson.adiSoyadi}
                  </h3>
                  <p className="text-[11px] text-blue-300 font-mono">
                    {selectedDetailPerson.kadroUnvani || 'Personel Detay Kartı'}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setSelectedDetailPerson(null)}
                className="p-1.5 text-slate-400 hover:text-white rounded-xl hover:bg-slate-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 space-y-3 text-xs font-sans">
              <div className="grid grid-cols-2 gap-3 bg-slate-950 p-3.5 rounded-2xl border border-slate-800 font-mono">
                <div><span className="text-slate-500">SIRA NO:</span> <span className="font-bold text-white ml-1">{selectedDetailPerson.siraNo}</span></div>
                <div><span className="text-slate-500">T.C. KİMLİK:</span> <span className="font-bold text-white ml-1">{selectedDetailPerson.tcKimlik || '-'}</span></div>
                <div><span className="text-slate-500">SİCİL NO:</span> <span className="font-bold text-white ml-1">{selectedDetailPerson.sicilNo || '-'}</span></div>
                <div><span className="text-slate-500">GÖREV YERİ:</span> <span className="font-bold text-white ml-1">{selectedDetailPerson.gorevYeri || '-'}</span></div>
                <div><span className="text-slate-500">DOĞUM T.:</span> <span className="font-bold text-white ml-1">{selectedDetailPerson.dogumTarihi || '-'}</span></div>
                <div><span className="text-slate-500">KAN GRUBU:</span> <span className="font-bold text-white ml-1">{selectedDetailPerson.kanGrubu || '-'}</span></div>
              </div>

              <div className="space-y-2 font-mono bg-slate-950 p-3.5 rounded-2xl border border-slate-800">
                <div><span className="text-slate-500">TELEFON:</span> <span className="font-bold text-emerald-400 ml-1">{selectedDetailPerson.telefonNo || '-'}</span></div>
                <div><span className="text-slate-500">ADRES:</span> <span className="font-bold text-slate-300 ml-1 text-[11px]">{selectedDetailPerson.adresBilgisi || '-'}</span></div>
                <div><span className="text-slate-500">YAKININ ADI:</span> <span className="font-bold text-white ml-1">{selectedDetailPerson.yakininAdiSoyadi || '-'}</span></div>
                <div><span className="text-slate-500">YAKIN TEL:</span> <span className="font-bold text-emerald-400 ml-1">{selectedDetailPerson.esTelefonNumaralari || '-'}</span></div>
                <div><span className="text-amber-400 font-bold">PERİYODİK MUAYENE (OGM):</span> <span className="font-bold text-amber-300 ml-1">{selectedDetailPerson.periyodikMuayeneTarihi || '-'}</span></div>
              </div>

              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => {
                    const target = selectedDetailPerson;
                    setSelectedDetailPerson(null);
                    handleOpenEditModal(target);
                  }}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-xl text-xs flex items-center gap-1.5"
                >
                  <Edit2 className="w-3.5 h-3.5" />
                  <span>Düzenle</span>
                </button>
                <button
                  type="button"
                  onClick={() => setSelectedDetailPerson(null)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold rounded-xl text-xs"
                >
                  Kapat
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};
