import React, { useState, useEffect, useMemo } from 'react';
import { 
  ShieldAlert, Plus, Search, RefreshCw, Download, Printer, 
  CheckCircle2, ArrowRightLeft, Building2, User, Calendar, 
  CheckSquare, Square, Trash2, Edit2, X, AlertTriangle, FileSpreadsheet,
  Layers, ChevronDown, Eye
} from 'lucide-react';
import * as XLSX from 'xlsx';

export interface YasamDestekRecord {
  id: string;
  siraNo?: number;
  personelAdi: string;
  gorevYeri: string;
  baslangicTarihi: string;
  bitisTarihi?: string;
  lifeVestSn: string;
  spareAirSn: string;
  helmetKitSn: string;
  durum: 'GÖREVDE' | 'DEVREDİLDİ' | 'ANKARA DEPO TESLİM' | string;
  devredilenPersonel?: string;
  teslimAlanPersonel?: string;
  notlar?: string;
  guncellenmeTarihi?: string;
}

interface YasamDestekPanelProps {
  records: YasamDestekRecord[];
  isLoading: boolean;
  onRefresh: () => Promise<void>;
  onSaveRecord: (record: YasamDestekRecord) => Promise<boolean>;
  onDeleteRecord: (id: string) => Promise<boolean>;
  showNotification: (msg: string) => void;
  triggerExport?: boolean;
  onExportComplete?: () => void;
}

export const YasamDestekPanel: React.FC<YasamDestekPanelProps> = ({
  records,
  isLoading,
  onRefresh,
  onSaveRecord,
  onDeleteRecord,
  showNotification,
  triggerExport,
  onExportComplete
}) => {
  // Listen for external export trigger from Parent (DepoManagementModal)
  useEffect(() => {
    if (triggerExport) {
      handleExportMasterExcel();
      if (onExportComplete) onExportComplete();
    }
  }, [triggerExport]);
  // Active Sub-Tab
  const [activeSubTab, setActiveSubTab] = useState<'can_yelegi' | 'spare_air' | 'helmet_kit' | 'zimmetler' | 'personnel'>('can_yelegi');

  // Stock Lists from Server
  const [canYelegiList, setCanYelegiList] = useState<any[]>([]);
  const [spareAirList, setSpareAirList] = useState<any[]>([]);
  const [helmetKitList, setHelmetKitList] = useState<any[]>([]);
  const [isStockLoading, setIsStockLoading] = useState(false);

  // Search & Filter
  const [searchQuery, setSearchQuery] = useState('');
  const [filterStatus, setFilterStatus] = useState<string>('ALL');

  // New Record Modal
  const [isNewModalOpen, setIsNewModalOpen] = useState(false);
  const [editingRecord, setEditingRecord] = useState<YasamDestekRecord | null>(null);

  // Form State
  const [formDurum, setFormDurum] = useState('DEPODA');
  const [formPersonel, setFormPersonel] = useState('');
  const [formGorevYeri, setFormGorevYeri] = useState('MİLAS');
  const [formCustomLocation, setFormCustomLocation] = useState('');
  const [formBaslangic, setFormBaslangic] = useState('');
  const [formLifeVest, setFormLifeVest] = useState('');
  const [formSpareAir, setFormSpareAir] = useState('');
  const [formHelmetKit, setFormHelmetKit] = useState('');
  const [formNotlar, setFormNotlar] = useState('');

  // Finish / Handover Modal
  const [finishTargetRecord, setFinishTargetRecord] = useState<YasamDestekRecord | null>(null);
  const [finishBitisTarihi, setFinishBitisTarihi] = useState('');
  const [finishActionType, setFinishActionType] = useState<'devret' | 'teslim'>('devret');
  const [devredilecekPersonel, setDevredilecekPersonel] = useState('');
  const [devredilecekGorevYeri, setDevredilecekGorevYeri] = useState('');
  const [devredilecekCustomLocation, setDevredilecekCustomLocation] = useState('');
  const [teslimAlanGorevli, setTeslimAlanGorevli] = useState('Ankara Depo Sorumlusu');
  const [checkLifeVest, setCheckLifeVest] = useState(true);
  const [checkSpareAir, setCheckSpareAir] = useState(true);
  const [checkHelmetKit, setCheckHelmetKit] = useState(true);
  const [finishNotes, setFinishNotes] = useState('');

  // Swap Item Modal State
  const [swapTargetRecord, setSwapTargetRecord] = useState<YasamDestekRecord | null>(null);
  const [swapItemType, setSwapItemType] = useState<'can_yelegi' | 'spare_air' | 'helmet_kit'>('can_yelegi');
  const [swapNewSn, setSwapNewSn] = useState('');
  const [swapNotes, setSwapNotes] = useState('');

  // Print Tutanak Modal
  const [printRecord, setPrintRecord] = useState<YasamDestekRecord | null>(null);

  // Viewing Personnel Zimmet Detail Modal
  const [viewingPersonnel, setViewingPersonnel] = useState<string | null>(null);

  // Personnel State
  const [personnelList, setPersonnelList] = useState<string[]>([]);
  const [personnelItems, setPersonnelItems] = useState<any[]>([]);
  const [personnelRoleFilter, setPersonnelRoleFilter] = useState<'ALL' | 'Teknisyen' | 'Pilot'>('ALL');

  // Fetch stocks & personnel
  const fetchAllStocks = async () => {
    try {
      setIsStockLoading(true);
      const res = await fetch('/api/yasam-destek/all-stocks');
      const json = await res.json();
      if (json && json.status === 'success') {
        setCanYelegiList(json.canYelegi || []);
        setSpareAirList(json.spareAir || []);
        setHelmetKitList(json.helmetKit || []);
      }
    } catch (e) {
      console.warn('Stock fetch error:', e);
    } finally {
      setIsStockLoading(false);
    }
  };

  const fetchPersonnel = async () => {
    try {
      const res = await fetch('/api/yasam-destek/personnel-list');
      const json = await res.json();
      console.log('DEBUG: Personnel list response:', json);
      if (json && json.status === 'success') {
        if (Array.isArray(json.personnel)) setPersonnelList(json.personnel);
        if (Array.isArray(json.items)) setPersonnelItems(json.items);
      } else {
        console.warn('DEBUG: Personnel list fetch unsuccessful:', json);
      }
    } catch (e) {
      console.error('DEBUG: Personnel list fetch error:', e);
    }
  };

  useEffect(() => {
    fetchAllStocks();
    fetchPersonnel();
    // Auto refresh every 10 minutes
    const interval = setInterval(() => {
      fetchAllStocks();
      fetchPersonnel();
    }, 10 * 60 * 1000);
    return () => clearInterval(interval);
  }, []);

  // Stock Row Edit Modal State
  const [editingStockItem, setEditingStockItem] = useState<{ type: 'can_yelegi' | 'spare_air' | 'helmet_kit', item: any, index: number } | null>(null);
  const [stockEditPersonel, setStockEditPersonel] = useState('');
  const [stockEditGorevYeri, setStockEditGorevYeri] = useState('MİLAS');
  const [stockEditSeriNo, setStockEditSeriNo] = useState('');
  const [stockEditSeriNo2, setStockEditSeriNo2] = useState('');
  const [stockEditStatus, setStockEditStatus] = useState('GÖREV BÖLGESİNDE');
  const [stockEditTarih, setStockEditTarih] = useState('');

  const openStockEditModal = (type: 'can_yelegi' | 'spare_air' | 'helmet_kit', item: any, index: number) => {
    setEditingStockItem({ type, item, index });
    setStockEditStatus(item.durum || (item.aciklamalar === 'DEPODA' ? 'DEPODA' : 'GÖREV BÖLGESİNDE'));
    
    if (item.durum === 'GÖREV BÖLGESİNDE') {
        setStockEditPersonel(item.aciklamalar || '');
        setStockEditGorevYeri(item.gorevYeri || 'MİLAS');
    } else {
        setStockEditPersonel('');
        setStockEditGorevYeri('MİLAS');
    }
    
    setStockEditSeriNo(item.seriNo || item.seriNo1 || '');
    setStockEditSeriNo2(item.seriNo2 || '');
    setStockEditTarih(item.tarih || new Date().toLocaleDateString('tr-TR'));
  };

  const handleSaveStockEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingStockItem) return;
    const { type, index } = editingStockItem;

    let targetList = type === 'can_yelegi' ? [...canYelegiList] : (type === 'spare_air' ? [...spareAirList] : [...helmetKitList]);
    if (index >= 0 && index < targetList.length) {
      const isDepot = stockEditStatus === 'BOŞTA' || stockEditStatus === 'DEPODA';
      const color = isDepot ? '#ffffff' : (stockEditStatus === 'ONARIMDA' ? '#ffff00' : (stockEditStatus === 'BAKIMA GİDECEK' ? '#00b0f0' : '#92d050'));
      
      targetList[index] = {
        ...targetList[index],
        aciklamalar: stockEditStatus === 'GÖREV BÖLGESİNDE' ? stockEditPersonel : 'DEPODA',
        gorevYeri: stockEditStatus === 'GÖREV BÖLGESİNDE' ? stockEditGorevYeri : '-',
        seriNo: stockEditSeriNo || targetList[index].seriNo,
        seriNo1: type === 'helmet_kit' ? (stockEditSeriNo || targetList[index].seriNo1) : targetList[index].seriNo1,
        seriNo2: type === 'helmet_kit' ? (stockEditSeriNo2 || targetList[index].seriNo2) : targetList[index].seriNo2,
        durum: stockEditStatus,
        color,
        tarih: stockEditTarih
      };

      if (type === 'can_yelegi') setCanYelegiList(targetList);
      else if (type === 'spare_air') setSpareAirList(targetList);
      else if (type === 'helmet_kit') setHelmetKitList(targetList);

      try {
        const res = await fetch('/api/yasam-destek/stock-update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ type, items: targetList })
        });
        const json = await res.json();
        if (json.status === 'success' && Array.isArray(json.records)) {
          await onRefresh();
        }
        await fetchAllStocks();
        await fetchPersonnel();
      } catch (e) {
        console.warn('Stock update fetch error:', e);
      }

      showNotification('✅ Stok kaydı ve teslim yeri güncellendi.');
    }
    setEditingStockItem(null);
  };

  const handleDeleteStockItem = async (type: 'can_yelegi' | 'spare_air' | 'helmet_kit', index: number) => {
    if (!window.confirm('Bu stok kaydını silmek istediğinizden emin misiniz?')) return;
    let targetList = type === 'can_yelegi' ? [...canYelegiList] : (type === 'spare_air' ? [...spareAirList] : [...helmetKitList]);
    targetList.splice(index, 1);
    if (type === 'can_yelegi') setCanYelegiList(targetList);
    else if (type === 'spare_air') setSpareAirList(targetList);
    else if (type === 'helmet_kit') setHelmetKitList(targetList);

    try {
      const res = await fetch('/api/yasam-destek/stock-update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type, items: targetList })
      });
      const json = await res.json();
      if (json.status === 'success' && Array.isArray(json.records)) {
        await onRefresh();
      }
      await fetchAllStocks();
      await fetchPersonnel();
    } catch (e) {
      console.warn('Stock delete error:', e);
    }
    showNotification('🗑️ Stok kaydı silindi.');
  };

  const handleQuickChangeStockStatus = async (type: 'can_yelegi' | 'spare_air' | 'helmet_kit', index: number, newStatus: string) => {
    let targetList = type === 'can_yelegi' ? [...canYelegiList] : (type === 'spare_air' ? [...spareAirList] : [...helmetKitList]);
    if (index >= 0 && index < targetList.length) {
      let aciklamaText = newStatus;
      if (newStatus === 'BOŞTA' || newStatus === 'DEPODA') {
        aciklamaText = 'DEPODA';
      }
      
      const isDepot = newStatus === 'BOŞTA' || newStatus === 'DEPODA';
      const color = isDepot ? '#ffffff' : (newStatus === 'ONARIMDA' ? '#ffff00' : (newStatus === 'BAKIMA GİDECEK' ? '#00b0f0' : (newStatus === 'KAYIT SİLME' ? '#ff0000' : '#92d050')));
      
      targetList[index] = {
        ...targetList[index],
        durum: isDepot ? 'DEPODA' : newStatus,
        aciklamalar: aciklamaText,
        color: color,
        tarih: isDepot ? '-' : new Date().toLocaleDateString('tr-TR')
      };

      if (type === 'can_yelegi') setCanYelegiList(targetList);
      else if (type === 'spare_air') setSpareAirList(targetList);
      else if (type === 'helmet_kit') setHelmetKitList(targetList);

      try {
        const res = await fetch('/api/yasam-destek/stock-update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ type, items: targetList })
        });
        const json = await res.json();
        if (json.status === 'success' && Array.isArray(json.records)) {
          await onRefresh();
        }
        await fetchAllStocks();
        await fetchPersonnel();
      } catch (e) {
        console.warn('Stock quick status error:', e);
      }

      showNotification('✅ Durum güncellendi ve zimmetler senkronize edildi.');
    }
  };

  // Search Filtered Lists
  const filterItemByStatus = (item: any) => {
    if (!filterStatus || filterStatus === 'ALL') return true;
    const a = (item.aciklamalar || '').toUpperCase().trim();
    const d = (item.durum || '').toUpperCase().trim();
    const isDepot = d === 'DEPODA' || d === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
    if (filterStatus === 'BOŞTA') return isDepot;
    if (filterStatus === 'ONARIMDA') return d === 'ONARIMDA' || a.includes('ONARIM');
    if (filterStatus === 'GÖREV BÖLGESİNDE') return !isDepot && d !== 'KAYIT SİLME' && d !== 'ONARIMDA' && d !== 'BAKIMA GİDECEK';
    if (filterStatus === 'KAYIT SİLME') return d === 'KAYIT SİLME' || a.includes('KAYIT SİLME');
    if (filterStatus === 'BAKIMA GİDECEK') return d === 'BAKIMA GİDECEK' || a.includes('BAKIM') || a.includes('DOLUM') || a.includes('ARIZALI');
    return true;
  };

  const filteredCanYelegi = useMemo(() => {
    const statusFiltered = canYelegiList.filter(filterItemByStatus);
    if (!searchQuery.trim()) return statusFiltered;
    const q = searchQuery.toLowerCase();
    return statusFiltered.filter(item => 
      (item.aciklamalar || '').toLowerCase().includes(q) ||
      (item.seriNo || '').toLowerCase().includes(q) ||
      (item.disNo || '').toLowerCase().includes(q) ||
      (item.durum || '').toLowerCase().includes(q) ||
      (item.tarih || '').toLowerCase().includes(q)
    );
  }, [canYelegiList, searchQuery, filterStatus]);

  const filteredSpareAir = useMemo(() => {
    const statusFiltered = spareAirList.filter(filterItemByStatus);
    if (!searchQuery.trim()) return statusFiltered;
    const q = searchQuery.toLowerCase();
    return statusFiltered.filter(item => 
      (item.aciklamalar || '').toLowerCase().includes(q) ||
      (item.seriNo || '').toLowerCase().includes(q) ||
      (item.durum || '').toLowerCase().includes(q) ||
      (item.tarih || '').toLowerCase().includes(q)
    );
  }, [spareAirList, searchQuery, filterStatus]);

  const filteredHelmetKit = useMemo(() => {
    const statusFiltered = helmetKitList.filter(filterItemByStatus);
    if (!searchQuery.trim()) return statusFiltered;
    const q = searchQuery.toLowerCase();
    return statusFiltered.filter(item => 
      (item.aciklamalar || '').toLowerCase().includes(q) ||
      (item.seriNo1 || '').toLowerCase().includes(q) ||
      (item.seriNo2 || '').toLowerCase().includes(q) ||
      (item.durum || '').toLowerCase().includes(q) ||
      (item.tarih || '').toLowerCase().includes(q)
    );
  }, [helmetKitList, searchQuery, filterStatus]);

  const filteredZimmetler = useMemo(() => {
    const activeRecords = records; // Filtreyi kaldırarak tüm kayıtları (ANKARA DEPO TESLİM dahil) gösteriyoruz.
    const statusFiltered = activeRecords.filter(r => {
      if (!filterStatus || filterStatus === 'ALL') return true;
      const d = (r.durum || '').toUpperCase().trim();
      if (filterStatus === 'GÖREV BÖLGESİNDE') return d === 'GÖREVDE' || d === 'GÖREV BÖLGESİNDE';
      return d === filterStatus;
    });
    if (!searchQuery.trim()) return statusFiltered;
    const q = searchQuery.toLowerCase();
    return statusFiltered.filter(r => 
      (r.personelAdi || '').toLowerCase().includes(q) ||
      (r.gorevYeri || '').toLowerCase().includes(q) ||
      (r.lifeVestSn || '').toLowerCase().includes(q) ||
      (r.spareAirSn || '').toLowerCase().includes(q) ||
      (r.helmetKitSn || '').toLowerCase().includes(q) ||
      (r.baslangicTarihi || '').toLowerCase().includes(q) ||
      (r.bitisTarihi || '').toLowerCase().includes(q) ||
      (r.durum || '').toLowerCase().includes(q) ||
      (r.devredilenPersonel || '').toLowerCase().includes(q) ||
      (r.notlar || '').toLowerCase().includes(q)
    );
  }, [records, searchQuery, filterStatus]);
  const STATUS_LEGEND = [
    { label: 'BOŞTA', bg: 'bg-white', text: 'text-slate-800', border: 'border-slate-300' },
    { label: 'ONARIMDA', bg: 'bg-[#ffff00]', text: 'text-slate-900', border: 'border-amber-400' },
    { label: 'GÖREV BÖLGESİNDE', bg: 'bg-[#92d050]', text: 'text-slate-900', border: 'border-emerald-600' },
    { label: 'KAYIT SİLME', bg: 'bg-[#ff0000]', text: 'text-white', border: 'border-rose-700' },
    { label: 'BAKIMA GİDECEK', bg: 'bg-[#00b0f0]', text: 'text-slate-900', border: 'border-cyan-600' }
  ];

  // Helper function to resolve row color
  const getRowColorClass = (item: any) => {
    // 1. Status based coloring (Priority)
    const status = (item.durum || '').toUpperCase();
    
    if (status === 'KAYIT SİLME') return 'bg-[#ff0000] text-white font-bold';
    if (status === 'BAKIMA GİDECEK') return 'bg-[#00b0f0] text-slate-900 font-bold';
    if (status === 'ONARIMDA') return 'bg-[#ffff00] text-slate-900 font-bold';
    
    // Check if it's explicitly in Depot status, regardless of aciklamalar
    if (status === 'DEPODA' || status === 'BOŞTA') {
      return 'bg-white hover:bg-slate-50 text-slate-800'; // Dolgusuz
    }

    // 2. Assignment status check
    const assigned = (item.aciklamalar || '').trim();
    const isDepot = !assigned || assigned === '-' || assigned.toUpperCase() === 'DEPODA' || assigned.toUpperCase() === 'BOŞTA';

    if (isDepot) {
      return 'bg-white hover:bg-slate-50 text-slate-800'; // Dolgusuz
    }
    
    // Default assigned row color
    return 'bg-[#92d050]/80 text-slate-900 font-bold hover:bg-[#92d050]';
  };

  // Open New Zimmet Form
  const openNewModal = () => {
    setEditingRecord(null);
    setFormDurum('DEPODA');
    setFormPersonel('');
    setFormGorevYeri('MİLAS');
    setFormCustomLocation('');
    const now = new Date();
    setFormBaslangic(now.toLocaleDateString('tr-TR'));
    
    // Default to first available stock serial numbers if available
    const firstVest = canYelegiList.find(c => c.durum !== 'KAYIT SİLME')?.seriNo || '';
    const firstAir = spareAirList.find(s => s.durum !== 'KAYIT SİLME')?.seriNo || '';
    const firstHelmet = helmetKitList.find(h => h.durum !== 'KAYIT SİLME')?.aciklamalar || '';
    
    setFormLifeVest(firstVest);
    setFormSpareAir(firstAir);
    setFormHelmetKit(firstHelmet);
    setFormNotlar('');
    setIsNewModalOpen(true);
  };

  // Save Zimmet
  const handleSaveForm = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (formDurum === 'GÖREV BÖLGESİNDE' && !formPersonel.trim()) {
      showNotification('⚠️ Lütfen personel adı ve soyadını belirtiniz.');
      return;
    }

    const effectiveLocation = formDurum === 'GÖREV BÖLGESİNDE' 
      ? (formGorevYeri === 'DİĞER' ? (formCustomLocation.trim().toUpperCase() || 'DİĞER') : formGorevYeri)
      : 'DEPODA';
    
    const personelAdi = formDurum === 'GÖREV BÖLGESİNDE' ? formPersonel.trim().toUpperCase() : '-';

    if (formDurum === 'GÖREV BÖLGESİNDE' && !editingRecord) {
      try {
        const res = await fetch('/api/yasam-destek/assign', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            personelAdi: personelAdi,
            lifeVestSn: formLifeVest.trim() || '-',
            spareAirSn: formSpareAir.trim() || '-',
            helmetKitSn: formHelmetKit.trim() || '-',
            gorevYeri: effectiveLocation,
            baslangicTarihi: formBaslangic.trim() || new Date().toLocaleDateString('tr-TR'),
            notlar: formNotlar.trim()
          })
        });
        const json = await res.json();
        if (json.status === 'success') {
          setIsNewModalOpen(false);
          showNotification(`✅ ${personelAdi} personeline zimmet verildi. Stoklar yeşil dolgu ile güncellendi.`);
          await onRefresh();
          await fetchAllStocks();
          return;
        }
      } catch (err: any) {
        console.warn('Assign API error:', err);
      }
    }

    const recordToSave: YasamDestekRecord = {
      id: editingRecord ? editingRecord.id : `yd_${Date.now()}`,
      siraNo: editingRecord ? editingRecord.siraNo : records.length + 1,
      personelAdi: personelAdi,
      gorevYeri: effectiveLocation,
      baslangicTarihi: formBaslangic.trim() || new Date().toLocaleDateString('tr-TR'),
      bitisTarihi: editingRecord?.bitisTarihi || '-',
      lifeVestSn: formLifeVest.trim() || '-',
      spareAirSn: formSpareAir.trim() || '-',
      helmetKitSn: formHelmetKit.trim() || '-',
      durum: editingRecord ? editingRecord.durum : (formDurum === 'GÖREV BÖLGESİNDE' ? 'GÖREVDE' : 'DEPODA'),
      devredilenPersonel: editingRecord?.devredilenPersonel || '-',
      teslimAlanPersonel: editingRecord?.teslimAlanPersonel || 'Depo Sorumlusu',
      notlar: formNotlar.trim(),
      guncellenmeTarihi: new Date().toLocaleDateString('tr-TR')
    };

    const success = await onSaveRecord(recordToSave);
    if (success) {
      setIsNewModalOpen(false);
      showNotification('✅ Yaşam Destek zimmet kaydı başarıyla oluşturuldu.');
      await fetchAllStocks();
    }
  };

  // Open Handover / Finish Modal
  const openFinishModal = (rec: YasamDestekRecord) => {
    setFinishTargetRecord(rec);
    const now = new Date();
    setFinishBitisTarihi(now.toLocaleDateString('tr-TR'));
    setFinishActionType('devret');
    setDevredilecekPersonel('');
    setDevredilecekGorevYeri(rec.gorevYeri || 'MİLAS');
    setDevredilecekCustomLocation('');
    setTeslimAlanGorevli('Ankara Depo Sorumlusu');
    setCheckLifeVest(true);
    setCheckSpareAir(true);
    setCheckHelmetKit(true);
    setFinishNotes('');
  };

  // Complete Handover / Return
  const handleCompleteFinish = async () => {
    if (!finishTargetRecord) return;

    if (finishActionType === 'devret' && !devredilecekPersonel.trim()) {
      showNotification('⚠️ Lütfen devredilecek personeli seçiniz.');
      return;
    }

    const effectiveDevirLocation = devredilecekGorevYeri === 'DİĞER' && devredilecekCustomLocation.trim()
      ? devredilecekCustomLocation.trim().toUpperCase()
      : (devredilecekGorevYeri || finishTargetRecord.gorevYeri || 'MİLAS');

    try {
      const res = await fetch('/api/yasam-destek/finish', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          zimmetId: finishTargetRecord.id,
          action: finishActionType,
          bitisTarihi: finishBitisTarihi,
          devredilenPersonel: devredilecekPersonel.trim().toUpperCase(),
          devredilenYer: effectiveDevirLocation,
          teslimAlanPersonel: teslimAlanGorevli.trim() || 'Depo Sorumlusu',
          notlar: finishNotes.trim()
        })
      });
      const json = await res.json();
      if (json.status === 'success') {
        setFinishTargetRecord(null);
        showNotification(finishActionType === 'devret' 
          ? `✅ Ekipmanlar ${devredilecekPersonel.trim().toUpperCase()} personeline başarıyla devredildi (yeşil dolgu).`
          : '✅ Ekipmanlar depoya teslim edildi ve stoklar beyaz dolguya çevrildi.');
        await onRefresh();
        await fetchAllStocks();
        return;
      }
    } catch (err: any) {
      console.warn('Finish API error:', err);
    }
  };

  // Open Swap Item Modal
  const openSwapModal = (rec: YasamDestekRecord) => {
    setSwapTargetRecord(rec);
    setSwapItemType('can_yelegi');
    const currentLv = (rec.lifeVestSn || '').trim();
    const availableVest = canYelegiList.find(c => {
      const a = (c.aciklamalar || '').toUpperCase().trim();
      const isDepot = c.durum === 'DEPODA' || c.durum === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
      return isDepot && c.durum !== 'KAYIT SİLME' && c.durum !== 'ONARIMDA' && c.durum !== 'BAKIMA GİDECEK' && c.seriNo !== currentLv;
    })?.seriNo || '';
    setSwapNewSn(availableVest);
    setSwapNotes('');
  };

  const handleSwapItemTypeChange = (type: 'can_yelegi' | 'spare_air' | 'helmet_kit') => {
    setSwapItemType(type);
    if (!swapTargetRecord) return;

    if (type === 'can_yelegi') {
      const currentLv = (swapTargetRecord.lifeVestSn || '').trim();
      const first = canYelegiList.find(c => {
        const a = (c.aciklamalar || '').toUpperCase().trim();
        const isDepot = c.durum === 'DEPODA' || c.durum === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
        return isDepot && c.durum !== 'KAYIT SİLME' && c.durum !== 'ONARIMDA' && c.durum !== 'BAKIMA GİDECEK' && c.seriNo !== currentLv;
      })?.seriNo || '';
      setSwapNewSn(first);
    } else if (type === 'spare_air') {
      const currentSa = (swapTargetRecord.spareAirSn || '').trim();
      const first = spareAirList.find(s => {
        const a = (s.aciklamalar || '').toUpperCase().trim();
        const isDepot = s.durum === 'DEPODA' || s.durum === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
        return isDepot && s.durum !== 'KAYIT SİLME' && s.durum !== 'ONARIMDA' && s.durum !== 'BAKIMA GİDECEK' && s.seriNo !== currentSa;
      })?.seriNo || '';
      setSwapNewSn(first);
    } else if (type === 'helmet_kit') {
      const currentHk = (swapTargetRecord.helmetKitSn || '').split(' (')[0].trim();
      const first = helmetKitList.find(h => {
        const a = (h.aciklamalar || '').toUpperCase().trim();
        const isDepot = h.durum === 'DEPODA' || h.durum === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
        const sn = (h.seriNo2 || h.seriNo1 || '').trim();
        return isDepot && h.durum !== 'KAYIT SİLME' && h.durum !== 'ONARIMDA' && h.durum !== 'BAKIMA GİDECEK' && sn !== currentHk && h.seriNo1 !== currentHk && h.seriNo2 !== currentHk;
      });
      const sn = first?.seriNo2 || first?.seriNo1 || '';
      setSwapNewSn(sn);
    }
  };

  const handleCompleteSwap = async () => {
    if (!swapTargetRecord || !swapNewSn) {
      showNotification('⚠️ Lütfen yeni seri numarasını seçiniz.');
      return;
    }

    let oldSn = '-';
    if (swapItemType === 'can_yelegi') oldSn = swapTargetRecord.lifeVestSn;
    else if (swapItemType === 'spare_air') oldSn = swapTargetRecord.spareAirSn;
    else if (swapItemType === 'helmet_kit') oldSn = swapTargetRecord.helmetKitSn;

    try {
      const res = await fetch('/api/yasam-destek/swap-item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          zimmetId: swapTargetRecord.id,
          personelAdi: swapTargetRecord.personelAdi,
          itemType: swapItemType,
          oldSn: oldSn,
          newSn: swapNewSn,
          notlar: swapNotes,
          tarih: new Date().toLocaleDateString('tr-TR')
        })
      });
      const json = await res.json();
      if (json.status === 'success') {
        setSwapTargetRecord(null);
        showNotification(`✅ ${swapTargetRecord.personelAdi} personeline ait ${swapItemType.toUpperCase()} başarıyla değiştirildi (${swapNewSn}). Eski parça depoya alındı.`);
        await onRefresh();
        await fetchAllStocks();
        return;
      }
    } catch (e: any) {
      showNotification('Hata: ' + e.message);
    }
  };

  // Export 4-Sheet Master Excel Workbook (Filtered & Styled with ExcelJS)
  const handleExportMasterExcel = async () => {
    try {
      const ExcelJS = await import('exceljs');
      const workbook = new ExcelJS.Workbook();
      
      const headerFill: any = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // slate-800
      const headerFont: any = { color: { argb: 'FFFFFFFF' }, bold: true };
      const borderStyle: any = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

      // Helper to apply common styles
      const styleSheet = (sheet: any, headers: string[]) => {
        const headerRow = sheet.getRow(1);
        headerRow.values = headers;
        headerRow.eachCell((cell: any) => {
          cell.fill = headerFill;
          cell.font = headerFont;
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = borderStyle;
        });
        sheet.columns.forEach((col: any) => {
          col.width = 20;
          col.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        });
      };

      // Sheet 1: CAN YELEĞİ STOK
      const ws1 = workbook.addWorksheet('CAN YELEĞİ STOK');
      const headers1 = ['S. NO', 'MALZEME ADI', 'PARÇA NO', 'SERİ NO', 'DIŞ NO', 'AÇIKLAMALAR', 'DURUM', 'TARİH'];
      styleSheet(ws1, headers1);
      filteredCanYelegi.forEach((c, i) => {
        const row = ws1.addRow([i + 1, c.malzemeAdi, c.parcaNo, c.seriNo, c.disNo || '-', c.aciklamalar, c.durum, c.tarih || '-']);
        const statusCell = row.getCell(7);
        if (c.durum === 'GÖREV BÖLGESİNDE') {
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
          statusCell.font = { color: { argb: 'FF000000' }, bold: true };
        }
        row.eachCell((cell: any) => { cell.border = borderStyle; });
      });

      // Sheet 2: SPARE AIR STOK
      const ws2 = workbook.addWorksheet('SPARE AIR STOK');
      const headers2 = ['S. NO', 'MALZEME ADI', 'PARÇA NO', 'SERİ NO', 'AÇIKLAMALAR', 'DURUM', 'TARİH'];
      styleSheet(ws2, headers2);
      filteredSpareAir.forEach((s, i) => {
        const row = ws2.addRow([i + 1, s.malzemeAdi, s.parcaNo, s.seriNo, s.aciklamalar, s.durum, s.tarih || '-']);
        const statusCell = row.getCell(6);
        if (s.durum === 'GÖREV BÖLGESİNDE') {
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
          statusCell.font = { color: { argb: 'FF000000' }, bold: true };
        }
        row.eachCell((cell: any) => { cell.border = borderStyle; });
      });

      // Sheet 3: HELMET KIT
      const ws3 = workbook.addWorksheet('HELMET KIT VE DAVID CLARK');
      const headers3 = ['S. NO', 'MALZEME ADI 1', 'PARÇA NO 1', 'SERİ NO 1', 'MALZEME ADI 2', 'PARÇA NO 2', 'SERİ NO 2', 'AÇIKLAMALAR', 'DURUM', 'TARİH'];
      styleSheet(ws3, headers3);
      filteredHelmetKit.forEach((h, i) => {
        const row = ws3.addRow([i + 1, h.malzemeAdi1, h.parcaNo1, h.seriNo1 || '-', h.malzemeAdi2, h.parcaNo2, h.seriNo2 || '-', h.aciklamalar, h.durum, h.tarih || '-']);
        const statusCell = row.getCell(9);
        if (h.durum === 'GÖREV BÖLGESİNDE') {
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
          statusCell.font = { color: { argb: 'FF000000' }, bold: true };
        }
        row.eachCell((cell: any) => { cell.border = borderStyle; });
      });

      // Sheet 4: GÖREV VE ZİMMETLER
      const ws4 = workbook.addWorksheet('GÖREV VE ZİMMETLER');
      const headers4 = ['SIRA NO', 'PERSONEL ADI SOYADI', 'GÖREV YERİ', 'BAŞLANGIÇ TARİHİ', 'BİTİŞ TARİHİ', 'LIFE VEST S/N', 'SPARE AIR S/N', 'HELMET KIT S/N', 'DURUM', 'DEVREDEN', 'TESLİM ALAN', 'NOTLAR'];
      styleSheet(ws4, headers4);
      filteredZimmetler.forEach((r, i) => {
        const row = ws4.addRow([i + 1, r.personelAdi, r.gorevYeri, r.baslangicTarihi, r.bitisTarihi || '-', r.lifeVestSn, r.spareAirSn, r.helmetKitSn, r.durum, r.devredilenPersonel || '-', r.teslimAlanPersonel || '-', r.notlar || '']);
        const statusCell = row.getCell(9);
        if (r.durum === 'GÖREVDE') {
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
          statusCell.font = { color: { argb: 'FF000000' }, bold: true };
        } else if (r.durum === 'DEVREDİLDİ') {
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } }; // amber-400
        }
        row.eachCell((cell: any) => { cell.border = borderStyle; });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `yasam_destek_at802_styled_${new Date().toISOString().slice(0, 10)}.xlsx`;
      link.click();

      showNotification('✅ 4 Sayfalı Stillendirilmiş Yaşam Destek Excel Dosyası Başarıyla İndirildi.');
    } catch (e: any) {
      showNotification('❌ Excel indirme hatası: ' + e.message);
    }
  };

  return (
    <div className="space-y-4 font-sans">
      
      {/* 1. RENK LEJANTI VE DURUM KODLARI (Tıklanabilir Filtreler) */}
      <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-xs">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-2.5">
          <div className="flex items-center gap-2">
            <span className="text-sm font-black text-slate-800 uppercase tracking-wider">
              🏷️ YAŞAM DESTEK DURUM VE RENK KODLARI (Filtrelemek için tıklayın)
            </span>
          </div>
          {filterStatus && filterStatus !== 'ALL' && (
            <button
              type="button"
              onClick={() => setFilterStatus('ALL')}
              className="px-3 py-1 bg-slate-900 text-white rounded-xl text-xs font-bold hover:bg-slate-800 transition cursor-pointer shadow-xs"
            >
              🔄 Filtreyi Kaldır ({filterStatus})
            </button>
          )}
        </div>

        {/* Renk Kartları (Tıklanabilir) */}
        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2 text-xs font-black">
          {STATUS_LEGEND.map((l, i) => {
            const isActive = filterStatus === l.label;
            return (
              <div 
                key={i} 
                onClick={() => setFilterStatus(isActive ? 'ALL' : l.label)}
                className={`p-2.5 rounded-xl border cursor-pointer transition-all ${l.border} ${l.bg} ${l.text} flex items-center justify-center text-center shadow-xs select-none hover:scale-[1.02] active:scale-95 ${isActive ? 'ring-4 ring-slate-900 shadow-lg scale-105 font-black underline' : ''}`}
                title="Tıklayarak bu durumda olan malzemeleri filtrele"
              >
                <span>{l.label} {isActive ? '✓' : ''}</span>
              </div>
            );
          })}
        </div>
      </div>

      {/* SEARCH BAR (KAYIT ARA) */}
      <div className="bg-white rounded-2xl p-3 border border-slate-200 shadow-xs flex flex-wrap items-center gap-3">
        <div className="relative flex-1 min-w-[240px]">
          <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="🔍 KAYIT ARA (Tarih, Personel Adı, Seri No, Görev Yeri)..."
            className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-emerald-600 focus:bg-white"
          />
        </div>
        {searchQuery && (
          <button
            onClick={() => setSearchQuery('')}
            className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl text-xs font-bold"
          >
            Temizle
          </button>
        )}
      </div>

      {/* 2. EXCEL SAYFA SEKMELERİ (Sheet Tabs) */}
      <div className="bg-slate-200/80 p-1.5 rounded-2xl flex flex-wrap items-center gap-1.5 border border-slate-300">
        <button
          type="button"
          onClick={() => setActiveSubTab('can_yelegi')}
          className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition flex items-center gap-2 cursor-pointer ${
            activeSubTab === 'can_yelegi'
              ? 'bg-[#0b3d1d] text-white shadow-md'
              : 'bg-white text-slate-700 hover:bg-slate-50'
          }`}
        >
          <span>🦺 CAN YELEĞİ STOK ({canYelegiList.length})</span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('spare_air')}
          className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition flex items-center gap-2 cursor-pointer ${
            activeSubTab === 'spare_air'
              ? 'bg-[#0b3d1d] text-white shadow-md'
              : 'bg-white text-slate-700 hover:bg-slate-50'
          }`}
        >
          <span>💨 SPARE AIR STOK ({spareAirList.length})</span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('helmet_kit')}
          className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition flex items-center gap-2 cursor-pointer ${
            activeSubTab === 'helmet_kit'
              ? 'bg-[#0b3d1d] text-white shadow-md'
              : 'bg-white text-slate-700 hover:bg-slate-50'
          }`}
        >
          <span>🪖 HELMET KIT VE DAVID CLARK KULAK VE DOLGU ({helmetKitList.length})</span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('personnel')}
          className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition flex items-center gap-2 cursor-pointer ${
            activeSubTab === 'personnel'
              ? 'bg-indigo-900 text-white shadow-md'
              : 'bg-indigo-100 text-indigo-950 hover:bg-indigo-200'
          }`}
        >
          <span>👨‍✈️ AT-802 PERSONEL VERİSİ ({personnelItems.length})</span>
        </button>

        <button
          type="button"
          onClick={() => setActiveSubTab('zimmetler')}
          className={`px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition flex items-center gap-2 cursor-pointer ml-auto ${
            activeSubTab === 'zimmetler'
              ? 'bg-amber-600 text-white shadow-md'
              : 'bg-amber-100 text-amber-900 hover:bg-amber-200'
          }`}
        >
          <span>📋 GÖREV & ZİMMETLER ({records.length})</span>
        </button>
      </div>

      {/* 3. AKTİF SAYFA İÇERİĞİ (HTML EXCEL TABLOSU) */}
      <div className="bg-white rounded-2xl border border-slate-300 shadow-sm overflow-hidden">
        
        {/* ========================================================================= */}
        {/* TAB 1: CAN YELEĞİ STOK TABLOSU */}
        {/* ========================================================================= */}
        {activeSubTab === 'can_yelegi' && (
          <div className="overflow-x-auto max-h-[65vh] overflow-y-auto">
            <div className="p-3 bg-slate-100 border-b border-slate-300 text-center font-black text-sm uppercase text-slate-800 tracking-wider">
              CAN YELEĞİ STOK ({filteredCanYelegi.length})
            </div>
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-[#b0b0b0] sticky top-0 z-20 border-b-2 border-slate-400 text-slate-900 font-black text-[11px] uppercase tracking-wider text-center select-none">
                <tr>
                  <th className="py-2.5 px-3 border-r border-slate-300 w-16">S. NO</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[180px]">MALZEME ADI</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px] font-mono">PARÇA NO</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px] font-mono">SERİ NO</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[100px] font-mono">DIŞ NO</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[160px] bg-slate-200">AÇIKLAMALAR</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[220px] bg-emerald-200/60 text-emerald-950">TESLİM EDİLEN PERSONEL/BÖLGE</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px]">TARİH / BİTİŞ</th>
                  <th className="py-2.5 px-3 min-w-[100px]">İŞLEMLER</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300 text-[11px] font-bold">
                {filteredCanYelegi.map((item, idx) => (
                  <tr key={idx} className={`border-b border-slate-300 ${getRowColorClass(item)} transition text-center`}>
                    <td className="py-2 px-2 border-r border-slate-300 font-mono">{item.sNo || idx + 1}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans">{item.malzemeAdi}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono">{item.parcaNo}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono font-black">{item.seriNo}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono font-black">{item.disNo || '-'}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans text-slate-700 font-medium">{item.notlar || item.aciklama || '-'}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans font-black text-slate-900">{item.aciklamalar || '-'}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono">{item.tarih || (item.durum === 'DEPODA' || item.durum === 'BOŞTA' ? '-' : 'DEVAM EDİYOR')}</td>
                    <td className="py-2 px-2">
                      <div className="flex flex-wrap items-center justify-center gap-1">
                        <button
                          type="button"
                          onClick={() => openStockEditModal('can_yelegi', item, idx)}
                          className="p-1 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-md text-xs cursor-pointer"
                          title="Düzenle"
                        >
                          <Edit2 className="w-3 h-3" />
                        </button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('can_yelegi', idx, 'ONARIMDA')} className="p-1 bg-yellow-300 hover:bg-yellow-400 text-slate-900 rounded-md text-[10px]" title="Onarıma Gönder">🛠️</button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('can_yelegi', idx, 'BAKIMA GİDECEK')} className="p-1 bg-cyan-300 hover:bg-cyan-400 text-slate-900 rounded-md text-[10px]" title="Bakıma Gönder">🔧</button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('can_yelegi', idx, 'KAYIT SİLME')} className="p-1 bg-red-400 hover:bg-red-500 text-white rounded-md text-[10px]" title="Kayıt Sil">🗑️</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* ========================================================================= */}
        {/* TAB 2: SPARE AIR STOK TABLOSU */}
        {/* ========================================================================= */}
        {activeSubTab === 'spare_air' && (
          <div className="overflow-x-auto max-h-[65vh] overflow-y-auto">
            <div className="p-3 bg-slate-100 border-b border-slate-300 text-center font-black text-sm uppercase text-slate-800 tracking-wider">
              SPARE AIR STOK ({filteredSpareAir.length})
            </div>
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-[#b0b0b0] sticky top-0 z-20 border-b-2 border-slate-400 text-slate-900 font-black text-[11px] uppercase tracking-wider text-center select-none">
                <tr>
                  <th className="py-2.5 px-3 border-r border-slate-300 w-16">S. NO</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[180px]">MALZEME ADI</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px] font-mono">PARÇA NO</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px] font-mono">SERİ NO</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[160px] bg-slate-200">AÇIKLAMALAR</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[220px] bg-cyan-200/60 text-cyan-950">TESLİM EDİLEN PERSONEL/BÖLGE</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px]">TARİH / BİTİŞ</th>
                  <th className="py-2.5 px-3 min-w-[100px]">İŞLEMLER</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300 text-[11px] font-bold">
                {filteredSpareAir.map((item, idx) => (
                  <tr key={idx} className={`border-b border-slate-300 ${getRowColorClass(item)} transition text-center`}>
                    <td className="py-2 px-2 border-r border-slate-300 font-mono">{item.sNo || idx + 1}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans">{item.malzemeAdi}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono">{item.parcaNo}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono font-black">{item.seriNo}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans text-slate-700 font-medium">{item.notlar || item.aciklama || '-'}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans font-black text-slate-900">{item.aciklamalar || '-'}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono">{item.tarih || (item.durum === 'DEPODA' || item.durum === 'BOŞTA' ? '-' : 'DEVAM EDİYOR')}</td>
                    <td className="py-2 px-2">
                      <div className="flex flex-wrap items-center justify-center gap-1">
                        <button
                          type="button"
                          onClick={() => openStockEditModal('spare_air', item, idx)}
                          className="p-1 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-md text-xs cursor-pointer"
                          title="Düzenle"
                        >
                          <Edit2 className="w-3 h-3" />
                        </button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('spare_air', idx, 'ONARIMDA')} className="p-1 bg-yellow-300 hover:bg-yellow-400 text-slate-900 rounded-md text-[10px]" title="Onarıma Gönder">🛠️</button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('spare_air', idx, 'BAKIMA GİDECEK')} className="p-1 bg-cyan-300 hover:bg-cyan-400 text-slate-900 rounded-md text-[10px]" title="Bakıma Gönder">🔧</button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('spare_air', idx, 'KAYIT SİLME')} className="p-1 bg-red-400 hover:bg-red-500 text-white rounded-md text-[10px]" title="Kayıt Sil">🗑️</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* ========================================================================= */}
        {/* TAB 3: HELMET KIT VE DAVID CLARK TABLOSU */}
        {/* ========================================================================= */}
        {activeSubTab === 'helmet_kit' && (
          <div className="overflow-x-auto max-h-[65vh] overflow-y-auto">
            <div className="p-3 bg-slate-100 border-b border-slate-300 text-center font-black text-sm uppercase text-slate-800 tracking-wider">
              HELMET KIT VE DAVID CLARK KULAK VE DOLGU ({filteredHelmetKit.length})
            </div>
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-[#b0b0b0] sticky top-0 z-20 border-b-2 border-slate-400 text-slate-900 font-black text-[11px] uppercase tracking-wider text-center select-none">
                <tr>
                  <th className="py-2.5 px-3 border-r border-slate-300 w-16">S. NO</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px]">MALZEME ADI 1</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px] font-mono">PARÇA NO 1</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px] font-mono">SERİ NO 1</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px]">MALZEME ADI 2</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px] font-mono">PARÇA NO 2</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px] font-mono text-indigo-950">KULAKLIK S/N</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[160px] bg-slate-200">AÇIKLAMALAR</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[200px] bg-purple-200/60 text-purple-950">TESLİM EDİLEN PERSONEL/BÖLGE</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px] font-mono">TARİH / BİTİŞ</th>
                  <th className="py-2.5 px-3 min-w-[100px]">İŞLEMLER</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300 text-[11px] font-bold">
                {filteredHelmetKit.map((item, idx) => (
                  <tr key={idx} className={`border-b border-slate-300 ${getRowColorClass(item)} transition text-center`}>
                    <td className="py-2 px-2 border-r border-slate-300 font-mono">{item.sNo || idx + 1}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-sans">{item.malzemeAdi1}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono">{item.parcaNo1}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono font-bold">{item.seriNo1 || '-'}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-sans">{item.malzemeAdi2}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono">{item.parcaNo2}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono font-black text-indigo-950">{item.seriNo2 || '-'}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans text-slate-700 font-medium">{item.notlar || item.aciklama || '-'}</td>
                    <td className="py-2 px-4 border-r border-slate-300 text-left font-sans font-black text-slate-900">{item.aciklamalar || '-'}</td>
                    <td className="py-2 px-3 border-r border-slate-300 font-mono">{item.tarih || (item.durum === 'DEPODA' || item.durum === 'BOŞTA' ? '-' : 'DEVAM EDİYOR')}</td>
                    <td className="py-2 px-2">
                      <div className="flex flex-wrap items-center justify-center gap-1">
                        <button
                          type="button"
                          onClick={() => openStockEditModal('helmet_kit', item, idx)}
                          className="p-1 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-md text-xs cursor-pointer"
                          title="Düzenle"
                        >
                          <Edit2 className="w-3 h-3" />
                        </button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('helmet_kit', idx, 'ONARIMDA')} className="p-1 bg-yellow-300 hover:bg-yellow-400 text-slate-900 rounded-md text-[10px]" title="Onarıma Gönder">🛠️</button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('helmet_kit', idx, 'BAKIMA GİDECEK')} className="p-1 bg-cyan-300 hover:bg-cyan-400 text-slate-900 rounded-md text-[10px]" title="Bakıma Gönder">🔧</button>
                        <button type="button" onClick={() => handleQuickChangeStockStatus('helmet_kit', idx, 'KAYIT SİLME')} className="p-1 bg-red-400 hover:bg-red-500 text-white rounded-md text-[10px]" title="Kayıt Sil">🗑️</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* ========================================================================= */}
        {/* TAB 4: GÖREV VE ZİMMETLER TABLOSU */}
        {/* ========================================================================= */}
        {activeSubTab === 'zimmetler' && (
          <div className="overflow-x-auto max-h-[65vh] overflow-y-auto">
            <div className="p-3 bg-amber-50 border-b border-amber-200 text-center font-black text-sm uppercase text-amber-950 tracking-wider">
              YAŞAM DESTEK CANLI GÖREV & ZİMMET LİSTESİ
            </div>
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-[#b0b0b0] sticky top-0 z-20 border-b-2 border-slate-400 text-slate-900 font-black text-[11px] uppercase tracking-wider text-center select-none">
                <tr>
                  <th className="py-2.5 px-3 border-r border-slate-300 w-16">SIRA NO</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[180px]">PERSONEL ADI SOYADI</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[110px]">GÖREV YERİ</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px]">BAŞLANGIÇ TARİHİ</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px]">BİTİŞ TARİHİ</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[130px] bg-amber-200/50">LIFE VEST S/N</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[130px] bg-cyan-200/50">SPARE AIR S/N</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[130px] bg-purple-200/50">HELMET KIT S/N</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[130px]">DURUM</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[160px] bg-slate-200">DEVREDEN / TESLİM EDEN</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[160px] bg-slate-200">DEVREDİLEN / TESLİM ALAN</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[160px] bg-slate-200">AÇIKLAMA / NOTLAR</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[150px]">GÜNCELLENME TARİHİ</th>
                  <th className="py-2.5 px-3 min-w-[140px] bg-slate-700 text-white font-black">İŞLEMLER</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-mono text-[11px]">
                {filteredZimmetler.length === 0 ? (
                  <tr>
                    <td colSpan={14} className="py-12 text-center text-slate-400 font-sans font-bold">
                      Aramanıza uygun kayıt bulunamadı veya henüz kayıtlı zimmet yok.
                    </td>
                  </tr>
                ) : (
                  filteredZimmetler.map((r, idx) => (
                    <tr key={r.id || idx} className="hover:bg-slate-50 transition border-b border-slate-200 text-center">
                      <td className="py-2.5 px-2 font-bold text-slate-500 border-r border-slate-200">{idx + 1}</td>
                      <td className="py-2.5 px-4 font-bold text-slate-900 border-r border-slate-200 text-left font-sans">{r.personelAdi}</td>
                      <td className="py-2.5 px-3 border-r border-slate-200 font-sans">
                        <span className="px-2 py-0.5 rounded bg-slate-100 font-bold border border-slate-300 text-[10px]">📍 {r.gorevYeri}</span>
                      </td>
                      <td className="py-2.5 px-3 border-r border-slate-200 text-slate-600">{r.baslangicTarihi}</td>
                      <td className="py-2.5 px-3 border-r border-slate-200 text-slate-600 font-bold text-emerald-800">{r.bitisTarihi || 'DEVAM EDİYOR'}</td>
                      <td className="py-2.5 px-3 border-r border-slate-200 font-bold text-amber-900 bg-amber-50/50">{r.lifeVestSn}</td>
                      <td className="py-2.5 px-3 border-r border-slate-200 font-bold text-cyan-900 bg-cyan-50/50">{r.spareAirSn}</td>
                      <td className="py-2.5 px-3 border-r border-slate-200 font-bold text-purple-900 bg-purple-50/50">{r.helmetKitSn}</td>
                      <td className="py-2.5 px-3 border-r border-slate-200 font-sans">
                        <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase ${
                          r.durum === 'GÖREVDE' ? 'bg-emerald-100 text-emerald-900 border border-emerald-300' :
                          r.durum === 'DEVREDİLDİ' ? 'bg-amber-100 text-amber-900 border border-amber-300' :
                          'bg-blue-100 text-blue-900 border border-blue-300'
                        }`}>
                          {r.durum}
                        </span>
                      </td>
                      <td className="py-2.5 px-4 border-r border-slate-200 text-left font-sans text-xs text-slate-700 font-medium">{r.devredilenPersonel || '-'}</td>
                      <td className="py-2.5 px-4 border-r border-slate-200 text-left font-sans text-xs text-slate-700 font-medium">{r.teslimAlanPersonel || '-'}</td>
                      <td className="py-2.5 px-4 border-r border-slate-200 text-left font-sans text-xs text-slate-700 font-medium">{r.notlar || '-'}</td>
                      <td className="py-2.5 px-3 border-r border-slate-200 font-bold text-slate-500">{r.guncellenmeTarihi || '-'}</td>
                      <td className="py-2.5 px-2 bg-slate-50/50">
                        <div className="flex items-center justify-center gap-1.5">
                          {r.durum === 'GÖREVDE' && (
                            <button
                              type="button"
                              onClick={() => openFinishModal(r)}
                              className="px-2 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-lg text-[10px] font-black uppercase flex items-center gap-1 cursor-pointer transition shadow-2xs"
                            >
                              <ArrowRightLeft className="w-3 h-3" />
                              <span>Devir</span>
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() => {
                              setPrintRecord(r);
                              setTimeout(() => window.print(), 300);
                            }}
                            className="p-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-xs transition cursor-pointer"
                            title="Tutanak Yazdır"
                          >
                            <Printer className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={async () => {
                              if (window.confirm('Bu kaydı silmek istediğinize emin misiniz?')) {
                                await onDeleteRecord(r.id);
                              }
                            }}
                            className="p-1.5 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded-lg text-xs transition cursor-pointer"
                            title="Sil"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* ========================================================================= */}
        {/* TAB 5: AT-802 PERSONEL VERİSİ (TEKNİSYENLER & PİLOTLAR) */}
        {/* ========================================================================= */}
        {activeSubTab === 'personnel' && (
          <div className="overflow-x-auto max-h-[65vh] overflow-y-auto">
            <div className="p-3 bg-indigo-50 border-b border-indigo-200 flex flex-wrap items-center justify-between gap-3">
              <div className="font-black text-sm uppercase text-indigo-950 tracking-wider">
                👨‍✈️ AT-802 PERSONEL VERİSİ ( CANLI LİSTE) ({personnelItems.length})
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setPersonnelRoleFilter('ALL')}
                  className={`px-3 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    personnelRoleFilter === 'ALL' ? 'bg-indigo-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  TÜMÜ
                </button>
                <button
                  type="button"
                  onClick={() => setPersonnelRoleFilter('Teknisyen')}
                  className={`px-3 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    personnelRoleFilter === 'Teknisyen' ? 'bg-indigo-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  🛠️ TEKNİSYENLER
                </button>
                <button
                  type="button"
                  onClick={() => setPersonnelRoleFilter('Pilot')}
                  className={`px-3 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    personnelRoleFilter === 'Pilot' ? 'bg-indigo-900 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  ✈️ PİLOTLAR
                </button>
                <button
                  type="button"
                  onClick={async () => {
                    const name = window.prompt('Yeni Personel Adı Soyadı:');
                    if (!name || !name.trim()) return;
                    const role = window.prompt('Unvanı (Teknisyen / Pilot):', 'Teknisyen');
                    if (!role) return;

                    await fetch('/api/yasam-destek/personnel-list', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ action: 'add', item: { adSoyad: name.trim().toUpperCase(), unvan: role.trim() } })
                    });
                    showNotification('✅ Personel başarıyla eklendi.');
                    fetchPersonnel();
                  }}
                  className="px-3 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>Personel Ekle</span>
                </button>
              </div>
            </div>
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-[#b0b0b0] sticky top-0 z-20 border-b-2 border-slate-400 text-slate-900 font-black text-[11px] uppercase tracking-wider text-center select-none">
                <tr>
                  <th className="py-2.5 px-3 border-r border-slate-300 w-16">SIRA</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[200px]">PERSONEL ADI SOYADI</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px]">ÜNVANI</th>
                  <th className="py-2.5 px-3 border-r border-slate-300 min-w-[120px]">GÖREV YERİ</th>
                  <th className="py-2.5 px-4 border-r border-slate-300 text-left min-w-[300px]">ZİMMETLİ EKİPMANLAR (SERİ NO)</th>
                  <th className="py-2.5 px-3 min-w-[140px]">İŞLEMLER</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-sans text-[11px]">
                {personnelItems
                  .filter(p => {
                    const matchesRole = personnelRoleFilter === 'ALL' || (p.unvan && p.unvan.toLowerCase() === personnelRoleFilter.toLowerCase());
                    const matchesQuery = !searchQuery.trim() || p.adSoyad.toLowerCase().includes(searchQuery.toLowerCase()) || (p.unvan || '').toLowerCase().includes(searchQuery.toLowerCase());
                    return matchesRole && matchesQuery;
                  })
                  .map((p, idx) => {
                    const rawZimmet = records.find(r => r.personelAdi.toUpperCase() === p.adSoyad.toUpperCase() && r.durum === 'GÖREVDE');
                    const hasActiveItems = rawZimmet && (
                      (rawZimmet.lifeVestSn && rawZimmet.lifeVestSn !== '-') ||
                      (rawZimmet.spareAirSn && rawZimmet.spareAirSn !== '-') ||
                      (rawZimmet.helmetKitSn && rawZimmet.helmetKitSn !== '-')
                    );
                    const activeZimmet = hasActiveItems ? rawZimmet : null;
                    return (
                      <tr key={p.id || idx} className="hover:bg-slate-50 transition border-b border-slate-200 text-center">
                        <td className="py-3 px-2 font-bold text-slate-500 border-r border-slate-200 font-mono">{idx + 1}</td>
                        <td className="py-3 px-4 font-black text-slate-900 border-r border-slate-200 text-left font-sans">
                          <span className="text-xs text-slate-950 font-black">{p.adSoyad}</span>
                        </td>
                        <td className="py-3 px-3 border-r border-slate-200 text-center font-sans font-bold">
                          <span className={`px-2.5 py-1 rounded-full text-xs ${
                            p.unvan === 'Teknisyen' ? 'bg-amber-100 text-amber-950 border border-amber-300' :
                            p.unvan === 'Pilot' ? 'bg-blue-100 text-blue-950 border border-blue-300' : 'bg-slate-100 text-slate-800 border border-slate-300'
                          }`}>
                            {p.unvan === 'Teknisyen' ? '🛠️ Teknisyen' : (p.unvan === 'Pilot' ? '✈️ Pilot' : '👤 ' + (p.unvan || 'Personel'))}
                          </span>
                        </td>
                        <td className="py-3 px-3 border-r border-slate-200 font-sans text-center">
                          {activeZimmet ? (
                            <span className="px-2.5 py-1 rounded-lg bg-emerald-100 text-emerald-950 font-black border border-emerald-300 text-xs">
                              📍 {activeZimmet.gorevYeri || 'MİLAS'}
                            </span>
                          ) : (
                            <span className="text-slate-300 select-none">-</span>
                          )}
                        </td>
                        <td className="py-3 px-4 border-r border-slate-200 text-left font-sans">
                          {activeZimmet ? (
                            <div className="flex flex-wrap items-center gap-1.5">
                              <span className="inline-flex items-center gap-1 text-[11px] bg-amber-100 text-amber-950 px-2.5 py-0.5 rounded-full border border-amber-300 font-mono font-black">
                                🦺 LV: {activeZimmet.lifeVestSn}
                              </span>
                              <span className="inline-flex items-center gap-1 text-[11px] bg-cyan-100 text-cyan-950 px-2.5 py-0.5 rounded-full border border-cyan-300 font-mono font-black">
                                💨 SA: {activeZimmet.spareAirSn}
                              </span>
                              <span className="inline-flex items-center gap-1 text-[11px] bg-purple-100 text-purple-950 px-2.5 py-0.5 rounded-full border border-purple-300 font-mono font-black">
                                🪖 HK: {activeZimmet.helmetKitSn}
                              </span>
                            </div>
                          ) : (
                            <span className="text-slate-400 font-bold text-xs italic">
                              aktif zimmet yok
                            </span>
                          )}
                        </td>
                        <td className="py-3 px-3 text-center">
                          <div className="flex items-center justify-center gap-1.5 flex-wrap">
                            {activeZimmet ? (
                              <>
                                <button
                                  type="button"
                                  onClick={() => openFinishModal(activeZimmet)}
                                  className="px-3 py-1.5 bg-amber-600 hover:bg-amber-700 active:scale-95 text-white rounded-xl text-xs font-black uppercase tracking-wider shadow-xs transition cursor-pointer flex items-center gap-1.5"
                                  title="Zimmeti devret veya depoya teslim et"
                                >
                                  <ArrowRightLeft className="w-3.5 h-3.5" />
                                  <span>Devret</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => openSwapModal(activeZimmet)}
                                  className="px-3 py-1.5 bg-indigo-700 hover:bg-indigo-800 active:scale-95 text-white rounded-xl text-xs font-black uppercase tracking-wider shadow-xs transition cursor-pointer flex items-center gap-1.5"
                                  title="Görevdeyken parça veya ekipman değiştir"
                                >
                                  <RefreshCw className="w-3.5 h-3.5" />
                                  <span>Değiştir</span>
                                </button>
                              </>
                            ) : (
                              <button
                                type="button"
                                onClick={() => {
                                  openNewModal();
                                  setFormPersonel(p.adSoyad);
                                  setFormDurum('GÖREV BÖLGESİNDE');
                                }}
                                className="px-4 py-1.5 bg-emerald-700 hover:bg-emerald-800 active:scale-95 text-white rounded-xl text-xs font-black uppercase tracking-wider shadow-xs transition cursor-pointer flex items-center gap-1.5 mx-auto"
                                title="Bu personele yeni zimmet ver"
                              >
                                <Plus className="w-3.5 h-3.5" />
                                <span>Zimmet Al</span>
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
        )}

      </div>

      {/* ========================================================================= */}
      {/* 4. MODAL: YENİ ZİMMET FORMU (GERÇEK STOK SERİ NO SEÇİMLERİ İLE) */}
      {/* ========================================================================= */}
      {isNewModalOpen && (
        <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="px-5 py-3.5 bg-[#0b3d1d] text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="text-xl">🚀</span>
                <h3 className="text-sm font-black uppercase tracking-wider">🚀 YENİ YAŞAM DESTEK ZİMMETİ VER</h3>
              </div>
              <button onClick={() => setIsNewModalOpen(false)} className="text-white/80 hover:text-white cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveForm} className="p-5 space-y-4">
              {/* Durum Seçimi */}
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Durum:</label>
                <select
                  value={formDurum}
                  onChange={(e) => setFormDurum(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 text-xs"
                >
                  <option value="DEPODA">DEPODA</option>
                  <option value="GÖREV BÖLGESİNDE">GÖREV BÖLGESİNDE</option>
                </select>
              </div>

              {formDurum === 'GÖREV BÖLGESİNDE' && (
                <>
                  {/* Personel Seçimi */}
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Personel Adı Soyadı: <span className="text-rose-500">*</span>
                    </label>
                    <select
                      required
                      value={formPersonel}
                      onChange={(e) => setFormPersonel(e.target.value)}
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 text-xs focus:outline-none focus:border-emerald-600 focus:bg-white"
                    >
                      <option value="">Personel Seçiniz...</option>
                      {personnelItems.map((p, i) => (
                        <option key={i} value={p.adSoyad}>
                          {p.adSoyad} ({p.unvan || 'Teknisyen'})
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Görev Yeri */}
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Görev Yeri:</label>
                    <select
                      value={formGorevYeri}
                      onChange={(e) => setFormGorevYeri(e.target.value)}
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 text-xs cursor-pointer"
                    >
                      <option value="MİLAS">MİLAS</option>
                      <option value="KARAİN">KARAİN</option>
                      <option value="ÇANAKKALE">ÇANAKKALE</option>
                      <option value="BURSA">BURSA</option>
                      <option value="ANKARA">ANKARA</option>
                      <option value="İZMİR">İZMİR</option>
                      <option value="MUĞLA">MUĞLA</option>
                      <option value="ANTALYA">ANTALYA</option>
                      <option value="DİĞER">✏️ DİĞER (Manuel Yazınız)</option>
                    </select>
                  </div>
                </>
              )}

              {formGorevYeri === 'DİĞER' && (
                <div>
                  <label className="block text-xs font-bold text-amber-800 mb-1">Özel Görev Yeri (Manuel Yazınız):</label>
                  <input
                    type="text"
                    value={formCustomLocation}
                    onChange={(e) => setFormCustomLocation(e.target.value)}
                    placeholder="Örn: ADANA, DALAMAN, SİNOP..."
                    className="w-full px-3 py-2 bg-amber-50 border border-amber-300 rounded-xl font-bold text-slate-900 text-xs uppercase"
                  />
                </div>
              )}

              {/* EKİPMAN SERİ NUMARALARI (GERÇEK STOKLARDAN SEÇİLEBİLİR) */}
              <div className="space-y-3 pt-2 border-t border-slate-100">
                <div className="text-xs font-black uppercase tracking-wider text-slate-800">
                  Zimmetlenecek Ekipmanlar ve Gerçek Stok Seri Noları
                </div>

                {/* Life Vest */}
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold text-slate-700">🦺 Life Vest (Can Yeleği) Seri No:</label>
                  </div>
                  <input
                    type="text"
                    required
                    value={formLifeVest}
                    onChange={(e) => setFormLifeVest(e.target.value)}
                    placeholder="Seri No girin..."
                    className="w-full px-3 py-1.5 bg-amber-50/50 border border-amber-300 rounded-xl font-mono font-bold text-slate-900 text-xs"
                  />
                </div>

                {/* Spare Air */}
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold text-slate-700">💨 Spare Air (Yedek Hava) Seri No:</label>
                  </div>
                  <input
                    type="text"
                    required
                    value={formSpareAir}
                    onChange={(e) => setFormSpareAir(e.target.value)}
                    placeholder="Seri No girin..."
                    className="w-full px-3 py-1.5 bg-cyan-50/50 border border-cyan-300 rounded-xl font-mono font-bold text-slate-900 text-xs"
                  />
                </div>

                {/* Helmet Kit */}
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold text-slate-700">🪖 Helmet Kit / David Clark Seri No:</label>
                  </div>
                  <input
                    type="text"
                    required
                    value={formHelmetKit}
                    onChange={(e) => setFormHelmetKit(e.target.value)}
                    placeholder="Seri No girin..."
                    className="w-full px-3 py-1.5 bg-purple-50/50 border border-purple-300 rounded-xl font-mono font-bold text-slate-900 text-xs"
                  />
                </div>

                {/* Helmet Kit */}
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold text-slate-700">🪖 Helmet Kit & David Clark:</label>
                    <span className="text-[10px] text-emerald-700 font-bold">{helmetKitList.length} Adet Stokta</span>
                  </div>
                  <div className="flex gap-2">
                    <select
                      onChange={(e) => setFormHelmetKit(e.target.value)}
                      className="px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:outline-none"
                    >
                      <option value="">Stoktan Seçin...</option>
                      {helmetKitList.map((h, i) => (
                        <option key={i} value={`${h.aciklamalar} (18852G-01 / H-10-13X)`}>
                          {h.aciklamalar} {h.tarih ? '(' + h.tarih + ')' : ''}
                        </option>
                      ))}
                    </select>
                    <input
                      type="text"
                      required
                      value={formHelmetKit}
                      onChange={(e) => setFormHelmetKit(e.target.value)}
                      placeholder="Tanım / Seri No..."
                      list="panelHelmetKitDatalist"
                      className="flex-1 px-3 py-1.5 bg-purple-50/50 border border-purple-300 rounded-xl font-bold text-slate-900 text-xs"
                    />
                    <datalist id="panelHelmetKitDatalist">
                      {helmetKitList.map((h, idx) => (
                        <option key={idx} value={`${h.aciklamalar} (18852G-01 / H-10-13X)`} />
                      ))}
                    </datalist>
                  </div>
                </div>
              </div>

              {/* Notlar */}
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Açıklama / Görev Notu:</label>
                <textarea
                  rows={2}
                  value={formNotlar}
                  onChange={(e) => setFormNotlar(e.target.value)}
                  placeholder="İntikal veya nöbet detayı..."
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs"
                />
              </div>

              <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIsNewModalOpen(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-[#0b3d1d] hover:bg-[#072813] text-white rounded-xl text-xs font-black uppercase tracking-wider cursor-pointer shadow-md"
                >
                  Zimmeti Kaydet ve Göreve Başlat
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 5. MODAL: GÖREVİ BİTİR / DEVİR & TESLİM */}
      {/* ========================================================================= */}
      {finishTargetRecord && (
        <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="px-5 py-3.5 bg-amber-600 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="text-xl">🔄</span>
                <h3 className="text-sm font-black uppercase tracking-wider">Görevi Bitir / Ekipman Devir & Teslim</h3>
              </div>
              <button onClick={() => setFinishTargetRecord(null)} className="text-white/80 hover:text-white cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 space-y-4">
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 text-xs">
                <div className="font-bold text-slate-400 uppercase text-[10px]">Görevi Tamamlayan:</div>
                <div className="text-sm font-black text-slate-900 mt-0.5">{finishTargetRecord.personelAdi}</div>
                <div className="text-slate-600 font-mono text-[11px] mt-0.5">
                  Mevcut Görev Yeri: <strong>{finishTargetRecord.gorevYeri}</strong> | Başlangıç: {finishTargetRecord.baslangicTarihi}
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Bitiş / Teslim Tarihi:</label>
                <input
                  type="text"
                  required
                  value={finishBitisTarihi}
                  onChange={(e) => setFinishBitisTarihi(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-mono text-slate-900 text-xs font-bold"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-2">İşlem Türünü Seçiniz:</label>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    type="button"
                    onClick={() => setFinishActionType('devret')}
                    className={`p-3 rounded-xl border-2 text-left transition cursor-pointer ${
                      finishActionType === 'devret'
                        ? 'border-indigo-600 bg-indigo-50/50 text-indigo-950 font-bold shadow-xs'
                        : 'border-slate-200 bg-white text-slate-600'
                    }`}
                  >
                    <div className="text-base">🤝</div>
                    <div className="text-xs font-black uppercase mt-1">Ekipmanları Devret</div>
                    <div className="text-[10px] text-slate-500 mt-0.5">Yeni personele aktar</div>
                  </button>

                  <button
                    type="button"
                    onClick={() => setFinishActionType('teslim')}
                    className={`p-3 rounded-xl border-2 text-left transition cursor-pointer ${
                      finishActionType === 'teslim'
                        ? 'border-blue-600 bg-blue-50/50 text-blue-950 font-bold shadow-xs'
                        : 'border-slate-200 bg-white text-slate-600'
                    }`}
                  >
                    <div className="text-base">🏢</div>
                    <div className="text-xs font-black uppercase mt-1">Ankara Depo Teslim</div>
                    <div className="text-[10px] text-slate-500 mt-0.5">Depoya iade et</div>
                  </button>
                </div>
              </div>

              {finishActionType === 'devret' ? (
                <div className="space-y-3 p-3 bg-indigo-50/40 rounded-xl border border-indigo-200">
                  <div>
                    <label className="block text-xs font-bold text-indigo-950 mb-1">
                      Devredilecek Personel: <span className="text-rose-500">*</span>
                    </label>
                    <select
                      required
                      value={devredilecekPersonel}
                      onChange={(e) => setDevredilecekPersonel(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-indigo-300 rounded-xl font-bold text-slate-900 text-xs"
                    >
                      <option value="">Devredilecek Personeli Seçiniz...</option>
                      {personnelItems
                        .filter(p => p.adSoyad.toUpperCase() !== finishTargetRecord.personelAdi.toUpperCase())
                        .map((p, i) => (
                          <option key={i} value={p.adSoyad}>
                            {p.adSoyad} ({p.unvan || 'Teknisyen'})
                          </option>
                        ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-indigo-950 mb-1">Yeni Görev Yeri:</label>
                    <select
                      value={devredilecekGorevYeri}
                      onChange={(e) => setDevredilecekGorevYeri(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-indigo-300 rounded-xl font-bold text-slate-900 text-xs cursor-pointer"
                    >
                      <option value="MİLAS">MİLAS</option>
                      <option value="KARAİN">KARAİN</option>
                      <option value="ÇANAKKALE">ÇANAKKALE</option>
                      <option value="BURSA">BURSA</option>
                      <option value="ANKARA">ANKARA</option>
                      <option value="İZMİR">İZMİR</option>
                      <option value="MUĞLA">MUĞLA</option>
                      <option value="ANTALYA">ANTALYA</option>
                      <option value="DİĞER">✏️ DİĞER (Manuel Yazınız)</option>
                    </select>
                    {devredilecekGorevYeri === 'DİĞER' && (
                      <input
                        type="text"
                        value={devredilecekCustomLocation}
                        onChange={(e) => setDevredilecekCustomLocation(e.target.value)}
                        placeholder="Örn: ADANA, DALAMAN, SİNOP..."
                        className="mt-2 w-full px-3 py-1.5 bg-amber-50 border-2 border-amber-400 rounded-xl font-bold text-slate-900 text-xs uppercase"
                      />
                    )}
                  </div>
                </div>
              ) : (
                <div className="p-3 bg-blue-50/40 rounded-xl border border-blue-200">
                  <label className="block text-xs font-bold text-blue-950 mb-1">Teslim Alan Depo Sorumlusu:</label>
                  <input
                    type="text"
                    value={teslimAlanGorevli}
                    onChange={(e) => setTeslimAlanGorevli(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-blue-300 rounded-xl font-bold text-slate-900 text-xs"
                  />
                </div>
              )}

              {/* Checklist */}
              <div>
                <div className="text-xs font-black uppercase text-slate-800 mb-2">Devredilecek Ekipmanlar:</div>
                <div className="space-y-2 text-xs">
                  <label className="flex items-center gap-2 p-2 bg-slate-50 rounded-lg border border-slate-200 cursor-pointer">
                    <input type="checkbox" checked={checkLifeVest} onChange={() => setCheckLifeVest(!checkLifeVest)} className="rounded" />
                    <span>Life Vest S/N: <strong className="font-mono text-amber-800">{finishTargetRecord.lifeVestSn}</strong></span>
                  </label>
                  <label className="flex items-center gap-2 p-2 bg-slate-50 rounded-lg border border-slate-200 cursor-pointer">
                    <input type="checkbox" checked={checkSpareAir} onChange={() => setCheckSpareAir(!checkSpareAir)} className="rounded" />
                    <span>Spare Air S/N: <strong className="font-mono text-cyan-800">{finishTargetRecord.spareAirSn}</strong></span>
                  </label>
                  <label className="flex items-center gap-2 p-2 bg-slate-50 rounded-lg border border-slate-200 cursor-pointer">
                    <input type="checkbox" checked={checkHelmetKit} onChange={() => setCheckHelmetKit(!checkHelmetKit)} className="rounded" />
                    <span>Helmet Kit: <strong className="font-mono text-purple-800">{finishTargetRecord.helmetKitSn}</strong></span>
                  </label>
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Not:</label>
                <input
                  type="text"
                  value={finishNotes}
                  onChange={(e) => setFinishNotes(e.target.value)}
                  placeholder="Ekipmanlar sağlam teslim edildi..."
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs"
                />
              </div>

              <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setFinishTargetRecord(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold cursor-pointer"
                >
                  Vazgeç
                </button>
                <button
                  type="button"
                  onClick={handleCompleteFinish}
                  className="px-5 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-xs font-black uppercase tracking-wider cursor-pointer shadow-md"
                >
                  {finishActionType === 'devret' ? 'Görevi Başlat ve Devret' : 'Depoya Teslim Al'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 5.5 MODAL: GÖREVDEYKEN PARÇA / EKİPMAN DEĞİŞTİR */}
      {/* ========================================================================= */}
      {swapTargetRecord && (
        <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="px-5 py-3.5 bg-indigo-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <RefreshCw className="w-4 h-4 text-emerald-300" />
                <h3 className="text-sm font-black uppercase tracking-wider">Görevdeyken Parça / Ekipman Değiştir</h3>
              </div>
              <button onClick={() => setSwapTargetRecord(null)} className="text-white/80 hover:text-white cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 space-y-4">
              <div className="p-3.5 bg-indigo-50 rounded-2xl border border-indigo-200 text-xs space-y-1">
                <div className="text-[10px] font-black uppercase text-indigo-900 tracking-wider">Personel ve Görev Bilgisi:</div>
                <div className="flex items-center justify-between">
                  <span className="text-sm font-black text-slate-900">{swapTargetRecord.personelAdi}</span>
                  <span className="px-2.5 py-0.5 bg-emerald-100 text-emerald-950 font-black rounded-lg border border-emerald-300 text-[11px]">
                    📍 {swapTargetRecord.gorevYeri || 'MİLAS'}
                  </span>
                </div>
                <div className="pt-2 mt-1 border-t border-indigo-100 grid grid-cols-3 gap-1.5 text-[10px] text-slate-600">
                  <div className="p-1 bg-white rounded border border-slate-200">🦺 LV: <strong className="font-mono text-amber-900">{swapTargetRecord.lifeVestSn || '-'}</strong></div>
                  <div className="p-1 bg-white rounded border border-slate-200">💨 SA: <strong className="font-mono text-cyan-900">{swapTargetRecord.spareAirSn || '-'}</strong></div>
                  <div className="p-1 bg-white rounded border border-slate-200">🪖 HK: <strong className="font-mono text-purple-900">{swapTargetRecord.helmetKitSn || '-'}</strong></div>
                </div>
              </div>

              <div>
                <label className="block text-xs font-black uppercase text-slate-700 mb-1.5">Değiştirilecek Ekipmanı Seçiniz: *</label>
                <select
                  value={swapItemType}
                  onChange={(e) => handleSwapItemTypeChange(e.target.value as any)}
                  className="w-full px-3.5 py-2.5 bg-white border-2 border-indigo-500 rounded-xl font-black text-xs text-slate-900 cursor-pointer shadow-2xs"
                >
                  <option value="can_yelegi">🦺 Can Yeleği (Life Vest - P/N: S-7200-511)</option>
                  <option value="spare_air">💨 Spare Air (Yedek Hava Tüpü - P/N: 175-001-CE)</option>
                  <option value="helmet_kit">🪖 Helmet Kit / David Clark Kulaklık</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-black uppercase text-slate-700 mb-1.5">
                  Depoda Müsait Yeni Seri Numarası: *
                </label>
                <select
                  value={swapNewSn}
                  onChange={(e) => setSwapNewSn(e.target.value)}
                  className="w-full px-3.5 py-2.5 bg-emerald-50 border-2 border-emerald-600 rounded-xl font-mono font-black text-xs text-emerald-950 cursor-pointer shadow-2xs"
                >
                  {swapItemType === 'can_yelegi' && (() => {
                    const currentLv = (swapTargetRecord?.lifeVestSn || '').trim();
                    const available = canYelegiList.filter(c => {
                      const a = (c.aciklamalar || '').toUpperCase().trim();
                      const isDepot = c.durum === 'DEPODA' || c.durum === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
                      return isDepot && c.durum !== 'KAYIT SİLME' && c.durum !== 'ONARIMDA' && c.durum !== 'BAKIMA GİDECEK' && c.seriNo !== currentLv;
                    });
                    if (available.length === 0) {
                      return <option value="">⚠️ Depoda müsait başka Can Yeleği bulunamadı</option>;
                    }
                    return available.map((c, i) => (
                      <option key={i} value={c.seriNo}>
                        {c.seriNo} ({c.disNo && c.disNo !== '-' ? c.disNo + ' - ' : ''}DEPODA/MÜSAİT)
                      </option>
                    ));
                  })()}

                  {swapItemType === 'spare_air' && (() => {
                    const currentSa = (swapTargetRecord?.spareAirSn || '').trim();
                    const available = spareAirList.filter(s => {
                      const a = (s.aciklamalar || '').toUpperCase().trim();
                      const isDepot = s.durum === 'DEPODA' || s.durum === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
                      return isDepot && s.durum !== 'KAYIT SİLME' && s.durum !== 'ONARIMDA' && s.durum !== 'BAKIMA GİDECEK' && s.seriNo !== currentSa;
                    });
                    if (available.length === 0) {
                      return <option value="">⚠️ Depoda müsait başka Spare Air bulunamadı</option>;
                    }
                    return available.map((s, i) => (
                      <option key={i} value={s.seriNo}>
                        {s.seriNo} (DEPODA/MÜSAİT)
                      </option>
                    ));
                  })()}

                  {swapItemType === 'helmet_kit' && (() => {
                    const currentHk = (swapTargetRecord?.helmetKitSn || '').split(' (')[0].trim();
                    const available = helmetKitList.filter(h => {
                      const a = (h.aciklamalar || '').toUpperCase().trim();
                      const isDepot = h.durum === 'DEPODA' || h.durum === 'BOŞTA' || a === 'DEPODA' || a === 'BOŞTA' || a === '-' || !a;
                      const sn = (h.seriNo2 || h.seriNo1 || '').trim();
                      return isDepot && h.durum !== 'KAYIT SİLME' && h.durum !== 'ONARIMDA' && h.durum !== 'BAKIMA GİDECEK' && sn !== currentHk && h.seriNo1 !== currentHk && h.seriNo2 !== currentHk;
                    });
                    if (available.length === 0) {
                      return <option value="">⚠️ Depoda müsait başka Helmet Kit bulunamadı</option>;
                    }
                    return available.map((h, i) => {
                      const sn = h.seriNo2 || h.seriNo1 || `H-10-13X-${i + 1}`;
                      const name = h.malzemeAdi2 || h.malzemeAdi1 || 'David Clark';
                      return (
                        <option key={i} value={sn}>
                          {sn} ({name} - DEPODA/MÜSAİT)
                        </option>
                      );
                    });
                  })()}
                </select>
                <p className="text-[10px] text-slate-500 mt-1">
                  * Seçilen yeni parça yeşil dolguya döner, personeldeki eski parça ise otomatik olarak depoya iade edilir (beyaz dolgu).
                </p>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Değişim Nedeni / Açıklama:</label>
                <input
                  type="text"
                  value={swapNotes}
                  onChange={(e) => setSwapNotes(e.target.value)}
                  placeholder="Örn: Arıza / Dolum değişimi, Sezonluk muayene..."
                  className="w-full px-3.5 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-medium"
                />
              </div>

              <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setSwapTargetRecord(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="button"
                  onClick={handleCompleteSwap}
                  className="px-5 py-2 bg-indigo-800 hover:bg-indigo-900 text-white rounded-xl text-xs font-black uppercase tracking-wider cursor-pointer shadow-md flex items-center gap-1.5"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Parçayı Değiştir ve Depoyu Güncelle</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* PRINT TUTANAK */}
      {printRecord && (
        <div className="hidden print:block fixed inset-0 z-[99999] bg-white p-8 font-sans">
          <div className="text-center pb-4 border-b-2 border-slate-800">
            <h2 className="text-lg font-black uppercase">T.C. TARIM VE ORMAN BAKANLIĞI</h2>
            <h3 className="text-sm font-bold uppercase">ORMAN GENEL MÜDÜRLÜĞÜ - HAVACILIK DAİRESİ BAŞKANLIĞI</h3>
            <h4 className="text-xs font-black uppercase mt-1 bg-slate-800 text-white py-1">
              AT-802 YAŞAM DESTEK EKİPMANLARI ZİMMET VE TESLİM TUTANAĞI
            </h4>
          </div>

          <div className="grid grid-cols-2 gap-4 my-6 text-xs">
            <div className="border border-slate-400 p-3 rounded">
              <div className="font-bold text-slate-500 mb-1">ZİMMET EDİLEN PERSONEL:</div>
              <div className="text-sm font-black">{printRecord.personelAdi}</div>
              <div className="mt-1">Görev Yeri: <strong>{printRecord.gorevYeri}</strong></div>
              <div>Başlangıç Tarihi: <strong>{printRecord.baslangicTarihi}</strong></div>
              <div>Bitiş Tarihi: <strong>{printRecord.bitisTarihi || 'DEVAM EDİYOR'}</strong></div>
            </div>

            <div className="border border-slate-400 p-3 rounded">
              <div className="font-bold text-slate-500 mb-1">DURUM VE İŞLEM BİLGİSİ:</div>
              <div className="text-sm font-black text-emerald-800">{printRecord.durum}</div>
              <div className="mt-1">Devredilen Personel: <strong>{printRecord.devredilenPersonel || '-'}</strong></div>
              <div>Teslim Alan: <strong>{printRecord.teslimAlanPersonel || '-'}</strong></div>
              <div>Açıklama: {printRecord.notlar || '-'}</div>
            </div>
          </div>

          <table className="w-full text-xs border border-slate-400 border-collapse mb-8">
            <thead>
              <tr className="bg-slate-200">
                <th className="border border-slate-400 p-2 text-center w-12">NO</th>
                <th className="border border-slate-400 p-2 text-left">EKİPMAN TANIMI</th>
                <th className="border border-slate-400 p-2 text-center w-40">SERİ NUMARASI</th>
                <th className="border border-slate-400 p-2 text-center w-28">DURUMU</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="border border-slate-400 p-2 text-center font-bold">1</td>
                <td className="border border-slate-400 p-2 font-bold">Life Vest (Can Yeleği)</td>
                <td className="border border-slate-400 p-2 text-center font-mono font-bold">{printRecord.lifeVestSn}</td>
                <td className="border border-slate-400 p-2 text-center">FAAL / SAĞLAM</td>
              </tr>
              <tr>
                <td className="border border-slate-400 p-2 text-center font-bold">2</td>
                <td className="border border-slate-400 p-2 font-bold">Spare Air (Yedek Hava Tüpü)</td>
                <td className="border border-slate-400 p-2 text-center font-mono font-bold">{printRecord.spareAirSn}</td>
                <td className="border border-slate-400 p-2 text-center">FAAL / SAĞLAM</td>
              </tr>
              <tr>
                <td className="border border-slate-400 p-2 text-center font-bold">3</td>
                <td className="border border-slate-400 p-2 font-bold">Helmet Kit (Kask ve Kulaklık Kiti)</td>
                <td className="border border-slate-400 p-2 text-center font-mono font-bold">{printRecord.helmetKitSn}</td>
                <td className="border border-slate-400 p-2 text-center">FAAL / SAĞLAM</td>
              </tr>
            </tbody>
          </table>

          <div className="grid grid-cols-2 gap-12 text-center text-xs mt-16">
            <div>
              <div className="font-bold text-slate-500 uppercase">TESLİM EDEN</div>
              <div className="font-black text-sm mt-1">{printRecord.personelAdi}</div>
              <div className="text-slate-400 text-[10px] mt-8">İmza: .......................................</div>
            </div>
            <div>
              <div className="font-bold text-slate-500 uppercase">TESLİM ALAN / DEVİR ALAN</div>
              <div className="font-black text-sm mt-1">
                {printRecord.devredilenPersonel && printRecord.devredilenPersonel !== '-' 
                  ? printRecord.devredilenPersonel 
                  : (printRecord.teslimAlanPersonel || 'Depo Sorumlusu')}
              </div>
              <div className="text-slate-400 text-[10px] mt-8">İmza: .......................................</div>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* 6. MODAL: STOK / ZİMMET SATIRI DÜZENLEME & DOLGU RENK SEÇİMİ */}
      {/* ========================================================================= */}
      {editingStockItem && (
        <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="px-5 py-3.5 bg-[#0b3d1d] text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Edit2 className="w-4 h-4 text-emerald-300" />
                <h3 className="text-sm font-black uppercase tracking-wider">
                  Stok & Teslim Yeri Düzenle
                </h3>
              </div>
              <button onClick={() => setEditingStockItem(null)} className="text-white/80 hover:text-white cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveStockEdit} className="p-5 space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Durum & Dolgu Rengi:</label>
                <select
                  value={stockEditStatus}
                  onChange={(e) => setStockEditStatus(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 text-xs"
                >
                  <option value="DEPODA">DEPODA (Dolgusuz / Beyaz - Depo Stoğunda)</option>
                  <option value="GÖREV BÖLGESİNDE">GÖREV BÖLGESİNDE (Yeşil - Görevde / Zimmetli)</option>
                  <option value="ONARIMDA">ONARIMDA (Sarı - Onarım Görmüş / İşlemde)</option>
                  <option value="BAKIMA GİDECEK">BAKIMA GİDECEK (Mavi - Dolum / Muayene)</option>
                  <option value="KAYIT SİLME">KAYIT SİLME (Kırmızı - Iskartaya Ayrıldı)</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Durum & Dolgu Rengi:</label>
                <select
                  value={stockEditStatus}
                  onChange={(e) => setStockEditStatus(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 text-xs"
                >
                  <option value="DEPODA">DEPODA (Dolgusuz / Beyaz - Depo Stoğunda)</option>
                  <option value="GÖREV BÖLGESİNDE">GÖREV BÖLGESİNDE (Yeşil - Görevde / Zimmetli)</option>
                  <option value="ONARIMDA">ONARIMDA (Sarı - Onarım Görmüş / İşlemde)</option>
                  <option value="BAKIMA GİDECEK">BAKIMA GİDECEK (Mavi - Dolum / Muayene)</option>
                  <option value="KAYIT SİLME">KAYIT SİLME (Kırmızı - Iskartaya Ayrıldı)</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Seri No / Parça kiti:</label>
                <input
                  type="text"
                  value={stockEditSeriNo}
                  onChange={(e) => setStockEditSeriNo(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-mono text-slate-900 text-xs font-bold"
                />
              </div>

              {editingStockItem.type === 'helmet_kit' && (
                <div>
                  <label className="block text-xs font-bold text-indigo-900 mb-1">
                    David Clark Kulaklık Seri No (H-10-13X):
                  </label>
                  <input
                    type="text"
                    value={stockEditSeriNo2}
                    onChange={(e) => setStockEditSeriNo2(e.target.value)}
                    placeholder="Örn: H-10-13X-1"
                    className="w-full px-3 py-2 bg-indigo-50 border border-indigo-300 rounded-xl font-mono font-black text-indigo-950 text-xs"
                  />
                </div>
              )}

              {stockEditStatus === 'GÖREV BÖLGESİNDE' && (
                <>
                    <div>
                        <label className="block text-xs font-bold text-slate-700 mb-1">
                          Teslim Edilen Personel:
                        </label>
                        <input
                          type="text"
                          required
                          list="panelPersonnelDatalist"
                          value={stockEditPersonel}
                          onChange={(e) => setStockEditPersonel(e.target.value)}
                          placeholder="Personel Adı..."
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 text-xs focus:outline-none focus:border-emerald-600 focus:bg-white"
                        />
                        <datalist id="panelPersonnelDatalist">
                          {personnelList.map((p, i) => (
                            <option key={i} value={p} />
                          ))}
                        </datalist>
                    </div>

                    <div>
                        <label className="block text-xs font-bold text-slate-700 mb-1">Görev Bölgesi:</label>
                        <select
                          value={stockEditGorevYeri}
                          onChange={(e) => setStockEditGorevYeri(e.target.value)}
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-bold text-slate-900 text-xs"
                        >
                          <option value="MİLAS">MİLAS</option>
                          <option value="KARAİN">KARAİN</option>
                          <option value="ÇANAKKALE">ÇANAKKALE</option>
                          <option value="BURSA">BURSA</option>
                          <option value="ANKARA">ANKARA</option>
                        </select>
                    </div>
                </>
              )}

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Teslim / Görev Tarihi:</label>
                <input
                  type="text"
                  value={stockEditTarih}
                  onChange={(e) => setStockEditTarih(e.target.value)}
                  placeholder="DD.MM.YYYY veya DEVAM EDİYOR"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl font-mono text-slate-900 text-xs font-bold"
                />
              </div>

              <div className="flex items-center justify-end gap-2.5 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setEditingStockItem(null)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold cursor-pointer"
                >
                  İptal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-[#0b3d1d] hover:bg-[#072813] text-white rounded-xl text-xs font-black uppercase tracking-wider cursor-pointer shadow-md"
                >
                  Kaydet
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 8. MODAL: PERSONEL ZİMMET DETAYLARI GÖRÜNTÜLE */}
      {viewingPersonnel && (() => {
        const rawZimmet = records.find(r => r.personelAdi.toUpperCase() === viewingPersonnel.toUpperCase() && r.durum === 'GÖREVDE');
        const hasActiveItems = rawZimmet && (
          (rawZimmet.lifeVestSn && rawZimmet.lifeVestSn !== '-') ||
          (rawZimmet.spareAirSn && rawZimmet.spareAirSn !== '-') ||
          (rawZimmet.helmetKitSn && rawZimmet.helmetKitSn !== '-')
        );
        const activeZimmet = hasActiveItems ? rawZimmet : null;
        const pastZimmets = records.filter(r => r.personelAdi.toUpperCase() === viewingPersonnel.toUpperCase() && r.durum !== 'GÖREVDE');
        return (
          <div className="fixed inset-0 z-[10000] bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-white rounded-3xl shadow-2xl max-w-lg w-full border border-slate-200 overflow-hidden animate-in fade-in duration-150">
              <div className="px-5 py-4 bg-[#0b3d1d] text-white flex items-center justify-between">
                <h3 className="text-sm font-black uppercase tracking-wider flex items-center gap-2">
                  <Eye className="w-4 h-4 text-emerald-300" />
                  <span>Personel Yaşam Destek Zimmet Detayı</span>
                </h3>
                <button
                  type="button"
                  onClick={() => setViewingPersonnel(null)}
                  className="text-white/80 hover:text-white cursor-pointer"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-6 space-y-4">
                {activeZimmet ? (
                  <>
                    <div className="p-4 bg-emerald-50 rounded-2xl border-2 border-emerald-300">
                      <div className="flex items-center justify-between pb-2 border-b border-emerald-200">
                        <div>
                          <div className="text-xs font-bold text-emerald-800 uppercase tracking-wider">Aktif Zimmet Sahibi</div>
                          <div className="text-lg font-black text-slate-900">{activeZimmet.personelAdi}</div>
                        </div>
                        <span className="px-3 py-1 rounded-full text-xs font-black bg-emerald-600 text-white shadow-sm">
                          ✅ GÖREVDE
                        </span>
                      </div>
                      <div className="grid grid-cols-2 gap-3 mt-3 text-xs">
                        <div>
                          <span className="text-slate-500 font-medium block">Görev Yeri:</span>
                          <span className="font-black text-slate-900 text-sm">📍 {activeZimmet.gorevYeri || 'MİLAS'}</span>
                        </div>
                        <div>
                          <span className="text-slate-500 font-medium block">Başlangıç Tarihi:</span>
                          <span className="font-bold text-slate-900">📅 {activeZimmet.baslangicTarihi}</span>
                        </div>
                      </div>
                    </div>

                    <div className="space-y-2.5 mt-4">
                      <div className="text-xs font-black uppercase text-slate-700 tracking-wider flex items-center justify-between">
                        <span>Teslim Edilen Yaşam Destek Ekipmanları:</span>
                        <span className="text-[10px] text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">🟢 Yeşil Dolgu</span>
                      </div>

                      <div className="p-3 bg-amber-50/70 rounded-xl border border-amber-200 flex items-center justify-between">
                        <div className="flex items-center gap-2.5">
                          <span className="text-xl">🦺</span>
                          <div>
                            <div className="text-[10px] font-bold text-amber-800 uppercase">Life West - Can Yeleği</div>
                            <div className="text-xs font-mono font-black text-amber-950">Seri No: {activeZimmet.lifeVestSn || 'Yok'}</div>
                          </div>
                        </div>
                        <span className="text-[10px] font-bold px-2 py-0.5 bg-amber-200 text-amber-900 rounded-full font-mono">P/N: S-7200-511</span>
                      </div>

                      <div className="p-3 bg-cyan-50/70 rounded-xl border border-cyan-200 flex items-center justify-between">
                        <div className="flex items-center gap-2.5">
                          <span className="text-xl">💨</span>
                          <div>
                            <div className="text-[10px] font-bold text-cyan-800 uppercase">Spare Air - Yedek Hava Tüpü</div>
                            <div className="text-xs font-mono font-black text-cyan-950">Seri No: {activeZimmet.spareAirSn || 'Yok'}</div>
                          </div>
                        </div>
                        <span className="text-[10px] font-bold px-2 py-0.5 bg-cyan-200 text-cyan-900 rounded-full font-mono">P/N: 175-001-CE</span>
                      </div>

                      <div className="p-3 bg-purple-50/70 rounded-xl border border-purple-200 flex items-center justify-between">
                        <div className="flex items-center gap-2.5">
                          <span className="text-xl">🪖</span>
                          <div>
                            <div className="text-[10px] font-bold text-purple-800 uppercase">Helmet Kit & David Clark Kulaklık</div>
                            <div className="text-xs font-mono font-black text-purple-950">Seri No: {activeZimmet.helmetKitSn || 'Yok'}</div>
                          </div>
                        </div>
                        <span className="text-[10px] font-bold px-2 py-0.5 bg-purple-200 text-purple-900 rounded-full font-mono">18852G-01 / H-10-13X</span>
                      </div>

                      {activeZimmet.notlar && (
                        <div className="p-2.5 bg-slate-50 rounded-xl border border-slate-200 text-xs mt-2">
                          <span className="font-bold text-slate-600 block text-[10px] uppercase">Görev Notu:</span>
                          <span className="text-slate-800 font-medium">{activeZimmet.notlar}</span>
                        </div>
                      )}
                    </div>

                    <div className="flex items-center justify-between gap-2 pt-4 border-t border-slate-200 mt-4">
                      <button
                        type="button"
                        onClick={() => {
                          setPrintRecord(activeZimmet);
                          setTimeout(() => window.print(), 200);
                        }}
                        className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl text-xs font-bold flex items-center gap-1.5 cursor-pointer"
                      >
                        <Printer className="w-3.5 h-3.5" />
                        <span>Tutanak Yazdır</span>
                      </button>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setViewingPersonnel(null)}
                          className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold cursor-pointer"
                        >
                          Kapat
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setViewingPersonnel(null);
                            openFinishModal(activeZimmet);
                          }}
                          className="px-5 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-xs font-black uppercase tracking-wider shadow-md cursor-pointer flex items-center gap-1.5 active:scale-95"
                        >
                          <span>🔄 Zimmeti Bitir / Devret</span>
                        </button>
                      </div>
                    </div>
                  </>
                ) : (
                  <div className="p-6 text-center">
                    <div className="w-16 h-16 mx-auto mb-3 rounded-full bg-slate-100 flex items-center justify-center text-3xl">
                      📦
                    </div>
                    <h4 className="text-base font-black text-slate-900">{viewingPersonnel}</h4>
                    <p className="text-xs text-slate-500 mt-1">Bu personele ait aktif bir yaşam destek zimmeti bulunmamaktadır. Ekipmanlar depodadır (dolgusuz/beyaz).</p>

                    {pastZimmets.length > 0 && (
                      <div className="mt-4 p-3 bg-slate-50 rounded-xl border border-slate-200 text-left">
                        <div className="text-[10px] font-bold text-slate-500 uppercase mb-1">Geçmiş Zimmet Kayıtları ({pastZimmets.length}):</div>
                        <div className="text-xs text-slate-700 font-mono">
                          Son görev: {pastZimmets[0].gorevYeri || '-'} ({pastZimmets[0].durum}) - Bitiş: {pastZimmets[0].bitisTarihi || '-'}
                        </div>
                      </div>
                    )}

                    <div className="mt-6 flex items-center justify-center gap-3">
                      <button
                        type="button"
                        onClick={() => setViewingPersonnel(null)}
                        className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold cursor-pointer"
                      >
                        Kapat
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setViewingPersonnel(null);
                          openNewModal();
                          setFormPersonel(viewingPersonnel);
                          setFormDurum('GÖREV BÖLGESİNDE');
                        }}
                        className="px-5 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-xl text-xs font-black uppercase tracking-wider shadow-md cursor-pointer flex items-center gap-1.5 active:scale-95"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>Bu Personele Zimmet Ver</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        );
      })()}

    </div>
  );
};
