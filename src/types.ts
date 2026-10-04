export interface AuditLogEntry {
  id: string;
  timestamp: string; // GG.AA.YYYY SS:DD
  unit: string;
  itemName: string;
  pn: string;
  fieldName: string;
  oldValue: string;
  newValue: string;
  author?: string;
  action?: string;
}

export type DepoTransactionType = 'GİRİŞ' | 'ÇIKIŞ' | 'TRANSFER' | 'BAKIM_GÖNDERİM' | 'KİT OLUŞTURMA' | 'GIRIS' | 'CIKIS' | 'BAKIM' | 'SAYIM DÜZELTME';

export interface DepoTransaction {
  id: string;
  timestamp: string; // GG.AA.YYYY SS:DD:SS
  date?: string; // GG.AA.YYYY (Excel formatı)
  type: DepoTransactionType | string;
  itemName: string;
  name?: string;
  pn: string;
  sn: string;
  itemDesc?: string;
  partNumber?: string;
  serialNumber?: string;
  adet?: number;
  category?: string;
  depoYeri?: string;
  islemTuru?: string;
  kuyrukKodu?: string;
  teslimAlan?: string;
  kabulYapan?: string;
  sourceLocation?: string;
  targetLocation?: string;
  sourceLoc?: string;
  destinationLoc?: string;
  tailNo?: string;
  workOrderNo?: string;
  quantity: number;
  unit?: string;
  operator?: string; // TESLİM ALAN
  receivedBy?: string; // KABUL YAPAN
  location?: string; // DEPO YERİ
  sheetName?: string; // SAYFA ADI (Örn: DEPO HAREKET GEÇMİŞİ-AT-802)
  notes?: string;
  kitComponentsData?: KitComponent[];
  kitCount?: number;
  isUndone?: boolean;
  isNewSessionTx?: boolean;
  sistemStok?: number;
  sayilanAdet?: number;
  fark?: number;
  isEdit?: boolean;
}

export interface KitComponent {
  pn: string;
  itemName: string;
  quantityPerKit: number;
  totalNeeded: number;
  currentStock: number;
}

export interface DepoKitDefinition {
  id: string;
  kitName: string;
  kitPn: string;
  quantity: number;
  location: string;
  createdAt: string;
  items: KitComponent[];
}

export interface DepoCertificateRecord {
  id: string;
  fileName: string;
  uploadDate: string;
  matchedKeywords: string[];
  fileDataUrl?: string;
}

export interface VehicleDocument {
  id: string;
  plate: string;
  vehiclePlate?: string;
  docType: 'ruhsat' | 'muayene' | 'sigorta_kasko' | 'diger' | string;
  title: string;
  fileName: string;
  fileSize?: string;
  uploadDate: string;
  fileDataUrl?: string;
}
