/**
 * Cross-System Communication Bridge for Hangar Portal and Standalone Sub-Systems
 * (Yer Destek & Özel Aletler, Hangar Depo & Yedek Parça, Kara Araçları Takip)
 */

export const HANGAR_CHANNEL_NAME = 'hangar_portal_cross_system_channel_v1';

export type StandaloneSystemType = 'yer-destek' | 'depo' | 'kara-araclari';

export interface CrossSystemEvent {
  type: 
    | 'DEPOT_SEARCH_REQUEST' 
    | 'DEPOT_SEARCH_RESPONSE' 
    | 'SYSTEM_DATA_UPDATED' 
    | 'DEPO_ITEM_TRANSFERRED'
    | 'KARA_ARACI_UPDATED'
    | 'REFRESH_ALL'
    | 'NAVIGATE_TO_ITEM' 
    | 'GUN_TAKIP_PING'
    | 'GUN_TAKIP_PONG';
  sender?: 'portal' | 'yer-destek' | 'depo' | 'kara-araclari';
  sourceSystem?: string;
  payload?: any;
  timestamp: number;
}

export interface DepoSearchResultItem {
  id?: string;
  name: string;
  partNo: string;
  serialNo?: string;
  quantity: string;
  location: string;
  category: string;
  unitKey?: string;
  status?: string;
  details?: string;
}

// Global broadcast channel instance with fallback
let channelInstance: BroadcastChannel | null = null;
try {
  if (typeof window !== 'undefined' && 'BroadcastChannel' in window) {
    channelInstance = new BroadcastChannel(HANGAR_CHANNEL_NAME);
  }
} catch (e) {
  console.warn('BroadcastChannel not supported or restricted, falling back to storage events:', e);
}

/**
 * Broadcast an event to all open tabs and frames
 */
export function publishCrossSystemEvent(type: CrossSystemEvent['type'], sender: CrossSystemEvent['sender'], payload?: any): void {
  const event: CrossSystemEvent = {
    type,
    sender,
    payload,
    timestamp: Date.now()
  };

  if (channelInstance) {
    try {
      channelInstance.postMessage(event);
    } catch (e) {
      console.warn('BroadcastChannel postMessage error:', e);
    }
  }

  // Also dispatch window storage event fallback for maximum reliability
  try {
    localStorage.setItem('hangar_cross_sys_event_bus', JSON.stringify(event));
  } catch (e) {
    // ignore
  }

  // Also dispatch local window event
  if (typeof window !== 'undefined') {
    window.dispatchEvent(new CustomEvent('hangar_cross_system_event', { detail: event }));
  }
}

/**
 * Subscribe to cross-system events
 */
export function subscribeCrossSystemEvents(handler: (event: CrossSystemEvent) => void): () => void {
  if (typeof window === 'undefined') return () => {};

  const handleBcMessage = (ev: MessageEvent) => {
    if (ev.data && ev.data.type) {
      handler(ev.data as CrossSystemEvent);
    }
  };

  const handleCustomEvent = (ev: Event) => {
    const custom = ev as CustomEvent<CrossSystemEvent>;
    if (custom.detail) {
      handler(custom.detail);
    }
  };

  const handleStorage = (ev: StorageEvent) => {
    if (ev.key === 'hangar_cross_sys_event_bus' && ev.newValue) {
      try {
        const parsed = JSON.parse(ev.newValue);
        if (parsed && parsed.type) {
          handler(parsed);
        }
      } catch (e) {
        // ignore
      }
    }
  };

  if (channelInstance) {
    channelInstance.addEventListener('message', handleBcMessage);
  }
  window.addEventListener('hangar_cross_system_event', handleCustomEvent);
  window.addEventListener('storage', handleStorage);

  return () => {
    if (channelInstance) {
      channelInstance.removeEventListener('message', handleBcMessage);
    }
    window.removeEventListener('hangar_cross_system_event', handleCustomEvent);
    window.removeEventListener('storage', handleStorage);
  };
}

/**
 * Open one of the 3 standalone index HTML pages with optional deep-link params
 */
export function openStandaloneSystem(
  system: StandaloneSystemType,
  options?: {
    target?: '_blank' | '_self';
    params?: Record<string, string>;
    search?: string;
    fleet?: string;
    unit?: string;
  }
): Window | null {
  const target = options?.target || '_blank';
  let path = '';

  switch (system) {
    case 'yer-destek':
      path = '/yer-destek.html';
      break;
    case 'depo':
      path = '/depo.html';
      break;
    case 'kara-araclari':
      path = '/kara-araclari.html';
      break;
  }

  const query = new URLSearchParams();
  if (options?.params) {
    for (const [k, v] of Object.entries(options.params)) {
      if (v !== undefined && v !== null && v !== '') {
        query.set(k, String(v));
      }
    }
  }
  if (options?.search) {
    query.set('search', options.search);
  }
  if (options?.fleet) {
    query.set('fleet', options.fleet);
  }
  if (options?.unit) {
    query.set('unit', options.unit);
  }

  const qStr = query.toString();
  if (qStr) {
    path += (path.includes('?') ? '&' : '?') + qStr;
  }

  return window.open(path, target);
}

/**
 * Searches across Hangar Depo spare parts in localStorage (or via live cross-system query)
 * Can be called from anywhere (e.g. Technical Publications, Gün Takip, or Main Portal)
 */
export function searchDepoSparePartsSync(query: string): DepoSearchResultItem[] {
  if (!query || !query.trim()) return [];
  const cleanQ = query.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  const rawQ = query.trim().toLowerCase();

  const results: DepoSearchResultItem[] = [];

  // 1. Search in excel_techizat_hangar_data (Main Hangar Warehouse Inventory)
  try {
    const hangarRaw = localStorage.getItem('excel_techizat_hangar_data');
    if (hangarRaw) {
      const rows: string[][] = JSON.parse(hangarRaw);
      if (Array.isArray(rows)) {
        rows.forEach((row, idx) => {
          const name = String(row[1] || '').trim();
          const pn = String(row[2] || '').trim();
          const sn = String(row[3] || '').trim();
          const qty = String(row[4] || '').trim();
          const loc = String(row[5] || '').trim();
          const status = String(row[6] || '').trim();
          const section = String(row[12] || row[13] || row[11] || 'Sarf ve Parça Depo').trim();

          const nameClean = name.toLowerCase().replace(/[^a-z0-9]/g, '');
          const pnClean = pn.toLowerCase().replace(/[^a-z0-9]/g, '');
          const snClean = sn.toLowerCase().replace(/[^a-z0-9]/g, '');

          if (
            nameClean.includes(cleanQ) ||
            pnClean.includes(cleanQ) ||
            snClean.includes(cleanQ) ||
            name.toLowerCase().includes(rawQ) ||
            pn.toLowerCase().includes(rawQ)
          ) {
            results.push({
              id: `hangar_${idx}`,
              name: name || 'Belirtilmemiş Malzeme',
              partNo: pn || '-',
              serialNo: sn || '-',
              quantity: qty || '0',
              location: loc || 'Hangar Depo',
              category: section.toUpperCase().includes('KİMYASAL') ? 'KİMYASAL DEPO' : 'SARF VE PARÇA DEPO',
              unitKey: 'hangar',
              status: status || 'FAAL',
              details: `Sıra #${row[0] || idx + 1}`
            });
          }
        });
      }
    }
  } catch (e) {
    console.warn('Depo sync search parse error for hangar data:', e);
  }

  // 2. Also search other aircraft fleet spare parts and special tools if relevant
  const units = ['at802', 'bell429', 't70', 'b360', 'c650'];
  units.forEach(u => {
    try {
      const raw = localStorage.getItem(`excel_techizat_${u}_data`);
      if (raw) {
        const rows: string[][] = JSON.parse(raw);
        if (Array.isArray(rows)) {
          rows.forEach((row, idx) => {
            const name = String(row[1] || '').trim();
            const pn = String(row[2] || '').trim();
            const sn = String(row[3] || '').trim();
            const qty = String(row[4] || '').trim();
            const loc = String(row[5] || '').trim();
            const status = String(row[6] || '').trim();

            const nameClean = name.toLowerCase().replace(/[^a-z0-9]/g, '');
            const pnClean = pn.toLowerCase().replace(/[^a-z0-9]/g, '');

            if (
              (cleanQ.length >= 3 && (pnClean.includes(cleanQ) || nameClean.includes(cleanQ))) ||
              pn.toLowerCase().includes(rawQ)
            ) {
              results.push({
                id: `${u}_${idx}`,
                name: name || 'Özel Alet / Parça',
                partNo: pn || '-',
                serialNo: sn || '-',
                quantity: qty || '1',
                location: loc || `${u.toUpperCase()} Atölyesi`,
                category: `${u.toUpperCase()} FİLOSU`,
                unitKey: u,
                status: status || 'FAAL',
                details: `${u.toUpperCase()} Yer Destek / Takımhane`
              });
            }
          });
        }
      }
    } catch (e) {
      // ignore
    }
  });

  return results;
}

/**
 * Calculates live item counts and warnings across all 3 systems
 */
export function getCrossSystemCounts(): {
  yerDestekCount: number;
  yerDestekUrgentCount: number;
  depoCount: number;
  depoCriticalCount: number;
  karaAraclariCount: number;
  karaAraclariUrgentCount: number;
} {
  let yerDestekCount = 0;
  let yerDestekUrgentCount = 0;
  let depoCount = 0;
  let depoCriticalCount = 0;
  let karaAraclariCount = 0;
  let karaAraclariUrgentCount = 0;

  if (typeof window === 'undefined') {
    return {
      yerDestekCount: 0,
      yerDestekUrgentCount: 0,
      depoCount: 0,
      depoCriticalCount: 0,
      karaAraclariCount: 0,
      karaAraclariUrgentCount: 0
    };
  }

  // 1. Yer Destek counts
  const units = ['at802', 'bell429', 't70', 't70_bumbi_backet', 't70_helitak', 'b360', 'c650'];
  units.forEach(u => {
    try {
      const raw = localStorage.getItem(`excel_techizat_${u}_data`);
      if (raw) {
        const rows: string[][] = JSON.parse(raw);
        if (Array.isArray(rows)) {
          yerDestekCount += rows.length;
          // check calibration/warning
          rows.forEach(r => {
            const nextDate = String(r[9] || '').trim();
            if (nextDate && isDateApproachingOrPast(nextDate, 90)) {
              yerDestekUrgentCount++;
            }
          });
        }
      }
    } catch (e) {}
  });

  // 2. Depo counts
  try {
    const raw = localStorage.getItem('excel_techizat_hangar_data');
    if (raw) {
      const rows: string[][] = JSON.parse(raw);
      if (Array.isArray(rows)) {
        depoCount = rows.length;
        rows.forEach(r => {
          const qtyStr = String(r[4] || '0').replace(/[^0-9]/g, '');
          const qty = parseInt(qtyStr, 10) || 0;
          if (qty <= 2) {
            depoCriticalCount++;
          }
        });
      }
    }
  } catch (e) {}

  // 3. Kara Araçları counts
  try {
    const raw = localStorage.getItem('excel_techizat_kara_araclari_data') || localStorage.getItem('techizat_kara_araclari_data');
    if (raw) {
      const rows: string[][] = JSON.parse(raw);
      if (Array.isArray(rows)) {
        karaAraclariCount = rows.length;
        rows.forEach(r => {
          const nextBakim = String(r[10] || '').trim();
          if (nextBakim && isDateApproachingOrPast(nextBakim, 90)) {
            karaAraclariUrgentCount++;
          }
        });
      }
    }
  } catch (e) {}

  return {
    yerDestekCount,
    yerDestekUrgentCount,
    depoCount,
    depoCriticalCount,
    karaAraclariCount,
    karaAraclariUrgentCount
  };
}

/**
 * Checks if a date (DD.MM.YYYY) is past or within specified days
 */
export function isDateApproachingOrPast(dateStr: string, daysThreshold = 90): boolean {
  if (!dateStr) return false;
  const parts = dateStr.split('.');
  if (parts.length !== 3) return false;
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const year = parseInt(parts[2], 10);
  if (isNaN(day) || isNaN(month) || isNaN(year)) return false;

  const targetDate = new Date(year, month, day);
  const now = new Date();
  const diffDays = Math.ceil((targetDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
  return diffDays <= daysThreshold;
}
