/**
 * OGM Havacılık Teknisyen ve Depo Personel Listesi Senkronizasyonu
 * Kaynak: Google Sheets "Personel" Sayfası (FULL_NAME sütunu)
 */

export const TEKNISYEN_PERSONEL_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzkIYFs3JvIEQkT3Kh-XdLXtKMdsWBaXg6XY91dk_5i16_bCJf6C9zkycklXubUTir5/exec";

export const DEFAULT_TEKNISYEN_PERSONEL_NAMES: string[] = [
  "MEHMET SEFA YÜCEDAĞ",
  "TEZCAN GÜZER",
  "ALPER ÖZMETİN",
  "ALİ ÖZKAVSAL",
  "İLYAS ARAT",
  "LEVENT ERDEM",
  "MEHMET IŞIK",
  "BAYRAM KORKMAZ",
  "UTKU GÖKGÖZ",
  "MUSTAFA GÜVEN",
  "OKAN AKÇA",
  "YILMAZ KARADENİZ",
  "DURSUN ŞİMŞEK",
  "SERKAN BOZDAĞ",
  "BARIŞ UZUN",
  "SERKAN KEBAPCI",
  "ÜMİT NAİLLİOĞLU",
  "VELİ ERSOY",
  "YUSUF BAŞPINAR",
  "HASAN AKSOY",
  "MUZAFFER SEMERCİ",
  "OĞUZ TALAY",
  "ÖNDER BİLGİN",
  "MURAT GÜNDÜZ",
  "GÖKHAN ÇETİNKAYA",
  "FERHAT ÖZCAN",
  "KEMAL CAN",
  "ÖMER ERSOY",
  "DOĞAN ÖZTÜRK",
  "HACI İBRAHİM YALABUK",
  "GÖKHAN KİRAZ",
  "AYCAN TAN",
  "HALİL AYGÖR",
  "SERDAR TOY",
  "ERSİN ÖNER",
  "TAŞKIN ÇALIŞIR",
  "FATİH YAŞBAY",
  "ÖZGÜR UYAR",
  "UĞUR GÜRCAN",
  "URAL UÇURUM",
  "YUNUS EROL",
  "SELÇUK İBİŞ",
  "ALİ TOPALAN",
  "KÜRŞAT ALKAN",
  "BAYKAN İSALAR",
  "FATİH PARLAK",
  "ABDULRAHİM DEMİR",
  "İMDAT SALMAN",
  "TALAT TÖNGÜŞ",
  "AHMET BİLGEN",
  "M.SERKAN KOTAN",
  "ZEKERİYA EVİRGEN",
  "GÖKTÜRK AKBUDAK",
  "ALKAN BALTACI",
  "FATİH ÇAMUR",
  "ESİN GÜNGÖR",
  "DİLARA USLU",
  "M.ZAHİD DURSUN",
  "ŞÜKRAN BALTACI",
  "MUSTAFA ÖZKAN",
  "EGE ERAKAY",
  "MAHİR YILMAZ",
  "YİĞİTHAN SARIHAN",
  "ABDULLAH YAPICI",
  "MUHAMMET DİLCİ",
  "YASİN DOĞAN",
  "RIDVAN ŞATIR",
  "BEYHAN SAYGIN",
  "NAZMİYE DEDE",
  "BEKİR EKİNCİ",
  "ENES KAHYAOĞLU",
  "KEREM DELİCE",
  "MUHAMMED FATİH ŞAHİN",
  "EREN KAPLAN",
  "MEHMET ADRAM",
  "FURKAN DUYURUCU",
  "BARIŞ BUĞRA KARACA",
  "YUSUF ARSLAN",
  "BARIŞ ÇELİKER",
  "FERDİ GÖL",
  "KEMAL ÇAKMAK"
];

const STORAGE_KEY = 'ogm_personel_teknisyen_list_v2';

/**
 * Get current list of personnel names from cache or default list
 */
export function getStoredPersonnelList(): string[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        // Merge with defaults to ensure complete list
        const merged = Array.from(new Set([...DEFAULT_TEKNISYEN_PERSONEL_NAMES, ...parsed]));
        return merged.sort();
      }
    }
  } catch (e) {
    console.warn('Personnel list parse error:', e);
  }
  return DEFAULT_TEKNISYEN_PERSONEL_NAMES.slice().sort();
}

/**
 * Fetch latest personnel list from the Google Script URL asynchronously
 */
export async function syncPersonnelListFromScript(): Promise<string[]> {
  try {
    const res = await fetch(`${TEKNISYEN_PERSONEL_SCRIPT_URL}?sheet=Personel&action=getPersonel`, {
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    });
    if (res.ok) {
      const data = await res.json();
      let names: string[] = [];
      if (Array.isArray(data)) {
        names = data.map((item: any) => {
          if (typeof item === 'string') return item.trim();
          if (item && typeof item === 'object') {
            return (item.FULL_NAME || item.full_name || item.name || item[2] || item.FullName || '').trim();
          }
          return '';
        }).filter(Boolean);
      } else if (data && Array.isArray(data.personel || data.data)) {
        const arr = data.personel || data.data;
        names = arr.map((item: any) => {
          if (typeof item === 'string') return item.trim();
          if (item && typeof item === 'object') {
            return (item.FULL_NAME || item.full_name || item.name || item[2] || '').trim();
          }
          return '';
        }).filter(Boolean);
      }

      if (names.length > 0) {
        const merged = Array.from(new Set([...DEFAULT_TEKNISYEN_PERSONEL_NAMES, ...names])).sort();
        localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
        return merged;
      }
    }
  } catch (err) {
    console.warn('Google Script personnel sync note:', err);
  }
  return getStoredPersonnelList();
}
