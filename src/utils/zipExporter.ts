import JSZip from 'jszip';

export interface StandaloneSiteConfig {
  id: string;
  name: string;
  htmlFileName: string;
  title: string;
  description: string;
}

/**
 * Downloads a standalone system as a deployable ZIP containing its HTML, assets, data, and deployment instructions.
 */
export async function downloadStandaloneSystemZip(systemId: 'yer-destek' | 'depo' | 'kara-araclari'): Promise<void> {
  const zip = new JSZip();

  let htmlFileName = 'yer-destek.html';
  let title = 'Hava Araçları Yer Destek ve Özel Aletler Takip Sistemi';
  let desc = 'Tüm filoların yer destek teçhizatları, kalibrasyon ve 90 gün bakım takip sistemi.';

  if (systemId === 'depo') {
    htmlFileName = 'depo.html';
    title = 'Hangar Yedek Parça ve Depo Yönetim Sistemi';
    desc = 'Sarf ve kimyasal depo envanteri, canlı P/N arama, transfer ve stok sayım sistemi.';
  } else if (systemId === 'kara-araclari') {
    htmlFileName = 'kara-araclari.html';
    title = 'Hangar Yer Destek Kara Araçları Takip Sistemi';
    desc = 'Araç filosu, görev emirleri, periyodik muayene ve 90 gün uyarı takip sistemi.';
  }

  // 1. Fetch current standalone HTML
  let htmlContent = '';
  try {
    const res = await fetch(`/${htmlFileName}`);
    htmlContent = await res.text();
  } catch (e) {
    htmlContent = `<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${title}</title>
</head>
<body>
  <div id="root"></div>
  <script type="module" src="/src/${systemId}/main.tsx"></script>
</body>
</html>`;
  }

  // Ensure index.html inside the zip so it can be hosted as root of the new website
  zip.file('index.html', htmlContent);
  zip.file(htmlFileName, htmlContent);

  // 2. Add local storage export of that system's data
  const exportedData: Record<string, any> = {};
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key) {
      if (
        (systemId === 'yer-destek' && (key.includes('techizat') || key.includes('at802') || key.includes('bell429') || key.includes('t70') || key.includes('b360') || key.includes('c650') || key.includes('hangar'))) ||
        (systemId === 'depo' && (key.includes('depo') || key.includes('sarf') || key.includes('kimyasal'))) ||
        (systemId === 'kara-araclari' && (key.includes('kara') || key.includes('arac') || key.includes('gorev')))
      ) {
        try {
          exportedData[key] = JSON.parse(localStorage.getItem(key) || '""');
        } catch {
          exportedData[key] = localStorage.getItem(key);
        }
      }
    }
  }
  zip.file('data-backup.json', JSON.stringify(exportedData, null, 2));

  // 3. Deployment instructions
  const readmeContent = `# ${title}
${desc}

## 🚀 Bağımsız Web Sitesi Kurulum ve Yayınlama Rehberi
Bu paket, ana portaldan bağımsız olarak tek başına çalışacak şekilde ayrıştırılmış müstakil bir web sitesidir.

### 1. Doğrudan Barındırma (Netlify, Vercel, Firebase Hosting, Apache, Nginx):
1. Bu ZIP dosyasını açın.
2. Web sunucunuzun kök dizinine (\`public_html\`, \`htdocs\` vb.) dosyaları yükleyin.
3. \`index.html\` sayfası ana giriş noktasıdır.

### 2. Bağımsız Çalışma Özellikleri:
- Kendi yerel önbelleğini (localStorage) kullanır, ana portala bağımlı olmadan çalışır.
- Google Drive senkronizasyon altyapısı mevcuttur.
- Excel indirme ve raporlama işlevleri dahildir.

---
OGM Havacılık Dairesi Başkanlığı
`;
  zip.file('README.md', readmeContent);

  // 4. Generate and trigger download
  const blob = await zip.generateAsync({ type: 'blob' });
  const downloadUrl = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = downloadUrl;
  a.download = `${systemId}-bagimsiz-sistem.zip`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(downloadUrl);
}
