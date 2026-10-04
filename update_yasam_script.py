import sys

with open('public/yasam-destek.html', 'r', encoding='utf-8') as f:
    text = f.read()

script_start = '<script>'
script_end = '</script>'

s_idx1 = text.index(script_start)
s_idx2 = text.rindex(script_end)

new_script = '''<script>
    let canYelegiList = [];
    let spareAirList = [];
    let helmetKitList = [];
    let zimmetList = [];
    let personnelList = [];
    let personnelItems = [];
    let activeRoleFilter = 'ALL';
    let currentModalAction = 'devret';
    let currentActiveRecord = null;
    let currentActiveTab = 'can_yelegi';

    // Renk Dolgu Mantığı (User Spec: DEPODA ise DOLGUSUZ / BEYAZ, Kişide Zimmetli ise YEŞİL DOLGU #92d050)
    function getRowColor(item) {
      const assigned = (item.aciklamalar || '').trim();
      const status = (item.durum || '').toUpperCase();
      const isEmpty = !assigned || assigned === '-' || assigned.toUpperCase() === 'DEPODA' || assigned.toUpperCase() === 'BOŞTA' || status === 'DEPODA';

      if (status === 'KAYIT SİLME' || (item.aciklamalar && item.aciklamalar.includes('KAYIT SİLME'))) {
        return 'bg-[#ff0000] text-white font-bold';
      }
      if (status === 'BAKIMA GİDECEK' || (item.aciklamalar && (item.aciklamalar.includes('DOLUM') || item.aciklamalar.includes('ARIZALI') || item.aciklamalar.includes('BAKIMA')))) {
        return 'bg-[#00b0f0] text-slate-900 font-bold';
      }
      if (status === 'ONARIMDA') {
        return 'bg-[#ffff00] text-slate-900 font-bold';
      }
      
      // Depodaki ürünler dolgusuz / beyaz
      if (isEmpty) {
        return 'bg-white hover:bg-slate-50 text-slate-800';
      }
      
      // Bir personele teslim edilmişse YEŞİL DOLGU (#92d050)
      return 'bg-[#92d050] text-slate-900 font-bold hover:bg-[#82c040]';
    }

    async function quickChangeStockStatus(type, index, newStatus) {
      let targetList = type === 'can_yelegi' ? canYelegiList : (type === 'spare_air' ? spareAirList : helmetKitList);
      if (index < 0 || index >= targetList.length) return;
      
      const item = targetList[index];
      const isDepot = newStatus === 'BOŞTA' || newStatus === 'DEPODA';
      const color = isDepot ? '#ffffff' : (newStatus === 'ONARIMDA' ? '#ffff00' : (newStatus === 'BAKIMA GİDECEK' ? '#00b0f0' : (newStatus === 'KAYIT SİLME' ? '#ff0000' : '#92d050')));
      
      targetList[index] = {
        ...item,
        durum: isDepot ? 'DEPODA' : newStatus,
        aciklamalar: isDepot ? 'DEPODA' : newStatus,
        color: color,
        tarih: isDepot ? '-' : new Date().toLocaleDateString('tr-TR')
      };
      
      await fetch('/api/yasam-destek/stock-update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type, items: targetList })
      });
      
      fetchData();
    }

    // TAB GEÇİŞLERİ (CAN YELEĞİ, SPARE AIR, HELMET KIT, ZİMMETLER, PERSONEL, YENİ ZİMMET)
    function switchPublicTab(t) {
      currentActiveTab = t;
      const allTabs = ['can_yelegi', 'spare_air', 'helmet_kit', 'zimmetler', 'personnel'];
      
      allTabs.forEach(id => {
        const sec = document.getElementById('section_' + id);
        const btn = document.getElementById('tab_' + id);
        if (id === t) {
          if (sec) sec.classList.remove('hidden');
          if (btn) {
            btn.className = "px-3 sm:px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider transition flex items-center gap-1.5 bg-[#0b3d1d] text-white shadow-xs cursor-pointer";
          }
        } else {
          if (sec) sec.classList.add('hidden');
          if (btn) {
            btn.className = "px-3 sm:px-4 py-2 rounded-xl text-xs font-bold uppercase tracking-wider transition flex items-center gap-1.5 bg-white text-slate-700 hover:bg-slate-100 cursor-pointer shadow-2xs";
          }
        }
      });

      const secNew = document.getElementById('section_new_form');
      const btnNew = document.getElementById('tab_new_zimmet');

      if (t === 'new') {
        allTabs.forEach(id => {
          const sec = document.getElementById('section_' + id);
          if (sec) sec.classList.add('hidden');
        });
        if (secNew) secNew.classList.remove('hidden');
        if (btnNew) {
          btnNew.className = "px-3.5 sm:px-4 py-2 rounded-xl bg-[#0b3d1d] text-white text-xs font-black uppercase tracking-wider transition flex items-center gap-1.5 cursor-pointer shadow-md";
        }
        const now = new Date();
        const inpBaslangic = document.getElementById('inpBaslangic');
        if (inpBaslangic && !inpBaslangic.value) {
          inpBaslangic.value = now.toLocaleDateString('tr-TR');
        }
        populateFormDropdowns();
      } else {
        if (secNew) secNew.classList.add('hidden');
        if (btnNew) {
          btnNew.className = "px-3.5 sm:px-4 py-2 rounded-xl bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-black uppercase tracking-wider transition flex items-center gap-1.5 cursor-pointer shadow-md active:scale-95";
        }
      }
    }

    function switchTab(t) {
      switchPublicTab(t);
    }

    // TÜM VERİLERİ ÇEK
    async function fetchData() {
      try {
        const r1 = await fetch('/api/yasam-destek/all-stocks');
        const j1 = await r1.json();
        if (j1 && j1.status === 'success') {
          canYelegiList = j1.canYelegi || [];
          spareAirList = j1.spareAir || [];
          helmetKitList = j1.helmetKit || [];
          renderStocks();
        }

        const r2 = await fetch('/api/yasam-destek');
        const j2 = await r2.json();
        if (j2 && j2.status === 'success') {
          zimmetList = j2.records || [];
          renderZimmetler();
        }

        const r3 = await fetch('/api/yasam-destek/personnel-list');
        const j3 = await r3.json();
        if (j3 && j3.status === 'success') {
          personnelList = j3.personnel || [];
          personnelItems = j3.items || [];
          const countP = document.getElementById('count_personnel');
          if (countP) countP.innerText = personnelItems.length;
          renderPersonnelTable();
          populateFormDropdowns();
        }
      } catch (e) {
        console.warn('Veri çekme hatası:', e);
      }
    }

    // FORM AÇILIR LİSTELERİNİ VE OTOMATİK SEÇİMLERİ DOLDUR
    function populateFormDropdowns() {
      // 1. Personel Select
      const selP = document.getElementById('selPersonel');
      if (selP) {
        selP.innerHTML = '<option value="">Listeden Personel Seçin...</option>' + 
          personnelItems.map(p => `<option value="${p.adSoyad}">${p.adSoyad} (${p.unvan || 'Teknisyen'})</option>`).join('');
      }
      const countEl = document.getElementById('formPersonnelCount');
      if (countEl) countEl.innerText = `${personnelItems.length} Personel`;

      // 2. Can Yeleği Select
      const selLV = document.getElementById('selLifeVest');
      if (selLV) {
        const availableCY = canYelegiList.filter(c => c.durum !== 'KAYIT SİLME');
        selLV.innerHTML = '<option value="">Depodaki Can Yeleklerinden Seç...</option>' + 
          availableCY.map(c => {
            const isDepot = c.durum === 'DEPODA' || c.aciklamalar === 'DEPODA' || !c.aciklamalar || c.aciklamalar === '-';
            const label = isDepot ? `${c.seriNo} [DEPODA / FAAL]` : `${c.seriNo} (${c.aciklamalar})`;
            return `<option value="${c.seriNo}">${label}</option>`;
          }).join('');
      }
      const countLV = document.getElementById('formLifeVestCount');
      if (countLV) {
        const depotLV = canYelegiList.filter(c => c.durum === 'DEPODA' || c.aciklamalar === 'DEPODA').length;
        countLV.innerText = `${depotLV} / ${canYelegiList.length} Depoda`;
      }

      // 3. Helmet Kit Select
      const selHK = document.getElementById('selHelmetKit');
      if (selHK) {
        const availableHK = helmetKitList.filter(h => h.durum !== 'KAYIT SİLME');
        selHK.innerHTML = '<option value="">Depodaki Kask & Kulaklıktan Seç...</option>' + 
          availableHK.map(h => {
            const isDepot = h.durum === 'DEPODA' || h.aciklamalar === 'DEPODA' || !h.aciklamalar || h.aciklamalar === '-';
            const snVal = h.seriNo2 && h.seriNo2 !== '-' ? h.seriNo2 : (h.seriNo1 || h.aciklamalar);
            const label = isDepot ? `${snVal} [DEPODA / FAAL]` : `${snVal} (${h.aciklamalar})`;
            return `<option value="${snVal}">${label}</option>`;
          }).join('');
      }
      const countHK = document.getElementById('formHelmetKitCount');
      if (countHK) {
        const depotHK = helmetKitList.filter(h => h.durum === 'DEPODA' || h.aciklamalar === 'DEPODA').length;
        countHK.innerText = `${depotHK} / ${helmetKitList.length} Depoda`;
      }

      // 4. Spare Air Select
      const selSA = document.getElementById('selSpareAir');
      if (selSA) {
        const availableSA = spareAirList.filter(s => s.durum !== 'KAYIT SİLME');
        selSA.innerHTML = '<option value="">Depodaki Spare Air\\'lerden Seç...</option>' + 
          availableSA.map(s => {
            const isDepot = s.durum === 'DEPODA' || s.aciklamalar === 'DEPODA' || !s.aciklamalar || s.aciklamalar === '-';
            const label = isDepot ? `${s.seriNo} [DEPODA / FAAL]` : `${s.seriNo} (${s.aciklamalar})`;
            return `<option value="${s.seriNo}">${label}</option>`;
          }).join('');
      }
      const countSA = document.getElementById('formSpareAirCount');
      if (countSA) {
        const depotSA = spareAirList.filter(s => s.durum === 'DEPODA' || s.aciklamalar === 'DEPODA').length;
        countSA.innerText = `${depotSA} / ${spareAirList.length} Depoda`;
      }
    }

    function onSelectPersonelChange(val) {
      const inp = document.getElementById('inpPersonel');
      if (inp) inp.value = val;
    }

    function handleLocationChange(val) {
      const customDiv = document.getElementById('customLocWrapper');
      if (customDiv) {
        customDiv.className = val === 'DİĞER' ? 'block' : 'hidden';
      }
    }

    // YENİ ZİMMET FORMU GÖNDERME (GÖREVİ BAŞLAT)
    async function handleFormSubmit(e) {
      e.preventDefault();
      const personel = document.getElementById('inpPersonel').value.trim().toUpperCase();
      const baslangic = document.getElementById('inpBaslangic').value.trim();
      const loc = document.getElementById('inpGorevYeri').value;
      const customLoc = document.getElementById('inpCustomLocation')?.value.trim().toUpperCase();
      const gorevYeri = loc === 'DİĞER' ? (customLoc || 'DİĞER') : loc;
      const lifeVest = document.getElementById('inpLifeVest').value.trim();
      const helmetKit = document.getElementById('inpHelmetKit').value.trim();
      const spareAir = document.getElementById('inpSpareAir').value.trim();
      const notlar = document.getElementById('inpNotlar').value.trim();

      if (!personel) {
        alert('Lütfen Personel Adı Soyadı seçiniz veya yazınız.');
        return;
      }
      if (!lifeVest) {
        alert('Lütfen Life Vest (Can Yeleği) seri numarasını seçiniz.');
        return;
      }
      if (!helmetKit) {
        alert('Lütfen Helmet Kit & Kulaklık seri numarasını seçiniz.');
        return;
      }
      if (!spareAir) {
        alert('Lütfen Spare Air (Yedek Hava) seri numarasını seçiniz.');
        return;
      }

      const btn = document.getElementById('btnGoreviBaslat');
      if (btn) btn.disabled = true;

      try {
        const res = await fetch('/api/yasam-destek/assign', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            personelAdi: personel,
            lifeVestSn: lifeVest,
            helmetKitSn: helmetKit,
            spareAirSn: spareAir,
            gorevYeri: gorevYeri,
            baslangicTarihi: baslangic || new Date().toLocaleDateString('tr-TR'),
            notlar: notlar
          })
        });

        const json = await res.json();
        if (json && json.status === 'success') {
          alert(`✅ Görev Başarıyla Başlatıldı!\\n\\nPersonel: ${personel}\\nGörev Yeri: ${gorevYeri}\\n\\nLife Vest (${lifeVest}), Kulaklık (${helmetKit}) ve Spare Air (${spareAir}) ekipmanları ${personel} adına zimmetlenerek stok tablolarında YEŞİL DOLGUYA çevrildi.`);
          
          // Formu temizle
          document.getElementById('inpPersonel').value = '';
          document.getElementById('selPersonel').value = '';
          document.getElementById('inpLifeVest').value = '';
          document.getElementById('selLifeVest').value = '';
          document.getElementById('inpHelmetKit').value = '';
          document.getElementById('selHelmetKit').value = '';
          document.getElementById('inpSpareAir').value = '';
          document.getElementById('selSpareAir').value = '';
          document.getElementById('inpNotlar').value = '';

          await fetchData();
          switchPublicTab('personnel');
        } else {
          alert('Hata: ' + (json.message || 'Kayıt yapılamadı.'));
        }
      } catch (err) {
        alert('Bağlantı hatası: ' + err.message);
      } finally {
        if (btn) btn.disabled = false;
      }
    }

    // BELİRLİ BİR PERSONELE DİREKT ZİMMET VER
    function openNewZimmetForPersonnel(personName) {
      switchPublicTab('new');
      setTimeout(() => {
        const inpP = document.getElementById('inpPersonel');
        const selP = document.getElementById('selPersonel');
        if (inpP) inpP.value = personName;
        if (selP) selP.value = personName;
      }, 50);
    }

    // ZİMMETİ GÖR MODALINI AÇ
    function openViewZimmetModal(personName) {
      const cleanName = personName.trim().toUpperCase();
      const pObj = personnelItems.find(p => p.adSoyad.toUpperCase() === cleanName) || { adSoyad: cleanName, unvan: 'Personel' };
      const activeZimmet = zimmetList.find(z => z.personelAdi.toUpperCase() === cleanName && z.durum === 'GÖREVDE');

      const modal = document.getElementById('viewZimmetModal');
      const content = document.getElementById('viewZimmetModalContent');
      if (!modal || !content) return;

      if (activeZimmet) {
        content.innerHTML = `
          <div class="p-4 bg-emerald-50 border-2 border-emerald-500/40 rounded-2xl space-y-3">
            <div class="flex items-center justify-between border-b border-emerald-200 pb-2.5">
              <div>
                <div class="text-[10px] font-black uppercase text-emerald-800 tracking-wider">ZİMMET SAHİBİ PERSONEL:</div>
                <div class="text-base font-black text-slate-900">${pObj.adSoyad}</div>
                <div class="text-xs font-bold text-slate-500">${pObj.unvan || 'Teknisyen'}</div>
              </div>
              <span class="px-3 py-1 rounded-full text-xs font-black uppercase bg-[#92d050] text-slate-950 border border-emerald-600 shadow-2xs">
                ● AKTİF GÖREVDE
              </span>
            </div>

            <div class="grid grid-cols-2 gap-2.5 text-xs pt-1">
              <div class="p-2.5 bg-white rounded-xl border border-slate-200">
                <div class="text-[10px] font-black text-slate-400 uppercase">Görev Yeri:</div>
                <div class="font-black text-slate-900 mt-0.5 text-sm">📍 ${activeZimmet.gorevYeri}</div>
              </div>
              <div class="p-2.5 bg-white rounded-xl border border-slate-200">
                <div class="text-[10px] font-black text-slate-400 uppercase">Başlangıç Tarihi:</div>
                <div class="font-mono font-bold text-slate-900 mt-0.5 text-sm">📅 ${activeZimmet.baslangicTarihi}</div>
              </div>
            </div>

            <div class="space-y-2 pt-2 border-t border-emerald-200">
              <div class="text-[11px] font-black uppercase text-slate-800">Zimmetli Ekipman Seri Numaraları:</div>
              <div class="p-2.5 bg-amber-50/80 rounded-xl border border-amber-300 flex items-center justify-between text-xs">
                <span class="font-bold text-amber-950">🦺 Life Vest (Can Yeleği):</span>
                <span class="font-mono font-black text-amber-900 text-sm">${activeZimmet.lifeVestSn || '-'}</span>
              </div>
              <div class="p-2.5 bg-purple-50/80 rounded-xl border border-purple-300 flex items-center justify-between text-xs">
                <span class="font-bold text-purple-950">🪖 Helmet Kit & Kulaklık:</span>
                <span class="font-mono font-black text-purple-900 text-sm">${activeZimmet.helmetKitSn || '-'}</span>
              </div>
              <div class="p-2.5 bg-cyan-50/80 rounded-xl border border-cyan-300 flex items-center justify-between text-xs">
                <span class="font-bold text-cyan-950">💨 Spare Air (Yedek Hava):</span>
                <span class="font-mono font-black text-cyan-900 text-sm">${activeZimmet.spareAirSn || '-'}</span>
              </div>
            </div>

            ${activeZimmet.notlar ? `
              <div class="p-2.5 bg-slate-50 rounded-xl border border-slate-200 text-xs text-slate-700">
                <div class="text-[10px] font-black text-slate-400 uppercase mb-0.5">Görev Notu:</div>
                <div>${activeZimmet.notlar}</div>
              </div>
            ` : ''}

            <div class="pt-4 flex flex-wrap items-center justify-end gap-2 border-t border-emerald-200">
              <button onclick="printTutanak('${activeZimmet.id}')" class="px-3.5 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer">
                <i class="fa-solid fa-print"></i>
                <span>Tutanak Yazdır</span>
              </button>
              <button onclick="closeViewZimmetModal(); openFinishModal('${activeZimmet.id}');" class="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-xl text-xs font-black uppercase tracking-wider transition flex items-center gap-1.5 cursor-pointer shadow-md">
                <i class="fa-solid fa-arrow-right-arrow-left"></i>
                <span>Zimmeti Bitir / Devret</span>
              </button>
            </div>
          </div>
        `;
      } else {
        content.innerHTML = `
          <div class="p-6 text-center space-y-4">
            <div class="w-14 h-14 bg-slate-100 text-slate-400 rounded-2xl flex items-center justify-center text-2xl mx-auto shadow-inner">
              🦺
            </div>
            <div>
              <div class="text-sm font-black text-slate-900 uppercase">${pObj.adSoyad} (${pObj.unvan || 'Teknisyen'})</div>
              <p class="text-xs text-slate-500 mt-1">Bu personele ait şu an aktif bir yaşam destek zimmeti bulunmamaktadır. Tüm ekipmanları depoda veya boştadır.</p>
            </div>
            <div class="pt-2 flex justify-center gap-2">
              <button onclick="closeViewZimmetModal()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold cursor-pointer">
                Kapat
              </button>
              <button onclick="closeViewZimmetModal(); openNewZimmetForPersonnel('${pObj.adSoyad}');" class="px-4 py-2 bg-[#0b3d1d] hover:bg-[#072813] text-white rounded-xl text-xs font-black uppercase tracking-wider cursor-pointer shadow-md">
                + Bu Personele Zimmet Ver
              </button>
            </div>
          </div>
        `;
      }

      modal.classList.remove('hidden');
    }

    function closeViewZimmetModal() {
      const modal = document.getElementById('viewZimmetModal');
      if (modal) modal.classList.add('hidden');
    }

    // GÖREVİ BİTİR / DEVİR & TESLİM MODALI
    function openFinishModal(id) {
      const rec = zimmetList.find(r => r.id === id);
      if (!rec) return;
      currentActiveRecord = rec;
      document.getElementById('modalCurrentPersonel').innerText = rec.personelAdi;
      document.getElementById('modalCurrentInfo').innerText = 'Yer: ' + rec.gorevYeri + ' | Başlangıç: ' + rec.baslangicTarihi + ' | LV: ' + rec.lifeVestSn + ' | HK: ' + rec.helmetKitSn + ' | SA: ' + rec.spareAirSn;
      const now = new Date();
      document.getElementById('modalBitisTarihi').value = now.toLocaleDateString('tr-TR');
      setModalAction('devret');
      document.getElementById('finishModal').classList.remove('hidden');
    }

    function closeFinishModal() {
      document.getElementById('finishModal').classList.add('hidden');
      currentActiveRecord = null;
    }

    function setModalAction(a) {
      currentModalAction = a;
      const btnDevret = document.getElementById('btnDevret');
      const btnTeslim = document.getElementById('btnTeslim');
      const secDevret = document.getElementById('secDevret');
      const secTeslim = document.getElementById('secTeslim');

      if (a === 'devret') {
        if (btnDevret) btnDevret.className = "p-3 rounded-xl border-2 border-indigo-600 bg-indigo-50/50 text-indigo-950 font-bold text-left cursor-pointer";
        if (btnTeslim) btnTeslim.className = "p-3 rounded-xl border-2 border-slate-200 bg-white text-slate-600 text-left cursor-pointer";
        if (secDevret) secDevret.classList.remove('hidden');
        if (secTeslim) secTeslim.classList.add('hidden');
      } else {
        if (btnDevret) btnDevret.className = "p-3 rounded-xl border-2 border-slate-200 bg-white text-slate-600 text-left cursor-pointer";
        if (btnTeslim) btnTeslim.className = "p-3 rounded-xl border-2 border-blue-600 bg-blue-50/50 text-blue-950 font-bold text-left cursor-pointer";
        if (secDevret) secDevret.classList.add('hidden');
        if (secTeslim) secTeslim.classList.remove('hidden');
      }
    }

    async function submitFinishModal() {
      if (!currentActiveRecord) return;
      const bitis = document.getElementById('modalBitisTarihi').value.trim();
      const notes = document.getElementById('modalNotes').value.trim();

      const devredilen = document.getElementById('modalDevredilenPersonel')?.value.trim().toUpperCase();
      const yer = document.getElementById('modalDevredilenYer')?.value;
      const teslimAlan = document.getElementById('modalTeslimAlan')?.value.trim();

      if (currentModalAction === 'devret' && !devredilen) {
        alert('Lütfen devredilecek personelin adını yazınız veya listeden seçiniz.');
        return;
      }

      try {
        const res = await fetch('/api/yasam-destek/finish', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            zimmetId: currentActiveRecord.id,
            action: currentModalAction,
            bitisTarihi: bitis || new Date().toLocaleDateString('tr-TR'),
            devredilenPersonel: devredilen,
            devredilenYer: yer,
            teslimAlanPersonel: teslimAlan,
            notlar: notes
          })
        });

        const json = await res.json();
        if (json && json.status === 'success') {
          if (currentModalAction === 'devret') {
            alert(`✅ Zimmet Başarıyla Devredildi!\\n\\nEski Görevli: ${currentActiveRecord.personelAdi}\\nYeni Görevli: ${devredilen}\\nYeni Görev Yeri: ${yer}\\n\\nStok tablolarında teslim edilen kişi ${devredilen} olarak güncellenip YEŞİL DOLGU (#92d050) olarak korunmuştur.`);
          } else {
            alert(`✅ Ekipmanlar Depoya Teslim Edildi!\\n\\nTeslim Alan: ${teslimAlan || 'Depo Sorumlusu'}\\n\\nLife Vest (${currentActiveRecord.lifeVestSn}), Kulaklık (${currentActiveRecord.helmetKitSn}) ve Spare Air (${currentActiveRecord.spareAirSn}) stok tablolarında "DEPODA" olarak işaretlenmiş ve BEYAZ DOLGUYA (dolgusuz) çevrilmiştir.`);
          }
          closeFinishModal();
          await fetchData();
          switchPublicTab('personnel');
        } else {
          alert('İşlem başarısız: ' + (json.message || 'Bilinmeyen hata'));
        }
      } catch (err) {
        alert('Hata: ' + err.message);
      }
    }

    // TÜM YAŞAM DESTEK STOKLARINI BAŞLANGIÇ DURUMUNA SIFIRLA (DEPODA VE BEYAZ DOLGU)
    async function resetAllStocksToDepot() {
      if (!confirm('Tüm Yaşam Destek stoklarını DEPODA ve BEYAZ DOLGU (dolgusuz) başlangıç durumuna getirmek istediğinize emin misiniz?\\n\\nAktif zimmetler depoya teslim edilmiş olarak işaretlenecektir.')) {
        return;
      }

      try {
        const res = await fetch('/api/yasam-destek/reset-to-depot', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' }
        });
        const json = await res.json();
        if (json && json.status === 'success') {
          alert('✅ Başarılı: Tüm Yaşam Destek stokları "DEPODA" ve BEYAZ DOLGU durumuna getirildi.');
          await fetchData();
        } else {
          alert('Hata: ' + (json.message || 'Sıfırlama yapılamadı.'));
        }
      } catch (e) {
        alert('Bağlantı hatası: ' + e.message);
      }
    }

    // TABLOLARI RENDER ET
    function renderStocks() {
      const query = (document.getElementById('htmlSearchInput')?.value || '').trim().toLowerCase();

      // 1. Can Yeleği
      const filteredCY = canYelegiList.filter(c => !query || 
        (c.aciklamalar || '').toLowerCase().includes(query) || 
        (c.notlar || '').toLowerCase().includes(query) || 
        (c.seriNo || '').toLowerCase().includes(query) || 
        (c.disNo || '').toLowerCase().includes(query) ||
        (c.durum || '').toLowerCase().includes(query)
      );
      const countCYEl = document.getElementById('count_cy');
      if (countCYEl) countCYEl.innerText = filteredCY.length;

      document.getElementById('tbody_can_yelegi').innerHTML = filteredCY.map((c, i) => {
        const isDepot = c.durum === 'DEPODA' || c.aciklamalar === 'DEPODA' || !c.aciklamalar || c.aciklamalar === '-';
        const displayAssigned = isDepot ? 'DEPODA' : c.aciklamalar;
        
        return `
          <tr class="border-b border-slate-300 ${getRowColor(c)} text-center transition">
            <td class="py-2 px-2 border-r border-slate-300 font-mono">${c.sNo || i + 1}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans">${c.malzemeAdi}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono">${c.parcaNo}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono font-black">${c.seriNo}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono font-black">${c.disNo || '-'}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans text-slate-700 font-medium">${c.notlar || c.aciklama || '-'}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans font-black ${isDepot ? 'text-slate-600' : 'text-slate-950'}">${displayAssigned}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono">${isDepot ? '-' : (c.tarih || 'DEVAM EDİYOR')}</td>
            <td class="py-1 px-1">
              <div class="flex flex-wrap items-center justify-center gap-1">
                <button onclick="quickChangeStockStatus('can_yelegi', ${canYelegiList.indexOf(c)}, 'ONARIMDA')" class="px-1.5 py-0.5 bg-yellow-400 hover:bg-yellow-500 text-slate-950 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">🔧 ONARIM</button>
                <button onclick="quickChangeStockStatus('can_yelegi', ${canYelegiList.indexOf(c)}, 'BAKIMA GİDECEK')" class="px-1.5 py-0.5 bg-cyan-400 hover:bg-cyan-500 text-slate-950 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">⚙️ BAKIM</button>
                <button onclick="quickChangeStockStatus('can_yelegi', ${canYelegiList.indexOf(c)}, 'DEPODA')" class="px-1.5 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">🏢 DEPOYA AL</button>
                <button onclick="promptInlineStockEdit('can_yelegi', ${canYelegiList.indexOf(c)})" class="p-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded text-xs border border-slate-300 cursor-pointer" title="Düzenle">✏️</button>
                <button onclick="deleteStockRowHTML('can_yelegi', ${canYelegiList.indexOf(c)})" class="p-1 bg-red-100 hover:bg-red-200 text-red-700 rounded text-xs border border-red-200 cursor-pointer" title="Sil">🗑️</button>
              </div>
            </td>
          </tr>
        `;
      }).join('');

      // 2. Spare Air
      const filteredSA = spareAirList.filter(s => !query || 
        (s.aciklamalar || '').toLowerCase().includes(query) || 
        (s.notlar || '').toLowerCase().includes(query) || 
        (s.seriNo || '').toLowerCase().includes(query) ||
        (s.durum || '').toLowerCase().includes(query)
      );
      const countSAEl = document.getElementById('count_sa');
      if (countSAEl) countSAEl.innerText = filteredSA.length;

      document.getElementById('tbody_spare_air').innerHTML = filteredSA.map((s, i) => {
        const isDepot = s.durum === 'DEPODA' || s.aciklamalar === 'DEPODA' || !s.aciklamalar || s.aciklamalar === '-';
        const displayAssigned = isDepot ? 'DEPODA' : s.aciklamalar;

        return `
          <tr class="border-b border-slate-300 ${getRowColor(s)} text-center transition">
            <td class="py-2 px-2 border-r border-slate-300 font-mono">${s.sNo || i + 1}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans">${s.malzemeAdi}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono">${s.parcaNo}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono font-black">${s.seriNo}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans text-slate-700 font-medium">${s.notlar || s.aciklama || '-'}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans font-black ${isDepot ? 'text-slate-600' : 'text-slate-950'}">${displayAssigned}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono">${isDepot ? '-' : (s.tarih || 'DEVAM EDİYOR')}</td>
            <td class="py-1 px-1">
              <div class="flex flex-wrap items-center justify-center gap-1">
                <button onclick="quickChangeStockStatus('spare_air', ${spareAirList.indexOf(s)}, 'ONARIMDA')" class="px-1.5 py-0.5 bg-yellow-400 hover:bg-yellow-500 text-slate-950 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">🔧 ONARIM</button>
                <button onclick="quickChangeStockStatus('spare_air', ${spareAirList.indexOf(s)}, 'BAKIMA GİDECEK')" class="px-1.5 py-0.5 bg-cyan-400 hover:bg-cyan-500 text-slate-950 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">⚙️ BAKIM</button>
                <button onclick="quickChangeStockStatus('spare_air', ${spareAirList.indexOf(s)}, 'DEPODA')" class="px-1.5 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">🏢 DEPOYA AL</button>
                <button onclick="promptInlineStockEdit('spare_air', ${spareAirList.indexOf(s)})" class="p-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded text-xs border border-slate-300 cursor-pointer" title="Düzenle">✏️</button>
                <button onclick="deleteStockRowHTML('spare_air', ${spareAirList.indexOf(s)})" class="p-1 bg-red-100 hover:bg-red-200 text-red-700 rounded text-xs border border-red-200 cursor-pointer" title="Sil">🗑️</button>
              </div>
            </td>
          </tr>
        `;
      }).join('');

      // 3. Helmet Kit
      const filteredHK = helmetKitList.filter(h => !query || 
        (h.aciklamalar || '').toLowerCase().includes(query) || 
        (h.notlar || '').toLowerCase().includes(query) || 
        (h.seriNo1 || '').toLowerCase().includes(query) ||
        (h.seriNo2 || '').toLowerCase().includes(query) ||
        (h.durum || '').toLowerCase().includes(query)
      );
      const countHKEl = document.getElementById('count_hk');
      if (countHKEl) countHKEl.innerText = filteredHK.length;

      document.getElementById('tbody_helmet_kit').innerHTML = filteredHK.map((h, i) => {
        const isDepot = h.durum === 'DEPODA' || h.aciklamalar === 'DEPODA' || !h.aciklamalar || h.aciklamalar === '-';
        const displayAssigned = isDepot ? 'DEPODA' : h.aciklamalar;

        return `
          <tr class="border-b border-slate-300 ${getRowColor(h)} text-center transition">
            <td class="py-2 px-2 border-r border-slate-300 font-mono">${h.sNo || i + 1}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-sans">${h.malzemeAdi1}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono">${h.parcaNo1}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono font-bold">${h.seriNo1 || '-'}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-sans">${h.malzemeAdi2}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono">${h.parcaNo2}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono font-black text-indigo-950">${h.seriNo2 || '-'}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans text-slate-700 font-medium">${h.notlar || h.aciklama || '-'}</td>
            <td class="py-2 px-4 border-r border-slate-300 text-left font-sans font-black ${isDepot ? 'text-slate-600' : 'text-slate-950'}">${displayAssigned}</td>
            <td class="py-2 px-3 border-r border-slate-300 font-mono">${isDepot ? '-' : (h.tarih || 'DEVAM EDİYOR')}</td>
            <td class="py-1 px-1">
              <div class="flex flex-wrap items-center justify-center gap-1">
                <button onclick="quickChangeStockStatus('helmet_kit', ${helmetKitList.indexOf(h)}, 'ONARIMDA')" class="px-1.5 py-0.5 bg-yellow-400 hover:bg-yellow-500 text-slate-950 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">🔧 ONARIM</button>
                <button onclick="quickChangeStockStatus('helmet_kit', ${helmetKitList.indexOf(h)}, 'BAKIMA GİDECEK')" class="px-1.5 py-0.5 bg-cyan-400 hover:bg-cyan-500 text-slate-950 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">⚙️ BAKIM</button>
                <button onclick="quickChangeStockStatus('helmet_kit', ${helmetKitList.indexOf(h)}, 'DEPODA')" class="px-1.5 py-0.5 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-[4px] text-[10px] font-black cursor-pointer shadow-2xs">🏢 DEPOYA AL</button>
                <button onclick="promptInlineStockEdit('helmet_kit', ${helmetKitList.indexOf(h)})" class="p-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded text-xs border border-slate-300 cursor-pointer" title="Düzenle">✏️</button>
                <button onclick="deleteStockRowHTML('helmet_kit', ${helmetKitList.indexOf(h)})" class="p-1 bg-red-100 hover:bg-red-200 text-red-700 rounded text-xs border border-red-200 cursor-pointer" title="Sil">🗑️</button>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    // GÖREV VE ZİMMETLER TABLOSU
    function renderZimmetler() {
      const query = (document.getElementById('htmlSearchInput')?.value || '').trim().toLowerCase();
      const filteredZM = zimmetList.filter(r => !query ||
        (r.personelAdi || '').toLowerCase().includes(query) ||
        (r.gorevYeri || '').toLowerCase().includes(query) ||
        (r.notlar || '').toLowerCase().includes(query) ||
        (r.lifeVestSn || '').toLowerCase().includes(query) ||
        (r.spareAirSn || '').toLowerCase().includes(query) ||
        (r.helmetKitSn || '').toLowerCase().includes(query) ||
        (r.baslangicTarihi || '').toLowerCase().includes(query) ||
        (r.bitisTarihi || '').toLowerCase().includes(query) ||
        (r.durum || '').toLowerCase().includes(query)
      );

      const countZM = document.getElementById('count_zm');
      if (countZM) countZM.innerText = filteredZM.length;

      const tbody = document.getElementById('tbody_zimmetler');
      if (!tbody) return;

      if (filteredZM.length === 0) {
        tbody.innerHTML = `<tr><td colspan="12" class="p-8 text-center text-slate-400 font-sans">Aramanıza uygun kayıt bulunamadı veya henüz aktif zimmet bulunmamaktadır.</td></tr>`;
        return;
      }

      tbody.innerHTML = filteredZM.map((r, i) => `
        <tr class="hover:bg-slate-50 border-b border-slate-200 text-center transition">
          <td class="py-2 px-2 font-bold text-slate-500 border-r border-slate-200">${i + 1}</td>
          <td class="py-2 px-4 font-bold text-slate-900 border-r border-slate-200 text-left font-sans">${r.personelAdi}</td>
          <td class="py-2 px-3 border-r border-slate-200 font-sans"><span class="px-2 py-0.5 rounded bg-slate-100 font-bold border border-slate-300 text-[10px]">📍 ${r.gorevYeri}</span></td>
          <td class="py-2 px-3 border-r border-slate-200 text-slate-600">${r.baslangicTarihi}</td>
          <td class="py-2 px-3 border-r border-slate-200 text-slate-600 font-bold text-emerald-800">${r.bitisTarihi || 'DEVAM EDİYOR'}</td>
          <td class="py-2 px-3 border-r border-slate-200 font-bold text-amber-900 bg-amber-50/50">${r.lifeVestSn}</td>
          <td class="py-2 px-3 border-r border-slate-200 font-bold text-cyan-900 bg-cyan-50/50">${r.spareAirSn}</td>
          <td class="py-2 px-3 border-r border-slate-200 font-bold text-purple-900 bg-purple-50/50">${r.helmetKitSn}</td>
          <td class="py-2 px-3 border-r border-slate-200 font-sans">
            <span class="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase ${
              r.durum === 'GÖREVDE' ? 'bg-[#92d050] text-slate-950 border border-emerald-600' :
              r.durum === 'DEVREDİLDİ' ? 'bg-amber-100 text-amber-900 border border-amber-300' : 'bg-blue-100 text-blue-900 border border-blue-300'
            }">${r.durum}</span>
          </td>
          <td class="py-2 px-4 border-r border-slate-200 text-left font-sans text-xs text-slate-700 font-medium">${r.notlar || '-'}</td>
          <td class="py-2 px-3 border-r border-slate-200 font-sans text-left text-[10px]">${r.devredilenPersonel && r.devredilenPersonel !== '-' ? r.devredilenPersonel : (r.teslimAlanPersonel || '-')}</td>
          <td class="py-2 px-2 bg-slate-50/50">
            <div class="flex items-center justify-center gap-1.5">
              ${r.durum === 'GÖREVDE' ? `
                <button onclick="openFinishModal('${r.id}')" class="px-2 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded text-[10px] font-black uppercase cursor-pointer">Devir/Teslim</button>
              ` : ''}
              <button onclick="printTutanak('${r.id}')" class="p-1 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded text-xs cursor-pointer" title="Yazdır"><i class="fa-solid fa-print"></i></button>
              <button onclick="deleteZimmetHTML('${r.id}')" class="p-1 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded text-xs cursor-pointer" title="Sil">🗑️</button>
            </div>
          </td>
        </tr>
      `).join('');
    }

    // PERSONEL LİSTESİ TABLOSU (User Spec: Yanında Zimmeti Gör Butonu ve Zimmeti Bitir Butonu)
    function filterPersonnelByRole(role) {
      activeRoleFilter = role;
      ['ALL', 'Teknisyen', 'Pilot'].forEach(r => {
        const btn = document.getElementById('btnRole' + (r === 'ALL' ? 'All' : r));
        if (btn) {
          if (r === role) {
            btn.className = "px-3 py-1 rounded-lg text-xs font-black bg-indigo-900 text-white cursor-pointer";
          } else {
            btn.className = "px-3 py-1 rounded-lg text-xs font-bold bg-slate-100 text-slate-700 hover:bg-slate-200 cursor-pointer";
          }
        }
      });
      renderPersonnelTable();
    }

    function renderPersonnelTable() {
      const query = (document.getElementById('htmlSearchInput')?.value || '').trim().toLowerCase();
      const filtered = personnelItems.filter(p => {
        const matchesRole = activeRoleFilter === 'ALL' || (p.unvan && p.unvan.toLowerCase() === activeRoleFilter.toLowerCase());
        const matchesQuery = !query || p.adSoyad.toLowerCase().includes(query) || (p.unvan || '').toLowerCase().includes(query);
        return matchesRole && matchesQuery;
      });

      const tbody = document.getElementById('tbody_personnel');
      if (!tbody) return;

      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="p-8 text-center text-slate-400">Aranan kriterlere uygun personel bulunamadı.</td></tr>`;
        return;
      }

      tbody.innerHTML = filtered.map((p, i) => {
        const activeZimmet = zimmetList.find(z => z.personelAdi.toUpperCase() === p.adSoyad.toUpperCase() && z.durum === 'GÖREVDE');
        
        return `
          <tr class="hover:bg-slate-50 border-b border-slate-200 text-center transition">
            <td class="py-2.5 px-2 font-bold text-slate-500 border-r border-slate-200 font-mono">${i + 1}</td>
            <td class="py-2.5 px-4 font-black text-slate-900 border-r border-slate-200 text-left font-sans">
              <div class="flex flex-col">
                <div class="flex items-center gap-2">
                  <span class="text-sm font-black">${p.adSoyad}</span>
                  ${activeZimmet ? `
                    <span class="px-2 py-0.5 rounded-full text-[9px] font-black uppercase bg-[#92d050] text-slate-950 border border-emerald-600">
                      📍 GÖREVDE: ${activeZimmet.gorevYeri}
                    </span>
                  ` : `
                    <span class="px-2 py-0.5 rounded-full text-[9px] font-bold uppercase bg-slate-100 text-slate-500 border border-slate-200">
                      DEPODA / BOŞTA
                    </span>
                  `}
                </div>
                ${activeZimmet ? `
                  <div class="mt-1.5 flex flex-wrap gap-1">
                    <span class="text-[10px] bg-amber-100 text-amber-900 px-2 py-0.5 rounded-lg border border-amber-300 font-mono font-bold">🦺 LV: ${activeZimmet.lifeVestSn}</span>
                    <span class="text-[10px] bg-purple-100 text-purple-900 px-2 py-0.5 rounded-lg border border-purple-300 font-mono font-bold">🪖 HK: ${activeZimmet.helmetKitSn}</span>
                    <span class="text-[10px] bg-cyan-100 text-cyan-900 px-2 py-0.5 rounded-lg border border-cyan-300 font-mono font-bold">💨 SA: ${activeZimmet.spareAirSn}</span>
                  </div>
                ` : ''}
              </div>
            </td>
            <td class="py-2.5 px-4 border-r border-slate-200 text-center font-sans font-bold">
              <span class="px-3 py-1 rounded-full text-xs ${
                p.unvan === 'Teknisyen' ? 'bg-amber-100 text-amber-950 border border-amber-300' :
                p.unvan === 'Pilot' ? 'bg-blue-100 text-blue-950 border border-blue-300' : 'bg-slate-100 text-slate-800 border border-slate-300'
              }">
                ${p.unvan === 'Teknisyen' ? '🛠️ Teknisyen' : (p.unvan === 'Pilot' ? '✈️ Pilot' : '👤 ' + (p.unvan || 'Personel'))}
              </span>
            </td>
            <td class="py-2.5 px-2">
              <div class="flex items-center justify-center gap-1.5">
                <!-- ZİMMETİ GÖR BUTONU (User Request) -->
                <button 
                  onclick="openViewZimmetModal('${p.adSoyad}')" 
                  class="px-3 py-1.5 bg-[#0b3d1d] hover:bg-[#072813] text-white rounded-lg text-[10px] font-black uppercase tracking-wider shadow-sm transition active:scale-95 flex items-center gap-1 cursor-pointer"
                  title="Personel üzerindeki tüm zimmet ekipmanlarını gör"
                >
                  <i class="fa-solid fa-eye"></i>
                  <span>Zimmeti Gör</span>
                </button>

                <!-- ZİMMETİ BİTİR BUTONU (User Request) -->
                ${activeZimmet ? `
                  <button 
                    onclick="openFinishModal('${activeZimmet.id}')" 
                    class="px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-[10px] font-black uppercase tracking-wider shadow-sm transition active:scale-95 flex items-center gap-1 cursor-pointer"
                    title="Zimmeti Devret veya Depoya Teslim Et"
                  >
                    <i class="fa-solid fa-arrow-right-arrow-left"></i>
                    <span>Zimmeti Bitir</span>
                  </button>
                ` : `
                  <button 
                    onclick="openNewZimmetForPersonnel('${p.adSoyad}')" 
                    class="px-2.5 py-1.5 bg-slate-100 hover:bg-emerald-50 text-slate-600 hover:text-emerald-800 border border-slate-200 rounded-lg text-[10px] font-bold transition cursor-pointer"
                    title="Bu personele yeni zimmet ver"
                  >
                    + Zimmet Ver
                  </button>
                `}

                <button onclick="promptEditPersonnel('${p.id}')" class="p-1.5 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-lg text-xs cursor-pointer" title="Personel Düzenle">✏️</button>
                <button onclick="deletePersonnelHTML('${p.id}')" class="p-1.5 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded-lg text-xs cursor-pointer" title="Personel Sil">🗑️</button>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    async function promptInlineStockEdit(type, idx) {
      let list = type === 'can_yelegi' ? canYelegiList : (type === 'spare_air' ? spareAirList : helmetKitList);
      if (idx < 0 || idx >= list.length) return;
      const current = list[idx];
      const newNotlar = prompt('Açıklama / Not Giriniz:', current.notlar || current.aciklama || '');
      if (newNotlar === null) return;

      const newPersonel = prompt('Teslim Edilen Personel / Görev Yeri Giriniz (Boşta ise DEPODA yazın):', current.aciklamalar || 'DEPODA');
      if (newPersonel === null) return;
      
      const isDepot = newPersonel.toUpperCase().includes('DEPO') || newPersonel.toUpperCase().includes('BOŞ') || newPersonel === '-';
      const newStatus = isDepot ? 'DEPODA' : 'GÖREV BÖLGESİNDE';

      list[idx] = {
        ...current,
        notlar: newNotlar,
        aciklama: newNotlar,
        aciklamalar: isDepot ? 'DEPODA' : newPersonel.trim().toUpperCase(),
        durum: newStatus,
        color: isDepot ? '#ffffff' : '#92d050',
        tarih: isDepot ? '-' : new Date().toLocaleDateString('tr-TR')
      };

      await fetch('/api/yasam-destek/stock-update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type, items: list })
      });
      alert('✅ Stok verisi başarıyla güncellendi.');
      fetchData();
    }

    async function deleteStockRowHTML(type, idx) {
      if (!confirm('Bu stok kaydını silmek istediğinize emin misiniz?')) return;
      let list = type === 'can_yelegi' ? canYelegiList : (type === 'spare_air' ? spareAirList : helmetKitList);
      list.splice(idx, 1);
      await fetch('/api/yasam-destek/stock-update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type, items: list })
      });
      alert('🗑️ Stok kaydı silindi.');
      fetchData();
    }

    async function deleteZimmetHTML(id) {
      if (!confirm('Bu zimmet kaydını silmek istediğinizden emin misiniz?')) return;
      await fetch('/api/yasam-destek/delete-record', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id })
      });
      alert('🗑️ Zimmet kaydı silindi.');
      fetchData();
    }

    async function promptAddPersonnel() {
      const name = prompt('Eklemek İstediğiniz Personel Adı Soyadı:');
      if (!name || !name.trim()) return;
      const role = prompt('Unvanı Giriniz (Teknisyen / Pilot):', 'Teknisyen');
      if (!role) return;

      await fetch('/api/yasam-destek/personnel-list', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'add',
          item: { adSoyad: name.trim().toUpperCase(), unvan: role.trim() }
        })
      });

      alert('✅ Personel başarıyla eklendi.');
      fetchData();
    }

    async function promptEditPersonnel(id) {
      const item = personnelItems.find(p => p.id === id);
      if (!item) return;
      const newName = prompt('Personel Adı Soyadı:', item.adSoyad);
      if (!newName || !newName.trim()) return;
      const newRole = prompt('Unvanı (Teknisyen / Pilot):', item.unvan || 'Teknisyen');
      if (!newRole) return;

      await fetch('/api/yasam-destek/personnel-list', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'edit',
          item: { id, adSoyad: newName.trim().toUpperCase(), unvan: newRole.trim() }
        })
      });

      alert('✅ Personel bilgisi güncellendi.');
      fetchData();
    }

    async function deletePersonnelHTML(id) {
      if (!confirm('Bu personeli silmek istediğinizden emin misiniz?')) return;
      await fetch('/api/yasam-destek/personnel-list', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'delete', id })
      });
      alert('🗑️ Personel silindi.');
      fetchData();
    }

    function printTutanak(id) {
      const rec = zimmetList.find(r => r.id === id);
      if (!rec) return;
      document.getElementById('prtPersonel').innerText = rec.personelAdi;
      document.getElementById('prtGorevYeri').innerText = rec.gorevYeri;
      document.getElementById('prtBaslangic').innerText = rec.baslangicTarihi;
      document.getElementById('prtBitis').innerText = rec.bitisTarihi || 'DEVAM EDİYOR';
      document.getElementById('prtDurum').innerText = rec.durum;
      document.getElementById('prtDevredilen').innerText = rec.devredilenPersonel || '-';
      document.getElementById('prtTeslimAlan').innerText = rec.teslimAlanPersonel || '-';
      document.getElementById('prtLifeVest').innerText = rec.lifeVestSn || '-';
      document.getElementById('prtSpareAir').innerText = rec.spareAirSn || '-';
      document.getElementById('prtHelmetKit').innerText = rec.helmetKitSn || '-';
      document.getElementById('prtSign1').innerText = rec.personelAdi;
      document.getElementById('prtSign2').innerText = rec.devredilenPersonel && rec.devredilenPersonel !== '-' ? rec.devredilenPersonel : (rec.teslimAlanPersonel || 'Depo Sorumlusu');
      window.print();
    }

    function exportExcel() {
      const wb = XLSX.utils.book_new();

      const ws1 = XLSX.utils.aoa_to_sheet([
        ['S. NO', 'MALZEME ADI', 'PARÇA NO', 'SERİ NO', 'DIŞ NO', 'AÇIKLAMA', 'TESLİM EDİLEN PERSONEL/BÖLGE', 'TARİH / BİTİŞ'],
        ...canYelegiList.map((c, i) => [i + 1, c.malzemeAdi, c.parcaNo, c.seriNo, c.disNo || '-', c.notlar || '-', c.aciklamalar || 'DEPODA', c.tarih || '-'])
      ]);
      XLSX.utils.book_append_sheet(wb, ws1, 'CAN YELEĞİ STOK');

      const ws2 = XLSX.utils.aoa_to_sheet([
        ['S. NO', 'MALZEME ADI', 'PARÇA NO', 'SERİ NO', 'AÇIKLAMA', 'TESLİM EDİLEN PERSONEL/BÖLGE', 'TARİH / BİTİŞ'],
        ...spareAirList.map((s, i) => [i + 1, s.malzemeAdi, s.parcaNo, s.seriNo, s.notlar || '-', s.aciklamalar || 'DEPODA', s.tarih || '-'])
      ]);
      XLSX.utils.book_append_sheet(wb, ws2, 'SPARE AIR STOK');

      const ws3 = XLSX.utils.aoa_to_sheet([
        ['S. NO', 'MALZEME ADI 1', 'PARÇA NO 1', 'SERİ NO 1', 'MALZEME ADI 2', 'PARÇA NO 2', 'KULAKLIK S/N', 'AÇIKLAMA', 'TESLİM EDİLEN PERSONEL/BÖLGE', 'TARİH / BİTİŞ'],
        ...helmetKitList.map((h, i) => [i + 1, h.malzemeAdi1, h.parcaNo1, h.seriNo1 || '-', h.malzemeAdi2, h.parcaNo2, h.seriNo2 || '-', h.notlar || '-', h.aciklamalar || 'DEPODA', h.tarih || '-'])
      ]);
      XLSX.utils.book_append_sheet(wb, ws3, 'HELMET KIT VE DAVID CLARK');

      const ws4 = XLSX.utils.aoa_to_sheet([
        ['SIRA NO', 'PERSONEL ADI SOYADI', 'GÖREV YERİ', 'BAŞLANGIÇ TARİHİ', 'BİTİŞ TARİHİ', 'LIFE VEST S/N', 'SPARE AIR S/N', 'HELMET KIT S/N', 'DURUM', 'DEVREDİLEN PERSONEL', 'TESLİM ALAN', 'AÇIKLAMA'],
        ...zimmetList.map((r, i) => [i + 1, r.personelAdi, r.gorevYeri, r.baslangicTarihi, r.bitisTarihi || '-', r.lifeVestSn, r.spareAirSn, r.helmetKitSn, r.durum, r.devredilenPersonel || '-', r.teslimAlanPersonel || '-', r.notlar || ''])
      ]);
      XLSX.utils.book_append_sheet(wb, ws4, 'GÖREV VE ZİMMETLER');

      const ws5 = XLSX.utils.aoa_to_sheet([
        ['SIRA NO', 'PERSONEL ADI SOYADI', 'ÜNVANI / GÖREVİ'],
        ...personnelItems.map((p, i) => [i + 1, p.adSoyad || '', p.unvan || 'Teknisyen'])
      ]);
      XLSX.utils.book_append_sheet(wb, ws5, 'AT-802 PERSONEL VERİSİ');

      XLSX.writeFile(wb, 'yasam_destek_at802_master.xlsx');
    }

    function handleExcelUploadWithProgress(event) {
      const file = event.target.files?.[0];
      if (!file) return;

      const modal = document.getElementById('excelProgressModal');
      const fileNameEl = document.getElementById('excelProgressFileName');
      const percentEl = document.getElementById('excelProgressPercent');
      const barInner = document.getElementById('excelProgressBarInner');
      const statusText = document.getElementById('excelProgressStatusText');
      const successBox = document.getElementById('excelSuccessAlert');
      const detailText = document.getElementById('excelSuccessDetailText');
      const closeBtn = document.getElementById('excelModalCloseBtn');

      fileNameEl.innerText = file.name;
      percentEl.innerText = '0%';
      barInner.style.width = '0%';
      statusText.innerText = 'Excel dosyası okunuyor...';
      successBox.classList.add('hidden');
      closeBtn.classList.add('hidden');
      modal.classList.remove('hidden');

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          percentEl.innerText = '25%';
          barInner.style.width = '25%';
          statusText.innerText = 'Sayfalar Ayrıştırılıyor...';

          setTimeout(() => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            percentEl.innerText = '60%';
            barInner.style.width = '60%';
            statusText.innerText = 'AT-802 Personel ve Depo Stok Verileri Okunuyor...';

            let parsedPersonnel = [];
            let parsedCanYelegi = [];
            let parsedSpareAir = [];
            let parsedHelmetKit = [];

            workbook.SheetNames.forEach(sheetName => {
              const sheet = workbook.Sheets[sheetName];
              const json = XLSX.utils.sheet_to_json(sheet);
              const upperSheet = sheetName.toUpperCase();

              if (upperSheet.includes('PERSONEL') || upperSheet.includes('AT-802') || upperSheet.includes('TEKNİSYEN') || upperSheet.includes('PİLOT')) {
                json.forEach((r, idx) => {
                  const rowStr = JSON.stringify(r).toUpperCase();
                  if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
                    return;
                  }
                  const name = (r['PERSONEL ADI SOYADI'] || r.adSoyad || r.Personel || r.Name || r['AD SOYAD'] || '').toString().trim().toUpperCase();
                  const role = (r['ÜNVANI / GÖREVİ'] || r.unvan || r.Unvan || r.Role || 'Teknisyen').toString().trim();
                  if (name) parsedPersonnel.push({ id: `p_${Date.now()}_${idx}`, adSoyad: name, unvan: role });
                });
              } else if (upperSheet.includes('YELEK') || upperSheet.includes('VEST') || upperSheet.includes('CAN')) {
                json.forEach((r, idx) => {
                  const rowStr = JSON.stringify(r).toUpperCase();
                  if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
                    return;
                  }
                  const sNo = r['S. NO'] || r['SIRA NO'] || r['SIRA'] || (idx + 1);
                  const malzemeAdi = r['MALZEME ADI'] || r['Malzeme'] || "LIFE WEST - CAN YELEĞİ";
                  const parcaNo = r['PARÇA NO'] || r['Parça No'] || "S-7200-511";
                  const seriNo = (r['SERİ NO'] || r['Seri No'] || r['S/N'] || r['SN'] || r['seriNo'] || r['KULAKLIK S/N'] || '').toString().trim();
                  const disNo = (r['DIŞ NO'] || r['Dış No'] || r['disNo'] || '-').toString().trim();
                  const aciklamalar = (r['AÇIKLAMALAR'] || r['Açıklamalar'] || r['aciklamalar'] || r['AÇIKLAMA'] || r['Açıklama'] || r['Not'] || r['notlar'] || '').toString().trim();
                  const tarih = (r['TARİH / BİTİŞ'] || r['Tarih'] || r['tarih'] || '-').toString().trim();
                  const isDepot = aciklamalar.toUpperCase().includes('DEPO') || aciklamalar.toUpperCase().includes('BOŞ') || !aciklamalar || aciklamalar === '-';
                  const durum = isDepot ? 'DEPODA' : 'GÖREV BÖLGESİNDE';
                  const color = isDepot ? '#ffffff' : '#92d050';
                  if (seriNo) {
                    parsedCanYelegi.push({ sNo, malzemeAdi, parcaNo, seriNo, disNo, aciklamalar: isDepot ? 'DEPODA' : aciklamalar, durum, color, tarih });
                  }
                });
              } else if (upperSheet.includes('SPARE') || upperSheet.includes('AIR') || upperSheet.includes('HAVA') || upperSheet.includes('TÜP')) {
                json.forEach((r, idx) => {
                  const rowStr = JSON.stringify(r).toUpperCase();
                  if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
                    return;
                  }
                  const sNo = r['S. NO'] || r['SIRA NO'] || r['SIRA'] || (idx + 1);
                  const malzemeAdi = r['MALZEME ADI'] || r['Malzeme'] || "SPARE AIR (HEED 3)";
                  const parcaNo = r['PARÇA NO'] || r['Parça No'] || "175-001-CE";
                  const seriNo = (r['SERİ NO'] || r['Seri No'] || r['S/N'] || r['SN'] || r['seriNo'] || '').toString().trim();
                  const aciklamalar = (r['AÇIKLAMALAR'] || r['Açıklamalar'] || r['aciklamalar'] || r['AÇIKLAMA'] || r['Açıklama'] || r['Not'] || r['notlar'] || '').toString().trim();
                  const tarih = (r['TARİH / BİTİŞ'] || r['Tarih'] || r['tarih'] || '-').toString().trim();
                  const isDepot = aciklamalar.toUpperCase().includes('DEPO') || aciklamalar.toUpperCase().includes('BOŞ') || !aciklamalar || aciklamalar === '-';
                  const durum = isDepot ? 'DEPODA' : 'GÖREV BÖLGESİNDE';
                  const color = isDepot ? '#ffffff' : '#92d050';
                  if (seriNo) {
                    parsedSpareAir.push({ sNo, malzemeAdi, parcaNo, seriNo, aciklamalar: isDepot ? 'DEPODA' : aciklamalar, durum, color, tarih });
                  }
                });
              } else if (upperSheet.includes('HELMET') || upperSheet.includes('KASK') || upperSheet.includes('KULAK') || upperSheet.includes('CLARK') || upperSheet.includes('KULAKLIK')) {
                json.forEach((r, idx) => {
                  const rowStr = JSON.stringify(r).toUpperCase();
                  if (rowStr.includes("KASK TAKIMI") || rowStr.includes("KOLTUK SAYISI") || rowStr.includes("DOUBLE HAVA") || rowStr.includes("SINGLE HAVA")) {
                    return;
                  }
                  const sNo = r['S. NO'] || r['SIRA NO'] || r['SIRA'] || (idx + 1);
                  const malzemeAdi1 = r['MALZEME ADI 1'] || r['Malzeme 1'] || "HELMET KIT";
                  const parcaNo1 = r['PARÇA NO 1'] || r['Parça No 1'] || "18852G-01";
                  const seriNo1 = (r['SERİ NO 1'] || r['Seri No 1'] || r['seriNo1'] || '-').toString().trim();
                  const malzemeAdi2 = r['MALZEME ADI 2'] || r['Malzeme 2'] || "DAVID CLARK";
                  const parcaNo2 = r['PARÇA NO 2'] || r['Parça No 2'] || "H-10-13X";
                  const seriNo2 = (r['KULAKLIK S/N'] || r['Seri No 2'] || r['seriNo2'] || r['KULAKLIK S.N.'] || '-').toString().trim();
                  const aciklamalar = (r['AÇIKLAMALAR'] || r['Açıklamalar'] || r['aciklamalar'] || r['AÇIKLAMA'] || r['Açıklama'] || r['Not'] || r['notlar'] || '').toString().trim();
                  const tarih = (r['TARİH / BİTİŞ'] || r['Tarih'] || r['tarih'] || '-').toString().trim();
                  const isDepot = aciklamalar.toUpperCase().includes('DEPO') || aciklamalar.toUpperCase().includes('BOŞ') || !aciklamalar || aciklamalar === '-';
                  const durum = isDepot ? 'DEPODA' : 'GÖREV BÖLGESİNDE';
                  const color = isDepot ? '#ffffff' : '#92d050';
                  parsedHelmetKit.push({ sNo, malzemeAdi1, parcaNo1, seriNo1, malzemeAdi2, parcaNo2, seriNo2, aciklamalar: isDepot ? 'DEPODA' : aciklamalar, durum, color, tarih });
                });
              }
            });

            setTimeout(async () => {
              percentEl.innerText = '90%';
              barInner.style.width = '90%';
              statusText.innerText = 'Server & Portala Kaydediliyor...';

              if (parsedPersonnel.length > 0) {
                await fetch('/api/yasam-destek/personnel-list', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ items: parsedPersonnel })
                });
              }

              if (parsedCanYelegi.length > 0 || parsedSpareAir.length > 0 || parsedHelmetKit.length > 0) {
                await fetch('/api/yasam-destek/upload-excel', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({
                    canYelegi: parsedCanYelegi,
                    spareAir: parsedSpareAir,
                    helmetKit: parsedHelmetKit
                  })
                });
              }

              percentEl.innerText = '100%';
              barInner.style.width = '100%';
              statusText.innerText = 'Yükleme Tamamlandı (%100)';

              successBox.classList.remove('hidden');
              
              let parts = [];
              if (parsedPersonnel.length > 0) parts.push(`${parsedPersonnel.length} personel`);
              if (parsedCanYelegi.length > 0) parts.push(`${parsedCanYelegi.length} can yeleği`);
              if (parsedSpareAir.length > 0) parts.push(`${parsedSpareAir.length} yedek hava`);
              if (parsedHelmetKit.length > 0) parts.push(`${parsedHelmetKit.length} kask kiti`);
              
              detailText.innerText = `Excel dosyasından ${parts.join(', ')} verisi başarıyla okunup "AT-802 YAŞAM DESTEK" sistemine ve depolarına eklendi.`;
              closeBtn.classList.remove('hidden');

              fetchData();
            }, 400);
          }, 300);
        } catch (err) {
          alert('Excel Okuma Hatası: ' + err.message);
          closeExcelProgressModal();
        }
      };
      reader.readAsArrayBuffer(file);
    }

    function closeExcelProgressModal() {
      document.getElementById('excelProgressModal').classList.add('hidden');
      document.getElementById('excelFileInput').value = '';
    }

    function showAutocomplete(inputId, type, value) {
      const suggestDiv = document.getElementById('suggest_' + inputId);
      if (!suggestDiv) return;

      const cleanVal = value.trim().toUpperCase();
      if (!cleanVal) {
        suggestDiv.classList.add('hidden');
        return;
      }

      let dataSource = [];
      if (type === 'personnel') {
        dataSource = personnelItems.map(p => p.adSoyad);
      } else if (type === 'life_vest') {
        dataSource = canYelegiList.filter(c => c.durum !== 'KAYIT SİLME').map(c => c.seriNo);
      } else if (type === 'spare_air') {
        dataSource = spareAirList.filter(s => s.durum !== 'KAYIT SİLME').map(s => s.seriNo);
      } else if (type === 'helmet_kit') {
        dataSource = helmetKitList.map(h => (h.seriNo2 && h.seriNo2 !== '-') ? h.seriNo2 : (h.seriNo1 || h.aciklamalar));
      }

      const matches = dataSource.filter(item => item && item.toUpperCase().includes(cleanVal));

      if (matches.length === 0) {
        suggestDiv.classList.add('hidden');
        return;
      }

      suggestDiv.innerHTML = matches.map(m => `
        <div onclick="selectAutocompleteValue('${inputId}', '${m}')" class="px-4 py-2 hover:bg-emerald-50 hover:text-emerald-950 cursor-pointer text-xs font-black text-slate-800 border-b border-slate-100 text-left transition select-none">
          ${m}
        </div>
      `).join('');
      
      suggestDiv.className = "absolute left-0 right-0 bg-white border border-slate-300 rounded-xl shadow-xl mt-1 max-h-48 overflow-y-auto z-50 divide-y divide-slate-100";
      suggestDiv.classList.remove('hidden');
    }

    function selectAutocompleteValue(inputId, value) {
      const inp = document.getElementById(inputId);
      if (inp) {
        inp.value = value;
        if (inputId === 'inpPersonel') {
          const selP = document.getElementById('selPersonel');
          if (selP) selP.value = value;
        }
        if (inputId === 'inpLifeVest') {
          const selLV = document.getElementById('selLifeVest');
          if (selLV) selLV.value = value;
        }
        if (inputId === 'inpSpareAir') {
          const selSA = document.getElementById('selSpareAir');
          if (selSA) selSA.value = value;
        }
        if (inputId === 'inpHelmetKit') {
          const selHK = document.getElementById('selHelmetKit');
          if (selHK) selHK.value = value;
        }
      }
      const suggestDiv = document.getElementById('suggest_' + inputId);
      if (suggestDiv) {
        suggestDiv.classList.add('hidden');
      }
    }

    // Modal ve Autocomplete dışı tıklamaları dinle
    document.addEventListener('click', function(e) {
      ['inpPersonel', 'inpLifeVest', 'inpSpareAir', 'inpHelmetKit', 'modalDevredilenPersonel'].forEach(id => {
        const suggestDiv = document.getElementById('suggest_' + id);
        const inputEl = document.getElementById(id);
        if (suggestDiv && !suggestDiv.contains(e.target) && e.target !== inputEl) {
          suggestDiv.classList.add('hidden');
        }
      });
    });

    // Sayfa Yüklendiğinde
    fetchData().then(() => {
      switchPublicTab('can_yelegi');
    });
  </script>'''

text = text[:s_idx1] + new_script + text[s_idx2 + len(script_end):]
with open('public/yasam-destek.html', 'w', encoding='utf-8') as f:
    f.write(text)

print('Script updated successfully!')
