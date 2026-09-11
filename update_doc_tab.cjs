const fs = require('fs');

let code = fs.readFileSync('src/App.tsx', 'utf8');

const targetBlock = `                    {/* Yüklenen Belgeler Tablosu */}
                    {(() => {
                      const docKey = \`\${(row[2] || row[1] || "").trim().toUpperCase()}\`;
                      const itemDocs = hangarPdfDocs.filter(d => d.itemKey === docKey);`;

const newBlock = `                    {/* Yüklenen Belgeler Tablosu & Google Drive Entegrasyonu */}
                    {(() => {
                      const docKey = \`\${(row[2] || row[1] || "").trim().toUpperCase()}\`;
                      const itemDocs = findMatchingDocs(row, hangarPdfDocs);
                      const filteredDriveDocs = hangarPdfDocs.filter(d => {
                        if (!driveDocSearchQuery.trim()) return true;
                        const q = driveDocSearchQuery.toLowerCase();
                        return (d.fileName || "").toLowerCase().includes(q) ||
                               (d.itemKey || "").toLowerCase().includes(q) ||
                               (d.firma || "").toLowerCase().includes(q) ||
                               (d.docType || "").toLowerCase().includes(q);
                      });
                      const displayDocs = docViewFilterTab === 'item' ? itemDocs : filteredDriveDocs;`;

if (!code.includes(targetBlock)) {
  console.error("targetBlock not found!");
  process.exit(1);
}

code = code.replace(targetBlock, newBlock);

// Now update header of the table container
const targetHeader = `<div className="px-5 py-3.5 bg-slate-900 text-white flex items-center justify-between">
                            <div className="flex items-center gap-2">
                              <FileText className="w-4 h-4 text-rose-400" />
                              <h5 className="text-xs font-black uppercase tracking-wider">
                                YÜKLENEN BELGELER ({itemDocs.length})
                              </h5>
                            </div>
                            <span className="text-[10px] font-mono text-slate-400">
                              Tıklayarak tam pencere önizleme yapabilirsiniz
                            </span>
                          </div>`;

const newHeader = `<div className="px-5 py-3 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-3">
                            <div className="flex items-center gap-2">
                              <div className="flex items-center bg-slate-800 p-1 rounded-xl border border-slate-700">
                                <button
                                  type="button"
                                  onClick={() => setDocViewFilterTab('item')}
                                  className={\`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer \${
                                    docViewFilterTab === 'item'
                                      ? 'bg-rose-600 text-white shadow-sm'
                                      : 'text-slate-300 hover:text-white'
                                  }\`}
                                >
                                  Bu Teçhizata Ait Belgeler (\${itemDocs.length})
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setDocViewFilterTab('all_drive')}
                                  className={\`px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer \${
                                    docViewFilterTab === 'all_drive'
                                      ? 'bg-emerald-700 text-white shadow-sm'
                                      : 'text-slate-300 hover:text-white'
                                  }\`}
                                >
                                  Tüm Google Drive Arşivi (\${hangarPdfDocs.length})
                                </button>
                              </div>
                            </div>
                            <div className="flex items-center gap-2">
                              {docViewFilterTab === 'all_drive' && (
                                <input
                                  type="text"
                                  value={driveDocSearchQuery}
                                  onChange={(e) => setDriveDocSearchQuery(e.target.value)}
                                  placeholder="Drive içinde ara..."
                                  className="px-3 py-1 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white placeholder:text-slate-400 focus:outline-none focus:border-emerald-500 w-40 sm:w-56"
                                />
                              )}
                              <button
                                type="button"
                                onClick={handleSyncDrivePdfs}
                                disabled={isSyncingDrivePdfs}
                                className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-600 active:scale-95 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 transition-all cursor-pointer shadow-sm disabled:opacity-50"
                                title="Google Drive'ı tara ve tüm geçmiş PDF evraklarını güncelle"
                              >
                                <RefreshCw className={\`w-3.5 h-3.5 \${isSyncingDrivePdfs ? 'animate-spin' : ''}\`} />
                                <span>{isSyncingDrivePdfs ? 'Taranıyor...' : 'Drive\\'ı Tara'}</span>
                              </button>
                            </div>
                          </div>`;

if (!code.includes(targetHeader)) {
  console.error("targetHeader not found!");
  process.exit(1);
}
code = code.replace(targetHeader, newHeader);

// Update itemDocs.length === 0 check to displayDocs.length === 0
code = code.replace(
  `{itemDocs.length === 0 ? (`,
  `{displayDocs.length === 0 ? (`
);

code = code.replace(
  `<p className="text-xs font-semibold">Henüz bu teçhizata ait yüklenmiş PDF evrak bulunmuyor.</p>`,
  `<p className="text-xs font-semibold">{docViewFilterTab === 'item' ? 'Bu teçhizatla eşleşen kayıtlı evrak bulunamadı.' : 'Google Drive üzerinde PDF evrak bulunamadı.'}</p>`
);

// Update itemDocs.map to displayDocs.map
code = code.replace(
  `{itemDocs.map((doc) => (`,
  `{displayDocs.map((doc) => (`
);

// Update download button href to handle Drive proxy
const targetDownload = `<a
                                            href={doc.fileData}
                                            download={doc.fileName}
                                            className="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-700 font-bold text-[11px] rounded-lg border border-slate-200 transition-all flex items-center gap-1 cursor-pointer"
                                            title="PDF İndir"
                                          >`;

const newDownload = `<a
                                            href={doc.fileData || (doc.driveFileId ? \`/api/pdf-proxy?fileId=\${doc.driveFileId}\` : doc.driveUrl || '#')}
                                            download={doc.fileName}
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            className="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-700 font-bold text-[11px] rounded-lg border border-slate-200 transition-all flex items-center gap-1 cursor-pointer"
                                            title="PDF İndir"
                                          >`;

if (!code.includes(targetDownload)) {
  console.error("targetDownload not found!");
  process.exit(1);
}
code = code.replace(targetDownload, newDownload);

fs.writeFileSync('src/App.tsx', code, 'utf8');
console.log('Successfully updated document tab!');
