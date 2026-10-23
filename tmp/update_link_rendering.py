import re

path = '/app/applet/src/components/OlayTakipCizelgesiModal.tsx'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = r"\{rawLinkVal\.split\('\\n'\)\.map\(\(lnk, lIdx\) => \{.*?\)\}\)"
replacement = """{rawLinkVal.split('\\n').map((lnk, lIdx) => {
                                          const u = lnk.trim();
                                          if (!u) return null;
                                          const match = u.match(/^(.*?)\\s*\\((https?:\\/\\/\\S+)\\)$/);
                                          const displayTitle = match ? match[1].trim() : u;
                                          const targetUrl = match ? match[2].trim() : (u.startsWith('http') ? u : `https://${u}`);

                                          return (
                                            <div key={lIdx} className="inline-flex items-center gap-1">
                                              <a
                                                href={targetUrl}
                                                target="_blank"
                                                rel="noopener noreferrer"
                                                className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md bg-blue-50 hover:bg-blue-100 text-blue-800 hover:text-blue-950 border border-blue-200 text-[11px] font-semibold transition-all shadow-2xs max-w-[220px] truncate"
                                                title={`${displayTitle} (${targetUrl}) - Yeni sekmede açmak için tıklayın`}
                                              >
                                                <ExternalLink className="w-3 h-3 text-blue-600 shrink-0" />
                                                <span className="truncate">{displayTitle}</span>
                                              </a>
                                            </div>
                                          );
                                        })}"""

new_content = re.sub(pattern, replacement, content, flags=re.DOTALL)
if new_content != content:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('SUCCESS REGEX')
else:
    print('REGEX NOT MATCHED')
