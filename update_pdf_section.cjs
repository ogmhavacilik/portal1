const fs = require('fs');

let code = fs.readFileSync('src/App.tsx', 'utf8');

// 1. Update line 10988-10990 in table cell
const targetCell = `const docKey = \`\${(targetRow[2] || targetRow[1] || "").trim().toUpperCase()}\`;
                                                const itemDocs = hangarPdfDocs.filter(d => d.itemKey === docKey);`;

const replaceCell = `const itemDocs = findMatchingDocs(targetRow, hangarPdfDocs);`;

if (code.includes(targetCell)) {
  code = code.replace(targetCell, replaceCell);
  console.log('Successfully replaced targetCell');
} else {
  console.log('targetCell not found directly, checking regex...');
  code = code.replace(/const docKey = `\${\(targetRow\[2\] \|\| targetRow\[1\] \|\| ""\)\.trim\(\)\.toUpperCase\(\)}`;\s*const itemDocs = hangarPdfDocs\.filter\(d => d\.itemKey === docKey\);/, `const itemDocs = findMatchingDocs(targetRow, hangarPdfDocs);`);
}

// 2. Update line 11782 tab badge
code = code.replace(
  `{hangarPdfDocs.filter(d => d.itemKey === (row[2] || row[1] || "").trim().toUpperCase()).length > 0 && (`,
  `{findMatchingDocs(row, hangarPdfDocs).length > 0 && (`
);

code = code.replace(
  `{hangarPdfDocs.filter(d => d.itemKey === (row[2] || row[1] || "").trim().toUpperCase()).length} Evrak Kayıtlı`,
  `{findMatchingDocs(row, hangarPdfDocs).length} Evrak Kayıtlı`
);

// 3. Update line 11938
code = code.replace(
  `const count = hangarPdfDocs.filter(d => d.itemKey === docKey).length;`,
  `const count = findMatchingDocs(row, hangarPdfDocs).length;`
);

fs.writeFileSync('src/App.tsx', code, 'utf8');
console.log('Finished updating App.tsx phase 1');
