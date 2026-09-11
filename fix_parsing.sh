#!/bin/bash
# We will use sed to replace the headerRowIdx detection logic in src/App.tsx

cat << 'INNER_EOF' > temp_replace.js
const fs = require('fs');
let content = fs.readFileSync('src/App.tsx', 'utf8');

// The block to replace:
/*
          let headerRowIdx = 0;
          for (let r = 0; r < Math.min(10, rawRows.length); r++) {
            const rCells = rawRows[r] || [];
            const filledCount = rCells.filter(c => String(c).trim() !== "").length;
            const hasKeywords = rCells.some(c => {
              const s = String(c).toLowerCase();
              return s.includes("sira") || s.includes("no") || s.includes("teçhizat") || s.includes("techizat") || s.includes("malzeme") || s.includes("parca") || s.includes("parça");
            });
            if (filledCount >= 3 && hasKeywords) {
              headerRowIdx = r;
              break;
            }
          }
*/

const oldLogic = `          let headerRowIdx = 0;
          for (let r = 0; r < Math.min(10, rawRows.length); r++) {
            const rCells = rawRows[r] || [];
            const filledCount = rCells.filter(c => String(c).trim() !== "").length;
            const hasKeywords = rCells.some(c => {
              const s = String(c).toLowerCase();
              return s.includes("sira") || s.includes("no") || s.includes("teçhizat") || s.includes("techizat") || s.includes("malzeme") || s.includes("parca") || s.includes("parça");
            });
            if (filledCount >= 3 && hasKeywords) {
              headerRowIdx = r;
              break;
            }
          }`;

const newLogic = `          // Robust Header Row Detection
          let headerRowIdx = 0;
          let maxScore = -1;
          for (let r = 0; r < Math.min(15, rawRows.length); r++) {
            const rCells = rawRows[r] || [];
            let score = 0;
            const rowStr = rCells.map(c => String(c).toLowerCase()).join(" ");
            
            if (rowStr.includes("sıra") || rowStr.includes("sira")) score += 2;
            if (rowStr.includes("teçhizat") || rowStr.includes("techizat") || rowStr.includes("malzeme") || rowStr.includes("araç")) score += 2;
            if (rowStr.includes("parça") || rowStr.includes("parca") || rowStr.includes("p/n")) score += 2;
            if (rowStr.includes("seri") || rowStr.includes("s/n")) score += 2;
            if (rowStr.includes("miktar") || rowStr.includes("kapasite")) score += 2;
            if (rowStr.includes("bulunduğu") || rowStr.includes("lokasyon")) score += 2;
            if (rowStr.includes("durum")) score += 1;
            if (rowStr.includes("bakım") || rowStr.includes("kalibrasyon") || rowStr.includes("kontrol")) score += 2;

            if (score > maxScore) {
              maxScore = score;
              headerRowIdx = r;
            }
          }
          
          if (maxScore < 4) {
            console.warn("Could not confidently find a header row. Falling back to row 0 or 1.");
          }`;

content = content.replace(oldLogic, newLogic);

// We need to also fix the fallbacks, just in case findColIdx fails.
// Let's replace the hardcoded fallbacks with an offset approach, or just trust findColIdx.
// If findColIdx finds the header row properly, it won't fail.

fs.writeFileSync('src/App.tsx', content, 'utf8');
INNER_EOF

node temp_replace.js
