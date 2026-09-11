const fs = require('fs');
let content = fs.readFileSync('src/App.tsx', 'utf8');

// The block to replace in both parseExcelWorksheetToRows and pullTechizatUnitFromDrive

const oldBlock1 = `            const rawSira = getVal(siraColIdx, "");
            const rawName = getVal(nameColIdx, rawRow[1] ? String(rawRow[1]).trim() : (rawRow[0] ? String(rawRow[0]).trim() : ""));
            const rawPn = getVal(pnColIdx, rawRow[2] ? String(rawRow[2]).trim() : "");
            const rawSn = getVal(snColIdx, rawRow[3] ? String(rawRow[3]).trim() : "-");
            const rawMiktar = getVal(miktarColIdx, rawRow[4] ? String(rawRow[4]).trim() : "1");
            const rawLoc = getVal(locColIdx, rawRow[5] ? String(rawRow[5]).trim() : "");`;

const oldBlock2 = `            const rawSira = getVal(siraColIdx, "");
            const rawName = getVal(nameColIdx, rawRow[1] ? String(rawRow[1]).trim() : "");
            const rawPn = getVal(pnColIdx, rawRow[2] ? String(rawRow[2]).trim() : "");
            const rawSn = getVal(snColIdx, rawRow[3] ? String(rawRow[3]).trim() : "-");
            const rawMiktar = getVal(miktarColIdx, rawRow[4] ? String(rawRow[4]).trim() : "1");
            const rawLoc = getVal(locColIdx, rawRow[5] ? String(rawRow[5]).trim() : "");`;

const newBlock = `            // Dynamic fallbacks to avoid hardcoded indices
            // If findColIdx failed for everything, we use the first few non-empty columns.
            const nonEmptyCols = rawRow.map((c, i) => ({ val: String(c || '').trim(), idx: i })).filter(c => c.val !== "");
            const guessIdx = (order) => nonEmptyCols.length > order ? nonEmptyCols[order].idx : -1;
            
            const rawSira = getVal(siraColIdx, getVal(guessIdx(0), ""));
            const rawName = getVal(nameColIdx, getVal(guessIdx(1), ""));
            const rawPn = getVal(pnColIdx, getVal(guessIdx(2), ""));
            const rawSn = getVal(snColIdx, getVal(guessIdx(3), "-"));
            const rawMiktar = getVal(miktarColIdx, getVal(guessIdx(4), "1"));
            const rawLoc = getVal(locColIdx, getVal(guessIdx(5), ""));`;

content = content.replace(oldBlock1, newBlock);
content = content.replace(oldBlock2, newBlock);

fs.writeFileSync('src/App.tsx', content, 'utf8');
