const groupMultiLocationRowsHelper = (
  rawRows,
  nameColIdx = 1,
  locColIdx = 5,
  miktarColIdx = 4,
  siraColIdx = 0,
  pnColIdx = 2,
  seriNoColIdx = 3
) => {
  const grouped = [];
  for (let i = 0; i < rawRows.length; i++) {
    const row = [...rawRows[i]];
    const rawSira = (row[siraColIdx] || "").trim();
    const rawName = (row[nameColIdx] || "").trim();
    const rawPn = pnColIdx >= 0 ? (row[pnColIdx] || "").trim() : "";
    const rawSn = seriNoColIdx >= 0 ? (row[seriNoColIdx] || "").trim() : "";
    const rawLoc = locColIdx >= 0 ? (row[locColIdx] || "").trim() : "";
    const rawMiktar = miktarColIdx >= 0 ? (row[miktarColIdx] || "").trim() : "";

    const lastRow = grouped.length > 0 ? grouped[grouped.length - 1] : null;
    const lastName = lastRow ? (lastRow[nameColIdx] || "").trim() : "";
    const lastPn = lastRow && pnColIdx >= 0 ? (lastRow[pnColIdx] || "").trim() : "";
    const lastSira = lastRow ? (lastRow[siraColIdx] || "").trim() : "";

    const isNameEmptyOrDash = !rawName || rawName === "-" || rawName === "--";
    const isSiraEmptyOrDash = !rawSira || rawSira === "-" || rawSira === "--";
    const isSamePn = rawPn !== "" && lastPn !== "" && rawPn.toUpperCase() === lastPn.toUpperCase();
    const isSameName = rawName !== "" && lastName !== "" && rawName.toUpperCase() === lastName.toUpperCase();
    const isSameSira = rawSira !== "" && lastSira !== "" && rawSira === lastSira;

    const isSubLocation =
      !!lastRow &&
      (
        isNameEmptyOrDash ||
        (isSameName && (isSamePn || !rawPn || !lastPn)) ||
        (isSameSira && (isSamePn || isSameName))
      );

    if (isSubLocation && lastRow) {
      const existingLocs = locColIdx >= 0 && lastRow[locColIdx] ? lastRow[locColIdx].split('\n') : [""];
      const existingMiktars = miktarColIdx >= 0 && lastRow[miktarColIdx] ? lastRow[miktarColIdx].split('\n') : ["1"];
      const existingSns = seriNoColIdx >= 0 && lastRow[seriNoColIdx] ? lastRow[seriNoColIdx].split('\n') : [""];

      const targetCount = Math.max(1, existingLocs.length, existingMiktars.length, existingSns.length);
      while (existingLocs.length < targetCount) existingLocs.push("");
      while (existingMiktars.length < targetCount) existingMiktars.push("1");
      while (existingSns.length < targetCount) existingSns.push("");

      const matchingLocIdx = rawLoc
        ? existingLocs.findIndex(l => l.trim().toUpperCase() === rawLoc.trim().toUpperCase())
        : -1;

      if (matchingLocIdx !== -1 && (!rawSn || rawSn === "-" || !existingSns[matchingLocIdx] || existingSns[matchingLocIdx] === "-")) {
        const eVal = Number(existingMiktars[matchingLocIdx]) || 0;
        const nVal = Number(rawMiktar) || 1;
        if (eVal > 0) {
          existingMiktars[matchingLocIdx] = String(eVal + nVal);
        } else {
          existingMiktars[matchingLocIdx] = rawMiktar || "1";
        }
        if (rawSn && rawSn !== "-") {
          existingSns[matchingLocIdx] = existingSns[matchingLocIdx] && existingSns[matchingLocIdx] !== "-" 
            ? `${existingSns[matchingLocIdx]}, ${rawSn}` 
            : rawSn;
        }
      } else {
        existingLocs.push(rawLoc);
        existingMiktars.push(rawMiktar || "1");
        existingSns.push(rawSn || "-");
      }

      lastRow[locColIdx] = existingLocs.join('\n');
      lastRow[miktarColIdx] = existingMiktars.join('\n');
      lastRow[seriNoColIdx] = existingSns.join('\n');
    } else {
      grouped.push(row);
    }
  }
  return grouped;
};
module.exports = { groupMultiLocationRowsHelper };
