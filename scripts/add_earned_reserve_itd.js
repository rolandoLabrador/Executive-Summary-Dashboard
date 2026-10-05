const fs = require('fs');
let code = fs.readFileSync('src/tabs/executive_dashboard.tab.ts', 'utf8');

// Also fix row.getCell(18).alignment in the top table
code = code.replace(
  /row\.getCell\(19\)\.numFmt = PERCENT;\n\s*row\.getCell\(18\)\.alignment = \{ horizontal: 'center' \};/g,
  "row.getCell(19).numFmt = PERCENT;\n      row.getCell(18).alignment = { horizontal: 'center' };\n      row.getCell(19).alignment = { horizontal: 'center' };"
);

// ITD header
code = code.replace(
  /headerRowITD\.getCell\(18\)\.value = 'Earned Loss Ratio';/g,
  "headerRowITD.getCell(18).value = 'Earned Reserve';\n    headerRowITD.getCell(19).value = 'Earned Loss Ratio';"
);

// ITD array
code = code.replace(
  /\[8, 9, 12, 13, 14, 15, 16, 17, 18\]\.forEach\(\(col\) => \{\n\s*const cell = headerRowITD\.getCell/g,
  "[8, 9, 12, 13, 14, 15, 16, 17, 18, 19].forEach((col) => {\n      const cell = headerRowITD.getCell"
);

// ITD row values
code = code.replace(
  /row\.getCell\(18\)\.value = itdDealer \? itdDealer\.earnedLossRatio : 0;/g,
  "row.getCell(18).value = itdDealer ? itdDealer.earnedReserve : 0;\n      row.getCell(18).numFmt = MONEY;\n      \n      row.getCell(19).value = itdDealer ? itdDealer.earnedLossRatio : 0;"
);

// ITD formatting
code = code.replace(
  /row\.getCell\(18\)\.numFmt = PERCENT;\n\s*row\.getCell\(18\)\.alignment = \{ horizontal: 'center' \};/g,
  "row.getCell(19).numFmt = PERCENT;\n      row.getCell(18).alignment = { horizontal: 'center' };\n      row.getCell(19).alignment = { horizontal: 'center' };"
);

fs.writeFileSync('src/tabs/executive_dashboard.tab.ts', code);
