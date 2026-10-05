const fs = require('fs');
let code = fs.readFileSync('src/tabs/executive_dashboard.tab.ts', 'utf8');

// Update header
code = code.replace(
  /headerRow\.getCell\(18\)\.value = 'Earned Loss Ratio';/,
  "headerRow.getCell(18).value = 'Earned Reserve';\n    headerRow.getCell(19).value = 'Earned Loss Ratio';"
);

// Update coloring array
code = code.replace(
  /\[8, 9, 12, 13, 14, 15, 16, 17, 18\]\.forEach\(\(col\) => \{/,
  '[8, 9, 12, 13, 14, 15, 16, 17, 18, 19].forEach((col) => {'
);

// Update row values
code = code.replace(
  /row\.getCell\(18\)\.value = dealer\.earnedLossRatio;/,
  "row.getCell(18).value = dealer.earnedReserve;\n      row.getCell(18).numFmt = MONEY;\n\n      row.getCell(19).value = dealer.earnedLossRatio;"
);

// Update percentage formatting
code = code.replace(
  /row\.getCell\(18\)\.numFmt = PERCENT;/,
  'row.getCell(19).numFmt = PERCENT;'
);

// Update conditional formatting rule
code = code.replace(
  /const lossRatioCell = row\.getCell\(18\);/,
  'const lossRatioCell = row.getCell(19);'
);

fs.writeFileSync('src/tabs/executive_dashboard.tab.ts', code);
