const fs = require('fs');

let excelCode = fs.readFileSync('src/services/excel.service.ts', 'utf8');

excelCode = excelCode.replace(/import \{\s*type type LossCodeMetric,\s*type MetricValues,\s*type type ReportConfig,\s*type ReportModel,\s*type \} from '\.\.\/models\/report\.types';/, 
  "import { type LossCodeMetric, type MetricValues, type ReportConfig, type ReportModel } from '../models/report.types';"
);

// One more attempt in case the regex slightly missed it:
const startIdx = excelCode.indexOf('import {\r\n  type type LossCodeMetric,');
if (startIdx > -1) {
  const endIdx = excelCode.indexOf("} from '../models/report.types';", startIdx);
  if (endIdx > -1) {
    excelCode = excelCode.substring(0, startIdx) + 
      "import { type LossCodeMetric, type MetricValues, type ReportConfig, type ReportModel } from '../models/report.types';" +
      excelCode.substring(endIdx + 32);
  }
}

// Check for \n instead of \r\n
const startIdx2 = excelCode.indexOf('import {\n  type type LossCodeMetric,');
if (startIdx2 > -1) {
  const endIdx2 = excelCode.indexOf("} from '../models/report.types';", startIdx2);
  if (endIdx2 > -1) {
    excelCode = excelCode.substring(0, startIdx2) + 
      "import { type LossCodeMetric, type MetricValues, type ReportConfig, type ReportModel } from '../models/report.types';" +
      excelCode.substring(endIdx2 + 32);
  }
}

fs.writeFileSync('src/services/excel.service.ts', excelCode);
console.log('Fixed import syntax!');
