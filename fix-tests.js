const fs = require('fs');

const path = 'tests/report.transformer.test.ts';
let code = fs.readFileSync(path, 'utf8');

if (!code.includes('asAsyncIterable')) {
  code =
    `
async function* asAsyncIterable(arr) {
  for (const item of arr) yield item;
}
` + code;
}

code = code.replace('let model: ReturnType<typeof transformer.transform>;', 'let model: any;');

code = code.replace('before(() => {', 'before(async () => {');
code = code.replace(
  'model = transformer.transform(contracts, cancellations, claims, sampleAudits);',
  'model = await transformer.transform(asAsyncIterable(contracts), asAsyncIterable(cancellations), asAsyncIterable(claims), sampleAudits);',
);

code = code.replace(
  /const report = new ReportTransformer\(config\)\.transform\((.*?),\s*(.*?),\s*(.*?),\s*(.*?)\);/g,
  'const report = await new ReportTransformer(config).transform(asAsyncIterable($1), asAsyncIterable($2), asAsyncIterable($3), $4);',
);

code = code.replace(
  /const report = new ReportTransformer\(config\)\.transform\((.*?),\s*(.*?),\s*(.*?)\);/g,
  'const report = await new ReportTransformer(config).transform(asAsyncIterable($1), asAsyncIterable($2), asAsyncIterable($3));',
);

code = code.replace(/it\('([^']+)', \(\) => {/g, (match, title) => {
  return `it('${title}', async () => {`;
});

fs.writeFileSync(path, code);
console.log('Fixed tests');
