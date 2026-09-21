const fs = require('fs');
let readme = fs.readFileSync('README.md', 'utf8');

const definitions = `## Metric rules (Definitions)

- **Premium**: Calculated as Net Admin + Net Written Reserve.
- **Underwriting Profit**: Calculated as Premium - Claims Paid.
- **Paid Loss Ratio**: Claims paid divided by Premium.
- **Earned Reserve**: Calculated per contract based on the effective date and its specific earning schedule curve up to the report date.
- **Active Contracts**: Distinct contracts whose latest snapshot has ContractStatus A and whose metadata.ActivationDate falls within the reporting period.
- **Cancellations Processed**: Distinct cancellation contracts whose metadata.CancelBillDate falls within the reporting period. This is activity, not a subtraction from the active cohort.
- **Net Written Reserve**: Included written reserve components less included cancelled reserve components.
- **Net Admin**: Included written admin components less included cancelled admin components.
- **Cancellation Timing**: Cancellation activity is recognized exclusively from metadata.CancelBillDate.
- **Contract Timing**: Written contract activity is recognized exclusively from metadata.ActivationDate.
- **Current Month**: Latest fully completed month, compared with the same calendar month in the prior year.
- **Year to Date**: January 1 through the latest completed month, compared with the same prior-year months.
- **Rolling 12 Months**: Latest completed month plus the preceding 11 months, compared with the preceding 12-month period.
- **Prior Full Calendar Year**: January 1 through December 31 of the calendar year immediately before the report as-of year.
- **Inception to Date (ITD)**: All recognized activity from the very first recorded date (inception) up to the latest completed month. ITD Loss Ratio is the total claims paid since inception divided by the ITD Premium.
- **Dealer Ranking**: Top dealers ranked by rolling-12 net written reserve.
- **Excluded Components**: Broadly excludes commission section plus components containing DEALER, DLR, COMMISSION, COMM, F&I, or PACK. Additionally, when calculating Reserve, specifically excludes CLIPFEE, PREMIUMTAX, CEEDINGFEE (CEDINGFEE), and ADMIN.
- **Claims**: Paid claim/payment records with a non-zero Total Paid Amount.
- **Snapshot Deduplication**: Contract and cancellation snapshots retain the newest record per Contract# and transaction type.
- **Claim Deduplication**: Claim count uses distinct Claim Number. Paid amounts retain the newest snapshot per payment/detail signature.
- **Privacy**: Customer identity, contact, address, and VIN fields are excluded at MongoDB extraction.

These rules are strictly enforced by the data pipeline and are embedded in every generated workbook on the \`Definitions\` tab.`;

readme = readme.replace(
  /## Metric rules[\s\S]*?## GitHub and Actions troubleshooting/,
  definitions + '\n\n## GitHub and Actions troubleshooting',
);

fs.writeFileSync('README.md', readme);
console.log('README.md updated successfully.');
