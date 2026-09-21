# OmniShield Executive Report

Automated warranty reporting from MongoDB to a sanitized Excel executive dashboard.

## Report contents

- Executive current-month, YTD, rolling-12, and year-over-year comparisons
- Top 20 dealers by rolling-12 net written reserve
- Agent and product dashboards
- Loss-code dashboard with paid-amount KPIs, component descriptions, and a rolling-12 pie chart
- Top 10 vehicle makes by rolling-12 paid claim amount on the Loss Code Dashboard
- Paid-loss-ratio data-bar visualizations
- Monthly trends
- Sanitized contract and paid-claim activity
- Data-quality and reconciliation results
- Controlled metric definitions

The report excludes customer names, contact details, addresses, VINs, the entire
commission section, and dealer/commission/pack component codes.

## Setup

1. Copy `.env.example` to `.env`.
2. Add the read-only MongoDB reporting URI.
3. Confirm the contract collection name. The placeholder is `ContractDatd`, matching
   the supplied MongoDB screenshot.
4. Add any exact component codes that must be excluded to
   `EXCLUDED_COMPONENT_CODES`.

Never commit `.env`.

## Run

```powershell
npm.cmd run build
npm.cmd test
npm.cmd run report -- --as-of 2026-07-31
```

The workbook is written to `output` unless `REPORT_OUTPUT_DIR` is changed.

## Memory Requirements (OOM Issues)

Because this report loads large historical datasets and builds massive uncompressed Excel arrays in memory before saving to disk, it requires significantly more RAM than standard Node.js limits (default ~1.4GB).

The `npm run report` and `npm start` scripts have been explicitly configured with `--max-old-space-size=8192` to allow the V8 engine to use up to 8GB of RAM. If you encounter an `Out of Memory` or `Heap out of memory` crash, ensure that the host machine executing this script has at least 8GB of free memory available.

If executing via CI/CD (like GitHub Actions standard Linux runners which have 7GB RAM), you may experience OS-level OOM kills. We recommend generating the report locally or on a dedicated reporting server.

## Optional SendGrid email delivery

The report is saved locally before email delivery is attempted. Configure these
values in `.env`:

```dotenv
EMAIL_ENABLED=true
SENDGRID_API_KEY=SG.your_real_key
EMAIL_FROM=verified-sender@example.com
EMAIL_TO=recipient1@example.com,recipient2@example.com
EMAIL_CC=
```

`EMAIL_FROM` must be a sender identity verified in SendGrid. Leave
`EMAIL_ENABLED=false` while testing locally or if email delivery is not required.

## Metric rules (Definitions)

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
- **Excluded Components**: Broadly excludes commission section plus components containing DEALER, DLR, COMMISSION, COMM, F&I, or PACK. Additionally, when calculating Reserve, specifically excludes CLIPFEE, PREMIUMTAX, CEDINGFEE, and ADMIN. When calculating Admin, specifically excludes ROADSIDEADMIN and LOANPMT.
- **Claims**: Paid claim/payment records with a non-zero Total Paid Amount.
- **Snapshot Deduplication**: Contract and cancellation snapshots retain the newest record per Contract# and transaction type.
- **Claim Deduplication**: Claim count uses distinct Claim Number. Paid amounts retain the newest snapshot per payment/detail signature.
- **Privacy**: Customer identity, contact, address, and VIN fields are excluded at MongoDB extraction.

These rules are strictly enforced by the data pipeline and are embedded in every generated workbook on the `Definitions` tab.

## GitHub and Actions troubleshooting

If a change made in GitHub is not working or a GitHub Actions check fails:

1. Open the failed workflow under the repository's **Actions** tab and expand the
   failed step to read the complete error message.
2. Update your local branch before making another fix:

   ```powershell
   git switch dev
   git pull --rebase origin dev
   ```

3. Reproduce the same clean installation and checks used by GitHub Actions:

   ```powershell
   npm.cmd ci
   npm.cmd run check
   ```

4. Commit the fix to `dev`, push it, and confirm that Actions passes before merging
   into `main`.
5. If Actions does not start, confirm that `.github/workflows/quality.yml` exists on
   the pushed branch and that GitHub Actions is enabled under **Settings → Actions**.

Do not commit `.env`, MongoDB credentials, SendGrid keys, generated `.xlsx` reports,
`node_modules`, `dist`, or `output`. Store any credentials required by a workflow in
**Settings → Secrets and variables → Actions**. Never paste credentials into workflow
files or GitHub issue comments.

## Scheduling

Use Windows Task Scheduler to invoke the report command with the desired month-end
date. Run the task under a service identity with read-only MongoDB access and write
access only to the report output location.
