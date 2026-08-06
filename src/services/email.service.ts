import * as sendGrid from '@sendgrid/mail';
import { type EmailConfig } from '../models/report.types';

export interface ReportEmail {
  companyName: string;
  asOfDate: Date;
  fileName: string;
  workbook: Buffer;
  dataQualityIssueCount: number;
}

function displayDate(value: Date): string {
  return value.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
}

export class EmailService {
  constructor(private readonly config: EmailConfig) {
    if (config.enabled) sendGrid.setApiKey(config.apiKey);
  }

  async sendReport(report: ReportEmail): Promise<boolean> {
    if (!this.config.enabled) {
      console.log('Email delivery is disabled. Set EMAIL_ENABLED=true after configuring SendGrid.');
      return false;
    }

    const asOf = displayDate(report.asOfDate);
    await sendGrid.send({
      to: this.config.to,
      cc: this.config.cc.length > 0 ? this.config.cc : undefined,
      from: this.config.from,
      subject: `${report.companyName} Executive Report — ${asOf}`,
      text: [
        `The ${report.companyName} executive warranty report through ${asOf} is attached.`,
        `Data-quality issues reported: ${report.dataQualityIssueCount}.`,
        'This report is confidential and contains sanitized operational information.',
      ].join('\n\n'),
      html: [
        `<p>The <strong>${report.companyName}</strong> executive warranty report through ${asOf} is attached.</p>`,
        `<p>Data-quality issues reported: <strong>${report.dataQualityIssueCount}</strong>.</p>`,
        '<p><em>This report is confidential and contains sanitized operational information.</em></p>',
      ].join(''),
      attachments: [
        {
          content: report.workbook.toString('base64'),
          filename: report.fileName,
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          disposition: 'attachment',
        },
      ],
    });
    console.log(`Report emailed to ${this.config.to.join(', ')}.`);
    return true;
  }
}
