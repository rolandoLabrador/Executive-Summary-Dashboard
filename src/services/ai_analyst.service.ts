import OpenAI from 'openai';

export interface TargetSummaryData {
  type: 'Agent' | 'Dealer';
  name: string;
  code?: string;
  tenureYears: number;
  tenureMonths: number;
  firstContractDate: Date | null;
  seasoningClassification:
    | 'Seasoned Program (≥2 Years)'
    | 'Developing Program (1-2 Years)'
    | 'Emerging / Unseasoned (<1 Year)';
  unearnedReserve: number;
  percentReserveEarned: number;
  rollingContracts: number;
  rollingReserve: number;
  rollingClaims: number;
  rollingLossRatio: number;
  itdContracts: number;
  itdReserve: number;
  itdEarnedReserve: number;
  itdClaims: number;
  itdLossRatio: number;
  topLossCodes: Array<{ code: string; description: string; amount: number }>;
}

export interface PortfolioSummaryData {
  companyName: string;
  asOfDate: Date;
  rollingContracts: number;
  rollingReserve: number;
  rollingEarnedReserve: number;
  rollingClaims: number;
  rollingLossRatio: number;
  rollingUnderwritingProfit: number;
  priorRollingContracts: number;
  priorRollingLossRatio: number;
  itdContracts: number;
  itdEarnedReserve: number;
  itdClaims: number;
  itdLossRatio: number;
  itdUnderwritingProfit: number;
  topAgentRisk?: {
    name: string;
    lossRatio: number;
    paidClaims: number;
    tenureYears?: number;
    firstDate?: Date | null;
  };
  topAgentProfitable?: {
    name: string;
    lossRatio: number;
    contracts: number;
    tenureYears?: number;
    firstDate?: Date | null;
  };
  topLossCode?: { code: string; description: string; amount: number };
}

export class AiAnalystService {
  private openai: OpenAI | null = null;
  private openaiModel: string;
  private geminiApiKey: string | null = null;
  private geminiModel: string;

  constructor() {
    // 1. Google Gemini (Free API Key from Google AI Studio)
    const geminiKey = process.env.GEMINI_API_KEY?.trim();
    if (geminiKey) {
      this.geminiApiKey = geminiKey;
    }
    this.geminiModel = process.env.GEMINI_MODEL?.trim() || 'gemini-3.8-flash';

    // 2. OpenAI / ChatGPT
    const openAiKey = process.env.OPENAI_API_KEY?.trim();
    if (openAiKey) {
      this.openai = new OpenAI({ apiKey: openAiKey, timeout: 15000 });
    }
    this.openaiModel = process.env.OPENAI_MODEL?.trim() || 'gpt-4o-mini';
  }

  /**
   * Generates an objective, factual executive overview (strictly less than 1 paragraph)
   * grounded in tenure, maturity, and hard empirical data.
   */
  async generateExecutiveOverview(target: TargetSummaryData): Promise<string> {
    const tenureStr = target.firstContractDate
      ? `${target.tenureYears} years (${target.tenureMonths} months active since ${target.firstContractDate.toLocaleDateString('en-US', { month: '2-digit', year: 'numeric' })})`
      : `${target.tenureYears} years`;

    const prompt = `You are an objective, non-emotional Actuarial Analyst providing strictly factual warranty intelligence.
STRICT RULES:
- NO subjective opinions, dramatic commentary, or pessimistic adjectives (do NOT use words like "alarming", "catastrophic", "severe distress", or prescriptive advice like "institute surcharge").
- Report hard facts, empirical metrics, dates, and mathematical realities only.
- CONTEXTUALIZE TENURE & INCEPTION-TO-DATE (CRITICAL):
  * If Seasoned (≥2 years, e.g. 5 years): State that with ${target.tenureYears} years of tenure and ${target.percentReserveEarned.toFixed(1)}% of reserves earned, Inception-to-Date metrics carry high statistical credibility. If the ITD loss ratio is unfavorable (>75%), state factually that this represents seasoned adverse experience that warrants attention.
  * If Emerging / New (<1-2 years or low volume): State that with only ${target.tenureMonths} months of operation and $${Math.round(target.unearnedReserve).toLocaleString()} in unearned reserve (${target.percentReserveEarned.toFixed(1)}% earned), early loss ratios reflect small-sample timing volatility and developing loss emergence rather than seasoned performance.
- LENGTH: Strictly LESS than one paragraph (2 to 4 factual sentences).

DATA FOR ${target.type.toUpperCase()}: ${target.name} ${target.code ? `(${target.code})` : ''}
- Program Tenure: ${tenureStr}
- Maturity Classification: ${target.seasoningClassification}
- Reserve Seasoning: ${target.percentReserveEarned.toFixed(1)}% of written reserves earned ($${Math.round(target.itdEarnedReserve).toLocaleString()} earned vs. $${Math.round(target.unearnedReserve).toLocaleString()} unearned)
- Inception-to-Date: ${target.itdContracts.toLocaleString()} total contracts, $${Math.round(target.itdEarnedReserve).toLocaleString()} Earned Reserve, $${Math.round(target.itdClaims).toLocaleString()} Paid Claims, Loss Ratio: ${(target.itdLossRatio * 100).toFixed(1)}%
- Rolling 12-Month: ${target.rollingContracts.toLocaleString()} written contracts, $${Math.round(target.rollingReserve).toLocaleString()} Net Reserve, $${Math.round(target.rollingClaims).toLocaleString()} Claims Paid, Loss Ratio: ${(target.rollingLossRatio * 100).toFixed(1)}%
- Primary Claim Exposure: ${target.topLossCodes[0] ? `${target.topLossCodes[0].code} ($${Math.round(target.topLossCodes[0].amount).toLocaleString()} paid)` : 'None recorded'}`;

    return this.executeAiPrompt(prompt, () => this.generateDeterministicOverview(target));
  }

  /**
   * Generates an objective, factual portfolio overview (strictly less than 1 paragraph)
   * grounded in enterprise-wide tenure, volume, and empirical performance.
   */
  async generatePortfolioOverview(portfolio: PortfolioSummaryData): Promise<string> {
    const prompt = `You are an objective Chief Actuary delivering factual, empirical portfolio reporting for ${portfolio.companyName}.
STRICT RULES:
- NO dramatic opinions, alarmism, or pessimistic adjectives. State empirical figures, mathematical developments, and factual volume breakdowns only.
- Distinguish between mature program trends (statistically credible, high earned reserve) and unseasoned cohort volatility (developing reserves).
- LENGTH: Strictly LESS than one paragraph (2 to 4 factual sentences).

PORTFOLIO METRICS:
- Inception-to-Date: ${portfolio.itdContracts.toLocaleString()} total contracts, $${Math.round(portfolio.itdEarnedReserve).toLocaleString()} Earned Reserve, $${Math.round(portfolio.itdClaims).toLocaleString()} Paid Claims, Cumulative Loss Ratio: ${(portfolio.itdLossRatio * 100).toFixed(1)}%, Cumulative Underwriting Profit: $${Math.round(portfolio.itdUnderwritingProfit).toLocaleString()}
- Rolling 12-Month: ${portfolio.rollingContracts.toLocaleString()} written contracts, $${Math.round(portfolio.rollingReserve).toLocaleString()} Net Reserve, $${Math.round(portfolio.rollingClaims).toLocaleString()} Claims Paid, Loss Ratio: ${(portfolio.rollingLossRatio * 100).toFixed(1)}%, Rolling Underwriting Profit: $${Math.round(portfolio.rollingUnderwritingProfit).toLocaleString()}
- Prior 12-Month Comparison: ${portfolio.priorRollingContracts.toLocaleString()} contracts at ${(portfolio.priorRollingLossRatio * 100).toFixed(1)}% loss ratio
${portfolio.topAgentRisk ? `- Primary Seasoned Program Loss Variance: ${portfolio.topAgentRisk.name} (${portfolio.topAgentRisk.tenureYears ? `${portfolio.topAgentRisk.tenureYears} yrs tenure, ` : ''}${(portfolio.topAgentRisk.lossRatio * 100).toFixed(1)}% loss ratio on $${Math.round(portfolio.topAgentRisk.paidClaims).toLocaleString()} claims)` : ''}
${portfolio.topLossCode ? `- Primary Portfolio Claim Exposure: Loss Code ${portfolio.topLossCode.code} ($${Math.round(portfolio.topLossCode.amount).toLocaleString()} paid)` : ''}`;

    return this.executeAiPrompt(prompt, () =>
      this.generateDeterministicPortfolioOverview(portfolio),
    );
  }

  private async executeAiPrompt(prompt: string, fallbackFn: () => string): Promise<string> {
    // Priority 1: Google Gemini (Free tier available)
    if (this.geminiApiKey) {
      try {
        const geminiOverview = await this.callGemini(prompt);
        if (geminiOverview) {
          return geminiOverview;
        }
      } catch (error) {
        console.warn(`[AiAnalystService] Gemini API call failed:`, error);
      }
    }

    // Priority 2: OpenAI / ChatGPT
    if (this.openai) {
      try {
        const response = await this.openai.chat.completions.create({
          model: this.openaiModel,
          messages: [
            {
              role: 'system',
              content:
                'You are an objective, data-driven actuarial analyst. Deliver strictly factual, non-emotional statements of 2 to 4 sentences without opinions, drama, or pessimism.',
            },
            { role: 'user', content: prompt },
          ],
          temperature: 0.1,
          max_tokens: 350,
        });

        const content = response.choices[0]?.message?.content?.trim();
        if (content) {
          return content;
        }
      } catch (error) {
        console.warn(`[AiAnalystService] OpenAI call failed:`, error);
      }
    }

    return fallbackFn();
  }

  private async callGemini(prompt: string): Promise<string | null> {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${this.geminiModel}:generateContent?key=${this.geminiApiKey}`;

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 25000);

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: controller.signal,
        body: JSON.stringify({
          contents: [
            {
              parts: [
                {
                  text:
                    'You are an objective, data-driven actuarial analyst. Deliver strictly factual, non-emotional statements of 2 to 4 sentences without opinions, drama, or pessimism.\n\n' +
                    prompt,
                },
              ],
            },
          ],
          generationConfig: {
            temperature: 0.1,
            maxOutputTokens: 2048,
            thinkingConfig: {
              thinkingBudget: 150,
            },
          },
        }),
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Gemini HTTP ${response.status}: ${errText}`);
      }

      const data = (await response.json()) as {
        candidates?: Array<{ content?: { parts?: Array<{ text?: string }> } }>;
      };

      const text = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
      return text || null;
    } finally {
      clearTimeout(timeout);
    }
  }

  private generateDeterministicOverview(target: TargetSummaryData): string {
    const lrPercent = (target.itdLossRatio * 100).toFixed(1);
    const rollingLrPercent = (target.rollingLossRatio * 100).toFixed(1);
    const dateStr = target.firstContractDate
      ? target.firstContractDate.toLocaleDateString('en-US', { month: '2-digit', year: 'numeric' })
      : 'inception';

    let maturityContext: string;
    if (target.seasoningClassification === 'Seasoned Program (≥2 Years)') {
      const statusNote =
        target.itdLossRatio > 0.75
          ? 'represents seasoned adverse underwriting experience warranting review.'
          : 'represents seasoned, stable underwriting performance.';
      maturityContext = `Operating for ${target.tenureYears} years (active since ${dateStr}) with ${target.itdContracts.toLocaleString()} contracts and ${target.percentReserveEarned.toFixed(1)}% of reserves earned, Inception-to-Date experience carries high statistical credibility and ${statusNote}`;
    } else {
      maturityContext = `Operating for ${target.tenureYears} years (${target.tenureMonths} months since ${dateStr}) with only ${target.percentReserveEarned.toFixed(1)}% of reserves earned ($${Math.round(target.unearnedReserve).toLocaleString()} unearned), Inception-to-Date loss ratios reflect early-cohort timing volatility rather than seasoned performance.`;
    }

    const topLossStr = target.topLossCodes[0]
      ? ` Primary claim exposure is ${target.topLossCodes[0].code} ($${Math.round(target.topLossCodes[0].amount).toLocaleString()} paid).`
      : '';

    return `${target.name} (${target.code || target.type}) records an Inception-to-Date Earned Loss Ratio of ${lrPercent}% on $${Math.round(target.itdEarnedReserve).toLocaleString()} earned reserve and $${Math.round(target.itdClaims).toLocaleString()} paid claims. Rolling 12-month volume is ${target.rollingContracts.toLocaleString()} written contracts at a ${rollingLrPercent}% loss ratio. ${maturityContext}${topLossStr}`.trim();
  }

  private generateDeterministicPortfolioOverview(portfolio: PortfolioSummaryData): string {
    const lrTrend = portfolio.rollingLossRatio > portfolio.itdLossRatio ? 'above' : 'below';
    const topRisk = portfolio.topAgentRisk
      ? ` Primary seasoned loss variance is driven by ${portfolio.topAgentRisk.name} (${portfolio.topAgentRisk.tenureYears ? `${portfolio.topAgentRisk.tenureYears} yrs tenure, ` : ''}${(portfolio.topAgentRisk.lossRatio * 100).toFixed(1)}% loss ratio on $${Math.round(portfolio.topAgentRisk.paidClaims).toLocaleString()} claims).`
      : '';

    return `${portfolio.companyName}'s portfolio maintains an Inception-to-Date Earned Loss Ratio of ${(portfolio.itdLossRatio * 100).toFixed(1)}% across ${portfolio.itdContracts.toLocaleString()} contracts, with cumulative underwriting profit standing at $${Math.round(portfolio.itdUnderwritingProfit).toLocaleString()}. Over the rolling 12 months, written volume totaled ${portfolio.rollingContracts.toLocaleString()} contracts at a ${(portfolio.rollingLossRatio * 100).toFixed(1)}% loss ratio (${lrTrend} inception levels), generating $${Math.round(portfolio.rollingUnderwritingProfit).toLocaleString()} in rolling underwriting profit.${topRisk}`.trim();
  }
}
