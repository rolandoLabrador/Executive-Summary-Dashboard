# AI & Data Strategy Options (PCMI Environment)

Given that the claims administration relies on **PCMI (PCRS)**—a notoriously closed system that is difficult to integrate with via modern APIs—we have to be highly strategic about how we deploy AI and ML. If we cannot easily write data *back* into PCMI in real-time, our solutions must operate either "on top of" the system (via bots) or "downstream" (via reporting and auditing).

Here is an exploration of grounded, margin-focused options that respect the PCMI limitation.

---

## Option 1: Downstream Leakage & Anomaly Auditing (Recommended First Step)
**The Concept:** Instead of trying to stop bad claims in real-time inside PCMI, we catch them mathematically the moment the data hits your MongoDB reporting replica, quantifying the exact dollar amount of "leakage" to present to executives.
*   **How it works:** We use the ML benchmarks (Median + 1.5x Standard Deviation for every Loss Code / Vehicle Make combination) that we generated. During the nightly dashboard run, the system tags every claim paid that day that exceeded the benchmark.
*   **The Action:** We output an "Anomalous Claims" report. The claims team uses this report to either issue chargebacks to the offending dealers or manually adjust the dealer's profile in PCMI to require manual supervisor review for all future claims.
*   **Pros:** Requires **zero** integration with PCMI. Extremely fast to build. Immediately proves ROI by putting a dollar figure on "money lost to overpayment this month."
*   **Cons:** It is retrospective. The money has technically already been paid, meaning you have to claw it back or absorb the loss for that specific claim.

## Option 2: Robotic Process Automation (RPA) for Straight-Through Processing
**The Concept:** If PCMI won't give you an API, you simulate a human. You deploy a "bot" that can see the PCMI screen.
*   **How it works:** Using an RPA tool (like UiPath, Microsoft Power Automate, or a headless browser automation script), a bot logs into PCMI 24/7. It reads the incoming claim amount, loss code, and VIN. It sends that data to our ML model. 
    *   If the model scores the claim as highly safe (e.g., $150 roadside tow), the bot literally clicks the "Approve" button on the PCMI screen.
    *   If the model flags it as anomalous, the bot leaves it in the queue for a human adjuster.
*   **Pros:** Achieves true "Straight-Through Processing" (STP). Drastically reduces claims headcount and operational expenses.
*   **Cons:** RPA is "brittle." If PCMI updates their user interface (e.g., changes the color or location of the "Approve" button), the bot breaks and has to be reprogrammed.

## Option 3: Dynamic Dealer Reserve Tiering & Risk Scoring
**The Concept:** PCMI handles the administration, but *you* dictate the reserve rates and contract pricing. We use ML to grade the dealers based on historical behavior.
*   **How it works:** We build a predictive model that looks at a dealer's first 90 days of behavior (cancellation rate, average claim severity, frequency) and predicts their lifetime Loss Ratio. 
*   **The Action:**
    *   **High-Risk Dealers:** You increase their reserve requirements or lower their commission structure.
    *   **Low-Risk Dealers:** You pitch them on a Reinsurance/DOWC (Dealer Owned Warranty Company) structure, trading your underwriting risk for guaranteed admin fees.
*   **Pros:** Pure margin optimization. Shifts the financial risk onto the bad actors without needing to fight them on individual claims in PCMI. 
*   **Cons:** Requires buy-in from the sales team, as raising rates on bad dealers can cause friction.

## Option 4: Pre-PCMI AI Document Extraction (OCR + LLM)
**The Concept:** The most expensive part of a claim is a human reading a messy PDF Repair Order (RO) from a dealership and manually typing the parts and labor lines into PCMI.
*   **How it works:** Dealerships email or upload the PDF RO. Before a human ever sees it, an AI vision model (like Gemini or Document AI) reads the invoice, extracts the exact Labor Hours, Parts Costs, and diagnostic codes, and standardizes them.
*   **The Action:** The structured data is handed to an RPA bot (or bulk-uploaded via CSV if PCMI allows batch uploads) to create the claim draft automatically.
*   **Pros:** Eliminates massive amounts of manual data entry. Prevents "fat-finger" typing errors that inflate claim costs.
*   **Cons:** Requires setting up a document ingestion pipeline before the data enters PCMI.

---

### Summary
If you are looking for the fastest way to make money without fighting PCMI's architecture:
1.  **Execute Option 1 immediately.** It costs almost nothing to build and highlights the exact amount of money being lost.
2.  **Use the results of Option 1 to justify Option 2.** Once executives see that $50,000 was lost last month to inflated claims, they will gladly fund an RPA bot to start clicking "Reject" automatically.
