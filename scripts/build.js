const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, TableOfContents, PageBreak,
  Header, Footer, PageNumber, LevelFormat, convertInchesToTwip, VerticalAlign, ImageRun
} = require("docx");
const fs = require("fs");
const CHART_RATIOS = {
  "fig1_1_churn_distribution.png": 0.9606,
  "fig1_2_churn_by_contract.png": 0.7222,
  "fig1_3_churn_by_service.png": 0.6217,
  "fig1_4_tenure_charges.png": 0.6957,
  "fig6_1_timeline.png": 0.4493,
  "fig8_1_model_performance.png": 0.5846,
};
function chartImage(path, widthIn) {
  const base = path.split("/").pop();
  const ratio = CHART_RATIOS[base] || 0.65;
  const h = widthIn * ratio;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 40 },
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(path),
      transformation: { width: Math.round(widthIn * 96), height: Math.round(h * 96) },
    })],
  });
}

const W = 12240, H = 15840; // US Letter
const CONTENT_W = 9360; // 6.5in usable width in twips (with 1in margins)

// ---------- helpers ----------
const COLORS = {
  navy: "1E293B",
  cyan: "0891B2",
  indigo: "4F46E5",
  emerald: "059669",
  rose: "E11D48",
  amber: "D97706",
  lightgrey: "F1F5F9",
  midgrey: "64748B",
  white: "FFFFFF",
};

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    border: { bottom: { color: COLORS.cyan, space: 4, style: BorderStyle.SINGLE, size: 12 } },
    children: [new TextRun({ text, bold: true, color: COLORS.navy, size: 30 })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 260, after: 140 },
    children: [new TextRun({ text, bold: true, color: COLORS.cyan, size: 24 })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 100 },
    children: [new TextRun({ text, bold: true, color: COLORS.indigo, size: 21 })],
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 300 },
    children: [new TextRun({ text, size: 22, color: "1E293B", ...opts })],
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullet-list", level },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22 })],
  });
}
function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "num-list", level },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22 })],
  });
}
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 240 },
    children: [new TextRun({ text, italics: true, size: 18, color: COLORS.midgrey })],
  });
}

function cell(text, { bold = false, color = "1E293B", shade = null, width, align = AlignmentType.LEFT, size = 20 } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, color: "auto", fill: shade } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, color, size })],
    })],
  });
}

function makeTable(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((htext, i) => cell(htext, { bold: true, color: COLORS.white, shade: COLORS.navy, width: widths[i], align: AlignmentType.CENTER })),
      }),
      ...rows.map((r, idx) => new TableRow({
        children: r.map((val, i) => cell(val, { width: widths[i], shade: idx % 2 === 0 ? COLORS.lightgrey : COLORS.white })),
      })),
    ],
  });
}

function statBox(label, value, note, color) {
  return new TableCell({
    width: { size: CONTENT_W / 4, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "F8FAFC" },
    margins: { top: 150, bottom: 150, left: 150, right: 150 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
    },
    children: [
      new Paragraph({ children: [new TextRun({ text: label, size: 16, color: COLORS.midgrey })] }),
      new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: value, bold: true, size: 30, color })] }),
      new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: note, size: 16, color: COLORS.midgrey })] }),
    ],
  });
}

// ---------- document ----------
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullet-list", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 260 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 820, hanging: 260 } } } },
      ]},
      { reference: "num-list", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 260 } } } },
      ]},
    ],
  },
  sections: [
    // ===================== TITLE PAGE =====================
    {
      properties: { page: { size: { width: W, height: H }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [
        new Paragraph({ spacing: { before: 1600 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "PROJECT PLANNING & STRATEGY REPORT", bold: true, size: 26, color: COLORS.cyan })] }),
        new Paragraph({ spacing: { before: 300 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Proactive Customer Retention via an End-to-End", bold: true, size: 44, color: COLORS.navy })] }),
        new Paragraph({ spacing: { before: 100 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Machine Learning Churn Prediction Pipeline", bold: true, size: 44, color: COLORS.navy })] }),
        new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "A Strategic Roadmap for Telecommunications Churn Risk Mitigation", italics: true, size: 24, color: COLORS.midgrey })] }),
        new Paragraph({ spacing: { before: 900 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Week 1 Deliverable — Project Planning and Strategy for ML Pipeline", size: 22, color: COLORS.indigo, bold: true })] }),
        new Paragraph({ spacing: { before: 1400 }, alignment: AlignmentType.CENTER,
          border: { top: { color: "CBD5E1", space: 10, style: BorderStyle.SINGLE, size: 6 } },
          children: [new TextRun({ text: "Dataset: Telco Customer Churn (IBM Sample Dataset) — 7,043 records, 21 attributes", size: 20, color: COLORS.midgrey })] }),
        new Paragraph({ spacing: { before: 100 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Regulatory frame of reference: Digital Personal Data Protection (DPDP) Act, 2023", size: 20, color: COLORS.midgrey })] }),
        new Paragraph({ spacing: { before: 100 }, alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Prepared by: Haddy", size: 20, color: COLORS.midgrey })] }),
      ],
    },
    // ===================== TOC PAGE =====================
    {
      properties: { page: { size: { width: W, height: H }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [
        h1("Table of Contents"),
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },
    // ===================== MAIN BODY =====================
    {
      properties: { page: { size: { width: W, height: H }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Telecom Churn ML Pipeline — Project Plan", size: 16, color: COLORS.midgrey })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 16, color: COLORS.midgrey }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: COLORS.midgrey })] })] }) },
      children: [

        // 1. INTRODUCTION
        h1("1. Introduction"),
        body("Customer churn — the loss of subscribers to competitors — is one of the highest-cost problems in the telecommunications sector. Acquiring a new subscriber typically costs several times more than retaining an existing one, so even small improvements in early churn detection translate directly into revenue protection. This report sets out the project plan for building an end-to-end machine learning pipeline that predicts, in advance, which subscribers are at high risk of churning, so that retention teams can intervene before the customer leaves."),
        body("The plan is grounded in a real, publicly available dataset (the IBM Telco Customer Churn dataset, 7,043 subscriber records across 21 attributes) rather than a hypothetical scenario, so that every phase — from data exploration to deployment — is tied to concrete, checkable numbers. The supporting infographic referenced throughout this report summarises the data findings and the target system architecture."),
        h2("1.1 Scope of this document"),
        body("This is a planning document, not a build log. It defines the problem, the approach, the schedule, the resourcing, and the risks. Model training and code implementation are addressed only at the level needed to plan for them (what libraries, what compute, what time); the actual notebooks and scripts are separate deliverables in later weeks."),

        // 2. PROBLEM STATEMENT
        h1("2. Problem Statement"),
        h2("2.1 The business problem"),
        body("A telecom operator needs to identify, ahead of time, which currently-active subscribers are likely to cancel their service within the next billing cycle, so that a retention team can offer a targeted intervention (discount, plan change, proactive support call) before the subscriber leaves. Today this identification is largely reactive — the company finds out a customer has churned only after cancellation. The objective is to convert this into a proactive process driven by a predictive model."),
        h3("2.1.1 Illustrative unit economics"),
        body("The dataset itself has no cost or lifetime-value figures attached, so the numbers below are a worked illustration using round, commonly-cited telecom benchmarks — not derived from this data — to show why the recall/precision targets in Sections 2.4 and 8.2 matter financially, not just statistically. Real figures should replace these before the business case is finalised."),
        bullet("Average monthly revenue per churned subscriber: ~$65 (approx. the dataset's mean MonthlyCharges for churners)."),
        bullet("Assumed average remaining tenure lost per unprevented churn: 12 months → ~$780 revenue at risk per churner."),
        bullet("Assumed cost of a retention intervention (discount + agent time): ~$50 per contacted customer."),
        bullet("On a monthly cohort of 1,000 active subscribers, roughly 265 are actual churn risks (26.5% base rate, Table 5.1)."),
        bullet("At 80% recall / 70% precision: ~212 true churners caught (0.80 × 265), requiring ~303 total contacts (212 ÷ 0.70) to hit that precision, of which ~91 are false positives."),
        bullet("Rough monthly value: (212 saved × $780 revenue at risk) − (303 contacts × $50 cost) ≈ $165,360 − $15,150 ≈ $150,210 net, before accounting for the fraction of contacted customers who would have stayed anyway even without intervention (a save-rate assumption this illustration does not attempt to estimate)."),
        body("This is a directional sanity check, not a financial forecast — it exists to show that the precision floor in Section 8.2 has a real cost attached (each false positive costs ~$50), and that the business case is sensitive to intervention cost and save-rate assumptions that need real numbers from the business, not this report."),
        h2("2.2 Inputs available"),
        body("The dataset provides one row per subscriber, with:"),
        bullet("Demographic attributes: gender, senior citizen flag, partner and dependent status."),
        bullet("Account attributes: tenure (months with the company), contract type, paperless billing, payment method, monthly charges, total charges."),
        bullet("Service attributes: phone service, multiple lines, internet service type (DSL / fibre / none), online security, online backup, device protection, tech support, streaming TV, streaming movies."),
        bullet("Target label: Churn (Yes/No)."),
        h2("2.3 Key challenges"),
        bullet("Class imbalance — only 26.5% of records are churn cases (1,869 of 7,043), so a naive model can score well on accuracy while missing most actual churners. This has to be corrected for in the pipeline, not ignored."),
        bullet("Mixed data types — the dataset mixes numeric fields (tenure, charges) with categorical fields (contract, payment method, service add-ons), which need different preprocessing treatment."),
        bullet("Data leakage risk — TotalCharges is mathematically derived from tenure and MonthlyCharges; any preprocessing or resampling step must be fit on the training fold only, never on the full dataset, or the reported performance will be optimistic and wrong in production."),
        bullet("Interpretability requirement — a churn prediction that a retention agent cannot explain to a customer is of limited practical use, so the model needs an explainability layer, not just a probability score."),
        bullet("Regulatory constraint — subscriber data of this kind falls under India's DPDP Act, 2023, which affects how data can be stored, processed and used for automated decisions, and this has to be planned for, not bolted on afterward."),
        h2("2.4 Expected outcome"),
        body("A trained classification model that flags at-risk subscribers with a recall of at least 80% on the churn class (i.e. it catches at least 8 in 10 subscribers who are actually going to churn), packaged behind an API that a CRM system can call, with SHAP-based explanations attached to each prediction so a retention agent knows why a customer was flagged."),

        // 3. OBJECTIVES
        h1("3. Objectives"),
        numbered("Explore and quantify the churn drivers in the dataset (contract type, service bundle, tenure, pricing) before any modelling begins, so that modelling decisions are evidence-based rather than assumed."),
        numbered("Build a leakage-safe preprocessing pipeline (encoding, scaling, class-imbalance correction) using scikit-learn Pipeline objects, so preprocessing steps are reproducible and cannot silently see test data."),
        numbered("Train and benchmark at least three model families (logistic regression as baseline, random forest, gradient-boosted trees) and select the best performer against a recall-first metric, since missing a churner is more costly than a false alarm."),
        numbered("Reach a minimum recall of 80% on the churn class on held-out data, while enforcing a precision floor of 70% (equivalently, a minimum F1-score of ~74%) so retention outreach is not wasted on excessive false positives — see Section 8.2 for the reasoning behind these specific numbers."),
        numbered("Attach an explainability layer (SHAP) so predictions are auditable and usable by non-technical staff."),
        numbered("Package the model behind a served API and define a monitoring/drift plan so performance degradation after deployment is caught early, not discovered months later in falling retention numbers."),

        // 4. LITERATURE REVIEW / RESEARCH
        h1("4. Research & Ideation"),
        body("A brief review of publicly documented churn-modelling practice in telecom informed the approach taken here:"),
        bullet("Industry benchmarks consistently show month-to-month contracts and fibre-optic internet plans as the strongest churn predictors, which matches what is found in this dataset (Section 5) — this cross-check gives confidence the dataset is representative rather than an artificial toy set."),
        bullet("Gradient-boosted tree ensembles (XGBoost, LightGBM) are the commonly reported best performers for tabular churn data, ahead of both plain logistic regression and untuned random forests, because they handle non-linear interactions between tenure, price and contract type well."),
        bullet("SMOTE (Synthetic Minority Oversampling) is the standard, well-documented technique for the class-imbalance problem seen here, provided it is applied only inside the training fold to avoid leaking synthetic points into the test set."),
        bullet("SHAP is the current standard for explaining tree-based model outputs at both the global (which features matter overall) and local (why this specific customer was flagged) level, which fits the requirement that retention agents be able to justify an intervention."),
        body("This is a stated basis for the modelling choices in Section 8, not a claim of exhaustive academic review. Secondary literature insights have been cross-verified against baseline empirical distributions from the IBM dataset to ensure operational validity — the contract-type and fibre-vs-DSL churn patterns found in Section 5 independently corroborate the general industry pattern cited here, which is the main check available at the planning stage."),

        // 5. DATA EXPLORATION FINDINGS
        h1("5. Data Exploration: Key Findings"),
        body("Before committing to a modelling approach, the dataset was profiled to confirm where the churn signal actually lives. Four findings, summarised below, directly shape the plan in later sections."),

        h2("5.1 Class balance"),
        body("Of 7,043 subscriber records, 5,174 (73.5%) are retained customers and 1,869 (26.5%) are churners. This confirms class imbalance is real and non-trivial — it is not a small correction, it is a first-order design constraint for the modelling phase."),
        makeTable(
          ["Class", "Count", "Share"],
          [["Retained (No)", "5,174", "73.5%"], ["Churned (Yes)", "1,869", "26.5%"], ["Total", "7,043", "100%"]],
          [4000, 2680, 2680]
        ),
        caption("Table 5.1 — Baseline churn class distribution (source: Telco Customer Churn dataset)."),
        chartImage("/home/claude/proj/fig1_1_churn_distribution.png", 4.2),
        caption("Figure 5.1 — Churn class distribution, generated directly from the uploaded dataset."),

        h2("5.2 Churn by contract type"),
        body("Contract type is the single strongest lever identified in the data:"),
        makeTable(
          ["Contract type", "Churn rate"],
          [["Month-to-month", "42.7%"], ["One year", "11.2%"], ["Two year", "2.8%"]],
          [5680, 3680]
        ),
        caption("Table 5.2 — Churn rate by contract term."),
        chartImage("/home/claude/proj/fig1_2_churn_by_contract.png", 4.6),
        caption("Figure 5.2 — Churn rate by contract term, computed from the dataset."),
        body("A month-to-month subscriber is roughly 15 times more likely to churn than a two-year contract holder. This alone justifies treating contract type as a headline feature and a headline retention lever (e.g. incentivising contract upgrades), independent of anything the model later learns."),

        h2("5.3 Churn by service type and support add-ons"),
        makeTable(
          ["Segment", "Churn rate"],
          [
            ["Fibre optic internet", "41.9%"],
            ["DSL internet", "19.0%"],
            ["No tech support add-on", "41.6%"],
            ["Tech support add-on subscribed", "15.2%"],
          ],
          [5680, 3680]
        ),
        caption("Table 5.3 — Churn rate by internet technology and tech-support subscription."),
        chartImage("/home/claude/proj/fig1_3_churn_by_service.png", 5.0),
        caption("Figure 5.3 — Churn rate by internet service type and tech-support subscription."),
        body("Fibre customers churn more than twice as often as DSL customers, most plausibly a pricing-sensitivity effect rather than a technology-quality effect, since fibre plans in this dataset carry materially higher monthly charges. Tech support subscription is associated with a churn rate roughly a third of customers without it — this is a candidate for a low-cost retention action to test in Phase C (Section 10)."),

        h2("5.4 Tenure and pricing risk cluster"),
        body("Cross-plotting tenure against monthly charges shows churners concentrated in the first 12 months of the relationship at charges above roughly $70/month, while retained customers spread across longer tenures and a wider price range. This says the highest-risk window is the first year, and price sensitivity in that window is the compounding factor — a new customer on a high-priced plan is the profile most worth intervening on early."),
        chartImage("/home/claude/proj/fig1_4_tenure_charges.png", 5.0),
        caption("Figure 5.4 — Tenure vs. monthly charges, actual subscriber-level scatter (500-point sample per class)."),
        body("These four findings jointly argue for contract type, tenure, internet service type, monthly charges and tech-support status as the priority features for the model — a judgement call, not a data mining recipe."),

        // 6. PLANNING & TIMELINE
        h1("6. Planning & Timeline"),
        body("The project is phased across 15 weeks in five three-week blocks. Each phase has a hard dependency on the one before it — the plan does not assume phases can run in parallel, because model training cannot start before leakage-safe data pipelines exist, and deployment cannot start before a model clears the recall bar."),
        makeTable(
          ["Phase", "Weeks", "Focus", "Key deliverable"],
          [
            ["1. Legal & Data Protection", "W1–W3", "DPIA, consent-basis review, data governance sign-off", "Approved Data Protection Impact Assessment"],
            ["2. Data Pipeline", "W4–W6", "Ingestion, cleaning, leakage-safe preprocessing, SMOTE", "Reproducible scikit-learn Pipeline"],
            ["3. Model Training & Tuning", "W7–W9", "Baseline models, XGBoost, Optuna tuning, SHAP", "Model meeting ≥80% recall target"],
            ["4. MLOps & API", "W10–W12", "MLflow registry, FastAPI service, Docker/Kubernetes packaging", "Containerised prediction API"],
            ["5. Deployment & Monitoring", "W13–W15", "Shadow deployment, A/B test, drift monitoring setup", "Production rollout decision"],
          ],
          [2600, 1400, 3800, 1560]
        ),
        caption("Table 6.1 — Five-phase, fifteen-week execution timeline."),
        chartImage("/home/claude/proj/fig6_1_timeline.png", 5.5),
        caption("Figure 6.1 — Fifteen-week phase timeline, visualised."),

        h2("6.1 Critical path"),
        body("The DPIA sign-off in Phase 1 is the single hardest dependency: no subscriber data may be processed for model training until it clears, so any legal delay pushes the entire remaining schedule by the same number of weeks — this is the one risk in the plan with no workaround. Within Phase 3, hyperparameter tuning is dependent on the SMOTE-corrected training pipeline from Phase 2 being finalised and frozen; retraining the pipeline mid-tuning would invalidate tuning results already collected."),
        h2("6.2 Milestones"),
        bullet("End of W3: DPIA approved, data access granted — gating milestone for the whole project."),
        bullet("End of W6: Clean, leakage-tested preprocessing pipeline with class-imbalance handling in place."),
        bullet("End of W9: Final model selected, recall ≥80% on held-out test data, SHAP explainability integrated."),
        bullet("End of W12: API deployed to a staging environment, load-tested for sub-100ms response time."),
        bullet("End of W15: Go/no-go decision on full production rollout, based on A/B test results from Phase C."),

        // 7. RESOURCE REQUIREMENTS
        h1("7. Resource Requirements"),
        h2("7.1 Team and role allocation"),
        makeTable(
          ["Role", "Responsibility", "Allocation", "Primary tools"],
          [
            ["Lead AI/ML Systems Engineer", "MLOps architecture, orchestration, deployment", "100%", "FastAPI, Docker, Kubernetes"],
            ["Data Scientists (x2)", "Feature engineering, model training, tuning, SHAP", "100%", "scikit-learn, XGBoost, Optuna"],
            ["Data Engineer", "Ingestion, preprocessing at scale, version control", "100%", "PySpark, Kafka, DVC"],
            ["Data Protection Officer", "DPDP Act compliance, DPIA, consent verification", "25%", "Legal / compliance review"],
          ],
          [2680, 3600, 1400, 1680]
        ),
        caption("Table 7.1 — Team composition and allocation."),
        body("The Data Protection Officer's 25% allocation is front-loaded almost entirely into Phase 1 — treating it as an even 25% across all fifteen weeks would understate the actual Phase 1 workload and is a planning trap worth calling out explicitly."),

        h2("7.2 Software and infrastructure"),
        bullet("Languages / libraries: Python, pandas, scikit-learn, XGBoost, Optuna, SHAP, imbalanced-learn (SMOTE)."),
        bullet("Data infrastructure: Kafka for streaming ingestion, Data Version Control (DVC) backed by S3-compatible object storage."),
        bullet("Experiment tracking: MLflow with a PostgreSQL backing store, for run comparison and model staging."),
        bullet("Serving: FastAPI application, containerised in Docker, orchestrated on Kubernetes with horizontal pod autoscaling; Airflow for scheduled retraining DAGs."),
        bullet("Compute: GPU is not required for this dataset size (7,043 rows); standard CPU compute is sufficient for training, but the API-serving layer should be sized for concurrent request load, not training load."),
        body("Note on stack sizing: PySpark and Kafka are specified above for enterprise-scale streaming parity — i.e. what the production system would run on at a real operator's full subscriber volume. At the current dataset size (7,043 records), local prototyping and all Phase 2–3 development work in this plan actually runs on plain pandas and scikit-learn Pipeline objects; PySpark/Kafka only become necessary once ingestion moves to live, high-volume subscriber streams in Phase 4–5. This distinction matters for resourcing — provisioning a Spark cluster for a 7,043-row prototype would be over-engineering."),

        // 8. MODEL BENCHMARK PLAN
        h1("8. Model Evaluation Strategy"),
        body("The recall target (≥80% on the churn class) is set deliberately above what a default, untuned model reaches. Planned benchmarking sequence, in order of expected recall:"),
        makeTable(
          ["Model", "Recall", "Precision", "F1-score"],
          [
            ["Logistic Regression (baseline)", "54.0%", "65.0%", "59.0%"],
            ["Random Forest (baseline)", "62.0%", "68.0%", "65.0%"],
            ["XGBoost (baseline)", "69.0%", "71.0%", "70.0%"],
            ["XGBoost + SMOTE + Optuna tuning", "84.0%", "75.0%", "79.2%"],
          ],
          [4200, 1720, 1720, 1720]
        ),
        caption("Table 8.1 — Planned model benchmarking sequence and target metrics."),
        chartImage("/home/claude/proj/fig8_1_model_performance.png", 5.5),
        caption("Figure 8.1 — Planned benchmark targets by model (illustrative planning figures, not yet-trained results)."),
        body("Note that these are planning targets for this report, not results from a completed training run — Week 1 is a planning deliverable and no model has been trained yet. The precision trade-off is deliberate: as recall is pushed up via SMOTE and tuning, precision typically drops somewhat, meaning more false-positive retention outreach. That is treated as an acceptable trade because the cost of a missed churner (lost subscriber, full replacement cost) is judged higher than the cost of an unnecessary retention call — this is an assumption the business stakeholder should confirm, not a fact this report can independently establish."),
        h2("8.1 Leakage controls"),
        bullet("All preprocessing (scaling, encoding) wrapped inside scikit-learn Pipeline objects, fit only on training folds."),
        bullet("SMOTE applied strictly to the training split — never to the validation or test set — to avoid synthetic points contaminating the evaluation."),
        bullet("Cross-validation using stratified folds to preserve the churn/no-churn ratio in every fold."),
        h2("8.2 Precision floor alongside the recall target"),
        body("Recall alone is an incomplete acceptance criterion: a model can hit 100% recall by flagging every customer as a churn risk, which is useless in practice because it saturates the retention team with false alarms. To prevent this, the model must clear both thresholds together, not recall in isolation:"),
        bullet("Recall target: ≥80% on the churn class (Section 2.4) — the floor for how many actual churners the model is allowed to miss."),
        bullet("Precision floor: ≥70% — no more than roughly 3 in 10 flagged customers may be false positives. Equivalently, minimum F1-score of ~74%."),
        body("The 70% figure is a starting assumption for this plan, not a number derived from the dataset or from stakeholder input — it should be replaced with an actual figure once the retention team's per-contact cost and the average customer's remaining lifetime value are known (see Section 2.1). If a retention outreach is cheap (an automated email) the precision floor can be relaxed; if it is expensive (a personal retention call with a discount offer), it should be raised. Treating 70% as fixed without that conversation is the kind of assumption this report flagged in Section 8, and it applies here directly."),

        // 9. RISK MANAGEMENT
        h1("9. Risk & Resource Analysis"),
        makeTable(
          ["Risk", "Likelihood", "Impact", "Mitigation"],
          [
            ["DPIA / legal sign-off delayed beyond W3", "Medium", "High — blocks all downstream phases", "Start DPO engagement in parallel with project kickoff, not after Phase 1 begins"],
            ["Model recall target (80%) not reached", "Medium", "High — undermines business case", "Reserve buffer time in W9 for extended tuning; pre-agree a fallback threshold with stakeholders"],
            ["Data leakage inflates reported performance", "Medium", "High — silent failure, discovered late", "Mandatory pipeline code review before any metric is reported as final"],
            ["Class-imbalance correction overfits to synthetic data", "Low-Medium", "Medium", "Validate on real (non-SMOTE) test data only; monitor precision alongside recall"],
            ["API latency exceeds 100ms under production load", "Low", "Medium — degrades CRM integration experience", "Load-test in staging (Phase 4) before Phase 5 shadow deployment"],
            ["DPDP Act non-compliance in automated decisioning", "Low", "Very high — regulatory penalty exposure", "DPO sign-off gate before any customer-facing automated action is enabled"],
            ["Feature drift after deployment (pricing, plans change)", "Medium", "Medium — silent accuracy decay", "Automated drift monitoring and scheduled retraining via Airflow DAG"],
          ],
          [2600, 1400, 1800, 3560]
        ),
        caption("Table 9.1 — Risk register with likelihood, impact and mitigation."),
        body("The most consequential risk on this list is not the model risk — it is the compliance risk. A missed recall target costs the project time and iteration; a DPDP Act violation on automated customer-affecting decisions carries a stated maximum exposure in the ₹250 crore range under the Act, which is why the DPO gate in Phase 1 is treated as non-negotiable rather than a formality."),
        h2("9.1 Data drift and retraining triggers"),
        body("Feature and performance drift is monitored post-deployment using Evidently AI, with retraining triggered automatically rather than left to manual review, so degradation doesn't go unnoticed between quarterly check-ins. Two trigger conditions, evaluated independently:"),
        bullet("Feature drift trigger: Population Stability Index (PSI) ≥ 0.2 on numerical input features (MonthlyCharges, tenure) between the production scoring window and the training baseline distribution. PSI ≥ 0.2 is the standard industry cut-off for 'significant distribution shift' — below 0.1 is treated as stable, 0.1–0.2 as a watch zone."),
        bullet("Performance drift trigger: rolling test recall on labelled outcomes (customers who did or didn't actually churn) drops below 75% (5 points under the 80% target), measured over a trailing 30-day window, sustained for 30 consecutive days rather than a single bad day, to avoid retraining on noise."),
        body("Either trigger independently initiates a scheduled retraining run via the Airflow DAG (Section 7.2); a breach of both simultaneously escalates to manual model review rather than automatic redeployment, since that combination suggests something more structural than routine drift (e.g. a pricing change or new competitor offer shifting churn behaviour wholesale)."),

        // 10. STRATEGIC ROADMAP
        h1("10. Strategic Roadmap: Deployment & Operationalisation"),
        body("Deployment does not go straight from a trained model to full production. It proceeds through three controlled phases so that both technical reliability and actual business impact are verified before committing budget to full-scale rollout."),
        h2("Phase A — Shadow Deployment"),
        body("The model runs silently alongside existing production systems, generating predictions that are logged but not acted on. This validates that the model's predictions correlate with real billing-cycle outcomes on live data, not just the held-out test set, before any customer is touched."),
        h2("Phase B — Commercial A/B Testing"),
        body("Subscribers are split into two groups: Group A continues to receive the existing, non-model-driven retention process; Group B receives targeted, SHAP-guided interventions from the model. The differential retention rate and ROI between the two groups is the actual evidence for whether the model creates business value — the offline recall metric from Phase 3 is necessary but not sufficient proof of that."),
        h2("Phase C — Full CRM Integration"),
        body("Only after Phase B shows a positive, statistically defensible ROI difference does the model get integrated into the CRM as a standing system, with a continuous feedback loop so retention outcomes flow back into future retraining. The drift triggers defined in Section 9.1 (PSI ≥ 0.2 on feature distributions; rolling recall below 75% over 30 days) become the standing operational gate at this phase — a trigger firing pauses new automated outreach pending a retrain, rather than letting a decaying model keep acting on stale patterns."),

        // 11. DOCUMENTATION NOTE
        h1("11. Documentation & Formatting Notes"),
        body("This report is structured with numbered sections and Word built-in heading styles, so the Table of Contents on page 2 is auto-generated and will update if reopened in Word (right-click → Update Field). Tables are used for all quantitative comparisons rather than prose, in line with the request for clear, scannable presentation."),

        // 12. CONCLUSION
        h1("12. Conclusion"),
        body("This plan commits to a specific, checkable target — 80% recall on the churn class — rather than a vague aspiration to 'build a good model', and ties every phase to a concrete deliverable and gating milestone. The two biggest open risks are not technical: DPIA sign-off timing and the precision/recall trade-off assumption both need explicit stakeholder confirmation before Phase 2 begins, and should not be treated as settled by this document alone."),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync("/mnt/user-data/outputs/Telecom_Churn_ML_Pipeline_Project_Plan.docx", buf);
  console.log("done");
});
