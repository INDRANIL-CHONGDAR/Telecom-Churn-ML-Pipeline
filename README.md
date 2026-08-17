# Telecom Churn ML Pipeline — Week 1: Project Planning

Week 1 deliverable for an end-to-end ML pipeline project on telecom customer churn.
This week is planning-only — no model has been trained yet. That happens in later weeks.

## What's here

- `docs/` — the Week 1 project plan (Word doc): problem statement, timeline, risk register, ROI illustration, DPDP Act compliance notes.
- `data/` — https://www.kaggle.com/datasets/blastchar/telco-customer-churn
            the IBM Telco Customer Churn dataset (7,043 records, 21 attributes) used for the exploratory charts.
  ** The project uses the Telco Customer Churn dataset originally provided as an IBM Sample Dataset and obtained through Kaggle.
  The raw CSV file is not redistributed in this repository. Please obtain the dataset from the original Kaggle source: **
- `scripts/make_charts.py` — regenerates the six exploratory charts in `outputs/` from `data/`.
- `scripts/build.js` — regenerates `docs/Telecom_Churn_ML_Pipeline_Project_Plan.docx` from scratch (Node.js, uses the `docx` npm package).
- `outputs/` — the six charts referenced in the project plan (churn distribution, churn by contract, churn by service, tenure vs. charges, timeline, planned model benchmarks).

## Running the scripts

**Charts** (Python 3, needs pandas + matplotlib):
```
pip install pandas matplotlib
python scripts/make_charts.py
```

**Word doc** (Node.js, needs the `docx` package):
```
npm install docx
node scripts/build.js
```

## Note on the model benchmark chart

`outputs/fig8_1_model_performance.png` shows *planned* recall/precision targets, not results from a trained model — labeled as such on the chart itself. No training has happened yet at this stage of the project.

## Status

- [x] Week 1 — Project planning and strategy
- [ ] Week 3 — Model training and evaluation (upcoming)
