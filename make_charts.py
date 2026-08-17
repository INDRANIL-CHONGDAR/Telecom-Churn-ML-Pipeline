import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams['font.family'] = 'DejaVu Sans'

df = pd.read_csv('/mnt/user-data/uploads/WA_Fn-UseC_-Telco-Customer-Churn.csv')
df['TotalCharges'] = pd.to_numeric(df['TotalCharges'], errors='coerce')

NAVY='#1E293B'; CYAN='#06b6d4'; ROSE='#f43f5e'; INDIGO='#6366f1'
PURPLE='#8b5cf6'; EMERALD='#10b981'; BLUE='#3b82f6'; GREY='#64748B'

def style_ax(ax):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#CBD5E1')
    ax.spines['bottom'].set_color('#CBD5E1')
    ax.tick_params(colors=NAVY, labelsize=10)

# Fig 1.1 doughnut
fig, ax = plt.subplots(figsize=(5,4), dpi=150)
counts = df['Churn'].value_counts()
vals=[counts['No'], counts['Yes']]
labels=[f'Retained\n({counts["No"]:,})', f'Churned\n({counts["Yes"]:,})']
ax.pie(vals, labels=labels, colors=[CYAN,ROSE], autopct='%1.1f%%', startangle=90,
       wedgeprops=dict(width=0.45, edgecolor='white'), textprops={'fontsize':10,'color':NAVY})
ax.set_title('Figure 1.1: Baseline Churn Class Distribution', fontsize=11, color=NAVY, fontweight='bold', pad=15)
plt.tight_layout(); plt.savefig('fig1_1_churn_distribution.png', bbox_inches='tight', facecolor='white'); plt.close()

# Fig 1.2 contract
fig, ax = plt.subplots(figsize=(5.5,4), dpi=150)
rates=(df.groupby('Contract')['Churn'].apply(lambda s:(s=='Yes').mean()*100).reindex(['Month-to-month','One year','Two year']))
bars=ax.bar(rates.index, rates.values, color=[ROSE,BLUE,EMERALD], width=0.55)
for b,v in zip(bars, rates.values):
    ax.text(b.get_x()+b.get_width()/2, v+1, f'{v:.1f}%', ha='center', fontsize=10, color=NAVY, fontweight='bold')
ax.set_ylabel('Churn Rate (%)'); ax.set_ylim(0,50)
ax.set_title('Figure 1.2: Churn Rate by Contract Term', fontsize=11, color=NAVY, fontweight='bold', pad=15)
style_ax(ax); plt.tight_layout(); plt.savefig('fig1_2_churn_by_contract.png', bbox_inches='tight', facecolor='white'); plt.close()

# Fig 1.3 service/support
fig, ax = plt.subplots(figsize=(6,4), dpi=150)
fd=(df[df['InternetService'].isin(['Fiber optic','DSL'])].groupby('InternetService')['Churn'].apply(lambda s:(s=='Yes').mean()*100))
sup=(df[df['TechSupport'].isin(['Yes','No'])].groupby('TechSupport')['Churn'].apply(lambda s:(s=='Yes').mean()*100))
cats=['Fiber Optic','DSL Internet','No Tech Support','Tech Support Subscribed']
vals=[fd['Fiber optic'], fd['DSL'], sup['No'], sup['Yes']]
colors=[PURPLE,CYAN,ROSE,EMERALD]
bars=ax.barh(cats, vals, color=colors, height=0.55)
for b,v in zip(bars, vals):
    ax.text(v+1, b.get_y()+b.get_height()/2, f'{v:.1f}%', va='center', fontsize=10, color=NAVY, fontweight='bold')
ax.set_xlabel('Churn Rate (%)'); ax.set_xlim(0,50); ax.invert_yaxis()
ax.set_title('Figure 1.3: Churn Rate by Service Type & Tech Support', fontsize=11, color=NAVY, fontweight='bold', pad=15)
style_ax(ax); plt.tight_layout(); plt.savefig('fig1_3_churn_by_service.png', bbox_inches='tight', facecolor='white'); plt.close()

# Fig 1.4 scatter
fig, ax = plt.subplots(figsize=(6,4.2), dpi=150)
churned=df[df['Churn']=='Yes'].sample(min(500,(df['Churn']=='Yes').sum()), random_state=42)
retained=df[df['Churn']=='No'].sample(min(500,(df['Churn']=='No').sum()), random_state=42)
ax.scatter(retained['tenure'], retained['MonthlyCharges'], s=10, alpha=0.4, color=EMERALD, label='Retained')
ax.scatter(churned['tenure'], churned['MonthlyCharges'], s=10, alpha=0.5, color=ROSE, label='Churned')
ax.set_xlabel('Tenure (Months)'); ax.set_ylabel('Monthly Charges ($)')
ax.set_title('Figure 1.4: Tenure vs. Monthly Charges Risk Clustering', fontsize=11, color=NAVY, fontweight='bold', pad=15)
ax.legend(loc='upper right', fontsize=9, frameon=False)
style_ax(ax); plt.tight_layout(); plt.savefig('fig1_4_tenure_charges.png', bbox_inches='tight', facecolor='white'); plt.close()

# Fig 8.1 model performance (planned targets)
fig, ax = plt.subplots(figsize=(7,4.2), dpi=150)
models=['Logistic\nRegression','Random Forest\n(Baseline)','XGBoost\n(Baseline)','XGBoost + SMOTE\n+ Optuna']
recall=[54.0,62.0,69.0,84.0]; precision=[65.0,68.0,71.0,75.0]; f1=[59.0,65.0,70.0,79.2]
x=np.arange(len(models)); w=0.25
ax.bar(x-w, recall, w, label='Recall (%)', color=EMERALD)
ax.bar(x, precision, w, label='Precision (%)', color=CYAN)
ax.bar(x+w, f1, w, label='F1-Score (%)', color=INDIGO)
ax.axhline(80, color=ROSE, linestyle='--', linewidth=1, alpha=0.7)
ax.text(len(models)-0.5, 81.5, 'Target Recall = 80%', color=ROSE, fontsize=8, ha='right')
ax.set_xticks(x); ax.set_xticklabels(models, fontsize=9)
ax.set_ylabel('Score (%)'); ax.set_ylim(0,100)
ax.set_title('Figure 8.1: Planned Model Benchmark Targets (illustrative — not yet trained)', fontsize=10.5, color=NAVY, fontweight='bold', pad=15)
ax.legend(loc='upper left', fontsize=9, frameon=False)
style_ax(ax); plt.tight_layout(); plt.savefig('fig8_1_model_performance.png', bbox_inches='tight', facecolor='white'); plt.close()

# Fig 6.1 timeline
fig, ax = plt.subplots(figsize=(7,3.2), dpi=150)
phases=['Phase 5: Deploy\n& Monitor (W13-15)','Phase 4: MLOps\n& API (W10-12)','Phase 3: Model\nTuning (W7-9)',
        'Phase 2: Data\nPipeline (W4-6)','Phase 1: Legal\n& DPIA (W1-3)']
starts=[12,9,6,3,0]
colors_g=[ROSE,EMERALD,PURPLE,INDIGO,CYAN]
ax.barh(phases, [3]*5, left=starts, color=colors_g, height=0.55)
ax.set_xlabel('Project Week'); ax.set_xlim(0,15)
ax.set_title('Figure 6.1: Fifteen-Week Phase Timeline', fontsize=11, color=NAVY, fontweight='bold', pad=15)
style_ax(ax); plt.tight_layout(); plt.savefig('fig6_1_timeline.png', bbox_inches='tight', facecolor='white'); plt.close()

print("done")
