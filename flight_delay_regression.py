"""
=============================================================================
ITEC 5040 - Week 2 Lab: Applied Advanced Regression Techniques
Multiple Linear Regression: Predicting Airline Arrival Delays
Author: Shruti Malik
Dataset: Chapter_06_flight_delay.csv
=============================================================================
"""

import sys, os, io, warnings
warnings.filterwarnings("ignore")

import pandas as pd
import numpy as np

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
from sklearn.inspection import permutation_importance

import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.diagnostic import het_breuschpagan
from scipy import stats

# ── seaborn style (works for both old and new versions) ─────────────────────
try:
    sns.set_theme(style="whitegrid", palette="muted")
except AttributeError:
    sns.set(style="whitegrid", palette="muted")

SEED = 42
BASE = os.path.dirname(os.path.abspath(__file__))
OUT  = os.path.join(BASE, "output")
os.makedirs(OUT, exist_ok=True)

# ── Tee stdout to log file ────────────────────────────────────────────────────
LOG_PATH = os.path.join(OUT, "analysis_log.txt")
_lf = open(LOG_PATH, "w", encoding="utf-8")
class _Tee:
    def __init__(self, a, b): self._a, self._b = a, b
    def write(self, x): self._a.write(x); self._b.write(x)
    def flush(self): self._a.flush(); self._b.flush()
sys.stdout = _Tee(sys.__stdout__, _lf)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 – LOAD AND EXPLORE THE DATA
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 65)
print("STEP 1: LOAD AND EXPLORE THE DATA")
print("=" * 65)

CSV = os.path.join(BASE, "Chapter_06_flight_delay.csv")
df  = pd.read_csv(CSV)

print(f"\nDataset shape : {df.shape[0]} rows x {df.shape[1]} columns")
print("\nColumn names  :", list(df.columns))
print("\nData types:\n", df.dtypes.to_string())
print("\nFirst 5 rows:\n", df.head().to_string())
print("\nDescriptive statistics:\n", df.describe().round(2).to_string())
print("\nMissing values per column:\n", df.isnull().sum().to_string())
print("\nCarrier distribution:\n", df["Carrier"].value_counts().to_string())

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 – INDIVIDUAL VARIABLE EXPLORATION & ASSUMPTION CHECKS
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 2: EXPLORING VARIABLES & ASSUMPTION CHECKS")
print("=" * 65)

num_cols = df.select_dtypes(include="number").columns.tolist()
print("\nNumeric columns:", num_cols)

# Histograms
fig, axes = plt.subplots(2, 5, figsize=(18, 7))
axes = axes.flatten()
for i, col in enumerate(num_cols):
    axes[i].hist(df[col].dropna().values, bins=30, edgecolor="white",
                 color="#4C72B0", alpha=0.85)
    axes[i].set_title(col, fontsize=10, fontweight="bold")
    axes[i].set_xlabel("Value"); axes[i].set_ylabel("Frequency")
plt.suptitle("Distribution of All Numeric Variables", fontsize=13,
             fontweight="bold", y=1.01)
plt.tight_layout()
plt.savefig(os.path.join(OUT, "01_variable_distributions.png"),
            dpi=150, bbox_inches="tight")
plt.close()
print("[Saved] 01_variable_distributions.png")

# Skewness & kurtosis
sk_kt = pd.DataFrame({
    "Skewness": df[num_cols].skew().round(3),
    "Kurtosis": df[num_cols].kurt().round(3)
})
print("\nSkewness & Kurtosis:\n", sk_kt.to_string())

# Shapiro-Wilk on Arr_Delay
sample = df["Arr_Delay"].sample(500, random_state=SEED).values
stat_sw, p_sw = stats.shapiro(sample)
print(f"\nShapiro-Wilk test on Arr_Delay (n=500): W={stat_sw:.4f}, p={p_sw:.4f}")
print("  Normality:", "Rejected (p<0.05)" if p_sw < 0.05 else "Supported (p>=0.05)")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 – RELATIONSHIPS BETWEEN VARIABLES
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 3: EVALUATING RELATIONSHIPS BETWEEN VARIABLES")
print("=" * 65)

corr = df[num_cols].corr().round(3)
print("\nCorrelation matrix:\n", corr.to_string())

fig, ax = plt.subplots(figsize=(11, 8))
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, mask=mask, annot=True, fmt=".2f", cmap="coolwarm",
            center=0, linewidths=0.5, ax=ax, annot_kws={"size": 9})
ax.set_title("Correlation Heatmap", fontsize=13, fontweight="bold", pad=12)
plt.tight_layout()
plt.savefig(os.path.join(OUT, "02_correlation_heatmap.png"),
            dpi=150, bbox_inches="tight")
plt.close()
print("[Saved] 02_correlation_heatmap.png")

# Scatter plots (key predictors vs target)
key_preds = ["Airport_Distance","Weather","Baggage_loading_time",
             "Late_Arrival_o","Security_o"]
fig, axes = plt.subplots(1, 5, figsize=(20, 4))
for ax, col in zip(axes, key_preds):
    ax.scatter(df[col].values, df["Arr_Delay"].values,
               alpha=0.35, color="#4C72B0", s=20, edgecolors="none")
    m, b = np.polyfit(df[col].values, df["Arr_Delay"].values, 1)
    x_line = np.linspace(df[col].min(), df[col].max(), 100)
    ax.plot(x_line, m*x_line + b, "r-", linewidth=1.5)
    r = corr.loc[col, "Arr_Delay"]
    ax.set_title(f"{col}\n(r = {r:.3f})", fontsize=9, fontweight="bold")
    ax.set_xlabel(col, fontsize=8); ax.set_ylabel("Arr_Delay", fontsize=8)
plt.suptitle("Key Predictors vs. Arrival Delay", fontsize=12,
             fontweight="bold", y=1.02)
plt.tight_layout()
plt.savefig(os.path.join(OUT, "03_scatter_plots.png"),
            dpi=150, bbox_inches="tight")
plt.close()
print("[Saved] 03_scatter_plots.png")

target_corr = corr["Arr_Delay"].drop("Arr_Delay").sort_values(key=abs, ascending=False)
print("\nCorrelations with Arr_Delay (sorted by |r|):\n", target_corr.to_string())

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 – FEATURE ENGINEERING & TRAIN / TEST SPLIT
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 4: FEATURE ENGINEERING & TRAIN/TEST SPLIT")
print("=" * 65)

df_enc = pd.get_dummies(df, columns=["Carrier"], drop_first=True)
# Convert boolean dummy columns to int for sklearn compatibility
bool_cols = df_enc.select_dtypes(include="bool").columns
df_enc[bool_cols] = df_enc[bool_cols].astype(int)

feat_cols = [c for c in df_enc.columns if c != "Arr_Delay"]
X = df_enc[feat_cols].astype(float)
y = df_enc["Arr_Delay"].astype(float)

X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.30, random_state=SEED)

print(f"\nFeatures after encoding : {X.shape[1]}")
print(f"Train set : {len(X_train)} rows ({len(X_train)/len(df)*100:.1f}%)")
print(f"Test  set : {len(X_test)} rows  ({len(X_test)/len(df)*100:.1f}%)")
print(f"\nFeature list:\n{list(feat_cols)}")

scaler = StandardScaler()
X_train_sc = pd.DataFrame(scaler.fit_transform(X_train), columns=X_train.columns)
X_test_sc  = pd.DataFrame(scaler.transform(X_test),      columns=X_test.columns)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 – VARIANCE INFLATION FACTOR (Multicollinearity)
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 5: MULTICOLLINEARITY CHECK (VIF)")
print("=" * 65)

X_tr_const = sm.add_constant(X_train_sc)
vif_df = pd.DataFrame({
    "Feature": X_tr_const.columns,
    "VIF":     [variance_inflation_factor(X_tr_const.values, i)
                for i in range(X_tr_const.shape[1])]
})
vif_df = vif_df[vif_df["Feature"] != "const"].sort_values("VIF", ascending=False)
print("\nVIF (all features, sorted):\n", vif_df.round(2).to_string(index=False))
hi_vif = vif_df[vif_df["VIF"] > 10]
if hi_vif.empty:
    print("\n  No features with VIF > 10 (multicollinearity acceptable).")
else:
    print("\n  High VIF (>10):\n", hi_vif.to_string(index=False))

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 – BUILD OLS REGRESSION MODEL
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 6: BUILDING THE MULTIPLE LINEAR REGRESSION MODEL")
print("=" * 65)

X_tr_ols = sm.add_constant(X_train)
ols      = sm.OLS(y_train, X_tr_ols).fit()
print(ols.summary())

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7 – EVALUATE MODEL ON TEST SET
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 7: MODEL EVALUATION ON TEST SET")
print("=" * 65)

X_te_ols = sm.add_constant(X_test, has_constant="add")
y_pred   = ols.predict(X_te_ols)

rmse        = float(np.sqrt(mean_squared_error(y_test, y_pred)))
mae         = float(mean_absolute_error(y_test, y_pred))
r2          = float(r2_score(y_test, y_pred))
adj_r2_tr   = float(ols.rsquared_adj)
fstat       = float(ols.fvalue)
f_pval      = float(ols.f_pvalue)

print(f"\n  R2  (test)        : {r2:.4f}")
print(f"  Adj R2 (train)    : {adj_r2_tr:.4f}")
print(f"  RMSE (test)       : {rmse:.2f} minutes")
print(f"  MAE  (test)       : {mae:.2f} minutes")
print(f"  F-statistic       : {fstat:.2f}  (p = {f_pval:.4e})")

# -- Predicted vs Actual
fig, ax = plt.subplots(figsize=(8, 6))
ax.scatter(y_test.values, y_pred.values, alpha=0.45, edgecolors="white",
           linewidth=0.4, color="#4C72B0", s=45, label="Predictions")
lo = min(y_test.min(), y_pred.min()) - 5
hi = max(y_test.max(), y_pred.max()) + 5
ax.plot([lo, hi], [lo, hi], "r--", linewidth=1.5, label="Perfect fit")
ax.set(xlabel="Actual Arrival Delay (min)",
       ylabel="Predicted Arrival Delay (min)",
       title=f"Predicted vs. Actual\n(R2={r2:.4f}, RMSE={rmse:.2f} min)",
       xlim=[lo, hi], ylim=[lo, hi])
ax.legend(); ax.set_aspect("equal")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "04_predicted_vs_actual.png"),
            dpi=150, bbox_inches="tight")
plt.close()
print("[Saved] 04_predicted_vs_actual.png")

# -- Residual diagnostics
residuals = y_test.values - y_pred.values
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
axes[0].scatter(y_pred.values, residuals, alpha=0.4, color="#DD8452",
                edgecolors="none", s=40)
axes[0].axhline(0, color="red", linestyle="--", linewidth=1.4)
axes[0].set(xlabel="Fitted Values", ylabel="Residuals",
            title="Residuals vs. Fitted Values")
sm.qqplot(residuals, line="s", ax=axes[1], alpha=0.5)
axes[1].set_title("Normal Q-Q Plot of Residuals")
plt.suptitle("Residual Diagnostics", fontsize=13, fontweight="bold")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "05_residual_diagnostics.png"),
            dpi=150, bbox_inches="tight")
plt.close()
print("[Saved] 05_residual_diagnostics.png")

# -- Durbin-Watson
dw = float(durbin_watson(ols.resid))
print(f"\n  Durbin-Watson     : {dw:.4f}")
print("  Autocorrelation   :", "None detected" if 1.5 < dw < 2.5
      else "Possible autocorrelation")

# -- Breusch-Pagan
bp_lm, bp_p, _, _ = het_breuschpagan(ols.resid, X_tr_ols)
print(f"\n  Breusch-Pagan LM  : {bp_lm:.4f}  (p = {bp_p:.4f})")
print("  Homoscedasticity  :", "Supported" if bp_p >= 0.05
      else "Rejected - heteroscedasticity present")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 8 – FEATURE IMPORTANCE
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 8: FEATURE IMPORTANCE")
print("=" * 65)

sk_model = LinearRegression().fit(X_train_sc, y_train)
coef_df  = pd.DataFrame({
    "Feature":   X_train_sc.columns,
    "Std_Coeff": sk_model.coef_
}).sort_values("Std_Coeff", key=abs, ascending=False).head(15)
print("\nTop 15 standardized coefficients:\n", coef_df.round(4).to_string(index=False))

fig, ax = plt.subplots(figsize=(9, 6))
colors = ["#4C72B0" if v > 0 else "#DD8452" for v in coef_df["Std_Coeff"]]
ax.barh(coef_df["Feature"].values[::-1],
        coef_df["Std_Coeff"].values[::-1],
        color=colors[::-1], edgecolor="white", height=0.65)
ax.axvline(0, color="black", linewidth=0.8)
ax.set_xlabel("Standardized Coefficient")
ax.set_title("Top 15 Features by Standardized OLS Coefficient", fontweight="bold")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "06_feature_importance.png"),
            dpi=150, bbox_inches="tight")
plt.close()
print("[Saved] 06_feature_importance.png")

# -- Permutation importance
perm = permutation_importance(sk_model, X_test_sc, y_test,
                              n_repeats=20, random_state=SEED, scoring="r2")
perm_df = pd.DataFrame({
    "Feature":    X_test_sc.columns,
    "Importance": perm.importances_mean
}).sort_values("Importance", ascending=False).head(10)
print("\nTop 10 permutation importances (R2 drop):\n",
      perm_df.round(4).to_string(index=False))

# ─────────────────────────────────────────────────────────────────────────────
# STEP 9 – FINAL SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("STEP 9: FINAL MODEL SUMMARY")
print("=" * 65)
print(f"""
  +-----------------------------------------+
  |     MULTIPLE LINEAR REGRESSION MODEL    |
  |   Predicting Airline Arrival Delays     |
  +-----------------------------------------+
  | Dataset         : 3,594 observations    |
  | Raw Features    : 10 columns            |
  | Encoded Features: {X_train.shape[1]:>2} (after one-hot)  |
  | Train / Test    : 70% / 30%             |
  +-----------------------------------------+
  | R2  (test)      : {r2:.4f}               |
  | Adj. R2 (train) : {adj_r2_tr:.4f}               |
  | RMSE            : {rmse:.2f} minutes        |
  | MAE             : {mae:.2f} minutes        |
  | F-statistic     : {fstat:.2f}               |
  | Durbin-Watson   : {dw:.4f}               |
  +-----------------------------------------+
""")
print(f"Charts saved to : {OUT}")
print(f"Log saved to    : {LOG_PATH}")
print("\nAnalysis complete.")

# restore stdout and close log
sys.stdout = sys.__stdout__
_lf.close()
