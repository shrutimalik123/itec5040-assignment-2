# ITEC 5040 Week 2 Lab: Applied Advanced Regression Techniques

This project demonstrates the application of multiple linear regression (MLR) to predict airline arrival delays using operational aviation data.

## 📋 Table of Contents
- [Project Overview](#project-overview)
- [Key Features](#key-features)
- [Dataset](#dataset)
- [Project Structure](#project-structure)
- [Analysis Methodology](#analysis-methodology)
- [Key Findings](#key-findings)
- [Requirements](#requirements)

## ✈️ Project Overview
The primary objective of this project is to build a predictive model that estimates airline arrival delays based on various operational factors such as airport distance, total daily flights, weather, and ground crew efficiency.

## 🚀 Key Features
- **Exploratory Data Analysis (EDA):** Comprehensive visualization of variables and their distributions.
- **Assumption Testing:** Formal tests for normality, linearity, homoscedasticity, and autocorrelation.
- **Multicollinearity Diagnostic:** VIF (Variance Inflation Factor) analysis to ensure coefficient stability.
- **Feature Engineering:** One-hot encoding for airline carrier categorical data.
- **Performance Evaluation:** Rigorous assessment using R-squared (R²), RMSE, and MAE on a 30% hold-out test set.

## 📊 Dataset
The analysis uses the `Chapter_06_flight_delay.csv` dataset, which contains 3,593 flight records with features including:
- **Predictors:** Carrier, Airport Distance, Number of Flights, Weather, Support Crew, Baggage Loading Time, Late Inbound Arrivals, Cleaning/Fueling/Security times.
- **Target:** Arrival Delay (`Arr_Delay`) in minutes.

## 📂 Project Structure
- `flight_delay_regression.py`: Core Python script for analysis and model building.
- `Week2_Lab_Shruti_Malik.docx`: Final comprehensive lab report.
- `Week2_Lab_Shruti_Malik.pdf`: PDF version of the final report.
- `output/`: Folder containing all visualization charts and the raw analysis log.
- `Chapter_06_flight_delay.csv`: Raw dataset.

## 🔬 Analysis Methodology
1. **Load and Explore:** Initial profiling and descriptive stats.
2. **Assumption Checks:** Symmetry and normality testing (Shapiro-Wilk).
3. **Correlation Study:** Heatmaps and scatter plots to identify key drivers.
4. **Data Preparation:** Encoding carriers and train-test splitting (70/30).
5. **Multicollinearity (VIF):** Ensuring features are sufficiently independent.
6. **Model Building:** OLS regression with summary statistics.
7. **Evaluation:** Diagnostic testing (Durbin-Watson, Breusch-Pagan) and error metrics (RMSE/MAE).
8. **Feature Importance:** Standardized coefficients and permutation importance.

## 📈 Key Findings
- **R² Score:** 0.82 (82% of delay variance explained).
- **RMSE:** 12.43 minutes.
- **Primary Drivers:** The number of scheduled hub flights and baggage loading time are the strongest predictors of delay, outweighing carrier identity and small-scale operational tasks.
- **Carrier Neutrality:** Once operational factors are controlled for, individual airline brand was found to be a non-significant predictor of delay magnitude.

## ⚙️ Requirements
To run the analysis script, you will need:
- `pandas`
- `numpy`
- `matplotlib`
- `seaborn`
- `scikit-learn`
- `statsmodels`
- `scipy`
- `aspose-words` (Optional, for PDF generation)

---
*Created by Shruti Malik for ITEC 5040: Predictive Analytics.*
