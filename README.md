# Employee Attrition Prediction

## Overview

This project focuses on predicting employee attrition using machine learning and identifying the key factors influencing workforce turnover. By leveraging HR analytics data, the goal is to shift from reactive HR decisions to proactive retention strategies.



## Objectives

* Identify key factors contributing to employee attrition
* Build predictive models to classify employees at risk
* Provide actionable insights for improving employee retention


## Tools & Technologies

* Python (Pandas, NumPy, Scikit-learn)
* Data Visualization (Matplotlib, Seaborn)
* Power BI (Dashboard)
* Google Colab



## Dataset

* IBM HR Analytics Dataset
* ~1470 employee records
* Features include age, job role, income, work-life balance, overtime, etc. 



## Approach

### 1. Data Preprocessing

* Handling missing values
* Encoding categorical variables
* Feature scaling
* Applied SMOTE to handle class imbalance

### 2. Exploratory Data Analysis

* Identified relationships between features and attrition
* Key patterns observed:

  * Overtime strongly linked to attrition
  * Lower income employees more likely to leave
  * Younger employees show higher attrition trends

### 3. Model Building

* Logistic Regression
* Decision Tree
* Random Forest



## Results

* Best Model: **Random Forest**
* Accuracy: **94.1%**
* AUC Score: **0.979**



## Key Insights

* Overtime is the strongest predictor of attrition
* Low job satisfaction and poor work-life balance increase risk
* Younger and lower-income employees are more likely to leave
* Certain roles (e.g., Sales Executive) show higher attrition



##  Business Impact

This project helps organizations:

* Identify high-risk employees
* Improve retention strategies
* Make data-driven HR decisions

---

## Project Structure

* Notebook: Data analysis and model building
* Dashboard: Visual insights
* Images: Key visualizations

---

## Conclusion

The project demonstrates how predictive analytics can transform HR decision-making by enabling proactive workforce management and improving organizational stability.
