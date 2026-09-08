# Excel Salary Dashboard

## Introduction

This Excel Salary Dashboard was created to help job seekers explore salary trends for their desired jobs and understand how salaries vary across different job titles, locations, and job schedule types.

The project uses a real-world data science jobs dataset from 2023 and demonstrates how Excel can be used to analyze data and present meaningful insights through an interactive dashboard.

### Dashboard File

My final dashboard is available in [01_Salary_Calculation_Project.xlsx](01_Salary_Calculation_Project.xlsx).

### Excel Skills Used

The following Excel skills were utilized for analysis:

- **📉 Charts**
- **🧮 Formulas and Functions**
- **❎ Data Validation**

### Data Jobs Dataset

The dataset used for this project contains real-world data science job information from 2023. It includes detailed information on:

- **👨‍💼 Job titles**
- **💰 Salaries**
- **📍 Locations**
- **🛠️ Skills**

## Dashboard Build

### 📉 Charts

#### 📊 Data Science Job Salaries - Bar Chart

- 🛠️ **Excel Features:** Utilized the bar chart feature with formatted salary values and an optimized layout for clarity.
- 🎨 **Design Choice:** Used a horizontal bar chart for visual comparison of median salaries.
- 📉 **Data Organization:** Sorted job titles by descending salary for improved readability.
- 💡 **Insights Gained:** The chart provides a quick view of salary trends and shows differences in median salaries across job roles.

#### 🗺️ Country Median Salaries - Map Chart

- 🛠️ **Excel Features:** Utilized Excel's Map Chart feature to display median salaries by country.
- 🎨 **Design Choice:** Used a color-coded map to visually differentiate salary levels across regions.
- 📊 **Data Representation:** Displayed median salary information for countries with available data.
- 👁️ **Visual Enhancement:** Improved readability and provided a clear geographic view of salary trends.
- 💡 **Insights Gained:** The map provides an overview of global salary differences across countries.

### 🧮 Formulas and Functions

#### 💰 Median Salary by Job Titles

The following formula calculates the median salary based on job title, country, and job schedule type:

```excel
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
