# Netflix Content Strategy Analysis
### Portfolio, Global Markets, Audience & Growth Trends

---

## Overview
A business analytics project on a Netflix catalogue dataset of **8,790 titles**. Python (pandas) is used to clean the data and build the analysis, and Microsoft Excel is used for formula-driven tables, charts and a trend dashboard. Each task ends with business insights and recommendations.

This project covers **4 of the 6 business analytics tasks** (Tasks 1–4).

## Tools & Skills
| Tools | Skills |
|---|---|
| Python (pandas) | Data cleaning & preparation |
| Python (openpyxl, matplotlib) | KPI identification & business reporting |
| Excel formulas (COUNTIFS, SUMIFS, AVERAGEIFS, INDEX/MATCH) | Market, audience and trend analysis |
| Excel charts & dashboard sheet | Data visualization & dashboard design |

## Dataset
`Dataset.csv` – columns: show_id, type, title, director, country, date_added, release_year, rating, duration, listed_in.

**Cleaning steps:** parsed `date_added` into dates and `year_added`; split `duration` into a numeric value; mapped 14 ratings to 5 audience segments; mapped countries to 8 regions; rows with country "Not Given" (287) are excluded from market-share calculations.

## Tasks Completed

### Task 1 – Netflix Content Portfolio Analysis
KPI summary, Movie vs TV Show mix, release-era breakdown and top 10 genres per type.
- 69.7% of titles are Movies and 30.3% are TV Shows.
- 67% of TV Shows have only one season.

### Task 2 – Global Market Performance Analysis
Country ranking, regional contribution, market tiers and emerging-market list with expansion recommendations.
- The US holds 38% of titles with a known country; the top 5 countries hold 66%.
- Egypt, Nigeria and the Philippines had 94–98% of their titles added in 2019–21.

### Task 3 – Audience Segmentation Analysis
Ratings grouped into Kids, Family, Teens, Adults and Unrated; Movies vs TV comparison and yearly segment trend.
- Adults (TV-MA / R / NC-17) are the largest segment at 45.6%; Teens follow at 30.1%.
- TV Shows skew towards Kids content (17.4% vs 7.2% of movies).

### Task 4 – Content Growth & Trend Analysis
Dashboard with KPI tiles, yearly additions, release-year trend and growth phases.
- Yearly additions peaked at 2,016 titles in 2019.
- Additions grew at about 67.9% per year from 2016 to 2019.

## Repository Structure
```
├── Task1_Netflix_Content_Portfolio_Analysis.xlsx
├── Task2_Global_Market_Performance_Analysis.xlsx
├── Task3_Audience_Segmentation_Analysis.xlsx
├── Task4_Content_Growth_Trend_Analysis.xlsx
├── build_analysis.py          # Python source code that builds all workbooks
├── Dataset_Cleaned.csv        # Cleaned dataset
├── charts/                    # PNG chart screenshots
└── README.md
```

## How to Run
```bash
pip install pandas openpyxl matplotlib
python build_analysis.py
```
Update the `SRC` and `OUT` paths at the top of the script to match your folders.

## Notes & Limitations
- The data ends on 25 Sep 2021, so 2021 is a partial year.
- The dataset shows catalogue availability by production country, not viewing hours or revenue.
- Insight text in each workbook is static; the tables and charts update by formula.

## Author
**Harsh Maradiya** – Business Analytics
