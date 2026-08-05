# OTA Girls U14 Ranking & Dashboard Web App

A lightweight, automated pipeline and interactive Streamlit web application designed to track, analyze, and visualize weekly Ontario Tennis Association (OTA) Girls Under 14 tennis rankings and player profiles. 

The application automatically scrapes tournament software ranking tables, performs parallel profile analysis to extract detailed match records and WTN (World Tennis Number) metrics, stores history inside an Excel workbook, and serves a rich analytics dashboard[cite: 3, 4].

---

## Repository Structure

```text
├── .github/
│   └── workflows/
│       └── run-scraper.yml     # GitHub Actions workflow for automated weekly cron runs
├── rankextract.py              # Automated web scraper and Excel workbook generator
├── app.py                      # Streamlit dashboard application
├── master.xlsx                 # Central database holding weekly ranking sheets and historical data
└── requirements.txt            # Project Python package dependencies
