# Hr-salary-analysis-dashboard
## 📌 Problem Statement
The HR leadership at a global enterprise needed clear visibility into salary distributions and overall workforce expenditure across (32K+) employee records. The objective of this project is to analyze where company funds are allocated and identify key spending patterns across job roles, employment types, and regions.
## 🛠️ Tools Used
Just Microsoft Excel doing heavy lifting. Powered entirely by dynamic functions and custom formulas built completely by hand. No shortcuts, just pure logic.
## 📊 The Dashboard


https://github.com/user-attachments/assets/d8fe6f88-1997-46e6-8a02-6156a5bbadd1
## 🧠 Technical Highlights (Advanced Formulas)
No pivot tables or hidden tricks—just pure Excel logic. Here are a few key dynamic formulas used:

1) For the "Median Salary" KPI card i used this clustering function:

```excel
=IFERROR(
  MEDIAN(
    FILTER(
      JOBS[salary_year_avg],
      (("All"=jobtitle)+(jobtitle=JOBS[job_title_short]))*
      (("All"=jobcountry)+(jobcountry=JOBS[job_country]))*
      (("All"=scheduletype)+ISNUMBER(FIND(scheduletype,JOBS[job_schedule_type])))*
      (("All"=DATE)+(DATE=JOBS[Month]))
    )
  ), "No result"
)
```

2) For "Job Counts" KPI card:

```excel
=COUNTA(FILTER(JOBS[job_title_short],
(("All"=jobtitle)+(jobtitle=JOBS[job_title_short]))*
(("All"=jobcountry)+(jobcountry=JOBS[job_country]))*
(("All"=scheduletype)+ISNUMBER(FIND(scheduletype,JOBS[job_schedule_type])))*
(("All"=DATE)+(DATE=JOBS[Month]))))
```

3) For "Top Platform" KPI card:

```excel
=COUNTIFS(JOBS[job_via],P2,JOBS[job_title_short],IF(jobtitle="All","*",jobtitle),JOBS[job_country],IF(jobcountry="All","*",jobcountry))
```
## 🔍 The Bottom Line & Key Insights

This dashboard cuts through the noise of (32K+) data points to directly answer three major business questions:

* **Where the Demand Is:** Immediately pinpoints high-demand roles like *Data Engineer* and *Data Analyst* across top job platforms.
* **True Salary Benchmarks:** Delivers realistic median salaries filtered on the fly by country, month, and contract type (Full-time, Contractor, etc.).
* **Global Talent Footprint:** Maps out regional hiring patterns geographically to spot global market trends.

---
⚡ *Built purely in Excel with custom dynamic array logic.*
## What's Next?
Follow me and watch me put into practice what I am learning.
