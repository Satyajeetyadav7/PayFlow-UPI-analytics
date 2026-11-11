# Resume and interview guide: PayFlow UPI Payments Analytics

For you, not for the repository. Read it fully before you apply.

## 0. Four rules that protect you

1. **Say the data is simulated.** Write "simulated payments dataset" in the resume or README. Interviewers respect it; being caught implying real company data ends the interview.
2. **Claim Power BI only after you build it.** The kit gives you every ingredient, but the `.pbix` is yours to create (2 to 3 hours with `powerbi/BUILD_GUIDE.md`). Build it, check your cards against the expected-values table, then claim it. Until then, list Excel and the web dashboard only.
3. **Do not say "published to Power BI Service" or "Publish to web" unless you did it.** Per Microsoft community answers, Publish to web needs a Pro licence and a work or school account. Say "built in Power BI Desktop, walkthrough video linked" instead.
4. **Make it yours.** Many students will find similar projects. Change the city list, rename the company, and add one analysis of your own (for example a retry-rate analysis or a bank ranking by value lost). A project you changed is easy to defend.

## 1. Resume entry (copy, then fill the links)

**PayFlow: UPI Payments Analytics** | Excel, Power BI, DAX, Power Query, Python
[Live Dashboard](https://YOUR-USERNAME.github.io/payflow-upi-analytics/) · [GitHub](https://github.com/YOUR-USERNAME/payflow-upi-analytics)

- Built a formula-driven Excel workbook (1,108 `COUNTIFS`/`SUMIFS`/`AVERAGEIFS` formulas on named ranges over a 60,000-row table) with a dropdown-driven KPI dashboard, conditional formatting and an outage finder for a simulated UPI payments dataset.
- Modelled a star schema in Power BI (1 fact, 5 dimensions) and wrote 36 DAX measures covering success and failure rates, time intelligence (MoM, YoY, YTD) and incident detection; cleaned and typed data in Power Query.
- Found that a co-operative bank's system failure rate was 5.8x the best private bank's (8.1% vs 1.4%), that small-finance and co-operative banks carried 18% of transactions but 35% of failures, and that failures put Rs 23.1 L (3.7%) of attempted value at risk.
- Detected two bank outages (60% and 63% failure rates against about 4.5% normal) with a binomial spike test (under 0.1% chance by luck) that flagged 6 of 5,096 bank-days, and showed fraud flags run at 75 bps for new customers against 23 bps for established ones.
- Published an interactive dashboard with filters, table views and dark mode; verified every Excel formula by independent recalculation.

**Short version (one-page resume):**
- Built an Excel and Power BI analytics project on simulated UPI payments (60,000 transactions): 1,108 formulas, a star-schema model and 36 DAX measures.
- Quantified bank reliability gaps (8.1% vs 1.4% system failure), detected two outages, and sized Rs 23.1 L of payment value at risk.
- Deployed a live interactive dashboard and a verified workbook; documented method and limits.

**ATS keywords covered:** Excel, advanced Excel, SUMIFS, COUNTIFS, named ranges, data validation, conditional formatting, dashboards, Power BI, DAX, Power Query, star schema, data modelling, time intelligence, KPI, anomaly detection, root-cause analysis, fraud analytics, payments, fintech, data cleaning, data visualisation, Python.

## 2. GitHub and LinkedIn text

**GitHub description:** UPI payments analytics: formula-driven Excel workbook, Power BI model with DAX, and a live dashboard. Simulated data.
**Topics:** `data-analysis` `excel` `power-bi` `dax` `fintech` `upi` `dashboard` `data-visualization`

**LinkedIn Featured blurb:** I built a payments analytics project (simulated data) that answers why UPI payments fail: which banks, hours and networks are weak, when outages happened, who gets flagged for fraud and what cashback costs. It has an Excel workbook, a Power BI model with 36 DAX measures and a live dashboard. Link below. Feedback welcome.

## 3. 60-second walkthrough

"I wanted a project that looks like real analyst work in fintech. I simulated a UPI payments company with realistic rules, because real payment data is private. In Excel I built formula-driven sheets with a dropdown dashboard; every number traces to a COUNTIFS or SUMIFS on a named range. In Power BI I modelled a star schema and wrote DAX for rates, time intelligence and incident detection. The main finding is that reliability depends on the bank and the signal, not the city: a co-operative bank fails for system reasons 5.8 times as often as the best private bank, and weak 2G or 3G connections fail at double the rate. I also built an outage finder that flags a day when the chance of that many failures at the bank's normal rate is under 0.1%, and it caught both injected outages. I checked the workbook by recalculating every formula independently, and I am clear that the data is simulated."

## 4. Interview questions with answers

**Why simulated data?** Payment data is private and public datasets rarely have banks, failure reasons, networks and fraud flags together. Simulating let me build known patterns and check my analysis recovers them. The cost is that the conclusions are illustrative; the same formulas and model work on real data with the same columns.

**Why COUNTIFS with a wildcard for "All"?** The dropdown cell holds "All". A helper cell turns it into `*`, which matches any text, so one formula serves both the filtered and unfiltered case without nested IFs.

**Why named ranges instead of cell ranges?** `SUMIFS(d_amt, d_bank, A5)` reads like a sentence, does not break when columns move, and is easy to audit.

**Why a helper column for monthly active users?** Distinct counts per month are slow with array formulas on 60,000 rows. I flag each customer's first transaction in a month in Python, then `SUMIFS` on that flag gives active users instantly. The trade-off is that the flag is a value, not a formula, and I say so in the dictionary.

**How do you know your formulas are right?** I computed every figure independently in pandas, wrote it as the cached value, then forced LibreOffice to recalculate all 1,108 formulas from scratch and compared. Zero mismatches. I also tested the method with a deliberately wrong cached value to be sure the recalculation was real.

**Explain your star schema.** One fact table of transactions and dimensions for date, customer, merchant, bank and city. Relationships are one-to-many from dimension to fact. City connects through customer, so a city filter reaches transactions without a second path.

**Why is one relationship inactive?** `dim_customer[signup_date]` to `dim_date[date]` would create two date relationships to the fact via customer. I keep it inactive and activate it with `USERELATIONSHIP` inside the `New Customers` measure.

**Explain `Chance of Spike by Luck`.** `Bank Normal System Failure Rate` removes the date filter (`REMOVEFILTERS(dim_date)`) to get the bank's all-period rate. Then `1 - BINOM.DIST(k-1, n, p, TRUE)` is the probability of at least k system failures in n payments if the bank were behaving normally. Below 0.1% is a flag.

**Why not a simple "2.5x the average" rule?** I tried it first. It flagged 17% of bank-days, because days with few payments swing wildly. A binomial test accounts for volume: 3 failures in 4 payments is not unusual, 14 in 20 is. It flags 6 of 5,096 bank-days and catches 4 of the 5 outage days with at least 5 transactions; the other 2 flags are false alarms, which is expected at 0.1% over thousands of tests. Say that out loud; it shows you checked your own method.

**Why at least 5 transactions?** Too few payments make any rate meaningless, so the day is not tested. The 0.1% threshold is a judgement call; on real data I would tune it on history.

**What is the difference between system, user and risk failures?** System failures are bank or network faults the company can fix. User failures are balance, PIN and limit errors, fixed with guidance. Risk failures are blocks by fraud rules. Mixing them hides where to act.

**Your Tier 2 finding is "no difference". Why include it?** Because the obvious guess (smaller cities fail more) is not supported. Weak networks are more common in Tier 2, but bank mix offsets it, so bank and network explain failures better than city. Reporting a null result honestly is part of the job.

**How reliable is the Rs 23.1 L at-risk figure?** It is an upper bound. Customers retry, so some value is only delayed. A real analysis would track retries within five minutes and net them off.

**What would you do next?** Measure retry success to size real loss, test whether cashback changes repeat use with an A/B test, and move the pipeline to a warehouse with a scheduled refresh and alerts on the spike rule.

**What are the limits?** The data is simulated, small groups are noisy (amounts above Rs 25,000), and I measured cashback cost but not its effect on behaviour.

## 5. Make it live and findable

- [ ] Public GitHub repository with this folder; add your live link and the screenshots to the README.
- [ ] GitHub Pages: *Settings > Pages > Deploy from a branch > main > /docs*. Open the link on your phone before you apply.
- [ ] Build the Power BI report from `powerbi/BUILD_GUIDE.md`; save `PayFlow_UPI_Dashboard.pbix` in `powerbi/`.
- [ ] Record a 60 to 90 second screen walkthrough of the Power BI report and upload it as an unlisted video; link it in the README and LinkedIn.
- [ ] Export 3 Power BI page screenshots into `docs/` and reference them in the README.
- [ ] Hyperlink the live demo and GitHub on your resume, near your name.
- [ ] Pin the repository on GitHub; add the project to LinkedIn Featured.
- [ ] Rehearse sections 3 and 4 aloud until you can do them without reading.

## 6. Roles to target

Data Analyst, Business Analyst, BI Analyst, Reporting Analyst, Risk or Fraud Analyst (junior), Product Analyst at fintech, payments, banking and e-commerce companies. Tailor the first bullet: lead with "fraud flags" for risk roles, "reliability" for product roles, "dashboard" for BI roles.
