# 21 Days Excel Challenge — Fuel Queue Project

A 21-day, real-project Excel challenge built around a synthetic **Nigerian fuel queue dataset** (part of the [Nigeria By Numbers](#) content series). Instead of isolated formula drills, this challenge follows the actual workflow of a data analyst: audit messy data, clean it, analyze it, and deliver real insights — end to end.

## Why this exists

Most "learn Excel" resources teach formulas in a vacuum. This challenge starts from a deliberately messy, realistic dataset (duplicates, missing values, inconsistent formatting, logical outliers) and works through the full analyst pipeline over 3 weeks, documenting every step publicly.

## The dataset

A synthetic dataset simulating 4 weeks of crowdsourced fuel queue submissions across 8 Nigerian states, with intentional real-world messiness:
- 710 rows, ~640 unique entries + intentional duplicates
- Fields: Submission ID, Date, Time, State, Station Name, Queue Time (mins), Price per Litre (NGN), Vehicle Type, Litres Purchased, Payment Method
- Realistic constraints: fuel prices ₦1150–1300 (varying by state), litres purchased scaled to vehicle type, queue times reflecting mostly-short waits with occasional scarcity-day spikes
- Deliberately injected issues: duplicate/near-duplicate rows, missing values, inconsistent date formats, inconsistent state/station name formatting, and logical outliers (negative/impossible queue times, out-of-range prices)

See [`data/`](./data) for the raw and cleaned versions.

## Roadmap

**Week 1 — Clean, Understand, Explore (Days 1–7)**
| Day | Focus |
|---|---|
| 1–2 | Data audit — duplicates, blanks, formatting, outliers |
| 3 | Data cleaning — TRIM, PROPER, TEXT-TO-COLUMNS, removing duplicates |
| 4 | Descriptive stats tied to real questions |
| 5 | Conditional logic — IF, IFS, COUNTIFS, SUMIFS, AVERAGEIFS |
| 6 | Lookups — VLOOKUP/XLOOKUP, INDEX-MATCH |
| 7 | Mini deliverable — "Which state has the worst fuel queue problem right now?" |

**Week 2 — Structure, Analyze, Visualize (Days 8–14)**
| Day | Focus |
|---|---|
| 8–9 | PivotTables |
| 10 | PivotCharts + conditional formatting |
| 11 | Advanced formulas — nested IFs, array formulas, SUMPRODUCT |
| 12 | Data validation |
| 13 | What-if analysis — Goal Seek / Data Tables |
| 14 | Mini deliverable — formatted Excel dashboard sheet |

**Week 3 — Real Analyst Deliverables (Days 15–21)**
| Day | Focus |
|---|---|
| 15–16 | Full dashboard — KPIs, slicers, dynamic charts |
| 17 | Automate with live-updating formulas |
| 18 | Written analyst insights report |
| 19 | Polish + documentation |
| 20 | Walkthrough writeup |
| 21 | Capstone — stakeholder-style presentation |

## Progress log

Daily findings and writeups live in [`logs/`](./logs), one file per day.

## Companion ebook: Excel Formulas Made Easy

Every day of this challenge doubles as a chapter draft for **Excel Formulas Made Easy** — a beginner-friendly ebook that teaches Excel through this exact real, messy dataset instead of generic examples. Chapters live in [`ebook/`](./ebook), and each one follows the same structure: Concept → Method → Findings → Common Mistakes → Key Takeaways → Interview Corner → Exercise.

- [Chapter 1: The Data Audit](./ebook/Chapter_01_The_Data_Audit.md)

## About

Part of an ongoing self-directed data analytics journey by [@onyemazuwa](https://github.com/onyemazuwa) — also documented on TikTok and LinkedIn under **Nigeria By Numbers**.
