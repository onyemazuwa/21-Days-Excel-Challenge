# Day 1–2: Data Audit

Real analyst work starts by interrogating the data before touching a single stat formula. Before any cleaning, I audited the raw fuel queue dataset for duplicates, missing values, inconsistent formatting, and outliers.

## Method

- **Duplicate Submission_IDs:** `=SUMPRODUCT((COUNTIF(A2:A711,A2:A711)>1)*1)`
- **True duplicates vs near-duplicates:** cross-checked full-row matches (Conditional Formatting → Duplicate Values on a concatenated helper column) against ID-only matches
- **Missing values:** `=COUNTBLANK()` per column
- **Date format check:** manual scan of the Date column
- **State / Station name variants:** Data → Filter → inspected the unique-values dropdown per column
- **Outliers:** `=MIN()` / `=MAX()` on Queue_Time_Minutes and Price_Per_Litre_NGN

## Findings

| Issue | Finding |
|---|---|
| Total row count | 711 rows |
| Duplicated Submission_IDs | 139 rows (66 true duplicates — fully identical; 73 near-duplicates — same ID, different data; max cluster size: 2) |
| Missing values | ~28 blanks each in State, Station_Name, Queue_Time_Minutes, Vehicle_Type, Payment_Method; 27 in Price_Per_Litre_NGN |
| Date format | At least 4 inconsistent formats in use: DD/MM/YYYY, MM-DD-YYYY, YYYY-MM-DD, YYYY/MM/DD |
| State names | ~10+ variations of 8 real states (casing, spacing, abbreviation differences — e.g. "Lagos" / "LAGOS" / "Lagos State") |
| Station names | ~14 variations of 10 real stations (casing, leading/trailing spaces, double-internal-spaces — e.g. "NNPC Mega Station" vs "NNPC  Mega Station") |
| Outliers — Queue_Time_Minutes | Min -15 (impossible), Max 999 (16+ hrs, unrealistic even on scarcity days) |
| Outliers — Price_Per_Litre_NGN | Min ₦350, Max ₦4500 (both far outside the realistic ₦1150–1300 band) |

## Key takeaway

Two things stood out that I wouldn't have caught without checking closely:

1. **Two cells can look identical and still not match.** "NNPC Mega Station" and "NNPC  Mega Station" (double space) render the same on screen but fail exact-match formulas like `COUNTIF`. This is a real "gotcha" — duplicate checks can silently fail if you don't clean whitespace first.
2. **Not all duplicates are the same problem.** 66 rows were true duplicates (safe to delete), but 73 shared an ID with a conflicting value elsewhere — a data-entry-error signal, not a duplicate-submission signal. These need investigation, not deletion.

Next: Day 3 — cleaning (TRIM, PROPER, TEXT-TO-COLUMNS, resolving duplicates, handling blanks and outliers).
