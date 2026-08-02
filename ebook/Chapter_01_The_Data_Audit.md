# Chapter 1: The Data Audit — Before You Touch a Single Formula

## Why This Chapter Comes First

Most Excel books open with a formula. `=SUM()`, `=AVERAGE()`, and off you go. This one doesn't — because that's not how real analyst work actually starts.

Before you calculate anything, you have to know what you're calculating *from*. Real data — data collected from real people, real forms, real submissions — is almost never clean. It has duplicates. It has blanks. It has dates typed three different ways by three different people. It has typos. It has values that are simply impossible.

If you skip straight to `=AVERAGE()` on data like that, you don't get an average. You get a confident-looking number that's quietly wrong — and nothing in Excel will warn you.

This chapter walks through a **data audit**: the process of interrogating a dataset before you trust it with a single calculation. We'll use a real example — a synthetic dataset built to track fuel queue times across Nigeria — and find every problem hiding inside it, the same way you would on your first day analyzing any new dataset.

---

## What Is a Data Audit?

A data audit is a structured check of your dataset for five common categories of problems:

1. **Duplicates** — the same record appearing more than once
2. **Missing values** — cells that are blank when they shouldn't be
3. **Inconsistent formatting** — the same kind of information written in different styles
4. **Outliers** — values that are technically present but don't make sense
5. **Structural issues** — wrong data types, misaligned columns, and similar problems

You're not fixing anything yet. You're just finding out what's actually in front of you, and writing it down. Think of it as a doctor running tests before writing a prescription — diagnosis before treatment.

---

## Our Example Dataset

Throughout this chapter, we'll use a fuel queue tracking dataset — 711 rows of (synthetic) submissions from people reporting how long they waited at fuel stations across 8 Nigerian states, what they paid per litre, and what vehicle they were driving.

The columns look like this:

| Column | What it holds |
|---|---|
| Submission_ID | A unique ID per entry, e.g. `FQ1001` |
| Date | The date of the submission |
| Time | The time of day |
| State | Which state the entry is from |
| Station_Name | The fuel station name |
| Queue_Time_Minutes | How long the person waited |
| Price_Per_Litre_NGN | What they paid per litre, in Naira |
| Vehicle_Type | Car, Motorcycle, SUV, etc. |
| Litres_Purchased | How much fuel they bought |
| Payment_Method | Cash, Card, or Transfer |

On the surface, it looks like a normal spreadsheet. It is not. Let's find out why.

---

## Step 1: Checking for Duplicate Records

### Why duplicates matter

Imagine one person submits the same fuel queue report twice by accident — maybe they double-clicked "Submit" on the form. If you don't catch that, your average queue time will be skewed, because that one entry now counts twice.

Worse: sometimes a duplicate isn't a perfect copy. Someone might submit the same ID twice with *slightly different values* — a typo the second time, or a genuine correction. These "near-duplicates" are more dangerous than exact duplicates, because they're harder to spot and raise the question: **which value is actually correct?**

### Finding duplicate IDs

Every submission in our dataset has a `Submission_ID` in column A. In theory, every ID should appear exactly once. Let's check.

In an empty column next to your data, enter:

```
=COUNTIF($A$2:$A$711,A2)
```

Here's what this formula does, piece by piece:
- `COUNTIF(range, criteria)` counts how many times a value appears within a range
- `$A$2:$A$711` is the full range of Submission_IDs — the dollar signs (`$`) **lock** this range so it doesn't shift when we copy the formula down (more on this in a moment)
- `A2` is the current row's ID — this one is *not* locked, so as you drag the formula down, it checks row 2's ID, then row 3's, then row 4's, and so on

If the result is `1`, that ID appears once — no problem. If the result is `2` or more, that ID is duplicated somewhere in the sheet.

**Understanding $ (absolute references) here matters a lot.** Without the dollar signs, dragging this formula down would shift *both* parts of the range, and your count would silently become wrong for every row after the first. This is one of the most common invisible bugs in beginner spreadsheets — the formula runs without error, but the numbers it gives you are meaningless.

### Counting total affected rows

To get a single number — how many rows total are involved in duplication — use:

```
=SUMPRODUCT((COUNTIF(A2:A711,A2:A711)>1)*1)
```

This is a slightly more advanced formula, so let's unpack it:
- `COUNTIF(A2:A711,A2:A711)` runs the same duplicate check as before, but for *every row at once*, producing a list of counts
- `>1` turns each of those counts into `TRUE` (if duplicated) or `FALSE` (if unique)
- `*1` converts `TRUE`/`FALSE` into `1`/`0`
- `SUMPRODUCT()` adds all of those 1s and 0s together, giving you a single total

On our fuel queue dataset, this returned **139** — meaning 139 out of 711 rows are involved in some kind of ID duplication. That's nearly 1 in 5 rows. Not a small problem.

### True duplicates vs. near-duplicates

139 rows being "duplicated" doesn't mean 139 rows are safe to delete. Some of those rows are **exact copies** — every single field matches. Others share only the ID, with different values elsewhere. These need to be handled completely differently:

- **True duplicates** (identical in every column) → safe to delete, no judgment call needed
- **Near-duplicates** (same ID, different data) → a red flag. Something went wrong at data entry — maybe a resubmission, a correction, or a genuine error. You can't just delete one; you need to *decide* which value is correct, or flag it for review.

To separate the two, you can build a helper column that combines every field into one long string:

```
=A2&"|"&B2&"|"&C2&"|"&D2&"|"&E2&"|"&F2&"|"&G2&"|"&H2&"|"&I2&"|"&J2
```

The `&` symbol joins text together, and the `"|"` characters are just separators so `"Lagos"` and `"NNPC"` next to each other don't accidentally look identical to `"LagosNNPC"`. Once every row is reduced to one combined string, running `COUNTIF` on *that* column tells you which rows are truly, 100% identical — not just similar.

In our dataset, this split the 139 duplicated rows into:
- **66 true duplicates** — fully identical, safe to remove
- **73 near-duplicates** — same ID, different data — flagged for investigation

That's a meaningfully different cleanup job than "delete 139 rows," and it's the kind of distinction that matters in real analyst work.

**A tip:** also check whether any ID appears *three or more* times, not just twice. A formula like `COUNTIF` treats a three-way duplicate the same as a two-way one unless you specifically check the count value — sort your helper column descending and scan the top values to be sure.

---

## Step 2: Checking for Missing Values

### Why missing values matter

A blank cell isn't automatically a problem — but it becomes one the moment you calculate something and don't realize a chunk of your data is silently absent. `AVERAGE()`, `SUM()`, and most Excel functions simply skip blank cells without warning you. So if 4% of your `Price_Per_Litre_NGN` column is empty, your average price is still calculated — just on 96% of the data, without telling you that.

### Finding missing values

For each column, use:

```
=COUNTBLANK(A2:A711)
```

This counts how many empty cells exist in that range. Run it once per column you care about — in our dataset, that meant checking `State`, `Station_Name`, `Queue_Time_Minutes`, `Price_Per_Litre_NGN`, `Vehicle_Type`, and `Payment_Method`.

The results came back remarkably consistent: **27–28 blank cells in each column.** That consistency is itself informative — it suggests the blanks weren't random accidents, but a systemic pattern (for example, a form field that was optional, or a data entry step that occasionally got skipped).

### What to do about missing values

You generally have three choices, and the right one depends on context:

1. **Leave them blank** — if the missing-ness itself is meaningful (e.g., "customer didn't select a payment method")
2. **Fill them** — with an average, a default, or a clearly-marked placeholder like `"Unknown"`
3. **Exclude those rows** — if the missing field is essential to your analysis and can't be reasonably guessed

There is no universally "correct" choice — this is analyst judgment, not a formula. What matters is that you make the decision deliberately, and document *why*.

---

## Step 3: Checking for Inconsistent Formatting

This is where spreadsheets get sneaky. A column can look perfectly fine to your eyes and still be a mess to Excel — because Excel doesn't see what you see. It compares text and numbers *literally*, character by character.

### Dates

Scan your date column. In our dataset, entries looked like this:

```
19/07/2026
07-21-2026
08-01-2026
2026-07-18
2026/07/18
```

At a glance, these all look like "dates." But Excel — and any formula, pivot table, or chart built on top of this column — will not treat them consistently unless they share one format. Worse, some of these are genuinely *ambiguous*: `08/01` could mean the 8th of January or the 1st of August, depending on which format was used for that specific entry, and once mixed together in a column, there's no way to be 100% certain which is which without going back to the source.

**The lesson:** always check your date column for consistency before doing anything else with it. A mixed-format date column isn't just untidy — it can quietly produce wrong sort orders, wrong `IF` comparisons, and wrong calculations.

### Text fields — the "invisible duplicate" trap

Here's a real mistake most people don't even know they're making until it costs them: **two cells can look 100% identical on your screen and still not match in a formula.**

In our dataset's `Station_Name` column, we found entries like:

```
NNPC Mega Station
NNPC  Mega Station     ← note the double space
```

To your eye, these are the same station. To Excel, they are two completely different text strings, because of one invisible extra space between "NNPC" and "Mega." Any `COUNTIF`, `VLOOKUP`, or duplicate check run on this column will treat them as different values — and quietly under-count how many times "NNPC Mega Station" really appears.

We found this same pattern across several entries:
- `" Mobil"` (leading space)
- `"mobil"` (lowercase)
- `"Total  Energies"` vs `"Total Energies"` (double space)

None of these are visible mistakes. All of them break formulas silently.

![Station Name filter dropdown showing near-identical duplicate entries](images/station_filter_dropdown.png)
*Figure 1.2 — The Station_Name filter dropdown. Look closely: "NNPC Mega Station" appears twice, once with a double space. "Total Energies" appears twice the same way. "Mobil" appears as itself, with a leading space, and in lowercase — three separate "unique" values for one real station.*

**How to check for this:** select your column, turn on **Data → Filter**, and click the column's filter dropdown. It lists every *unique* value currently in the column — if you see more entries than you expect (more "unique" stations than actually exist, more "unique" states than actually exist), you've found formatting inconsistency. In our `State` column, this method revealed roughly 10 variations of what were really only 8 actual states — different casing, different abbreviations, and stray whitespace.

![State filter dropdown showing multiple variants of the same states](images/state_filter_dropdown.png)
*Figure 1.1 — The filter dropdown on the State column. Notice "abuja (fct)" and "FCT Abuja" are the same state; "LAGOS" and "Lagos State" are the same state. Excel treats each of these as a completely separate value until they're standardized.*

We'll fix all of this with the `TRIM()` and `PROPER()` functions in the next chapter. For now, the job is just to notice it.

---

## Step 4: Checking for Outliers

### Why outliers matter

An outlier is a value that's *technically* present in your data but doesn't make logical sense. These aren't formatting problems — the cell isn't blank, isn't misspelled, isn't duplicated. It's just wrong.

### Finding outliers

The fastest first check for any numeric column is simply:

```
=MIN(range)
=MAX(range)
```

These two formulas instantly show you the extremes of a column — and extremes are usually where problems hide.

In our dataset:
- `Queue_Time_Minutes` ranged from **-15 to 999**
- `Price_Per_Litre_NGN` ranged from **₦350 to ₦4500**

Both of these should stop you immediately. A negative queue time is not just unusual — it's physically impossible; time can't be negative. A 999-minute queue (over 16 hours) is far beyond even the worst realistic fuel scarcity scenario. And a fuel price of ₦350 or ₦4500 falls dramatically outside the real-world range we know applies (roughly ₦1150–1300 at time of writing).

### What to do about outliers

Like missing values, this requires judgment, not just a formula:
- Is it a data entry mistake (e.g., a misplaced decimal, an extra digit)?
- Is it a genuine but rare event worth investigating rather than discarding?
- Is it clearly impossible, and safe to flag or remove?

The instinct to "just delete anything unusual" is a mistake in itself — sometimes the outlier is the most interesting data point in the whole set. The job of a data audit is to *notice* it and ask why, not to erase anything inconvenient.

---

## Putting It All Together: The Audit Summary

By the end of a proper data audit, you shouldn't have "fixed" anything yet — but you should have a clear, written record of every issue you found. For our fuel queue dataset, that summary looked like this:

| Issue | Finding |
|---|---|
| Total rows | 711 |
| Duplicated Submission_IDs | 139 rows (66 true duplicates, 73 near-duplicates) |
| Missing values | ~27–28 blanks per affected column |
| Date format | At least 4 inconsistent formats in use |
| State name variants | ~10 versions of 8 real states |
| Station name variants | ~14 versions of 10 real stations |
| Outliers — queue time | -15 to 999 minutes |
| Outliers — price per litre | ₦350 to ₦4500 |

This table is the deliverable. Not a chart. Not a dashboard. A clear, honest account of what's actually in the data — because every decision you make after this point depends on knowing that first.

![The completed audit summary logged in Excel](images/audit_summary_excel.png)
*Figure 1.3 — This audit, logged directly in the spreadsheet. A simple two-column table like this is often more valuable to a team than a polished chart, because it's the honest record everything else gets built on.*

---

## Common Mistakes to Avoid

A few things beginners get wrong here that aren't obvious until they've already caused a problem:

- **Forgetting the `$` in `COUNTIF` ranges.** Without locking the range with absolute references, dragging the formula down silently shifts the range for every row — the formula still runs, still returns a number, and that number is simply wrong. Excel won't warn you.
- **Deleting duplicates before separating true duplicates from near-duplicates.** If you run "Remove Duplicates" on the whole dataset immediately, you might keep the wrong version of a near-duplicate pair — or worse, lose the one record that had the correct value.
- **Trusting `AVERAGE()` or `SUM()` without checking for blanks first.** These functions skip blank cells automatically and give no warning that they did — meaning your "average" might quietly be based on far fewer rows than you think.
- **Assuming a filter dropdown's unique-value list is exhaustive proof of cleanliness.** It's a great first check, but very close variants (like a single trailing space) can sometimes be visually indistinguishable in the dropdown list itself — always cross-check suspicious columns with a `TRIM()`-based comparison too, which we'll cover in Chapter 2.
- **Deleting outliers on sight.** An implausible value is worth investigating, not automatically discarding — sometimes it's a data entry error, but sometimes it's the most important row in the sheet.

---

## Key Takeaways

1. **Always audit before you analyze.** A beautiful chart built on dirty data is still wrong — it just looks confident while being wrong.
2. **Two cells can look identical and still not match.** Whitespace and casing are invisible to your eyes but very visible to Excel.
3. **Not all duplicates are the same problem.** True duplicates are safe to delete; near-duplicates need investigation.
4. **`MIN()` and `MAX()` are your fastest outlier detectors.** Run them on every numeric column before trusting an average.
5. **A data audit produces a written record, not a fix.** Diagnosis comes before treatment.

In the next chapter, we'll take everything we found here and actually clean it: standardizing text with `TRIM()` and `PROPER()`, resolving duplicates, fixing dates, and deciding what to do with the blanks and outliers we've flagged.

---

## Interview Corner: Questions You Might Get Asked

Everything in this chapter isn't just practical skill — it's also exactly the kind of thing interviewers ask about, because it separates people who've memorized formulas from people who've actually cleaned real data. Here are questions tied to what we just covered, with the kind of answer that shows real understanding.

**Q: "How would you check for duplicates in a large dataset?"**

> Start with `COUNTIF` on the unique identifier column to flag which values repeat, then confirm with a full-row comparison — because a repeated ID doesn't always mean a repeated record. I'd distinguish between exact duplicates, which are safe to remove, and records that share an ID but have conflicting values elsewhere, which need investigation rather than deletion.

**Q: "What's the difference between a duplicate and a near-duplicate, and why does it matter?"**

> A true duplicate is identical across every field — it's redundant data with no ambiguity, safe to delete. A near-duplicate shares a key identifier but differs in other columns — that's usually a sign of a data entry error, a resubmission, or a legitimate correction, and deleting one at random risks throwing away the correct value instead of the wrong one.

**Q: "How do you handle missing data?"**

> It depends on why the value is missing and how important that field is to the analysis. Sometimes the absence itself is informative and should be left as-is. Sometimes it makes sense to fill with an average or a clear placeholder like "Unknown." And sometimes, if the field is essential and can't be reasonably estimated, the row needs to be excluded. The wrong approach is applying one rule to every column without thinking about what the data represents.

**Q: "Why might two cells that look identical not match in a formula?"**

> Usually whitespace or casing. Excel compares text literally — an extra space, a trailing space, or different capitalization makes two visually identical strings register as different values to something like `COUNTIF` or `VLOOKUP`. That's why `TRIM()` and consistent casing matter before running any duplicate or lookup logic — otherwise you can silently undercount or fail to match records that are actually the same.

**Q: "How would you spot outliers in a dataset without visualizing it first?"**

> `MIN()` and `MAX()` are the fastest first pass — they immediately surface extreme values on any numeric column. From there, I'd sanity-check the extremes against what's actually plausible for that field — a negative time or a price wildly outside a known real-world range is a signal to dig deeper before trusting any summary statistic built on that column.

**Q: "Walk me through how you'd approach a messy dataset you've never seen before."**

> Before I calculate anything, I audit it: check for duplicate records, scan for missing values column by column, look for inconsistent formatting in dates and text fields, and check the min/max of numeric columns for anything implausible. I write down what I find before fixing anything — that record becomes the basis for every cleaning decision that follows, and it's also useful documentation if anyone questions the analysis later.

---

*Exercise: Before moving to Chapter 2, try running the `COUNTBLANK()` and `MIN()`/`MAX()` checks on a spreadsheet of your own — even something small, like a list of expenses or contacts. You'll likely find at least one surprise.*
