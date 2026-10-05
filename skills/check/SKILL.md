---
name: check
description: Check whether the numbers in a document, slide deck, spreadsheet, PDF or pasted text add up. Use when the user asks to check, verify, sanity-check or proof the figures, totals, sums, percentages, calculations, dates or units in a report, budget, quote, invoice, proposal, table, deck or spreadsheet; asks "do the numbers add up" or "is the maths right"; wants figures cross-checked across several files; or wants a last numbers check before sending something. Applies in any language, for example "Zahlen prüfen", "stimmen die Summen", "rakamları kontrol et". Finds wrong totals, percentages that do not match their counts, broken spreadsheet formulas, unit and number-format mix-ups, impossible dates and figures that contradict each other, and shows the arithmetic for each one.
---

# Adds Up

Check the arithmetic and the internal consistency of the material and report what does not add up, with the working shown.

The job is narrow on purpose: numbers, units, dates, counts, and whether statements about them agree with each other. Leave writing style, grammar, structure and the quality of the argument alone unless the user asks for those as well.

A good check has three properties. Every finding can be verified by the reader in under a minute, because it says where the figure is, what the document states and what the arithmetic gives. Nothing is reported that the material itself does not support, because one false alarm costs more trust than a missed rounding difference. And the report says what was covered and what could not be seen.

## Steps

### 1. Get the material

Work on whatever the user points at: attached or named files, a folder of files, or text pasted into the conversation. If the request names nothing that can be checked, ask which file. Ask nothing else up front. The defaults are: check everything, and report in the language of the request.

### 2. Extract

When a code tool is available, run the bundled script on the files:

```
python3 ${CLAUDE_SKILL_DIR}/scripts/adds_up.py extract <file> [<file> ...]
```

The script is `scripts/adds_up.py`, next to this file. It uses the Python standard library only, makes no network requests and changes no file. It reads:

| File | What it reads | What it leaves out |
|------|---------------|--------------------|
| `.docx` | Paragraphs, tables, text boxes, footnotes, page headers and footers, the stored data of charts | Images, comments, tracked deletions |
| `.pptx` | Slide text, tables, speaker notes, the stored data of charts and the chart type | Pictures, the workbook embedded behind a chart |
| `.xlsx`, `.xlsm` | Every sheet including hidden ones, stored values, the formula behind each cell, how a cell is displayed | Charts, pivot caches, macros |
| `.pdf` | Text by page and line, ruled tables, and tables without rules inferred from where the numbers sit | Scanned pages and images |
| `.csv`, `.tsv`, `.md`, `.txt` | Everything, with Markdown tables as tables | |

The output has four parts for each file:

- **Content**, with a location label on every line: `P12` paragraph, `T3 R4 C2` table cell, `S5 text 2` slide text, `Sheet!B7` cell, `p3 L14` PDF line.
- **Leads**, in reading order: places where a rule fired, such as a total that differs from its parts, a product that breaks the pattern of its column, a `SUM` range that stops a row short, a weekday that does not match its date.
- **What the rules covered**: how many totals, row calculations, weekdays and so on the rules tested and how many agreed. This shows what silence means. A rule that tested nothing has confirmed nothing.
- **Notes**: what was read, what was not, hidden rows and sheets, ambiguous number formats.

Tables longer than 60 rows are printed in part: the top, the bottom and every row a lead points to. The rules still test every row. Add `--all-rows` when the rows in between need reading, for example to look for outliers, and `--json` if structured output is easier to work with. If a file type is not supported, or the script cannot run, read the file by other means and carry on. For scanned pages and figures inside images, read them visually and say in the report that those figures were read by eye.

With pasted text, or with no code tool at all, work directly from the text.

### 3. Read everything and note the figures

Read the full content, not only the leads. The rules see tables and a few sentence patterns. They do not see that the summary says 34% where the chart says 36%, that "five regions" is followed by four names, or that revenue "up 12%" is 9.6% by the table.

While reading, keep track of each figure that matters: what it measures, its unit, the period and scope it refers to, and where it appears. Pay attention to every place the same quantity appears twice.

### 4. Confirm the leads

A lead is a candidate, not a finding. For each one, look at the cell or sentence it points to and decide whether the rule applies there. Is that row really the total of those rows? Are the figures exact counts or rounded amounts? Is the "share" column a share of that total, or of something else? Then recompute it yourself.

Drop the leads that do not hold. Then work out causes before writing anything:

- **Merge leads that share a cause.** One wrong line item produces a wrong line total, subtotal and grand total. That is one finding with its knock-on effects listed.
- **Find which figure is the faulty one.** When a total disagrees with its parts, either can be wrong. Test the figures that depend on it. If a subtotal fails but the grand total equals the correct sum plus tax, the subtotal is the slip and the grand total is right.
- **Say when a suspicious figure is correct.** If a lead points at a cell that turns out to be right, such as a grand total that only looks wrong because a row total beside it is wrong, state in the finding that it is correct and should not be changed. The script marks many of these itself: a lead that says "Knock-on" or "the text cell is the cause" points at a correct cell, and belongs inside the finding for the cell it names, never in a row of its own.
- **Recompute a knock-on from the corrected value** before calling it one. A figure downstream of an error is not automatically wrong.

### 5. Run the checks the script cannot

Go through `references/checks.md` family by family: sums and derived figures stated in prose, percentages and growth rates, averages, counts and enumerations, durations and deadlines, units and magnitudes, text against tables and charts, summary against body, basic statistics, spreadsheet structure. Use the families that fit the material and skip the rest.

When a percentage or a mean is reported with a sample size but without the underlying count, test whether it can come from whole numbers at all:

```
python3 ${CLAUDE_SKILL_DIR}/scripts/adds_up.py grim --n 12 --percent 37.5
python3 ${CLAUDE_SKILL_DIR}/scripts/adds_up.py grim --n 25 --mean 3.48
```

This test applies to proportions of n and to means of whole-number values such as scores or counts. It does not apply to growth rates, to shares of money, or when the count itself is given, in which case divide the count by n. It has little power when n is large and the figure has few decimals.

### 6. Compute, never estimate

Run every sum, product, ratio, percentage and date difference as a calculation. With a code tool, use code: `Decimal` for money, `datetime` for dates. Without one, write the working out step by step and mark the finding "worked by hand". Do not report a figure from mental arithmetic, and compute a discrepancy a second way before reporting it if the result surprised you.

### 7. Several files

When more than one file is given, match the figures that measure the same thing (same quantity, period, unit and scope) across the files and compare them. Report a disagreement as a conflict and name both places. Do not decide which file is right, unless one is plainly the source of the other, such as the spreadsheet a slide was built from. In that case say so.

### 8. Report

Use the format below.

## Judging a finding

Give every finding one of three verdicts.

- **Wrong.** The document's own numbers prove it: the parts and the stated result are both there and do not agree.
- **Conflict.** Two places state the same quantity differently and the material does not settle which one is right. If a sum or a figure that depends on it does settle it, for example the table adds up to one of the two values, the other one is Wrong.
- **Query.** It looks off, but the answer depends on something the material does not contain: whether figures are rounded, how a term is defined, an input that is not shown. Say what would settle it.

A typed total that disagrees with its parts is Wrong even when nothing else depends on it. Say that the material cannot show whether a part changed after the total was typed or the total itself was mistyped.

When one cause produces both a proven error and a disagreement, report one finding and give it the stronger verdict, in the order Wrong, Conflict, Query. When the material points to one reading without proving it, say which reading reconciles the figures and that it is likely, not proven. "With 34% the slices total 100%, so the chart value is the likelier slip."

**Rounding.** A total of rounded figures can differ from the sum of the displayed parts. With n parts the gap can reach n/2 units of the last displayed digit in the worst case and is usually far smaller, so a gap of one unit is ordinary and a gap of three units across four parts is not. Counts of people, items, days or events are exact, and any difference in them is Wrong. Do not list ordinary rounding gaps as findings.

**No outside facts.** Judge the material against itself. Do not "correct" a market size, an exchange rate, a price or a clinical value from memory. If the right figure cannot be derived from the material, say which input is missing. Calendar facts (the weekday of a date, the days in a month), unit conversions and arithmetic are fair to use. A tax rate is fair only when the document states it.

**Ambiguous notation.** A figure such as `1.234` or `1,234` can be read two ways, and so can a date such as `03/04/2026`. Read it the way the rest of the document reads and say so only when the reading changes a finding. Two number formats in one document deserve a Query only when a reader could misread a figure.

**Do not change the file.** The check ends with a report. If the user then asks for the errors to be fixed, change only the figures agreed, leave everything else untouched and list each change. Where a fix needs a choice, for example which of two conflicting numbers is correct, ask.

## Report format

Answer in the language of the user's request. Translate the column headings, the verdicts and the "Checked" labels into that language and keep the title "Adds Up". Use these terms so that reports stay consistent:

| English | German | Turkish |
|---------|--------|---------|
| Wrong | Falsch | Yanlış |
| Conflict | Widerspruch | Çelişki |
| Query | Rückfrage | Teyit gerekli |
| Where / It says / The numbers give / Verdict | Wo / Dort steht / Die Zahlen ergeben / Urteil | Nerede / Belgede yazan / Hesap sonucu / Sonuç |
| Checked / Not checked / Worth knowing | Geprüft / Nicht geprüft / Gut zu wissen | Kontrol edildi / Kontrol edilemedi / Bilmekte fayda var |

Write numbers the way the document writes them. For a spreadsheet, use what the cell displays where the script shows it ("shown as 49.0%") and otherwise plain digits with thousands separators. Use `x` and `-` for the operators, or the typographic signs if the user's own writing uses them.

Structure:

1. **Title and opening line.** The opening line counts findings, where one finding is one root cause: "3 findings: 2 wrong, 1 to confirm", or "14 findings, all wrong". Add one sentence on the effect on the document's main figures when there is one. Give a range when the effect depends on how a finding is settled.
2. **Table of findings.** Order by verdict (Wrong, Conflict, Query). Inside each verdict, put the findings that change the document's main totals or headline claims first and the rest in reading order.
3. **Working.** One short paragraph per finding: the calculation, the size and direction of the error, the knock-on effects, and what cannot be told from the document.
4. **Checked / Not checked.** One or two sentences. Count roughly: tables, statements with figures, dates. Under "Not checked" name what could not be read (images, scanned pages) and inputs with no source inside the material, but only inputs that carry a headline figure.
5. **Worth knowing**, only when it applies and never more than three items. Anything a calculation shows to be wrong, including a weekday, a duration or a count, is a finding and belongs in the table. This line is for things that are not errors but will make the recipient stop, such as hidden rows that feed a total, a missing currency or unit, or a relation between the document's own figures that looks implausible (costs above revenue in three units). State the observation and leave the judgement to the reader.

End there. Add no praise, no general advice and no closing offer.

```
**Adds Up: quote-2026-117.docx**

3 findings: 2 wrong, 1 to confirm. The order total is 200.00 too high.

| # | Where | It says | The numbers give | Verdict |
|---|-------|---------|------------------|---------|
| 1 | Price table, row "Service plan" | Amount 7,400.00 | 15 x 480.00 = 7,200.00 | Wrong |
| 2 | Summary, second paragraph | "up 12% on Q2" | (1,480 - 1,350) / 1,350 = 9.6% | Wrong |
| 3 | Segment table, "Share" column | Shares total 99.8% | Rounding of four parts explains 0.2 | Query |

**1. Service plan amount.** 15 x 480.00 = 7,200.00. The table states 7,400.00, which is 200.00 too high. The table total of 14,620.00 adds the stated amounts correctly, so it carries the same error and becomes 14,420.00. The document does not show whether the quantity, the unit price or the amount was mistyped.

**2. Growth on Q2.** Q3 revenue is 1,480 and Q2 is 1,350, both in the regional table. (1,480 - 1,350) / 1,350 = 9.6%, not 12%.

**3. Segment shares.** 20.1 + 39.9 + 30.2 + 9.6 = 99.8. Rounding produces gaps of this size. Confirm only if the shares are meant to be exact.

**Checked:** 3 tables, 14 statements with figures, 2 dates. **Not checked:** the chart on page 4 is an image, so its values could not be read.

**Worth knowing:** the quote names no currency.
```

When nothing is wrong, keep the title, say so in one sentence and give the coverage:

```
**Adds Up: offsite-budget.docx**

No findings. Every figure I could test agrees with the others.

**Checked:** 2 tables (line amounts, totals, shares), 2 statements with figures, 1 date. **Not checked:** nothing was left unread.
```

Rules for the report:

- Give the location in the reader's terms: section heading, table caption or the table's first header cell, row label, slide number, sheet and cell. Word files carry no page numbers in the extraction, so use headings and row labels. The script's labels (`P12`, `T3`) may follow in brackets but do not help on their own.
- One finding per root cause, with its knock-on effects inside it. When no single correction reconciles a table, report each independent discrepancy as its own finding and say that they may be linked.
- When a correction depends on how another finding is settled, say so in one sentence.
- For spreadsheets, name the cell and quote the formula. When a main total is hit by more than one finding, add one line after the working with the combined effect.
- Keep queries to what a careful reader would want confirmed. A report with twenty queries hides the two real errors.
- With more than about fifteen findings, lead with the ones that change a main total and group the rest by table, slide or sheet.
- State plainly when part of the material could not be checked and why. Coverage that is left unsaid reads as "all clear".
