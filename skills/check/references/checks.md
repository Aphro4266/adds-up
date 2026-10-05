# Check families

What to look for beyond the leads the script produces, family by family. Each entry says what goes wrong and how to test it. Use the families that fit the material.

## 1. Sums and totals

- **Totals in prose.** "The three work packages cost 40k, 65k and 30k, 145k in all." Add the parts: 135k.
- **Totals that live in a different place from their parts.** A total in the executive summary, the parts in an appendix table. Find both and compare.
- **Subtotals and grand totals.** A grand total must equal the sum of the subtotals and the sum of all detail rows. A grand total that also adds the subtotals in counts everything twice.
- **Two-way tables.** The bottom-right cell must equal the sum of the row totals, the sum of the column totals and the sum of all the cells inside the table. Add up the inside cells first. If that sum equals the corner cell, the inside of the table is sound and the faulty figures are row or column totals: recompute each one. If a row total and a column total both fail and one change to the cell where they cross repairs both, that cell is the error. If no single change repairs both, they are separate slips.
- **"Of which" rows.** Rows marked "of which", "thereof", "davon", "darunter" are part of the row above, not extra parts of the total.
- **Balances.** Assets equal liabilities plus equity. Opening balance plus movements equals closing balance. A closing balance should reappear as the next period's opening balance.

## 2. Derived figures

- **Line items.** Quantity × unit price = amount, for every row.
- **Tax and discounts.** Net + tax = gross, with the rate the document states. A discount is taken from the right base: "10% off, then 19% VAT" is not the same as 9% on top.
- **A rate that is named but never applied.** The text grants "12.5% discount on the list price" and no table row shows it. That is a Query: say which reading makes the table right (prices already discounted) and what the totals would be under the other one.
- **Margins and markups.** Margin = profit / revenue. Markup = profit / cost. A 25% markup is a 20% margin. Check which one the document claims and which one the numbers give.
- **Averages.** Mean = sum / count. An average of group averages is only correct when the groups are the same size; otherwise it has to be weighted. "Average order value 84" with revenue 12,600 and 140 orders gives 90.
- **Rates and ratios.** Per-unit, per-head and per-day figures: divide and compare. Conversion rate = conversions / visitors, not the other way round.
- **Currency conversion.** Amount × stated rate = converted amount. Check the direction of the rate.
- **Compound growth.** Start × (1 + rate)^years = end. A "CAGR of 20%" over five years multiplies the start by 2.49, not by 2.

## 3. Percentages

- **Percentage of a base.** n / N × 100 must give the stated percentage. Check that N is the base the sentence names: "40% of respondents" needs the number of respondents, not the number invited.
- **Parts of a whole.** Shares of one whole add to 100% within rounding. Answers to a multiple-choice question where several can be ticked do not have to.
- **Percent change.** (new − old) / old × 100. Common slips: dividing by the new value, or reversing old and new.
- **Percentage points against percent.** A rise from 28% to 34% is 6 percentage points and 21.4 percent. "An increase of 6%" is wrong for it.
- **Changes do not cancel.** Down 50% and then up 50% leaves 75% of the start. Up 20% twice is 44%, not 40%.
- **Impossible values.** A share above 100%, a decrease of more than 100%, a response rate above 100%.
- **Percentages of small samples.** With n = 12 only multiples of 8.33% can occur. Use the `grim` command for any percentage or mean reported with its n.

## 4. Counts and enumerations

- **Announced counts.** "Five reasons", "the following three steps", "seven sites" must be followed by that many.
- **Rows against a stated n.** "Table 2 lists all 14 sites" needs 14 rows.
- **Flows.** Screened − excluded = enrolled. Enrolled − withdrawn = completed. Each step has to follow from the one before, and the groups have to add up to the whole.
- **Denominators that drift.** n = 120 in the methods, 118 in the results table, 121 in the abstract. Either a reason is given or it is a Conflict.
- **Numbering.** Section, figure and table numbers that skip or repeat, and references to "Table 5" where there are four tables.

## 5. Dates and durations

- **Weekdays.** The script checks dates written with a year. For a date without a year, use the year the document implies and say that you assumed it.
- **Durations.** Compute end minus start. "1 March to 31 May, 14 weeks" is 91 days, 13 weeks. Check whether both end days are meant to be counted.
- **Order.** Milestones in sequence, a deadline after the start, a delivery date after the order date, a signature date not after the effective date unless the text says so.
- **Deadlines from rules.** "Payment within 30 days of the invoice date" against the due date that is printed.
- **Ages and periods.** Born 1972 and "aged 51" in a 2026 document. "Over the last five years (2020 to 2026)".
- **Quarters and fiscal years.** Q3 is July to September in a calendar year. If the document defines a different fiscal year, use its definition throughout.
- **Dates against the document's own time.** A launch "planned" for a date that lies before the report's own period or date. Raise it as a Query.

## 6. Units and magnitudes

- **Unit switches.** mg and g, ml and l, cm and m, days and weeks, per month and per year, net and gross in one column or between a table and the text that describes it.
- **Thousands and millions.** "Revenue 4.2m" in the text, "4,200" in a table headed "in thousands": consistent. "4,200,000" in that table: a thousand times too large. In German, "Mio." is million and "Mrd." is billion; "Billion" in German is a million million.
- **Decimal marks.** `1.250` is 1250 in German and Turkish notation and 1.25 in English notation. Decide from the surrounding figures, and raise a Query when a document mixes both.
- **Order of magnitude.** A unit price of 0.45 with 2,000 units and a line amount of 9,000. A dose ten times the others in its column.
- **Currency.** A table in EUR and a total quoted in USD with no rate given.
- **No unit at all.** A budget or quote that names no currency, or a table whose unit appears only in the text. This is not an error. Mention it under "Worth knowing".

## 7. The same figure in more than one place

- **Summary against body.** Every number in the abstract, executive summary, key messages or covering email should match the body exactly, including the rounding.
- **Text against table.** A sentence that describes a table has to quote the table's numbers.
- **Chart against text.** Chart data read by the script comes from the chart's stored values. Compare it with the title, the labels and the sentence that describes the chart. Stored shares arrive as fractions (0.36); write them as percentages when the text does.
- **Pie charts.** A pie always fills the circle. If its values total 102%, every slice is drawn smaller than its label says, and the chart matches none of the stated figures.
- **A count in the text that drives a table.** "18 people confirmed" and a quantity of 18 in every per-head row. If the headcount changes, every row changes.
- **Slide against notes.** Speaker notes often hold an older version of the number on the slide.
- **Repeated tables.** The same table in the main text and an appendix, or in the deck and the spreadsheet behind it.
- **Rounding consistency.** 12.4% in one place and 12.6% in another is a Conflict. 12.4% and "about 12%" is not.

- **Rankings and superlatives.** "The largest contributor", "the fastest-growing unit", "more than half of". Compute the ranking or the comparison from the document's figures. If the claim fails under the natural reading of its terms, it is Wrong; state the reading you used.

## 8. Basic statistics

Use only when the material reports statistics, and only what can be tested from the reported values.

- **Counts and percentages.** n / N against the stated percentage, in every table cell of the form "45 (37.5%)". The N is usually in the column header.
- **Confidence intervals.** The lower bound is below the upper bound and the point estimate lies between them. For a symmetric interval the estimate is in the middle.
- **Intervals and p-values.** A 95% interval for a difference that excludes 0, or for a ratio that excludes 1, goes with p < 0.05, and the reverse. A mismatch is a Query, because the test may differ from the interval method.
- **Ranges.** Minimum ≤ median ≤ maximum, and the mean lies between minimum and maximum.
- **Group sizes.** The groups add to the total, and the degrees of freedom fit the stated n.
- **Whole-number data.** A mean of 3.47 from 25 whole-number scores cannot occur. The `grim` command tests it. It applies only when every single value is a whole number, and it says little when n is large: with n = 240 every whole percentage is possible.

## 9. Spreadsheets

The script lists most of these as leads. Confirm each by looking at the formula and its neighbours.

- **Typed numbers where a formula belongs.** A constant between two cells with the same formula pattern. The number may have been right when typed and stale since.
- **Formulas that break the pattern.** One cell in a row or column that uses a different formula from its neighbours, such as `=B4+C4` among `=B2*C2`.
- **Ranges that stop short.** `=SUM(C2:C5)` with a number in C6, or one total covering fewer rows than its sibling totals.
- **Numbers stored as text.** `SUM` skips them, so the total is lower than it looks.
- **Hidden rows, columns and sheets.** They still feed totals. Say that they exist; do not treat them as errors.
- **Error values.** `#REF!`, `#DIV/0!`, `#N/A`, `#VALUE!` in cells that feed other cells.
- **No stored results.** A file written by a script can hold formulas with no values. The check then covers typed numbers and formula structure only; say so.
- **Units per column.** One column, one unit. A percentage column holding both 0.12 and 12.
- **Label against formula.** "Margin (% of revenue)" computed as profit divided by cost. "Average per unit" divided by the wrong count. Read each formula on a summary sheet against its label; the rules do not.
- **Typed figures on a summary sheet.** A count, a total or a rate that is typed where it could be calculated from another sheet. Recalculate it from that sheet.

## 10. Across files and sheets

- Match figures on what they measure, the period, the unit and the scope before comparing. "Revenue 2026" in a budget and "Revenue 2026" in a deck may be plan and forecast.
- **Scope that differs inside one calculation.** A profit line that subtracts twelve months of cost from ten months of revenue, or a rate built from a numerator and a denominator that cover different rows. Raise it as a Query and say what would settle it.
- Version labels and dates that differ between files ("v3" on the title page, "v2" in the footer; "as of 30 June" against "as of 30 September") explain many conflicts. Report the version mismatch first and the figures under it.
- A figure present in one file and absent from the other is not a conflict.
