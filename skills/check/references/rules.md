# Rules for tables and sentences

The tests to run in step 3, on every table, sheet and sentence with figures. Each failure is a lead to confirm in step 5, not a finding. For each test, keep count of how many cases it checked and how many agreed.

## Rounding tolerance

Use one tolerance throughout. For a stated result built from n parts:

- `part_step` is the unit of the last digit shown in the parts (0.1 for 12.4, 1 for 1,280, 0.01 for 7,200.00); `stated_step` is the same for the stated result.
- `worst = part_step x n / 2 + stated_step / 2`, the largest gap rounding can produce.
- `spread = sqrt((part_step² x n + stated_step²) / 12)`, the gap rounding usually produces.

Classify the gap between the stated result and the computed one:

| Class | Condition | Treat as |
|-------|-----------|----------|
| exact | gap is zero | agrees |
| rounding | gap ≤ worst and gap ≤ 2.5 x spread | agrees |
| stretch | gap ≤ worst but > 2.5 x spread | Query at most, only if the figures matter |
| off | gap > worst | lead |

Counts of people, items, days and events have no rounding: any gap is a lead. For a single derived figure (one product, one percentage), the stated value is close enough when it matches the computed value rounded to the stated number of decimals; allow one unit of the last digit for a ratio or a percentage, half a unit for a sum or a difference.

## Totals and subtotals

- **Recognising a total row.** The label contains a total word: total, grand total, sum, overall, subtotal; Gesamt, Gesamtsumme, Summe, insgesamt, Zwischensumme, Teilsumme; toplam, genel toplam, ara toplam; total général, sous-total, somme; suma, total; totale; totaal. An unlabelled last row also counts as a total when the table has at least four detail rows and at least two other columns of that row equal the sum of the rows above exactly.
- **Test.** Each column of a total row against the sum of the detail rows above it, back to the previous total or subtotal. A grand total against the sum of the subtotals, and against the sum of all detail rows. Skip "of which", "thereof", "davon" and "darunter" rows when summing.
- **Total columns.** A column that equals the sum of its neighbours in most rows. Test it in every row.
- **Knock-on.** When a part is itself a lead, substitute its recomputed value and test the total again. If the total now agrees, it is correct and belongs inside the finding for that part.
- **Average rows.** A row labelled average, mean, Durchschnitt, Mittelwert, ortalama, moyenne, media: plain mean of the detail rows above.

## Relations between columns

- **Row relations.** Look for a relation that holds in most rows of a table: sum (C = A + B), difference (C = A - B), product (C = A x B), ratio (C = A / B), ratio in % (C = A / B x 100), change in % (C = (B - A) / A x 100). A relation counts as established when it holds in at least two rows (three for ratios and changes) and in more rows than it fails. Every row where it fails is a lead.
- **Rate rows.** A row whose label names a rate ("VAT 19%", "discount 10%", "MwSt. 19 %", "KDV %20"). Test it as each of: rate of the base, base plus rate, base minus rate, rate contained in a gross base, gross base less rate. If none holds with the stated base, try the sum of the detail rows as the base before raising a lead.
- **Share columns.** A column of value / base x 100 against a total. Established when it holds in at least two of three or more rows. Each failing row is a lead.
- **Percentages that should reach 100.** A share column whose detail rows sum to between 95 and 105 but not to 100 within the rounding tolerance. Shares stored as fractions (0.36): sum between 0.95 and 1.05 but not 1. Skip multiple-answer questions.
- **Mixed units.** Two or more different units in one numeric column (mg and g, k and m, % and plain numbers).

## Spreadsheets

Read formulas, not only values. Compare formulas in relative form: `=B2*C2` in row 2 and `=B3*C3` in row 3 are the same formula.

- **Error values.** `#REF!`, `#DIV/0!`, `#N/A`, `#VALUE!`, `#NAME?` in any cell, more so in cells that feed others.
- **No stored result.** A formula cell with no cached value. Values then cannot be checked; say so.
- **Number stored as text.** A text cell that parses as a number, is not a year, and sits between numeric cells. If a total is short by exactly its value, the text cell is the cause.
- **SUM ranges.** A `SUM` whose range stops above the last number in the column, skips a numeric cell, or covers fewer rows than its sibling totals.
- **Typed over a formula.** A typed number between two cells with the same formula.
- **Formula pattern.** A formula that differs from neighbours that share one pattern.
- **Typed among formulas.** A typed number in a column with at least three formula cells, where typed cells are no more than a fifth of the column.
- **Ratio base.** A label that names a base ("% of revenue", "per unit", "pro Kopf", "je Stück") against the cell the formula actually divides by.

## Sentences

- **Weekdays.** A weekday next to a full date: compute the weekday. Weekday names in English, German, Turkish, French and Spanish.
- **Durations.** A stated duration next to a start and an end date. Weeks: allow three days either way of n x 7, counting the end day or not. Days: n equals the difference, or the difference plus one. Months: whole months, plus or minus one.
- **n of N.** "45 of 120 (37.5%)" or "45/120 (37.5%)": n x 100 / N against the stated percentage.
- **Percentage points.** A change between two percentages called "x%": test whether x is the difference in points or the relative change, and flag the one stated if it is the other.
- **Number formats.** Count clear decimal-comma figures (`1.250,00`) and clear decimal-point figures (`1,250.00`). Both present is a lead for a Query only if a reader could misread a figure.

## GRIM test

Use it when a percentage or a mean of whole-number values is reported with its sample size n but without the count.

- **Percentage.** Possible counts are the whole numbers k from 0 to n. The percentage is possible when some k gives k / n x 100 that rounds to the stated value at the stated number of decimals.
- **Mean of whole-number values** (scores, counts). The mean is possible when some whole number S gives S / n that rounds to the stated mean at the stated number of decimals. Test the two whole numbers nearest to mean x n.
- **Limits.** Not for growth rates, shares of money, or when the count itself is given (then divide the count by n). It says little when n is large and the figure has few decimals: when 10^decimals ≤ n for a mean, or 100 x 10^decimals ≤ n for a percentage, every value is possible and the test proves nothing.
