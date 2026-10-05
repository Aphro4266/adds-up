# Adds Up

Adds Up is a plugin for Claude that checks whether the numbers in a document add up. Give it a report, a quote, a slide deck, a spreadsheet or a PDF and it tells you which totals are wrong, which percentages do not match their counts and which figures contradict each other, with the arithmetic shown for every finding.

It exists because a spell-checker will not tell you that the table total is 200 too high, that 156 of 240 is 65% and not 62%, or that the meeting on "Wednesday, 13 October" falls on a Tuesday.

## What it checks

| Family | Examples |
|--------|----------|
| Sums | Row and column totals, subtotals, grand totals, totals quoted in the text |
| Derived figures | Quantity × price, tax and discounts, averages, growth rates, margins |
| Percentages | n of N against the stated %, shares that should reach 100%, percent against percentage points |
| Counts | "Five regions" followed by four names, patient or order flows, a sample size that drifts |
| Dates | Weekday against date, dates that do not exist, durations, deadlines |
| Units and formats | mg against g, thousands against millions, `1.250,00` next to `1,250.00` |
| Consistency | Summary against body, text against table, chart against caption, slide against notes, file against file |
| Spreadsheets | Typed numbers where a formula belongs, formulas that break a pattern, `SUM` ranges that stop a row short, numbers stored as text |
| Statistics | Percentages and means that cannot come from the stated sample size, confidence intervals, ranges |

Every finding gets one of three verdicts. **Wrong** means the document's own numbers prove it. **Conflict** means two places disagree and the document does not say which is right. **Query** means the answer depends on something the document does not contain, such as whether figures were rounded.

## How to use it

Attach or point to the files and ask in your own words:

- "Do the numbers in this proposal add up?"
- "Check the totals and percentages in the attached budget."
- "Cross-check the figures in the deck against the spreadsheet."
- "Stimmen die Summen in diesem Angebot?"

In Cowork and Claude Code you can also call the skill directly with `/adds-up:check`.

It reads Word (`.docx`), PowerPoint (`.pptx`), Excel (`.xlsx`, `.xlsm`), CSV, Markdown, plain text and PDF, and it works on text pasted into the conversation. Number formats with a decimal point and with a decimal comma are both understood, and so are total labels and weekday and month names in English, German, Turkish, French and Spanish.

## What you get back

```
Adds Up: q3-report.docx

4 findings: 3 wrong, 1 to confirm. The order total is 200.00 too high.

#  Where                           It says             The numbers give                 Verdict
1  Order book, row "Service plan"  Amount 7,400.00     15 x 480.00 = 7,200.00           Wrong
2  Summary, first paragraph        "up 12% on Q2"      (1,480 - 1,350) / 1,350 = 9.6%   Wrong
3  Revenue by region, row "East"   Total 1,290         410 + 425 + 445 = 1,280          Wrong
4  Segment table, "Share" column   Shares total 99.8%  Rounding of four parts           Query

Checked: 3 tables, 14 statements with figures, 2 dates.
Not checked: the chart on page 4 is an image, so its values could not be read.
Worth knowing: the report names no currency.
```

Each finding comes with its working and its knock-on effects, for example the table total that carries a wrong line item, and with a note when a figure that looks wrong is in fact correct and should be left alone. The report ends by saying what was covered and what could not be seen. It is written in the language you asked in. Adds Up does not change your file. If you want the errors corrected, ask after you have read the report.

## What it does not do

It judges a document against itself. It does not know whether your market size, exchange rate or clinical result is true, and it will not replace a figure from memory. It does not comment on wording, style or strategy. Figures that exist only inside pictures are read by eye and reported as such. It is a second pair of eyes, not an audit, and a clean report means that nothing was found in what could be checked.

## How it works, and what it runs

The plugin contains one skill and one Python script, `skills/check/scripts/adds_up.py`. There is no connector, no hook and no server.

When a code tool is available, Claude runs the script on the files you named. The script pulls out the text, the tables, the chart data, the speaker notes and, for spreadsheets, the formula behind each cell. It then applies fixed rules, for example "a row labelled Total should equal the rows above it", and prints the places where a rule fired. Claude confirms each of those against the document, reads the rest for what rules cannot see, recalculates every figure with code and writes the report.

The script uses the Python standard library only. It reads the files it is given and prints to the conversation. It writes no file, installs nothing, sends nothing anywhere and makes no network request. For PDFs it uses `pdfplumber` or `pypdf` if one of them is already installed and otherwise leaves the PDF for Claude to read directly. Your documents go nowhere beyond the Claude session you are already working in, and the plugin stores nothing.

Where no code tool is available, Claude works from the text and writes out the arithmetic step by step.

## Try it

`examples/quarterly-report.md` is a one-page report with ten planted errors. Ask Claude to check it, or paste its contents into a chat. A full check finds:

1. 13 October 2026 is a Tuesday, not a Wednesday.
2. Growth on Q2 is 9.6%, not 12%.
3. 156 of 240 is 65%, not 62%.
4. "Five regions" is followed by four.
5. The East row adds up to 1,280, not 1,290. The grand total of 4,120 is correct.
6. The service plan line is 15 × 480.00 = 7,200.00, not 7,400.00, and the order total carries the error.
7. Small business is 72 of 240, which is 30%, not 35%, so the shares add up to 105%.
8. 1 March to 31 May 2027 is 13 weeks, not 14.
9. 31 June does not exist.
10. A rise from 28% to 34% is 6 percentage points, or 21.4%, not "6%".

## How much it helps

I tested the plugin on two generated files with a known answer key: a 25-table annual review with 13 planted errors and a three-sheet workbook with 8. Each setting ran once or twice, so read these as an indication and not as a benchmark.

| Model | Annual review, without | with Adds Up | Workbook, without | with Adds Up |
|-------|-----------------------:|-------------:|------------------:|-------------:|
| Claude's fastest model (Haiku) | 3 to 4 of 13 | 12 to 13 of 13 | 4 of 8 | 6 of 8 |
| Mid-size model (Sonnet) | 13 of 13 | 13 of 13 | 8 of 8 | 8 of 8 |
| Largest model | 13 of 13 | 13 of 13 | 8 of 8 | 8 of 8 |

With a code tool, the larger models found every planted error with or without the plugin. What the plugin changed for them was the report: one finding per cause, a verdict on each, a statement of what was and was not checked, and no rewriting of the file. For the fastest model it changed what was found. The files were written by the plugin's author, which favours the plugin, and none was longer than 300 rows or 60 paragraphs.

## Requirements

Python 3.8 or later for the script, in any environment where Claude can run code. Nothing needs to be installed.

## Development

```
python3 -m unittest discover -s tests -v
python3 skills/check/scripts/adds_up.py extract examples/quarterly-report.md
```

Bug reports are most useful with the smallest table or sentence that shows the problem: a rule that fired where it should not have, or an error it walked past.

## Licence

MIT. See `LICENSE`.
