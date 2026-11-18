# Reading the files

What to read in each file type in step 2, and what is easy to miss. With a code tool, read the file with code; libraries the environment already has, such as python-docx, python-pptx, openpyxl, pdfplumber or pypdf, are fine. Install nothing without asking. Word, PowerPoint and Excel files are zip archives of XML, so the standard library's `zipfile` and an XML parser reach everything listed below when no library is present.

## Word (`.docx`, `.docm`)

| Read | Where it sits in the archive |
|------|------------------------------|
| Paragraphs and tables, in reading order | `word/document.xml` |
| Text boxes | inside `document.xml`, in drawing and text-box elements; libraries often skip them |
| Footnotes | `word/footnotes.xml` |
| Headers and footers | `word/header*.xml`, `word/footer*.xml` |
| Chart data (categories, series, stored values) | `word/charts/chart*.xml` |

Leave out tracked deletions. Comments are not content. Images cannot be read as text: look at them, or list them under "Not checked". Word files carry no page numbers, so locate figures by heading, paragraph and table row.

## PowerPoint (`.pptx`, `.pptm`)

| Read | Where it sits in the archive |
|------|------------------------------|
| Slide text and tables, by slide number | `ppt/slides/slide*.xml`, ordered as listed in `ppt/presentation.xml` |
| Speaker notes | `ppt/notesSlides/notesSlide*.xml` |
| Chart data and chart type | `ppt/charts/chart*.xml` |

Note the chart type: a pie or doughnut always fills the circle, which matters when its values do not reach 100%. The workbook embedded behind a chart is not read; the stored values in the chart part are what the slide shows.

## Excel (`.xlsx`, `.xlsm`)

Read every sheet, including hidden ones, and for each cell:

- the stored value, and the formula behind it if there is one;
- how the cell is displayed (its number format, so that 0.49 can be reported as "shown as 49.0%");
- whether the cell is text, a number, an error value, or a formula with no stored result.

With openpyxl, load the workbook twice: once for formulas and once with `data_only=True` for stored values. Note hidden rows, columns and sheets, because they still feed totals. Charts, pivot caches and macros are not read.

## PDF

Read text by page and line. Take ruled tables as tables. Where a table has no rules, line up the numbers by their horizontal position to rebuild the columns, and say in the report that the table structure was inferred. Scanned pages have no text layer: read them visually and say so.

## CSV, TSV, Markdown, plain text

Read everything. Treat Markdown tables as tables. For CSV, check the delimiter and the decimal mark before parsing: a German export often uses `;` as the delimiter and `,` as the decimal mark.

## Pasted text

Work from the text as given. A table pasted from a spreadsheet arrives as tab-separated columns.

## Large tables

Apply every rule to every row. When writing out content for your own reading, more than about 60 rows can be shown in part (the top, the bottom and every row a lead points to), but read the rows in between whenever outliers or a broken pattern could hide there.
