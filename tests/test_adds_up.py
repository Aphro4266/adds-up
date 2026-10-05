"""Tests for the Adds Up script. Standard library only.

Run from the repository root:
    python3 -m unittest discover -s tests -v
"""

import os
import sys
import tempfile
import unittest
import zipfile
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "skills", "check", "scripts"))

import adds_up as au  # noqa: E402


def leads_for(rows, conv="us", **kw):
    au.FMT["conv"] = conv
    au.FMT["group"] = True
    return au.check_table(au.Table("T1", rows, **kw), conv)


def kinds(leads, strength=None):
    return sorted(l["kind"] for l in leads if strength is None or l["strength"] == strength)


class ParseNumbers(unittest.TestCase):
    def val(self, text, conv="us"):
        n = au.parse_cell(text, conv)
        return None if n is None else n.value

    def test_plain_and_grouped(self):
        self.assertEqual(self.val("1,234.56"), Decimal("1234.56"))
        self.assertEqual(self.val("1.234,56"), Decimal("1234.56"))
        self.assertEqual(self.val("1 234,5"), Decimal("1234.5"))
        self.assertEqual(self.val("1'234"), Decimal("1234"))
        self.assertEqual(self.val("12,5"), Decimal("12.5"))
        self.assertEqual(self.val("0,125"), Decimal("0.125"))
        self.assertEqual(self.val("1,234,567"), Decimal("1234567"))
        self.assertEqual(self.val("1.234.567"), Decimal("1234567"))

    def test_ambiguous_follows_convention(self):
        self.assertEqual(self.val("1.234", "us"), Decimal("1.234"))
        self.assertEqual(self.val("1.234", "eu"), Decimal("1234"))
        self.assertEqual(self.val("1,234", "us"), Decimal("1234"))
        self.assertEqual(self.val("1,234", "eu"), Decimal("1.234"))
        self.assertEqual(au.parse_cell("1,234").style, "amb")

    def test_signs_units_percent(self):
        self.assertEqual(self.val("(700)"), Decimal("-700"))
        self.assertEqual(self.val("− 12.5"), Decimal("-12.5"))
        self.assertEqual(self.val("€ 3,40", "eu"), Decimal("3.40"))
        self.assertEqual(self.val("850,00 €", "eu"), Decimal("850.00"))
        self.assertEqual(self.val("EUR 1,200"), Decimal("1200"))
        n = au.parse_cell("37.5%")
        self.assertTrue(n.pct)
        self.assertEqual(n.decimals, 1)
        self.assertEqual(au.parse_cell("250 mg").unit, "mg")

    def test_not_numbers(self):
        for text in ("12.03.2026", "2026-03-12", "10:30", "5-7", "Q1 2026", "1.2.3", "n/a", "", "12 working days"):
            self.assertIsNone(au.parse_cell(text), text)

    def test_four_digit_decimals_are_numbers_not_dates(self):
        self.assertEqual(self.val("1333.3"), Decimal("1333.3"))
        self.assertEqual(self.val("2026.5"), Decimal("2026.5"))
        self.assertIsNone(au.parse_cell("2026-03"))
        self.assertIsNone(au.parse_cell("2026.03.12"))

    def test_yearlike(self):
        self.assertTrue(au.parse_cell("2026").yearlike)
        self.assertFalse(au.parse_cell("2,026").yearlike)


class TotalRows(unittest.TestCase):
    def test_correct_table_is_silent(self):
        rows = [["Item", "Q1", "Q2"], ["A", "10", "20"], ["B", "30", "40"], ["Total", "40", "60"]]
        self.assertEqual(leads_for(rows), [])

    def test_wrong_total(self):
        rows = [["Item", "Q1", "Q2"], ["A", "100", "200"], ["B", "300", "400"], ["Total", "400", "650"]]
        leads = leads_for(rows)
        self.assertEqual(kinds(leads), ["total-row"])
        self.assertIn("600", leads[0]["message"])
        self.assertEqual(leads[0]["where"], "T1 R4 C3")

    def test_german_labels_and_format(self):
        rows = [["Position", "Betrag"], ["A", "1.250,00 €"], ["B", "708,00 €"], ["C", "850,00 €"],
                ["Summe", "2.818,00 €"]]
        leads = leads_for(rows, "eu")
        self.assertEqual(kinds(leads), ["total-row"])
        self.assertIn("2.808,00", leads[0]["message"])

    def test_turkish_label(self):
        rows = [["Kalem", "Tutar"], ["A", "100"], ["B", "200"], ["Toplam", "310"]]
        self.assertEqual(kinds(leads_for(rows)), ["total-row"])

    def test_subtotals_and_grand_total(self):
        rows = [["Item", "Cost"], ["Laptop", "900"], ["Screen", "300"], ["Subtotal hardware", "1,200"],
                ["Licence", "250"], ["Support", "150"], ["Subtotal software", "400"], ["Grand total", "1,600"]]
        self.assertEqual(leads_for(rows), [])
        rows[-1][1] = "1,700"
        self.assertEqual(kinds(leads_for(rows)), ["total-row"])

    def test_sections_with_their_own_totals(self):
        rows = [["Item", "Amount"], ["Product sales", "900"], ["Services", "300"], ["Total revenue", "1,200"],
                ["Salaries", "500"], ["Rent", "200"], ["Total costs", "700"], ["Profit", "500"]]
        self.assertEqual(leads_for(rows), [])

    def test_total_stated_first(self):
        rows = [["Group", "n"], ["Total", "120"], ["Women", "70"], ["Men", "50"]]
        self.assertEqual(leads_for(rows), [])
        rows[1][1] = "125"
        self.assertEqual(kinds(leads_for(rows)), ["total-row"])

    def test_of_which_rows_are_not_parts(self):
        rows = [["Region", "Sales"], ["Europe", "500"], ["of which Germany", "200"], ["Asia", "300"], ["Total", "800"]]
        self.assertEqual(leads_for(rows), [])

    def test_small_counts_off_by_one_are_reported(self):
        rows = [["Team", "2025", "2026"], ["Sales", "12", "15"], ["Support", "8", "11"], ["Total", "20", "27"]]
        leads = leads_for(rows)
        self.assertEqual(kinds(leads, "normal"), ["total-row"])

    def test_rounded_figures_are_weak(self):
        rows = [["Region", "Revenue (m)"], ["A", "12.3"], ["B", "8.4"], ["C", "5.6"], ["D", "3.3"], ["Total", "29.7"]]
        leads = leads_for(rows)   # parts sum to 29.6
        self.assertEqual(kinds(leads, "weak"), ["total-row"])
        self.assertEqual(kinds(leads, "normal"), [])

    def test_financial_statement_is_silent(self):
        rows = [["", "2025", "2026"], ["Revenue", "1,200", "1,350"], ["Cost of sales", "(700)", "(790)"],
                ["Gross profit", "500", "560"], ["Operating expenses", "(300)", "(320)"],
                ["Operating profit", "200", "240"]]
        self.assertEqual(leads_for(rows), [])

    def test_named_quantities_are_not_sums(self):
        rows = [["Item", "Value"], ["Total revenue", "427,575"], ["Total cost", "284,189"], ["Profit", "143,386"],
                ["Units sold", "11,455"], ["Average revenue per unit", "37.33"], ["Cost centres", "40"]]
        self.assertEqual(leads_for(rows), [])
        rows = [["Item", "Value"], ["Units", "100"], ["Price", "5"], ["Total revenue", "500"]]
        self.assertEqual(leads_for(rows), [])

    def test_annual_total_column_called_year(self):
        rows = [["Cost centre", "H1", "H2", "Year"], ["A", "1,200", "1,300", "2,500"], ["B", "800", "900", "1,700"],
                ["C", "400", "450", "950"], ["D", "300", "300", "600"]]
        self.assertEqual(kinds(leads_for(rows)), ["total-column"])

    def test_rate_rows(self):
        rows = [["Item", "Amount"], ["Goods", "1,000.00"], ["Service", "500.00"], ["Subtotal", "1,500.00"],
                ["VAT 19%", "285.00"], ["Total", "1,785.00"]]
        self.assertEqual(leads_for(rows), [])
        rows[4][1] = "295.00"
        leads = leads_for(rows)
        self.assertIn("rate-row", kinds(leads))


class ColumnsAndRelations(unittest.TestCase):
    def test_corner_total_that_agrees_with_the_detail_cells_is_left_alone(self):
        rows = [["Region", "Q1", "Q2", "Q3", "Total"], ["North", "310", "330", "365", "1,005"],
                ["South", "280", "295", "340", "915"], ["East", "410", "425", "445", "1,290"],
                ["West", "290", "300", "330", "920"], ["Total", "1,290", "1,350", "1,480", "4,120"]]
        leads = leads_for(rows)
        self.assertEqual([l["where"] for l in leads], ["T1 R4 C5"])

    def test_total_that_is_only_a_knock_on_is_marked(self):
        rows = [["Segment", "Customers", "Share"], ["A", "48", "20%"], ["B", "96", "40%"], ["C", "72", "35%"],
                ["D", "24", "10%"], ["Total", "240", "100%"]]
        leads = leads_for(rows)
        total = [l for l in leads if l["where"] == "T1 R6 C3"][0]
        self.assertEqual(total["strength"], "weak")
        self.assertIn("Knock-on", total["message"])
        self.assertIn("should not be changed", total["message"])

    def test_total_column(self):
        rows = [["Region", "Q1", "Q2", "Q3", "Total"], ["N", "1", "2", "3", "6"], ["S", "4", "5", "6", "15"],
                ["E", "7", "8", "9", "25"]]
        leads = leads_for(rows)
        self.assertEqual(kinds(leads), ["total-column"])
        self.assertEqual(leads[0]["where"], "T1 R4 C5")

    def test_quantity_times_price(self):
        rows = [["Item", "Qty", "Price", "Amount"], ["A", "40", "125.00", "5,000.00"],
                ["B", "120", "18.50", "2,220.00"], ["C", "15", "480.00", "7,400.00"]]
        leads = leads_for(rows)
        self.assertEqual(kinds(leads), ["row-relation"])
        self.assertIn("7,200.00", leads[0]["message"])

    def test_share_column(self):
        rows = [["Segment", "Customers", "Share"], ["A", "48", "20%"], ["B", "96", "40%"], ["C", "72", "35%"],
                ["D", "24", "10%"], ["Total", "240", "100%"]]
        leads = leads_for(rows)
        self.assertIn("share", kinds(leads))

    def test_percent_column_without_total(self):
        rows = [["Answer", "Share"], ["Yes", "45%"], ["No", "35%"], ["Unsure", "25%"]]
        self.assertEqual(kinds(leads_for(rows)), ["percent-sum"])
        rows[3][1] = "20%"
        self.assertEqual(leads_for(rows), [])

    def test_growth_rates_are_not_shares(self):
        rows = [["Year", "Growth"], ["2023", "4%"], ["2024", "6%"], ["2025", "5%"]]
        self.assertEqual(leads_for(rows), [])

    def test_mixed_units(self):
        rows = [["Product", "Dose"], ["A", "250 mg"], ["B", "500 mg"], ["C", "1 g"]]
        self.assertEqual(kinds(leads_for(rows)), ["mixed-units"])

    def test_plain_data_table_is_silent(self):
        rows = [["Site", "Age", "Weight", "Height"], ["1", "34", "71", "168"], ["2", "51", "83", "175"],
                ["3", "29", "64", "181"], ["4", "45", "90", "170"]]
        self.assertEqual(leads_for(rows), [])


class Prose(unittest.TestCase):
    def check(self, text):
        return au.check_prose([("P1", text)])

    def test_weekdays(self):
        self.assertEqual(self.check("Meeting on Monday, 5 October 2026."), [])
        self.assertEqual(kinds(self.check("Meeting on Tuesday, 5 October 2026.")), ["weekday"])
        self.assertEqual(self.check("Termin am Montag, 05.10.2026."), [])
        self.assertEqual(kinds(self.check("Termin am Dienstag, den 5. Oktober 2026.")), ["weekday"])
        self.assertEqual(self.check("Toplantı 5 Ekim 2026 Pazartesi günü."), [])
        self.assertEqual(kinds(self.check("Toplantı 5 Ekim 2026 Salı günü.")), ["weekday"])
        self.assertEqual(self.check("Due Monday, October 5, 2026."), [])
        # 10/05/2026 is a Monday read month first and a Sunday read day first: either reading is accepted
        self.assertEqual(self.check("Due Monday 10/05/2026."), [])
        self.assertEqual(kinds(self.check("Due Friday 10/05/2026.")), ["weekday"])

    def test_impossible_dates(self):
        self.assertEqual(kinds(self.check("Launch on 31 June 2026.")), ["date"])
        self.assertEqual(kinds(self.check("Frist: 30.02.2026")), ["date"])
        self.assertEqual(kinds(self.check("Leap day 29 February 2027")), ["date"])
        self.assertEqual(self.check("Leap day 29 February 2028"), [])

    def test_durations(self):
        self.assertEqual(self.check("The project runs from 1 March 2026 to 31 May 2026, a duration of 13 weeks."), [])
        self.assertEqual(kinds(self.check("The project runs from 1 March 2026 to 31 May 2026, a duration of 14 weeks.")),
                         ["duration"])
        self.assertEqual(self.check("Laufzeit: 01.03.2026 bis 31.05.2026 (92 Tage)."), [])
        self.assertEqual(kinds(self.check("Laufzeit: 01.03.2026 bis 31.05.2026 (95 Tage).")), ["duration"])
        self.assertEqual(kinds(self.check("Valid from 2026-06-01 to 2026-05-01.")), ["duration"])

    def test_percentage_points(self):
        leads = self.check("Our share grew from 28% to 34%, an increase of 6%.")
        self.assertEqual(kinds(leads), ["percentage-points"])
        self.assertIn("21.4%", leads[0]["message"])
        self.assertEqual(self.check("Our share grew from 28% to 34%, an increase of 6 percentage points."), [])
        self.assertEqual(self.check("Our share grew from 28% to 34%, an increase of 21%."), [])
        self.assertEqual(kinds(self.check("Der Anteil stieg von 28 % auf 34 %, ein Anstieg um 6 %.")),
                         ["percentage-points"])

    def test_n_of_n(self):
        self.assertEqual(self.check("45 of 120 patients (37.5%) healed."), [])
        self.assertEqual(self.check("45 of 120 patients (38%) healed."), [])
        self.assertEqual(kinds(self.check("45 of 120 patients (40%) healed.")), ["n-of-N"])
        self.assertEqual(kinds(self.check("Response rate 62% (156/240).")), ["n-of-N"])
        self.assertEqual(self.check("Response rate 65% (n=156/240)."), [])
        self.assertEqual(kinds(self.check("18 von 24 Teilnehmern (70 %) haben zugesagt.")), ["n-of-N"])


class Spreadsheets(unittest.TestCase):
    def test_relative_formula_patterns(self):
        pattern = au.to_r1c1("SUM(B3:D3)", 3, 5)
        self.assertEqual(pattern, au.to_r1c1("SUM(B5:D5)", 5, 5))
        self.assertEqual(au.from_r1c1(pattern, 4, 5), "SUM(B4:D4)")
        self.assertEqual(au.from_r1c1(au.to_r1c1("D2/$D$5", 2, 5), 3, 5), "D3/$D$5")

    def test_display_formats(self):
        self.assertEqual(au._fmt_shape("0.0%"), (1, False))
        self.assertEqual(au._fmt_shape("#,##0.00"), (2, True))
        self.assertIsNone(au._fmt_shape("General"))
        self.assertEqual(au.shown_as("0.489901160292", (1, False), True), "49.0%")
        self.assertEqual(au.shown_as("1234567.891", (2, True), False), "1,234,567.89")


class Grim(unittest.TestCase):
    def test_percent(self):
        self.assertFalse(au.grim(12, "37.5", True)[0])
        self.assertTrue(au.grim(12, "41.7", True)[0])
        self.assertTrue(au.grim(28, "36", True)[0])
        self.assertTrue(au.grim(8, "37,5", True)[0])
        self.assertEqual(au.grim(240, "65", True)[0], [155, 156, 157])

    def test_mean(self):
        self.assertTrue(au.grim(25, "3.48", False)[0])
        self.assertFalse(au.grim(25, "3.47", False)[0])


class Files(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def write(self, name, text):
        path = os.path.join(self.dir, name)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)
        return path

    def test_markdown(self):
        path = self.write("a.md", "# Costs\n\n24 of 32 (80%) attended.\n\n"
                                  "| Item | Units | Unit cost | Cost |\n|---|---|---|---|\n"
                                  "| Venue | 1 | 1,800.00 | 1,800.00 |\n| Catering | 32 | 42.50 | 1,360.00 |\n"
                                  "| Print | 32 | 11.25 | 350.00 |\n| Travel | 3 | 310.00 | 930.00 |\n"
                                  "| **Total** | | | **4,440.00** |\n")
        res = au.process(path)
        self.assertIsNone(res["error"])
        self.assertEqual(kinds(res["leads"]), ["n-of-N", "row-relation"])
        self.assertIn("extraction", au.render([res], 50))

    def test_semicolon_csv_with_decimal_commas(self):
        path = self.write("k.csv", "Kostenstelle;Plan;Ist\nPersonal;120.000,00;124.500,00\n"
                                   "Miete;18.000,00;18.000,00\nReisen;9.500,00;7.250,00\nSumme;147.500,00;149.850,00\n")
        res = au.process(path)
        self.assertEqual(res["convention"], "eu")
        self.assertEqual([l["where"] for l in res["leads"]], ["C5"])

    def test_mixed_formats_are_noted(self):
        path = self.write("m.txt", "Preis 1.250,00 EUR, Rabatt 12.5% und Versand 4,90 EUR.\n")
        res = au.process(path)
        self.assertEqual(kinds(res["leads"]), ["number-format"])

    def test_unsupported_and_missing(self):
        path = self.write("old.doc", "x")
        self.assertIn("Unsupported", au.process(path)["error"])

    def workbook(self, name, sheets):
        """Write a small workbook. sheets: {sheet name: [row, ...]}, a cell is text, number or ("=FORMULA", value)."""
        ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        path = os.path.join(self.dir, name)
        with zipfile.ZipFile(path, "w") as z:
            tags, rels = [], []
            for i, (title, rows) in enumerate(sheets.items(), 1):
                xml = []
                for r, row in enumerate(rows, 1):
                    cells = []
                    for c, v in enumerate(row):
                        ref = "%s%d" % (au.num_to_col(c + 1), r)
                        if v is None:
                            continue
                        if isinstance(v, tuple):
                            cells.append('<c r="%s"><f>%s</f><v>%s</v></c>' % (ref, v[0].lstrip("="), v[1]))
                        elif isinstance(v, str):
                            cells.append('<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, v))
                        else:
                            cells.append('<c r="%s"><v>%s</v></c>' % (ref, v))
                    xml.append('<row r="%d">%s</row>' % (r, "".join(cells)))
                z.writestr("xl/worksheets/sheet%d.xml" % i,
                           "<worksheet %s><sheetData>%s</sheetData></worksheet>" % (ns, "".join(xml)))
                tags.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (title, i, i))
                rels.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>'
                            % (i, rel, i))
            z.writestr("xl/workbook.xml", '<workbook %s xmlns:r="%s"><sheets>%s</sheets></workbook>'
                       % (ns, rel, "".join(tags)))
            z.writestr("xl/_rels/workbook.xml.rels",
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s'
                       "</Relationships>" % "".join(rels))
        return path

    def test_summary_sheet_rules(self):
        path = self.workbook("s.xlsx", {"Summary": [
            ["Item", "Value"],
            ["Total revenue", ("=Sales!F9", 5000)],
            ["Total cost", ("=Costs!N9", 4000)],
            ["Profit", ("=B2-B3", 1000)],
            ["Margin (% of revenue)", ("=B4/B3", 0.25)],
            ["Units sold", ("=Sales!D9", 200)],
            ["Average revenue per unit", ("=B2/B6", 25)],
            ["Cost centres", 41],
        ]})
        res = au.process(path)
        self.assertEqual(kinds(res["leads"]), ["ratio-base", "typed-among-formulas"])
        self.assertEqual([l["where"] for l in res["leads"]], ["Summary!B5", "Summary!B8"])

    def test_text_number_is_named_as_the_cause(self):
        path = self.workbook("t.xlsx", {"Costs": [
            ["Centre", "Jan", "Feb", "Mar", "Q1"],
            ["A", 100, 110, 120, ("=SUM(B2:D2)", 330)],
            ["B", 200, "210", 220, ("=SUM(B3:D3)", 420)],
            ["C", 300, 310, 320, ("=SUM(B4:D4)", 930)],
            ["Total", ("=SUM(B2:B4)", 600), ("=SUM(C2:C4)", 420), ("=SUM(D2:D4)", 660), ("=SUM(E2:E4)", 1680)],
        ]})
        res = au.process(path)
        by_cell = dict((l["where"], l) for l in res["leads"] if l["kind"] != "number-as-text")
        self.assertIn("text cell is the cause", by_cell["Costs!E3"]["message"])
        self.assertIn("text cell is the cause", by_cell["Costs!C5"]["message"])

    def test_minimal_workbook(self):
        sheet = (
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
            '<row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Jan</t></is></c></row>'
            '<row r="2"><c r="A2" t="inlineStr"><is><t>Rent</t></is></c><c r="B2"><v>100</v></c></row>'
            '<row r="3"><c r="A3" t="inlineStr"><is><t>Power</t></is></c><c r="B3"><v>50</v></c></row>'
            '<row r="4"><c r="A4" t="inlineStr"><is><t>Water</t></is></c><c r="B4"><v>25</v></c></row>'
            '<row r="5"><c r="A5" t="inlineStr"><is><t>Total</t></is></c><c r="B5"><f>SUM(B2:B3)</f><v>150</v></c></row>'
            "</sheetData></worksheet>")
        workbook = (
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Costs" sheetId="1" r:id="rId1"/></sheets></workbook>')
        rels = (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/></Relationships>')
        path = os.path.join(self.dir, "w.xlsx")
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("xl/workbook.xml", workbook)
            z.writestr("xl/_rels/workbook.xml.rels", rels)
            z.writestr("xl/worksheets/sheet1.xml", sheet)
        res = au.process(path)
        self.assertIsNone(res["error"])
        self.assertEqual(kinds(res["leads"]), ["sum-range", "total-row"])
        self.assertEqual(res["leads"][0]["where"], "Costs!B5")


if __name__ == "__main__":
    unittest.main()
