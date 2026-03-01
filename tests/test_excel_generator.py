"""Tests for Excel report generation."""

import io

from openpyxl import load_workbook

from mortgage_mcp.services.excel_generator import generate_excel, generate_excel_base64


class TestGenerateExcel:
    def test_returns_bytes(self, sample_extraction):
        result = generate_excel(sample_extraction)
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_valid_xlsx(self, sample_extraction):
        data = generate_excel(sample_extraction)
        wb = load_workbook(io.BytesIO(data))
        assert "Resume" in wb.sheetnames
        assert "Detail mensuel" in wb.sheetnames
        assert "Depots" in wb.sheetnames

    def test_resume_sheet_content(self, sample_extraction):
        data = generate_excel(sample_extraction)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Resume"]

        assert ws["A1"].value == "Grille d'analyse — Revenu de travailleur autonome"
        assert ws["B3"].value == "Jean Tremblay"
        assert ws["B4"].value == "Desjardins"
        assert ws["B7"].value == 3
        assert ws["B11"].value == 18700.00
        assert ws["B14"].value == 74800.00

    def test_monthly_sheet_rows(self, sample_extraction):
        data = generate_excel(sample_extraction)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Detail mensuel"]

        # Header row
        assert ws.cell(row=1, column=1).value == "Mois"
        # Data rows: 3 months
        assert ws.cell(row=2, column=1).value == "2025-01"
        assert ws.cell(row=3, column=1).value == "2025-02"
        assert ws.cell(row=4, column=1).value == "2025-03"
        # Business deposits
        assert ws.cell(row=2, column=3).value == 6000.00
        assert ws.cell(row=3, column=3).value == 5500.00

    def test_monthly_sheet_formulas(self, sample_extraction):
        data = generate_excel(sample_extraction)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Detail mensuel"]

        # TOTAL row should be at row 5 (3 months + 1 header + 1)
        assert ws.cell(row=5, column=1).value == "TOTAL"
        # Should contain SUM formula
        assert "SUM" in str(ws.cell(row=5, column=2).value)
        # MOYENNE row at 6
        assert ws.cell(row=6, column=1).value == "MOYENNE"
        assert "AVERAGE" in str(ws.cell(row=6, column=2).value)

    def test_deposits_sheet_rows(self, sample_extraction):
        data = generate_excel(sample_extraction)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Depots"]

        assert ws.cell(row=1, column=1).value == "Date"
        # Total deposits across all months: 4 + 3 + 3 = 10
        total_deposits = sum(
            len(m.deposits) for m in sample_extraction.monthly_breakdown
        )
        # Check last deposit row has data
        assert ws.cell(row=total_deposits + 1, column=1).value is not None

    def test_base64_output(self, sample_extraction):
        result = generate_excel_base64(sample_extraction)
        assert isinstance(result, str)
        # Should be valid base64
        import base64
        decoded = base64.b64decode(result)
        assert decoded[:2] == b"PK"  # ZIP/XLSX magic bytes
