"""Tests for downpayment Excel report generation."""

import base64
import io

import pytest
from openpyxl import load_workbook

from mortgage_mcp.services.downpayment_excel import generate_dp_excel, generate_dp_excel_base64


class TestExcelGeneration:
    def test_generates_bytes(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        assert isinstance(data, bytes)
        assert len(data) > 0

    def test_generates_base64(self, dp_audit_result):
        b64 = generate_dp_excel_base64(dp_audit_result)
        # Should be valid base64
        raw = base64.b64decode(b64)
        assert len(raw) > 0

    def test_five_sheets_present(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        expected = {"Résumé", "Comptes", "Transactions clés", "Transferts", "Flags & Demandes"}
        assert set(wb.sheetnames) == expected

    def test_resume_sheet_content(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Résumé"]
        # Title
        assert "Audit de la mise de fonds" in (ws.cell(1, 1).value or "")
        # Borrower name should appear
        values = [ws.cell(r, 2).value for r in range(1, 20)]
        assert "Jean Tremblay" in values

    def test_comptes_sheet_content(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Comptes"]
        # Header row
        assert ws.cell(1, 1).value == "ID"
        assert ws.cell(1, 2).value == "Institution"
        # Data rows (2 accounts)
        assert ws.cell(2, 1).value == "A1"
        assert ws.cell(3, 1).value == "A2"

    def test_key_transactions_sheet(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Transactions clés"]
        assert ws.cell(1, 1).value == "ID"
        # Should have flagged transactions
        row_count = sum(1 for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value)
        assert row_count >= 1

    def test_transfers_sheet(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Transferts"]
        assert ws.cell(1, 1).value == "De (compte)"
        # One transfer match
        assert ws.cell(2, 1).value == "A1"
        assert ws.cell(2, 2).value == "A2"

    def test_flags_sheet(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Flags & Demandes"]
        # Should have "Drapeaux d'audit" title
        assert ws.cell(1, 1).value == "Drapeaux d'audit"
        # Should have flag data rows
        assert ws.cell(3, 1).value is not None  # First flag type

    def test_flags_sheet_has_requests(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Flags & Demandes"]
        # Find "Demandes au client" section
        found = False
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row, 1).value
            if val and "Demandes au client" in str(val):
                found = True
                break
        assert found

    def test_currency_formatting(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        ws = wb["Transferts"]
        # Amount column should have currency format
        if ws.cell(2, 3).value is not None:
            assert "$" in (ws.cell(2, 3).number_format or "")

    def test_empty_result(self):
        """Test with minimal audit result."""
        from mortgage_mcp.models.downpayment import DPAuditResult, DPSummary, SourceBreakdown
        result = DPAuditResult(
            summary=DPSummary(dp_target=50000, source_breakdown=SourceBreakdown()),
            borrower_name="Test",
        )
        data = generate_dp_excel(result)
        wb = load_workbook(io.BytesIO(data))
        assert len(wb.sheetnames) == 5
