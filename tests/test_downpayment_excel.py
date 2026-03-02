"""Tests for downpayment Excel report generation."""

import base64
import io

import pytest
from openpyxl import load_workbook

from mortgage_mcp.models.downpayment import (
    DPAuditResult,
    DPSummary,
    FlagSeverity,
    SourceBreakdown,
)
from mortgage_mcp.services.downpayment_excel import (
    FLAG_TYPE_LABELS,
    SEVERITY_LABELS,
    generate_dp_excel,
    generate_dp_excel_base64,
)


class TestExcelGeneration:
    def test_generates_bytes(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        assert isinstance(data, bytes)
        assert len(data) > 0

    def test_generates_base64(self, dp_audit_result):
        b64 = generate_dp_excel_base64(dp_audit_result)
        raw = base64.b64decode(b64)
        assert len(raw) > 0

    def test_three_sheets_present(self, dp_audit_result):
        data = generate_dp_excel(dp_audit_result)
        wb = load_workbook(io.BytesIO(data))
        expected = {"Résumé", "Demandes au client", "Détail"}
        assert set(wb.sheetnames) == expected


class TestResumeSheet:
    def test_title(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        assert "Audit de la mise de fonds" in (ws.cell(1, 1).value or "")

    def test_verdict_present(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        # Verdict is in row 3 (after title + blank row)
        verdict = ws.cell(3, 1).value
        assert verdict in ("CONFORME", "À VÉRIFIER", "RÉVISION REQUISE")

    def test_verdict_red_when_critical(self, dp_audit_result):
        """Result with critical flags should show RÉVISION REQUISE."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        assert ws.cell(3, 1).value == "RÉVISION REQUISE"

    def test_verdict_green_when_conforme(self):
        result = DPAuditResult(
            summary=DPSummary(
                dp_target=50000, dp_explained_amount=55000,
                source_breakdown=SourceBreakdown(payroll=55000),
            ),
            borrower_name="Test",
        )
        wb = load_workbook(io.BytesIO(generate_dp_excel(result)))
        ws = wb["Résumé"]
        assert ws.cell(3, 1).value == "CONFORME"

    def test_percentage_progress(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        progress = ws.cell(4, 1).value or ""
        assert "%" in progress
        assert "expliqué" in progress

    def test_borrower_name_present(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        values = [ws.cell(r, 2).value for r in range(1, 30)]
        assert "Jean Tremblay" in values

    def test_no_co_borrower_when_empty(self, dp_audit_result):
        """Co-emprunteur line should not appear when empty."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        labels = [ws.cell(r, 1).value for r in range(1, 30)]
        assert "Co-emprunteur:" not in labels

    def test_accounts_inline(self, dp_audit_result):
        """Accounts should appear inline on Résumé, not a separate sheet."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        values = [ws.cell(r, 1).value for r in range(1, 50)]
        assert "Comptes analysés" in values
        # Institution names should appear
        all_vals = []
        for r in range(1, 50):
            for c in range(1, 7):
                v = ws.cell(r, c).value
                if v:
                    all_vals.append(str(v))
        assert any("Desjardins" in v for v in all_vals)

    def test_zero_sources_hidden(self):
        """Source lines with zero value should not appear."""
        result = DPAuditResult(
            summary=DPSummary(
                dp_target=50000,
                dp_explained_amount=20000,
                unexplained_amount=30000,
                needs_review=True,
                source_breakdown=SourceBreakdown(payroll=20000, gift=0, investment_sale=0, unexplained=30000),
            ),
            borrower_name="Test",
        )
        wb = load_workbook(io.BytesIO(generate_dp_excel(result)))
        ws = wb["Résumé"]
        labels = [ws.cell(r, 1).value for r in range(1, 50)]
        assert "Dons:" not in labels
        assert "Vente de placements:" not in labels
        assert "Accumulation salariale:" in labels

    def test_transfers_inline(self, dp_audit_result):
        """Transfers should appear inline on Résumé."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        values = [ws.cell(r, 1).value for r in range(1, 60)]
        assert "Transferts inter-comptes détectés" in values

    def test_flags_french_labels(self, dp_audit_result):
        """Flag types and severities should use French labels."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Résumé"]
        all_vals = set()
        for r in range(1, 60):
            for c in range(1, 5):
                v = ws.cell(r, c).value
                if v:
                    all_vals.add(str(v))
        # Should have French labels, not English enum values
        assert any("Dépôt important" in v for v in all_vals)
        assert any("Critique" in v for v in all_vals)
        # Should NOT have raw English values
        assert not any(v == "large_deposit" for v in all_vals)
        assert not any(v == "critical" for v in all_vals)


class TestDemandesSheet:
    def test_title(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Demandes au client"]
        assert ws.cell(1, 1).value == "Demandes au client"

    def test_request_count(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Demandes au client"]
        subtitle = ws.cell(2, 1).value or ""
        assert "1 document(s)" in subtitle

    def test_request_content(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Demandes au client"]
        all_vals = []
        for r in range(1, 20):
            for c in range(1, 3):
                v = ws.cell(r, c).value
                if v:
                    all_vals.append(str(v))
        assert any("Lettre de don" in v for v in all_vals)

    def test_human_readable_tx_refs(self, dp_audit_result):
        """Transaction references should be human-readable, not raw IDs."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Demandes au client"]
        all_vals = []
        for r in range(1, 20):
            for c in range(1, 3):
                v = ws.cell(r, c).value
                if v:
                    all_vals.append(str(v))
        # Should contain amount + description, not just "A1-007"
        has_readable_ref = any("25,000" in v or "DON PARENTS" in v for v in all_vals)
        assert has_readable_ref

    def test_empty_requests(self):
        result = DPAuditResult(
            summary=DPSummary(dp_target=50000, source_breakdown=SourceBreakdown()),
            borrower_name="Test",
        )
        wb = load_workbook(io.BytesIO(generate_dp_excel(result)))
        ws = wb["Demandes au client"]
        assert "Aucune demande requise" in (ws.cell(2, 1).value or "")


class TestDetailSheet:
    def test_title(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Détail"]
        assert "Transactions nécessitant" in (ws.cell(1, 1).value or "")

    def test_headers_french(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Détail"]
        # Header row is row 3 (after title + blank)
        headers = [ws.cell(3, c).value for c in range(1, 7)]
        assert "Date" in headers
        assert "Compte" in headers
        assert "Catégorie" in headers
        assert "Drapeaux" in headers

    def test_institution_names_instead_of_ids(self, dp_audit_result):
        """Account column should show institution names, not raw IDs like 'A1'."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Détail"]
        all_vals = []
        for r in range(4, 20):
            v = ws.cell(r, 2).value  # Compte column
            if v:
                all_vals.append(str(v))
        # Should have institution names
        assert all(v != "A1" and v != "A2" for v in all_vals)

    def test_french_categories(self, dp_audit_result):
        """Category column should use French labels."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Détail"]
        for r in range(4, 20):
            v = ws.cell(r, 5).value  # Catégorie column
            if v:
                assert v in ("Salaire", "Don", "Espèces", "Revenu d'affaires",
                             "Transfert", "Gouvernement", "Placement", "Prêt",
                             "Remboursement", "Facture", "Achat", "Autre"), f"Unexpected category: {v}"

    def test_french_flag_labels(self, dp_audit_result):
        """Flag column should use French labels."""
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Détail"]
        for r in range(4, 20):
            v = ws.cell(r, 6).value  # Drapeaux column
            if v:
                # Should NOT contain English enum values
                assert "large_deposit" not in str(v)
                assert "cash_deposit" not in str(v)

    def test_no_info_only_transactions(self, dp_audit_result):
        """Transactions only flagged with INFO should not appear in Détail."""
        from mortgage_mcp.models.downpayment import (
            DPFlag,
            DPTransaction,
            FlagType,
            TransactionCategory,
            TransactionType,
        )
        result = DPAuditResult(
            transactions=[
                DPTransaction(id="A1-001", date="2025-01-15", description="SMALL RECURRING",
                              amount=500, type=TransactionType.DEPOSIT,
                              category=TransactionCategory.OTHER, account_id="A1"),
            ],
            flags=[
                DPFlag(type=FlagType.NON_PAYROLL_RECURRING, severity=FlagSeverity.INFO,
                       rationale="test", supporting_transaction_ids=["A1-001"]),
            ],
            summary=DPSummary(dp_target=50000, source_breakdown=SourceBreakdown()),
            borrower_name="Test",
        )
        wb = load_workbook(io.BytesIO(generate_dp_excel(result)))
        ws = wb["Détail"]
        # Should show "Aucune transaction" message since only INFO flags
        all_vals = [ws.cell(r, 1).value for r in range(1, 10)]
        assert any("Aucune transaction" in str(v) for v in all_vals if v)

    def test_currency_formatting(self, dp_audit_result):
        wb = load_workbook(io.BytesIO(generate_dp_excel(dp_audit_result)))
        ws = wb["Détail"]
        # Find a row with amount data
        for r in range(4, 20):
            if ws.cell(r, 4).value is not None:
                assert "$" in (ws.cell(r, 4).number_format or "")
                break

    def test_empty_result(self):
        """Minimal result should still generate valid Excel."""
        result = DPAuditResult(
            summary=DPSummary(dp_target=50000, source_breakdown=SourceBreakdown()),
            borrower_name="Test",
        )
        data = generate_dp_excel(result)
        wb = load_workbook(io.BytesIO(data))
        assert len(wb.sheetnames) == 3


class TestDateFormatting:
    def test_french_date_format(self):
        from mortgage_mcp.services.downpayment_excel import _format_date_short
        assert _format_date_short("2025-02-20") == "20 févr. 2025"
        assert _format_date_short("2025-01-01") == "1 janv. 2025"
        assert _format_date_short("2025-12-25") == "25 déc. 2025"

    def test_invalid_date_passthrough(self):
        from mortgage_mcp.services.downpayment_excel import _format_date_short
        assert _format_date_short("invalid") == "invalid"
        assert _format_date_short("") == ""
