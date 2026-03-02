"""Tests for downpayment audit tool orchestration (pipeline mocked)."""

import base64
import json
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mortgage_mcp.models.downpayment import (
    DPAccountInfo,
    DPAuditResult,
    DPExtraction,
    DPSummary,
    DPTransaction,
    SourceBreakdown,
    TransactionCategory,
    TransactionType,
)
from mortgage_mcp.tools.downpayment_audit import audit_downpayment


@pytest.fixture
def mock_ctx():
    ctx = MagicMock()
    ctx.report_progress = AsyncMock()
    ctx.info = AsyncMock()
    return ctx


@pytest.fixture
def valid_doc():
    """A valid PDF document dict."""
    pdf_bytes = b"%PDF-1.4 fake content"
    return {
        "data": base64.b64encode(pdf_bytes).decode(),
        "mime_type": "application/pdf",
    }


@pytest.fixture
def mock_extraction():
    return DPExtraction(
        accounts=[DPAccountInfo(account_id="A1", institution="Desjardins")],
        transactions=[
            DPTransaction(
                id="A1-001", date="2025-01-15", description="PAIE", amount=3500.00,
                type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1",
            ),
        ],
    )


class TestAuditDownpayment:
    @pytest.mark.asyncio
    async def test_document_error(self, mock_ctx):
        """Invalid document should return error TextContent."""
        bad_doc = {"data": "not-base64!!!", "mime_type": "application/pdf"}
        result = await audit_downpayment(
            [bad_doc], 80000, "2025-06-15", "Jean", mock_ctx,
        )
        assert len(result) == 1
        assert "Erreur" in result[0].text

    @pytest.mark.asyncio
    async def test_gemini_error(self, mock_ctx, valid_doc):
        """Gemini failure should return error TextContent."""
        with patch("mortgage_mcp.tools.downpayment_audit.extract_dp_transactions",
                    new_callable=AsyncMock, side_effect=Exception("Gemini down")):
            result = await audit_downpayment(
                [valid_doc], 80000, "2025-06-15", "Jean", mock_ctx,
            )
        assert len(result) == 1
        assert "Gemini down" in result[0].text

    @pytest.mark.asyncio
    async def test_full_pipeline(self, mock_ctx, valid_doc, mock_extraction):
        """Full pipeline should return 3 items: summary, JSON, Excel."""
        with patch("mortgage_mcp.tools.downpayment_audit.extract_dp_transactions",
                    new_callable=AsyncMock, return_value=mock_extraction):
            result = await audit_downpayment(
                [valid_doc], 80000, "2025-06-15", "Jean Tremblay", mock_ctx,
            )
        assert len(result) == 3
        # First: markdown summary
        assert "Audit de la mise de fonds" in result[0].text
        # Second: JSON
        json_data = json.loads(result[1].text)
        assert "summary" in json_data
        # Third: Excel resource
        assert result[2].type == "resource"
        assert result[2].resource.mimeType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    @pytest.mark.asyncio
    async def test_summary_contains_borrower(self, mock_ctx, valid_doc, mock_extraction):
        """Summary should contain borrower name."""
        with patch("mortgage_mcp.tools.downpayment_audit.extract_dp_transactions",
                    new_callable=AsyncMock, return_value=mock_extraction):
            result = await audit_downpayment(
                [valid_doc], 80000, "2025-06-15", "Jean Tremblay", mock_ctx,
            )
        assert "Jean Tremblay" in result[0].text

    @pytest.mark.asyncio
    async def test_json_is_valid(self, mock_ctx, valid_doc, mock_extraction):
        """JSON output should be valid and contain expected fields."""
        with patch("mortgage_mcp.tools.downpayment_audit.extract_dp_transactions",
                    new_callable=AsyncMock, return_value=mock_extraction):
            result = await audit_downpayment(
                [valid_doc], 80000, "2025-06-15", "Jean", mock_ctx,
            )
        data = json.loads(result[1].text)
        assert "accounts" in data
        assert "transactions" in data
        assert "transfers" in data
        assert "flags" in data
        assert "summary" in data

    @pytest.mark.asyncio
    async def test_excel_filename(self, mock_ctx, valid_doc, mock_extraction):
        """Excel filename should include borrower slug."""
        with patch("mortgage_mcp.tools.downpayment_audit.extract_dp_transactions",
                    new_callable=AsyncMock, return_value=mock_extraction):
            result = await audit_downpayment(
                [valid_doc], 80000, "2025-06-15", "Jean Tremblay", mock_ctx,
            )
        uri = str(result[2].resource.uri)
        assert "jean_tremblay" in uri

    @pytest.mark.asyncio
    async def test_progress_reporting(self, mock_ctx, valid_doc, mock_extraction):
        """Should report progress multiple times."""
        with patch("mortgage_mcp.tools.downpayment_audit.extract_dp_transactions",
                    new_callable=AsyncMock, return_value=mock_extraction):
            await audit_downpayment(
                [valid_doc], 80000, "2025-06-15", "Jean", mock_ctx,
            )
        assert mock_ctx.report_progress.call_count >= 4
