"""Tests for downpayment Vertex AI extraction (Gemini mocked)."""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mortgage_mcp.models.downpayment import DPAccountInfo, DPExtraction, DPTransaction, TransactionCategory, TransactionType
from mortgage_mcp.services.document_parser import ParsedDocument
from mortgage_mcp.services.downpayment_vertex import (
    DP_EXTRACTION_PROMPT,
    _build_dp_contents,
    extract_dp_transactions,
)


@pytest.fixture
def sample_docs():
    return [
        ParsedDocument(data=b"%PDF-fake", mime_type="application/pdf"),
    ]


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


class TestBuildContents:
    def test_basic_content(self, sample_docs):
        contents = _build_dp_contents(sample_docs)
        assert len(contents) == 1
        parts = contents[0].parts
        # Should have prompt + document
        assert len(parts) >= 2

    def test_context_included(self, sample_docs):
        contents = _build_dp_contents(
            sample_docs, borrower_name="Jean", co_borrower_name="Marie",
            closing_date="2025-06-15", deal_notes="Achat condo",
        )
        parts = contents[0].parts
        # First part should be context
        context_text = parts[0].text
        assert "Jean" in context_text
        assert "Marie" in context_text
        assert "2025-06-15" in context_text
        assert "Achat condo" in context_text

    def test_no_context_when_empty(self, sample_docs):
        contents = _build_dp_contents(sample_docs)
        parts = contents[0].parts
        # First part should be the prompt directly
        assert DP_EXTRACTION_PROMPT in parts[0].text

    def test_multiple_documents(self):
        docs = [
            ParsedDocument(data=b"%PDF-fake1", mime_type="application/pdf"),
            ParsedDocument(data=b"\x89PNG-fake", mime_type="image/png"),
        ]
        contents = _build_dp_contents(docs)
        parts = contents[0].parts
        # prompt + 2 docs = 3 parts
        assert len(parts) == 3


class TestExtractDPTransactions:
    @pytest.mark.asyncio
    async def test_extract_returns_dp_extraction(self, sample_docs, mock_extraction):
        mock_response = MagicMock()
        mock_response.text = mock_extraction.model_dump_json()

        mock_generate = AsyncMock(return_value=mock_response)
        mock_aio = MagicMock()
        mock_aio.models.generate_content = mock_generate

        mock_client = MagicMock()
        mock_client.aio = mock_aio

        with patch("mortgage_mcp.services.downpayment_vertex.genai.Client", return_value=mock_client), \
             patch("mortgage_mcp.services.downpayment_vertex.settings") as mock_settings:
            mock_settings.gemini_model = "gemini-2.5-flash"
            mock_settings.google_cloud_project = "test-project"
            mock_settings.google_cloud_location = "northamerica-northeast1"
            mock_settings.setup_gcp_credentials = MagicMock()

            result = await extract_dp_transactions(sample_docs, borrower_name="Jean")

        assert isinstance(result, DPExtraction)
        assert len(result.accounts) == 1
        assert result.transactions[0].id == "A1-001"
