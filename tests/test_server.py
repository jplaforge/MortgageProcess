"""Tests for MCP server tool registration."""

import base64
from unittest.mock import AsyncMock, patch

import pytest

from mortgage_mcp.server import mcp


@pytest.fixture
def mock_ctx():
    """A mock MCP Context with async logging and progress methods."""
    ctx = AsyncMock()
    ctx.info = AsyncMock()
    ctx.warning = AsyncMock()
    ctx.report_progress = AsyncMock()
    return ctx


class TestServerToolRegistration:
    def test_tools_registered(self):
        """Verify both tools are registered on the MCP server."""
        tools = mcp._tool_manager._tools
        tool_names = list(tools.keys())
        assert "analyze_bank_statements" in tool_names
        assert "health_check" in tool_names

    def test_analyze_tool_has_description(self):
        tools = mcp._tool_manager._tools
        tool = tools["analyze_bank_statements"]
        assert "relevés bancaires" in tool.description.lower()

    def test_health_tool_has_description(self):
        tools = mcp._tool_manager._tools
        tool = tools["health_check"]
        assert "vertex ai" in tool.description.lower() or "statut" in tool.description.lower()


class TestAnalyzeBankStatementsTool:
    @pytest.mark.asyncio
    async def test_invalid_document_returns_error(self, mock_ctx):
        """Calling with bad documents should return an error message, not raise."""
        from mortgage_mcp.tools.analyze_bank_statements import analyze_bank_statements

        result = await analyze_bank_statements(
            documents=[{"data": "bad-base64!!!", "mime_type": "application/pdf"}],
            ctx=mock_ctx,
        )
        assert len(result) == 1
        assert "Erreur" in result[0].text

    @pytest.mark.asyncio
    async def test_unsupported_mime_returns_error(self, mock_ctx):
        from mortgage_mcp.tools.analyze_bank_statements import analyze_bank_statements

        result = await analyze_bank_statements(
            documents=[
                {"data": base64.b64encode(b"data").decode(), "mime_type": "application/zip"}
            ],
            ctx=mock_ctx,
        )
        assert len(result) == 1
        assert "Erreur" in result[0].text

    @pytest.mark.asyncio
    async def test_successful_analysis_with_mock(self, sample_extraction, mock_ctx):
        """With mocked Vertex AI, should return summary + Excel."""
        from mortgage_mcp.tools.analyze_bank_statements import analyze_bank_statements

        pdf_b64 = base64.b64encode(b"%PDF-1.4 fake content").decode()

        with patch(
            "mortgage_mcp.tools.analyze_bank_statements.extract_bank_statements",
            new_callable=AsyncMock,
            return_value=sample_extraction,
        ):
            result = await analyze_bank_statements(
                documents=[{"data": pdf_b64, "mime_type": "application/pdf"}],
                ctx=mock_ctx,
                borrower_name="Jean Tremblay",
            )

        assert len(result) == 2
        # First item: text summary
        assert result[0].type == "text"
        assert "Jean Tremblay" in result[0].text
        assert "74,800.00" in result[0].text or "74800" in result[0].text
        # Second item: embedded Excel resource
        assert result[1].type == "resource"
        assert "jean_tremblay" in str(result[1].resource.uri)
