"""Tests for document parsing and validation."""

import base64

import pytest

from mortgage_mcp.services.document_parser import (
    DocumentParseError,
    parse_document,
    parse_documents,
)


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode()


class TestParseDocument:
    def test_valid_pdf(self):
        pdf_bytes = b"%PDF-1.4 fake pdf content"
        result = parse_document(_b64(pdf_bytes), "application/pdf")
        assert result.data == pdf_bytes
        assert result.mime_type == "application/pdf"

    def test_valid_jpeg(self):
        jpeg_bytes = b"\xff\xd8\xff\xe0fake jpeg"
        result = parse_document(_b64(jpeg_bytes), "image/jpeg")
        assert result.data == jpeg_bytes
        assert result.mime_type == "image/jpeg"

    def test_valid_png(self):
        png_bytes = b"\x89PNG\r\n\x1a\nfake png"
        result = parse_document(_b64(png_bytes), "image/png")
        assert result.data == png_bytes

    def test_valid_csv(self):
        csv_bytes = b"date,description,amount\n2025-01-01,Test,100.00"
        result = parse_document(_b64(csv_bytes), "text/csv")
        assert result.data == csv_bytes

    def test_unsupported_mime_type(self):
        with pytest.raises(DocumentParseError, match="Type MIME non supporté"):
            parse_document(_b64(b"data"), "application/json")

    def test_invalid_base64(self):
        with pytest.raises(DocumentParseError, match="Décodage base64"):
            parse_document("not-valid-base64!!!", "application/pdf")

    def test_empty_document(self):
        with pytest.raises(DocumentParseError, match="vide"):
            parse_document(_b64(b""), "application/pdf")

    def test_magic_bytes_mismatch(self):
        with pytest.raises(DocumentParseError, match="ne correspond pas"):
            parse_document(_b64(b"this is not a pdf"), "application/pdf")

    def test_csv_no_magic_check(self):
        """CSV has no magic bytes, so any content should pass."""
        result = parse_document(_b64(b"anything"), "text/csv")
        assert result.data == b"anything"


class TestParseDocuments:
    def test_valid_list(self):
        docs = [
            {"data": _b64(b"%PDF-1.4 content"), "mime_type": "application/pdf"},
            {"data": _b64(b"\xff\xd8\xff\xe0img"), "mime_type": "image/jpeg"},
        ]
        result = parse_documents(docs)
        assert len(result) == 2

    def test_missing_data(self):
        with pytest.raises(DocumentParseError, match="Document 1"):
            parse_documents([{"mime_type": "application/pdf"}])

    def test_missing_mime_type(self):
        with pytest.raises(DocumentParseError, match="Document 1"):
            parse_documents([{"data": _b64(b"test")}])
