"""Decode and validate base64-encoded documents."""

import base64

ALLOWED_MIME_TYPES = {
    "application/pdf",
    "image/jpeg",
    "image/png",
    "text/csv",
}

# Magic bytes for file type validation
MAGIC_BYTES = {
    "application/pdf": b"%PDF",
    "image/jpeg": b"\xff\xd8\xff",
    "image/png": b"\x89PNG",
}


class DocumentParseError(Exception):
    pass


class ParsedDocument:
    def __init__(self, data: bytes, mime_type: str):
        self.data = data
        self.mime_type = mime_type


def parse_document(base64_data: str, mime_type: str) -> ParsedDocument:
    """Decode a base64 document and validate its type.

    Args:
        base64_data: Base64-encoded file content.
        mime_type: Declared MIME type of the document.

    Returns:
        ParsedDocument with decoded bytes and validated MIME type.

    Raises:
        DocumentParseError: If MIME type is unsupported or content doesn't match.
    """
    if mime_type not in ALLOWED_MIME_TYPES:
        raise DocumentParseError(
            f"Type MIME non supporté: {mime_type}. "
            f"Types acceptés: {', '.join(sorted(ALLOWED_MIME_TYPES))}"
        )

    try:
        data = base64.b64decode(base64_data)
    except Exception as exc:
        raise DocumentParseError(f"Décodage base64 échoué: {exc}") from exc

    if len(data) == 0:
        raise DocumentParseError("Le document est vide")

    # Validate magic bytes for binary formats
    expected_magic = MAGIC_BYTES.get(mime_type)
    if expected_magic and not data[:len(expected_magic)] == expected_magic:
        raise DocumentParseError(
            f"Le contenu ne correspond pas au type MIME déclaré ({mime_type})"
        )

    return ParsedDocument(data=data, mime_type=mime_type)


def parse_documents(
    documents: list[dict],
) -> list[ParsedDocument]:
    """Parse a list of document dicts with 'data' and 'mime_type' keys."""
    parsed = []
    for i, doc in enumerate(documents):
        data = doc.get("data")
        mime = doc.get("mime_type")
        if not data or not mime:
            raise DocumentParseError(
                f"Document {i + 1}: 'data' et 'mime_type' sont requis"
            )
        parsed.append(parse_document(data, mime))
    return parsed
