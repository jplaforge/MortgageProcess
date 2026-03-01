"""Bank statement analysis tool — orchestrates parsing, AI extraction, and Excel generation."""

from mcp.types import EmbeddedResource, TextContent, BlobResourceContents

from mortgage_mcp.services.document_parser import DocumentParseError, parse_documents
from mortgage_mcp.services.excel_generator import generate_excel_base64
from mortgage_mcp.services.vertex_ai import extract_bank_statements


def _format_summary(extraction) -> str:
    """Build a human-readable French summary of the extraction."""
    info = extraction.account_info
    lines = [
        "# Analyse des relevés bancaires — Travailleur autonome",
        "",
        f"**Titulaire:** {info.account_holder}",
        f"**Institution:** {info.institution}",
        f"**Période:** {info.statement_period_start} au {info.statement_period_end}",
        f"**Mois couverts:** {extraction.months_covered}",
        "",
        "## Sommaire des revenus",
        "",
        f"| Métrique | Montant |",
        f"|----------|---------|",
        f"| Dépôts totaux | {extraction.total_deposits:,.2f} $ |",
        f"| Revenu d'affaires total | {extraction.total_business_income:,.2f} $ |",
        f"| Retraits totaux | {extraction.total_withdrawals:,.2f} $ |",
        f"| Revenu mensuel moyen (affaires) | {extraction.average_monthly_business_income:,.2f} $ |",
        f"| **Revenu annualisé (affaires)** | **{extraction.annualized_business_income:,.2f} $** |",
        "",
    ]

    if extraction.monthly_breakdown:
        lines.append("## Ventilation mensuelle")
        lines.append("")
        lines.append("| Mois | Dépôts affaires | Retraits | Nb dépôts |")
        lines.append("|------|-----------------|----------|-----------|")
        for m in extraction.monthly_breakdown:
            lines.append(
                f"| {m.month} | {m.business_deposits:,.2f} $ | "
                f"{m.total_withdrawals:,.2f} $ | {m.deposit_count} |"
            )
        lines.append("")

    if extraction.confidence_notes:
        lines.append("## Notes pour le courtier")
        lines.append("")
        for note in extraction.confidence_notes:
            lines.append(f"- {note}")
        lines.append("")

    lines.append("---")
    lines.append("*Le fichier Excel détaillé est joint ci-dessous.*")

    return "\n".join(lines)


async def analyze_bank_statements(
    documents: list[dict],
    borrower_name: str | None = None,
    business_name: str | None = None,
    business_type: str | None = None,
) -> list[TextContent | EmbeddedResource]:
    """Analyze bank statements and return summary + Excel file.

    Args:
        documents: List of dicts with 'data' (base64) and 'mime_type' keys.
        borrower_name: Optional borrower name for context.
        business_name: Optional business name for context.
        business_type: Optional business type for context.

    Returns:
        List containing a TextContent summary and an EmbeddedResource Excel file.
    """
    # 1. Parse and validate documents
    try:
        parsed = parse_documents(documents)
    except DocumentParseError as exc:
        return [TextContent(type="text", text=f"Erreur de document: {exc}")]

    # 2. Extract via Vertex AI
    try:
        extraction = await extract_bank_statements(
            parsed, borrower_name, business_name, business_type
        )
    except Exception as exc:
        return [
            TextContent(
                type="text",
                text=f"Erreur lors de l'analyse IA: {exc}\n\nVeuillez réessayer ou vérifier la qualité des documents.",
            )
        ]

    # 3. Generate Excel
    try:
        excel_b64 = generate_excel_base64(extraction)
    except Exception as exc:
        return [
            TextContent(
                type="text",
                text=f"L'analyse a réussi mais la génération Excel a échoué: {exc}\n\n{_format_summary(extraction)}",
            )
        ]

    # 4. Build response
    summary = _format_summary(extraction)

    borrower_slug = (borrower_name or "emprunteur").replace(" ", "_").lower()
    filename = f"analyse_revenu_{borrower_slug}.xlsx"

    return [
        TextContent(type="text", text=summary),
        EmbeddedResource(
            type="resource",
            resource=BlobResourceContents(
                uri=f"data:///{filename}",
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                blob=excel_b64,
            ),
        ),
    ]
