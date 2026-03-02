"""Bank statement analysis tool — orchestrates parsing, AI extraction, and Excel generation."""

from mcp.server.fastmcp import Context
from mcp.types import EmbeddedResource, TextContent, BlobResourceContents

from mortgage_mcp.services.document_parser import DocumentParseError, parse_documents
from mortgage_mcp.services.excel_generator import generate_excel_base64
from mortgage_mcp.services.vertex_ai import extract_bank_statements


def _format_summary(extraction) -> str:
    """Build a human-readable French summary of the extraction."""
    info = extraction.account_info
    # Calculate total personal transfers across all months
    total_personal_transfers = sum(
        m.personal_transfers for m in extraction.monthly_breakdown
    )

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
        f"| Transferts inter-comptes exclus | {total_personal_transfers:,.2f} $ |",
        f"| Retraits totaux | {extraction.total_withdrawals:,.2f} $ |",
        f"| Revenu mensuel moyen (affaires) | {extraction.average_monthly_business_income:,.2f} $ |",
        f"| **Revenu annualisé (affaires)** | **{extraction.annualized_business_income:,.2f} $** |",
        "",
    ]

    if extraction.monthly_breakdown:
        lines.append("## Ventilation mensuelle")
        lines.append("")
        lines.append("| Mois | Dépôts affaires | Transferts | Retraits | Nb dépôts |")
        lines.append("|------|-----------------|------------|----------|-----------|")
        for m in extraction.monthly_breakdown:
            lines.append(
                f"| {m.month} | {m.business_deposits:,.2f} $ | "
                f"{m.personal_transfers:,.2f} $ | "
                f"{m.total_withdrawals:,.2f} $ | {m.deposit_count} |"
            )
        lines.append("")

    if total_personal_transfers > 0:
        lines.append("## Transferts inter-comptes")
        lines.append("")
        lines.append(
            f"**{total_personal_transfers:,.2f} $** en transferts entre comptes de l'emprunteur "
            "ont été détectés et **exclus** du revenu d'affaires."
        )
        lines.append("")
        # List transfer-related confidence notes
        transfer_notes = [
            n for n in (extraction.confidence_notes or [])
            if "transfert inter-comptes" in n.lower()
        ]
        if transfer_notes:
            lines.append("Détails:")
            for note in transfer_notes:
                lines.append(f"- {note}")
            lines.append("")

    if extraction.nsf_events:
        lines.append("## Indicateurs de risque")
        lines.append("")
        lines.append(f"**Événements NSF/découverts:** {len(extraction.nsf_events)}")
        lines.append(f"**Frais NSF totaux:** {extraction.nsf_total_fees:,.2f} $")
        lines.append("")
        lines.append("| Date | Description | Montant |")
        lines.append("|------|-------------|---------|")
        for nsf in extraction.nsf_events:
            lines.append(f"| {nsf.date} | {nsf.description} | {nsf.amount:,.2f} $ |")
        lines.append("")

    if extraction.recurring_obligations:
        lines.append("## Obligations récurrentes")
        lines.append("")
        lines.append("| Bénéficiaire | Montant mensuel | Type |")
        lines.append("|--------------|-----------------|------|")
        for ob in extraction.recurring_obligations:
            lines.append(f"| {ob.payee} | {ob.monthly_amount:,.2f} $ | {ob.category} |")
        lines.append("")
        lines.append(f"**Total obligations mensuelles:** {extraction.total_monthly_obligations:,.2f} $")
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
    ctx: Context,
    borrower_name: str | None = None,
    business_name: str | None = None,
    business_type: str | None = None,
) -> list[TextContent | EmbeddedResource]:
    """Analyze bank statements and return summary + Excel file.

    Args:
        documents: List of dicts with 'data' (base64) and 'mime_type' keys.
        ctx: MCP Context for progress reporting and logging.
        borrower_name: Optional borrower name for context.
        business_name: Optional business name for context.
        business_type: Optional business type for context.

    Returns:
        List containing a TextContent summary and an EmbeddedResource Excel file.
    """
    # 1. Parse and validate documents
    await ctx.report_progress(progress=0, total=4)
    await ctx.info(f"Réception de {len(documents)} document(s) à analyser")
    try:
        parsed = parse_documents(documents)
    except DocumentParseError as exc:
        return [TextContent(type="text", text=f"Erreur de document: {exc}")]

    # 2. Extract via Vertex AI
    await ctx.report_progress(progress=1, total=4)
    await ctx.info("Extraction IA en cours via Vertex AI (Gemini)")
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
    await ctx.report_progress(progress=2, total=4)
    await ctx.info("Génération du rapport Excel")
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
    await ctx.report_progress(progress=3, total=4)
    summary = _format_summary(extraction)

    borrower_slug = (borrower_name or "emprunteur").replace(" ", "_").lower()
    filename = f"analyse_revenu_{borrower_slug}.xlsx"

    await ctx.report_progress(progress=4, total=4)
    await ctx.info(f"Analyse terminée — {extraction.months_covered} mois couverts, revenu annualisé: {extraction.annualized_business_income:,.2f} $")

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
