"""Downpayment audit tool — orchestrates parsing, AI extraction, analysis, and Excel generation."""

from mcp.server.fastmcp import Context
from mcp.types import BlobResourceContents, EmbeddedResource, TextContent

from mortgage_mcp.services.document_parser import DocumentParseError, parse_documents
from mortgage_mcp.services.downpayment_analyzer import analyze
from mortgage_mcp.services.downpayment_excel import generate_dp_excel_base64
from mortgage_mcp.services.downpayment_vertex import extract_dp_transactions


def _format_dp_summary(result) -> str:
    """Build a human-readable French summary of the audit result."""
    s = result.summary
    sb = s.source_breakdown

    lines = [
        "# Audit de la mise de fonds — Provenance des fonds",
        "",
        f"**Emprunteur:** {result.borrower_name}",
    ]
    if result.co_borrower_name:
        lines.append(f"**Co-emprunteur:** {result.co_borrower_name}")
    lines += [
        f"**Date de clôture:** {result.closing_date}",
        f"**Mise de fonds cible:** {s.dp_target:,.2f} $",
        "",
        "## Ventilation des sources",
        "",
        "| Source | Montant |",
        "|--------|---------|",
        f"| Épargne salariale | {sb.payroll:,.2f} $ |",
        f"| Dons | {sb.gift:,.2f} $ |",
        f"| Vente de placements | {sb.investment_sale:,.2f} $ |",
        f"| Vente de propriété | {sb.property_sale:,.2f} $ |",
        f"| Autres sources expliquées | {sb.other_explained:,.2f} $ |",
        f"| **Sources non expliquées** | **{sb.unexplained:,.2f} $** |",
        "",
        f"**Montant expliqué:** {s.dp_explained_amount:,.2f} $",
        f"**Montant non expliqué:** {s.unexplained_amount:,.2f} $",
        f"**Nécessite révision:** {'OUI' if s.needs_review else 'NON'}",
        "",
    ]

    # Flags by severity
    if result.flags:
        lines.append("## Drapeaux d'audit")
        lines.append("")
        for severity_label, severity_val in [("Critiques", "critical"), ("Avertissements", "warning"), ("Informations", "info")]:
            sev_flags = [f for f in result.flags if f.severity.value == severity_val]
            if sev_flags:
                lines.append(f"### {severity_label} ({len(sev_flags)})")
                lines.append("")
                for f in sev_flags:
                    lines.append(f"- {f.rationale}")
                lines.append("")

    # Transfers
    if result.transfers:
        lines.append(f"## Transferts inter-comptes ({len(result.transfers)})")
        lines.append("")
        lines.append("| De | Vers | Montant | Score |")
        lines.append("|----|------|---------|-------|")
        for tm in result.transfers:
            lines.append(f"| {tm.from_account_id} | {tm.to_account_id} | {tm.amount:,.2f} $ | {tm.match_score:.2f} |")
        lines.append("")

    # Client requests
    if result.client_requests:
        lines.append("## Demandes au client")
        lines.append("")
        for i, req in enumerate(result.client_requests, 1):
            lines.append(f"### {i}. {req.title}")
            lines.append(f"{req.reason}")
            lines.append("")
            for doc in req.required_docs:
                lines.append(f"- {doc}")
            lines.append("")

    lines.append("---")
    lines.append("*Le fichier Excel détaillé et le JSON structuré sont joints ci-dessous.*")

    return "\n".join(lines)


def _format_dp_json(result) -> str:
    """Serialize the full result to JSON."""
    return result.model_dump_json(indent=2)


async def audit_downpayment(
    documents: list[dict],
    target_downpayment_amount: float,
    closing_date: str,
    borrower_name: str,
    ctx: Context,
    supporting_documents: list[dict] | None = None,
    co_borrower_name: str | None = None,
    deal_notes: str | None = None,
) -> list[TextContent | EmbeddedResource]:
    """Audit downpayment source of funds and return summary + JSON + Excel.

    Args:
        documents: Bank statements as base64 dicts with 'data' and 'mime_type'.
        target_downpayment_amount: Target downpayment in CAD.
        closing_date: Expected closing date (YYYY-MM-DD).
        borrower_name: Borrower name.
        ctx: MCP Context for progress reporting and logging.
        supporting_documents: Optional supporting docs (gift letters, etc.).
        co_borrower_name: Optional co-borrower name.
        deal_notes: Optional deal notes.

    Returns:
        List containing TextContent summary, TextContent JSON, and EmbeddedResource Excel.
    """
    total_steps = 5

    # 1. Parse documents
    await ctx.report_progress(progress=0, total=total_steps)
    await ctx.info(f"Réception de {len(documents)} relevé(s) bancaire(s)")
    try:
        parsed = parse_documents(documents)
        if supporting_documents:
            await ctx.info(f"+ {len(supporting_documents)} document(s) justificatif(s)")
            parsed += parse_documents(supporting_documents)
    except DocumentParseError as exc:
        return [TextContent(type="text", text=f"Erreur de document: {exc}")]

    # 2. Gemini extraction
    await ctx.report_progress(progress=1, total=total_steps)
    await ctx.info("Extraction IA en cours via Vertex AI (Gemini) — extraction de toutes les transactions")
    try:
        extraction = await extract_dp_transactions(
            parsed, borrower_name, co_borrower_name, closing_date, deal_notes
        )
    except Exception as exc:
        return [
            TextContent(
                type="text",
                text=f"Erreur lors de l'extraction IA: {exc}\n\nVeuillez réessayer ou vérifier la qualité des documents.",
            )
        ]

    # 3. Python post-processing
    await ctx.report_progress(progress=2, total=total_steps)
    await ctx.info(
        f"Post-traitement: {len(extraction.transactions)} transactions, "
        f"{len(extraction.accounts)} compte(s)"
    )
    try:
        result = analyze(
            extraction,
            target_downpayment=target_downpayment_amount,
            closing_date=closing_date,
            borrower_name=borrower_name,
            co_borrower_name=co_borrower_name,
            deal_notes=deal_notes,
        )
    except Exception as exc:
        return [
            TextContent(
                type="text",
                text=f"Erreur lors du post-traitement: {exc}",
            )
        ]

    # 4. Generate Excel
    await ctx.report_progress(progress=3, total=total_steps)
    await ctx.info("Génération du rapport Excel d'audit")
    try:
        excel_b64 = generate_dp_excel_base64(result)
    except Exception as exc:
        summary = _format_dp_summary(result)
        return [
            TextContent(
                type="text",
                text=f"L'analyse a réussi mais la génération Excel a échoué: {exc}\n\n{summary}",
            )
        ]

    # 5. Build response
    await ctx.report_progress(progress=4, total=total_steps)
    summary = _format_dp_summary(result)
    json_output = _format_dp_json(result)

    borrower_slug = borrower_name.replace(" ", "_").lower()
    filename = f"audit_mise_de_fonds_{borrower_slug}.xlsx"

    await ctx.report_progress(progress=5, total=total_steps)
    flag_count = len(result.flags)
    transfer_count = len(result.transfers)
    await ctx.info(
        f"Audit terminé — {flag_count} drapeau(x), {transfer_count} transfert(s) détecté(s), "
        f"mise de fonds: {result.summary.dp_explained_amount:,.2f} $ expliqué sur {result.summary.dp_target:,.2f} $"
    )

    return [
        TextContent(type="text", text=summary),
        TextContent(type="text", text=json_output),
        EmbeddedResource(
            type="resource",
            resource=BlobResourceContents(
                uri=f"data:///{filename}",
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                blob=excel_b64,
            ),
        ),
    ]
