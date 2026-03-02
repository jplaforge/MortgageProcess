"""Generate the Excel downpayment audit report from analysis results."""

import base64
import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from mortgage_mcp.models.downpayment import (
    DPAuditResult,
    FlagSeverity,
)

HEADER_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT_WHITE = Font(bold=True, size=11, color="FFFFFF")
CURRENCY_FORMAT = '#,##0.00 $'
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

SEVERITY_FILLS = {
    FlagSeverity.CRITICAL: PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    FlagSeverity.WARNING: PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    FlagSeverity.INFO: PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid"),
}

SEVERITY_FONTS = {
    FlagSeverity.CRITICAL: Font(color="9C0006"),
    FlagSeverity.WARNING: Font(color="9C6500"),
    FlagSeverity.INFO: Font(color="1F4E79"),
}


def _apply_header_style(ws, row: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def _set_col_widths(ws, widths: dict[str, int]) -> None:
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


# ── Sheet fillers ─────────────────────────────────────────────────────────


def _fill_resume(ws, result: DPAuditResult) -> None:
    """Fill the Résumé (summary) sheet."""
    row = 1
    ws.cell(row=row, column=1, value="Audit de la mise de fonds — Provenance des fonds").font = Font(bold=True, size=14)
    row += 2

    # Deal info
    deal_data = [
        ("Emprunteur:", result.borrower_name),
        ("Co-emprunteur:", result.co_borrower_name or "N/A"),
        ("Date de clôture:", result.closing_date),
        ("Mise de fonds cible:", result.summary.dp_target),
    ]
    for label, value in deal_data:
        ws.cell(row=row, column=1, value=label).font = HEADER_FONT
        cell = ws.cell(row=row, column=2, value=value)
        if isinstance(value, float):
            cell.number_format = CURRENCY_FORMAT
        row += 1
    if result.deal_notes:
        ws.cell(row=row, column=1, value="Notes:").font = HEADER_FONT
        ws.cell(row=row, column=2, value=result.deal_notes)
        row += 1
    row += 1

    # Source breakdown
    ws.cell(row=row, column=1, value="Ventilation des sources").font = Font(bold=True, size=12)
    row += 1
    sb = result.summary.source_breakdown
    sources = [
        ("Épargne salariale:", sb.payroll),
        ("Dons:", sb.gift),
        ("Vente de placements:", sb.investment_sale),
        ("Vente de propriété:", sb.property_sale),
        ("Autres sources expliquées:", sb.other_explained),
        ("Sources non expliquées:", sb.unexplained),
    ]
    for label, value in sources:
        ws.cell(row=row, column=1, value=label).font = HEADER_FONT
        cell = ws.cell(row=row, column=2, value=value)
        cell.number_format = CURRENCY_FORMAT
        if label.startswith("Sources non") and value > 0:
            cell.font = Font(color="9C0006", bold=True)
        row += 1
    row += 1

    # Summary
    ws.cell(row=row, column=1, value="Évaluation globale").font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=1, value="Montant expliqué:").font = HEADER_FONT
    ws.cell(row=row, column=2, value=result.summary.dp_explained_amount).number_format = CURRENCY_FORMAT
    row += 1
    ws.cell(row=row, column=1, value="Montant non expliqué:").font = HEADER_FONT
    cell = ws.cell(row=row, column=2, value=result.summary.unexplained_amount)
    cell.number_format = CURRENCY_FORMAT
    row += 1
    ws.cell(row=row, column=1, value="Nécessite révision:").font = HEADER_FONT
    ws.cell(row=row, column=2, value="OUI" if result.summary.needs_review else "NON")
    row += 1

    if result.summary.review_notes:
        row += 1
        ws.cell(row=row, column=1, value="Notes de révision:").font = HEADER_FONT
        row += 1
        for note in result.summary.review_notes:
            ws.cell(row=row, column=1, value=f"• {note}")
            row += 1

    # Flag summary
    if result.flags:
        row += 1
        ws.cell(row=row, column=1, value="Résumé des drapeaux").font = Font(bold=True, size=12)
        row += 1
        critical = sum(1 for f in result.flags if f.severity == FlagSeverity.CRITICAL)
        warning = sum(1 for f in result.flags if f.severity == FlagSeverity.WARNING)
        info = sum(1 for f in result.flags if f.severity == FlagSeverity.INFO)
        if critical:
            ws.cell(row=row, column=1, value=f"Critiques: {critical}").font = Font(color="9C0006", bold=True)
            row += 1
        if warning:
            ws.cell(row=row, column=1, value=f"Avertissements: {warning}").font = Font(color="9C6500", bold=True)
            row += 1
        if info:
            ws.cell(row=row, column=1, value=f"Informations: {info}").font = Font(color="1F4E79")
            row += 1

    _set_col_widths(ws, {"A": 35, "B": 40})


def _fill_comptes(ws, result: DPAuditResult) -> None:
    """Fill the Comptes sheet."""
    headers = ["ID", "Institution", "Titulaire", "Période début", "Période fin",
               "Solde ouverture", "Solde fermeture", "Confiance"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    _apply_header_style(ws, 1, len(headers))

    for i, acct in enumerate(result.accounts, start=2):
        row_data = [
            acct.account_id, acct.institution, acct.holder_name,
            acct.period_start, acct.period_end,
            acct.opening_balance, acct.closing_balance, acct.confidence,
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = THIN_BORDER
            if col in (6, 7):
                cell.number_format = CURRENCY_FORMAT

    _set_col_widths(ws, {"A": 8, "B": 25, "C": 25, "D": 14, "E": 14, "F": 18, "G": 18, "H": 12})


def _fill_key_transactions(ws, result: DPAuditResult) -> None:
    """Fill the Transactions clés sheet — flagged deposits and large deposits."""
    headers = ["ID", "Date", "Compte", "Description", "Montant", "Catégorie", "Page", "Flags"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    _apply_header_style(ws, 1, len(headers))

    # Collect flagged transaction IDs
    flagged_ids: dict[str, list[str]] = {}
    for flag in result.flags:
        for tid in flag.supporting_transaction_ids:
            flagged_ids.setdefault(tid, []).append(flag.type.value)

    # Key transactions: flagged or large deposits
    transfer_tx_ids = {m.from_transaction_id for m in result.transfers} | {m.to_transaction_id for m in result.transfers}
    key_txns = [
        t for t in result.transactions
        if t.id in flagged_ids or (t.type.value == "deposit" and t.amount >= 5000 and t.id not in transfer_tx_ids)
    ]
    key_txns.sort(key=lambda t: t.date)

    for i, txn in enumerate(key_txns, start=2):
        flag_text = ", ".join(flagged_ids.get(txn.id, []))
        row_data = [
            txn.id, txn.date, txn.account_id, txn.description,
            txn.amount, txn.category.value, txn.page_source, flag_text,
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = THIN_BORDER
            if col == 5:
                cell.number_format = CURRENCY_FORMAT

    _set_col_widths(ws, {"A": 10, "B": 12, "C": 10, "D": 45, "E": 16, "F": 18, "G": 8, "H": 30})


def _fill_transfers(ws, result: DPAuditResult) -> None:
    """Fill the Transferts sheet."""
    headers = ["De (compte)", "Vers (compte)", "Montant", "ID retrait", "ID(s) dépôt", "Delta jours", "Score", "Split"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    _apply_header_style(ws, 1, len(headers))

    for i, tm in enumerate(result.transfers, start=2):
        dep_ids = ", ".join(tm.to_transaction_ids) if tm.to_transaction_ids else tm.to_transaction_id
        row_data = [
            tm.from_account_id, tm.to_account_id, tm.amount,
            tm.from_transaction_id, dep_ids,
            tm.date_delta_days, tm.match_score,
            "OUI" if tm.is_split else "",
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = THIN_BORDER
            if col == 3:
                cell.number_format = CURRENCY_FORMAT

    _set_col_widths(ws, {"A": 14, "B": 14, "C": 16, "D": 14, "E": 20, "F": 12, "G": 10, "H": 8})


def _fill_flags_and_requests(ws, result: DPAuditResult) -> None:
    """Fill the Flags & Demandes sheet."""
    row = 1
    ws.cell(row=row, column=1, value="Drapeaux d'audit").font = Font(bold=True, size=12)
    row += 1

    flag_headers = ["Type", "Sévérité", "Explication", "Transactions", "Docs recommandés"]
    for col, h in enumerate(flag_headers, 1):
        ws.cell(row=row, column=col, value=h)
    _apply_header_style(ws, row, len(flag_headers))
    row += 1

    for flag in result.flags:
        ws.cell(row=row, column=1, value=flag.type.value).border = THIN_BORDER
        sev_cell = ws.cell(row=row, column=2, value=flag.severity.value)
        sev_cell.border = THIN_BORDER
        sev_cell.fill = SEVERITY_FILLS.get(flag.severity, PatternFill())
        sev_cell.font = SEVERITY_FONTS.get(flag.severity, Font())
        ws.cell(row=row, column=3, value=flag.rationale).border = THIN_BORDER
        ws.cell(row=row, column=4, value=", ".join(flag.supporting_transaction_ids)).border = THIN_BORDER
        ws.cell(row=row, column=5, value=", ".join(flag.recommended_documents)).border = THIN_BORDER
        row += 1

    # Client requests
    row += 2
    ws.cell(row=row, column=1, value="Demandes au client").font = Font(bold=True, size=12)
    row += 1

    req_headers = ["Titre", "Raison", "Documents requis", "Transactions"]
    for col, h in enumerate(req_headers, 1):
        ws.cell(row=row, column=col, value=h)
    _apply_header_style(ws, row, len(req_headers))
    row += 1

    for req in result.client_requests:
        ws.cell(row=row, column=1, value=req.title).border = THIN_BORDER
        ws.cell(row=row, column=2, value=req.reason).border = THIN_BORDER
        ws.cell(row=row, column=3, value=", ".join(req.required_docs)).border = THIN_BORDER
        ws.cell(row=row, column=4, value=", ".join(req.supporting_transaction_ids)).border = THIN_BORDER
        row += 1

    _set_col_widths(ws, {"A": 25, "B": 20, "C": 55, "D": 25, "E": 45})


# ── Public API ────────────────────────────────────────────────────────────


def generate_dp_excel(result: DPAuditResult) -> bytes:
    """Generate the downpayment audit Excel workbook and return bytes."""
    wb = Workbook()

    ws_resume = wb.active
    ws_resume.title = "Résumé"
    ws_comptes = wb.create_sheet("Comptes")
    ws_key_txns = wb.create_sheet("Transactions clés")
    ws_transfers = wb.create_sheet("Transferts")
    ws_flags = wb.create_sheet("Flags & Demandes")

    _fill_resume(ws_resume, result)
    _fill_comptes(ws_comptes, result)
    _fill_key_transactions(ws_key_txns, result)
    _fill_transfers(ws_transfers, result)
    _fill_flags_and_requests(ws_flags, result)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_dp_excel_base64(result: DPAuditResult) -> str:
    """Generate Excel and return as base64 string."""
    return base64.b64encode(generate_dp_excel(result)).decode()
