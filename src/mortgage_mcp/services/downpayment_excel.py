"""Generate the Excel downpayment audit report from analysis results.

Produces a 3-sheet workbook designed for mortgage brokers:
1. Résumé — verdict, key numbers, accounts, sources, transfers, flags
2. Demandes au client — actionable document requests
3. Détail — flagged transactions needing verification
"""

import base64
import io
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from mortgage_mcp.models.downpayment import (
    DPAccountInfo,
    DPAuditResult,
    DPTransaction,
    FlagSeverity,
    FlagType,
    TransactionCategory,
    TransactionType,
)

# ── Styling constants ─────────────────────────────────────────────────────

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

VERDICT_STYLES = {
    "green": (
        PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        Font(bold=True, size=14, color="006100"),
    ),
    "yellow": (
        PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        Font(bold=True, size=14, color="9C6500"),
    ),
    "red": (
        PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        Font(bold=True, size=14, color="9C0006"),
    ),
}

# ── French label maps ─────────────────────────────────────────────────────

FLAG_TYPE_LABELS: dict[FlagType, str] = {
    FlagType.LARGE_DEPOSIT: "Dépôt important",
    FlagType.CASH_DEPOSIT: "Dépôt en espèces",
    FlagType.NON_PAYROLL_RECURRING: "Récurrent non-salarial",
    FlagType.MULTI_HOP_TRANSFER: "Chaîne de transferts",
    FlagType.PERIOD_GAP: "Couverture insuffisante",
    FlagType.ROUND_AMOUNT: "Montant rond",
    FlagType.RAPID_SUCCESSION: "Succession rapide",
    FlagType.UNEXPLAINED_SOURCE: "Source inexpliquée",
}

SEVERITY_LABELS: dict[FlagSeverity, str] = {
    FlagSeverity.CRITICAL: "Critique",
    FlagSeverity.WARNING: "Avertissement",
    FlagSeverity.INFO: "Information",
}

CATEGORY_LABELS: dict[TransactionCategory, str] = {
    TransactionCategory.PAYROLL: "Salaire",
    TransactionCategory.BUSINESS_INCOME: "Revenu d'affaires",
    TransactionCategory.TRANSFER: "Transfert",
    TransactionCategory.CASH: "Espèces",
    TransactionCategory.GOVERNMENT: "Gouvernement",
    TransactionCategory.INVESTMENT: "Placement",
    TransactionCategory.GIFT: "Don",
    TransactionCategory.LOAN: "Prêt",
    TransactionCategory.REFUND: "Remboursement",
    TransactionCategory.BILL_PAYMENT: "Facture",
    TransactionCategory.PURCHASE: "Achat",
    TransactionCategory.OTHER: "Autre",
}

MONTH_NAMES_FR = {
    1: "janv.", 2: "févr.", 3: "mars", 4: "avr.", 5: "mai", 6: "juin",
    7: "juill.", 8: "août", 9: "sept.", 10: "oct.", 11: "nov.", 12: "déc.",
}

_SEVERITY_ORDER = {FlagSeverity.CRITICAL: 0, FlagSeverity.WARNING: 1, FlagSeverity.INFO: 2}


# ── Helpers ───────────────────────────────────────────────────────────────


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


def _format_date_short(date_str: str) -> str:
    """Format YYYY-MM-DD as '20 févr. 2025'."""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return f"{dt.day} {MONTH_NAMES_FR[dt.month]} {dt.year}"
    except (ValueError, KeyError):
        return date_str


def _build_account_lookup(accounts: list[DPAccountInfo]) -> dict[str, str]:
    """Map account_id -> 'Institution (XXXX)' for human-readable labels."""
    lookup = {}
    for acct in accounts:
        label = acct.institution
        if acct.account_number_last4:
            label += f" ({acct.account_number_last4})"
        lookup[acct.account_id] = label
    return lookup


def _build_tx_lookup(
    transactions: list[DPTransaction],
    acct_lookup: dict[str, str],
) -> dict[str, str]:
    """Map transaction ID -> human-readable summary."""
    lookup = {}
    for tx in transactions:
        acct_label = acct_lookup.get(tx.account_id, tx.account_id)
        date_short = _format_date_short(tx.date)
        desc = tx.description[:30] + "..." if len(tx.description) > 30 else tx.description
        lookup[tx.id] = f"{tx.amount:,.0f} $ — {desc} ({date_short}, {acct_label})"
    return lookup


def _section_header(ws, row: int, title: str) -> int:
    """Write a section header and return the next row."""
    ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=12)
    return row + 1


def _all_transfer_tx_ids(result: DPAuditResult) -> set[str]:
    """Collect all transaction IDs involved in matched transfers."""
    ids: set[str] = set()
    for m in result.transfers:
        ids.add(m.from_transaction_id)
        if m.to_transaction_id:
            ids.add(m.to_transaction_id)
        ids.update(m.to_transaction_ids)
    return ids


# ── Sheet 1: Résumé ──────────────────────────────────────────────────────


def _fill_resume(ws, result: DPAuditResult) -> None:
    """Fill the Résumé sheet: verdict, key numbers, accounts, sources, transfers, flags."""
    acct_lookup = _build_account_lookup(result.accounts)
    summary = result.summary

    row = 1
    ws.cell(row=row, column=1, value="Audit de la mise de fonds — Provenance des fonds").font = Font(bold=True, size=14)
    row += 2

    # ── Verdict ──
    has_critical = any(f.severity == FlagSeverity.CRITICAL for f in result.flags)
    pct = summary.dp_explained_amount / summary.dp_target if summary.dp_target > 0 else 0

    if not summary.needs_review:
        verdict_text, style_key = "CONFORME", "green"
    elif has_critical:
        verdict_text, style_key = "RÉVISION REQUISE", "red"
    else:
        verdict_text, style_key = "À VÉRIFIER", "yellow"

    verdict_fill, verdict_font = VERDICT_STYLES[style_key]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    verdict_cell = ws.cell(row=row, column=1, value=verdict_text)
    verdict_cell.font = verdict_font
    verdict_cell.fill = verdict_fill
    verdict_cell.alignment = Alignment(horizontal="center", vertical="center")
    verdict_cell.border = THIN_BORDER
    for col in range(2, 4):
        ws.cell(row=row, column=col).fill = verdict_fill
        ws.cell(row=row, column=col).border = THIN_BORDER
    row += 1

    # Progress line
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    progress_text = f"{pct:.0%} expliqué — {summary.dp_explained_amount:,.0f} $ / {summary.dp_target:,.0f} $"
    progress_cell = ws.cell(row=row, column=1, value=progress_text)
    progress_cell.font = Font(bold=True, size=11)
    progress_cell.alignment = Alignment(horizontal="center")
    row += 2

    # ── Deal info ──
    row = _section_header(ws, row, "Informations du dossier")
    deal_data: list[tuple[str, str | float]] = [("Emprunteur:", result.borrower_name)]
    if result.co_borrower_name:
        deal_data.append(("Co-emprunteur:", result.co_borrower_name))
    deal_data.append(("Date de clôture:", _format_date_short(result.closing_date)))
    deal_data.append(("Mise de fonds cible:", summary.dp_target))

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

    # ── Accounts ──
    if result.accounts:
        row = _section_header(ws, row, "Comptes analysés")
        acct_headers = ["Institution", "Titulaire", "No compte", "Période", "Solde ouverture", "Solde fermeture"]
        for col, h in enumerate(acct_headers, 1):
            ws.cell(row=row, column=col, value=h)
        _apply_header_style(ws, row, len(acct_headers))
        row += 1

        for acct in result.accounts:
            period = ""
            if acct.period_start and acct.period_end:
                period = f"{_format_date_short(acct.period_start)} — {_format_date_short(acct.period_end)}"
            row_data = [
                acct.institution, acct.holder_name,
                acct.account_number_last4 or "—",
                period,
                acct.opening_balance, acct.closing_balance,
            ]
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.border = THIN_BORDER
                if col in (5, 6):
                    cell.number_format = CURRENCY_FORMAT
            row += 1
        row += 1

    # ── Source breakdown (skip zeros) ──
    row = _section_header(ws, row, "Ventilation des sources")
    sb = summary.source_breakdown
    sources = [
        ("Accumulation salariale:", sb.payroll),
        ("Dons:", sb.gift),
        ("Vente de placements:", sb.investment_sale),
        ("Vente de propriété:", sb.property_sale),
        ("Autres sources expliquées:", sb.other_explained),
    ]
    for label, value in sources:
        if value > 0:
            ws.cell(row=row, column=1, value=label).font = HEADER_FONT
            ws.cell(row=row, column=2, value=value).number_format = CURRENCY_FORMAT
            row += 1

    if sb.unexplained > 0:
        ws.cell(row=row, column=1, value="Sources non expliquées:").font = HEADER_FONT
        cell = ws.cell(row=row, column=2, value=sb.unexplained)
        cell.number_format = CURRENCY_FORMAT
        cell.font = Font(color="9C0006", bold=True)
        row += 1
    row += 1

    # ── Transfers (inline, only if any) ──
    if result.transfers:
        row = _section_header(ws, row, "Transferts inter-comptes détectés")
        tf_headers = ["De", "Vers", "Montant", "Détail"]
        for col, h in enumerate(tf_headers, 1):
            ws.cell(row=row, column=col, value=h)
        _apply_header_style(ws, row, len(tf_headers))
        row += 1

        for tm in result.transfers:
            from_label = acct_lookup.get(tm.from_account_id, tm.from_account_id)
            to_label = acct_lookup.get(tm.to_account_id, tm.to_account_id)
            if tm.is_split:
                detail = f"Fractionné en {len(tm.to_transaction_ids)} dépôts"
            elif tm.date_delta_days > 0:
                detail = f"Délai: {tm.date_delta_days} jour(s)"
            else:
                detail = "Même jour"
            for col, val in enumerate([from_label, to_label, tm.amount, detail], 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.border = THIN_BORDER
                if col == 3:
                    cell.number_format = CURRENCY_FORMAT
            row += 1
        row += 1

    # ── Flags ──
    if result.flags:
        row = _section_header(ws, row, "Drapeaux d'audit")
        critical = sum(1 for f in result.flags if f.severity == FlagSeverity.CRITICAL)
        warning = sum(1 for f in result.flags if f.severity == FlagSeverity.WARNING)
        info = sum(1 for f in result.flags if f.severity == FlagSeverity.INFO)
        counts = []
        if critical:
            counts.append(f"{critical} critique(s)")
        if warning:
            counts.append(f"{warning} avertissement(s)")
        if info:
            counts.append(f"{info} information(s)")
        ws.cell(row=row, column=1, value=" — ".join(counts)).font = Font(bold=True, size=10)
        row += 1

        flag_headers = ["Type", "Sévérité", "Explication"]
        for col, h in enumerate(flag_headers, 1):
            ws.cell(row=row, column=col, value=h)
        _apply_header_style(ws, row, len(flag_headers))
        row += 1

        sorted_flags = sorted(result.flags, key=lambda f: _SEVERITY_ORDER.get(f.severity, 3))
        for flag in sorted_flags:
            ws.cell(row=row, column=1, value=FLAG_TYPE_LABELS.get(flag.type, flag.type.value)).border = THIN_BORDER
            sev_cell = ws.cell(row=row, column=2, value=SEVERITY_LABELS.get(flag.severity, flag.severity.value))
            sev_cell.border = THIN_BORDER
            sev_cell.fill = SEVERITY_FILLS.get(flag.severity, PatternFill())
            sev_cell.font = SEVERITY_FONTS.get(flag.severity, Font())
            ws.cell(row=row, column=3, value=flag.rationale).border = THIN_BORDER
            row += 1
        row += 1

    # ── Review notes ──
    if summary.review_notes:
        row = _section_header(ws, row, "Notes de révision")
        for note in summary.review_notes:
            ws.cell(row=row, column=1, value=f"• {note}")
            row += 1

    _set_col_widths(ws, {"A": 35, "B": 30, "C": 25, "D": 20, "E": 18, "F": 18})


# ── Sheet 2: Demandes au client ──────────────────────────────────────────


def _fill_demandes(ws, result: DPAuditResult) -> None:
    """Fill the Demandes au client sheet — numbered action cards."""
    acct_lookup = _build_account_lookup(result.accounts)
    tx_lookup = _build_tx_lookup(result.transactions, acct_lookup)

    row = 1
    ws.cell(row=row, column=1, value="Demandes au client").font = Font(bold=True, size=14)
    row += 1

    if not result.client_requests:
        ws.cell(row=row, column=1, value="Aucune demande requise.").font = Font(italic=True, size=11)
        _set_col_widths(ws, {"A": 50})
        return

    ws.cell(
        row=row, column=1,
        value=f"{len(result.client_requests)} document(s) à demander au client",
    ).font = Font(size=11)
    row += 2

    for i, req in enumerate(result.client_requests, 1):
        ws.cell(row=row, column=1, value=f"{i}. {req.title}").font = Font(bold=True, size=12)
        row += 1

        ws.cell(row=row, column=1, value="Raison:").font = HEADER_FONT
        ws.cell(row=row, column=2, value=req.reason)
        row += 1

        ws.cell(row=row, column=1, value="Documents requis:").font = HEADER_FONT
        for doc in req.required_docs:
            ws.cell(row=row, column=2, value=f"• {doc}")
            row += 1

        if req.supporting_transaction_ids:
            ws.cell(row=row, column=1, value="Transactions concernées:").font = HEADER_FONT
            for tid in req.supporting_transaction_ids:
                ref = tx_lookup.get(tid, tid)
                ws.cell(row=row, column=2, value=f"• {ref}")
                row += 1

        row += 1  # blank line between requests

    _set_col_widths(ws, {"A": 30, "B": 70})


# ── Sheet 3: Détail ──────────────────────────────────────────────────────


def _fill_detail(ws, result: DPAuditResult) -> None:
    """Fill the Détail sheet — significant transactions needing attention.

    Only includes WARNING+ flagged transactions and large unflagged deposits.
    INFO-only flags (like non_payroll_recurring) are excluded to reduce noise.
    """
    acct_lookup = _build_account_lookup(result.accounts)

    row = 1
    ws.cell(row=row, column=1, value="Transactions nécessitant une vérification").font = Font(bold=True, size=14)
    row += 2

    # Build flagged IDs map: tx_id -> [(french_label, severity)]
    flagged_ids: dict[str, list[tuple[str, FlagSeverity]]] = {}
    for flag in result.flags:
        for tid in flag.supporting_transaction_ids:
            flagged_ids.setdefault(tid, []).append(
                (FLAG_TYPE_LABELS.get(flag.type, flag.type.value), flag.severity)
            )

    transfer_tx_ids = _all_transfer_tx_ids(result)

    # Filter: WARNING+ flagged transactions, or large deposits (>=5k) not in transfers
    key_txns: list[DPTransaction] = []
    for t in result.transactions:
        if t.id in flagged_ids:
            highest_sev = min(
                (sev for _, sev in flagged_ids[t.id]),
                key=lambda s: _SEVERITY_ORDER[s],
            )
            if highest_sev != FlagSeverity.INFO:
                key_txns.append(t)
        elif (
            t.type == TransactionType.DEPOSIT
            and t.amount >= 5000
            and t.id not in transfer_tx_ids
        ):
            key_txns.append(t)
    key_txns.sort(key=lambda t: t.date)

    if not key_txns:
        ws.cell(row=row, column=1, value="Aucune transaction nécessitant une vérification.").font = Font(
            italic=True, size=11,
        )
        _set_col_widths(ws, {"A": 50})
        return

    headers = ["Date", "Compte", "Description", "Montant", "Catégorie", "Drapeaux"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=row, column=col, value=h)
    _apply_header_style(ws, row, len(headers))
    row += 1

    for txn in key_txns:
        acct_label = acct_lookup.get(txn.account_id, txn.account_id)
        flag_info = flagged_ids.get(txn.id, [])
        flag_text = ", ".join(label for label, _ in flag_info) if flag_info else ""

        highest_sev = None
        if flag_info:
            highest_sev = min(
                (sev for _, sev in flag_info),
                key=lambda s: _SEVERITY_ORDER[s],
            )

        row_data = [
            _format_date_short(txn.date),
            acct_label,
            txn.description,
            txn.amount,
            CATEGORY_LABELS.get(txn.category, txn.category.value),
            flag_text,
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.border = THIN_BORDER
            if col == 4:
                cell.number_format = CURRENCY_FORMAT
            if col == 6 and highest_sev:
                cell.fill = SEVERITY_FILLS.get(highest_sev, PatternFill())
                cell.font = SEVERITY_FONTS.get(highest_sev, Font())
        row += 1

    _set_col_widths(ws, {"A": 18, "B": 28, "C": 45, "D": 16, "E": 18, "F": 30})


# ── Public API ────────────────────────────────────────────────────────────


def generate_dp_excel(result: DPAuditResult) -> bytes:
    """Generate the downpayment audit Excel workbook and return bytes."""
    wb = Workbook()

    ws_resume = wb.active
    ws_resume.title = "Résumé"
    ws_demandes = wb.create_sheet("Demandes au client")
    ws_detail = wb.create_sheet("Détail")

    _fill_resume(ws_resume, result)
    _fill_demandes(ws_demandes, result)
    _fill_detail(ws_detail, result)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_dp_excel_base64(result: DPAuditResult) -> str:
    """Generate Excel and return as base64 string."""
    return base64.b64encode(generate_dp_excel(result)).decode()
