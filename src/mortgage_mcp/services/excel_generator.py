"""Generate the Excel income analysis report from extraction data."""

import base64
import io
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from mortgage_mcp.models.bank_statement import BankStatementExtraction

TEMPLATE_PATH = Path(__file__).parent.parent / "templates" / "grille_revenu_autonome.xlsx"

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


def _apply_header_style(ws, row: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def generate_excel(extraction: BankStatementExtraction) -> bytes:
    """Populate the Excel template with extraction data and return bytes.

    If the template file exists, it is used as-is. Otherwise, sheets are
    created from scratch with proper formatting.
    """
    if TEMPLATE_PATH.exists():
        wb = load_workbook(TEMPLATE_PATH)
    else:
        wb = _create_workbook_from_scratch()

    _fill_resume(wb, extraction)
    _fill_monthly(wb, extraction)
    _fill_deposits(wb, extraction)
    _fill_withdrawals(wb, extraction)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_excel_base64(extraction: BankStatementExtraction) -> str:
    """Generate Excel and return as base64 string."""
    return base64.b64encode(generate_excel(extraction)).decode()


def _create_workbook_from_scratch():
    """Create a new workbook with the three required sheets."""
    from openpyxl import Workbook

    wb = Workbook()
    # Rename default sheet
    ws_resume = wb.active
    ws_resume.title = "Resume"
    wb.create_sheet("Detail mensuel")
    wb.create_sheet("Depots")
    wb.create_sheet("Retraits")
    return wb


def _fill_resume(wb, extraction: BankStatementExtraction) -> None:
    """Fill the Resume (summary) sheet."""
    ws = wb["Resume"]

    # Clear existing data rows (keep structure)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    info = extraction.account_info
    current_row = 1

    # Title
    ws.cell(row=current_row, column=1, value="Grille d'analyse — Revenu de travailleur autonome").font = Font(bold=True, size=14)
    current_row += 2  # skip a blank row

    # Borrower info
    borrower_data = [
        ("Titulaire du compte:", info.account_holder),
        ("Institution financière:", info.institution),
        ("Numéro de compte:", f"***{info.account_number_last4}" if info.account_number_last4 else "N/D"),
        ("Période analysée:", f"{info.statement_period_start} au {info.statement_period_end}"),
        ("Nombre de mois:", extraction.months_covered),
    ]
    for label, value in borrower_data:
        ws.cell(row=current_row, column=1, value=label).font = HEADER_FONT
        ws.cell(row=current_row, column=2, value=value)
        current_row += 1
    current_row += 1  # blank row

    # Financial summary
    ws.cell(row=current_row, column=1, value="Sommaire financier").font = Font(bold=True, size=12)
    current_row += 1

    fin_data = [
        ("Dépôts totaux:", extraction.total_deposits),
        ("Revenu d'affaires total:", extraction.total_business_income),
        ("Retraits totaux:", extraction.total_withdrawals),
        ("Revenu mensuel moyen (affaires):", extraction.average_monthly_business_income),
        ("Revenu annualisé (affaires):", extraction.annualized_business_income),
    ]
    for label, value in fin_data:
        ws.cell(row=current_row, column=1, value=label).font = HEADER_FONT
        cell = ws.cell(row=current_row, column=2, value=value)
        cell.number_format = CURRENCY_FORMAT
        current_row += 1
    current_row += 1  # blank row

    # Risk indicators (NSF)
    if extraction.nsf_events:
        ws.cell(row=current_row, column=1, value="Indicateurs de risque").font = Font(bold=True, size=12)
        current_row += 1
        ws.cell(row=current_row, column=1, value="Événements NSF/découverts:").font = HEADER_FONT
        ws.cell(row=current_row, column=2, value=len(extraction.nsf_events))
        current_row += 1
        ws.cell(row=current_row, column=1, value="Frais NSF totaux:").font = HEADER_FONT
        ws.cell(row=current_row, column=2, value=extraction.nsf_total_fees).number_format = CURRENCY_FORMAT
        current_row += 2  # blank row

    # Recurring obligations
    if extraction.recurring_obligations:
        ws.cell(row=current_row, column=1, value="Obligations récurrentes détectées").font = Font(bold=True, size=12)
        current_row += 1
        # Mini table headers
        for col, header in enumerate(["Bénéficiaire", "Montant mensuel", "Type"], 1):
            cell = ws.cell(row=current_row, column=col, value=header)
            cell.font = HEADER_FONT
            cell.border = THIN_BORDER
        current_row += 1
        for obligation in extraction.recurring_obligations:
            ws.cell(row=current_row, column=1, value=obligation.payee).border = THIN_BORDER
            cell_amt = ws.cell(row=current_row, column=2, value=obligation.monthly_amount)
            cell_amt.number_format = CURRENCY_FORMAT
            cell_amt.border = THIN_BORDER
            ws.cell(row=current_row, column=3, value=obligation.category).border = THIN_BORDER
            current_row += 1
        ws.cell(row=current_row, column=1, value="Total obligations mensuelles:").font = HEADER_FONT
        ws.cell(row=current_row, column=2, value=extraction.total_monthly_obligations).number_format = CURRENCY_FORMAT
        current_row += 2  # blank row

    # Confidence notes
    if extraction.confidence_notes:
        ws.cell(row=current_row, column=1, value="Notes et observations").font = Font(bold=True, size=12)
        current_row += 1
        for note in extraction.confidence_notes:
            ws.cell(row=current_row, column=1, value=f"• {note}")
            current_row += 1

    # Column widths
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 20


def _fill_monthly(wb, extraction: BankStatementExtraction) -> None:
    """Fill the Detail mensuel (monthly breakdown) sheet."""
    ws = wb["Detail mensuel"]

    # Clear
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    headers = [
        "Mois", "Dépôts bruts", "Dépôts affaires", "Transferts personnels",
        "Gouvernement", "Remboursements", "Prêts/crédit", "Autres",
        "Retraits", "Revenu net", "Nb dépôts",
    ]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    _apply_header_style(ws, 1, len(headers))

    currency_cols = len(headers) - 1  # all except last (Nb dépôts)

    for i, month in enumerate(extraction.monthly_breakdown, start=2):
        net = month.business_deposits - month.total_withdrawals
        row_data = [
            month.month,
            month.total_deposits,
            month.business_deposits,
            month.personal_transfers,
            month.government_deposits,
            month.refund_deposits,
            month.loan_credit_deposits,
            month.other_deposits,
            month.total_withdrawals,
            net,
            month.deposit_count,
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = THIN_BORDER
            if 2 <= col <= currency_cols:
                cell.number_format = CURRENCY_FORMAT

    # Summary rows
    n = len(extraction.monthly_breakdown)
    if n > 0:
        summary_row = n + 2
        last_data_row = n + 1

        ws.cell(row=summary_row, column=1, value="TOTAL").font = HEADER_FONT
        for col in range(2, currency_cols + 1):
            col_letter = chr(64 + col)
            cell = ws.cell(
                row=summary_row, column=col,
                value=f"=SUM({col_letter}2:{col_letter}{last_data_row})"
            )
            cell.number_format = CURRENCY_FORMAT
            cell.font = HEADER_FONT
            cell.border = THIN_BORDER

        avg_row = summary_row + 1
        ws.cell(row=avg_row, column=1, value="MOYENNE").font = HEADER_FONT
        for col in range(2, currency_cols + 1):
            col_letter = chr(64 + col)
            cell = ws.cell(
                row=avg_row, column=col,
                value=f"=AVERAGE({col_letter}2:{col_letter}{last_data_row})"
            )
            cell.number_format = CURRENCY_FORMAT
            cell.font = HEADER_FONT
            cell.border = THIN_BORDER

    # Column widths
    widths = [12, 16, 16, 20, 16, 18, 16, 14, 16, 16, 12]
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w


def _fill_deposits(wb, extraction: BankStatementExtraction) -> None:
    """Fill the Depots (all deposits) sheet."""
    ws = wb["Depots"]

    # Clear
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    headers = ["Date", "Compte", "Description", "Montant", "Catégorie"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    _apply_header_style(ws, 1, len(headers))

    row_num = 2
    for month in extraction.monthly_breakdown:
        for dep in month.deposits:
            ws.cell(row=row_num, column=1, value=dep.date).border = THIN_BORDER
            ws.cell(row=row_num, column=2, value=dep.account).border = THIN_BORDER
            ws.cell(row=row_num, column=3, value=dep.description).border = THIN_BORDER
            cell_amt = ws.cell(row=row_num, column=4, value=dep.amount)
            cell_amt.number_format = CURRENCY_FORMAT
            cell_amt.border = THIN_BORDER
            ws.cell(row=row_num, column=5, value=dep.category.value).border = THIN_BORDER
            row_num += 1

    # Column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 20


def _fill_withdrawals(wb, extraction: BankStatementExtraction) -> None:
    """Fill the Retraits (all withdrawals) sheet."""
    if "Retraits" not in wb.sheetnames:
        wb.create_sheet("Retraits")
    ws = wb["Retraits"]

    # Clear
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    headers = ["Date", "Compte", "Description", "Montant", "Catégorie"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    _apply_header_style(ws, 1, len(headers))

    row_num = 2
    for month in extraction.monthly_breakdown:
        for wd in month.withdrawals:
            ws.cell(row=row_num, column=1, value=wd.date).border = THIN_BORDER
            ws.cell(row=row_num, column=2, value=wd.account).border = THIN_BORDER
            ws.cell(row=row_num, column=3, value=wd.description).border = THIN_BORDER
            cell_amt = ws.cell(row=row_num, column=4, value=wd.amount)
            cell_amt.number_format = CURRENCY_FORMAT
            cell_amt.border = THIN_BORDER
            ws.cell(row=row_num, column=5, value=wd.category).border = THIN_BORDER
            row_num += 1

    # Column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 20
