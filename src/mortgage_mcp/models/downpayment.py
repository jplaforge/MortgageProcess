"""Pydantic models for downpayment / source of funds audit.

These models serve double duty:
1. Define the structured output schema for Gemini extraction (DPExtraction)
2. Type the data flowing through the post-processing pipeline (DPAuditResult)
"""

from enum import Enum

from pydantic import BaseModel, Field


# ── Gemini extraction enums & models ──────────────────────────────────────


class TransactionType(str, Enum):
    DEPOSIT = "deposit"
    WITHDRAWAL = "withdrawal"


class TransactionCategory(str, Enum):
    PAYROLL = "payroll"
    BUSINESS_INCOME = "business_income"
    TRANSFER = "transfer"
    CASH = "cash"
    GOVERNMENT = "government"
    INVESTMENT = "investment"
    GIFT = "gift"
    LOAN = "loan"
    REFUND = "refund"
    BILL_PAYMENT = "bill_payment"
    PURCHASE = "purchase"
    OTHER = "other"


class DPTransaction(BaseModel):
    """A single transaction extracted from a bank statement."""

    id: str = Field(description="Identifiant unique (ex: A1-001)")
    date: str = Field(description="Date de la transaction (YYYY-MM-DD)")
    description: str = Field(description="Description telle qu'affichée sur le relevé")
    amount: float = Field(description="Montant en CAD (toujours positif)")
    type: TransactionType = Field(description="Type: deposit ou withdrawal")
    category: TransactionCategory = Field(description="Catégorie de la transaction")
    account_id: str = Field(description="Identifiant du compte (ex: A1)")
    page_source: int = Field(default=0, description="Page source dans le document")
    confidence: float = Field(default=1.0, description="Score de confiance 0-1")
    normalized_description: str = Field(default="", description="Description normalisée")
    merchant_guess: str = Field(default="", description="Estimation du commerçant/source")


class DPAccountInfo(BaseModel):
    """Metadata for a single bank account."""

    account_id: str = Field(description="Identifiant interne du compte (ex: A1)")
    institution: str = Field(description="Nom de l'institution financière")
    account_number_last4: str = Field(default="", description="4 derniers chiffres du numéro de compte")
    holder_name: str = Field(default="", description="Nom du titulaire")
    period_start: str = Field(default="", description="Début de la période couverte (YYYY-MM-DD)")
    period_end: str = Field(default="", description="Fin de la période couverte (YYYY-MM-DD)")
    opening_balance: float = Field(default=0.0, description="Solde d'ouverture")
    closing_balance: float = Field(default=0.0, description="Solde de fermeture")
    confidence: float = Field(default=1.0, description="Score de confiance 0-1")


class DPExtraction(BaseModel):
    """Full Gemini extraction result for downpayment audit."""

    accounts: list[DPAccountInfo] = Field(description="Comptes bancaires identifiés")
    transactions: list[DPTransaction] = Field(description="Toutes les transactions extraites")


# ── Post-processing enums & models ────────────────────────────────────────


class FlagSeverity(str, Enum):
    CRITICAL = "critical"
    WARNING = "warning"
    INFO = "info"


class FlagType(str, Enum):
    LARGE_DEPOSIT = "large_deposit"
    CASH_DEPOSIT = "cash_deposit"
    NON_PAYROLL_RECURRING = "non_payroll_recurring"
    MULTI_HOP_TRANSFER = "multi_hop_transfer"
    PERIOD_GAP = "period_gap"
    ROUND_AMOUNT = "round_amount"
    RAPID_SUCCESSION = "rapid_succession"
    UNEXPLAINED_SOURCE = "unexplained_source"


class TransferMatch(BaseModel):
    """A matched inter-account transfer (1:1 or 1:N split)."""

    from_account_id: str = Field(description="Compte source du retrait")
    to_account_id: str = Field(description="Compte destination du dépôt")
    amount: float = Field(description="Montant du transfert (côté retrait)")
    from_transaction_id: str = Field(description="ID de la transaction retrait")
    to_transaction_id: str = Field(default="", description="ID de la transaction dépôt (1:1 match)")
    to_transaction_ids: list[str] = Field(default_factory=list, description="IDs des dépôts (1:N split match)")
    date_delta_days: int = Field(description="Écart en jours entre retrait et dépôt")
    match_score: float = Field(description="Score de correspondance 0-1")
    is_split: bool = Field(default=False, description="True si transfert fractionné (1:N)")


class DPFlag(BaseModel):
    """A flag raised during audit analysis."""

    type: FlagType = Field(description="Type de drapeau")
    severity: FlagSeverity = Field(description="Sévérité: critical, warning, info")
    rationale: str = Field(description="Explication en français")
    supporting_transaction_ids: list[str] = Field(default_factory=list, description="IDs des transactions concernées")
    recommended_documents: list[str] = Field(default_factory=list, description="Documents recommandés à obtenir")


class ClientRequest(BaseModel):
    """A document request to send to the client."""

    title: str = Field(description="Titre de la demande")
    reason: str = Field(description="Raison de la demande")
    required_docs: list[str] = Field(default_factory=list, description="Documents requis")
    supporting_transaction_ids: list[str] = Field(default_factory=list, description="IDs des transactions concernées")


class SourceBreakdown(BaseModel):
    """Breakdown of downpayment funding sources."""

    savings: float = Field(default=0.0, description="Épargne accumulée (paie)")
    gift: float = Field(default=0.0, description="Dons")
    investment_sale: float = Field(default=0.0, description="Vente de placements")
    property_sale: float = Field(default=0.0, description="Vente de propriété")
    payroll: float = Field(default=0.0, description="Accumulation salariale")
    other_explained: float = Field(default=0.0, description="Autres sources expliquées")
    unexplained: float = Field(default=0.0, description="Sources non expliquées")


class DPSummary(BaseModel):
    """High-level audit summary."""

    dp_target: float = Field(description="Mise de fonds cible")
    dp_explained_amount: float = Field(default=0.0, description="Montant expliqué")
    unexplained_amount: float = Field(default=0.0, description="Montant non expliqué")
    needs_review: bool = Field(default=False, description="Nécessite une révision manuelle")
    review_notes: list[str] = Field(default_factory=list, description="Notes de révision")
    source_breakdown: SourceBreakdown = Field(default_factory=SourceBreakdown, description="Ventilation des sources")


class DPAuditResult(BaseModel):
    """Complete downpayment audit result."""

    accounts: list[DPAccountInfo] = Field(default_factory=list, description="Comptes analysés")
    transactions: list[DPTransaction] = Field(default_factory=list, description="Toutes les transactions")
    transfers: list[TransferMatch] = Field(default_factory=list, description="Transferts inter-comptes matchés")
    flags: list[DPFlag] = Field(default_factory=list, description="Drapeaux d'audit")
    client_requests: list[ClientRequest] = Field(default_factory=list, description="Demandes au client")
    summary: DPSummary = Field(description="Résumé de l'audit")
    borrower_name: str = Field(default="", description="Nom de l'emprunteur")
    co_borrower_name: str = Field(default="", description="Nom du co-emprunteur")
    closing_date: str = Field(default="", description="Date de clôture prévue (YYYY-MM-DD)")
    deal_notes: str = Field(default="", description="Notes sur le dossier")
