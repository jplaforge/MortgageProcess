"""Shared test fixtures."""

import pytest

from mortgage_mcp.models.bank_statement import (
    AccountInfo,
    BankStatementExtraction,
    Deposit,
    DepositCategory,
    MonthlyBreakdown,
    NSFEvent,
    RecurringObligation,
    Withdrawal,
)
from mortgage_mcp.models.downpayment import (
    ClientRequest,
    DPAccountInfo,
    DPAuditResult,
    DPExtraction,
    DPFlag,
    DPSummary,
    DPTransaction,
    FlagSeverity,
    FlagType,
    SourceBreakdown,
    TransactionCategory,
    TransactionType,
    TransferMatch,
)


@pytest.fixture
def multi_account_extraction() -> BankStatementExtraction:
    """Multi-account extraction with inter-account transfers properly detected."""
    return BankStatementExtraction(
        account_info=AccountInfo(
            account_holder="Martin Girard",
            institution="Banque Boréale / Caisse Laurentienne",
            account_number_last4="XX89",
            statement_period_start="2025-01-01",
            statement_period_end="2025-03-31",
        ),
        monthly_breakdown=[
            MonthlyBreakdown(
                month="2025-01",
                total_deposits=17500.00,
                business_deposits=14000.00,
                personal_transfers=3500.00,
                government_deposits=0.00,
                refund_deposits=0.00,
                loan_credit_deposits=0.00,
                other_deposits=0.00,
                total_withdrawals=9500.00,
                deposit_count=8,
                deposits=[
                    Deposit(
                        date="2025-01-03",
                        description="VIREMENT INTERAC - CLIENT RENOVATION",
                        amount=5000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Boréale XX89",
                    ),
                    Deposit(
                        date="2025-01-10",
                        description="DEPOT CHEQUE - CONTRAT PLOMBERIE",
                        amount=4000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Boréale XX89",
                    ),
                    Deposit(
                        date="2025-01-15",
                        description="VIREMENT INTERAC - CLIENT ELECTRIQUE",
                        amount=3000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Laurentienne XX77",
                    ),
                    Deposit(
                        date="2025-01-17",
                        description="DEPOT CHEQUE - SOUS-CONTRAT",
                        amount=2000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Laurentienne XX77",
                    ),
                    # Inter-account transfers (detected)
                    Deposit(
                        date="2025-01-15",
                        description="VIREMENT TFR COMPTE BOREALE",
                        amount=2000.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                        account="Laurentienne XX77",
                    ),
                    Deposit(
                        date="2025-01-28",
                        description="TRANSFERT DEPUIS LAURENTIENNE",
                        amount=1500.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                        account="Boréale XX89",
                    ),
                ],
                withdrawals=[
                    Withdrawal(
                        date="2025-01-15",
                        description="VIREMENT VERS LAURENTIENNE",
                        amount=2000.00,
                        category="transfert",
                        account="Boréale XX89",
                    ),
                    Withdrawal(
                        date="2025-01-28",
                        description="TFR VERS BOREALE",
                        amount=1500.00,
                        category="transfert",
                        account="Laurentienne XX77",
                    ),
                ],
            ),
            MonthlyBreakdown(
                month="2025-02",
                total_deposits=16000.00,
                business_deposits=13000.00,
                personal_transfers=3000.00,
                government_deposits=0.00,
                refund_deposits=0.00,
                loan_credit_deposits=0.00,
                other_deposits=0.00,
                total_withdrawals=8200.00,
                deposit_count=7,
                deposits=[
                    Deposit(
                        date="2025-02-05",
                        description="VIREMENT INTERAC - CLIENT CUISINE",
                        amount=6000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Boréale XX89",
                    ),
                    Deposit(
                        date="2025-02-12",
                        description="DEPOT CHEQUE - CONTRAT SALLE BAIN",
                        amount=4000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Laurentienne XX77",
                    ),
                    Deposit(
                        date="2025-02-20",
                        description="VIREMENT INTERAC - CLIENT TERRASSE",
                        amount=3000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Boréale XX89",
                    ),
                    # Inter-account transfers (detected)
                    Deposit(
                        date="2025-02-10",
                        description="TRANSFERT DEPUIS BOREALE",
                        amount=2000.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                        account="Laurentienne XX77",
                    ),
                    Deposit(
                        date="2025-02-25",
                        description="VIREMENT TFR LAURENTIENNE",
                        amount=1000.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                        account="Boréale XX89",
                    ),
                ],
                withdrawals=[
                    Withdrawal(
                        date="2025-02-10",
                        description="VIREMENT VERS LAURENTIENNE",
                        amount=2000.00,
                        category="transfert",
                        account="Boréale XX89",
                    ),
                    Withdrawal(
                        date="2025-02-25",
                        description="TFR VERS BOREALE",
                        amount=1000.00,
                        category="transfert",
                        account="Laurentienne XX77",
                    ),
                ],
            ),
            MonthlyBreakdown(
                month="2025-03",
                total_deposits=16000.00,
                business_deposits=13000.00,
                personal_transfers=3000.00,
                government_deposits=0.00,
                refund_deposits=0.00,
                loan_credit_deposits=0.00,
                other_deposits=0.00,
                total_withdrawals=9300.00,
                deposit_count=7,
                deposits=[
                    Deposit(
                        date="2025-03-04",
                        description="VIREMENT INTERAC - CLIENT TOITURE",
                        amount=7000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Boréale XX89",
                    ),
                    Deposit(
                        date="2025-03-15",
                        description="DEPOT CHEQUE - CONTRAT PEINTURE",
                        amount=3500.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Laurentienne XX77",
                    ),
                    Deposit(
                        date="2025-03-22",
                        description="VIREMENT INTERAC - CLIENT DECK",
                        amount=2500.00,
                        category=DepositCategory.BUSINESS_INCOME,
                        account="Boréale XX89",
                    ),
                    # Inter-account transfer (detected)
                    Deposit(
                        date="2025-03-18",
                        description="TRANSFERT DEPUIS BOREALE",
                        amount=3000.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                        account="Laurentienne XX77",
                    ),
                ],
                withdrawals=[
                    Withdrawal(
                        date="2025-03-18",
                        description="VIREMENT VERS LAURENTIENNE",
                        amount=3000.00,
                        category="transfert",
                        account="Boréale XX89",
                    ),
                ],
            ),
        ],
        total_business_income=40000.00,
        total_deposits=49500.00,
        total_withdrawals=27000.00,
        months_covered=3,
        average_monthly_business_income=13333.33,
        annualized_business_income=160000.00,
        confidence_notes=[
            "Transfert inter-comptes détecté: retrait de 2 000 $ sur Boréale XX89 (2025-01-15) → dépôt de 2 000 $ sur Laurentienne XX77 (2025-01-15)",
            "Transfert inter-comptes détecté: retrait de 1 500 $ sur Laurentienne XX77 (2025-01-28) → dépôt de 1 500 $ sur Boréale XX89 (2025-01-28)",
            "Transfert inter-comptes détecté: retrait de 2 000 $ sur Boréale XX89 (2025-02-10) → dépôt de 2 000 $ sur Laurentienne XX77 (2025-02-10)",
            "Transfert inter-comptes détecté: retrait de 1 000 $ sur Laurentienne XX77 (2025-02-25) → dépôt de 1 000 $ sur Boréale XX89 (2025-02-25)",
            "Transfert inter-comptes détecté: retrait de 3 000 $ sur Boréale XX89 (2025-03-18) → dépôt de 3 000 $ sur Laurentienne XX77 (2025-03-18)",
        ],
    )


@pytest.fixture
def sample_extraction() -> BankStatementExtraction:
    """A realistic extraction result for testing."""
    return BankStatementExtraction(
        account_info=AccountInfo(
            account_holder="Jean Tremblay",
            institution="Desjardins",
            account_number_last4="4321",
            statement_period_start="2025-01-01",
            statement_period_end="2025-03-31",
        ),
        monthly_breakdown=[
            MonthlyBreakdown(
                month="2025-01",
                total_deposits=8500.00,
                business_deposits=6000.00,
                personal_transfers=2000.00,
                government_deposits=500.00,
                refund_deposits=0.00,
                loan_credit_deposits=0.00,
                other_deposits=0.00,
                total_withdrawals=4200.00,
                deposit_count=12,
                deposits=[
                    Deposit(
                        date="2025-01-05",
                        description="VIREMENT INTERAC - CLIENT ABC",
                        amount=3000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                    ),
                    Deposit(
                        date="2025-01-15",
                        description="VIREMENT INTERAC - CLIENT XYZ",
                        amount=3000.00,
                        category=DepositCategory.BUSINESS_INCOME,
                    ),
                    Deposit(
                        date="2025-01-20",
                        description="VIREMENT - COMPTE EPARGNE",
                        amount=2000.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                    ),
                    Deposit(
                        date="2025-01-25",
                        description="CREDIT TPS/TVH",
                        amount=500.00,
                        category=DepositCategory.GOVERNMENT,
                    ),
                ],
                withdrawals=[
                    Withdrawal(
                        date="2025-01-10",
                        description="LOYER BUREAU",
                        amount=1500.00,
                        category="rent",
                    ),
                ],
            ),
            MonthlyBreakdown(
                month="2025-02",
                total_deposits=7000.00,
                business_deposits=5500.00,
                personal_transfers=1000.00,
                government_deposits=500.00,
                refund_deposits=0.00,
                loan_credit_deposits=0.00,
                other_deposits=0.00,
                total_withdrawals=3800.00,
                deposit_count=10,
                deposits=[
                    Deposit(
                        date="2025-02-03",
                        description="DEPOT CHEQUE - CONTRAT MAINT",
                        amount=5500.00,
                        category=DepositCategory.BUSINESS_INCOME,
                    ),
                    Deposit(
                        date="2025-02-10",
                        description="VIREMENT CONJOINT",
                        amount=1000.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                    ),
                    Deposit(
                        date="2025-02-28",
                        description="CREDIT TPS/TVH",
                        amount=500.00,
                        category=DepositCategory.GOVERNMENT,
                    ),
                ],
                withdrawals=[
                    Withdrawal(
                        date="2025-02-01",
                        description="PAIEMENT HYPOTHEQUE",
                        amount=1200.00,
                        category="hypotheque",
                    ),
                    Withdrawal(
                        date="2025-02-15",
                        description="BELL MOBILITE",
                        amount=85.00,
                        category="telecom",
                    ),
                ],
            ),
            MonthlyBreakdown(
                month="2025-03",
                total_deposits=9200.00,
                business_deposits=7200.00,
                personal_transfers=1500.00,
                government_deposits=500.00,
                refund_deposits=0.00,
                loan_credit_deposits=0.00,
                other_deposits=0.00,
                total_withdrawals=5100.00,
                deposit_count=14,
                deposits=[
                    Deposit(
                        date="2025-03-01",
                        description="VIREMENT INTERAC - GROS CLIENT",
                        amount=7200.00,
                        category=DepositCategory.BUSINESS_INCOME,
                    ),
                    Deposit(
                        date="2025-03-15",
                        description="TRANSFERT COMPTE CONJOINT",
                        amount=1500.00,
                        category=DepositCategory.PERSONAL_TRANSFER,
                    ),
                    Deposit(
                        date="2025-03-20",
                        description="CREDIT TPS/TVH",
                        amount=500.00,
                        category=DepositCategory.GOVERNMENT,
                    ),
                ],
                withdrawals=[
                    Withdrawal(
                        date="2025-03-01",
                        description="PAIEMENT HYPOTHEQUE",
                        amount=1200.00,
                        category="hypotheque",
                    ),
                    Withdrawal(
                        date="2025-03-10",
                        description="ASSURANCE AUTO",
                        amount=150.00,
                        category="assurance",
                    ),
                ],
            ),
        ],
        total_business_income=18700.00,
        total_deposits=24700.00,
        total_withdrawals=13100.00,
        months_covered=3,
        average_monthly_business_income=6233.33,
        annualized_business_income=74800.00,
        confidence_notes=[
            "Le dépôt de 7 200 $ en mars est inhabituellement élevé — vérifier le contrat.",
            "Transferts personnels réguliers du conjoint — non inclus dans le revenu d'affaires.",
        ],
    )


@pytest.fixture
def dp_extraction() -> DPExtraction:
    """Realistic DPExtraction: 2 accounts, 3 months, transfers, cash deposits, 25k gift."""
    return DPExtraction(
        accounts=[
            DPAccountInfo(
                account_id="A1",
                institution="Desjardins",
                account_number_last4="4321",
                holder_name="Jean Tremblay",
                period_start="2025-01-01",
                period_end="2025-03-31",
                opening_balance=15000.00,
                closing_balance=42000.00,
            ),
            DPAccountInfo(
                account_id="A2",
                institution="Banque Nationale",
                account_number_last4="8765",
                holder_name="Jean Tremblay",
                period_start="2025-01-15",
                period_end="2025-03-31",
                opening_balance=8000.00,
                closing_balance=22000.00,
            ),
        ],
        transactions=[
            # A1 payroll deposits
            DPTransaction(id="A1-001", date="2025-01-15", description="PAIE EMPLOYEUR INC", amount=3500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1"),
            DPTransaction(id="A1-002", date="2025-01-31", description="PAIE EMPLOYEUR INC", amount=3500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1"),
            DPTransaction(id="A1-003", date="2025-02-15", description="PAIE EMPLOYEUR INC", amount=3500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1"),
            DPTransaction(id="A1-004", date="2025-02-28", description="PAIE EMPLOYEUR INC", amount=3500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1"),
            DPTransaction(id="A1-005", date="2025-03-15", description="PAIE EMPLOYEUR INC", amount=3500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1"),
            DPTransaction(id="A1-006", date="2025-03-31", description="PAIE EMPLOYEUR INC", amount=3500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1"),
            # A1 gift deposit (25k)
            DPTransaction(id="A1-007", date="2025-02-20", description="VIREMENT INTERAC DON PARENTS", amount=25000.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.GIFT, account_id="A1"),
            # A1 cash deposit
            DPTransaction(id="A1-008", date="2025-03-10", description="DEPOT COMPTANT GUICHET", amount=2500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.CASH, account_id="A1"),
            # A1 -> A2 transfer (withdrawal side)
            DPTransaction(id="A1-009", date="2025-02-05", description="VIREMENT VERS NATIONALE", amount=5000.00,
                          type=TransactionType.WITHDRAWAL, category=TransactionCategory.TRANSFER, account_id="A1"),
            # A1 bill payments
            DPTransaction(id="A1-010", date="2025-01-20", description="PAIEMENT HYPOTHEQUE", amount=1800.00,
                          type=TransactionType.WITHDRAWAL, category=TransactionCategory.BILL_PAYMENT, account_id="A1"),
            DPTransaction(id="A1-011", date="2025-02-20", description="PAIEMENT HYPOTHEQUE", amount=1800.00,
                          type=TransactionType.WITHDRAWAL, category=TransactionCategory.BILL_PAYMENT, account_id="A1"),
            # A2 deposits
            DPTransaction(id="A2-001", date="2025-02-06", description="TRANSFERT DEPUIS DESJARDINS", amount=5000.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.TRANSFER, account_id="A2"),
            DPTransaction(id="A2-002", date="2025-03-01", description="DEPOT CHEQUE", amount=8000.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.OTHER, account_id="A2"),
            # A2 large unknown deposit
            DPTransaction(id="A2-003", date="2025-03-20", description="DEPOT CHEQUE INCONNU", amount=10000.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.OTHER, account_id="A2"),
            # A2 withdrawal
            DPTransaction(id="A2-004", date="2025-03-05", description="ACHAT POS RONA", amount=450.00,
                          type=TransactionType.WITHDRAWAL, category=TransactionCategory.PURCHASE, account_id="A2"),
        ],
    )


@pytest.fixture
def dp_audit_result() -> DPAuditResult:
    """Complete DPAuditResult with DP target of 80k$."""
    return DPAuditResult(
        accounts=[
            DPAccountInfo(account_id="A1", institution="Desjardins",
                          holder_name="Jean Tremblay", period_start="2025-01-01",
                          period_end="2025-03-31", opening_balance=15000.00, closing_balance=42000.00),
            DPAccountInfo(account_id="A2", institution="Banque Nationale",
                          holder_name="Jean Tremblay", period_start="2025-01-15",
                          period_end="2025-03-31", opening_balance=8000.00, closing_balance=22000.00),
        ],
        transactions=[
            DPTransaction(id="A1-001", date="2025-01-15", description="PAIE EMPLOYEUR INC", amount=3500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.PAYROLL, account_id="A1"),
            DPTransaction(id="A1-007", date="2025-02-20", description="VIREMENT INTERAC DON PARENTS", amount=25000.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.GIFT, account_id="A1"),
            DPTransaction(id="A1-008", date="2025-03-10", description="DEPOT COMPTANT GUICHET", amount=2500.00,
                          type=TransactionType.DEPOSIT, category=TransactionCategory.CASH, account_id="A1"),
        ],
        transfers=[
            TransferMatch(from_account_id="A1", to_account_id="A2", amount=5000.00,
                          from_transaction_id="A1-009", to_transaction_id="A2-001",
                          date_delta_days=1, match_score=0.85),
        ],
        flags=[
            DPFlag(type=FlagType.LARGE_DEPOSIT, severity=FlagSeverity.CRITICAL,
                   rationale="Dépôt important de 25 000,00 $ le 2025-02-20",
                   supporting_transaction_ids=["A1-007"],
                   recommended_documents=["Preuve de provenance des fonds"]),
            DPFlag(type=FlagType.CASH_DEPOSIT, severity=FlagSeverity.WARNING,
                   rationale="Dépôt en espèces de 2 500,00 $ le 2025-03-10",
                   supporting_transaction_ids=["A1-008"],
                   recommended_documents=["Lettre explicative pour dépôt en espèces"]),
        ],
        client_requests=[
            ClientRequest(title="Lettre de don notariée", reason="Don de 25 000 $ détecté",
                          required_docs=["Lettre de don notariée"], supporting_transaction_ids=["A1-007"]),
        ],
        summary=DPSummary(
            dp_target=80000.00,
            dp_explained_amount=49500.00,
            unexplained_amount=30500.00,
            needs_review=True,
            review_notes=["2 drapeau(x) critique(s) détecté(s)", "30 500,00 $ de la mise de fonds non expliqué"],
            source_breakdown=SourceBreakdown(
                payroll=21000.00, gift=25000.00, investment_sale=0.00,
                property_sale=0.00, other_explained=3500.00, unexplained=30500.00,
            ),
        ),
        borrower_name="Jean Tremblay",
        closing_date="2025-06-15",
    )
