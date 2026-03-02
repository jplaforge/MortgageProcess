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
