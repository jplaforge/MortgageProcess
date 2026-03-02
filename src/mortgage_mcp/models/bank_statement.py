"""Pydantic models for bank statement extraction.

These models serve double duty:
1. Define the structured output schema for Gemini extraction
2. Type the data flowing through the MCP tool pipeline
"""

from enum import Enum

from pydantic import BaseModel, Field


class DepositCategory(str, Enum):
    BUSINESS_INCOME = "business_income"
    PERSONAL_TRANSFER = "personal_transfer"
    GOVERNMENT = "government"
    LOAN_CREDIT = "loan_credit"
    REFUND = "refund"
    OTHER = "other"


class Deposit(BaseModel):
    date: str = Field(description="Date of the deposit (YYYY-MM-DD)")
    description: str = Field(description="Transaction description as shown on statement")
    amount: float = Field(description="Deposit amount in CAD")
    category: DepositCategory = Field(description="Categorization of the deposit source")


class Withdrawal(BaseModel):
    date: str = Field(description="Date of the withdrawal (YYYY-MM-DD)")
    description: str = Field(description="Transaction description as shown on statement")
    amount: float = Field(description="Withdrawal amount in CAD (positive number)")
    category: str = Field(description="Expense category (e.g. rent, supplies, telecom)")


class MonthlyBreakdown(BaseModel):
    month: str = Field(description="Month in YYYY-MM format")
    total_deposits: float = Field(description="Sum of all deposits for the month")
    business_deposits: float = Field(description="Sum of deposits categorized as business income")
    personal_transfers: float = Field(description="Sum of personal/inter-account transfers")
    government_deposits: float = Field(description="Sum of government payments (GST/HST credits, etc.)")
    refund_deposits: float = Field(description="Sum of deposits categorized as refunds")
    loan_credit_deposits: float = Field(description="Sum of deposits categorized as loans/credit")
    other_deposits: float = Field(description="Sum of deposits not fitting any other category")
    total_withdrawals: float = Field(description="Sum of all withdrawals for the month")
    deposit_count: int = Field(description="Number of deposits in the month")
    deposits: list[Deposit] = Field(default_factory=list, description="Individual deposits")
    withdrawals: list[Withdrawal] = Field(default_factory=list, description="Individual withdrawals")


class AccountInfo(BaseModel):
    account_holder: str = Field(description="Name of the account holder")
    institution: str = Field(description="Financial institution name")
    account_number_last4: str = Field(default="", description="Last 4 digits of account number if visible")
    statement_period_start: str = Field(description="Start date of the statement period (YYYY-MM-DD)")
    statement_period_end: str = Field(description="End date of the statement period (YYYY-MM-DD)")


class BankStatementExtraction(BaseModel):
    """Full extraction result from bank statement analysis.

    Used as both the Gemini response_schema and the internal data model.
    """

    account_info: AccountInfo = Field(description="Account and statement metadata")
    monthly_breakdown: list[MonthlyBreakdown] = Field(description="Per-month financial breakdown")
    total_business_income: float = Field(description="Total business income across all months")
    total_deposits: float = Field(description="Total of all deposits across all months")
    total_withdrawals: float = Field(description="Total of all withdrawals across all months")
    months_covered: int = Field(description="Number of months covered by the statements")
    average_monthly_business_income: float = Field(
        description="Average monthly business income (total_business_income / months_covered)"
    )
    annualized_business_income: float = Field(
        description="Projected annual business income (average_monthly * 12)"
    )
    confidence_notes: list[str] = Field(
        default_factory=list,
        description="Notes about data quality, assumptions, or flags for broker review",
    )
