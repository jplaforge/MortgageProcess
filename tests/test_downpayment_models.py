"""Tests for downpayment Pydantic models."""

import json

import pytest

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


class TestTransactionEnums:
    def test_transaction_type_values(self):
        assert TransactionType.DEPOSIT == "deposit"
        assert TransactionType.WITHDRAWAL == "withdrawal"

    def test_transaction_category_values(self):
        assert TransactionCategory.PAYROLL == "payroll"
        assert TransactionCategory.CASH == "cash"
        assert TransactionCategory.GIFT == "gift"
        assert TransactionCategory.OTHER == "other"

    def test_flag_severity_values(self):
        assert FlagSeverity.CRITICAL == "critical"
        assert FlagSeverity.WARNING == "warning"
        assert FlagSeverity.INFO == "info"

    def test_flag_type_values(self):
        assert FlagType.LARGE_DEPOSIT == "large_deposit"
        assert FlagType.CASH_DEPOSIT == "cash_deposit"
        assert FlagType.UNEXPLAINED_SOURCE == "unexplained_source"


class TestDPTransaction:
    def test_create_deposit(self):
        t = DPTransaction(
            id="A1-001", date="2025-01-15", description="PAIE",
            amount=3500.00, type=TransactionType.DEPOSIT,
            category=TransactionCategory.PAYROLL, account_id="A1",
        )
        assert t.id == "A1-001"
        assert t.type == TransactionType.DEPOSIT
        assert t.amount == 3500.00

    def test_defaults(self):
        t = DPTransaction(
            id="A1-001", date="2025-01-15", description="TEST",
            amount=100.0, type=TransactionType.DEPOSIT,
            category=TransactionCategory.OTHER, account_id="A1",
        )
        assert t.page_source == 0
        assert t.confidence == 1.0
        assert t.normalized_description == ""
        assert t.merchant_guess == ""


class TestDPExtraction:
    def test_create_extraction(self, dp_extraction):
        assert len(dp_extraction.accounts) == 2
        assert len(dp_extraction.transactions) > 0

    def test_serialization_roundtrip(self, dp_extraction):
        json_str = dp_extraction.model_dump_json()
        restored = DPExtraction.model_validate_json(json_str)
        assert len(restored.accounts) == len(dp_extraction.accounts)
        assert len(restored.transactions) == len(dp_extraction.transactions)
        assert restored.transactions[0].id == dp_extraction.transactions[0].id


class TestDPAuditResult:
    def test_create_audit_result(self, dp_audit_result):
        assert dp_audit_result.borrower_name == "Jean Tremblay"
        assert dp_audit_result.summary.dp_target == 80000.00
        assert len(dp_audit_result.flags) == 2

    def test_serialization_roundtrip(self, dp_audit_result):
        json_str = dp_audit_result.model_dump_json()
        data = json.loads(json_str)
        assert data["borrower_name"] == "Jean Tremblay"
        assert data["summary"]["dp_target"] == 80000.00
        restored = DPAuditResult.model_validate_json(json_str)
        assert restored.summary.dp_target == dp_audit_result.summary.dp_target

    def test_source_breakdown_fields(self):
        sb = SourceBreakdown(payroll=21000, gift=25000, unexplained=34000)
        assert sb.savings == 0.0
        assert sb.property_sale == 0.0
        assert sb.payroll == 21000

    def test_transfer_match_fields(self):
        tm = TransferMatch(
            from_account_id="A1", to_account_id="A2", amount=5000,
            from_transaction_id="A1-009", to_transaction_id="A2-001",
            date_delta_days=1, match_score=0.85,
        )
        assert tm.match_score == 0.85
        assert tm.date_delta_days == 1
