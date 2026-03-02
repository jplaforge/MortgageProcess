"""Tests for downpayment analyzer (deterministic post-processing)."""

import pytest

from mortgage_mcp.models.downpayment import (
    DPAccountInfo,
    DPExtraction,
    DPFlag,
    DPTransaction,
    FlagSeverity,
    FlagType,
    SourceBreakdown,
    TransactionCategory,
    TransactionType,
    TransferMatch,
)
from mortgage_mcp.services.downpayment_analyzer import (
    analyze,
    build_summary,
    calculate_source_breakdown,
    detect_flags,
    generate_client_requests,
    match_transfers,
)


def _tx(id, date, desc, amount, tx_type, category, account_id):
    """Shorthand to create a DPTransaction."""
    return DPTransaction(
        id=id, date=date, description=desc, amount=amount,
        type=tx_type, category=category, account_id=account_id,
    )


# ── Transfer matching tests ──────────────────────────────────────────────


class TestMatchTransfers:
    def test_exact_amount_match(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT VERS A2", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-15", "TRANSFERT DEPUIS A1", 5000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
        ]
        matches = match_transfers(txns)
        assert len(matches) == 1
        assert matches[0].amount == 5000
        assert matches[0].from_account_id == "A1"
        assert matches[0].to_account_id == "A2"

    def test_amount_tolerance(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-15", "TRANSFERT", 5002, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
        ]
        matches = match_transfers(txns)
        assert len(matches) == 1

    def test_amount_outside_tolerance(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-15", "DEPOT", 5500, TransactionType.DEPOSIT, TransactionCategory.OTHER, "A2"),
        ]
        matches = match_transfers(txns)
        assert len(matches) == 0

    def test_date_window(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-17", "TRANSFERT", 5000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
        ]
        matches = match_transfers(txns)
        assert len(matches) == 1
        assert matches[0].date_delta_days == 2

    def test_date_outside_window(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-25", "TRANSFERT", 5000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
        ]
        matches = match_transfers(txns)
        assert len(matches) == 0

    def test_same_account_excluded(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A1-002", "2025-01-15", "TRANSFERT", 5000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A1"),
        ]
        matches = match_transfers(txns)
        assert len(matches) == 0

    def test_keyword_boost_score(self):
        # With keywords
        txns_keywords = [
            _tx("A1-001", "2025-01-15", "VIREMENT INTERAC VERS A2", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-15", "TRANSFERT DEPUIS A1", 5000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
        ]
        # Without keywords
        txns_no_keywords = [
            _tx("A1-001", "2025-01-15", "SORTIE FONDS", 5000, TransactionType.WITHDRAWAL, TransactionCategory.OTHER, "A1"),
            _tx("A2-001", "2025-01-15", "ENTREE FONDS", 5000, TransactionType.DEPOSIT, TransactionCategory.OTHER, "A2"),
        ]
        matches_kw = match_transfers(txns_keywords)
        matches_no = match_transfers(txns_no_keywords)
        assert len(matches_kw) == 1
        assert len(matches_no) == 1
        assert matches_kw[0].match_score > matches_no[0].match_score

    def test_greedy_one_to_one(self):
        """Two withdrawals, two deposits — each should match only once."""
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT 1", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A1-002", "2025-01-20", "VIREMENT 2", 3000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-15", "TRANSFERT 1", 5000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
            _tx("A2-002", "2025-01-20", "TRANSFERT 2", 3000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
        ]
        matches = match_transfers(txns)
        assert len(matches) == 2
        matched_wd_ids = {m.from_transaction_id for m in matches}
        matched_dep_ids = {m.to_transaction_id for m in matches}
        assert len(matched_wd_ids) == 2
        assert len(matched_dep_ids) == 2


# ── Flag detection tests ─────────────────────────────────────────────────


class TestDetectFlags:
    def test_large_deposit(self):
        txns = [
            _tx("A1-001", "2025-01-15", "GROS DEPOT", 10000, TransactionType.DEPOSIT, TransactionCategory.OTHER, "A1"),
        ]
        flags = detect_flags(txns, [], [], 80000)
        types = [f.type for f in flags]
        assert FlagType.LARGE_DEPOSIT in types

    def test_cash_deposit(self):
        txns = [
            _tx("A1-001", "2025-01-15", "DEPOT COMPTANT GUICHET", 2500, TransactionType.DEPOSIT, TransactionCategory.CASH, "A1"),
        ]
        flags = detect_flags(txns, [], [], 80000)
        cash_flags = [f for f in flags if f.type == FlagType.CASH_DEPOSIT]
        assert len(cash_flags) == 1
        assert cash_flags[0].severity == FlagSeverity.WARNING

    def test_cash_deposit_critical_above_10k(self):
        txns = [
            _tx("A1-001", "2025-01-15", "CASH DEPOSIT", 15000, TransactionType.DEPOSIT, TransactionCategory.CASH, "A1"),
        ]
        flags = detect_flags(txns, [], [], 80000)
        cash_flags = [f for f in flags if f.type == FlagType.CASH_DEPOSIT]
        assert len(cash_flags) >= 1
        assert cash_flags[0].severity == FlagSeverity.CRITICAL

    def test_round_amount(self):
        txns = [
            _tx("A1-001", "2025-01-15", "DEPOT", 10000, TransactionType.DEPOSIT, TransactionCategory.BUSINESS_INCOME, "A1"),
        ]
        flags = detect_flags(txns, [], [], 80000)
        round_flags = [f for f in flags if f.type == FlagType.ROUND_AMOUNT]
        assert len(round_flags) == 1
        assert round_flags[0].severity == FlagSeverity.INFO

    def test_rapid_succession(self):
        txns = [
            _tx("A1-001", "2025-01-15", "DEPOT 1", 5000, TransactionType.DEPOSIT, TransactionCategory.OTHER, "A1"),
            _tx("A1-002", "2025-01-16", "DEPOT 2", 4000, TransactionType.DEPOSIT, TransactionCategory.OTHER, "A1"),
        ]
        flags = detect_flags(txns, [], [], 80000)
        rapid_flags = [f for f in flags if f.type == FlagType.RAPID_SUCCESSION]
        assert len(rapid_flags) >= 1

    def test_period_gap(self):
        accts = [
            DPAccountInfo(account_id="A1", institution="Desjardins",
                          period_start="2025-01-01", period_end="2025-02-15"),
        ]
        flags = detect_flags([], [], accts, 80000)
        gap_flags = [f for f in flags if f.type == FlagType.PERIOD_GAP]
        assert len(gap_flags) == 1

    def test_no_period_gap_with_90_days(self):
        accts = [
            DPAccountInfo(account_id="A1", institution="Desjardins",
                          period_start="2025-01-01", period_end="2025-04-01"),
        ]
        flags = detect_flags([], [], accts, 80000)
        gap_flags = [f for f in flags if f.type == FlagType.PERIOD_GAP]
        assert len(gap_flags) == 0

    def test_unexplained_source(self):
        txns = [
            _tx("A1-001", "2025-01-15", "DEPOT INCONNU", 8000, TransactionType.DEPOSIT, TransactionCategory.OTHER, "A1"),
        ]
        flags = detect_flags(txns, [], [], 80000)
        unexplained = [f for f in flags if f.type == FlagType.UNEXPLAINED_SOURCE]
        assert len(unexplained) == 1
        assert unexplained[0].severity == FlagSeverity.CRITICAL

    def test_transfer_deposits_not_flagged_as_large(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT", 10000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-15", "TRANSFERT", 10000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
        ]
        transfers = [TransferMatch(
            from_account_id="A1", to_account_id="A2", amount=10000,
            from_transaction_id="A1-001", to_transaction_id="A2-001",
            date_delta_days=0, match_score=0.9,
        )]
        flags = detect_flags(txns, transfers, [], 80000)
        large_flags = [f for f in flags if f.type == FlagType.LARGE_DEPOSIT]
        flagged_ids = [tid for f in large_flags for tid in f.supporting_transaction_ids]
        assert "A2-001" not in flagged_ids

    def test_multi_hop_transfer(self):
        transfers = [
            TransferMatch(from_account_id="A1", to_account_id="A2", amount=5000,
                          from_transaction_id="A1-001", to_transaction_id="A2-001",
                          date_delta_days=0, match_score=0.9),
            TransferMatch(from_account_id="A2", to_account_id="A3", amount=5000,
                          from_transaction_id="A2-002", to_transaction_id="A3-001",
                          date_delta_days=1, match_score=0.8),
        ]
        flags = detect_flags([], transfers, [], 80000)
        hop_flags = [f for f in flags if f.type == FlagType.MULTI_HOP_TRANSFER]
        assert len(hop_flags) >= 1

    def test_non_payroll_recurring(self):
        txns = [
            _tx("A1-001", "2025-01-15", "DEPOT CLIENT X", 3000, TransactionType.DEPOSIT, TransactionCategory.BUSINESS_INCOME, "A1"),
            _tx("A1-002", "2025-02-15", "DEPOT CLIENT X", 3050, TransactionType.DEPOSIT, TransactionCategory.BUSINESS_INCOME, "A1"),
        ]
        flags = detect_flags(txns, [], [], 80000)
        recurring = [f for f in flags if f.type == FlagType.NON_PAYROLL_RECURRING]
        assert len(recurring) >= 1


# ── Source breakdown tests ───────────────────────────────────────────────


class TestSourceBreakdown:
    def test_category_mapping(self):
        txns = [
            _tx("A1-001", "2025-01-15", "PAIE", 7000, TransactionType.DEPOSIT, TransactionCategory.PAYROLL, "A1"),
            _tx("A1-002", "2025-02-20", "DON", 25000, TransactionType.DEPOSIT, TransactionCategory.GIFT, "A1"),
            _tx("A1-003", "2025-03-01", "PLACEMENT", 10000, TransactionType.DEPOSIT, TransactionCategory.INVESTMENT, "A1"),
        ]
        sb = calculate_source_breakdown(txns, [], 80000)
        assert sb.payroll == 7000
        assert sb.gift == 25000
        assert sb.investment_sale == 10000

    def test_unexplained_residual(self):
        txns = [
            _tx("A1-001", "2025-01-15", "PAIE", 5000, TransactionType.DEPOSIT, TransactionCategory.PAYROLL, "A1"),
        ]
        sb = calculate_source_breakdown(txns, [], 80000)
        assert sb.unexplained == 75000

    def test_transfers_excluded(self):
        txns = [
            _tx("A1-001", "2025-01-15", "VIREMENT", 5000, TransactionType.WITHDRAWAL, TransactionCategory.TRANSFER, "A1"),
            _tx("A2-001", "2025-01-15", "TRANSFERT", 5000, TransactionType.DEPOSIT, TransactionCategory.TRANSFER, "A2"),
            _tx("A1-002", "2025-01-20", "PAIE", 3000, TransactionType.DEPOSIT, TransactionCategory.PAYROLL, "A1"),
        ]
        transfers = [TransferMatch(
            from_account_id="A1", to_account_id="A2", amount=5000,
            from_transaction_id="A1-001", to_transaction_id="A2-001",
            date_delta_days=0, match_score=0.9,
        )]
        sb = calculate_source_breakdown(txns, transfers, 80000)
        assert sb.payroll == 3000
        # Transfer deposit should not be counted
        assert sb.unexplained == 77000

    def test_withdrawals_ignored(self):
        txns = [
            _tx("A1-001", "2025-01-15", "PAIE", 3000, TransactionType.DEPOSIT, TransactionCategory.PAYROLL, "A1"),
            _tx("A1-002", "2025-01-20", "ACHAT", 1000, TransactionType.WITHDRAWAL, TransactionCategory.PURCHASE, "A1"),
        ]
        sb = calculate_source_breakdown(txns, [], 10000)
        assert sb.payroll == 3000
        assert sb.unexplained == 7000

    def test_business_income_as_other_explained(self):
        txns = [
            _tx("A1-001", "2025-01-15", "CLIENT X", 5000, TransactionType.DEPOSIT, TransactionCategory.BUSINESS_INCOME, "A1"),
        ]
        sb = calculate_source_breakdown(txns, [], 10000)
        assert sb.other_explained == 5000


# ── Client requests tests ────────────────────────────────────────────────


class TestClientRequests:
    def test_cash_deposit_request(self):
        flags = [
            DPFlag(type=FlagType.CASH_DEPOSIT, severity=FlagSeverity.WARNING,
                   rationale="Dépôt en espèces", supporting_transaction_ids=["A1-001"]),
        ]
        reqs = generate_client_requests(flags, [])
        assert len(reqs) >= 1
        assert any("espèces" in r.title.lower() for r in reqs)

    def test_large_deposit_request(self):
        flags = [
            DPFlag(type=FlagType.LARGE_DEPOSIT, severity=FlagSeverity.CRITICAL,
                   rationale="Gros dépôt", supporting_transaction_ids=["A1-001"]),
        ]
        reqs = generate_client_requests(flags, [])
        assert any("provenance" in r.title.lower() for r in reqs)

    def test_gift_transaction_request(self):
        txns = [
            _tx("A1-001", "2025-01-15", "DON", 25000, TransactionType.DEPOSIT, TransactionCategory.GIFT, "A1"),
        ]
        reqs = generate_client_requests([], txns)
        assert any("don" in r.title.lower() for r in reqs)

    def test_period_gap_request(self):
        flags = [
            DPFlag(type=FlagType.PERIOD_GAP, severity=FlagSeverity.WARNING,
                   rationale="Couverture insuffisante"),
        ]
        reqs = generate_client_requests(flags, [])
        assert any("complémentaires" in r.title.lower() for r in reqs)

    def test_references_transaction_ids(self):
        flags = [
            DPFlag(type=FlagType.CASH_DEPOSIT, severity=FlagSeverity.WARNING,
                   rationale="Cash", supporting_transaction_ids=["A1-008"]),
        ]
        reqs = generate_client_requests(flags, [])
        assert reqs[0].supporting_transaction_ids == ["A1-008"]


# ── Summary tests ────────────────────────────────────────────────────────


class TestBuildSummary:
    def test_needs_review_with_critical_flags(self):
        sb = SourceBreakdown(payroll=50000, unexplained=30000)
        flags = [DPFlag(type=FlagType.LARGE_DEPOSIT, severity=FlagSeverity.CRITICAL, rationale="test")]
        summary = build_summary(80000, sb, flags)
        assert summary.needs_review is True
        assert summary.dp_target == 80000

    def test_no_review_when_clean(self):
        sb = SourceBreakdown(payroll=80000, unexplained=0)
        flags = [DPFlag(type=FlagType.ROUND_AMOUNT, severity=FlagSeverity.INFO, rationale="info")]
        summary = build_summary(80000, sb, flags)
        assert summary.needs_review is False

    def test_review_notes_count(self):
        sb = SourceBreakdown(payroll=50000, unexplained=30000)
        flags = [
            DPFlag(type=FlagType.LARGE_DEPOSIT, severity=FlagSeverity.CRITICAL, rationale="crit"),
            DPFlag(type=FlagType.CASH_DEPOSIT, severity=FlagSeverity.WARNING, rationale="warn"),
        ]
        summary = build_summary(80000, sb, flags)
        # Should have notes for critical, warning, and unexplained
        assert len(summary.review_notes) >= 2


# ── Full pipeline test ───────────────────────────────────────────────────


class TestAnalyze:
    def test_full_pipeline(self, dp_extraction):
        result = analyze(
            dp_extraction,
            target_downpayment=80000,
            closing_date="2025-06-15",
            borrower_name="Jean Tremblay",
        )
        assert result.borrower_name == "Jean Tremblay"
        assert result.closing_date == "2025-06-15"
        assert len(result.accounts) == 2
        assert len(result.transactions) > 0
        assert len(result.transfers) >= 1  # A1-009 -> A2-001
        assert len(result.flags) > 0
        assert result.summary.dp_target == 80000

    def test_pipeline_detects_transfer(self, dp_extraction):
        result = analyze(dp_extraction, 80000, "2025-06-15", "Jean Tremblay")
        transfer_pairs = [(t.from_transaction_id, t.to_transaction_id) for t in result.transfers]
        assert ("A1-009", "A2-001") in transfer_pairs

    def test_pipeline_flags_cash(self, dp_extraction):
        result = analyze(dp_extraction, 80000, "2025-06-15", "Jean Tremblay")
        cash_flags = [f for f in result.flags if f.type == FlagType.CASH_DEPOSIT]
        assert len(cash_flags) >= 1

    def test_pipeline_flags_gift(self, dp_extraction):
        result = analyze(dp_extraction, 80000, "2025-06-15", "Jean Tremblay")
        # Gift should generate client request
        gift_reqs = [r for r in result.client_requests if "don" in r.title.lower()]
        assert len(gift_reqs) >= 1
