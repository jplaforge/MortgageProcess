"""Deterministic post-processing for downpayment audit.

Takes a DPExtraction (Gemini output) + deal parameters and produces
a DPAuditResult with transfer matching, flag detection, source
breakdown, and client requests.
"""

from collections import defaultdict
from datetime import datetime, timedelta

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

# ── Configurable constants ────────────────────────────────────────────────

TRANSFER_AMOUNT_TOLERANCE = 0.005  # 0.5%
TRANSFER_AMOUNT_ABS_TOLERANCE = 1.0  # ±1$
TRANSFER_DATE_WINDOW_DAYS = 3
SPLIT_TRANSFER_AMOUNT_TOLERANCE = 0.02  # 2% for split transfers (looser)
SPLIT_TRANSFER_DATE_WINDOW_DAYS = 5
LARGE_DEPOSIT_ABS_THRESHOLD = 5_000
LARGE_DEPOSIT_RELATIVE_FACTOR = 0.25  # 25% of monthly average
CASH_KEYWORDS = {"CASH", "COMPTANT", "GUICHET", "ATM", "DEPOT ESPECES", "ESPECES", "DEPOT COMPTANT"}
TRANSFER_KEYWORDS = {"VIREMENT", "TRANSFERT", "INTERAC", "TFR", "TRANSFER", "VIR"}
ROUND_AMOUNTS = {5_000, 10_000, 25_000, 50_000, 100_000}
RAPID_SUCCESSION_THRESHOLD = 3_000
RAPID_SUCCESSION_HOURS = 48
MINIMUM_COVERAGE_DAYS = 90


def _parse_date(date_str: str) -> datetime:
    return datetime.strptime(date_str, "%Y-%m-%d")


def _has_keywords(description: str, keywords: set[str]) -> bool:
    desc_upper = description.upper()
    return any(kw in desc_upper for kw in keywords)


def _all_transfer_tx_ids(transfers: list[TransferMatch]) -> set[str]:
    """Collect all transaction IDs involved in matched transfers (1:1 and 1:N)."""
    ids: set[str] = set()
    for m in transfers:
        ids.add(m.from_transaction_id)
        if m.to_transaction_id:
            ids.add(m.to_transaction_id)
        ids.update(m.to_transaction_ids)
    return ids


# ── 1. Transfer matching ─────────────────────────────────────────────────


def match_transfers(transactions: list[DPTransaction]) -> list[TransferMatch]:
    """Match inter-account transfer pairs (withdrawal → deposit).

    Two passes:
    1. Greedy one-to-one matching by descending score.
    2. Split transfer matching: unmatched withdrawals → multiple deposits summing to ~same amount.
    """
    withdrawals = [t for t in transactions if t.type == TransactionType.WITHDRAWAL]
    deposits = [t for t in transactions if t.type == TransactionType.DEPOSIT]

    # ── Pass 1: 1:1 matching ──────────────────────────────────────────────

    candidates: list[tuple[float, DPTransaction, DPTransaction]] = []

    for wd in withdrawals:
        for dep in deposits:
            if wd.account_id == dep.account_id:
                continue
            if wd.amount == 0:
                continue
            amount_diff = abs(wd.amount - dep.amount)
            relative_diff = amount_diff / wd.amount
            if relative_diff > TRANSFER_AMOUNT_TOLERANCE and amount_diff > TRANSFER_AMOUNT_ABS_TOLERANCE:
                continue
            try:
                wd_date = _parse_date(wd.date)
                dep_date = _parse_date(dep.date)
            except ValueError:
                continue
            date_delta = abs((dep_date - wd_date).days)
            if date_delta > TRANSFER_DATE_WINDOW_DAYS:
                continue

            amount_score = max(0, 0.5 - (relative_diff * 10))
            date_score = 0.3 * (1 - date_delta / (TRANSFER_DATE_WINDOW_DAYS + 1))
            keyword_score = 0.0
            if _has_keywords(wd.description, TRANSFER_KEYWORDS) or _has_keywords(dep.description, TRANSFER_KEYWORDS):
                keyword_score = 0.2
            score = amount_score + date_score + keyword_score
            candidates.append((score, wd, dep))

    candidates.sort(key=lambda x: x[0], reverse=True)

    used_wd: set[str] = set()
    used_dep: set[str] = set()
    matches: list[TransferMatch] = []

    for score, wd, dep in candidates:
        if wd.id in used_wd or dep.id in used_dep:
            continue
        used_wd.add(wd.id)
        used_dep.add(dep.id)
        try:
            date_delta = abs((_parse_date(dep.date) - _parse_date(wd.date)).days)
        except ValueError:
            date_delta = 0
        matches.append(TransferMatch(
            from_account_id=wd.account_id,
            to_account_id=dep.account_id,
            amount=wd.amount,
            from_transaction_id=wd.id,
            to_transaction_id=dep.id,
            to_transaction_ids=[dep.id],
            date_delta_days=date_delta,
            match_score=round(score, 3),
        ))

    # ── Pass 2: 1:N split transfer matching ───────────────────────────────

    unmatched_wd = [w for w in withdrawals if w.id not in used_wd and w.amount >= 1000]
    available_dep = [d for d in deposits if d.id not in used_dep]

    for wd in unmatched_wd:
        if wd.amount == 0:
            continue
        try:
            wd_date = _parse_date(wd.date)
        except ValueError:
            continue

        # Only consider deposits on different accounts, within date window, with transfer keywords
        dep_candidates = []
        for dep in available_dep:
            if dep.account_id == wd.account_id:
                continue
            try:
                dep_date = _parse_date(dep.date)
            except ValueError:
                continue
            date_delta = abs((dep_date - wd_date).days)
            if date_delta > SPLIT_TRANSFER_DATE_WINDOW_DAYS:
                continue
            if not (_has_keywords(wd.description, TRANSFER_KEYWORDS) or _has_keywords(dep.description, TRANSFER_KEYWORDS)):
                continue
            dep_candidates.append((dep, date_delta))

        if not dep_candidates:
            continue

        # Try subset-sum: find combination of deposits that sum to ~withdrawal amount
        # Sort by amount descending for greedy approach
        dep_candidates.sort(key=lambda x: x[0].amount, reverse=True)
        selected: list[tuple[DPTransaction, int]] = []
        remaining = wd.amount

        for dep, delta in dep_candidates:
            if dep.id in used_dep:
                continue
            if dep.amount <= remaining * (1 + SPLIT_TRANSFER_AMOUNT_TOLERANCE):
                selected.append((dep, delta))
                remaining -= dep.amount
                if abs(remaining) <= wd.amount * SPLIT_TRANSFER_AMOUNT_TOLERANCE or remaining <= TRANSFER_AMOUNT_ABS_TOLERANCE:
                    break

        if not selected:
            continue

        total_selected = sum(d.amount for d, _ in selected)
        amount_diff = abs(total_selected - wd.amount)
        relative_diff = amount_diff / wd.amount if wd.amount > 0 else 1.0

        if relative_diff <= SPLIT_TRANSFER_AMOUNT_TOLERANCE or amount_diff <= TRANSFER_AMOUNT_ABS_TOLERANCE:
            max_delta = max(delta for _, delta in selected)
            # Score: lower than 1:1 matches since split is less certain
            score = 0.3 + (0.2 if relative_diff < 0.005 else 0.1) + (0.1 if max_delta <= 2 else 0.0)

            dep_ids = [d.id for d, _ in selected]
            to_accounts = {d.account_id for d, _ in selected}

            used_wd.add(wd.id)
            for d, _ in selected:
                used_dep.add(d.id)

            matches.append(TransferMatch(
                from_account_id=wd.account_id,
                to_account_id=", ".join(sorted(to_accounts)),
                amount=wd.amount,
                from_transaction_id=wd.id,
                to_transaction_id=dep_ids[0],
                to_transaction_ids=dep_ids,
                date_delta_days=max_delta,
                match_score=round(score, 3),
                is_split=True,
            ))

    return matches


# ── 2. Flag detection ────────────────────────────────────────────────────


def detect_flags(
    transactions: list[DPTransaction],
    transfers: list[TransferMatch],
    accounts: list[DPAccountInfo],
    target_dp: float,
) -> list[DPFlag]:
    """Detect audit flags on transactions."""
    flags: list[DPFlag] = []

    deposits = [t for t in transactions if t.type == TransactionType.DEPOSIT]
    transfer_tx_ids = _all_transfer_tx_ids(transfers)

    # Monthly average for relative thresholds
    monthly_totals: dict[str, float] = defaultdict(float)
    for dep in deposits:
        month = dep.date[:7] if len(dep.date) >= 7 else ""
        monthly_totals[month] += dep.amount
    avg_monthly = sum(monthly_totals.values()) / max(len(monthly_totals), 1)

    relative_threshold = avg_monthly * LARGE_DEPOSIT_RELATIVE_FACTOR

    for dep in deposits:
        # Skip transfer deposits for most flags
        is_transfer = dep.id in transfer_tx_ids

        # Large deposit
        if dep.amount > LARGE_DEPOSIT_ABS_THRESHOLD or dep.amount > relative_threshold:
            if not is_transfer:
                severity = FlagSeverity.CRITICAL if dep.amount > target_dp * 0.25 else FlagSeverity.WARNING
                flags.append(DPFlag(
                    type=FlagType.LARGE_DEPOSIT,
                    severity=severity,
                    rationale=f"Dépôt important de {dep.amount:,.2f} $ le {dep.date} — {dep.description}",
                    supporting_transaction_ids=[dep.id],
                    recommended_documents=["Preuve de provenance des fonds", "Relevé de compte source"],
                ))

        # Cash deposit
        if _has_keywords(dep.description, CASH_KEYWORDS) or dep.category == TransactionCategory.CASH:
            severity = FlagSeverity.CRITICAL if dep.amount >= 10_000 else FlagSeverity.WARNING
            flags.append(DPFlag(
                type=FlagType.CASH_DEPOSIT,
                severity=severity,
                rationale=f"Dépôt en espèces de {dep.amount:,.2f} $ le {dep.date} — {dep.description}",
                supporting_transaction_ids=[dep.id],
                recommended_documents=["Lettre explicative pour dépôt en espèces"],
            ))

        # Round amount
        if dep.amount in ROUND_AMOUNTS and not is_transfer:
            flags.append(DPFlag(
                type=FlagType.ROUND_AMOUNT,
                severity=FlagSeverity.INFO,
                rationale=f"Montant rond de {dep.amount:,.2f} $ le {dep.date} — {dep.description}",
                supporting_transaction_ids=[dep.id],
                recommended_documents=[],
            ))

    # Rapid succession: multiple large deposits (>3k) within 48h
    large_deposits = sorted(
        [d for d in deposits if d.amount > RAPID_SUCCESSION_THRESHOLD and d.id not in transfer_tx_ids],
        key=lambda d: d.date,
    )
    for i, dep in enumerate(large_deposits):
        group = [dep]
        try:
            dep_date = _parse_date(dep.date)
        except ValueError:
            continue
        for j in range(i + 1, len(large_deposits)):
            try:
                other_date = _parse_date(large_deposits[j].date)
            except ValueError:
                continue
            if (other_date - dep_date) <= timedelta(hours=RAPID_SUCCESSION_HOURS):
                group.append(large_deposits[j])
            else:
                break
        if len(group) >= 2:
            group_ids = [g.id for g in group]
            # Avoid duplicate flags for same group
            already_flagged = any(
                f.type == FlagType.RAPID_SUCCESSION and set(f.supporting_transaction_ids) == set(group_ids)
                for f in flags
            )
            if not already_flagged:
                total = sum(g.amount for g in group)
                flags.append(DPFlag(
                    type=FlagType.RAPID_SUCCESSION,
                    severity=FlagSeverity.WARNING,
                    rationale=f"{len(group)} dépôts importants totalisant {total:,.2f} $ en moins de 48h",
                    supporting_transaction_ids=group_ids,
                    recommended_documents=["Preuve de provenance des fonds"],
                ))

    # Non-payroll recurring (same amount within ±5%, appears 2+ times, not payroll)
    non_payroll_deposits = [d for d in deposits if d.category != TransactionCategory.PAYROLL and d.id not in transfer_tx_ids]
    amount_groups: dict[int, list[DPTransaction]] = defaultdict(list)
    for dep in non_payroll_deposits:
        bucket = round(dep.amount / 100) * 100  # bucket by nearest 100
        amount_groups[bucket].append(dep)
    for bucket, group in amount_groups.items():
        if len(group) >= 2:
            flags.append(DPFlag(
                type=FlagType.NON_PAYROLL_RECURRING,
                severity=FlagSeverity.INFO,
                rationale=f"Dépôt récurrent non-salarial: ~{bucket:,.0f} $ apparaît {len(group)} fois",
                supporting_transaction_ids=[g.id for g in group],
                recommended_documents=[],
            ))

    # Multi-hop transfer: detect chains A→B→C in transfer graph
    # Build adjacency: account -> list of (to_account, transfer)
    adjacency: dict[str, list[tuple[str, TransferMatch]]] = defaultdict(list)
    for tm in transfers:
        adjacency[tm.from_account_id].append((tm.to_account_id, tm))
    for src, edges in adjacency.items():
        for to_acct, tm1 in edges:
            if to_acct in adjacency:
                for final_acct, tm2 in adjacency[to_acct]:
                    if final_acct != src:
                        flags.append(DPFlag(
                            type=FlagType.MULTI_HOP_TRANSFER,
                            severity=FlagSeverity.WARNING,
                            rationale=f"Chaîne de transfert détectée: {src} → {to_acct} → {final_acct}",
                            supporting_transaction_ids=[
                                tm1.from_transaction_id, tm1.to_transaction_id,
                                tm2.from_transaction_id, tm2.to_transaction_id,
                            ],
                            recommended_documents=["Relevés complets de tous les comptes intermédiaires"],
                        ))

    # Period gap: check coverage
    if accounts:
        all_starts: list[datetime] = []
        all_ends: list[datetime] = []
        for acct in accounts:
            if acct.period_start:
                try:
                    all_starts.append(_parse_date(acct.period_start))
                except ValueError:
                    pass
            if acct.period_end:
                try:
                    all_ends.append(_parse_date(acct.period_end))
                except ValueError:
                    pass
        if all_starts and all_ends:
            earliest = min(all_starts)
            latest = max(all_ends)
            coverage_days = (latest - earliest).days
            if coverage_days < MINIMUM_COVERAGE_DAYS:
                flags.append(DPFlag(
                    type=FlagType.PERIOD_GAP,
                    severity=FlagSeverity.WARNING,
                    rationale=f"Couverture de {coverage_days} jours seulement (minimum recommandé: {MINIMUM_COVERAGE_DAYS} jours)",
                    supporting_transaction_ids=[],
                    recommended_documents=["Relevés bancaires couvrant au moins 90 jours"],
                ))

    # Unexplained source: large deposits with category "other" or unmatched "transfer" category
    for dep in deposits:
        if dep.id in transfer_tx_ids:
            continue
        if dep.amount < LARGE_DEPOSIT_ABS_THRESHOLD:
            continue
        if dep.category == TransactionCategory.OTHER:
            flags.append(DPFlag(
                type=FlagType.UNEXPLAINED_SOURCE,
                severity=FlagSeverity.CRITICAL,
                rationale=f"Dépôt de {dep.amount:,.2f} $ le {dep.date} sans source identifiable — {dep.description}",
                supporting_transaction_ids=[dep.id],
                recommended_documents=["Preuve de provenance des fonds", "Lettre explicative"],
            ))
        elif dep.category == TransactionCategory.TRANSFER:
            # Transfer-category deposit not matched as inter-account → needs verification
            flags.append(DPFlag(
                type=FlagType.UNEXPLAINED_SOURCE,
                severity=FlagSeverity.WARNING,
                rationale=f"Dépôt catégorisé comme transfert ({dep.amount:,.2f} $) le {dep.date} "
                          f"sans retrait correspondant identifié — {dep.description}",
                supporting_transaction_ids=[dep.id],
                recommended_documents=["Relevé du compte source montrant le retrait"],
            ))

    return flags


# ── 3. Source breakdown ──────────────────────────────────────────────────


def calculate_source_breakdown(
    transactions: list[DPTransaction],
    transfers: list[TransferMatch],
    target_dp: float,
) -> SourceBreakdown:
    """Map transaction categories to source breakdown fields."""
    transfer_tx_ids = _all_transfer_tx_ids(transfers)

    deposits = [
        t for t in transactions
        if t.type == TransactionType.DEPOSIT and t.id not in transfer_tx_ids
    ]

    payroll = 0.0
    gift = 0.0
    investment_sale = 0.0
    other_explained = 0.0

    for dep in deposits:
        if dep.category == TransactionCategory.PAYROLL:
            payroll += dep.amount
        elif dep.category == TransactionCategory.GIFT:
            gift += dep.amount
        elif dep.category == TransactionCategory.INVESTMENT:
            investment_sale += dep.amount
        elif dep.category in (
            TransactionCategory.BUSINESS_INCOME,
            TransactionCategory.GOVERNMENT,
            TransactionCategory.REFUND,
            TransactionCategory.TRANSFER,
        ):
            # Unmatched transfer-category deposits are counted as other_explained
            # (they represent incoming funds even if source account is unknown)
            other_explained += dep.amount
        # other categories (OTHER, CASH, LOAN, etc.) not counted as explained

    explained = payroll + gift + investment_sale + other_explained
    unexplained = max(0, target_dp - explained)

    return SourceBreakdown(
        savings=0.0,
        gift=gift,
        investment_sale=investment_sale,
        property_sale=0.0,
        payroll=payroll,
        other_explained=other_explained,
        unexplained=unexplained,
    )


# ── 4. Client requests ──────────────────────────────────────────────────


def generate_client_requests(
    flags: list[DPFlag],
    transactions: list[DPTransaction],
) -> list[ClientRequest]:
    """Generate document requests based on detected flags."""
    requests: list[ClientRequest] = []
    seen_types: set[FlagType] = set()

    for flag in flags:
        if flag.type == FlagType.CASH_DEPOSIT and FlagType.CASH_DEPOSIT not in seen_types:
            seen_types.add(FlagType.CASH_DEPOSIT)
            requests.append(ClientRequest(
                title="Lettre explicative pour dépôt(s) en espèces",
                reason="Un ou plusieurs dépôts en espèces ont été détectés. "
                       "Le prêteur exige une explication écrite de la provenance de ces fonds.",
                required_docs=[
                    "Lettre explicative signée par l'emprunteur",
                    "Pièces justificatives (reçus de vente, retrait d'un autre compte, etc.)",
                ],
                supporting_transaction_ids=flag.supporting_transaction_ids,
            ))

        elif flag.type in (FlagType.LARGE_DEPOSIT, FlagType.UNEXPLAINED_SOURCE) and FlagType.LARGE_DEPOSIT not in seen_types:
            seen_types.add(FlagType.LARGE_DEPOSIT)
            requests.append(ClientRequest(
                title="Preuve de provenance des fonds",
                reason="Un ou plusieurs dépôts importants nécessitent une preuve de provenance "
                       "pour satisfaire les exigences du prêteur.",
                required_docs=[
                    "Relevé du compte source montrant le retrait",
                    "Contrat de vente ou reçu (si applicable)",
                    "Lettre explicative signée",
                ],
                supporting_transaction_ids=flag.supporting_transaction_ids,
            ))

        elif flag.type == FlagType.PERIOD_GAP and FlagType.PERIOD_GAP not in seen_types:
            seen_types.add(FlagType.PERIOD_GAP)
            requests.append(ClientRequest(
                title="Relevés bancaires complémentaires",
                reason="La période couverte par les relevés est insuffisante. "
                       "Le prêteur exige un minimum de 90 jours de relevés.",
                required_docs=["Relevés bancaires couvrant au moins 90 jours"],
                supporting_transaction_ids=[],
            ))

    # Check for gift transactions
    gift_txs = [t for t in transactions if t.category == TransactionCategory.GIFT]
    if gift_txs:
        requests.append(ClientRequest(
            title="Lettre de don notariée",
            reason="Un ou plusieurs dons ont été identifiés dans les relevés. "
                   "Le prêteur exige une lettre de don notariée et la preuve de la capacité financière du donateur.",
            required_docs=[
                "Lettre de don notariée",
                "Preuve de capacité financière du donateur (relevé bancaire ou état financier)",
                "Preuve du transfert de fonds",
            ],
            supporting_transaction_ids=[t.id for t in gift_txs],
        ))

    return requests


# ── 5. Summary builder ───────────────────────────────────────────────────


def build_summary(
    target_dp: float,
    source_breakdown: SourceBreakdown,
    flags: list[DPFlag],
) -> DPSummary:
    """Build the high-level audit summary."""
    explained = (
        source_breakdown.payroll
        + source_breakdown.gift
        + source_breakdown.investment_sale
        + source_breakdown.property_sale
        + source_breakdown.other_explained
        + source_breakdown.savings
    )
    unexplained = source_breakdown.unexplained

    has_critical = any(f.severity == FlagSeverity.CRITICAL for f in flags)
    has_warning = any(f.severity == FlagSeverity.WARNING for f in flags)
    needs_review = has_critical or unexplained > 0

    review_notes: list[str] = []
    if has_critical:
        critical_count = sum(1 for f in flags if f.severity == FlagSeverity.CRITICAL)
        review_notes.append(f"{critical_count} drapeau(x) critique(s) détecté(s)")
    if has_warning:
        warning_count = sum(1 for f in flags if f.severity == FlagSeverity.WARNING)
        review_notes.append(f"{warning_count} avertissement(s) détecté(s)")
    if unexplained > 0:
        review_notes.append(f"{unexplained:,.2f} $ de la mise de fonds non expliqué")

    return DPSummary(
        dp_target=target_dp,
        dp_explained_amount=explained,
        unexplained_amount=unexplained,
        needs_review=needs_review,
        review_notes=review_notes,
        source_breakdown=source_breakdown,
    )


# ── Main entry point ─────────────────────────────────────────────────────


def analyze(
    extraction: DPExtraction,
    target_downpayment: float,
    closing_date: str,
    borrower_name: str,
    co_borrower_name: str | None = None,
    deal_notes: str | None = None,
) -> DPAuditResult:
    """Full post-processing pipeline.

    Args:
        extraction: Raw Gemini extraction.
        target_downpayment: Target downpayment amount in CAD.
        closing_date: Expected closing date (YYYY-MM-DD).
        borrower_name: Borrower name.
        co_borrower_name: Optional co-borrower name.
        deal_notes: Optional deal notes.

    Returns:
        Complete DPAuditResult.
    """
    txns = extraction.transactions
    accts = extraction.accounts

    transfers = match_transfers(txns)
    flags = detect_flags(txns, transfers, accts, target_downpayment)
    breakdown = calculate_source_breakdown(txns, transfers, target_downpayment)
    client_reqs = generate_client_requests(flags, txns)
    summary = build_summary(target_downpayment, breakdown, flags)

    return DPAuditResult(
        accounts=accts,
        transactions=txns,
        transfers=transfers,
        flags=flags,
        client_requests=client_reqs,
        summary=summary,
        borrower_name=borrower_name,
        co_borrower_name=co_borrower_name or "",
        closing_date=closing_date,
        deal_notes=deal_notes or "",
    )
