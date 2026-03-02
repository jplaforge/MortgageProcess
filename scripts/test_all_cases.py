#!/usr/bin/env python3
"""Run audit_downpayment against all 10 test cases and produce a broker report.

Usage:
    python scripts/test_all_cases.py --token TOKEN [--url URL] [--cases CASE01,CASE05]
"""

import argparse
import asyncio
import base64
import json
import sys
from pathlib import Path

BASE = Path(__file__).parent.parent / "mcp_bank_statements_testset"

MIME_MAP = {
    ".pdf": "application/pdf",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
}

# Case configurations: borrower, target downpayment, closing date, supporting doc folders
CASE_CONFIG = {
    "CASE01": {
        "folder": "CASE01_Baseline_PDF_Text",
        "borrower": "Martin Girard",
        "target": 40000.0,
        "closing_date": "2026-04-01",
        "notes": "Cas de référence — PDF texte natif, aucun flag attendu.",
    },
    "CASE02": {
        "folder": "CASE02_Scanned_PDF_Rotated",
        "borrower": "Martin Girard",
        "target": 35000.0,
        "closing_date": "2026-04-01",
        "notes": "PDF scannés (rotation). Teste la robustesse OCR de Gemini.",
    },
    "CASE03": {
        "folder": "CASE03_PNG_Statements",
        "borrower": "Sophie Nguyen",
        "target": 30000.0,
        "closing_date": "2026-04-01",
        "notes": "Relevés en format PNG — images haute résolution.",
    },
    "CASE04": {
        "folder": "CASE04_Transfer_Chain_3Banks",
        "borrower": "Sophie Nguyen",
        "target": 50000.0,
        "closing_date": "2026-04-01",
        "notes": "Chaîne de transferts entre 3 banques. Flag attendu: transfer_chain 8 000 $ en 2026-02.",
    },
    "CASE05": {
        "folder": "CASE05_Split_Transfer",
        "borrower": "Nadia Tremblay",
        "target": 45000.0,
        "closing_date": "2026-04-01",
        "notes": "Transfert fractionné (split transfer). Flag attendu: split_transfer 6 000 $ en 2026-02.",
    },
    "CASE06": {
        "folder": "CASE06_Gift_Deposit_With_Letter",
        "borrower": "Martin Girard",
        "target": 50000.0,
        "closing_date": "2026-04-01",
        "notes": "Don familial 20 000 $ avec lettre de don. Flag attendu: gift_deposit.",
        "supporting_docs": True,  # includes Gift_Letter
    },
    "CASE07": {
        "folder": "CASE07_Cash_Deposit_and_Crypto",
        "borrower": "Sophie Nguyen",
        "target": 35000.0,
        "closing_date": "2026-04-01",
        "notes": "Dépôt comptant 3 500 $ + source crypto 4 800 $. 2 flags attendus.",
    },
    "CASE08": {
        "folder": "CASE08_Missing_Page_PDF",
        "borrower": "Nadia Tremblay",
        "target": 30000.0,
        "closing_date": "2026-04-01",
        "notes": "Page manquante dans le relevé 2026-01. L'outil doit signaler l'incomplétude.",
    },
    "CASE09": {
        "folder": "CASE09_USD_Account_ForeignCurrency",
        "borrower": "Martin Girard",
        "target": 25000.0,
        "closing_date": "2026-04-01",
        "notes": "Compte USD — devise étrangère. Flag attendu: fx 4 200 USD en 2026-01.",
    },
    "CASE10": {
        "folder": "CASE10_SelfEmployed_Business_Bordel",
        "borrower": "Nadia Tremblay",
        "target": 40000.0,
        "closing_date": "2026-04-01",
        "notes": "Compte d'affaires travailleur autonome. Gros dépôt 12 500 $ en 2026-01.",
    },
}


def load_docs(folder: Path, include_gift_letter: bool = False) -> tuple[list[dict], list[dict]]:
    docs = []
    support = []
    for f in sorted(folder.iterdir()):
        ext = f.suffix.lower()
        mime = MIME_MAP.get(ext)
        if mime is None:
            continue
        b64 = base64.b64encode(f.read_bytes()).decode()
        entry = {"data": b64, "mime_type": mime}
        # Gift letter goes to supporting_documents
        if "gift" in f.name.lower() and include_gift_letter:
            support.append(entry)
            print(f"    [support] {f.name}")
        else:
            docs.append(entry)
            print(f"    [doc]     {f.name}")
    return docs, support


async def run_case(session, case_id: str, cfg: dict) -> dict:
    folder = BASE / cfg["folder"]
    print(f"\n{'='*65}")
    print(f"  {case_id}: {cfg['folder']}")
    print(f"  Emprunteur: {cfg['borrower']} | Cible: {cfg['target']:,.0f} $")
    print(f"  {cfg['notes']}")
    print(f"{'='*65}")

    docs, support = load_docs(folder, include_gift_letter=cfg.get("supporting_docs", False))
    print(f"  {len(docs)} relevé(s), {len(support)} doc(s) de soutien\n")

    kwargs = {
        "documents": docs,
        "target_downpayment_amount": cfg["target"],
        "closing_date": cfg["closing_date"],
        "borrower_name": cfg["borrower"],
        "deal_notes": cfg["notes"],
    }
    if support:
        kwargs["supporting_documents"] = support

    result = await session.call_tool("audit_downpayment", kwargs)

    text_parts = [c.text for c in result.content if hasattr(c, "text")]
    excel_saved = False
    for c in result.content:
        if hasattr(c, "resource") and hasattr(c.resource, "blob"):
            fname = f"audit_{case_id}_{cfg['borrower'].lower().replace(' ', '_')}.xlsx"
            Path(fname).write_bytes(base64.b64decode(c.resource.blob))
            excel_saved = True
            print(f"  Excel sauvegardé: {fname}")

    summary = "\n".join(text_parts)
    print(summary[:2000])
    if len(summary) > 2000:
        print(f"  ... [{len(summary)-2000} chars tronqués]")

    return {
        "case_id": case_id,
        "borrower": cfg["borrower"],
        "summary": summary,
        "excel_saved": excel_saved,
    }


async def main(url: str, token: str | None, selected_cases: list[str] | None):
    from mcp import ClientSession
    from mcp.client.streamable_http import streamablehttp_client

    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    cases_to_run = selected_cases or list(CASE_CONFIG.keys())
    results = []

    print(f"\nServeur: {url}")
    print(f"Cas à tester: {', '.join(cases_to_run)}")

    async with streamablehttp_client(url, headers=headers) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()
            print("Connecté au serveur MCP.\n")

            for case_id in cases_to_run:
                if case_id not in CASE_CONFIG:
                    print(f"CAS INCONNU: {case_id}, ignoré.")
                    continue
                try:
                    r = await run_case(session, case_id, CASE_CONFIG[case_id])
                    results.append(r)
                except Exception as e:
                    print(f"\nERREUR {case_id}: {e}")
                    results.append({"case_id": case_id, "error": str(e)})

    # Final summary
    print("\n" + "="*65)
    print("RÉSUMÉ DES CAS TESTÉS")
    print("="*65)
    for r in results:
        status = "ERREUR" if "error" in r else ("OK" if r.get("excel_saved") else "OK (texte seulement)")
        print(f"  {r['case_id']}: {status}")
    print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="https://mcp-mortgageprocess.onrender.com/mcp")
    parser.add_argument("--token", default=None)
    parser.add_argument("--cases", default=None, help="Comma-separated list of cases, e.g. CASE01,CASE05")
    args = parser.parse_args()

    selected = [c.strip().upper() for c in args.cases.split(",")] if args.cases else None
    asyncio.run(main(args.url, args.token, selected))
