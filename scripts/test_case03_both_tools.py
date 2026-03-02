#!/usr/bin/env python3
"""Run both analyze_bank_statements and audit_downpayment for CASE03 (Sophie Nguyen / PNG).

Usage:
    python scripts/test_case03_both_tools.py --token TOKEN [--url URL]
"""

import argparse
import asyncio
import base64
from pathlib import Path

FOLDER = Path(__file__).parent.parent / "mcp_bank_statements_testset/CASE03_PNG_Statements"

BORROWER = "Sophie Nguyen"
TARGET_DOWNPAYMENT = 30000.0
CLOSING_DATE = "2026-04-01"


def load_docs(folder: Path) -> list[dict]:
    docs = []
    for f in sorted(folder.iterdir()):
        if f.suffix.lower() == ".png":
            b64 = base64.b64encode(f.read_bytes()).decode()
            docs.append({"data": b64, "mime_type": "image/png"})
            print(f"  Loaded: {f.name} ({f.stat().st_size:,} bytes)")
    return docs


async def run(url: str, token: str | None):
    from mcp import ClientSession
    from mcp.client.streamable_http import streamablehttp_client

    headers = {"Authorization": f"Bearer {token}"} if token else {}

    print(f"\nConnecting to {url} ...")
    async with streamablehttp_client(url, headers=headers) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()
            print("Connected.\n")

            docs = load_docs(FOLDER)
            print(f"\n{len(docs)} file(s) loaded.\n")

            # ── Tool 1: analyze_bank_statements ──────────────────────────────
            print("=" * 65)
            print("TOOL 1: analyze_bank_statements")
            print("=" * 65)
            result1 = await session.call_tool(
                "analyze_bank_statements",
                {
                    "documents": docs,
                    "borrower_name": BORROWER,
                    "business_name": "Nguyen Conseil",
                    "business_type": "Services-conseils",
                },
            )
            for c in result1.content:
                if hasattr(c, "text"):
                    print(c.text)
                elif hasattr(c, "resource") and hasattr(c.resource, "blob"):
                    fname = f"analyse_revenu_sophie_nguyen.xlsx"
                    Path(fname).write_bytes(base64.b64decode(c.resource.blob))
                    print(f"\nExcel saved: {Path(fname).resolve()}")

            # ── Tool 2: audit_downpayment ────────────────────────────────────
            print("\n" + "=" * 65)
            print("TOOL 2: audit_downpayment")
            print("=" * 65)
            result2 = await session.call_tool(
                "audit_downpayment",
                {
                    "documents": docs,
                    "target_downpayment_amount": TARGET_DOWNPAYMENT,
                    "closing_date": CLOSING_DATE,
                    "borrower_name": BORROWER,
                    "deal_notes": "Relevés PNG haute résolution — test robustesse OCR.",
                },
            )
            for c in result2.content:
                if hasattr(c, "text"):
                    print(c.text)
                elif hasattr(c, "resource") and hasattr(c.resource, "blob"):
                    fname = f"audit_CASE03_sophie_nguyen.xlsx"
                    Path(fname).write_bytes(base64.b64decode(c.resource.blob))
                    print(f"\nExcel saved: {Path(fname).resolve()}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="https://mcp-mortgageprocess.onrender.com/mcp")
    parser.add_argument("--token", default=None)
    args = parser.parse_args()

    import os
    token = args.token or os.environ.get("MCP_AUTH_TOKEN")
    asyncio.run(run(args.url, token))
