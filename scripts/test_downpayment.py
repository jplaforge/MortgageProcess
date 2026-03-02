#!/usr/bin/env python3
"""Test the MCP server downpayment audit tool with bank statement files.

Usage:
    python scripts/test_downpayment.py /path/to/folder/ --dp 80000 --closing 2025-06-15 --borrower "Jean Tremblay"
    python scripts/test_downpayment.py /path/to/folder/ --dp 80000 --closing 2025-06-15 --borrower "Jean" --url https://... --token TOKEN
"""

import argparse
import asyncio
import base64
import json
from pathlib import Path


MIME_MAP = {
    ".pdf": "application/pdf",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".csv": "text/csv",
}


def load_documents(folder: Path) -> list[dict]:
    """Load all supported files from a folder as base64 documents."""
    docs = []
    for f in sorted(folder.iterdir()):
        ext = f.suffix.lower()
        mime = MIME_MAP.get(ext)
        if mime is None:
            print(f"  Skipping unsupported file: {f.name}")
            continue
        b64 = base64.b64encode(f.read_bytes()).decode()
        docs.append({"data": b64, "mime_type": mime})
        print(f"  Loaded: {f.name} ({mime}, {f.stat().st_size:,} bytes)")
    return docs


async def run(
    url: str,
    token: str | None,
    folder: Path,
    dp_amount: float,
    closing_date: str,
    borrower_name: str,
    co_borrower: str | None,
    notes: str | None,
):
    from mcp import ClientSession
    from mcp.client.streamable_http import streamablehttp_client

    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    print(f"\nConnecting to {url} ...")
    async with streamablehttp_client(url, headers=headers) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()
            print("Connected.\n")

            # Load documents
            print(f"Loading files from {folder}/")
            docs = load_documents(folder)
            print(f"\n{len(docs)} document(s) loaded.\n")

            if not docs:
                print("No supported files found.")
                return

            # Build tool arguments
            tool_args = {
                "documents": docs,
                "target_downpayment_amount": dp_amount,
                "closing_date": closing_date,
                "borrower_name": borrower_name,
            }
            if co_borrower:
                tool_args["co_borrower_name"] = co_borrower
            if notes:
                tool_args["deal_notes"] = notes

            print(f"Calling audit_downpayment (DP={dp_amount:,.0f}$, closing={closing_date}) ...")
            result = await session.call_tool("audit_downpayment", tool_args)

            print("\n" + "=" * 70)
            print("RESULTS")
            print("=" * 70)
            for content in result.content:
                if hasattr(content, "text"):
                    # Check if it's JSON
                    try:
                        data = json.loads(content.text)
                        # Save JSON to file
                        borrower_slug = borrower_name.replace(" ", "_").lower()
                        json_path = Path(f"audit_mise_de_fonds_{borrower_slug}.json")
                        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False))
                        print(f"\nJSON saved: {json_path.resolve()}")
                        # Print summary stats
                        summary = data.get("summary", {})
                        print(f"  DP target: {summary.get('dp_target', 'N/A'):,.2f} $")
                        print(f"  Explained: {summary.get('dp_explained_amount', 'N/A'):,.2f} $")
                        print(f"  Unexplained: {summary.get('unexplained_amount', 'N/A'):,.2f} $")
                        print(f"  Needs review: {summary.get('needs_review', 'N/A')}")
                        print(f"  Flags: {len(data.get('flags', []))}")
                        print(f"  Transfers: {len(data.get('transfers', []))}")
                    except (json.JSONDecodeError, TypeError):
                        # It's the markdown summary
                        print(content.text)
                elif hasattr(content, "resource"):
                    uri = str(content.resource.uri)
                    filename = uri.split("///")[-1] if "///" in uri else "audit_output.xlsx"
                    excel_bytes = base64.b64decode(content.resource.blob)
                    out_path = Path(filename)
                    out_path.write_bytes(excel_bytes)
                    print(f"\nExcel saved: {out_path.resolve()} ({len(excel_bytes):,} bytes)")
            print()


def main():
    parser = argparse.ArgumentParser(description="Test MCP downpayment audit tool")
    parser.add_argument("folder", help="Folder containing bank statement files")
    parser.add_argument("--dp", type=float, required=True, help="Target downpayment amount in CAD")
    parser.add_argument("--closing", required=True, help="Closing date (YYYY-MM-DD)")
    parser.add_argument("--borrower", required=True, help="Borrower name")
    parser.add_argument("--co-borrower", default=None, help="Co-borrower name (optional)")
    parser.add_argument("--notes", default=None, help="Deal notes (optional)")
    parser.add_argument("--url", default="https://mcp-mortgageprocess.onrender.com/mcp")
    parser.add_argument("--token", default=None, help="Auth token (defaults to MCP_AUTH_TOKEN env)")
    args = parser.parse_args()

    import os
    token = args.token or os.environ.get("MCP_AUTH_TOKEN")

    asyncio.run(run(
        args.url, token, Path(args.folder),
        args.dp, args.closing, args.borrower,
        args.co_borrower, args.notes,
    ))


if __name__ == "__main__":
    main()
