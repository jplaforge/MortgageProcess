#!/usr/bin/env python3
"""Test the MCP server with multiple bank statement files from one client.

Usage:
    python scripts/test_multi_files.py /path/to/folder/
    python scripts/test_multi_files.py /path/to/folder/ --url https://... --token TOKEN
"""

import argparse
import asyncio
import base64
import mimetypes
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


async def run(url: str, token: str | None, folder: Path):
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

            # Call analyze
            print("Calling analyze_bank_statements ...")
            result = await session.call_tool(
                "analyze_bank_statements",
                {
                    "documents": docs,
                    "borrower_name": "Martin Girard",
                    "business_name": "Girard Consultation TI",
                    "business_type": "Consultation informatique",
                },
            )

            print("\n" + "=" * 70)
            print("RESULTS")
            print("=" * 70)
            for content in result.content:
                if hasattr(content, "text"):
                    print(content.text)
                elif hasattr(content, "resource"):
                    uri = str(content.resource.uri)
                    filename = uri.split("///")[-1] if "///" in uri else "output.xlsx"
                    excel_bytes = base64.b64decode(content.resource.blob)
                    out_path = Path(filename)
                    out_path.write_bytes(excel_bytes)
                    print(f"\nExcel saved: {out_path.resolve()} ({len(excel_bytes):,} bytes)")
            print()


def main():
    parser = argparse.ArgumentParser(description="Test MCP with multiple files")
    parser.add_argument("folder", help="Folder containing bank statement files")
    parser.add_argument("--url", default="https://mcp-mortgageprocess.onrender.com/mcp")
    parser.add_argument("--token", default=None, help="Auth token (defaults to MCP_AUTH_TOKEN env)")
    args = parser.parse_args()

    import os
    token = args.token or os.environ.get("MCP_AUTH_TOKEN")

    asyncio.run(run(args.url, token, Path(args.folder)))


if __name__ == "__main__":
    main()
