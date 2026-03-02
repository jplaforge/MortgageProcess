#!/usr/bin/env python3
"""Manual test client for the MCP server.

Usage:
    # Test against local server
    python scripts/test_client.py

    # Test against remote Render deployment
    python scripts/test_client.py --url https://your-app.onrender.com/mcp --token YOUR_TOKEN

    # Test with a real PDF
    python scripts/test_client.py --file /path/to/statement.pdf
"""

import argparse
import asyncio
import base64
import sys
from pathlib import Path


async def test_with_mcp_client(url: str, token: str | None, pdf_path: str | None):
    """Test using the MCP client SDK."""
    from mcp import ClientSession
    from mcp.client.streamable_http import streamablehttp_client

    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    async with streamablehttp_client(url, headers=headers) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # List tools
            tools = await session.list_tools()
            print("Available tools:")
            for tool in tools.tools:
                print(f"  - {tool.name}: {tool.description[:80]}...")
            print()

            # Health check
            print("Running health_check...")
            result = await session.call_tool("health_check", {})
            for content in result.content:
                print(f"  {content.text}")
            print()

            # Analyze bank statements (with sample or real PDF)
            if pdf_path:
                pdf_data = Path(pdf_path).read_bytes()
                mime_type = "application/pdf"
            else:
                # Create a minimal test PDF
                pdf_data = b"%PDF-1.4 sample"
                mime_type = "application/pdf"
                print("(Using sample PDF data — for real testing, use --file)")

            b64_data = base64.b64encode(pdf_data).decode()

            print("Running analyze_bank_statements...")
            result = await session.call_tool(
                "analyze_bank_statements",
                {
                    "documents": [{"data": b64_data, "mime_type": mime_type}],
                    "borrower_name": "Test Emprunteur",
                },
            )

            for content in result.content:
                if hasattr(content, "text"):
                    print(content.text[:500])
                elif hasattr(content, "resource"):
                    uri = str(content.resource.uri)
                    filename = uri.split("///")[-1] if "///" in uri else "output.xlsx"
                    excel_bytes = base64.b64decode(content.resource.blob)
                    out_path = Path(filename)
                    out_path.write_bytes(excel_bytes)
                    print(f"  Excel saved: {out_path.resolve()} ({len(excel_bytes)} bytes)")
            print()


def main():
    parser = argparse.ArgumentParser(description="Test MCP mortgage analyzer")
    parser.add_argument("--url", default="http://localhost:8000/mcp", help="MCP server URL")
    parser.add_argument("--token", default=None, help="Auth bearer token")
    parser.add_argument("--file", default=None, help="Path to a real PDF bank statement")
    args = parser.parse_args()

    asyncio.run(test_with_mcp_client(args.url, args.token, args.file))


if __name__ == "__main__":
    main()
