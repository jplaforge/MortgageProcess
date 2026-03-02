"""MCP server for self-employed mortgage income analysis."""

import os
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from mcp.server.auth.provider import AccessToken, TokenVerifier
from mcp.server.auth.settings import AuthSettings
from mcp.server.fastmcp import Context, FastMCP
from mcp.types import ToolAnnotations
from pydantic import AnyHttpUrl

from mortgage_mcp.config import settings

port = int(os.environ.get("PORT", settings.port))

RENDER_URL = "https://mcp-mortgageprocess.onrender.com"


class BearerTokenVerifier(TokenVerifier):
    """Verify requests against the MCP_AUTH_TOKEN environment variable."""

    async def verify_token(self, token: str) -> AccessToken | None:
        if not settings.mcp_auth_token:
            return None
        if token != settings.mcp_auth_token:
            return None
        return AccessToken(token=token, client_id="mcp-client", scopes=[])


@asynccontextmanager
async def server_lifespan(server: FastMCP) -> AsyncIterator[None]:
    """Run once at startup: set up GCP credentials."""
    settings.setup_gcp_credentials()
    yield


_use_auth = bool(settings.mcp_auth_token)

mcp = FastMCP(
    "WelcomeSpaces Mortgage Analyzer",
    instructions=(
        "Ce serveur offre deux outils d'analyse hypothécaire:\n"
        "1. analyze_bank_statements — Analyse les relevés bancaires de travailleurs autonomes "
        "pour calculer le revenu admissible.\n"
        "2. audit_downpayment — Audite la provenance de la mise de fonds: "
        "trace les transferts, détecte les dépôts suspects, et génère un dossier d'audit complet."
    ),
    host="0.0.0.0",
    port=port,
    lifespan=server_lifespan,
    auth=AuthSettings(
        issuer_url=AnyHttpUrl(RENDER_URL),
        resource_server_url=AnyHttpUrl(RENDER_URL),
    ) if _use_auth else None,
    token_verifier=BearerTokenVerifier() if _use_auth else None,
)


@mcp.tool(
    title="Analyser les relevés bancaires",
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    ),
)
async def analyze_bank_statements(
    documents: list[dict],
    ctx: Context,
    borrower_name: str | None = None,
    business_name: str | None = None,
    business_type: str | None = None,
) -> list:
    """Analyse les relevés bancaires d'un travailleur autonome pour calculer le revenu admissible.

    Envoie les documents à Gemini pour extraction structurée, puis génère un fichier Excel
    détaillé avec ventilation mensuelle des revenus.

    Args:
        documents: Liste de documents encodés en base64. Chaque élément doit avoir:
            - data: contenu du fichier en base64
            - mime_type: type MIME (application/pdf, image/jpeg, image/png, text/csv)
        borrower_name: Nom de l'emprunteur (optionnel, pour le rapport)
        business_name: Nom de l'entreprise (optionnel, pour contexte IA)
        business_type: Type d'entreprise (optionnel, pour contexte IA)

    Returns:
        Résumé textuel en français + fichier Excel encodé en base64.
    """
    from mortgage_mcp.tools.analyze_bank_statements import (
        analyze_bank_statements as _analyze,
    )

    return await _analyze(documents, ctx, borrower_name, business_name, business_type)


@mcp.tool(
    title="Auditer la mise de fonds",
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    ),
)
async def audit_downpayment(
    documents: list[dict],
    target_downpayment_amount: float,
    closing_date: str,
    borrower_name: str,
    ctx: Context,
    supporting_documents: list[dict] | None = None,
    co_borrower_name: str | None = None,
    deal_notes: str | None = None,
) -> list:
    """Audite la provenance de la mise de fonds pour un dossier hypothécaire.

    Analyse les relevés bancaires pour tracer chaque dollar: transferts inter-comptes,
    dépôts en espèces, dons, et sources non expliquées. Génère un dossier d'audit complet.

    Args:
        documents: Relevés bancaires encodés en base64. Chaque élément doit avoir:
            - data: contenu du fichier en base64
            - mime_type: type MIME (application/pdf, image/jpeg, image/png, text/csv)
        target_downpayment_amount: Montant cible de la mise de fonds en CAD.
        closing_date: Date de clôture prévue (YYYY-MM-DD).
        borrower_name: Nom de l'emprunteur.
        supporting_documents: Documents justificatifs optionnels (lettres de don, etc.).
        co_borrower_name: Nom du co-emprunteur (optionnel).
        deal_notes: Notes sur le dossier (optionnel).

    Returns:
        Résumé textuel en français + JSON structuré + fichier Excel d'audit.
    """
    from mortgage_mcp.tools.downpayment_audit import (
        audit_downpayment as _audit,
    )

    return await _audit(
        documents, target_downpayment_amount, closing_date, borrower_name, ctx,
        supporting_documents, co_borrower_name, deal_notes,
    )


@mcp.tool(
    title="Vérifier l'état du serveur",
    annotations=ToolAnnotations(
        readOnlyHint=True,
        destructiveHint=False,
        idempotentHint=True,
        openWorldHint=False,
    ),
)
async def health_check(ctx: Context) -> str:
    """Vérifie la connectivité avec Vertex AI et retourne le statut du serveur."""
    from mortgage_mcp.tools.health import health_check as _health_check

    return await _health_check(ctx)


def main() -> None:
    mcp.run(transport="streamable-http")


if __name__ == "__main__":
    main()
