"""MCP server for self-employed mortgage income analysis."""

import os

from mcp.server.auth.provider import AccessToken, TokenVerifier
from mcp.server.fastmcp import FastMCP

from mortgage_mcp.config import settings

port = int(os.environ.get("PORT", settings.port))


class BearerTokenVerifier(TokenVerifier):
    """Verify requests against the MCP_AUTH_TOKEN environment variable."""

    async def verify_token(self, token: str) -> AccessToken | None:
        if not settings.mcp_auth_token:
            return None
        if token != settings.mcp_auth_token:
            return None
        return AccessToken(token=token, client_id="mcp-client", scopes=[])


token_verifier = BearerTokenVerifier() if settings.mcp_auth_token else None

mcp = FastMCP(
    "WelcomeSpaces Mortgage Analyzer",
    instructions=(
        "Ce serveur analyse les relevés bancaires de travailleurs autonomes "
        "pour calculer le revenu admissible aux fins d'une demande de prêt hypothécaire. "
        "Envoyez des relevés bancaires en base64 via l'outil analyze_bank_statements."
    ),
    host="0.0.0.0",
    port=port,
    token_verifier=token_verifier,
)


@mcp.tool()
async def analyze_bank_statements(
    documents: list[dict],
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

    return await _analyze(documents, borrower_name, business_name, business_type)


@mcp.tool()
async def health_check() -> str:
    """Vérifie la connectivité avec Vertex AI et retourne le statut du serveur."""
    from mortgage_mcp.tools.health import health_check as _health_check

    return await _health_check()


def main() -> None:
    settings.setup_gcp_credentials()

    mcp.run(transport="streamable-http")


if __name__ == "__main__":
    main()
