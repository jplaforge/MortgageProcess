"""Health check tool for MCP server."""

from mortgage_mcp.services.vertex_ai import check_vertex_ai_connection


async def health_check() -> str:
    """Check Vertex AI connectivity and return status."""
    result = check_vertex_ai_connection()
    if hasattr(result, "__await__"):
        result = await result

    if result["status"] == "connected":
        return (
            f"Statut: Connecté\n"
            f"Modèle: {result['model']}\n"
            f"Région: {result['location']}\n"
            f"Réponse: {result['response']}"
        )
    return (
        f"Statut: Erreur\n"
        f"Modèle: {result['model']}\n"
        f"Région: {result['location']}\n"
        f"Erreur: {result['error']}"
    )
