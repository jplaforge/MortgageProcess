"""Health check tool for MCP server."""

from mcp.server.fastmcp import Context

from mortgage_mcp.services.vertex_ai import check_vertex_ai_connection


async def health_check(ctx: Context) -> str:
    """Check Vertex AI connectivity and return status."""
    await ctx.info("Vérification de la connectivité Vertex AI")
    result = check_vertex_ai_connection()
    if hasattr(result, "__await__"):
        result = await result

    if result["status"] == "connected":
        await ctx.info(f"Vertex AI connecté — modèle: {result['model']}")
        return (
            f"Statut: Connecté\n"
            f"Modèle: {result['model']}\n"
            f"Région: {result['location']}\n"
            f"Réponse: {result['response']}"
        )
    await ctx.warning(f"Vertex AI non disponible: {result.get('error', 'inconnu')}")
    return (
        f"Statut: Erreur\n"
        f"Modèle: {result['model']}\n"
        f"Région: {result['location']}\n"
        f"Erreur: {result['error']}"
    )
