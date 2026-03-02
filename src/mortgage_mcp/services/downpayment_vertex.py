"""Vertex AI Gemini integration for downpayment / source of funds extraction."""

from google import genai
from google.genai import types

from mortgage_mcp.config import settings
from mortgage_mcp.models.downpayment import DPExtraction
from mortgage_mcp.services.document_parser import ParsedDocument

DP_EXTRACTION_PROMPT = """\
Tu es un analyste spécialisé dans la vérification de la provenance de la mise de fonds \
pour des demandes de prêt hypothécaire au Québec.

Analyse les relevés bancaires fournis et extrais TOUTES les transactions (dépôts ET retraits) \
de manière exhaustive. L'objectif est de tracer la provenance de chaque dollar de la mise de fonds.

## Instructions

### 1. Comptes bancaires
Pour CHAQUE compte bancaire identifié dans les documents, crée une entrée dans `accounts`:
- `account_id`: identifiant court (A1, A2, A3...)
- `institution`: nom de l'institution financière
- `account_number_last4`: 4 derniers chiffres du numéro de compte (si visible)
- `holder_name`: nom du titulaire
- `period_start` / `period_end`: période couverte par les relevés (YYYY-MM-DD)
- `opening_balance` / `closing_balance`: soldes d'ouverture et de fermeture
- `confidence`: score de confiance 0-1

### 2. Transactions
Pour CHAQUE transaction sur CHAQUE compte, crée une entrée dans `transactions`:
- `id`: identifiant unique au format "account_id-NNN" (ex: A1-001, A1-002, A2-001)
  Les IDs doivent être séquentiels par compte.
- `date`: date (YYYY-MM-DD)
- `description`: description complète telle qu'affichée
- `amount`: montant en CAD (toujours positif)
- `type`: "deposit" ou "withdrawal"
- `category`: catégorisation basique:
  * `payroll`: salaire, paie régulière (mots-clés: PAIE, SALAIRE, PAYROLL, DIRECT DEPOSIT)
  * `business_income`: revenus d'entreprise, clients, honoraires
  * `transfer`: virements entre comptes, transferts (VIREMENT, TRANSFERT, TFR, INTERAC entre comptes)
  * `cash`: dépôts en espèces (CASH, COMPTANT, GUICHET, ATM, DEPOT ESPECES)
  * `government`: paiements gouvernementaux (TPS, TVH, ARC, CRA, PRESTATIONS)
  * `investment`: placements, REER, CELI (PLACEMENT, REER, CELI, INVESTISSEMENT)
  * `gift`: dons identifiables (DON, CADEAU)
  * `loan`: prêts, marges de crédit (PRET, MARGE, LOC, EMPRUNT)
  * `refund`: remboursements (REMBOURSEMENT, REFUND, CREDIT)
  * `bill_payment`: paiements de factures (PAIEMENT, FACTURE, BELL, HYDRO, etc.)
  * `purchase`: achats (ACHAT, POS, DEBIT)
  * `other`: tout ce qui ne correspond à aucune catégorie ci-dessus
- `account_id`: correspond au account_id du compte
- `page_source`: numéro de page dans le document source
- `confidence`: score de confiance 0-1
- `normalized_description`: description simplifiée/normalisée
- `merchant_guess`: estimation de la source/commerçant

### 3. Règles importantes
- Extrais TOUTES les transactions sans exception, pas seulement les dépôts
- Les montants sont toujours POSITIFS — le champ `type` distingue dépôts et retraits
- Les dates en format YYYY-MM-DD
- Sois conservateur dans la catégorisation: en cas de doute, utilise `other`
- NE FAIS PAS le matching de transferts inter-comptes — le post-traitement Python s'en charge
- NE FAIS PAS la détection de drapeaux — le post-traitement Python s'en charge
- Concentre-toi uniquement sur l'extraction précise et exhaustive
"""


def _build_dp_contents(
    documents: list[ParsedDocument],
    borrower_name: str | None = None,
    co_borrower_name: str | None = None,
    closing_date: str | None = None,
    deal_notes: str | None = None,
) -> list[types.Content]:
    """Build the multimodal content list for Gemini."""
    parts: list[types.Part] = []

    context_lines = []
    if borrower_name:
        context_lines.append(f"Nom de l'emprunteur: {borrower_name}")
    if co_borrower_name:
        context_lines.append(f"Nom du co-emprunteur: {co_borrower_name}")
    if closing_date:
        context_lines.append(f"Date de clôture prévue: {closing_date}")
    if deal_notes:
        context_lines.append(f"Notes sur le dossier: {deal_notes}")

    if context_lines:
        parts.append(types.Part.from_text(
            text="Contexte du dossier:\n" + "\n".join(context_lines) + "\n"
        ))

    parts.append(types.Part.from_text(text=DP_EXTRACTION_PROMPT))

    for doc in documents:
        parts.append(types.Part.from_bytes(data=doc.data, mime_type=doc.mime_type))

    return [types.Content(role="user", parts=parts)]


async def extract_dp_transactions(
    documents: list[ParsedDocument],
    borrower_name: str | None = None,
    co_borrower_name: str | None = None,
    closing_date: str | None = None,
    deal_notes: str | None = None,
) -> DPExtraction:
    """Send documents to Gemini and get structured extraction for downpayment audit.

    Returns:
        DPExtraction with accounts and all transactions.
    """
    settings.setup_gcp_credentials()

    client = genai.Client(
        vertexai=True,
        project=settings.google_cloud_project,
        location=settings.google_cloud_location,
    )

    contents = _build_dp_contents(
        documents, borrower_name, co_borrower_name, closing_date, deal_notes
    )

    response = await client.aio.models.generate_content(
        model=settings.gemini_model,
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=DPExtraction,
            temperature=0.1,
        ),
    )

    return DPExtraction.model_validate_json(response.text)
