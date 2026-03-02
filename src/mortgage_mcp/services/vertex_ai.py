"""Vertex AI Gemini integration for bank statement extraction."""

from google import genai
from google.genai import types

from mortgage_mcp.config import settings
from mortgage_mcp.models.bank_statement import BankStatementExtraction
from mortgage_mcp.services.document_parser import ParsedDocument

EXTRACTION_PROMPT = """\
Tu es un analyste financier spécialisé dans l'évaluation de revenus de travailleurs autonomes pour des demandes de prêt hypothécaire au Québec.

Analyse les relevés bancaires fournis et extrais les informations suivantes de façon exhaustive et précise:

1. **Informations du compte**: Titulaire, institution financière, numéro de compte (derniers 4 chiffres), période couverte.

2. **Dépôts**: Pour CHAQUE dépôt, identifie:
   - La date exacte
   - La description complète
   - Le montant
   - La catégorie:
     * `business_income`: Revenus d'entreprise, paiements de clients, virements Interac de clients, honoraires professionnels, revenus de contrats
     * `personal_transfer`: Transferts personnels entre comptes, virements d'un conjoint/famille
     * `government`: Crédits TPS/TVH, remboursements d'impôt, prestations gouvernementales
     * `loan_credit`: Prêts, marges de crédit, avances
     * `refund`: Remboursements de fournisseurs, retours
     * `other`: Tout autre dépôt non classifiable
   - Le compte source (`account`): identifiant court du compte bancaire d'où provient la transaction, au format "Institution XX99" (ex: "Desjardins XX43", "Banque Boréale XX12"). Utilise le nom abrégé de l'institution + les 2-4 derniers chiffres du compte.

3. **Retraits**: Pour CHAQUE retrait significatif, identifie la date, description, montant, catégorie de dépense, et le compte source (`account`) au même format que pour les dépôts.

4. **Ventilation mensuelle**: Pour chaque mois, calcule:
   - `total_deposits`: somme de tous les dépôts
   - `business_deposits`: somme des dépôts `business_income`
   - `personal_transfers`: somme des dépôts `personal_transfer`
   - `government_deposits`: somme des dépôts `government`
   - `refund_deposits`: somme des dépôts `refund`
   - `loan_credit_deposits`: somme des dépôts `loan_credit`
   - `other_deposits`: somme des dépôts `other` uniquement
   - `total_withdrawals`: somme de tous les retraits

5. **Totaux**: Calcule le revenu d'affaires total, le revenu mensuel moyen, et le revenu annualisé (moyenne × 12).

6. **Comptes multiples**: L'emprunteur peut fournir des relevés de PLUSIEURS comptes bancaires (institutions différentes). Tu dois analyser TOUS les documents fournis et inclure les transactions de CHAQUE compte. Combine les données de tous les comptes dans une seule ventilation mensuelle.

7. **Déduplication**: Uniquement au sein d'un MÊME compte, si des captures d'écran ou pages se chevauchent:
   - Un doublon potentiel est défini comme: même date ET même description ET même montant ET même compte. Si les dates diffèrent, ce n'est PAS un doublon.
   - Ne retire un doublon QUE si tu es certain qu'il s'agit de la même transaction (ex: pages qui se chevauchent dans un même relevé PDF).
   - En cas de doute, GARDE la transaction et signale-la dans `confidence_notes` avec la mention "Doublon potentiel".
   - NE PAS éliminer des transactions provenant de comptes différents, même si elles ont le même montant ou la même date.

9. **Détection NSF / Découverts**: Identifie tous les frais de fonds insuffisants (NSF), découverts, et items retournés dans les retraits. Cherche les descriptions contenant: "NSF", "FONDS INSUFFISANTS", "DÉCOUVERT", "ITEM RETOURNÉ", "PROVISION INSUFFISANTE", "CHÈQUE RETOURNÉ", "FRAIS RETOUR". Pour chaque événement, extrais la date, description, montant des frais et le compte. Calcule le total des frais NSF.

10. **Obligations récurrentes**: Identifie les retraits qui se répètent mensuellement (même bénéficiaire, montant similaire ±5%). Classe-les par type:
      - `hypotheque`: paiements hypothécaires ou loyer
      - `pret_auto`: prêt automobile
      - `marge_credit`: marge de crédit, carte de crédit (paiement minimum récurrent)
      - `assurance`: primes d'assurance
      - `telecom`: téléphone, internet, câble
      - `pension_alimentaire`: pension alimentaire, soutien aux enfants
      - `autre`: autre obligation récurrente
      Calcule le total mensuel de toutes les obligations récurrentes identifiées.
      Note: une obligation récurrente doit apparaître au moins 2 fois sur la période analysée.

8. **Notes de confiance**: Signale tout élément nécessitant une vérification par le courtier:
   - Dépôts inhabituellement élevés
   - Revenus irréguliers
   - Transferts ambigus entre revenus et transferts personnels
   - Périodes sans activité
   - Qualité des documents (illisible, pages manquantes, etc.)
   - Doublons potentiels détectés (signaler, ne pas retirer sauf si certain)

IMPORTANT:
- Tous les montants en dollars canadiens (CAD)
- Les dates en format YYYY-MM-DD
- Les mois en format YYYY-MM
- Sois conservateur: en cas de doute sur la catégorie d'un dépôt, classe-le comme `other` et ajoute une note
- Un travailleur autonome typique reçoit des paiements de clients variés par virement, chèque ou Interac
- Si les relevés proviennent de plusieurs institutions, indique toutes les institutions séparées par " / " dans le champ institution (ex: "Banque Boréale / Caisse Laurentienne")
- Analyse TOUS les documents fournis sans en ignorer aucun
"""


def _build_contents(
    documents: list[ParsedDocument],
    borrower_name: str | None = None,
    business_name: str | None = None,
    business_type: str | None = None,
) -> list[types.Content]:
    """Build the multimodal content list for Gemini."""
    parts: list[types.Part] = []

    # Add context if provided
    context_lines = []
    if borrower_name:
        context_lines.append(f"Nom de l'emprunteur: {borrower_name}")
    if business_name:
        context_lines.append(f"Nom de l'entreprise: {business_name}")
    if business_type:
        context_lines.append(f"Type d'entreprise: {business_type}")

    if context_lines:
        parts.append(types.Part.from_text(
            text="Contexte:\n" + "\n".join(context_lines) + "\n"
        ))

    parts.append(types.Part.from_text(text=EXTRACTION_PROMPT))

    # Add each document as a binary part
    for doc in documents:
        parts.append(types.Part.from_bytes(data=doc.data, mime_type=doc.mime_type))

    return [types.Content(role="user", parts=parts)]


async def extract_bank_statements(
    documents: list[ParsedDocument],
    borrower_name: str | None = None,
    business_name: str | None = None,
    business_type: str | None = None,
) -> BankStatementExtraction:
    """Send documents to Gemini and get structured extraction.

    Returns:
        BankStatementExtraction with all extracted data.
    """
    settings.setup_gcp_credentials()

    client = genai.Client(
        vertexai=True,
        project=settings.google_cloud_project,
        location=settings.google_cloud_location,
    )

    contents = _build_contents(
        documents, borrower_name, business_name, business_type
    )

    response = await client.aio.models.generate_content(
        model=settings.gemini_model,
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=BankStatementExtraction,
            temperature=0.1,
        ),
    )

    return BankStatementExtraction.model_validate_json(response.text)


async def check_vertex_ai_connection() -> dict:
    """Verify connectivity to Vertex AI. Returns status dict."""
    settings.setup_gcp_credentials()

    try:
        client = genai.Client(
            vertexai=True,
            project=settings.google_cloud_project,
            location=settings.google_cloud_location,
        )
        response = await client.aio.models.generate_content(
            model=settings.gemini_model,
            contents="Réponds uniquement: OK",
            config=types.GenerateContentConfig(
                max_output_tokens=10,
                temperature=0.0,
            ),
        )
        return {
            "status": "connected",
            "model": settings.gemini_model,
            "location": settings.google_cloud_location,
            "response": (response.text or "").strip(),
        }
    except Exception as exc:
        return {
            "status": "error",
            "model": settings.gemini_model,
            "location": settings.google_cloud_location,
            "error": str(exc),
        }
