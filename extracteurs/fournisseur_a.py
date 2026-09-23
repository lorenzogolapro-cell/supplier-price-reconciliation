# -*- coding: utf-8 -*-
"""
Extractor for the FOURNISSEUR A price catalogue (Excel file,
tarif_fournisseur.xlsx).

Quirk of this file: the first ~15 rows are a courtesy header (address,
shipping terms, free-shipping threshold, a "TARIF DISTRIBUTEUR
CONFIDENTIEL" notice). The real header row is detected dynamically by
looking for the first row containing both "Reference" and "Designation",
so the code survives a shift in later editions.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

from extracteurs.base import (
    calculer_prix_colis,
    chemin_lisible,
    controler_coherence_prix,
    economie_palier,
    finaliser,
    nettoyer_ean,
    nettoyer_nombre,
    nettoyer_texte,
    normaliser_conditionnement,
    resumer_paliers,
    trouver_ligne_entete,
    _normaliser,
)

FOURNISSEUR = "FOURNISSEUR A"

# Keywords used to locate the real header row
MOTS_CLES_ENTETE = ["Référence", "Désignation"]

# Mapping from file heading -> normalised schema field.
# The key is the normalised heading (lower case, no accents, no
# punctuation). We match by NAME and not by position: that is what avoids
# confusing "Votre tarif net HT unitaire" with the tiers "Votre tarif net
# HT 2/3".
ENTETES_VERS_CHAMPS = {
    "reference": "ref_fournisseur",
    "designation": "designation",
    # CAREFUL: in the 2026 price list this column is headed "Tarif public
    # conseille TTC" -> it really does include VAT, verified against the
    # data (purchase price = incl. VAT / (1 + VAT) x (1 - discount)).
    "tarifpublicconseillettc": "tarif_public_ttc",
    "tva": "tva_taux",
    "remise": "remise_taux",
    "votretarifnethtunitaire": "prix_achat_unitaire_ht",
    # Volume pricing: 878 rows have a tier 2, 468 have a tier 3. These
    # are UNIT PRICES applying from the threshold quantity on, not lot
    # prices.
    "qte2": "palier2_qte",
    "votretarifnetht2": "palier2_prix_ht",
    "qte3": "palier3_qte",
    "votretarifnetht3": "palier3_prix_ht",
    "codelppr": "code_lppr",
    "montantlppr": "montant_lppr",
    "ecopartht": "eco_part_ht",
    "codeean": "ean",
    "conditionnement": "conditionnement",
    "dm": "dispositif_medical",
    "origine": "origine",
}

# Safety net: 0-based positions expected from the 2026 price list. Used
# only for the fields that name matching failed to find (renamed heading,
# merged cell, and so on).
INDEX_SECOURS = {
    "ref_fournisseur": 1,
    "designation": 2,
    "tarif_public_ttc": 3,
    "tva_taux": 4,
    "remise_taux": 5,
    "prix_achat_unitaire_ht": 6,
    "palier2_qte": 7,
    "palier2_prix_ht": 8,
    "palier3_qte": 9,
    "palier3_prix_ht": 10,
    "code_lppr": 11,
    "montant_lppr": 12,
    "eco_part_ht": 14,
    "ean": 16,
    "conditionnement": 17,
    "dispositif_medical": 19,
    "origine": 21,
}

# Columns deliberately ignored:
#   0  Page              -> layout of the printed catalogue
#   13 Packaging retail
#   15 Eco Part TTC      -> recomputable from eco_part_ht
#   18 Poids brut        -> gross weight
#   20 Code douanier     -> customs code

# Fields to convert to numbers / to text
CHAMPS_NUM = {
    "tarif_public_ttc",
    "tva_taux",
    "remise_taux",
    "prix_achat_unitaire_ht",
    "palier2_qte",
    "palier2_prix_ht",
    "palier3_qte",
    "palier3_prix_ht",
    "montant_lppr",
    "eco_part_ht",
}


def _construire_mapping(feuille, ligne_entete: int) -> dict[str, int]:
    """Maps each schema field to a 0-based column index."""
    entetes = next(
        feuille.iter_rows(
            min_row=ligne_entete, max_row=ligne_entete, values_only=True
        )
    )

    mapping: dict[str, int] = {}
    for index, entete in enumerate(entetes):
        if entete is None:
            continue
        champ = ENTETES_VERS_CHAMPS.get(_normaliser(str(entete)))
        # First found wins: prevents a duplicated heading from
        # overwriting the right column.
        if champ and champ not in mapping:
            mapping[champ] = index

    # Fill in the missing fields from the known positions
    for champ, index in INDEX_SECOURS.items():
        if champ not in mapping and index < len(entetes):
            mapping[champ] = index

    return mapping


def extract(path) -> "pd.DataFrame":  # noqa: F821 (type resolved at runtime)
    """Extracts the FOURNISSEUR A catalogue into the common schema."""
    path = Path(path)
    # read_only: the file has ~3100 rows, avoid loading it all into RAM.
    # data_only: we want computed values, not formulas.
    classeur = openpyxl.load_workbook(
        chemin_lisible(path), read_only=True, data_only=True
    )
    feuille = classeur[classeur.sheetnames[0]]

    ligne_entete = trouver_ligne_entete(feuille, MOTS_CLES_ENTETE)
    mapping = _construire_mapping(feuille, ligne_entete)

    manquants = [c for c in ("ref_fournisseur", "designation") if c not in mapping]
    if manquants:
        raise ValueError(f"Essential columns not found: {manquants}")

    lignes: list[dict] = []
    for valeurs in feuille.iter_rows(min_row=ligne_entete + 1, values_only=True):
        if valeurs is None:
            continue

        def cellule(champ):
            """Raw value of a field, or None if the column is missing."""
            index = mapping.get(champ)
            if index is None or index >= len(valeurs):
                return None
            return valeurs[index]

        ref = nettoyer_texte(cellule("ref_fournisseur"))
        designation = nettoyer_texte(cellule("designation"))

        # Rows without a reference or a label are skipped: section
        # separators, blank rows, page footers.
        if not ref or not designation:
            continue

        ligne = {
            "fournisseur": FOURNISSEUR,
            "ref_fournisseur": ref,
            "designation": designation,
            "ean": nettoyer_ean(cellule("ean")),
            "code_lppr": nettoyer_texte(cellule("code_lppr")),
            "dispositif_medical": nettoyer_texte(cellule("dispositif_medical")),
            "origine": nettoyer_texte(cellule("origine")),
            "fichier_source": path.name,
        }
        for champ in CHAMPS_NUM:
            ligne[champ] = nettoyer_nombre(cellule(champ))

        # Pack size: missing or zero -> 1, and the fact is recorded.
        (
            ligne["conditionnement"],
            ligne["conditionnement_suppose"],
        ) = normaliser_conditionnement(cellule("conditionnement"))

        # Full case price = unit price x number of units per case.
        ligne["prix_colis_ht"] = calculer_prix_colis(
            ligne["prix_achat_unitaire_ht"], ligne["conditionnement"]
        )

        # Check: does the net price in the file really match the VAT
        # inclusive list price, less VAT and then less the discount?
        (
            ligne["prix_recalcule"],
            ligne["ecart_prix"],
            ligne["prix_coherent"],
        ) = controler_coherence_prix(
            ligne["tarif_public_ttc"],
            ligne["tva_taux"],
            ligne["remise_taux"],
            ligne["prix_achat_unitaire_ht"],
        )

        # Volume tiers: unit prices applying beyond a threshold
        prix_base = ligne["prix_achat_unitaire_ht"]
        ligne["economie_palier2_pct"] = economie_palier(
            prix_base, ligne["palier2_prix_ht"]
        )
        ligne["economie_palier3_pct"] = economie_palier(
            prix_base, ligne["palier3_prix_ht"]
        )
        ligne["paliers"] = resumer_paliers(
            prix_base,
            [
                (ligne["palier2_qte"], ligne["palier2_prix_ht"]),
                (ligne["palier3_qte"], ligne["palier3_prix_ht"]),
            ],
        )

        lignes.append(ligne)

    classeur.close()
    return finaliser(lignes)
