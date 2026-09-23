# -*- coding: utf-8 -*-
"""
Generic extractor for Excel price lists.

Suppliers share no layout whatsoever, but they all name their columns the
same way: a reference, a label, a price, and sometimes an EAN. This module
locates those columns by their heading rather than by their position,
which saves writing one extractor per supplier.

It does not replace the dedicated extractors: when a price list has a
logic of its own (tiers named in the header at Supplier S, a discount
applied to the net price at Supplier A), the specific extractor stays more
accurate. The generic one is there to cover the rest quickly.

Main precaution: never mistake a selling price for a purchase price. The
"prix public", "PVC", "tarif conseille" and "location" columns are
explicitly excluded.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd

from extracteurs.base import (
    calculer_prix_colis,
    chemin_lisible,
    economie_palier,
    finaliser,
    nettoyer_ean,
    nettoyer_nombre,
    nettoyer_texte,
    normaliser_conditionnement,
    resumer_paliers,
    valider_paliers,
)

# --- column vocabulary ------------------------------------------------------

REFERENCE = [
    "referencefournisseur", "referencesfournisseur", "reffournisseur",
    "referencearticle", "referencefabricant", "codearticle", "codeproduit",
    "reference", "references", "ref", "refart", "code", "codes", "article",
    "codeean",  # some price lists only have the EAN as a reference
]

DESIGNATION = [
    "designation", "designations", "libelle", "libelles", "denomination",
    "produit", "produits", "description", "nomduproduit", "intitule",
    "articles", "nom",
]

EAN = ["ean", "gtin", "codeean", "codegtin", "codebarre", "eancode", "ean13"]

# What, in a heading, designates the CASE rather than the piece: "EAN
# carton", "GTIN IUD (CDT)", "EAN Boîte". Supplier CV goes as far as
# three levels: unit, box, carton.
COLIS = ["carton", "cdt", "conditionnement", "colis", "palette", "caisse",
         "boite", "outer"]

# Headings that designate our purchase price UNAMBIGUOUSLY, even when
# they contain an otherwise suspect word. "Prix de vente client" is the
# price at which the supplier sells to US: that really is our purchase
# price, whereas the exclusion filter below would reject the word
# "vente".
PRIX_ACHAT_EXPLICITE = [
    "prixdeventeclient", "prixventeclient", "votretarif", "votreprix",
    "tarifnet", "prixnet", "prixdachat", "prixachat", "franco",
    # "Prix Vente REMISE" (18/09): the word "remise" (discount) lifts the
    # ambiguity that "vente" (sale) introduces. A column that says REMISE
    # is never a list price: it is the price after our negotiated
    # discount, hence ours.
    #
    # Without this entry, Supplier CX's price list yielded the "Prix
    # base" column: the cooler bags (REF0001 to REF0004) went out at
    # 22.00 EUR while the column next to it says 15.00, and 15.00 is both
    # the pricing manager's former purchase price AND the price actually
    # paid on order. Three sources agreeing against one.
    # (references and amounts above: sample values)
    #
    # We do NOT add plain "prixvente": at another supplier that would be
    # the resale price, and the exclusion filter is right to distrust it.
    # It is the word "remise" that licenses the exception.
    "prixventeremise", "prixdeventeremise", "prixventeremis",
]

# Purchase price: what we are looking for
PRIX_ACHAT = [
    "prixachat", "prixdachat", "tarifnet", "prixnet", "netht", "prixnetht",
    "tarifachat", "prixremise", "prixcession", "achatht", "prixhtnet",
    "votretarif", "tarifs", "tarif", "prixht", "prix",
]

# A per-case price is not a unit price: taking it multiplies the purchase
# price by the pack size, which has already produced a x109 gap.
AU_COLIS = ["carton", "palette", "colis", "boiteau", "parlot", "aulot",
            "sachetde"]

# What designates the price actually paid per unit: that one, or the
# discounted unit price, never the list price.
A_LUNITE = ["unitaire", "alunite", "lunite", "unitee", "remise", "net"]

# Selling price: what must on no account be taken for a purchase price
PRIX_EXCLUS = [
    "public", "pvc", "conseille", "conseil", "ttc", "vente", "revente",
    "location", "loyer", "lpp", "lppr", "remboursement", "psl", "pvp",
    "constate", "detail",
]

TVA = ["tauxtva", "tva", "codetva"]
ECO_PART = ["ecoparticipation", "ecopart", "ecotaxe", "deee", "d3e"]
CONDITIONNEMENT = [
    "conditionnement", "colisage", "uv", "unitedevente", "parcarton",
    "qtecarton", "nbparcolis", "pcb",
]
CODE_LPPR = ["codelppr", "codelpp", "lpprindividuel", "nouveaucodelppr"]

# A tier quantity column: "x 10", "par 12", "a partir de 6"
QUANTITE = re.compile(r"(?:x|par|apartirde|des)\s*(\d+)")


def _normaliser(valeur) -> str:
    """Lower case, without accents, spaces or punctuation."""
    if valeur is None:
        return ""
    texte = unicodedata.normalize("NFKD", str(valeur))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texte.lower())


def _correspond(normalise: str, vocabulaire: list[str]) -> bool:
    """Does the heading start with one of the vocabulary terms?"""
    return any(normalise.startswith(terme) for terme in vocabulaire)


def _contient(normalise: str, termes: list[str]) -> bool:
    return any(terme in normalise for terme in termes)


def trouver_entete(brut: pd.DataFrame, max_lignes: int = 25) -> int | None:
    """Index of the header row, or None if none is recognised.

    A row makes a credible header when it carries at least a label and
    either a reference or a price. We keep the first one, failing that
    the one that recognises the most columns.
    """
    meilleur, score_max = None, 0

    for position in range(min(max_lignes, len(brut))):
        cellules = [_normaliser(v) for v in brut.iloc[position].tolist()]
        cellules = [c for c in cellules if c]
        if len(cellules) < 2:
            continue

        a_designation = any(_correspond(c, DESIGNATION) for c in cellules)
        a_reference = any(_correspond(c, REFERENCE) for c in cellules)
        a_prix = any(
            _correspond(c, PRIX_ACHAT) and not _contient(c, PRIX_EXCLUS)
            for c in cellules
        )
        a_ean = any(_contient(c, EAN) for c in cellules)

        if not a_designation or not (a_reference or a_prix):
            continue

        score = sum((a_designation, a_reference, a_prix, a_ean))
        if score > score_max:
            meilleur, score_max = position, score

    return meilleur


def analyser_colonnes(entetes) -> dict:
    """Maps each schema field to a column index.

    The purchase price is chosen with care: among the price columns that
    were not excluded, the one with the most explicit heading wins
    ("tarif net HT" before "prix"). The others become tiers.
    """
    mapping = {}
    prix_candidats = []

    for index, entete in enumerate(entetes):
        normalise = _normaliser(entete)
        if not normalise:
            continue

        if _contient(normalise, EAN):
            # Many price lists publish two codes: one for the piece and
            # one for the case ("EAN unité" / "EAN carton", "GTIN IUD
            # (Unité)" / "(CDT)"). The first is the one scanned at
            # picking, and it is the one `ean` carries; the second is
            # kept separately for goods-in.
            if _contient(normalise, COLIS):
                mapping.setdefault("ean_conditionnement", index)
                continue
            if "ean" not in mapping:
                mapping["ean"] = index
                continue
            # A second code with no mention of packaging level: keep it
            # anyway rather than lose it.
            mapping.setdefault("ean_conditionnement", index)
            continue
        if _correspond(normalise, CODE_LPPR) and "code_lppr" not in mapping:
            mapping["code_lppr"] = index
            continue
        if _correspond(normalise, TVA) and "tva_taux" not in mapping:
            mapping["tva_taux"] = index
            continue
        if _correspond(normalise, ECO_PART) and "eco_part_ht" not in mapping:
            mapping["eco_part_ht"] = index
            continue
        if (_correspond(normalise, CONDITIONNEMENT)
                and "conditionnement" not in mapping):
            mapping["conditionnement"] = index
            continue
        if _correspond(normalise, DESIGNATION) and "designation" not in mapping:
            mapping["designation"] = index
            continue

        explicite = _correspond(normalise, PRIX_ACHAT_EXPLICITE)
        if explicite or _correspond(normalise, PRIX_ACHAT):
            # An explicit heading escapes the exclusion filter
            if not explicite and _contient(normalise, PRIX_EXCLUS):
                continue  # selling price, rental, LPP: not a purchase price
            # The more explicit the heading, the higher its priority
            priorite = next(
                (len(PRIX_ACHAT) - rang
                 for rang, terme in enumerate(PRIX_ACHAT)
                 if normalise.startswith(terme)),
                0,
            )
            if explicite:
                priorite += 100

            # Purchasing rule, given by the buyers: the price to keep is
            # the UNIT selling price or the unit DISCOUNTED price, never
            # the list price and never a per-case price. Supplier CV
            # publishes "prix indicatif de la boite au Carton" right next
            # to "prix indicatif de l'unitee a la boite": both carry the
            # same keywords, only this weighting separates them.
            if any(m in normalise for m in AU_COLIS):
                priorite -= 50
            if any(m in normalise for m in A_LUNITE):
                priorite += 30

            quantite = 1.0
            correspondance = QUANTITE.search(normalise)
            if correspondance:
                quantite = float(correspondance.group(1))
            prix_candidats.append((index, quantite, priorite))
            continue

        if _correspond(normalise, REFERENCE) and "ref_fournisseur" not in mapping:
            mapping["ref_fournisseur"] = index
            continue

    # The unit price is the most explicit one at quantity 1; columns with
    # a higher quantity become the tiers.
    prix_candidats.sort(key=lambda c: (c[1], -c[2]))
    mapping["_prix"] = prix_candidats
    return mapping


def extract(path, fournisseur: str | None = None) -> "pd.DataFrame":  # noqa: F821
    """Extracts any Excel price list into the common schema.

    `fournisseur` names the source; failing that the parent folder name
    is used.
    """
    path = Path(path)
    nom_fournisseur = fournisseur or path.parent.name.upper()

    feuilles = pd.read_excel(
        chemin_lisible(path), dtype=str, header=None, sheet_name=None
    )

    lignes: list[dict] = []
    for onglet, brut in feuilles.items():
        if brut.empty:
            continue
        position = trouver_entete(brut)
        if position is None:
            continue

        entetes = brut.iloc[position].tolist()
        mapping = analyser_colonnes(entetes)
        prix_candidats = mapping.pop("_prix", [])

        # The name of the chosen column travels with the price. Without
        # it there is no way to check afterwards that we really took the
        # discounted price and not the list price: it can only be
        # assumed, and an assumption about a price column is paid for in
        # wrong purchase prices.
        colonne_prix = ""
        if prix_candidats:
            index = prix_candidats[0][0]
            if index < len(entetes):
                colonne_prix = str(entetes[index]).strip()

        # With neither price nor EAN, the sheet brings nothing
        if not prix_candidats and "ean" not in mapping:
            continue
        if "designation" not in mapping:
            continue

        for _, valeurs in brut.iloc[position + 1:].iterrows():
            def cellule(champ):
                index = mapping.get(champ)
                return None if index is None else valeurs.iloc[index]

            designation = nettoyer_texte(cellule("designation"))
            ref = nettoyer_texte(cellule("ref_fournisseur"))
            ean = nettoyer_ean(cellule("ean"))

            # A row with no label, or with no identifier at all, is a
            # section title or a layout row.
            if not designation or not (ref or ean):
                continue

            tarifs = [
                (quantite, nettoyer_nombre(valeurs.iloc[index]))
                for index, quantite, _ in prix_candidats
            ]
            tarifs = [(q, p) for q, p in tarifs if p is not None and p > 0]
            if not tarifs:
                continue

            # The best scoring column gives the unit price. The others
            # only become tiers if they carry a quantity above 1 AND a
            # lower price: otherwise they are other price columns (list
            # price, a different discount), not tiers.
            prix_unitaire = tarifs[0][1]
            paliers = valider_paliers(prix_unitaire, tarifs[1:])

            ligne = {
                "fournisseur": nom_fournisseur,
                "ref_fournisseur": ref or ean,
                "designation": designation,
                "ean": ean,
                "ean_conditionnement": nettoyer_ean(
                    cellule("ean_conditionnement")),
                "prix_achat_unitaire_ht": prix_unitaire,
                "colonne_prix": colonne_prix,
                "eco_part_ht": nettoyer_nombre(cellule("eco_part_ht")),
                "tva_taux": nettoyer_nombre(cellule("tva_taux")),
                "code_lppr": nettoyer_texte(cellule("code_lppr")),
                "montant_lppr": None,
                "tarif_public_ttc": None,
                "remise_taux": None,
                "origine": None,
                "dispositif_medical": None,
                "fichier_source": path.name,
                "palier2_qte": paliers[0][0] if len(paliers) > 0 else None,
                "palier2_prix_ht": paliers[0][1] if len(paliers) > 0 else None,
                "palier3_qte": paliers[1][0] if len(paliers) > 1 else None,
                "palier3_prix_ht": paliers[1][1] if len(paliers) > 1 else None,
                "prix_recalcule": None,
                "ecart_prix": None,
                "prix_coherent": None,
            }

            (
                ligne["conditionnement"],
                ligne["conditionnement_suppose"],
            ) = normaliser_conditionnement(cellule("conditionnement"))
            ligne["prix_colis_ht"] = calculer_prix_colis(
                prix_unitaire, ligne["conditionnement"]
            )
            ligne["economie_palier2_pct"] = economie_palier(
                prix_unitaire, ligne["palier2_prix_ht"]
            )
            ligne["economie_palier3_pct"] = economie_palier(
                prix_unitaire, ligne["palier3_prix_ht"]
            )
            ligne["paliers"] = resumer_paliers(
                prix_unitaire,
                [
                    (ligne["palier2_qte"], ligne["palier2_prix_ht"]),
                    (ligne["palier3_qte"], ligne["palier3_prix_ht"]),
                ],
            )
            lignes.append(ligne)

    return finaliser(lignes)
