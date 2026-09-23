# -*- coding: utf-8 -*-
"""
Extracteur des tarifs dont les paliers degressifs sont eclates en LIGNES.

Plusieurs fournisseurs publient leur grille ainsi : une ligne par couple
(reference, quantite), le prix variant d'une ligne a l'autre.

    FOURNISSEUR P   REFERENCE ARTICLE | DESIGNATION    | QUANTITE | PAHT
                    REF0001           | PRODUIT EXEMPLE|        1 | 9.00
                    REF0001           | PRODUIT EXEMPLE|        6 | 7.50
                    REF0001           | PRODUIT EXEMPLE|       12 | 7.00

    FOURNISSEUR L   Référence | Désignation     | Qté | Prix de vente
                    REF0002   | PRODUIT EXEMPLE |  20 | 5.00
                    REF0002   | PRODUIT EXEMPLE |  80 | 4.50

    (references, libelles et montants ci-dessus : valeurs d'exemple)

Le traitement est le meme dans tous les cas : regrouper par reference, la
plus petite quantite donnant le prix unitaire et les suivantes les paliers.

Deux precautions :
  - la plus petite quantite publiee n'est pas toujours 1 (le Fournisseur L
    commence a 20 sur certaines references) : elle devient alors le
    conditionnement,
    et le prix reste un prix unitaire ;
  - a quantite egale, le tarif le plus bas l'emporte, certains fichiers
    repetant une meme ligne avec des conditions differentes.
"""

from __future__ import annotations

import re
import unicodedata
import warnings
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

# Vocabulaire des colonnes, teste sur le debut de l'intitule normalise.
# L'ordre compte : les intitules les plus longs sont essayes en premier.
VOCABULAIRE = {
    "ref_fournisseur": [
        "referencearticle", "referencefournisseur", "reference", "ref",
        "codearticle", "code",
    ],
    "designation": [
        "designationarticle", "designation", "libelle", "produit",
    ],
    "quantite": ["quantite", "qte", "qty"],
    "prix_achat_unitaire_ht": [
        "paht", "prixachatht", "prixachat", "prixdevente", "tarifnet",
        "prixnet", "prixunitaire", "prix", "tarif",
    ],
    "ean": ["codeean", "ean", "gtin"],
    "tva_taux": ["tauxtva", "tva"],
    "code_lppr": ["codelppr", "codelpp"],
    "montant_lppr": ["montantlpp", "lppht", "tariflppr", "lpp"],
    "conditionnement": ["uniteminiemballage", "uniteminiembalage",
                        "conditionnement", "colisage", "pcb"],
    "eco_part_ht": ["montantecoparticipation", "ecoparticipation",
                    "ecopart", "ecotaxe"],
}

# Colonnes de prix a ne jamais confondre avec un prix d'achat
PRIX_EXCLUS = ["public", "pvc", "conseille", "ttc", "location", "remise"]


def _normaliser(valeur) -> str:
    if valeur is None:
        return ""
    texte = unicodedata.normalize("NFKD", str(valeur))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texte.lower())


def _trouver_entete(brut: pd.DataFrame, max_lignes: int = 20) -> int | None:
    """Ligne portant a la fois une reference, une designation et un prix."""
    for position in range(min(max_lignes, len(brut))):
        cellules = [_normaliser(v) for v in brut.iloc[position].tolist()]
        cellules = [c for c in cellules if c]
        if len(cellules) < 3:
            continue

        def porte(champ):
            return any(
                c.startswith(terme)
                for c in cellules
                for terme in VOCABULAIRE[champ]
            )

        if porte("ref_fournisseur") and porte("designation") and porte(
                "prix_achat_unitaire_ht"):
            return position
    return None


def _analyser_colonnes(entetes) -> dict:
    """Associe chaque champ a un index de colonne."""
    mapping = {}
    for index, entete in enumerate(entetes):
        normalise = _normaliser(entete)
        if not normalise:
            continue
        for champ, termes in VOCABULAIRE.items():
            if champ in mapping:
                continue
            for terme in sorted(termes, key=len, reverse=True):
                if not normalise.startswith(terme):
                    continue
                if champ == "prix_achat_unitaire_ht" and any(
                        exclu in normalise for exclu in PRIX_EXCLUS):
                    break
                mapping[champ] = index
                break
            if champ in mapping:
                break
    return mapping


def extract(path, fournisseur: str | None = None,
            onglet: str | None = None) -> "pd.DataFrame":  # noqa: F821
    """Extrait un tarif a paliers en lignes vers le schema commun."""
    path = Path(path)
    nom = fournisseur or path.parent.name.upper()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        lisible = chemin_lisible(path)
        try:
            classeur = pd.ExcelFile(lisible)
        except Exception:
            # Certains classeurs portent une feuille de style qu'openpyxl
            # refuse (Fournisseur BF) : calamine les lit sans broncher.
            classeur = pd.ExcelFile(lisible, engine="calamine")
        onglets = [onglet] if onglet else classeur.sheet_names

        produits: dict[str, dict] = {}
        for feuille in onglets:
            brut = pd.read_excel(classeur, sheet_name=feuille, dtype=str,
                                 header=None)
            if brut.empty:
                continue
            position = _trouver_entete(brut)
            if position is None:
                continue

            mapping = _analyser_colonnes(brut.iloc[position].tolist())
            if "ref_fournisseur" not in mapping or \
                    "prix_achat_unitaire_ht" not in mapping:
                continue

            def valeur(ligne, champ):
                index = mapping.get(champ)
                return None if index is None else ligne.iloc[index]

            for _, ligne in brut.iloc[position + 1:].iterrows():
                ref = nettoyer_texte(valeur(ligne, "ref_fournisseur"))
                designation = nettoyer_texte(valeur(ligne, "designation"))
                prix = nettoyer_nombre(valeur(ligne, "prix_achat_unitaire_ht"))
                if not ref or not designation or prix is None or prix <= 0:
                    continue

                quantite = nettoyer_nombre(valeur(ligne, "quantite")) or 1.0

                produit = produits.setdefault(ref, {
                    "designation": designation,
                    "tarifs": {},
                    "ean": nettoyer_ean(valeur(ligne, "ean")),
                    "tva_taux": nettoyer_nombre(valeur(ligne, "tva_taux")),
                    "code_lppr": nettoyer_texte(valeur(ligne, "code_lppr")),
                    "montant_lppr": nettoyer_nombre(
                        valeur(ligne, "montant_lppr")),
                    "eco_part_ht": nettoyer_nombre(
                        valeur(ligne, "eco_part_ht")),
                    "conditionnement": nettoyer_nombre(
                        valeur(ligne, "conditionnement")),
                })
                ancien = produit["tarifs"].get(quantite)
                if ancien is None or prix < ancien:
                    produit["tarifs"][quantite] = prix

    lignes = []
    for ref, produit in produits.items():
        tarifs = sorted(produit["tarifs"].items())
        quantite_base, prix_unitaire = tarifs[0]
        # Un tarif qui augmente avec la quantite n'est pas un palier :
        # le Fournisseur P publie ainsi des variantes plus cheres sur la
        # meme reference. On ne retient que la degressivite reelle.
        paliers = valider_paliers(prix_unitaire, tarifs[1:])

        ligne = {
            "fournisseur": nom,
            "ref_fournisseur": ref,
            "designation": produit["designation"],
            "ean": produit["ean"],
            "prix_achat_unitaire_ht": prix_unitaire,
            "tva_taux": produit["tva_taux"],
            "code_lppr": produit["code_lppr"],
            "montant_lppr": produit["montant_lppr"],
            "eco_part_ht": produit["eco_part_ht"],
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

        # Quand le tarif publie un conditionnement, il prime ; sinon la plus
        # petite quantite tarifee en tient lieu.
        conditionnement = produit["conditionnement"] or quantite_base
        ligne["conditionnement"], ligne["conditionnement_suppose"] = (
            normaliser_conditionnement(conditionnement)
        )
        ligne["prix_colis_ht"] = calculer_prix_colis(
            prix_unitaire, ligne["conditionnement"]
        )
        ligne["economie_palier2_pct"] = economie_palier(
            prix_unitaire, ligne["palier2_prix_ht"])
        ligne["economie_palier3_pct"] = economie_palier(
            prix_unitaire, ligne["palier3_prix_ht"])
        ligne["paliers"] = resumer_paliers(
            prix_unitaire,
            [
                (ligne["palier2_qte"], ligne["palier2_prix_ht"]),
                (ligne["palier3_qte"], ligne["palier3_prix_ht"]),
            ],
            quantite_base=quantite_base,
        )
        lignes.append(ligne)

    return finaliser(lignes)
