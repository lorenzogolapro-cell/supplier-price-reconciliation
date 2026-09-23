# -*- coding: utf-8 -*-
"""
Extracteur generique pour les tarifs Excel.

Les fournisseurs ne partagent aucune mise en page, mais ils nomment tous
leurs colonnes de la meme facon : une reference, un libelle, un prix, et
parfois un code EAN. Ce module repere ces colonnes par leur intitule plutot
que par leur position, ce qui evite d'ecrire un extracteur par fournisseur.

Il ne remplace pas les extracteurs dedies : quand un tarif a une logique
propre (paliers nommes dans l'en-tete chez le Fournisseur S, remise
appliquee au HT chez le Fournisseur A), l'extracteur specifique reste plus
juste. Le generique
sert a couvrir rapidement le reste.

Precaution principale : ne jamais confondre un prix de vente avec un prix
d'achat. Les colonnes "prix public", "PVC", "tarif conseille" ou "location"
sont explicitement exclues.
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

# --- vocabulaire des colonnes ---------------------------------------------

REFERENCE = [
    "referencefournisseur", "referencesfournisseur", "reffournisseur",
    "referencearticle", "referencefabricant", "codearticle", "codeproduit",
    "reference", "references", "ref", "refart", "code", "codes", "article",
    "codeean",  # certains tarifs n'ont que l'EAN comme reference
]

DESIGNATION = [
    "designation", "designations", "libelle", "libelles", "denomination",
    "produit", "produits", "description", "nomduproduit", "intitule",
    "articles", "nom",
]

EAN = ["ean", "gtin", "codeean", "codegtin", "codebarre", "eancode", "ean13"]

# Ce qui, dans l'intitule, designe le COLIS et non la piece : "EAN
# carton", "GTIN IUD (CDT)", "EAN Boîte". Le Fournisseur CV va jusqu'a
# trois niveaux — unite, boite, carton.
COLIS = ["carton", "cdt", "conditionnement", "colis", "palette", "caisse",
         "boite", "outer"]

# Intitules qui designent SANS AMBIGUITE notre prix d'achat, meme s'ils
# contiennent un mot par ailleurs suspect. "Prix de vente client" est le
# prix auquel le fournisseur NOUS vend : c'est bien notre prix d'achat,
# alors que le filtre d'exclusion ci-dessous ecarterait le mot "vente".
PRIX_ACHAT_EXPLICITE = [
    "prixdeventeclient", "prixventeclient", "votretarif", "votreprix",
    "tarifnet", "prixnet", "prixdachat", "prixachat", "franco",
    # « Prix Vente REMISE » (18/09) — le mot « remise » leve l'ambiguite que
    # « vente » introduit. Une colonne qui dit REMISE n'est jamais un prix
    # public : c'est le prix apres notre remise negociee, donc le notre.
    #
    # Sans cette entree, le tarif du Fournisseur CX sortait le « Prix
    # base » : les sacs isothermes (REF0001 a REF0004) partaient a
    # 22,00 EUR alors que la colonne d'a cote annonce 15,00 — et 15,00 est
    # a la fois l'ancien PA du responsable des prix ET le prix reellement
    # paye en commande. Trois sources d'accord contre une.
    # (references et montants ci-dessus : valeurs d'exemple)
    #
    # On n'ajoute PAS « prixvente » tout court : chez un autre fournisseur
    # ce serait le prix de revente, et le filtre d'exclusion a raison de
    # s'en mefier. C'est le mot « remise » qui autorise l'exception.
    "prixventeremise", "prixdeventeremise", "prixventeremis",
]

# Prix d'achat : ce qu'on cherche
PRIX_ACHAT = [
    "prixachat", "prixdachat", "tarifnet", "prixnet", "netht", "prixnetht",
    "tarifachat", "prixremise", "prixcession", "achatht", "prixhtnet",
    "votretarif", "tarifs", "tarif", "prixht", "prix",
]

# Un prix au colis n'est pas un prix unitaire : le retenir multiplie le
# PA par le conditionnement, ce qui a deja produit un ecart x109.
AU_COLIS = ["carton", "palette", "colis", "boiteau", "parlot", "aulot",
            "sachetde"]

# Ce qui designe le prix reellement paye a l'unite : c'est celui-la, ou
# le prix remise unitaire, jamais le prix public.
A_LUNITE = ["unitaire", "alunite", "lunite", "unitee", "remise", "net"]

# Prix de vente : ce qu'il ne faut surtout pas prendre pour un prix d'achat
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

# Une colonne de quantite de palier : "x 10", "par 12", "a partir de 6"
QUANTITE = re.compile(r"(?:x|par|apartirde|des)\s*(\d+)")


def _normaliser(valeur) -> str:
    """Minuscule sans accent, espace ni ponctuation."""
    if valeur is None:
        return ""
    texte = unicodedata.normalize("NFKD", str(valeur))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texte.lower())


def _correspond(normalise: str, vocabulaire: list[str]) -> bool:
    """L'intitule commence-t-il par l'un des termes du vocabulaire ?"""
    return any(normalise.startswith(terme) for terme in vocabulaire)


def _contient(normalise: str, termes: list[str]) -> bool:
    return any(terme in normalise for terme in termes)


def trouver_entete(brut: pd.DataFrame, max_lignes: int = 25) -> int | None:
    """Indice de la ligne d'en-tete, ou None si aucune n'est reconnue.

    Une ligne fait un en-tete credible quand elle porte au moins une
    designation et une reference ou un prix. On retient la premiere, et a
    defaut celle qui reconnait le plus de colonnes.
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
    """Associe chaque champ du schema a un index de colonne.

    Le prix d'achat est choisi avec soin : parmi les colonnes de prix non
    exclues, celle dont l'intitule est le plus explicite l'emporte
    ("tarif net HT" avant "prix"). Les autres deviennent des paliers.
    """
    mapping = {}
    prix_candidats = []

    for index, entete in enumerate(entetes):
        normalise = _normaliser(entete)
        if not normalise:
            continue

        if _contient(normalise, EAN):
            # Beaucoup de tarifs publient deux codes : celui de la piece
            # et celui du colis (« EAN unité » / « EAN carton », « GTIN
            # IUD (Unité) » / « (CDT) »). Le premier est celui qu'on
            # scanne au picking, et c'est lui que porte `ean` ; le second
            # est garde a part pour la reception.
            if _contient(normalise, COLIS):
                mapping.setdefault("ean_conditionnement", index)
                continue
            if "ean" not in mapping:
                mapping["ean"] = index
                continue
            # Un second code sans mention de colisage : on le garde
            # quand meme plutot que de le perdre.
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
            # Un intitule explicite echappe au filtre d'exclusion
            if not explicite and _contient(normalise, PRIX_EXCLUS):
                continue  # prix de vente, location, LPP : pas un prix d'achat
            # Plus l'intitule est explicite, plus il est prioritaire
            priorite = next(
                (len(PRIX_ACHAT) - rang
                 for rang, terme in enumerate(PRIX_ACHAT)
                 if normalise.startswith(terme)),
                0,
            )
            if explicite:
                priorite += 100

            # Regle d'achat, donnee par les appros : le prix a retenir est
            # le prix UNITAIRE de vente ou le prix REMISE unitaire, jamais
            # le prix public et jamais un prix au colis. Le Fournisseur CV
            # publie « prix indicatif de la boite au Carton » juste a cote de
            # « prix indicatif de l'unitee a la boite » : les deux portent
            # les memes mots-cles, seule cette ponderation les separe.
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

    # Le prix unitaire est celui de quantite 1 le plus explicite ;
    # les colonnes de quantite superieure deviennent les paliers.
    prix_candidats.sort(key=lambda c: (c[1], -c[2]))
    mapping["_prix"] = prix_candidats
    return mapping


def extract(path, fournisseur: str | None = None) -> "pd.DataFrame":  # noqa: F821
    """Extrait un tarif Excel quelconque vers le schema commun.

    `fournisseur` permet de nommer la source ; a defaut le nom du dossier
    parent est utilise.
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

        # Le nom de la colonne retenue voyage avec le prix. Sans lui, on
        # ne peut pas verifier apres coup qu'on a bien pris le prix
        # remise et non le tarif public : on ne peut que le supposer, et
        # une supposition sur une colonne de prix se paie en PA faux.
        colonne_prix = ""
        if prix_candidats:
            index = prix_candidats[0][0]
            if index < len(entetes):
                colonne_prix = str(entetes[index]).strip()

        # Sans prix ni EAN, la feuille n'apporte rien
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

            # Une ligne sans libelle, ou sans aucun identifiant, est un
            # titre de rubrique ou une ligne de mise en page.
            if not designation or not (ref or ean):
                continue

            tarifs = [
                (quantite, nettoyer_nombre(valeurs.iloc[index]))
                for index, quantite, _ in prix_candidats
            ]
            tarifs = [(q, p) for q, p in tarifs if p is not None and p > 0]
            if not tarifs:
                continue

            # La colonne la mieux notee donne le prix unitaire. Les autres
            # ne deviennent des paliers que si elles portent une quantite
            # superieure a 1 ET un prix inferieur : sinon ce sont d'autres
            # colonnes de prix (public, remise differente), pas des paliers.
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
