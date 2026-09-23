# -*- coding: utf-8 -*-
"""
Socle commun aux lecteurs d'accuses de reception fournisseurs.

Un AR dit le prix que le fournisseur s'engage a facturer, pour une
quantite donnee, a une date donnee. C'est la seule source qui soit un
FAIT : un tarif n'est qu'une intention, et le PA du WMS n'est que ce
que quelqu'un a bien voulu y ressaisir.

Deux consequences pour le schema ci-dessous :

  - on garde TOUJOURS la quantite a cote du prix. Un meme article se
    negocie 9,00 EUR par 5 et 7,50 EUR par 20 (valeurs d'exemple) :
    comparer un prix d'AR a un PA unique n'a de sens qu'a quantite
    comparable. C'est la lecon la plus couteuse de ce chantier.
  - on garde le prix affiche ET le prix net. Le premier sert a relire
    l'AR, le second est celui qu'on paie.

Chaque lecteur (ar/<fournisseur>.py) expose :
    extract(chemin) -> pandas.DataFrame
respectant COLONNES, dans cet ordre.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd
import pdfplumber

# ---------------------------------------------------------------------------
# Schema de sortie normalise
# ---------------------------------------------------------------------------
COLONNES = [
    "fournisseur",        # nom du fournisseur (ex: FOURNISSEUR L)
    "fichier_ar",         # PDF d'origine, pour pouvoir remonter a la preuve
    "num_ar",             # numero d'accuse cote fournisseur
    "commande_wms",       # notre numero de commande ("Votre reference")
    "ref_fournisseur",    # reference article du fournisseur (TOUJOURS du texte)
    "designation_ar",     # libelle tel qu'il figure sur l'AR
    "quantite",           # quantite commandee — indissociable du prix
    "prix_unitaire_ar",   # prix unitaire affiche, avant remise
    "remise_pct",         # taux de remise en POURCENTAGE (25.0 = 25 %)
    "prix_net_ar",        # prix reellement paye : unitaire moins la remise
    "montant_ligne",      # total de la ligne, quand l'AR le publie
    "delai_livraison",    # date de livraison annoncee, quand elle figure
    # DATE DE COMMANDE, telle que l'AR l'ecrit — pas la date de reception
    # du mail (celle-la est dans index_ar.csv et peut en differer).
    #
    # Ajoutee le 21/09/2026, et elle n'est pas un confort : UN AR NE
    # PROUVE RIEN HORS DE SA DATE. Chez le fournisseur L, confronter le
    # tarif de l'annee a l'ensemble des AR donnait un rapport de 1,12 —
    # de quoi croire qu'on surevaluait tout le catalogue de 12 %. Reparti
    # par mois, le rapport vaut 1,1200 jusqu'en fevrier puis 1,0000 a
    # partir de mars : ce n'etait pas une erreur, c'etait la hausse
    # annuelle, et nos prix etaient justes. Sans la date, on corrigeait
    # a l'envers.  (ratios et mois de bascule : valeurs d'exemple)
    #
    # Meme lecon chez le fournisseur E, ou une seconde remise apparait
    # en cours d'annee, et chez le fournisseur G, ou la hausse tombe a
    # un autre mois : le mois de bascule est propre a chaque fournisseur.
    "date_commande",
]


def _sans_invisibles(texte: str) -> str:
    """Retire les caracteres de formatage invisibles.

    Le fournisseur H seme des traits d'union conditionnels (U+00AD) au
    milieu de ses references. Ils ne se voient pas, s'impriment comme rien, et font
    echouer tout rapprochement sans qu'on comprenne pourquoi.
    """
    normalise = unicodedata.normalize("NFKC", texte)
    return "".join(c for c in normalise
                   if unicodedata.category(c) != "Cf")


def nettoyer_texte(valeur) -> str:
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return ""
    return " ".join(_sans_invisibles(str(valeur)).split())


def nombre(texte) -> float | None:
    """"1 111,36" -> 1111.36 ; les AR ecrivent a la francaise.

    L'espace insecable est le piege habituel des PDF : il ressemble a une
    espace mais n'en est pas une, et float() s'en etrangle.
    """
    if texte is None:
        return None
    brut = _sans_invisibles(str(texte))
    brut = brut.replace(" ", "").replace(" ", "")
    brut = brut.replace(" ", "").replace(",", ".")
    if not brut:
        return None
    try:
        return float(brut)
    except ValueError:
        return None


def prix_net(unitaire: float | None, remise: float | None) -> float | None:
    """Prix reellement paye. Une remise est toujours prise en valeur
    absolue : les AR l'ecrivent tantot "25%", tantot "-25%"."""
    if unitaire is None:
        return None
    if remise is None or remise == 0:
        return round(unitaire, 4)
    return round(unitaire * (1 - abs(remise) / 100), 4)


def texte_pdf(chemin) -> str:
    """Texte integral du PDF. Aucun AR rencontre ne demande d'OCR."""
    with pdfplumber.open(Path(chemin)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def premier_groupe(motif: re.Pattern, texte: str) -> str:
    trouve = motif.search(texte)
    return trouve.group(1) if trouve else ""


def cadrer(lignes: list[dict], fournisseur: str, chemin) -> pd.DataFrame:
    """Met les lignes lues au schema commun, colonnes manquantes comprises."""
    df = pd.DataFrame(lignes)
    if df.empty:
        return pd.DataFrame(columns=COLONNES)
    df["fournisseur"] = fournisseur
    df["fichier_ar"] = Path(chemin).name
    for colonne in COLONNES:
        if colonne not in df.columns:
            df[colonne] = None
    return df[COLONNES]
