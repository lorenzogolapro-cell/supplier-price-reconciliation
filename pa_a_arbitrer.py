# -*- coding: utf-8 -*-
"""
Chantier PA — tout ce qui attend une décision humaine, en un seul classeur.

    python pa_a_arbitrer.py

POURQUOI
    Les questions ouvertes se sont retrouvées éparpillées dans quatre
    fichiers, chacun produit par un outil différent. Personne ne peut
    arbitrer en ouvrant quatre classeurs et en recoupant à la main.

    Celui-ci ne calcule rien : il rassemble ce qui a déjà été produit et le
    présente dans l'ordre où les décisions doivent être prises. La première
    feuille se lit seule, et dit ce qu'il faut trancher.

CE QU'IL NE FAIT PAS
    Il n'applique aucun prix et n'écrit dans aucun classeur de travail. Tant
    qu'une question n'est pas tranchée, la réponse reste « on ne fait pas ».
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

SORTIE = RACINE / "sortie" / "3-achats"

# Les questions de fond, celles qui ne se règlent pas ligne à ligne. Elles
# valent plus que les listes : une seule réponse débloque des dizaines
# d'articles d'un coup.
QUESTIONS = [
    {
        "Sujet": "FOURNISSEUR O — quelle colonne ?",
        # Valeurs d'exemple : le trio tarif / remise / prix net est reproduit
        # dans ses proportions, pas dans ses montants reels.
        "Question": "Le tarif porte « Tarif de Base HT » (82,50), « Remise » "
                    "(0,1800) et « Prix HT boîte Revendeurs » (67,65). Nous "
                    "lisons la première, le responsable la troisième.",
        "Ce qui est établi": "82,50 x (1 - 0,1800) = 67,65 au centime. La "
                             "règle d'achat dit « prix remisé unitaire », donc "
                             "le responsable a raison et c'est notre lecture "
                             "qui est à corriger.",
        "Enjeu": "2 articles en divergence, mais la colonne vaut pour tout "
                 "le tarif du Fournisseur O",
        "Décision attendue": "Confirmer qu'on bascule sur « Prix HT boîte "
                             "Revendeurs »",
    },
    {
        "Sujet": "FOURNISSEUR I — unitaire ou franco ?",
        "Question": "Deux colonnes : « TARIF UNITAIRE » (le prix d'une "
                    "casaque) et « FRANCO » (le prix du carton de 40).",
        "Ce qui est établi": "1,2500 x 40 = 50,00 : les deux sont justes, ils "  # valeurs d'exemple
                             "ne répondent pas à la même question. Sur 104 "
                             "articles, 88 collent au FRANCO et 16 au TARIF "
                             "UNITAIRE — et ces 16 portent tous « UNITE » ou "
                             "« UNITAIRE » dans leur libellé WMS.",
        "Enjeu": "16 articles",
        "Décision attendue": "Valider la règle : libellé WMS « UNITE » -> "
                             "tarif unitaire, sinon franco",
    },
    {
        "Sujet": "FOURNISSEUR I — quel fichier fait foi ?",
        "Question": "Trois tarifs 2026 coexistent. « TARIF REVENDEUR "
                    "PRIVILEGE JUILLET 2026 » donne 5,00 pour une seringue, "  # valeurs d'exemple
                    "« TARIFREVENDEUR2026 » donne 5,60.",
        "Ce qui est établi": "Le DPA actuel est 4,850, donc plus proche de "  # valeurs d'exemple
                             "5,00 (x1,03) que de 5,60 (x1,15). Aujourd'hui "
                             "c'est l'ordre alphabétique qui tranche, ce qui "
                             "n'est pas une décision.",
        "Enjeu": "tout le tarif du Fournisseur I",
        "Décision attendue": "Désigner le fichier qui fait foi",
    },
    {
        "Sujet": "FOURNISSEUR D — 60 CHUT à x2,80",
        "Question": "Soixante chaussures « UNITE » ressortent à 2,80 fois "  # valeur d'exemple
                    "leur DPA.",
        "Ce qui est établi": "2,80 vaut à peu près 2 x 1,40 : la paire ET la "  # valeurs d'exemple
                             "remise de 25 % connue chez ce fournisseur. Mais "
                             "il a été dit qu'il cote 1 au catalogue.",
        "Enjeu": "60 articles",
        "Décision attendue": "Le tarif cote-t-il la paire ou l'unité, et "
                             "la remise est-elle déjà dedans ?",
    },
    {
        "Sujet": "FOURNISSEUR AZ — le peson de lève-personne",
        "Question": "La référence WMS est A151002, le tarif porte A151000 "
                    "(« Peson pour lève personne », 880,00 EUR).",  # valeur d'exemple
        "Ce qui est établi": "880,00 = 1100,00 x 0,80, comme toutes les lignes "  # valeurs d'exemple
                             "du tarif. Le DPA est 800,00, soit x1,10. Le "
                             "fichier « Lève personne pour le peson » "
                             "suggère un même peson décliné selon l'attache.",
        "Enjeu": "1 article",
        "Décision attendue": "A151002 et A151000 sont-ils le même peson ?",
    },
    {
        "Sujet": "FOURNISSEUR B — les coquilles hors XXL",
        "Question": "Le WMS a des COQUILLE MODELE M1 en T38, T44, T48 ; le "
                    "tarif ne porte qu'une ligne « MODELE M1 XXL ».",
        "Ce qui est établi": "Le rapport x1,61 le dit : ce n'est pas le même "  # valeur d'exemple
                             "produit. Les tailles courantes manquent au "
                             "tarif. Idem pour le MODELE M2.",
        "Enjeu": "une dizaine d'articles",
        "Décision attendue": "Question à poser au Fournisseur B, pas au code",
    },
    {
        "Sujet": "FOURNISSEUR H — 147 articles mal rattachés",
        "Question": "Des articles des Fournisseurs AM, BG et BY sont rattachés "
                    "au Fournisseur H, et leur champ référence contient des "
                    "désignations en texte libre.",
        "Ce qui est établi": "Ce n'est pas un problème de tarif mais de "
                             "rattachement fournisseur dans le WMS.",
        "Enjeu": "147 articles",
        "Décision attendue": "Corriger les fiches, ou accepter qu'ils restent "
                             "sans prix",
    },
]


SOURCES = [
    ("Conditionnement", "pa_conditionnement_a_trancher.xlsx", "À trancher", 3,
     "Prix dont l'unité est douteuse. Un diviseur négatif veut dire "
     "MULTIPLIER : le WMS compte par lot, le tarif cote à l'unité."),
    ("Fournisseur B par modèle", "pa_par_modele.xlsx", "Par modèle", 3,
     "Rapprochés par NOM DE MODÈLE, faute de référence au tarif. Plusieurs "
     "lignes par article : c'est voulu, l'arbitrage est humain."),
    ("Libellé", "pa_libelle_relecture.xlsx", "À relire", 3,
     "Candidats trouvés par ressemblance de libellé. Mesuré à 96 % de "
     "justesse : environ un sur vingt-cinq est faux."),
    ("Divergences avec le responsable", "reconciliation_responsable.xlsx",
     "Divergences", 0,
     "Les deux pipelines ont un prix et ils diffèrent. Toutes au palier 1, "
     "donc réellement opposables."),
]


def lire(nom_fichier: str, onglet: str, entete: int) -> pd.DataFrame:
    from extracteurs.base import chemin_lisible

    chemin = SORTIE / nom_fichier
    if not chemin.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(chemin_lisible(chemin), sheet_name=onglet,
                             header=entete)
    except Exception as err:
        print(f"    ! {nom_fichier} / {onglet} : {type(err).__name__} {err}")
        return pd.DataFrame()


def main() -> None:
    import mise_en_forme

    tables = {}
    resume = []
    for titre, fichier, onglet, entete, explication in SOURCES:
        table = lire(fichier, onglet, entete)
        tables[titre] = table
        resume.append({
            "Sujet": titre,
            "Lignes": len(table),
            "Ce que c'est": explication,
            "Fichier d'origine": fichier,
        })
        print(f"  {titre:<24} {len(table):>5} lignes")

    synthese = pd.DataFrame(resume)
    questions = pd.DataFrame(QUESTIONS)

    chemin = mise_en_forme.chemin_ecriture(SORTIE / "pa_A_ARBITRER.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as writeur:
        questions.to_excel(writeur, sheet_name="1 - Questions de fond",
                           index=False, startrow=3)
        synthese.to_excel(writeur, sheet_name="2 - Ce qu'il y a dedans",
                          index=False, startrow=3)
        for titre, table in tables.items():
            if table.empty:
                continue
            # Excel limite les noms d'onglet à 31 caractères.
            table.to_excel(writeur, sheet_name=titre[:31], index=False,
                           startrow=3)

    mises = {
        "1 - Questions de fond": {
            "ligne_entete": 4, "figer_colonne": 2,
            "titre": "Les décisions à prendre, par ordre d'importance",
            "sous_titre": "Une seule réponse débloque souvent des dizaines "
                          "d'articles. Rien n'est appliqué sans elle."},
        "2 - Ce qu'il y a dedans": {
            "ligne_entete": 4, "figer_colonne": 2,
            "titre": "Les listes détaillées, feuille par feuille",
            "sous_titre": "Chacune vient d'un outil du chantier"},
    }
    for titre, _, _, _, explication in SOURCES:
        if titre in tables and not tables[titre].empty:
            mises[titre[:31]] = {
                "ligne_entete": 4, "figer_colonne": 2,
                "titre": titre, "sous_titre": explication}
    mise_en_forme.formater(chemin, mises)

    print()
    print(f"  {len(questions)} questions de fond, "
          f"{sum(len(t) for t in tables.values())} lignes de détail")
    print(f"  -> {chemin}")


if __name__ == "__main__":
    main()
