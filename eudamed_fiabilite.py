# -*- coding: utf-8 -*-
"""
Quand un code venu d'EUDAMED mérite qu'on le garde.

EUDAMED n'est pas en cause : la façon dont on l'a interrogé l'est. On lui
a soumis la référence de l'article, et le registre rend le dispositif qui
porte cette référence — chez N'IMPORTE QUEL fabricant d'Europe. Une
référence à quatre chiffres en désigne des centaines.

Ce que cela donne dans le fichier, sans rien inventer :

    CHAUSSURE MODELE-B 39  → réf « 2975 » → 5012345678900
    CHAUSSURE MODELE-B 40  → réf « 2975 » → 5012345678900
    CHAUSSURE MODELE-B 41  → réf « 2975 » → 5012345678900

Trois pointures sous un seul code, préfixe britannique sur une chaussure
française.

La mesure
---------
Chaque code confronté au préfixe GS1 relevé dans le tarif du MÊME
fournisseur (V16, 557 codes testables) :

    référence de 3 caractères    28 codes     0 % de concordance
                 4                23          4 %
                 5                27          7 %
                 6               148         11 %
                 7               111         30 %
                 8               114          8 %
                 9 et plus        96         96 %

La rupture est franche et ne doit rien au hasard : au-delà de neuf
caractères une référence est distinctive, elle ne ramène qu'un dispositif
et c'est le bon. En deçà, c'est une loterie.

Le test a un biais qu'il faut connaître : EUDAMED indexe le FABRICANT, et
un fournisseur qui distribue sans fabriquer portera légitimement un
préfixe étranger. Ce biais joue contre EUDAMED plus que contre les autres
sources — mais il n'explique ni le 0 % à trois caractères, ni le saut à
96 % dès qu'on passe neuf.

Ce que la règle NE fait pas
----------------------------
Elle ne supprime rien. Les codes écartés partent dans
`perimetre/retours/`, d'où l'on pourra les reprendre le jour où on saura
les vérifier. Elle ne touche pas non plus « EAN distributeur - Vn.xlsx », qui
reste le journal de ce que la chaîne a produit : c'est à la SORTIE, dans
ce qu'on lit et dans ce qu'on emporte en entrepôt, que le tri se fait.
"""

from __future__ import annotations

import pandas as pd

# Neuf, parce que c'est là que la mesure bascule — pas parce que c'est un
# chiffre rond. Le relâcher à 8 ferait rentrer 114 codes à 8 % de
# concordance.
LONGUEUR_MINIMALE = 9

SOURCE = "eudamed"

MOTIF = (f"référence de moins de {LONGUEUR_MINIMALE} caractères : EUDAMED "
         f"a rendu le dispositif d'un autre fabricant")


def reference_fiable(reference) -> bool:
    """Cette référence était-elle assez distinctive pour interroger EUDAMED ?"""
    if reference is None or (isinstance(reference, float) and pd.isna(reference)):
        return False
    return len(str(reference).strip()) >= LONGUEUR_MINIMALE


def douteux(df: pd.DataFrame, colonne_source: str = "Source",
            colonne_reference: str = "Réf. fournisseur") -> pd.Series:
    """Les lignes dont le code vient d'EUDAMED sur une référence trop courte.

    Rend un masque, jamais un cadre tronqué : à l'appelant de décider s'il
    écarte, s'il déclasse ou s'il se contente de compter.
    """
    if colonne_source not in df.columns:
        return pd.Series(False, index=df.index)
    vient_deudamed = df[colonne_source].fillna("").str.strip().eq(SOURCE)
    trop_court = ~df[colonne_reference].map(reference_fiable)
    return vient_deudamed & trop_court
