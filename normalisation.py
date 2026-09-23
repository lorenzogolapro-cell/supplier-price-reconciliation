# -*- coding: utf-8 -*-
"""
Normalisation de la référence article, pour le rapprochement.

Une seule fonction, volontairement isolée : elle est appelée des deux côtés
de chaque jointure du pipeline. Deux normalisations qui divergent d'un
espace ou d'une casse produisent un rapprochement vide, sans lever
d'erreur — c'est exactement le genre de panne silencieuse que ce module
existe pour empêcher.
"""

import re


def normaliser_reference(valeur) -> str | None:
    """Référence interne sous sa forme de rapprochement.

    Le slash de tête est un artefact d'export : « /51280 » et « 51280 »
    désignent le même article. Tout autre slash est conservé — il peut
    porter du sens, et rien ne dit qu'il ne s'agit pas d'une référence
    fabricant glissée dans la colonne.
    """
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return None
    texte = str(valeur).strip()
    if not texte or texte.lower() in ("nan", "none", "-"):
        return None
    texte = texte.lstrip("/").strip()
    texte = re.sub(r"\s+", " ", texte).upper()
    return texte or None
