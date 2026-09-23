# -*- coding: utf-8 -*-
"""
Normalisation of the article reference, for matching.

A single function, deliberately isolated: it is called on both sides of
every join in the pipeline. Two normalisations that diverge by one space
or one letter case produce an empty match without raising any error,
which is exactly the kind of silent failure this module exists to
prevent.
"""

import re


def normaliser_reference(valeur) -> str | None:
    """Internal reference in its matching form.

    The leading slash is an export artefact: "/51280" and "51280"
    designate the same article. Any other slash is kept: it may carry
    meaning, and nothing says it is not a manufacturer reference that
    slipped into the column.
    """
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return None
    texte = str(valeur).strip()
    if not texte or texte.lower() in ("nan", "none", "-"):
        return None
    texte = texte.lstrip("/").strip()
    texte = re.sub(r"\s+", " ", texte).upper()
    return texte or None
