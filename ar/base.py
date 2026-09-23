# -*- coding: utf-8 -*-
"""
Common base for the supplier order acknowledgement readers.

An order acknowledgement states the price the supplier commits to
invoicing, for a given quantity, on a given date. It is the only source
that is a FACT: a price list is only an intention, and the WMS purchase
price is only what somebody was willing to key in.

Two consequences for the schema below:

  - we ALWAYS keep the quantity next to the price. One and the same
    article is negotiated at 9.00 EUR per 5 and 7.50 EUR per 20 (example
    values): comparing an acknowledgement price with a single purchase
    price only makes sense at a comparable quantity. It is the most
    expensive lesson of this project.
  - we keep the displayed price AND the net price. The first one is
    there to re-read the acknowledgement, the second is what we pay.

Each reader (ar/<fournisseur>.py) exposes:
    extract(chemin) -> pandas.DataFrame
respecting COLONNES, in that order.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import pandas as pd
import pdfplumber

# ---------------------------------------------------------------------------
# Normalised output schema
# ---------------------------------------------------------------------------
COLONNES = [
    "fournisseur",        # supplier name (e.g. FOURNISSEUR L)
    "fichier_ar",         # source PDF, so we can get back to the evidence
    "num_ar",             # acknowledgement number on the supplier's side
    "commande_wms",       # our own order number ("Votre reference")
    "ref_fournisseur",    # supplier's article reference (ALWAYS text)
    "designation_ar",     # label as it appears on the acknowledgement
    "quantite",           # quantity ordered, inseparable from the price
    "prix_unitaire_ar",   # unit price displayed, before discount
    "remise_pct",         # discount rate as a PERCENTAGE (25.0 = 25 %)
    "prix_net_ar",        # price actually paid: unit price less discount
    "montant_ligne",      # line total, when the acknowledgement gives it
    "delai_livraison",    # delivery date announced, when it is there
    # ORDER DATE, as the acknowledgement writes it, not the date the mail
    # was received (that one is in index_ar.csv and can differ from it).
    #
    # Added on 21/09/2026, and it is not a convenience: AN ORDER
    # ACKNOWLEDGEMENT PROVES NOTHING OUTSIDE ITS OWN DATE. At supplier L,
    # comparing the year's price list with all the acknowledgements gave
    # a ratio of 1.12, enough to believe we were overvaluing the whole
    # catalogue by 12 %. Broken down by month, the ratio is 1.1200 until
    # February then 1.0000 from March on: it was not an error, it was the
    # annual increase, and our prices were right. Without the date, we
    # were correcting the wrong way round.  (ratios and switch month:
    # example values)
    #
    # Same lesson at supplier E, where a second discount appears partway
    # through the year, and at supplier G, where the increase falls in a
    # different month: the switch month belongs to each supplier.
    "date_commande",
]


def _sans_invisibles(texte: str) -> str:
    """Removes the invisible formatting characters.

    Supplier H scatters soft hyphens (U+00AD) in the middle of its
    references. They cannot be seen, they print as nothing, and they
    make every match fail without anyone understanding why.
    """
    normalise = unicodedata.normalize("NFKC", texte)
    return "".join(c for c in normalise
                   if unicodedata.category(c) != "Cf")


def nettoyer_texte(valeur) -> str:
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return ""
    return " ".join(_sans_invisibles(str(valeur)).split())


def nombre(texte) -> float | None:
    """"1 111,36" -> 1111.36; acknowledgements write the French way.

    The non-breaking space is the usual PDF trap: it looks like a space
    but is not one, and float() chokes on it.
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
    """Price actually paid. A discount is always taken as an absolute
    value: acknowledgements write it sometimes "25%", sometimes "-25%"."""
    if unitaire is None:
        return None
    if remise is None or remise == 0:
        return round(unitaire, 4)
    return round(unitaire * (1 - abs(remise) / 100), 4)


def texte_pdf(chemin) -> str:
    """Full text of the PDF. No acknowledgement met so far needs OCR."""
    with pdfplumber.open(Path(chemin)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def premier_groupe(motif: re.Pattern, texte: str) -> str:
    trouve = motif.search(texte)
    return trouve.group(1) if trouve else ""


def cadrer(lignes: list[dict], fournisseur: str, chemin) -> pd.DataFrame:
    """Fits the rows read to the common schema, missing columns included."""
    df = pd.DataFrame(lignes)
    if df.empty:
        return pd.DataFrame(columns=COLONNES)
    df["fournisseur"] = fournisseur
    df["fichier_ar"] = Path(chemin).name
    for colonne in COLONNES:
        if colonne not in df.columns:
            df[colonne] = None
    return df[COLONNES]
