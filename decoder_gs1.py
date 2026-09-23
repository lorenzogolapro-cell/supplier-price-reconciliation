# -*- coding: utf-8 -*-
"""
Decodes a GS1-128 / GS1 DataMatrix barcode read with a hand scanner.

Medical device labels do not carry a plain EAN-13 but a structured
string, where every piece of data is preceded by an application
identifier (AI) in parentheses:

    (01)01234567890128(17)270329(10)L4521

    (01) GTIN      -> the product code, 14 digits
    (17) 270329    -> expiry, 29 March 2027
    (10) L4521     -> batch number

What we are after is the GTIN. When its first digit, the packaging
indicator, is 0, it designates the sales unit and the GTIN is an EAN-13
preceded by a padding zero: that thirteen-digit code is the one to look
up in the WMS. An indicator from 1 to 8 designates a case or a pallet,
which is not the article we sell.

    python decoder_gs1.py "(01)01234567890128(17)270329"

Also useful to check a scan that returned nothing during picking: if the
scanner transmits the whole string instead of the GTIN alone, the WMS
finds nothing even though the code is correct.
"""

from __future__ import annotations

import re
import sys

from extracteurs.base import cle_controle_ean, ean_est_valide

# Fixed length of the most common AIs. Those not listed here are of
# variable length and end at the separator or at the next opening
# parenthesis.
LONGUEURS_FIXES = {
    "00": 18,   # SSCC, logistic unit
    "01": 14,   # GTIN
    "02": 14,   # GTIN of the contained items
    "11": 6,    # production date
    "13": 6,    # packaging date
    "15": 6,    # best-before date
    "17": 6,    # expiry date
    "20": 2,    # product variant
}

LIBELLES = {
    "00": "logistic unit (SSCC)",
    "01": "GTIN (product code)",
    "02": "GTIN of the contained items",
    "10": "batch number",
    "11": "production date",
    "15": "best before",
    "17": "expiry date",
    "21": "serial number",
    "30": "quantity",
    "37": "number of items contained",
    "240": "manufacturer reference",
    "241": "customer reference",
}

# GS1 dates are written YYMMDD
DATE = re.compile(r"^(\d{2})(\d{2})(\d{2})$")


def _lisible(ai: str, valeur: str) -> str:
    """Value formatted according to what the AI holds."""
    if ai in ("11", "13", "15", "17"):
        trouve = DATE.match(valeur)
        if trouve:
            annee, mois, jour = trouve.groups()
            # GS1: 00-49 -> 2000-2049, 50-99 -> 1950-1999
            siecle = "20" if int(annee) <= 49 else "19"
            jour = jour if jour != "00" else "last day of the month"
            return f"{jour}/{mois}/{siecle}{annee}"
    return valeur


def decoder(chaine: str) -> dict:
    """Splits a GS1 string into {AI: value}.

    Accepts both writings: with parentheses, as the labels print them,
    or raw with FNC1 separators, as some scanners emit them.
    """
    if not chaine:
        return {}
    texte = chaine.strip()

    # Readers prefix the string with a symbology identifier: ]C1 for a
    # GS1-128, ]d2 for a GS1 DataMatrix, ]e0 for an RSS. It announces the
    # barcode type and is not part of the data.
    texte = re.sub(r"^\][A-Za-z]\d", "", texte)

    # Parenthesised form: the simplest, the AIs are delimited
    if "(" in texte:
        elements = re.findall(r"\((\d{2,4})\)([^(]*)", texte)
        return {ai: valeur.strip() for ai, valeur in elements}

    # Raw form: we move forward using the fixed lengths, and stop at the
    # separator for the variable-length AIs.
    texte = texte.replace("\x1d", "|")
    resultat: dict[str, str] = {}
    position = 0
    while position < len(texte) - 1:
        ai = texte[position:position + 2]
        position += 2
        longueur = LONGUEURS_FIXES.get(ai)
        if longueur:
            resultat[ai] = texte[position:position + longueur]
            position += longueur
        else:
            fin = texte.find("|", position)
            fin = len(texte) if fin == -1 else fin
            resultat[ai] = texte[position:fin]
            position = fin + 1
    return resultat


def ean_depuis_gs1(chaine: str) -> str | None:
    """EAN-13 of the sales unit held in a GS1 code, if there is one.

    Returns None when the GTIN designates a case or a pallet: that is
    then not the code of the article we sell.
    """
    gtin = decoder(chaine).get("01")
    if not gtin:
        return None
    gtin = re.sub(r"\D", "", gtin)
    if len(gtin) == 13 and ean_est_valide(gtin):
        return gtin
    if len(gtin) != 14:
        return None
    if gtin[0] != "0":
        return None  # case packing, not the sales unit
    candidat = gtin[1:]
    return candidat if ean_est_valide(candidat) else None


def expliquer(chaine: str) -> str:
    """Readable description of a scanned code."""
    elements = decoder(chaine)
    if not elements:
        return "empty or unreadable string"

    lignes = []
    for ai, valeur in elements.items():
        libelle = LIBELLES.get(ai, f"AI {ai}")
        lignes.append(f"  ({ai}) {libelle:<32} {_lisible(ai, valeur)}")

    gtin = elements.get("01", "")
    if gtin:
        ean = ean_depuis_gs1(chaine)
        if ean:
            lignes.append(f"\n  -> EAN-13 to look up in the WMS: {ean}")
        elif len(gtin) == 14 and gtin[0] != "0":
            lignes.append(
                f"\n  -> indicator {gtin[0]}: case packing code "
                f"(case or pallet), not the sales unit"
            )
        else:
            attendue = cle_controle_ean(gtin)
            lignes.append(
                f"\n  -> invalid GTIN: check digit read {gtin[-1]}, "
                f"expected {attendue}"
            )
    return "\n".join(lignes)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    for argument in sys.argv[1:]:
        print(f"\n{argument}")
        print(expliquer(argument))
