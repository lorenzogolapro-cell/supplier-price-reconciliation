# -*- coding: utf-8 -*-
"""
Extractor for the FOURNISSEUR C price list (PDF, "tarif_fournisseur.pdf",
25 pages).

CPAP/ventilation devices, masks, consumables and accessories. The PDF has
a clean text layer, so no OCR is needed.

ONE SINGLE PRICE COLUMN, AND IT IS THE PURCHASE PRICE
    The document is the named price list sent to the distributor. It
    publishes no list price, no recommended price and no LPPR: one single
    column of amounts, headed "Tarif 2026 HT en Euros" (mask and device
    pages) or "HT en €" (accessory pages, in landscape).

    Check, made before writing the first regex: out of the 1,567 lines in
    the text layer, ONE SINGLE line carries two amounts, and it is a
    service subscription ("1,00€/semaine soit 4,33€/mois", sample
    values). There is therefore no second price column to confuse with
    the first: the usual trap (net list price / gross list price /
    discount) does not arise here.

    Cross-check against known market orders of magnitude: a MODELE
    EXEMPLE device at 600.00 EUR (observed retail around 900 EUR), a
    MODELE EXEMPLE mask at 100.00 EUR (retail ~ 160 EUR). These really
    are distributor purchase prices, not retail prices. (amounts: sample
    values)

WHY GEOMETRY RATHER THAN A REGEX OVER THE TEXT
    The raw text from pypdfium2 runs the label and the price together,
    and the price list contains labels that END in a number:

        K10000  Kit d'alimentation MODELE EXEMPLE 10  500,00 EUR

    Read as text, "EXEMPLE 10 500,00" yields 10,500.00 EUR for a power
    kit worth 500 EUR: a factor of 21 that would slip through
    pa_completer's x10 safety net one time in two. The x coordinates
    settle it without ambiguity: the "10" is at x=204, inside the label
    column, and the price starts at x=443. So we read the PDF with
    pdfplumber, by columns.

    Second reason: on the accessory pages, pdfplumber cuts amounts
    anywhere it likes. "24,00" comes out as two words, "2" then "4,00",
    and "413,00" as "4" then "13,00". Only concatenating EVERYTHING found
    in the right-hand column rebuilds the number.

THREE LAYOUTS IN THE SAME FILE
    masks (portrait)       : Reference | (A fuite) | Label | Price | COMMENTAIRES
    devices (portrait)     : Reference | Label | Price
    accessories (landscape): Reference | Label | 13 machine compatibility
                             columns | Price

    The boundaries are therefore not hard-coded: the price column is
    located page by page from the x coordinate of the "€" symbols, which
    are all aligned. What sits to the RIGHT of the "€" is not a price but
    the COMMENTAIRES column ("ARRET DE COMMERCIALISATION 31 MAI 2025"):
    it goes out as an alert, never as a label.

    The symbol is not always there, though: the DIVERS page carries none,
    and three product-range blocks omit it row by row. The column is then
    aligned on the right edge of the amounts, without which a whole page
    used to disappear, including the REF0001 chin strap, which is in
    scope.

    The machine compatibility crosses (isolated "x" characters, between
    the label and the price) are stripped out: they are not words of the
    label.

ROWS ARE READ ACROSS SEVERAL BANDS
    One row in five has its label on a different band from its reference
    and its price. Three shapes, all of them present:

        band 220: REF0002  x x x x  2 2,00 EUR    price before label
        band 221:        Filtre MODELE EXEMPLE, Std, pack de 12

        band 385: Systeme de polysomnographie ambulatoire MODELE EXEMPLE
        band 391: REF0003-KA                        15 000,00 EUR
        band 397: chargeur, piles et kit de demarrage Adultes

    Hence the grouping by vertical proximity (TOLERANCE_RANG) rather than
    a split into text lines. The completeness check is simple and it
    comes out exactly right: 551 price cells in the PDF, 551 rows read.

WHAT WE DO NOT DIVIDE
    Many references are lots: "Filtre MODELE EXEMPLE (par 12)" at
    15.00 EUR. These are not pack sizes to divide by but DISTINCT
    REFERENCES: the same filter exists as a single unit (REF0004,
    1.50 EUR), by 2 (REF0005), by 12 (REF0006), by 50 (REF0007), each
    with its own code, which the WMS stores as is. Dividing would give a
    price no order could ever be placed at. So the lot is flagged as an
    alert and the price stays the price of the reference. (references and
    amounts: sample values)

REPETITIONS ARE NOT DUPLICATES
    The same accessory is reprinted under every compatible product range:
    the REF0008 magnetic clip comes back in eight sections, the REF0009
    template in six. 61 of the 551 rows read are of this kind, always at
    the same price: no reference in the price list carries two different
    prices, verified. They are collapsed to one row per (reference,
    price).

WHAT IS DELIBERATELY LEFT ASIDE
    - The last 5 product pages are headed "References en fin de
      commercialisation": no price column at all, only "Jusqu'a
      epuisement des stocks". Nothing to extract.
    - Amounts at 0.00 EUR (free measuring templates, software supplied at
      "- €"): a zero purchase price pushed into the WMS is more dangerous
      than a missing one. They are counted, not returned.
    - Service subscriptions are quoted "1€/semaine soit 4,33€/mois":
      that is not a unit price for an item.
"""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

import pandas as pd

from extracteurs.base import chemin_lisible, finaliser, nettoyer_texte

FOURNISSEUR = "FOURNISSEUR C"

# Maximum vertical gap, in PDF points, between two fragments of the same
# price list row. Grouping proceeds step by step: when the label spans
# two levels, the reference and the price sit BETWEEN the two, so the row
# reads in this order
#
#     band 385   Systeme de polysomnographie ambulatoire MODELE EXEMPLE
#     band 391   REF0003-KA                                 15 000,00 EUR
#     band 397   chargeur, piles et kit de demarrage Adultes
#
# that is, jumps of 6 pt at most, while two neighbouring price list rows
# are always at least 7 pt apart (11 pt on the device pages). 6 pt
# therefore separates the two cases without ever confusing them.
TOLERANCE_RANG = 6.0

# X coordinate where the "References" column ends. It starts at x=36 on
# portrait pages as on landscape ones; the label never starts before
# x=80. This boundary also separates the section titles ("MASQUES
# FACIAUX", flush left) from the item rows.
X_FIN_REFERENCES = 70.0

# Vertical distance beyond which a note from the COMMENTAIRES column is
# no longer attached to a row: it then applies to a whole section, not to
# a single item.
ECART_COMMENTAIRE = 8.0

# X coordinate where the machine compatibility columns of the accessory
# pages begin (one column per compatible range), ticked with an isolated
# lower-case "x". The first is at x=243, the longest label stops at
# x=231: 230 separates the two without confusing them.
#
# The coordinate condition is not a luxury. A filter based on the
# character alone erased the "X" from "Coussins narinaires - X Small"
# (REF0010), which then became the exact twin of the Small (REF0011):
# two distinct items made indistinguishable by their label.
X_DEBUT_COMPATIBILITE = 230.0

# Maximum width of a price cell, in points. Used to walk back from the
# "€" to the first digit of the amount: "16 000,00" (sample value) takes
# up 38 pt, so 45 pt is taken as the margin. Beyond that the previous
# column would start.
LARGEUR_CELLULE_PRIX = 45.0

# A FOURNISSEUR C reference: 12345, 12345-KB, 10000001, R100-200, K10000,
# 27000K, 7000001-PK12, 7000002-3 (sample values). At least three digits:
# that is what tells a reference from a section title starting with a
# digit ("2 NIVEAUX DE PRESSION") or from a subscription
# ("ABOEXEMPLE01").
REFERENCE = re.compile(r"^[A-Z]{0,4}\d[A-Z0-9]*(?:[-/][A-Z0-9]+)*$")

# Amount rebuilt after gluing back the fragments of the right column.
MONTANT = re.compile(r"^\d{1,6},\d{2}$")

# Bands to ignore: headers repeated at the top of each page, and footers.
BRUIT = re.compile(
    r"Internal Use|R[ée]f[ée]rences\s+D[ée]signation|COMMENTAIRES"
    r"|TARIFS? (MASQUES|APPAREILS|ACCESSOIRES)|en Euros|Tarif 2026 HT"
    r"|fin de commercialisation",
    re.IGNORECASE,
)

# Lots sold under their own reference: "(par 12)", "pack de 50", "(x12)",
# "(25)". Only used to raise an alert, see the module docstring.
LOT = re.compile(
    r"\(\s*par\s+(\d+)\s*\)|pack\s+de\s+(\d+)|\(\s*x\s*(\d+)\s*\)|\(\s*(\d{2,})\s*\)",
    re.IGNORECASE,
)


def _mot_texte(mot) -> str:
    return mot["text"].strip()


def _rangs(mots: list[dict]) -> list[list[dict]]:
    """Groups the words of a page into visual bands.

    pdfplumber returns words out of order: we sort by y coordinate, then
    chain them as long as the gap stays below TOLERANCE_RANG.
    """
    rangs: list[list[dict]] = []
    courant: list[dict] = []
    precedente = None
    for mot in sorted(mots, key=lambda m: (m["top"], m["x0"])):
        # Chaining uses the y coordinate of the PREVIOUS word, not that of
        # the first word of the band: a two-level row advances in 4 pt
        # steps and spreads over 8 pt in total.
        if precedente is not None and mot["top"] - precedente > TOLERANCE_RANG:
            rangs.append(courant)
            courant = []
        courant.append(mot)
        precedente = mot["top"]
    if courant:
        rangs.append(courant)
    return rangs


def _colonne_prix(mots: list[dict]) -> tuple[float, float] | None:
    """Bounds (left, right) of the page's price column.

    The "€" symbols are all aligned on the same x coordinate; the amount
    is set to their left, any comment to their right. We take their
    MEDIAN and not their minimum: the header of the accessory pages also
    contains a "€", 12 points further left than the column, and it would
    shift the boundary right into the compatibility crosses.

    Not every page carries the symbol: the DIVERS page has none, and
    three product-range blocks omit it row by row. Failing that, the
    column is aligned on the RIGHT edge of the amounts, which is exactly
    where the symbol would stand. Without that fallback a whole page used
    to disappear, including the REF0001 chin strap, which is in scope.

    Returns None on a page without a single amount ("References en fin de
    commercialisation", terms and conditions of sale).
    """
    abscisses = sorted(m["x0"] for m in mots if _mot_texte(m) == "€")
    if not abscisses:
        bords = sorted(m["x1"] for m in mots if MONTANT.match(_mot_texte(m)))
        if not bords:
            return None
        # The missing symbol would sit two points after the amount.
        abscisses = [bords[len(bords) // 2] + 2.0]
    mediane = abscisses[len(abscisses) // 2]
    return mediane - LARGEUR_CELLULE_PRIX, mediane + 10.0


def _intitule(page_texte: str) -> str:
    """EXACT heading of the price column, as it is printed.

    Two layouts, two wordings: the accessory pages (landscape) shorten it
    to "HT en €" where the mask and device pages write "Tarif 2026 HT en
    Euros". Both are returned verbatim, they go into colonne_prix.
    """
    if "TARIF ACCESSOIRES APPAREILS 2026" in page_texte:
        return "HT en €"
    return "Tarif 2026 HT en Euros"


def _decouper(rang: list[dict], bornes: tuple[float, float]) -> dict:
    """Sorts the words of a band into the three zones of the grid."""
    gauche, _ = bornes
    zone = {"reference": [], "designation": [], "prix": []}
    # Sort by y THEN by x: a label spread over two levels reads top to
    # bottom, not left to right; its two fragments start at the same x
    # coordinate.
    for mot in sorted(rang, key=lambda m: (m["top"], m["x0"])):
        texte = _mot_texte(mot)
        if not texte:
            continue
        if mot["x0"] >= gauche:
            if texte != "€":
                zone["prix"].append(texte)
        elif mot["x0"] < X_FIN_REFERENCES:
            zone["reference"].append(texte)
        elif not (texte == "x" and mot["x0"] >= X_DEBUT_COMPATIBILITE):
            # An isolated lower-case "x", past the label, is a machine
            # compatibility cross and not a word of the label.
            zone["designation"].append(texte)
    return zone


def _commentaires(mots: list[dict], droite: float) -> list[tuple[float, str]]:
    """COMMENTAIRES column of the mask pages, by y coordinate.

    It sits to the RIGHT of the price and is not aligned on the rows: a
    note such as "ARRET DE COMMERCIALISATION 31 MAI 2025" lands halfway
    between two items, 4 points from each. Leaving it in the band
    grouping would weld the two neighbouring rows into one and make every
    other item disappear. So it is set aside here, then attached to the
    nearest row.
    """
    groupes: dict[float, list[dict]] = {}
    for mot in mots:
        if mot["x0"] < droite or not _mot_texte(mot):
            continue
        clef = next((c for c in groupes if abs(c - mot["top"]) <= 2), mot["top"])
        groupes.setdefault(clef, []).append(mot)
    return [
        (clef, " ".join(_mot_texte(m) for m in sorted(v, key=lambda m: m["x0"])))
        for clef, v in sorted(groupes.items())
    ]


def _montant(morceaux: list[str]) -> float | None:
    """Glues the fragments of the right column back into an amount.

    On the accessory pages, "24,00" arrives as two words ("2", "4,00")
    and "413,00" as "4" + "13,00": we concatenate without a space,
    otherwise the amount loses its leading digit.
    """
    brut = "".join(morceaux).replace(" ", "").replace(" ", "").replace(" ", "")
    if not MONTANT.match(brut):
        return None
    return float(brut.replace(",", "."))


def _lot(designation: str) -> int | None:
    trouve = LOT.search(designation or "")
    if not trouve:
        return None
    for valeur in trouve.groups():
        if valeur:
            quantite = int(valeur)
            # "(25x10mm)" and other dimensions are not lots.
            return quantite if 1 < quantite <= 500 else None
    return None


def extract(path) -> pd.DataFrame:
    chemin = Path(path)
    try:
        import pdfplumber
    except ImportError:  # pragma: no cover
        return finaliser([])

    brutes: list[dict] = []
    sans_prix = 0
    a_zero = 0
    # Last section title seen. A handful of rows have an empty label
    # cell: the price list then makes do with the title just above
    # (REF0012 under "Tuyau avec bague de fuite MODELE EXEMPLE").
    rubrique: str | None = None

    with pdfplumber.open(str(chemin_lisible(chemin))) as pdf:
        for page in pdf.pages:
            mots = page.extract_words()
            bornes = _colonne_prix(mots)
            if bornes is None:
                continue
            intitule = _intitule(page.extract_text() or "")
            grille = [m for m in mots if m["x0"] < bornes[1]]

            lignes_page: list[dict] = []
            courante: dict | None = None
            for rang in _rangs(grille):
                texte_rang = " ".join(_mot_texte(m) for m in rang)
                if BRUIT.search(texte_rang):
                    # A header closes the current row, it does not cancel
                    # it: the last row of a page sits just above the
                    # footer and would be lost without this.
                    if courante:
                        lignes_page.append(courante)
                    courante = None
                    continue

                zone = _decouper(rang, bornes)
                reference = next(
                    (r for r in zone["reference"] if REFERENCE.match(r) and
                     sum(c.isdigit() for c in r) >= 3),
                    None,
                )

                # A section title is flush left, in the reference column,
                # but it SPILLS OVER into the label column ("Kit Bulle
                # MODELE EXEMPLE F40" runs from x=37 to x=96). Detecting
                # it from its starting x alone avoids sticking "F40" onto
                # the end of the previous row's label.
                debut = min(m["x0"] for m in rang)
                if reference is None and debut < X_FIN_REFERENCES:
                    rubrique = nettoyer_texte(texte_rang)
                    if courante:
                        lignes_page.append(courante)
                    courante = None
                    continue

                if reference is not None:
                    if courante:
                        lignes_page.append(courante)
                    courante = {
                        "reference": reference,
                        "designation": list(zone["designation"]),
                        "prix": list(zone["prix"]),
                        "commentaire": [],
                        "colonne_prix": intitule,
                        "rubrique": rubrique,
                        # y coordinate of the row, so its comment can be
                        # hooked back on once the whole page is read.
                        "y": min(m["top"] for m in rang),
                    }
                    continue

                if courante is None:
                    continue
                # Band with no reference: continuation of the previous row.
                courante["designation"] += zone["designation"]
                courante["prix"] += zone["prix"]

            if courante:
                lignes_page.append(courante)

            for hauteur, texte in _commentaires(mots, bornes[1]):
                if BRUIT.search(texte) or not lignes_page:
                    continue
                proche = min(lignes_page, key=lambda l: abs(l["y"] - hauteur))
                if abs(proche["y"] - hauteur) <= ECART_COMMENTAIRE:
                    proche["commentaire"].append(texte)

            brutes += lignes_page

    lignes = []
    for brute in brutes:
        prix = _montant(brute["prix"])
        designation = nettoyer_texte(" ".join(brute["designation"]))
        if prix is None:
            sans_prix += 1
            continue
        if prix == 0:
            # Free templates, software at "- €": a zero purchase price
            # pushed into the WMS would be taken for a real one.
            a_zero += 1
            continue

        alertes = []
        if not designation:
            designation = brute.get("rubrique")
            alertes.append("designation absente de la ligne, reprise du titre "
                           "de rubrique")
        commentaire = nettoyer_texte(" ".join(brute["commentaire"]))
        if commentaire:
            alertes.append(commentaire)
        quantite = _lot(designation)
        if quantite:
            alertes.append(
                f"reference vendue par {quantite} — prix NON divise, "
                f"c'est le prix de la reference telle qu'elle se commande"
            )

        lignes.append({
            "fournisseur": FOURNISSEUR,
            "ref_fournisseur": brute["reference"],
            "designation": designation,
            "prix_achat_unitaire_ht": prix,
            "colonne_prix": brute["colonne_prix"],
            "fichier_source": chemin.name,
            "alerte": " | ".join(alertes) or None,
        })

    # The price list repeats the same reference under every compatible
    # range: the REF0008 clip comes back in eight sections. Only one row
    # per (reference, price) is kept, and if a reference carries two
    # different prices, both go out flagged rather than being silently
    # arbitrated.
    par_reference = defaultdict(set)
    for ligne in lignes:
        par_reference[ligne["ref_fournisseur"]].add(ligne["prix_achat_unitaire_ht"])

    vues = set()
    retenues = []
    for ligne in lignes:
        clef = (ligne["ref_fournisseur"], ligne["prix_achat_unitaire_ht"])
        if clef in vues:
            continue
        vues.add(clef)
        prix_multiples = par_reference[ligne["ref_fournisseur"]]
        if len(prix_multiples) > 1:
            liste = " / ".join(f"{p:.2f}" for p in sorted(prix_multiples))
            complement = f"reference a plusieurs prix dans le tarif : {liste}"
            ligne["alerte"] = (
                f"{ligne['alerte']} | {complement}" if ligne["alerte"] else complement
            )
        retenues.append(ligne)

    print(f"    {len(retenues)} rows kept out of {len(lignes)} read "
          f"({len(lignes) - len(retenues)} repeats from one range to the "
          f"next), {sans_prix} with no price, {a_zero} at 0.00 EUR")
    return finaliser(retenues)


if __name__ == "__main__":
    import sys

    defaut = (Path(__file__).resolve().parent.parent / "catalogues"
              / "fournisseur_c" / "2026" / "tarif_fournisseur.pdf")
    cible = Path(sys.argv[1]) if len(sys.argv) > 1 else defaut
    table = extract(cible)
    pd.set_option("display.width", 220)
    pd.set_option("display.max_colwidth", 70)
    pd.set_option("display.max_rows", 400)
    print(f"{len(table)} rows extracted from {cible.name}")
    if not table.empty:
        prix = table["prix_achat_unitaire_ht"]
        print(f"price: min {prix.min():.2f} / median {prix.median():.2f} / "
              f"max {prix.max():.2f} EUR")
        colonnes = ["ref_fournisseur", "designation", "prix_achat_unitaire_ht",
                    "colonne_prix"]
        print(table[colonnes].head(60).to_string(index=False))
        print("...")
        print(table[colonnes].tail(40).to_string(index=False))
