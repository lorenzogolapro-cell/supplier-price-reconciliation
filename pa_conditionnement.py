# -*- coding: utf-8 -*-
"""
PA workstream: bring a pack price back to a unit price.

    python pa_conditionnement.py            diagnostic over the whole scope
    python pa_conditionnement.py FOURNISSEUR_G  diagnostic on one supplier

THE PROBLEM
    Some suppliers quote by the lot - per 50, per 100, the box, the carton,
    the pair - while the WMS reasons by the unit. The price read is
    correct, but it
    does not answer the same question, and nothing in the file says so.

    An x2 at Supplier G is not an increase: it is a shoe sold by the pair
    while the WMS counts shoes one by one.

THE DIVISOR, IN ORDER OF RELIABILITY
    1. COLUMN    the price list carries an explicit column: "Nbre unités
                 par boîte" at Supplier O, "PCB" at Supplier AL, "CDT" at
                 Supplier BH, "UNITE DE VENTE" at Supplier BM. The
                 extractor stores it in
                 `conditionnement`: it is authoritative.
    2. DESIGNATION  the price list label carries it: "20X100ML", "BTE/50",
                 "LOT DE 25". This is text parsing, hence strict patterns:
                 a size "T38" or a model "1234567" must never pass for a
                 pack size.
    3. WMS       the pack size on the article record.
    4. INFERRED  the ratio to the DPA lands on a round integer.

    **Rank 4 NEVER applies on its own.** An inference from the ratio must
    be corroborated, either by another source or by its systematic
    recurrence: the same integer over at least five articles of the same
    supplier. A divisor inferred on an isolated article is REPORTED, it is
    not applied.

    This is the Supplier BP wall: their price list mixes articles quoted by
    the unit and
    others quoted by the pack, with no indicator. The factor is not
    inferable there. Do not try to force it.

THREE CONFIGURATIONS, NOT TO BE CONFUSED
    A. The price list gives several prices by quantity, one of them unitary.
       We take the unit column and **divide nothing**. The other columns
       are genuine price breaks, already expressed in units.

    B. The price list's only price is a pack price.
       We divide by N. The "per 100" is not a price break, it is the sales
       unit: there is only one price, so there is no threshold to convert.
       The pack size becomes a **minimum order quantity**, stored
       separately - it serves replenishment, where the MOQ beats the max.

    C. The price list quotes by the pack AND offers breaks on the NUMBER of
       packs. Here only, if the price is divided by N, the break is
       **multiplied by N**: 10 cartons of 100 are 1,000 units, not 10.
       Writing the volume price against a break of 10 would wrongly price
       every order between 10 and 999.

    The sorting test: does the price list carry more than one price for the
    same reference? No, and it is per pack -> B. Yes, one of them unitary
    -> A. Yes, none unitary -> C.

WHAT COMES OUT
    `diviseur_conditionnement`, `source_diviseur`, `prix_avant_division`,
    `preuve_diviseur`, `minimum_de_commande`. Without these columns, a
    wrong price cannot be detected afterwards.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

# --------------------------------------------------------------------------
# Rank 4: what we agree to infer

# Plausible pack sizes. We do not test just any integer: a ratio of 7.0 is
# more probably an increase than a lot of seven.
CONDITIONNEMENTS_PLAUSIBLES = (2, 3, 4, 5, 6, 8, 10, 12, 20, 24, 25, 30, 36,
                               48, 50, 100, 144, 200, 250, 500, 1000)

# Tolerance around the integer. 5% lets a moderate price increase ride on
# top of the pack size (x2.05 at Supplier G is still an x2)
# without swallowing a
# clearly different ratio.
TOLERANCE = 0.06

# Below this, a recurrence is a coincidence, not a supplier rule.
ARTICLES_MINIMUM = 5

# Maximum tolerated gap between an article's ratio and the factor
# established for its price list. 1.5 lets a price increase ride on top of
# the pack size without swallowing a ratio of a different order of
# magnitude.
MARGE_FACTEUR = 1.5


# --------------------------------------------------------------------------
# Rank 2: reading a pack size from a label
#
# Each pattern is anchored and requires a keyword or an unambiguous shape.
# Without that, "MODELE EXEMPLE 1 T3" or the reference "1234567" would
# produce a pack size, and the price would be divided by a model number.

MOTIFS_DESIGNATION = [
    # "20X100ML", "12X1L", "6X500G": N sub-units of volume/mass.
    # N is the pack size, not the volume.
    (re.compile(r"\b(\d{1,4})\s*[xX*]\s*\d+(?:[.,]\d+)?\s*(?:ML|L|G|KG|CL)\b"),
     "N x volume"),
    # "4 bidons de 5 L", "12 flacons de 500 ml", "6 boîtes de 100".
    # The shape is "N <container> DE <quantity>": N is what counts, never
    # the quantity that follows.
    #
    # ORDER MATTERS: this pattern comes BEFORE the keyword one below.
    # On "6 boîtes de 100", the keyword pattern would read "BOITES DE 100"
    # and return 100 - the contents of one box - where the price list sells
    # six boxes. The number BEFORE the container beats the one after it.
    # "BOITE DE 100", with no number in front, correctly falls through to
    # the next pattern.
    #
    # Found on 16/09 at Supplier I: "PRODUIT EXEMPLE 1 - 4 bidons de 5 L
    # + 1 pompe" at EUR 92.00 against a DPA of 21.30 - 92.00/4 = 23.00,
    # i.e. 1.08 times the DPA (example values).
    # Two guard rails, taken from a sweep of the scope's 5,082 labels,
    # where this pattern produced two false positives:
    #   - NOT preceded by "N°": "DOIGTIER TAILLE M N°3 SACHET DE 100"
    #     returned 3, which is a SIZE;
    #   - NOT preceded by "x" or "X": "COMPRESSE 10 X 10 CARTON DE"
    #     returned 10, which is a DIMENSION.
    (re.compile(r"(?<![\d.,°])(?<![xX]\s)"
                r"\b(\d{1,4})\s*(?:BIDONS?|FLACONS?|BOITES?|BOÎTES?|CARTONS?"
                r"|SACHETS?|POCHES?|TUBES?|PACKS?|BRIQUES?|BOUTEILLES?"
                r"|CUPS?|ROULEAUX?)\s+DE\b"),
     "N contenants de"),
    # "BTE/50", "BTE DE 50", "BOITE DE 100", "CARTON DE 24", "SACHET/4".
    # The keyword is what makes the pattern safe: it is what tells a pack
    # size apart from any other number.
    (re.compile(r"\b(?:BTE|BOITE|BOÎTE|CARTON|CT|SACHET|SACH|PAQUET|PQT|LOT|"
                r"POCHE|ETUI|DISTRIB(?:UTEUR)?)\s*(?:DE\s*|D[E']\s*|/\s*|\.\s*)?"
                r"(\d{1,4})\b"),
     "boîte / carton / lot"),
    # "PAR 100", "PAR 50"
    (re.compile(r"\bPAR\s+(\d{1,4})\b"), "« par N »"),
    # "- 5 PCS", "42PCS", "10 UNITES", "16 GALETTES".
    # Found on 16/09 at Supplier S - "DRAP DE REHAUSSEMENT 140x110cm -
    # 5 PCS" at EUR 133.00 against a DPA of 30.20: 133.00/5 = 26.60, i.e.
    # 0.88 times the DPA (example values). The lot was written down, nobody
    # was reading it.
    #
    # "PIECES" SPELLED OUT IN FULL IS EXCLUDED, and that is not an
    # oversight. Across the scope's 5,082 labels, the long form NEVER
    # designated a quantity: it designates the TYPE of product -
    # "POCHE MODELE EXEMPLE 2 2 PIECES" is a two-piece system,
    # "COLLIER TRACHEO 2 PIECES" a two-part collar. Halving their price was
    # the only possible effect of that word.
    # The abbreviation "PCS", by contrast, is always a quantity.
    # The number must be preceded NEITHER by a digit, NOR by a letter, NOR
    # by a "°": "CHUT MODELE EXEMPLE 3 NOIR T38 UNITE" would otherwise
    # return 38, which is a SHOE SIZE, and divide a shoe's price by 38.
    (re.compile(r"(?<![\d,°A-Z])\b(\d{1,4})\s*(?:PCS\b|UNITES?\b"
                r"|UNITÉS?\b|UVC\b|GALETTES?\b)"),
     "N pièces"),
    # "PACK X36", "X24 PCS", "PROTÈGE SEAU X20".
    #
    # The most profitable and the most dangerous pattern. Two guard rails,
    # taken from a pass over the scope's 5,082 designations:
    #   - NOT followed by a unit of measure: "130 X 90 CM" is a sheet, not
    #     a lot of 90; "2 X 1 KG" is a mass;
    #   - NOT preceded by a number: "10 X 12 CM", "70 X 110" are
    #     dimensions, and the second value there is never a pack size.
    (re.compile(r"(?<![\d.,])\s[xX]\s?(\d{1,4})\b"
                r"(?!\s*(?:CM|MM|M\b|KG|G\b|ML|L\b|CL|PO|POUCE))"),
     "« xN »"),
    # The bare "/N" pattern was REMOVED. On real designations it only
    # brought back sizes - "T: 54/56" returned 56, "380/405" returned 405,
    # "T39/41" returned 41. Twenty-four false positives, not one true one:
    # dividing a price by a size is precisely the silent error this module
    # exists to prevent. The genuine forms "BTE/100" and "SACHET/50" are
    # already caught by the keyword pattern above.
]

# A pack size read outside these bounds is almost certainly something else:
# a size, a length, a year, a reference.
BORNES_DESIGNATION = (2, 5000)

# "LA PAIRE", "PAIRE DE ...". The word is recognised, but it never PROPOSES
# a division - it only serves to recognise a WMS label that is already at
# the lot, that is, to PREVENT a division. See
# diviseur_depuis_designation, which carries the measurement that settled
# the matter.
MOTIF_PAIRE = re.compile(r"\b(?:LA\s+PAIRE|PAR\s+PAIRE|PAIRES?)\b")


def _conditionnement_brut(libelle) -> tuple[int | None, str]:
    """The unit count a label announces, without interpreting anything."""
    if libelle is None or (isinstance(libelle, float) and pd.isna(libelle)):
        return None, ""
    texte = str(libelle).upper()
    for motif, nom in MOTIFS_DESIGNATION:
        trouve = motif.search(texte)
        if not trouve:
            continue
        try:
            valeur = int(trouve.group(1))
        except (IndexError, ValueError):
            continue
        if BORNES_DESIGNATION[0] <= valeur <= BORNES_DESIGNATION[1]:
            return valeur, nom
    if MOTIF_PAIRE.search(texte):
        return 2, "paire"
    return None, ""


def diviseur_depuis_designation(designation_tarif,
                                designation_wms=None) -> tuple[int | None, str]:
    """The pack size to DIVIDE by, read from the labels.

    The delicate point: a pack size read in a label does not say that a
    division is needed. It says what the row represents. What commands the
    division is the GAP between what the price list quotes and what the WMS
    counts.

        price list "BTE/50", WMS by the unit    -> divide by 50
        price list "LA PAIRE"                   -> NEVER divide here

    WHY "PAIRE" NO LONGER DIVIDES (16/09)
        A pair of ramps is ONE product: a single ramp is not sold on its
        own. The word most often describes the nature of the article, not
        its packaging - and sometimes just a component, as at Supplier AA
        where "1 paire de bottes" is part of a EUR 2,500 compression
        therapy device (example value).

        The counter-check through the WMS label was not enough: it required
        the WORD "paire" on both sides, whereas the WMS says the same thing
        in the PLURAL - "RAMPES TELESCOPIQUES 152CM".

        Measurement of 16/09 over the full pass: this rank divided
        5 articles, and all 5 were wrong. At Supplier H (REF-EXEMPLE-1) and
        Supplier CS (REF-EXEMPLE-2), the price read was equal to the WMS's
        DPA TO THE CENT - the proof that no division was needed was right
        there in the data.
        None of the 5 ratios to the DPA corroborated an x2.

        And nothing depended on it: Supplier G's 50 rows, the case the rule
        had been written for, all go through rank 4, where the division is
        corroborated by the ratio to the DPA. That is where "paire" has to
        prove itself, row by row, or not at all.

    The word is still recognised by `_conditionnement_brut`: on the WMS
    label it serves to say "this article is ALREADY at the lot", hence to
    prevent a division. 296 articles in the scope depend on it.

    When only the price list label is available, we draw no conclusion: it
    is up to the caller to supply the matching WMS label.
    """
    cond_tarif, nom = _conditionnement_brut(designation_tarif)
    if not cond_tarif:
        return None, ""
    # "Paire" never proposes a division. See above: 5 divisions, 5 of them
    # wrong, and the only supplier concerned goes through rank 4.
    if nom == "paire":
        return None, ""
    cond_wms, _ = _conditionnement_brut(designation_wms)
    # Both sides announce the same pack size: same basis, nothing to do.
    if cond_wms and cond_wms == cond_tarif:
        return None, ""
    # The WMS announces a DIFFERENT pack size: the ratio of the two is the
    # only defensible divisor (price list per 100, WMS per 10 -> divide by
    # 10).
    if cond_wms and cond_wms != cond_tarif:
        if cond_tarif % cond_wms == 0:
            return (cond_tarif // cond_wms,
                    f"désignation ({nom}, net du conditionnement du WMS)")
        return None, ""
    return cond_tarif, f"désignation ({nom})"


# --------------------------------------------------------------------------
# Rank 4: inference from the ratio, never on its own


def entier_rond_le_plus_proche(rapport) -> int | None:
    """The plausible pack size `rapport` is close to, if there is one.

    Used to ESTABLISH a price list's factor, so it is deliberately strict:
    we only want to count as evidence the ratios that land squarely.
    """
    if rapport is None or pd.isna(rapport) or rapport <= 0:
        return None
    for n in CONDITIONNEMENTS_PLAUSIBLES:
        if abs(rapport - n) <= TOLERANCE * n:
            return n
    return None


def suit_le_facteur(rapport, facteur) -> bool:
    """Does this article follow the pack size already established for its
    price list?

    A different question from the previous one, and that is the whole
    point. Once we know the price list quotes by the lot of N - established
    over dozens of articles - it is no longer about asking every row to
    prove the factor again. It is about deciding between two hypotheses: is
    this row at the lot, or already unitary?

    We decide by whichever of the two is closer in RATIO, not in absolute
    difference: the lower boundary is the geometric mean of 1 and N. For
    N = 2 that is 1.41 - a ratio of 1.86 is at the lot, a ratio of 1.05 is
    not. Without that, four Supplier G shoes whose price list had moved by
    9%
    stayed at the price of the pair.

    BUT the geometric mean alone is ruinous on large factors: for N = 100
    it accepts any ratio between 10 and 1,000. At Supplier AE, prices at
    377 times the DPA were divided by 100 that way, and came out at
    EUR 0.0050 against EUR 0.50 (example values) - the brief's "divided
    twice", only worse. So we add a SECOND condition, multiplicative and
    bounded: the ratio must stay within half again more or less than the
    factor. For N = 100 that gives [66.7; 150] rather than [10; 1000].

    The two conditions are cumulative. A ratio that fails them is not
    divided: it is reported, and that is all.
    """
    if rapport is None or pd.isna(rapport) or rapport <= 0 or not facteur:
        return False
    n = float(facteur)
    plus_proche_du_lot = rapport > n ** 0.5
    dans_la_fourchette = (n / MARGE_FACTEUR) <= rapport <= (n * MARGE_FACTEUR)
    return plus_proche_du_lot and dans_la_fourchette


def deduire_par_systematicite(rapports: pd.Series) -> tuple[int | None, str, int]:
    """The pack size the supplier applies systematically.

    We look not at the size of a gap but at its RECURRENCE. The same
    integer recurring over dozens of articles is not an increase: it is a
    sales unit. A genuine price change never lands on the same factor for
    that many references.

    Returns (divisor, proof, number of articles concerned).
    """
    utiles = rapports.dropna()
    utiles = utiles[utiles > 0]
    if len(utiles) < ARTICLES_MINIMUM:
        return None, "", 0
    candidats = utiles.map(entier_rond_le_plus_proche).dropna()
    if candidats.empty:
        return None, "", 0
    comptes = candidats.value_counts()
    meilleur = int(comptes.index[0])
    combien = int(comptes.iloc[0])
    if combien < ARTICLES_MINIMUM:
        return None, "", combien
    # The recurrence has to dominate: if only half the articles follow the
    # factor, the price list mixes sales units and nothing can be inferred
    # wholesale. That is Supplier BP's case, and it must not be forced.
    part = combien / len(utiles)
    if part < 0.60:
        return None, (f"facteur {meilleur} sur {combien} des {len(utiles)} "
                      f"articles seulement — tarif mixte, non déductible"), combien
    preuve = (f"rapport au DPA proche de {meilleur} sur {combien} des "
              f"{len(utiles)} articles comparables ({part:.0%})")
    return meilleur, preuve, combien


# --------------------------------------------------------------------------
# Choosing the divisor, across all ranks


def choisir_diviseur(conditionnement_tarif, conditionnement_supposé,
                     designation, conditionnement_wms,
                     deduit=None, preuve_deduite="") -> tuple[float, str, str]:
    """The divisor kept, its source and its proof, in order of reliability.

    Always returns a usable triple: (1, "", "") when nothing is
    established. Dividing by default would be worse than doing nothing.
    """
    # Rank 1 - REMOVED, and the reason has to be stated so nobody puts it
    # back.
    #
    # The brief asked for the price list's pack size column ("Nbre unités
    # par boîte" at Supplier O, "PCB" at Supplier AL) to be used as
    # the most reliable divisor. That is wrong IN THIS PROJECT, because
    # extracteurs/base.py's schema is explicit:
    #
    #     prix_achat_unitaire_ht   purchase price of ONE unit, excl. tax
    #     conditionnement          number of units per pack
    #     prix_colis_ht            prix_achat_unitaire_ht x conditionnement
    #
    # The price is ALREADY unitary: the extractors have done the division.
    # The `conditionnement` column serves to go back up to the pack price,
    # not to come back down from it. That is also why Suppliers V, O, AL,
    # BM and T already came out at a median of 1.000 in check C2.
    #
    # This rank cost dearly before it was spotted: at Supplier AE,
    # nineteen prices were divided a second time and came out at
    # EUR 0.0050 against a DPA of EUR 0.50 (example values). Scope
    # coverage lost 91 articles as a result.
    #
    # So the schema's pack size feeds `minimum_de_commande`, which serves
    # replenishment, and nothing else.

    # Rank 2 - the price list label
    depuis_libelle, nom = diviseur_depuis_designation(designation)
    if depuis_libelle:
        return float(depuis_libelle), "désignation", nom

    # Rank 3 - the WMS article record
    if (conditionnement_wms is not None
            and not pd.isna(conditionnement_wms)
            and float(conditionnement_wms) > 1):
        return (float(conditionnement_wms), "WMS",
                "conditionnement de la fiche article")

    # Rank 4 - the inference, only if it was corroborated upstream
    if deduit:
        return float(deduit), "déduit", preuve_deduite

    return 1.0, "", ""


# --------------------------------------------------------------------------
# Application


def configuration(nb_prix_pour_la_reference: int, a_un_prix_unitaire: bool) -> str:
    """A, B or C - the brief's sorting test, written down once and for all."""
    if nb_prix_pour_la_reference <= 1:
        return "B"
    return "A" if a_un_prix_unitaire else "C"


def appliquer(table: pd.DataFrame,
              colonne_prix: str = "prix_achat_unitaire_ht",
              colonne_palier: str | None = "paliers") -> pd.DataFrame:
    """Divides prices by their pack size, and records everything it does.

    The table must already carry `diviseur_conditionnement` and
    `source_diviseur`. The division and the multiplication of the price
    break happen HERE, in the same function and with the same N: separating
    them would risk one day having one applied without the other.
    """
    t = table.copy()
    diviseur = pd.to_numeric(t["diviseur_conditionnement"],
                             errors="coerce").fillna(1.0)
    prix = pd.to_numeric(t[colonne_prix], errors="coerce")

    t["prix_avant_division"] = prix
    # A divisor BELOW 1 is a multiplication: the price list quotes by the
    # unit while the WMS counts by the lot. Dividing by 1/10 amounts to
    # multiplying by 10, and the same formula serves both directions.
    a_convertir = (diviseur > 1) | ((diviseur > 0) & (diviseur < 1))

    # Case A: the extractor already picked a unit column, there is nothing
    # to divide. We say so in a column rather than let it look like an
    # oversight.
    if "config_conditionnement" in t.columns:
        cas_a = t["config_conditionnement"] == "A"
        a_convertir = a_convertir & ~cas_a

    a_diviser = a_convertir
    t[colonne_prix] = prix.where(~a_convertir, prix / diviseur)

    # Case C, and only case C: the price breaks are expressed in numbers of
    # packs, so they convert to units with the same N.
    if colonne_palier and colonne_palier in t.columns and "config_conditionnement" in t.columns:
        cas_c = (t["config_conditionnement"] == "C") & a_diviser
        if cas_c.any():
            palier = pd.to_numeric(t.loc[cas_c, colonne_palier], errors="coerce")
            t.loc[cas_c, colonne_palier] = palier * diviseur[cas_c]

    # The pack size remains a minimum order quantity: you cannot buy 3
    # units of a product sold by 100. That constraint serves
    # replenishment, not the price break, so it gets its own column.
    t["minimum_de_commande"] = diviseur.where(diviseur > 1)
    return t


# --------------------------------------------------------------------------
# Integration into the matching cascade


def appliquer_au_rapprochement(fusion: pd.DataFrame,
                               colonne_prix="prix_achat_unitaire_ht",
                               colonne_dpa="pa_wms") -> pd.DataFrame:
    """Sets the divisor on a matching table, and divides what must be divided.

    Called from pa_completer.py, once the price list is matched and the DPA
    joined. It adds five columns and only changes the price when it can say
    WHY. Without these columns, a wrong price cannot be detected
    afterwards.

    Deliberate caution on price breaks: when the price list carries breaks,
    we cannot tell from here whether they count units (case A, nothing to
    divide) or packs (case C, where the break would also have to be
    multiplied). The two look alike and go wrong silently, so we REPORT
    instead of deciding. Case C is rare; getting it wrong is not.
    """
    t = fusion.copy()
    n = len(t)
    vide = pd.Series([None] * n, index=t.index, dtype="object")
    t["diviseur_conditionnement"] = 1.0
    t["source_diviseur"] = ""
    t["preuve_diviseur"] = ""
    t["config_conditionnement"] = ""
    t["prix_avant_division"] = pd.to_numeric(t.get(colonne_prix), errors="coerce")

    if colonne_prix not in t.columns:
        return t

    cond_tarif = (t["conditionnement"] if "conditionnement" in t.columns else vide)
    suppose = (t["conditionnement_suppose"]
               if "conditionnement_suppose" in t.columns
               else pd.Series([False] * n, index=t.index))
    desi_tarif = (t["designation"] if "designation" in t.columns
                  else t.get("Désignation tarif", vide))
    desi_wms = t.get("Libellé déclinaison ^(1)", vide)

    # The WMS's DPA is the only counter-check available row by row. Without
    # it, a division read from a label has nothing to be verified against.
    prix_lu = pd.to_numeric(t.get(colonne_prix), errors="coerce")
    dpa_lu = (pd.to_numeric(t[colonne_dpa], errors="coerce")
              if colonne_dpa in t.columns
              else pd.Series([float("nan")] * n, index=t.index))

    # --- ranks 1 to 3, row by row ---
    diviseurs, sources, preuves, ecartees = [], [], [], []
    for i in t.index:
        # Has the extractor ALREADY brought this price back to the unit?
        # That is the case as soon as it fills in a pack size read from the
        # price list - at Supplier O, the "Nbre unités par boîte" column.
        # Re-parsing "Boîte de 30" out of the designation would then divide
        # a second time: penile sheath REF-EXEMPLE-3 fell from EUR 1.900 to
        # EUR 0.0633 (example values), the same error as at Supplier AE.
        deja_divise = False
        valeur_cond = cond_tarif.get(i)
        if (valeur_cond is not None and not pd.isna(valeur_cond)
                and float(valeur_cond) > 1 and not bool(suppose.get(i, False))):
            deja_divise = True

        d, s, p = choisir_diviseur(
            cond_tarif.get(i), suppose.get(i, False),
            None if deja_divise else desi_tarif.get(i), None)
        # The WMS label serves as a counter-check: if it announces the same
        # pack size as the price list, the two count the same thing.
        if s == "désignation":
            d2, s2 = diviseur_depuis_designation(desi_tarif.get(i),
                                                 desi_wms.get(i))
            if not d2:
                d, s, p = 1.0, "", ""
            else:
                d = float(d2)

        # SECOND counter-check, and this is the one that decides: the RATIO
        # TO THE DPA.
        #
        # The label says what the row represents; it does not say that the
        # WMS counts differently. Measurement of 16/09 over the full pass:
        # of thirteen divisions derived from a label, ELEVEN started from a
        # price already equal to the DPA - often to the cent - and divided
        # it anyway.
        #
        #   Supp. AF  "Téterelles MODELE EXEMPLE 5 XL (Lot de 2)"
        #            DPA 6.00  price 6.00  -> proposed 3.00. The WMS's DPA
        #            was already the one for the lot of two.
        #   Supp. K   "Rétroviseur (la paire) pour MODELE EXEMPLE 6 X4"
        #            divided by FOUR: the "X4" is the model name.
        #            DPA 15.00  price 15.00  -> proposed 3.75.
        #   Supp. BK  "Séparateurs x6 Petits": DPA 9.00 for all six.
        #            (prices: example values)
        #
        # The two divisions that survive are the ones the ratio
        # corroborates: Supplier J's "12X1L" starts at 12.01 times the DPA
        # and comes back to 1.001; Supplier D's "sachet de 4" starts at
        # 2.67.
        #
        # Deliberately the same judge as at rank 4: a division proves
        # itself the same way whatever source suggested it.
        ecartee = ""
        # A hand-made match escapes the counter-check: the person validated
        # both the price list line AND the expected price, which is worth
        # more than a ratio. See pa_correspondances.
        validee = bool(t.at[i, "correspondance_validee"]) \
            if "correspondance_validee" in t.columns else False
        if s == "désignation" and d > 1 and not validee:
            rapport = None
            if pd.notna(dpa_lu.get(i)) and float(dpa_lu.get(i)) > 0 \
                    and pd.notna(prix_lu.get(i)):
                rapport = float(prix_lu.get(i)) / float(dpa_lu.get(i))
            if rapport is None:
                # No DPA: nothing corroborates, nothing contradicts. We
                # divide - the label remains an indication - but we SAY so,
                # as at rank 4, so that review knows where to look.
                p = f"{p} — aucun DPA, division non corroborée"
            elif not suit_le_facteur(rapport, d):
                ecartee = (f"division par {d:g} écartée : le tarif est à "
                           f"{rapport:.3f}× le DPA, ce qui ne soutient pas "
                           f"un lot de {d:g} — lu « {p} »")
                d, s, p = 1.0, "", ""

        # The reverse case, and it is just as costly: the price list quotes
        # by the unit while the WMS counts by the LOT. "POCHES MODELE
        # EXEMPLE 7 1,5 L 10 U" has a DPA of EUR 16.00 for ten bags
        # (example value); proposing EUR 1.60 would value it ten times too
        # low. So we MULTIPLY instead of dividing.
        if d == 1.0 and deja_divise:
            lot_wms, nom_lot = _conditionnement_brut(desi_wms.get(i))
            if lot_wms and lot_wms > 1 and abs(lot_wms - float(valeur_cond)) < 0.01:
                d = 1.0 / float(lot_wms)
                s = "WMS au lot"
                p = (f"le WMS compte par {int(lot_wms)} ({nom_lot}) et le tarif "
                     f"à l'unité — prix REMULTIPLIÉ, pas divisé")

        diviseurs.append(d)
        sources.append(s)
        preuves.append(p)
        ecartees.append(ecartee)
    t["diviseur_conditionnement"] = diviseurs
    t["source_diviseur"] = sources
    # A refused division does not disappear: it is readable in a column.
    # Without that, the row would look like a row no rule ever touched, and
    # nobody would know a pack size had been seen there and set aside.
    t["division_ecartee"] = ecartees
    t["preuve_diviseur"] = preuves

    # --- rank 4: the inference, wholesale and only if it is systematic ---
    prix = pd.to_numeric(t[colonne_prix], errors="coerce")
    dpa = pd.to_numeric(t.get(colonne_dpa), errors="coerce") if colonne_dpa in t.columns else None
    if dpa is not None:
        sans_source = t["source_diviseur"] == ""
        rapports = (prix / dpa).where(sans_source & (dpa > 0) & (prix > 0))
        deduit, preuve, combien = deduire_par_systematicite(rapports)
        if deduit:
            # An article whose WMS label ALREADY announces a lot does not
            # enter the inference: its ratio is explained otherwise.
            deja_au_lot = desi_wms.map(lambda x: _conditionnement_brut(x)[0] or 0)
            eligible = sans_source & (deja_au_lot <= 1)

            # The factor is established for the PRICE LIST, but it does not
            # apply blindly to every row. Three situations, three
            # treatments:
            #
            #   - the article's ratio matches the factor -> we divide, and
            #     the article corroborates the division itself;
            #   - the article has NO DPA -> no counter-check is possible,
            #     but the price list is established at the lot: we divide
            #     and say so. Without that, the CHUT MODELE EXEMPLE 8 XTRA
            #     came out at EUR 60.00 (example value) - the price
            #     of the pair - for the sole reason that they had never
            #     been bought;
            #   - the ratio is close to 1 -> this article is ALREADY
            #     unitary in the price list. Dividing it would make a price
            #     twice too low. That is the mixed price list, and it gets
            #     reported instead of forced.
            colle = rapports.map(lambda r: suit_le_facteur(r, deduit))
            sans_dpa = eligible & rapports.isna()
            deja_unitaire = eligible & rapports.notna() & ~colle

            vise = eligible & colle
            t.loc[vise, "diviseur_conditionnement"] = float(deduit)
            t.loc[vise, "source_diviseur"] = "déduit"
            t.loc[vise, "preuve_diviseur"] = preuve

            t.loc[sans_dpa, "diviseur_conditionnement"] = float(deduit)
            t.loc[sans_dpa, "source_diviseur"] = "déduit (sans DPA)"
            t.loc[sans_dpa, "preuve_diviseur"] = (
                preuve + " — article sans DPA : diviseur du tarif appliqué, "
                         "non corroboré individuellement")

            t.loc[deja_unitaire, "preuve_diviseur"] = (
                f"NON divisé : le fournisseur cote au lot de {deduit}, mais le "
                f"rapport au DPA de cet article ne le suit pas — ligne "
                f"probablement déjà unitaire au tarif")
            print(f"  pack size inferred /{deduit} - {int(vise.sum())} "
                  f"articles corroborated, {int(sans_dpa.sum())} without a DPA, "
                  f"{int(deja_unitaire.sum())} left as they are")
        elif preuve:
            print(f"  pack size NOT inferred: {preuve}")

    # --- configuration A / B / C ---
    a_des_paliers = pd.Series(False, index=t.index)
    for nom in ("paliers", "Paliers"):
        if nom in t.columns:
            a_des_paliers |= t[nom].notna() & (t[nom].astype(str).str.strip() != "")
    t["config_conditionnement"] = "B"
    t.loc[a_des_paliers, "config_conditionnement"] = "A ou C — à vérifier"

    # We only divide case B: a single price, per pack.
    a_diviser = (pd.to_numeric(t["diviseur_conditionnement"],
                               errors="coerce").fillna(1) > 1)
    bloque = a_diviser & a_des_paliers
    if bloque.any():
        t.loc[bloque, "preuve_diviseur"] = (
            t.loc[bloque, "preuve_diviseur"]
            + " — NON APPLIQUÉ : le tarif porte des paliers, vérifier s'ils "
              "comptent des unités (cas A) ou des conditionnements (cas C)")
        print(f"  {int(bloque.sum())} divisors not applied: price breaks in "
              f"the price list, case A or C to be settled")
        t.loc[bloque, "diviseur_conditionnement"] = 1.0

    t = appliquer(t, colonne_prix=colonne_prix, colonne_palier=None)
    applique = (pd.to_numeric(t["diviseur_conditionnement"],
                              errors="coerce").fillna(1) > 1)
    if applique.any():
        print(f"  {int(applique.sum())} prices brought back to the unit "
              f"(sources: {t.loc[applique, 'source_diviseur'].value_counts().to_dict()})")
    return t


# --------------------------------------------------------------------------
# Diagnostic: report, apply nothing


def diagnostic(motif: str | None = None) -> pd.DataFrame:
    """What the pack size workstream would find, supplier by supplier.

    Reads the per-article state, hence the prices ALREADY matched,
    including those the x10 guard rail refused: those are precisely the
    ones a divisor can unblock.
    """
    from extracteurs.base import chemin_lisible

    chemin = RACINE / "sortie" / "3-achats" / "pa_etat_par_article.xlsx"
    if not chemin.exists():
        print("pa_etat_par_article.xlsx missing: run run_pa.py first")
        return pd.DataFrame()
    d = pd.read_excel(chemin_lisible(chemin), sheet_name=0, header=3)

    d["ancien"] = pd.to_numeric(d["Ancien PA"], errors="coerce")
    d["neuf"] = pd.to_numeric(d["Nouveau PA"], errors="coerce")
    d["ecarte"] = pd.to_numeric(d["PA écarté (à vérifier)"], errors="coerce")
    # The price list price, whether it was kept or refused by the guard rail.
    d["tarif"] = d["neuf"].fillna(d["ecarte"])
    if motif:
        d = d[d["Fournisseur"].fillna("").str.upper().str.contains(motif.upper())]

    lignes = []
    for nom, sous in d.groupby("Fournisseur"):
        comparables = sous[(sous["ancien"] > 0) & (sous["tarif"] > 0)]
        if comparables.empty:
            continue
        rapports = comparables["tarif"] / comparables["ancien"]
        mediane = rapports.median()
        deduit, preuve, combien = deduire_par_systematicite(rapports)

        # Careful: pa_etat_par_article only carries the WMS designation,
        # not the price list one. So what we read here says what the WMS
        # counts - and a WMS article already labelled "LA PAIRE" or
        # "BTE/50" is ALREADY at the lot:
        # it must on no account be divided. The column serves to recognise
        # those cases, never to propose a divisor.
        cote_wms = (sous["Désignation"].map(
            lambda x: _conditionnement_brut(x)[0]).dropna())
        libelle_dominant = (int(cote_wms.mode().iloc[0])
                            if not cote_wms.empty else None)
        part_au_lot = len(cote_wms) / len(sous) if len(sous) else 0

        lignes.append({
            "Fournisseur": nom,
            "Comparables": len(comparables),
            "Médiane": round(mediane, 3),
            "Hors fourchette": not (0.70 <= mediane <= 1.30),
            "Diviseur déduit": deduit,
            "Articles concernés": combien,
            "WMS déjà au lot": libelle_dominant,
            "Part au lot": round(part_au_lot, 2),
            "Médiane après division": (round(mediane / deduit, 3)
                                       if deduit else None),
            "Preuve": preuve,
        })

    res = pd.DataFrame(lignes)
    if res.empty:
        return res
    return res.sort_values(["Hors fourchette", "Comparables"], ascending=False)


def a_trancher() -> pd.DataFrame:
    """The articles where the pack size calls for a human decision.

    Three signals, all drawn from what we already have to hand:

      1. The WMS label says "UNITE" or "UNITAIRE" while the price kept is
         several times the DPA. This is the surest signal: at Supplier I,
         the 16 articles whose unit price beats the carton price ALL carry
         "UNITE" in their WMS label.
      2. The WMS label announces a lot ("BTE/50", "LOT DE 10") and the
         price is far below the DPA: the price list is then quoting by the
         unit and what is needed is a MULTIPLICATION, not a division.
      3. The ratio to the DPA lands near a round integer without any source
         having justified a division.

    We decide nothing: we lay it out.
    """
    from extracteurs.base import chemin_lisible

    chemin = RACINE / "sortie" / "3-achats" / "pa_etat_par_article.xlsx"
    if not chemin.exists():
        print("pa_etat_par_article.xlsx missing: run run_pa.py first")
        return pd.DataFrame()
    d = pd.read_excel(chemin_lisible(chemin), sheet_name=0, header=3)

    d["ancien"] = pd.to_numeric(d["Ancien PA"], errors="coerce")
    d["neuf"] = pd.to_numeric(d["Nouveau PA"], errors="coerce")
    d["ecarte"] = pd.to_numeric(d["PA écarté (à vérifier)"], errors="coerce")
    d["tarif"] = d["neuf"].fillna(d["ecarte"])
    d = d[(d["ancien"] > 0) & (d["tarif"] > 0)].copy()
    d["rapport"] = d["tarif"] / d["ancien"]

    libelle = d["Désignation"].fillna("").str.upper()
    dit_unite = libelle.str.contains(r"\bUNITE\b|\bUNITAIRE\b", regex=True)
    lot_wms = d["Désignation"].map(lambda x: _conditionnement_brut(x)[0] or 0)

    lignes = []
    for i in d.index:
        r = d.at[i, "rapport"]
        signal = ""
        propose = None
        if dit_unite[i] and r >= 2:
            signal = ("le WMS dit « unité » mais le prix vaut "
                      f"{r:.1f}× le DPA — le tarif cote sans doute au lot")
            propose = entier_rond_le_plus_proche(r) or round(r)
        elif lot_wms[i] > 1 and r <= 0.7:
            signal = (f"le WMS compte par {int(lot_wms[i])} et le prix vaut "
                      f"{r:.2f}× le DPA — le tarif cote à l'unité, il faudrait "
                      f"MULTIPLIER")
            propose = -int(lot_wms[i])
        elif r >= 1.8:
            rond = entier_rond_le_plus_proche(r)
            if rond:
                signal = (f"rapport ×{r:.2f}, proche de {rond} — "
                          f"conditionnement possible, non confirmé")
                propose = rond
        if not signal:
            continue
        lignes.append({
            "Code article": d.at[i, "Code article"],
            "Désignation": d.at[i, "Désignation"],
            "Fournisseur": d.at[i, "Fournisseur"],
            "Ancien PA": round(d.at[i, "ancien"], 4),
            "Prix retenu": round(d.at[i, "tarif"], 4),
            "Rapport": round(r, 2),
            "Diviseur proposé": propose,
            "Prix si appliqué": (round(d.at[i, "tarif"] / propose, 4)
                                 if propose and propose > 0
                                 else (round(d.at[i, "tarif"] * -propose, 4)
                                       if propose else None)),
            "À trancher": signal,
            "Fichier source": d.at[i, "Fichier source PA"],
        })
    res = pd.DataFrame(lignes)
    if res.empty:
        return res
    # Biggest gap first: that is where the error costs the most.
    return res.sort_values("Rapport", ascending=False)


def ecrire_a_trancher() -> Path | None:
    import mise_en_forme
    from main import SORTIE_ACHATS

    res = a_trancher()
    if res.empty:
        print("no case to settle")
        return None
    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_ACHATS / "pa_conditionnement_a_trancher.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        res.to_excel(w, sheet_name="À trancher", index=False, startrow=3)
    mise_en_forme.formater(chemin, {
        "À trancher": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Conditionnement — les cas qui demandent un œil",
            "sous_titre": "Un diviseur négatif veut dire MULTIPLIER : le "
                          "WMS compte par lot, le tarif cote à l'unité"},
    })
    print(f"  {len(res)} cases to settle -> {chemin}")
    return chemin


def main() -> None:
    if "--a-trancher" in sys.argv:
        ecrire_a_trancher()
        return
    motif = sys.argv[1] if len(sys.argv) > 1 else None
    res = diagnostic(motif)
    if res.empty:
        print("nothing to diagnose")
        return

    pd.set_option("display.width", 220)
    hors = res[res["Hors fourchette"]]
    print(f"{len(res)} comparable suppliers, "
          f"{len(hors)} outside the 0.70-1.30 band")
    print()
    if not hors.empty:
        print("=== OUTSIDE THE BAND - check C2 ===")
        print(hors[["Fournisseur", "Comparables", "Médiane", "Diviseur déduit",
                    "Articles concernés", "Médiane après division",
                    "WMS déjà au lot", "Part au lot"]].to_string(index=False))
        print()

    proposables = res[res["Diviseur déduit"].notna()]
    if not proposables.empty:
        print("=== DIVISORS CORROBORATED BY SYSTEMATIC RECURRENCE ===")
        for _, r in proposables.iterrows():
            # If the WMS ALREADY counts by the lot, the ratio does not come
            # from there and dividing would be a mistake. We say so instead
            # of proposing.
            reserve = ""
            if r["WMS déjà au lot"] and r["Part au lot"] >= 0.5:
                reserve = (f"  [WARNING: the WMS is already at the lot of "
                           f"{int(r['WMS déjà au lot'])} on "
                           f"{r['Part au lot']:.0%} of the articles - do not "
                           f"divide without checking]")
            print(f"  {r['Fournisseur'][:44]:<46} /{int(r['Diviseur déduit']):<4} "
                  f"{r['Médiane']:>7.3f} -> {r['Médiane après division']:<7.3f}{reserve}")
            print(f"      {r['Preuve']}")
    else:
        print("No corroborated divisor: nothing applies wholesale.")

    print()
    print("This diagnostic REPORTS. Nothing is applied here.")


if __name__ == "__main__":
    main()
