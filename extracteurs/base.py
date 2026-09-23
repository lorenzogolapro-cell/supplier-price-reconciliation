# -*- coding: utf-8 -*-
"""
Common foundation shared by every supplier catalogue extractor.

Provides:
  - the normalised output schema (COLONNES)
  - the cleaning helpers (text, number, EAN)
  - EAN-8 / EAN-13 check digit validation
  - automatic detection of the header row in an Excel sheet

Every extractor (extracteurs/<fournisseur>.py) must expose:
    extract(path) -> pandas.DataFrame
matching COLONNES exactly, in that order.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import unicodedata
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Normalised output schema, identical for every supplier
# ---------------------------------------------------------------------------

# The names are deliberately explicit: a price always says whether it is
# per unit or per case, a rate always says it is stored as a decimal.
# Order matters: the columns used day to day come first.
COLONNES = [
    "fournisseur",         # supplier name (e.g. FOURNISSEUR A)
    "ref_fournisseur",     # supplier catalogue reference (ALWAYS text)
    # Our own item code, when the supplier file happens to carry it.
    # That is rare (only files produced by an earlier reconciliation have
    # it) but it is the safest link to the WMS: it removes the need to go
    # through the supplier reference, which is precisely the identifier
    # that diverges most often.
    "code_article_wms",
    "designation",         # product label
    "ean",                 # EAN of the SELLING UNIT (always text)
    # Many price lists publish two codes: one for the piece and one for
    # the case (Supplier AE "Unite"/"CDT", Supplier T "sachet"/"carton",
    # Supplier CV goes as far as three: unit, box, carton). `ean` always
    # carries the selling unit, the one scanned at picking; the case code
    # is kept separately, it is used at goods-in.
    "ean_conditionnement",
    # --- purchase price -----------------------------------------------------
    "prix_achat_unitaire_ht",  # purchase price of ONE unit, excl. VAT
    # EXACT heading of the column this price came from, spelled as it is
    # written in the price list. A price list often publishes three prices
    # side by side (list, discounted, per case) and picking the wrong
    # column raises no error at all: it just yields a wrong purchase
    # price. This is the only way to check afterwards that we took the
    # right one.
    "colonne_prix",
    "conditionnement",         # number of units per case
    "prix_colis_ht",           # prix_achat_unitaire_ht x conditionnement
    "conditionnement_suppose",  # True if the pack size was missing -> 1
    # --- how the price is built, for the margin check -----------------------
    "tarif_public_ttc",    # recommended retail price, incl. VAT (never net)
    "tva_taux",            # VAT rate as a DECIMAL (0.2 = 20 %)
    "remise_taux",         # discount rate as a DECIMAL, applied to the net
    "prix_recalcule",      # (tarif_public_ttc / (1 + tva)) x (1 - remise)
    "ecart_prix",          # |prix_recalcule - prix_achat_unitaire_ht|
    "prix_coherent",       # True if the gap is below TOLERANCE_PRIX
    # --- volume tiers: these are UNIT PRICES, not lot prices ----------------
    "palier2_qte",         # from this quantity on, the unit price...
    "palier2_prix_ht",     # ... drops to this value
    "economie_palier2_pct",  # unit saving vs base price, as a decimal
    "palier3_qte",
    "palier3_prix_ht",
    "economie_palier3_pct",
    "paliers",             # readable summary: "1:30.40 | 6:28.70 | 75:17.60"
    # --- product attributes -------------------------------------------------
    "eco_part_ht",         # eco-contribution, excl. VAT
    "code_lppr",           # LPPR reimbursement code if present (text)
    "montant_lppr",        # LPPR amount if present
    "dispositif_medical",  # Y / N (medical device)
    "origine",             # country of origin
    "fichier_source",      # name of the file the row came from
]

# Tolerated gap between the net price in the file and the recomputed
# price, in euros.
TOLERANCE_PRIX = 0.05

# Columns that must stay text when written to Excel
COLONNES_TEXTE = [
    "fournisseur",
    "ref_fournisseur",
    "code_article_wms",
    "designation",
    "ean",
    "ean_conditionnement",
    "colonne_prix",
    "paliers",
    "code_lppr",
    "dispositif_medical",
    "origine",
    "fichier_source",
]

# Numeric columns (float). No rounding is applied here: amounts stay at
# full precision and are only rounded for display, by the Excel cell
# format.
COLONNES_NUM = [
    "prix_achat_unitaire_ht",
    "conditionnement",
    "prix_colis_ht",
    "tarif_public_ttc",
    "tva_taux",
    "remise_taux",
    "prix_recalcule",
    "ecart_prix",
    "palier2_qte",
    "palier2_prix_ht",
    "economie_palier2_pct",
    "palier3_qte",
    "palier3_prix_ht",
    "economie_palier3_pct",
    "eco_part_ht",
    "montant_lppr",
]

# Boolean columns
COLONNES_BOOL = ["prix_coherent", "conditionnement_suppose"]


# ---------------------------------------------------------------------------
# Reading the source files
# ---------------------------------------------------------------------------

def chemin_lisible(path) -> Path:
    """A usable path to read a workbook, even one open in Excel.

    Excel takes an exclusive lock on open files: any read then fails with
    PermissionError, which aborts the whole run for a reason that has
    nothing to do with the data. Copying, however, is still allowed, so we
    work on a temporary duplicate.

    Returns the original path when it can be read directly.
    """
    path = Path(path)
    try:
        with open(path, "rb"):
            return path
    except PermissionError:
        pass

    copie = Path(tempfile.gettempdir()) / f"catalogue_verrouille_{path.name}"
    shutil.copy2(path, copie)
    print(f"    (file open in Excel, reading a temporary copy)")
    return copie


# ---------------------------------------------------------------------------
# Cleaning helpers
# ---------------------------------------------------------------------------

def nettoyer_texte(valeur) -> str | None:
    """Turns any cell into clean text, or None when empty.

    Important detail: references and EANs sometimes come out of Excel as
    floats (812176.0, 3.760123e+12). We convert back to an integer before
    stringifying, so no stray '.0' or scientific notation is carried over.
    """
    if valeur is None:
        return None
    if isinstance(valeur, float):
        # NaN
        if valeur != valeur:
            return None
        # whole float -> drop the stray .0
        if valeur.is_integer():
            return str(int(valeur))
        return repr(valeur)
    if isinstance(valeur, int):
        return str(valeur)
    texte = str(valeur).strip()
    # Non-breaking and doubled spaces left behind by Excel exports
    texte = texte.replace(" ", " ")
    texte = re.sub(r"\s+", " ", texte)
    return texte or None


def nettoyer_nombre(valeur) -> float | None:
    """Turns a cell into a float, or None if empty / not convertible.

    Handles the comma decimal separator just in case (some catalogues
    export prices as French text), currency symbols, thousands spaces and
    percentages written '20 %' -> 0.2.
    """
    if valeur is None:
        return None
    if isinstance(valeur, bool):
        return None
    if isinstance(valeur, (int, float)):
        if isinstance(valeur, float) and valeur != valeur:  # NaN
            return None
        return float(valeur)

    texte = str(valeur).strip()
    if not texte:
        return None

    pourcentage = "%" in texte
    # Keep only what can make up a number
    texte = texte.replace(" ", "").replace(" ", "")
    texte = texte.replace("€", "").replace("%", "")
    texte = texte.replace(",", ".")
    texte = re.sub(r"[^0-9.\-]", "", texte)
    if texte in ("", "-", ".", "-."):
        return None
    try:
        nombre = float(texte)
    except ValueError:
        return None
    return nombre / 100 if pourcentage else nombre


def nettoyer_ean(valeur) -> str | None:
    """Normalises an EAN code into TEXT.

    - runs through nettoyer_texte to kill scientific notation
    - strips everything that is not a digit (hyphens, spaces)
    - restores the leading zeros Excel dropped:
        11 or 12 digits -> 13 (US UPC-A codes, whose leading zero or
        zeros are systematically lost on export)
        7 digits        -> 8
      Zero-padding is only applied if it makes the check digit come out
      right; otherwise the value is kept as is, so that a genuine data
      entry error is not hidden.
    - folds back to EAN-13 the GTIN-14 codes whose packaging indicator is
      0: such a code does not designate a carton but the selling unit,
      written over fourteen positions. The check digit is the same on
      both sides, a leading zero weighing nothing in the weighted sum.
    - discards anything too short to be a barcode (cells holding "0",
      layout leftovers)
    Performs NO validity filtering: validity is assessed separately by
    ean_est_valide(), so it can be reported without dropping the row.
    """
    texte = nettoyer_texte(valeur)
    if texte is None:
        return None
    chiffres = re.sub(r"\D", "", texte)
    if not chiffres:
        return None

    if len(chiffres) in (11, 12):
        candidat = chiffres.zfill(13)
        if cle_controle_ean(candidat) == int(candidat[-1]):
            return candidat
    elif len(chiffres) == 7:
        candidat = chiffres.zfill(8)
        if cle_controle_ean(candidat) == int(candidat[-1]):
            return candidat
    elif len(chiffres) == 14 and chiffres[0] == "0":
        # Indicator 0: selling unit, not a packaging level
        candidat = chiffres[1:]
        if cle_controle_ean(candidat) == int(candidat[-1]):
            return candidat

    # Below eight digits no barcode exists: this is an empty cell or a
    # leftover, not data worth reporting.
    if len(chiffres) < 8:
        return None

    return chiffres


def cle_controle_ean(chiffres: str) -> int | None:
    """Computes the expected check digit for an EAN-8 or EAN-13.

    Standard GS1 algorithm: 3/1 weighted sum starting from the right
    (excluding the check digit), then complement to the next ten.
    """
    if not chiffres or not chiffres.isdigit() or len(chiffres) not in (8, 12, 13, 14):
        return None
    corps = chiffres[:-1]
    somme = 0
    # The weight alternates 3 then 1, going back from the last body digit
    for position, caractere in enumerate(reversed(corps)):
        poids = 3 if position % 2 == 0 else 1
        somme += int(caractere) * poids
    return (10 - somme % 10) % 10


def ean_est_valide(ean) -> bool:
    """True if the EAN is 8 or 13 digits long AND its check digit is right.

    A GTIN-14 (packaging code) is deliberately treated as NOT valid here:
    it is syntactically correct but does not designate the selling unit,
    so it is not what we want to push into the WMS. diagnostic_ean()
    tells it apart from a genuine error.
    """
    if not ean or not isinstance(ean, str):
        return False
    if not ean.isdigit() or len(ean) not in (8, 13):
        return False
    # "0000000000000" passes the check digit: a zero sum is divisible by
    # ten. It is nonetheless an empty box filled in by hand, not a code:
    # Supplier I's price list carries 48 of them, 34 of which would have
    # OVERWRITTEN a real code in the WMS. The check digit proves that a
    # number is internally consistent, never that it names a product.
    if set(ean) == {"0"}:
        return False
    return cle_controle_ean(ean) == int(ean[-1])


def diagnostic_ean(ean) -> str:
    """Explains why an EAN is rejected, to make the report actionable.

    Returns one short string among (values kept verbatim, they end up in
    the quality report):
      "OK"                      -> EAN-8 / EAN-13 with a valid check digit
      "absent"                  -> missing
      "caracteres non numeriques" -> non-numeric characters
      "GTIN-14 (colisage), cle correcte" -> carton code, not the selling
                                            unit
      "longueur N inattendue"   -> unexpected length N
      "cle de controle incorrecte (attendue X)" -> wrong check digit,
                                                   X expected
    """
    if ean is None or not isinstance(ean, str) or ean == "":
        return "absent"
    if not ean.isdigit():
        return "caracteres non numeriques"

    longueur = len(ean)
    cle_attendue = cle_controle_ean(ean)

    if longueur == 14:
        if cle_attendue != int(ean[-1]):
            return "longueur 14 et cle incorrecte"
        # The leading indicator tells the carton from the selling unit
        if ean[0] == "0":
            return "GTIN-14 d'unite (a ramener en EAN-13)"
        return "GTIN-14 (colisage), cle correcte"
    if longueur == 12:
        if cle_attendue == int(ean[-1]):
            return "UPC-A 12 (a completer en EAN-13)"
        return "longueur 12 et cle incorrecte"
    if longueur not in (8, 13):
        return f"longueur {longueur} inattendue"
    if cle_attendue != int(ean[-1]):
        return f"cle de controle incorrecte (attendue {cle_attendue})"
    return "OK"


# ---------------------------------------------------------------------------
# Pricing logic
# ---------------------------------------------------------------------------
#
# The supplier price computation chain, verified against Supplier A's 2026
# price list:
#
#     prix_ht_public = tarif_public_ttc / (1 + tva_taux)
#     prix_achat     = prix_ht_public * (1 - remise_taux)
#
# Examples: 480.00 / 1.20 = 400.00 ; x (1 - 0.30000) = 280.00   # sample values
#           527.50 / 1.055 = 500.00 ; x (1 - 0.40000) = 300.00  # sample values
#
# The list price in column 3 includes VAT, never excludes it. That is the
# number one source of error on this file.

def normaliser_conditionnement(valeur) -> tuple[float, bool]:
    """A usable pack size, plus a flag saying the value was assumed.

    A pack size that is empty, zero, negative or non-numeric is brought
    back to 1 (the product is then taken to be ordered by the unit), and
    the second element of the return is True so the assumption stays on
    record.
    """
    nombre = nettoyer_nombre(valeur)
    if nombre is None or nombre <= 0:
        return 1.0, True
    return float(nombre), False


def calculer_prix_colis(prix_unitaire_ht, conditionnement) -> float | None:
    """Price of a full case: unit price x number of units per case.

    No rounding: the value stays at full precision, rounding is the job
    of the Excel display format.
    """
    if prix_unitaire_ht is None or conditionnement is None:
        return None
    return prix_unitaire_ht * conditionnement


def controler_coherence_prix(
    tarif_public_ttc, tva_taux, remise_taux, prix_achat_unitaire_ht
) -> tuple[float | None, float | None, bool | None]:
    """Recomputes the theoretical purchase price and compares it to the file.

    Returns (prix_recalcule, absolute gap, coherent). A gap larger than
    TOLERANCE_PRIX signals either a supplier error or a special case (a
    negotiated price outside the grid): in both cases the row must be
    checked before an order is based on it.

    Returns (None, None, None) when one of the required values is
    missing: coherence can then be neither confirmed nor denied.
    """
    if (
        tarif_public_ttc is None
        or tva_taux is None
        or remise_taux is None
        or prix_achat_unitaire_ht is None
        or tva_taux <= -1
    ):
        return None, None, None

    prix_recalcule = (tarif_public_ttc / (1 + tva_taux)) * (1 - remise_taux)
    ecart = abs(prix_recalcule - prix_achat_unitaire_ht)
    return prix_recalcule, ecart, ecart < TOLERANCE_PRIX


def valider_paliers(prix_unitaire, paliers) -> list:
    """Keeps only the tiers that really are tiers.

    A volume tier implies two things: a quantity above 1, and a price
    below the unit price. Anything that fails those conditions is ANOTHER
    price column, most often the recommended retail price, which the
    supplier publishes next to its net price.

    Confusing the two would produce negative "discounts" and would skew
    the whole quantity-based purchasing analysis.

    `paliers` is a list of (quantity, price) pairs. The return is the same
    list, filtered and sorted by increasing quantity.
    """
    if prix_unitaire is None or prix_unitaire <= 0:
        return []

    retenus = [
        (quantite, prix)
        for quantite, prix in paliers
        if quantite is not None and prix is not None
        and quantite > 1 and 0 < prix < prix_unitaire
    ]
    return sorted(retenus, key=lambda couple: couple[0])


def economie_palier(prix_base, prix_palier) -> float | None:
    """Unit saving of a tier against the base price, as a decimal.

    A negative value signals a tier that is more expensive than the base
    price, that is, an anomaly in the supplier price list.
    """
    if prix_base is None or prix_palier is None or prix_base <= 0:
        return None
    return 1 - prix_palier / prix_base


def resumer_paliers(prix_unitaire, paliers, quantite_base=1) -> str | None:
    """Readable summary of the tier grid: "1:30.40 | 6:28.70 | 75:17.60".

    `paliers` is a list of (threshold quantity, unit price at that tier)
    pairs. This format is for human reading only: the palier2_qte /
    palier2_prix_ht columns remain the source for any computation.

    `quantite_base` is the quantity from which the unit price applies. It
    is 1 in most price lists, but some only quote from a lot upwards
    (Supplier L starts at 20 on walking sticks): printing "1:" would then
    be misleading.
    """
    if prix_unitaire is None:
        return None

    valides = [
        (quantite, prix)
        for quantite, prix in paliers
        if quantite is not None and prix is not None
    ]
    if not valides:
        return None

    valides.sort(key=lambda couple: couple[0])
    morceaux = [f"{quantite_base:g}:{prix_unitaire:.2f}"]
    morceaux += [f"{quantite:g}:{prix:.2f}" for quantite, prix in valides]
    return " | ".join(morceaux)


def anomalies_paliers(prix_base, palier2, palier3) -> list[str]:
    """Checks the internal consistency of the tier grid.

    `palier2` and `palier3` are (quantity, unit price) pairs. A sound grid
    is strictly decreasing in price as the quantity rises. Returns the
    list of anomalies found (French wording, it goes straight into the
    quality report), empty when all is well.
    """
    anomalies = []
    q2, p2 = palier2
    q3, p3 = palier3

    if prix_base is not None:
        for nom, prix in (("palier 2", p2), ("palier 3", p3)):
            if prix is not None and prix > prix_base:
                anomalies.append(f"{nom} plus cher que le prix de base")

    if q2 is not None and q3 is not None:
        if q3 <= q2:
            anomalies.append("quantite du palier 3 inferieure ou egale au palier 2")
            if p2 is not None and p3 is not None and p3 < p2:
                anomalies.append("palier 3 moins cher a quantite inferieure")
        elif p2 is not None and p3 is not None and p3 > p2:
            anomalies.append("palier 3 plus cher que le palier 2")

    return anomalies


def prix_pour_quantite(ligne, quantite: int) -> dict:
    """Best applicable price for a given quantity.

    `ligne` is a catalogue row (dict or pandas Series) following the
    common schema. The requested quantity is rounded up to the next
    multiple of the pack size: you cannot order 3 units of a product sold
    by 6. The tier is then picked from the quantity actually ordered, not
    from the one requested.

    Returns:
        prix_unitaire_ht    unit price of the applied tier
        total_ht            prix_unitaire_ht x quantite_commandee
        palier_applique     "base", "palier 2" or "palier 3"
        quantite_demandee   the quantity passed as an argument
        quantite_commandee  after rounding up to the pack size
        arrondi_conditionnement  True if the quantity had to be raised
        conditionnement     the pack size used
    """
    def valeur(champ):
        """Lenient read: dict, pandas Series, NaN values."""
        brut = ligne[champ] if champ in ligne else None
        if brut is None or (isinstance(brut, float) and brut != brut):
            return None
        return brut

    if quantite is None or quantite <= 0:
        raise ValueError("The quantity must be a strictly positive integer")

    prix_base = valeur("prix_achat_unitaire_ht")
    if prix_base is None:
        raise ValueError(
            f"No unit price for reference {valeur('ref_fournisseur')}"
        )

    conditionnement, _ = normaliser_conditionnement(valeur("conditionnement"))

    # Only whole cases can be ordered
    nb_colis = -(-quantite // conditionnement)  # ceiling integer division
    quantite_commandee = nb_colis * conditionnement
    arrondi = quantite_commandee != quantite

    # The best tier reached by the quantity actually ordered. We walk the
    # thresholds in increasing order and keep the last one reached, which
    # stays correct even if the grid is badly ordered.
    prix_unitaire = prix_base
    palier_applique = "base"
    candidats = [
        ("palier 2", valeur("palier2_qte"), valeur("palier2_prix_ht")),
        ("palier 3", valeur("palier3_qte"), valeur("palier3_prix_ht")),
    ]
    for nom, seuil, prix in sorted(
        (c for c in candidats if c[1] is not None and c[2] is not None),
        key=lambda c: c[1],
    ):
        if quantite_commandee >= seuil and prix < prix_unitaire:
            prix_unitaire = prix
            palier_applique = nom

    return {
        "prix_unitaire_ht": prix_unitaire,
        "total_ht": prix_unitaire * quantite_commandee,
        "palier_applique": palier_applique,
        "quantite_demandee": quantite,
        "quantite_commandee": quantite_commandee,
        "arrondi_conditionnement": arrondi,
        "conditionnement": conditionnement,
    }


# ---------------------------------------------------------------------------
# Reference normalisation, for matching against the WMS database
# ---------------------------------------------------------------------------

def cle_ref_stricte(valeur) -> str | None:
    """Strict matching key: upper case, no spaces.

    The trailing '.0' is stripped: it appears as soon as a purely numeric
    reference has passed through a float type.
    """
    texte = nettoyer_texte(valeur)
    if texte is None:
        return None
    texte = texte.upper().replace(" ", "")
    if texte.endswith(".0"):
        texte = texte[:-2]
    return texte or None


def cle_ref_souple(valeur) -> str | None:
    """Lenient matching key: as above, plus no . - _ / separators.

    Indispensable here: the WMS stores variants with a dot (123456.M,
    789012.B) where the supplier catalogue runs them together (123456M,
    789012N). That single tolerance alone gains about 250 matches on
    Supplier A.
    """
    texte = cle_ref_stricte(valeur)
    if texte is None:
        return None
    texte = re.sub(r"[.\-_/]", "", texte)
    return texte or None


# ---------------------------------------------------------------------------
# Header row detection
# ---------------------------------------------------------------------------

def _normaliser(texte: str) -> str:
    """Lower case, no accents, no punctuation, to compare headings."""
    texte = unicodedata.normalize("NFKD", texte)
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texte.lower())


def trouver_entete_df(brut, mots_cles, max_lignes: int = 60) -> int | None:
    """Same search as trouver_ligne_entete, but on a raw DataFrame.

    Returns the index (0-based, pandas) of the first row containing all
    the keywords, or None. Useful when the file was read with
    `header=None` rather than opened through openpyxl.
    """
    cles = [_normaliser(m) for m in mots_cles]
    for position in range(min(max_lignes, len(brut))):
        cellules = [
            _normaliser(str(valeur))
            for valeur in brut.iloc[position].tolist()
            if str(valeur) != "nan"
        ]
        if all(any(cle in cellule for cellule in cellules) for cle in cles):
            return position
    return None


def colonne_par_intitule(entetes, *fragments: str) -> int | None:
    """Position of the first column whose heading contains a fragment.

    Price list headings often span several lines and are riddled with
    typos ("Descirption"), hence the comparison on a normalised fragment
    rather than on equality.
    """
    cibles = [_normaliser(fragment) for fragment in fragments]
    for position, intitule in enumerate(entetes):
        texte = _normaliser(str(intitule))
        if any(cible in texte for cible in cibles):
            return position
    return None


def trouver_ligne_entete(feuille, mots_cles, max_lignes: int = 60) -> int:
    """Returns the number (1-based, openpyxl) of the header row.

    We look for the FIRST row that contains all the given keywords
    (comparison insensitive to case and accents). This avoids hard-coding
    an offset that would differ from one catalogue to the next.

    Raises ValueError if nothing is found in the first `max_lignes` rows.
    """
    cles = [_normaliser(m) for m in mots_cles]
    for ligne in feuille.iter_rows(min_row=1, max_row=max_lignes, values_only=False):
        cellules = [_normaliser(str(c.value)) for c in ligne if c.value is not None]
        if all(any(cle in cellule for cellule in cellules) for cle in cles):
            return ligne[0].row
    raise ValueError(
        f"Header row not found (keywords searched: {mots_cles})"
    )


# ---------------------------------------------------------------------------
# Finalising the DataFrame
# ---------------------------------------------------------------------------

def finaliser(lignes: list[dict]) -> pd.DataFrame:
    """Builds the final DataFrame: columns in order, with the right types.

    Text columns are forced to 'object' holding real strings, never
    numbers, so that writing to Excel preserves leading zeros and
    alphanumeric references (e.g. 123456N).
    """
    df = pd.DataFrame(lignes, columns=COLONNES)

    for colonne in COLONNES_TEXTE:
        df[colonne] = df[colonne].astype("object").where(df[colonne].notna(), None)

    for colonne in COLONNES_NUM:
        df[colonne] = pd.to_numeric(df[colonne], errors="coerce")

    # Booleans stay 'object': prix_coherent is None when coherence could
    # not be assessed, which a bool dtype would flatten away.
    for colonne in COLONNES_BOOL:
        df[colonne] = df[colonne].astype("object")

    return df
