# -*- coding: utf-8 -*-
"""
Extractor for FOURNISSEUR B's CONDITIONS (Excel), the 2026 price list.

WHY THIS FILE RATHER THAN THE PDF
    The project used to read "tarif_fournisseur.pdf" and got only 52 rows
    out of it, leaving 130 in-scope items without a price. The same price
    list exists as Excel ("tarif_fournisseur.xlsx") with twelve product
    tabs and roughly 195 priced references.

    It was being discarded for a reason unrelated to its contents: its
    NAME says "2025", so the inventory assigned it the 2025 edition,
    while its first page is titled "CONDITIONS 2026 (applicable au 1er
    decembre 2025 sous nouvelle nomenclature)". The PDF, by contrast, had
    been picked by hand. This is the edition-year trap documented in the
    project rules, in its most expensive form.

THE LAYOUT, WHICH IS "A BIT OF A MESS"
    Each tab mixes three kinds of row:

      - a nomenclature CATEGORY row, with no price
        "FMP - Fauteuil non-modulaire a propulsion manuelle ou a pousser"
      - a PRODUCT row, with its price at quantity 1
        "MODELE EXEMPLE - Accoudoirs relevables"   150,00  145,00  3*
      - one or more TIER rows, with no label, which attach to the product
        above
        (empty)                                            140,00  6*

    (labels and amounts above: sample values)

    A product whose tier price was read instead of its unit price would
    be valued too low: that is the 08/09 incident at the pricing
    manager's, in reverse. So we take the "Qte 1" column, and the tiers
    go into the `paliers` column, for the record.

THE PERCENTAGE TRAP
    Some rows carry not a price but a DISCOUNT RATE off the list price,
    which sits in the last column. MODELE EXEMPLE has no price: it has
    "15 %" and a list price of 600,00, hence 510,00.

    Three notations coexist, and it is the cell FORMAT that tells them
    apart, not their value:

        150             format "0 €"    -> a price
        0,15            format "0%"     -> a rate of 15 %
        "10+5%"         text            -> a COMPOUND discount

    The compound discount is the real trap. "10+5%" is not 15 % but
    14.5 %: 1 - (0.90 x 0.95). The file demonstrates this itself, writing
    elsewhere "14,5% (10+5%)", "19% (10+10%)", "24% (20+5%)", the same
    rule three times over. An early version read "10+1%" as 1 % and
    valued a wheelchair at 99 % of its list price.

    Without the format, 0,15 could have passed for fifteen cents.

    (every rate and amount in this docstring: sample values)

NO SUPPLIER REFERENCE
    These tabs carry none, just like Supplier E. Matching will therefore
    go through the label, with the precautions taken in pa_libelle.py.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from extracteurs.base import chemin_lisible, finaliser, nettoyer_texte

FOURNISSEUR = "FOURNISSEUR B"

# Tabs that carry no product at all.
HORS_PRODUITS = {
    "page de garde", "administratifs", "cgv", "sommaire", "feuil1",
}

# "151.5", "151,50": a price. At most two decimals are allowed.
PRIX = re.compile(r"^\d{1,6}(?:[.,]\d{1,2})?$")
# "14,5% (10+5%)", "24% (20+5%)"  # sample values
POURCENT = re.compile(r"(\d{1,2}(?:[.,]\d+)?)\s*%")
# "3*", "9*", "3 et+", "2"
QUANTITE = re.compile(r"(\d{1,4})")

# Below this amount, a value is not a wheelchair price but a rate that
# Excel stored as a number (0.25 for 25 %).
PLANCHER_PRIX = 1.0


def _texte(valeur) -> str:
    if valeur is None or (isinstance(valeur, float) and pd.isna(valeur)):
        return ""
    return re.sub(r"\s+", " ", str(valeur)).strip()


def _nombre(valeur):
    brut = _texte(valeur).replace(" ", "").replace("\xa0", "")
    brut = brut.replace(" ", "").replace(",", ".")
    if not brut:
        return None
    try:
        return float(brut)
    except ValueError:
        return None


def _taux_depuis_texte(brut: str) -> float | None:
    """The discount rate a piece of text announces, as a percentage.

    Three forms, in decreasing order of reliability:

        "14,5% (10+5%)"   the rate is written before the bracket: read it
        "10+5%"           COMPOUND discount: 1 - (0.90 x 0.95) = 14.5 %
        "15%"             a simple rate

    (rates above: sample values)
    """
    texte = brut.replace(",", ".")
    # Form "X% (...)": the total is already computed, do not recompute.
    avant = re.match(r"\s*(\d{1,2}(?:\.\d+)?)\s*%\s*\(", texte)
    if avant:
        return float(avant.group(1))

    # Compound form "10+5%", "10+5+2%"  # sample values
    if "+" in texte:
        parts = re.findall(r"(\d{1,2}(?:\.\d+)?)", texte)
        if len(parts) >= 2:
            reste = 1.0
            for part in parts:
                valeur = float(part)
                if not 0 < valeur < 100:
                    return None
                reste *= (1 - valeur / 100)
            return round((1 - reste) * 100, 4)
        return None

    simple = POURCENT.search(texte)
    if simple:
        return float(simple.group(1))
    return None


def _prix_ou_taux(valeur, format_cellule, public):
    """(net price, how it was obtained), or (None, reason).

    `format_cellule` is the Excel format of the cell. THAT is what says
    whether 0,15 means fifteen cents or fifteen percent: the value alone
    does not say so, and guessing from a threshold only worked by luck.

    The second element is kept in French: it lands in the colonne_prix
    output column of the reconciliation report.
    """
    brut = _texte(valeur)
    if not brut:
        return None, ""

    format_cellule = format_cellule or ""
    est_pourcent = "%" in format_cellule
    valeur_num = _nombre(brut)

    # Text carrying a "%" is a rate, whatever the format says.
    if "%" in brut:
        taux = _taux_depuis_texte(brut)
        if taux is None:
            return None, "taux illisible"
        if public is None:
            return None, "taux sans prix public"
        return round(public * (1 - taux / 100), 4), f"public - {taux:g}%"

    # A numeric value in a percentage-formatted cell IS a rate: 0,15
    # under the "0%" format displays as "15 %".
    if est_pourcent and valeur_num is not None:
        if not 0 < valeur_num < 1:
            return None, "taux hors bornes"
        if public is None:
            return None, "taux sans prix public"
        return (round(public * (1 - valeur_num), 4),
                f"public - {valeur_num * 100:g}%")

    if valeur_num is None:
        return None, ""
    # That leaves the ordinary case: a price, in a euro-formatted cell.
    if valeur_num < PLANCHER_PRIX:
        return None, "valeur trop faible pour un prix"
    return round(valeur_num, 4), "tarif net unitaire HT"


def _colonnes(brut: pd.DataFrame) -> dict | None:
    """Locates the header row and maps each role to an index.

    We read the HEADINGS rather than fixed positions: the tabs do not all
    share the same width, and an inserted column would silently shift
    every price.
    """
    for ligne in range(min(12, len(brut))):
        valeurs = [_texte(v).lower() for v in brut.iloc[ligne]]
        joint = " | ".join(valeurs)
        if "prix public" not in joint:
            continue
        roles = {"entete": ligne}
        for index, valeur in enumerate(valeurs):
            if "prix public" in valeur:
                roles["public"] = index
            elif "qt" in valeur and "partir" in valeur:
                roles["quantite"] = index
            elif "qt" in valeur and ("1" in valeur or "unitaire" in valeur):
                roles.setdefault("prix", index)
            elif "tarif net" in valeur or "unitaire net" in valeur:
                roles.setdefault("palier", index)
            elif "nomenclature" in valeur or "produits" in valeur:
                roles["designation"] = index
            elif "lpp" in valeur:
                roles.setdefault("lpp", index)
        if "prix" not in roles:
            continue
        # The label always comes before the first price column.
        roles.setdefault("designation", max(roles["prix"] - 1, 0))
        roles.setdefault("palier", roles["prix"] + 1)
        return roles
    return None


def extract(path) -> pd.DataFrame:
    chemin = Path(path)
    lisible = chemin_lisible(chemin)
    classeur = pd.ExcelFile(lisible)
    # data_only=True gives the cached value of formulas and, above all,
    # `number_format`, without which a rate cannot be told from a price.
    from openpyxl import load_workbook
    livre = load_workbook(lisible, data_only=True)

    lignes = []
    for feuille in classeur.sheet_names:
        if feuille.strip().lower() in HORS_PRODUITS:
            continue
        brut = pd.read_excel(lisible, sheet_name=feuille, header=None, dtype=str)
        if brut.empty or brut.shape[1] < 3:
            continue
        roles = _colonnes(brut)
        if roles is None:
            continue
        onglet = livre[feuille] if feuille in livre.sheetnames else None

        courant = None
        categorie = ""
        for position in range(roles["entete"] + 1, len(brut)):
            ligne = brut.iloc[position]

            def case(role):
                index = roles.get(role)
                if index is None or index >= len(ligne):
                    return None
                return ligne.iloc[index]

            def format_de(role):
                """The Excel format of the cell, or '' if not found."""
                index = roles.get(role)
                if onglet is None or index is None:
                    return ""
                cellule = onglet.cell(row=position + 1, column=index + 1)
                return cellule.number_format or ""

            designation = _texte(case("designation"))
            public = _nombre(case("public"))
            prix, origine = _prix_ou_taux(case("prix"), format_de("prix"), public)
            prix_palier, _ = _prix_ou_taux(case("palier"), format_de("palier"),
                                           public)
            quantite = _texte(case("quantite"))

            if designation and prix is not None:
                # New PRODUCT row
                courant = {
                    "fournisseur": FOURNISSEUR,
                    # The category joins the label: it is what carries
                    # FMP / FRM, which the WMS echoes in its own labels.
                    "designation": nettoyer_texte(
                        f"{categorie} | {designation}" if categorie
                        else designation),
                    "prix_achat_unitaire_ht": prix,
                    "tarif_public_ttc": None,
                    "colonne_prix": origine,
                    "fichier_source": chemin.name,
                    "_public_ht": public,
                    "_paliers": [f"1:{prix:.2f}"],
                    "code_lppr": nettoyer_texte(_texte(case("lpp"))) or None,
                }
                lignes.append(courant)
                if prix_palier is not None and quantite:
                    seuil = QUANTITE.search(quantite)
                    if seuil:
                        courant["_paliers"].append(
                            f"{seuil.group(1)}:{prix_palier:.2f}")
            elif not designation and prix_palier is not None and courant:
                # TIER row attached to the product above
                seuil = QUANTITE.search(quantite) if quantite else None
                if seuil:
                    courant["_paliers"].append(
                        f"{seuil.group(1)}:{prix_palier:.2f}")
            elif designation and prix is None:
                # Nomenclature CATEGORY row: it carries no price, it
                # closes the previous product and applies to every
                # product that follows.
                #
                # It is not decorative: "FMP - Fauteuil non-modulaire"
                # and "FRM - Fauteuil modulaire" distinguish two families
                # that the WMS names too (VPH FMP ..., VPH FRM ...).
                # Without it, a modular wheelchair can end up with the
                # price of a non-modular one.
                courant = None
                categorie = designation

    for ligne in lignes:
        paliers = ligne.pop("_paliers", [])
        ligne.pop("_public_ht", None)
        ligne["paliers"] = " | ".join(paliers) if len(paliers) > 1 else None

    return finaliser(lignes)


if __name__ == "__main__":
    import sys

    defaut = (Path(__file__).resolve().parent.parent / "catalogues"
              / "fournisseur_b" / "tarif_fournisseur.xlsx")
    cible = Path(sys.argv[1]) if len(sys.argv) > 1 else defaut
    table = extract(cible)
    pd.set_option("display.width", 230)
    pd.set_option("display.max_colwidth", 60)
    print(f"{len(table)} rows extracted from {cible.name}")
    colonnes = [c for c in ("designation", "prix_achat_unitaire_ht",
                            "colonne_prix", "paliers") if c in table.columns]
    print(table[colonnes].head(30).to_string(index=False))
