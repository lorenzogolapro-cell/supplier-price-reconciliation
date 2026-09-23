# -*- coding: utf-8 -*-
"""
PA workstream: repair damaged identifiers before any fuzzy matching.

    python pa_identifiants.py               diagnostic over the whole scope
    python pa_identifiants.py FOURNISSEUR_E one supplier

WHY THIS COMES BEFORE THE LABEL
    A match proved by an EAN check digit is an EXACT match, not a candidate:
    the GS1 3/1 weighted sum only falls into place by chance one time in ten.
    So it can be applied, where a label match can only ever propose.

WHAT WE DO
    1. Several identifiers inside a single cell.
       "1234567890128+1234567890135": two valid EANs separated by a "+".
       We split them and try each one. No digit is ever added.

    2. GTIN-14 whose leading indicator is zero.
       All thirteen digits of the EAN are already there; we drop the padding.

WE REBUILD NO IDENTIFIER (decision of 16/09)
    Three repairs used to live here: a truncated EAN whose check digit we
    recomputed, a restored leading zero, a wrong check digit corrected. All
    removed. A check digit that falls into place proves less than it looks
    like: it does so by chance one time in ten, and over thousands of
    references that happens often enough to set prices that are both wrong
    and silent. A damaged identifier is REPORTED; it goes back to label
    matching, which proposes instead of applying.

WHAT WE DO NOT REPAIR EITHER
    The reference stem: the numeric radical of at least six digits, the one
    that brings "1234567HR" and "1234567" together at Supplier J. The price
    owner REPORTS it, never applies it, and we keep his rule: a range suffix
    often tells two real products apart. It only becomes usable once the
    price corroborates it, and it then comes out as "À RELIRE", never as
    something injectable.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

from extracteurs.base import (  # noqa: E402
    cle_ref_stricte,
    ean_est_valide,
)

# A "+" separates two identifiers; so do two or more spaces, when data entry
# has run two columns together. We stay deliberately narrow: a comma or a
# slash often belongs to the reference itself ("123.456").
SEPARATEURS = re.compile(r"\s*\+\s*|\s{2,}")


def morceaux(reference) -> list[str]:
    """The distinct identifiers a single cell carries."""
    if reference is None or (isinstance(reference, float) and pd.isna(reference)):
        return []
    texte = str(reference).strip()
    if not texte:
        return []
    parts = [p.strip() for p in SEPARATEURS.split(texte) if p and p.strip()]
    return parts or [texte]


def variantes_ean(valeur: str) -> list[tuple[str, str]]:
    """The plausible EANs this value could designate, with their proof.

    Returns ONLY codes whose check digit is correct. A candidate whose check
    digit does not fall into place is not an EAN: proposing it would amount
    to manufacturing a wrong and silent price.
    """
    texte = str(valeur or "").strip()
    if not texte:
        return []

    # A barcode contains ONLY digits. "1234-56-7.0" is a supplier reference:
    # stripping its hyphens to get eight digits, then computing a check digit
    # for them, proves nothing - it manufactures an EAN that never existed.
    # So we refuse anything that is not already numeric.
    if not texte.isdigit():
        return []
    chiffres = texte

    trouves: list[tuple[str, str]] = []

    def ajouter(code: str, preuve: str) -> None:
        if ean_est_valide(code) and all(code != c for c, _ in trouves):
            trouves.append((code, preuve))

    # As read
    ajouter(chiffres, "EAN lu tel quel")

    # GTIN-14 of a sales unit: the leading indicator is zero, all thirteen
    # digits of the EAN are already there. We remove one padding digit, we
    # invent none - that is why this case survives the rule.
    if len(chiffres) == 14 and chiffres[0] == "0":
        ajouter(chiffres[1:], "GTIN-14 d'unité ramené en EAN-13")

    # REMOVED ON 16/09, by decision: WE DO NOT REBUILD IDENTIFIERS.
    #
    # Three repairs used to live here, all of them proved by the check
    # digit, all of them deleted:
    #   - EAN truncated to 12 digits, check digit recomputed;
    #   - leading zero lost by Excel, restored with zfill;
    #   - check digit wrong by one digit, corrected.
    #
    # The argument was that a check digit only falls into place by chance one
    # time in ten. That is true and it is not enough: over thousands of
    # references, one time in ten happens often, and the price then set is
    # wrong AND silent. A damaged identifier gets reported, not repaired.
    #
    # What becomes of such a code: it stays without a price, and the next
    # match - label, manufacturer - takes it over by PROPOSING.

    return trouves


def candidats_identifiant(reference) -> list[tuple[str, str]]:
    """Every provable EAN behind a reference, with the cell split apart."""
    resultats: list[tuple[str, str]] = []
    parts = morceaux(reference)
    for i, part in enumerate(parts):
        prefixe = "" if len(parts) == 1 else f"cellule à {len(parts)} identifiants — "
        for code, preuve in variantes_ean(part):
            if all(code != c for c, _ in resultats):
                resultats.append((code, prefixe + preuve))
    return resultats


# --------------------------------------------------------------------------
# Reference stem: reported, never applied


MOTIF_SOCLE = re.compile(r"(\d{6,})")


def socle_reference(reference) -> str | None:
    """The numeric radical of at least six digits, as the price owner
    uses it.

    "1234567HR" and "1234567" share the stem "1234567". That is a clue, not
    a proof: at Supplier J, "UG" and "6C" tell apart sub-ranges that are not
    the same product.
    """
    if reference is None or (isinstance(reference, float) and pd.isna(reference)):
        return None
    trouve = MOTIF_SOCLE.search(str(reference))
    return trouve.group(1) if trouve else None


# --------------------------------------------------------------------------
# Applied to the matching cascade


def rapprocher_par_identifiant(fusion: pd.DataFrame,
                               tarif: pd.DataFrame,
                               champs: list[str]) -> pd.DataFrame:
    """Fifth pass: EANs AS READ, over whatever is still without a price.

    Called from pa_completer.py after the four existing keys. Since 16/09 it
    only matches on codes entirely present in the cell: multi-identifier
    cell, unpadded GTIN-14. No reconstruction any more, therefore no price
    set on an invented digit.
    """
    if "prix_achat_unitaire_ht" not in fusion.columns:
        return fusion
    absents = fusion["prix_achat_unitaire_ht"].isna()
    if not absents.any():
        return fusion

    # Index of the price list's EANs, every valid one of them.
    index: dict[str, int] = {}
    for colonne in ("ean", "ref_fournisseur"):
        if colonne not in tarif.columns:
            continue
        for position, valeur in tarif[colonne].items():
            chiffres = re.sub(r"\D", "", str(valeur or ""))
            if ean_est_valide(chiffres):
                index.setdefault(chiffres, position)
    if not index:
        return fusion

    sources = []
    for colonne in ("ref_fournisseur_wms", "ref_fabricant"):
        if colonne in fusion.columns:
            sources.append(colonne)
    if not sources:
        return fusion

    trouves = 0
    for i in fusion.index[absents]:
        for colonne in sources:
            gagne = False
            for code, preuve in candidats_identifiant(fusion.at[i, colonne]):
                if code in index:
                    ligne = index[code]
                    for champ in champs:
                        if champ in tarif.columns:
                            fusion.at[i, champ] = tarif.at[ligne, champ]
                    fusion.at[i, "Clé de rapprochement"] = "EAN lu"
                    fusion.at[i, "preuve_identifiant"] = f"{colonne} : {preuve} -> {code}"
                    trouves += 1
                    gagne = True
                    break
            if gagne:
                break
    if trouves:
        print(f"  {trouves} matched by EAN as read")
    return fusion


# --------------------------------------------------------------------------
# Diagnostic


def diagnostic(motif: str | None = None) -> None:
    """What identifier repair would find, applying nothing."""
    from extracteurs.base import chemin_lisible

    chemin = RACINE / "sortie" / "3-achats" / "pa_etat_par_article.xlsx"
    if not chemin.exists():
        print("pa_etat_par_article.xlsx missing: run run_pa.py first")
        return
    d = pd.read_excel(chemin_lisible(chemin), sheet_name=0, header=3, dtype=str)
    if motif:
        d = d[d["Fournisseur"].fillna("").str.upper().str.contains(motif.upper())]

    cible = d[d["Motif"].fillna("").str.contains("pas dans son tarif", case=False)]
    print(f"{len(cible)} articles in their supplier's price list but unmatched")
    print()

    compteurs: dict[str, int] = {}
    exemples: dict[str, list[str]] = {}
    multi = 0
    for _, r in cible.iterrows():
        for colonne in ("Réf. article fournisseur", "Référence fabricant"):
            if colonne not in cible.columns:
                continue
            brut = r.get(colonne)
            if len(morceaux(brut)) > 1:
                multi += 1
            for code, preuve in candidats_identifiant(brut):
                if preuve.endswith("EAN lu tel quel"):
                    continue
                cle = preuve.split(" — ")[-1]
                compteurs[cle] = compteurs.get(cle, 0) + 1
                exemples.setdefault(cle, []).append(f"{brut} -> {code}")

    print("=== possible reconstructions ===")
    if compteurs:
        for cle, n in sorted(compteurs.items(), key=lambda x: -x[1]):
            print(f"  {n:>4}  {cle}")
            for e in exemples[cle][:4]:
                print(f"          {e}")
    else:
        print("  none")
    print(f"\n  cells carrying multiple identifiers: {multi}")

    # The stem: we count it, we do not apply it.
    socles = cible["Réf. article fournisseur"].map(socle_reference).dropna()
    print(f"\n=== usable reference stems: {len(socles)} ===")
    print("  (reported, never applied - corroboration by the price required)")


def main() -> None:
    diagnostic(sys.argv[1] if len(sys.argv) > 1 else None)


if __name__ == "__main__":
    main()
