# -*- coding: utf-8 -*-
"""
Working scope: which WMS items are worth documenting.

The database holds 23,698 items, but a good share of them are not
ordered like a catalogue product: spare parts, service packages, the
rental fleet, discontinued products. Documenting them would serve no
purpose.

The filter is applied to the "Type" column of the WMS export, the only
one that tells these kinds of item apart reliably.
"""

from __future__ import annotations

import unicodedata

import pandas as pd

# Types excluded from the workstream, with the reason for exclusion.
TYPES_EXCLUS = {
    "SE": "écoulement de stock",
    "LO": "location de dispositif médical",
    "79": "produit arrêté fabricant",
    "19": "parc matériel",
    "FO": "forfait de prestation",
    "SV": "pièces détachées SAV",
}

# Families with no meaning for a product catalogue: these are not items
# that get sold, they are administrative rows.
FAMILLES_EXCLUES = {
    "N903": "articles de gestion administrative",
}

# --- wheelchairs under the old nomenclature --------------------------------
# The move to the new nomenclature (December 2025) created labels
# prefixed with "VPH". The old "FR ..." and "FAUT ..." labels duplicate
# them and must no longer be documented.
#
# The rule applies ONLY TO wheelchair families: a label starting with
# "FAUT" in the "Bath and shower" family denotes a bath chair, which
# stays inside the scope.
PREFIXES_ANCIENNE_NOMENCLATURE = ("FR ", "FAUT")
PREFIXE_CONSERVE = "VPH"
FAMILLES_FAUTEUILS = ("FAUTEUIL ROULANT",)


def code_type(valeur) -> str:
    """Short code of a WMS type: "CP - CATALOGUE PROFESSIONNEL" -> "CP"."""
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return ""
    return str(valeur).split("-")[0].strip().upper()


def _sans_accent(valeur) -> str:
    """Uppercase, accent-free, to compare labels and families."""
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return ""
    texte = unicodedata.normalize("NFKD", str(valeur))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return texte.upper().strip()


def est_fauteuil_ancienne_nomenclature(libelle, famille_libelle) -> bool:
    """Wheelchair still carrying an old-nomenclature label."""
    famille = _sans_accent(famille_libelle)
    if not any(cible in famille for cible in FAMILLES_FAUTEUILS):
        return False

    texte = _sans_accent(libelle)
    if texte.startswith(PREFIXE_CONSERVE):
        return False
    return texte.startswith(PREFIXES_ANCIENNE_NOMENCLATURE)


def motif_hors_perimetre(ligne) -> str | None:
    """Reason to exclude an item, or None if it is inside the scope."""
    code = code_type(ligne.get("Type") or ligne.get("type_article"))
    if code in TYPES_EXCLUS:
        return f"{code} — {TYPES_EXCLUS[code]}"

    famille = ligne.get("Code famille") or ligne.get("code_famille")
    if famille and str(famille).strip().upper() in FAMILLES_EXCLUES:
        cle = str(famille).strip().upper()
        return f"{cle} — {FAMILLES_EXCLUES[cle]}"

    libelle = (ligne.get("Libellé déclinaison ^(1)")
               or ligne.get("designation_wms"))
    famille_libelle = ligne.get("Libellé famille") or ligne.get("famille")
    if est_fauteuil_ancienne_nomenclature(libelle, famille_libelle):
        return "fauteuil roulant d'ancienne nomenclature (pas de VPH devant)"

    return None


def annoter(df: pd.DataFrame, colonne_type: str = "Type") -> pd.DataFrame:
    """The COMPLETE master data, plus the exclusion reason.

    This is the form to use upstream of a ranking: the starting
    population is the whole WMS master data, and whatever used to
    exclude a row becomes a column that qualifies it. No row
    disappears.

    A filter applied upstream shows up in the total and nowhere else:
    there is no longer any way to know what was removed, or why, or how
    much. A column can be counted, broken down and challenged.

    Two columns added:
      - `motif_hors_perimetre` : the reason in plain words, "" if the
                                 item stays
      - `au_perimetre`         : boolean, to filter DOWNSTREAM if needed

    The reason comes from `motif_hors_perimetre()`, the same function
    used everywhere else: one single definition of the scope, not two
    that would quietly drift apart.
    """
    travail = df.copy()
    if colonne_type != "Type" and colonne_type in travail.columns:
        # `motif_hors_perimetre` reads "Type" or "type_article". A source
        # that names its column differently has to say so here, otherwise
        # the type is never read and the filter fails without raising.
        travail["Type"] = travail[colonne_type]

    motifs = travail.apply(motif_hors_perimetre, axis=1)
    resultat = df.copy()
    resultat["motif_hors_perimetre"] = motifs.fillna("")
    resultat["au_perimetre"] = motifs.isna()
    return resultat


def filtrer(df: pd.DataFrame, colonne_type: str = "Type") -> pd.DataFrame:
    """Subset of the items to document.

    WARNING: this function DROPS rows. It suits the EAN workstream,
    which knowingly works on a subset. For a ranking of the master
    data, use `annoter()`: the starting population must stay whole and
    the exclusion must become a column.

    `colonne_type` must point at a column carrying the SHORT CODE of the
    type. On the stock export, passing the label filters almost nothing,
    without raising the slightest error.
    """
    codes = df[colonne_type].map(code_type)
    garde = ~codes.isin(TYPES_EXCLUS)

    colonne_famille = next(
        (c for c in ("Code famille", "code_famille") if c in df.columns), None
    )
    if colonne_famille:
        familles = df[colonne_famille].fillna("").str.strip().str.upper()
        garde &= ~familles.isin(FAMILLES_EXCLUES)

    # Wheelchairs under the old nomenclature
    colonne_libelle = next(
        (c for c in ("Libellé déclinaison ^(1)", "Libellé article ^(1)",
                     "designation_wms") if c in df.columns), None
    )
    colonne_famille_libelle = next(
        (c for c in ("Libellé famille", "famille") if c in df.columns), None
    )
    if colonne_libelle and colonne_famille_libelle:
        anciens = df.apply(
            lambda r: est_fauteuil_ancienne_nomenclature(
                r[colonne_libelle], r[colonne_famille_libelle]
            ),
            axis=1,
        )
        garde &= ~anciens

    return df[garde]


def resumer(df: pd.DataFrame, colonne_type: str = "Type") -> pd.DataFrame:
    """Count of items kept and excluded, by type."""
    codes = df[colonne_type].map(code_type)
    table = df.assign(_code=codes).groupby("_code").size().rename("articles")
    table = table.reset_index().rename(columns={"_code": "type"})
    table["statut"] = table["type"].map(
        lambda c: f"écarté ({TYPES_EXCLUS[c]})" if c in TYPES_EXCLUS
        else "retenu"
    )
    return table.sort_values("articles", ascending=False)
