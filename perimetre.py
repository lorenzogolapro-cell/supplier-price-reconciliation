# -*- coding: utf-8 -*-
"""
Perimetre de travail : quels articles du WMS meritent d'etre documentes.

La base compte 23 698 articles, mais une bonne part ne se commande pas
comme un produit de catalogue : pieces detachees, forfaits de prestation,
parc de location, produits arretes. Les documenter n'aurait aucun usage.

Le filtre porte sur la colonne "Type" de l'export du WMS, la seule qui
distingue ces natures d'article de facon fiable.
"""

from __future__ import annotations

import unicodedata

import pandas as pd

# Types ecartes du chantier, avec la raison de leur exclusion.
TYPES_EXCLUS = {
    "SE": "écoulement de stock",
    "LO": "location de dispositif médical",
    "79": "produit arrêté fabricant",
    "19": "parc matériel",
    "FO": "forfait de prestation",
    "SV": "pièces détachées SAV",
}

# Familles sans objet pour un catalogue produit : ce ne sont pas des
# articles vendus mais des lignes de gestion.
FAMILLES_EXCLUES = {
    "N903": "articles de gestion administrative",
}

# --- fauteuils roulants d'ancienne nomenclature ----------------------------
# Le passage a la nouvelle nomenclature (decembre 2025) a fait creer des
# libelles prefixes "VPH". Les anciens libelles "FR ..." et "FAUT ..." font
# double emploi avec eux et ne doivent plus etre documentes.
#
# La regle ne s'applique QU'AUX familles de fauteuils roulants : un libelle
# commencant par "FAUT" dans la famille "Bain et douche" designe un fauteuil
# de bain, qui reste au perimetre.
PREFIXES_ANCIENNE_NOMENCLATURE = ("FR ", "FAUT")
PREFIXE_CONSERVE = "VPH"
FAMILLES_FAUTEUILS = ("FAUTEUIL ROULANT",)


def code_type(valeur) -> str:
    """Code court d'un type WMS : "CP - CATALOGUE PROFESSIONNEL" -> "CP"."""
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return ""
    return str(valeur).split("-")[0].strip().upper()


def _sans_accent(valeur) -> str:
    """Majuscules sans accent, pour comparer libelles et familles."""
    if valeur is None or (isinstance(valeur, float) and valeur != valeur):
        return ""
    texte = unicodedata.normalize("NFKD", str(valeur))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return texte.upper().strip()


def est_fauteuil_ancienne_nomenclature(libelle, famille_libelle) -> bool:
    """Fauteuil roulant portant encore un libelle d'ancienne nomenclature."""
    famille = _sans_accent(famille_libelle)
    if not any(cible in famille for cible in FAMILLES_FAUTEUILS):
        return False

    texte = _sans_accent(libelle)
    if texte.startswith(PREFIXE_CONSERVE):
        return False
    return texte.startswith(PREFIXES_ANCIENNE_NOMENCLATURE)


def motif_hors_perimetre(ligne) -> str | None:
    """Raison d'ecarter un article, ou None s'il est dans le perimetre."""
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
    """Le referentiel COMPLET, augmente du motif d'exclusion.

    C'est la forme a utiliser en amont d'un classement : la population de
    depart est le referentiel WMS entier, et ce qui ecartait une ligne
    devient une colonne qui la qualifie. Aucune ligne ne disparait.

    Un filtre applique en amont se voit dans le total et nulle part
    ailleurs : on ne sait plus ce qui a ete retire, ni pourquoi, ni
    combien. Une colonne se compte, se ventile et se conteste.

    Deux colonnes ajoutees :
      - `motif_hors_perimetre` : la raison en clair, "" si l'article reste
      - `au_perimetre`          : booleen, pour filtrer EN AVAL si besoin

    Le motif vient de `motif_hors_perimetre()`, la meme fonction que celle
    utilisee ailleurs : une seule definition du perimetre, pas deux qui
    divergeraient en silence.
    """
    travail = df.copy()
    if colonne_type != "Type" and colonne_type in travail.columns:
        # `motif_hors_perimetre` lit "Type" ou "type_article". Une source
        # qui nomme sa colonne autrement doit le dire ici, sinon le type
        # n'est jamais lu et le filtre echoue sans lever d'erreur.
        travail["Type"] = travail[colonne_type]

    motifs = travail.apply(motif_hors_perimetre, axis=1)
    resultat = df.copy()
    resultat["motif_hors_perimetre"] = motifs.fillna("")
    resultat["au_perimetre"] = motifs.isna()
    return resultat


def filtrer(df: pd.DataFrame, colonne_type: str = "Type") -> pd.DataFrame:
    """Sous-ensemble des articles a documenter.

    ATTENTION : cette fonction SUPPRIME des lignes. Elle convient au
    chantier EAN, qui travaille sciemment sur un sous-ensemble. Pour un
    classement du referentiel, utiliser `annoter()` : la population de
    depart doit rester entiere et l'exclusion devenir une colonne.

    `colonne_type` doit designer une colonne portant le CODE COURT du type.
    Sur l'export stock, passer le libelle ne filtre presque rien, sans
    lever la moindre erreur.
    """
    codes = df[colonne_type].map(code_type)
    garde = ~codes.isin(TYPES_EXCLUS)

    colonne_famille = next(
        (c for c in ("Code famille", "code_famille") if c in df.columns), None
    )
    if colonne_famille:
        familles = df[colonne_famille].fillna("").str.strip().str.upper()
        garde &= ~familles.isin(FAMILLES_EXCLUES)

    # Fauteuils roulants d'ancienne nomenclature
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
    """Compte des articles retenus et ecartes, par type."""
    codes = df[colonne_type].map(code_type)
    table = df.assign(_code=codes).groupby("_code").size().rename("articles")
    table = table.reset_index().rename(columns={"_code": "type"})
    table["statut"] = table["type"].map(
        lambda c: f"écarté ({TYPES_EXCLUS[c]})" if c in TYPES_EXCLUS
        else "retenu"
    )
    return table.sort_values("articles", ascending=False)
