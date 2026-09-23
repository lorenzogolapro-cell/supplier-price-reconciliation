# -*- coding: utf-8 -*-
"""
THE EAN code file: every source merged, one row per article.

    python ean_global.py

Output: sortie/2-chantier-ean/EAN distributeur.xlsx

Until now the codes were scattered between the consolidation, the VIDAL
readings and the warehouse scanner, each in its own file. This one
gathers them and decides: for each article, the best code available and
where it comes from.

Order of preference of the sources, from the safest to the least safe:

    1. entrepôt        read off the packaging with a scanner: the product
                       itself gave it
    2. tarif           the supplier publishes it in its own file
    3. référence=EAN   the reference keyed into the WMS IS a barcode
    4. eudamed         the manufacturer's UDI declaration, by reference
    5. vidal           matched by label on the VIDAL website
    6. eudamed-modèle  matched by model name: the least safe

A code read at the warehouse therefore beats everything else, including
the price list: if the two differ, the box is right.

Three sheets:
    "Codes EAN"    every article in scope, code and provenance
    "À compléter"  what is missing, starting with the active articles
    "Synthèse"     the count by source and by supplier
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import pandas as pd

import activite  # noqa: E402
import mise_en_forme  # noqa: E402
import perimetre  # noqa: E402
from extracteurs.base import chemin_lisible, ean_est_valide, nettoyer_texte  # noqa: E402
from main import SORTIE_CHANTIER, WMS_EXTRACTS  # noqa: E402

EXTRACT = "extract article 28_08.xlsx"
ANNEE = 2026

# From the safest to the least safe: the order decides conflicts
PRIORITES = ["entrepôt", "fournisseur", "tarif", "référence=EAN", "eudamed",
             "vidal", "libellé", "eudamed-modèle"]

COLONNES = [
    "Code article", "Référence", "Désignation", "Fournisseur",
    # "Origine référence" says which of the two WMS columns supplied the
    # reference: the supplier's, or, failing that, the manufacturer's. A
    # match made on the second one therefore stays identifiable.
    "Réf. fournisseur", "Origine référence",
    "Code EAN", "Source", "Fiabilité", "Concordance",
    "Détail source",
    "Actif", "Motif activité", "Stock", "Emplacement", "Type", "Famille",
]

# Three levels rather than two: a label match is not a certainty, but it
# is not doubtful for all that: it is right in roughly four cases out of
# five. Confusing it with the uncertain would reject more than a thousand
# usable codes outright.
FIABLES = {"entrepôt", "fournisseur", "tarif", "référence=EAN", "eudamed"}
PROBABLES = {"libellé", "vidal"}

# Above this score, a label match is deemed probable; below it, it calls
# for a second reading.
SCORE_PROBABLE = 0.8


def _fiabilite(source, score) -> str:
    """Confidence level of a code, from its source and its score."""
    if not source:
        return ""
    if source in FIABLES:
        return "sûr"
    if source in PROBABLES:
        # A known score sharpens the judgement; with no score (VIDAL,
        # whose matching is already filtered), we stay at "probable".
        if score is None or pd.isna(score):
            return "probable"
        return "probable" if float(score) >= SCORE_PROBABLE else "à vérifier"
    return "à vérifier"


def _consolidation() -> dict:
    """Codes from the price lists and from EUDAMED, with their provenance."""
    fichier = mise_en_forme.derniere_version(
        SORTIE_CHANTIER, "catalogue_wms_complete.xlsx")
    if fichier is None:
        return {}
    df = pd.read_excel(chemin_lisible(fichier), sheet_name="Articles",
                       skiprows=3, dtype=str)
    df = df[df["code_article"].notna() & df["ean"].notna()]

    # The consolidation names its sources differently: we bring them back
    # to the vocabulary shared by the whole file.
    equivalences = {
        "tarif": "tarif",
        "référence = EAN": "référence=EAN",
        "eudamed-référence": "eudamed",
        "eudamed-modèle": "eudamed-modèle",
        # Match on the label: the product is the right one by its name,
        # but the reference did not confirm it. To be reviewed.
        "libellé": "libellé",
    }
    resultat = {}
    for _, ligne in df.iterrows():
        source = equivalences.get(str(ligne.get("source_ean")), "tarif")
        # For a code from a price list, the exact file lets us trace back
        # to the original row in case of doubt; for EUDAMED, it is the
        # name under which the manufacturer declared the product.
        if source == "tarif":
            detail = ligne.get("fichier_source")
        elif source == "libellé":
            # The supplier's label: it is the one we will re-read to
            # confirm the product really is ours.
            detail = ligne.get("libelle_fournisseur")
        else:
            detail = ligne.get("eudamed_nom")
        score = ligne.get("concordance_libelle")
        resultat[ligne["code_article"]] = (
            ligne["ean"], source, detail if pd.notna(detail) else "",
            score if pd.notna(score) else None)
    print(f"  consolidation : {len(resultat)} codes")
    return resultat


def _vidal() -> dict:
    """Codes read from the VIDAL website."""
    fichier = mise_en_forme.derniere_version(SORTIE_CHANTIER,
                                             "ean_vidal.xlsx")
    if fichier is None:
        return {}
    try:
        df = pd.read_excel(chemin_lisible(fichier),
                           sheet_name="Rapprochements", skiprows=3,
                           dtype=str)
    except Exception:
        return {}
    df = df[df["code_article"].notna() & df["ean"].notna()]
    resultat = {}
    for _, ligne in df.iterrows():
        # The matched VIDAL label: it is the one we will re-read to check
        # that the product really is ours.
        resultat[ligne["code_article"]] = (
            ligne["ean"], "vidal", str(ligne.get("libelle_vidal") or "")[:80],
            ligne.get("concordance"))
    print(f"  VIDAL         : {len(resultat)} codes")
    return resultat


def _fournisseurs() -> dict:
    """Codes sent by the suppliers, in answer to our requests.

    This is the best source after a code read off the packaging: the
    manufacturer answered about THE line we submitted to it. No matching,
    no deduction, no possible homonym, the three ways the other sources
    get it wrong.

    The file is produced by `reponses_ean.py`, which has already
    discarded wrong GS1 check digits and case codes answered in place of
    the sales unit.
    """
    fichier = mise_en_forme.derniere_version(SORTIE_CHANTIER,
                                             "ean_fournisseurs.xlsx")
    if fichier is None:
        return {}
    try:
        df = pd.read_excel(chemin_lisible(fichier), sheet_name="Codes EAN",
                           skiprows=3, dtype=str)
    except Exception:
        return {}
    df = df[(df["Verdict"] == "retenu") & df["Code article"].notna()]
    resultat = {}
    for _, ligne in df.iterrows():
        resultat[ligne["Code article"]] = (
            ligne["Code EAN"], "fournisseur",
            f"communiqué par {ligne['Fournisseur']} "
            f"le {str(ligne.get('Reçu le'))[:10]}",
            None)
    print(f"  suppliers     : {len(resultat)} codes")
    return resultat


def _eudamed_reference() -> dict:
    """Codes read from EUDAMED, queried reference by reference.

    This collection runs in the background for several days and writes
    its own workbook as it goes. It is not part of the catalogue
    consolidation, it has nothing to do with a price list, and so it has
    to be read here; otherwise its findings stay in a file nobody joins.

    It is a match on an EXACT correspondence of the manufacturer
    reference: the safest source after a code read off the packaging.
    """
    fichier = mise_en_forme.derniere_version(SORTIE_CHANTIER,
                                             "ean_eudamed_reference.xlsx")
    if fichier is None:
        return {}
    try:
        df = pd.read_excel(chemin_lisible(fichier), sheet_name="Codes EAN",
                           skiprows=3, dtype=str)
    except Exception:
        return {}
    df = df[df["Code article"].notna() & df["Code EAN"].notna()]
    resultat = {}
    for _, ligne in df.iterrows():
        resultat[ligne["Code article"]] = (
            ligne["Code EAN"], "eudamed",
            f"référence {ligne.get('Référence fournisseur')} "
            f"— relevé le {str(ligne.get('Relevé le'))[:10]}",
            None)
    print(f"  EUDAMED (ref) : {len(resultat)} codes")
    return resultat


def _entrepot() -> dict:
    """Codes read with a scanner, off the packaging.

    The most recent reading of an article wins: if we scanned twice, it
    is because the first reading needed correcting.
    """
    fichier = RACINE / "data" / "collectes.csv"
    if not fichier.exists():
        return {}
    resultat = {}
    with fichier.open(encoding="utf-8-sig") as flux:
        for ligne in csv.DictReader(flux, delimiter=";"):
            code = (ligne.get("code_article") or "").strip()
            ean = (ligne.get("ean") or "").strip()
            if code and ean:
                nature = (ligne.get("nature") or "").strip()
                horodatage = (ligne.get("horodatage") or "")[:10]
                resultat[code] = (
                    ean, "entrepôt",
                    f"scanné le {horodatage}"
                    + (f" — {nature}" if nature else ""), None)
    print(f"  warehouse     : {len(resultat)} codes")
    return resultat


# What cannot be a manufacturer reference, checked against the 2 947
# candidate values of the 28/08 extract.
REFERENCE_DOUTEUSE = re.compile(
    # Excel converted "8001-08-01" into a date: the corruption shows
    r"^\d{4}-\d{2}-\d{2}"
    # Free text typed in instead of a reference
    r"|\b(?:SELON|VOIR|DIVERS|AUCUN|SANS|NEANT|N/?A|NC)\b",
    re.IGNORECASE)


def _reference_fabricant_utilisable(valeur) -> str | None:
    """The manufacturer reference, when it really is one.

    It serves as a fallback when the supplier reference is missing, and
    that one is missing on 3 543 articles, half of what is left without a
    code. It is even the best key for EUDAMED, which indexes the
    MANUFACTURER's declarations and not the distributor's.

    Three shapes are discarded: dates born of an Excel conversion, free
    text such as "SELON TAILLE", and anything shorter than four
    characters: "81" or "AT" designate nothing and would waste one call
    each.
    """
    texte = nettoyer_texte(valeur)
    if not texte or len(texte) < 4:
        return None
    if REFERENCE_DOUTEUSE.search(texte):
        return None
    return texte


def _reference(article) -> tuple[str | None, str]:
    """(reference to query, where it comes from) for a WMS article."""
    fournisseur = nettoyer_texte(article.get("Ref. art. four."))
    if fournisseur:
        return fournisseur, "fournisseur"
    fabricant = _reference_fabricant_utilisable(
        article.get("Référence fabricant"))
    if fabricant:
        return fabricant, "fabricant"
    return None, ""


def _articles() -> pd.DataFrame:
    """The scope, with stock, activity and location."""
    df = pd.read_excel(chemin_lisible(WMS_EXTRACTS / EXTRACT), dtype=str)
    df = perimetre.filtrer(df)

    fichiers = sorted(WMS_EXTRACTS.glob("stock_par-article*.xlsx"))
    if fichiers:
        stock = pd.read_excel(chemin_lisible(fichiers[-1]), dtype=str)
        if {"Code article", "Stock"} <= set(stock.columns):
            index = stock.dropna(subset=["Code article"]).drop_duplicates(
                "Code article").set_index("Code article")["Stock"]
            df["Stock"] = df["Code article"].map(index)

    return activite.qualifier(df, annee=ANNEE)


def main() -> None:
    print("Sources:")
    sources = {}
    # From the least safe to the safest: the best overwrite the others
    for lecture in (_vidal, _consolidation, _eudamed_reference,
                    _fournisseurs, _entrepot):
        for code, valeur in lecture().items():
            precedent = sources.get(code)
            if precedent is None or (
                PRIORITES.index(valeur[1]) < PRIORITES.index(precedent[1])
            ):
                sources[code] = valeur

    # eudamed-modèle is the only match we know gets it wrong en masse: we
    # keep it, flagged, never given priority.
    articles = _articles()
    lignes = []
    for _, article in articles.iterrows():
        code = article.get("Code article")
        ean, source, detail, score = sources.get(code, (None, None, "", None))
        if ean and not ean_est_valide(str(ean)):
            ean, source, detail, score = None, None, "", None
        lignes.append({
            "Code article": code,
            "Référence": article.get("Référence"),
            "Désignation": article.get("Libellé déclinaison ^(1)"),
            "Fournisseur": article.get("Nom fournisseur"),
            # The supplier reference first: it is the one that joins the
            # price lists. Failing that, the manufacturer's, which exists
            # on 2 947 articles where the other one is missing.
            #
            # `nettoyer_texte` and not a plain `or`: an empty cell read by
            # pandas is NaN, and NaN is TRUE in Python. The fallback would
            # never have fired.
            "Réf. fournisseur": _reference(article)[0],
            "Origine référence": _reference(article)[1],
            "Code EAN": ean,
            "Source": source,
            "Fiabilité": _fiabilite(source, score),
            "Concordance": score,
            "Détail source": detail,
            "Actif": "OUI" if article.get("actif") else "",
            "Motif activité": article.get("motif_activite"),
            "Stock": article.get("Stock"),
            "Emplacement": article.get("emplacement"),
            "Type": article.get("Type"),
            "Famille": article.get("Libellé famille"),
        })

    tout = pd.DataFrame(lignes, columns=COLONNES)
    avec = tout[tout["Code EAN"].notna()]
    manque = tout[tout["Code EAN"].isna()].sort_values(
        ["Actif", "Fournisseur"], ascending=[False, True])

    synthese = avec.groupby(["Source", "Fiabilité"]).size().rename(
        "codes").reset_index().sort_values("codes", ascending=False)

    par_fournisseur = tout.groupby("Fournisseur").agg(
        articles=("Code article", "size"),
        avec_ean=("Code EAN", lambda s: int(s.notna().sum())),
        actifs=("Actif", lambda s: int((s == "OUI").sum())),
    ).reset_index()
    par_fournisseur["taux_pct"] = (
        100 * par_fournisseur["avec_ean"] / par_fournisseur["articles"]
    ).round(1)
    par_fournisseur = par_fournisseur.sort_values("articles", ascending=False)

    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_CHANTIER / "EAN distributeur.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        tout.to_excel(writer, sheet_name="Codes EAN", index=False, startrow=3)
        manque.to_excel(writer, sheet_name="À compléter", index=False)
        synthese.to_excel(writer, sheet_name="Synthèse", index=False)
        par_fournisseur.to_excel(writer, sheet_name="Synthèse", index=False,
                                 startrow=len(synthese) + 3)

    actifs = tout[tout["Actif"] == "OUI"]
    actifs_ok = int(actifs["Code EAN"].notna().sum())
    mise_en_forme.formater(
        chemin,
        options={"Codes EAN": {
            "ligne_entete": 4,
            "figer_colonne": 3,
            "titre": "Codes EAN du distributeur — toutes sources",
            "sous_titre": (
                f"{len(tout)} articles | {len(avec)} avec code EAN "
                f"({100 * len(avec) / len(tout):.0f} %) | "
                f"actifs : {actifs_ok}/{len(actifs)} "
                f"({100 * actifs_ok / max(len(actifs), 1):.0f} %)"
            ),
        }},
    )

    print(f"\n{len(tout)} articles | {len(avec)} with an EAN "
          f"({100 * len(avec) / len(tout):.0f} %)")
    print(f"  of which active: {actifs_ok} / {len(actifs)}")
    for niveau in ("sûr", "probable", "à vérifier"):
        nombre = int((avec["Fiabilité"] == niveau).sum())
        if nombre:
            print(f"  {niveau:<12} {nombre}")
    print(f"\n{synthese.to_string(index=False)}")
    print(f"\n  -> {chemin}")


if __name__ == "__main__":
    main()
