# -*- coding: utf-8 -*-
"""
LE fichier des codes EAN : toutes les sources reunies, une ligne par article.

    python ean_global.py

Sortie : sortie/2-chantier-ean/EAN distributeur.xlsx

Jusqu'ici les codes etaient eparpilles entre la consolidation, les releves
VIDAL et le scanner d'entrepot, chacun dans son fichier. Celui-ci les
rassemble et tranche : pour chaque article, le meilleur code disponible et
d'ou il vient.

Ordre de preference des sources, du plus sur au moins sur :

    1. entrepot        releve sur l'emballage, au scanner : c'est le produit
                       lui-meme qui l'a donne
    2. tarif           le fournisseur le publie dans son fichier
    3. référence=EAN   la reference saisie dans le WMS EST un code-barres
    4. eudamed         declaration UDI du fabricant, par reference
    5. vidal           rapproche par libelle sur le site VIDAL
    6. eudamed-modèle  rapproche par nom de modele : le moins sur

Un code releve a l'entrepot l'emporte donc sur tout le reste, y compris
sur le tarif : si les deux different, c'est le carton qui a raison.

Trois onglets :
    "Codes EAN"    tous les articles du perimetre, code et provenance
    "À compléter"  ce qui manque, en commencant par les articles actifs
    "Synthèse"     le compte par source et par fournisseur
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

# Du plus sur au moins sur : l'ordre decide en cas de conflit
PRIORITES = ["entrepôt", "fournisseur", "tarif", "référence=EAN", "eudamed",
             "vidal", "libellé", "eudamed-modèle"]

COLONNES = [
    "Code article", "Référence", "Désignation", "Fournisseur",
    # « Origine référence » dit laquelle des deux colonnes du WMS a
    # fourni la référence : celle du fournisseur, ou, à défaut, celle du
    # fabricant. Un rapprochement fait sur la seconde reste ainsi
    # identifiable.
    "Réf. fournisseur", "Origine référence",
    "Code EAN", "Source", "Fiabilité", "Concordance",
    "Détail source",
    "Actif", "Motif activité", "Stock", "Emplacement", "Type", "Famille",
]

# Trois niveaux plutot que deux : le rapprochement par libelle n'est pas
# une certitude, mais il n'est pas douteux pour autant — il tombe juste
# dans environ quatre cas sur cinq. Le confondre avec l'incertain ferait
# rejeter en bloc plus de mille codes exploitables.
FIABLES = {"entrepôt", "fournisseur", "tarif", "référence=EAN", "eudamed"}
PROBABLES = {"libellé", "vidal"}

# Au-dessus de ce score, un rapprochement par libelle est tenu pour
# probable ; en dessous, il demande une relecture.
SCORE_PROBABLE = 0.8


def _fiabilite(source, score) -> str:
    """Niveau de confiance d'un code, selon sa source et son score."""
    if not source:
        return ""
    if source in FIABLES:
        return "sûr"
    if source in PROBABLES:
        # Un score connu affine le jugement ; sans score (VIDAL, dont le
        # rapprochement est deja filtre), on s'en tient a "probable".
        if score is None or pd.isna(score):
            return "probable"
        return "probable" if float(score) >= SCORE_PROBABLE else "à vérifier"
    return "à vérifier"


def _consolidation() -> dict:
    """Codes issus des tarifs et d'EUDAMED, avec leur provenance."""
    fichier = mise_en_forme.derniere_version(
        SORTIE_CHANTIER, "catalogue_wms_complete.xlsx")
    if fichier is None:
        return {}
    df = pd.read_excel(chemin_lisible(fichier), sheet_name="Articles",
                       skiprows=3, dtype=str)
    df = df[df["code_article"].notna() & df["ean"].notna()]

    # La consolidation nomme ses sources autrement : on les ramene au
    # vocabulaire commun a tout le fichier.
    equivalences = {
        "tarif": "tarif",
        "référence = EAN": "référence=EAN",
        "eudamed-référence": "eudamed",
        "eudamed-modèle": "eudamed-modèle",
        # Rapprochement par le libelle : le produit est le bon selon son
        # nom, mais la reference n'a pas confirme. A relire.
        "libellé": "libellé",
    }
    resultat = {}
    for _, ligne in df.iterrows():
        source = equivalences.get(str(ligne.get("source_ean")), "tarif")
        # Pour un code venu d'un tarif, le fichier exact permet de
        # remonter a la ligne d'origine en cas de doute ; pour EUDAMED,
        # c'est le nom sous lequel le fabricant a declare le produit.
        if source == "tarif":
            detail = ligne.get("fichier_source")
        elif source == "libellé":
            # Le libelle du fournisseur : c'est lui qu'on relira pour
            # confirmer que le produit est bien le notre.
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
    """Codes releves sur le site VIDAL."""
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
        # Le libelle VIDAL rapproche : c'est lui qu'on relira pour
        # verifier que le produit est bien le notre.
        resultat[ligne["code_article"]] = (
            ligne["ean"], "vidal", str(ligne.get("libelle_vidal") or "")[:80],
            ligne.get("concordance"))
    print(f"  VIDAL         : {len(resultat)} codes")
    return resultat


def _fournisseurs() -> dict:
    """Codes communiques par les fournisseurs, en reponse a nos demandes.

    C'est la meilleure source apres un code lu sur l'emballage : le
    fabricant a repondu sur LA ligne qu'on lui a soumise. Ni
    rapprochement, ni deduction, ni homonymie possible — les trois
    facons dont les autres sources se trompent.

    Le fichier est produit par `reponses_ean.py`, qui a deja ecarte les
    cles GS1 fausses et les codes de carton repondus a la place de
    l'unite de vente.
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
    print(f"  fournisseurs  : {len(resultat)} codes")
    return resultat


def _eudamed_reference() -> dict:
    """Codes releves sur EUDAMED, interroges reference par reference.

    Cette collecte tourne en tache de fond pendant plusieurs jours et
    ecrit son propre classeur au fil de l'eau. Elle n'entre pas dans la
    consolidation des catalogues — elle n'a rien a voir avec un tarif —
    et doit donc etre lue ici, sans quoi ses trouvailles restent dans un
    fichier que personne ne rapproche.

    C'est un rapprochement par correspondance EXACTE sur la reference
    fabricant : la source la plus sure apres un code lu sur l'emballage.
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
    print(f"  EUDAMED (réf) : {len(resultat)} codes")
    return resultat


def _entrepot() -> dict:
    """Codes releves au scanner, sur l'emballage.

    Le dernier releve d'un article l'emporte : si on a scanne deux fois,
    c'est que la premiere lecture etait a corriger.
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
    print(f"  entrepôt      : {len(resultat)} codes")
    return resultat


# Ce qui ne peut pas etre une reference fabricant, verifie sur les 2 947
# valeurs candidates de l'extract du 28/08.
REFERENCE_DOUTEUSE = re.compile(
    # Excel a converti « 8001-08-01 » en date : la corruption est visible
    r"^\d{4}-\d{2}-\d{2}"
    # Mentions libres saisies a la place d'une reference
    r"|\b(?:SELON|VOIR|DIVERS|AUCUN|SANS|NEANT|N/?A|NC)\b",
    re.IGNORECASE)


def _reference_fabricant_utilisable(valeur) -> str | None:
    """La reference fabricant, quand elle en est vraiment une.

    Elle sert de repli quand la reference fournisseur manque — et elle
    manque sur 3 543 articles, soit la moitie de ce qui reste sans code.
    C'est meme la meilleure clef pour EUDAMED, qui indexe les
    declarations du FABRICANT et non du distributeur.

    Trois formes sont ecartees : les dates nees d'une conversion Excel,
    les mentions libres du type « SELON TAILLE », et tout ce qui tient
    en moins de quatre caracteres — « 81 », « AT » ne designent rien et
    feraient perdre un appel chacun.
    """
    texte = nettoyer_texte(valeur)
    if not texte or len(texte) < 4:
        return None
    if REFERENCE_DOUTEUSE.search(texte):
        return None
    return texte


def _reference(article) -> tuple[str | None, str]:
    """(reference a interroger, d'ou elle vient) pour un article du WMS."""
    fournisseur = nettoyer_texte(article.get("Ref. art. four."))
    if fournisseur:
        return fournisseur, "fournisseur"
    fabricant = _reference_fabricant_utilisable(
        article.get("Référence fabricant"))
    if fabricant:
        return fabricant, "fabricant"
    return None, ""


def _articles() -> pd.DataFrame:
    """Le perimetre, avec stock, activite et emplacement."""
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
    print("Sources :")
    sources = {}
    # Du moins sur au plus sur : les meilleurs ecrasent les autres
    for lecture in (_vidal, _consolidation, _eudamed_reference,
                    _fournisseurs, _entrepot):
        for code, valeur in lecture().items():
            precedent = sources.get(code)
            if precedent is None or (
                PRIORITES.index(valeur[1]) < PRIORITES.index(precedent[1])
            ):
                sources[code] = valeur

    # eudamed-modèle est le seul rapprochement dont on sait qu'il se
    # trompe en masse : on le garde, signale, jamais prioritaire.
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
            # La reference fournisseur d'abord — c'est elle qui rapproche
            # les tarifs. A defaut, celle du fabricant, qui existe sur
            # 2 947 articles la ou l'autre manque.
            #
            # `nettoyer_texte` et non un simple `or` : une cellule vide
            # lue par pandas vaut NaN, et NaN est VRAI en Python. Le
            # repli ne se serait jamais declenche.
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

    print(f"\n{len(tout)} articles | {len(avec)} avec EAN "
          f"({100 * len(avec) / len(tout):.0f} %)")
    print(f"  dont actifs : {actifs_ok} / {len(actifs)}")
    for niveau in ("sûr", "probable", "à vérifier"):
        nombre = int((avec["Fiabilité"] == niveau).sum())
        if nombre:
            print(f"  {niveau:<12} {nombre}")
    print(f"\n{synthese.to_string(index=False)}")
    print(f"\n  -> {chemin}")


if __name__ == "__main__":
    main()
