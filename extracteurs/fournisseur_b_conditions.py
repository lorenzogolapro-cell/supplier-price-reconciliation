# -*- coding: utf-8 -*-
"""
Extracteur des CONDITIONS du FOURNISSEUR B (Excel) — le tarif 2026.

POURQUOI CE FICHIER PLUTOT QUE LE PDF
    Le projet lisait « tarif_fournisseur.pdf » et n'en tirait que 52 lignes,
    laissant 130 articles du perimetre sans prix. Le meme tarif existe en
    Excel — « tarif_fournisseur.xlsx » — avec douze onglets produits et
    environ 195 references chiffrees.

    Il etait ecarte pour une raison sans rapport avec son contenu : son NOM
    porte « 2025 », donc l'inventaire lui donnait le millesime 2025, alors
    que sa premiere page s'intitule « CONDITIONS 2026 (applicable au 1er
    decembre 2025 sous nouvelle nomenclature) ». Le PDF, lui, avait ete
    retenu par choix manuel. C'est le piege de millesime documente dans les
    regles du projet, dans sa version la plus couteuse.

LA MISE EN PAGE, QUI EST « UN PEU LE BORDEL »
    Chaque onglet melange trois sortes de lignes :

      - une ligne de CATEGORIE de nomenclature, sans prix
        « FMP - Fauteuil non-modulaire a propulsion manuelle ou a pousser »
      - une ligne PRODUIT, avec son prix a la quantite 1
        « MODELE EXEMPLE - Accoudoirs relevables »  150,00  145,00  3*
      - une ou plusieurs lignes de PALIER, sans designation, qui se
        rattachent au produit du dessus
        (vide)                                             140,00  6*

    (libelles et montants ci-dessus : valeurs d'exemple)

    Un produit dont on lirait le palier au lieu du prix unitaire serait
    valorise trop bas — c'est l'incident du 08/09 chez le responsable des
    prix, a l'envers. On prend donc la colonne « Qte 1 », et les paliers
    partent en colonne `paliers`, pour memoire.

LE PIEGE DES POURCENTAGES
    Certaines lignes ne portent pas un prix mais un TAUX DE REMISE sur le
    prix public, qui est en derniere colonne. Le MODELE EXEMPLE n'a pas de
    prix : il a « 15 % » et un public de 600,00, donc 510,00.

    Trois ecritures coexistent, et c'est le FORMAT de la cellule qui les
    departage — pas leur valeur :

        150             format « 0 € »   -> un prix
        0,15            format « 0% »    -> un taux de 15 %
        « 10+5% »       texte            -> une remise COMPOSEE

    La remise composee est le piege veritable. « 10+5% » ne fait pas 15 %
    mais 14,5 % : 1 - (0,90 x 0,95). Le fichier le demontre lui-meme, en
    ecrivant ailleurs « 14,5% (10+5%) », « 19% (10+10%) », « 24% (20+5%) »
    — trois fois la meme regle. Une premiere version lisait « 10+1% » comme
    1 % et valorisait un fauteuil a 99 % de son prix public.

    Sans le format, 0,15 aurait pu passer pour quinze centimes.

    (tous les taux et montants de ce docstring : valeurs d'exemple)

PAS DE REFERENCE FOURNISSEUR
    Ces onglets n'en portent pas — comme le Fournisseur E. Le rapprochement
    se fera donc par le libelle, avec les precautions de pa_libelle.py.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from extracteurs.base import chemin_lisible, finaliser, nettoyer_texte

FOURNISSEUR = "FOURNISSEUR B"

# Onglets qui ne portent aucun produit.
HORS_PRODUITS = {
    "page de garde", "administratifs", "cgv", "sommaire", "feuil1",
}

# « 151.5 », « 151,50 » — un prix. On exige au plus deux decimales.
PRIX = re.compile(r"^\d{1,6}(?:[.,]\d{1,2})?$")
# « 14,5% (10+5%) », « 24% (20+5%) »  # valeurs d'exemple
POURCENT = re.compile(r"(\d{1,2}(?:[.,]\d+)?)\s*%")
# « 3* », « 9* », « 3 et+ », « 2 »
QUANTITE = re.compile(r"(\d{1,4})")

# En dessous de ce montant, une valeur n'est pas un prix de fauteuil roulant
# mais un taux que Excel a stocke en nombre (0,25 pour 25 %).
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
    """Le taux de remise qu'annonce un texte, en pourcentage.

    Trois formes, dans cet ordre de fiabilite :

        « 14,5% (10+5%) »   le taux est ecrit avant la parenthese : on le lit
        « 10+5% »           remise COMPOSEE : 1 - (0,90 x 0,95) = 14,5 %
        « 15% »             un taux simple

    (taux ci-dessus : valeurs d'exemple)
    """
    texte = brut.replace(",", ".")
    # Forme « X% (...) » : le total est deja calcule, on ne recalcule pas.
    avant = re.match(r"\s*(\d{1,2}(?:\.\d+)?)\s*%\s*\(", texte)
    if avant:
        return float(avant.group(1))

    # Forme composee « 10+5% », « 10+5+2% »  # valeurs d'exemple
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
    """(prix HT, comment il a ete obtenu) — ou (None, motif).

    `format_cellule` est le format Excel de la cellule. C'est LUI qui dit si
    0,15 vaut quinze centimes ou quinze pour cent : la valeur seule
    ne le dit pas, et deviner au seuil marchait par chance.
    """
    brut = _texte(valeur)
    if not brut:
        return None, ""

    format_cellule = format_cellule or ""
    est_pourcent = "%" in format_cellule
    valeur_num = _nombre(brut)

    # Un texte qui porte un « % » est un taux, quel que soit le format.
    if "%" in brut:
        taux = _taux_depuis_texte(brut)
        if taux is None:
            return None, "taux illisible"
        if public is None:
            return None, "taux sans prix public"
        return round(public * (1 - taux / 100), 4), f"public - {taux:g}%"

    # Une valeur numerique dans une cellule formatee en pourcentage EST un
    # taux : 0,15 sous le format « 0% » s'affiche « 15 % ».
    if est_pourcent and valeur_num is not None:
        if not 0 < valeur_num < 1:
            return None, "taux hors bornes"
        if public is None:
            return None, "taux sans prix public"
        return (round(public * (1 - valeur_num), 4),
                f"public - {valeur_num * 100:g}%")

    if valeur_num is None:
        return None, ""
    # Reste le cas ordinaire : un prix, dans une cellule formatee en euros.
    if valeur_num < PLANCHER_PRIX:
        return None, "valeur trop faible pour un prix"
    return round(valeur_num, 4), "tarif net unitaire HT"


def _colonnes(brut: pd.DataFrame) -> dict | None:
    """Repere la ligne d'en-tete et associe chaque role a un index.

    On lit les INTITULES plutot que des positions figees : les onglets ne
    partagent pas tous la meme largeur, et une colonne inseree decalerait
    silencieusement tous les prix.
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
        # La designation precede toujours la premiere colonne de prix.
        roles.setdefault("designation", max(roles["prix"] - 1, 0))
        roles.setdefault("palier", roles["prix"] + 1)
        return roles
    return None


def extract(path) -> pd.DataFrame:
    chemin = Path(path)
    lisible = chemin_lisible(chemin)
    classeur = pd.ExcelFile(lisible)
    # data_only=True donne la valeur en cache des formules, et surtout
    # `number_format`, sans lequel on ne peut pas distinguer un taux d'un prix.
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
                """Le format Excel de la cellule, ou '' si introuvable."""
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
                # Nouvelle ligne PRODUIT
                courant = {
                    "fournisseur": FOURNISSEUR,
                    # La categorie rejoint la designation : c'est elle qui
                    # porte FMP / FRM, que le WMS reprend dans ses libelles.
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
                # Ligne de PALIER rattachee au produit du dessus
                seuil = QUANTITE.search(quantite) if quantite else None
                if seuil:
                    courant["_paliers"].append(
                        f"{seuil.group(1)}:{prix_palier:.2f}")
            elif designation and prix is None:
                # Ligne de CATEGORIE de nomenclature : elle ne porte pas de
                # prix, elle ferme le produit precedent et s'applique a tous
                # les produits qui suivent.
                #
                # Elle n'est pas decorative : « FMP - Fauteuil non-modulaire »
                # et « FRM - Fauteuil modulaire » distinguent deux familles
                # que le WMS nomme lui aussi (VPH FMP ..., VPH FRM ...). Sans
                # elle, un fauteuil modulaire peut recevoir le prix d'un
                # non-modulaire.
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
    print(f"{len(table)} lignes extraites de {cible.name}")
    colonnes = [c for c in ("designation", "prix_achat_unitaire_ht",
                            "colonne_prix", "paliers") if c in table.columns]
    print(table[colonnes].head(30).to_string(index=False))
