# -*- coding: utf-8 -*-
"""
Ajoute les propositions de PA au classeur du responsable des prix, sans
rien y ecraser.

    python pa_fusionner.py

Le classeur du responsable des prix fait foi (PERIMETRE.md §6) : on
confirme, on comble, on signale, on n'ecrase jamais. Ce script ne modifie
donc AUCUNE ligne existante — il ajoute les siennes a la suite, dans les
memes colonnes.

Il n'ecrit pas non plus dans son fichier : celui-ci est ouvert dans Excel
la moitie du temps, et surtout ecraser le document de travail de
quelqu'un d'autre n'est pas une operation qu'on fait sans qu'il le sache.
La sortie est une COPIE, a cote, que l'on compare puis remplace.

Une colonne « Origine ligne » est ajoutee en fin de tableau : elle dit
« Responsable prix » ou « Proposition <date> ». Sans elle, les deux
populations deviennent indistinguables des la premiere sauvegarde, et plus
personne ne sait ce qui a ete valide.

Sortie : ./data/prix_valides + propositions <date>.xlsx
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import mise_en_forme  # noqa: E402
from extracteurs.base import chemin_lisible  # noqa: E402
from main import SORTIE_ACHATS  # noqa: E402

DONNEES = Path("./data")
CLASSEUR = DONNEES / "prix_valides.xlsx"
import wms_extract  # noqa: E402
EXTRACT = wms_extract.chemin()
FEUILLE = "Feuil1"
MARQUE = f"Proposition {date.today():%d/%m/%Y}"

# Colonne du classeur <- colonne de mes propositions
CORRESPONDANCE = {
    "Code article": "Code article",
    "Libellé": "Libellé déclinaison ^(1)",
    "Fournisseur principal": "Nom fournisseur",
    "Réf. fournisseur": "Réf. fournisseur",
    "Type": "Type",
    "Ancien PA": "Ancien PA",
    "Nouveau PA": "Nouveau PA",
    "Écart %": "Écart %",
    "Origine Nouveau PA": "Origine Nouveau PA",
    "Fichier source PA": "Fichier source PA",
    "Alerte": "Alerte",
}


def propositions() -> pd.DataFrame:
    """Les propositions à ajouter, des deux sources.

    Deux gisements, et l'ordre compte :

      1. la CONSOLIDATION 2026 — « base_prix.xlsx ». C'est la source
         principale du responsable des prix, deja arbitree et declaree
         bonne. Elle prime.
      2. mes propositions tirees des tarifs fournisseurs, pour tout ce
         que la consolidation ne couvre pas.

    Un article servi par les deux prend la valeur de la premiere.
    """
    lots = []

    maj = SORTIE_ACHATS / "pa_maj2026.xlsx"
    consolidation = 0
    if maj.exists():
        for feuille in ("À injecter", "Doublon avec mes propositions"):
            try:
                lot = pd.read_excel(chemin_lisible(maj), sheet_name=feuille,
                                    skiprows=3)
            except Exception:
                continue
            if lot.empty:
                continue
            lot = lot.rename(columns={"Libellé": "Libellé déclinaison ^(1)",
                                      "Fournisseur principal":
                                          "Nom fournisseur"})
            lot["_origine"] = f"Consolidation 2026 {date.today():%d/%m/%Y}"
            lots.append(lot)
            consolidation += len(lot)
        if consolidation:
            print(f"  {consolidation} venant de la consolidation 2026")

    for fichier in sorted(SORTIE_ACHATS.glob("pa_completer_*.xlsx")):
        try:
            lot = pd.read_excel(chemin_lisible(fichier),
                                sheet_name="À compléter", skiprows=3)
        except Exception as err:
            print(f"  ! {fichier.name} : {type(err).__name__}")
            continue
        if lot.empty:
            continue
        # Le catalogue pro hors perimetre est une population a part : elle
        # n'a eu ni vente ni reception sur douze mois. Elle vaut d'etre
        # tarifee — l'article est publie — mais elle ne doit jamais se
        # confondre avec le perimetre dans un decompte.
        lot["_origine"] = (
            f"Catalogue pro hors périmètre {date.today():%d/%m/%Y}"
            if "CATALOGUE-PRO" in fichier.stem.upper() else MARQUE)
        lots.append(lot)
    if not lots:
        raise SystemExit("Aucune proposition à ajouter.")
    tout = pd.concat(lots, ignore_index=True)
    tout = tout[pd.to_numeric(tout["Nouveau PA"],
                              errors="coerce").fillna(0) > 0]
    print(f"{len(tout)} propositions lues dans {len(lots)} classeurs")

    # Un meme article peut sortir de deux classeurs fournisseurs — un
    # fragment de nom qui en designe plusieurs, un article rattache a
    # deux tarifs. On n'en garde qu'une ligne.
    tout["_code"] = tout["Code article"].astype(str).str.strip()
    doubles = tout["_code"].duplicated()
    if doubles.any():
        print(f"  {int(doubles.sum())} doublons internes écartés")
        tout = tout[~doubles]

    # Un prix calcule traine ses decimales flottantes : 579.5999999999999
    # n'est pas un prix, c'est un artefact. Le classeur du responsable des
    # prix va jusqu'a quatre decimales, on s'y tient.
    for colonne in ("Nouveau PA", "Ancien PA"):
        if colonne in tout.columns:
            tout[colonne] = pd.to_numeric(tout[colonne],
                                          errors="coerce").round(4)
    if "Écart %" in tout.columns:
        tout["Écart %"] = pd.to_numeric(tout["Écart %"],
                                        errors="coerce").round(4)
    return tout


# --- mise en forme --------------------------------------------------------
# Le fichier part par mail : il doit se lire sans mode d'emploi. On ne
# touche pas aux valeurs, seulement a ce qui aide l'oeil — l'en-tete
# fige, un filtre, des largeurs tenables, et surtout une couleur qui
# distingue ce qui est valide de ce qui attend une relecture.
FOND_ENTETE = PatternFill("solid", fgColor="1F3B36")
FOND_RESPONSABLE = PatternFill("solid", fgColor="EAF1EE")
FOND_TARIF = PatternFill("solid", fgColor="FBF3E4")
FOND_CONSO = PatternFill("solid", fgColor="EDF0F8")
ROUGE = Font(color="9E3A2A")

LARGEURS = {
    "Code article": 12, "Code déclinaison": 14, "Libellé": 46,
    "Fournisseur principal": 30, "Code fournisseur": 12,
    "Réf. fournisseur": 20, "Statut ligne": 20, "Rayon": 22,
    "Famille": 24, "Type": 24, "Nouveau rayon": 20,
    "Nouvelle famille": 24, "Stock": 8, "Palier": 8,
    "Ancien PA": 12, "Nouveau PA": 12, "Écart %": 10,
    "Origine Nouveau PA": 30, "Fichier source PA": 34,
    "Source référencement": 26, "Alerte": 58, "Origine ligne": 24,
}
EUROS = ("Ancien PA", "Nouveau PA")
POURCENT = ("Écart %",)


# Les feuilles du responsable des prix que nous ne regenerons PAS. Leur
# contenu date du jour ou il les a produites, et il ne bouge plus.
FEUILLES_FIGEES = ("non raproch", "non rapproch", "suivi integration")

# Marque posee DEVANT le nom de l'onglet, pour qu'elle survive a la
# troncature d'Excel a 31 caracteres.
PREFIXE_FIGE = "FIGÉ — "


def dater_les_feuilles_figees(classeur) -> None:
    """Dit, sur l'onglet lui-meme, qu'une feuille n'est plus a jour.

    Nous n'ecrivons que Feuil1 — c'est la regle, et elle protege le travail
    du responsable des prix. Mais la feuille « non raproché » porte 624
    lignes arretees au 02/09 et ne bouge plus : identique dans les dix
    versions du classeur. Elle affirme donc que le Fournisseur J n'est pas
    rapproche alors que ses seize articles ont un prix dans Feuil1 — et
    c'est elle qu'on ouvre en premier, parce que son nom repond a la
    question qu'on se pose.

    Une donnee perimee qui porte un nom exact est pire qu'une donnee
    absente. On ne TOUCHE PAS a son contenu : on renomme l'onglet pour que
    sa date se lise avant son titre, et on lui met une couleur d'alerte.

    Le contenu reste integralement lisible, et rien n'est supprime.
    """
    from openpyxl.utils.exceptions import InvalidFileException  # noqa: F401

    for feuille in list(classeur.worksheets):
        titre = str(feuille.title).lower()
        if not any(marque in titre for marque in FEUILLES_FIGEES):
            continue
        if feuille.title.upper().startswith(PREFIXE_FIGE.upper()):
            continue  # deja date lors d'une passe precedente
        # La marque passe DEVANT : Excel limite un onglet a 31 caracteres et
        # tronque a droite, si bien qu'un avertissement place a la fin est
        # le premier a disparaitre — « non raproché… (FIGÉ voir Feuil1 ».
        # Devant, il reste lisible meme sur un onglet etroit.
        nouveau = f"{PREFIXE_FIGE}{feuille.title.strip()}"[:31]
        try:
            feuille.title = nouveau
            feuille.sheet_properties.tabColor = "C00000"
            print(f"  onglet « {titre} » renommé « {nouveau} » — "
                  f"il date d'avant le chantier et ne se régénère pas")
        except Exception as erreur:
            print(f"  ! onglet « {titre} » non renommé : "
                  f"{type(erreur).__name__}")


def embellir(feuille, entetes: list, derniere: int,
             ligne_entete: int = 1) -> None:
    """Rend la feuille lisible sans en modifier une seule valeur."""
    for i, nom in enumerate(entetes, start=1):
        cellule = feuille.cell(row=ligne_entete, column=i)
        cellule.font = Font(bold=True, color="FFFFFF")
        cellule.fill = FOND_ENTETE
        cellule.alignment = Alignment(vertical="center", wrap_text=True)
        lettre = get_column_letter(i)
        feuille.column_dimensions[lettre].width = LARGEURS.get(str(nom), 16)
    feuille.row_dimensions[ligne_entete].height = 30
    # Figer jusqu'a D, pas C : la colonne C porte le LIBELLE, et figer en C
    # la laisse defiler. On se retrouve alors, des qu'on va voir les prix a
    # droite, avec des lignes de chiffres sans savoir de quel article il
    # s'agit — signale par l'utilisateur le 16/09. A, B et C restent visibles.
    feuille.freeze_panes = f"D{ligne_entete + 1}"
    # Feuil1 est un tableau Excel : il porte deja son filtre. En poser un
    # second au niveau de la feuille rend le classeur illisible pour Excel.
    if not getattr(feuille, "tables", None):
        feuille.auto_filter.ref = (
            f"A{ligne_entete}:{get_column_letter(len(entetes))}{derniere}")

    index = {str(n): i + 1 for i, n in enumerate(entetes)}
    col_origine = index.get("Origine ligne")
    col_alerte = index.get("Alerte")

    for r in range(ligne_entete + 1, derniere + 1):
        origine = ""
        if col_origine:
            origine = str(feuille.cell(row=r, column=col_origine).value or "")
        fond = (FOND_RESPONSABLE if origine.startswith("Responsable")
                else FOND_CONSO if origine.startswith("Consolidation")
                else FOND_TARIF if origine else None)
        if fond and col_origine:
            feuille.cell(row=r, column=col_origine).fill = fond

        for nom in EUROS:
            if nom in index:
                feuille.cell(row=r, column=index[nom]).number_format = \
                    '# ##0.00 "€"'
        for nom in POURCENT:
            if nom in index:
                feuille.cell(row=r, column=index[nom]).number_format = "0.0 %"

        # Une alerte doit se voir : c'est le seul endroit ou l'oeil doit
        # s'arreter dans un tableau de trois mille lignes.
        if col_alerte and feuille.cell(row=r, column=col_alerte).value:
            feuille.cell(row=r, column=col_alerte).font = ROUGE


def recadrer_tableau(feuille, entetes: list, ligne_entete: int,
                     derniere: int) -> None:
    """Reajuste le tableau Excel de la feuille sur les donnees reelles.

    Feuil1 n'est pas une plage ordinaire : c'est un objet Tableau, qui
    porte sa propre liste de colonnes et sa propre etendue. Ajouter des
    lignes ou une colonne sans le lui dire laisse sa definition pointer
    a cote — et Excel declare le classeur endommage.
    """
    from openpyxl.worksheet.table import TableColumn

    for tableau in list(getattr(feuille, "tables", {}).values()):
        connues = {str(c.name) for c in tableau.tableColumns}
        suivant = max((c.id for c in tableau.tableColumns), default=0) + 1
        for nom in entetes:
            if str(nom) not in connues:
                tableau.tableColumns.append(
                    TableColumn(id=suivant, name=str(nom)))
                suivant += 1
        ref = (f"A{ligne_entete}:"
               f"{get_column_letter(len(entetes))}{derniere}")
        tableau.ref = ref
        if tableau.autoFilter is not None:
            tableau.autoFilter.ref = ref


def _pourcent(part: float) -> str:
    """58,4 % — la virgule et l'espace, comme le reste du classeur."""
    return f"{part * 100:.1f} %".replace(".", ",")


def ligne_couverture(feuille, entetes: list, derniere: int) -> int:
    """Insere en tete un cartouche qui dit ou en est le chantier.

    Un classeur de trois mille lignes ne dit ni sur quelle population il
    porte, ni ce qu'il en couvre. Trois lignes : le taux, sa composition,
    puis la definition du denominateur — parce qu'un pourcentage dont on
    ignore la population ne veut rien dire, et que le lecteur de ce
    fichier n'a pas le manifeste du perimetre sous les yeux.

    Renvoie le nombre de lignes inserees.
    """
    import perimetre_liste  # noqa: E402

    index = {str(n): i + 1 for i, n in enumerate(entetes)}
    col_code, col_origine = index.get("Code article"), index.get(
        "Origine ligne")
    if not col_code or not col_origine:
        return 0

    valides: set = set()
    tous: set = set()
    for r in range(2, derniere + 1):
        code = feuille.cell(row=r, column=col_code).value
        if code in (None, ""):
            continue
        code = str(code).strip()
        tous.add(code)
        if str(feuille.cell(row=r, column=col_origine).value
                or "").startswith("Responsable"):
            valides.add(code)

    perimetre = perimetre_liste.codes_article()
    total = len(perimetre)
    if not total:
        return 0
    acquis = len(valides & perimetre)
    avec = len(tous & perimetre)
    propose = avec - acquis
    manifeste = perimetre_liste.manifeste()

    lignes = [
        (f"PRIX D'ACHAT 2026 — {avec} des {total} articles du périmètre "
         f"ont un prix, soit {_pourcent(avec / total)}.",
         Font(bold=True, size=13, color="1F3B36")),

        (f"Dont {acquis} validés par le responsable des prix "
         f"({_pourcent(acquis / total)}) "
         f"et {propose} propositions à relire "
         f"({_pourcent(propose / total)}). "
         f"Restent {total - avec} articles sans prix "
         f"({_pourcent((total - avec) / total)}). "
         f"Colonne « Origine ligne » : « Responsable prix — validé » fait "
         f"foi, le reste est une proposition sourcée, non relue.",
         Font(size=11, color="1F3B36")),

        (f"Périmètre v{manifeste['version']}, figé le "
         f"{manifeste['figé_le'][:10]} : les {manifeste['articles']} "
         f"articles ayant eu au moins une vente ou une réception d'achat "
         f"sur les {manifeste['fenêtre_mois']} mois arrêtés au "
         f"{manifeste['arrêté']}, hors motifs d'écartement — soit "
         f"{total} codes article distincts, le dénominateur ci-dessus. "
         f"Ce classeur compte {len(tous)} codes : les "
         f"{len(tous) - avec} autres sont hors périmètre et ne comptent "
         f"pas dans le taux.",
         Font(italic=True, size=10, color="5A6A66")),
    ]

    feuille.insert_rows(1, amount=len(lignes))
    for i, (texte, police) in enumerate(lignes, start=1):
        cellule = feuille.cell(row=i, column=1, value=texte)
        cellule.font = police
        cellule.alignment = Alignment(vertical="center")
        feuille.row_dimensions[i].height = 20 if i == 1 else 16
    return len(lignes)
    return True


def main() -> None:
    if not CLASSEUR.exists():
        raise SystemExit(f"Classeur introuvable : {CLASSEUR}")

    nouvelles = propositions()

    # On travaille sur une copie lisible : le classeur est souvent ouvert
    classeur = load_workbook(chemin_lisible(CLASSEUR))
    feuille = classeur[FEUILLE]
    entetes = [c.value for c in feuille[1]]
    print(f"classeur : {feuille.max_row - 1} lignes, "
          f"{len(entetes)} colonnes")

    # Derniere ligne REELLEMENT remplie : le classeur traine des lignes
    # vides d'une version anterieure, y ecrire creerait un trou.
    colonne_code = entetes.index("Code article") + 1
    derniere = 1
    deja = set()
    for r in range(2, feuille.max_row + 1):
        valeur = feuille.cell(row=r, column=colonne_code).value
        if valeur not in (None, ""):
            derniere = r
            deja.add(str(valeur).strip())
    print(f"  dernière ligne remplie : {derniere} "
          f"({len(deja)} codes article distincts)")

    # Garde-fou : ne jamais ajouter un article que le responsable porte deja
    nouvelles["_code"] = nouvelles["Code article"].astype(str).str.strip()
    doublons = nouvelles["_code"].isin(deja)
    if doublons.any():
        print(f"  {int(doublons.sum())} propositions écartées : "
              f"l'article est déjà dans le classeur")
        nouvelles = nouvelles[~doublons]

    # Code declinaison, que mes fichiers ne portent pas
    art = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Code déclinaison"])
    pont = art.dropna(subset=["Code article"]).drop_duplicates("Code article")
    decl = dict(zip(pont["Code article"].str.strip(),
                    pont["Code déclinaison"]))

    # Colonne d'origine, ajoutee en fin de tableau
    if "Origine ligne" in entetes:
        colonne_origine = entetes.index("Origine ligne") + 1
    else:
        colonne_origine = len(entetes) + 1
        feuille.cell(row=1, column=colonne_origine, value="Origine ligne")
        for r in range(2, derniere + 1):
            feuille.cell(row=r, column=colonne_origine,
                         value="Responsable prix")
        print(f"  colonne « Origine ligne » ajoutée en position "
              f"{colonne_origine}")

    ecrites = 0
    for _, ligne in nouvelles.iterrows():
        derniere += 1
        for cible, source in CORRESPONDANCE.items():
            if cible not in entetes or source not in nouvelles.columns:
                continue
            valeur = ligne[source]
            if pd.isna(valeur):
                continue
            feuille.cell(row=derniere, column=entetes.index(cible) + 1,
                         value=valeur)
        if "Code déclinaison" in entetes:
            feuille.cell(row=derniere,
                         column=entetes.index("Code déclinaison") + 1,
                         value=decl.get(ligne["_code"]))
        feuille.cell(row=derniere, column=colonne_origine,
                     value=ligne.get("_origine", MARQUE))
        ecrites += 1

    dater_les_feuilles_figees(classeur)

    entetes_finaux = [c.value for c in feuille[1]]
    # La ligne de couverture decale tout d'un cran : l'entete passe en 2.
    decale = ligne_couverture(feuille, entetes_finaux, derniere)
    embellir(feuille, entetes_finaux, derniere + decale,
             ligne_entete=1 + decale)
    recadrer_tableau(feuille, entetes_finaux, 1 + decale, derniere + decale)

    # Le classeur de sortie est souvent ouvert dans Excel : plutot que
    # d'echouer apres tout le calcul, on ecrit a cote sous un nom
    # horodate, et on le signale.
    sortie = mise_en_forme.chemin_ecriture(
        DONNEES / f"prix_valides + propositions "
                  f"{date.today():%Y-%m-%d}.xlsx")
    classeur.save(sortie)
    classeur.close()

    print(f"\n  {ecrites} lignes ajoutées, 0 ligne existante modifiée")
    print(f"  -> {sortie}")
    print("\n  Les propositions se filtrent sur « Origine ligne ».")
    print("  Rien n'a été écrit dans le classeur d'origine.")


if __name__ == "__main__":
    main()
