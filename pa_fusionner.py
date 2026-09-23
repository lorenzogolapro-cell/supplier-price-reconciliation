# -*- coding: utf-8 -*-
"""
Adds the purchase-price proposals to the price owner's workbook, without
overwriting anything in it.

    python pa_fusionner.py

The price owner's workbook is authoritative, by governance rule: we
confirm, we fill gaps, we flag, we never overwrite. So this script
modifies NO existing row: it appends its own below them, in the same
columns.

Nor does it write into their file: that file is open in Excel half the
time, and above all, overwriting somebody else's working document is not
an operation you carry out without them knowing. The output is a COPY,
alongside, which they compare and then replace.

An "Origine ligne" column is added at the end of the table: it says
"Responsable prix" or "Proposition <date>". Without it the two
populations become indistinguishable from the very first save, and
nobody knows any more what has been validated.

Output: ./data/prix_valides + propositions <date>.xlsx
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

# Workbook column <- column of my proposals
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
    """The proposals to add, from both sources.

    Two seams, and the order matters:

      1. the 2026 CONSOLIDATION, "base_prix.xlsx". This is the price
         owner's main source, already arbitrated and declared good. It
         takes precedence.
      2. my proposals drawn from the supplier price lists, for everything
         the consolidation does not cover.

    An item served by both takes the value from the first.
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
            print(f"  {consolidation} coming from the 2026 consolidation")

    for fichier in sorted(SORTIE_ACHATS.glob("pa_completer_*.xlsx")):
        try:
            lot = pd.read_excel(chemin_lisible(fichier),
                                sheet_name="À compléter", skiprows=3)
        except Exception as err:
            print(f"  ! {fichier.name}: {type(err).__name__}")
            continue
        if lot.empty:
            continue
        # The pro catalogue outside the scope is a population of its own:
        # it has had neither a sale nor a receipt over twelve months. It
        # deserves a price (the item is published) but it must never be
        # mixed up with the scope in a count.
        lot["_origine"] = (
            f"Catalogue pro hors périmètre {date.today():%d/%m/%Y}"
            if "CATALOGUE-PRO" in fichier.stem.upper() else MARQUE)
        lots.append(lot)
    if not lots:
        raise SystemExit("No proposal to add.")
    tout = pd.concat(lots, ignore_index=True)
    tout = tout[pd.to_numeric(tout["Nouveau PA"],
                              errors="coerce").fillna(0) > 0]
    print(f"{len(tout)} proposals read from {len(lots)} workbooks")

    # The same item can come out of two supplier workbooks: a name
    # fragment that matches several of them, an item attached to two
    # price lists. We keep only one row.
    tout["_code"] = tout["Code article"].astype(str).str.strip()
    doubles = tout["_code"].duplicated()
    if doubles.any():
        print(f"  {int(doubles.sum())} internal duplicates dropped")
        tout = tout[~doubles]

    # A computed price drags its floating decimals along: 579.5999999999999
    # is not a price, it is an artefact. The price owner's workbook goes to
    # four decimals, so we stick to that.
    for colonne in ("Nouveau PA", "Ancien PA"):
        if colonne in tout.columns:
            tout[colonne] = pd.to_numeric(tout[colonne],
                                          errors="coerce").round(4)
    if "Écart %" in tout.columns:
        tout["Écart %"] = pd.to_numeric(tout["Écart %"],
                                        errors="coerce").round(4)
    return tout


# --- formatting -----------------------------------------------------------
# The file goes out by email: it must be readable with no instructions.
# We do not touch the values, only what helps the eye: a frozen header, a
# filter, workable widths, and above all a colour that tells what has
# been validated from what is waiting for review.
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


# The price owner's sheets that we do NOT regenerate. Their content dates
# from the day they produced them, and it no longer moves.
FEUILLES_FIGEES = ("non raproch", "non rapproch", "suivi integration")

# Mark placed IN FRONT of the tab name, so that it survives Excel's
# truncation at 31 characters.
PREFIXE_FIGE = "FIGÉ — "


def dater_les_feuilles_figees(classeur) -> None:
    """Says, on the tab itself, that a sheet is no longer up to date.

    We only write Feuil1: that is the rule, and it protects the price
    owner's work. But the "non raproché" sheet carries 624 rows frozen at
    02/09 and no longer moves: identical across the ten versions of the
    workbook. It therefore claims that Supplier J is not matched when in
    fact its sixteen items have a price in Feuil1, and it is the sheet
    people open first, because its name answers the question they are
    asking.

    Stale data under an accurate name is worse than missing data. We do
    NOT TOUCH its content: we rename the tab so that its date reads
    before its title, and we give it an alert colour.

    The content stays fully readable, and nothing is deleted.
    """
    from openpyxl.utils.exceptions import InvalidFileException  # noqa: F401

    for feuille in list(classeur.worksheets):
        titre = str(feuille.title).lower()
        if not any(marque in titre for marque in FEUILLES_FIGEES):
            continue
        if feuille.title.upper().startswith(PREFIXE_FIGE.upper()):
            continue  # already dated on a previous run
        # The mark goes IN FRONT: Excel caps a tab at 31 characters and
        # truncates on the right, so a warning placed at the end is the
        # first thing to disappear ("non raproché... (FIGÉ voir Feuil1").
        # In front, it stays readable even on a narrow tab.
        nouveau = f"{PREFIXE_FIGE}{feuille.title.strip()}"[:31]
        try:
            feuille.title = nouveau
            feuille.sheet_properties.tabColor = "C00000"
            print(f"  tab '{titre}' renamed '{nouveau}': it predates the "
                  f"workstream and is not regenerated")
        except Exception as erreur:
            print(f"  ! tab '{titre}' not renamed: "
                  f"{type(erreur).__name__}")


def embellir(feuille, entetes: list, derniere: int,
             ligne_entete: int = 1) -> None:
    """Makes the sheet readable without changing a single value."""
    for i, nom in enumerate(entetes, start=1):
        cellule = feuille.cell(row=ligne_entete, column=i)
        cellule.font = Font(bold=True, color="FFFFFF")
        cellule.fill = FOND_ENTETE
        cellule.alignment = Alignment(vertical="center", wrap_text=True)
        lettre = get_column_letter(i)
        feuille.column_dimensions[lettre].width = LARGEURS.get(str(nom), 16)
    feuille.row_dimensions[ligne_entete].height = 30
    # Freeze up to D, not C: column C carries the LABEL, and freezing at C
    # leaves it scrolling. As soon as you go and look at the prices on the
    # right you end up with rows of figures without knowing which item
    # they belong to (reported by the user on 16/09). A, B and C stay
    # visible.
    feuille.freeze_panes = f"D{ligne_entete + 1}"
    # Feuil1 is an Excel table: it already carries its own filter. Adding a
    # second one at sheet level makes the workbook unreadable for Excel.
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

        # An alert has to be seen: it is the only place where the eye
        # should stop in a three-thousand-row table.
        if col_alerte and feuille.cell(row=r, column=col_alerte).value:
            feuille.cell(row=r, column=col_alerte).font = ROUGE


def recadrer_tableau(feuille, entetes: list, ligne_entete: int,
                     derniere: int) -> None:
    """Realigns the sheet's Excel table onto the real data.

    Feuil1 is not an ordinary range: it is a Table object, carrying its
    own column list and its own extent. Adding rows or a column without
    telling it leaves its definition pointing elsewhere, and Excel then
    declares the workbook damaged.
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
    """58,4 % : comma and space, like the rest of the workbook."""
    return f"{part * 100:.1f} %".replace(".", ",")


def ligne_couverture(feuille, entetes: list, derniere: int) -> int:
    """Inserts a cover block at the top saying where the work stands.

    A three-thousand-row workbook says neither which population it covers
    nor how much of it is covered. Three lines: the rate, what it is made
    of, then the definition of the denominator, because a percentage
    whose population is unknown means nothing, and the reader of this
    file does not have the scope manifest in front of them.

    Returns the number of rows inserted.
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
        raise SystemExit(f"Workbook not found: {CLASSEUR}")

    nouvelles = propositions()

    # We work on a readable copy: the workbook is often open
    classeur = load_workbook(chemin_lisible(CLASSEUR))
    feuille = classeur[FEUILLE]
    entetes = [c.value for c in feuille[1]]
    print(f"workbook: {feuille.max_row - 1} rows, "
          f"{len(entetes)} columns")

    # Last row ACTUALLY filled: the workbook drags empty rows along from
    # an earlier version, and writing into them would leave a hole.
    colonne_code = entetes.index("Code article") + 1
    derniere = 1
    deja = set()
    for r in range(2, feuille.max_row + 1):
        valeur = feuille.cell(row=r, column=colonne_code).value
        if valeur not in (None, ""):
            derniere = r
            deja.add(str(valeur).strip())
    print(f"  last filled row: {derniere} "
          f"({len(deja)} distinct item codes)")

    # Guardrail: never add an item the price owner already carries
    nouvelles["_code"] = nouvelles["Code article"].astype(str).str.strip()
    doublons = nouvelles["_code"].isin(deja)
    if doublons.any():
        print(f"  {int(doublons.sum())} proposals dropped: "
              f"the item is already in the workbook")
        nouvelles = nouvelles[~doublons]

    # Variant code, which my files do not carry
    art = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Code déclinaison"])
    pont = art.dropna(subset=["Code article"]).drop_duplicates("Code article")
    decl = dict(zip(pont["Code article"].str.strip(),
                    pont["Code déclinaison"]))

    # Origin column, added at the end of the table
    if "Origine ligne" in entetes:
        colonne_origine = entetes.index("Origine ligne") + 1
    else:
        colonne_origine = len(entetes) + 1
        feuille.cell(row=1, column=colonne_origine, value="Origine ligne")
        for r in range(2, derniere + 1):
            feuille.cell(row=r, column=colonne_origine,
                         value="Responsable prix")
        print(f"  column 'Origine ligne' added at position "
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
    # The cover block shifts everything down: the header moves to row 2.
    decale = ligne_couverture(feuille, entetes_finaux, derniere)
    embellir(feuille, entetes_finaux, derniere + decale,
             ligne_entete=1 + decale)
    recadrer_tableau(feuille, entetes_finaux, 1 + decale, derniere + decale)

    # The output workbook is often open in Excel: rather than failing
    # after all the computation, we write alongside under a timestamped
    # name, and say so.
    sortie = mise_en_forme.chemin_ecriture(
        DONNEES / f"prix_valides + propositions "
                  f"{date.today():%Y-%m-%d}.xlsx")
    classeur.save(sortie)
    classeur.close()

    print(f"\n  {ecrites} rows added, 0 existing row modified")
    print(f"  -> {sortie}")
    print("\n  The proposals are filtered on 'Origine ligne'.")
    print("  Nothing has been written into the original workbook.")


if __name__ == "__main__":
    main()
