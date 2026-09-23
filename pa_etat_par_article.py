# -*- coding: utf-8 -*-
"""
State of the 2026 purchase price, item by item, across the whole scope.

    python pa_etat_par_article.py

The other workbooks in this workstream each show one piece: a supplier,
a sub-population, the proposals. None of them answers the question
people actually ask in front of an item: where does this one stand, and
if nothing has been done, why?

One row per item of the frozen scope, all 4,731 of them, including those
there is nothing to say about: their number is precisely what counts. No
filter is applied: anything that could be one becomes a column instead,
so that sorting does not require rerunning a job.

Output: sortie/3-achats/pa_etat_par_article.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import mise_en_forme  # noqa: E402
import perimetre_liste  # noqa: E402
from extracteurs.base import chemin_lisible  # noqa: E402
from main import SORTIE_ACHATS  # noqa: E402

DONNEES = Path("./data")
import wms_extract  # noqa: E402
EXTRACT = wms_extract.chemin()

REPRISES = ["Ancien PA", "Nouveau PA", "Écart %", "Origine Nouveau PA",
            # The reference as it appears IN THE PRICE LIST: that is the
            # one used to order, and it does not always coincide with the
            # one on the item record.
            "Réf. fournisseur",
            "Fichier source PA", "Alerte", "Clé de rapprochement"]


def code_article(colonne: pd.Series) -> pd.Series:
    """The item code as text, without the ".0" pandas sticks onto it.

    An integer column containing one blank turns into a float column, and
    "32818" then reads "32818.0". The join drops to zero without raising
    the slightest error: the kind of failure that only shows in the final
    result, when it is too late.
    """
    texte = colonne.astype(str).str.strip()
    return texte.str.replace(r"\.0$", "", regex=True)


def classeur_fusionne() -> Path:
    """The latest validated-prices + proposals workbook produced."""
    trouves = sorted(DONNEES.glob("prix_valides + propositions *.xlsx"),
                     key=lambda f: f.stat().st_mtime)
    if not trouves:
        raise SystemExit("Run pa_fusionner.py first")
    return trouves[-1]


def prix_connus() -> pd.DataFrame:
    """What the consolidated workbook knows, per item code.

    We read the workbook produced rather than redoing the consolidation:
    the published figure and the one in this state view must be the same,
    otherwise one of the two is lying.
    """
    chemin = classeur_fusionne()
    print(f"  price source: {chemin.name}")
    feuille = load_workbook(chemin_lisible(chemin), read_only=True)["Feuil1"]
    lignes = list(feuille.iter_rows(values_only=True))
    # The cover block takes up the first rows: the header is the first
    # row carrying "Code article".
    depart = next(i for i, l in enumerate(lignes)
                  if l and "Code article" in [str(c) for c in l])
    entetes = [str(c) for c in lignes[depart]]
    table = pd.DataFrame(lignes[depart + 1:], columns=entetes)
    table = table.loc[:, ~table.columns.duplicated()]
    table["Code article"] = code_article(table["Code article"])
    table = table[table["Code article"].ne("") & table["Code article"].ne("None")]
    garde = ["Code article", "Origine ligne"] + [c for c in REPRISES
                                                 if c in table.columns]
    return table[garde].drop_duplicates("Code article")


def motifs() -> pd.DataFrame:
    """Why an item has no price, according to the sweeps.

    The three causes are not equivalent and are not handled the same way:
    an excluded price exists and is waiting for a decision, an item
    missing from the price list calls for a reference check, a supplier
    with no price list calls for a letter. Confusing them would lose the
    only useful ranking.
    """
    lots = []
    for fichier in sorted(SORTIE_ACHATS.glob("pa_completer_*.xlsx")):
        for feuille, motif in (
                ("Écartés", "prix trouvé mais écarté — remise non portée "
                            "par le tarif, à arbitrer"),
                ("Sans tarif", "fournisseur tarifé, mais l'article n'est "
                               "pas dans son tarif")):
            try:
                lot = pd.read_excel(chemin_lisible(fichier),
                                    sheet_name=feuille, skiprows=3)
            except Exception:
                continue
            if lot.empty or "Code article" not in lot.columns:
                continue
            lots.append(pd.DataFrame({
                "Code article": code_article(lot["Code article"]),
                "Motif": motif,
                "PA écarté (à vérifier)": lot.get("PA écarté (à vérifier)"),
                "Alerte balayage": lot.get("Alerte"),
            }))
    if not lots:
        return pd.DataFrame(columns=["Code article", "Motif"])
    # A decidable gap wins over an absence: it is the more advanced of
    # the two leads.
    tout = pd.concat(lots, ignore_index=True)
    tout["_rang"] = tout["Motif"].str.startswith("prix trouvé").map(
        {True: 0, False: 1})
    return (tout.sort_values("_rang").drop_duplicates("Code article")
            .drop(columns="_rang"))


def main() -> None:
    dedans = perimetre_liste.charger().copy()
    print(f"  {perimetre_liste.entete()}")
    dedans["PERIMETRE_VERSION"] = perimetre_liste.manifeste()["version"]

    wms = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Nom fabriquant",
                                 "Référence fabricant", "Ref. art. four.",
                                 "Nom fournisseur"])
    wms = wms.dropna(subset=["Code article"]).drop_duplicates("Code article")
    wms["Code article"] = wms["Code article"].str.strip()
    wms = wms.rename(columns={"Nom fournisseur": "Fournisseur (fiche WMS)",
                              "Ref. art. four.": "Réf. article fournisseur"})
    etat = dedans.merge(wms, on="Code article", how="left")

    etat = etat.merge(prix_connus(), on="Code article", how="left")
    etat = etat.merge(motifs(), on="Code article", how="left")

    # The Ancien PA came from the consolidated workbook alone, but an item
    # with no 2026 price does not appear in it, and so came out with no
    # CURRENT price either. That is exactly the opposite of what we want
    # to see: on an item still to be done, knowing what we pay today is
    # the first useful piece of information. We fill it from the WMS.
    try:
        import pa_completer
        actuels = (pa_completer.pa_actuels()
                   .drop_duplicates("code_article")
                   .set_index("code_article")["pa_wms"])
        if "Ancien PA" not in etat.columns:
            etat["Ancien PA"] = None
        vide = etat["Ancien PA"].isna()
        etat.loc[vide, "Ancien PA"] = etat.loc[vide, "Code article"].map(
            actuels)
        print(f"  Ancien PA filled from the WMS on "
              f"{int(vide.sum() - etat['Ancien PA'].isna().sum())} items")
    except Exception as err:                       # pragma: no cover
        print(f"  ! WMS purchase price unavailable: {type(err).__name__}")

    # HUMAN DECISIONS, APPLIED HERE AND NOWHERE FURTHER DOWN (18/09)
    #
    # Until now they were applied in the final view only. Consequence: NO
    # other view saw them, neither the reconciliation, which re-reported
    # already-settled divergences on every run, nor the check against
    # orders, nor the decision sheets. An alarm that still rings after the
    # problem has been dealt with ends up being ignored, and that is the
    # last thing we want on a tool that is rerun at every pass.
    #
    # Since `pa_etat_par_article` is the base from which ALL the views
    # derive, this is where the corrections must enter. They do so BEFORE
    # the status is computed, so that a corrected item stops being "sans
    # prix", and NEVER on one of the price owner's rows.
    try:
        import pa_corrections_17_09 as corrections

        codes = code_article(etat["Code article"])
        est_valide = etat["Origine ligne"].fillna("").str.startswith(
            "Responsable")
        nouveau = pd.to_numeric(etat["Nouveau PA"], errors="coerce")
        ancien = pd.to_numeric(etat["Ancien PA"], errors="coerce")
        poses, gardes = 0, 0

        for entree in corrections.PRIX:
            cible = (codes == str(entree["code_article"])) & ~est_valide
            if not cible.any():
                continue
            etat.loc[cible, "Nouveau PA"] = entree["prix"]
            etat.loc[cible, "Origine Nouveau PA"] = "Arbitrage manuel"
            etat.loc[cible, "Fichier source PA"] = "pa_corrections_17_09.py"
            etat.loc[cible, "Clé de rapprochement"] = "arbitrage manuel"
            poses += int(cible.sum())

        # Prices written straight into the last column of the deliverable,
        # recorded in the durable register. Same rule: never on a row
        # already validated by the price owner.
        import pa_annotations_livrable

        for code, prix in pa_annotations_livrable.prix_annotes().items():
            cible = (codes == str(code)) & ~est_valide
            if not cible.any():
                continue
            etat.loc[cible, "Nouveau PA"] = float(prix)
            etat.loc[cible, "Origine Nouveau PA"] = (
                "Annotation du livrable")
            etat.loc[cible, "Fichier source PA"] = (
                "pa_annotations_livrable.csv")
            etat.loc[cible, "Clé de rapprochement"] = "annotation livrable"
            poses += int(cible.sum())

        # "GARDE": the proposed candidate was wrong, the former purchase
        # price takes its place back. With no former price there is
        # nothing to take back: we do not post a price we do not have.
        for entree in corrections.GARDE:
            cible = ((codes == str(entree["code_article"])) & ~est_valide
                     & ancien.notna())
            if not cible.any():
                continue
            etat.loc[cible, "Nouveau PA"] = ancien[cible]
            etat.loc[cible, "Origine Nouveau PA"] = (
                "Arbitrage manuel — ancien PA confirmé")
            etat.loc[cible, "Fichier source PA"] = "pa_corrections_17_09.py"
            gardes += int(cible.sum())

        if poses or gardes:
            # The gap is recomputed on the rows touched, otherwise it
            # keeps the trace of the previous price and talks nonsense.
            n = pd.to_numeric(etat["Nouveau PA"], errors="coerce")
            a = pd.to_numeric(etat["Ancien PA"], errors="coerce")
            calculable = n.notna() & a.notna() & (a != 0)
            etat.loc[calculable, "Écart %"] = ((n - a) / a).round(4)
            print(f"  manual decisions: {poses} prices posted, "
                  f"{gardes} former purchase prices confirmed")
    except Exception as err:                           # pragma: no cover
        print(f"  ! manual decisions unavailable: {type(err).__name__}: "
              f"{err}")

    # The status is the only column that reads at a glance: three values,
    # no more, and what the price owner has validated is kept apart from
    # the rest because it alone has been reviewed.
    origine = etat["Origine ligne"].fillna("")
    a_un_prix = pd.to_numeric(etat["Nouveau PA"], errors="coerce").notna()
    etat["Statut PA"] = "3 — SANS PRIX"
    etat.loc[a_un_prix, "Statut PA"] = "2 — PROPOSÉ (à relire)"
    etat.loc[origine.str.startswith("Responsable"), "Statut PA"] = \
        "1 — ACQUIS (validé par le responsable)"

    # An item with no price AND no known reason has never been swept: its
    # supplier sent no price list at all. Saying so is more useful than
    # leaving the cell empty.
    sans_motif = (etat["Statut PA"] == "3 — SANS PRIX") & etat["Motif"].isna()
    sans_fournisseur = (etat["Fournisseur"].fillna("").str.strip() == "")
    etat.loc[sans_motif & ~sans_fournisseur, "Motif"] = \
        "aucun tarif reçu de ce fournisseur"
    etat.loc[sans_motif & sans_fournisseur, "Motif"] = (
        "aucun fournisseur au périmètre — voir le fabricant")
    etat.loc[etat["Statut PA"] != "3 — SANS PRIX", "Motif"] = ""

    # Suppliers we have left leave the workstream: their items are not
    # "sans prix" waiting for a price list, they are OUT OF SCOPE OF THE
    # WORK. The distinction is not cosmetic: an item with no price calls
    # for work, this one does not. Without it, they come back into the
    # to-do lists at every run, and somebody eventually goes looking for
    # their price list.
    from fournisseurs_arretes import est_arrete, motif as motif_arret

    arretes = etat["Fournisseur"].fillna("").map(est_arrete)
    a_marquer = arretes & (etat["Statut PA"] == "3 — SANS PRIX")
    etat.loc[a_marquer, "Motif"] = etat.loc[a_marquer, "Fournisseur"].map(
        motif_arret)
    etat.loc[arretes, "Statut PA"] = "4 — HORS CHANTIER (fournisseur arrêté)"
    if arretes.any():
        print(f"    {int(arretes.sum())} items out of the workstream "
              f"(supplier dropped), of which "
              f"{int(a_marquer.sum())} had no price")

    # WHAT WE SET ASIDE IS NOT WHAT RESISTS (18/09)
    #
    # "3 — SANS PRIX" mixed two things that have nothing to do with each
    # other: the items we cannot manage to price, and those we have
    # DECIDED not to handle (beds, wheelchairs, VPH, discontinued
    # products). Out of 1,078 "sans prix", 522 were in the second case.
    # Reading that figure as work remaining is being wrong by a factor of
    # two, and that is what happened on 18/09 in a note meant for an
    # outside reader.
    #
    # Rule no. 3 of the workstream says what to do: "every filter becomes
    # a column". So setting aside becomes a COLUMN OF THE BASE VIEW,
    # rather than information you have to go and find in a second file by
    # cross-referencing two sources.
    #
    # The classification comes from `pa_ecartes.classer()`, the SAME
    # function that produces the exclusion sheets, never a copy: two
    # written forms of the same rule would end up contradicting each
    # other.
    #
    # `Statut PA` is not touched. The existing views that rely on
    # "3 — SANS PRIX" keep working; those that want the real work
    # remaining filter on an empty "Mise à l'écart".
    try:
        import pa_ecartes

        annot = pa_ecartes.annotations()
        avec = etat.merge(annot, on="Code article", how="left")
        etat["Mise à l'écart"] = pa_ecartes.classer(avec)["motif"].values
        ecartes = etat["Mise à l'écart"].fillna("") != ""
        sans_prix = etat["Statut PA"] == "3 — SANS PRIX"
        print(f"    {int(ecartes.sum())} items set aside "
              f"(bed/wheelchair/VPH, supplier dropped, discontinued)")
        print(f"    REAL work remaining: "
              f"{int((sans_prix & ~ecartes).sum())} without a price in the "
              f"active scope, out of {int(sans_prix.sum())} 'sans prix' "
              f"in total")
    except Exception as err:                           # pragma: no cover
        etat["Mise à l'écart"] = ""
        print(f"  ! set-aside unavailable: {type(err).__name__}: {err}")

    colonnes = [
        "PERIMETRE_VERSION", "Réf. interne", "Code article",
        "Code déclinaison", "Désignation", "Type", "Famille",
        "Fournisseur", "Nom fabriquant", "Réf. article fournisseur",
        "Référence fabricant", "Réf. fournisseur",
        "Statut PA", "Mise à l'écart", "Motif",
        "Ancien PA", "Nouveau PA", "Écart %", "Origine ligne",
        "Origine Nouveau PA", "Fichier source PA",
        "Clé de rapprochement", "Alerte", "PA écarté (à vérifier)",
        "Alerte balayage",
        "VENTE_12M", "ACHAT_12M", "PORTE",
    ]
    sortie = etat.reindex(columns=[c for c in colonnes if c in etat.columns])
    for c in ("Ancien PA", "Nouveau PA", "Écart %",
              "PA écarté (à vérifier)"):
        if c in sortie.columns:
            sortie[c] = pd.to_numeric(sortie[c], errors="coerce").round(4)

    total = len(sortie)
    compte = sortie["Statut PA"].value_counts()
    acquis = int(compte.get("1 — ACQUIS (validé par le responsable)", 0))
    propose = int(compte.get("2 — PROPOSÉ (à relire)", 0))
    print(f"\n{total} items")
    for statut, n in compte.sort_index().items():
        print(f"  {statut:<30} {n:>5}  ({n / total:.1%})")
    print("\nreasons for the items without a price:")
    for motif, n in (sortie.loc[sortie["Statut PA"] == "3 — SANS PRIX",
                                "Motif"].value_counts().items()):
        print(f"  {n:>5}  {motif}")

    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_ACHATS / "pa_etat_par_article.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as writeur:
        sortie.sort_values(["Statut PA", "Fournisseur", "Désignation"]
                           ).to_excel(writeur, sheet_name="Périmètre",
                                      index=False, startrow=3)
        (sortie[sortie["Statut PA"] == "3 — SANS PRIX"]
         .sort_values(["Motif", "Fournisseur"])
         .to_excel(writeur, sheet_name="Sans prix", index=False, startrow=3))
    mise_en_forme.formater(chemin, {
        "Périmètre": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Prix d'achat 2026 — état article par article",
            "sous_titre": (
                f"{total} articles du périmètre figé — "
                f"{acquis} validés par le responsable ({acquis / total:.1%}), "
                f"{propose} proposés ({propose / total:.1%}), "
                f"{total - acquis - propose} sans prix "
                f"({(total - acquis - propose) / total:.1%}). "
                f"Seul « ACQUIS » a été relu.")},
        "Sans prix": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Les articles qui n'ont aucun prix 2026",
            "sous_titre": "Triés par motif : c'est le motif qui dit à qui "
                          "revient l'action"},
    })
    print(f"\n  -> {chemin}")


if __name__ == "__main__":
    sys.exit(main())
